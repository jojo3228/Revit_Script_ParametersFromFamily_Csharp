using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

namespace ParametersFromFamily
{
    [Transaction(TransactionMode.Manual)]
    public class ExportFamilyParametersToCSV : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получение текущего документа
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // Проверка, является ли документ семейством
            if (!doc.IsFamilyDocument)
            {
                TaskDialog.Show("Ошибка", "Скрипт работает только с документами семейств");
                return Result.Failed;
            }

            // Получение текущего семейства
            FamilyManager familyManager = doc.FamilyManager;

            // Список параметров для экспорта
            var parametersList = new List<ParameterData>();

            var excludedNames = new List<string> {
                "Отметка по умолчанию",
                "URL",
                "Группа модели",
                "Изготовитель",
                "Изображение типоразмера",
                "Ключевая пометка",
                "Код по классификатору",
                "Комментарии к типоразмеру",
                "Описание",
                "Стоимость",
                "Ключ имени сечения",
                "Огнестойкость"
            };

            // Словарь для преобразования значений Group в произвольные значения
            string resourceName = "ParametersFromFamily.groupMappings.json";
            var json = LoadJsonFromEmbeddedResource(resourceName);
            var groupMappings = JsonSerializer.Deserialize<Dictionary<string, string>>(json);

            // Перебор всех параметров семейства
            foreach (FamilyParameter param in familyManager.Parameters)
            {
                // Если у параметра есть формула, пропускаем его
                if (!string.IsNullOrEmpty(param.Formula))
                {
                    continue;
                }

                // Если имя параметра есть в списке исключений, пропускаем его
                if (excludedNames.Contains(param.Definition.Name))
                {
                    continue;
                }

                string group = param.Definition.ParameterGroup.ToString();

                // Получение текущего значения параметра
                string value = GetParameterValue(param, familyManager, doc);

                // Добавление параметра в список
                parametersList.Add(new ParameterData
                {
                    Name = param.Definition.Name,
                    Value = value,
                    DescriptionField = "Добавить описание",
                    ImageField = "Добавить картинку",
                    Group = group,
                    IsInstance = param.IsInstance,
                });
            }

            // Группировка параметров по полю Group
            var groupedParameters = parametersList
                .OrderBy(p => groupMappings.ContainsKey(p.Group) ? groupMappings.Keys.ToList().IndexOf(p.Group) : int.MaxValue)
                .ThenBy(p => p.Name)
                .ToList();

            // Генерация CSV
            var csvData = new List<string>
            {
                "Group,Name,Value,DescriptionField,ImageField,IsInstance" // Заголовки
            };

            foreach (var param in groupedParameters)
            {
                csvData.Add(string.Join(",",
                    EscapeCsv(param.Group),
                    EscapeCsv(param.Name),
                    EscapeCsv(param.Value),
                    EscapeCsv(param.DescriptionField),
                    EscapeCsv(param.ImageField),
                    param.IsInstance.ToString()
                ));
            }

            // Открытие диалогового окна для выбора пути сохранения файла
            string revitFileName = Path.GetFileNameWithoutExtension(doc.Title);
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "CSV files (*.csv)|*.csv",
                Title = "Сохранить файл как",
                FileName = $"{revitFileName}_FamilyParameters_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.csv",
                DefaultExt = "csv",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = saveFileDialog.FileName;

                // Сохранение данных в CSV-файл
                File.WriteAllLines(filePath, csvData, Encoding.UTF8);

                TranslateCsvGroups(filePath, groupMappings);

                TaskDialog.Show("Выполнено", "Файл сохранен");

                return Result.Succeeded;
            }
            else
            {
                TaskDialog.Show("Отмена", "Операция была отменена пользователем.");
                return Result.Cancelled;
            }
        }

    // Метод для получения названия материала или имени файла изображения
    private string GetElementName(Document doc, ElementId id)
        {
            if (id == null || id == ElementId.InvalidElementId)
                return "None";

            Element element = doc.GetElement(id); // Получаем элемент по его ID

            if (element == null)
                return "None";

            // Проверяем, является ли элемент материалом
            if (element is Material material)
            {
                return material.Name; // Возвращаем имя материала
            }

            // Проверяем, является ли элемент изображением
            if (element is ImageType imageType)
            {
                return imageType.Name; // Возвращаем имя изображения
            }

            return element.Name ?? "Unknown"; // Возвращаем общее имя элемента, если это другой тип
        }

    // Основной метод получения значения параметра
    private string GetParameterValue(FamilyParameter param, FamilyManager familyManager, Document doc)
    {
        try
        {
            var value = familyManager.CurrentType.AsValueString(param);
            if (value == null)
            {
                switch (param.StorageType)
                {
                    case StorageType.Double:
                        double? doubleValue = familyManager.CurrentType.AsDouble(param);
                        return doubleValue.HasValue ? doubleValue.Value.ToString("0.##") : "None";

                    case StorageType.Integer:
                        int? intValue = familyManager.CurrentType.AsInteger(param);
                        if (param.Definition.ParameterType == ParameterType.YesNo)
                        {
                            return "Да/Нет";
                        }
                        return intValue.HasValue ? intValue.Value.ToString() : "None";

                        case StorageType.String:
                        return familyManager.CurrentType.AsString(param) ?? "None";

                    case StorageType.ElementId:
                        var id = familyManager.CurrentType.AsElementId(param);
                        return GetElementName(doc, id); // Вызываем метод для получения имени элемента

                    default:
                        return "Unknown";
                }
            }
            return value;
        }
        catch (Exception ex)
        {
            return "Error: " + ex.Message;
        }
    }


    // Метод для экранирования данных в CSV
    private string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                value = $"\"{value.Replace("\"", "\"\"")}\"";
            }
            return value;
        }

        // Метод для перевода групп в сгенерированном CSV
        private void TranslateCsvGroups(string filePath, Dictionary<string, string> groupMappings)
        {
            // Читаем все строки из CSV
            var lines = File.ReadAllLines(filePath, Encoding.UTF8).ToList();

            // Обрабатываем строки начиная со второй (пропуская заголовок)
            for (int i = 1; i < lines.Count; i++)
            {
                var columns = lines[i].Split(',');

                if (columns.Length > 0)
                {
                    string groupName = columns[0]; // Первая колонка - это имя группы

                    // Если группа есть в словаре, заменяем её на русский
                    if (groupMappings.ContainsKey(groupName))
                    {
                        columns[0] = groupMappings[groupName];
                        lines[i] = string.Join(",", columns); // Обновляем строку
                    }
                }
            }

            // Перезаписываем CSV с обновлёнными данными
            File.WriteAllLines(filePath, lines, Encoding.UTF8);
        }

        private string LoadJsonFromEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new FileNotFoundException($"Ресурс {resourceName} не найден в сборке.");
                }
                using (var reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }

    // Класс для представления параметра
    public class ParameterData
    {
        public string Name { get; set; }
        public string DescriptionField { get; set; }
        public string ImageField { get; set; }
        public string Group { get; set; }
        public bool IsInstance { get; set; }
        public string Value { get; set; }
    }
}


