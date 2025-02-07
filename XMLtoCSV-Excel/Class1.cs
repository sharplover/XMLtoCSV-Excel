using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace XMLtoCSV_Excel
{
    public class XmlProcessor
    {

        public void ProcessXmlToCsvAndExcel(
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
            string filePath,
            bool isOddSelected,
            bool isEvenSelected,
            bool isCombinedSelected,
            bool isNormativeSelected,
            bool isVariantSelected,
            bool isFirstDecadeSelected,
            bool isSecondDecadeSelected,
            bool isThirdDecadeSelected,
            bool isWholeMonthSelected,
            bool isMainWaySelected,
            bool isIndicatorsSelected,
            bool isDailyIndicatorsSelected)
        {
            // Загружаем XML
            var xmlDoc = XDocument.Load(filePath);
            var sourceData = ParseXmlData(xmlDoc);

            // Применяем выбранные параметры
            if (isFirstDecadeSelected) FillFirstDecade(sourceData);
            if (isSecondDecadeSelected) FillSecondDecade(sourceData);
            if (isThirdDecadeSelected) FillThirdDecade(sourceData);
            if (isWholeMonthSelected) AddMonthlySummary(sourceData);
            if (isIndicatorsSelected) var (totalOddSourceSpeed, totalEvenSourceSpeed, totalCombinedSourceSpeed,
     totalOddVariantSpeed, totalEvenVariantSpeed, totalCombinedVariantSpeed) = AddAllRoad(sourceData);

            // Логирование
            Console.WriteLine("Обработанные данные:");
            foreach (var key in sourceData.Keys)
            {
                Console.WriteLine($"Ключ: {key}, Значения: {sourceData[key]}");
            }

            // Формируем пути для сохранения
            string csvFilePath = filePath.Replace(".xml", ".csv");
            string excelFilePath = filePath.Replace(".xml", ".xlsx");

            // Сохранение в CSV и Excel
            SaveToCsv(csvFilePath, sourceData, isOddSelected, isEvenSelected, isCombinedSelected, isNormativeSelected, isVariantSelected);
            SaveToExcel(excelFilePath, sourceData, isOddSelected, isEvenSelected, isCombinedSelected, isNormativeSelected, isVariantSelected);
        }

        // Метод для сохранения данных в CSV с учётом активных параметров
        private void SaveToCsv(string csvFilePath, Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
                               bool enableOdd, bool enableEven, bool enableCombined, bool enableSourceSpeed, bool enableVariantSpeed)
        {
            using (var writer = new StreamWriter(csvFilePath, false, System.Text.Encoding.UTF8))
            {
                writer.WriteLine("Участок;Дата;Тип параметра;Значение");

                foreach (var date in data.Keys)
                {
                    foreach (var area in data[date].Keys)
                    {
                        var values = data[date][area];

                        if (enableOdd) writer.WriteLine($"{area};{date};OddSource;{Math.Round(values.OddSource, 1)}");
                        if (enableEven) writer.WriteLine($"{area};{date};EvenSource;{Math.Round(values.EvenSource, 1)}");
                        if (enableCombined) writer.WriteLine($"{area};{date};CombinatedSourceThreads;{Math.Round(values.CombinatedSourceThreads, 1)}");
                        if (enableOdd && enableVariantSpeed) writer.WriteLine($"{area};{date};OddVariant;{Math.Round(values.OddVariant, 1)}");
                        if (enableEven && enableVariantSpeed) writer.WriteLine($"{area};{date};EvenVariant;{Math.Round(values.EvenVariant, 1)}");
                        if (enableCombined && enableVariantSpeed) writer.WriteLine($"{area};{date};CombinatedVariantThreads;{Math.Round(values.CombinatedVariantThreads, 1)}");
                    }
                }
            }

            Console.WriteLine("CSV файл сохранён.");
        }

        // Метод для сохранения данных в Excel с учётом активных параметров
        private void SaveToExcel(
            string excelFilePath,
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
            bool enableOdd,
            bool enableEven,
            bool enableCombined,
            bool enableSourceSpeed,
            bool enableVariantSpeed)
        {
            using (var package = new ExcelPackage())
            {
                // Лист для нормативных скоростей (Source)
                if (enableSourceSpeed)
                {
                    var sourceSheet = package.Workbook.Worksheets.Add("Source");
                    FillSheet(sourceSheet, data, enableOdd, enableEven, enableCombined, isSource: true);
                }

                // Лист для вариантных скоростей (Variant)
                if (enableVariantSpeed)
                {
                    var variantSheet = package.Workbook.Worksheets.Add("Variant");
                    FillSheet(variantSheet, data, enableOdd, enableEven, enableCombined, isSource: false);
                }

                package.SaveAs(new FileInfo(excelFilePath));
            }
            Console.WriteLine("Excel файл сохранён.");
        }

        private void FillSheet(
                   ExcelWorksheet sheet,
                   Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads,
                       double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
                   bool enableOdd,
                   bool enableEven,
                   bool enableCombined,
                   bool isSource)
        {
            try
            {
                // Очищаем лист и удаляем все объединения
                sheet.Cells[sheet.Dimension?.Address ?? "A1"].Clear();
                RemoveAllMergedCells(sheet);

                CreateHeader(sheet, data.Keys, enableOdd, enableEven, enableCombined, isSource);
                FillDataRows(sheet, data, enableOdd, enableEven, enableCombined, isSource);

                // Автоподбор ширины столбцов
                sheet.Cells.AutoFitColumns();
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Ошибка при заполнении листа {sheet.Name}", ex);
            }
        }

        private void RemoveAllMergedCells(ExcelWorksheet sheet)
        {
            var mergedCells = sheet.MergedCells.ToList();
            foreach (var address in mergedCells)
            {
                sheet.Cells[address].Merge = false;
            }
        }

        private void CreateHeader(
            ExcelWorksheet sheet,
            IEnumerable<string> dates,
            bool enableOdd,
            bool enableEven,
            bool enableCombined,
            bool isSource)
        {
            int column = 1;
            sheet.Cells[1, column].Value = "Участок";
            sheet.Column(column).Width = 25;

            foreach (var date in dates.OrderBy(d => d))
            {
                int columnsCount = 0;
                if (enableCombined) columnsCount++;
                if (enableOdd) columnsCount++;
                if (enableEven) columnsCount++;

                if (columnsCount == 0) continue;

                // Объединение ячеек для даты
                var headerCell = sheet.Cells[1, column + 1, 1, column + columnsCount];
                headerCell.Merge = true;
                headerCell.Value = date;
                headerCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                // Подзаголовки
                int subCol = column + 1;
                if (enableCombined)
                {
                    sheet.Cells[2, subCol].Value = isSource ? "Комбинир." : "Вариант комб.";
                    subCol++;
                }
                if (enableOdd)
                {
                    sheet.Cells[2, subCol].Value = isSource ? "Нечетный" : "Вариант неч.";
                    subCol++;
                }
                if (enableEven)
                {
                    sheet.Cells[2, subCol].Value = isSource ? "Четный" : "Вариант чет.";
                    subCol++;
                }

                column += columnsCount;
            }

            // Стиль заголовков
            using (var headerRange = sheet.Cells[1, 1, 2, column])
            {
                headerRange.Style.Font.Bold = true;
                headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                headerRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            }
        }

        private void FillDataRows(
            ExcelWorksheet sheet,
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads,
                double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
            bool enableOdd,
            bool enableEven,
            bool enableCombined,
            bool isSource)
        {
            int row = 3;

            foreach (var area in data.First().Value.Keys.OrderBy(a => a))
            {
                sheet.Cells[row, 1].Value = area;
                int column = 2;

                foreach (var date in data.Keys.OrderBy(d => d))
                {
                    if (!data[date].ContainsKey(area)) continue;

                    var values = data[date][area];
                    int startCol = column;

                    try
                    {
                        if (enableCombined)
                        {
                            sheet.Cells[row, column++].Value = Math.Round(isSource
                                ? values.CombinatedSourceThreads
                                : values.CombinatedVariantThreads, 1);
                        }
                        if (enableOdd)
                        {
                            sheet.Cells[row, column++].Value = Math.Round(isSource
                                ? values.OddSource
                                : values.OddVariant, 1);
                        }
                        if (enableEven)
                        {
                            sheet.Cells[row, column++].Value = Math.Round(isSource
                                ? values.EvenSource
                                : values.EvenVariant, 1);
                        }

                        // Объединение если несколько значений
                        if (column - startCol > 1)
                        {
                            sheet.Cells[row, startCol, row, column - 1].Merge = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"Ошибка в строке {row}, колонки {startCol}-{column} | {ex.Message}");
                    }
                }
                row++;
            }
        }


        public Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> ParseXmlData(XDocument xmlDoc)
        {
            var result = new Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>>();
            var originalDataDictionary = new Dictionary<string, Dictionary<string, (double OddSourceKm, double OddSourceHours, double EvenSourceKm, double EvenSourceHours, double OddVariantKm, double OddVariantHours, double EvenVariantKm, double EvenVariantHours)>>();

            foreach (var areaElement in xmlDoc.Descendants("Area"))
            {
                string areaName = areaElement.Attribute("Name")?.Value;
                if (string.IsNullOrEmpty(areaName)) continue;

                foreach (var cellElement in areaElement.Descendants("Cell"))
                {
                    string dateStr = cellElement.Attribute("Date")?.Value;
                    if (string.IsNullOrEmpty(dateStr))
                    {
                        Console.WriteLine("Не найдено значение атрибута 'Date'");
                        continue;
                    }

                    if (!DateTime.TryParseExact(dateStr, "dd.MM.yyyy HH:mm:ss", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) &&
                        !DateTime.TryParse(dateStr, out date))
                    {
                        Console.WriteLine($"Ошибка при парсинге даты: {dateStr}");
                        continue;
                    }

                    string dateFormatted = date.ToString("yyyy-MM-dd");

                    if (!result.ContainsKey(dateFormatted))
                        result[dateFormatted] = new Dictionary<string, (double, double, double, double, double, double)>();

                    if (!result[dateFormatted].ContainsKey(areaName))
                        result[dateFormatted][areaName] = (0, 0, 0, 0, 0, 0);

                    var cellInfoElement = cellElement.Element("CellInformation");

                    double oddSourceKm = 0, oddSourceHours = 0;
                    double evenSourceKm = 0, evenSourceHours = 0;
                    double oddVariantKm = 0, oddVariantHours = 0;
                    double evenVariantKm = 0, evenVariantHours = 0;

                    var oddSourceElement = cellInfoElement.Element("OddSourceThreads");
                    var evenSourceElement = cellInfoElement.Element("EvenSourceThreads");
                    var oddVariantElement = cellInfoElement.Element("OddVariantThreads");
                    var evenVariantElement = cellInfoElement.Element("EvenVariantThreads");

                    if (oddSourceElement != null)
                    {
                        foreach (var section in oddSourceElement.Descendants("SectionInformations"))
                        {
                            oddSourceKm += double.Parse(section.Attribute("TrainKm")?.Value ?? "0");
                            oddSourceHours += double.Parse(section.Attribute("TrainHoursWithStops")?.Value ?? "0");
                        }
                    }

                    if (evenSourceElement != null)
                    {
                        foreach (var section in evenSourceElement.Descendants("SectionInformations"))
                        {
                            evenSourceKm += double.Parse(section.Attribute("TrainKm")?.Value ?? "0");
                            evenSourceHours += double.Parse(section.Attribute("TrainHoursWithStops")?.Value ?? "0");
                        }
                    }

                    if (oddVariantElement != null)
                    {
                        foreach (var section in oddVariantElement.Descendants("SectionInformations"))
                        {
                            oddVariantKm += double.Parse(section.Attribute("TrainKm")?.Value ?? "0");
                            oddVariantHours += double.Parse(section.Attribute("TrainHoursWithStops")?.Value ?? "0");
                        }
                    }

                    if (evenVariantElement != null)
                    {
                        foreach (var section in evenVariantElement.Descendants("SectionInformations"))
                        {
                            evenVariantKm += double.Parse(section.Attribute("TrainKm")?.Value ?? "0");
                            evenVariantHours += double.Parse(section.Attribute("TrainHoursWithStops")?.Value ?? "0");
                        }
                    }

                    double combinedSourceThreads = (oddSourceHours + evenSourceHours) > 0
                        ? (oddSourceKm + evenSourceKm) / (oddSourceHours + evenSourceHours)
                        : 0;

                    double combinedVariantThreads = (oddVariantHours + evenVariantHours) > 0
                        ? (oddVariantKm + evenVariantKm) / (oddVariantHours + evenVariantHours)
                        : 0;

                    result[dateFormatted][areaName] = (
                        oddSourceKm / Math.Max(oddSourceHours, 1),
                        evenSourceKm / Math.Max(evenSourceHours, 1),
                        combinedSourceThreads,
                        oddVariantKm / Math.Max(oddVariantHours, 1),
                        evenVariantKm / Math.Max(evenVariantHours, 1),
                        combinedVariantThreads
                    );

                    // Сохраняем оригинальные данные в отдельный словарь для дальнейшего использования
                    if (!originalDataDictionary.ContainsKey(dateFormatted))
                        originalDataDictionary[dateFormatted] = new Dictionary<string, (double, double, double, double, double, double, double, double)>();

                    originalDataDictionary[dateFormatted][areaName] = (
                        oddSourceKm,          // OddSource Km
                        oddSourceHours,       // OddSource Hours
                        evenSourceKm,         // EvenSource Km
                        evenSourceHours,      // EvenSource Hours
                        oddVariantKm,         // OddVariant Km
                        oddVariantHours,      // OddVariant Hours
                        evenVariantKm,        // EvenVariant Km
                        evenVariantHours      // EvenVariant Hours
                    );
                }
            }

            // Теперь передаем оригинальные данные из словаря в HandleMainWay
            foreach (var dateFormatted in originalDataDictionary.Keys)
            {

                //  Console.WriteLine($"dateFormatted: {dateFormatted}");

                // Суммарные данные для пути
                var startWayData = (0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0);

                foreach (var areaName in originalDataDictionary[dateFormatted].Keys)
                {

                    //    Console.WriteLine($"areaName: {areaName}");

                    var isMainWay = xmlDoc.Descendants("Area")
                        .Where(a => a.Attribute("Name")?.Value == areaName)
                        .FirstOrDefault()?
                        .Attribute("IsMainWay")?.Value;


                    //   Console.WriteLine($"isMainWay: {isMainWay}");

                    if (isMainWay == "1")
                    {

                        // Суммируем данные для каждого участка с использованием originalData
                        var (oddSourceKm, oddSourceHours, evenSourceKm, evenSourceHours, oddVariantKm, oddVariantHours, evenVariantKm, evenVariantHours) = originalDataDictionary[dateFormatted][areaName];




                        // Суммируем значения для MainWay
                        startWayData = (
                            startWayData.Item1 + oddSourceKm,
                            startWayData.Item2 + oddSourceHours,
                            startWayData.Item3 + evenSourceKm,
                            startWayData.Item4 + evenSourceHours,
                            startWayData.Item5 + oddVariantKm,
                            startWayData.Item6 + oddVariantHours,
                            startWayData.Item7 + evenVariantKm,
                            startWayData.Item8 + evenVariantHours
                        );

                        //    Console.WriteLine($"startWayData: {startWayData}");


                        // Расчет комбинированных значений для SourceThreads и VariantThreads
                        var combinedSourceThreads = (startWayData.Item2 + startWayData.Item4) > 0
                            ? (startWayData.Item1 + startWayData.Item3) / (startWayData.Item2 + startWayData.Item4)
                            : 0;

                        var combinedVariantThreads = (startWayData.Item6 + startWayData.Item8) > 0
                            ? (startWayData.Item5 + startWayData.Item7) / (startWayData.Item6 + startWayData.Item8)
                            : 0;

                        // Добавляем данные для MainWay
                        result[dateFormatted]["MainWay"] = (
                startWayData.Item1 / Math.Max(startWayData.Item2, 1),
                startWayData.Item3 / Math.Max(startWayData.Item4, 1),
                combinedSourceThreads,
                startWayData.Item5 / Math.Max(startWayData.Item6, 1),
                startWayData.Item7 / Math.Max(startWayData.Item8, 1),
                combinedVariantThreads
            );
                    }
                }
            }


            return result;
        }

        // Методы для направлений
        public void FillData(Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
                             bool enableOdd, bool enableEven, bool enableCombined, bool enableSourceSpeed, bool enableVariantSpeed)
        {
            foreach (var dateEntry in data)
            {
                foreach (var areaEntry in dateEntry.Value.Keys.ToList()) // Копируем ключи, чтобы избежать изменений при итерации
                {
                    var values = dateEntry.Value[areaEntry];

                    double oddSource = enableOdd ? values.OddSource : 0;
                    double evenSource = enableEven ? values.EvenSource : 0;
                    double combinedSource = enableCombined ? values.CombinatedSourceThreads : 0;
                    double oddVariant = enableOdd && enableVariantSpeed ? values.OddVariant : 0;
                    double evenVariant = enableEven && enableVariantSpeed ? values.EvenVariant : 0;
                    double combinedVariant = enableCombined && enableVariantSpeed ? values.CombinatedVariantThreads : 0;

                    data[dateEntry.Key][areaEntry] = (oddSource, evenSource, combinedSource, oddVariant, evenVariant, combinedVariant);
                }
            }

            Console.WriteLine("Данные заполнены с учётом выбранных параметров.");
        }

        // Методы для периодов
        // Метод для обработки первой декады
        public void FillFirstDecade(Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var targetMonth = GetTargetMonth(data);
            var firstDecadeDates = GetDatesForDecade(data, targetMonth, 1, 10);
            var result = ProcessDecade(firstDecadeDates, data);

            Console.WriteLine("Первая декада обработана.");
        }

        // Метод для обработки второй декады
        public void FillSecondDecade(Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var targetMonth = GetTargetMonth(data);
            var secondDecadeDates = GetDatesForDecade(data, targetMonth, 11, 20);
            var result = ProcessDecade(secondDecadeDates, data);

            Console.WriteLine("Вторая декада обработана.");
        }

        // Метод для обработки третьей декады
        public void FillThirdDecade(Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var targetMonth = GetTargetMonth(data);
            var daysInMonth = DateTime.DaysInMonth(targetMonth.Year, targetMonth.Month);
            var thirdDecadeDates = GetDatesForDecade(data, targetMonth, 21, daysInMonth);
            var result = ProcessDecade(thirdDecadeDates, data);

            Console.WriteLine("Третья декада обработана.");
        }

        // Вспомогательный метод: определение основного месяца
        private (int Year, int Month) GetTargetMonth(
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var monthCounts = data.Keys
                .Select(date => DateTime.Parse(date))
                .GroupBy(date => (date.Year, date.Month)) // Используем кортеж вместо анонимного типа
                .ToDictionary(g => g.Key, g => g.Count());

            return monthCounts.OrderByDescending(kvp => kvp.Value).First().Key;
        }


        // Вспомогательный метод: получение дат для указанной декады
        private List<DateTime> GetDatesForDecade(
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data,
            (int Year, int Month) targetMonth,
            int startDay,
            int endDay)
        {
            return data.Keys
                .Select(date => DateTime.Parse(date))
                .Where(date => date.Year == targetMonth.Year && date.Month == targetMonth.Month && date.Day >= startDay && date.Day <= endDay)
                .OrderBy(date => date)
                .ToList();
        }

        // Вспомогательный метод: обработка данных за декаду
        private Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)> ProcessDecade(
            List<DateTime> dates,
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var aggregatedData = new Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>();

            foreach (var area in data.First().Value.Keys)
            {
                double oddSourceKmSum = 0.0;
                double oddSourceHoursSum = 0.0;
                double evenSourceKmSum = 0.0;
                double evenSourceHoursSum = 0.0;
                double oddVariantKmSum = 0.0;
                double oddVariantHoursSum = 0.0;
                double evenVariantKmSum = 0.0;
                double evenVariantHoursSum = 0.0;

                foreach (var date in dates)
                {
                    var dateKey = date.ToString("yyyy-MM-dd");
                    if (data.TryGetValue(dateKey, out var dayData) && dayData.TryGetValue(area, out var values))
                    {
                        oddSourceKmSum += values.OddSource * values.OddSource;
                        oddSourceHoursSum += values.OddSource;

                        evenSourceKmSum += values.EvenSource * values.EvenSource;
                        evenSourceHoursSum += values.EvenSource;

                        oddVariantKmSum += values.OddVariant * values.OddVariant;
                        oddVariantHoursSum += values.OddVariant;

                        evenVariantKmSum += values.EvenVariant * values.EvenVariant;
                        evenVariantHoursSum += values.EvenVariant;
                    }
                }

                double oddSourceSpeed = oddSourceHoursSum > 0 ? Math.Round(oddSourceKmSum / oddSourceHoursSum, 1) : 0;
                double evenSourceSpeed = evenSourceHoursSum > 0 ? Math.Round(evenSourceKmSum / evenSourceHoursSum, 1) : 0;

                double oddVariantSpeed = oddVariantHoursSum > 0 ? Math.Round(oddVariantKmSum / oddVariantHoursSum, 1) : 0;
                double evenVariantSpeed = evenVariantHoursSum > 0 ? Math.Round(evenVariantKmSum / evenVariantHoursSum, 1) : 0;

                double combinedSourceThreads = (oddSourceKmSum + evenSourceKmSum) /
                                               (oddSourceHoursSum + evenSourceHoursSum > 0 ? oddSourceHoursSum + evenSourceHoursSum : 1);

                double combinedVariantThreads = (oddVariantKmSum + evenVariantKmSum) /
                                                (oddVariantHoursSum + evenVariantHoursSum > 0 ? oddVariantHoursSum + evenVariantHoursSum : 1);

                aggregatedData[area] = (oddSourceSpeed, evenSourceSpeed, combinedSourceThreads, oddVariantSpeed, evenVariantSpeed, combinedVariantThreads);
            }

            return aggregatedData;
        }

        public void AddMonthlySummary(Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            var monthlySummary = new Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>();

            // Проходим по всем участкам
            foreach (var areaName in data.First().Value.Keys)
            {
                double oddSourceKmSum = 0.0;
                double oddSourceHoursSum = 0.0;
                double evenSourceKmSum = 0.0;
                double evenSourceHoursSum = 0.0;
                double oddVariantKmSum = 0.0;
                double oddVariantHoursSum = 0.0;
                double evenVariantKmSum = 0.0;
                double evenVariantHoursSum = 0.0;

                foreach (var date in data.Keys)
                {
                    if (data[date].TryGetValue(areaName, out var values))
                    {
                        oddSourceKmSum += values.OddSource * values.OddSource;
                        oddSourceHoursSum += values.OddSource;

                        evenSourceKmSum += values.EvenSource * values.EvenSource;
                        evenSourceHoursSum += values.EvenSource;

                        oddVariantKmSum += values.OddVariant * values.OddVariant;
                        oddVariantHoursSum += values.OddVariant;

                        evenVariantKmSum += values.EvenVariant * values.EvenVariant;
                        evenVariantHoursSum += values.EvenVariant;
                    }
                }

                // Рассчитываем итоговые скорости
                double oddSourceSpeed = oddSourceHoursSum > 0 ? Math.Round(oddSourceKmSum / oddSourceHoursSum, 1) : 0;
                double evenSourceSpeed = evenSourceHoursSum > 0 ? Math.Round(evenSourceKmSum / evenSourceHoursSum, 1) : 0;

                double oddVariantSpeed = oddVariantHoursSum > 0 ? Math.Round(oddVariantKmSum / oddVariantHoursSum, 1) : 0;
                double evenVariantSpeed = evenVariantHoursSum > 0 ? Math.Round(evenVariantKmSum / evenVariantHoursSum, 1) : 0;

                double combinedSourceThreads = (oddSourceKmSum + evenSourceKmSum) /
                                               (oddSourceHoursSum + evenSourceHoursSum > 0 ? oddSourceHoursSum + evenSourceHoursSum : 1);

                double combinedVariantThreads = (oddVariantKmSum + evenVariantKmSum) /
                                                (oddVariantHoursSum + evenVariantHoursSum > 0 ? oddVariantHoursSum + evenVariantHoursSum : 1);

                monthlySummary[areaName] = (oddSourceSpeed, evenSourceSpeed, combinedSourceThreads, oddVariantSpeed, evenVariantSpeed, combinedVariantThreads);
            }

            // Добавляем итог за месяц в словарь
            data.Add("Итог за месяц", monthlySummary);
        }

        // Дополнительные методы
        public void FillMainWay()
        {
            // Логика обработки данных по главному ходу из оригинального кода
        }

        public (double totalOddSourceSpeed, double totalEvenSourceSpeed, double totalCombinedSourceSpeed, double totalOddVariantSpeed, double totalEvenVariantSpeed, double totalCombinedVariantSpeed) AddAllRoad(
            Dictionary<string, Dictionary<string, (double OddSource, double EvenSource, double CombinatedSourceThreads, double OddVariant, double EvenVariant, double CombinatedVariantThreads)>> data)
        {
            // Инициализация сумм километров и часов
            double oddSourceKmSum = 0.0, oddSourceHoursSum = 0.0;
            double evenSourceKmSum = 0.0, evenSourceHoursSum = 0.0;
            double oddVariantKmSum = 0.0, oddVariantHoursSum = 0.0;
            double evenVariantKmSum = 0.0, evenVariantHoursSum = 0.0;

            // Суммируем данные по всем датам и участкам
            foreach (var dateEntry in data)
            {
                foreach (var areaEntry in dateEntry.Value)
                {
                    if (areaEntry.Key == "MainWay")
                        continue; // Игнорируем "MainWay"

                    var values = areaEntry.Value;

                    // Суммируем километры и часы для OddSource и EvenSource
                    oddSourceKmSum += values.OddSource * values.OddSource;
                    oddSourceHoursSum += values.OddSource;

                    evenSourceKmSum += values.EvenSource * values.EvenSource;
                    evenSourceHoursSum += values.EvenSource;

                    // Суммируем километры и часы для OddVariant и EvenVariant
                    oddVariantKmSum += values.OddVariant * values.OddVariant;
                    oddVariantHoursSum += values.OddVariant;

                    evenVariantKmSum += values.EvenVariant * values.EvenVariant;
                    evenVariantHoursSum += values.EvenVariant;
                }
            }


            // Рассчитываем итоговые скорости
            double totalOddSourceSpeed = oddSourceHoursSum > 0 ? oddSourceKmSum / oddSourceHoursSum : 0;
            double totalEvenSourceSpeed = evenSourceHoursSum > 0 ? evenSourceKmSum / evenSourceHoursSum : 0;
            double totalCombinedSourceSpeed = (oddSourceHoursSum + evenSourceHoursSum) > 0
                ? (oddSourceKmSum + evenSourceKmSum) / (oddSourceHoursSum + evenSourceHoursSum)
                : 0;

            double totalOddVariantSpeed = oddVariantHoursSum > 0 ? oddVariantKmSum / oddVariantHoursSum : 0;
            double totalEvenVariantSpeed = evenVariantHoursSum > 0 ? evenVariantKmSum / evenVariantHoursSum : 0;
            double totalCombinedVariantSpeed = (oddVariantHoursSum + evenVariantHoursSum) > 0
                ? (oddVariantKmSum + evenVariantKmSum) / (oddVariantHoursSum + evenVariantHoursSum)
                : 0;

            // Возвращаем итоговые скорости
            return (totalOddSourceSpeed, totalEvenSourceSpeed, totalCombinedSourceSpeed, totalOddVariantSpeed, totalEvenVariantSpeed, totalCombinedVariantSpeed);
        }

    }

}
