
DateTime beginDate = DateTime.Now;

var sheet = WorkbookNonExcel.Worksheets.FirstOrDefault();

string attr1 = "Источник затрат установки ПУ";
string attr1_val = ReportParams.attr1;

// Максимальное количество данных
var maxDataCount = ReportParams.maxDataCount;

// Дата начала интервала
var startDate = ReportParams.startDate;
// Количество месяцев
int monthsCount = ReportParams.monthsCount;
// Дата окончания интервала
var stopDate = startDate.AddMonths(monthsCount);

// Выводимые параметры
var parametr = TariffZoneBasedParameter.Instances.EnergyActiveForwardMonth; // Энергия А+ за месяц

// Первая строка таблицы
var startRow = 5;
// Первая строка
int row = 1;
// Первый стобец с данными
var lastCol = 7;

foreach (ClassifierItem classifierItem in ReportParams.ClassifierItems)
{
	// Цикл по точкам учёта
	foreach (MeterPoint meterPoint in classifierItem.GetAllChildrenOfClass(MeterPoint.GetClassInfo()))
	{
		if (meterPoint.AttributeConsumer != null) continue;
		var electricityMeter = meterPoint.AttributeElectricityMeter == null ? "Нет прибора" : meterPoint.AttributeElectricityMeter.Caption;
		if (electricityMeter == "Нет прибора") continue;
		bool hasMeterModel = false;
		if (ReportParams.isMeterMt880 || ReportParams.isMeterAm550) {
			if (ReportParams.isMeterMt880)
			{
				if (meterPoint.AttributeElectricityMeter is Mt880Meter)
					hasMeterModel = true;
			}
			if (ReportParams.isMeterAm550)
			{
				if (meterPoint.AttributeElectricityMeter is Am550EMeter || meterPoint.AttributeElectricityMeter is Am550TMeter)
					hasMeterModel = true;
			}
		} else hasMeterModel = true;
		
		if (!hasMeterModel) continue;
		string attr1Src = meterPoint.ReadValueByAttributeCaption(attr1) != null ? meterPoint.ReadValueByAttributeCaption(attr1).ToString() : "н/д";
		if (attr1Src != attr1_val) continue;
		// Получение данных по точке учёта
		var recCount = 0;
		double dataValueSum = -1.0;
		string dataInfo = "";
		for (int i = 1; i <= monthsCount; i++)
		{
			var interval = new DayIntervalData(){ StartDt = startDate.AddMonths(i - 1), EndDt = startDate.AddMonths(i).AddDays(0) }; // .AddDays(-1)
			// Получение данных по точке учёта
			// var data = meterPoint.GetMeterPointFinalData(parametr, interval);
			var data = meterPoint.GetMeterPointFinalData(parametr, interval).ToArray();
			// var record = data.FirstOrDefault();
			var record = data.FirstOrDefault(x => x.ValueDt != null);
			if (record != null) {
				double dataValue = record.Value / 1000.0;
				// Тесты
				dataInfo = String.Concat(dataInfo, string.Format("{0} [{1} - {2}]", dataValue.ToString(), startDate.AddMonths(i - 1).ToString(), startDate.AddMonths(i).AddDays(0).ToString()), "; ");
				if (dataValueSum == -1.0) dataValueSum = 0;
				dataValueSum = dataValueSum + dataValue;
				recCount++;
			}
		}
		if (dataValueSum > -1.0 && dataValueSum <= 1.0)
		{
			sheet.Cells[startRow + row, 0].Value = row.ToString();
			sheet.Cells[startRow + row, 1].Value = meterPoint.Caption;
			sheet.Cells[startRow + row, 2].Value = meterPoint.AttributeElectricityMeter.AttributeSerialNumber.ToString();
			sheet.Cells[startRow + row, 3].Value = meterPoint.AttributeElectricityMeter.ToString();
			sheet.Cells[startRow + row, 4].Value = attr1Src;
			sheet.Cells[startRow + row, 5].Value = dataValueSum.ToString();
			sheet.Cells[startRow + row, 6].Value = recCount.ToString();
			sheet.Cells[startRow + row, 7].Value = dataInfo.ToString();

			row++;

			if (maxDataCount > 0)
				if (row >= maxDataCount) break;
		}
	}
}

int rowsCount = row - 1;

// Установить рамки вокруг всех ячеек
sheet.Cells.GetSubrangeAbsolute(startRow, 0, startRow + rowsCount, lastCol).SetBorder();
// В заголовке таблицы установить жирный шрифт и устанавливаем цвет ячеек
var header = sheet.Cells.GetSubrangeAbsolute(startRow, 0, startRow, lastCol);
header.Style.FillPattern.SetSolid(Color.FromArgb(0, 189, 215, 238));
header.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
header.Style.VerticalAlignment = VerticalAlignmentStyle.Center;
// Переносим стороки
header.Style.WrapText = true;
// Заголовки
sheet.Cells[startRow, 0].Value = "№";
sheet.Cells[startRow, 1].Value = "Наименование ТУ";
sheet.Cells[startRow, 2].Value = "Номер ПУ";
sheet.Cells[startRow, 3].Value = "Тип ПУ";
sheet.Cells[startRow, 4].Value = attr1;
sheet.Cells[startRow, 5].Value = parametr.Caption;
sheet.Cells[startRow, 6].Value = "Кол-во мес-ев";
sheet.Cells[startRow, 7].Value = "Данные";

// Включить автоматическую ширину первого столбца
for (int i = 0; i <= lastCol; i++)
	sheet.Columns[i].AutoFit();
sheet.Columns[1].Width = 15*1024;
sheet.Columns[1].Style.WrapText = true;
sheet.Columns[7].Width = 25*1024;
sheet.Columns[7].Style.WrapText = true;

// Выводим информацию по параметрам формирования отчёта
List<string> listClassifierItems = new List<string>();
foreach(var entry in ReportParams.ClassifierItems)
{
	if (entry != null)
		listClassifierItems.Add(entry.Caption);
}

sheet.Cells[0, 0].Value = string.Format("Параметр: «{0}». Атрибут: «{1}»=«{2}»", parametr.Caption, attr1, attr1_val);
sheet.Cells[1, 0].Value = string.Format("За период с {0:dd.MM.yyyy H:mm:ss} по {1:dd.MM.yyyy H:mm:ss}", startDate, stopDate);
sheet.Cells[2, 0].Value = string.Format("Элемент(ы) классификатора: {0}. Это МТ880?: {1}. Это АМ550?: {2}", string.Join("; ", listClassifierItems), ReportParams.isMeterMt880.ToString(), ReportParams.isMeterAm550.ToString());
sheet.Cells[3, 0].Value = string.Format("Сформирован {0:dd.MM.yyyy H:mm:ss}. Старт {1:dd.MM.yyyy H:mm:ss}. Раница: {2}:{3} ", DateTime.Now, beginDate, (DateTime.Now - beginDate).Hours, (DateTime.Now - beginDate).Minutes);
sheet.Cells.GetSubrangeAbsolute(0, 0, startRow - 2, 0).Style.Font.Weight = ExcelFont.BoldWeight;
