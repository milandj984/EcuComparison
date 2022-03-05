using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EcuComparison.Constants;
using EcuComparison.Models;
using OfficeOpenXml;

namespace EcuComparison.Domain
{
	public class DoubleStatisticGenerator
	{
		private readonly ExcelWorkbook _workbook;

		public DoubleStatisticGenerator(ExcelWorkbook workbook)
		{
			_workbook = workbook;
			SheetData = new List<EcuModel>();
			SheetData2 = new List<EcuModel>();
		}

		public List<EcuModel> SheetData { get; set; }
		
		public List<EcuModel> SheetData2 { get; set; }

		public Task GenerateAsync()
		{
			if (SheetData.Count <= 0 && SheetData2.Count <= 0) return Task.CompletedTask;

			string sheetName = SheetData[0].SheetName.Replace("te.VBV_", "");
			string sheetName2 = SheetData2[0].SheetName.Replace("te.VBV_", "");
			string statisticSheetName = $"{sheetName} vs {sheetName2}";
			Console.WriteLine($"Writing data for sheet: {statisticSheetName}");

			// Generate new sheet
			ExcelWorksheet newSheet = _workbook.Worksheets.Add(statisticSheetName);

			// Generate header
			GenerateHeader(newSheet);

			// Generate result data
			IEnumerable<AggregatedVsEcuModel> data = Calculate();

			// Write data to sheet
			WriteDataToSheet(newSheet, data);

			// Apply table style
			SingleStatisticGenerator.ApplyTableStyle(newSheet);
			newSheet.Cells.AutoFitColumns();

			return Task.CompletedTask;
		}

		private IEnumerable<AggregatedVsEcuModel> Calculate()
		{
			List<EcuModel> combinedData = SheetData.Concat(SheetData2).ToList();
			Dictionary<string, int> startingNumberOfEcusPerDx = combinedData
				.GroupBy(v => v.Dx)
				.ToDictionary(k => k.Key, v => v.Count());

			Dictionary<string, string> ecuSource = combinedData
				.GroupBy(v => v.Dx)
				.ToDictionary(k => k.Key, v => CheckSourceOrDefault(v.ToList()));
			
			// Groups data by Dx and gets statistic
			IEnumerable<AggregatedVsEcuModel> result = combinedData
				.GroupBy(v => new
				{
					v.Dx,
					v.Sw,
					v.Hw
				})
				.Select(v => v.FirstOrDefault())
				.GroupBy(v => v.Dx)
				.Select(v =>
				{
					List<EcuModel> ecuModels = v.ToList();

					return new AggregatedVsEcuModel
					{
						Dx = v.Key,
						SwVariance = SingleStatisticGenerator.CheckSoftwareVariance(ecuModels),
						HwVariance = SingleStatisticGenerator.CheckHardwareVariance(ecuModels),
						NumberOfVariance = ecuModels.Count,
						StartingNumber = startingNumberOfEcusPerDx.GetValueOrDefault(v.Key),
						Source = ecuSource.GetValueOrDefault(v.Key)
					};
				});

			return result;
		}

		private static void GenerateHeader(ExcelWorksheet sheet)
		{
			sheet.Cells[1, 1].Value = "Dx";
			sheet.Cells[1, 2].Value = "Starting Number";
			sheet.Cells[1, 3].Value = "Number Of Variance";
			sheet.Cells[1, 4].Value = "Source";
			sheet.Cells[1, 5].Value = "Overall Variance";
			sheet.Cells[1, 6].Value = "SW Variance";
			sheet.Cells[1, 7].Value = "HW Variance";
		}

		private static void WriteDataToSheet(ExcelWorksheet sheet, IEnumerable<AggregatedVsEcuModel> data)
		{
			int row = 2;

			foreach (AggregatedVsEcuModel model in data)
			{
				sheet.Cells[row, 1].Value = model.Dx;
				sheet.Cells[row, 2].Value = model.StartingNumber;
				sheet.Cells[row, 3].Value = model.NumberOfVariance;
				sheet.Cells[row, 4].Value = model.Source;
				sheet.Cells[row, 5].Value = model.OverallVariance;
				sheet.Cells[row, 6].Value = model.SwVariance;
				sheet.Cells[row, 7].Value = model.HwVariance;
				row++;
			}
		}

		private string CheckSourceOrDefault(List<EcuModel> ecuModels)
		{
			if (ecuModels.All(v => v.SheetName == SheetData[0].SheetName))
			{
				return VsPairsState.Omitted;
			}
			
			if (ecuModels.All(v => v.SheetName == SheetData2[0].SheetName))
			{
				return VsPairsState.New;
			}

			return default;
		}
	}
}