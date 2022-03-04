using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using EcuComparison.Constants;
using EcuComparison.Models;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace EcuComparison.Domain
{
	public class SingleStatisticGenerator
	{
		private readonly ExcelWorkbook _workbook;

		private List<EcuModel> _sheetData;

		private const string Prefix = " - aggregated";

		private Dictionary<string, int> _startingNumberOfEcusPerDx;

		public SingleStatisticGenerator(ExcelWorkbook workbook, List<EcuModel> sheetData)
		{
			_workbook = workbook;
			_sheetData = sheetData;
			_startingNumberOfEcusPerDx = sheetData
				.GroupBy(v => v.Dx)
				.ToDictionary(k => k.Key, v => v.Count());
		}

		public SingleStatisticGenerator(ExcelWorkbook workbook)
		{
			_workbook = workbook;
			_sheetData = new List<EcuModel>();
			_startingNumberOfEcusPerDx = new Dictionary<string, int>();
		}

		public List<EcuModel> SheetData
		{
			get => _sheetData;
			set
			{
				_sheetData = value;
				_startingNumberOfEcusPerDx = _sheetData
					.GroupBy(v => v.Dx)
					.ToDictionary(k => k.Key, v => v.Count());
			}
		}

		public Task GenerateAsync()
		{
			if (_sheetData.Count <= 0) return Task.CompletedTask;
			
			string sheetName = _sheetData.First().SheetName;
			string statisticSheetName = sheetName + Prefix;
			Console.WriteLine($"Writing data for sheet: {statisticSheetName}");
			
			// Generate new sheet
			ExcelWorksheet newSheet = _workbook.Worksheets.Add(statisticSheetName);
			
			// Generate header
			GenerateHeader(newSheet);
			
			// Generate result data
			IEnumerable<AggregatedEcuModel> data = Calculate();
			
			// Write data to sheet
			WriteDataToSheet(newSheet, data);
			
			// Apply table style
			ApplyTableStyle(newSheet);
			newSheet.Cells.AutoFitColumns();

			return Task.CompletedTask;
		}

		private static void ApplyTableStyle(ExcelWorksheet sheet)
		{
			ExcelRange range = sheet.Cells[1, 1, sheet.Dimension.Rows, 6];
			ExcelTable table = sheet.Tables.Add(range, "");
			table.TableStyle = TableStyles.Medium6;
		}

		private static void WriteDataToSheet(ExcelWorksheet sheet, IEnumerable<AggregatedEcuModel> data)
		{
			int row = 2;

			foreach (AggregatedEcuModel model in data)
			{
				sheet.Cells[row, 1].Value = model.Dx;
				sheet.Cells[row, 2].Value = model.StartingNumber;
				sheet.Cells[row, 3].Value = model.NumberOfVariance;
				sheet.Cells[row, 4].Value = model.OverallVariance;
				sheet.Cells[row, 5].Value = model.SwVariance;
				sheet.Cells[row, 6].Value = model.HwVariance;
				row++;
			}
		}

		private static void GenerateHeader(ExcelWorksheet sheet)
		{
			sheet.Cells[1, 1].Value = "Dx";
			sheet.Cells[1, 2].Value = "Starting Number";
			sheet.Cells[1, 3].Value = "Number Of Variance";
			sheet.Cells[1, 4].Value = "Overall Variance";
			sheet.Cells[1, 5].Value = "SW Variance";
			sheet.Cells[1, 6].Value = "HW Variance";
		}
		
		private IEnumerable<AggregatedEcuModel> Calculate()
		{
			// Groups data by Dx and gets statistic
			IEnumerable<AggregatedEcuModel> result = _sheetData
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

					return new AggregatedEcuModel
					{
						Dx = v.Key,
						SwVariance = CheckSoftwareVariance(ecuModels),
						HwVariance = CheckHardwareVariance(ecuModels),
						NumberOfVariance = ecuModels.Count,
						StartingNumber = _startingNumberOfEcusPerDx.GetValueOrDefault(v.Key)
					};
				});

			return result;
		}

		private static string CheckSoftwareVariance(IReadOnlyList<EcuModel> model)
		{
			string sw = model[0].Sw;

			for (int i = 1; i < model.Count; i++)
			{
				if (model[i].Sw != sw)
				{
					return Variance.ModSw;
				}
			}

			return Variance.CopSw;
		}

		private static string CheckHardwareVariance(IReadOnlyList<EcuModel> model)
		{
			string hw = model[0].Hw;

			for (int i = 1; i < model.Count; i++)
			{
				if (model[i].Hw != hw)
				{
					return Variance.ModHw;
				}
			}

			return Variance.CopHw;
		}
	}
}