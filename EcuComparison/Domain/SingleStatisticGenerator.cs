using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EcuComparison.Constants;
using EcuComparison.Models;

namespace EcuComparison.Domain
{
	public class SingleStatisticGenerator
	{
		private readonly XLWorkbook _workbook;

		private List<EcuModel> _sheetData;

		private const string Prefix = " - aggregated";

		private Dictionary<string, int> _startingNumberOfEcusPerDx;

		public SingleStatisticGenerator(XLWorkbook workbook, List<EcuModel> sheetData)
		{
			_workbook = workbook;
			_sheetData = sheetData;
			_startingNumberOfEcusPerDx = sheetData
				.GroupBy(v => v.Dx)
				.ToDictionary(k => k.Key, v => v.Count());
		}

		public SingleStatisticGenerator(XLWorkbook workbook)
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
			IXLWorksheet newSheet = _workbook.AddWorksheet(statisticSheetName);
			
			// Generate header
			GenerateHeader(newSheet);
			
			// Generate result data
			IEnumerable<AggregatedEcuModel> data = Calculate();
			
			// Write data to sheet
			WriteDataToSheet(data, newSheet);

			return Task.CompletedTask;
		}

		private static void WriteDataToSheet(IEnumerable<AggregatedEcuModel> data, IXLWorksheet sheet)
		{
			int row = 2;

			foreach (AggregatedEcuModel model in data)
			{
				sheet.Cell(row, 1).SetValue(model.Dx);
				sheet.Cell(row, 2).SetValue(model.StartingNumber);
				sheet.Cell(row, 3).SetValue(model.NumberOfVariance);
				sheet.Cell(row, 4).SetValue(model.OverallVariance);
				sheet.Cell(row, 5).SetValue(model.SwVariance);
				sheet.Cell(row, 6).SetValue(model.HwVariance);
				row++;
			}
		}

		private static void GenerateHeader(IXLWorksheet sheet)
		{
			sheet.Cell(1, 1).SetValue("Dx");
			sheet.Cell(1, 2).SetValue("Starting Number");
			sheet.Cell(1, 3).SetValue("Number Of Variance");
			sheet.Cell(1, 4).SetValue("Overall Variance");
			sheet.Cell(1, 5).SetValue("SW Variance");
			sheet.Cell(1, 6).SetValue("HW Variance");
		}
		
		private IEnumerable<AggregatedEcuModel> Calculate()
		{
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
			string sw = model.First().Sw;

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
			string hw = model.First().Hw;

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