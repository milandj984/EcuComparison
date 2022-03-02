using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using EcuComparison.Models;

namespace EcuComparison.Domain
{
	public class DataCollector
	{
		private readonly string _filePath;
		
		private static string[] _sheetNames = { "te.VBV_AERO B EU", "te.VBV_ID.4 EU", "te.VBV_ID.Buzz EU", "te.VBV_ID.4 NAR", "te.VBV_AERO B CN" };

		public DataCollector(string filePath)
		{
			_filePath = filePath;
		}

		public List<List<EcuModel>> Collect()
		{
			List<List<EcuModel>> result = new List<List<EcuModel>>();
			using XLWorkbook workbook = new XLWorkbook(_filePath);

			foreach (string sheetName in _sheetNames)
			{
				Console.WriteLine($"Collecting data for sheet: {sheetName}");
				List<EcuModel> data = new List<EcuModel>();
				
				if (workbook.TryGetWorksheet(sheetName, out IXLWorksheet sheet))
				{
					IXLRows rows = sheet.RowsUsed();
					data.AddRange(rows.Skip(1).Select(MapRowToModel));
				}
				else
				{
					Console.WriteLine($"Couldn't find sheet with name: {sheetName}");
					continue;
				}
				
				result.Add(data);
			}

			return result;
		}

		private static EcuModel MapRowToModel(IXLRow row)
		{
			List<IXLCell> cells = row.Cells().ToList();

			EcuModel model = new EcuModel()
			{
				Project = cells[0].GetString(),
				Region = cells[1].GetString(),
				Dx = cells[2].GetString(),
				Grundsteurgerat = cells[3].GetString(),
				SgTnrHwTnr = cells[4].GetString(),
				Comment = cells[5].GetString(),
				Sw = cells[6].GetString(),
				Hw = cells[7].GetString(),
				SwTermin = cells[8].GetString()
			};

			return model;
		}
	}
}