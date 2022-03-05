using System;
using System.Collections.Generic;
using EcuComparison.Constants;
using EcuComparison.Models;
using OfficeOpenXml;

namespace EcuComparison.Domain
{
	public class DataCollector
	{
		private readonly ExcelWorkbook _workbook;

		private readonly string[] _sheetNames = { Sheets.AeroBEu, Sheets.Id4Eu, Sheets.IdBuzzEu, Sheets.Id4Nar, Sheets.AeroBCn };

		public DataCollector(ExcelWorkbook workbook)
		{
			_workbook = workbook;
		}

		public List<List<EcuModel>> Collect()
		{
			List<List<EcuModel>> result = new List<List<EcuModel>>();

			foreach (string sheetName in _sheetNames)
			{
				Console.WriteLine($"Collecting data for sheet: {sheetName}");
				ExcelWorksheet sheet = _workbook.Worksheets[sheetName];

				if (sheet is not null)
				{
					List<EcuModel> rows = new List<EcuModel>();
					int totalRows = sheet.Dimension.Rows;

					for (int i = 2; i <= totalRows; i++)
					{
						EcuModel mappedRow = MapRowToModel(sheet, i);

						if (!string.IsNullOrEmpty(mappedRow.Dx))
						{
							rows.Add(mappedRow);
						}
					}
					
					result.Add(rows);
				}
				else
				{
					Console.WriteLine($"Couldn't find sheet with name: {sheetName}");
				}
			}

			return result;
		}

		private static EcuModel MapRowToModel(ExcelWorksheet sheet, int rowIndex)
		{
			EcuModel model = new EcuModel()
			{
				SheetName = sheet.Name,
				Project = sheet.Cells[rowIndex, 1].Text,
				Region = sheet.Cells[rowIndex, 2].Text,
				Dx = sheet.Cells[rowIndex, 3].Text,
				Grundsteurgerat = sheet.Cells[rowIndex, 4].Text,
				SgTnrHwTnr = sheet.Cells[rowIndex, 5].Text,
				Comment = sheet.Cells[rowIndex, 6].Text,
				Sw = sheet.Cells[rowIndex, 7].Text,
				Hw = sheet.Cells[rowIndex, 8].Text,
				SwTermin = sheet.Cells[rowIndex, 9].Text
			};

			return model;
		}
	}
}