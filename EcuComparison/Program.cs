using System;
using System.Collections.Generic;
using ClosedXML.Excel;
using EcuComparison.Domain;
using EcuComparison.Models;

namespace EcuComparison
{
	class Program
	{
		static void Main(string[] args)
		{
			const string filePath = @"C:\Milan\Work\C#\Projects\EcuComparison\test.xlsx";
			Console.WriteLine("Loading excel file...");
			using XLWorkbook workbook = new XLWorkbook(filePath);
			
			// Collecting data from excel sheets
			DataCollector dataCollector = new DataCollector(workbook);
			List<List<EcuModel>> excelSheetData = dataCollector.Collect();

			SingleStatisticGenerator generator = new SingleStatisticGenerator(workbook);
			
			// Statistic for single sheet
			foreach (List<EcuModel> ecuModels in excelSheetData)
			{
				generator.SheetData = ecuModels;
				generator.GenerateAsync();
			}
			
			// Save changes in workbook
			Console.WriteLine("Saving...");
			workbook.Save();
		}
	}
}