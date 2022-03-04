using System;
using System.Collections.Generic;
using System.IO;
using EcuComparison.Domain;
using EcuComparison.Models;
using OfficeOpenXml;

namespace EcuComparison
{
	class Program
	{
		static void Main(string[] args)
		{
			const string filePath = @"C:\Milan\Work\C#\Projects\EcuComparison\test.xlsx";
			Console.WriteLine("Loading excel file...");
			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			FileInfo fileInfo = new FileInfo(filePath);
			using ExcelPackage package = new ExcelPackage(fileInfo);
			using ExcelWorkbook workbook = package.Workbook;

			// Collecting data from excel sheets
			DataCollector dataCollector = new DataCollector(workbook);
			List<List<EcuModel>> excelSheetData = dataCollector.Collect();

			SingleStatisticGenerator generator = new SingleStatisticGenerator(workbook);
			
			// Writing statistic for single sheet
			foreach (List<EcuModel> sheetData in excelSheetData)
			{
				generator.SheetData = sheetData;
				generator.GenerateAsync();
			}
			
			// Save changes in workbook
			Console.WriteLine("Saving...");
			package.Save();
		}
	}
}