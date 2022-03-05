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

			// Writing statistic for single sheet
			SingleStatisticGenerator generator = new SingleStatisticGenerator(workbook);
			
			foreach (List<EcuModel> sheetData in excelSheetData)
			{
				generator.SheetData = sheetData;
				generator.GenerateAsync();
			}
			
			// Writing statistic for Vs sheet
			DoubleStatisticGenerator generator2 = new DoubleStatisticGenerator(workbook);

			// AERO B EU vs. ID.4 EU
			generator2.SheetData = excelSheetData[0];
			generator2.SheetData2 = excelSheetData[1];
			generator2.GenerateAsync();

			// AERO B EU vs. ID.Buzz EU
			generator2.SheetData = excelSheetData[0];
			generator2.SheetData2 = excelSheetData[2];
			generator2.GenerateAsync();

			// ID.4 EU vs. ID.Buzz EU
			generator2.SheetData = excelSheetData[1];
			generator2.SheetData2 = excelSheetData[2];
			generator2.GenerateAsync();

			// ID.4 EU vs. ID.4 NAR
			generator2.SheetData = excelSheetData[1];
			generator2.SheetData2 = excelSheetData[3];
			generator2.GenerateAsync();

			// AERO B EU vs. AERO B CN
			generator2.SheetData = excelSheetData[0];
			generator2.SheetData2 = excelSheetData[4];
			generator2.GenerateAsync();

			// Save changes in workbook
			Console.WriteLine("Saving...");
			package.Save();
		}
	}
}