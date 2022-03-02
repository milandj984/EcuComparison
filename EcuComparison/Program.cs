using System.Collections.Generic;
using EcuComparison.Domain;
using EcuComparison.Models;

namespace EcuComparison
{
	class Program
	{
		static void Main(string[] args)
		{
			const string filePath = @"C:\Milan\Work\C#\Projects\EcuComparison\test.xlsx";
			DataCollector dataCollector = new DataCollector(filePath);
			List<List<EcuModel>> excelData = dataCollector.Collect();
		}
	}
}