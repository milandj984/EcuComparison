namespace EcuComparison.Models
{
	public class AggregatedEcuModel
	{
		public string Dx { get; set; }

		public string SwVariance { get; set; }

		public string HwVariance { get; set; }

		public int NumberOfVariance { get; set; }
		
		public int StartingNumber { get; set; }

		public string OverallVariance => $"{SwVariance}, {HwVariance}";
	}
}