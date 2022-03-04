using EcuComparison.Constants;

namespace EcuComparison.Models
{
	public class AggregatedEcuModel
	{
		public string Dx { get; init; }

		public string SwVariance { get; init; }

		public string HwVariance { get; init; }

		public int NumberOfVariance { get; init; }
		
		public int StartingNumber { get; init; }

		public string OverallVariance => SwVariance == Variance.CopSw && HwVariance == Variance.CopHw ? Variance.NoVariance : $"{SwVariance}, {HwVariance}";
	}
}