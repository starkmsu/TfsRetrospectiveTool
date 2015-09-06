using System;

namespace TfsRetrospectiveTool
{
	[Serializable]
	public class Config
	{
		public string TfsUrl { get; set; }

		public string AreaPath { get; set; }

		public string Iteration { get; set; }

		public Config Copy()
		{
			return new Config
			{
				TfsUrl = TfsUrl,
				AreaPath = AreaPath,
				Iteration = Iteration,
			};
		}

		public bool Equals(Config other)
		{
			if (other == null)
				return false;
			return TfsUrl == other.TfsUrl
				&& AreaPath == other.AreaPath
				&& Iteration == other.Iteration;
		}
	}
}
