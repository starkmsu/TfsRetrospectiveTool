using System;
using System.Collections.Generic;
using System.Linq;

namespace TfsRetrospectiveTool
{
	[Serializable]
	public class Config
	{
		public string TfsUrl { get; set; }

		public string AreaPath { get; set; }

		public List<string> AllAreaPaths { get; set; }

		public string Iteration { get; set; }

		public Config()
		{
			AllAreaPaths = new List<string>();
		}

		public Config Copy()
		{
			return new Config
			{
				TfsUrl = TfsUrl,
				AreaPath = AreaPath,
				Iteration = Iteration,
				AllAreaPaths = AllAreaPaths,
			};
		}

		public bool Equals(Config other)
		{
			if (other == null)
				return false;
			return TfsUrl == other.TfsUrl
				&& AreaPath == other.AreaPath
				&& Iteration == other.Iteration
				&& AllAreaPaths.Count == other.AllAreaPaths.Count
				&& AllAreaPaths.All(a => other.AllAreaPaths.Any(o => a == o))
				&& other.AllAreaPaths.All(o => AllAreaPaths.Any(a => a == o));
		}
	}
}
