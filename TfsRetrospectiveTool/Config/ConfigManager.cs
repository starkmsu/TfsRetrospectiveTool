using System;
using System.IO;
using System.Xml.Serialization;

namespace TfsRetrospectiveTool
{
	internal class ConfigManager
	{
		private const string s_configExtension = "cfg";

		private static Config m_lastConfig;

		internal static Config LoadConfig()
		{
			string userAppDaraPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			string appName = AppDomain.CurrentDomain.FriendlyName;
			while (appName.Contains("."))
			{
				appName = Path.GetFileNameWithoutExtension(appName);
			}
			appName += "." + s_configExtension;
			string path = Path.Combine(userAppDaraPath, appName);

			// search in local folder
			if (!File.Exists(path))
				return new Config();

			using (var fs = new FileStream(path, FileMode.OpenOrCreate, FileAccess.Read, FileShare.Read))
			{
				try
				{
					var configObj = new XmlSerializer(typeof(Config)).Deserialize(fs);
					var config = configObj as Config;
					if (config != null)
						m_lastConfig = config.Copy();
					return config;
				}
				catch (Exception)
				{
					return new Config();
				}
			}
		}

		internal static void SaveConfig(Config config)
		{
			if (config.Equals(m_lastConfig))
				return;

			string userAppDaraPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
			string appName = AppDomain.CurrentDomain.FriendlyName;
			while (appName.Contains("."))
			{
				appName = Path.GetFileNameWithoutExtension(appName);
			}
			appName += "." + s_configExtension;
			string path = Path.Combine(userAppDaraPath, appName);

			using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.Write))
			{
				var serializer = new XmlSerializer(typeof(Config));
				serializer.Serialize(fs, config);
			}
		}
	}
}
