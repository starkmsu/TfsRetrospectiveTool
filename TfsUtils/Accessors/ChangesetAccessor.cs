using Microsoft.TeamFoundation.VersionControl.Client;

namespace TfsUtils.Accessors
{
	public class ChangesetAccessor
	{
		private readonly VersionControlServer m_versionControlServer;

		public ChangesetAccessor(TfsAccessor accessor)
		{
			m_versionControlServer = accessor.GetVersionControlServer();
		}

		public Changeset GetChangesetById(int changesetId)
		{
			return m_versionControlServer.GetChangeset(changesetId, true, false, true);
		}
	}
}
