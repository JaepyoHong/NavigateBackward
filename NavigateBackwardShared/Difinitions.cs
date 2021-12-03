using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NavigateBackward
{
	public struct PositionToNavigate
	{
		public string path;
		public int anchorLine;
		public int anchorColumn;
		public int endLine;
		public int endColumn;
	}
}
