using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Task = System.Threading.Tasks.Task;
using EnvDTE;
using Microsoft.VisualStudio.Threading;
using System.Threading.Tasks;

namespace NavigateBackward
{
	public class NavigateBackwardService
	{
		private readonly AsyncPackage package;
		private readonly List<PositionToNavigate> historyToNavigate = new List<PositionToNavigate>();

		public NavigateBackwardService(AsyncPackage package)
		{
			this.package = package ?? throw new ArgumentNullException(nameof(package));
		}

		public void GoToDefinitionCommandFired(object sender, EventArgs e)
		{
			JoinableTask joinableTask = package.JoinableTaskFactory.RunAsync(async delegate
			{
				PositionToNavigate positionBefore;
				positionBefore = await GetCurrentPositionAsync();

				PostBuiltInCommand((uint)VSConstants.VSStd97CmdID.GotoDefn);
				await Task.Delay(200);

				PositionToNavigate positionAfter;
				positionAfter = await GetCurrentPositionAsync();

				if (positionBefore.path != positionAfter.path ||
					positionBefore.anchorLine != positionAfter.anchorLine || positionBefore.anchorColumn != positionAfter.anchorColumn ||
					positionBefore.endLine != positionAfter.endLine || positionBefore.endColumn != positionAfter.endColumn)
				{
					historyToNavigate.Add(positionBefore);
				}
			});
		}

		public void NavigateBackwardCommandFired(object sender, EventArgs e)
		{
			ThreadHelper.ThrowIfNotOnUIThread();

			if (historyToNavigate.Count == 0)
				return;

			PositionToNavigate position = historyToNavigate.Last();
			historyToNavigate.RemoveAt(historyToNavigate.Count - 1);

			try
			{
				ServiceProvider sp = ServiceProvider.GlobalProvider;
				object service = sp.GetService(typeof(DTE));
				if (service != null)
				{
					DTE dte = service as DTE;
					if (dte.get_IsOpenFile(EnvDTE.Constants.vsViewKindCode, position.path))     
					{
						//Document document = dte.Documents.Item(position.path);
						//document.Activate();

						//foreach (Document document in dte.Documents)
						//{
						//	if (document.FullName == position.path)
						//	{
						//		document.Activate();
						//		break;
						//	}
						//}	
						Window win = dte.OpenFile(EnvDTE.Constants.vsViewKindCode, position.path);
						win.Visible = true;
						win.SetFocus();
					}
					else
					{
						Window win = dte.OpenFile(EnvDTE.Constants.vsViewKindCode, position.path);
						win.Visible = true;
						win.SetFocus();
					}

					service = sp.GetService(typeof(SVsTextManager));
					if (service != null)
					{
						IVsTextManager2 textManager = service as IVsTextManager2;
						if (textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out IVsTextView view) == VSConstants.S_OK)
						{
							view.SetSelection(position.anchorLine, position.anchorColumn, position.endLine, position.endColumn);
						}
					}
				}
			}
			catch (Exception ex)
			{
				VsShellUtilities.ShowMessageBox(
					package,
					string.Format(CultureInfo.CurrentCulture, "Inside {0}.NavigateBackwardCommandFired(), {1}", this.GetType().FullName, ex.Message),
					"Exception",
					OLEMSGICON.OLEMSGICON_WARNING,
					OLEMSGBUTTON.OLEMSGBUTTON_OK,
					OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
			}
		}

		private async Task<PositionToNavigate> GetCurrentPositionAsync()
		{
			var p = new PositionToNavigate();

			await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

			// get position to navigate backward
			if (await package.GetServiceAsync(typeof(DTE)) is DTE dte)
			{
				p.path = dte.ActiveDocument.FullName;

				if (await package.GetServiceAsync(typeof(SVsTextManager)) is IVsTextManager2 textManager)
				{
					if (textManager.GetActiveView2(1, null, (uint)_VIEWFRAMETYPE.vftCodeWindow, out IVsTextView view) == VSConstants.S_OK)
					{
						view.GetSelection(out p.anchorLine, out p.anchorColumn, out p.endLine, out p.endColumn);
					}
				}
			}

			return p;
		}

		private void PostBuiltInCommand(uint commandId)
		{
			JoinableTask joinableTask = package.JoinableTaskFactory.RunAsync(async delegate
			{
				await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

				if (await package.GetServiceAsync(typeof(SVsUIShell)) is IVsUIShell shell)
				{
					// execute native GoToDifinition command
					object obj = null;
					ErrorHandler.ThrowOnFailure(shell.PostExecCommand(
						VSConstants.GUID_VSStandardCommandSet97, commandId,
						(uint)Microsoft.VisualStudio.OLE.Interop.OLECMDEXECOPT.OLECMDEXECOPT_DONTPROMPTUSER, ref obj));
				}
			});
		}
	}
}
