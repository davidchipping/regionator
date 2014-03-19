using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using Microsoft.Win32;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using EnvDTE;

namespace RegionatorExtension
{

    static class GuidList
    {
        public const string guidRegionatorExtensionPkgString = "4e1470ea-f78f-4ea7-b793-bc940b4f9805";
        public const string guidRegionatorExtensionCmdSetString = "b0ccedc1-2776-4981-baa9-ce145ef22c2a";

        public static readonly Guid guidRegionatorExtensionCmdSet = new Guid(guidRegionatorExtensionCmdSetString);
    };
    static class PkgCmdIDList
    {
        public const uint RegionatorCommandId = 0x100;
    };

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GuidList.guidRegionatorExtensionPkgString)]
    public sealed class RegionatorExtensionPackage : Package
    {

        protected override void Initialize()
        {
            base.Initialize();

            OleMenuCommandService mcs = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if ( null != mcs )
            {
                CommandID menuCommandID = new CommandID(GuidList.guidRegionatorExtensionCmdSet, (int)PkgCmdIDList.RegionatorCommandId);
                MenuCommand menuItem = new MenuCommand(MenuItemCallback, menuCommandID );
                mcs.AddCommand( menuItem );
            }
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            string errorUnableToOpenDocument = "Error: unable to open the current document.";
            DTE dte = GetService(typeof(SDTE)) as DTE;
            if (dte == null)
            {
                ShowFeedBack(errorUnableToOpenDocument);
                return;
            }
            if (dte.ActiveDocument == null)
            {
                ShowFeedBack(errorUnableToOpenDocument);
                return;
            }
            TextDocument currentDocument = dte.ActiveDocument.Object("TextDocument") as TextDocument;
            if (currentDocument == null)
            {
                ShowFeedBack(errorUnableToOpenDocument);
                return;
            }

            dte.UndoContext.Open("Removed Regions");
            try
            {

                //Replace regions in c#
                currentDocument.ReplacePattern(@"^[\s]*#region(.*)(\r\n?|\n)", string.Empty, vsFindOptionsValue: (int)vsFindOptions.vsFindOptionsRegularExpression);
                currentDocument.ReplacePattern(@"^[\s]*#endregion(.*)(\r\n?|\n)", string.Empty, vsFindOptionsValue: (int)vsFindOptions.vsFindOptionsRegularExpression);

                //Replace regions in vb.net
                currentDocument.ReplacePattern(@"^[\s]*#Region(.*)(\r\n?|\n)", string.Empty, vsFindOptionsValue: (int)vsFindOptions.vsFindOptionsRegularExpression);
                currentDocument.ReplacePattern(@"^[\s]*#End Region(.*)(\r\n?|\n)", string.Empty, vsFindOptionsValue: (int)vsFindOptions.vsFindOptionsRegularExpression);
            }
            finally
            {
                dte.UndoContext.Close();
            }
        }

        private void ShowFeedBack(string message)
        {
            IVsUIShell uiShell = (IVsUIShell)GetService(typeof(SVsUIShell));
            Guid clsid = Guid.Empty;
            int result;
            Microsoft.VisualStudio.ErrorHandler.ThrowOnFailure(uiShell.ShowMessageBox(
                       0, ref clsid, "Regionator", string.Format(CultureInfo.CurrentCulture, message, this.ToString()),
                       string.Empty, 0, OLEMSGBUTTON.OLEMSGBUTTON_OK, OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST,
                       OLEMSGICON.OLEMSGICON_INFO, 0, out result));
        }

    }
}
