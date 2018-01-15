using System;
using System.ComponentModel.Design;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace QuickAttach.src
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(GUIDList.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(UIContextGuids80.SolutionExists)]
    public sealed class QAttachPackage : Package, IVsSolutionEvents3, IVsDebuggerEvents
    {
        private DTE2 _applicationObject;

        private IVsSolution solution = null;
        private IVsDebugger debugger = null;
        private uint _hSolutionEvents = uint.MaxValue;
        private uint _hDebuggerEvents = uint.MaxValue;

        private MenuCommand dynamicCmdMenuCommand;

        public QAttachPackage(){ }

        #region Package Members

        protected override void Initialize()
        {
            OleMenuCommandService commandService = GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(GUIDList.CommandSet, 0x0100);
                this.dynamicCmdMenuCommand = new MenuCommand(this.StartQuick, menuCommandID);
                commandService.AddCommand(dynamicCmdMenuCommand);
            }

            AdviseSolutionEvents();
            base.Initialize();

        }

        protected override void Dispose(bool disposing)
        {
            UnadviseSolutionEvents();
            base.Dispose(disposing);
        }

        private void AdviseSolutionEvents()
        {
            UnadviseSolutionEvents();

            solution = GetService(typeof(SVsSolution)) as IVsSolution;
            debugger = GetService(typeof(SVsShellDebugger)) as IVsDebugger;

            if (solution != null)
            {
                solution.AdviseSolutionEvents(this, out _hSolutionEvents);
                debugger.AdviseDebuggerEvents(this, out _hDebuggerEvents);
            }
        }

        private void UnadviseSolutionEvents()
        {
            if (solution != null)
            {
                if (_hSolutionEvents != uint.MaxValue)
                {
                    solution.UnadviseSolutionEvents(_hSolutionEvents);
                    _hSolutionEvents = uint.MaxValue;
                }
                solution = null;
            }

            if (debugger != null)
            {
                if (_hDebuggerEvents != uint.MaxValue)
                {
                    debugger.UnadviseDebuggerEvents(_hDebuggerEvents);
                    _hDebuggerEvents = uint.MaxValue;
                }
                debugger = null;
            }
        }

        public int OnAfterOpenSolution(object pUnkReserved, int fNewSolution)
        {
            dynamicCmdMenuCommand.Enabled = true;
            return VSConstants.S_OK;
        }

        public int OnAfterCloseSolution(object pUnkReserved) { return VSConstants.S_OK; }
        public int OnAfterClosingChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnAfterLoadProject(IVsHierarchy pStubHierarchy, IVsHierarchy pRealHierarchy) { return VSConstants.S_OK; }
        public int OnAfterMergeSolution(object pUnkReserved) { return VSConstants.S_OK; }
        public int OnAfterOpenProject(IVsHierarchy pHierarchy, int fAdded) { return VSConstants.S_OK; }
        public int OnAfterOpeningChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeCloseProject(IVsHierarchy pHierarchy, int fRemoved) { return VSConstants.S_OK; }
        public int OnBeforeCloseSolution(object pUnkReserved) { return VSConstants.S_OK; }
        public int OnBeforeClosingChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeOpeningChildren(IVsHierarchy pHierarchy) { return VSConstants.S_OK; }
        public int OnBeforeUnloadProject(IVsHierarchy pRealHierarchy, IVsHierarchy pStubHierarchy) { return VSConstants.S_OK; }
        public int OnQueryCloseProject(IVsHierarchy pHierarchy, int fRemoving, ref int pfCancel) { return VSConstants.S_OK; }
        public int OnQueryCloseSolution(object pUnkReserved, ref int pfCancel) { return VSConstants.S_OK; }
        public int OnQueryUnloadProject(IVsHierarchy pRealHierarchy, ref int pfCancel) { return VSConstants.S_OK; }

        public int OnModeChange(DBGMODE dbgmodeNew)
        {
            switch (dbgmodeNew)
            {
                case DBGMODE.DBGMODE_Run:
                    dynamicCmdMenuCommand.Enabled = false;
                    break;
                case DBGMODE.DBGMODE_Design:
                    dynamicCmdMenuCommand.Enabled = true;
                    break;
            }

            return VSConstants.S_OK;

        }
        #endregion

        private void StartQuick(object sender, EventArgs e)
        {
            _applicationObject = (DTE2)GetService(typeof(DTE));
            EnvDTE.ProjectItem oProjectItem = _applicationObject.SelectedItems.Item(1).ProjectItem;

            _DTE solution = _applicationObject.Solution.SolutionBuild.DTE;
            VisualStudioAttacher.AttachVisualStudioToProcess(solution, "w3wp");
        }
    }
}
