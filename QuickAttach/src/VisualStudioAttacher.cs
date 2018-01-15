using EnvDTE;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Security;
using System.Windows.Forms;
using DTEProcess = EnvDTE.Process;
using Process = System.Diagnostics.Process;

// Icon made by http://www.famfamfam.com/lab/icons/silk/
namespace QuickAttach.src
{
    public static class VisualStudioAttacher
    {
        [DllImport("ole32.dll")]
        public static extern int CreateBindCtx(int reserved, out IBindCtx ppbc);

        [DllImport("ole32.dll")]
        public static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable prot);

        public static void AttachVisualStudioToProcess(_DTE instanceSolution, string ProcessName)
        {
            Process[] prss = Process.GetProcessesByName(ProcessName);

            if (prss.Length <= 0)
            {
                MessageBox.Show("Failed to find w3wp.exe process! Make sure you opened the page first!", "Quick Attach", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            DTEProcess processToAttachTo = instanceSolution.Debugger.LocalProcesses.Cast<DTEProcess>().FirstOrDefault(process => process.ProcessID == prss[0].Id);

            try
            {

                if (processToAttachTo != null)
                    processToAttachTo.Attach();
                else
                {
                    MessageBox.Show("Failed to attach process to solution!", "Quick Attach", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

            }
            catch (SecurityException)
            {
                MessageBox.Show("Admin Permissions are required to attach the process", "Quick Attach", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }

        private static bool TryGetVsInstance(int processId, out _DTE instance)
        {
            IntPtr numFetched = IntPtr.Zero;
            IRunningObjectTable runningObjectTable;
            IEnumMoniker monikerEnumerator;
            IMoniker[] monikers = new IMoniker[1];

            GetRunningObjectTable(0, out runningObjectTable);
            runningObjectTable.EnumRunning(out monikerEnumerator);
            monikerEnumerator.Reset();

            while (monikerEnumerator.Next(1, monikers, numFetched) == 0)
            {
                IBindCtx ctx;
                CreateBindCtx(0, out ctx);

                string runningObjectName;
                monikers[0].GetDisplayName(ctx, null, out runningObjectName);

                object runningObjectVal;
                runningObjectTable.GetObject(monikers[0], out runningObjectVal);

                if (runningObjectVal is _DTE && runningObjectName.StartsWith("!VisualStudio"))
                {
                    int currentProcessId = int.Parse(runningObjectName.Split(':')[1]);

                    if (currentProcessId == processId)
                    {
                        instance = (_DTE)runningObjectVal;
                        return true;
                    }
                }
            }

            instance = null;
            return false;
        }
    }
}