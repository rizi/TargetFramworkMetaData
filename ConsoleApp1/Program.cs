using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text.RegularExpressions;

using EnvDTE;

using Process = System.Diagnostics.Process;
using Thread = System.Threading.Thread;

namespace ConsoleApp1
{
    internal class Program
    {
        #region Private Methods

        private static void Main(string[] args)
        {
            try
            {
                DTE dte = GetDte();

                if (dte != null)
                {
                    //give VS some time bevore asking for the projects, otherwise you'll get a "Busy Exception..."
                    Thread.Sleep(1000);

                    Project project = dte.Solution.Projects.Item(1);
                    dynamic targetFrameworkMoniker = project.Properties.Item("TargetFrameworkMoniker").Value;
                    dynamic targetFrameworkId = project.Properties.Item("TargetFramework").Value;

                    Console.WriteLine("TargetFrameworkMoniker: " + targetFrameworkMoniker + "\r\nTargetFrameworkId: " + targetFrameworkId);
                    Console.ReadKey();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }


        [DllImport("ole32.dll")]
        private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);


        /// <summary>
        /// Gets the DTE object from any devenv process.
        /// </summary>
        /// <returns>
        /// Retrieved DTE object or
        /// <see langword="null">
        /// if not found.
        /// </see>
        /// </returns>
        private static DTE GetDte()
        {
            object runningObject = null;

            IBindCtx bindCtx = null;
            IRunningObjectTable rot = null;
            IEnumMoniker enumMonikers = null;

            try
            {
                Process process = Process.GetProcessesByName("devenv").SingleOrDefault();

                Marshal.ThrowExceptionForHR(CreateBindCtx(0, out bindCtx));
                bindCtx.GetRunningObjectTable(out rot);
                rot.EnumRunning(out enumMonikers);

                IMoniker[] moniker = new IMoniker[1];
                IntPtr numberFetched = IntPtr.Zero;
                while (enumMonikers.Next(1, moniker, numberFetched) == 0)
                {
                    IMoniker runningObjectMoniker = moniker[0];

                    string name = null;

                    try
                    {
                        runningObjectMoniker?.GetDisplayName(bindCtx, null, out name);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        // Do nothing, there is something in the ROT that we do not have access to.
                    }

                    Regex monikerRegex = new Regex(@"!VisualStudio.DTE\.\d+\.\d+\:" + process.Id, RegexOptions.IgnoreCase);
                    if (!string.IsNullOrEmpty(name) && monikerRegex.IsMatch(name))
                    {
                        Marshal.ThrowExceptionForHR(rot.GetObject(runningObjectMoniker, out runningObject));
                        break;
                    }
                }
            }
            finally
            {
                if (enumMonikers != null)
                {
                    Marshal.ReleaseComObject(enumMonikers);
                }

                if (rot != null)
                {
                    Marshal.ReleaseComObject(rot);
                }

                if (bindCtx != null)
                {
                    Marshal.ReleaseComObject(bindCtx);
                }
            }

            return runningObject as DTE;
        }

        #endregion
    }
}