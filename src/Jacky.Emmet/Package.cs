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
using EnvDTE80;

namespace Jacky.Emmet
{
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [ProvideAutoLoad(UIContextGuids80.EmptySolution)]
    [ProvideAutoLoad(UIContextGuids80.NoSolution)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasSingleProject)]
    [ProvideAutoLoad(UIContextGuids80.SolutionHasMultipleProjects)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    [Guid(Guids.guidEmmetPkgString)]
    public sealed class Package : Microsoft.VisualStudio.Shell.Package, IDisposable
    {
        private ScriptEngine _engine = null;
        private Console _console = null;
        private Context _context = null;

        public Package()
        {
        }

        public T GetService<T>()
        {
            return (T)GetService(typeof(T));
        }

        #region Package Members
        protected override void Initialize()
        {
            base.Initialize();
            try
            {
                _engine = new ScriptEngine();
                _console = new Console(this, _engine);
                _context = new Context(this, _engine);
                _engine.Bind("console", _console);
                _engine.Bind("context", _context);
                _engine.Exec(_context.Root + "\\Assets\\startup.js");
            }
            catch
            {

            }
        }
        #endregion

        public void Dispose()
        {
            _engine.Dispose();
        }
        /// <summary>
        /// This function is the callback used to execute a command when the a menu item is clicked.
        /// See the Initialize method to see how the menu item is associated to this function using
        /// the OleMenuCommandService service and the MenuCommand class.
        /// </summary>

    }
}
