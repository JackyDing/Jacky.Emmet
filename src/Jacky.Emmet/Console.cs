using System.Diagnostics;
using System.Runtime.InteropServices;
using EnvDTE;
using EnvDTE80;

namespace Jacky.Emmet
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Console
    {
        private DTE _dte = null;
        private Package _pkg;
        private ScriptEngine _engine = null;
        
        public DTE App
        {
            get
            {
                if (_dte == null)
                {
                    _dte = _pkg.GetService<DTE>();
                }
                return _dte;
            }
        }

        public Console(Package package, ScriptEngine engine)
        {
            _pkg = package;
            _engine = engine;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        public void log(string text)
        {
            Debug.Print(text);
        }
    }
}
