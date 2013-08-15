using System;
using System.IO;
using System.Reflection;
using System.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace Jacky.Emmet
{
    /// <summary>
    /// Represents a Windows JScript Engine.
    /// </summary>
    public sealed class ScriptEngine : IDisposable
    {
        [Flags]
        private enum ScriptText
        {
            None = 0,
            DelayExecution = 1,
            IsVisible = 2,
            IsExpression = 32,
            IsPersistent = 64,
            HostManageSource = 128
        }

        [Flags]
        private enum ScriptInfo
        {
            None = 0,
            IUnknown = 1,
            ITypeInfo = 2
        }

        [Flags]
        private enum ScriptItem
        {
            None = 0,
            IsVisible = 2,
            IsSource = 4,
            GlobalMembers = 8,
            IsPersistent = 64,
            CodeOnly = 512,
            NoCode = 1024
        }

        private enum ScriptThreadState
        {
            NotInScript = 0,
            Running = 1
        }

        private enum ScriptState
        {
            Uninitialized = 0,
            Started = 1,
            Connected = 2,
            Disconnected = 3,
            Closed = 4,
            Initialized = 5
        }

        [Guid("BB1A2AE1-A4F9-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScript
        {
            [PreserveSig]
            int SetScriptSite(IActiveScriptSite pass);
            [PreserveSig]
            int GetScriptSite(Guid riid, out IntPtr site);
            [PreserveSig]
            int SetScriptState(ScriptState state);
            [PreserveSig]
            int GetScriptState(out ScriptState scriptState);
            [PreserveSig]
            int Close();
            [PreserveSig]
            int AddNamedItem(string name, ScriptItem flags);
            [PreserveSig]
            int AddTypeLib(Guid typeLib, uint major, uint minor, uint flags);
            [PreserveSig]
            int GetScriptDispatch(string itemName, out IntPtr dispatch);
            [PreserveSig]
            int GetCurrentScriptThreadID(out uint thread);
            [PreserveSig]
            int GetScriptThreadID(uint win32ThreadId, out uint thread);
            [PreserveSig]
            int GetScriptThreadState(uint thread, out ScriptThreadState state);
            [PreserveSig]
            int InterruptScriptThread(uint thread, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo, uint flags);
            [PreserveSig]
            int Clone(out IActiveScript script);
        }

        [Guid("4954E0D0-FBC7-11D1-8410-006008C3FBFC"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScriptProperty
        {
            [PreserveSig]
            int GetProperty(int dwProperty, IntPtr pvarIndex, out object pvarValue);
            [PreserveSig]
            int SetProperty(int dwProperty, IntPtr pvarIndex, ref object pvarValue);
        }

        [Guid("DB01A1E3-A42B-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScriptSite
        {
            [PreserveSig]
            int GetLCID(out int lcid);
            [PreserveSig]
            int GetItemInfo(string name, ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo);
            [PreserveSig]
            int GetDocVersionString(out string version);
            [PreserveSig]
            int OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
            [PreserveSig]
            int OnStateChange(ScriptState scriptState);
            [PreserveSig]
            int OnScriptError(IActiveScriptError scriptError);
            [PreserveSig]
            int OnEnterScript();
            [PreserveSig]
            int OnLeaveScript();
        }

        [Guid("EAE1BA61-A4ED-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScriptError
        {
            [PreserveSig]
            int GetExceptionInfo(out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
            [PreserveSig]
            int GetSourcePosition(out uint sourceContext, out int lineNumber, out int characterPosition);
            [PreserveSig]
            int GetSourceLineText(out string sourceLine);
        }

        [Guid("BB1A2AE2-A4F9-11cf-8F20-00805F2CD064"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScriptParse32
        {
            [PreserveSig]
            int InitNew();
            [PreserveSig]
            int AddScriptlet(string defaultName, string code, string itemName, string subItemName, string eventName, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out string name, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
            [PreserveSig]
            int ParseScriptText(string code, string itemName, IntPtr context, string delimiter, int sourceContextCookie, uint startingLineNumber, ScriptText flags, out object result, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        }

        [Guid("C7EF7658-E1EE-480E-97EA-D52CB4D76D17"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IActiveScriptParse64
        {
            [PreserveSig]
            int InitNew();
            [PreserveSig]
            int AddScriptlet(string defaultName, string code, string itemName, string subItemName, string eventName, string delimiter, IntPtr sourceContextCookie, uint startingLineNumber, ScriptText flags, out string name, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
            [PreserveSig]
            int ParseScriptText(string code, string itemName, IntPtr context, string delimiter, long sourceContextCookie, uint startingLineNumber, ScriptText flags, out object result, out System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo);
        }

        private IActiveScript _engine;
        private IActiveScriptParse32 _parse32;
        private IActiveScriptParse64 _parse64;

        [Serializable]
        internal class ScriptException : Exception
        {
            internal int Line { get; set; }
            internal int Column { get; set; }
            internal int Number { get; set; }
            internal string Desc { get; set; }
            internal string Text { get; set; }
            internal ScriptException(string message)
                : base(message)
            {
            }
        }

        internal class ScriptSite : IActiveScriptSite
        {
            internal ScriptException LastException;
            internal Dictionary<string, object> NamedItems = new Dictionary<string, object>();

            int IActiveScriptSite.GetLCID(out int lcid)
            {
                lcid = Thread.CurrentThread.CurrentCulture.LCID;
                return 0;
            }

            int IActiveScriptSite.GetItemInfo(string name, ScriptInfo returnMask, out IntPtr item, IntPtr typeInfo)
            {
                item = IntPtr.Zero;
                if ((returnMask & ScriptInfo.ITypeInfo) == ScriptInfo.ITypeInfo)
                    return -2147467263;

                object value;
                if (!NamedItems.TryGetValue(name, out value))
                    return unchecked((int)0x8002802B);

                item = Marshal.GetIUnknownForObject(value);
                return 0;
            }

            int IActiveScriptSite.GetDocVersionString(out string version)
            {
                version = null;
                return 0;
            }

            int IActiveScriptSite.OnScriptTerminate(object result, System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo)
            {
                return 0;
            }

            int IActiveScriptSite.OnStateChange(ScriptState scriptState)
            {
                return 0;
            }

            int IActiveScriptSite.OnScriptError(IActiveScriptError scriptError)
            {
                string sourceLine = null;
                try
                {
                    scriptError.GetSourceLineText(out sourceLine);
                }
                catch
                {
                    // happens sometimes... 
                }
                uint sourceContext;
                int lineNumber;
                int characterPosition;
                scriptError.GetSourcePosition(out sourceContext, out lineNumber, out characterPosition);
                lineNumber++;
                characterPosition++;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
                scriptError.GetExceptionInfo(out exceptionInfo);

                string message;
                if (!string.IsNullOrEmpty(sourceLine))
                {
                    message = "Script exception: {1}. Error number {0} (0x{0:X8}): {2} at line {3}, column {4}. Source line: '{5}'.";
                }
                else
                {
                    message = "Script exception: {1}. Error number {0} (0x{0:X8}): {2} at line {3}, column {4}.";
                }
                LastException = new ScriptException(string.Format(message, exceptionInfo.scode, exceptionInfo.bstrSource, exceptionInfo.bstrDescription, lineNumber, characterPosition, sourceLine));
                LastException.Line = lineNumber;
                LastException.Column = characterPosition;
                LastException.Number = exceptionInfo.scode;
                LastException.Text = sourceLine;
                LastException.Desc = exceptionInfo.bstrDescription;
                return 0;
            }

            int IActiveScriptSite.OnEnterScript()
            {
                return 0;
            }

            int IActiveScriptSite.OnLeaveScript()
            {
                return 0;
            }
        }

        internal ScriptSite Site;

        /// <summary> 
        /// Initializes a new instance of the <see cref="ScriptEngine"/> class. 
        /// </summary>
        public ScriptEngine()
        {
            try
            {
                _engine = Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("{16d51579-a30b-4c8b-a276-0ff4dc41e755}"), true)) as IActiveScript;
            }
            catch
            {
            	_engine = Activator.CreateInstance(Type.GetTypeFromProgID("javascript", true)) as IActiveScript;
            }

            Site = new ScriptSite();
            _engine.SetScriptSite(Site);

            // support 32-bit & 64-bit process 
            if (IntPtr.Size == 4)
            {
                _parse32 = (IActiveScriptParse32)_engine;
                _parse32.InitNew();
            }
            else
            {
                _parse64 = (IActiveScriptParse64)_engine;
                _parse64.InitNew();
            }
        }

        /// <summary> 
        /// Adds the name of a root-level item to the scripting engine's name space. 
        /// </summary> 
        /// <param name="name">The name. May not be null.</param> 
        /// <param name="item">The value. It must be a ComVisible object.</param> 
        public void Bind(string name, object item)
        {
            if (name == null)
                throw new ArgumentNullException("name");

            _engine.AddNamedItem(name, ScriptItem.IsVisible | ScriptItem.IsSource);
            Site.NamedItems[name] = item;
        }

        /// <summary> 
        /// Evaluates an expression. 
        /// </summary> 
        /// <param name="code">The code. May not be null.</param> 
        /// <returns>The result of the evaluation.</returns> 
        public Script Eval(string code)
        {
            if (code == null)
                throw new ArgumentNullException("code");

            return Parse(code);
        }

        /// <summary> 
        /// Evaluates a file. 
        /// </summary> 
        /// <param name="file">The file. May not be null.</param> 
        /// <returns>The result of the evaluation.</returns> 
        public Script Exec(string file)
        {
            string code = "";
            using (StreamReader reader = new StreamReader(file))
            {
                code = reader.ReadToEnd();
            }
            return Parse(code);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="code"></param>
        /// <returns></returns>
        public Script Parse(string code)
        {
            _engine.SetScriptState(ScriptState.Connected);
            
            try
            { 
                object result;
                System.Runtime.InteropServices.ComTypes.EXCEPINFO exceptionInfo;
                if (_parse32 != null)
                {
                    _parse32.ParseScriptText(code, null, IntPtr.Zero, null, 0, 0, ScriptText.None, out result, out exceptionInfo);
                }
                else
                {
                    _parse64.ParseScriptText(code, null, IntPtr.Zero, null, 0, 0, ScriptText.None, out result, out exceptionInfo);
                }
            }
            catch
            {
                if (Site.LastException != null)
                    throw Site.LastException;

                throw;
            }

            if (Site.LastException != null)
                throw Site.LastException;

            IntPtr dispatch;
            _engine.GetScriptDispatch(null, out dispatch);
            Script script = new Script(this, dispatch);
            return script;
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_parse32 != null)
            {
                Marshal.ReleaseComObject(_parse32);
                _parse32 = null;
            }

            if (_parse64 != null)
            {
                Marshal.ReleaseComObject(_parse64);
                _parse64 = null;
            }

            if (_engine != null)
            {
                Marshal.ReleaseComObject(_engine);
                _engine = null;
            }
        }
    }

    /// <summary>
    /// Defines a script object that can be evaluated at runtime.
    /// </summary>
    public sealed class Script : IDisposable
    {
        private object _dispatch;
        private readonly ScriptEngine _engine;

        internal Script(ScriptEngine engine, IntPtr dispatch)
        {
            _engine = engine;
            _dispatch = Marshal.GetObjectForIUnknown(dispatch);
        }

        /// <summary>
        /// Calls a method.
        /// </summary>
        /// <param name="method">The method name. May not be null.</param>
        /// <param name="arguments">The optional arguments.</param>
        /// <returns>The call result.</returns>
        public object Call(string method, params object[] arguments)
        {
            if (_dispatch == null)
                throw new InvalidOperationException();

            if (method == null)
                throw new ArgumentNullException("method");

            try
            {
                return _dispatch.GetType().InvokeMember(method, BindingFlags.InvokeMethod, null, _dispatch, arguments);
            }
            catch
            {
                if (_engine.Site.LastException != null)
                    throw _engine.Site.LastException;

                throw;
            }
        }

        public void Dispose()
        {
            if (_dispatch != null)
            {
                Marshal.ReleaseComObject(_dispatch);
                _dispatch = null;
            }
        }
    }
}
