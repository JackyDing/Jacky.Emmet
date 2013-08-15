using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Text;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualBasic;
using Microsoft.VisualStudio.CommandBars;
using Microsoft.VisualStudio.Shell;
using System.ComponentModel.Design;

namespace Jacky.Emmet
{
    /// <summary>
    /// </summary>
    [DataContract]
    public class Action
    {
        /// <summary>
        /// </summary>
        [DataMember]
        public string type { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public string name { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public string label { get; set; }
        /// <summary>
        /// </summary>
        [DataMember]
        public Action[] items { get; set; }
    }

    /// <summary>
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Context
    {
        private Package _pkg = null;
        private DTE _dte = null;
        private ScriptEngine _engine = null;
        private string _root = null;
        private Dictionary<int, Action> _actions = new Dictionary<int, Action>();

        internal class Data
        {
            internal String Text { get; set; }
            internal String Line { get; set; }
            internal int Active { get; set; }
            internal int Anchor { get; set; }
            internal int LineBegOffset { get; set; }
            internal int LineEndOffset { get; set; }
        }

        private Data _data = new Data();

        /// <summary>
        /// 
        /// </summary>
        public string Root
        {
            get
            {
                if (_root == null)
                {
                    Uri uri = new Uri(Assembly.GetExecutingAssembly().EscapedCodeBase);
                    _root = System.IO.Path.GetDirectoryName(Uri.UnescapeDataString(uri.AbsolutePath));
                }
                return _root;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Path
        {
            get
            {
                return App.ActiveDocument.FullName;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Text
        {
            get
            {
                return _data.Text;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Line
        {
            get
            {
                return _data.Line;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Anchor
        {
            get
            {
                return _data.Anchor;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int Active
        {
            get
            {
                return _data.Active;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LineBegOffset
        {
            get
            {
                return _data.LineBegOffset;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public int LineEndOffset
        {
            get
            {
                return _data.LineEndOffset;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Syntax
        {
            get
            {
                return "html";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Profile
        {
            get
            {
                return "html";
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public string Selection
        {
            get
            {
                TextSelection selection = App.ActiveDocument.Selection as TextSelection;
                return selection.Text;
            }
            set
            {
                TextSelection selection = App.ActiveDocument.Selection as TextSelection;
                selection.Text = value;
            }
        }

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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="engine"></param>
        /// <returns></returns>
        public Context(Package package, ScriptEngine engine)
        {
            _pkg = package;
            _engine = engine;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="beg"></param>
        /// <param name="end"></param>
        /// <returns></returns>
        public bool select(int beg, int end)
        {
            TextSelection selection = App.ActiveDocument.Selection as TextSelection;
            selection.MoveToAbsoluteOffset(toAbsolute(beg), false);
            if (beg != end)
            {
                selection.MoveToAbsoluteOffset(toAbsolute(end), true);
                return true;
            }
            return false;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="beg"></param>
        /// <param name="end"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public bool replace(int beg, int end, string value)
        {
            TextSelection selection = App.ActiveDocument.Selection as TextSelection;
            VirtualPoint active = selection.ActivePoint;
            EditPoint point0 = active.CreateEditPoint();
            EditPoint point1 = active.CreateEditPoint();
            point0.MoveToAbsoluteOffset(toAbsolute(beg));
            point1.MoveToAbsoluteOffset(toAbsolute(end));
            point0.ReplaceText(point1, value, 1);
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public bool require(string path)
        {
            try
            {
                _engine.Exec(path);
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="json"></param>
        /// <returns></returns>
        public bool startup(string json)
        {
            try
            {
                OleMenuCommandService mcs = _pkg.GetService<IMenuCommandService>() as OleMenuCommandService;
                if (null != mcs)
                {
                    Array luids = Enum.GetValues(typeof(Luids));

                    foreach (object luid in luids)
                    {
                        CommandID menuLuid = new CommandID(Guids.guidEmmetSet, (int)luid);
                        OleMenuCommand menuItem = new OleMenuCommand(MenuItemCallback, menuLuid);
                        menuItem.BeforeQueryStatus += MenuItemQueryStatus;
                        mcs.AddCommand(menuItem);
                    }
                    
                    DataContractJsonSerializer serializer = new DataContractJsonSerializer(typeof(Action[]));
                    MemoryStream stream = new MemoryStream(Encoding.UTF8.GetBytes(json));
                    Action[] actions = serializer.ReadObject(stream) as Action[];

                    int i = 0;
                    foreach (var action in actions)
                    {
                        if (i >= luids.Length)
                            break;
                        if (action.type == "action")
                        {
                            _actions.Add((int)luids.GetValue(i++), action);
                        }
                        else
                        {
                            if (action.items != null)
                            {
                                foreach (var item in action.items)
                                {
                                    _actions.Add((int)luids.GetValue(i++), item);
                                }
                            }
                        }
                    }
                }
            }
            catch
            {
                return false;
            }
            return true;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public string prompt(string title, string value)
        {
            return Interaction.InputBox(title, title, null);
        }

        private int toPosition(int absolute)
        {
            TextSelection selection = App.ActiveDocument.Selection as TextSelection;
            VirtualPoint active = selection.ActivePoint;
            EditPoint point0 = active.CreateEditPoint();
            EditPoint point1 = active.CreateEditPoint();
            point0.StartOfDocument();
            point1.MoveToAbsoluteOffset(absolute);
            string text = point0.GetText(point1);
            return (absolute - 1) + number(text, "\r\n", text.Length - 1);
        }

        private int toAbsolute(int position)
        {
            return (position + 1) - number(Text, "\r\n", position);
        }

        private int number(string text, string omit, int last)
        {
            int num = 0;
            int idx = 0;
            int len = omit.Length;
            while (true)
            {
                int tmp = text.IndexOf(omit, idx, StringComparison.Ordinal);
                if (tmp == -1 || tmp >= last)
                {
                    return num;
                }
                num++;
                idx = tmp + len;
            }
        }

        private void enterCallback()
        {
            TextDocument doc = App.ActiveDocument.Object("TextDocument") as TextDocument;
            TextSelection selection = doc.Selection as TextSelection;
            VirtualPoint active = selection.ActivePoint;
            VirtualPoint anchor = selection.AnchorPoint;
            EditPoint point0 = active.CreateEditPoint();
            EditPoint point1 = active.CreateEditPoint();
            point0.StartOfLine();
            point1.EndOfLine();
            _data.Text = doc.StartPoint.CreateEditPoint().GetText(doc.EndPoint);
            _data.Line = point0.GetText(point1);
            _data.Anchor = toPosition(anchor.AbsoluteCharOffset);
            _data.Active = toPosition(active.AbsoluteCharOffset);
            _data.LineBegOffset = toPosition(point0.AbsoluteCharOffset);
            _data.LineEndOffset = toPosition(point1.AbsoluteCharOffset);
        }

        private void leaveCallback()
        {
            _data.Text = "";
            _data.Line = "";
            _data.Anchor = 1;
            _data.Active = 1;
            _data.LineBegOffset = 1;
            _data.LineEndOffset = 1;
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            OleMenuCommand command = sender as OleMenuCommand;
            Action action = null;
            if (_actions.TryGetValue(command.CommandID.ID, out action))
            {
                enterCallback();
                _engine.Eval("emmet.require('actions').run('" + action.name + "', editor);");
                leaveCallback();
            }
        }

        private void MenuItemQueryStatus(object sender, EventArgs e)
        {
            OleMenuCommand command = sender as OleMenuCommand;
            if (command != null)
            {
                Action action = null;
                if (_actions.TryGetValue(command.CommandID.ID, out action))
                {
                    command.Text = action.label;
                    command.Visible = true;
                    command.Enabled = App.ActiveDocument != null && (action.name == "wrap_with_abbreviation" ? this.Selection != "" : true);
                }
                else
                {
                    command.Visible = false;
                }
            }
        }
    }
}
