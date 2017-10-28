using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;
using System.Windows.Forms;

namespace k90_macro_maker
{
    abstract class MacroItem
    {
        public abstract string GetHexCode();
    }

    class MacroDelay : MacroItem
    {
        public MacroDelay(int ms)
        {
            Milliseconds = ms;
        }

        public int Milliseconds { get; set; }

        public override string GetHexCode()
        {
            return string.Format("ff{0:x4}", Milliseconds);
        }
    }

    class MacroKeyDown : MacroItem
    {
        public MacroKeyDown(Keys key)
        {
            Key = key;
        }

        public Keys Key { get; set; }

        public override string GetHexCode()
        {
            return string.Format("{0:x2}0001", (int)Key);
        }
    }

    class MacroKeyUp : MacroItem
    {
        public MacroKeyUp(Keys key)
        {
            Key = key;
        }

        public Keys Key { get; set; }

        public override string GetHexCode()
        {
            return string.Format("{0:x2}0000", (int)Key);
        }
    }

    class MacroKeyPress : MacroItem
    {
        public MacroKeyPress(Keys key, int msDelay)
        {
            Key = key;
            MsDelay = msDelay;
        }

        public Keys Key { get; set; }
        public int MsDelay { get; set; }

        public override string GetHexCode()
        {
            return string.Format("{0:x2}0001", (int)Key)
                 + string.Format("ff{0:x4}", MsDelay)
                 + string.Format("{0:x2}0000", (int)Key);
        }
    }

    class Program
    {
        static void Process(TextReader r, XmlWriter w)
        {
            w.WriteStartElement("GKEYINFO");
            w.WriteElementString("Info", "LaverGKey");
            w.WriteElementString("LoopType", "0");
            w.WriteElementString("ButtonFunction", "48");
            w.WriteElementString("ButtonID", "1");

            string macroName = "NewMacro";
            int delay = 15;

            List<MacroItem> macro = new List<MacroItem>();

            bool haveDelay = true;

            string line;
            while ((line = r.ReadLine()) != null)
            {
                string[] parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length == 0)
                    continue;

                switch (parts[0])
                {
                    case "delay":
                        {
                            int ms;
                            if (parts.Length != 2 || !int.TryParse(parts[1], out ms))
                                ms = delay;
                            haveDelay = true;
                            macro.Add(new MacroDelay(ms));
                            break;
                        }

                    case "defaultdelay":
                    case "default_delay":
                        {
                            int ms;
                            if (parts.Length != 2 || !int.TryParse(parts[0], out ms) || ms < 1 || ms > 999)
                            {
                                Console.Error.WriteLine("ignoring malformed line '{0}'", line);
                                continue;
                            }
                            delay = ms;
                            break;
                        }

                    case "nodelay":
                        {
                            haveDelay = true;
                            break;
                        }

                    case "name":
                        {
                            macroName = line.Substring(parts[0].Length + 1);
                            if (macroName == "")
                                macroName = "NewMacro";
                            break;
                        }

                    case "press":
                    case "keypress":
                        {
                            Keys key;
                            int ms;

                            if (parts.Length <= 1 || parts.Length > 3)
                            {
                                Console.Error.WriteLine("ignoring malformed line '{0}'", line);
                                continue;
                            }

                            key = StringToKey(parts[1]);
                            if (key == Keys.None)
                                continue;

                            if (parts.Length != 3 || !int.TryParse(parts[2], out ms))
                                ms = delay;

                            if (!haveDelay)
                                macro.Add(new MacroDelay(delay));
                            macro.Add(new MacroKeyPress(key, ms));
                            haveDelay = false;
                            break;
                        }

                    case "down":
                    case "keydown":
                        {
                            if (parts.Length != 2)
                            {
                                Console.Error.WriteLine("ignoring malformed line '{0}'", line);
                                continue;
                            }

                            Keys key = StringToKey(parts[1]);
                            if (key == Keys.None)
                                continue;

                            if (!haveDelay)
                                macro.Add(new MacroDelay(delay));
                            macro.Add(new MacroKeyDown(key));
                            haveDelay = false;
                            break;
                        }

                    case "up":
                    case "keyup":
                        {
                            if (parts.Length != 2)
                            {
                                Console.Error.WriteLine("ignoring malformed line '{0}'", line);
                                continue;
                            }

                            Keys key = StringToKey(parts[1]);
                            if (key == Keys.None)
                                continue;

                            if (!haveDelay)
                                macro.Add(new MacroDelay(delay));
                            macro.Add(new MacroKeyUp(key));
                            haveDelay = false;
                            break;
                        }

                    default:
                        {
                            Keys key = StringToKey(parts[0]);
                            if (key == Keys.None)
                                continue;

                            if (!haveDelay)
                                macro.Add(new MacroDelay(delay));
                            macro.Add(new MacroKeyPress(key, delay));
                            haveDelay = false;
                            break;
                        }
                }
            }

            w.WriteElementString("DefaultDelayTime", delay.ToString());
            w.WriteElementString("DelayType", "2");
            w.WriteElementString("FixMacroDelay", delay.ToString());
            w.WriteElementString("LaunchPath", "");
            w.WriteElementString("LoopNumber", "1");
            w.WriteElementString("MacroName", macroName);
            w.WriteElementString("RandomDelayTime", "1000");
            w.WriteElementString("MacroInfo", GetMacroInfo(macro));

            w.WriteEndElement();
        }

        static string GetMacroInfo(List<MacroItem> items)
        {
            const int maxSize = 1360 * 6;
            StringBuilder result = new StringBuilder();

            foreach (var item in items)
            {
                result.Append(item.GetHexCode());
            }
            while (result.Length < maxSize)
                result.Append("000000");
            if (result.Length > maxSize)
                Console.Error.WriteLine("warning: macro maximum size exceeded");
            return result.ToString();
        }

        static Keys StringToKey(string name)
        {
            if (name.Length == 1)
            {
                char ch = name[0];
                if (ch >= 'a' && ch <= 'z')
                    return (Keys)(((int)Keys.A) + (int)ch - (int)'a');

                if (ch >= 'A' && ch <= 'Z')
                    return (Keys)(((int)Keys.A) + (int)ch - (int)'A');

                if (ch >= '0' && ch <= '9')
                    return (Keys)(((int)Keys.D0) + (int)ch - (int)'0');
            }

            switch (name)
            {
                case "`":
                case "~":
                    return Keys.Oemtilde;

                case "shift":
                    return Keys.ShiftKey;
                case "lshift":
                case "leftshift":
                case "left_shift":
                    return Keys.LShiftKey;
                case "rshift":
                case "rightshift":
                case "right_shift":
                    return Keys.RShiftKey;

                case "ctrl":
                case "control":
                    return Keys.ControlKey;
                case "lctrl":
                case "leftctrl":
                case "left_ctrl":
                case "lcontrol":
                case "leftcontrol":
                case "left_control":
                    return Keys.LControlKey;
                case "rctrl":
                case "rightctrl":
                case "right_ctrl":
                case "rcontrol":
                case "rightcontrol":
                case "right_control":
                    return Keys.RControlKey;

                case "alt":
                    return Keys.Menu;
                case "lalt":
                case "leftalt":
                case "left_alt":
                    return Keys.LMenu;
                case "altgr":
                case "ralt":
                case "rightalt":
                case "right_alt":
                    return Keys.RMenu;

                case "tilde":
                    return Keys.Oemtilde;

                case "enter":
                case "return":
                    return Keys.Enter;

                case "esc":
                case "escape":
                    return Keys.Escape;

                case "ins":
                case "insert":
                    return Keys.Insert;

                case "del":
                case "delete":
                    return Keys.Delete;

                case "bkspc":
                case "backspace":
                    return Keys.Back;

                case "home":
                    return Keys.Home;

                case "end":
                    return Keys.End;

                case "left":
                    return Keys.Left;

                case "right":
                    return Keys.Right;

                case "up":
                    return Keys.Up;

                case "down":
                    return Keys.Down;

                case "pgup":
                case "pageup":
                    return Keys.PageUp;

                case "pgdn":
                case "pagedown":
                    return Keys.PageDown;

                case "space":
                case "spacebar":
                case "spc":
                    return Keys.Space;

                case ".":
                    return Keys.OemPeriod;

                case ",":
                    return Keys.Oemcomma;

                default:
                    try
                    {
                        return (Keys)Enum.Parse(typeof(Keys), name);
                    }
                    catch (Exception)
                    {
                        Console.Error.WriteLine("warning: could not parse key name '{0}', ignoring", name);
                        return Keys.None;
                    }
            }
        }

        static void Main(string[] args)
        {
            foreach (string arg in args)
            {
                string outArg = Path.ChangeExtension(arg, ".xml");
                if (outArg.ToLower() == arg.ToLower())
                {
                    Console.Error.WriteLine("skipping argument {0}, already has .xml extension", arg);
                    continue;
                }

                try
                {
                    using (TextReader r = File.OpenText(arg))
                    using (XmlTextWriter w = new XmlTextWriter(outArg, Encoding.UTF8))
                    {
                        w.Formatting = Formatting.Indented;
                        w.Indentation = 0;
                        Process(r, w);
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("error in {0}: {1}", arg, ex.Message);
                }
            }
        }
    }
}