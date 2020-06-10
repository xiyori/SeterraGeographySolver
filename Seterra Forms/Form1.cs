using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows;
using System.IO;

namespace Seterra_Forms
{
    public partial class Form1 : Form
    {
        double dpi = DPITools.GetScalingFactor();
        double screenwidth = Screen.PrimaryScreen.Bounds.Width, screenheight = Screen.PrimaryScreen.Bounds.Height;
        Size defaultsize;
        string res = "..//..//res/", mapsfolder;
        GConsole console;
        Mode mode = Mode.Default;
        string loadedmapname = "";
        Rectangle currentbounds, currentkeybounds, currentfindkeybounds;
        List<Bitmap> testkeys;
        List<double[]> clicks;

        enum Mode { Default, Creating, FinishCreating, Solving, FinishSolving }

        public Form1()
        {
            InitializeComponent();
            screenwidth *= dpi;
            screenheight *= dpi;
            EmulateClick.Intitialize((int)screenwidth, (int)screenheight, 10);
            MouseHook.MouseDown += MouseHook_MouseDown;
            DPITools.DpiFix();
            Width = (int)(Width * dpi);
            Height = (int)(Height * dpi);
            defaultsize = Size;
            mapsfolder = res + "maps/";
            console = new GConsole(this, "Seterra Geography Solver [Version 1.00.00]\r\n" +
                "(c) 2018 Foma Shipilov. All rights reserved.", "SGS:\\" + Environment.UserName + ">");
            console.CommandReceived += Console_CommandReceived;
        }

        private void Console_CommandReceived(object sender, CommandReceivedEventArgs e)
        {
            if (mode == Mode.Creating)
            {
                if (e.Command[0].ToLower() == "y") StartMapAdding();
                else CancelMapAdding();
            }
            else if (mode == Mode.FinishCreating)
            {
                if (e.Command[0].ToLower() == "y") FinishMapAdding();
                else ResumeMapAdding();
            }
            else if (mode == Mode.Solving)
            {
                if (e.Command[0].ToLower() == "y") StartMapSolving();
                else CancelMapSolving();
            }
            else if (mode == Mode.FinishSolving)
            {
                console.BlockCommandList = false;
                mode = Mode.Default;
                DefaultMode();
            }
            else if (mode == Mode.Default)
            {
                if (e.Command[0] == "show") DispayMapsList();
                else if (e.Command[0] == "add")
                {
                    if (e.Command.Length < 2 || e.Command[1] == "") console.WriteLine("Can't add a map without a name!");
                    else PrepareMapAdding(e.Command[1]);
                }
                else if (e.Command[0] == "remove")
                {
                    if (e.Command.Length < 2 || e.Command[1] == "") console.WriteLine("Can't remove a map without a name!");
                    else RemoveMap(e.Command[1]);
                }
                else if (e.Command[0] == "rename")
                {
                    if (e.Command.Length < 3 || e.Command[2] == "") console.WriteLine("Can't rename a map without a name!");
                    else RenameMap(e.Command[1], e.Command[2]);
                }
                else if (e.Command[0] == "solve")
                {
                    if (e.Command.Length < 2 || e.Command[1] == "") console.WriteLine("Can't solve a map without a name!");
                    else PrepareMapSolving(e.Command[1]);
                }
                else if (e.Command[0] == "cut")
                {
                    if (e.Command.Length < 6 || e.Command[5] == "") console.WriteLine("Missing parameters!");
                    else CutShrink(e.Command[1], bool.Parse(e.Command[2]), int.Parse(e.Command[3]),
                        bool.Parse(e.Command[4]), int.Parse(e.Command[5]));
                }
                else if (e.Command[0] == "exit")
                {
                    Close();
                    return;
                }
                else console.WriteLine("\'" + e.Command[0] + "\' is not recognized as a command.\r\n");
            }
            if (mode == Mode.Default) console.WriteFolder();
        }
        
        #region SolvingUtils

        private void StartMapSolving()
        {
            console.BlockInput = true;
            AnalizeBrowser();
            MouseHook.InstallHook();
            console.WriteLine("Hook successfully installed. Awaiting input...\r\n\r\n" +
                "Please restart seterra test.\r\n");
        }

        private void CancelMapSolving()
        {
            mode = Mode.Default;
            console.WriteLine("Operation cancelled by user.");
            console.BlockCommandList = false;
            DefaultMode();
        }

        private void PrepareMapSolving(string mapname)
        {
            if (!Directory.GetDirectories(mapsfolder).Contains(mapsfolder + mapname))
            {
                console.WriteLine("Map \"" + mapname + "\" does not exist!");
                return;
            }
            LoadMap(mapname);
            SideMode();
            console.Write("\r\nPlease open seterra map in Internet browser.\r\n" +
                "The program will analize your browser, then it will automatically start solving once you press \'restart\' button.\r\n\r\n" +
                "Are you ready? (y/n) -> ");
            mode = Mode.Solving;
            console.BlockCommandList = true;
        }
        #endregion
        #region AddingUtils

        private void FinishMapAdding()
        {
            for (int i = 0; i < testkeys.Count; i++)
                testkeys[i].Save(mapsfolder + loadedmapname + "/" + clicks[i][0] + " " + clicks[i][1] + " .bmp");
            mode = Mode.Default;
            console.BlockInput = false;
            console.BlockCommandList = false;
            console.WriteLine("Successfully saved \"" + loadedmapname + "\".");
            DefaultMode();
        }

        private void ResumeMapAdding()
        {
            MouseHook.InstallHook();
            mode = Mode.Creating;
            console.BlockInput = true;
            console.WriteLine("\r\nOperation resumed.");
        }

        private void StartMapAdding()
        {
            console.BlockInput = true;
            AnalizeBrowser();
            clicks = new List<double[]>();
            testkeys = new List<Bitmap>();
            MouseHook.InstallHook();
            console.WriteLine("Hook successfully installed. Awaiting input...\r\n\r\n" +
                "Though you are not limited by time you have to complete the test without a single mistake, so please be careful!\r\n" +
                "Once the test is finished simply click on the program's window to save map!\r\n");
        }

        private void CancelMapAdding()
        {
            mode = Mode.Default;
            console.WriteLine("Operation cancelled by user.");
            RemoveMap(loadedmapname);
            loadedmapname = "";
            console.BlockCommandList = false;
            DefaultMode();
        }

        private void PrepareMapAdding(string mapname)
        {
            if (Directory.GetDirectories(mapsfolder).Contains(mapsfolder + mapname))
            {
                console.WriteLine("Map \"" + mapname + "\" already exists!");
                return;
            }
            Directory.CreateDirectory(mapsfolder + mapname);
            loadedmapname = mapname;
            SideMode();
            console.Write("Creating new map \"" + mapname + "\"...\r\n\r\nPlease open seterra map in Internet browser.\r\n" +
                "The program will analize your browser, then it will be ready for education. SGS will remember your clicks to use them in automatic mode.\r\n\r\n" +
                "You should just pass the test correctly.\r\n\r\n" +
                "Are you ready? (y/n) -> ");
            mode = Mode.Creating;
            console.BlockCommandList = true;
        }
        #endregion
        #region WindowUtils

        void SideMode()
        {
            Width = (int)screenwidth / 5;
            Height = (int)(screenheight * 0.7);
            Left = (int)screenwidth - Width;
            Top = ((int)screenheight - Height) / 2;
            TopMost = true;
        }

        void DefaultMode()
        {
            Size = defaultsize;
            Left = ((int)screenwidth - Width) / 2;
            Top = ((int)screenheight - Height) / 2;
            TopMost = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (mode != Mode.Default && mode != Mode.FinishSolving) e.Cancel = true;
        }
        #endregion
        #region IO

        private void LoadMap(string mapname)
        {
            if (loadedmapname == mapname)
            {
                console.WriteLine("Map \"" + mapname + "\" is already loaded.");
                return;
            }
            console.WriteLine("Loading map \"" + mapname + "\"...");
            string[] mapfiles = Directory.GetFiles(mapsfolder + mapname);
            testkeys = new List<Bitmap>();
            clicks = new List<double[]>();
            foreach (string mapfile in mapfiles)
            {
                testkeys.Add(new Bitmap(mapfile));
                string[] click = mapfile.Substring(mapfile.LastIndexOf('\\') + 1).Split(' ');
                clicks.Add(new double[] { double.Parse(click[0]), double.Parse(click[1]) });
            }
            loadedmapname = mapname;
        }

        private void RemoveMap(string mapname)
        {
            if (Directory.GetDirectories(mapsfolder).Contains(mapsfolder + mapname))
            {
                Directory.Delete(mapsfolder + mapname, true);
                console.WriteLine("Successfully removed \"" + mapname + "\".");
            }
            else console.WriteLine("Map \"" + mapname + "\" has already been removed or does not exist!");
        }

        private void RenameMap(string mapname, string newname)
        {
            if (Directory.GetDirectories(mapsfolder).Contains(mapsfolder + mapname))
            {
                if (Directory.GetDirectories(mapsfolder).Contains(mapsfolder + newname))
                    console.WriteLine("Map \"" + newname + "\" already exists!");
                else
                {
                    Directory.Move(mapsfolder + mapname, mapsfolder + newname);
                    console.WriteLine("Successfully renamed \"" + mapname + "\" to \"" + newname + "\".");
                }
            }
            else console.WriteLine("Map \"" + mapname + "\" does not exist!");
        }

        private void DispayMapsList()
        {
            string[] maps = Directory.GetDirectories(mapsfolder);
            if (maps.Length == 0) console.WriteLine("There are no maps to display yet! Try creating a new one!");
            else
            {
                for (int i = 0; i < maps.Length; i++) maps[i] = maps[i].Substring(maps[i].LastIndexOf('/') + 1);
                console.WriteLine(maps);
            }
        }

        private void CutShrink(string mapname, bool cutw, int w, bool cuth, int h)
        {
            LoadMap(mapname);
            List<Bitmap> newkeys = new List<Bitmap>();
            for (int i = 0; i < testkeys.Count; i++)
            {
                Bitmap newkey = new Bitmap(w, h);
                using (Graphics g = Graphics.FromImage(newkey))
                {
                    g.DrawImage(testkeys[i], new Rectangle(0, 0, w, h), new Rectangle(0, 0, cutw ? w : testkeys[i].Width, cuth ? h : testkeys[i].Height), GraphicsUnit.Pixel);
                }
                newkeys.Add(newkey);
            }
            testkeys = newkeys;
            loadedmapname = mapname + "_c";
            Directory.CreateDirectory(mapsfolder + loadedmapname);
            FinishMapAdding();
        }
        #endregion
        #region Analizing

        private void AnalizeBrowser()
        {
            console.WriteLine("\r\nAnalizing browser...");
            currentbounds = GetSeterraBounds();
            currentkeybounds = new Rectangle(0, (int)Math.Round(currentbounds.Width / 48.5 + currentbounds.Y),
                (int)Math.Round(currentbounds.Width / 10.582), (int)Math.Round(currentbounds.Width / 145.5));
            currentfindkeybounds = new Rectangle((int)Math.Round(currentbounds.Width / 9.312 + currentbounds.X),
                (int)Math.Round(currentbounds.Width / 37.548 + currentbounds.Y), (int)Math.Round(currentbounds.Width / 6.4667), 1);
        }

        private Rectangle GetSeterraBounds()
        {
            Rectangle output = new Rectangle();
            Bitmap screen = Screenshot();
            int searchmode;
            Color pixel;
            for (int y = 100; y < screen.Height; y++)
            {
                searchmode = 0;
                for (int x = 0; x < screen.Width; x++)
                {
                    pixel = screen.GetPixel(x, y);
                    if (searchmode == 0)
                    {
                        if (x >= screen.Width / 2) break;
                        if (pixel.B == 255 && pixel.G == 255 && pixel.R == 255) searchmode = 1;
                    }
                    else if (searchmode == 1)
                    {
                        if (x >= screen.Width / 2) break;
                        if (pixel.B != 255 || pixel.G != 255 || pixel.R != 255)
                        {
                            output.X = x;
                            output.Y = y;
                            searchmode = 2;
                        }
                    }
                    else if (searchmode == 2 && pixel.B == 255 && pixel.G == 255 && pixel.R == 255)
                    {
                        output.Width = x - output.X;
                        break;
                    }
                }
                if (searchmode == 2) break;
            }
            int ys;
            for (ys = output.Y; ys < screen.Height; ys++)
            {
                pixel = screen.GetPixel(output.X, ys);
                if (pixel.B == 255 && pixel.G == 255 && pixel.R == 255) break;
            }
            output.Height = ys - output.Y;
            return output;
        }
        #endregion
        #region Solving

        private void SolveMap(MouseEventArgs e)
        {
            MouseHook.UnInstallHook();
            console.WriteLine("Hook disabled. Process started.\r\n");
            EmulateClick.LeftClick(e.X, e.Y);
            System.Threading.Thread.Sleep(500);
            for (int i = 0; i < testkeys.Count * 5; i++)
            {
                int ind;
                try { ind = SearchKey(GetKey()).X; }
                catch { break; }
                EmulateClick.LeftClick((int)Math.Round((double)currentbounds.Width / clicks[ind][0]) + currentbounds.X,
                    (int)Math.Round((double)currentbounds.Width / clicks[ind][1]) + currentbounds.Y); 
            }
            console.Write("Done! Press \'Enter\' to continue...");
            mode = Mode.FinishSolving;
            console.BlockInput = false;
        }
        #endregion
        #region Processing

        private void MouseHook_MouseDown(object sender, MouseEventArgs e)
        {
            if (mode == Mode.Creating) AddClick(e);
            else if (mode == Mode.Solving) SolveMap(e);
        }

        private void AddClick(MouseEventArgs e)
        {
            if (!currentbounds.Contains(e.Location))
            {
                MouseHook.UnInstallHook();
                mode = Mode.FinishCreating;
                console.BlockInput = false;
                console.Write("Hook disabled.\r\n\r\n" +
                    "You've clicked out of seterra bounds. Do you want to finish education and save map? (y/n) -> ");
                return;
            }
            Bitmap newkey = GetKey();
            double[] mpoint = new double[]{ (double)currentbounds.Width / (double)(e.X - currentbounds.X),
                    (double)currentbounds.Width / (double)(e.Y - currentbounds.Y) };
            Point p = SearchKey(newkey);
            if (p.Y == 0)
            {
                clicks[p.X] = mpoint;
                testkeys[p.X] = newkey;
                console.WriteLine("Old value replaced.");
            }
            else
            {
                clicks.Add(mpoint);
                testkeys.Add(newkey);
                console.WriteLine("New value added. Difference = " + p.Y);
            }
        }

        private Bitmap GetKey()
        {
            currentkeybounds.X = GetKeyX();
            Bitmap newkey = new Bitmap(73, 8);
            using (Graphics g = Graphics.FromImage(newkey))
            {
                g.DrawImage(Screenshot(currentkeybounds), new Rectangle(0, 0, 73, 8));
            }
            Color pixel;
            for (int y = 0; y < 8; y++)
                for (int x = 0; x < 73; x++)
                {
                    pixel = newkey.GetPixel(x, y);
                    if (pixel.R + pixel.G + pixel.B >= 220) newkey.SetPixel(x, y, Color.White);
                    else newkey.SetPixel(x, y, Color.Black);
                }
            //newkey.Save("..//..//res/key.bmp");
            return newkey;
        }

        private int GetKeyX()
        {
            Bitmap findkey = Screenshot(currentfindkeybounds);
            int left = 0, right = 0, prev = 0;
            bool fixprev = false;
            Color pixel;
            //findkey.Save("..//..//res/b.bmp");
            for (int i = 0; i < findkey.Width; i++)
            {
                pixel = findkey.GetPixel(i, 0);
                if (pixel.R + pixel.G + pixel.B <= 100)
                {
                    if (!fixprev)
                    {
                        prev = right;
                        left = i;
                        fixprev = true;
                    }
                    right = i;
                }
                else
                {
                    if (i - right == 1) fixprev = false;
                    if (i - right >= currentbounds.Width / 72.75 && left - prev >= currentbounds.Width / 72.75)
                    {
                        /*using (Graphics g = Graphics.FromImage(findkey))
                        {
                            g.DrawLine(new Pen(Brushes.Red, 1), new Point(left, 0), new Point(left, 9));
                            g.DrawLine(new Pen(Brushes.Green, 1), new Point(prev, 0), new Point(prev, 9));
                        }
                        findkey.Save("..//..//res/b.bmp");*/
                        return (int)Math.Round(currentfindkeybounds.X + right + currentbounds.Width / 9.95);
                    }
                }
            }
            throw new InvalidOperationException();
        }
        #endregion
        #region BitmapSearch

        private Point SearchKey(Bitmap key)
        {
            if (testkeys.Count == 0) return new Point(0, -1);
            int mindiff = Subtract(testkeys[0], key), diff, minind = 0;
            for (int i = 1; i < testkeys.Count; i++)
            {
                diff = Subtract(testkeys[i], key);
                if (mindiff > diff)
                {
                    mindiff = diff;
                    minind = i;
                }
            }
            return new Point(minind, mindiff);
        }

        private int Subtract(Bitmap img1, Bitmap img2)
        {
            int diff = 0;
            Color pixel1, pixel2;
            for (int y = 0; y < img1.Height; y++)
            {
                for (int x = 0; x < img1.Width; x++)
                {
                    pixel1 = img1.GetPixel(x, y);
                    pixel2 = img2.GetPixel(x, y);
                    diff += Math.Abs(pixel1.R - pixel2.R) + Math.Abs(pixel1.G - pixel2.G) + Math.Abs(pixel1.B - pixel2.B);
                }
            }
            return diff;
        }
        #endregion
        #region ScreenshotTools

        private Bitmap Screenshot() => Screenshot(0, 0, (int)screenwidth, (int)screenheight);
        private Bitmap Screenshot(Rectangle rect) => Screenshot(rect.X, rect.Y, rect.Width, rect.Height);

        private Bitmap Screenshot(int x, int y, int width, int height)
        {
            Bitmap shot = new Bitmap(width, height);
            Graphics g = Graphics.FromImage(shot);
            g.CopyFromScreen(x, y, 0, 0, new Size(width, height), CopyPixelOperation.SourceCopy);
            return shot;
        }
        #endregion
    }

    public class GConsole : TextBox
    {
        private int caretpos, syspos;
        private string systext, title, folder, lasttext;
        private bool blocktextchangedevent;
        private List<string> commandhistory;
        private int commandindex = 0;
        public event EventHandler<CommandReceivedEventArgs> CommandReceived;

        public GConsole(Form form, string title, string folder)
        {
            BackColor = Color.Black;
            Cursor = Cursors.Default;
            Font = new Font("Consolas", 13F, FontStyle.Regular, GraphicsUnit.Point, 0);
            ForeColor = Color.FromArgb(224, 224, 224);
            Location = new Point(0, 0);
            Multiline = true;
            Name = "consoleBox";
            Size = form.ClientSize;
            ScrollBars = ScrollBars.Vertical;
            ContextMenu = new ContextMenu();
            ShortcutsEnabled = false;
            this.title = title;
            this.folder = folder;
            commandhistory = new List<string>();
            Text = systext = lasttext = title + "\r\n\r\n" + folder;
            SelectionStart = syspos = caretpos = Text.Length;
            MouseUp += GConsole_MouseUp;
            KeyDown += GConsole_KeyDown;
            TextChanged += GConsole_TextChanged;
            form.SizeChanged += Parent_SizeChanged;
            form.Controls.Add(this);
        }

        private void Parent_SizeChanged(object sender, EventArgs e) => Size = ((Form)sender).ClientSize;

        private void GConsole_TextChanged(object sender, EventArgs e)
        {
            if (!blocktextchangedevent)
            {
                if (BlockInput || (SelectionStart < syspos && Text.Length < lasttext.Length))
                {
                    blocktextchangedevent = true;
                    Text = lasttext;
                    ReturnCaret();
                }
                else
                {
                    if (Text.Length - lasttext.Length == 2)
                    {
                        string command = Text.Substring(syspos).Replace("\r\n", "");
                        if (!BlockCommandList)
                        {
                            if (commandhistory.Count == 0 || !command.Equals(commandhistory[commandhistory.Count - 1]))
                                commandhistory.Add(command);
                            commandindex = commandhistory.Count;
                        }
                        blocktextchangedevent = true;
                        if (command.StartsWith("clear"))
                        {
                            Text = systext = lasttext = "\r\n" + folder;
                            syspos = caretpos = Text.Length;
                            ReturnCaret();
                        }
                        else
                        {
                            Text = systext = lasttext = systext + command + "\r\n";
                            blocktextchangedevent = false;
                            syspos = caretpos = Text.Length;
                            ReturnCaret();
                            CommandReceived(command, new CommandReceivedEventArgs
                            {
                                Command = command.Split(' ')
                            });
                        }
                    }
                    else
                    {
                        lasttext = Text;
                        caretpos = SelectionStart;
                    }
                }
            }
            else blocktextchangedevent = false;
        }

        private void GConsole_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down)
            {
                e.Handled = true;
                if (!BlockInput && !BlockCommandList)
                {
                    string command;
                    if (e.KeyCode == Keys.Up)
                    {
                        if (commandindex == 0) return;
                        else command = commandhistory[--commandindex];
                    }
                    else
                    {
                        if (commandindex >= commandhistory.Count - 1) return;
                        else command = commandhistory[++commandindex];
                    }
                    blocktextchangedevent = true;
                    Text = lasttext = systext + command;
                    caretpos = Text.Length;
                    ReturnCaret();
                }
            }
            else if (e.KeyCode == Keys.Left)
            {
                if (SelectionStart <= syspos || e.Modifiers == Keys.Shift) e.Handled = true;
                else caretpos--;
            }
            else if (e.KeyCode == Keys.Right)
            {
                if (e.Modifiers == Keys.Shift) e.Handled = true;
                else caretpos++;
            }
            else if (e.KeyCode == Keys.V && e.Modifiers == Keys.Control) e.Handled = true;
        }

        private void GConsole_MouseUp(object sender, MouseEventArgs e) => ReturnCaret();

        private void ReturnCaret()
        {
            SelectionStart = caretpos;
            SelectionLength = 0;
            ScrollToCaret();
        }

        public void WriteFolder()
        {
            if (Text.EndsWith("\r\n")) Write(folder);
            else Write("\r\n" + folder);
        }

        public void Write(string[] strings)
        {
            string text = strings[0];
            for (int i = 1; i < strings.Length; i++) text += "\r\n" + strings[i];
            Write(text);
        }

        public void WriteLine(string[] strings)
        {
            strings[strings.Length - 1] += "\r\n";
            Write(strings);
        }

        public void Write(string text)
        {
            if (text == "") return;
            blocktextchangedevent = true;
            AppendText(text);
            SelectionStart = syspos = caretpos = Text.Length;
            systext = lasttext = Text;
        }

        public void WriteLine(string text) => Write(text + "\r\n");

        public bool BlockInput { get; set; }
        public bool BlockCommandList { get; set; }
    }

    public class CommandReceivedEventArgs : EventArgs
    {
        public string[] Command { get; set; }
    }

    public static class DPITools
    {
        [DllImport("gdi32.dll")]
        static extern int GetDeviceCaps(IntPtr hdc, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117
        }


        public static float GetScalingFactor()
        {
            Graphics g = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr desktop = g.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(desktop, (int)DeviceCap.DESKTOPVERTRES);

            float ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;

            return ScreenScalingFactor; // 1.25 = 125%
        }

        public static void DpiFix()
        {
            if (Environment.OSVersion.Version.Major >= 6) SetProcessDPIAware();
        }

        [DllImport("user32.dll")]
        private static extern bool SetProcessDPIAware();
    }

    public static class EmulateClick
    {
        static int ScreenWidth, ScreenHeight, SleepTime;

        [DllImport("user32.dll")]
        public static extern void mouse_event(uint dwFlags, int dx, int dy, uint dwData, int dwExtraInfo);

        [Flags]
        public enum MouseEventFlags
        {
            LEFTDOWN = 0x00000002,
            LEFTUP = 0x00000004,
            MIDDLEDOWN = 0x00000020,
            MIDDLEUP = 0x00000040,
            MOVE = 0x00000001,
            ABSOLUTE = 0x00008000,
            RIGHTDOWN = 0x00000008,
            RIGHTUP = 0x00000010
        }

        public static void Intitialize(int screenwidth, int screenheight, int sleeptime)
        {
            ScreenWidth = screenwidth;
            ScreenHeight = screenheight;
            SleepTime = sleeptime;
        }

        public static void LeftClick(int x, int y)
        {
            x += 5;
            y += 6;
            x *= 65536 / ScreenWidth;
            y *= 65536 / ScreenHeight;
            mouse_event((int)(MouseEventFlags.MOVE | MouseEventFlags.ABSOLUTE), x, y, 0, 0);
            //mouse_event((int)(MouseEventFlags.LEFTDOWN), x, y, 0, 0);
            mouse_event((int)(MouseEventFlags.LEFTDOWN), x, y, 0, 0);
            System.Threading.Thread.Sleep(SleepTime);
            mouse_event((int)(MouseEventFlags.LEFTUP), x, y, 0, 0);
        }

        public static void RightClick(int x, int y)
        {
            x *= 65536 / ScreenWidth;
            y *= 65536 / ScreenHeight;
            mouse_event((int)(MouseEventFlags.MOVE | MouseEventFlags.ABSOLUTE), x, y, 0, 0);
            mouse_event((int)(MouseEventFlags.RIGHTDOWN), x, y, 0, 0);
            System.Threading.Thread.Sleep(SleepTime);
            mouse_event((int)(MouseEventFlags.RIGHTUP), x, y, 0, 0);
        }
    }

    public static class MouseHook
    {
        #region Declarations
        public static event MouseEventHandler MouseDown;
        public static event MouseEventHandler MouseUp;
        public static event MouseEventHandler MouseMove;

        [StructLayout(LayoutKind.Sequential)]
        struct MOUSEHOOKSTRUCT
        {
            public POINT pt;
            public IntPtr hwnd;
            public int wHitTestCode;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public int mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator Point(POINT p)
            {
                return new Point(p.X, p.Y);
            }

            public static implicit operator POINT(Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        const int WM_LBUTTONDOWN = 0x201;
        const int WM_LBUTTONUP = 0x202;
        const int WM_MOUSEMOVE = 0x0200;
        const int WM_MOUSEWHEEL = 0x020A;
        const int WM_RBUTTONDOWN = 0x0204;
        const int WM_RBUTTONUP = 0x0205;
        const int WM_MBUTTONUP = 0x208;
        const int WM_MBUTTONDOWN = 0x207;
        const int WM_XBUTTONDOWN = 0x20B;
        const int WM_XBUTTONUP = 0x20C;

        static IntPtr hHook = IntPtr.Zero;
        static IntPtr hModule = IntPtr.Zero;
        static bool hookInstall = false;
        static bool localHook = false;
        static API.HookProc hookDel;
        #endregion

        /// <summary>
        /// Hook install method.
        /// </summary>
        public static void InstallHook()
        {
            if (IsHookInstalled)
                return;

            hModule = Marshal.GetHINSTANCE(AppDomain.CurrentDomain.GetAssemblies()[0].GetModules()[0]);
            hookDel = new API.HookProc(HookProcFunction);

            if (localHook)
                hHook = API.SetWindowsHookEx(API.HookType.WH_MOUSE,
                    hookDel, IntPtr.Zero, AppDomain.GetCurrentThreadId()); // Если подчеркивает не обращай внимание, так надо.
            else
                hHook = API.SetWindowsHookEx(API.HookType.WH_MOUSE_LL,
                    hookDel, hModule, 0);

            if (hHook != IntPtr.Zero)
                hookInstall = true;
            else
                throw new Win32Exception("Can't install low level keyboard hook!");
        }
        /// <summary>
        /// If hook installed return true, either false.
        /// </summary>
        public static bool IsHookInstalled
        {
            get { return hookInstall && hHook != IntPtr.Zero; }
        }
        /// <summary>
        /// Module handle in which hook was installed.
        /// </summary>
        public static IntPtr ModuleHandle
        {
            get { return hModule; }
        }
        /// <summary>
        /// If true local hook will installed, either global.
        /// </summary>
        public static bool LocalHook
        {
            get { return localHook; }
            set
            {
                if (value != localHook)
                {
                    if (IsHookInstalled)
                        throw new Win32Exception("Can't change type of hook than it install!");
                    localHook = value;
                }
            }
        }
        /// <summary>
        /// Uninstall hook method.
        /// </summary>
        public static void UnInstallHook()
        {
            if (IsHookInstalled)
            {
                if (!API.UnhookWindowsHookEx(hHook))
                    throw new Win32Exception("Can't uninstall low level keyboard hook!");
                hHook = IntPtr.Zero;
                hModule = IntPtr.Zero;
                hookInstall = false;
            }
        }
        /// <summary>
        /// Hook process messages.
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        static IntPtr HookProcFunction(int nCode, IntPtr wParam, [In] IntPtr lParam)
        {
            if (nCode == 0)
            {
                if (localHook)
                {
                    MOUSEHOOKSTRUCT mhs = (MOUSEHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MOUSEHOOKSTRUCT));
                    #region switch
                    switch (wParam.ToInt32())
                    {
                        case WM_LBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Left,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_LBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Left,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Middle,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Middle,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MOUSEMOVE:
                            if (MouseMove != null)
                                MouseMove(null,
                                    new MouseEventArgs(MouseButtons.None,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MOUSEWHEEL:
                            // Данный хук не позволяет узнать куда вращается колесо мыши.
                            break;
                        case WM_RBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Right,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_RBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Right,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        default:
                            //Debug.WriteLine(string.Format("X:{0}; Y:{1}; Handle:{2}; HitTest:{3}; EI:{4}; wParam:{5}; lParam:{6}",
                            //    mhs.pt.X, mhs.pt.Y, mhs.hwnd, mhs.wHitTestCode, mhs.dwExtraInfo, wParam.ToString(), lParam.ToString()));
                            break;
                    }
                    #endregion
                }
                else
                {
                    MSLLHOOKSTRUCT mhs = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));

                    #region switch
                    switch (wParam.ToInt32())
                    {
                        case WM_LBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Left,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_LBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Left,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Middle,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Middle,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MOUSEMOVE:
                            if (MouseMove != null)
                                MouseMove(null,
                                    new MouseEventArgs(MouseButtons.None,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_MOUSEWHEEL:
                            if (MouseMove != null)
                                MouseMove(null,
                                    new MouseEventArgs(MouseButtons.None, mhs.time,
                                        mhs.pt.X, mhs.pt.Y, mhs.mouseData >> 16));
                            //Debug.WriteLine(string.Format("X:{0}; Y:{1}; MD:{2}; Time:{3}; EI:{4}; wParam:{5}; lParam:{6}",
                            //            mhs.pt.X, mhs.pt.Y, mhs.mouseData, mhs.time, mhs.dwExtraInfo, wParam.ToString(), lParam.ToString()));
                            break;
                        case WM_RBUTTONDOWN:
                            if (MouseDown != null)
                                MouseDown(null,
                                    new MouseEventArgs(MouseButtons.Right,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        case WM_RBUTTONUP:
                            if (MouseUp != null)
                                MouseUp(null,
                                    new MouseEventArgs(MouseButtons.Right,
                                        1,
                                        mhs.pt.X,
                                        mhs.pt.Y,
                                        0));
                            break;
                        default:

                            break;
                    }
                    #endregion
                }
            }

            return API.CallNextHookEx(hHook, nCode, wParam, lParam);
        }
    }

    static class API
    {
        public delegate IntPtr HookProc(int nCode, IntPtr wParam, [In] IntPtr lParam);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, [In] IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(HookType hookType, HookProc lpfn,
        IntPtr hMod, int dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern IntPtr GetModuleHandle(string lpModuleName);

        public enum HookType : int
        {
            WH_JOURNALRECORD = 0,
            WH_JOURNALPLAYBACK = 1,
            WH_KEYBOARD = 2,
            WH_GETMESSAGE = 3,
            WH_CALLWNDPROC = 4,
            WH_CBT = 5,
            WH_SYSMSGFILTER = 6,
            WH_MOUSE = 7,
            WH_HARDWARE = 8,
            WH_DEBUG = 9,
            WH_SHELL = 10,
            WH_FOREGROUNDIDLE = 11,
            WH_CALLWNDPROCRET = 12,
            WH_KEYBOARD_LL = 13,
            WH_MOUSE_LL = 14
        }
    }
}
