/*Modify List:
 * 20050725,Color&Font add
 * 
 */
using System;
using System.Drawing;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Forms;
using mshtml;
using System.Runtime.InteropServices;
using System.Diagnostics;

using System.Reflection;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Process
	/// </summary>
	public class Shell
	{
		protected IntPtr shellHandle = IntPtr.Zero;
		protected DataRow shellInfo = null;
		protected const int MILLISECONDS_TIMEOUT = 1000 * 15;
		//base only used handle collect,static for callback
		private static DataTable tblWinTMP = null;
		//base and parent used handle collect
		protected DataTable tblWin = null;
		//base only used for filter process
		private static string sProcessNameFilterTMP = null;
		//base and parent used for filter process
		protected string sProcessNameFilter = null;

		#region Windows Api Define
		[DllImport("oleacc", CharSet=CharSet.Ansi, SetLastError=true, ExactSpelling=true)]
		public static extern int ObjectFromLresult(int lResult, ref Guid riid, int wParam, ref HTMLDocument ppvObject);

		[DllImport("user32", EntryPoint="SendMessageTimeoutW", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
		public static extern int SendMessageTimeout(IntPtr hWnd, int msg, int wParam, int lParam, int fuFlags, int uTimeout, ref int lpdwResult);

		[DllImport("user32.dll")]
		static extern IntPtr GetDesktopWindow();

		[DllImport("user32", EntryPoint="GetWindowTextW", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true,CallingConvention=CallingConvention.Winapi)]
		public static extern void GetWindowText(int h, StringBuilder s, int nMaxCount);

		[DllImport("user32", EntryPoint="GetClassNameW", CharSet=CharSet.Unicode, SetLastError=true, ExactSpelling=true)]
		public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

		[DllImport("user32.dll", SetLastError=true)]
		static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

		[DllImport("user32.dll", EntryPoint="ShowWindow", CharSet=CharSet.Auto)]
		public static extern int ShowWindow(IntPtr hwnd,int nCmdShow);

		// Nested Types
		public delegate int EnumChildProc(IntPtr hWnd, ref IntPtr lParam);
		[DllImport("user32", CharSet=CharSet.Ansi, SetLastError=true, ExactSpelling=true)]
		public static extern int EnumChildWindows(IntPtr hWndParent, EnumChildProc lpEnumFunc, ref IntPtr lParam);

		[DllImport("user32", EntryPoint="RegisterWindowMessageA", CharSet=CharSet.Ansi, SetLastError=true, ExactSpelling=true)]
		public static extern int RegisterWindowMessage([MarshalAs(UnmanagedType.VBByRefStr)] ref string lpString);

		[DllImport("user32")]
		private static extern int ShowWindow(int hwnd, int nCmdShow);

		[DllImport("user32")]
		public static extern int IsWindow(int hwnd);

		[DllImport("user32.dll")]
		private static extern bool SetForegroundWindow(IntPtr hWnd);

		[DllImport("user32.dll")]
		static extern IntPtr SetFocus(IntPtr hWnd);

		[DllImport("user32.dll")]
		static extern uint WaitForInputIdle(IntPtr hProcess, uint dwMilliseconds);

		[DllImport("user32")]
		public static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

		[DllImport("user32")]
		public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
		/*
		Scroll page up
		SendMessage(textBox1.Handle,WM_VSCROLL,(IntPtr)SB_PAGEUP,IntPtr.Zero);

		Scroll page down
		SendMessage(textBox1.Handle,WM_VSCROLL,(IntPtr)SB_PAGEDOWN,IntPtr.Zero);
		*/
		private const int WM_SCROLL = 276; // Horizontal scroll
		private const int WM_VSCROLL = 277; // Vertical scroll
		private const int SB_LINEUP = 0; // Scrolls one line up
		private const int SB_LINELEFT = 0;// Scrolls one cell left
		private const int SB_LINEDOWN = 1; // Scrolls one line down
		private const int SB_LINERIGHT = 1;// Scrolls one cell right
		private const int SB_PAGEUP = 2; // Scrolls one page up
		private const int SB_PAGELEFT = 2;// Scrolls one page left
		private const int SB_PAGEDOWN = 3; // Scrolls one page down
		private const int SB_PAGERIGTH = 3; // Scrolls one page right
		private const int SB_PAGETOP = 6; // Scrolls to the upper left
		private const int SB_LEFT = 6; // Scrolls to the left
		private const int SB_PAGEBOTTOM = 7; // Scrolls to the upper right
		private const int SB_RIGHT = 7; // Scrolls to the right
		private const int SB_ENDSCROLL = 8; // Ends scroll
		private const int WM_COPY = 0x301;
		private const int WM_CUT = 0x300;
		private const int WM_PASTE = 0x302;
		[DllImport("user32.dll",CharSet=CharSet.Auto)]
		private static extern int SendMessage(IntPtr hWnd, int wMsg, IntPtr wParam, IntPtr lParam);
		#endregion

		// Win32 Constants
		// Constants from WinUser.h
		const int GWL_STYLE = -16;
		const int GWL_EXSTYLE = -20;
		const uint WS_CAPTION = 0xC00000;
		const uint WS_VISIBLE = 0x10000000;
		const uint MY_INNERWIN = 0x40000000;
		const uint MY_INNERWINSUB = 0x80000000;
		[DllImport("user32.dll", SetLastError=true)]
		static extern int GetWindowLong(IntPtr hWnd, int nItem);

		private void _Init()
		{
			if(tblWin == null)
			{
				tblWinTMP = new DataTable();
				tblWinTMP.Columns.Add("handle");
				tblWinTMP.Columns.Add("classname");
				tblWinTMP.Columns.Add("title");
				tblWinTMP.Columns.Add("style");
				tblWinTMP.Columns.Add("exstyle");
				tblWinTMP.Columns.Add("pid");
				tblWinTMP.Columns.Add("process");
				tblWinTMP.Columns.Add("processname");
				tblWinTMP.Columns.Add("processfile");
				tblWin = tblWinTMP.Copy();
			}
			shellHandle = IntPtr.Zero;
		}

		public Shell()
		{
			_Init();
		}
 
		public void Dispose()
		{
		}

		public virtual bool NewWindow(string sFullPathName)
		{
			return NewWindow(sFullPathName, "", ProcessWindowStyle.Normal);
		}
		public virtual bool NewWindow(string sFullPathName, string sArguments)
		{
			return NewWindow(sFullPathName, sArguments, ProcessWindowStyle.Normal);
		}
		public virtual bool NewWindow(string sFullPathName, string sArguments, ProcessWindowStyle pWindowStyle)
		{
			ProcessStartInfo psi = new ProcessStartInfo();
			psi.FileName = sFullPathName;
			psi.Arguments = sArguments;
			psi.WindowStyle = pWindowStyle;
			shellHandle = IntPtr.Zero;
			bool bOK = false;
			try
			{
				//wait for nCnt/10 seconeds
				int nCnt = MILLISECONDS_TIMEOUT / 100;
				Process process = Process.Start(psi);
				bOK = getWindowByProcessID(process.Id);
				while(!bOK && nCnt > 0)
				{
					nCnt--;
					Application.DoEvents();
					System.Threading.Thread.Sleep(100);
					bOK = getWindowByProcessID(process.Id);
				}
			}
			catch
			{
				return false;
			}
			return bOK;
		}

		protected virtual void getAllWindow()
		{
			//init various
			tblWinTMP.Clear();
			tblWin.Clear();
			//perhaps set by parent,set to static various
			sProcessNameFilterTMP = sProcessNameFilter;

			//for get all windows,start from desktop
			IntPtr hand = GetDesktopWindow();
			IntPtr handPara = hand;
			//loop for get all chile windows
			EnumChildWindows(hand, new EnumChildProc(EnumChildProcEntry), ref handPara);
			//copy from static to various
			tblWin = tblWinTMP.Copy();
			tblWinTMP.Clear();
			if(sProcessNameFilterTMP != null && sProcessNameFilterTMP.Equals("EXCEL"))
			{
				//treate Excel activaty window and MS-SDIa,XLMAIN
				//one excel file have two handle,del it(XLMAIN handle)
				Hashtable tblExcelFile = new Hashtable();
				foreach(DataRow dr in tblWin.Rows)
				{
					if(dr["classname"].ToString().Equals("MS-SDIa"))
					{
						tblExcelFile[dr["title"].ToString()] = "true";
					}
				}
				for(int i = tblWin.Rows.Count - 1; i >= 0; i--)
				{
					if(tblWin.Rows[i]["classname"].ToString().Equals("XLMAIN"))
					{
						string sTitle = tblWin.Rows[i]["title"].ToString();
						if(sTitle.StartsWith("Microsoft Excel - "))
						{
							sTitle = sTitle.Substring(18);
						}
						if(tblExcelFile.ContainsKey(sTitle))
						{
							tblWin.Rows.RemoveAt(i);
						}
					}
				}
			}
		}
		private static int EnumChildProcEntry(IntPtr hwnd, ref IntPtr lParam)
		{
			int style = GetWindowLong(hwnd, GWL_STYLE);
			int exstyle = GetWindowLong(hwnd, GWL_EXSTYLE);
			if((style & WS_VISIBLE) == WS_VISIBLE
				&& (style & WS_CAPTION) == WS_CAPTION
				&& (style & MY_INNERWIN) != MY_INNERWIN
				)
			{
				StringBuilder sbClassName = new StringBuilder(256);
				GetClassName(hwnd, sbClassName, sbClassName.Capacity);
				string sClassName = sbClassName.ToString();
				if("IDEOwner".Equals(sClassName)
					//|| "XLMAIN".Equals(sClassName)
					)
				{
					return 1;
				}
				uint pid;
				GetWindowThreadProcessId(hwnd, out pid);
				Process process = System.Diagnostics.Process.GetProcessById((int)pid);
				if(sProcessNameFilterTMP != null && !sProcessNameFilterTMP.Equals(process.ProcessName))
				{
					return 1;
				}

				StringBuilder stitle = new StringBuilder(1024);
				GetWindowText((int)hwnd, stitle, stitle.Capacity);
				DataRow dr = tblWinTMP.NewRow();
				dr["handle"] = (int)hwnd;
				dr["title"] = stitle.ToString();
				dr["classname"] = sClassName;
				dr["style"] = style;
				dr["exstyle"] = exstyle;
				dr["pid"] = pid;
				dr["process"] = process;
				dr["processname"] = process.ProcessName;
				dr["processfile"] = process.MainModule.FileName;
				tblWinTMP.Rows.Add(dr);
			}
			return 1;
		}

		public virtual bool getWindowByProcessName(string sProcessName)
		{ 
			return getWindowByProcessName(sProcessName, 0);
		}
		public virtual bool getWindowByProcessName(string sProcessName, int nIndex)
		{ 
			this.shellHandle = IntPtr.Zero;
			this.shellInfo = null;
			getAllWindow();
			foreach(DataRow dr in tblWin.Rows)
			{
				if(dr["processname"].ToString().Equals(sProcessName))
				{
					if(nIndex == 0)
					{
						this.shellHandle = (IntPtr)int.Parse(dr["handle"].ToString());
						this.shellInfo = dr;
						return true;
					}
					else
					{
						nIndex--;
					}
				}
			}
			return false;
		}

		public virtual bool getWindowByTitle(string sTitle)
		{ 
			return getWindowByTitle(sTitle, 0);
		}
		public virtual bool getWindowByTitle(string sTitle, int nIndex)
		{ 
			this.shellHandle = IntPtr.Zero;
			this.shellInfo = null;
			getAllWindow();
			foreach(DataRow dr in tblWin.Rows)
			{
				if(dr["title"].ToString().Equals(sTitle))
				{
					if(nIndex == 0)
					{
						this.shellHandle = (IntPtr)int.Parse(dr["handle"].ToString());
						this.shellInfo = dr;
						return true;
					}
					else
					{
						nIndex--;
					}
				}
			}
			return false;
		}

		public virtual bool getWindowByProcessID(int nProcessID)
		{ 
			this.shellHandle = IntPtr.Zero;
			this.shellInfo = null;
			getAllWindow();
			foreach(DataRow dr in tblWin.Rows)
			{
				if(int.Parse(dr["pid"].ToString()) == nProcessID)
				{
					this.shellHandle = (IntPtr)int.Parse(dr["handle"].ToString());
					this.shellInfo = dr;
					return true;
				}
			}
			return false;
		}

		#region The CmdShow
		/// <summary>
		/// The CmdShow
		/// </summary>
		/// <remarks>to show windows</remarks>
		public enum CmdShow
		{
			/// <summary>Hide the window</summary>
			SW_HIDE = 0,
			/// <summary>Maximize the window</summary>
			SW_MAXIMIZE = 3,
			/// <summary>Minimize the window</summary>
			SW_MINIMIZE = 6,
			/// <summary>Restore the window (not maximized nor minimized)</summary>
			SW_RESTORE = 9,
			/// <summary>Show the window</summary>
			SW_SHOW = 5,
			/// <summary>Show the window maximized</summary>
			SW_SHOWMAXIMIZED = 3,
			/// <summary>Show the window minimized</summary>
			SW_SHOWMINIMIZED = 2,
			/// <summary>Show the window minimized but do not activate it</summary>
			SW_SHOWMINNOACTIVE = 7,
			/// <summary>Show the window in its current state but do not activate it</summary>
			SW_SHOWNA = 8,
			/// <summary>Show the window in its most recent size and position but do not activate it</summary>
			SW_SHOWNOACTIVATE = 4,
			/// <summary>Show the window and activate it (as usual)</summary>
			SW_SHOWNORMAL = 1
		}
		#endregion
		public virtual bool ShowWindow(CmdShow nCmdShow)
		{
			if(this.shellHandle == IntPtr.Zero)
			{
				return false;
			}
			ShowWindow(this.shellHandle, (int)nCmdShow);
			WaitForInputIdle(this.shellHandle, MILLISECONDS_TIMEOUT);
			return true;
		}
		[DllImport("user32.dll")]
		static extern IntPtr GetForegroundWindow();

		public virtual bool isLives()
		{
			if(this.shellHandle == IntPtr.Zero)
			{
				return false;
			}
			if(IsWindow(this.shellHandle.ToInt32()) > 0)
			{
				return true;
			}
			return false;
		}

		public const int WM_ACTIVATE = 0x6;
		public virtual bool SetForegroundWindow()
		{
			if(!isLives())
			{
				return false;
			}
			bool bWait = false;
			if(GetForegroundWindow().ToInt32() != this.shellHandle.ToInt32())
			{
				bWait = true;
			}
			SetForegroundWindow(this.shellHandle);
			//WA_CLICKACTIVE:2
			SendMessage(this.shellHandle, WM_ACTIVATE, (IntPtr)2, (IntPtr)0);
			SetFocus(this.shellHandle);
			if(bWait)
			{
				wait(300);
			}
			return true;
		}

		/*
			Key Code 
			BACKSPACE {BACKSPACE}, {BS}, or {BKSP} 
			BREAK {BREAK} 
			CAPS LOCK {CAPSLOCK} 
			DEL or DELETE {DELETE} or {DEL} 
			DOWN ARROW {DOWN} 
			END {END} 
			ENTER {ENTER}or ~ 
			ESC {ESC} 
			HELP {HELP} 
			HOME {HOME} 
			INS or INSERT {INSERT} or {INS} 
			LEFT ARROW {LEFT} 
			NUM LOCK {NUMLOCK} 
			PAGE DOWN {PGDN} 
			PAGE UP {PGUP} 
			PRINT SCREEN {PRTSC} (reserved for future use) 
			RIGHT ARROW {RIGHT} 
			SCROLL LOCK {SCROLLLOCK} 
			TAB {TAB} 
			UP ARROW {UP} 
			F1 {F1} 
			...
			F16 {F16} 
			Keypad add {ADD} 
			Keypad subtract {SUBTRACT} 
			Keypad multiply {MULTIPLY} 
			Keypad divide {DIVIDE} 

			to hold down SHIFT while E and C are pressed, use "+(EC)"
			to hold down SHIFT while E is pressed, followed by C without SHIFT, use "+EC".
			SHIFT + 
			CTRL ^ 
			ALT % 
		*/
		public virtual bool SendKey(string sKeys)
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			SendKeys.SendWait(sKeys);
			return true;
		}

		private const int KEYEVENTF_KEYUP = 0x0002;
		//private const int VK_MENU = 0x12;
		public virtual bool KeyEvent(Keys keys)
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			KeyEventArgs keye = new KeyEventArgs(keys);
			if(keye.Control)
			{
				keybd_event((byte)Keys.ControlKey, 0, 0, 0);//down Ctrl
			}
			if(keye.Alt)
			{
				keybd_event((byte)Keys.Menu, 0, 0, 0);//down Ctrl
			}
			if(keye.Shift)
			{
				keybd_event((byte)Keys.ShiftKey, 0, 0, 0);//down Ctrl
			}
			keybd_event((byte)keye.KeyCode, 0, 0, 0);
			keybd_event((byte)keye.KeyCode, 0, KEYEVENTF_KEYUP, 0);
			if(keye.Control)
			{
				keybd_event((byte)Keys.ControlKey, 0, KEYEVENTF_KEYUP, 0);//up Ctrl
			}
			if(keye.Alt)
			{
				keybd_event((byte)Keys.Menu, 0, KEYEVENTF_KEYUP, 0);//up Ctrl
			}
			if(keye.Shift)
			{
				keybd_event((byte)Keys.ShiftKey, 0, KEYEVENTF_KEYUP, 0);//up Ctrl
			}
			return true;
		}

		private const int WM_LBUTTONDOWN = 0x201;
		private const int WM_LBUTTONUP = 0x202;
		private const int WM_RBUTTONDOWN = 0x204;
		private const int WM_RBUTTONUP = 0x205;
		/*
		Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
		Public Const MOUSEEVENTF_MIDDLEUP = &H40
		Public Const MOUSEEVENTF_RIGHTDOWN = &H8
		Public Const MOUSEEVENTF_RIGHTUP = &H10
		Public Const MOUSEEVENTF_MOVE = &H1
		*/
		private const int MOUSEEVENTF_LEFTDOWN = 0x2;
		private const int MOUSEEVENTF_LEFTUP = 0x4;
		public virtual bool MouseEvent(int x, int y)
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			Cursor.Position = new Point(x, y);
			mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
			mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);

			//SendMessage(this.shellHandle, WM_LBUTTONDOWN, (IntPtr)0, (IntPtr)0);
			//SendMessage(this.shellHandle, WM_LBUTTONUP, (IntPtr)0, (IntPtr)0);
			return true;
		}

		public virtual bool EditCopy()
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			KeyEvent(Keys.C | Keys.Control);
			//SendMessage(this.shellHandle, WM_COPY, (IntPtr)0, (IntPtr)0);
			return true;
		}
		public virtual bool EditCut()
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			KeyEvent(Keys.X | Keys.Control);
			//SendMessage(this.shellHandle, WM_CUT, (IntPtr)0, (IntPtr)0);
			return true;
		}
		public virtual bool EditPaste()
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			KeyEvent(Keys.V | Keys.Control);
			//SendMessage(this.shellHandle, WM_PASTE, (IntPtr)0, (IntPtr)0);
			return true;
		}

		public void wait(int nMillisecondsTimeout)
		{
			int nCnt = nMillisecondsTimeout / 100;
			while(nCnt >= 0)
			{
				nCnt--;
				Application.DoEvents();
				WaitForInputIdle(this.shellHandle, MILLISECONDS_TIMEOUT);
				System.Threading.Thread.Sleep(100);
			}
		}

		public virtual bool WindowToClipboard()
		{
			if(!SetForegroundWindow())
			{
				return false;
			}
			Bitmap bmp = cc.Capture.CaptureWindow(this.shellHandle);
			if(bmp == null)
			{
				Clipboard.SetDataObject(string.Empty, true);
				return false;
			}
			else
			{
				Clipboard.SetDataObject(bmp, true);
			}
			return true;
		}

		public static void ScreenToClipboard()
		{
			Clipboard.SetDataObject(cc.Capture.CaptureScreen(), true);
		}


		public virtual DataTable getAllWindowInfoTbl()
		{ 
			getAllWindow();
			return tblWin;
		}

		public virtual string getAllWindowInfo()
		{ 
			getAllWindow();
			StringBuilder sb = new StringBuilder();
			foreach(DataRow dr in tblWin.Rows)
			{
				for(int i = 0; i < tblWin.Columns.Count; i++)
				{
					sb.Append(tblWin.Columns[i].ColumnName + ":" + dr[i] + "\r\n");
				}
				sb.Append("\r\n");
			}
			return sb.ToString();
		}

	}

	public class IE : Shell
	{
		private static HTMLDocument htmlDocument = null;

		public HTMLDocument Document
		{
			get
			{
				return htmlDocument;
			}
			set
			{
				htmlDocument = value;
			}
		}

		public IE()
		{
			sProcessNameFilter = "iexplore";
		}
 
		new public void Dispose()
		{
		}

		public class htmlItem
		{
			private HTMLDocument htmlDoc = null;
			private string htmlID = null;
			public htmlItem(HTMLDocument htmlDoc, string htmlID)
			{
				this.htmlDoc = htmlDoc;
				this.htmlID = htmlID;
			}
			//return Encode for JScript
			static private string JSEncode(string sTxt)
			{
				if(sTxt == null)
				{
					return "";
				}
				return sTxt.Replace("\\", "\\\\").Replace("\r", "").Replace("\n", "\\r\\n").Replace("'", "\\'");
			}

			public string Text
			{
				get
				{
					try
					{
						string sScript = @"
if(!document.all('JS_ReturnValue')){
	var s = '<textarea style=\'display:none\' id=\'JS_ReturnValue\' name=\'JS_ReturnValue\'></textarea>';
	element = document.createElement(s);
	document.body.appendChild(element);
}
document.all('JS_ReturnValue').innerText = document." + htmlID + ";";
						htmlDoc.parentWindow.execScript(sScript, "JScript");
						return htmlDoc.getElementById("JS_ReturnValue").innerText;
					}
					catch(Exception exp)
					{
						Console.WriteLine("Err htmlItem:" + exp.Message);
					}
					return null;
				}
				set
				{
					try
					{
						string sScript = "document." + htmlID + " = '" + JSEncode(value) + "'";
						htmlDoc.parentWindow.execScript(sScript, "JScript");
					}
					catch(Exception exp)
					{
						Console.WriteLine("Err htmlItem:" + exp.Message);
					}
				}
			}
		}
		public htmlItem this[string htmlID]
		{
			get
			{
				return new htmlItem(htmlDocument, htmlID);
			}
		}
		public IHTMLElement item(string name)
		{
			return (IHTMLElement)(htmlDocument.all.item(name, 0));
		}
		public IHTMLElement item(string name, int index)
		{
			return (IHTMLElement)(htmlDocument.all.item(name, index));
		}
		public IHTMLInputElement formitem(string name)
		{
			return (IHTMLInputElement)(htmlDocument.forms.item(name, 0));
		}
		public IHTMLInputElement formitem(string name, int index)
		{
			return (IHTMLInputElement)(htmlDocument.all.item(name, index));
		}

		private bool _getIEDocument()
		{
			IntPtr hwnd = shellHandle;

			IntPtr handPara = IntPtr.Zero;
			htmlDocument = null;
			EnumChildWindows(hwnd, new EnumChildProc(EnumChildProcEntry), ref handPara);
			if(htmlDocument != null)
			{
				return true;
			}
			return false;
		}
		private static int EnumChildProcEntry(IntPtr hwnd, ref IntPtr lParam)
		{
			StringBuilder sbClassName = new StringBuilder(256);
			GetClassName(hwnd, sbClassName, sbClassName.Capacity);

			if(sbClassName.ToString().Equals("Internet Explorer_Server"))
			{
				string text1 = "WM_HTML_GETOBJECT";
				int num1 = RegisterWindowMessage(ref text1);
				int num2 = 0;
				SendMessageTimeout(hwnd, num1, 0, 0, 2, 1000, ref num2);
				if(num2 > 0)
				{
					HTMLDocument htmlDoc = null;
					Guid guid1 = new Guid("626FC520-A41E-11CF-A731-00A0C9082637");
					int num3 = ObjectFromLresult(num2, ref guid1, 0, ref htmlDoc);
					if (num3 == 0 && htmlDoc != null)
					{
						htmlDocument = htmlDoc;
						return 0;
					}
				}
			}
			return 1;
		}

		public bool NewWindow()
		{
			return NewWindow("about:blank", ProcessWindowStyle.Normal);
		}
		new public bool NewWindow(string sURL)
		{
			return NewWindow(sURL, ProcessWindowStyle.Normal);
		}
		public bool NewWindow(string sURL, ProcessWindowStyle pWindowStyle)
		{
			if(base.NewWindow("IExplore.exe", sURL, pWindowStyle))
			{
				if(_getIEDocument())
				{
					return true;
				}
			}
			return false;
		}

		public override bool getWindowByTitle(string sTitle)
		{ 
			return getWindowByTitle(sTitle, 0);
		}

		public override bool getWindowByTitle(string sTitle, int nIndex)
		{ 
			if(!base.getWindowByTitle(sTitle, nIndex))
			{
				if(!base.getWindowByTitle(sTitle + " - Microsoft Internet Explorer", nIndex))
				{
					return false;
				}
			}
			return _getIEDocument();
		}

		public bool getWindowByURL(string sURL)
		{
			getAllWindow();
			if(sURL.Substring(1, 2).Equals(":\\") || sURL.Substring(1, 2).Equals("://"))
			{
				sURL = "file://" + sURL;
			}
			foreach(DataRow dr in tblWin.Rows)
			{
				this.shellHandle = (IntPtr)int.Parse(dr["handle"].ToString());
				if(_getIEDocument())
				{
					if(htmlDocument.url.Equals(sURL))
					{
						this.shellInfo = dr;
						return true;
					}
				}
			}
			this.shellHandle = IntPtr.Zero;
			this.shellInfo = null;
			htmlDocument = null;
			return false;
		}

		public bool execScript(string sCode)
		{
			return execScript(sCode, "JScript");
		}
		public bool execScript(string sCode, string sLanguage)
		{
			if(shellHandle == IntPtr.Zero || Document == null)
			{
				return false;
			}
			try
			{
				Document.parentWindow.execScript(sCode, sLanguage);
				return true;
			}
			catch(Exception exp)
			{
				Console.WriteLine("Err htmlItem:" + exp.Message);
			}
			return false;
		}

		public override bool isLives()
		{
			if(!base.isLives())
			{
				return false;
			}
			try
			{
				string s = Document.url;
				if(Document.url != null)
				{
					return true;
				}
			}
			catch
			{
			}
			return false;
		}

		public void waitWhileBusy()
		{
			waitWhileBusy(MILLISECONDS_TIMEOUT);
		}
		public void waitWhileBusy(int nMillisecondsTimeout)
		{
			if(shellHandle == IntPtr.Zero || Document == null)
			{
				return;
			}
			int nCnt = nMillisecondsTimeout / 100;
			while(nCnt >= 0)
			{
				if(false)
				{
					return;
				}
				nCnt--;
				Application.DoEvents();
				System.Threading.Thread.Sleep(100);
			}
		}

		public override string getAllWindowInfo()
		{ 
			HTMLDocument tmphtmlDoc = htmlDocument;
			IntPtr tmpHandle = shellHandle;
			getAllWindow();
			StringBuilder sb = new StringBuilder();
			foreach(DataRow dr in tblWin.Rows)
			{
				for(int i = 0; i < tblWin.Columns.Count; i++)
				{
					sb.Append(tblWin.Columns[i].ColumnName + ":" + dr[i] + "\r\n");
				}
				shellHandle = (IntPtr)int.Parse(dr["handle"].ToString());
				if(_getIEDocument())
				{
					sb.Append("URL:" + htmlDocument.url + "\r\n");
				}
				sb.Append("\r\n");
			}
			htmlDocument = tmphtmlDoc;
			shellHandle = tmpHandle;
			return sb.ToString();
		}

	}

	public class Excel : Shell
	{
		private cc.Variant excel = null;
		public Excel()
		{
			sProcessNameFilter = "EXCEL";

		}
 
		new public void Dispose()
		{
			if(excel != null)
			{
				excel.ReleaseComObject();
			}
		}

		private bool _getExcelFromHandle()
		{
			//excel = null;
			//Variant excelv = Variant.CreateComInstance("Excel.Application");
			//Process process = System.Diagnostics.Process.GetProcessById(int.Parse(shellInfo["pid"].ToString()));
			if(shellHandle == IntPtr.Zero)
			{
				return false;
			}
			return true;
		}

		new public bool NewWindow(string sFullPathName)
		{
			return NewWindow(sFullPathName, ProcessWindowStyle.Normal);
		}
		public bool NewWindow(string sFullPathName, ProcessWindowStyle pWindowStyle)
		{
			if(!File.Exists(sFullPathName))
			{
				try
				{
					//if not create empty file,then open error
					File.Create(sFullPathName, 1).Close();
				}
				catch
				{
					return false;
				}
			}
			if(base.NewWindow("EXCEL.EXE", sFullPathName, pWindowStyle))
			{
				return _getExcelFromHandle();
			}
			return false;
		}

		public bool getWindowByExcelFileName(string sExcelFileName)
		{ 
			string sNameOnly = Path.GetFileName(sExcelFileName);
			if(!base.getWindowByTitle(sNameOnly))
			{
				if(!base.getWindowByTitle("Microsoft Excel - " + sNameOnly))
				{
					return false;
				}
			}
			return _getExcelFromHandle();
		}

		public object this[string propertyName, params object[] parameters]
		{
			get
			{
				if(excel == null)
				{
					return null;
				}
				else
				{
					return excel[propertyName, parameters];
				}
			}
			set
			{
				if(excel != null)
				{
					excel[propertyName, parameters] = value;
				}
			}
		}

	}

	#region class for access to Capture
	class Capture
	{
		[Serializable, StructLayout(LayoutKind.Sequential)]
			public struct RECT
		{
			public int Left;    
			public int Top;    
			public int Right;    
			public int Bottom;
		}

		[DllImport("user32.dll")]
		private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

		[DllImport("user32.dll", EntryPoint="GetDesktopWindow")]
		private static extern IntPtr GetDesktopWindow();

		[DllImport("user32.dll")]
		static extern IntPtr GetForegroundWindow();

		[DllImport("user32.dll", EntryPoint="GetWindowDC")]
		private static extern IntPtr GetWindowDC(Int32 ptr);

		[DllImport("user32.dll", EntryPoint="ReleaseDC")]
		private static extern IntPtr ReleaseDC(IntPtr hWnd,IntPtr hDc);

		[System.Runtime.InteropServices.DllImportAttribute("gdi32.dll")]
		private static extern bool BitBlt(
			IntPtr hdcDest, // handle to destination DC
			int nXDest,  // x-coord of destination upper-left corner
			int nYDest,  // y-coord of destination upper-left corner
			int nWidth,  // width of destination rectangle
			int nHeight, // height of destination rectangle
			IntPtr hdcSrc,  // handle to source DC
			int nXSrc,   // x-coordinate of source upper-left corner
			int nYSrc,   // y-coordinate of source upper-left corner
			System.Int32 dwRop  // raster operation code
			);

		public static Bitmap CaptureWindow(IntPtr hWnd)
		{
			//get the window size
			RECT rect;
			GetWindowRect(hWnd, out rect);
			return CaptureWindow(hWnd, 0, 0, rect.Right - rect.Left, rect.Bottom - rect.Top);
		}

		/// <summary>
		/// Captures the window or part thereof to a bitmap image.
		/// </summary>
		/// <param name="wndHWND">window handle</param>
		/// <param name="x">x location in window</param>
		/// <param name="y">y location in window</param>
		/// <param name="width">width of capture area</param>
		/// <param name="height">height of capture area</param>
		/// <returns>window bitmap</returns>
		public static Bitmap CaptureWindow(IntPtr hWnd, int x, int y, int width, int height)
		{
			//Here we get the handle to the desktop device context.
			IntPtr hDC = GetWindowDC(hWnd.ToInt32());
			if(hDC == IntPtr.Zero)
			{
				return null;
			}
			//Here we make a compatible device context in memory for screen device context.
			Graphics g1 = Graphics.FromHdc(hDC);
			Bitmap img = new Bitmap(width, height, g1);
			Graphics g2 = Graphics.FromImage(img);

			IntPtr dc1 = g1.GetHdc();
			IntPtr dc2 = g2.GetHdc();
			BitBlt(dc2, 0, 0, width, height, dc1, 0, 0, 13369376);
			g1.ReleaseHdc(dc1);
			g2.ReleaseHdc(dc2);
			Bitmap bmp = System.Drawing.Image.FromHbitmap(img.GetHbitmap());
			return bmp;
		}
		public static Bitmap CaptureScreen()
		{
			return CaptureWindow(GetDesktopWindow());
		}
		public static Bitmap CaptureScreen(int x, int y, int width, int height)
		{
			return CaptureWindow(GetDesktopWindow(), x, y, width, height);
		}
		public static Bitmap CaptureActivateWindow()
		{
			return CaptureWindow(GetForegroundWindow());
		}

	}
	#endregion

	#region class for access to object(Variant)
	//class for access to object
	// see sample in Test
	public class Variant : IDisposable
	{
		object _obj;
		Type _objType;
		bool _isComObj = false;

		Variant()
		{
		}

		public static Variant CreateComInstance(string progID)
		{
			Variant variant = new Variant();
			variant._objType = Type.GetTypeFromProgID(progID);
			variant._obj = Activator.CreateInstance(variant._objType);
			variant._isComObj = true;
			return variant;
		}
		public static Variant GetComInstance(string progID)
		{
			Variant variant = new Variant();
			variant._objType = Type.GetTypeFromProgID(progID);
			variant._obj = System.Runtime.InteropServices.Marshal.GetActiveObject(progID);
			variant._isComObj = true;
			return variant;
		}
		public static Variant CreateInstance(string objectTypeName)
		{
			Variant variant = new Variant();
			variant._objType = Type.GetType(objectTypeName);
			variant._obj = Activator.CreateInstance(variant._objType);
			return variant;
		}
		public static Variant CreateInstance(object obj)
		{
			Variant variant = new Variant();
			variant._objType = obj.GetType();
			variant._obj = obj;
			variant._isComObj = obj.GetType().ToString() == "System.__ComObject";
			return variant;
		}
		public int ReleaseComObject()
		{
			try
			{
				if (_isComObj)
					return Marshal.ReleaseComObject(_obj);
			}
			catch
			{
			}
			return -1;
		}
		public object InvokeMethod(string methodName,params object[] parameters)
		{
			try
			{
				return _objType.InvokeMember(methodName,
					BindingFlags.InvokeMethod,
					null, _obj, parameters);
			}
			catch
			{
				throw;
			}
		}

		public object GetPropertyValue(string propertyName, params object[] parameters)
		{
			object oretu = null;
			try
			{
				oretu = _objType.InvokeMember(propertyName,
					BindingFlags.GetProperty, null, _obj, parameters);
			}
			catch
			{
			}
			return oretu;
		}
		public void SetPropertyValue(string propertyName, object value, params object[] parameters)
		{
			if (value as DBNull != null)
				value = null;
			string s = value as string;
			if (s != null && s.Length == 0)
			{
				value = null;
			}

			if (_isComObj)
			{
				try
				{
					object[] args = new object[parameters.Length + 1];
					Array.Copy(parameters, 0, args, 1, parameters.Length);
					args[0] = value;
					_objType.InvokeMember(propertyName,
						BindingFlags.SetProperty,
						null, _obj, args);
				}
				catch
				{
				}
				return;
			}

			PropertyInfo pi = _objType.GetProperty(propertyName);
			if (pi != null)
			{
				if (pi.CanWrite)
				{
					try
					{
						pi.SetValue(_obj, null == value ?
							null :
							(pi.PropertyType.IsInterface ?
						value :
							Convert.ChangeType(value, pi.PropertyType)
							),
							parameters);
					}
					catch (Exception exp)
					{
						throw new Exception(String.Format("err for set:{0}, property type:{1}, value type:{2}, msgÅF{3}", pi.Name, pi.PropertyType.FullName, value.GetType().FullName, null == exp.InnerException ? exp.Message : exp.InnerException.Message));
					}
				}
				else
					throw new Exception(pi.PropertyType.ToString());
			}
			else
				throw new Exception(String.Format("property:{0} is not defined in {1}", pi.Name, _obj.GetType().FullName));
		}

		public object this[string propertyName]
		{
			get
			{
				return GetPropertyValue(propertyName, new object[]{});
			}
			set
			{
				SetPropertyValue(propertyName, value, new object[]{});
			}
		}

		public object this[string propertyName, params object[] parameters]
		{
			get
			{
				return GetPropertyValue(propertyName, parameters);
			}
			set
			{
				SetPropertyValue(propertyName, value, parameters);
			}
		}

		void IDisposable.Dispose()
		{
			if (_isComObj)
				Marshal.ReleaseComObject(_obj);
		}

		public static void Test()
		{
			//sample for excel
			Variant excel = Variant.CreateComInstance("Excel.Application");
			excel["Visible"] = true;
			Variant workBooks = Variant.CreateInstance(excel["Workbooks"]);
			using (Variant workBook = Variant.CreateInstance(workBooks.InvokeMethod("Add")))
			{
				for (int i = 1; i < 10; i++)
				{
					for (int j = 1; j < 10; j++)
					{
						using (Variant cell = Variant.CreateInstance(excel["Cells", i, j]))
						{
							cell["Value"] = i * j;
							Console.WriteLine(cell["Value"]);
						}
					}
				}
			}
			//excel.InvokeMethod("Quit");

			Variant varie = Variant.CreateComInstance("InternetExplorer.Application");
			varie["Visible"] = true;
			varie.InvokeMethod("Navigate", new object[]{"about:blank"});
			bool isBusy = varie["Busy"].ToString().Equals("true");
			Console.WriteLine("isBusy:" + isBusy);
			while(varie["Busy"].ToString().Equals("true"))
			{
				Application.DoEvents();
			}

			Variant varIEDocument = Variant.CreateInstance(varie["Document"]);
			Variant varIEall = Variant.CreateInstance(varIEDocument["all"]);

			varIEDocument.InvokeMethod("write", new object[]{"xyz"});
			//this.all["codetitle"] = "abc";
			//Console.WriteLine(this.all["HTML", "innerHTML"]);

			//using (Variant cell = Variant.CreateInstance(all["codetitle"]))
			{
				//cell["Value"] = "abccc";
				//Console.WriteLine(cell["Value"]);
			}

			//varie.InvokeMethod("Quit");
		}
	}
	#endregion

}
