/*Modify List:
 * 
 * 
 */
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.ComponentModel;
using System.Collections;
using System.Diagnostics;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Util
	/// </summary>
	/// <summary>
	/// Inherits from System.Windows.Form.NativeWindow. Provides an Event for Message handling
	/// </summary>
	public class HotKey: System.Windows.Forms.NativeWindow
	{
		[DllImport("user32.dll", SetLastError=true)]
		private static extern bool RegisterHotKey(
			IntPtr hWnd, //handle to window
			int id, //hot key identifier
			int fsModifiers, //key-modifier options
			int vk //virtual-key code
			);
		
		[DllImport("user32.dll", SetLastError=true)]
		private static extern bool UnregisterHotKey(
			IntPtr hWnd, //handle to window
			int id //hot key identifier
			);

		public HotKey()
		{
			CreateParams parms = new CreateParams();
			this.CreateHandle(parms);
		}
 
		public void Dispose()
		{
			if(this.Handle != IntPtr.Zero)
			{
				this.DestroyHandle();
			}
		}

		/// <summary>
		/// Defines a delegate for Message handling
		/// </summary>
		public delegate void HotKeyEventHandler(KeyEventArgs e);
		private Hashtable htHotKey = null;
		private static string ARGS_PRE = "args";

		private enum Modifiers
		{
			MOD_ALT = 0x0001,
			MOD_CONTROL = 0x0002,
			MOD_SHIFT = 0x0004,
			MOD_WIN = 0x0008
		}

		public bool RegisterHotkey(Keys keys, HotKeyEventHandler ehandler)
		{
			if(htHotKey == null)
			{
				htHotKey = new Hashtable();
			}

			KeyEventArgs keye = new KeyEventArgs(keys);
			int htkey = int.Parse(keye.Modifiers.GetHashCode().ToString() + keye.KeyCode.GetHashCode().ToString());
			if(htHotKey.ContainsKey(htkey))
			{
				//have registerd this hotkey
				return false;
			}

			int keyModifiers = 0;
			if(keye.Alt)
			{
				keyModifiers += (int)Modifiers.MOD_ALT;
			}
			if(keye.Control)
			{
				keyModifiers += (int)Modifiers.MOD_CONTROL;
			}
			if(keye.Shift)
			{
				keyModifiers += (int)Modifiers.MOD_SHIFT;
			}
			if((keye.Modifiers & Keys.LWin) == Keys.LWin || (keye.Modifiers & Keys.RWin) == Keys.RWin)
			{
				keyModifiers += (int)Modifiers.MOD_WIN;
			}

			bool isOK = RegisterHotKey(this.Handle, htkey, keyModifiers, (int)keye.KeyCode);
			if(isOK)
			{
				htHotKey[htkey] = ehandler;
				htHotKey[ARGS_PRE + htkey] = keye;
			}
			return isOK;
		}

		public bool UnregisterHotkey(Keys keys)
		{
			//unregister hotkey
			if(htHotKey == null)
			{
				return false;
			}

			KeyEventArgs keye = new KeyEventArgs(keys);
			int htkey = int.Parse(keye.Modifiers.GetHashCode().ToString() + keye.KeyCode.GetHashCode().ToString());
			if(!htHotKey.ContainsKey(htkey))
			{
				//have registerd this hotkey
				return false;
			}
			bool isOK = UnregisterHotKey(this.Handle, htkey);
			if(isOK)
			{
				htHotKey.Remove(htkey);
				htHotKey.Remove(ARGS_PRE + htkey);
			}
			if(htHotKey.Count < 1)
			{
				htHotKey = null;
			}
			return isOK;
		}

		public bool UnregisterHotkeyAll()
		{
			//unregister hotkey
			if(htHotKey == null)
			{
				return false;
			}

			bool isAllOK = true;
			foreach(object okey in htHotKey.Keys)
			{
				if(!okey.ToString().StartsWith(ARGS_PRE))
				{
					bool isOK = UnregisterHotKey(this.Handle, (int)okey);
					if(!isOK)
					{
						isAllOK = false;
					}
				}
			}
			htHotKey = null;
			return isAllOK;
		}

		const int WM_HOTKEY = 0x0312;
		protected override void WndProc(ref Message m)
		{
			if(htHotKey != null && m.Msg == WM_HOTKEY)
			{
				int htkey = m.WParam.ToInt32();
				HotKeyEventHandler HotKeyPressed = (HotKeyEventHandler)htHotKey[htkey];
				if(HotKeyPressed != null)
				{
					HotKeyPressed((KeyEventArgs)htHotKey[ARGS_PRE + htkey]);
				}
			}
			base.WndProc(ref m);
		}

	}

}
