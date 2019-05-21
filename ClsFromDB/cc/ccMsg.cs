/*Modify List:
 * 20050725,Color&Font add
 * 
 */
using System;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Msg
	/// </summary>
	public class Msg
	{
		private System.Windows.Forms.RichTextBox txtMsg;
		//デフォルトFont, Color
		private System.Drawing.Font defaultFont;
		private System.Drawing.Color defaultColor;
		public Msg(System.Windows.Forms.RichTextBox txtMsg)
		{
			this.txtMsg = txtMsg;
			//デフォルトFont, Colorの設定
			defaultFont = txtMsg.Font;
			defaultColor = txtMsg.ForeColor;
		}

		/// <summary>
		/// デフォルトFontの設定と取得
		/// </summary>
		public System.Drawing.Font DefaultFont
		{
			get
			{
				return defaultFont;
			}
			set
			{
				defaultFont = value;
			}
		}

		/// <summary>
		/// デフォルトColorの設定と取得
		/// </summary>
		public System.Drawing.Color DefaultColor
		{
			get
			{
				return defaultColor;
			}
			set
			{
				defaultColor = value;
			}
		}

		/// <summary>
		/// メッセージにフォーカスをセット
		/// </summary>
		public void Focus()
		{
			txtMsg.Focus();
		}

		/// <summary>
		/// メッセージをクリアする
		/// </summary>
		public void clear()
		{
			txtMsg.Text = "";
		}

		/// <summary>
		/// メッセージを出力
		/// </summary>
		public void print(string msg)
		{
			print(msg, defaultColor, defaultFont);
		}
		/// <summary>
		/// 色を指定し、メッセージを出力
		/// </summary>
		public void print(string msg, System.Drawing.Color cor)
		{
			print(msg, cor, defaultFont);
		}
		/// <summary>
		/// フォントを指定し、メッセージを出力
		/// </summary>
		public void print(string msg, System.Drawing.Font font)
		{
			print(msg, defaultColor, font);
		}
		/// <summary>
		/// 色、フォントを指定し、メッセージを出力
		/// </summary>
		public void print(string msg, System.Drawing.Color cor, System.Drawing.Font font)
		{
			//出力
			foreach(string s in msg.Split('\n'))
			{
				//指定色を設定
				txtMsg.SelectionLength = 0;
				txtMsg.SelectionStart = txtMsg.Text.Length;
				txtMsg.SelectionColor = cor;
				txtMsg.SelectionFont = font;
				txtMsg.AppendText(s);
				System.Windows.Forms.Application.DoEvents();
			}
		}

		/// <summary>
		/// 改行し、メッセージを出力
		/// </summary>
		public void println(string msg)
		{
			print(msg + "\r\n", defaultColor, defaultFont);
		}
		/// <summary>
		/// 色を指定し、改行し、メッセージを出力
		/// </summary>
		public void println(string msg, System.Drawing.Color cor)
		{
			print(msg + "\r\n", cor, defaultFont);
		}
		/// <summary>
		/// フォントを指定し、改行し、メッセージを出力
		/// </summary>
		public void println(string msg, System.Drawing.Font font)
		{
			print(msg + "\r\n", defaultColor, font);
		}
		/// <summary>
		/// 色、フォントを指定し、改行し、メッセージを出力
		/// </summary>
		public void println(string msg, System.Drawing.Color cor, System.Drawing.Font font)
		{
			print(msg + "\r\n", cor, font);
		}
	}

}
