using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Data.SqlClient;

namespace ClsFromDB
{
	/// <summary>
	/// PageFromServer の概要の説明です。
	/// </summary>
	public class PageJSDebug : System.Windows.Forms.UserControl
	{
		public Form1 frmMain = null;

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnRun;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.Button btnNew;
		private System.Windows.Forms.Button btnClose;
		private ClsFromDB.PageJSDebugPage jspage;
		private System.Windows.Forms.Label label1;

		public PageJSDebug()
		{
			// この呼び出しは、Windows.Forms フォーム デザイナで必要です。
			InitializeComponent();
		}
		/// <summary>
		/// 使用されているリソースに後処理を実行します。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region コンポーネント デザイナで生成されたコード 
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.btnRun = new System.Windows.Forms.Button();
			this.btnCancel = new System.Windows.Forms.Button();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.jspage = new ClsFromDB.PageJSDebugPage();
			this.btnNew = new System.Windows.Forms.Button();
			this.btnClose = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			this.SuspendLayout();
			// 
			// btnRun
			// 
			this.btnRun.Location = new System.Drawing.Point(0, 0);
			this.btnRun.Name = "btnRun";
			this.btnRun.Size = new System.Drawing.Size(72, 24);
			this.btnRun.TabIndex = 12;
			this.btnRun.Text = "RunJScript";
			this.btnRun.Click += new System.EventHandler(this.btnRuns_Click);
			// 
			// btnCancel
			// 
			this.btnCancel.Enabled = false;
			this.btnCancel.Location = new System.Drawing.Point(72, 0);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.Size = new System.Drawing.Size(56, 24);
			this.btnCancel.TabIndex = 16;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// tabControl1
			// 
			this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Location = new System.Drawing.Point(0, 24);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(600, 248);
			this.tabControl1.TabIndex = 18;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.jspage);
			this.tabPage1.Location = new System.Drawing.Point(4, 21);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Size = new System.Drawing.Size(592, 223);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "JScript";
			// 
			// jspage
			// 
			this.jspage.Dock = System.Windows.Forms.DockStyle.Fill;
			this.jspage.Location = new System.Drawing.Point(0, 0);
			this.jspage.Name = "jspage";
			this.jspage.Size = new System.Drawing.Size(592, 223);
			this.jspage.TabIndex = 0;
			// 
			// btnNew
			// 
			this.btnNew.Location = new System.Drawing.Point(192, 5);
			this.btnNew.Name = "btnNew";
			this.btnNew.Size = new System.Drawing.Size(56, 19);
			this.btnNew.TabIndex = 12;
			this.btnNew.Text = "NewTab";
			this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
			// 
			// btnClose
			// 
			this.btnClose.Location = new System.Drawing.Point(128, 5);
			this.btnClose.Name = "btnClose";
			this.btnClose.Size = new System.Drawing.Size(64, 19);
			this.btnClose.TabIndex = 12;
			this.btnClose.Text = "CloseTab";
			this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(248, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(248, 23);
			this.label1.TabIndex = 19;
			this.label1.Text = "add \'js=js;\' at end line,then can see js various.";
			// 
			// PageJSDebug
			// 
			this.Controls.Add(this.label1);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnRun);
			this.Controls.Add(this.btnNew);
			this.Controls.Add(this.btnClose);
			this.Name = "PageJSDebug";
			this.Size = new System.Drawing.Size(600, 272);
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			if(btnCancel.Text.Equals("Cancel"))
			{
				btnCancel.Text = "Cancel...";
			}
		}

		private void btnRuns_Click(object sender, System.EventArgs e)
		{
			btnCancel.Enabled = true;
			btnRun.Enabled = false;
			JSTestRun_main();
			btnCancel.Enabled = false;
			btnRun.Enabled = true;
			btnCancel.Text = "Cancel";
		}

		//for calculate time
		DateTime MainTime = System.DateTime.Now;
		private void JSTestRun_main()
		{
			//Start do something
			if(MessageBox.Show("Start RunJscript?", "Msg...", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question) != DialogResult.Yes)
			{
				return;
			}
			//create new IPara
			IPara ipara = frmMain.CreateIPara();
			//set filed value
			ipara = frmMain.CreateSampleFiled(ipara);
			ipara = ClassExt.TreateFieldANDJScript(ipara);

			System.Text.StringBuilder sbJSTxt = (System.Text.StringBuilder)ipara.SystemVarious["StringBuilderJScriptTxt"];
			string sJSTxtHead = (string)ipara.SystemVarious["JScriptTxtHEAD"];
			string sJSTxtEnd = (string)ipara.SystemVarious["JScriptTxtEND"];

			jspage = (PageJSDebugPage)tabControl1.SelectedTab.Controls[0];
			frmMain.msg.clear();
			frmMain.msg.Focus();
			frmMain.labStatus.Text = "RunJscript.";
			MainTime = System.DateTime.Now;
			frmMain.msg.println("Start:" + MainTime);
			string ErrorString = "";
			for(int i = 0; i < jspage.numRunCnt.Value; i++)
			{
				if((int)(i/100) == i/100)
				{
					frmMain.labStatus.Text ="RunCount:" + i;
				}
				string sJS;
				//not check for include define JS()
				if(jspage.chkAddJS.Checked)
				{
					sJS = sbJSTxt.ToString() + "\r\n" + sJSTxtHead + "\r\n" + jspage.txtJS.Text + "\r\n" + sJSTxtEnd;
				}
				else
				{
					sJS = jspage.txtJS.Text;
				}

				bool isReset = jspage.chkReset.Checked;
				object obj = cc.Eval.JSEvaluateToObject(sJS, isReset, out ErrorString);
				Application.DoEvents();
				if(btnCancel.Text.Equals("Cancel..."))
				{
					frmMain.msg.println("User Cancel.");
					return;
				}
				if(i == 0)
				{
					if(ErrorString != null)
					{
						frmMain.msg.println(ErrorString, Color.Red);
						break;
					}
					if(obj == null)
					{
						frmMain.msg.println("Run JScript error, no value return.");
						break;
					}
					frmMain.msg.println("Return Type:" + obj.GetType());
					if(!obj.GetType().FullName.Equals("Microsoft.JScript.JSObject"))
					{
						frmMain.msg.println("Return Value:" + obj);
					}
					else
					{
						GetScriptObjectValue((Microsoft.JScript.ScriptObject)obj, "");
					}
				}
			}
			frmMain.labStatus.Text ="RunCount:" + jspage.numRunCnt.Value;
			frmMain.msg.println("End:" + System.DateTime.Now + "(elapsed:" + (int)((System.DateTime.Now - MainTime).TotalMilliseconds/1000) + " Seconds)");
		}

		private void GetScriptObjectValue(Microsoft.JScript.ScriptObject jsobj, string sSpace)
		{
			try
			{
				sSpace += "  ";
				System.Reflection.FieldInfo[] jsFields = jsobj.GetFields(System.Reflection.BindingFlags.GetField);
				foreach(System.Reflection.FieldInfo jsField in jsFields)
				{
					frmMain.msg.println(sSpace + "jsField Type:" +  jsField.GetType());
					frmMain.msg.println(sSpace + "jsField Name:" +  jsField.Name);
					object obj = jsField.GetValue(jsobj);
					if(!obj.GetType().FullName.Equals("Microsoft.JScript.JSObject"))
					{
						frmMain.msg.println(sSpace + "Return Value:" + obj);
					}
					else
					{
						GetScriptObjectValue((Microsoft.JScript.ScriptObject)obj, sSpace);
					}
				}
			}
			catch(Exception exp)
			{
				frmMain.msg.println(sSpace + "Error:" + exp.Message, Color.Red);
			}
		}

		public void Config_Load(System.Collections.Specialized.NameValueCollection coll)
		{
			for(int i = 0; i < 1000; i++)
			{
				string JSnumRunCnt = coll.Get("JSnumRunCnt" + i);
				string JStxtText = coll.Get("JStxtText" + i);
				string JSchkReset = coll.Get("JSchkReset" + i);
				string JSchkAddJS = coll.Get("JSchkAddJS" + i);
				if(JSnumRunCnt == null || JStxtText == null
					|| JSchkReset == null || JSchkAddJS == null)
				{
					break;
				}
				if(i != 0)
				{
					btnNew_Click(null, null);
				}
				jspage = (PageJSDebugPage)tabControl1.TabPages[i].Controls[0];
				try
				{
					jspage.numRunCnt.Value = int.Parse(JSnumRunCnt);
				}
				catch
				{
				}
				jspage.txtJS.Text = JStxtText.Replace("[#NEW_LINE#]", "\r\n");
				jspage.chkReset.Checked = JSchkReset.Equals("1");
				jspage.chkAddJS.Checked = JSchkAddJS.Equals("1");
			}
			try
			{
				tabControl1.SelectedIndex = int.Parse(coll.Get("JSDebugSelectedIndex"));
			}
			catch
			{
			}
		}

		public void Config_Save(System.IO.StreamWriter sw)
		{
			sw.Write("#JSDebug setting\r\n");
			sw.Write("JSDebugSelectedIndex={0}\r\n", tabControl1.SelectedIndex);
			for(int i = 0; i < tabControl1.Controls.Count; i++)
			{
				jspage = (PageJSDebugPage)tabControl1.TabPages[i].Controls[0];
				string JStxtText = jspage.txtJS.Text.Replace("\r", "").Replace("\n", "[#NEW_LINE#]");
				sw.Write("JSchkReset{0}={1}\r\n", i, jspage.chkReset.Checked ? "1" : "0");
				sw.Write("JSnumRunCnt{0}={1}\r\n", i, jspage.numRunCnt.Value);
				sw.Write("JSchkAddJS{0}={1}\r\n", i, jspage.chkAddJS.Checked ? "1" : "0");
				sw.Write("JStxtText{0}={1}\r\n", i, JStxtText);
			}
		}

		int nTabPageCnt = 0;
		private void btnNew_Click(object sender, System.EventArgs e)
		{
			jspage = new PageJSDebugPage();
			tabControl1.TabPages.Add(new TabPage());
			int nTabPageInd = tabControl1.TabPages.Count - 1;
			tabControl1.TabPages[nTabPageInd].Controls.Add(jspage);
			jspage.Height = tabControl1.TabPages[nTabPageInd].Height;
			jspage.Width = tabControl1.TabPages[nTabPageInd].Width;
			jspage.Anchor = ((System.Windows.Forms.AnchorStyles)((((
				System.Windows.Forms.AnchorStyles.Top 
				| System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			nTabPageCnt++;
			tabControl1.TabPages[nTabPageInd].Text = "" + nTabPageCnt;
		}

		private void btnClose_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			jspage = (PageJSDebugPage)tabControl1.SelectedTab.Controls[0];
			if(tabControl1.SelectedIndex == 0)
			{
				if(!jspage.txtJS.Text.Trim().Equals(""))
				{
					//clear it?
					if(MessageBox.Show("Can not close this TabPage, Clear it?", "Msg...", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) != DialogResult.Yes)
					{
						return;
					}
					jspage.txtJS.Text = "";
				}
				else
				{
					MessageBox.Show("Can not close this TabPage.", "Msg...", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				frmMain.labStatus.Text = "Can not close this TabPage";
				return;
			}
			if(!jspage.txtJS.Text.Trim().Equals(""))
			{
				//close it?
				if(MessageBox.Show("Close this TabPage?", "Msg...", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question) != DialogResult.Yes)
				{
					return;
				}
			}
			tabControl1.Controls.Remove(tabControl1.SelectedTab);
		}

	}
}
