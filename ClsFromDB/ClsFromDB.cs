using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Xml;
using System.Data.SqlClient;

namespace ClsFromDB
{
	/// <summary>
	/// Form1 の概要の説明です。
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private static string TAB_PAGE_HEAD = "●　";
		//ext data that saved to TabPage's Tag
		public class TabPageDatas
		{
			private Hashtable tbl = new Hashtable();
			public int ActiveIndex = -1;
			public class TabPageData
			{
				public int msgHeight = 0;
				public bool isVisible = false;
			}

			/// <summary>
			/// get/set value of one page
			/// </summary>
			public TabPageData this[int nIndex]
			{
				get
				{
					if(tbl[nIndex] == null)
					{
						tbl[nIndex] = new TabPageData();
					}
					return (TabPageData)tbl[nIndex];
				}
				set
				{
					tbl[nIndex] = value;
				}
			}
		}
		private TabPageDatas tabData = new TabPageDatas();
		public cc.Msg msg;

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabMain;
		private System.Windows.Forms.Button btnExit;
		private System.Windows.Forms.Button btnCreateCls;
		private System.Windows.Forms.RichTextBox txtMsg;
		private System.Windows.Forms.TextBox txtSaveTo;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button btnSaveToDir;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.RadioButton radFromExcel;
		private System.Windows.Forms.RadioButton radFromServer;
		private System.Windows.Forms.Button btnMsgHide;
		private System.Windows.Forms.Button btnMsgClear;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Splitter splitter1;
		private System.Windows.Forms.TabPage FromExcel;
		private System.Windows.Forms.TabPage FromServer;
		private System.Windows.Forms.TabPage ConfigTemplate;
		public System.Windows.Forms.Label labStatus;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.TabPage setVarious;
		private System.Windows.Forms.TabPage JSDebug;
		private ClsFromDB.PageFromExcel frmExcel;
		private ClsFromDB.PageFromServer frmServer;
		private ClsFromDB.PageConfigTemplate frmTemplate;
		private ClsFromDB.PageVarious frmVarious;
		private ClsFromDB.PageJSDebug frmJSDebug;

		public Form1()
		{
			//
			// Windows フォーム デザイナ サポートに必要です。
			//
			InitializeComponent();
		}

		/// <summary>
		/// 使用されているリソースに後処理を実行します。
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows フォーム デザイナで生成されたコード 
		/// <summary>
		/// デザイナ サポートに必要なメソッドです。このメソッドの内容を
		/// コード エディタで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabMain = new System.Windows.Forms.TabPage();
			this.btnExit = new System.Windows.Forms.Button();
			this.btnCreateCls = new System.Windows.Forms.Button();
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.radFromExcel = new System.Windows.Forms.RadioButton();
			this.txtSaveTo = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.btnSaveToDir = new System.Windows.Forms.Button();
			this.radFromServer = new System.Windows.Forms.RadioButton();
			this.label9 = new System.Windows.Forms.Label();
			this.FromExcel = new System.Windows.Forms.TabPage();
			this.frmExcel = new ClsFromDB.PageFromExcel();
			this.FromServer = new System.Windows.Forms.TabPage();
			this.frmServer = new ClsFromDB.PageFromServer();
			this.ConfigTemplate = new System.Windows.Forms.TabPage();
			this.frmTemplate = new ClsFromDB.PageConfigTemplate();
			this.setVarious = new System.Windows.Forms.TabPage();
			this.frmVarious = new ClsFromDB.PageVarious();
			this.JSDebug = new System.Windows.Forms.TabPage();
			this.frmJSDebug = new ClsFromDB.PageJSDebug();
			this.txtMsg = new System.Windows.Forms.RichTextBox();
			this.labStatus = new System.Windows.Forms.Label();
			this.btnMsgHide = new System.Windows.Forms.Button();
			this.btnMsgClear = new System.Windows.Forms.Button();
			this.panel1 = new System.Windows.Forms.Panel();
			this.splitter1 = new System.Windows.Forms.Splitter();
			this.tabControl1.SuspendLayout();
			this.tabMain.SuspendLayout();
			this.FromExcel.SuspendLayout();
			this.FromServer.SuspendLayout();
			this.ConfigTemplate.SuspendLayout();
			this.setVarious.SuspendLayout();
			this.JSDebug.SuspendLayout();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabMain);
			this.tabControl1.Controls.Add(this.FromExcel);
			this.tabControl1.Controls.Add(this.FromServer);
			this.tabControl1.Controls.Add(this.ConfigTemplate);
			this.tabControl1.Controls.Add(this.setVarious);
			this.tabControl1.Controls.Add(this.JSDebug);
			this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.tabControl1.Location = new System.Drawing.Point(0, 0);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(584, 298);
			this.tabControl1.TabIndex = 0;
			this.tabControl1.SelectedIndexChanged += new System.EventHandler(this.tabControl1_SelectedIndexChanged);
			// 
			// tabMain
			// 
			this.tabMain.Controls.Add(this.btnExit);
			this.tabMain.Controls.Add(this.btnCreateCls);
			this.tabMain.Controls.Add(this.textBox1);
			this.tabMain.Controls.Add(this.radFromExcel);
			this.tabMain.Controls.Add(this.txtSaveTo);
			this.tabMain.Controls.Add(this.label4);
			this.tabMain.Controls.Add(this.btnSaveToDir);
			this.tabMain.Controls.Add(this.radFromServer);
			this.tabMain.Controls.Add(this.label9);
			this.tabMain.Location = new System.Drawing.Point(4, 21);
			this.tabMain.Name = "tabMain";
			this.tabMain.Size = new System.Drawing.Size(576, 273);
			this.tabMain.TabIndex = 0;
			this.tabMain.Text = "Main";
			// 
			// btnExit
			// 
			this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnExit.Location = new System.Drawing.Point(424, 40);
			this.btnExit.Name = "btnExit";
			this.btnExit.Size = new System.Drawing.Size(120, 24);
			this.btnExit.TabIndex = 2;
			this.btnExit.Text = "Exit";
			this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
			// 
			// btnCreateCls
			// 
			this.btnCreateCls.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnCreateCls.Location = new System.Drawing.Point(424, 8);
			this.btnCreateCls.Name = "btnCreateCls";
			this.btnCreateCls.Size = new System.Drawing.Size(120, 23);
			this.btnCreateCls.TabIndex = 1;
			this.btnCreateCls.Text = "CreateClass";
			this.btnCreateCls.Click += new System.EventHandler(this.btnCreateCls_Click);
			// 
			// textBox1
			// 
			this.textBox1.Anchor = System.Windows.Forms.AnchorStyles.Top;
			this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.textBox1.Font = new System.Drawing.Font("MS UI Gothic", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.textBox1.Location = new System.Drawing.Point(64, 72);
			this.textBox1.Multiline = true;
			this.textBox1.Name = "textBox1";
			this.textBox1.ReadOnly = true;
			this.textBox1.Size = new System.Drawing.Size(448, 120);
			this.textBox1.TabIndex = 14;
			this.textBox1.Text = "クラス作成補充ツール：\r\n\r\n１、Excel、またDBServerから、Templateにで定義されたクラスを作成\r\n２、Templateの定義によって、Java" +
				"、C#、Jsp、SQLの作成も可能です\r\n※Oracleより取得はまだです。SqlServerより一部パラメータが取得していない。";
			// 
			// radFromExcel
			// 
			this.radFromExcel.Checked = true;
			this.radFromExcel.Location = new System.Drawing.Point(80, 16);
			this.radFromExcel.Name = "radFromExcel";
			this.radFromExcel.Size = new System.Drawing.Size(80, 24);
			this.radFromExcel.TabIndex = 12;
			this.radFromExcel.TabStop = true;
			this.radFromExcel.Text = "FromExcel";
			// 
			// txtSaveTo
			// 
			this.txtSaveTo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSaveTo.Location = new System.Drawing.Point(80, 40);
			this.txtSaveTo.Name = "txtSaveTo";
			this.txtSaveTo.Size = new System.Drawing.Size(312, 19);
			this.txtSaveTo.TabIndex = 9;
			this.txtSaveTo.Text = "c:\\IFClass\\";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 42);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 8;
			this.label4.Text = "Save to:";
			// 
			// btnSaveToDir
			// 
			this.btnSaveToDir.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSaveToDir.Font = new System.Drawing.Font("MS UI Gothic", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSaveToDir.Location = new System.Drawing.Point(392, 40);
			this.btnSaveToDir.Name = "btnSaveToDir";
			this.btnSaveToDir.Size = new System.Drawing.Size(16, 19);
			this.btnSaveToDir.TabIndex = 10;
			this.btnSaveToDir.Text = "...";
			this.btnSaveToDir.Click += new System.EventHandler(this.btnSaveToDir_Click);
			// 
			// radFromServer
			// 
			this.radFromServer.Location = new System.Drawing.Point(160, 16);
			this.radFromServer.Name = "radFromServer";
			this.radFromServer.Size = new System.Drawing.Size(96, 24);
			this.radFromServer.TabIndex = 13;
			this.radFromServer.Text = "FromServer";
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(8, 21);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 8;
			this.label9.Text = "TableLayout:";
			// 
			// FromExcel
			// 
			this.FromExcel.Controls.Add(this.frmExcel);
			this.FromExcel.Location = new System.Drawing.Point(4, 21);
			this.FromExcel.Name = "FromExcel";
			this.FromExcel.Size = new System.Drawing.Size(576, 273);
			this.FromExcel.TabIndex = 5;
			this.FromExcel.Tag = "";
			this.FromExcel.Text = "FromExcel";
			// 
			// frmExcel
			// 
			this.frmExcel.Dock = System.Windows.Forms.DockStyle.Fill;
			this.frmExcel.Location = new System.Drawing.Point(0, 0);
			this.frmExcel.Name = "frmExcel";
			this.frmExcel.Size = new System.Drawing.Size(576, 273);
			this.frmExcel.TabIndex = 0;
			// 
			// FromServer
			// 
			this.FromServer.Controls.Add(this.frmServer);
			this.FromServer.Location = new System.Drawing.Point(4, 21);
			this.FromServer.Name = "FromServer";
			this.FromServer.Size = new System.Drawing.Size(576, 273);
			this.FromServer.TabIndex = 6;
			this.FromServer.Tag = "";
			this.FromServer.Text = "FromServer";
			// 
			// frmServer
			// 
			this.frmServer.Dock = System.Windows.Forms.DockStyle.Fill;
			this.frmServer.Location = new System.Drawing.Point(0, 0);
			this.frmServer.Name = "frmServer";
			this.frmServer.Size = new System.Drawing.Size(576, 273);
			this.frmServer.TabIndex = 0;
			// 
			// ConfigTemplate
			// 
			this.ConfigTemplate.AutoScroll = true;
			this.ConfigTemplate.Controls.Add(this.frmTemplate);
			this.ConfigTemplate.Location = new System.Drawing.Point(4, 21);
			this.ConfigTemplate.Name = "ConfigTemplate";
			this.ConfigTemplate.Size = new System.Drawing.Size(576, 273);
			this.ConfigTemplate.TabIndex = 4;
			this.ConfigTemplate.Tag = "";
			this.ConfigTemplate.Text = "ConfigTemplate";
			// 
			// frmTemplate
			// 
			this.frmTemplate.Dock = System.Windows.Forms.DockStyle.Fill;
			this.frmTemplate.Location = new System.Drawing.Point(0, 0);
			this.frmTemplate.Name = "frmTemplate";
			this.frmTemplate.Size = new System.Drawing.Size(576, 273);
			this.frmTemplate.TabIndex = 0;
			// 
			// setVarious
			// 
			this.setVarious.Controls.Add(this.frmVarious);
			this.setVarious.Location = new System.Drawing.Point(4, 21);
			this.setVarious.Name = "setVarious";
			this.setVarious.Size = new System.Drawing.Size(576, 273);
			this.setVarious.TabIndex = 7;
			this.setVarious.Tag = "";
			this.setVarious.Text = "setVarious";
			// 
			// frmVarious
			// 
			this.frmVarious.Dock = System.Windows.Forms.DockStyle.Fill;
			this.frmVarious.Location = new System.Drawing.Point(0, 0);
			this.frmVarious.Name = "frmVarious";
			this.frmVarious.Size = new System.Drawing.Size(576, 273);
			this.frmVarious.TabIndex = 0;
			// 
			// JSDebug
			// 
			this.JSDebug.Controls.Add(this.frmJSDebug);
			this.JSDebug.Location = new System.Drawing.Point(4, 21);
			this.JSDebug.Name = "JSDebug";
			this.JSDebug.Size = new System.Drawing.Size(576, 273);
			this.JSDebug.TabIndex = 9;
			this.JSDebug.Tag = "";
			this.JSDebug.Text = "JScriptDebug";
			// 
			// frmJSDebug
			// 
			this.frmJSDebug.Dock = System.Windows.Forms.DockStyle.Fill;
			this.frmJSDebug.Location = new System.Drawing.Point(0, 0);
			this.frmJSDebug.Name = "frmJSDebug";
			this.frmJSDebug.Size = new System.Drawing.Size(576, 273);
			this.frmJSDebug.TabIndex = 0;
			// 
			// txtMsg
			// 
			this.txtMsg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtMsg.Location = new System.Drawing.Point(1, 0);
			this.txtMsg.Name = "txtMsg";
			this.txtMsg.Size = new System.Drawing.Size(580, 52);
			this.txtMsg.TabIndex = 0;
			this.txtMsg.Text = "";
			this.txtMsg.TextChanged += new System.EventHandler(this.txtMsg_TextChanged);
			// 
			// labStatus
			// 
			this.labStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.labStatus.Location = new System.Drawing.Point(0, 56);
			this.labStatus.Name = "labStatus";
			this.labStatus.Size = new System.Drawing.Size(448, 14);
			this.labStatus.TabIndex = 1;
			this.labStatus.Text = "Status:";
			// 
			// btnMsgHide
			// 
			this.btnMsgHide.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnMsgHide.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnMsgHide.Location = new System.Drawing.Point(520, 54);
			this.btnMsgHide.Name = "btnMsgHide";
			this.btnMsgHide.Size = new System.Drawing.Size(64, 18);
			this.btnMsgHide.TabIndex = 11;
			this.btnMsgHide.Text = "Hide Log";
			this.btnMsgHide.Click += new System.EventHandler(this.btnMsgHide_Click);
			// 
			// btnMsgClear
			// 
			this.btnMsgClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.btnMsgClear.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnMsgClear.Location = new System.Drawing.Point(456, 54);
			this.btnMsgClear.Name = "btnMsgClear";
			this.btnMsgClear.Size = new System.Drawing.Size(64, 18);
			this.btnMsgClear.TabIndex = 11;
			this.btnMsgClear.Text = "Clear Log";
			this.btnMsgClear.Click += new System.EventHandler(this.btnMsgClear_Click);
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.labStatus);
			this.panel1.Controls.Add(this.btnMsgClear);
			this.panel1.Controls.Add(this.btnMsgHide);
			this.panel1.Controls.Add(this.txtMsg);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel1.Location = new System.Drawing.Point(0, 301);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(584, 72);
			this.panel1.TabIndex = 13;
			// 
			// splitter1
			// 
			this.splitter1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.splitter1.Location = new System.Drawing.Point(0, 298);
			this.splitter1.Name = "splitter1";
			this.splitter1.Size = new System.Drawing.Size(584, 3);
			this.splitter1.TabIndex = 14;
			this.splitter1.TabStop = false;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 12);
			this.ClientSize = new System.Drawing.Size(584, 373);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.splitter1);
			this.Controls.Add(this.panel1);
			this.Name = "Form1";
			this.Text = "ClsFromDB";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Form1_Closing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.tabControl1.ResumeLayout(false);
			this.tabMain.ResumeLayout(false);
			this.FromExcel.ResumeLayout(false);
			this.FromServer.ResumeLayout(false);
			this.ConfigTemplate.ResumeLayout(false);
			this.setVarious.ResumeLayout(false);
			this.JSDebug.ResumeLayout(false);
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// アプリケーションのメイン エントリ ポイントです。
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			System.Reflection.Assembly ass = System.Reflection.Assembly.GetExecutingAssembly();
			string sver = System.Diagnostics.FileVersionInfo.GetVersionInfo(ass.Location).FileVersion;
			this.Text = this.Text + " (Version:" + sver + ") - NewCon.ShuKK";
		
			msg = new cc.Msg(txtMsg);

			//all sub pages, for set frmMain
			for(int i = 0; i < tabControl1.TabPages.Count; i++)
			{
				if(tabControl1.TabPages[i].Controls.Count > 0)
				{
					Control userCtl = tabControl1.TabPages[i].Controls[0];
					System.Type ctlType = userCtl.GetType();
					if(ctlType.BaseType != null && ctlType.BaseType.Name.Equals("UserControl"))
					{
						System.Reflection.FieldInfo field = ctlType.GetField("frmMain");
						if(field != null)
						{
							field.SetValue(userCtl, this);
						}
					}
				}
			}

			Config_Load();
			tabControl1_SelectedIndexChanged(sender, e);
		}

		private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if(!btnExit.Text.Equals("Exit") && !btnExit.Text.Equals("Exit...OK"))
			{
				btnExit.Text = "Exit...";
				e.Cancel = true;
			}
			else
			{
				Config_Save();
			}
		}

		private void btnExit_Click(object sender, System.EventArgs e)
		{
			if(btnExit.Text.Equals("Cancel"))
			{
				btnExit.Text = "Cancel...";
			}
			if(btnExit.Text.Equals("Exit") || btnExit.Text.Equals("Exit...OK"))
			{
				Close();
			}
		}

		private void btnCreateCls_Click(object sender, System.EventArgs e)
		{
			btnCreateCls.Enabled = false;
			btnExit.Text = "Cancel";

			try
			{
				createcls_main();
			}
			catch(cc.AppException exp)
			{
				msg.println("Error when create class:" + exp.MessageAll);
			}
			catch(Exception exp)
			{
				msg.println("Error when create class:" + exp.Message);
			}

			if(isMainStart)
			{
				//if output Start,then out "end time"
				msg.println("End:" + System.DateTime.Now + "(elapsed:" + (int)((System.DateTime.Now - MainTime).TotalMilliseconds/1000) + " Seconds)");
				isMainStart = false;
			}
			if(btnExit.Text.Equals("Exit..."))
			{
				btnExit.Text = "Exit...OK";
				Close();
			}
			btnCreateCls.Enabled = true;
			btnExit.Text = "Exit";
		}

		//for calculate time
		DateTime MainTime = System.DateTime.Now;
		bool isMainStart = false;
		private void createcls_main()
		{
			//not for frmipara == null
			if(frmExcel == null || frmServer == null || frmTemplate == null || 
				frmVarious == null || frmJSDebug == null)
			{
				msg.println("Please config Excel,SqlServer,Template and Various information first.", Color.Red);
				return;
			}

			//check template
			if(frmTemplate.lstTemp.CheckedItems.Count < 1)
			{
				msg.println("no selected template to create class.");
				return;
			}
			XmlDocument xmldoc = frmTemplate.getXMLTemplate();
			if(xmldoc == null)
			{
				//if not get or config/template.count < 1,error msg is out at getXMLTemplate
				return;
			}
			XmlNodeList nodeList = xmldoc.SelectNodes("config/template");

			string soutpath = txtSaveTo.Text.Trim();
			if(soutpath.Equals(""))
			{
				msg.println("need input out directory.");
				return;
			}
			if(!soutpath.EndsWith("\\"))
			{
				soutpath += "\\";
			}

			//Start do samething
			if(MessageBox.Show("Start create class from selected tables?", "Msg...", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question) != DialogResult.Yes)
			{
				return;
			}
			msg.clear();
			msg.Focus();
			labStatus.Text = "create class.";
			isMainStart = true;
			MainTime = System.DateTime.Now;
			msg.println("Start:" + MainTime);

			if(!Directory.Exists(soutpath))
			{
				try
				{
					Directory.CreateDirectory(soutpath);
				}
				catch
				{
				}
				if(!Directory.Exists(soutpath))
				{
					msg.println("can not create dir:" + soutpath);
					return;
				}
			}
			Application.DoEvents();

			//Interface of ipara
			IPara ipara = CreateIPara();
			ipara.OutPath = soutpath;
			ipara.TemplateCount = nodeList.Count;

			if(radFromExcel.Checked)
			{
				//check excel template
				string sFilePath = frmExcel.txtExcelTemp.Text;
				if(sFilePath.StartsWith(".\\") || sFilePath.StartsWith("./"))
				{
					sFilePath = Application.StartupPath + sFilePath.Substring(1);
				}
				if(!File.Exists(sFilePath))
				{
					msg.println("selected excel file(table layout) is not exist.");
					return;
				}
				if(frmExcel.lstExcelTemp.CheckedItems.Count != 1)
				{
					msg.println("please select the sheet(only one) of excel file with table layout defined.");
					return;
				}

				//check excel file
				if(frmExcel.lstExcel.CheckedItems.Count < 1)
				{
					msg.println("please select the excel file need treate.");
					return;
				}

				//check if the sheets in ExcelTemp File is same to list
				string sConnExcel = cc.OleDB.ConnStringExcel(sFilePath, false);
				cc.OleDB dbExcel = new cc.OleDB(sConnExcel);
				if(dbExcel.Error())
				{
					msg.println("can not get excel:\r\n  " + sFilePath + "\r\n" + dbExcel.Exception.Message, Color.Red);
					dbExcel.Dispose();
					return;
				}
				string[] sSheetsName = dbExcel.GetExcelSheetsName();
				bool isExcelTempListOK = false;
				if(sSheetsName != null && frmExcel.lstExcelTemp.Items.Count == sSheetsName.Length)
				{
					isExcelTempListOK = true;
					for(int i = 0; i < frmExcel.lstExcelTemp.Items.Count; i++)
					{
						if(!frmExcel.lstExcelTemp.Items[i].ToString().Equals(sSheetsName[i]))
						{
							isExcelTempListOK = false;
							break;
						}
					}
				}
				if(!isExcelTempListOK)
				{
					msg.println("the sheets in Excel Template File is not same to list,need refresh list:\r\n" + sFilePath);
					dbExcel.Dispose();
					return;
				}

				//get excel template info
				string sSheetName = "";
				for(int i = 0; i < frmExcel.lstExcelTemp.Items.Count; i++)
				{
					if(frmExcel.lstExcelTemp.GetItemChecked(i))
					{
						sSheetName = frmExcel.lstExcelTemp.Items[i].ToString();
						break;
					}
				}
				System.Data.DataTable tblExcelTemp = dbExcel.GetExcelSheet(sSheetName);
				if(tblExcelTemp == null)
				{
					msg.println("the sheets in Excel Template File can not get:" + sSheetName);
					dbExcel.Dispose();
					return;
				}
				HashXY hashxy = ClassExt.GetExcelTempInfo(tblExcelTemp);
				tblExcelTemp.Dispose();
				dbExcel.Dispose();

				if(hashxy == null || hashxy["[#TABLE_NAME#]"] == null || hashxy["[#FIELD_START_Y#]"] == null)
				{
					msg.println("the selected Excel Template Sheet is not include enough info.");
					msg.println("at least this is need:");
					msg.println("  [#TABLE_NAME#]],[#FIELD_NAME#],[#FIELD_TYPE#],[#FIELD_INGETER#],[#FIELD_DECIMAL#]");
					msg.println("also you can define like:");
					msg.println("  [#FIELD_TYPE_INGETER_DECIMAL#]],[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#],[#FIELD_INGETER_DECIMAL#],[#FIELD_INGETERDECIMAL_DECIMAL#]");
					return;
				}
				if(hashxy["[#FIELD_START_Y#]"].Y < 0)
				{
					msg.println("the filed info like [#FIELD_NAME#],[#FIELD_TYPE#],and other [#FIELD_VARIOUS#]... should begin at the same line.");
					return;
				}

				msg.println("Create file via Template from table information.", Color.Blue);
				cc.OleDB dbExcelF = null; //for open each excel file
				System.Data.DataTable tblSheet = null; //for each sheet of excel

				//get table count
				int nTableCount = 0;
				sFilePath = "";
				for(int i = 0; i < frmExcel.lstExcel.Items.Count; i++)
				{
					sFilePath = frmExcel.lstExcel.Items[i].ToString();
					if(sFilePath.Substring(1, 2).Equals(":\\"))
					{
						if(!File.Exists(sFilePath))
						{
							sFilePath = "";
						}
						continue;
					}
					if(frmExcel.lstExcel.GetItemChecked(i) && !sFilePath.Equals(""))
					{
						nTableCount++;
					}
				}

				//need treated table count
				ipara.TableCount = nTableCount;
				int nTableIndex = 0;
				for(int i = 0; i < frmExcel.lstExcel.Items.Count; i++)
				{
					sSheetName = frmExcel.lstExcel.Items[i].ToString();
					/* the structure of frmExcel.lstExcel:
					 * D:\\Table\\Table1.xls
					 * 　　Sheet1
					 * 　　Sheet2
					 * 　　...
					*/
					//now is Excel File,so set null to it,let next checked sheet to open it
					if(sSheetName.Substring(1, 2).Equals(":\\"))
					{
						sFilePath = sSheetName;
						if(dbExcelF != null)
						{
							dbExcelF.Dispose();
							dbExcelF = null;
						}
						continue;
					}
					//then treate each sheet of the Excel file,remove "　　" of head
					sSheetName = sSheetName.Substring(2);
					if(frmExcel.lstExcel.GetItemChecked(i))
					{
						if(dbExcelF == null)
						{
							sConnExcel = cc.OleDB.ConnStringExcel(sFilePath, false);
							dbExcelF = new cc.OleDB(sConnExcel);
							if(dbExcelF.Error())
							{
								msg.println("Open Excel File Error:");
								msg.println("  " + sFilePath);
								msg.println(dbExcelF.Exception.Message, Color.Red);
								continue;
							}
						}
						if(dbExcelF == null || dbExcelF.Error())
						{
							msg.println("  Skip:" + sSheetName);
							continue;
						}
						if(tblSheet != null)
						{
							tblSheet.Dispose();
							tblSheet = null;
						}
						tblSheet = dbExcelF.GetExcelSheet(sSheetName);
						if(dbExcelF.Error() || tblSheet == null)
						{
							msg.println("  get data error,Skip:" + sSheetName);
							continue;
						}

						//current treated table index
						ipara.TableIndex = nTableIndex;
						nTableIndex++;
						labStatus.Text = "Complete:" + nTableIndex + "/" + nTableCount;
						string smsg = ClassExt.GetTBLInfoFromExcel(hashxy, tblSheet, ipara);
						if(smsg != null)
						{
							msg.println(smsg);
							continue;
						}

						smsg = "Excel:" + Path.GetFileName(sFilePath)
							+ ", Sheet:" + sSheetName + ", Table:"
							+ ipara.UserVarious["[#TABLE_NAME#]"];
						msg.println(smsg);
						//output with this table info
						if(!createcls_main_onetable(nodeList, ipara))
						{
							if(dbExcelF != null)
							{
								dbExcelF.Dispose();
								dbExcelF = null;
							}
							return;
						}
					}
				}
				//close last excel file
				if(dbExcelF != null)
				{
					dbExcelF.Dispose();
					dbExcelF = null;
				}
			}
			else
			{
				if(frmServer.lstDBTBL.Items.Count < 1)
				{
					msg.println("no from tables to create class.");
					return;
				}

				//conn server
				cc.DB cdb = frmServer.ConnDB();
				if(cdb == null)
				{
					msg.println("SQLServerに接続...できませんでした。", Color.Red);
					return;
				}

				msg.println("Create file via Template from table information.", Color.Blue);
				//need treated table count
				ipara.TableCount = frmServer.lstDBTBL.Items.Count;
				for(int loopi = 0; loopi < frmServer.lstDBTBL.Items.Count; loopi++)
				{
					labStatus.Text = "Complete:" + (loopi + 1) + "/" + frmServer.lstDBTBL.Items.Count;

					int npos = frmServer.lstDBTBL.Items[loopi].ToString().IndexOf(" . ");
					string clsdb = frmServer.lstDBTBL.Items[loopi].ToString().Substring(0, npos);
					string clstbl = frmServer.lstDBTBL.Items[loopi].ToString().Substring(npos + 3);

					ipara.UserVarious["[#DB_NAME#]"] = clsdb;
					ipara.UserVarious["[#TABLE_NAME#]"] = clstbl;
					//current treated table index
					ipara.TableIndex = loopi;
					msg.println(clsdb + " . " + clstbl + ":");
					string smsg = ClassExt.GetTBLInfoFromServer(cdb, ipara);
					if(smsg != null)
					{
						msg.println(smsg);
						continue;
					}

					smsg = "DataBase:" + clsdb + ", Table:" + clstbl;
					msg.println(smsg);
					//output with this table info
					if(!createcls_main_onetable(nodeList, ipara))
					{
						if(cdb != null)
						{
							cdb.Dispose();
						}
						return;
					}
				}

				if(cdb != null)
				{
					cdb.Dispose();
				}
			}

		}

		private bool createcls_main_onetable(XmlNodeList nodeList, IPara ipara)
		{
			//use tableinfo,treat every XML.Template and output it
			for(int loopj = 0; loopj < nodeList.Count; loopj++)
			{
				//if user want cancel?
				Application.DoEvents();
				if(!btnExit.Text.Equals("Cancel"))
				{
					if(MessageBox.Show("Cancel?", "Msg...", MessageBoxButtons.YesNoCancel,MessageBoxIcon.Question) == DialogResult.Yes)
					{
						msg.println("User Cancel.");
						return false;
					}
					else
					{
						btnExit.Text = "Cancel";
					}
				}

				ipara.TemplateIndex = loopj;
				ipara.TemplateNode = nodeList[loopj];
				string smsg = ClassExt.CreateClsFromTemp(ipara);
				if(smsg != null)
				{
					msg.println(smsg);
					continue;
				}
			}
			return true;
		}

		public IPara CreateIPara()
		{
			IPara ipara = new IPara();
			ipara.msg = new cc.Msg(this.txtMsg);

			//add user defined various & add user defined typies
			ipara.tblUserType = frmVarious.UserTypeContrast();

			//get user defined various
			ipara.UserVarious = frmVarious.UserVarious();

			//get command JScript that add to head or end of JScript
			string sJSTxtHead = frmTemplate.getJScriptHead();
			if(sJSTxtHead == null)
			{
				//if checked and read file error,print it
				msg.println("Read JScriptFile1 error", Color.Red);
				sJSTxtHead = "";
			}
			string sJSTxtEnd = frmTemplate.getJScriptEnd();
			if(sJSTxtEnd == null)
			{
				//if checked and read file error,print it
				msg.println("Read JScriptFile2 error", Color.Red);
				sJSTxtEnd = "";
			}
			ipara.SystemVarious["JScriptTxtHEAD"] = sJSTxtHead;
			ipara.SystemVarious["JScriptTxtEND"] = sJSTxtEnd;

			return ipara;
		}

		public IPara CreateSampleFiled(IPara ipara)
		{
			//need various
			ipara.UserVarious["[#DB_NAME#]"] = "SampleDB";
			ipara.UserVarious["[#TABLE_NAME#]"] = "SampleTBL";
			ipara.UserVarious["[#TABLE_COMMENT#]"] = "サンプルテーブル";
			ipara.UserVarious["[#TEMPLATE_SUBDIR#]"] = "";
			ipara.UserVarious["[#TEMPLATE_FILENAME#]"] = "c:\temp.txt";
			ipara.UserVarious["[#TEMPLATE_LANGUAGE#]"] = "java";

			//ipara.tblField.AddNewRow
			ipara.tblField.Columns.Add("FIELD_COMMENT");
			ipara.tblField.Columns.Add("FIELD_PK");
			ipara.tblField.Columns.Add("FIELD_NULL");
			DataRow curRow = ipara.tblField.NewRow();
			curRow["FIELD_NAME"] = "Field1";
			curRow["FIELD_TYPE_EXCEL"] = "char"; //get from excel file or server
			curRow["FIELD_INGETER"] = "5";
			curRow["FIELD_DECIMAL"] = "";
			curRow["FIELD_COMMENT"] = "this is Field1";
			curRow["FIELD_PK"] = "PK";
			curRow["FIELD_NULL"] = "NOT NULL";
			ipara.tblField.Rows.Add(curRow);

			curRow = ipara.tblField.NewRow();
			curRow["FIELD_NAME"] = "Field2";
			curRow["FIELD_TYPE_EXCEL"] = "varchar"; //get from excel file or server
			curRow["FIELD_INGETER"] = "3";
			curRow["FIELD_DECIMAL"] = "";
			curRow["FIELD_COMMENT"] = "this is Field2";
			curRow["FIELD_PK"] = "PK";
			curRow["FIELD_NULL"] = "NOT NULL";
			ipara.tblField.Rows.Add(curRow);

			curRow = ipara.tblField.NewRow();
			curRow["FIELD_NAME"] = "Field3";
			curRow["FIELD_TYPE_EXCEL"] = "datetime"; //get from excel file or server
			curRow["FIELD_INGETER"] = "";
			curRow["FIELD_DECIMAL"] = "";
			curRow["FIELD_COMMENT"] = "this is Field3";
			curRow["FIELD_NULL"] = "NOT NULL";
			ipara.tblField.Rows.Add(curRow);

			curRow = ipara.tblField.NewRow();
			curRow["FIELD_NAME"] = "Field4";
			curRow["FIELD_TYPE_EXCEL"] = "number"; //get from excel file or server
			curRow["FIELD_INGETER"] = "6";
			curRow["FIELD_DECIMAL"] = "2";
			curRow["FIELD_COMMENT"] = "this is Field4";
			ipara.tblField.Rows.Add(curRow);
			return ipara;
		}

		private void btnSaveToDir_Click(object sender, System.EventArgs e)
		{
			string sPath = cc.Util.DirSelect("Please Select Folder:", txtSaveTo.Text);
			if(sPath != null)
			{
				txtSaveTo.Text = sPath;
				if(!txtSaveTo.Text.EndsWith("\\"))
				{
					txtSaveTo.Text += "\\";
				}
			}
		}

		private void tabControl1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			int nOldIndex = tabData.ActiveIndex;
			int nNewIndex = tabControl1.SelectedIndex;
			if(nOldIndex == -1)
			{
				tabControl1.SelectedTab.Text = TAB_PAGE_HEAD + tabControl1.SelectedTab.Text;
				tabData.ActiveIndex = nNewIndex;
				if(!tabData[nNewIndex].isVisible)
				{
					btnMsgHide_Click(sender, e);
				}
				else
				{
					int nHeight = tabData[nNewIndex].msgHeight;
					if(nHeight > this.Height - 80)
					{
						nHeight = this.Height - 80;
					}
					panel1.Height = nHeight;
				}
				return;
			}
			if(tabControl1.TabPages[nOldIndex].Text.StartsWith(TAB_PAGE_HEAD))
			{
				tabControl1.TabPages[nOldIndex].Text = 
					tabControl1.TabPages[nOldIndex].Text.Substring(TAB_PAGE_HEAD.Length);
			}
			if(nOldIndex != nNewIndex && txtMsg.Visible)
			{
				tabData[nOldIndex].msgHeight = panel1.Height;
			}
			tabData.ActiveIndex = nNewIndex;
			tabControl1.SelectedTab.Text = TAB_PAGE_HEAD + tabControl1.SelectedTab.Text;

			if(!tabData[nNewIndex].isVisible && !txtMsg.Visible)
			{
				return;
			}
			if((tabData[nNewIndex].isVisible && !txtMsg.Visible)
				|| (!tabData[nNewIndex].isVisible && txtMsg.Visible))
			{
				panel1.Height = tabData[nNewIndex].msgHeight;
				btnMsgHide_Click(sender, e);
			}
			else
			{
				if(tabData[nNewIndex].msgHeight == 0
					|| tabData[nNewIndex].msgHeight == panel1.Height)
				{
					return;
				}
				int nHeight = tabData[nNewIndex].msgHeight;
				if(nHeight > this.Height - 80)
				{
					nHeight = this.Height - 80;
				}
				panel1.Height = nHeight;
			}
		}

		private void btnMsgHide_Click(object sender, System.EventArgs e)
		{
			int nNewIndex = tabControl1.SelectedIndex;
			if(txtMsg.Visible)
			{
				tabData[nNewIndex].msgHeight = panel1.Height;
				panel1.Height = btnMsgHide.Height;
				btnMsgHide.Text = "Show Log";
				splitter1.Enabled = false;
				tabData[nNewIndex].isVisible = false;
				txtMsg.Visible = false;
			}
			else
			{
				int nHeight = tabData[nNewIndex].msgHeight;
				if(nHeight > this.Height - 80)
				{
					nHeight = this.Height - 80;
				}
				if(nHeight < btnMsgHide.Height)
				{
					nHeight = btnMsgHide.Height + 20;
				}
				panel1.Height = nHeight;
				btnMsgHide.Text = "Hide Log";
				splitter1.Enabled = true;
				tabData[nNewIndex].isVisible = true;
				txtMsg.Visible = true;
			}
			btnMsgHide.ForeColor = Color.Black;
		}

		private void txtMsg_TextChanged(object sender, System.EventArgs e)
		{
			if(txtMsg.Visible == false && !txtMsg.Text.Equals(""))
			{
				btnMsgHide.ForeColor = Color.Blue;
			}
		}

		private void btnMsgClear_Click(object sender, System.EventArgs e)
		{
			txtMsg.Text = "";
		}

		void Config_Load()
		{
			string sFileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\" + this.CompanyName + ".ini";
			System.Collections.Specialized.NameValueCollection coll = cc.Util.ReadIni(sFileName);
			if(coll == null)
			{
				coll = new System.Collections.Specialized.NameValueCollection();
			}
			try
			{
				if(coll.Get("windows_x") != null && coll.Get("windows_y") != null)
				{
					this.Location = new System.Drawing.Point(int.Parse(coll.Get("windows_y"))
						, int.Parse(coll.Get("windows_y")));
				}
				if(coll.Get("windows_w") != null && coll.Get("windows_h") != null)
				{
					this.Size = new System.Drawing.Size(int.Parse(coll.Get("windows_w"))
						, int.Parse(coll.Get("windows_h")));
				}
				if(coll["activate_tab"] != null)
				{
					tabControl1.SelectedIndex = int.Parse(coll["activate_tab"]);
				}
				if(coll.Get("windows_state") != null)
				{
					if(coll.Get("windows_state").Equals("Maximized"))
					{
						WindowState = FormWindowState.Maximized;
					}
					if(coll.Get("windows_state").Equals("Minimized"))
					{
						WindowState = FormWindowState.Minimized;
					}
				}
			}
			catch
			{
			}
			if(coll.Get("txtSaveTo") != null && !coll.Get("txtSaveTo").Equals(""))
			{
				txtSaveTo.Text = coll.Get("txtSaveTo");
			}
			if(coll.Get("radFromExcel") != null)
			{
				if(coll.Get("radFromExcel").Equals("1"))
				{
					radFromExcel.Checked = true;
					radFromServer.Checked = false;
				}
				else
				{
					radFromExcel.Checked = false;
					radFromServer.Checked = true;
				}
			}

			//get all tabpage's msgHeight
			for(int i = 0; i < tabControl1.TabCount; i++)
			{
				try
				{
					if(coll["msgHeight" + i] != null)
					{
						tabData[i].msgHeight = int.Parse(coll["msgHeight" + i]);
					}
					if(coll["isVisible" + i] != null)
					{
						tabData[i].isVisible = coll["isVisible" + i].Equals("True");
					}
				}
				catch
				{
				}
			}

			//all sub pages
			for(int i = 0; i < tabControl1.TabPages.Count; i++)
			{
				if(tabControl1.TabPages[i].Controls.Count > 0)
				{
					Control userCtl = tabControl1.TabPages[i].Controls[0];
					System.Type ctlType = userCtl.GetType();
					if(ctlType.BaseType != null && ctlType.BaseType.Name.Equals("UserControl"))
					{
						System.Reflection.MethodInfo method = ctlType.GetMethod("Config_Load");
						if(method != null)
						{
							method.Invoke(userCtl, new object[]{coll});
						}
					}
				}
			}
		}

		void Config_Save()
		{
			string sFileName = Path.GetDirectoryName(Application.ExecutablePath) + "\\" + this.CompanyName + ".ini";
			System.IO.StreamWriter sw = new System.IO.StreamWriter(sFileName, false, System.Text.Encoding.Default);
			sw.Write("#ini file for " + this.CompanyName + " - made by NewCon.ShuKK\r\n");
			sw.Write("windows_state={0}\r\n", this.WindowState);
			if(WindowState != FormWindowState.Normal)
			{
				WindowState = FormWindowState.Normal;
			}
			sw.Write("windows_x={0}\r\n", this.Location.X);
			sw.Write("windows_y={0}\r\n", this.Location.Y);
			sw.Write("windows_w={0}\r\n", this.Size.Width);
			sw.Write("windows_h={0}\r\n", this.Size.Height);
			sw.Write("windows_hr={0}\r\n", this.panel1.Height);
			sw.Write("activate_tab={0}\r\n", tabControl1.SelectedIndex);
			sw.Write("radFromExcel={0}\r\n", radFromExcel.Checked ? "1" : "0");
			sw.Write("txtSaveTo={0}\r\n", txtSaveTo.Text);

			//save all tabpage's msgHeight
			if(txtMsg.Visible)
			{
				tabData[tabControl1.SelectedIndex].msgHeight = panel1.Height;
			}
			for(int i = 0; i < tabControl1.TabCount; i++)
			{
				sw.Write("msgHeight{0}={1}\r\n", i, tabData[i].msgHeight);
				sw.Write("isVisible{0}={1}\r\n", i, tabData[i].isVisible);
			}

			//all sub pages
			for(int i = 0; i < tabControl1.TabPages.Count; i++)
			{
				if(tabControl1.TabPages[i].Controls.Count > 0)
				{
					Control userCtl = tabControl1.TabPages[i].Controls[0];
					System.Type ctlType = userCtl.GetType();
					if(ctlType.BaseType != null && ctlType.BaseType.Name.Equals("UserControl"))
					{
						System.Reflection.MethodInfo method = ctlType.GetMethod("Config_Save");
						if(method != null)
						{
							method.Invoke(userCtl, new object[]{sw});
						}
					}
				}
			}

			sw.Flush();
			sw.Close();
		}

	}

}
