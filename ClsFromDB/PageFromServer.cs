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
	public class PageFromServer : System.Windows.Forms.UserControl
	{
		public Form1 frmMain = null;

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnSelAllR;
		private System.Windows.Forms.Button btnSelUNR;
		private System.Windows.Forms.Button btnGetDB;
		private System.Windows.Forms.Button btnTBLAdd;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox txtSqlUser;
		private System.Windows.Forms.ComboBox lstServer;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox txtSqlPass;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Button btnDBDel;
		private System.Windows.Forms.Button btnGetTBL;
		private System.Windows.Forms.TextBox txtSqlServer;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.TextBox txtSqlString;
		private System.Windows.Forms.TextBox txtSqlDB;
		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.CheckBox chkConnType;
		private System.Windows.Forms.Button btnSelUNL;
		private System.Windows.Forms.Button btnSelAllL;
		public System.Windows.Forms.ListBox lstTBL;
		public System.Windows.Forms.ListBox lstDB;
		public System.Windows.Forms.ListBox lstDBTBL;

		public PageFromServer()
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
			this.btnSelAllR = new System.Windows.Forms.Button();
			this.btnSelUNR = new System.Windows.Forms.Button();
			this.btnGetDB = new System.Windows.Forms.Button();
			this.btnTBLAdd = new System.Windows.Forms.Button();
			this.lstTBL = new System.Windows.Forms.ListBox();
			this.label5 = new System.Windows.Forms.Label();
			this.txtSqlUser = new System.Windows.Forms.TextBox();
			this.lstServer = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.txtSqlPass = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.lstDB = new System.Windows.Forms.ListBox();
			this.lstDBTBL = new System.Windows.Forms.ListBox();
			this.btnDBDel = new System.Windows.Forms.Button();
			this.btnGetTBL = new System.Windows.Forms.Button();
			this.txtSqlServer = new System.Windows.Forms.TextBox();
			this.label8 = new System.Windows.Forms.Label();
			this.txtSqlString = new System.Windows.Forms.TextBox();
			this.txtSqlDB = new System.Windows.Forms.TextBox();
			this.label10 = new System.Windows.Forms.Label();
			this.chkConnType = new System.Windows.Forms.CheckBox();
			this.btnSelUNL = new System.Windows.Forms.Button();
			this.btnSelAllL = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// btnSelAllR
			// 
			this.btnSelAllR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSelAllR.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelAllR.Location = new System.Drawing.Point(532, 56);
			this.btnSelAllR.Name = "btnSelAllR";
			this.btnSelAllR.Size = new System.Drawing.Size(32, 16);
			this.btnSelAllR.TabIndex = 35;
			this.btnSelAllR.Text = "All";
			this.btnSelAllR.Click += new System.EventHandler(this.btnSelAllR_Click);
			// 
			// btnSelUNR
			// 
			this.btnSelUNR.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.btnSelUNR.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelUNR.Location = new System.Drawing.Point(564, 56);
			this.btnSelUNR.Name = "btnSelUNR";
			this.btnSelUNR.Size = new System.Drawing.Size(43, 16);
			this.btnSelUNR.TabIndex = 33;
			this.btnSelUNR.Text = "UnAll";
			this.btnSelUNR.Click += new System.EventHandler(this.btnSelUNR_Click);
			// 
			// btnGetDB
			// 
			this.btnGetDB.Location = new System.Drawing.Point(0, 72);
			this.btnGetDB.Name = "btnGetDB";
			this.btnGetDB.Size = new System.Drawing.Size(64, 23);
			this.btnGetDB.TabIndex = 29;
			this.btnGetDB.Text = "get DB";
			this.btnGetDB.Click += new System.EventHandler(this.btnGetDB_Click);
			// 
			// btnTBLAdd
			// 
			this.btnTBLAdd.Location = new System.Drawing.Point(249, 192);
			this.btnTBLAdd.Name = "btnTBLAdd";
			this.btnTBLAdd.Size = new System.Drawing.Size(14, 23);
			this.btnTBLAdd.TabIndex = 27;
			this.btnTBLAdd.Text = ">";
			this.btnTBLAdd.Click += new System.EventHandler(this.btnTBLAdd_Click);
			// 
			// lstTBL
			// 
			this.lstTBL.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.lstTBL.ItemHeight = 12;
			this.lstTBL.Location = new System.Drawing.Point(64, 168);
			this.lstTBL.Name = "lstTBL";
			this.lstTBL.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.lstTBL.Size = new System.Drawing.Size(184, 136);
			this.lstTBL.TabIndex = 24;
			this.lstTBL.DoubleClick += new System.EventHandler(this.lstTBL_DoubleClick);
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(64, 56);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(64, 16);
			this.label5.TabIndex = 23;
			this.label5.Text = "DB List:";
			// 
			// txtSqlUser
			// 
			this.txtSqlUser.Location = new System.Drawing.Point(64, 32);
			this.txtSqlUser.Name = "txtSqlUser";
			this.txtSqlUser.Size = new System.Drawing.Size(104, 19);
			this.txtSqlUser.TabIndex = 19;
			this.txtSqlUser.Text = "sa";
			// 
			// lstServer
			// 
			this.lstServer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.lstServer.Items.AddRange(new object[] {
														   "Oracle",
														   "SQLServer"});
			this.lstServer.Location = new System.Drawing.Point(64, 8);
			this.lstServer.Name = "lstServer";
			this.lstServer.Size = new System.Drawing.Size(104, 20);
			this.lstServer.TabIndex = 15;
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(0, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 11;
			this.label1.Text = "DBMS:";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(0, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(72, 16);
			this.label2.TabIndex = 12;
			this.label2.Text = "USER:";
			// 
			// txtSqlPass
			// 
			this.txtSqlPass.Location = new System.Drawing.Point(216, 32);
			this.txtSqlPass.Name = "txtSqlPass";
			this.txtSqlPass.PasswordChar = '*';
			this.txtSqlPass.Size = new System.Drawing.Size(88, 19);
			this.txtSqlPass.TabIndex = 17;
			this.txtSqlPass.Text = "licsadmin";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(176, 32);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(72, 16);
			this.label3.TabIndex = 10;
			this.label3.Text = "PASS:";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(64, 152);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(64, 16);
			this.label6.TabIndex = 22;
			this.label6.Text = "TBL List:";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(264, 56);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(200, 16);
			this.label7.TabIndex = 21;
			this.label7.Text = "Create Class From TBL List:";
			// 
			// lstDB
			// 
			this.lstDB.ItemHeight = 12;
			this.lstDB.Location = new System.Drawing.Point(64, 72);
			this.lstDB.Name = "lstDB";
			this.lstDB.Size = new System.Drawing.Size(184, 76);
			this.lstDB.TabIndex = 26;
			this.lstDB.DoubleClick += new System.EventHandler(this.lstDB_DoubleClick);
			// 
			// lstDBTBL
			// 
			this.lstDBTBL.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lstDBTBL.ItemHeight = 12;
			this.lstDBTBL.Location = new System.Drawing.Point(264, 72);
			this.lstDBTBL.Name = "lstDBTBL";
			this.lstDBTBL.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
			this.lstDBTBL.Size = new System.Drawing.Size(344, 232);
			this.lstDBTBL.Sorted = true;
			this.lstDBTBL.TabIndex = 25;
			this.lstDBTBL.DoubleClick += new System.EventHandler(this.lstDBTBL_DoubleClick);
			// 
			// btnDBDel
			// 
			this.btnDBDel.Location = new System.Drawing.Point(249, 216);
			this.btnDBDel.Name = "btnDBDel";
			this.btnDBDel.Size = new System.Drawing.Size(14, 23);
			this.btnDBDel.TabIndex = 28;
			this.btnDBDel.Text = "<";
			this.btnDBDel.Click += new System.EventHandler(this.btnDBDel_Click);
			// 
			// btnGetTBL
			// 
			this.btnGetTBL.Location = new System.Drawing.Point(0, 96);
			this.btnGetTBL.Name = "btnGetTBL";
			this.btnGetTBL.Size = new System.Drawing.Size(64, 23);
			this.btnGetTBL.TabIndex = 30;
			this.btnGetTBL.Text = "get TBL";
			this.btnGetTBL.Click += new System.EventHandler(this.btnGetTBL_Click);
			// 
			// txtSqlServer
			// 
			this.txtSqlServer.Location = new System.Drawing.Point(352, 32);
			this.txtSqlServer.Name = "txtSqlServer";
			this.txtSqlServer.Size = new System.Drawing.Size(88, 19);
			this.txtSqlServer.TabIndex = 16;
			this.txtSqlServer.Text = "10.6.155.143";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(312, 32);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(48, 16);
			this.label8.TabIndex = 13;
			this.label8.Text = "Server:";
			// 
			// txtSqlString
			// 
			this.txtSqlString.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSqlString.Enabled = false;
			this.txtSqlString.Location = new System.Drawing.Point(264, 8);
			this.txtSqlString.Name = "txtSqlString";
			this.txtSqlString.Size = new System.Drawing.Size(344, 19);
			this.txtSqlString.TabIndex = 20;
			this.txtSqlString.Text = "Data Source=10.6.155.143;User ID=sa;Password=licsadmin;Initial Catalog=na_kazo_r3" +
				"db";
			// 
			// txtSqlDB
			// 
			this.txtSqlDB.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSqlDB.Location = new System.Drawing.Point(472, 32);
			this.txtSqlDB.Name = "txtSqlDB";
			this.txtSqlDB.Size = new System.Drawing.Size(136, 19);
			this.txtSqlDB.TabIndex = 18;
			this.txtSqlDB.Text = "na_kazo_r3db";
			// 
			// label10
			// 
			this.label10.Location = new System.Drawing.Point(448, 32);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(32, 16);
			this.label10.TabIndex = 14;
			this.label10.Text = "DB:";
			// 
			// chkConnType
			// 
			this.chkConnType.Location = new System.Drawing.Point(176, 8);
			this.chkConnType.Name = "chkConnType";
			this.chkConnType.Size = new System.Drawing.Size(88, 24);
			this.chkConnType.TabIndex = 31;
			this.chkConnType.Text = "接続文字列";
			this.chkConnType.CheckedChanged += new System.EventHandler(this.chkConnType_CheckedChanged);
			// 
			// btnSelUNL
			// 
			this.btnSelUNL.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelUNL.Location = new System.Drawing.Point(205, 152);
			this.btnSelUNL.Name = "btnSelUNL";
			this.btnSelUNL.Size = new System.Drawing.Size(43, 16);
			this.btnSelUNL.TabIndex = 32;
			this.btnSelUNL.Text = "UnAll";
			this.btnSelUNL.Click += new System.EventHandler(this.btnSelUNL_Click);
			// 
			// btnSelAllL
			// 
			this.btnSelAllL.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelAllL.Location = new System.Drawing.Point(173, 152);
			this.btnSelAllL.Name = "btnSelAllL";
			this.btnSelAllL.Size = new System.Drawing.Size(32, 16);
			this.btnSelAllL.TabIndex = 34;
			this.btnSelAllL.Text = "All";
			this.btnSelAllL.Click += new System.EventHandler(this.btnSelAllL_Click);
			// 
			// PageFromServer
			// 
			this.Controls.Add(this.btnSelAllR);
			this.Controls.Add(this.btnSelUNR);
			this.Controls.Add(this.btnGetDB);
			this.Controls.Add(this.btnTBLAdd);
			this.Controls.Add(this.lstTBL);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.txtSqlUser);
			this.Controls.Add(this.lstServer);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.txtSqlPass);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.lstDB);
			this.Controls.Add(this.lstDBTBL);
			this.Controls.Add(this.btnDBDel);
			this.Controls.Add(this.btnGetTBL);
			this.Controls.Add(this.txtSqlServer);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.txtSqlString);
			this.Controls.Add(this.txtSqlDB);
			this.Controls.Add(this.label10);
			this.Controls.Add(this.chkConnType);
			this.Controls.Add(this.btnSelUNL);
			this.Controls.Add(this.btnSelAllL);
			this.Name = "PageFromServer";
			this.Size = new System.Drawing.Size(608, 296);
			this.ResumeLayout(false);

		}
		#endregion

		private void chkConnType_CheckedChanged(object sender, System.EventArgs e)
		{
			if(chkConnType.Checked)
			{
				txtSqlString.Enabled = true;
				txtSqlUser.Enabled = false;
				txtSqlPass.Enabled = false;
				txtSqlServer.Enabled = false;
				txtSqlDB.Enabled = false;
			}
			else
			{
				txtSqlString.Enabled = false;
				txtSqlUser.Enabled = true;
				txtSqlPass.Enabled = true;
				txtSqlServer.Enabled = true;
				txtSqlDB.Enabled = true;
			}
		}

		private void btnSelAllL_Click(object sender, System.EventArgs e)
		{
			for(int i = 0; i < lstTBL.Items.Count; i++)
			{
				lstTBL.SetSelected(i, true);
			}
		}

		private void btnSelUNL_Click(object sender, System.EventArgs e)
		{
			lstTBL.ClearSelected();
		}

		private void btnSelAllR_Click(object sender, System.EventArgs e)
		{
			for(int i = 0; i < lstDBTBL.Items.Count; i++)
			{
				lstDBTBL.SetSelected(i, true);
			}
		}

		private void btnSelUNR_Click(object sender, System.EventArgs e)
		{
			lstDBTBL.ClearSelected();
		}

		private void btnGetDB_Click(object sender, System.EventArgs e)
		{
			//get DB list
			lstDB.Items.Clear();
			lstTBL.Items.Clear();
			lstDBTBL.Items.Clear();
			cc.DB cdb = ConnDB();
			if(cdb == null)
			{
				return;
			}
			try
			{
				DataTable tbl = cdb.GetTable("SELECT name FROM master.dbo.sysdatabases order by name");
				if(cdb.Error())
				{
					frmMain.labStatus.Text = "DATABASEの取得にできませんでした。";
					throw new Exception();
				}
				for(int i = 0; i < tbl.Rows.Count; i++)
				{
					lstDB.Items.Add(tbl.Rows[i].ItemArray[0].ToString());
				}
				frmMain.labStatus.Text = "DATABASEを取得しました。";
			}
			catch(Exception exp)
			{
				frmMain.msg.println("DATABASEの取得にエラーが発生しました：");
				frmMain.msg.println(exp.Message, Color.Red);
				frmMain.labStatus.Text = "DATABASEの取得にできませんでした。";
			}
			finally
			{
				if(cdb != null)
				{
					cdb.Dispose();
				}
			}
		}

		private void btnGetTBL_Click(object sender, System.EventArgs e)
		{
			const string SQLDBTBL= "select name, xtype, id from [{0}].dbo.sysobjects where name like ('{1}%') and xtype in ('U','V') order by name";
			//get TBL list
			lstTBL.Items.Clear();
			if(lstDB.SelectedIndex < 0)
			{
				frmMain.labStatus.Text = "Please select DB first.";
				return;
			}
			sfromdb = lstDB.Items[lstDB.SelectedIndex].ToString();

			cc.DB cdb = ConnDB();
			if(cdb == null)
			{
				return;
			}
			try
			{
				DataTable tbl = cdb.GetTable(String.Format(SQLDBTBL, sfromdb, ""));
				if(cdb.Error())
				{
					frmMain.labStatus.Text = "TBLの取得にできませんでした。";
					throw new Exception();
				}
				for(int i = 0; i < tbl.Rows.Count; i++)
				{
					lstTBL.Items.Add(tbl.Rows[i].ItemArray[0].ToString());
				}
				frmMain.labStatus.Text = "TBLを取得しました。";
			}
			catch(Exception exp)
			{
				frmMain.msg.println("TBLの取得にエラーが発生しました：");
				frmMain.msg.println(exp.Message, Color.Red);
				frmMain.labStatus.Text = "TBLの取得にできませんでした。";
			}
			finally
			{
				if(cdb != null)
				{
					cdb.Dispose();
				}
			}
		}

		private void lstDB_DoubleClick(object sender, System.EventArgs e)
		{
			btnGetTBL_Click(sender, e);
		}

		private void lstTBL_DoubleClick(object sender, System.EventArgs e)
		{
			btnTBLAdd_Click(sender, e);
		}

		private void lstDBTBL_DoubleClick(object sender, System.EventArgs e)
		{
			btnDBDel_Click(sender, e);
		}

		private void btnTBLAdd_Click(object sender, System.EventArgs e)
		{
			int ncnt = 0;
			for(int i = lstTBL.Items.Count - 1; i >= 0; i--)
			{
				if(lstTBL.GetSelected(i))
				{
					bool isExist = false;
					string s = sfromdb + " . " + lstTBL.Items[i].ToString();
					for(int j = 0; j < lstDBTBL.Items.Count; j++)
					{
						if(lstDBTBL.Items[j].ToString().Equals(s))
						{
							isExist = true;
							break;
						}
					}
					if(!isExist)
					{
						lstDBTBL.Items.Add(s);
						ncnt++;
					}
				}
			}
			frmMain.labStatus.Text = "added item:" + ncnt;
		}

		private void btnDBDel_Click(object sender, System.EventArgs e)
		{
			int ncnt = 0;
			for(int i = lstDBTBL.Items.Count - 1; i >= 0; i--)
			{
				if(lstDBTBL.GetSelected(i))
				{
					lstDBTBL.Items.RemoveAt(i);
					ncnt++;
				}
			}
			frmMain.labStatus.Text = "removed item:" + ncnt;
		}

		private string sfromdb = "";
		public cc.DB ConnDB()
		{
			//get DB list
			string sConnString;
			if(chkConnType.Checked)
			{
				sConnString = txtSqlString.Text.Trim();
			}
			else
			{
				string sqlpath = txtSqlServer.Text.Trim();
				string sqluser = txtSqlUser.Text.Trim();
				string sqlpass = txtSqlPass.Text.Trim();
				string sqldb = "Initial Catalog=" + txtSqlDB.Text.Trim();
				sConnString = "Data Source=" + sqlpath + ";User ID=" + sqluser + ";Password=" + sqlpass + ";" + sqldb;
			}
			cc.DB cdb = null;
			try
			{
				frmMain.labStatus.Text = "SQLServerに接続...";
				cdb = new cc.DB(sConnString);
				if(cdb.Error())
				{
					frmMain.labStatus.Text = "SQLServerに接続...できませんでした。";
					throw new Exception();
				}
				frmMain.labStatus.Text = "SQLServerに接続...しました。";
				Application.DoEvents();
			}
			catch(Exception exp)
			{
				frmMain.msg.println("SQLServerの接続にエラーが発生しました：");
				frmMain.msg.println(exp.Message, Color.Red);
				cdb = null;
			}
			return cdb;
		}

		public void Config_Load(System.Collections.Specialized.NameValueCollection coll)
		{
			if(coll.Get("lstServer") != null)
			{
				lstServer.SelectedIndex = Int16.Parse(coll.Get("lstServer"));
			}
			if(coll.Get("txtSqlString") != null && !coll.Get("txtSqlString").Equals(""))
			{
				txtSqlString.Text = coll.Get("txtSqlString");
			}
			if(coll.Get("txtSqlUser") != null && !coll.Get("txtSqlUser").Equals(""))
			{
				txtSqlUser.Text = coll.Get("txtSqlUser");
			}
			if(coll.Get("txtSqlPass") != null)
			{
				txtSqlPass.Text = coll.Get("txtSqlPass");
			}
			if(coll.Get("txtSqlServer") != null)
			{
				txtSqlServer.Text = coll.Get("txtSqlServer");
			}
			if(coll.Get("txtSqlDB") != null)
			{
				txtSqlDB.Text = coll.Get("txtSqlDB");
			}
			if(coll.Get("chkConnType") != null)
			{
				if(coll.Get("chkConnType").Equals("0"))
				{
					chkConnType.Checked = false;
				}
				else
				{
					//chkConnType_CheckedChanged(null, null);
					chkConnType.Checked = true;
				}
			}
			//restore server table list
			if(coll["DBServerTableList"] != null)
			{
				string[] slist = coll["DBServerTableList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					lstDBTBL.Items.Add(slist[i]);
				}
			}
		}

		public void Config_Save(System.IO.StreamWriter sw)
		{
			sw.Write("lstServer={0}\r\n", lstServer.SelectedIndex);
			sw.Write("txtSqlString={0}\r\n", txtSqlString.Text);
			sw.Write("chkConnType={0}\r\n", chkConnType.Checked ? "1" : "0");
			sw.Write("txtSqlUser={0}\r\n", txtSqlUser.Text);
			sw.Write("txtSqlPass={0}\r\n", txtSqlPass.Text);
			sw.Write("txtSqlServer={0}\r\n", txtSqlServer.Text);
			sw.Write("txtSqlDB={0}\r\n", txtSqlDB.Text);
			//save server table list
			for(int i = 0; i < lstDBTBL.Items.Count; i++)
			{
				sw.Write("DBServerTableList={0}\r\n", lstDBTBL.Items[i].ToString());
			}
		}

	}
}
