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
	/// PageFormExcel の概要の説明です。
	/// </summary>
	public class PageFromExcel : System.Windows.Forms.UserControl
	{
		public Form1 frmMain = null;

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnLoadExcel;
		private System.Windows.Forms.Button btnExcelTemp;
		private System.Windows.Forms.Button btnSelAllExcel;
		private System.Windows.Forms.Button btnSelUNExcel;
		private System.Windows.Forms.Button btnClearExcel;
		public System.Windows.Forms.CheckedListBox lstExcelTemp;
		public System.Windows.Forms.CheckedListBox lstExcel;
		public System.Windows.Forms.TextBox txtExcelTemp;

		public PageFromExcel()
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
			this.lstExcel = new System.Windows.Forms.CheckedListBox();
			this.txtExcelTemp = new System.Windows.Forms.TextBox();
			this.btnLoadExcel = new System.Windows.Forms.Button();
			this.btnExcelTemp = new System.Windows.Forms.Button();
			this.btnSelAllExcel = new System.Windows.Forms.Button();
			this.btnSelUNExcel = new System.Windows.Forms.Button();
			this.lstExcelTemp = new System.Windows.Forms.CheckedListBox();
			this.btnClearExcel = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// lstExcel
			// 
			this.lstExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lstExcel.CheckOnClick = true;
			this.lstExcel.Location = new System.Drawing.Point(80, 88);
			this.lstExcel.Name = "lstExcel";
			this.lstExcel.Size = new System.Drawing.Size(512, 158);
			this.lstExcel.TabIndex = 22;
			this.lstExcel.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstExcel_ItemCheck);
			// 
			// txtExcelTemp
			// 
			this.txtExcelTemp.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtExcelTemp.Location = new System.Drawing.Point(80, 8);
			this.txtExcelTemp.Name = "txtExcelTemp";
			this.txtExcelTemp.ReadOnly = true;
			this.txtExcelTemp.Size = new System.Drawing.Size(464, 12);
			this.txtExcelTemp.TabIndex = 21;
			this.txtExcelTemp.Text = "Please Load ExcelTemplate File for Excel File Infomation, and select one sheet be" +
				"low.";
			// 
			// btnLoadExcel
			// 
			this.btnLoadExcel.Location = new System.Drawing.Point(0, 88);
			this.btnLoadExcel.Name = "btnLoadExcel";
			this.btnLoadExcel.TabIndex = 19;
			this.btnLoadExcel.Text = "LoadExcel";
			this.btnLoadExcel.Click += new System.EventHandler(this.btnLoadExcel_Click);
			// 
			// btnExcelTemp
			// 
			this.btnExcelTemp.Location = new System.Drawing.Point(0, 8);
			this.btnExcelTemp.Name = "btnExcelTemp";
			this.btnExcelTemp.TabIndex = 15;
			this.btnExcelTemp.Text = "LoadTemp";
			this.btnExcelTemp.Click += new System.EventHandler(this.btnExcelTemp_Click);
			// 
			// btnSelAllExcel
			// 
			this.btnSelAllExcel.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelAllExcel.Location = new System.Drawing.Point(0, 112);
			this.btnSelAllExcel.Name = "btnSelAllExcel";
			this.btnSelAllExcel.Size = new System.Drawing.Size(32, 16);
			this.btnSelAllExcel.TabIndex = 16;
			this.btnSelAllExcel.Text = "All";
			this.btnSelAllExcel.Click += new System.EventHandler(this.btnSelAllExcel_Click);
			// 
			// btnSelUNExcel
			// 
			this.btnSelUNExcel.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelUNExcel.Location = new System.Drawing.Point(32, 112);
			this.btnSelUNExcel.Name = "btnSelUNExcel";
			this.btnSelUNExcel.Size = new System.Drawing.Size(43, 16);
			this.btnSelUNExcel.TabIndex = 17;
			this.btnSelUNExcel.Text = "UnAll";
			this.btnSelUNExcel.Click += new System.EventHandler(this.btnSelUNExcel_Click);
			// 
			// lstExcelTemp
			// 
			this.lstExcelTemp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lstExcelTemp.CheckOnClick = true;
			this.lstExcelTemp.Location = new System.Drawing.Point(80, 32);
			this.lstExcelTemp.Name = "lstExcelTemp";
			this.lstExcelTemp.Size = new System.Drawing.Size(512, 46);
			this.lstExcelTemp.TabIndex = 20;
			this.lstExcelTemp.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstExcelTemp_ItemCheck);
			// 
			// btnClearExcel
			// 
			this.btnClearExcel.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnClearExcel.Location = new System.Drawing.Point(0, 128);
			this.btnClearExcel.Name = "btnClearExcel";
			this.btnClearExcel.Size = new System.Drawing.Size(75, 16);
			this.btnClearExcel.TabIndex = 18;
			this.btnClearExcel.Text = "Remove All";
			this.btnClearExcel.Click += new System.EventHandler(this.btnClearExcel_Click);
			// 
			// PageFromExcel
			// 
			this.Controls.Add(this.lstExcel);
			this.Controls.Add(this.txtExcelTemp);
			this.Controls.Add(this.btnLoadExcel);
			this.Controls.Add(this.btnExcelTemp);
			this.Controls.Add(this.btnSelAllExcel);
			this.Controls.Add(this.btnSelUNExcel);
			this.Controls.Add(this.lstExcelTemp);
			this.Controls.Add(this.btnClearExcel);
			this.Name = "PageFromExcel";
			this.Size = new System.Drawing.Size(592, 240);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnExcelTemp_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			string initpath = Path.GetDirectoryName(Application.ExecutablePath);
			string sfilename = txtExcelTemp.Text.Trim();
			if(sfilename.StartsWith(".\\"))
			{
				sfilename = initpath + sfilename.Substring(1);
			}
			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			if(File.Exists(sfilename))
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(sfilename);
			}
			else
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(Application.ExecutablePath);
			}
			openFileDialog1.Filter = "テーブルレイアウトTemplateファイル(*.xls)|*.xls";
			openFileDialog1.Title = "テーブルレイアウトTemplateファイルを選択してください。";

			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				lstExcelTemp.Items.Clear();

				string sExcelName = openFileDialog1.FileName;
				if(Path.GetDirectoryName(sExcelName).StartsWith(initpath))
				{
					txtExcelTemp.Text = sExcelName.Substring(initpath.Length);
					if(txtExcelTemp.Text.StartsWith("\\"))
					{
						txtExcelTemp.Text = "." + txtExcelTemp.Text;
					}
				}
				else
				{
					txtExcelTemp.Text = sExcelName;
				}

				frmMain.labStatus.Text = "Now is getting excle file information...";
				Application.DoEvents();

				string sConnExcel = cc.OleDB.ConnStringExcel(sExcelName, false);
				cc.OleDB dbExcel = new cc.OleDB(sConnExcel);
				if(dbExcel.Error())
				{
					frmMain.msg.println("Open Excel File Error:");
					frmMain.msg.println("File:" + sExcelName);
					frmMain.msg.println(dbExcel.Exception.Message, Color.Red);
					return;
				}
				string[] sSheetsName = dbExcel.GetExcelSheetsName();
				if(sSheetsName != null)
				{
					foreach(string s in sSheetsName)
					{
						lstExcelTemp.Items.Add(s);
					}
				}
				dbExcel.Dispose();
				if(lstExcelTemp.Items.Count > 0)
				{
					lstExcelTemp.SetItemChecked(0, true);
					frmMain.labStatus.Text = "Now is getting excle file information...OK";
				}
				else
				{
					frmMain.msg.println("This Excel file not include Template information.", Color.Red);
					frmMain.labStatus.Text = "Now is getting excle file information...no Template included";
				}
			}
		}

		private void btnLoadExcel_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			if(lstExcel.Tag != null && File.Exists(lstExcel.Tag.ToString()))
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(lstExcel.Tag.ToString());
			}
			else
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(Application.ExecutablePath);
			}
			openFileDialog1.Filter = "テーブルレイアウトファイル(*.xls)|*.xls";
			openFileDialog1.Title = "テーブルレイアウトファイルを選択してください。";
			openFileDialog1.Multiselect = true;

			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				string[] sExcelName = openFileDialog1.FileNames;
				lstExcel.Tag = sExcelName[0];
				for(int i = 0; i < sExcelName.Length; i++)
				{
					if(lstExcel.FindString(sExcelName[i]) >= 0)
					{
						continue;
					}
					lstExcel.Items.Add(sExcelName[i]);
					frmMain.labStatus.Text = "Now is getting excle file information...";
					Application.DoEvents();

					string sConnExcel = cc.OleDB.ConnStringExcel(sExcelName[i], false);
					cc.OleDB dbExcel = new cc.OleDB(sConnExcel);
					if(dbExcel.Error())
					{
						frmMain.msg.println("Open Excel File Error:");
						frmMain.msg.println("File:" + sExcelName[i]);
						frmMain.msg.println(dbExcel.Exception.Message, Color.Red);
						return;
					}
					string[] sSheetsName = dbExcel.GetExcelSheetsName();
					if(sSheetsName != null)
					{
						foreach(string s in sSheetsName)
						{
							lstExcel.Items.Add("　　" + s);
						}
					}
					frmMain.labStatus.Text = "Now is getting excle file information...OK";
					dbExcel.Dispose();
				}
			}
		}

		private void btnSelAllExcel_Click(object sender, System.EventArgs e)
		{
			for(int i = 0; i < lstExcel.Items.Count; i++)
			{
				lstExcel.SetItemChecked(i, true);
			}
		}

		private void btnSelUNExcel_Click(object sender, System.EventArgs e)
		{
			for(int i = 0; i < lstExcel.Items.Count; i++)
			{
				lstExcel.SetItemChecked(i, false);
			}
		}

		private void btnClearExcel_Click(object sender, System.EventArgs e)
		{
			lstExcel.Items.Clear();
		}

		bool bOnlyOneIn = false; //for not reloop in CheckListBox_ItemCheck
		private void lstExcelTemp_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			if(bOnlyOneIn)
			{
				return;
			}
			if(lstExcelTemp.SelectedIndex < 0)
			{
				return;
			}
			bOnlyOneIn = true;
			if(e.NewValue == CheckState.Checked)
			{
				for(int i = 0; i < lstExcelTemp.Items.Count; i++)
				{
					if(i != lstExcelTemp.SelectedIndex)
					{
						lstExcelTemp.SetItemChecked(i, false);
					}
				}
			}
			bOnlyOneIn = false;
		}

		private void lstExcel_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			if(bOnlyOneIn)
			{
				return;
			}
			int nindex = lstExcel.SelectedIndex;
			if(nindex < 0)
			{
				return;
			}
			bOnlyOneIn = true;
			if(lstExcel.Items[nindex].ToString().Substring(1, 2).Equals(":\\"))
			{
				//all sheets of File select or unselect
				for(int i = nindex + 1; i < lstExcel.Items.Count; i++)
				{
					if(lstExcel.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						break;
					}
					lstExcel.SetItemChecked(i, (e.NewValue == CheckState.Checked));
				}
			}
			else
			{
				//if all sheets of File select or unselect,then select or unselect File
				int nend = nindex;
				bool balselall = true;
				bool balselno = true;
				if(e.NewValue == CheckState.Checked)
				{
					balselno = false;
				}
				else
				{
					balselall = false;
				}
				for(int i = nindex + 1; i < lstExcel.Items.Count; i++)
				{
					if(lstExcel.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						break;
					}
					nend = i;
				}
				for(int i = nend; i >= 0; i--)
				{
					if(lstExcel.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						if(balselall && !balselno)
						{
							lstExcel.SetItemChecked(i, true);
						}
						else
						{
							lstExcel.SetItemChecked(i, false);
						}
						break;
					}
					if(i != nindex)
					{
						if(lstExcel.GetItemChecked(i))
						{
							balselno = false;
						}
						else
						{
							balselall = false;
						}
					}
				}
			}
			bOnlyOneIn = false;
		}

		public void Config_Load(System.Collections.Specialized.NameValueCollection coll)
		{
			if(coll.Get("txtExcelTemp") != null && !coll.Get("txtExcelTemp").Equals(""))
			{
				txtExcelTemp.Text = coll.Get("txtExcelTemp");
			}
			//restore excel file(template) list
			if(coll["ExcelTemplateList"] != null)
			{
				string[] slist = coll["ExcelTemplateList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					string line = slist[i];
					int npos = line.LastIndexOf("=");
					if(npos > 0)
					{
						lstExcelTemp.Items.Add(line.Substring(0, npos));
						if(line.Substring(npos + 1).Equals("on"))
						{
							lstExcelTemp.SetItemChecked(lstExcelTemp.Items.Count - 1, true);
						}
					}
				}
			}
			if(coll.Get("txtExcel") != null && !coll.Get("txtExcel").Equals(""))
			{
				lstExcel.Tag = coll.Get("txtExcel");
			}
			//restore excel file(table) list
			if(coll["ExcelTableList"] != null)
			{
				string[] slist = coll["ExcelTableList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					string line = slist[i];
					int npos = line.LastIndexOf("=");
					if(npos > 0)
					{
						lstExcel.Items.Add(line.Substring(0, npos));
						if(line.Substring(npos + 1).Equals("on"))
						{
							lstExcel.SetItemChecked(lstExcel.Items.Count - 1, true);
						}
					}
				}
			}
		}

		public void Config_Save(System.IO.StreamWriter sw)
		{
			sw.Write("txtExcelTemp={0}\r\n", txtExcelTemp.Text);
			//save excel file(template) list
			for(int i = 0; i < lstExcelTemp.Items.Count; i++)
			{
				if(lstExcelTemp.GetItemChecked(i))
				{
					sw.Write("ExcelTemplateList={0}={1}\r\n", lstExcelTemp.Items[i].ToString(), "on");
				}
				else
				{
					sw.Write("ExcelTemplateList={0}={1}\r\n", lstExcelTemp.Items[i].ToString(), "off");
				}
			}
			sw.Write("txtExcel={0}\r\n", lstExcel.Tag);
			//save excel file(table) list
			for(int i = 0; i < lstExcel.Items.Count; i++)
			{
				if(lstExcel.GetItemChecked(i))
				{
					sw.Write("ExcelTableList={0}={1}\r\n", lstExcel.Items[i].ToString(), "on");
				}
				else
				{
					sw.Write("ExcelTableList={0}={1}\r\n", lstExcel.Items[i].ToString(), "off");
				}
			}
		}

	}
}
