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
	public class PageVarious : System.Windows.Forms.UserControl
	{
		public Form1 frmMain = null;
		public System.Collections.Specialized.NameValueCollection nvcSysVar = 
			new System.Collections.Specialized.NameValueCollection();

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.Button btnVarAdd;
		private System.Windows.Forms.Button btnVarMod;
		private System.Windows.Forms.Button btnVarDel;
		private System.Windows.Forms.TextBox txtVarName;
		public System.Windows.Forms.TextBox txtVarLable;
		private System.Windows.Forms.TextBox txtVarEdit;
		public System.Windows.Forms.TextBox textBox2;
		public System.Windows.Forms.ListBox lstVar;
		private System.Windows.Forms.Button btnTypeAdd;
		private System.Windows.Forms.Button btnTypeMod;
		private System.Windows.Forms.Button btnTypeDel;
		public System.Windows.Forms.ListBox lstType;
		private System.Windows.Forms.ComboBox comTypeLang;
		private System.Windows.Forms.ComboBox comTypeSys;
		private System.Windows.Forms.ComboBox comTypeUser;
		private System.Windows.Forms.ComboBox comInitVal;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.ComboBox comMaxVal;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.TextBox txtDefaultLen;

		public PageVarious()
		{
			// この呼び出しは、Windows.Forms フォーム デザイナで必要です。
			InitializeComponent();

			defineSysVar();
			comMaxVal.SelectedIndex = 0;
		}

		//ad sys various
		private void defineSysVar()
		{
			nvcSysVar["[#SYS_DATE#]"] = System.DateTime.Now.ToString("yyyy/MM/dd");
			nvcSysVar["[#SYS_TIME#]"] = System.DateTime.Now.ToString("HH:mm:ss");
			nvcSysVar["[#CREATE_DATE#]"] = "";
			nvcSysVar["[#CREATE_VERSION#]"] = "";
			nvcSysVar["[#CREATE_AUTHOR#]"] = "";
			nvcSysVar["[#DB_NAME#]"] = "";
			nvcSysVar["[#TABLE_NAME#]"] = "";
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
			this.txtVarName = new System.Windows.Forms.TextBox();
			this.txtVarLable = new System.Windows.Forms.TextBox();
			this.btnVarAdd = new System.Windows.Forms.Button();
			this.btnVarMod = new System.Windows.Forms.Button();
			this.btnVarDel = new System.Windows.Forms.Button();
			this.btnTypeAdd = new System.Windows.Forms.Button();
			this.btnTypeMod = new System.Windows.Forms.Button();
			this.btnTypeDel = new System.Windows.Forms.Button();
			this.comTypeLang = new System.Windows.Forms.ComboBox();
			this.txtVarEdit = new System.Windows.Forms.TextBox();
			this.textBox2 = new System.Windows.Forms.TextBox();
			this.lstVar = new System.Windows.Forms.ListBox();
			this.lstType = new System.Windows.Forms.ListBox();
			this.comTypeSys = new System.Windows.Forms.ComboBox();
			this.comTypeUser = new System.Windows.Forms.ComboBox();
			this.comInitVal = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.comMaxVal = new System.Windows.Forms.ComboBox();
			this.label9 = new System.Windows.Forms.Label();
			this.txtDefaultLen = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// txtVarName
			// 
			this.txtVarName.Location = new System.Drawing.Point(56, 24);
			this.txtVarName.Name = "txtVarName";
			this.txtVarName.Size = new System.Drawing.Size(176, 19);
			this.txtVarName.TabIndex = 23;
			this.txtVarName.Text = "";
			// 
			// txtVarLable
			// 
			this.txtVarLable.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtVarLable.Location = new System.Drawing.Point(8, 24);
			this.txtVarLable.Name = "txtVarLable";
			this.txtVarLable.ReadOnly = true;
			this.txtVarLable.Size = new System.Drawing.Size(88, 12);
			this.txtVarLable.TabIndex = 21;
			this.txtVarLable.Text = "Various:";
			// 
			// btnVarAdd
			// 
			this.btnVarAdd.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnVarAdd.Location = new System.Drawing.Point(232, 24);
			this.btnVarAdd.Name = "btnVarAdd";
			this.btnVarAdd.Size = new System.Drawing.Size(32, 16);
			this.btnVarAdd.TabIndex = 16;
			this.btnVarAdd.Text = "add";
			this.btnVarAdd.Click += new System.EventHandler(this.btnVarAdd_Click);
			// 
			// btnVarMod
			// 
			this.btnVarMod.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnVarMod.Location = new System.Drawing.Point(232, 40);
			this.btnVarMod.Name = "btnVarMod";
			this.btnVarMod.Size = new System.Drawing.Size(32, 16);
			this.btnVarMod.TabIndex = 16;
			this.btnVarMod.Text = "mod";
			this.btnVarMod.Click += new System.EventHandler(this.btnVarMod_Click);
			// 
			// btnVarDel
			// 
			this.btnVarDel.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnVarDel.Location = new System.Drawing.Point(232, 56);
			this.btnVarDel.Name = "btnVarDel";
			this.btnVarDel.Size = new System.Drawing.Size(32, 16);
			this.btnVarDel.TabIndex = 16;
			this.btnVarDel.Text = "del";
			this.btnVarDel.Click += new System.EventHandler(this.btnVarDel_Click);
			// 
			// btnTypeAdd
			// 
			this.btnTypeAdd.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnTypeAdd.Location = new System.Drawing.Point(576, 24);
			this.btnTypeAdd.Name = "btnTypeAdd";
			this.btnTypeAdd.Size = new System.Drawing.Size(32, 16);
			this.btnTypeAdd.TabIndex = 16;
			this.btnTypeAdd.Text = "add";
			this.btnTypeAdd.Click += new System.EventHandler(this.btnTypeAdd_Click);
			// 
			// btnTypeMod
			// 
			this.btnTypeMod.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnTypeMod.Location = new System.Drawing.Point(576, 40);
			this.btnTypeMod.Name = "btnTypeMod";
			this.btnTypeMod.Size = new System.Drawing.Size(32, 16);
			this.btnTypeMod.TabIndex = 16;
			this.btnTypeMod.Text = "mod";
			this.btnTypeMod.Click += new System.EventHandler(this.btnTypeMod_Click);
			// 
			// btnTypeDel
			// 
			this.btnTypeDel.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnTypeDel.Location = new System.Drawing.Point(576, 56);
			this.btnTypeDel.Name = "btnTypeDel";
			this.btnTypeDel.Size = new System.Drawing.Size(32, 16);
			this.btnTypeDel.TabIndex = 16;
			this.btnTypeDel.Text = "del";
			this.btnTypeDel.Click += new System.EventHandler(this.btnTypeDel_Click);
			// 
			// comTypeLang
			// 
			this.comTypeLang.Items.AddRange(new object[] {
															 "Default",
															 "Java",
															 "Jsp",
															 "HTML",
															 "C#",
															 "SqlServerScript",
															 "OracleScript"});
			this.comTypeLang.Location = new System.Drawing.Point(336, 48);
			this.comTypeLang.Name = "comTypeLang";
			this.comTypeLang.Size = new System.Drawing.Size(80, 20);
			this.comTypeLang.TabIndex = 24;
			this.comTypeLang.Text = "Java";
			this.comTypeLang.Leave += new System.EventHandler(this.comTypeLang_Leave);
			// 
			// txtVarEdit
			// 
			this.txtVarEdit.Location = new System.Drawing.Point(56, 48);
			this.txtVarEdit.Name = "txtVarEdit";
			this.txtVarEdit.Size = new System.Drawing.Size(176, 19);
			this.txtVarEdit.TabIndex = 23;
			this.txtVarEdit.Text = "";
			// 
			// textBox2
			// 
			this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.textBox2.Location = new System.Drawing.Point(8, 48);
			this.textBox2.Name = "textBox2";
			this.textBox2.ReadOnly = true;
			this.textBox2.Size = new System.Drawing.Size(88, 12);
			this.textBox2.TabIndex = 21;
			this.textBox2.Text = "Value:";
			// 
			// lstVar
			// 
			this.lstVar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left)));
			this.lstVar.ItemHeight = 12;
			this.lstVar.Location = new System.Drawing.Point(8, 72);
			this.lstVar.Name = "lstVar";
			this.lstVar.Size = new System.Drawing.Size(256, 172);
			this.lstVar.TabIndex = 25;
			this.lstVar.SelectedIndexChanged += new System.EventHandler(this.lstVar_SelectedIndexChanged);
			// 
			// lstType
			// 
			this.lstType.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lstType.ItemHeight = 12;
			this.lstType.Items.AddRange(new object[] {
														 "[C#]datetime => DateTime[;]=InitVal[;]new DateTime(2005, 6, 30, 0, 0, 0)[;]10",
														 "[Default]char => String[;]char[;]\"\"[;]",
														 "[Default]datetime => String[;]=InitVal[;]\"2005/06/30\"[;]10",
														 "[Default]int => int[;]int[;]0[;]",
														 "[Default]number => int[;]int[;]0[;]",
														 "[Default]smallint => int[;]int[;]0[;]",
														 "[Default]varchar => String[;]char[;]\"\"[;]",
														 "[Default]varchar2 => String[;]char[;]\"\"[;]",
														 "[Java]datetime => java.util.Date[;]=InitVal[;]new Date(\"2005/06/22\")[;]10",
														 "[OracleScript]char => CHAR[;]char[;]\"\"[;]",
														 "[OracleScript]datetime => DATETIME[;]=InitVal[;]TO_DATE(\'20050630\', \'YYYYMMDD\')[;" +
														 "]10",
														 "[OracleScript]int => NUMBER[;]int[;]0[;]",
														 "[OracleScript]varchar => VARCHAR2[;]char[;]\"\"[;]",
														 "[SqlServerScript]char => CHAR[;]char[;]\"\"[;]",
														 "[SqlServerScript]datetime => DATETIME[;]=InitVal[;]cast(\'2005-06-30\'as datetime)[" +
														 ";]10",
														 "[SqlServerScript]int => NUMBER[;]int[;]0[;]",
														 "[SqlServerScript]varchar => VARCHAR[;]char[;]\"\"[;]"});
			this.lstType.Location = new System.Drawing.Point(280, 96);
			this.lstType.Name = "lstType";
			this.lstType.Size = new System.Drawing.Size(352, 148);
			this.lstType.Sorted = true;
			this.lstType.TabIndex = 25;
			this.lstType.SelectedIndexChanged += new System.EventHandler(this.lstType_SelectedIndexChanged);
			// 
			// comTypeSys
			// 
			this.comTypeSys.Items.AddRange(new object[] {
															"String",
															"string",
															"Boolean",
															"boolean",
															"bool",
															"Integer",
															"Int",
															"int",
															"Date",
															"ARRAY",
															"Array",
															"BFILE",
															"BigDecimal",
															"BLOB",
															"Blob",
															"Byte",
															"byte",
															"CHAR",
															"CLOB",
															"Clob",
															"CustomDatum",
															"Date",
															"DATE",
															"Datum",
															"Double",
															"double",
															"Float",
															"float",
															"int",
															"Integer",
															"Long",
															"long",
															"NUMBER",
															"RAW",
															"ROWID",
															"Short",
															"short",
															"String",
															"STRUCT",
															"Struct",
															"Time",
															"Timestamp"});
			this.comTypeSys.Location = new System.Drawing.Point(504, 24);
			this.comTypeSys.Name = "comTypeSys";
			this.comTypeSys.Size = new System.Drawing.Size(72, 20);
			this.comTypeSys.TabIndex = 24;
			this.comTypeSys.Text = "String";
			// 
			// comTypeUser
			// 
			this.comTypeUser.Items.AddRange(new object[] {
															 "(Oracle)dec",
															 "(Oracle)decimal",
															 "(Oracle)double",
															 "(Oracle)float",
															 "(Oracle)int",
															 "(Oracle)integer",
															 "(Oracle)numeric",
															 "(Oracle)number",
															 "(Oracle)real",
															 "(Oracle)smallint",
															 "(Oracle)char",
															 "(Oracle)varchar2",
															 "(Oracle)long",
															 "(Oracle)raw",
															 "(Oracle)long raw",
															 "(Oracle)date",
															 "(Oracle)timestamp",
															 "(Oracle)interval year",
															 "(Oracle)interval day",
															 "(Oracle)rowid",
															 "(Oracle)urowid",
															 "(Oracle)nchar",
															 "(Oracle)nvarchar2",
															 "(Oracle)bfile",
															 "(Oracle)blob",
															 "(Oracle)clob",
															 "(Oracle)nclob",
															 "(SqlServer)Binary",
															 "(SqlServer)Varbinary",
															 "(SqlServer)Char",
															 "(SqlServer)Varchar",
															 "(SqlServer)Nchar",
															 "(SqlServer)Nvarchar",
															 "(SqlServer)Datetime",
															 "(SqlServer)Smalldatetime",
															 "(SqlServer)Decimal",
															 "(SqlServer)Numeric",
															 "(SqlServer)Float",
															 "(SqlServer)Real",
															 "(SqlServer)Int",
															 "(SqlServer)Smallint",
															 "(SqlServer)Tinyint",
															 "(SqlServer)Money",
															 "(SqlServer)Smallmoney",
															 "(SqlServer)Bit",
															 "(SqlServer)Cursor",
															 "(SqlServer)Sysname",
															 "(SqlServer)Timestamp",
															 "(SqlServer)Uniqueidentifier",
															 "(SqlServer)Text",
															 "(SqlServer)Image",
															 "(SqlServer)Ntext"});
			this.comTypeUser.Location = new System.Drawing.Point(336, 24);
			this.comTypeUser.Name = "comTypeUser";
			this.comTypeUser.Size = new System.Drawing.Size(112, 20);
			this.comTypeUser.TabIndex = 24;
			this.comTypeUser.Text = "varchar2";
			this.comTypeUser.Leave += new System.EventHandler(this.comTypeUser_Leave);
			// 
			// comInitVal
			// 
			this.comInitVal.Items.AddRange(new object[] {
															"\"\"",
															"0",
															"-1",
															"null",
															"String.Empty",
															"new Date(\"2005/06/30\")"});
			this.comInitVal.Location = new System.Drawing.Point(336, 72);
			this.comInitVal.Name = "comInitVal";
			this.comInitVal.Size = new System.Drawing.Size(80, 20);
			this.comInitVal.TabIndex = 24;
			this.comInitVal.Text = "\"\"";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(7, 6);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(232, 16);
			this.label1.TabIndex = 27;
			this.label1.Text = "define various used by this prg.";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(279, 6);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(105, 16);
			this.label2.TabIndex = 27;
			this.label2.Text = "table type contrast:";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(416, 51);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(88, 16);
			this.label3.TabIndex = 27;
			this.label3.Text = "MaxVal method:";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(452, 28);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(96, 16);
			this.label4.TabIndex = 27;
			this.label4.Text = "To Type:";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(279, 28);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(96, 16);
			this.label5.TabIndex = 27;
			this.label5.Text = "User Type:";
			// 
			// label6
			// 
			this.label6.Location = new System.Drawing.Point(279, 52);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(96, 16);
			this.label6.TabIndex = 27;
			this.label6.Text = "Language:";
			// 
			// label7
			// 
			this.label7.Location = new System.Drawing.Point(279, 76);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(96, 16);
			this.label7.TabIndex = 27;
			this.label7.Text = "InitVal:";
			// 
			// label8
			// 
			this.label8.Location = new System.Drawing.Point(384, 5);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(219, 16);
			this.label8.TabIndex = 27;
			this.label8.Text = "(if empty MaxValue,then equals InitVal)";
			// 
			// comMaxVal
			// 
			this.comMaxVal.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comMaxVal.Items.AddRange(new object[] {
														   "char",
														   "int",
														   "=InitVal"});
			this.comMaxVal.Location = new System.Drawing.Point(504, 48);
			this.comMaxVal.Name = "comMaxVal";
			this.comMaxVal.Size = new System.Drawing.Size(72, 20);
			this.comMaxVal.TabIndex = 24;
			// 
			// label9
			// 
			this.label9.Location = new System.Drawing.Point(430, 74);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 16);
			this.label9.TabIndex = 27;
			this.label9.Text = "Show Length:";
			// 
			// txtDefaultLen
			// 
			this.txtDefaultLen.Location = new System.Drawing.Point(504, 72);
			this.txtDefaultLen.MaxLength = 9999;
			this.txtDefaultLen.Name = "txtDefaultLen";
			this.txtDefaultLen.Size = new System.Drawing.Size(72, 19);
			this.txtDefaultLen.TabIndex = 28;
			this.txtDefaultLen.Text = "";
			// 
			// PageVarious
			// 
			this.Controls.Add(this.txtDefaultLen);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.lstVar);
			this.Controls.Add(this.comTypeLang);
			this.Controls.Add(this.txtVarName);
			this.Controls.Add(this.txtVarLable);
			this.Controls.Add(this.btnVarAdd);
			this.Controls.Add(this.btnVarMod);
			this.Controls.Add(this.btnVarDel);
			this.Controls.Add(this.txtVarEdit);
			this.Controls.Add(this.textBox2);
			this.Controls.Add(this.lstType);
			this.Controls.Add(this.comTypeSys);
			this.Controls.Add(this.btnTypeMod);
			this.Controls.Add(this.btnTypeDel);
			this.Controls.Add(this.btnTypeAdd);
			this.Controls.Add(this.comTypeUser);
			this.Controls.Add(this.comInitVal);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.comMaxVal);
			this.Controls.Add(this.label9);
			this.Name = "PageVarious";
			this.Size = new System.Drawing.Size(632, 240);
			this.ResumeLayout(false);

		}
		#endregion

		public string VAR_SPLIT_STRING = " => ";
		public string VAR_SPLIT_STRING2 = "[;]";
		public void Config_Load(System.Collections.Specialized.NameValueCollection coll)
		{
			if(coll.Get("txtExcelTemp") != null && !coll.Get("txtExcelTemp").Equals(""))
			{
				//txtExcelTemp.Text = coll.Get("txtExcelTemp");
			}

			//add system various
			lstVar.Items.Clear();
			for(int i = 0; i < nvcSysVar.Count; i++)
			{
				lstVar.Items.Add(nvcSysVar.GetKey(i) + VAR_SPLIT_STRING + nvcSysVar[i]);
			}
			//restore excel file(template) list
			if(coll["VarList"] != null)
			{
				//if is defined by system and not has values,then replace it
				string[] slist = coll["VarList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					string line = slist[i];
					int npos = line.IndexOf(VAR_SPLIT_STRING);
					if(npos > 0)
					{
						string sName = line.Substring(0, npos);
						for(int j = 0; j < lstVar.Items.Count; j++)
						{
							if(lstVar.Items[j].ToString().StartsWith(sName + VAR_SPLIT_STRING))
							{
								if(lstVar.Items[j].ToString().Equals(sName + VAR_SPLIT_STRING))
								{
									lstVar.Items[j] = line;
								}
								line = "";
								break;
							}
						}
						if(!line.Equals(""))
						{
							lstVar.Items.Add(line);
						}
					}
				}
			}
			//restore excel file(table) list
			if(coll["TypeList"] != null)
			{
				lstType.Items.Clear();
				string[] slist = coll["TypeList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					string line = slist[i];
					int npos = line.IndexOf(VAR_SPLIT_STRING);
					if(npos > 0)
					{
						lstType.Items.Add(line);
					}
				}
			}
		}

		public void Config_Save(System.IO.StreamWriter sw)
		{
			//sw.Write("txtExcelTemp={0}\r\n", txtExcelTemp.Text);
			//save excel file(template) list
			for(int i = 0; i < lstVar.Items.Count; i++)
			{
				sw.Write("VarList={0}\r\n", lstVar.Items[i].ToString());
			}
			//save excel file(table) list
			for(int i = 0; i < lstType.Items.Count; i++)
			{
				sw.Write("TypeList={0}\r\n", lstType.Items[i].ToString());
			}
		}

		public System.Collections.Specialized.NameValueCollection UserVarious()
		{
			//add user defined various
			System.Collections.Specialized.NameValueCollection nvcUserVar = 
				new System.Collections.Specialized.NameValueCollection();
			for(int i = 0; i < lstVar.Items.Count; i++)
			{
				string str = lstVar.Items[i].ToString();
				int npos = str.IndexOf(VAR_SPLIT_STRING);
				if(npos > 0)
				{
					nvcUserVar[str.Substring(0, npos)] = str.Substring(npos + VAR_SPLIT_STRING.Length);
				}
			}
			return nvcUserVar;
		}

		//ad sys various
		private Hashtable UserTypeContrastSplite(string str)
		{
			//TypeUser,Lang,TypeSys,MaxVal,InitVal,Len
			Hashtable htbl = new Hashtable();
			int nlang = str.IndexOf("]");
			int npos1 = str.IndexOf(VAR_SPLIT_STRING);
			int npos2 = -1;
			int npos3 = -1;
			int npos4 = -1;
			if(npos1 > 0)
			{
				npos2 = str.IndexOf(VAR_SPLIT_STRING2, npos1 + 1);
			}
			if(npos2 > 0)
			{
				npos3 = str.IndexOf(VAR_SPLIT_STRING2, npos2 + 1);
			}
			if(npos3 > 0)
			{
				npos4 = str.IndexOf(VAR_SPLIT_STRING2, npos3 + 1);
			}
			if(nlang >= 0 && npos4 > 0)
			{
				htbl["usertype"] = str.Substring(nlang + 1, npos1 - nlang - 1);
				htbl["language"] = str.Substring(1, nlang - 1);
				htbl["totype"] = str.Substring(npos1 + VAR_SPLIT_STRING.Length, npos2 - npos1 - VAR_SPLIT_STRING.Length);
				htbl["maxvalue"] = str.Substring(npos2 + VAR_SPLIT_STRING2.Length, npos3 - npos2 - VAR_SPLIT_STRING2.Length);
				htbl["defaultvalue"] = str.Substring(npos3 + VAR_SPLIT_STRING2.Length, npos4 - npos3 - VAR_SPLIT_STRING2.Length);
				htbl["defaultlength"] = str.Substring(npos4 + VAR_SPLIT_STRING2.Length);
				return htbl;
			}
			return null;
		}

		public cc.Table UserTypeContrast()
		{
			//add user defined typies
			IPara ipara = new IPara();
			cc.Table tblUserType = ipara.tblUserType;

			for(int i = 0; i < lstType.Items.Count; i++)
			{
				Hashtable htbl = UserTypeContrastSplite(lstType.Items[i].ToString());
				if(htbl != null)
				{
					DataRow dr = tblUserType.NewRow();
					dr["usertype"] = htbl["usertype"];
					dr["language"] = htbl["language"].ToString().ToLower();
					dr["totype"] = htbl["totype"];
					dr["maxvalue"] = htbl["maxvalue"];
					dr["defaultvalue"] = htbl["defaultvalue"];
					dr["defaultlength"] = htbl["defaultlength"];
					tblUserType.Rows.Add(dr);
				}
			}
			return tblUserType;
		}

		private void lstVar_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if(lstVar.SelectedIndex >= 0)
			{
				string str = lstVar.Items[lstVar.SelectedIndex].ToString();
				int npos = str.IndexOf(VAR_SPLIT_STRING);
				txtVarName.Text = str.Substring(0, npos);
				txtVarEdit.Text = str.Substring(npos + VAR_SPLIT_STRING.Length);
			}
		}

		private void btnVarAdd_Click(object sender, System.EventArgs e)
		{
			string sName = txtVarName.Text.Trim();
			if(!sName.Equals(""))
			{
				for(int i = 0; i < lstVar.Items.Count; i++)
				{
					if(lstVar.Items[i].ToString().StartsWith(sName + VAR_SPLIT_STRING))
					{
						frmMain.labStatus.Text = "this Various is exist.";
						return;
					}
				}
				lstVar.Items.Add(sName + VAR_SPLIT_STRING + txtVarEdit.Text.Trim());
			}
			else
			{
				frmMain.labStatus.Text = "Various name can not be empty.";
			}
		}

		private void btnVarMod_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			if(lstVar.SelectedIndex < 0)
			{
				frmMain.labStatus.Text = "Please select one item first!";
				return;
			}

			string sName = txtVarName.Text.Trim();
			//if is defined by system various,can not modify name
			string str = lstVar.Items[lstVar.SelectedIndex].ToString();
			int npos = str.IndexOf(VAR_SPLIT_STRING);
			if(nvcSysVar[str.Substring(0, npos)] != null && !str.Substring(0, npos).Equals(sName))
			{
				frmMain.labStatus.Text = "this is system various,can not modify it's name.";
				sName = str.Substring(0, npos);
			}
			if(!sName.Equals(""))
			{
				for(int i = 0; i < lstVar.Items.Count; i++)
				{
					if(i != lstVar.SelectedIndex && lstVar.Items[i].ToString().StartsWith(sName + VAR_SPLIT_STRING))
					{
						frmMain.labStatus.Text = "this Various is exist.";
						return;
					}
				}
				lstVar.Items[lstVar.SelectedIndex] = sName + VAR_SPLIT_STRING + txtVarEdit.Text.Trim();
			}
			else
			{
				frmMain.labStatus.Text = "Various name can not be empty.";
			}
		}

		private void btnVarDel_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			if(lstVar.SelectedIndex < 0)
			{
				frmMain.labStatus.Text = "Please select one item first!";
				return;
			}
			string str = lstVar.Items[lstVar.SelectedIndex].ToString();
			int npos = str.IndexOf(VAR_SPLIT_STRING);
			if(nvcSysVar[str.Substring(0, npos)] != null)
			{
				frmMain.labStatus.Text = "this is system various,can not be delete.";
			}
			else
			{
				lstVar.Items.RemoveAt(lstVar.SelectedIndex);
			}
		}

		private void lstType_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			if(lstType.SelectedIndex >= 0)
			{
				Hashtable htbl = UserTypeContrastSplite(lstType.Items[lstType.SelectedIndex].ToString());
				if(htbl != null)
				{
					comTypeUser.Text = (string)htbl["usertype"];
					comTypeLang.Text = (string)htbl["language"];
					comTypeSys.Text = (string)htbl["totype"];

					string sMaxVal = (string)htbl["maxvalue"];
					comMaxVal.SelectedIndex = -1;
					for(int i = 0; i < comMaxVal.Items.Count; i++)
					{
						if(comMaxVal.Items[i].ToString().Equals(sMaxVal))
						{
							comMaxVal.SelectedIndex = i;
							break;
						}
					}
					comInitVal.Text = (string)htbl["defaultvalue"];
					txtDefaultLen.Text = (string)htbl["defaultlength"];
				}
			}
		}

		private void btnTypeAdd_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			string sName = comTypeUser.Text.Trim();
			string sLang = comTypeLang.Text.Trim();
			string sSys = comTypeSys.Text.Trim();
			string sLen = comMaxVal.Text.Trim();
			string sVal = comInitVal.Text.Trim();
			string sDefaultLen = txtDefaultLen.Text.Trim();
			string strKey = "[" + sLang + "]" + sName + VAR_SPLIT_STRING;
			string str = strKey + sSys + VAR_SPLIT_STRING2 + sLen + VAR_SPLIT_STRING2 + sVal + VAR_SPLIT_STRING2 + sDefaultLen;
			if(sName.Equals("") || sLang.Equals("") || sSys.Equals(""))
			{
				frmMain.labStatus.Text = "define for type contrast need user 'type,language,to type'";
			}
			else
			{
				for(int i = 0; i < lstType.Items.Count; i++)
				{
					if(lstType.Items[i].ToString().StartsWith(strKey))
					{
						frmMain.labStatus.Text = "has defined for this.";
						return;
					}
				}
				lstType.Items.Add(str);
			}
		}

		private void btnTypeMod_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			if(lstType.SelectedIndex < 0)
			{
				frmMain.labStatus.Text = "Please select one item first!";
				return;
			}

			string sName = comTypeUser.Text.Trim();
			string sLang = comTypeLang.Text.Trim();
			string sSys = comTypeSys.Text.Trim();
			string sLen = comMaxVal.Text.Trim();
			string sVal = comInitVal.Text.Trim();
			string sDefaultLen = txtDefaultLen.Text.Trim();
			string strKey = "[" + sLang + "]" + sName + VAR_SPLIT_STRING;
			string str = strKey + sSys + VAR_SPLIT_STRING2 + sLen + VAR_SPLIT_STRING2 + sVal + VAR_SPLIT_STRING2 + sDefaultLen;
			if(sName.Equals("") || sLang.Equals("") || sSys.Equals(""))
			{
				frmMain.labStatus.Text = "define for type contrast need user 'type,language,to type'";
			}
			else
			{
				for(int i = 0; i < lstType.Items.Count; i++)
				{
					if(i != lstType.SelectedIndex && lstType.Items[i].ToString().StartsWith(strKey))
					{
						frmMain.labStatus.Text = "has defined for this.";
						return;
					}
				}
				lstType.Items[lstType.SelectedIndex] = str;
			}
		}

		private void btnTypeDel_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			int nindex = lstType.SelectedIndex;
			if(nindex < 0)
			{
				frmMain.labStatus.Text = "Please select one item first!";
				return;
			}
			lstType.Items.RemoveAt(lstType.SelectedIndex);
			if(lstType.Items.Count > 0)
			{
				if(nindex < lstType.Items.Count)
				{
					lstType.SetSelected(nindex, true);
				}
				else
				{
					lstType.SetSelected(lstType.Items.Count - 1, true);
				}
			}
		}

		private void comTypeUser_Leave(object sender, System.EventArgs e)
		{
			if(comTypeUser.SelectedIndex < 0)
			{
				return;
			}
			string str = comTypeUser.Items[comTypeUser.SelectedIndex].ToString().ToLower();
			comTypeUser.Text = str.Substring(str.IndexOf(")") + 1);
		}

		private void comTypeLang_Leave(object sender, System.EventArgs e)
		{
			if(comTypeUser.SelectedIndex < 0)
			{
				return;
			}
			comTypeUser.Text = comTypeUser.Items[comTypeUser.SelectedIndex].ToString().ToLower();
		}

	}
}
