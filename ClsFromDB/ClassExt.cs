using System;
using System.Drawing;
using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Xml;
using System.Text;
using System.Data.SqlClient;

namespace ClsFromDB
{
	/// <summary>
	/// IPara
	/// </summary>
	public class IPara
	{
		private string sOutPath = "c:\\temp\\";

		public cc.Msg msg = null;
		public cc.Table tblField;
		public cc.Table tblUserType;
		//when search excel type is not defined in tblUserType,
		//then save the result to UserTypeToOutType
		//[language]excel_type
		public NameValueCollection UserTypeToOutType = new NameValueCollection();
		public NameValueCollection UserVarious = new NameValueCollection();
		public Hashtable SystemVarious = new Hashtable();
		public XmlNode TemplateNode = null;
		public int TemplateCount = -1;
		public int TemplateIndex = -1;
		public int TableIndex = -1;
		public int TableCount = -1;
		public NameValueCollection TemplateFieldTexts = new NameValueCollection();

		public string TemplateFieldText = null;
		public int TemplateFieldIndex = -1;
		public int TemplateFieldCount = -1;
		public int TableFieldIndex = -1;
		public int TableFieldCount = -1;
		public string TemplateMainText = null;

		public IPara()
		{
			tblField = new cc.Table();
			tblField.Columns.Add("FIELD_TYPE_EXCEL"); //get from excel file or server
			tblField.Columns.Add("FIELD_TYPE_INT_OR_CHAR"); //for calcute max value,int or char or =InitVal
			tblField.Columns.Add("FIELD_NAME");
			tblField.Columns.Add("FIELD_INGETER");
			tblField.Columns.Add("FIELD_DECIMAL");
			tblField.Columns.Add("FIELD_INGETER_DECIMAL"); //(INGETER) or (INGETER, DECIMAL)
			tblField.Columns.Add("FIELD_INGETER_PLUS_DECIMAL"); //INGETER+DECIMAL
			tblField.Columns.Add("FIELD_SHOW_LENGTH"); //INGETER+DECIMAL,if 0,then use defaultlength
			tblField.Columns.Add("FIELD_TYPE"); //for output,user need type(java,c#,sqlscript,...)
			tblField.Columns.Add("FIELD_VALUE_INIT");
			tblField.Columns.Add("FIELD_VALUE_MIDDLE");
			tblField.Columns.Add("FIELD_VALUE_MAX");
			tblField.Columns.Add("FIELD_VALUE_INIT_NO_QUOTES");
			tblField.Columns.Add("FIELD_VALUE_MIDDLE_NO_QUOTES");
			tblField.Columns.Add("FIELD_VALUE_MAX_NO_QUOTES");

			tblUserType = new cc.Table();
			tblUserType.Columns.Add("usertype");
			tblUserType.Columns.Add("language");
			tblUserType.Columns.Add("totype");
			tblUserType.Columns.Add("maxvalue");
			tblUserType.Columns.Add("defaultvalue");
			tblUserType.Columns.Add("defaultlength");
		}

		/// <summary>
		/// OutPath
		/// </summary>
		public string OutPath
		{
			get
			{
				return sOutPath;
			}
			set
			{
				sOutPath = value;
				if(sOutPath == null || sOutPath.Equals(""))
				{
					sOutPath = "c:\\temp\\";
				}
				if(!sOutPath.EndsWith("\\"))
				{
					sOutPath += "\\";
				}
			}
		}

	}

	/// <summary>
	/// HashXY
	/// </summary>
	public class HashXY
	{
		/// <summary>
		/// Point
		/// </summary>
		public class XY
		{
			private int x;
			private int y;
			private string text;

			public XY(int x, int y)
			{
				this.x = x;
				this.y = y;
				this.text = null;
			}

			public int X
			{
				get
				{
					return x;
				}
				set
				{
					x = value;
				}
			}

			public int Y
			{
				get
				{
					return y;
				}
				set
				{
					y = value;
				}
			}

			public string Text
			{
				get
				{
					return text;
				}
				set
				{
					text = value;
				}
			}

		}

		private Hashtable htbl = new Hashtable();

		/// <summary>
		/// get member
		/// </summary>
		public XY this[string oKey]
		{
			get
			{
				return (XY)htbl[oKey];
			}
			set
			{
				htbl[oKey] = value;
			}
		}

		/// <summary>
		/// Hashtable
		/// </summary>
		public Hashtable Hashtable
		{
			get
			{
				return htbl;
			}
		}

	}

	/// <summary>
	/// Template's Ext Class
	/// </summary>
	public class ClassExt
	{
		public ClassExt()
		{
		}

		//for define COMMENT,PK,NULL...
		static public string sEXCEL_VARIOUS = "[#USER_VARIOUS#]";
		static public string sFIELD_VARIOUS = "[#FIELD_VARIOUS#]";
		static public string sFIELD_END = "[#FIELD_END_IF_EQUALS#]";
		/// <summary>
		/// getExcelTempInfo,get position of defined item and value
		/// </summary>
		/// <param name="excelFile">System.Data.DataTable tbl</param>
		/// <returns>Hashtable(Object:new int[]{x,y})</returns>
		static public HashXY GetExcelTempInfo(System.Data.DataTable tblExcelTemp)
		{
			if(tblExcelTemp == null)
			{
				return null;
			}

			HashXY hashxy = new HashXY();
			hashxy["[#DB_NAME#]"] = null;
			hashxy["[#TABLE_NAME#]"] = null;
			hashxy["[#FIELD_NAME#]"] = null;
			hashxy["[#FIELD_TYPE#]"] = null;
			hashxy["[#FIELD_TYPE_INGETER_DECIMAL#]"] = null;
			hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"] = null;
			hashxy["[#FIELD_INGETER#]"] = null;
			hashxy["[#FIELD_DECIMAL#]"] = null;
			hashxy["[#FIELD_INGETER_DECIMAL#]"] = null;
			hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"] = null;

			hashxy["[#FIELD_START_Y#]"] = new HashXY.XY(-1, -1);
			for(int y = 0; y < tblExcelTemp.Rows.Count; y++)
			{
				for(int x = 0; x < tblExcelTemp.Columns.Count; x++)
				{
					string sCell = "" + tblExcelTemp.Rows[y][x];
					sCell = sCell.Trim();
					if(!sCell.Equals("") && sCell.StartsWith("[#"))
					{
						if(hashxy.Hashtable.ContainsKey(sCell))
						{
							hashxy[sCell] = new HashXY.XY(x,y);
							hashxy[sCell].Text = null;
						}
						else
						{
							if(sCell.StartsWith(sFIELD_VARIOUS) && sCell.Length > sFIELD_VARIOUS.Length)
							{
								//sFIELD_VARIOUS:get field various, this info only used in XML's "field" tag.
								hashxy[sCell] = new HashXY.XY(x,y);
								hashxy[sCell].Text = sCell.Substring(sFIELD_VARIOUS.Length).Trim();
							}
							else if(sCell.StartsWith(sEXCEL_VARIOUS) && sCell.Length > sEXCEL_VARIOUS.Length)
							{
								//sTABLE_VARIOUS:get table various
								hashxy[sCell] = new HashXY.XY(x,y);
								hashxy[sCell].Text = sCell.Substring(sFIELD_VARIOUS.Length).Trim();
							}
							else if(sCell.StartsWith(sFIELD_END))
							{
								//sFIELD_END:if cell'value = this,then end field loop
								hashxy["[#FIELD_START_Y#]"] = new HashXY.XY(x, y);
								hashxy["[#FIELD_START_Y#]"].Text = sCell.Substring(sFIELD_END.Length).Trim();
							}
						}
					}
				}
			}

			bool hasFIELD_NAME = false;
			bool hasFIELD_TYPE = false;
			bool hasFIELD_INGETER = false;
			bool hasFIELD_DECIMAL = false;
			if(hashxy["[#FIELD_NAME#]"] != null)
			{
				hasFIELD_NAME = true;
			}
			if(hashxy["[#FIELD_TYPE#]"] != null)
			{
				hasFIELD_TYPE = true;
			}
			if(hashxy["[#FIELD_TYPE_INGETER_DECIMAL#]"] != null || hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"] != null)
			{
				hasFIELD_TYPE = true;
				hasFIELD_INGETER = true;
				hasFIELD_DECIMAL = true;
			}
			if(hashxy["[#FIELD_INGETER#]"] != null)
			{
				hasFIELD_INGETER = true;
			}
			if(hashxy["[#FIELD_DECIMAL#]"] != null)
			{
				hasFIELD_DECIMAL = true;
			}
			if(hashxy["[#FIELD_INGETER_DECIMAL#]"] != null || hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"] != null)
			{
				hasFIELD_INGETER = true;
				hasFIELD_DECIMAL = true;
			}
			int nFieldStart_Y = -1;
			foreach(string sKey in hashxy.Hashtable.Keys)
			{
				if(sKey.StartsWith("[#FIELD_") && hashxy[sKey] != null)
				{
					if(nFieldStart_Y == -1)
					{
						nFieldStart_Y = hashxy[sKey].Y;
					}
					else
					{
						if(nFieldStart_Y != hashxy[sKey].Y)
						{
							nFieldStart_Y = -2;
							break;
						}
					}
				}
			}
			//nFieldStart_Y = -1:not find any define for field
			//nFieldStart_Y = -2:field define is not at same line
			//nFieldStart_Y = -3:at least needed field is not defined
			hashxy["[#FIELD_START_Y#]"].Y = nFieldStart_Y;
			if(!(hasFIELD_NAME && hasFIELD_TYPE && hasFIELD_INGETER && hasFIELD_DECIMAL))
			{
				hashxy["[#FIELD_START_Y#]"].Y = -3;
			}
			return hashxy;
		}

		/// <summary>
		/// return String created from tablelayout and template.
		/// </summary>
		/// <param name="excelFile">XmlNode node</param>
		/// <returns>String</returns>
		static public string GetTBLInfoFromExcel(HashXY hashxy, 
			System.Data.DataTable tblSheet, IPara ipara)
		{
			if(hashxy == null || tblSheet == null || ipara == null)
			{
				return "need excel template info and excel sheet data and other need info in function GetTBLInfoFromExcel";
			}
			if(hashxy["[#FIELD_START_Y#]"] == null || hashxy["[#FIELD_START_Y#]"].Y == -1 || hashxy["[#FIELD_START_Y#]"].Y == -3)
			{
				return "at least need define:FIELD_NAME,FIELD_TYPE,FIELD_INGETER,FIELD_DECIMAL";
			}
			if(hashxy["[#FIELD_START_Y#]"].Y == -2)
			{
				return "need define any field info at same line,also:[#FIELD_END_IF_EQUALS#],[#FIELD_VARIOUS#]...";
			}

			//need use new ITable
			IPara ipara2 = new IPara();
			cc.Table tblField = ipara2.tblField;
			int nFieldStart = hashxy["[#FIELD_START_Y#]"].Y;
			string sFieldEndString = hashxy["[#FIELD_START_Y#]"].Text;
			//for judge end of field:if FIELD_NAME is empty,then end.
			//but if defined [#FIELD_END_IF_EQUALS#],then if cell'value=(defined),then end
			int nFieldEndX;
			if(sFieldEndString == null)
			{
				nFieldEndX = hashxy["[#FIELD_NAME#]"].X;
				sFieldEndString = "";
			}
			else
			{
				nFieldEndX = hashxy["[#FIELD_START_Y#]"].X;
			}
			//add sFIELD_VARIOUS to tblField'Column and add table various(GlobalVarious)
			foreach(string sKey in hashxy.Hashtable.Keys)
			{
				if(sKey.StartsWith(sFIELD_VARIOUS))
				{
					//add new colum
					tblField.Columns.Add(hashxy[sKey].Text);
				}
				else if(sKey.StartsWith(sEXCEL_VARIOUS))
				{
					//add table various(GlobalVarious)
					ipara.UserVarious["[#" + sKey.Substring(sEXCEL_VARIOUS.Length) + "#]"] = GetSheetData(tblSheet, hashxy[sKey].X, hashxy[sKey].Y);
				}
			}
			//loop for treate every rows*cols
			for(int loopi = nFieldStart; loopi < tblSheet.Rows.Count; loopi++)
			{
				string fieldName = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_NAME#]"].X]).ToString().Trim();
				string fieldType = null;
				string fieldInt = null;
				string fieldDec = null;
				if(fieldName.Equals("") || tblSheet.Rows[loopi][nFieldEndX].ToString().Trim().Equals(sFieldEndString))
				{
					break;
				}

				//get field type
				if(hashxy["[#FIELD_TYPE#]"] != null)
				{
					fieldType = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_TYPE#]"].X]).ToString().Trim();
				}

				//get field length int
				if(hashxy["[#FIELD_INGETER#]"] != null)
				{
					fieldInt = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_INGETER#]"].X]).ToString().Trim();
				}

				//get field length dec
				if(hashxy["[#FIELD_DECIMAL#]"] != null)
				{
					fieldDec = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_DECIMAL#]"].X]).ToString().Trim();
				}

				//FIELD_TYPE_INGETER_DECIMAL,FIELD_TYPE_INGETERDECIMAL_DECIMAL
				if((fieldType == null || fieldInt == null || fieldDec == null) 
					&& (hashxy["[#FIELD_TYPE_INGETER_DECIMAL#]"] != null || hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"] != null))
				{
					string strcellval = null;
					if(hashxy["[#FIELD_TYPE_INGETER_DECIMAL#]"] != null)
					{
						strcellval = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_TYPE_INGETER_DECIMAL#]"].X]).ToString().Trim();
					}
					if(hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"] != null)
					{
						strcellval = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"].X]).ToString().Trim();
					}
					string[] strcellarr = SplitTypeLen(strcellval);
					if(hashxy["[#FIELD_TYPE_INGETERDECIMAL_DECIMAL#]"] != null)
					{
						//INGETER = INGETERDECIMAL - DECIMAL
						strcellarr[1] = "" + (cc.Util.toInt(strcellarr[1]) - cc.Util.toInt(strcellarr[2]));
					}
					if(fieldType == null)
					{
						fieldType = strcellarr[0];
					}
					if(fieldInt == null)
					{
						fieldInt = strcellarr[1];
					}
					if(fieldDec == null)
					{
						fieldDec = strcellarr[2];
					}
				}

				//FIELD_INGETER_DECIMAL,FIELD_INGETERDECIMAL_DECIMAL
				if((fieldInt == null || fieldDec == null) 
					&& (hashxy["[#FIELD_INGETER_DECIMAL#]"] != null || hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"] != null))
				{
					string strcellval = null;
					if(hashxy["[#FIELD_INGETER_DECIMAL#]"] != null)
					{
						strcellval = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_INGETER_DECIMAL#]"].X]).ToString().Trim();
					}
					if(hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"] != null)
					{
						strcellval = ("" + tblSheet.Rows[loopi][hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"].X]).ToString().Trim();
					}
					string[] strcellarr = SplitTypeLen(strcellval);
					if(hashxy["[#FIELD_INGETERDECIMAL_DECIMAL#]"] != null)
					{
						//INGETER = INGETERDECIMAL - DECIMAL
						strcellarr[0] = "" + (cc.Util.toInt(strcellarr[0]) - cc.Util.toInt(strcellarr[1]));
					}
					if(fieldInt == null)
					{
						fieldInt = strcellarr[0];
					}
					if(fieldDec == null)
					{
						fieldDec = strcellarr[1];
					}
				}

				//add base various(column),defined by system
				DataRow curRow = tblField.NewRow();
				curRow["FIELD_NAME"] = fieldName;
				curRow["FIELD_TYPE_EXCEL"] = fieldType; //get from excel file or server
				curRow["FIELD_INGETER"] = fieldInt;
				curRow["FIELD_DECIMAL"] = fieldDec;
				tblField.Rows.Add(curRow);

				//add UserDefined various(sFIELD_VARIOUS) in excel
				foreach(string sKey in hashxy.Hashtable.Keys)
				{
					if(sKey.StartsWith(sFIELD_VARIOUS))
					{
						//add field value
						curRow[hashxy[sKey].Text] = ("" + tblSheet.Rows[loopi][hashxy[sKey].X]).ToString().Trim();
					}
				}
			}
			if(hashxy["[#DB_NAME#]"] != null)
			{
				ipara.UserVarious["[#DB_NAME#]"] = GetSheetData(tblSheet, 
					hashxy["[#DB_NAME#]"].X, hashxy["[#DB_NAME#]"].Y);
			}
			if(hashxy["[#TABLE_NAME#]"] != null)
			{
				ipara.UserVarious["[#TABLE_NAME#]"] = GetSheetData(tblSheet, 
					hashxy["[#TABLE_NAME#]"].X, hashxy["[#TABLE_NAME#]"].Y);
			}
			ipara.tblField = tblField;
			return null;
		}

		/// <summary>
		/// return String created from tablelayout and template.
		/// </summary>
		/// <param name="excelFile">XmlNode node</param>
		/// <returns>String</returns>
		static public string GetTBLInfoFromServer(cc.DB cdb, IPara ipara)
		{
			/*
			* SqlServer
			127	[f1] [bigint]
			173	[f2] [binary] (50)
			104	[f3] [bit]
			175	[f4] [char]
			61	[f5] [datetime]
			106	[f6] [decimal](18, 0)
			62	[f7] [float]
			34	[f8] [image]
			56	[f9] [int]
			60	[f10] [money]
			239	[f11] [nchar] (10) 
			99	[f12] [ntext] 
			108	[f13] [numeric](18, 0)
			231	[f14] [nvarchar] (50)
			59	[f15] [real]
			58	[f16] [smalldatetime]
			52	[f17] [smallint]
			122	[f18] [smallmoney]
			98	[f19] [sql_variant]
			35	[f20] [text] 
			189	[f21] [timestamp]
			48	[f22] [tinyint]
			36	[f23] [uniqueidentifier]
			165	[f24] [varbinary] (50)
			167	[f25] [varchar] (50) 
			*/
			string sdb = ipara.UserVarious["[#DB_NAME#]"];
			string stbl = ipara.UserVarious["[#TABLE_NAME#]"];

			//need use new ITable
			IPara ipara2 = new IPara();
			cc.Table tblField = ipara2.tblField;
			const string SQLDBObjectProperties =  "select C.* from {0}.dbo.sysobjects O join {0}.dbo.syscolumns C on C.id = O.id where O.name = '{1}' and (O.xtype ='U' or O.xtype ='V') order by C.name";
			try
			{
				DataTable tbl = cdb.GetTable(String.Format(SQLDBObjectProperties, sdb, stbl));
				if(cdb.Error())
				{
					return "Table InformationÇÃéÊìæÇ…Ç≈Ç´Ç‹ÇπÇÒÇ≈ÇµÇΩÅB";
				}

				//TBL info
				for(int i = 0; i < tbl.Rows.Count; i++)
				{
					string sName = "" + tbl.Rows[i]["name"];
					string sType = "" + tbl.Rows[i]["xtype"];
					switch(sType)
					{
						case "127":
							sType = "bigint";
							break;
						case "173":
							sType = "binary";
							break;
						case "104":
							sType = "bit";
							break;
						case "175":
							sType = "char";
							break;
						case "61":
							sType = "datetime";
							break;
						case "106":
							sType = "decimal";
							break;
						case "62":
							sType = "float";
							break;
						case "34":
							sType = "image";
							break;
						case "56":
							sType = "int";
							break;
						case "60":
							sType = "money";
							break;
						case "239":
							sType = "nchar";
							break;
						case "99":
							sType = "ntext";
							break;
						case "108":
							sType = "numeric";
							break;
						case "231":
							sType = "nvarchar";
							break;
						case "59":
							sType = "real";
							break;
						case "58":
							sType = "smalldatetime";
							break;
						case "52":
							sType = "smallint";
							break;
						case "122":
							sType = "smallmoney";
							break;
						case "98":
							sType = "sql_variant";
							break;
						case "35":
							sType = "text";
							break;
						case "189":
							sType = "timestamp";
							break;
						case "48":
							sType = "tinyint";
							break;
						case "36":
							sType = "uniqueidentifier";
							break;
						case "165":
							sType = "varbinary";
							break;
						case "167":
							sType = "varchar";
							break;
						default:
							sType = "Unknown" + sType;
							break;
					}
					string sInt = "" + tbl.Rows[i]["length"];

					//TODO

					string sDec = "" + tbl.Rows[i]["length"];
					//add base various(column),defined by system
					DataRow curRow = tblField.NewRow();
					curRow["FIELD_NAME"] = sName;
					curRow["FIELD_TYPE_EXCEL"] = sType; //get from excel file or server
					curRow["FIELD_INGETER"] = sInt;
					curRow["FIELD_DECIMAL"] = sDec;
					tblField.Rows.Add(curRow);
				}
			}
			catch
			{
				return "TBLèÓïÒÇÃéÊìæÇ…Ç≈Ç´Ç‹ÇπÇÒÇ≈ÇµÇΩÅB";
			}
			ipara.tblField = tblField;
			return null;
		}

		//return sheet.cells(x, y)
		static public string GetSheetData(System.Data.DataTable tblSheet, int x, int y)
		{
			if(x >= tblSheet.Columns.Count)
			{
				return "";
			}
			if(y >= tblSheet.Rows.Count)
			{
				return "";
			}
			if(tblSheet.Rows[y][x] == null)
			{
				return "";
			}
			else
			{
				return ("" + tblSheet.Rows[y][x]).ToString().Trim();
			}
		}

		//split string by not IsLetterOrDigit&"_"&"-"
		static public string[] SplitTypeLen(string str)
		{
			string[] retu = new string[]{"", "", ""};
			if(str == null)
			{
				return retu;
			}
			str = str.Trim();
			bool isStart = false;
			int arrCnt = 0;
			int nStart = 0;
			for(int i = 0; i < str.Length + 2; i++)
			{
				if(i < str.Length && (char.IsLetterOrDigit(str[i]) || str[i] == '_' || str[i] == '-'))
				{
					isStart = true;
					if(nStart == -1)
					{
						nStart = i;
					}
					continue;
				}
				else
				{
					if(isStart || i >= str.Length - 1)
					{
						retu[arrCnt] = str.Substring(nStart, i - nStart);
						arrCnt++;
						nStart = -1;
						if(arrCnt > 2 || i >= str.Length - 1)
						{
							break;
						}
					}
					isStart = false;
				}
			}
			return retu;
		}

		//replace var,and replace "[!LOWER!]","[!UPPER!]","[!FIRSTUPPER!]"
		static public string ReplaceVar(string str, string sfrom, string sto)
		{
			if(str == null)
			{
				return "";
			}
			if(sfrom == null || sto == null || sfrom.Equals(""))
			{
				return str;
			}
			str = str.Replace(sfrom + "[!LOWER!]", sto.ToLower());
			str = str.Replace(sfrom + "[!UPPER!]", sto.ToUpper());
			if(sto.Length > 1)
			{
				str = str.Replace(sfrom + "[!FIRSTUPPER!]", sto.Substring(0, 1).ToUpper() + sto.Substring(1).ToLower());
			}
			else
			{
				str = str.Replace(sfrom + "[!FIRSTUPPER!]", sto.ToUpper());
			}
			str = str.Replace(sfrom, sto);
			return str;
		}

		//ChangeVar var,and change to "[!LOWER!]","[!UPPER!]","[!FIRSTUPPER!]"
		static public string ChangeVar(string str, string changetype)
		{
			if(str == null)
			{
				return "";
			}
			changetype = changetype.Trim().ToLower();
			if(changetype.Equals("lower"))
			{
				str = str.ToLower();
			}
			else if(changetype.Equals("upper"))
			{
				str = str.ToUpper();
			}
			else if(changetype.Equals("firstupper"))
			{
				if(str.Length > 1)
				{
					str = str.Substring(0, 1).ToUpper() + str.Substring(1).ToLower();
				}
				else
				{
					str = str.ToUpper();
				}
			}
			return str;
		}

		//return Encode for JScript
		static public string JSEncode(string sTxt)
		{
			if(sTxt == null)
			{
				return "";
			}
			return sTxt.Replace("\\", "\\\\").Replace("\r", "").Replace("\n", "\\r\\n").Replace("'", "\\'");
		}

		//for JS EVAL:start:[!JS!],end:[!JS_END!]
		static public string Eval(IPara ipara, System.Text.StringBuilder sbJSTxt, ref string sTxt, string sJSTxtEnd)
		{
			//JS EVAL
			string ErrorString = null;
			int nOldLen = sbJSTxt.Length;
			int nEval = sTxt.IndexOf("[!JS!]");
			int nOldPos = 0;
			//get all JScript
			while(nEval >= 0)
			{
				int nEval2 = sTxt.IndexOf("[!JS_END!]", nEval + 6);
				if(nEval2 > nEval)
				{
					if(nEval - nOldPos > 0)
					{
						sbJSTxt.Append("writetxt('" + JSEncode(sTxt.Substring(nOldPos, nEval - nOldPos)) + "');\r\n");
					}
					nOldPos = nEval2 + 10;
					sbJSTxt.Append(sTxt.Substring(nEval + 6, nEval2 - nEval - 6));
					nEval = sTxt.IndexOf("[!JS!]", nEval2 + 10);
				}
				else
				{
					return "found [!JS!],but not right with [!JS_END!]";
				}
			}
			//if found [!JS!] and [!JS_END!],then evel it
			if(sbJSTxt.Length > nOldLen)
			{
				sbJSTxt.Append("writetxt('" + JSEncode(sTxt.Substring(nOldPos)) + "');\r\n");
				sbJSTxt.Append(sJSTxtEnd);
				sbJSTxt.Append("\r\njs = js;\r\n");
				Microsoft.JScript.ScriptObject jsobj = (Microsoft.JScript.ScriptObject)cc.Eval.JSEvaluateToObject(sbJSTxt.ToString(), true, out ErrorString);
				if(ErrorString != null)
				{
					return ErrorString;
				}
				if(jsobj == null)
				{
					return "Run JScript Error:no value return.";
				}
				object jswritetxt = jsobj["writetxt"];
				if(jswritetxt != null && jswritetxt.GetType().FullName.Equals("Microsoft.JScript.ConcatString"))
				{
					sTxt = "" + jswritetxt;
				}
				else
				{
					sTxt = null;
				}
				object jsinfo = jsobj["info"];
				if(jsinfo != null && jsinfo.GetType().FullName.Equals("Microsoft.JScript.JSObject"))
				{
					Microsoft.JScript.JSObject jsobjinfo = (Microsoft.JScript.JSObject)jsinfo;
					if(jsobjinfo["TEMPLATE_SUBDIR"] != null)
					{
						ipara.UserVarious["[#TEMPLATE_SUBDIR#]"] = jsobjinfo["TEMPLATE_SUBDIR"].ToString();
					}
					if(jsobjinfo["TEMPLATE_FILENAME"] != null)
					{
						ipara.UserVarious["[#TEMPLATE_FILENAME#]"] = jsobjinfo["TEMPLATE_FILENAME"].ToString();
					}
				}
				object jsmsg = jsobj["writemsg"];
				if(jsmsg != null && !("" + jsmsg).Equals(""))
				{
					return "Run JScript Message:" + jsmsg;
				}
			}
			return null;
		}

		/// <summary>
		/// return String created from tablelayout and template.
		/// </summary>
		/// <param name="excelFile">XmlNode node</param>
		/// <returns>String</returns>
		static public string CreateClsFromTemp(IPara ipara)
		{
			XmlNode xmlnode = ipara.TemplateNode;
			cc.Msg msg = ipara.msg;
			string outpath = ipara.OutPath;
			if(xmlnode == null)
			{
				return "xmlnode is null";
			}
			if(msg == null)
			{
				return "msg is null";
			}
			if(outpath == null || outpath.Equals(""))
			{
				return "need define outpath";
			}
			if(!outpath.EndsWith("\\"))
			{
				outpath += "\\";
			}
			string clsdb = ipara.UserVarious["[#DB_NAME#]"];
			if(clsdb == null || clsdb.Equals(""))
			{
				clsdb = "DB_NAME_NOT_DEFINED";
			}
			string clstbl = ipara.UserVarious["[#TABLE_NAME#]"];
			if(clstbl == null || clstbl.Equals(""))
			{
				clstbl = "TBL_NAME_NOT_DEFINED";
			}

			if(xmlnode.Attributes["subdir"] != null)
			{
				string subdir = xmlnode.Attributes["subdir"].InnerText;
				//turn dbname and tblname to various
				subdir = ReplaceVar(subdir, "[#DB_NAME#]", clsdb);
				subdir = ReplaceVar(subdir, "[#TABLE_NAME#]", clstbl);
				ipara.UserVarious["[#TEMPLATE_SUBDIR#]"] = subdir;
			}
			string filename = "";
			if(xmlnode.Attributes["filename"] != null)
			{
				filename = xmlnode.Attributes["filename"].InnerText;
				//turn dbname and tblname to various
				filename = ReplaceVar(filename, "[#DB_NAME#]", clsdb);
				filename = ReplaceVar(filename, "[#TABLE_NAME#]", clstbl);
			}
			if(filename.Equals(""))
			{
				return "no out filename is defined.";
			}
			ipara.UserVarious["[#TEMPLATE_FILENAME#]"] = filename;

			//default name,change for lower or upper defined
			if(xmlnode.Attributes["dbnametype"] != null)
			{
				ipara.UserVarious["[#DB_NAME#]"] = ChangeVar(clsdb, xmlnode.Attributes["dbnametype"].InnerText);
			}
			if(xmlnode.Attributes["tablenametype"] != null)
			{
				ipara.UserVarious["[#TABLE_NAME#]"] = ChangeVar(clstbl, xmlnode.Attributes["tablenametype"].InnerText);
			}
			if(xmlnode.Attributes["fieldnametype"] != null)
			{
				string changenametype = xmlnode.Attributes["fieldnametype"].InnerText;
				for(int i = 0; i < ipara.tblField.Rows.Count; i++)
				{
					ipara.tblField[i, "FIELD_NAME"] = ChangeVar(ipara.tblField[i, "FIELD_NAME"], changenametype);
				}
			}

			try
			{
				//Templateïœä∑
				string sFileTxt = CreateClsFromTempString(ipara);
				string subdir2 = ipara.UserVarious["[#TEMPLATE_SUBDIR#]"];
				string filename2 = ipara.UserVarious["[#TEMPLATE_FILENAME#]"];
				if(sFileTxt != null && !filename2.Equals(""))
				{
					string clsfullname = outpath;
					if(!subdir2.Equals(""))
					{
						if(cc.Util.create_subdir(outpath, subdir2, ref clsfullname) != null)
						{
							return "can not create subdir defined in XML template:" + subdir2;
						}
					}
					clsfullname += filename2;
					msg.println("  " + clsfullname);
					System.IO.StreamWriter sw = new System.IO.StreamWriter(clsfullname, false, System.Text.Encoding.Default);
					sw.Write(sFileTxt);
					sw.Close();
				}
				return null;
			}
			catch(Exception exp)
			{
				return exp.Message;
			}
		}

		/// <summary>
		/// return IPara ipara,and treate every field and add JScript to ipara.
		/// </summary>
		/// <param name="excelFile">IPara ipara</param>
		/// <returns>IPara ipara</returns>
		static public IPara TreateFieldANDJScript(IPara ipara)
		{
			XmlNode xmlnode = ipara.TemplateNode;
			cc.Table tblField = ipara.tblField;
			cc.Table tblUserType = ipara.tblUserType;
			cc.Msg msg = ipara.msg;
			string language = "";
			if(xmlnode != null && xmlnode.Attributes["language"] != null)
			{
				language = xmlnode.Attributes["language"].InnerText.ToLower();
			}
			else
			{
				language = "default";
			}
			ipara.UserVarious["[#TEMPLATE_LANGUAGE#]"] = language;

			//change excel or server FIELD_TYPE to user "table type contrast:"
			NameValueCollection UserTypeToOutType = ipara.UserTypeToOutType;
			for(int fieldi = 0; fieldi < tblField.Rows.Count; fieldi++)
			{
				string sType = tblField[fieldi, "FIELD_TYPE_EXCEL"];
				string sLangType = "[" + language + "]" + sType;
				int nUserType = -1;
				if(UserTypeToOutType[sLangType] != null)
				{
					nUserType = int.Parse(UserTypeToOutType[sLangType]);
				}
				else
				{
					//if same to sToType and language,then use it,otherwise use language:Default
					for(int vari = 0; vari < tblUserType.Rows.Count; vari++)
					{
						string sUserType = tblUserType[vari, "usertype"];
						string sLang = tblUserType[vari, "language"];
						if(sType.Equals(sUserType))
						{
							if(language.Equals(sLang))
							{
								nUserType = vari;
								break;
							}
							else
							{
								if(sLang.Equals("default"))
								{
									nUserType = vari;
								}
							}
						}
					}
					UserTypeToOutType[sLangType] = "" + nUserType;
					if(nUserType < 0)
					{
						msg.println(" Waring:not defined type:" + sLangType);
					}
				}

				if(nUserType >= 0)
				{
					//Found userDefined
					string sDefaultVal = tblUserType[nUserType, "defaultvalue"];
					//maxvalue:char,int,=InitVal
					string sIniOrChar = tblUserType[nUserType, "maxvalue"];

					//INGETER and DECIMAL
					string singeter = tblField[fieldi, "FIELD_INGETER"];
					int ningeter = cc.Util.toInt(singeter);
					string sdecimal = tblField[fieldi, "FIELD_DECIMAL"];
					int ndecimal = cc.Util.toInt(sdecimal);
					//for sql script(or others):FIELD_INGETER_DECIMAL
					//out:(int, dec)
					string sfieldintdec = "";
					if(!singeter.Equals(""))
					{
						sfieldintdec += "(" + singeter;
						if(!sdecimal.Equals(""))
						{
							sfieldintdec += ", " + sdecimal;
						}
						sfieldintdec += ")";
					}
					//for sql script(or others):FIELD_INGETER_PLUS_DECIMAL
					//out:int_plus_dec,if singeter is empty,then empty
					string sfieldintplusdec = singeter;
					if((ningeter + ndecimal) > 0)
					{
						sfieldintplusdec = "" + (ningeter + ndecimal);
					}
					//use sDefaultLen only no (ningeter + ndecimal) defined
					string sDefaultLen = tblUserType[nUserType, "defaultlength"];
					if((ningeter + ndecimal) > 0)
					{
						sDefaultLen = sfieldintplusdec;
					}

					//create FIELD_VALUE_MAX and FIELD_VALUE_MIDDLE
					//sDefaultVal:char,int,=InitVal
					string sMidVal = "";
					string sMaxVal = "";
					if(sIniOrChar.Equals("char"))
					{
						if(ningeter > 0)
						{
							for(int tmpi = 0; tmpi < ningeter; tmpi++)
							{
								if(tmpi.ToString().EndsWith("0"))
								{
									sMaxVal += "A";
								}
								else
								{
									sMaxVal += "X";
								}
								if(tmpi <= (ningeter - 1)/2)
								{
									sMidVal = sMaxVal;
								}
							}
						}
						sMaxVal = "\"" + sMaxVal + "\"";
						sMidVal = "\"" + sMidVal + "\"";
					}
					else if(sIniOrChar.Equals("int"))
					{
						for(int tmpi = 0; tmpi < ningeter; tmpi++)
						{
							if(tmpi.ToString().EndsWith("0"))
							{
								sMaxVal += "9";
							}
							else
							{
								sMaxVal += "1";
							}
							if(tmpi <= (ningeter - 1)/2)
							{
								sMidVal = sMaxVal;
							}
						}
						if(ndecimal > 0)
						{
							if(sMaxVal.Equals(""))
							{
								sMaxVal = "0";
							}
							if(sMidVal.Equals(""))
							{
								sMidVal = "0";
							}
							sMaxVal = sMaxVal + ".";
							sMidVal = sMidVal + ".";
							for(int tmpi = 0; tmpi < ndecimal; tmpi++)
							{
								if(tmpi.ToString().EndsWith("0"))
								{
									sMaxVal += "9";
								}
								else
								{
									sMaxVal += "1";
								}
								if(tmpi <= (ndecimal - 1)/2)
								{
									sMidVal = sMaxVal;
								}
							}
						}
						if(sMaxVal.Equals(""))
						{
							sMaxVal = "0";
						}
						if(sMidVal.Equals(""))
						{
							sMidVal = "0";
						}
					}
					else
					{
						sMaxVal = sDefaultVal;
						sMidVal = sDefaultVal;
					}
					string sDefaultVal_NO_QUOTES = sDefaultVal.Trim('"');
					string sMaxVal_NO_QUOTES = sMaxVal.Trim('"');
					string sMidVal_NO_QUOTES = sMidVal.Trim('"');

					tblField[fieldi, "FIELD_TYPE"] = tblUserType[nUserType, "totype"];
					tblField[fieldi, "FIELD_VALUE_INIT"] = sDefaultVal;
					tblField[fieldi, "FIELD_VALUE_MIDDLE"] = sMidVal;
					tblField[fieldi, "FIELD_VALUE_MAX"] = sMaxVal;
					tblField[fieldi, "FIELD_VALUE_INIT_NO_QUOTES"] = sDefaultVal_NO_QUOTES;
					tblField[fieldi, "FIELD_VALUE_MIDDLE_NO_QUOTES"] = sMidVal_NO_QUOTES;
					tblField[fieldi, "FIELD_VALUE_MAX_NO_QUOTES"] = sMaxVal_NO_QUOTES;
					tblField[fieldi, "FIELD_INGETER_DECIMAL"] = sfieldintdec;
					tblField[fieldi, "FIELD_INGETER_PLUS_DECIMAL"] = sfieldintplusdec;
					tblField[fieldi, "FIELD_SHOW_LENGTH"] = sDefaultLen;
				}
				else
				{
					tblField[fieldi, "FIELD_TYPE"] = sType;
				}
			}

			//for EVAL JScript:
			System.Text.StringBuilder sbJSTxt = new System.Text.StringBuilder(2048);
			sbJSTxt.Append("class JS{\r\n");
			sbJSTxt.Append("	var info = new Object;\r\n");
			sbJSTxt.Append("	var writemsg = '';\r\n");
			sbJSTxt.Append("	var writetxt = '';\r\n");
			sbJSTxt.Append("	var field = new Array();\r\n");
			sbJSTxt.Append("	function JS(){\r\n");
			for(int fieldi = 0; fieldi < tblField.Rows.Count; fieldi++)
			{
				sbJSTxt.Append("		field[" + fieldi + "] = new Object;\r\n");
				for(int fieldx = 2; fieldx < tblField.Columns.Count; fieldx++)
				{
					sbJSTxt.Append("		field[" + fieldi + "]['" + tblField.Columns[fieldx].ColumnName + "'] = " +
						"'" + JSEncode("" + tblField[fieldi, fieldx]) + "';\r\n");
				}
			}
			//add user define various
			for(int vari = 0; vari < ipara.UserVarious.Count; vari++)
			{
				string skey = ipara.UserVarious.GetKey(vari);
				sbJSTxt.Append("		info['" + skey.Substring(2, skey.Length - 4) + "'] = '" + 
					JSEncode(ipara.UserVarious[vari]) + "';\r\n");
			}
			sbJSTxt.Append("	}\r\n");
			sbJSTxt.Append("}\r\n");
			sbJSTxt.Append("var js = new JS();\r\n");
			sbJSTxt.Append("function writetxt(s){	\r\njs.writetxt += s;\r\n}\r\n");
			sbJSTxt.Append("function writemsg(s){	\r\njs.writemsg += s;\r\n}\r\n");
			ipara.SystemVarious["StringBuilderJScriptTxt"] = sbJSTxt;
			return ipara;
		}

		/// <summary>
		/// return String created from tablelayout and template.
		/// </summary>
		/// <param name="excelFile">IPara ipara</param>
		/// <returns>string</returns>
		static public string CreateClsFromTempString(IPara ipara)
		{
			XmlNode xmlnode = ipara.TemplateNode;
			XmlNodeList nodeListField = xmlnode.SelectNodes("field");
			cc.Table tblField = ipara.tblField;
			cc.Table tblUserType = ipara.tblUserType;
			cc.Msg msg = ipara.msg;

			string clsdb = ipara.UserVarious["[#DB_NAME#]"];
			string clstbl = ipara.UserVarious["[#TABLE_NAME#]"];
			string sID;
			string sTxt;
			string sField;

			//TreateField(type,len,value) and add JScript various
			TreateFieldANDJScript(ipara);

			ipara.TemplateFieldCount = nodeListField.Count;
			//treate each XML.Template.Field
			for(int i = 0; i < nodeListField.Count; i++)
			{
				sTxt = "";
				if(nodeListField[i].Attributes["id"] != null)
				{
					sID = nodeListField[i].Attributes["id"].InnerText;
				}
				else
				{
					//no id,skip
					continue;
				}
				string sFieldAddWhileNotEnd;
				if(nodeListField[i].Attributes["addwhilenotend"] != null)
				{
					sFieldAddWhileNotEnd = nodeListField[i].Attributes["addwhilenotend"].InnerText;
				}
				else
				{
					sFieldAddWhileNotEnd = "";
				}

				ipara.TemplateFieldIndex = i;
				ipara.TableFieldCount = tblField.Rows.Count;

				//use every Excel.Field to replace Template.Field,and add every replaced string
				for(int fieldi = 0; fieldi < tblField.Rows.Count; fieldi++)
				{
					sField = nodeListField[i].InnerText;

					//replace Template.Field's Various
					sField = ReplaceVar(sField, "[#FIELD_ROWCOUNT#]", "" + tblField.Rows.Count);
					sField = ReplaceVar(sField, "[#FIELD_ROWINDEX#]", "" + fieldi);

					//tblField all Column defined by system:
					//tblField = new cc.ITable();
					//tblField.AddColumn("FIELD_TYPE_EXCEL"); //get from excel file or server
					//tblField.AddColumn("FIELD_TYPE_INT_OR_CHAR"); //for calcute max value,int or char or =InitVal
					//tblField.AddColumn("FIELD_NAME");
					//tblField.AddColumn("FIELD_INGETER");
					//tblField.AddColumn("FIELD_DECIMAL");
					//tblField.AddColumn("FIELD_INGETER_DECIMAL");
					//tblField.AddColumn("FIELD_INGETER_PLUS_DECIMAL");
					//tblField.AddColumn("FIELD_SHOW_LENGTH"); //INGETER+DECIMAL,if 0,then use defaultlength
					//tblField.AddColumn("FIELD_TYPE"); //for output,user need type(java,c#,sqlscript,...)
					//tblField.AddColumn("FIELD_VALUE_INIT");
					//tblField.AddColumn("FIELD_VALUE_MIDDLE");
					//tblField.AddColumn("FIELD_VALUE_MAX");
					//tblField.AddColumn("FIELD_VALUE_INIT_NO_QUOTES");
					//tblField.AddColumn("FIELD_VALUE_MIDDLE_NO_QUOTES");
					//tblField.AddColumn("FIELD_VALUE_MAX_NO_QUOTES");
					for(int fieldx = 2; fieldx < tblField.Columns.Count; fieldx++)
					{
						sField = ReplaceVar(sField, "[#" + tblField.Columns[fieldx].ColumnName + "#]", "" + tblField[fieldi, fieldx]);
					}

					sTxt = sTxt + sField + sFieldAddWhileNotEnd;
				}

				if(sFieldAddWhileNotEnd.Length > 0 && sTxt.Length > 0)
				{
					sTxt = sTxt.Remove(sTxt.Length - sFieldAddWhileNotEnd.Length, sFieldAddWhileNotEnd.Length);
				}
				//replace user define various
				for(int vari = 0; vari < ipara.UserVarious.Count; vari++)
				{
					sTxt = ReplaceVar(sTxt, ipara.UserVarious.GetKey(vari), ipara.UserVarious[vari]);
				}
				ipara.TemplateFieldTexts[sID] = sTxt;
			}

			ipara.TemplateFieldText = null;
			ipara.TemplateFieldIndex = -1;
			ipara.TableFieldIndex = -1;

			sTxt = xmlnode.SelectSingleNode("main").InnerText;
			sTxt = ReplaceVar(sTxt, "[#FIELD_ROWINDEX#]", "-1");
			sTxt = ReplaceVar(sTxt, "[#FIELD_ROWCOUNT#]", "" + tblField.Rows.Count);
			sTxt = ReplaceVar(sTxt, "[#DB_NAME#]", clsdb);
			sTxt = ReplaceVar(sTxt, "[#TABLE_NAME#]", clstbl);
			//replace user define various
			for(int vari = 0; vari < ipara.UserVarious.Count; vari++)
			{
				sTxt = ReplaceVar(sTxt, ipara.UserVarious.GetKey(vari), ipara.UserVarious[vari]);
			}

			//replace template field
			for(int i = 0; i < ipara.TemplateFieldTexts.Count; i++)
			{
				sTxt = sTxt.Replace("[#CLASS_FIELD_" + ipara.TemplateFieldTexts.GetKey(i) + "#]", ipara.TemplateFieldTexts[i]);
			}

			//for EVAL:
			System.Text.StringBuilder sbJSTxt = (System.Text.StringBuilder)ipara.SystemVarious["StringBuilderJScriptTxt"];
			string sJSTxtHead = (string)ipara.SystemVarious["JScriptTxtHEAD"];
			string sJSTxtEnd = (string)ipara.SystemVarious["JScriptTxtEND"];
			if(!sJSTxtHead.Equals(""))
			{
				sbJSTxt.Append("\r\n");
				sbJSTxt.Append(sJSTxtHead);
			}
			string MErrorString = Eval(ipara, sbJSTxt, ref sTxt, sJSTxtEnd);
			if(MErrorString != null)
			{
				msg.println(MErrorString, Color.Red);
			}
			if(sTxt == null)
			{
				//at ipara or JScript,cancel it
				msg.println(" this file is Canceled by JScript", Color.Red);
				return null;
			}
			string filename2 = ipara.UserVarious["[#TEMPLATE_FILENAME#]"];
			if(filename2 == null || filename2.Equals(""))
			{
				//at ipara or JScript,cancel it
				msg.println(" this file name is set Empty by JScript, Canceled it", Color.Red);
				return null;
			}
			//for out:[#various#\],replace \] ==> ]
			return sTxt.Replace("\\]", "]");
		}

	}
}
