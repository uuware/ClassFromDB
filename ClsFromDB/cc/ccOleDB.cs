using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;

	/// <summary>
	/// 汎用クラス
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// OleDbサーバにデータ操作クラス
	/// connect to Csv(TXT),Excel,...
	/// </summary>
	public class OleDB
	{ 
		private OleDbConnection m_conn = null;
		private OleDbTransaction m_tran = null;
		private Exception m_exp = null;
		private const string MSG_DB_NOT_CONN = "OleDBサーバに接続していません。";

		/// <summary>
		/// SQL Server の接続情報（SqlConnection）を提供して、DBクラスをつくります
		/// </summary>
		public OleDB(OleDbConnection conn)
		{
			if(conn == null || conn.State != ConnectionState.Open)
			{
				m_exp = new Exception(MSG_DB_NOT_CONN);
			}
			else
			{
				m_conn = conn;
			}
		}

		/// <summary>
		/// SQL Server の接続文字列を提供して、DBクラスをつくります
		/// </summary>
		public OleDB(string sConnString)
		{
			try
			{
				m_conn = new OleDbConnection(sConnString);
				m_conn.Open();
			}
			catch(Exception exp)
			{
				m_conn = null;
				m_exp = exp;
				Console.WriteLine(exp);
			}
		}

		~OleDB()
		{
		}

		public void Dispose()
		{
			this.Close();
		}

		/// <summary>
		/// CSVに接続文字列を戻す
		/// Sample:select * from aa.txt where 1 = 1
		/// </summary>
		static public string ConnStringCsv(string sFilePathOnly, bool isHasHead)
		{
			string sSql = @"Provider=Microsoft.Jet.OLEDB.4.0;" + 
				@"Data Source=" + sFilePathOnly  + ";" + 
				@"Extended Properties=" + Convert.ToChar(34).ToString() + 
				@"Text;"; 
			if(isHasHead)
			{
				sSql += "HDR=Yes;";
			}
			else
			{
				sSql += "HDR=No;";
			}
			sSql += Convert.ToChar(34).ToString();
			return sSql;
		}

		/// <summary>
		/// Excelに接続文字列を戻す
		/// </summary>
		static public string ConnStringExcel(string sFilePathName, bool isHasHead)
		{
			//Connection String. Change the excel file to the file you will search.
			string sSql = @"Provider=Microsoft.Jet.OLEDB.4.0;" + 
				@"Data Source=" + sFilePathName  + ";" + 
				@"Extended Properties=" + Convert.ToChar(34).ToString() + 
				@"Excel 8.0;IMEX=1;"; 
			if(isHasHead)
			{
				sSql += "HDR=Yes;";
			}
			else
			{
				sSql += "HDR=No;";
			}
			sSql += Convert.ToChar(34).ToString();
			return sSql;
		}

		/// <summary>
		/// Accessに接続文字列を戻す
		/// </summary>
		static public string ConnStringAccess(string sFilePathName, string sUser, string sPass)
		{
			string sSql = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" 
				+ sFilePathName + ";"
				+ ";User Id=" + sUser 
				+ ";Password=" + sPass + ";";
			return sSql;
		}

		/// <summary>
		/// Oracleに接続文字列を戻す
		/// </summary>
		static public string ConnStringOracle(string sServer, string sUser, string sPass, string sDB)
		{
			string sSql = @"Provider=msdaora;Data Source=" + sServer 
				+ ";User Id=" + sUser 
				+ ";Password=" + sPass + ";";
			return sSql;
		}

		/// <summary>
		/// SQLServerに接続文字列を戻す
		/// </summary>
		static public string ConnStringSQLServer(string sServer, string sUser, string sPass, string sDB)
		{
			string sSql = @"Provider=sqloledb;Data Source=" + sServer 
				+ ";Initial Catalog=" + sDB 
				+ ";User Id=" + sUser 
				+ ";Password=" + sPass + ";";
			return sSql;
		}

		/// <summary>
		/// SQLServerに接続文字列を戻す(Trusted Connection)
		/// </summary>
		static public string ConnStringSQLServer(string sServer, string sDB)
		{
			string sSql = @"Provider=sqloledb;Data Source=" + sServer 
				+ ";Initial Catalog=" + sDB 
				+ ";Integrated Security=SSPI;";
			return sSql;
		}																	

		/// <summary>
		/// This mehtod retrieves the excel sheets information from 
		/// an excel workbook.
		/// </summary>
		/// <param name="excelFile">The excel file.</param>
		/// <param name="isGetSheetsNameOnly">if true,return SheetsNames only</param>
		/// <returns>System.Data.DataSet (SheetsNames,Sheet1,Sheet2...)</returns>
		public string[] GetExcelSheetsName()
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return null;
			}
			try
			{
				//Get the data table containg the schema guid.
				System.Data.DataTable dt = m_conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
 
				if(dt == null)
				{
					return null;
				}

				System.Collections.Specialized.NameValueCollection coll = new System.Collections.Specialized.NameValueCollection();
				foreach(DataRow row in dt.Rows)
				{
					string s = row["TABLE_NAME"].ToString();
					int npos = s.LastIndexOf("$");
					if(s.StartsWith("'"))
					{
						s = s.Substring(1, npos - 1);
					}
					else
					{
						s = s.Substring(0, npos);
					}
					//SheetName:$'Print_Area,$'Print_Titles
					if(npos > row["TABLE_NAME"].ToString().Length - 3)
					{
						coll.Set(s, row["TABLE_NAME"].ToString());
					}
				}
				dt.Dispose();
				return coll.AllKeys;
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".chgDB:\r\n" + exp);
				m_exp = exp;
				return null;
			}
		}

		/// <summary>
		/// This mehtod retrieves the excel sheet information from 
		/// an excel workbook.
		/// </summary>
		/// <param name="sSheetName">The sheet name.</param>
		/// <returns>System.Data.DataTable</returns>
		public System.Data.DataTable GetExcelSheet(string sSheetName)
		{
			string sSql = "SELECT * FROM " + "[" + sSheetName + "$]";
			DataSet ds = GetDataSet(sSql);
			if(ds != null)
			{
				return ds.Tables[0];
			}
			else
			{
				return null;
			}
		}

		/// <summary>
		/// 接続がオープンかないかを判断
		/// </summary>
		public bool isOpen()
		{
			if(m_conn != null && m_conn.State == ConnectionState.Open)
			{
				return true;
			}
			return false;
		}

		/// <summary>
		/// 接続をクローズします
		/// </summary>
		private void Close()
		{
			try
			{
				if(m_conn != null && m_conn.State == ConnectionState.Open)
				{
					if(m_tran != null)
					{
						if(!Commit())
						{
							Rollback();
						}
					}
					m_conn.Close();
					m_conn.Dispose();
					m_conn = null;
				}
			}
			catch(Exception exp)
			{
				m_exp = exp;
				Console.WriteLine(this.GetType().FullName + ".Close:\r\n" + exp);
			}
		}

		/// <summary>
		/// データ ソースでトランザクションを開始します
		/// </summary>
		public bool BeginTransaction()
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return false;
			}
			if(m_tran != null)
			{
				return false;
			}
			m_tran = m_conn.BeginTransaction();
			return true;
		}
	
		/// <summary>
		/// データ ソースでCommitします
		/// </summary>
		public bool Commit()
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return false;
			}
			if(m_tran == null)
			{
				return false;
			}
			try
			{
				m_tran.Commit();
				m_tran = null;
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".Commit:\r\n" + exp);
				m_exp = exp;
				return false;
			}
			return true;
		}
	
		/// <summary>
		/// データ ソースでRollbackします
		/// </summary>
		public bool Rollback()
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return false;
			}
			if(m_tran == null)
			{
				return false;
			}
			try
			{
				m_tran.Rollback();
				m_tran = null;
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".Rollback:\r\n" + exp);
				m_exp = exp;
				return false;
			}
			return true;
		}
	
		/// <summary>
		/// SQLを実行し、結果（DataSet）を戻します
		/// </summary>
		public DataSet GetDataSet(string sSql)
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return null;
			}
			m_exp = null;
			DataSet ds = new DataSet();
			try
			{
				OleDbDataAdapter da = new OleDbDataAdapter();
				if(m_tran != null)
				{
					da.SelectCommand = new OleDbCommand(sSql, m_conn, m_tran);
				}
				else
				{
					da.SelectCommand = new OleDbCommand(sSql, m_conn);
				}
				da.Fill(ds);
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".GetDataSet:\r\n" + exp);
				m_exp = exp;
				return null;
			}
			return ds;
		}

		/// <summary>
		/// SQLを実行し、結果（Table）を戻します
		/// </summary>
		public DataTable GetTable(string sSql)
		{
			DataSet ds = GetDataSet(sSql);
			if(ds != null)
			{
				return ds.Tables[0];
			}
			else
			{
				return null;
			}
		}

		/// <summary>
		/// SQLを実行し、結果（Row）を戻します
		/// </summary>
		public DataRow GetRow(string sSql)
		{
			DataSet ds = GetDataSet(sSql);
			if(ds != null)
			{
				return ds.Tables[0].Rows[0];
			}
			else
			{
				return null;
			}
		}

		/// <summary>
		/// SQLを実行し、結果（ExecuteScalar：一つ値を戻す）を戻します
		/// </summary>
		public Object GetRow0Col0(string sSql)
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return null;
			}
			m_exp = null;
			try
			{
				OleDbCommand sc = new OleDbCommand(sSql, m_conn);
				if(m_tran != null)
				{
					sc.Transaction = m_tran;
				}
				return sc.ExecuteScalar();
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".GetRow0Col0:\r\n" + exp);
				m_exp = exp;
			}
			return null;
		}

		/// <summary>
		/// SQLを実行し、戻る結果がなし
		/// </summary>
		public int ExecuteNonQuery(string sSql)
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return -1;
			}
			m_exp = null;
			try
			{
				OleDbCommand sc = new OleDbCommand(sSql, m_conn);
				if(m_tran != null)
				{
					sc.Transaction = m_tran;
				}
				return sc.ExecuteNonQuery();
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".ExecuteNonQuery:\r\n" + exp);
				m_exp = exp;
			}
			return -1;
		}

		/// <summary>
		/// SQL Serverにデータベースを変わります
		/// </summary>
		public bool chgDB(string sDatabase)
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return false;
			}
			if(m_conn.Database.Equals(sDatabase))
			{
				return true;
			}
			try
			{
				m_conn.ChangeDatabase(sDatabase);
			}
			catch(Exception exp)
			{
				Console.WriteLine(this.GetType().FullName + ".chgDB:\r\n" + exp);
				m_exp = exp;
				return false;
			}
			return true;
		}

		/// <summary>
		/// SQLを実行し、エラーがあった場合、Exceptionを取得します
		/// </summary>
		public Exception Exception
		{
			get
			{
				if(m_conn == null || m_conn.State != ConnectionState.Open)
				{
					return new Exception(MSG_DB_NOT_CONN);
				}
				return m_exp;
			}
		}

		/// <summary>
		/// SQLを実行し、エラーがあるかどうかの判断
		/// </summary>
		public bool Error()
		{
			if(m_conn == null || m_conn.State != ConnectionState.Open)
			{
				return true;
			}
			if(m_exp != null)
			{
				return true;
			}
			return false;
		}
	}

}
