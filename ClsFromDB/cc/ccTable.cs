/*Modify List:
 * 20050725,Color&Font add
 * 
 */
using System;
using System.Data;
using System.Text;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Interface of Table
	/// not return null but "" at any case
	/// </summary>
	public class Table : System.Data.DataTable
	{
		protected static string ERR_NOT_EXIST_ROW = "not exist row:{0}";
		protected static string ERR_NOT_EXIST_COLUMN = "not exist column:{0}";
		public Table() : base()
		{
		}

		public Table(string sTblName) : base(sTblName)
		{
		}

		/// <summary>
		/// ColumnExist
		/// </summary>
		public bool ColumnExist(string sColumnName)
		{
			sColumnName = sColumnName.ToUpper();
			for(int xi = 0; xi < this.Columns.Count; xi++)
			{
				if(this.Columns[xi].ColumnName.ToUpper().Equals(sColumnName))
				{
					return true;
				}
			}
			return false;
		}

		/// <summary>
		/// ExistColumn
		/// </summary>
		public string ColumnName(int nCol)
		{
			if(nCol >= this.Columns.Count)
			{
				throw new cc.AppException(String.Format(ERR_NOT_EXIST_COLUMN, nCol));
			}
			return this.Columns[nCol].ColumnName;
		}

		/// <summary>
		/// search from Table
		/// </summary>
		public DataRow SearchRow(string sColumnName, string sValue)
		{
			sColumnName = sColumnName.ToUpper();
			if(!ColumnExist(sColumnName))
			{
				return null;
			}
			for(int yi = 0; yi < this.Rows.Count; yi++)
			{
				object obj = this.Rows[yi][sColumnName];
				if((obj == null && sValue == null) || (obj != null && obj.ToString().Equals(sValue)))
				{
					return this.Rows[yi];
				}
			}
			return null;
		}

		/// <summary>
		/// search from Table
		/// </summary>
		public DataRow SearchRow(int nCol, string sValue)
		{
			if(nCol >= this.Columns.Count)
			{
				return null;
			}
			for(int yi = 0; yi < this.Rows.Count; yi++)
			{
				object obj = this.Rows[yi][nCol];
				if((obj == null && sValue == null) || (obj != null && obj.ToString().Equals(sValue)))
				{
					return this.Rows[yi];
				}
			}
			return null;
		}

		/// <summary>
		/// get/set value of ITable
		/// </summary>
		public DataRow this[int nLine]
		{
			get
			{
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				return this.Rows[nLine];
			}
			set
			{
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				for(int i = 0; i < this.Columns.Count; i++)
				{
					this.Rows[nLine][i] = value[i];
				}
			}
		}

		/// <summary>
		/// get/set value(object) of ITable
		/// </summary>
		public object this[int nLine, int nCol]
		{
			get
			{
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				if(nCol >= this.Columns.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_COLUMN, nCol));
				}
				return this.Rows[nLine][nCol];
			}
			set
			{
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				if(nCol >= this.Columns.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_COLUMN, nCol));
				}
				this.Rows[nLine][nCol] = value;
			}
		}

		/// <summary>
		/// get/set value of ITable
		/// </summary>
		public string this[int nLine, string sColumnName]
		{
			get
			{
				sColumnName = sColumnName.ToUpper();
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				if(!ColumnExist(sColumnName))
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_COLUMN, sColumnName));
				}
				return "" + this.Rows[nLine][sColumnName];
			}
			set
			{
				sColumnName = sColumnName.ToUpper();
				if(nLine >= this.Rows.Count)
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_ROW, nLine));
				}
				if(!ColumnExist(sColumnName))
				{
					throw new cc.AppException(String.Format(ERR_NOT_EXIST_COLUMN, sColumnName));
				}
				this.Rows[nLine][sColumnName] = value;
			}
		}

		/// <summary>
		/// return all value of Table
		/// </summary>
		public override string ToString()
		{
			return ToString("; ");
		}

		/// <summary>
		/// return all value of Table
		/// </summary>
		public string ToString(string sSeparator)
		{
			StringBuilder sb = new StringBuilder();
			for(int xi = 0; xi < this.Columns.Count; xi++)
			{
				sb.Append(this.Columns[xi].ColumnName + sSeparator);
			}
			sb.Append("\r\n");
			for(int yi = 0; yi < this.Rows.Count; yi++)
			{
				for(int xi = 0; xi < this.Columns.Count; xi++)
				{
					sb.Append("" + this.Rows[yi][xi] + sSeparator);
				}
				sb.Append("\r\n");
			}
			return sb.ToString();
		}

	}

}
