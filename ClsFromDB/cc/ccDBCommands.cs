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
	/// SQLServerサーバにデータSQLString、コメントとして預かっている
	/// </summary>
	class DBCommands
	{
		#region QueryStrings
		const string QueryString_GetDatabaseObjects = @"declare @SQL varchar(1000) 
															declare @DBName varchar(255) 
															declare SYSDB cursor for 
															select name 
															from master.dbo.sysdatabases 
															where has_dbaccess(name) = 1 
															order by name 

															open SYSDB 
															fetch next from SYSDB into  @DBName 
															while @@fetch_status = 0 
															begin 
															set @SQL = 'use [' + @DBName +']  select '''
																+ @DBName + ''', O.name, xtype from [' + @DBName + '].dbo.sysobjects O
																where id >1000 
																and xtype in (''U'',''P'',''V'',''FN'') 
																and name not like (''dt_%'') 
																order by xtype , O.name' 
															exec( @SQL ) 
															fetch next from SYSDB into  @DBName 
															end 
															close SYSDB 
															Deallocate SYSDB ";
			
		const string QueryString_GetDatabaseObject= "select name, xtype, id from [{0}].dbo.sysobjects where name like ('{1}%') and xtype in ('U','P','FN','TF','V') order by name";
		const string QueryString_DatabaseObjectProperties =  "select C.* from {0}.dbo.sysobjects O join {0}.dbo.syscolumns C on C.id = O.id where O.name = '{1}' and (O.xtype ='U' or O.xtype ='V') order by C.name";
		const string QueryString_CreateScript = "select text from sysobjects o join  syscomments c on c.id = o.id where o.name = '{0}' order by o.name";
		//const string QueryString_CreateScript = @"select " +
		//										"	case xtype " +
		//										"		when 'P' then 'DROP PROCEDURE {0}"+"\n"+"GO"+"\n"+"' + text  " +
		//										"		when 'FN' then 'DROP FUNCTION {0}"+"\n"+"GO"+"\n"+"' + text  " +
		//										"		when 'TF' then 'DROP FUNCTION {0}"+"\n"+"GO"+"\n"+"' + text  " +
		//										"		when 'V' then 'DROP VIEW {0}"+"\n"+"GO"+"\n"+"' + text  " +
		//										"		else text  " +
		//										"	end  " +
		//										"from sysobjects o  " +
		//										"join  syscomments c on c.id = o.id  " +
		//										"where o.name = '{0}' order by o.name " ;

		const string QueryString_GetJoiningOptions = @"select 
															o.name, fc.name,ro.name, c.name, fk.*
															from sysobjects o 
															join sysforeignkeys fk on fk.fkeyid = o.id
															join sysobjects ro on ro.id = fk.rkeyid
															join syscolumns c on c.id = ro.id and c.colid = fk.rkey																				  join syscolumns fc on fc.id = o.id and fc.colid = fk.fkey
															where o.name = '{0}'";
		const string QueryString_AllObjects = @"SELECT 1 as Tag, NULL as Parent,
							o.name as [DBObject!1!Name],
							o.xtype as [DBObject!1!Type],
							null as [DBObjectAttribute!2!Name] ,
							null as [DBObjectAttribute!2!Type], 
							null as [DBObjectAttribute!2!Length] ,
							null as [DBObjectAttribute!2!Precision]
						from 	sysobjects o 
						where 	o.xType != 'S'
						union all
						SELECT 2,1,
							o.name,
							o.xtype,
							c.name,
							t.name,
							c.length,
							c.prec
						from 	sysobjects o 
						join	sysColumns c on c.id = o.id
						join 	systypes t on t.xtype = c.xtype
						where 	o.xType != 'S'
						and 	len(c.Name)>0
						ORDER BY [DBObject!1!Name],[DBObjectAttribute!2!Name]
						FOR XML EXPLICIT";
		const string QueryString_ReferenceObjects = @"SELECT distinct o.name, o.xtype
						from syscomments c
						join sysobjects o on o.id = c.id
						where o.name != '{0}'
						and	c.text like'%{0} %' or c.text like'%{0}(%'";
		#endregion

		static public string CreateScript(string objectName)
		{
			return String.Format(QueryString_CreateScript, objectName);
		}
		static public string AllDatabaseObjects_()
		{
			return QueryString_AllObjects;
		}
		static public string DatabaseObjects()
		{
			return QueryString_GetDatabaseObjects;
		}
		static public string DatabaseObject(string DBName, string likeChar)
		{
			return String.Format(QueryString_GetDatabaseObject, DBName, likeChar);
		}
		static public string DatabasesObjectProperties(string DBName, string objectName)
		{
			return String.Format(QueryString_DatabaseObjectProperties, DBName,objectName);
		}
		static public string DatabasesReferenceObjects(string objectName)
		{
			return String.Format(QueryString_ReferenceObjects, objectName);
		}
		static public string GetJoiningOptions_(string objectName)
		{
			return String.Format(QueryString_GetJoiningOptions, objectName);
		}
	}
}
