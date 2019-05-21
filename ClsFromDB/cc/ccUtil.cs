/*Modify List:
 * split Msg to single file
 * 
 */
using System;
using System.Drawing;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Forms;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Util
	/// </summary>
	public class Util
	{
		public Util()
		{
		}
 
		public void Dispose()
		{
		}

		/// <summary>
		/// Select Folder.
		/// when return is not null, then this is error message
		/// </summary>
		static public string create_subdir(string pdir, string subdir, ref string fulldir)
		{
			if(!Directory.Exists(pdir))
			{
				return "can not create sub dir for parent dir is not exist:" + pdir;
			}
			while(subdir.StartsWith("\\") || subdir.StartsWith("/"))
			{
				subdir = subdir.Substring(1);
			}
			if(subdir.Equals(""))
			{
				return null;
			}
			if(subdir.IndexOf(":") >= 0)
			{
				return "the subdir is not valid for include \":\":" + subdir;
			}
			pdir = pdir.Replace("/", "\\");
			subdir = subdir.Replace("/", "\\");
			if(!pdir.EndsWith("\\"))
			{
				pdir += "\\";
			}
			if(!subdir.EndsWith("\\"))
			{
				subdir += "\\";
			}
			int npos = subdir.IndexOf("\\");
			while(npos > 0)
			{
				string newdir = pdir + subdir.Substring(0, npos);
				try
				{
					if(!Directory.Exists(newdir))
					{
						Directory.CreateDirectory(newdir);
					}
					if(!Directory.Exists(newdir))
					{
						return "can not create dir:" + newdir;
					}
					fulldir = newdir;
				}
				catch(Exception exp)
				{
					return exp.Message;
				}
				npos = subdir.IndexOf("\\", npos + 1);
			}
			if(!fulldir.EndsWith("\\"))
			{
				fulldir += "\\";
			}
			return null;
		}

		/// <summary>
		/// Replace but IgnoreCase
		/// </summary>
		static public string ReplaceIgnoreCase(string src, string sfrom, string swith)
		{
			int nposold = 0;
			int nposnew = 0;
			System.Text.StringBuilder sb = new System.Text.StringBuilder();
			string srcUp = src.ToUpper();
			string sfromUp = sfrom.ToUpper();
			nposnew = srcUp.IndexOf(sfromUp, nposold);
			while(nposnew >= 0)
			{
				sb.Append(src.Substring(nposold,  nposnew - nposold) + swith);
				nposold = nposnew + sfrom.Length;
				nposnew = srcUp.IndexOf(sfromUp, nposold);
			}
			sb.Append(src.Substring(nposold));
			return sb.ToString();
		}

		static public System.Collections.Specialized.NameValueCollection ReadIni(string sFileName)
		{
			System.Collections.Specialized.NameValueCollection coll = null;
			string sTxt = cc.Util.readAll(sFileName);
			if(sTxt != null)
			{
				coll = new System.Collections.Specialized.NameValueCollection();
				string[] lines = sTxt.Replace("\n","").Split('\r');
				for (int i = 0; i < lines.Length; i++)
				{
					string line = lines[i].Trim();
					if(line.Equals("") || line.StartsWith("#") || line.StartsWith(";"))
					{
						continue;
					}

					int npos;
					npos = line.IndexOf("=");
					if(npos > 0)
					{
						string skey = line.Substring(0, npos);
						if(coll.Get(skey) != null)
						{
							coll.Set(skey, coll.Get(skey) + "\r\n" + line.Substring(npos + 1));
						}
						else
						{
							coll.Set(skey, line.Substring(npos + 1));
						}
					}
				}
			}
			return coll;
		}

		static public bool writeFile(string sFileFullName, string sWriteString)
		{
			return writeFile(sFileFullName, sWriteString, false);
		}

		static public bool writeFile(string sFileFullName, string sWriteString, bool isAppend)
		{
			try
			{
				string initpath = Path.GetDirectoryName(Application.ExecutablePath);
				if(sFileFullName.StartsWith(".\\") || sFileFullName.StartsWith("./"))
				{
					//sFileFullName = ".\\filename.ext";
					sFileFullName = initpath + sFileFullName.Substring(1);
				}
				else if(sFileFullName.IndexOf("\\") < 0 && sFileFullName.IndexOf("/") < 0)
				{
					//sFileFullName = "filename.ext";
					sFileFullName = initpath + "\\" + sFileFullName;
				}
				System.IO.StreamWriter sw = new System.IO.StreamWriter(sFileFullName, isAppend, System.Text.Encoding.Default);
				sw.Write(sWriteString);
				sw.Flush();
				sw.Close();
			}
			catch
			{
				return false;
			}
			return true;
		}

		static public string readAll(string sFileFullName)
		{
			System.Text.Encoding encoding = null;
			return readAll(sFileFullName, ref encoding);
		}

		static public string readAll(string sFileFullName, ref System.Text.Encoding encoding)
		{
			string sTxt = null;
			try
			{
				string initpath = Path.GetDirectoryName(Application.ExecutablePath);
				if(sFileFullName.StartsWith(".\\") || sFileFullName.StartsWith("./"))
				{
					//sFileFullName = ".\\filename.ext";
					sFileFullName = initpath + sFileFullName.Substring(1);
				}
				else if(sFileFullName.IndexOf("\\") < 0 && sFileFullName.IndexOf("/") < 0)
				{
					//sFileFullName = "filename.ext";
					sFileFullName = initpath + "\\" + sFileFullName;
				}
				if(File.Exists(sFileFullName))
				{
					FileStream fs = new FileStream(sFileFullName, FileMode.Open, FileAccess.Read);
					long filesize = fs.Length;
					BinaryReader reader = new BinaryReader(fs);
					byte[] srcByte = reader.ReadBytes((int)fs.Length);
					fs.Close();
					sTxt = AutoEncoding(srcByte, ref encoding);
					//System.IO.StreamReader sr = new System.IO.StreamReader(sFileFullName, encoding);
					//sTxt = sr.ReadToEnd();
					//sr.Close();
				}
			}
			catch(Exception exp)
			{
				throw exp;
			}
			return sTxt;
		}

		static public string readResource(string sFileFullName)
		{
			System.Reflection.Assembly asse = System.Reflection.Assembly.GetExecutingAssembly();
			//System.IO.Stream file = asse.GetManifestResourceStream("AssemblyName.ImageFile.jpg");
			//this.pictureBox1.Image = Image.FromStream(file);
			return readResource(asse, sFileFullName);
		}

		static public string readResource(System.Reflection.Assembly asse, string sFileFullName)
		{
			string sTxt = string.Empty;
			try
			{
				System.IO.Stream strm = asse.GetManifestResourceStream(sFileFullName);
				if(strm == null)
				{
					for(int i = 0; i < asse.GetManifestResourceNames().Length; i++)
					{
						if(asse.GetManifestResourceNames()[i].EndsWith("." + sFileFullName))
						{
							strm = asse.GetManifestResourceStream(asse.GetManifestResourceNames()[i]);
							break;
						}
					}
				}
				if(strm != null)
				{

					byte[] srcByte = new byte[strm.Length];
					strm.Read(srcByte, 0, (int)strm.Length);
					//sTxt = new System.IO.StreamReader(strm,System.Text.Encoding.Default).ReadToEnd();
					strm.Close();
					sTxt = AutoEncoding(srcByte);
				}
			}
			catch(Exception exp)
			{
				throw exp;
			}
			return sTxt;
		}

		static public string AutoEncoding(byte[] bt)
		{
			System.Text.Encoding enc = null;
			return AutoEncoding(bt, ref enc);
		}

		static public string AutoEncoding(byte[] bt,ref System.Text.Encoding enc)
		{
			System.Text.Encoding[] arrenc = {
												System.Text.Encoding.Default,
												System.Text.Encoding.GetEncoding("ascii"),
												System.Text.Encoding.GetEncoding("UTF-8"),
												System.Text.Encoding.GetEncoding("gb2312"),//cannot detect gb2312&big5
												System.Text.Encoding.GetEncoding("big5"),
												System.Text.Encoding.GetEncoding("shift_jis"),
												System.Text.Encoding.GetEncoding("euc-jp"),
												System.Text.Encoding.GetEncoding("iso-2022-jp"),
												System.Text.Encoding.GetEncoding("utf-16")
											};
			byte[] btenu;
			enc = System.Text.Encoding.Unicode;
			btenu = enc.GetPreamble();
			if(bt.Length > 2 && bt[0] == btenu[0] && bt[1] == btenu[1])
			{
				return enc.GetString(bt, 2, bt.Length - 2);
			}
			enc = System.Text.Encoding.UTF8;
			btenu = enc.GetPreamble();
			if(bt.Length > 3 && bt[0] == btenu[0] && bt[1] == btenu[1] && bt[2] == btenu[2])
			{
				return enc.GetString(bt, 3, bt.Length - 3);
			}
			for(int i = 0; i < arrenc.Length; i++)
			{
				enc = arrenc[i];
				string sTxt = enc.GetString(bt);
				byte[] bt2 = enc.GetBytes(sTxt);
				if(bt.Length == bt2.Length)
				{
					for(int j = 0; j < bt.Length; j++)
					{
						if(bt[j] != bt2[j])
						{
							break;
						}
						if(j == bt.Length - 1)
						{
							return sTxt;
						}
					}
				}
			}
			enc = System.Text.Encoding.Default;
			return enc.GetString(bt);
		}

		/// <summary>
		/// Copy directory structure recursively
		/// </summary>
		public static void DirCopy(string src, string dst)
		{
			if(dst[dst.Length-1] != Path.DirectorySeparatorChar)
			{
				dst += Path.DirectorySeparatorChar;
			}
			if(!Directory.Exists(dst))
			{
				Directory.CreateDirectory(dst);
			}
			string[] files = Directory.GetFiles(src);
			foreach(string element in files)
			{
				//files in directory
				File.Copy(element, dst + Path.GetFileName(element), true);
			}
			string[] dirs = Directory.GetDirectories(src);
			foreach(string element in dirs)
			{
				//Sub directories
				DirCopy(element, dst + Path.GetFileName(element));
			}
		}

		/// <summary>
		/// Select Folder.
		/// when return is not null,then this is selected path
		/// need:System.Windows.Forms.dll
		/// use like this:
		///		string sFullPath = cc.DirSelect("Dialog Title Goes Here", "c:\\");
		///     if(sFullPath != null)
		///     {
		///     	MessageBox.Show(sFullPath);
		///     }
		/// </summary>
		public static string DirSelect(string sTitle, string sLocation)
		{
			FolderBrowserDialog fbd = new FolderBrowserDialog();
			fbd.Description = sTitle;
			fbd.RootFolder = System.Environment.SpecialFolder.Desktop;
			fbd.SelectedPath = sLocation;
			fbd.ShowNewFolderButton = true;
			if(fbd.ShowDialog() == DialogResult.OK)
			{
				return fbd.SelectedPath;
			}
			else
			{
				return null;
			}
		}

		//to int,if null or "",then 0
		//if "123.456" or "123aaa" then retu "123"
		static public int toInt(string str)
		{
			int nretu = 0;
			if(str == null || str.Trim().Equals(""))
			{
				return nretu;
			}
			str = "0" + str.Trim();
			int nLen = 0;
			char[] ch = str.ToCharArray();
			for(nLen = 0; nLen < ch.Length; nLen++)
			{
				if(ch[nLen] == '.' || ch[nLen] < 48 || ch[nLen] > 67)
				{
					break;
				}
			}
			try
			{
				nretu = int.Parse(str.Substring(0, nLen));
			}
			catch
			{
			}
			return nretu;
		}

	}

}
