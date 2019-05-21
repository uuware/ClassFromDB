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
	public class PageConfigTemplate : System.Windows.Forms.UserControl
	{
		public Form1 frmMain = null;
		private XmlDocument xmldoc = null;

		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		private System.Windows.Forms.TextBox txtXmlPath;
		private System.Windows.Forms.TextBox txtTempTest;
		public System.Windows.Forms.CheckedListBox lstTemp;
		private System.Windows.Forms.Button btnLoadXML;
		private System.Windows.Forms.Button btnTempTest;
		private System.Windows.Forms.Button btnSelAll;
		private System.Windows.Forms.Button btnSelAllUN;
		private System.Windows.Forms.Button btnRemvAll;
		private System.Windows.Forms.CheckBox chkJScriptFile1;
		private System.Windows.Forms.TextBox txtJScriptFile1;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox txtJScriptFile2;
		private System.Windows.Forms.CheckBox chkJScriptFile2;
		private System.Windows.Forms.Button btnJScriptFile1;
		private System.Windows.Forms.Button btnJScriptFile2;

		public PageConfigTemplate()
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
			this.txtXmlPath = new System.Windows.Forms.TextBox();
			this.txtTempTest = new System.Windows.Forms.TextBox();
			this.lstTemp = new System.Windows.Forms.CheckedListBox();
			this.btnLoadXML = new System.Windows.Forms.Button();
			this.btnTempTest = new System.Windows.Forms.Button();
			this.btnSelAll = new System.Windows.Forms.Button();
			this.btnSelAllUN = new System.Windows.Forms.Button();
			this.btnRemvAll = new System.Windows.Forms.Button();
			this.chkJScriptFile1 = new System.Windows.Forms.CheckBox();
			this.txtJScriptFile1 = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.btnJScriptFile1 = new System.Windows.Forms.Button();
			this.btnJScriptFile2 = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.txtJScriptFile2 = new System.Windows.Forms.TextBox();
			this.chkJScriptFile2 = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// txtXmlPath
			// 
			this.txtXmlPath.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtXmlPath.BorderStyle = System.Windows.Forms.BorderStyle.None;
			this.txtXmlPath.Location = new System.Drawing.Point(80, 8);
			this.txtXmlPath.Name = "txtXmlPath";
			this.txtXmlPath.ReadOnly = true;
			this.txtXmlPath.Size = new System.Drawing.Size(400, 12);
			this.txtXmlPath.TabIndex = 15;
			this.txtXmlPath.Text = "XML Template List:";
			// 
			// txtTempTest
			// 
			this.txtTempTest.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtTempTest.Location = new System.Drawing.Point(80, 168);
			this.txtTempTest.Multiline = true;
			this.txtTempTest.Name = "txtTempTest";
			this.txtTempTest.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtTempTest.Size = new System.Drawing.Size(568, 104);
			this.txtTempTest.TabIndex = 14;
			this.txtTempTest.Text = "";
			// 
			// lstTemp
			// 
			this.lstTemp.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.lstTemp.CheckOnClick = true;
			this.lstTemp.Location = new System.Drawing.Point(80, 32);
			this.lstTemp.Name = "lstTemp";
			this.lstTemp.Size = new System.Drawing.Size(568, 88);
			this.lstTemp.TabIndex = 13;
			this.lstTemp.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstTemp_ItemCheck);
			// 
			// btnLoadXML
			// 
			this.btnLoadXML.Location = new System.Drawing.Point(0, 8);
			this.btnLoadXML.Name = "btnLoadXML";
			this.btnLoadXML.TabIndex = 12;
			this.btnLoadXML.Text = "LoadXML";
			this.btnLoadXML.Click += new System.EventHandler(this.btnLoadXML_Click);
			// 
			// btnTempTest
			// 
			this.btnTempTest.Location = new System.Drawing.Point(0, 168);
			this.btnTempTest.Name = "btnTempTest";
			this.btnTempTest.TabIndex = 8;
			this.btnTempTest.Text = "CreateTest";
			this.btnTempTest.Click += new System.EventHandler(this.btnTempTest_Click);
			// 
			// btnSelAll
			// 
			this.btnSelAll.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelAll.Location = new System.Drawing.Point(0, 32);
			this.btnSelAll.Name = "btnSelAll";
			this.btnSelAll.Size = new System.Drawing.Size(32, 16);
			this.btnSelAll.TabIndex = 11;
			this.btnSelAll.Text = "All";
			this.btnSelAll.Click += new System.EventHandler(this.btnSelAll_Click);
			// 
			// btnSelAllUN
			// 
			this.btnSelAllUN.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnSelAllUN.Location = new System.Drawing.Point(32, 32);
			this.btnSelAllUN.Name = "btnSelAllUN";
			this.btnSelAllUN.Size = new System.Drawing.Size(43, 16);
			this.btnSelAllUN.TabIndex = 10;
			this.btnSelAllUN.Text = "UnAll";
			this.btnSelAllUN.Click += new System.EventHandler(this.btnSelAllUN_Click);
			// 
			// btnRemvAll
			// 
			this.btnRemvAll.Font = new System.Drawing.Font("MS UI Gothic", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnRemvAll.Location = new System.Drawing.Point(0, 48);
			this.btnRemvAll.Name = "btnRemvAll";
			this.btnRemvAll.Size = new System.Drawing.Size(75, 16);
			this.btnRemvAll.TabIndex = 19;
			this.btnRemvAll.Text = "Remove All";
			this.btnRemvAll.Click += new System.EventHandler(this.btnRemvAll_Click);
			// 
			// chkJScriptFile1
			// 
			this.chkJScriptFile1.Location = new System.Drawing.Point(408, 120);
			this.chkJScriptFile1.Name = "chkJScriptFile1";
			this.chkJScriptFile1.Size = new System.Drawing.Size(216, 24);
			this.chkJScriptFile1.TabIndex = 44;
			this.chkJScriptFile1.Text = "Valid(Add to Head of Each Template)";
			// 
			// txtJScriptFile1
			// 
			this.txtJScriptFile1.Location = new System.Drawing.Point(80, 120);
			this.txtJScriptFile1.Name = "txtJScriptFile1";
			this.txtJScriptFile1.Size = new System.Drawing.Size(312, 19);
			this.txtJScriptFile1.TabIndex = 42;
			this.txtJScriptFile1.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 127);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(72, 16);
			this.label4.TabIndex = 41;
			this.label4.Text = "JScriptFile1:";
			// 
			// btnJScriptFile1
			// 
			this.btnJScriptFile1.Font = new System.Drawing.Font("MS UI Gothic", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnJScriptFile1.Location = new System.Drawing.Point(392, 120);
			this.btnJScriptFile1.Name = "btnJScriptFile1";
			this.btnJScriptFile1.Size = new System.Drawing.Size(16, 19);
			this.btnJScriptFile1.TabIndex = 43;
			this.btnJScriptFile1.Text = "...";
			this.btnJScriptFile1.Click += new System.EventHandler(this.btnJScriptFile1_Click);
			// 
			// btnJScriptFile2
			// 
			this.btnJScriptFile2.Font = new System.Drawing.Font("MS UI Gothic", 6F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(128)));
			this.btnJScriptFile2.Location = new System.Drawing.Point(392, 144);
			this.btnJScriptFile2.Name = "btnJScriptFile2";
			this.btnJScriptFile2.Size = new System.Drawing.Size(16, 19);
			this.btnJScriptFile2.TabIndex = 43;
			this.btnJScriptFile2.Text = "...";
			this.btnJScriptFile2.Click += new System.EventHandler(this.btnJScriptFile2_Click);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 149);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(72, 16);
			this.label1.TabIndex = 41;
			this.label1.Text = "JScriptFile2:";
			// 
			// txtJScriptFile2
			// 
			this.txtJScriptFile2.Location = new System.Drawing.Point(80, 144);
			this.txtJScriptFile2.Name = "txtJScriptFile2";
			this.txtJScriptFile2.Size = new System.Drawing.Size(312, 19);
			this.txtJScriptFile2.TabIndex = 42;
			this.txtJScriptFile2.Text = "";
			// 
			// chkJScriptFile2
			// 
			this.chkJScriptFile2.Location = new System.Drawing.Point(408, 144);
			this.chkJScriptFile2.Name = "chkJScriptFile2";
			this.chkJScriptFile2.Size = new System.Drawing.Size(216, 24);
			this.chkJScriptFile2.TabIndex = 44;
			this.chkJScriptFile2.Text = "Valid(Add to End of Each Template)";
			// 
			// PageConfigTemplate
			// 
			this.Controls.Add(this.txtJScriptFile1);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btnJScriptFile1);
			this.Controls.Add(this.btnRemvAll);
			this.Controls.Add(this.txtXmlPath);
			this.Controls.Add(this.txtTempTest);
			this.Controls.Add(this.lstTemp);
			this.Controls.Add(this.btnLoadXML);
			this.Controls.Add(this.btnTempTest);
			this.Controls.Add(this.btnSelAll);
			this.Controls.Add(this.btnSelAllUN);
			this.Controls.Add(this.btnJScriptFile2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtJScriptFile2);
			this.Controls.Add(this.chkJScriptFile1);
			this.Controls.Add(this.chkJScriptFile2);
			this.Name = "PageConfigTemplate";
			this.Size = new System.Drawing.Size(648, 272);
			this.ResumeLayout(false);

		}
		#endregion

		private void btnTempTest_Click(object sender, System.EventArgs e)
		{
			//frmMain.msg.clear();
			frmMain.labStatus.Text = "";
			if(lstTemp.CheckedItems.Count < 1)
			{
				frmMain.labStatus.Text = "no template be selected in list.";
				return;
			}

			XmlDocument xmldoc2 = getXMLTemplate();
			if(xmldoc2 == null)
			{
				frmMain.labStatus.Text = "XML file is not valid.";
				return;
			}
			XmlNodeList nodeList = xmldoc2.SelectNodes("config/template");

			//create new IPara
			IPara ipara = frmMain.CreateIPara();

			//set filed value
			ipara = frmMain.CreateSampleFiled(ipara);

			//テンプレートごとに出力処理
			txtTempTest.Text = "";
			for(int loopj = 0; loopj < nodeList.Count; loopj++)
			{
				ipara.TemplateNode = nodeList[loopj];
				string stitle = "";
				if(nodeList[loopj].Attributes["title"] != null)
				{
					stitle = nodeList[loopj].Attributes["title"].InnerText;
					if(nodeList[loopj].Attributes["language"] != null)
					{
						stitle = "[" + nodeList[loopj].Attributes["language"].InnerText + "] " + nodeList[loopj].Attributes["title"].InnerText;
					}
				}
				txtTempTest.AppendText("##################################################\r\n");
				txtTempTest.AppendText("# " + stitle + "\r\n");
				txtTempTest.AppendText("##################################################\r\n");
				string sFileTxt = ClassExt.CreateClsFromTempString(ipara);
				if(sFileTxt != null)
				{
					txtTempTest.AppendText(sFileTxt);
				}
				else
				{
					txtTempTest.AppendText("  this file is Canceled by ipara\r\n");
				}
			}
			frmMain.labStatus.Text = "only the selected template is created.";
		}

		private void btnLoadXML_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			frmMain.msg.clear();
			OpenFileDialog openFileDialog1 = new OpenFileDialog();
			if(lstTemp.Tag != null && File.Exists(lstTemp.Tag.ToString()))
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(lstTemp.Tag.ToString());
			}
			else
			{
				openFileDialog1.InitialDirectory = Path.GetFullPath(Application.ExecutablePath);
			}
			openFileDialog1.Filter = "クラステンプレート配置情報ファイル(*.xml)|*.xml";
			openFileDialog1.Title = "クラステンプレートファイルを選択してください。";
			openFileDialog1.Multiselect = true;

			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				string[] sSelFileName = openFileDialog1.FileNames;
				lstTemp.Tag = sSelFileName[0];
				for(int i = 0; i < sSelFileName.Length; i++)
				{
					if(lstTemp.FindString(sSelFileName[i]) >= 0)
					{
						continue;
					}

					frmMain.labStatus.Text = "Now is getting XML file information...";
					Application.DoEvents();
					XmlDocument xmldoc2 = getXMLTemplate(sSelFileName[i]);
					if(xmldoc2 == null)
					{
						frmMain.msg.println("Can not Load, Skip:" + sSelFileName[i]);
						continue;
					}

					XmlNodeList nodeList = xmldoc2.SelectNodes("config/template");
					if(nodeList.Count > 0)
					{
						lstTemp.Items.Add(sSelFileName[i]);
					}
					for(int nodei = 0; nodei < nodeList.Count; nodei++)
					{
						if(nodeList[nodei].Attributes["title"] != null)
						{
							string slang = "";
							if(nodeList[nodei].Attributes["language"] != null)
							{
								slang = "[" + nodeList[nodei].Attributes["language"].InnerText + "] ";
							}
							lstTemp.Items.Add("　　" + slang + nodeList[nodei].Attributes["title"].InnerText);
						}
					}
					frmMain.labStatus.Text = "Now is getting XML file information...OK";
				}
			}
		}

		private void btnSelAll_Click(object sender, System.EventArgs e)
		{
			for(int i = 0; i < lstTemp.Items.Count; i++)
			{
				lstTemp.SetItemChecked(i, true);
			}
		}

		private void btnSelAllUN_Click(object sender, System.EventArgs e)
		{
			//lstTemp.ClearSelected();
			for(int i = 0; i < lstTemp.Items.Count; i++)
			{
				lstTemp.SetItemChecked(i, false);
			}
		}

		private void btnRemvAll_Click(object sender, System.EventArgs e)
		{
			lstTemp.Items.Clear();
		}

		public XmlDocument getXMLTemplate()
		{
			//reload every time
			try
			{
				xmldoc = new XmlDocument();
				XmlDocument xmldoc2 = null;
				string sFullFileName = null;
				for(int i = 0; i < lstTemp.Items.Count; i++)
				{
					string sItem = lstTemp.Items[i].ToString();
					if(sItem.Substring(1, 2).Equals(":\\"))
					{
						sFullFileName = sItem;
						xmldoc2 = null;
						continue;
					}
					else if(lstTemp.GetItemChecked(i))
					{
						if(sFullFileName != null && xmldoc2 == null)
						{
							try
							{
								xmldoc2 = new XmlDocument();
								xmldoc2.Load(sFullFileName);
							}
							catch
							{
								sFullFileName = null;
							}
						}
						if(sFullFileName == null || xmldoc2 == null)
						{
							frmMain.msg.println("Can not Load XML, Skip:" + sItem);
							continue;
						}

						//see list item is or not below to xml item
						XmlNodeList nodeList = xmldoc2.SelectNodes("config/template");
						bool hasFound = false;
						for(int nodei = 0; nodei < nodeList.Count; nodei++)
						{
							if(nodeList[nodei].Attributes["title"] != null)
							{
								string slang = "";
								if(nodeList[nodei].Attributes["language"] != null)
								{
									slang = "[" + nodeList[nodei].Attributes["language"].InnerText + "] ";
								}
								string sNodeItem = "　　" + slang + nodeList[nodei].Attributes["title"].InnerText;
								if(sNodeItem.Equals(sItem))
								{
									//copy node
									hasFound = true;
									if(xmldoc.InnerXml.Equals(""))
									{
										xmldoc.InnerXml = "<config></config>";
									}
									xmldoc.DocumentElement.AppendChild(xmldoc.ImportNode(nodeList[nodei], true));
								}
							}
						}
						if(!hasFound)
						{
							frmMain.msg.println("not found Template:" + sItem);
						}
					}
				}
				if(xmldoc!= null && xmldoc.SelectNodes("config/template").Count < 1)
				{
					xmldoc = null;
				}
			}
			catch(Exception exp)
			{
				xmldoc = null;
				frmMain.msg.println("XMLのLoadにエラーが発生しました：");
				frmMain.msg.println(exp.Message, Color.Red);
				frmMain.labStatus.Text = "XMLfile is not valid.";
			}
			return xmldoc;
		}

		public XmlDocument getXMLTemplate(string sFullFileName)
		{
			//reload every time
			XmlDocument xmldoc2 = null;
			if(!File.Exists(sFullFileName))
			{
				frmMain.msg.println("XML file not exist.");
				return null;
			}
			try
			{
				xmldoc2 = new XmlDocument();
				xmldoc2.Load(sFullFileName);
			}
			catch(Exception exp)
			{
				xmldoc2 = null;
				frmMain.msg.println("Error when load XML,maybe need save as UTF-8：");
				frmMain.msg.println(exp.Message, Color.Red);
				frmMain.labStatus.Text = "XMLfile is not valid.";
			}
			return xmldoc2;
		}

		public void Config_Load(System.Collections.Specialized.NameValueCollection coll)
		{
			//JScript File
			if(coll.Get("txtJScriptFile1") != null)
			{
				txtJScriptFile1.Text = coll.Get("txtJScriptFile1");
			}
			if(coll.Get("txtJScriptFile1Valid") != null)
			{
				chkJScriptFile1.Checked = coll.Get("txtJScriptFile1Valid").Equals("1") ? true : false;
			}
			if(coll.Get("txtJScriptFile2") != null)
			{
				txtJScriptFile2.Text = coll.Get("txtJScriptFile2");
			}
			if(coll.Get("txtJScriptFile2Valid") != null)
			{
				chkJScriptFile2.Checked = coll.Get("txtJScriptFile2Valid").Equals("1") ? true : false;
			}
			//last load Xml path
			if(coll.Get("txtXmlPath") != null)
			{
				lstTemp.Tag = coll.Get("txtXmlPath");
			}
			//restore excel file(template) list
			if(coll["XMLTemplateList"] != null)
			{
				string[] slist = coll["XMLTemplateList"].Replace("\n","").Split('\r');
				for (int i = 0; i < slist.Length; i++)
				{
					string line = slist[i];
					int npos = line.LastIndexOf("=");
					if(npos > 0)
					{
						lstTemp.Items.Add(line.Substring(0, npos));
						if(line.Substring(npos + 1).Equals("on"))
						{
							lstTemp.SetItemChecked(lstTemp.Items.Count - 1, true);
						}
					}
				}
			}
		}

		public void Config_Save(System.IO.StreamWriter sw)
		{
			//JScript File
			sw.Write("txtJScriptFile1={0}\r\n", txtJScriptFile1.Text);
			sw.Write("txtJScriptFile1Valid={0}\r\n", chkJScriptFile1.Checked ? "1" : "0");
			sw.Write("txtJScriptFile2={0}\r\n", txtJScriptFile2.Text);
			sw.Write("txtJScriptFile2Valid={0}\r\n", chkJScriptFile2.Checked ? "1" : "0");
			//last load XML path
			sw.Write("txtXmlPath={0}\r\n", "" + lstTemp.Tag);
			//save excel file(template) list
			for(int i = 0; i < lstTemp.Items.Count; i++)
			{
				if(lstTemp.GetItemChecked(i))
				{
					sw.Write("XMLTemplateList={0}={1}\r\n", lstTemp.Items[i].ToString(), "on");
				}
				else
				{
					sw.Write("XMLTemplateList={0}={1}\r\n", lstTemp.Items[i].ToString(), "off");
				}
			}
		}

		bool bOnlyOneIn = false; //for not reloop in CheckListBox_ItemCheck
		private void lstTemp_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			if(bOnlyOneIn)
			{
				return;
			}
			int nindex = lstTemp.SelectedIndex;
			if(nindex < 0)
			{
				return;
			}
			bOnlyOneIn = true;
			if(lstTemp.Items[nindex].ToString().Substring(1, 2).Equals(":\\"))
			{
				//all sheets of File select or unselect
				for(int i = nindex + 1; i < lstTemp.Items.Count; i++)
				{
					if(lstTemp.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						break;
					}
					lstTemp.SetItemChecked(i, (e.NewValue == CheckState.Checked));
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
				for(int i = nindex + 1; i < lstTemp.Items.Count; i++)
				{
					if(lstTemp.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						break;
					}
					nend = i;
				}
				for(int i = nend; i >= 0; i--)
				{
					if(lstTemp.Items[i].ToString().Substring(1, 2).Equals(":\\"))
					{
						if(balselall && !balselno)
						{
							lstTemp.SetItemChecked(i, true);
						}
						else
						{
							lstTemp.SetItemChecked(i, false);
						}
						break;
					}
					if(i != nindex)
					{
						if(lstTemp.GetItemChecked(i))
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

		public string getJScriptHead()
		{
			//get string that add to head of JScript
			string sJSTxtHead = "";
			string fName = txtJScriptFile1.Text.Trim();
			if(chkJScriptFile1.Checked && !fName.Equals(""))
			{
				if(fName.StartsWith(".\\") || fName.StartsWith("./"))
				{
					fName = Path.GetDirectoryName(Application.ExecutablePath) + fName.Substring(1);
				}
				if(File.Exists(fName))
				{
					//if read error then null
					sJSTxtHead = cc.Util.readAll(fName);
				}
				else
				{
					//for out some msg
					return null;
				}
			}
			return sJSTxtHead;
		}

		public string getJScriptEnd()
		{
			//get string that add to end of JScript
			string sJSTxtEnd = "";
			string fName = txtJScriptFile2.Text.Trim();
			if(chkJScriptFile2.Checked && !fName.Equals(""))
			{
				if(fName.StartsWith(".\\") || fName.StartsWith("./"))
				{
					fName = Path.GetDirectoryName(Application.ExecutablePath) + fName.Substring(1);
				}
				if(File.Exists(fName))
				{
					//if read error then null
					sJSTxtEnd = cc.Util.readAll(fName);
				}
				else
				{
					//for out some msg
					return null;
				}
			}
			return sJSTxtEnd;
		}

		private void btnJScriptFile1_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			string initpath = Path.GetDirectoryName(Application.ExecutablePath);
			string sfilename = txtJScriptFile1.Text.Trim();
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
				openFileDialog1.InitialDirectory = initpath;
			}
			openFileDialog1.Filter = "JScriptファイル(*.txt; *.js)|*.txt; *.js|JScriptファイル(*.*)|*.*";
			openFileDialog1.Title = "JScriptファイルを選択してください。";

			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				if(Path.GetDirectoryName(openFileDialog1.FileName).StartsWith(initpath))
				{
					txtJScriptFile1.Text = openFileDialog1.FileName.Substring(initpath.Length);
					if(txtJScriptFile1.Text.StartsWith("\\"))
					{
						txtJScriptFile1.Text = "." + txtJScriptFile1.Text;
					}
				}
				else
				{
					txtJScriptFile1.Text = openFileDialog1.FileName;
				}
			}
		}

		private void btnJScriptFile2_Click(object sender, System.EventArgs e)
		{
			frmMain.labStatus.Text = "";
			string initpath = Path.GetDirectoryName(Application.ExecutablePath);
			string sfilename = txtJScriptFile2.Text.Trim();
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
				openFileDialog1.InitialDirectory = initpath;
			}
			openFileDialog1.Filter = "JScriptファイル(*.txt; *.js)|*.txt; *.js|JScriptファイル(*.*)|*.*";
			openFileDialog1.Title = "JScriptファイルを選択してください。";

			if(openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				if(Path.GetDirectoryName(openFileDialog1.FileName).StartsWith(initpath))
				{
					txtJScriptFile2.Text = openFileDialog1.FileName.Substring(initpath.Length);
					if(txtJScriptFile2.Text.StartsWith("\\"))
					{
						txtJScriptFile2.Text = "." + txtJScriptFile2.Text;
					}
				}
				else
				{
					txtJScriptFile2.Text = openFileDialog1.FileName;
				}
			}
		}

	}
}
