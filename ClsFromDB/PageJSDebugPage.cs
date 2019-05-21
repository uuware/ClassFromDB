using System;
using System.Windows.Forms;
using System.IO;

namespace ClsFromDB
{
	/// <summary>
	/// PageFromServer の概要の説明です。
	/// </summary>
	public class PageJSDebugPage : System.Windows.Forms.UserControl
	{
		/// <summary>
		/// 必要なデザイナ変数です。
		/// </summary>
		private System.ComponentModel.Container components = null;
		public System.Windows.Forms.CheckBox chkReset;
		public System.Windows.Forms.NumericUpDown numRunCnt;
		private System.Windows.Forms.Label label1;
		public System.Windows.Forms.CheckBox chkAddJS;
		private System.Windows.Forms.Button btnSaveTo;
		private System.Windows.Forms.Button btnOpenFile;
		public System.Windows.Forms.RichTextBox txtJS;

		public PageJSDebugPage()
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
			this.chkReset = new System.Windows.Forms.CheckBox();
			this.numRunCnt = new System.Windows.Forms.NumericUpDown();
			this.label1 = new System.Windows.Forms.Label();
			this.txtJS = new System.Windows.Forms.RichTextBox();
			this.chkAddJS = new System.Windows.Forms.CheckBox();
			this.btnSaveTo = new System.Windows.Forms.Button();
			this.btnOpenFile = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.numRunCnt)).BeginInit();
			this.SuspendLayout();
			// 
			// chkReset
			// 
			this.chkReset.Location = new System.Drawing.Point(136, 0);
			this.chkReset.Name = "chkReset";
			this.chkReset.Size = new System.Drawing.Size(96, 24);
			this.chkReset.TabIndex = 11;
			this.chkReset.Text = "Engine Reset";
			// 
			// numRunCnt
			// 
			this.numRunCnt.Location = new System.Drawing.Point(80, 0);
			this.numRunCnt.Maximum = new System.Decimal(new int[] {
																	  99999,
																	  0,
																	  0,
																	  0});
			this.numRunCnt.Minimum = new System.Decimal(new int[] {
																	  1,
																	  0,
																	  0,
																	  0});
			this.numRunCnt.Name = "numRunCnt";
			this.numRunCnt.Size = new System.Drawing.Size(48, 19);
			this.numRunCnt.TabIndex = 15;
			this.numRunCnt.Value = new System.Decimal(new int[] {
																	1,
																	0,
																	0,
																	0});
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(0, 5);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(88, 16);
			this.label1.TabIndex = 13;
			this.label1.Text = "TestRunCount:";
			// 
			// txtJS
			// 
			this.txtJS.AcceptsTab = true;
			this.txtJS.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtJS.AutoWordSelection = true;
			this.txtJS.HideSelection = false;
			this.txtJS.Location = new System.Drawing.Point(0, 24);
			this.txtJS.Name = "txtJS";
			this.txtJS.Size = new System.Drawing.Size(600, 248);
			this.txtJS.TabIndex = 17;
			this.txtJS.Text = "";
			// 
			// chkAddJS
			// 
			this.chkAddJS.Location = new System.Drawing.Point(248, 0);
			this.chkAddJS.Name = "chkAddJS";
			this.chkAddJS.Size = new System.Drawing.Size(176, 24);
			this.chkAddJS.TabIndex = 11;
			this.chkAddJS.Text = "Include inner JScript Define";
			// 
			// btnSaveTo
			// 
			this.btnSaveTo.Location = new System.Drawing.Point(488, 2);
			this.btnSaveTo.Name = "btnSaveTo";
			this.btnSaveTo.Size = new System.Drawing.Size(56, 19);
			this.btnSaveTo.TabIndex = 23;
			this.btnSaveTo.Text = "SaveTo";
			this.btnSaveTo.Click += new System.EventHandler(this.btnSaveTo_Click);
			// 
			// btnOpenFile
			// 
			this.btnOpenFile.Location = new System.Drawing.Point(424, 2);
			this.btnOpenFile.Name = "btnOpenFile";
			this.btnOpenFile.Size = new System.Drawing.Size(64, 19);
			this.btnOpenFile.TabIndex = 22;
			this.btnOpenFile.Text = "OpenFile";
			this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);
			// 
			// PageJSDebugPage
			// 
			this.Controls.Add(this.btnSaveTo);
			this.Controls.Add(this.btnOpenFile);
			this.Controls.Add(this.chkReset);
			this.Controls.Add(this.numRunCnt);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtJS);
			this.Controls.Add(this.chkAddJS);
			this.Name = "PageJSDebugPage";
			this.Size = new System.Drawing.Size(600, 272);
			((System.ComponentModel.ISupportInitialize)(this.numRunCnt)).EndInit();
			this.ResumeLayout(false);

		}
		#endregion

		private void btnOpenFile_Click(object sender, System.EventArgs e)
		{
			string initpath = Path.GetDirectoryName(Application.ExecutablePath);
			string sfilename = "" + txtJS.Tag;
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
				txtJS.Tag = openFileDialog1.FileName;
				txtJS.Text = cc.Util.readAll(openFileDialog1.FileName);
			}
		}

		private void btnSaveTo_Click(object sender, System.EventArgs e)
		{
			string initpath = Path.GetDirectoryName(Application.ExecutablePath);
			string sfilename = "" + txtJS.Tag;
			if(sfilename.StartsWith(".\\"))
			{
				sfilename = initpath + sfilename.Substring(1);
			}
			SaveFileDialog saveFileDialog1 = new SaveFileDialog();
			if(File.Exists(sfilename))
			{
				saveFileDialog1.InitialDirectory = Path.GetFullPath(sfilename);
			}
			else
			{
				saveFileDialog1.InitialDirectory = initpath;
			}
			saveFileDialog1.Filter = "JScriptファイル(*.txt; *.js)|*.txt; *.js|JScriptファイル(*.*)|*.*";
			saveFileDialog1.Title = "JScriptファイルを選択してください。";

			if(saveFileDialog1.ShowDialog() == DialogResult.OK)
			{
				txtJS.Tag = saveFileDialog1.FileName;
				cc.Util.writeFile(saveFileDialog1.FileName, txtJS.Text);
			}
		}

	}
}
