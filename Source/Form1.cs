using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
// Download by http://www.codefans.net
using System.Data;
using Word;
namespace MyFile
{
	/// <summary>
	/// Form1 ��ժҪ˵����
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.SaveFileDialog saveFileDialog1;
		/// <summary>
		/// ����������������
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Windows ���������֧���������
			//
			InitializeComponent();

			//
			// TODO: �� InitializeComponent ���ú�����κι��캯������
			//
		}

		/// <summary>
		/// ������������ʹ�õ���Դ��
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows ������������ɵĴ���
		/// <summary>
		/// �����֧������ķ��� - ��Ҫʹ�ô���༭���޸�
		/// �˷��������ݡ�
		/// </summary>
		private void InitializeComponent()
		{
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(104, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "������������Ϣ��";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(184, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(144, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "����ΪWORD�ļ���ʽ";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// richTextBox1
			// 
			this.richTextBox1.Location = new System.Drawing.Point(8, 32);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(320, 176);
			this.richTextBox1.TabIndex = 2;
			this.richTextBox1.Text = "";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(6, 14);
			this.ClientSize = new System.Drawing.Size(336, 214);
			this.Controls.Add(this.richTextBox1);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label1);
			this.MaximizeBox = false;
			this.Name = "Form1";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "��ʾ����WORD�ļ�";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// Ӧ�ó��������ڵ㡣
		/// </summary>
		[STAThread]
		static void Main() 
		{
			System.Windows.Forms.Application.Run(new Form1());
		}

		private void button1_Click(object sender, System.EventArgs e)
		{//����ΪWORD�ļ�
            //if (this.richTextBox1.Text == "")
            //{
            //    //File.Word frmWord = new File.Word();
            //    //frmWord.ShowDialog();
            //    return;
            //}
			if(this.saveFileDialog1.ShowDialog()==DialogResult.Cancel)
				return;
			string FileName=this.saveFileDialog1.FileName;
			if(FileName.Length<1)
				return;
			FileName+=".doc";
			try
			{
				Word.ApplicationClass MyWord=new Word.ApplicationClass(); 
				Word.Document MyDoc;				
				Object Nothing=System.Reflection.Missing.Value; 
				MyDoc=MyWord.Documents.Add(ref Nothing,ref Nothing,ref Nothing,ref Nothing);
                //MyDoc.Paragraphs.CloseUp();
                object pBreak = (int)WdBreakType.wdSectionBreakContinuous; //wdSectionBreakNextPage
                //para.Range.InsertBreak(ref pBreak);

                //MyDoc.Paragraphs.Last.Range.InsertBreak(pBreak);
                //MyDoc.Paragraphs.Last.Range.InsertParagraphAfter();
                //MyDoc.Reload();
                //MyDoc.Paragraphs.Last.Range.InsertParagraphAfter();

                Range curretRange = MyDoc.Paragraphs.Last.Range;
                curretRange.Text = "Ĭ��խ������SS";

                MyDoc.Paragraphs.First.Range.Text = "��ҳ";

                //MyDoc.Paragraphs.Last.Range.Text="Ĭ��խ������22";
                //MyDoc.Paragraphs.Last.Range.Text = "Ĭ��խ������";
                //MyDoc.Paragraphs.Last.Range.Text = "Ĭ��խ������SS";
                
                //gotoLastLine(MyDoc);
                //MyDoc.Paragraphs.Last.Range.MoveEnd();
                //object pBreak = (int)WdBreakType.wdSectionBreakNextPage;
                //para.Range.InsertBreak(ref pBreak);
                MyDoc.Paragraphs.Add(ref Nothing);
                MyDoc.Paragraphs.Last.Range.InsertBreak(pBreak);
                //MyDoc.Paragraphs.Last.Range.InsertParagraphAfter();
                Range curretRange2 = MyDoc.Paragraphs.Last.Range;
                curretRange2.Text = "�����ʾ������";
                //MyDoc.Paragraphs.Last.Range.Text = "�����ʾ������";
                MyDoc.Paragraphs.Last.Range.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                
                ////MyDoc.Paragraphs.Last.Range.MoveEnd();
                //gotoLastLine(MyDoc);

                MyDoc.Paragraphs.Add(ref Nothing);
                MyDoc.Paragraphs.Last.Range.InsertBreak(pBreak);
                //MyDoc.Paragraphs.Last.Range.InsertParagraphAfter();
                //MyDoc.Paragraphs.Last.Range.Text = "���������µ����ݰ�";
                Range curretRange3 = MyDoc.Paragraphs.Last.Range;
                curretRange3.Text = "���������µ����ݰ�";
                MyDoc.Paragraphs.Last.Range.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                //MyDoc.Paragraphs.First.Range.Text = "��ҳ";

                //MyDoc.Paragraphs[0].Range.Text = "Ĭ��խ������SS";
                //MyDoc.Paragraphs[1].Range.Text = "�����ʾ������";
                //MyDoc.Paragraphs[2].Range.Text = "���������µ����ݰ�";
                //gotoLastLine(MyDoc);

                //Range rg=new 
                //MyDoc.Paragraphs.Add()
                //MyDoc.Paragraphs[2].Range.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                //MyDoc = MyWord.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                //MyDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                //MyDoc.Paragraphs.Last.Range.Text = "����һ�¿�ҳ";

                //MyDoc = MyWord.Documents.Add(ref Nothing, ref Nothing, ref Nothing, ref Nothing);
                //MyDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
                //MyDoc.Paragraphs.Last.Range.Text = this.richTextBox1.Text;


                //object o1=0;
                //object o2=1;

                //Word.Range newRange = MyDoc.Range(ref o1, ref o2);
                //newRange.PageSetup.Orientation = WdOrientation.wdOrientPortrait;

                //o1 = 2;
                //o2 = 2;

                //newRange = MyDoc.Range(ref o1, ref o2);
                //newRange.PageSetup.Orientation = WdOrientation.wdOrientLandscape;

                //Word.Range newRange2 = MyDoc.Paragraphs[1].Range;
                //newRange2.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
                object MyFileName=FileName;
				//��WordDoc�ĵ���������ݱ���ΪDOC�ĵ� 
				MyDoc.SaveAs(ref MyFileName,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing,ref Nothing); 
				//�ر�WordDoc�ĵ����� 
				MyDoc.Close(ref Nothing, ref Nothing, ref Nothing); 
				//�ر�WordApp������� 
				MyWord.Quit(ref Nothing, ref Nothing, ref Nothing); 		
				MessageBox.Show("WORD�ļ�����ɹ�","��Ϣ��ʾ",MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
			catch(Exception Err)
			{
				MessageBox.Show("WORD�ļ��������ʧ�ܣ�"+Err.Message,"��Ϣ��ʾ",MessageBoxButtons.OK,MessageBoxIcon.Information);
			}
		}

        private void gotoLastLine(Document thisDocument)
        {

            object dummy = System.Reflection.Missing.Value;

            object what = WdGoToItem.wdGoToLine;

            object which = WdGoToDirection.wdGoToLast;

            object count = 99999999;

            thisDocument.Application.Selection.GoTo(ref what, ref which, ref count, ref dummy);

        }
	}
}
