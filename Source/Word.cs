using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
//using System.Windows.Forms;

using Word;

namespace File
{
    public partial class Word : System.Windows.Forms.Form
    {
        public Word()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            object oFileName = @"C:\ddddddddd.doc";
            object oReadOnly = false;
            object oMissing = System.Reflection.Missing.Value;
            object Nothing = System.Reflection.Missing.Value;
            Application oWord;
            Document oDoc;
            oWord = new Application();
            oWord.Visible = true;//只是为了方便观察
            oDoc = oWord.Documents.Open(ref oFileName, ref oMissing, ref oReadOnly, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            //MessageBox.Show(oDoc.Tables.Count.ToString());

            Table newTable = oDoc.Tables[1];
            //设置表格样式 
            newTable.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleThickThinLargeGap;
            newTable.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            newTable.Columns[1].Width = 100f;
            newTable.Columns[2].Width = 220f;
            newTable.Columns[3].Width = 105f;

            //填充表格内容 
            newTable.Cell(1, 1).Range.Text = "产品详细信息表";
            newTable.Cell(1, 1).Range.Bold = 2; //设置单元格中字体为粗体 

            //删除行
            newTable.Rows[newTable.Rows.Count].Delete();

            //移动到下一页
            WdGoToItem goPage = WdGoToItem.wdGoToPage;
            oWord.Selection.GoToNext(goPage);

            //插入图片
            string FileName = @"C:\dsdddd.jpg";//图片所在路径
            object LinkToFile = false;
            object SaveWithDocument = true;
            object Anchor = oWord.Application.Selection.Range;
            oWord.Selection.Document.Shapes.AddPicture(FileName, ref LinkToFile, ref SaveWithDocument, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Anchor);

            oWord.Selection.TypeText("我们的村村sfsfsdfdddddddddd");



            object filename = @"C:\eeeeee.doc";
            //文件保存 
            oDoc.SaveAs(ref filename, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                          ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing, ref Nothing,
                          ref Nothing, ref Nothing, ref Nothing);
            oDoc.Close(ref Nothing, ref Nothing, ref Nothing);
            oWord.Quit(ref Nothing, ref Nothing, ref Nothing);
        }

        private void OpenWord_Click(object sender, EventArgs e)
        {
            Application wordApp = new ApplicationClass(); try
            {
                // Word程序可见
                wordApp.Visible = true; object missing = System.Reflection.Missing.Value; object fileName = @"C:\a.doc";                // 打开Word文档
                wordApp.Documents.Open(
                        ref fileName, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing
                    ); Range range1, range2;                //// 取得总页数
                //int pageNumber = wordApp.ActiveDocument.ComputeStatistics(
                //        Word.WdStatistic.wdStatisticPages, 
                //        ref missing
                //    );                // 跳转种类
                object objWhat = WdGoToItem.wdGoToPage;
                // 跳转位置
                object objWhich = WdGoToDirection.wdGoToLast;                // 转向最后一页
                wordApp.Selection.GoTo(ref objWhat, ref objWhich, ref missing, ref missing);                // Range取得
                range1 = wordApp.Selection.Range;
                range2 = wordApp.ActiveDocument.Range(ref missing, ref missing); object start = range1.Start;
                object end = range2.End;                // 删除最后一页
                wordApp.ActiveDocument.Range(ref start, ref end).Delete(ref missing, ref missing);                // 跳转至最前
                objWhich = WdGoToDirection.wdGoToFirst;
                wordApp.Selection.GoTo(ref objWhat, ref objWhich, ref missing, ref missing); wordApp.ActiveDocument.Save(); wordApp.Quit(ref missing, ref missing, ref missing);
            }
            finally
            {
                if (null != wordApp)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }
        }


        //自定义变量
        private ApplicationClass WordApp = null; //word
        private Document document = null;//document  
        private object missing = null;//缺省值

        /// <summary>
        /// 将光标移动至文件头部
        /// </summary>
        public void moveStart()
        {
            object story = WdUnits.wdStory;
            WordApp.Selection.HomeKey(ref story, ref missing);//向上移到顶
        }

        //移动到下一页
        private void movePage(int count)
        {
            for (int i = 0; i < count; i++)
            {
                WdGoToItem goPage = WdGoToItem.wdGoToPage;
                WordApp.Selection.GoToNext(goPage);
            }
        }

        /// <summary>
        /// 插入文本
        /// </summary>
        /// <param name="text"></param>
        public void insertText(string text)
        {
            WordApp.Selection.TypeText(text);
        }

        /// <summary>
        /// 删除文本
        /// </summary>
        /// <param name="count"></param>
        public void deleteText(object count)
        {
            WordApp.Selection.Delete(ref missing, ref count);
        }




        /// <summary>
        /// 选择所有图形
        /// </summary>
        public void sel()
        {
            document.Shapes.SelectAll();
        }
    }
}
