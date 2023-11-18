using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn2
{
    public partial class ThisAddIn
    {
        private Dictionary<int, SlideKeywordInfo> slideKeywordsDictionary = new Dictionary<int, SlideKeywordInfo>();//
        private MyUserControl myUserControl1;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;//CustomTaskPaneのインスタンスをThisAddInクラスのメンバとして宣言

        public class SlideKeywordInfo
        {
            public int SlideIndex { get; set; }//slideインデックス
            public List<string> keywords { get; set; }//マークされたキーワードのリスト

            public SlideKeywordInfo(int slideIndex)
            {
                SlideIndex = slideIndex;
                keywords = new List<string>();
            }
        }


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //作業ウィンドの作成
            myUserControl1 = new MyUserControl();
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;

            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
            //slide切り替え時のイベントハンドラを設定
            this.Application.SlideSelectionChanged += Application_SlideSelectionChanged;
        }
        //ハイライト
        public void HighlightKeyword(PowerPoint.Selection Sel)
        {
            if(Sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
               
            }
        }

        private void Application_SlideSelectionChanged(PowerPoint.SlideRange SldRange)
        {
            myUserControl1.listBox1.Items.Clear();
        }

        //パワポ内のドキュメントのテキストを選択したときのイベント
        //選択したオブジェクトのSelが引数
        //Sel.Typeで制御(テキストか図形かなど)
        //TextRange:選択されたテキストのテキスト範囲を取得

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionText)//テキストが選択されている場合
            {
                System.Diagnostics.Debug.WriteLine(Sel.TextRange.Text);
            }else if(Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)//図形が選択されている場合
            {
                System.Diagnostics.Debug.WriteLine("図形");
            }
        }
        //選択したテキストとそのslide情報をディクショナリに追加するメソッド
        //スライドごとにキーワード情報を更新
        public void AddSelectedTextToSlideKeyWords(string selectedText,int slideIndex)
        {
            if (!slideKeywordsDictionary.ContainsKey(slideIndex))
            {
                slideKeywordsDictionary[slideIndex] = new SlideKeywordInfo(slideIndex);
            }
            slideKeywordsDictionary[slideIndex].keywords.Add(selectedText);
        }
        //slideのキーワードリストを取得するメソッド
        public List<string> GetKeywordsForSlide(int slideIndex)
        {
            if (slideKeywordsDictionary.ContainsKey(slideIndex))
            {
                return slideKeywordsDictionary[slideIndex].keywords;
            }
            else
            {
                return new List<string>();//スライド内にマーク付けされたキーワードがない場合は空のリストを返す
            }
        }


 

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
