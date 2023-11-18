using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn2
{
    public partial class MyUserControl : UserControl
    {
        public MyUserControl()
        {
            InitializeComponent();
        }


        private void MarkRedButton_Click(object sender, EventArgs e)
        {
            PowerPoint.Selection selection = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            //PowerPoint.Application pptApp = Globals.ThisAddIn.Application;
            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                PowerPoint.TextRange selectedText = selection.TextRange;

                selectedText.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red); // フォントの色を赤に設定
                selectedText.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue; // テキストを強調表示（太字）
                selectedText.Font.Underline = Microsoft.Office.Core.MsoTriState.msoTrue; // アンダーラインを引く


                // 選択範囲の背景を赤くする方法はサポートされてなさそう

                //現在のスライドのインデックスを取得
                int currentSlideIndex = Globals.ThisAddIn.Application.ActiveWindow.View.Slide.SlideIndex;
                //マークしたキーワードを追加
                string keyword = selectedText.Text;
                //ThisAddIn.csのメソッドを呼び出しキーワードをスライドごとに保存
                Globals.ThisAddIn.AddSelectedTextToSlideKeyWords(keyword, currentSlideIndex);

                //リストボックスにキーワードを表示
                List<string> keywordsForCurrentSlide = Globals.ThisAddIn.GetKeywordsForSlide(currentSlideIndex);

                listBox1.Items.Clear();
                listBox1.Items.AddRange(keywordsForCurrentSlide.ToArray());
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
 
    
}
