using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddIn2
{
    using System;

    public class SlideKeywordInfo
    {
        public int SlideIndex { get; set; }//slideインデックス
        public List<string> keywords { get;set; }//マークされたキーワードのリスト

        public SlideKeywordInfo(int slideIndex)
        {
            SlideIndex = slideIndex;
            keywords = new List<string>();
        }
    }
}
