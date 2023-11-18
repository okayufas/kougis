using System;

public class SlideKeywordInfo
{
	public int SlideIndex { get; set; }//slideインデックス
	public List<string> keywords { get;,set; }//マークされたキーワードのリスト

	public SlideKeywordInfo(int slideIndex)
	{
		SlideIndex = slideIndex;
		keywords = new List<string>();
	}
}
