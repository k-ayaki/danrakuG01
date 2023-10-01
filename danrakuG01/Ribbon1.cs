using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using danrakuG01;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace danrakuG01
{
    public partial class Ribbon1
    {
        public Document doc;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }
        private void AddDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            Range rangeSave = doc.Application.Selection.Range;
            DelDanraku delDanraku = new DelDanraku(doc);
            if(delDanraku.m_cancel == false && delDanraku.m_error == false)
            {
                AddDanraku addDanraku = new AddDanraku(doc);
                addDanraku.Dispose();
            }
            delDanraku.Dispose();
            rangeSave.Select();
        }

        private void RenumDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            Range rangeSave = doc.Application.Selection.Range;
            RenumDanraku renumDanraku = new RenumDanraku(doc);
            renumDanraku.Dispose();
            rangeSave.Select();
        }

        private void DelDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            Range rangeSave = doc.Application.Selection.Range;
            DelDanraku delDanraku = new DelDanraku(doc);
            delDanraku.Dispose();
            rangeSave.Select();
        }
        // プログラムによって文書内のテキストを検索および置換する
        // https://docs.microsoft.com/ja-jp/visualstudio/vsto/how-to-programmatically-search-for-and-replace-text-in-documents?view=vs-2019
        // プログラムによって検索後に選択を復元する
        // https://docs.microsoft.com/ja-jp/visualstudio/vsto/how-to-programmatically-restore-selections-after-searches?view=vs-2019
        public void 垂直タブを改行に(Document doc)
        {
            object missing = null;

            Range rangeSave = doc.Application.Selection.Range;
            doc.Application.Selection.WholeStory();
            Find findObject = doc.Application.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = "^l";
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = "^p";
            findObject.MatchFuzzy = false;
            findObject.Forward = true;
            object findtext = "^l";
            object replacetext = "^p";
            object replaceAll = WdReplace.wdReplaceAll;
            findObject.Execute(ref findtext, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref replacetext,
                ref replaceAll, ref missing, ref missing, ref missing, ref missing);
            findObject.ClearFormatting();
            rangeSave.Select();
        }
    }
}
