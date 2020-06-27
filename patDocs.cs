using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;

namespace danrakuG01
{
    class patDoc
    {
        private const string 最初のブックマーク = "BK_738c8dc9_af9f_4e29_adde_934af51b07e2";
        private const string 次のブックマーク = "BK_97414a23_4a86_4371_a3b7_c00da9f7751a";

        private long e_counter = 1;

        public void G_段落番号振直(Document doc)
        {
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            Range rng = 書類名の範囲("明細書", doc);
            if (rng == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                return;
            }
            long counter;
            counter = 0;
            rng.Find.MatchWildcards = true;
            while (rng.Find.Execute("【[０-９]@】"))
            {
                counter++;
                rng.Text = 段落番号文字列の生成(counter);
                rng.SetRange(rng.End, rng.End);
            }
        }

        public string 段落番号文字列の生成(
                    long counter)
        {
            string 段落番号文字列;

            段落番号文字列 = "【" + Strings.StrConv(counter.ToString("0000"), VbStrConv.Wide, 0) + "】";
            return 段落番号文字列;
        }
        public void G_段落番号削除(Document doc)
        {
            Range rng;
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            rng = 書類名の範囲("明細書", doc);
            if (rng == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                return;
            }
            rng.Find.MatchWildcards = true;
            while (rng.Find.Execute("【[０-９]@】"))
            {
                rng.Text = "";
                rng.SetRange(rng.End, rng.End);
                if (パラグラフが空白か判定(rng.Paragraphs[1]))
                {
                    rng.Paragraphs[1].Range.Delete();
                }
            }
        }

        public void G_段落番号付与(Document doc)
        {
            Paragraph paraPrev;
            Range rng;

            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            垂直タブを改行に(doc);
            Range 付与範囲 = 書類名の範囲("明細書", doc);
            e_counter = 1;
            if (付与範囲 == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                return;
            }
            foreach (Paragraph paraCurr in 付与範囲.Paragraphs)
            {

                if (項目の判定(paraCurr))
                {
                    if (パラグラフが段落番号付与対象項目か判定(paraCurr))
                    {
                        rng = 直後への段落番号の挿入(paraCurr);
                        string theParaText = paraCurr.Range.Text;
                        if (theParaText.IndexOf("【符号の説明】") >= 0)
                        {
                            break;
                        }
                    }
                    else if (パラグラフが数化表項目か判定(paraCurr))
                    {
                        paraPrev = テキスト記載パラグラフ取得(paraCurr, -1);
                        if (パラグラフが数化表項目か判定(paraPrev))
                        {
                            直前への段落番号挿入(paraCurr);
                        }
                        paraPrev = null;
                    }
                    else if (パラグラフが段落番号付与非対象項目か判定(paraCurr))
                    {
                        if (パラグラフが不正な段落番号か判定(paraCurr))
                        {
                            paraCurr.Range.Delete();
                        }
                    }
                    else
                    {
                        rng = 直後への段落番号の挿入(paraCurr);
                    }
                }
                else
                {
                    段落番号付与判定(paraCurr);
                }
            }
            G_段落番号振直(doc);
        }

        public void 段落番号付与判定(Paragraph paraCurr)
        {
            Paragraph paraPrev = テキスト記載パラグラフ取得(paraCurr, -1);
            Paragraph paraPrev2 = 有効パラグラフ取得(paraCurr, -1);
            if (パラグラフが空白か判定(paraCurr) == true)
            {
                ; // skip
            }
            else if (パラグラフが数式を含むか判定(paraCurr) == true
            || パラグラフが表を含むか判定(paraCurr) == true
            || パラグラフが画像を含むか判定(paraCurr) == true)
            {
                直前への段落番号挿入(paraCurr);
            }
            else if (パラグラフが数化表項目か判定(paraPrev) == true)
            {
                直前への段落番号挿入(paraCurr);
            }
            else if (paraPrev == null)
            {
                ; // skip
            }
            else
            {
                long 項目までの行数 = 項目までの行数取得(paraCurr, -1);
                if ((見出し判定(paraCurr) == true || 図説明の判定(paraCurr) == true)
                && 3 <= 項目までの行数)
                {
                    直前への段落番号挿入(paraCurr);
                }
                else if (文末の句点判定(paraPrev) == true && 6 <= 項目までの行数)
                {
                    直前への段落番号挿入(paraCurr);
                }
                else if (10 <= 項目までの行数)
                {
                    直前への段落番号挿入(paraCurr);
                }
            }
        }
        public bool 図説明の判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"^[　]*?図[１-９－Ａ-Ｚａ-ｚ]+"))
            {
                return true;
            }
            if (Regex.IsMatch(para.Range.Text, @"図[１-９－Ａ-Ｚａ-ｚ]+(（[Ａ-Ｚａｚ]）|)[はがをで]"))
            {
                return true;
            }
            return false;
        }
        public bool 見出し判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"^[　]*?(（.*?）|《.*?》|＜.*?＞|≪.*?≫|［.*?］|〔.*?〕|｛.*?｝|〈.*?〉|\(.*?\)|\[.*?\])[　]*?[\r\n]$"))
            {
                return true;
            }
            return false;
        }
        public bool 文末の句点判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"。[\r\n\x0b\x0c]$"))
            {
                return true;
            }
            return false;
        }

        public Range 直前への段落番号挿入(Paragraph paraCurr)
        {
            String 段落番号;
            Paragraph paraPrev2 = 有効パラグラフ取得(paraCurr, -1);
            if (パラグラフが段落番号か判定(paraPrev2))
            {
                return null;
            }

            Range rng2 = paraCurr.Range;
            段落番号 = "　" + 段落番号文字列の生成(e_counter);
            Range rng = 直前への文字列挿入(paraCurr, 段落番号);
            if (rng != null)
            {
                e_counter++;
            }
            return rng;
        }

        public Range 直前への文字列挿入(Paragraph paraCurr, String strIns)
        {
            Range rng = paraCurr.Range;
            rng.Collapse(WdCollapseDirection.wdCollapseStart);
            if (rng.Tables.Count > 0
            || rng.OMaths.Count > 0)
            {
                Paragraph paraPrev = パラグラフの取得(paraCurr, -1);
                Range rng2 = paraPrev.Range;
                if (rng2.Tables.Count > 0
                || rng2.OMaths.Count > 0)
                {
                    return null;
                }
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd);
                rng2.InsertAfter("\r" + strIns);
                return rng;
            }
            rng.InsertBefore(strIns + "\r");
            return rng;
        }

        public Range 直後への段落番号の挿入(Paragraph paraCurr)
        {
            Paragraph paraNext2 = パラグラフの取得(paraCurr, 1);
            if (パラグラフが段落番号か判定(paraNext2))
            {
                return null;
            }
            String 段落番号 = "　" + 段落番号文字列の生成(e_counter);
            Range rng = 直後への文字列の挿入(paraCurr, 段落番号);
            if (rng != null)
            {
                e_counter++;
            }
            return rng;
        }
        public Range 直後への文字列の挿入(Paragraph paraCurr, String strIns)
        {
            Range rng;
            Paragraph paraNext = パラグラフの取得(paraCurr, 1);
            if (paraNext.Range.Tables.Count > 0
            || paraNext.Range.OMaths.Count > 0)
            {
                // 次のパラグラフが表を含むときに末尾からInsertAfterすると、表内に書き込まれてしまう
                // これを回避するため、改行文字を領域からトリムして、末尾に改行+文字列を挿入する
                /*
                rng = paraCurr.Range;
                領域終端のトリム(ref rng);
                rng.InsertAfter("\r" + strIns);
                paraNext = パラグラフの取得(paraNext, -1);
                return paraNext.Range;
                */
                paraCurr.Range.Text += strIns + "\r";
                return paraCurr.Range;
            }
            rng = paraCurr.Range;
            rng.Collapse(WdCollapseDirection.wdCollapseEnd);
            if (rng.OMaths.Count > 0
            || rng.Tables.Count > 0)
            {
                return null;
            }
            rng.InsertAfter(strIns + "\r");
            return rng;
        }
        public bool 項目の判定(Paragraph paraCurr)
        {
            Range rng = paraCurr.Range;
            rng.Find.MatchWildcards = true;
            return rng.Find.Execute("【[!】]@】");
        }
        public bool パラグラフが段落番号か判定(Paragraph paraCurr)
        {
            Range rng = paraCurr.Range;
            rng.Find.MatchWildcards = true;
            return rng.Find.Execute("【[０-９]@】");
        }
        public bool パラグラフが空白か判定(Paragraph para)
        {
            string ckStr;
            ckStr = para.Range.Text;
            ckStr = ckStr.Replace("　", "");
            ckStr = ckStr.Replace("\r\n", "");
            ckStr = ckStr.Replace("\r", "");
            ckStr = ckStr.Replace("\n", "");
            ckStr = ckStr.Replace("\x0c", "");
            ckStr = ckStr.Replace("\x0b", "");
            if (ckStr.Length == 0)
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが数式を含むか判定(Paragraph para)
        {
            if (para == null)
            {
                return false;
            }
            if (para.Range.OMaths.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool パラグラフが表を含むか判定(Paragraph para)
        {
            if (para == null)
            {
                return false;
            }
            if (para.Range.Tables.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool パラグラフが画像を含むか判定(Paragraph para)
        {
            if (para == null)
            {
                return false;
            }
            if (para.Range.InlineShapes.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        public bool パラグラフが数化表項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text,
                @"(【化[０-９Ａ-Ｚａ-ｚ．－（）]+】" +
                 "|【数[０-９Ａ-Ｚａ-ｚ．－（）]+】" +
                 "|【表[０-９Ａ-Ｚａ-ｚ．－（）]+】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが数式項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"(【数[０-９Ａ-Ｚａ-ｚ．－（）]+】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが化学式項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"(【化[０-９Ａ-Ｚａ-ｚ．－（）]+】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが表項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"(【表[０-９Ａ-Ｚａ-ｚ．－（）]+】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが段落番号付与対象項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text,
                @"(【技術分野】" +
                 "|【背景技術】" +
                 "|【特許文献】" +
                 "|【非特許文献】" +
                 "|【発明が解決しようとする課題】" +
                 "|【課題を解決するための手段】" +
                 "|【発明の効果】" +
                 "|【図面の簡単な説明】" +
                 "|【発明を実施するための形態】" +
                 "|【実施例[０-９]+】" +
                 "|【産業上の利用可能性】" +
                 "|【符号の説明】" +
                 "|【受託番号】" +
                 "|【配列表フリーテキスト】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが段落番号付与非対象項目か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text,
                @"(【書類名】" +
                 "|【発明の名称】" +
                 "|【[０-９]+】" +
                 "|【特許文献[０-９]+】" +
                 "|【非特許文献[０-９]+】" +
                 "|【図[０-９Ａ-Ｚａ-ｚ．－（）]+】" +
                 "|【先行技術文献】" +
                 "|【発明の概要】" +
                 "|【特許請求の範囲】" +
                 "|【請求項[０-９]+】" +
                 "|【要約】" +
                 "|【課題】" +
                 "|【解決手段】" +
                 "|【選択図】)"))
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが不正な段落番号か判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"(【[０-９]+】)"))
            {
                Paragraph paraPrev = テキスト記載パラグラフ取得(para, -1);
                if (パラグラフが段落番号付与非対象項目か判定(paraPrev))
                {
                    return true;
                }
            }
            return false;
        }
        //指定された方向において、テキストが記載されたパラグラフを取得
        public Paragraph テキスト記載パラグラフ取得(Paragraph paraCurr, long initialcount)
        {
            long movecount = initialcount;
            Paragraph target = パラグラフの取得(paraCurr, movecount);
            while (target != null)
            {
                if (パラグラフが空白か判定(target) == false
                && パラグラフが数式を含むか判定(target) == false
                && パラグラフが表を含むか判定(target) == false
                && パラグラフが画像を含むか判定(target) == false)
                {
                    break;
                }
                movecount += initialcount;
                target = パラグラフの取得(paraCurr, movecount);
            }
            return target;
        }
        // 指定された方向において、テキストやオブジェクトが記載されたパラグラフを取得
        // 2018/3/20
        public Paragraph 有効パラグラフ取得(Paragraph paraCurr, long initialcount)
        {
            long movecount = initialcount;
            Paragraph target = パラグラフの取得(paraCurr, movecount);
            while (target != null)
            {
                if (パラグラフが空白か判定(target) == false
                || パラグラフが数式を含むか判定(target) == true
                || パラグラフが表を含むか判定(target) == true
                || パラグラフが画像を含むか判定(target) == true)
                {
                    break;
                }
                movecount += initialcount;
                target = パラグラフの取得(paraCurr, movecount);
            }
            return target;
        }
        public long 項目までの行数取得(Paragraph paraCurr, long initialcount)
        {
            long 行数 = 0;
            long movecount = initialcount;
            Paragraph target = パラグラフの取得(paraCurr, movecount);
            while (target != null)
            {
                if (項目の判定(target) == true)
                {
                    break;
                }
                long バイト数 = LenB(target.Range.Text);
                行数 += ((バイト数 + 77) / 80);
                movecount += initialcount;
                target = パラグラフの取得(paraCurr, movecount);
            }
            return 行数;
        }
        public int LenB(string stTarget)
        {
            return System.Text.Encoding.GetEncoding("Shift_JIS").GetByteCount(stTarget);
        }
        // paraCurrを基準として､movecountで指定された位置の段落を返す｡
        public Paragraph パラグラフの取得(Paragraph paraCurr, long movecount)
        {
            Range rng = paraCurr.Range;
            if (movecount == 0)
            {
                return paraCurr;
            }
            rng.StartOf(WdUnits.wdParagraph, WdMovementType.wdMove);
            if (rng.Move(WdUnits.wdParagraph, movecount) != 0)
            {
                return rng.Paragraphs[1];
            }
            return null;
        }
        /// <summary>
        ///
        /// </summary>
        /// 


        public void 領域拡張(ref Range rng)
        {
            rng.StartOf(WdUnits.wdParagraph, WdMovementType.wdExtend);
            rng.EndOf(WdUnits.wdParagraph, WdMovementType.wdExtend);
        }
        public void 垂直タブを改行に(Document doc)
        {
            Range rng = doc.Range(0, 0);
            rng.Find.MatchWildcards = false;
            while (rng.Find.Execute("\x0b"))
            {
                rng.Find.Text = "\r";
                rng.SetRange(rng.End, rng.End);
            }
        }

        // 開始位置をトリミング
        // matchesで指定された文字以外が出現するまでトリミング
        public void 領域始端のトリム(ref Range rng)
        {
            string c;
            int max, i;

            max = rng.Characters.Count;
            for (i = 1; i <= max; i++)
            {
                c = rng.Characters[i].Text;
                string spc = "　 \r\n";
                if (spc.IndexOf(c) >= 0)
                {
                    rng.MoveStartUntil(c);
                    break;
                }
            }
        }

        //  終了位置をトリミング
        //  matchesで指定された文字以外が出現するまでトリミング
        public void 領域終端のトリム(ref Range rng)
        {
            string c;
            int max;
            int i;

            max = rng.Characters.Count;
            for (i = max; i >= 1; i--)
            {
                c = rng.Characters[i].Text;
                string spc = "　 \r\n";
                if (spc.IndexOf(c) >= 0)
                {
                    rng.MoveEndUntil(c, WdConstants.wdBackward);
                    break;
                }
            }
        }

        public Microsoft.Office.Interop.Word.Bookmark 次の書類名にブックマーク2(
            Document doc,
            long spos,
            long epos,
            string ブックマーク名)
        {
            Microsoft.Office.Interop.Word.Bookmark bm;
            Range rng;
            bm = null;

            rng = doc.Range(spos, spos);
            rng.Find.Forward = true;
            rng.Find.MatchWildcards = true;

            if (rng.Find.Execute("【書類名】"))
            {
                if (rng.End < epos)
                {
                    bm = rng.Bookmarks.Add(ブックマーク名);
                }
            }
            return bm;
        }
        public bool 書類名の範囲を選択(
            Document doc,
            ref Range rng,
            long spos = 0,
            long epos = -1
            )
        {
            Microsoft.Office.Interop.Word.Bookmark bmS;
            bmS = 次の書類名にブックマーク2(doc, spos, epos, 最初のブックマーク);
            if (bmS == null)
            {
                return false;
            }
            Microsoft.Office.Interop.Word.Bookmark bmE;
            bmE = 次の書類名にブックマーク2(doc, bmS.Range.End, epos, 次のブックマーク);

            long endpos;
            rng = doc.Range(0, 0);
            if (bmE == null)
            {
                endpos = doc.Content.End;
            }
            else
            {
                endpos = bmE.Range.Start;
            }

            rng = doc.Range(bmS.Range.Start, endpos);
            bmS.Delete();
            if (bmE != null)
            {
                bmE.Delete();
            }
            return true;
        }
        public Range 書類名の範囲(
            string docname,
            Document doc,
            long spos = 0,
            long epos = -1
            )
        {
            Range rng;
            if (epos == -1)
            {
                epos = doc.Content.End;
            }
            rng = null;
            while (書類名の範囲を選択(doc, ref rng, spos, epos) == true)
            {
                if (rng == null)
                {
                    break;
                }
                string para1 = rng.Paragraphs[1].Range.Text;

                if (para1.IndexOf(docname) > 0)
                {
                    領域始端のトリム(ref rng);
                    領域終端のトリム(ref rng);
                    領域拡張(ref rng);
                    break;
                }
                spos = rng.End;
            }
            return rng;
        }
    }
}
