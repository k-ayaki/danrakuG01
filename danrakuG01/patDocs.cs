using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using System.Text.RegularExpressions;

namespace danrakuG01
{
    class patDoc : IDisposable
    {
        private const string 最初のブックマーク = "BK_738c8dc9_af9f_4e29_adde_934af51b07e2";
        private const string 次のブックマーク = "BK_97414a23_4a86_4371_a3b7_c00da9f7751a";

        private long e_counter = 1;
        private bool disposedValue;
        public Document m_doc { get; set; }
        public Range m_rng { get; set; }

        private const long PARA_FORWARD = 1;
        private const long PARA_BACKWARD = -1;
        public patDoc(Document aDoc)
        {
            m_doc = aDoc;
            m_rng = 書類名の範囲("明細書");
        }
        public void 段落挿入処理(Paragraph paraCurr)
        {
            if (パラグラフが数式を含む(paraCurr)
            || パラグラフが表を含む(paraCurr)
            || パラグラフが画像を含む(paraCurr))
            {
                ;
            }
            else
            if (項目の判定(paraCurr))
            {
                項目直後段落番号付与判定(paraCurr);
            }
            else
            {
                文章中段落番号付与判定(paraCurr);
            }

        }
        public void 項目直後段落番号付与判定(Paragraph paraCurr)
        {
            if (パラグラフが段落番号付与対象項目(paraCurr))
            {
                直後段落番号挿入(paraCurr);
            }
            else if (パラグラフが段落番号付与非対象項目(paraCurr))
            {
                ;
            }
            else if (パラグラフが数化表項目(paraCurr))
            {
                Paragraph paraNext1 = パラグラフの取得(paraCurr, PARA_FORWARD);
                if (パラグラフが数式を含む(paraNext1) == false
                && パラグラフが表を含む(paraNext1) == false
                && パラグラフが画像を含む(paraNext1) == false)
                {
                    Paragraph paraNext2 = パラグラフの取得(paraNext1, PARA_FORWARD);
                    if (パラグラフが数式を含む(paraNext2) == false
                    && パラグラフが表を含む(paraNext2) == false
                    && パラグラフが画像を含む(paraNext2) == false)
                    {
                        if (数式判定(paraNext1))
                        {
                            直後段落番号挿入(paraNext1);
                        }
                        else
                        {
                            直後段落番号挿入(paraCurr);
                        }
                    }
                }
            }
            else // 未定義の項目
            {
                直後段落番号挿入(paraCurr);
            }

        }
        public void 文章中段落番号付与判定(Paragraph paraCurr)
        {
            Paragraph paraPrev1 = パラグラフの取得(paraCurr, PARA_BACKWARD);
            if (パラグラフが数式を含む(paraPrev1) == true
            || パラグラフが表を含む(paraPrev1) == true
            || パラグラフが画像を含む(paraPrev1) == true)
            {
                直前段落番号挿入(paraCurr);
            }
            else
            {
                long 項目までの行数 = 項目までの行数取得(paraCurr, PARA_BACKWARD);
                Paragraph paraPrev = テキスト記載パラグラフ取得(paraCurr, PARA_BACKWARD);

                if ((見出し判定(paraCurr) == true || 図説明の判定(paraCurr) == true)
                && 3 <= 項目までの行数)
                {
                    直前段落番号挿入(paraCurr);
                }
                else if (文末の句点判定(paraPrev) == true && 4 <= 項目までの行数)
                {
                    直前段落番号挿入(paraCurr);
                }
                else if (5 <= 項目までの行数)
                {
                    直前段落番号挿入(paraCurr);
                }
            }
        }
        private void 直前段落番号挿入(Paragraph para)
        {
            string strIns = "　" + 段落番号文字列(e_counter);
            para.Range.Collapse(WdCollapseDirection.wdCollapseStart);
            para.Range.InsertBefore(strIns + "\r");
            e_counter++;
        }
        private void 直後段落番号挿入(Paragraph para)
        {
            string strIns = "　" + 段落番号文字列(e_counter);
            para.Range.Collapse(WdCollapseDirection.wdCollapseStart);
            para.Range.InsertBefore(para.Range.Text);
            para.Range.InsertBefore(strIns + "\r");
            para.Range.Collapse(WdCollapseDirection.wdCollapseEnd);
            para.Range.Delete();
            e_counter++;
        }
        public bool 図説明の判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text, @"^[　]*?図[１-９－Ａ-Ｚａ-ｚ]+"))
            {
                return true;
            }
            if (Regex.IsMatch(para.Range.Text, @"図[１-９－Ａ-Ｚａ-ｚ]+(（[Ａ-Ｚａ-ｚ]）|)[はがをで]"))
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
            if (Regex.IsMatch(para.Range.Text.TrimEnd(), @"。$"))
            {
                return true;
            }
            return false;
        }
        public bool 数式判定(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text.TrimEnd(), @"[（(][０-９0-9]+[）)]$"))
            {
                return true;
            }
            if (Regex.IsMatch(para.Range.Text.TrimEnd(), @"[=＝≒≠≡＜<≦＞>≧]"))
            {
                return true;
            }
            return false;
        }
        public string 段落番号文字列(long counter)
        {
            string 段落番号文字列;

            段落番号文字列 = "【" + Strings.StrConv(counter.ToString("0000"), VbStrConv.Wide, 0) + "】";
            return 段落番号文字列;
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
        public bool パラグラフが空白(Paragraph para)
        {
            string ckStr;
            char[] charsToTrim = { '　', ' ', '\r', '\n', '\t', '\x0b', '\x0c', '\x0f' };

            ckStr = para.Range.Text;
            ckStr = ckStr.Trim(charsToTrim);
            if (ckStr.Length == 0)
            {
                return true;
            }
            return false;
        }
        public bool パラグラフが数式を含む(Paragraph para)
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
        public bool パラグラフが表を含む(Paragraph para)
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
        public bool パラグラフが画像を含む(Paragraph para)
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
        public bool パラグラフが数化表項目(Paragraph para)
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
        public bool パラグラフが段落番号付与対象項目(Paragraph para)
        {
            if (Regex.IsMatch(para.Range.Text,
                @"(【技術分野】" +
                 "|【背景技術】" +
                 "|【従来の技術】" +
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
        public bool パラグラフが段落番号付与非対象項目(Paragraph para)
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
                 "|【請求の範囲】" +
                 "|【特許請求の範囲】" +
                 "|【実用新案登録請求の範囲】" +
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
        //指定された方向において、テキストが記載されたパラグラフを取得
        public Paragraph テキスト記載パラグラフ取得(Paragraph paraCurr, long direction)
        {
            long movecount = direction;
            Paragraph target = パラグラフの取得(paraCurr, movecount);
            while (target != null)
            {
                if (パラグラフが空白(target) == false
                && パラグラフが数式を含む(target) == false
                && パラグラフが表を含む(target) == false
                && パラグラフが画像を含む(target) == false)
                {
                    break;
                }
                movecount += direction;
                target = パラグラフの取得(paraCurr, movecount);
            }
            return target;
        }
        public long 項目までの行数取得(Paragraph paraCurr, long direction)
        {
            long 行数 = 0;
            Paragraph target = パラグラフの取得(paraCurr, direction);
            while (target != null)
            {
                if (項目の判定(target) == true)
                {
                    break;
                }
                long バイト数 = LenB(target.Range.Text);
                行数 += ((バイト数 + 77) / 80);
                target = パラグラフの取得(target, direction);
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

        public void 領域拡張(ref Range rng)
        {
            rng.StartOf(WdUnits.wdParagraph, WdMovementType.wdExtend);
            rng.EndOf(WdUnits.wdParagraph, WdMovementType.wdExtend);
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

        public Microsoft.Office.Interop.Word.Bookmark 次の書類名にブックマーク(
            long spos,
            long epos,
            string ブックマーク名)
        {
            Microsoft.Office.Interop.Word.Bookmark bm;
            Range rng;
            bm = null;

            rng = m_doc.Range(spos, spos);
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
            ref Range rng,
            long spos = 0,
            long epos = -1
            )
        {
            Microsoft.Office.Interop.Word.Bookmark bmS;
            bmS = 次の書類名にブックマーク(spos, epos, 最初のブックマーク);
            if (bmS == null)
            {
                return false;
            }
            Microsoft.Office.Interop.Word.Bookmark bmE;
            bmE = 次の書類名にブックマーク(bmS.Range.End, epos, 次のブックマーク);

            long endpos;
            rng = m_doc.Range(0, 0);
            if (bmE == null)
            {
                endpos = m_doc.Content.End;
            }
            else
            {
                endpos = bmE.Range.Start;
            }

            rng = m_doc.Range(bmS.Range.Start, endpos);
            bmS.Delete();
            if (bmE != null)
            {
                bmE.Delete();
            }
            return true;
        }
        public Range 書類名の範囲(
            string docname,
            long spos = 0,
            long epos = -1
            )
        {
            Range rng;
            if (epos == -1)
            {
                epos = m_doc.Content.End;
            }
            rng = null;
            while (書類名の範囲を選択(ref rng, spos, epos) == true)
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

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: マネージド状態を破棄します (マネージド オブジェクト)
                }

                // TODO: アンマネージド リソース (アンマネージド オブジェクト) を解放し、ファイナライザーをオーバーライドします
                // TODO: 大きなフィールドを null に設定します
                disposedValue = true;
            }
        }

        // // TODO: 'Dispose(bool disposing)' にアンマネージド リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします
        // ~patDoc()
        // {
        //     // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
