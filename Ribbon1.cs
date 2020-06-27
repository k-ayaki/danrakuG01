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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private void AddDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;

            var myPatDoc = new patDoc();
            //myPatDoc.G_段落番号付与(doc);
            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            myPatDoc.垂直タブを改行に(doc);
            //ProgressDialogオブジェクトを作成する
            ProgressDialog pd = new ProgressDialog("段落番号の付与",
                new DoWorkEventHandler(ProgressDialog_Add_DoWork),
                16);
            //進行状況ダイアログを表示する
            DialogResult result = pd.ShowDialog();
            //結果を取得する
            if (result == DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルされました");
                //後始末
                pd.Dispose();
                return;
            }
            else if (result == DialogResult.Abort)
            {
                //エラー情報を取得する
                Exception ex = pd.Error;
                MessageBox.Show("エラー: " + ex.Message);
                //後始末
                pd.Dispose();
                return;
            }
            else if (result == DialogResult.OK)
            {
                //結果を取得する
                int stopTime = (int)pd.Result;
                //MessageBox.Show("成功しました: " + stopTime.ToString());
            }
            //後始末
            pd.Dispose();

            //ProgressDialogオブジェクトを作成する
            pd = new ProgressDialog("段落番号の振り直し",
                new DoWorkEventHandler(ProgressDialog_Renum_DoWork),
                16);
            //進行状況ダイアログを表示する
            result = pd.ShowDialog();
            //結果を取得する
            if (result == DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルされました");
            }
            else if (result == DialogResult.Abort)
            {
                //エラー情報を取得する
                Exception ex = pd.Error;
                MessageBox.Show("エラー: " + ex.Message);
            }
            else if (result == DialogResult.OK)
            {
                //結果を取得する
                int stopTime = (int)pd.Result;
                //MessageBox.Show("成功しました: " + stopTime.ToString());
            }
            //後始末
            pd.Dispose();
        }
        //DoWorkイベントハンドラ
        // 段落の追加
        private void ProgressDialog_Add_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;

            //パラメータを取得する
            int stopTime = (int)e.Argument;

            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            var myPatDoc = new patDoc();
            Range rng = myPatDoc.書類名の範囲("明細書", doc);
            if (rng == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                e.Result = 0;
                return;
            }
            int counter = 0;
            int i = 0;
            int lastTick = Environment.TickCount;
            int currTick;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");

            foreach (Paragraph paraCurr in rng.Paragraphs)
            {
                string tmpdebug = paraCurr.Range.Text;
                if (myPatDoc.項目の判定(paraCurr))
                {
                    if (myPatDoc.パラグラフが段落番号付与対象項目か判定(paraCurr))
                    {
                        Range rng2 = myPatDoc.直後への段落番号の挿入(paraCurr);
                        if (paraCurr.Range.Text.IndexOf("【符号の説明】") >= 0)
                        {
                            break;
                        }
                    }
                    else if (myPatDoc.パラグラフが数化表項目か判定(paraCurr))
                    {
                        Paragraph paraPrev = myPatDoc.テキスト記載パラグラフ取得(paraCurr, -1);
                        if (myPatDoc.パラグラフが数化表項目か判定(paraPrev))
                        {
                            myPatDoc.直前への段落番号挿入(paraCurr);
                        }
                        paraPrev = null;
                    }
                    else if (myPatDoc.パラグラフが段落番号付与非対象項目か判定(paraCurr))
                    {
                        if (myPatDoc.パラグラフが不正な段落番号か判定(paraCurr))
                        {
                            paraCurr.Range.Delete();
                        }
                    }
                    else
                    {
                        Range rng2 = myPatDoc.直後への段落番号の挿入(paraCurr);
                    }
                }
                else
                {
                    myPatDoc.段落番号付与判定(paraCurr);
                }
                counter++;
                //キャンセルされたか調べる
                if (bw.CancellationPending)
                {
                    //キャンセルされたとき
                    e.Cancel = true;
                    return;
                }
                //指定された時間待機する
                //System.Threading.Thread.Sleep(stopTime);

                currTick = Environment.TickCount;
                if (currTick - lastTick > 1000)
                {
                    //ProgressChangedイベントハンドラを呼び出し、
                    //コントロールの表示を変更する
                    i = (int)(paraCurr.Range.End * 100 / rng.End);
                    bw.ReportProgress(i, i.ToString() + "% 終了しました");
                    lastTick = currTick;
                }
            }
            i = 100;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");
            System.Threading.Thread.Sleep(500);
            //結果を設定する
            e.Result = counter;
        }

        private void RenumDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            var myPatDoc = new patDoc();

            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            myPatDoc.垂直タブを改行に(doc);
            //ProgressDialogオブジェクトを作成する
            ProgressDialog pd = new ProgressDialog("段落番号の振り直し",
                new DoWorkEventHandler(ProgressDialog_Renum_DoWork),
                16);
            //進行状況ダイアログを表示する
            DialogResult result = pd.ShowDialog();
            //結果を取得する
            if (result == DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルされました");
            }
            else if (result == DialogResult.Abort)
            {
                //エラー情報を取得する
                Exception ex = pd.Error;
                MessageBox.Show("エラー: " + ex.Message);
            }
            else if (result == DialogResult.OK)
            {
                //結果を取得する
                int stopTime = (int)pd.Result;
                //MessageBox.Show("成功しました: " + stopTime.ToString());
            }

            //後始末
            pd.Dispose();
        }
        //DoWorkイベントハンドラ
        // 段落振り直し
        private void ProgressDialog_Renum_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;

            //パラメータを取得する
            int stopTime = (int)e.Argument;

            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            var myPatDoc = new patDoc();
            Range rng = myPatDoc.書類名の範囲("明細書", doc);
            if (rng == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                e.Result = 0;
                return;
            }
            long endpos = rng.End;
            int counter = 0;
            int i = 0;
            int lastTick = Environment.TickCount;
            int currTick;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");

            rng.Find.MatchWildcards = true;
            while (rng.Find.Execute("【[０-９]@】"))
            {
                counter++;
                rng.Text = myPatDoc.段落番号文字列の生成(counter);
                rng.SetRange(rng.End, rng.End);

                //キャンセルされたか調べる
                if (bw.CancellationPending)
                {
                    //キャンセルされたとき
                    e.Cancel = true;
                    return;
                }
                //指定された時間待機する
                //System.Threading.Thread.Sleep(16);

                currTick = Environment.TickCount;
                if (currTick - lastTick > 1000)
                {
                    //ProgressChangedイベントハンドラを呼び出し、
                    //コントロールの表示を変更する
                    i = (int)(rng.End * 100 / endpos);
                    bw.ReportProgress(i, i.ToString() + "% 終了しました");
                    lastTick = currTick;
                }
            }
            i = 100;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");
            System.Threading.Thread.Sleep(500);
            //結果を設定する
            e.Result = counter;
        }

        private void DelDanraku_Click(object sender, RibbonControlEventArgs e)
        {
            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            var myPatDoc = new patDoc();

            if (doc.TrackRevisions == true)
            {
                System.Windows.Forms.MessageBox.Show("変更履歴の記録をオフしてくたさい");
                return;
            }
            myPatDoc.垂直タブを改行に(doc);
            ProgressDialog pd = new ProgressDialog("段落の削除",
                new DoWorkEventHandler(ProgressDialog_Del_DoWork),
                16);
            //進行状況ダイアログを表示する
            DialogResult result = pd.ShowDialog();
            //結果を取得する
            if (result == DialogResult.Cancel)
            {
                MessageBox.Show("キャンセルされました");
            }
            else if (result == DialogResult.Abort)
            {
                //エラー情報を取得する
                Exception ex = pd.Error;
                MessageBox.Show("エラー: " + ex.Message);
            }
            else if (result == DialogResult.OK)
            {
                //結果を取得する
                int stopTime = (int)pd.Result;
                //MessageBox.Show("成功しました: " + stopTime.ToString());
            }
            //後始末
            pd.Dispose();
        }
        
        // 段落の削除
        private void ProgressDialog_Del_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;

            //パラメータを取得する
            int stopTime = (int)e.Argument;

            Document doc = danrakuG01.Globals.ThisAddIn.Application.ActiveDocument;
            var myPatDoc = new patDoc();
            Range rng = myPatDoc.書類名の範囲("明細書", doc);
            if (rng == null)
            {
                System.Windows.Forms.MessageBox.Show("明細書が記載されていません。", "警告");
                e.Result = 0;
                return;
            }
            long endpos = rng.End;
            int counter = 0;
            int i = 0;
            rng.Find.MatchWildcards = true;

            int lastTick = Environment.TickCount;
            int currTick;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");

            while (rng.Find.Execute("【[０-９]@】"))
            {
                counter++;

                rng.Text = "";
                rng.SetRange(rng.End, rng.End);
                if (myPatDoc.パラグラフが空白か判定(rng.Paragraphs[1]))
                {
                    rng.Paragraphs[1].Range.Delete();
                }
                //キャンセルされたか調べる
                if (bw.CancellationPending)
                {
                    //キャンセルされたとき
                    e.Cancel = true;
                    return;
                }
                currTick = Environment.TickCount;
                if( currTick - lastTick > 1000 )
                {
                    //指定された時間待機する
                    //System.Threading.Thread.Sleep(stopTime);

                    //ProgressChangedイベントハンドラを呼び出し、
                    //コントロールの表示を変更する
                    i = (int)(rng.End * 100 / endpos);
                    bw.ReportProgress(i, i.ToString() + "% 終了しました");
                    lastTick = currTick;
                }
            }
            i = 100;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");
            System.Threading.Thread.Sleep(500);
            //結果を設定する
            e.Result = counter;
        }
    }
}
