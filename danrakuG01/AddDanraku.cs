using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Tools.Ribbon;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace danrakuG01
{
    internal class AddDanraku : IDisposable
    {
        private bool disposedValue;

        public Document doc;
        public patDoc myPatDoc;

        public AddDanraku(Document aDoc)
        {
            doc = aDoc;
            myPatDoc = new patDoc(doc);

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
            myPatDoc.Dispose();
        }
        //DoWorkイベントハンドラ
        // 段落の追加
        private void ProgressDialog_Add_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;

            //パラメータを取得する
            int stopTime = (int)e.Argument;

            if (myPatDoc.m_rng == null)
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

            foreach (Paragraph paraCurr in myPatDoc.m_rng.Paragraphs)
            {
                string linebuf = paraCurr.Range.Text;
                myPatDoc.段落挿入処理(paraCurr);
                if (linebuf.IndexOf("【符号の説明】") >= 0)
                {
                    break;
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
                    i = (int)(paraCurr.Range.End * 100 / myPatDoc.m_rng.End);
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
        // ~AddDanraku()
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
