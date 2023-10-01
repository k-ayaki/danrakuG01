using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;

namespace danrakuG01
{
    internal class RenumDanraku : IDisposable
    {
        private bool disposedValue;

        public Document doc;
        public patDoc myPatDoc;
        public RenumDanraku(Document aDoc)
        {
            doc = aDoc;
            myPatDoc = new patDoc(doc);

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
            myPatDoc.Dispose();
        }
        //DoWorkイベントハンドラ
        // 段落振り直し
        private void ProgressDialog_Renum_DoWork(object sender, DoWorkEventArgs e)
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
            long endpos = myPatDoc.m_rng.End;
            int counter = 0;
            int i = 0;
            int lastTick = Environment.TickCount;
            int currTick;
            bw.ReportProgress(i, i.ToString() + "% 終了しました");

            myPatDoc.m_rng.Find.MatchWildcards = true;
            while (myPatDoc.m_rng.Find.Execute("【[０-９]@】"))
            {
                counter++;
                myPatDoc.m_rng.Text = myPatDoc.段落番号文字列(counter);
                myPatDoc.m_rng.SetRange(myPatDoc.m_rng.End, myPatDoc.m_rng.End);

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
                    i = (int)(myPatDoc.m_rng.End * 100 / endpos);
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
        // ~RenumDanraku()
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
