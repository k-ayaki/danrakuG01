namespace danrakuG01
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Ribbon1));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.danrakuG = this.Factory.CreateRibbonGroup();
            this.addDanraku = this.Factory.CreateRibbonButton();
            this.renumDanraku = this.Factory.CreateRibbonButton();
            this.delDanraku = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.danrakuG.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.danrakuG);
            this.tab1.Label = "AppLint";
            this.tab1.Name = "tab1";
            // 
            // danrakuG
            // 
            this.danrakuG.Items.Add(this.addDanraku);
            this.danrakuG.Items.Add(this.renumDanraku);
            this.danrakuG.Items.Add(this.delDanraku);
            this.danrakuG.Label = "段落生成";
            this.danrakuG.Name = "danrakuG";
            // 
            // addDanraku
            // 
            this.addDanraku.Image = ((System.Drawing.Image)(resources.GetObject("addDanraku.Image")));
            this.addDanraku.Label = "段落付与";
            this.addDanraku.Name = "addDanraku";
            this.addDanraku.ScreenTip = "ver.1.0.0.8:段落番号の新規付与";
            this.addDanraku.ShowImage = true;
            this.addDanraku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddDanraku_Click);
            // 
            // renumDanraku
            // 
            this.renumDanraku.Image = ((System.Drawing.Image)(resources.GetObject("renumDanraku.Image")));
            this.renumDanraku.Label = "番号振直";
            this.renumDanraku.Name = "renumDanraku";
            this.renumDanraku.ScreenTip = "ver.1.0.0.8:段落番号の振り直し";
            this.renumDanraku.ShowImage = true;
            this.renumDanraku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.RenumDanraku_Click);
            // 
            // delDanraku
            // 
            this.delDanraku.Image = ((System.Drawing.Image)(resources.GetObject("delDanraku.Image")));
            this.delDanraku.Label = "段落削除";
            this.delDanraku.Name = "delDanraku";
            this.delDanraku.ScreenTip = "ver.1.0.0.8:段落番号の削除";
            this.delDanraku.ShowImage = true;
            this.delDanraku.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.DelDanraku_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.danrakuG.ResumeLayout(false);
            this.danrakuG.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup danrakuG;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton addDanraku;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton renumDanraku;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton delDanraku;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
