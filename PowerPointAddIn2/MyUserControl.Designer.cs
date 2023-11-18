namespace PowerPointAddIn2
{
    partial class MyUserControl
    {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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
            this.MarkRedButton = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // MarkRedButton
            // 
            this.MarkRedButton.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 13.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MarkRedButton.Location = new System.Drawing.Point(213, 626);
            this.MarkRedButton.Name = "MarkRedButton";
            this.MarkRedButton.Size = new System.Drawing.Size(200, 102);
            this.MarkRedButton.TabIndex = 0;
            this.MarkRedButton.Text = "マーク付けボタン";
            this.MarkRedButton.UseVisualStyleBackColor = true;
            this.MarkRedButton.Click += new System.EventHandler(this.MarkRedButton_Click);
            // 
            // listBox1
            // 
            this.listBox1.Font = new System.Drawing.Font("ＭＳ Ｐゴシック", 13.875F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBox1.FormattingEnabled = true;
            this.listBox1.ItemHeight = 37;
            this.listBox1.Location = new System.Drawing.Point(96, 156);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(434, 374);
            this.listBox1.TabIndex = 2;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // MyUserControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 25F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.MarkRedButton);
            this.Name = "MyUserControl";
            this.Size = new System.Drawing.Size(666, 1011);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button MarkRedButton;
        public System.Windows.Forms.ListBox listBox1;
    }
}
