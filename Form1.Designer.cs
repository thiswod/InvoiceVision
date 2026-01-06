namespace InvoiceVision
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnSelectImages = new Button();
            btnStart = new Button();
            btnExport = new Button();
            listBoxImages = new ListBox();
            columnHeader1 = new ColumnHeader();
            columnHeader2 = new ColumnHeader();
            columnHeader3 = new ColumnHeader();
            columnHeader4 = new ColumnHeader();
            columnHeader5 = new ColumnHeader();
            columnHeader6 = new ColumnHeader();
            columnHeader7 = new ColumnHeader();
            columnHeader8 = new ColumnHeader();
            columnHeader9 = new ColumnHeader();
            columnHeader10 = new ColumnHeader();
            columnHeader11 = new ColumnHeader();
            columnHeader12 = new ColumnHeader();
            columnHeader13 = new ColumnHeader();
            columnHeader14 = new ColumnHeader();
            columnHeader15 = new ColumnHeader();
            columnHeader16 = new ColumnHeader();
            columnHeader17 = new ColumnHeader();
            columnHeader18 = new ColumnHeader();
            columnHeader19 = new ColumnHeader();
            columnHeader20 = new ColumnHeader();
            superListView = new SuperListView();
            progressBar = new ProgressBar();
            labelStatus = new Label();
            SuspendLayout();
            // 
            // btnSelectImages
            // 
            btnSelectImages.Location = new Point(12, 12);
            btnSelectImages.Name = "btnSelectImages";
            btnSelectImages.Size = new Size(120, 35);
            btnSelectImages.TabIndex = 0;
            btnSelectImages.Text = "选择文件";
            btnSelectImages.UseVisualStyleBackColor = true;
            btnSelectImages.Click += BtnSelectImages_Click;
            // 
            // btnStart
            // 
            btnStart.Enabled = false;
            btnStart.Location = new Point(138, 12);
            btnStart.Name = "btnStart";
            btnStart.Size = new Size(120, 35);
            btnStart.TabIndex = 1;
            btnStart.Text = "开始识别";
            btnStart.UseVisualStyleBackColor = true;
            btnStart.Click += BtnStart_Click;
            // 
            // btnExport
            // 
            btnExport.Enabled = false;
            btnExport.Location = new Point(264, 12);
            btnExport.Name = "btnExport";
            btnExport.Size = new Size(120, 35);
            btnExport.TabIndex = 2;
            btnExport.Text = "导出Excel";
            btnExport.UseVisualStyleBackColor = true;
            btnExport.Click += BtnExport_Click;
            // 
            // listBoxImages
            // 
            listBoxImages.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left;
            listBoxImages.FormattingEnabled = true;
            listBoxImages.ItemHeight = 17;
            listBoxImages.Location = new Point(12, 53);
            listBoxImages.Name = "listBoxImages";
            listBoxImages.Size = new Size(372, 361);
            listBoxImages.TabIndex = 3;
            // 
            // columnHeader1
            // 
            columnHeader1.Text = "发票号码";
            columnHeader1.Width = 150;
            // 
            // columnHeader2
            // 
            columnHeader2.Text = "发票代码";
            columnHeader2.Width = 120;
            // 
            // columnHeader3
            // 
            columnHeader3.Text = "开票日期";
            columnHeader3.Width = 120;
            // 
            // columnHeader4
            // 
            columnHeader4.Text = "购买方名称";
            columnHeader4.Width = 200;
            // 
            // columnHeader5
            // 
            columnHeader5.Text = "销售方名称";
            columnHeader5.Width = 200;
            // 
            // columnHeader6
            // 
            columnHeader6.Text = "金额合计";
            columnHeader6.Width = 100;
            // 
            // columnHeader7
            // 
            columnHeader7.Text = "税额";
            columnHeader7.Width = 100;
            // 
            // columnHeader8
            // 
            columnHeader8.Text = "价税合计";
            columnHeader8.Width = 100;
            // 
            // columnHeader9
            // 
            columnHeader9.Text = "文件路径";
            columnHeader9.Width = 350;
            // 
            // columnHeader10
            // 
            columnHeader10.Text = "购买方税号";
            columnHeader10.Width = 200;
            // 
            // columnHeader11
            // 
            columnHeader11.Text = "销售方税号";
            columnHeader11.Width = 200;
            // 
            // columnHeader12
            // 
            columnHeader12.Text = "商品名称";
            columnHeader12.Width = 200;
            // 
            // columnHeader13
            // 
            columnHeader13.Text = "规格型号";
            columnHeader13.Width = 150;
            // 
            // columnHeader14
            // 
            columnHeader14.Text = "单位";
            columnHeader14.Width = 80;
            // 
            // columnHeader15
            // 
            columnHeader15.Text = "数量";
            columnHeader15.Width = 80;
            // 
            // columnHeader16
            // 
            columnHeader16.Text = "单价";
            columnHeader16.Width = 100;
            // 
            // columnHeader17
            // 
            columnHeader17.Text = "金额";
            columnHeader17.Width = 100;
            // 
            // columnHeader18
            // 
            columnHeader18.Text = "税率";
            columnHeader18.Width = 80;
            // 
            // columnHeader19
            // 
            columnHeader19.Text = "税额";
            columnHeader19.Width = 100;
            // 
            // columnHeader20
            // 
            columnHeader20.Text = "文件路径";
            columnHeader20.Width = 350;
            // 
            // superListView
            // 
            superListView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            superListView.Columns.AddRange(new ColumnHeader[] { columnHeader1, columnHeader2, columnHeader3, columnHeader4, columnHeader5, columnHeader10, columnHeader11, columnHeader6, columnHeader7, columnHeader8, columnHeader12, columnHeader13, columnHeader14, columnHeader15, columnHeader16, columnHeader17, columnHeader18, columnHeader19, columnHeader20 });
            superListView.FullRowSelect = true;
            superListView.GridLines = true;
            superListView.Location = new Point(390, 53);
            superListView.Name = "superListView";
            superListView.OwnerDraw = true;
            superListView.Size = new Size(1022, 361);
            superListView.TabIndex = 4;
            superListView.UseCompatibleStateImageBehavior = false;
            superListView.View = View.Details;
            // 
            // progressBar
            // 
            progressBar.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new Point(12, 420);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(1400, 23);
            progressBar.TabIndex = 5;
            progressBar.Visible = false;
            // 
            // labelStatus
            // 
            labelStatus.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            labelStatus.AutoSize = true;
            labelStatus.Location = new Point(12, 450);
            labelStatus.Name = "labelStatus";
            labelStatus.Size = new Size(32, 17);
            labelStatus.TabIndex = 6;
            labelStatus.Text = "就绪";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 17F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1424, 476);
            Controls.Add(labelStatus);
            Controls.Add(progressBar);
            Controls.Add(superListView);
            Controls.Add(listBoxImages);
            Controls.Add(btnExport);
            Controls.Add(btnStart);
            Controls.Add(btnSelectImages);
            MinimumSize = new Size(800, 500);
            Name = "Form1";
            Text = "发票识别工具";
            Load += Form1_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnSelectImages;
        private Button btnStart;
        private Button btnExport;
        private ListBox listBoxImages;
        private SuperListView superListView;
        private ProgressBar progressBar;
        private Label labelStatus;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ColumnHeader columnHeader4;
        private System.Windows.Forms.ColumnHeader columnHeader5;
        private System.Windows.Forms.ColumnHeader columnHeader6;
        private System.Windows.Forms.ColumnHeader columnHeader7;
        private System.Windows.Forms.ColumnHeader columnHeader8;
        private System.Windows.Forms.ColumnHeader columnHeader9;
        private System.Windows.Forms.ColumnHeader columnHeader10;
        private System.Windows.Forms.ColumnHeader columnHeader11;
        private System.Windows.Forms.ColumnHeader columnHeader12;
        private System.Windows.Forms.ColumnHeader columnHeader13;
        private System.Windows.Forms.ColumnHeader columnHeader14;
        private System.Windows.Forms.ColumnHeader columnHeader15;
        private System.Windows.Forms.ColumnHeader columnHeader16;
        private System.Windows.Forms.ColumnHeader columnHeader17;
        private System.Windows.Forms.ColumnHeader columnHeader18;
        private System.Windows.Forms.ColumnHeader columnHeader19;
        private System.Windows.Forms.ColumnHeader columnHeader20;
    }
}
