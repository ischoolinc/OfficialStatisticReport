
namespace ClassAndStudentInfo
{
    partial class PrintBranch
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.txtTitle = new System.Windows.Forms.TextBox();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.lstCourseKind = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.chkUnGraduate = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.SuspendLayout();
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(73, 338);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(120, 45);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 2;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.Location = new System.Drawing.Point(212, 338);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(120, 45);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // txtTitle
            // 
            this.txtTitle.Location = new System.Drawing.Point(185, 285);
            this.txtTitle.Name = "txtTitle";
            this.txtTitle.Size = new System.Drawing.Size(181, 27);
            this.txtTitle.TabIndex = 18;
            this.txtTitle.TextChanged += new System.EventHandler(this.txtTitle_TextChanged);
            // 
            // labelX2
            // 
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(12, 285);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(179, 37);
            this.labelX2.TabIndex = 17;
            this.labelX2.Text = "請輸入欲列印報表標題";
            // 
            // lstCourseKind
            // 
            // 
            // 
            // 
            this.lstCourseKind.Border.Class = "ListViewBorder";
            this.lstCourseKind.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lstCourseKind.CheckBoxes = true;
            this.lstCourseKind.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1,
            this.columnHeader2});
            this.lstCourseKind.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lstCourseKind.HideSelection = false;
            this.lstCourseKind.Location = new System.Drawing.Point(12, 55);
            this.lstCourseKind.Name = "lstCourseKind";
            this.lstCourseKind.Size = new System.Drawing.Size(354, 177);
            this.lstCourseKind.TabIndex = 16;
            this.lstCourseKind.UseCompatibleStateImageBehavior = false;
            this.lstCourseKind.View = System.Windows.Forms.View.Details;
            this.lstCourseKind.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.lstCourseKind_ItemChecked);
            this.lstCourseKind.SelectedIndexChanged += new System.EventHandler(this.lstCourseKind_SelectedIndexChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "課程類型";
            this.columnHeader1.Width = 300;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "ID";
            this.columnHeader2.Width = 0;
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(301, 37);
            this.labelX1.TabIndex = 15;
            this.labelX1.Text = "請選擇欲列印課程類型";
            // 
            // chkUnGraduate
            // 
            this.chkUnGraduate.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkUnGraduate.BackgroundStyle.Class = "";
            this.chkUnGraduate.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkUnGraduate.Location = new System.Drawing.Point(13, 242);
            this.chkUnGraduate.Name = "chkUnGraduate";
            this.chkUnGraduate.Size = new System.Drawing.Size(354, 37);
            this.chkUnGraduate.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.chkUnGraduate.TabIndex = 24;
            this.chkUnGraduate.Text = "取得修業資格統計包含364輔導延修";
            // 
            // PrintBranch
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(379, 393);
            this.Controls.Add(this.chkUnGraduate);
            this.Controls.Add(this.txtTitle);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.lstCourseKind);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnPrint);
            this.DoubleBuffered = true;
            this.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Name = "PrintBranch";
            this.Text = "表-7 高中職學校班級及學生概況（一）";
            this.Load += new System.EventHandler(this.PrintBranch_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private System.Windows.Forms.TextBox txtTitle;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ListViewEx lstCourseKind;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkUnGraduate;
    }
}