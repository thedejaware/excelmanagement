namespace ExcelProcess
{
    partial class Form1
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
            this.grdExcel = new System.Windows.Forms.DataGridView();
            this.btnResults = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.grdExcel)).BeginInit();
            this.SuspendLayout();
            // 
            // grdExcel
            // 
            this.grdExcel.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.grdExcel.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.grdExcel.Location = new System.Drawing.Point(0, 76);
            this.grdExcel.Name = "grdExcel";
            this.grdExcel.Size = new System.Drawing.Size(945, 305);
            this.grdExcel.TabIndex = 0;
            // 
            // btnResults
            // 
            this.btnResults.Location = new System.Drawing.Point(220, 30);
            this.btnResults.Name = "btnResults";
            this.btnResults.Size = new System.Drawing.Size(552, 23);
            this.btnResults.TabIndex = 1;
            this.btnResults.Text = "Sonuçları Getir";
            this.btnResults.UseVisualStyleBackColor = true;
            this.btnResults.Click += new System.EventHandler(this.btnResults_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(945, 381);
            this.Controls.Add(this.btnResults);
            this.Controls.Add(this.grdExcel);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.grdExcel)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView grdExcel;
        private System.Windows.Forms.Button btnResults;
    }
}

