namespace ExcelHelper
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tb_Path = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btn_OpenFile = new System.Windows.Forms.Button();
            this.btn_Calc = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tb_Path
            // 
            this.tb_Path.BackColor = System.Drawing.Color.WhiteSmoke;
            this.tb_Path.Location = new System.Drawing.Point(12, 49);
            this.tb_Path.Multiline = true;
            this.tb_Path.Name = "tb_Path";
            this.tb_Path.ReadOnly = true;
            this.tb_Path.Size = new System.Drawing.Size(605, 45);
            this.tb_Path.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 24);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Path:";
            // 
            // btn_OpenFile
            // 
            this.btn_OpenFile.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btn_OpenFile.Location = new System.Drawing.Point(638, 49);
            this.btn_OpenFile.Name = "btn_OpenFile";
            this.btn_OpenFile.Size = new System.Drawing.Size(114, 45);
            this.btn_OpenFile.TabIndex = 2;
            this.btn_OpenFile.Text = "...";
            this.btn_OpenFile.UseVisualStyleBackColor = false;
            this.btn_OpenFile.Click += new System.EventHandler(this.btn_OpenFile_Click);
            // 
            // btn_Calc
            // 
            this.btn_Calc.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btn_Calc.Location = new System.Drawing.Point(12, 109);
            this.btn_Calc.Name = "btn_Calc";
            this.btn_Calc.Size = new System.Drawing.Size(197, 45);
            this.btn_Calc.TabIndex = 3;
            this.btn_Calc.Text = "Calculate";
            this.btn_Calc.UseVisualStyleBackColor = false;
            this.btn_Calc.Click += new System.EventHandler(this.btn_Calc_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1121, 569);
            this.Controls.Add(this.btn_Calc);
            this.Controls.Add(this.btn_OpenFile);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tb_Path);
            this.Name = "Form1";
            this.Text = "ExcelHelper";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_Path;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btn_OpenFile;
        private System.Windows.Forms.Button btn_Calc;
    }
}

