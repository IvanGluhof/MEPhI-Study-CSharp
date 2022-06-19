namespace SpecDep
{
    partial class Vvod_kursa
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
            this.button_save = new System.Windows.Forms.Button();
            this.button_cancel = new System.Windows.Forms.Button();
            this.comboBox_vybor_val = new System.Windows.Forms.ComboBox();
            this.textBox_vvod_kursa = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // button_save
            // 
            this.button_save.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.button_save.Location = new System.Drawing.Point(194, 117);
            this.button_save.Name = "button_save";
            this.button_save.Size = new System.Drawing.Size(92, 36);
            this.button_save.TabIndex = 0;
            this.button_save.Text = "Принять";
            this.button_save.UseVisualStyleBackColor = false;
            this.button_save.Click += new System.EventHandler(this.button_save_Click);
            // 
            // button_cancel
            // 
            this.button_cancel.BackColor = System.Drawing.Color.DeepSkyBlue;
            this.button_cancel.Location = new System.Drawing.Point(41, 117);
            this.button_cancel.Name = "button_cancel";
            this.button_cancel.Size = new System.Drawing.Size(92, 36);
            this.button_cancel.TabIndex = 1;
            this.button_cancel.Text = "Отмена";
            this.button_cancel.UseVisualStyleBackColor = false;
            this.button_cancel.Click += new System.EventHandler(this.button_cancel_Click);
            // 
            // comboBox_vybor_val
            // 
            this.comboBox_vybor_val.FormattingEnabled = true;
            this.comboBox_vybor_val.Location = new System.Drawing.Point(12, 56);
            this.comboBox_vybor_val.Name = "comboBox_vybor_val";
            this.comboBox_vybor_val.Size = new System.Drawing.Size(151, 21);
            this.comboBox_vybor_val.TabIndex = 2;
            this.comboBox_vybor_val.SelectedIndexChanged += new System.EventHandler(this.comboBox_vybor_val_SelectedIndexChanged);
            // 
            // textBox_vvod_kursa
            // 
            this.textBox_vvod_kursa.Location = new System.Drawing.Point(169, 57);
            this.textBox_vvod_kursa.Name = "textBox_vvod_kursa";
            this.textBox_vvod_kursa.Size = new System.Drawing.Size(151, 20);
            this.textBox_vvod_kursa.TabIndex = 3;
            this.textBox_vvod_kursa.TextChanged += new System.EventHandler(this.textBox_vvod_kursa_TextChanged);
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.SkyBlue;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Location = new System.Drawing.Point(12, 37);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(151, 13);
            this.textBox2.TabIndex = 4;
            this.textBox2.Text = "Выбор валюты";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.Color.SkyBlue;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Location = new System.Drawing.Point(169, 37);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(151, 13);
            this.textBox3.TabIndex = 5;
            this.textBox3.Text = "Ввод курса";
            this.textBox3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // Vvod_kursa
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.SkyBlue;
            this.ClientSize = new System.Drawing.Size(335, 170);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox_vvod_kursa);
            this.Controls.Add(this.comboBox_vybor_val);
            this.Controls.Add(this.button_cancel);
            this.Controls.Add(this.button_save);
            this.Name = "Vvod_kursa";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Ввод курса валюты";
            this.Load += new System.EventHandler(this.Vvod_kursa_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_save;
        private System.Windows.Forms.Button button_cancel;
        private System.Windows.Forms.ComboBox comboBox_vybor_val;
        private System.Windows.Forms.TextBox textBox_vvod_kursa;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
    }
}