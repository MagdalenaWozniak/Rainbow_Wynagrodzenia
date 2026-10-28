using System.Windows.Forms;

namespace ONE012
{
    partial class ONE012WIN
    {
        /// <summary>
        /// Wymagana zmienna projektanta.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Wyczyść wszystkie używane zasoby.
        /// </summary>
        /// <param name="disposing">prawda, jeżeli zarządzane zasoby powinny zostać zlikwidowane; Fałsz w przeciwnym wypadku.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Kod generowany przez Projektanta formularzy systemu Windows

        /// <summary>
        /// Metoda wymagana do obsługi projektanta — nie należy modyfikować
        /// jej zawartości w edytorze kodu.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ONE012WIN));
            this.title1 = new System.Windows.Forms.Label();
            this.tB1 = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.lb5 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // title1
            // 
            this.title1.AutoSize = true;
            this.title1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F);
            this.title1.Location = new System.Drawing.Point(12, 22);
            this.title1.Name = "title1";
            this.title1.Size = new System.Drawing.Size(370, 34);
            this.title1.TabIndex = 1;
            this.title1.Text = "Proszę wybrać arkusz zawierający listę dodatków płac do \r\nimportu.";
            // 
            // tB1
            // 
            this.tB1.ForeColor = System.Drawing.SystemColors.WindowFrame;
            this.tB1.Location = new System.Drawing.Point(108, 108);
            this.tB1.Name = "tB1";
            this.tB1.ReadOnly = true;
            this.tB1.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.tB1.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
            this.tB1.Size = new System.Drawing.Size(237, 20);
            this.tB1.TabIndex = 3;
            this.tB1.Text = "nie wybrano pliku";
            this.tB1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.TextBox1_MouseDoubleClick);
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.5F);
            this.button1.Location = new System.Drawing.Point(15, 148);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(395, 45);
            this.button1.TabIndex = 4;
            this.button1.Text = "Wykonaj";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.Button1_Click);
            // 
            // lb5
            // 
            this.lb5.AutoSize = true;
            this.lb5.ForeColor = System.Drawing.Color.Firebrick;
            this.lb5.Location = new System.Drawing.Point(15, 200);
            this.lb5.Name = "lb5";
            this.lb5.Size = new System.Drawing.Size(0, 13);
            this.lb5.TabIndex = 5;
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(351, 107);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(61, 22);
            this.button2.TabIndex = 0;
            this.button2.Text = "Wybierz";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.Button2_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F);
            this.label1.Location = new System.Drawing.Point(14, 108);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 17);
            this.label1.TabIndex = 2;
            this.label1.Text = "Plik arkusza:";
            // 
            // ONE012WIN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(426, 205);
            this.Controls.Add(this.title1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.lb5);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tB1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ONE012WIN";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ONE: Import listy dodatków płac";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label title1;
        private System.Windows.Forms.TextBox tB1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lb5;
        private Button button2;
        private Label label1;
    }
}

