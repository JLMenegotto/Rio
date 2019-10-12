namespace Rio
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
		protected override void Dispose ( bool disposing )
			{
			if (disposing && ( components != null ))
				{
				components.Dispose ( );
				}
			base.Dispose ( disposing );
			}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent ()
			{
            this.Processar = new System.Windows.Forms.Button();
            this.Arv_Temas = new System.Windows.Forms.TreeView();
            this.ProcessarTemas = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.ProcessarBairros = new System.Windows.Forms.Button();
            this.Arv_Bairros = new System.Windows.Forms.TreeView();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // Processar
            // 
            this.Processar.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Processar.Location = new System.Drawing.Point(310, 560);
            this.Processar.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Processar.Name = "Processar";
            this.Processar.Size = new System.Drawing.Size(115, 30);
            this.Processar.TabIndex = 16;
            this.Processar.Text = "Processar";
            this.Processar.UseVisualStyleBackColor = true;
            this.Processar.Click += new System.EventHandler(this.Processar_Click);
            // 
            // Arv_Temas
            // 
            this.Arv_Temas.BackColor = System.Drawing.SystemColors.Window;
            this.Arv_Temas.CheckBoxes = true;
            this.Arv_Temas.Font = new System.Drawing.Font("Arial Narrow", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Arv_Temas.Location = new System.Drawing.Point(8, 37);
            this.Arv_Temas.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Arv_Temas.Name = "Arv_Temas";
            this.Arv_Temas.Size = new System.Drawing.Size(347, 436);
            this.Arv_Temas.TabIndex = 17;
            // 
            // ProcessarTemas
            // 
            this.ProcessarTemas.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProcessarTemas.Location = new System.Drawing.Point(8, 481);
            this.ProcessarTemas.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ProcessarTemas.Name = "ProcessarTemas";
            this.ProcessarTemas.Size = new System.Drawing.Size(348, 37);
            this.ProcessarTemas.TabIndex = 18;
            this.ProcessarTemas.Text = "Processar Temas";
            this.ProcessarTemas.UseMnemonic = false;
            this.ProcessarTemas.UseVisualStyleBackColor = true;
            this.ProcessarTemas.Click += new System.EventHandler(this.ProcessarTemas_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Arial Narrow", 8F, System.Drawing.FontStyle.Bold);
            this.label3.Location = new System.Drawing.Point(5, 522);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(39, 15);
            this.label3.TabIndex = 19;
            this.label3.Text = "TEMAS";
            // 
            // ProcessarBairros
            // 
            this.ProcessarBairros.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProcessarBairros.Location = new System.Drawing.Point(361, 481);
            this.ProcessarBairros.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.ProcessarBairros.Name = "ProcessarBairros";
            this.ProcessarBairros.Size = new System.Drawing.Size(348, 37);
            this.ProcessarBairros.TabIndex = 21;
            this.ProcessarBairros.Text = "Processar Bairros";
            this.ProcessarBairros.UseVisualStyleBackColor = true;
            this.ProcessarBairros.Click += new System.EventHandler(this.ProcessarBarri_Click);
            // 
            // Arv_Bairros
            // 
            this.Arv_Bairros.CheckBoxes = true;
            this.Arv_Bairros.Font = new System.Drawing.Font("Arial Narrow", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Arv_Bairros.Location = new System.Drawing.Point(361, 37);
            this.Arv_Bairros.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Arv_Bairros.Name = "Arv_Bairros";
            this.Arv_Bairros.Size = new System.Drawing.Size(347, 436);
            this.Arv_Bairros.TabIndex = 20;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Arial Narrow", 8F, System.Drawing.FontStyle.Bold);
            this.label1.Location = new System.Drawing.Point(357, 522);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 15);
            this.label1.TabIndex = 22;
            this.label1.Text = "BAIRROS";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(717, 603);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ProcessarBairros);
            this.Controls.Add(this.Arv_Bairros);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ProcessarTemas);
            this.Controls.Add(this.Arv_Temas);
            this.Controls.Add(this.Processar);
            this.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(3, 4, 3, 4);
            this.Name = "Form1";
            this.Text = "Município do Rio de Janeiro";
            this.ResumeLayout(false);
            this.PerformLayout();

			}

		#endregion

        private System.Windows.Forms.TreeView Arv_Temas;
        private System.Windows.Forms.TreeView Arv_Bairros;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button Processar;
        private System.Windows.Forms.Button ProcessarTemas;
        private System.Windows.Forms.Button ProcessarBairros;

		}
}