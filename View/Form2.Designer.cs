namespace Lab_data.View
{
    partial class Form2
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
            DataGridViewCellStyle dataGridViewCellStyle1 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle2 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle3 = new DataGridViewCellStyle();
            DataGridViewCellStyle dataGridViewCellStyle4 = new DataGridViewCellStyle();
            materialDataTable1 = new MaterialSkin2DotNet.Controls.MaterialDataTable();
            Column1 = new DataGridViewTextBoxColumn();
            Column2 = new DataGridViewTextBoxColumn();
            materialButton1 = new MaterialSkin2DotNet.Controls.MaterialButton();
            ((System.ComponentModel.ISupportInitialize)materialDataTable1).BeginInit();
            SuspendLayout();
            // 
            // materialDataTable1
            // 
            materialDataTable1.AllowUserToDeleteRows = false;
            materialDataTable1.AllowUserToResizeRows = false;
            materialDataTable1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            materialDataTable1.BackgroundColor = Color.FromArgb(255, 255, 255);
            materialDataTable1.BorderStyle = BorderStyle.None;
            materialDataTable1.CellBorderStyle = DataGridViewCellBorderStyle.SunkenHorizontal;
            materialDataTable1.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dataGridViewCellStyle1.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = Color.FromArgb(255, 255, 255);
            dataGridViewCellStyle1.Font = new Font("Roboto Medium", 14F, FontStyle.Bold, GraphicsUnit.Pixel);
            dataGridViewCellStyle1.ForeColor = Color.FromArgb(222, 0, 0, 0);
            dataGridViewCellStyle1.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = DataGridViewTriState.True;
            materialDataTable1.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            materialDataTable1.ColumnHeadersHeight = 56;
            materialDataTable1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            materialDataTable1.Columns.AddRange(new DataGridViewColumn[] { Column1, Column2 });
            dataGridViewCellStyle2.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle2.BackColor = SystemColors.Window;
            dataGridViewCellStyle2.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            dataGridViewCellStyle2.ForeColor = Color.FromArgb(222, 0, 0, 0);
            dataGridViewCellStyle2.SelectionBackColor = Color.FromArgb(63, 81, 181);
            dataGridViewCellStyle2.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = DataGridViewTriState.False;
            materialDataTable1.DefaultCellStyle = dataGridViewCellStyle2;
            materialDataTable1.Depth = 0;
            materialDataTable1.Dock = DockStyle.Fill;
            materialDataTable1.EnableHeadersVisualStyles = false;
            materialDataTable1.Font = new Font("Roboto", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            materialDataTable1.GridColor = Color.FromArgb(239, 239, 239);
            materialDataTable1.Location = new Point(3, 64);
            materialDataTable1.MouseState = MaterialSkin2DotNet.MouseState.HOVER;
            materialDataTable1.Name = "materialDataTable1";
            dataGridViewCellStyle3.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle3.BackColor = SystemColors.Control;
            dataGridViewCellStyle3.Font = new Font("Roboto", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
            dataGridViewCellStyle3.ForeColor = SystemColors.WindowText;
            dataGridViewCellStyle3.SelectionBackColor = SystemColors.Highlight;
            dataGridViewCellStyle3.SelectionForeColor = SystemColors.HighlightText;
            dataGridViewCellStyle3.WrapMode = DataGridViewTriState.True;
            materialDataTable1.RowHeadersDefaultCellStyle = dataGridViewCellStyle3;
            materialDataTable1.RowHeadersVisible = false;
            dataGridViewCellStyle4.BackColor = Color.FromArgb(80, 80, 80);
            dataGridViewCellStyle4.ForeColor = Color.FromArgb(222, 0, 0, 0);
            materialDataTable1.RowsDefaultCellStyle = dataGridViewCellStyle4;
            materialDataTable1.RowTemplate.Height = 52;
            materialDataTable1.ScrollBars = ScrollBars.None;
            materialDataTable1.ShowVerticalScroll = false;
            materialDataTable1.Size = new Size(794, 383);
            materialDataTable1.TabIndex = 0;
            // 
            // Column1
            // 
            Column1.HeaderText = "Название";
            Column1.Name = "Column1";
            // 
            // Column2
            // 
            Column2.HeaderText = "Количество";
            Column2.Name = "Column2";
            // 
            // materialButton1
            // 
            materialButton1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            materialButton1.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            materialButton1.Density = MaterialSkin2DotNet.Controls.MaterialButton.MaterialButtonDensity.Default;
            materialButton1.Depth = 0;
            materialButton1.HighEmphasis = true;
            materialButton1.Icon = null;
            materialButton1.Location = new Point(629, 25);
            materialButton1.Margin = new Padding(4, 6, 4, 6);
            materialButton1.MouseState = MaterialSkin2DotNet.MouseState.HOVER;
            materialButton1.Name = "materialButton1";
            materialButton1.NoAccentTextColor = Color.Empty;
            materialButton1.Size = new Size(169, 36);
            materialButton1.TabIndex = 1;
            materialButton1.Text = "Сохранить в excel";
            materialButton1.Type = MaterialSkin2DotNet.Controls.MaterialButton.MaterialButtonType.Contained;
            materialButton1.UseAccentColor = false;
            materialButton1.UseVisualStyleBackColor = true;
            materialButton1.Click += materialButton1_Click;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
            Controls.Add(materialButton1);
            Controls.Add(materialDataTable1);
            Name = "Form2";
            Text = "Form2";
            ((System.ComponentModel.ISupportInitialize)materialDataTable1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private MaterialSkin2DotNet.Controls.MaterialDataTable materialDataTable1;
        private DataGridViewTextBoxColumn Column1;
        private DataGridViewTextBoxColumn Column2;
        private MaterialSkin2DotNet.Controls.MaterialButton materialButton1;
    }
}