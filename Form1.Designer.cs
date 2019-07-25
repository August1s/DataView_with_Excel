namespace DataView_with_Excel
{
    partial class Form1
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.Button_update = new System.Windows.Forms.Button();
            this.Button_insert = new System.Windows.Forms.Button();
            this.Button_delete = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // Button_update
            // 
            this.Button_update.Location = new System.Drawing.Point(36, 426);
            this.Button_update.Name = "Button_update";
            this.Button_update.Size = new System.Drawing.Size(75, 23);
            this.Button_update.TabIndex = 0;
            this.Button_update.Text = "update";
            this.Button_update.UseVisualStyleBackColor = true;
            this.Button_update.Click += new System.EventHandler(this.Button_update_Click);
            // 
            // Button_insert
            // 
            this.Button_insert.Location = new System.Drawing.Point(117, 426);
            this.Button_insert.Name = "Button_insert";
            this.Button_insert.Size = new System.Drawing.Size(75, 23);
            this.Button_insert.TabIndex = 1;
            this.Button_insert.Text = "insert";
            this.Button_insert.UseVisualStyleBackColor = true;
            this.Button_insert.Click += new System.EventHandler(this.Button_insert_Click);
            // 
            // Button_delete
            // 
            this.Button_delete.Location = new System.Drawing.Point(198, 426);
            this.Button_delete.Name = "Button_delete";
            this.Button_delete.Size = new System.Drawing.Size(75, 23);
            this.Button_delete.TabIndex = 2;
            this.Button_delete.Text = "delete";
            this.Button_delete.UseVisualStyleBackColor = true;
            this.Button_delete.Click += new System.EventHandler(this.Button_delete_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(36, 43);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(457, 342);
            this.dataGridView1.TabIndex = 3;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(540, 470);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.Button_delete);
            this.Controls.Add(this.Button_insert);
            this.Controls.Add(this.Button_update);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Button_update;
        private System.Windows.Forms.Button Button_insert;
        private System.Windows.Forms.Button Button_delete;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}

