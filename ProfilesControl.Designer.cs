namespace Cliver.PdfMailer2
{
    partial class ProfilesControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.bDelete = new System.Windows.Forms.Button();
            this.Names = new System.Windows.Forms.ComboBox();
            this.bSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // bDelete
            // 
            this.bDelete.BackColor = System.Drawing.SystemColors.Control;
            this.bDelete.Location = new System.Drawing.Point(170, 9);
            this.bDelete.Name = "bDelete";
            this.bDelete.Size = new System.Drawing.Size(24, 23);
            this.bDelete.TabIndex = 55;
            this.bDelete.Text = "-";
            this.bDelete.UseVisualStyleBackColor = false;
            this.bDelete.Click += new System.EventHandler(this.bDelete_Click);
            // 
            // Names
            // 
            this.Names.FormattingEnabled = true;
            this.Names.Location = new System.Drawing.Point(13, 10);
            this.Names.Name = "Names";
            this.Names.Size = new System.Drawing.Size(121, 21);
            this.Names.TabIndex = 53;
            // 
            // bSave
            // 
            this.bSave.BackColor = System.Drawing.SystemColors.Control;
            this.bSave.Location = new System.Drawing.Point(140, 9);
            this.bSave.Name = "bSave";
            this.bSave.Size = new System.Drawing.Size(24, 23);
            this.bSave.TabIndex = 54;
            this.bSave.Text = "+";
            this.bSave.UseVisualStyleBackColor = false;
            this.bSave.Click += new System.EventHandler(this.bSave_Click);
            // 
            // ProfilesControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlDark;
            this.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.Controls.Add(this.bDelete);
            this.Controls.Add(this.Names);
            this.Controls.Add(this.bSave);
            this.Name = "ProfilesControl";
            this.Size = new System.Drawing.Size(404, 42);
            this.ResumeLayout(false);

        }

        #endregion
        internal System.Windows.Forms.ComboBox Names;
        internal System.Windows.Forms.Button bDelete;
        internal System.Windows.Forms.Button bSave;
    }
}
