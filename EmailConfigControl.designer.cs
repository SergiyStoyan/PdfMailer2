namespace Cliver.PdfMailer2
{
    partial class EmailConfigControl
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
            this.tabControl3 = new System.Windows.Forms.TabControl();
            this.tabPage10 = new System.Windows.Forms.TabPage();
            this.groupBox15 = new System.Windows.Forms.GroupBox();
            this.SmtpHost = new System.Windows.Forms.TextBox();
            this.label36 = new System.Windows.Forms.Label();
            this.label40 = new System.Windows.Forms.Label();
            this.SmtpPort = new System.Windows.Forms.TextBox();
            this.SmtpPassword = new System.Windows.Forms.TextBox();
            this.label38 = new System.Windows.Forms.Label();
            this.label39 = new System.Windows.Forms.Label();
            this.EmailSenderEmail = new System.Windows.Forms.TextBox();
            this.EmailServerProfiles = new Cliver.PdfMailer2.ProfilesControl();
            this.tabPage12 = new System.Windows.Forms.TabPage();
            this.UseRandomDelay = new System.Windows.Forms.CheckBox();
            this.label46 = new System.Windows.Forms.Label();
            this.MaxRandomDelayMss = new System.Windows.Forms.TextBox();
            this.MinRandomDelayMss = new System.Windows.Forms.TextBox();
            this.label47 = new System.Windows.Forms.Label();
            this.tabControl3.SuspendLayout();
            this.tabPage10.SuspendLayout();
            this.groupBox15.SuspendLayout();
            this.tabPage12.SuspendLayout();
            this.SuspendLayout();
            // 
            // group_box
            // 
            this.group_box.Size = new System.Drawing.Size(460, 332);
            this.group_box.Text = "TestCustom";
            // 
            // toolTip1
            // 
            this.toolTip1.AutoPopDelay = 100000;
            this.toolTip1.InitialDelay = 500;
            this.toolTip1.ReshowDelay = 100;
            // 
            // tabControl3
            // 
            this.tabControl3.Appearance = System.Windows.Forms.TabAppearance.FlatButtons;
            this.tabControl3.Controls.Add(this.tabPage10);
            this.tabControl3.Controls.Add(this.tabPage12);
            this.tabControl3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl3.Location = new System.Drawing.Point(0, 0);
            this.tabControl3.Name = "tabControl3";
            this.tabControl3.SelectedIndex = 0;
            this.tabControl3.Size = new System.Drawing.Size(460, 332);
            this.tabControl3.TabIndex = 144;
            // 
            // tabPage10
            // 
            this.tabPage10.Controls.Add(this.groupBox15);
            this.tabPage10.Controls.Add(this.label39);
            this.tabPage10.Controls.Add(this.EmailSenderEmail);
            this.tabPage10.Controls.Add(this.EmailServerProfiles);
            this.tabPage10.Location = new System.Drawing.Point(4, 25);
            this.tabPage10.Name = "tabPage10";
            this.tabPage10.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage10.Size = new System.Drawing.Size(452, 303);
            this.tabPage10.TabIndex = 1;
            this.tabPage10.Text = "Email";
            this.tabPage10.UseVisualStyleBackColor = true;
            // 
            // groupBox15
            // 
            this.groupBox15.Controls.Add(this.SmtpHost);
            this.groupBox15.Controls.Add(this.label36);
            this.groupBox15.Controls.Add(this.label40);
            this.groupBox15.Controls.Add(this.SmtpPort);
            this.groupBox15.Controls.Add(this.SmtpPassword);
            this.groupBox15.Controls.Add(this.label38);
            this.groupBox15.Location = new System.Drawing.Point(15, 70);
            this.groupBox15.Name = "groupBox15";
            this.groupBox15.Size = new System.Drawing.Size(256, 122);
            this.groupBox15.TabIndex = 59;
            this.groupBox15.TabStop = false;
            this.groupBox15.Text = "SMTP";
            // 
            // SmtpHost
            // 
            this.SmtpHost.Location = new System.Drawing.Point(16, 41);
            this.SmtpHost.Name = "SmtpHost";
            this.SmtpHost.Size = new System.Drawing.Size(183, 20);
            this.SmtpHost.TabIndex = 0;
            // 
            // label36
            // 
            this.label36.Location = new System.Drawing.Point(202, 25);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(48, 13);
            this.label36.TabIndex = 58;
            this.label36.Text = "Port:";
            this.label36.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label40
            // 
            this.label40.Location = new System.Drawing.Point(13, 25);
            this.label40.Name = "label40";
            this.label40.Size = new System.Drawing.Size(83, 13);
            this.label40.TabIndex = 50;
            this.label40.Text = "Host:";
            this.label40.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // SmtpPort
            // 
            this.SmtpPort.Location = new System.Drawing.Point(205, 41);
            this.SmtpPort.Name = "SmtpPort";
            this.SmtpPort.Size = new System.Drawing.Size(36, 20);
            this.SmtpPort.TabIndex = 57;
            // 
            // SmtpPassword
            // 
            this.SmtpPassword.Location = new System.Drawing.Point(16, 86);
            this.SmtpPassword.Name = "SmtpPassword";
            this.SmtpPassword.PasswordChar = '*';
            this.SmtpPassword.Size = new System.Drawing.Size(183, 20);
            this.SmtpPassword.TabIndex = 53;
            this.SmtpPassword.UseSystemPasswordChar = true;
            // 
            // label38
            // 
            this.label38.Location = new System.Drawing.Point(13, 72);
            this.label38.Name = "label38";
            this.label38.Size = new System.Drawing.Size(101, 13);
            this.label38.TabIndex = 54;
            this.label38.Text = "Password:";
            this.label38.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label39
            // 
            this.label39.Location = new System.Drawing.Point(28, 216);
            this.label39.Name = "label39";
            this.label39.Size = new System.Drawing.Size(155, 13);
            this.label39.TabIndex = 52;
            this.label39.Text = "Sender Address:";
            this.label39.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // EmailSenderEmail
            // 
            this.EmailSenderEmail.Location = new System.Drawing.Point(31, 235);
            this.EmailSenderEmail.Name = "EmailSenderEmail";
            this.EmailSenderEmail.Size = new System.Drawing.Size(225, 20);
            this.EmailSenderEmail.TabIndex = 51;
            // 
            // EmailServerProfiles
            // 
            this.EmailServerProfiles.BackColor = System.Drawing.SystemColors.ControlDark;
            this.EmailServerProfiles.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.EmailServerProfiles.Dock = System.Windows.Forms.DockStyle.Top;
            this.EmailServerProfiles.Location = new System.Drawing.Point(3, 3);
            this.EmailServerProfiles.Name = "EmailServerProfiles";
            this.EmailServerProfiles.Size = new System.Drawing.Size(446, 31);
            this.EmailServerProfiles.TabIndex = 59;
            // 
            // tabPage12
            // 
            this.tabPage12.Controls.Add(this.UseRandomDelay);
            this.tabPage12.Controls.Add(this.label46);
            this.tabPage12.Controls.Add(this.MaxRandomDelayMss);
            this.tabPage12.Controls.Add(this.MinRandomDelayMss);
            this.tabPage12.Controls.Add(this.label47);
            this.tabPage12.Location = new System.Drawing.Point(4, 25);
            this.tabPage12.Name = "tabPage12";
            this.tabPage12.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage12.Size = new System.Drawing.Size(452, 303);
            this.tabPage12.TabIndex = 3;
            this.tabPage12.Text = "Mailer";
            this.tabPage12.UseVisualStyleBackColor = true;
            // 
            // UseRandomDelay
            // 
            this.UseRandomDelay.AutoSize = true;
            this.UseRandomDelay.Location = new System.Drawing.Point(16, 26);
            this.UseRandomDelay.Name = "UseRandomDelay";
            this.UseRandomDelay.Size = new System.Drawing.Size(211, 17);
            this.UseRandomDelay.TabIndex = 59;
            this.UseRandomDelay.Text = "Use Random Delay Within The Range:";
            this.UseRandomDelay.UseVisualStyleBackColor = true;
            this.UseRandomDelay.CheckedChanged += new System.EventHandler(this.UseRandomDelay_CheckedChanged);
            // 
            // label46
            // 
            this.label46.Location = new System.Drawing.Point(166, 52);
            this.label46.Name = "label46";
            this.label46.Size = new System.Drawing.Size(121, 13);
            this.label46.TabIndex = 58;
            this.label46.Text = "milliseconds";
            this.label46.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // MaxRandomDelayMss
            // 
            this.MaxRandomDelayMss.Location = new System.Drawing.Point(106, 49);
            this.MaxRandomDelayMss.Name = "MaxRandomDelayMss";
            this.MaxRandomDelayMss.Size = new System.Drawing.Size(54, 20);
            this.MaxRandomDelayMss.TabIndex = 57;
            // 
            // MinRandomDelayMss
            // 
            this.MinRandomDelayMss.Location = new System.Drawing.Point(16, 49);
            this.MinRandomDelayMss.Name = "MinRandomDelayMss";
            this.MinRandomDelayMss.Size = new System.Drawing.Size(54, 20);
            this.MinRandomDelayMss.TabIndex = 0;
            // 
            // label47
            // 
            this.label47.Location = new System.Drawing.Point(76, 52);
            this.label47.Name = "label47";
            this.label47.Size = new System.Drawing.Size(24, 13);
            this.label47.TabIndex = 50;
            this.label47.Text = "<>";
            this.label47.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // EmailConfigControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.Controls.Add(this.tabControl3);
            this.Name = "EmailConfigControl";
            this.Size = new System.Drawing.Size(460, 332);
            this.Controls.SetChildIndex(this.group_box, 0);
            this.Controls.SetChildIndex(this.tabControl3, 0);
            this.tabControl3.ResumeLayout(false);
            this.tabPage10.ResumeLayout(false);
            this.tabPage10.PerformLayout();
            this.groupBox15.ResumeLayout(false);
            this.groupBox15.PerformLayout();
            this.tabPage12.ResumeLayout(false);
            this.tabPage12.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }





        #endregion

        private System.Windows.Forms.TabControl tabControl3;
        private System.Windows.Forms.TabPage tabPage10;
        private System.Windows.Forms.GroupBox groupBox15;
        private System.Windows.Forms.TextBox SmtpHost;
        private System.Windows.Forms.Label label36;
        private System.Windows.Forms.Label label40;
        private System.Windows.Forms.TextBox SmtpPort;
        private System.Windows.Forms.TextBox SmtpPassword;
        private System.Windows.Forms.Label label38;
        private System.Windows.Forms.Label label39;
        private System.Windows.Forms.TextBox EmailSenderEmail;
        private PdfMailer2.ProfilesControl EmailServerProfiles;
        private System.Windows.Forms.TabPage tabPage12;
        private System.Windows.Forms.CheckBox UseRandomDelay;
        private System.Windows.Forms.Label label46;
        private System.Windows.Forms.TextBox MaxRandomDelayMss;
        private System.Windows.Forms.TextBox MinRandomDelayMss;
        private System.Windows.Forms.Label label47;
    }
}
