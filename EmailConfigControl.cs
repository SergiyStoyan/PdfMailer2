using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace Cliver.PdfMailer2
{
    public partial class EmailConfigControl : Cliver.BotGui.ConfigControl
    {
        override public string Section { get { return "Email"; } }

        public EmailConfigControl()
        {
            InitializeComponent();
        }

        override protected void SetToolTip()
        {
            //toolTip1.SetToolTip(this.DbConnectionString, "Database connection string");
        }

        override protected void Set()
        {
            set_group_box_values_from_config();

            EmailServerProfiles.Add = EmailServerProfiles_Add;
            EmailServerProfiles.Select = EmailServerProfiles_Select;
            EmailServerProfiles.Delete = () => { Settings.Email.EmailServerProfileNames2EmailServerProfile.Remove(EmailServerProfiles.Names.Text); };
            foreach (string name in Settings.Email.EmailServerProfileNames2EmailServerProfile.Keys)
                EmailServerProfiles.Names.Items.Add(name);

            EmailServerProfiles.Names.SelectedItem = Settings.Email.EmailServerProfileName;
            
            UseRandomDelay_CheckedChanged(null, null);
        }

        private bool EmailServerProfiles_Add()
        {
            string m1 = "";
            string m2 = " is not set.";

            Settings.EmailServerProfile p = new Settings.EmailServerProfile();

            p._ProfileName = EmailServerProfiles.Names.Text;
            p.SenderEmail = EmailSenderEmail.Text;
            p.SmtpHost = SmtpHost.Text;
            p.SmtpPassword = SmtpPassword.Text;
            if (string.IsNullOrWhiteSpace(SmtpPort.Text))
            {
                Message.Exclaim(m1 + "SmtpPort" + m2);
                return false;
            }
            if (!int.TryParse(SmtpPort.Text, out p.SmtpPort))
            {
                Message.Exclaim("SmtpPort is not number.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Escrow ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.SenderEmail))
            {
                Message.Exclaim(m1 + "SenderEmail" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.SmtpHost))
            {
                Message.Exclaim(m1 + "SmtpHost" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.SmtpPassword))
            {
                Message.Exclaim(m1 + "SmtpPassword" + m2);
                return false;
            }

            Settings.Email.EmailServerProfileNames2EmailServerProfile[EmailServerProfiles.Names.Text] = p;
            return true;
        }

        private void EmailServerProfiles_Select()
        {
            if (EmailServerProfiles.Names.SelectedItem == null)
            {
                EmailSenderEmail.Text = "";
                SmtpHost.Text = "";
                SmtpPassword.Text = "";
                SmtpPort.Text = "";
                return;
            }
            Settings.EmailServerProfile p = Settings.Email.EmailServerProfileNames2EmailServerProfile[(string)EmailServerProfiles.Names.SelectedItem];
            EmailSenderEmail.Text = p.SenderEmail;
            SmtpHost.Text = p.SmtpHost;
            SmtpPassword.Text = p.SmtpPassword;
            SmtpPort.Text = p.SmtpPort.ToString();
        }

        private void UseRandomDelay_CheckedChanged(object sender, EventArgs e)
        {
            MinRandomDelayMss.Enabled = UseRandomDelay.Checked;
            MaxRandomDelayMss.Enabled = UseRandomDelay.Checked;
        }

        override protected bool Get()
        {
            string m1 = "";
            string m2 = " is not set.";

            if (!EmailServerProfiles_Add())
                return false;

            if (string.IsNullOrWhiteSpace(EmailServerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EmailServerProfileName" + m2);
                return false;
            }
            Settings.Email.EmailServerProfileName = EmailServerProfiles.Names.Text;

            if (UseRandomDelay.Checked)
            {
                if (string.IsNullOrWhiteSpace(MinRandomDelayMss.Text))
                {
                    Message.Exclaim(m1 + "MinRandomDelayMss" + m2);
                    return false;
                }
                if (!int.TryParse(MinRandomDelayMss.Text, out Settings.Email.MinRandomDelayMss))
                {
                    Message.Exclaim("MinRandomDelayMss is not number.");
                    return false;
                }

                if (string.IsNullOrWhiteSpace(MaxRandomDelayMss.Text))
                {
                    Message.Exclaim(m1 + "MaxRandomDelayMss" + m2);
                    return false;
                }
                if (!int.TryParse(MaxRandomDelayMss.Text, out Settings.Email.MaxRandomDelayMss))
                {
                    Message.Exclaim("MaxRandomDelayMss is not number.");
                    return false;
                }

                if (Settings.Email.MinRandomDelayMss >= Settings.Email.MaxRandomDelayMss)
                {
                    Message.Exclaim("MinRandomDelayMss should be less than MaxRandomDelayMss.");
                    return false;
                }
            }

            Settings.Email.UseRandomDelay = UseRandomDelay.Checked;

            put_control_values_to_config(Name, group_box);
            return true;
        }
    }
}

