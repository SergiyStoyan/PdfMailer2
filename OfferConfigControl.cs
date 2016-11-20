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
    public partial class OfferConfigControl : Cliver.BotGui.ConfigControl
    {
        override public string Section { get { return "Offer"; } }

        public OfferConfigControl()
        {
            InitializeComponent();
        }

        override protected void SetToolTip()
        {
            //toolTip1.SetToolTip(this.DbConnectionString, "Database connection string");
        }

        override protected void Set()
        {
            EmailTemplateProfiles.Add = EmailTemplateProfiles_Add;
            EmailTemplateProfiles.Select = EmailTemplateProfiles_Select;
            EmailTemplateProfiles.Delete = () => { Settings.Offer.EmailTemplateProfileNames2EmailTemplateProfileProfile.Remove(EmailTemplateProfiles.Names.Text); };
            foreach (string name in Settings.Offer.EmailTemplateProfileNames2EmailTemplateProfileProfile.Keys)
                EmailTemplateProfiles.Names.Items.Add(name);

            EmailTemplateProfiles.Names.SelectedItem = Settings.Offer.EmailTemplateProfileName;

            foreach (string file in Settings.Offer.AttachmentFiles)
                Attachments.Items.Add(file);
            foreach (int id in Settings.Offer.SelectedAttachmentIds)
                Attachments.SetItemChecked(id, true);

            CloseOfEscrow.Value = Settings.Offer.CloseOfEscrow;
            Emd.Text = Settings.Offer.Emd;
            ShortSaleAddendum.Checked = Settings.Offer.ShortSaleAddendum;
            OtherAddendum1.Checked = Settings.Offer.OtherAddendum1;
            OtherAddendum2.Checked = Settings.Offer.OtherAddendum2;
        }

        private bool EmailTemplateProfiles_Add()
        {
            Settings.EmailTemplateProfile p = new Settings.EmailTemplateProfile();

            p._ProfileName = EmailTemplateProfiles.Names.Text;
            p.Body = EmailBody.Text;
            p.Subject = EmailSubject.Text;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "EmailTemplate ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Body))
            {
                Message.Exclaim(m1 + "EmailBody" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Subject))
            {
                Message.Exclaim(m1 + "EmailSubject" + m2);
                return false;
            }

            Settings.Offer.EmailTemplateProfileNames2EmailTemplateProfileProfile[EmailTemplateProfiles.Names.Text] = p;
            return true;
        }

        private void EmailTemplateProfiles_Select()
        {
            if (EmailTemplateProfiles.Names.SelectedItem == null)
            {
                EmailBody.Text = "";
                EmailSubject.Text = "";
                return;
            }
            Settings.EmailTemplateProfile p = Settings.Offer.EmailTemplateProfileNames2EmailTemplateProfileProfile[(string)EmailTemplateProfiles.Names.SelectedItem];
            EmailBody.Text = p.Body;
            EmailSubject.Text = p.Subject;
        }

        private void bImportAttachment_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Pick an attachment file";
            //d.Filter = "Filter tree files (*." + Settings.FilterTreeFileExtension + ")|*." + Settings.FilterTreeFileExtension + "|All files (*.*)|*.*";
            if (d.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(d.FileName))
                return;
            Attachments.Items.Add(d.FileName);
        }

        private void DeleteAttachment_Click(object sender, EventArgs e)
        {
            Attachments.Items.Remove(Attachments.SelectedItem);
        }

        override protected bool Get()
        {
            string m1 = "";
            string m2 = " is not set.";

            if (!EmailTemplateProfiles_Add())
                return false;

            if (string.IsNullOrWhiteSpace(EmailTemplateProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EmailServerProfileName" + m2);
                return false;
            }
            Settings.Offer.EmailTemplateProfileName = EmailTemplateProfiles.Names.Text;

            if (CloseOfEscrow.Value == null)
            {
                Message.Exclaim(m1 + "CloseOfEscrow" + m2);
                return false;
            }
            Settings.Offer.CloseOfEscrow = CloseOfEscrow.Value;

            if (string.IsNullOrWhiteSpace(Emd.Text))
            {
                Message.Exclaim(m1 + "Emd" + m2);
                return false;
            }
            Settings.Offer.Emd = Emd.Text;

            Settings.Offer.ShortSaleAddendum = ShortSaleAddendum.Checked;
            Settings.Offer.OtherAddendum1 = OtherAddendum1.Checked;
            Settings.Offer.OtherAddendum2 = OtherAddendum2.Checked;

            Settings.Offer.SelectedAttachmentIds = new int[Attachments.CheckedIndices.Count];
            Settings.Offer.AttachmentFiles = new string[Attachments.Items.Count];
            Attachments.Items.CopyTo(Settings.Offer.AttachmentFiles, 0);
            Settings.Offer.SelectedAttachmentIds = new int[Attachments.CheckedIndices.Count];
            Attachments.CheckedIndices.CopyTo(Settings.Offer.SelectedAttachmentIds, 0);

            return true;
        }
    }
}