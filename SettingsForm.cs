using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;

namespace Cliver.PdfMailer2
{
    public partial class SettingsForm : BaseForm//Form//
    {
        public SettingsForm()
        {
            InitializeComponent();

            Text = "Campaign Settings";
            Name = "Campaign";

            PartyProfiles.Add = PartyProfiles_Add;
            PartyProfiles.Select = PartyProfiles_Select;
            PartyProfiles.Delete = () => { Program.Settings.PartyProfileNames2PartyProfile.Remove(PartyProfiles.Names.Text); };
            foreach (string name in Program.Settings.PartyProfileNames2PartyProfile.Keys)
                PartyProfiles.Names.Items.Add(name);

            BuyerProfiles.Add = BuyerProfiles_Add;
            BuyerProfiles.Select = BuyerProfiles_Select;
            BuyerProfiles.Delete = () => { Program.Settings.BuyerProfileNames2BuyerProfile.Remove(BuyerProfiles.Names.Text); };
            foreach (string name in Program.Settings.BuyerProfileNames2BuyerProfile.Keys)
                BuyerProfiles.Names.Items.Add(name);

            BrokerProfiles.Add = BrokerProfiles_Add;
            BrokerProfiles.Select = BrokerProfiles_Select;
            BrokerProfiles.Delete = () => { Program.Settings.BrokerProfileNames2BrokerProfile.Remove(BrokerProfiles.Names.Text); };
            foreach (string name in Program.Settings.BrokerProfileNames2BrokerProfile.Keys)
                BrokerProfiles.Names.Items.Add(name);

            AgentProfiles.Add = AgentProfiles_Add;
            AgentProfiles.Select = AgentProfiles_Select;
            AgentProfiles.Delete = () => { Program.Settings.AgentProfileNames2AgentProfile.Remove(AgentProfiles.Names.Text); };
            foreach (string name in Program.Settings.AgentProfileNames2AgentProfile.Keys)
                AgentProfiles.Names.Items.Add(name);

            EscrowProfiles.Add = EscrowProfiles_Add;
            EscrowProfiles.Select = EscrowProfiles_Select;
            EscrowProfiles.Delete = () => { Program.Settings.EscrowProfileNames2EscrowProfile.Remove(EscrowProfiles.Names.Text); };
            foreach (string name in Program.Settings.EscrowProfileNames2EscrowProfile.Keys)
                EscrowProfiles.Names.Items.Add(name);

            EmailTemplateProfiles.Add = EmailTemplateProfiles_Add;
            EmailTemplateProfiles.Select = EmailTemplateProfiles_Select;
            EmailTemplateProfiles.Delete = () => { Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile.Remove(EmailTemplateProfiles.Names.Text); };
            foreach (string name in Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile.Keys)
                EmailTemplateProfiles.Names.Items.Add(name);

            EmailServerProfiles.Add = EmailServerProfiles_Add;
            EmailServerProfiles.Select = EmailServerProfiles_Select;
            EmailServerProfiles.Delete = () => { Program.Settings.EmailServerProfileNames2EmailServerProfile.Remove(EmailServerProfiles.Names.Text); };
            foreach (string name in Program.Settings.EmailServerProfileNames2EmailServerProfile.Keys)
                EmailServerProfiles.Names.Items.Add(name);

            PartyProfiles.Names.SelectedItem = Program.Settings.PartyProfileName;
            BuyerProfiles.Names.SelectedItem = Program.Settings.BuyerProfileName;
            BrokerProfiles.Names.SelectedItem = Program.Settings.BrokerProfileName;
            AgentProfiles.Names.SelectedItem = Program.Settings.AgentProfileName;
            EscrowProfiles.Names.SelectedItem = Program.Settings.EscrowProfileName;
            EmailTemplateProfiles.Names.SelectedItem = Program.Settings.EmailTemplateProfileName;
            EmailServerProfiles.Names.SelectedItem = Program.Settings.EmailServerProfileName;

            CloseOfEscrow.Value = Program.Settings.CloseOfEscrow;
            Emd.Text = Program.Settings.Emd;
            ShortSaleAddendum.Checked = Program.Settings.ShortSaleAddendum;
            OtherAddendum1.Checked = Program.Settings.OtherAddendum1;
            OtherAddendum2.Checked = Program.Settings.OtherAddendum2;

            UseRandomDelay.Checked = Program.Settings.UseRandomDelay;
            UseRandomDelay_CheckedChanged(null, null);
            MinRandomDelay.Text = Program.Settings.MinRandomDelayMss.ToString();
            MaxRandomDelay.Text = Program.Settings.MaxRandomDelayMss.ToString();

            foreach (string file in Program.Settings.AttachmentFiles)
                Attachments.Items.Add(file);
            foreach (int id in Program.Settings.SelectedAttachmentIds)
                Attachments.SetItemChecked(id, true);
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
            Program.EmailServerProfile p = Program.Settings.EmailServerProfileNames2EmailServerProfile[(string)EmailServerProfiles.Names.SelectedItem];
            EmailSenderEmail.Text = p.SenderEmail;
            SmtpHost.Text = p.SmtpHost;
            SmtpPassword.Text = p.SmtpPassword;
            SmtpPort.Text = p.SmtpPort.ToString();
        }

        private void EmailTemplateProfiles_Select()
        {
            if (EmailTemplateProfiles.Names.SelectedItem == null)
            {
                EmailBody.Text = "";
                EmailSubject.Text = "";
                return;
            }
            Program.EmailTemplateProfile p = Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[(string)EmailTemplateProfiles.Names.SelectedItem];
            EmailBody.Text = p.Body;
            EmailSubject.Text = p.Subject;
        }

        private void EscrowProfiles_Select()
        {
            if (EscrowProfiles.Names.SelectedItem == null)
            {
                EscrowOfficer.Text = "";
                EscrowTitleCompany.Text = "";
                return;
            }
            Program.EscrowProfile p = Program.Settings.EscrowProfileNames2EscrowProfile[(string)EscrowProfiles.Names.SelectedItem];
            EscrowOfficer.Text = p.Officer;
            EscrowTitleCompany.Text = p.TitleCompany;
        }

        private void AgentProfiles_Select()
        {
            if (AgentProfiles.Names.SelectedItem == null)
            {
                AgentEmail.Text = "";
                AgentInitial.ImageLocation = "";
                AgentLicenseNo.Text = "";
                AgentName.Text = "";
                AgentSignature.ImageLocation = "";
                return;
            }
            Program.AgentProfile p = Program.Settings.AgentProfileNames2AgentProfile[(string)AgentProfiles.Names.SelectedItem];
            AgentEmail.Text = p.Email;
            AgentInitial.ImageLocation = p.InitialFile;
            AgentLicenseNo.Text = p.LicenseNo;
            AgentName.Text = p.Name;
            AgentSignature.ImageLocation = p.SignatureFile;
        }

        private void BrokerProfiles_Select()
        {
            if (BrokerProfiles.Names.SelectedItem == null)
            {
                BrokerAddress.Text = "";
                BrokerCity.Text = "";
                BrokerCompany.Text = "";
                BrokerLicenseNo.Text = "";
                BrokerName.Text = "";
                BrokerPhone.Text = "";
                BrokerState.Text = "";
                BrokerZip.Text = "";
                return;
            }
            Program.BrokerProfile p = Program.Settings.BrokerProfileNames2BrokerProfile[(string)BrokerProfiles.Names.SelectedItem];
            BrokerAddress.Text = p.Address;
            BrokerCity.Text = p.City;
            BrokerCompany.Text = p.Company;
            BrokerLicenseNo.Text = p.LicenseNo;
            BrokerName.Text = p.Name;
            BrokerPhone.Text = p.Phone;
            BrokerState.Text = p.State;
            BrokerZip.Text = p.Zip;
        }

        private void BuyerProfiles_Select()
        {
            if (BuyerProfiles.Names.SelectedItem == null)
            {
                CoBuyerInitial.ImageLocation = null;
                CoBuyerName.Text = null;
                CoBuyerSignature.ImageLocation = null;
                BuyerInitial.ImageLocation = null;
                BuyerName.Text = null;
                BuyerSignature.ImageLocation = null;
                UseCoBuyer.Checked = false;
                UseCoBuyer_CheckedChanged(null, null);
                return;
            }
            Program.BuyerProfile p = Program.Settings.BuyerProfileNames2BuyerProfile[(string)BuyerProfiles.Names.SelectedItem];
            CoBuyerInitial.ImageLocation = p.CoBuyerInitialFile;
            CoBuyerName.Text = p.CoBuyerName;
            CoBuyerSignature.ImageLocation = p.CoBuyerSignatureFile;
            BuyerInitial.ImageLocation = p.InitialFile;
            BuyerName.Text = p.Name;
            BuyerSignature.ImageLocation = p.SignatureFile;
            UseCoBuyer.Checked = p.UseCoBuyer;
            UseCoBuyer_CheckedChanged(null, null);
        }

        private void PartyProfiles_Select()
        {
            if (PartyProfiles.Names.SelectedItem == null)
            {
                AgentProfiles.Names.SelectedItem = null;
                BrokerProfiles.Names.SelectedItem = null;
                BuyerProfiles.Names.SelectedItem = null;
                EscrowProfiles.Names.SelectedItem = null;
                return;
            }
            Program.PartyProfile p = Program.Settings.PartyProfileNames2PartyProfile[(string)PartyProfiles.Names.SelectedItem];
            AgentProfiles.Names.SelectedItem = p.AgentProfileName;
            BrokerProfiles.Names.SelectedItem = p.BrokerProfileName;
            BuyerProfiles.Names.SelectedItem = p.BuyerProfileName;
            EscrowProfiles.Names.SelectedItem = p.EscrowProfileName;
        }

        private bool EmailServerProfiles_Add()
        {
            string m1 = "";
            string m2 = " is not set.";

            Program.EmailServerProfile p = new Program.EmailServerProfile();

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

            Program.Settings.EmailServerProfileNames2EmailServerProfile[EmailServerProfiles.Names.Text] = p;
            return true;
        }

        private bool EmailTemplateProfiles_Add()
        {
            Program.EmailTemplateProfile p = new Program.EmailTemplateProfile();

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

            Program.Settings.EmailTemplateProfileNames2EmailTemplateProfileProfile[EmailTemplateProfiles.Names.Text] = p;
            return true;
        }

        private bool EscrowProfiles_Add()
        {
            Program.EscrowProfile p = new Program.EscrowProfile();

            p._ProfileName = EscrowProfiles.Names.Text;
            p.Officer = EscrowOfficer.Text;
            p.TitleCompany = EscrowTitleCompany.Text;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Escrow ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Officer))
            {
                Message.Exclaim(m1 + "EscrowOfficer" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.TitleCompany))
            {
                Message.Exclaim(m1 + "EscrowTitleCompany" + m2);
                return false;
            }

            Program.Settings.EscrowProfileNames2EscrowProfile[EscrowProfiles.Names.Text] = p;
            return true;
        }

        private bool AgentProfiles_Add()
        {
            Program.AgentProfile p = new Program.AgentProfile();

            p._ProfileName = AgentProfiles.Names.Text;
            p.Name = AgentName.Text;
            p.LicenseNo = AgentLicenseNo.Text;
            p.Email = AgentEmail.Text;
            p.InitialFile = AgentInitial.ImageLocation;
            p.SignatureFile = AgentSignature.ImageLocation;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Agent ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Name))
            {
                Message.Exclaim(m1 + "AgentName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.InitialFile))
            {
                Message.Exclaim(m1 + "AgentInitial" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.SignatureFile))
            {
                Message.Exclaim(m1 + "AgentSignature" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.LicenseNo))
            {
                Message.Exclaim(m1 + "AgentLicenseNo" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Email))
            {
                Message.Exclaim(m1 + "AgentEmail" + m2);
                return false;
            }

            Program.Settings.AgentProfileNames2AgentProfile[AgentProfiles.Names.Text] = p;
            return true;
        }

        private bool BrokerProfiles_Add()
        {
            Program.BrokerProfile p = new Program.BrokerProfile();

            p._ProfileName = BrokerProfiles.Names.Text;
            p.Name = BrokerName.Text;
            p.Address = BrokerAddress.Text;
            p.City = BrokerCity.Text;
            p.Company = BrokerCompany.Text;
            p.LicenseNo = BrokerLicenseNo.Text;
            p.Phone = BrokerPhone.Text;
            p.State = BrokerState.Text;
            p.Zip = BrokerZip.Text;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Broker ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Name))
            {
                Message.Exclaim(m1 + "BrokerName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Address))
            {
                Message.Exclaim(m1 + "BrokerAddress" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.City))
            {
                Message.Exclaim(m1 + "BrokerCity" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Company))
            {
                Message.Exclaim(m1 + "BrokerCompany" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.LicenseNo))
            {
                Message.Exclaim(m1 + "BrokerLicenseNo" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Phone))
            {
                Message.Exclaim(m1 + "BrokerPhone" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.State))
            {
                Message.Exclaim(m1 + "BrokerState" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Zip))
            {
                Message.Exclaim(m1 + "BrokerZip" + m2);
                return false;
            }

            Program.Settings.BrokerProfileNames2BrokerProfile[BrokerProfiles.Names.Text] = p;
            return true;
        }

        private bool BuyerProfiles_Add()
        {
            Program.BuyerProfile p = new Program.BuyerProfile();

            p._ProfileName = BuyerProfiles.Names.Text;
            p.Name = BuyerName.Text;
            p.InitialFile = BuyerInitial.ImageLocation;
            p.SignatureFile = BuyerSignature.ImageLocation;
            p.UseCoBuyer = UseCoBuyer.Checked;
            p.CoBuyerName = CoBuyerName.Text;
            p.CoBuyerInitialFile = CoBuyerInitial.ImageLocation;
            p.CoBuyerSignatureFile = CoBuyerSignature.ImageLocation;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Buyer ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.Name))
            {
                Message.Exclaim(m1 + "BuyerName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.InitialFile))
            {
                Message.Exclaim(m1 + "BuyerInitial" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.SignatureFile))
            {
                Message.Exclaim(m1 + "BuyerSignature" + m2);
                return false;
            }
            if (p.UseCoBuyer)
            {
                if (string.IsNullOrWhiteSpace(p.CoBuyerName))
                {
                    Message.Exclaim(m1 + "CoBuyerName" + m2);
                    return false;
                }
                if (string.IsNullOrWhiteSpace(p.CoBuyerInitialFile))
                {
                    Message.Exclaim(m1 + "CoBuyerInitial" + m2);
                    return false;
                }
                if (string.IsNullOrWhiteSpace(p.CoBuyerSignatureFile))
                {
                    Message.Exclaim(m1 + "CoBuyerSignature" + m2);
                    return false;
                }
            }

            Program.Settings.BuyerProfileNames2BuyerProfile[BuyerProfiles.Names.Text] = p;
            return true;
        }

        private bool PartyProfiles_Add()
        {
            Program.PartyProfile p = new Program.PartyProfile();

            if (!AgentProfiles_Add()
                || !BrokerProfiles_Add()
                || !BuyerProfiles_Add()
                || !EscrowProfiles_Add()
                )
                return false;

            p._ProfileName = PartyProfiles.Names.Text;
            p.AgentProfileName = AgentProfiles.Names.Text;
            p.BrokerProfileName = BrokerProfiles.Names.Text;
            p.BuyerProfileName = BuyerProfiles.Names.Text;
            p.EscrowProfileName = EscrowProfiles.Names.Text;

            string m1 = "";
            string m2 = " is not set.";
            if (string.IsNullOrWhiteSpace(p._ProfileName))
            {
                Message.Exclaim(m1 + "Party ProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.AgentProfileName))
            {
                Message.Exclaim(m1 + "AgentProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.BrokerProfileName))
            {
                Message.Exclaim(m1 + "BrokerProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.BuyerProfileName))
            {
                Message.Exclaim(m1 + "BuyerProfileName" + m2);
                return false;
            }
            if (string.IsNullOrWhiteSpace(p.EscrowProfileName))
            {
                Message.Exclaim(m1 + "EscrowProfileName" + m2);
                return false;
            }

            Program.Settings.PartyProfileNames2PartyProfile[PartyProfiles.Names.Text] = p;
            return true;
        }

        private void selectImage(PictureBox pictureBox)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Pick an image file";
            //d.Filter = "Filter tree files (*." + Program.FilterTreeFileExtension + ")|*." + Program.FilterTreeFileExtension + "|All files (*.*)|*.*";
            if (d.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(d.FileName))
                return;
            pictureBox.ImageLocation = d.FileName;
        }

        private void selectBuyerInitial_Click(object sender, EventArgs e)
        {
            selectImage(BuyerInitial);
        }

        private void selectBuyerSignature_Click(object sender, EventArgs e)
        {
            selectImage(BuyerSignature);
        }

        private void selectCoBuyerSignature_Click(object sender, EventArgs e)
        {
            selectImage(CoBuyerSignature);
        }

        private void selectAgentInitial_Click(object sender, EventArgs e)
        {
            selectImage(AgentInitial);
        }

        private void selectAgentSignature_Click(object sender, EventArgs e)
        {
            selectImage(AgentSignature);
        }

        private void selectCoBuyerInitial_Click(object sender, EventArgs e)
        {
            selectImage(CoBuyerInitial);
        }

        private void UseCoBuyer_CheckedChanged(object sender, EventArgs e)
        {
            gCoBuyer.Enabled = UseCoBuyer.Checked;
        }

        private bool save()
        {
            string m1 = "";
            string m2 = " is not set.";

            if (!PartyProfiles_Add())
                return false;
            if (!EmailServerProfiles_Add())
                return false;
            if (!EmailTemplateProfiles_Add())
                return false;
            
            if (string.IsNullOrWhiteSpace(PartyProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "PartyProfileName" + m2);
                return false;
            }
            Program.Settings.PartyProfileName = PartyProfiles.Names.Text;

            if (string.IsNullOrWhiteSpace(BuyerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "BrokerProfileName" + m2);
                return false;
            }
            Program.Settings.BuyerProfileName = BuyerProfiles.Names.Text;

            if (string.IsNullOrWhiteSpace(BrokerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "BuyerProfileName" + m2);
                return false;
            }
            Program.Settings.BrokerProfileName = BrokerProfiles.Names.Text;

            if (string.IsNullOrWhiteSpace(AgentProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "AgentProfileName" + m2);
                return false;
            }
            Program.Settings.AgentProfileName = AgentProfiles.Names.Text;

            if (string.IsNullOrWhiteSpace(EscrowProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EscrowProfileName" + m2);
                return false;
            }
            Program.Settings.EscrowProfileName = EscrowProfiles.Names.Text;
            
            if (string.IsNullOrWhiteSpace(EmailTemplateProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EmailTemplateProfileName" + m2);
                return false;
            }
            Program.Settings.EmailTemplateProfileName = EmailTemplateProfiles.Names.Text;

            if (string.IsNullOrWhiteSpace(EmailServerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EmailServerProfileName" + m2);
                return false;
            }
            Program.Settings.EmailServerProfileName = EmailServerProfiles.Names.Text;

            if (CloseOfEscrow.Value == null)
            {
                Message.Exclaim(m1 + "CloseOfEscrow" + m2);
                return false;
            }
            Program.Settings.CloseOfEscrow = CloseOfEscrow.Value;

            if (string.IsNullOrWhiteSpace(Emd.Text))
            {
                Message.Exclaim(m1 + "Emd" + m2);
                return false;
            }
            Program.Settings.Emd = Emd.Text;

            if (string.IsNullOrWhiteSpace(MinRandomDelay.Text))
            {
                Message.Exclaim(m1 + "MinRandomDelay" + m2);
                return false;
            }
            if (!int.TryParse(SmtpPort.Text, out Program.Settings.MinRandomDelayMss))
            {
                Message.Exclaim("MinRandomDelay is not number.");
                return false;
            }

            if (string.IsNullOrWhiteSpace(MaxRandomDelay.Text))
            {
                Message.Exclaim(m1 + "MaxRandomDelay" + m2);
                return false;
            }
            if (!int.TryParse(SmtpPort.Text, out Program.Settings.MaxRandomDelayMss))
            {
                Message.Exclaim("MaxRandomDelay is not number.");
                return false;
            }
            
            Program.Settings.ShortSaleAddendum = ShortSaleAddendum.Checked;
            Program.Settings.OtherAddendum1 = OtherAddendum1.Checked;
            Program.Settings.OtherAddendum2 = OtherAddendum2.Checked;
            Program.Settings.UseRandomDelay = UseRandomDelay.Checked;
            Program.Settings.SelectedAttachmentIds = new int[Attachments.CheckedIndices.Count];
            Program.Settings.AttachmentFiles = new string[Attachments.Items.Count];
            Attachments.Items.CopyTo(Program.Settings.AttachmentFiles, 0);
            Program.Settings.SelectedAttachmentIds = new int[Attachments.CheckedIndices.Count];
            Attachments.CheckedIndices.CopyTo(Program.Settings.SelectedAttachmentIds, 0);

            Program.Settings.Save();
            return true;
        }

        private void bImportAttachment_Click(object sender, EventArgs e)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Pick an attachment file";
            //d.Filter = "Filter tree files (*." + Program.FilterTreeFileExtension + ")|*." + Program.FilterTreeFileExtension + "|All files (*.*)|*.*";
            if (d.ShowDialog(this) != DialogResult.OK || string.IsNullOrWhiteSpace(d.FileName))
                return;
            Attachments.Items.Add(d.FileName);
        }

        private void DeleteAttachment_Click(object sender, EventArgs e)
        {
            Attachments.Items.Remove(Attachments.SelectedItem);
        }

        private void UseRandomDelay_CheckedChanged(object sender, EventArgs e)
        {
            MinRandomDelay.Enabled = UseRandomDelay.Checked;
            MaxRandomDelay.Enabled = UseRandomDelay.Checked;
        }

        private void bSave_Click(object sender, EventArgs e)
        {
            if (!save())
                return;
            Close();
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}