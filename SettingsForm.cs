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
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();

            PartyProfiles.Save= PartyProfiles_Save;
            foreach (string name in Program.Profiles.PartyProfileNames2PartyProfile.Keys)
                PartyProfiles.Names.Items.Add(name);

            BuyerProfiles.Save = BuyerProfiles_Save;
            foreach (string name in Program.Profiles.BuyerProfileNames2BuyerProfile.Keys)
                BuyerProfiles.Names.Items.Add(name);

            BrokerProfiles.Save = BrokerProfiles_Save;
            foreach (string name in Program.Profiles.BrokerProfileNames2BrokerProfile.Keys)
                BrokerProfiles.Names.Items.Add(name);

            AgentProfiles.Save = AgentProfiles_Save;
            foreach (string name in Program.Profiles.AgentProfileNames2AgentProfile.Keys)
                AgentProfiles.Names.Items.Add(name);

            EscrowProfiles.Save = EscrowProfiles_Save;
            foreach (string name in Program.Profiles.EscrowProfileNames2EscrowProfile.Keys)
                EscrowProfiles.Names.Items.Add(name);

            EmailTemplateProfiles.Save = EmailTemplateProfiles_Save;
            foreach (string name in Program.Profiles.EmailTemplateProfileNames2EmailTemplateProfileProfile.Keys)
                EmailTemplateProfiles.Names.Items.Add(name);

            EmailServerProfiles.Save = EmailServerProfiles_Save;
            foreach (string name in Program.Profiles.EmailServerProfileNames2EmailServerProfile.Keys)
                EmailServerProfiles.Names.Items.Add(name);
        }

        private bool EmailServerProfiles_Save()
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

            Program.Profiles.EmailServerProfileNames2EmailServerProfile[EmailServerProfiles.Names.Text] = p;
            EmailServerProfiles.Names.Items.Remove(EmailServerProfiles.Names.Text);
            EmailServerProfiles.Names.Items.Insert(0, EmailServerProfiles.Names.Text);
            EmailServerProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool EmailTemplateProfiles_Save()
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

            Program.Profiles.EmailTemplateProfileNames2EmailTemplateProfileProfile[EmailTemplateProfiles.Names.Text] = p;
            EmailTemplateProfiles.Names.Items.Remove(EmailTemplateProfiles.Names.Text);
            EmailTemplateProfiles.Names.Items.Insert(0, EmailTemplateProfiles.Names.Text);
            EmailTemplateProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool EscrowProfiles_Save()
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

            Program.Profiles.EscrowProfileNames2EscrowProfile[EscrowProfiles.Names.Text] = p;
            EscrowProfiles.Names.Items.Remove(EscrowProfiles.Names.Text);
            EscrowProfiles.Names.Items.Insert(0, EscrowProfiles.Names.Text);
            EscrowProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool AgentProfiles_Save()
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

            Program.Profiles.AgentProfileNames2AgentProfile[AgentProfiles.Names.Text] = p;
            AgentProfiles.Names.Items.Remove(AgentProfiles.Names.Text);
            AgentProfiles.Names.Items.Insert(0, AgentProfiles.Names.Text);
            AgentProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool BrokerProfiles_Save()
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

            Program.Profiles.BrokerProfileNames2BrokerProfile[BrokerProfiles.Names.Text] = p;
            BrokerProfiles.Names.Items.Remove(BrokerProfiles.Names.Text);
            BrokerProfiles.Names.Items.Insert(0, BrokerProfiles.Names.Text);
            BrokerProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool BuyerProfiles_Save()
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

            Program.Profiles.BuyerProfileNames2BuyerProfile[BuyerProfiles.Names.Text] = p;
            BuyerProfiles.Names.Items.Remove(BuyerProfiles.Names.Text);
            BuyerProfiles.Names.Items.Insert(0, BuyerProfiles.Names.Text);
            BuyerProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private bool PartyProfiles_Save()
        {
            Program.PartyProfile p = new Program.PartyProfile();

            if (!AgentProfiles_Save()
                || !BrokerProfiles_Save()
                || !BuyerProfiles_Save()
                || !EscrowProfiles_Save()
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

            Program.Profiles.PartyProfileNames2PartyProfile[PartyProfiles.Names.Text] = p;
            PartyProfiles.Names.Items.Remove(PartyProfiles.Names.Text);
            PartyProfiles.Names.Items.Insert(0, PartyProfiles.Names.Text);
            PartyProfiles.Names.SelectedIndex = 0;
            return true;
        }

        private void selectBuyerInitial_Click(object sender, EventArgs e)
        {

        }

        private void selectBuyerSignature_Click(object sender, EventArgs e)
        {

        }

        private void selectCobuyerSignature_Click(object sender, EventArgs e)
        {

        }

        private void useCobuyerName_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void deleteBroker_Click(object sender, EventArgs e)
        {

        }

        private void saveBroker_Click(object sender, EventArgs e)
        {

        }

        private void useCoBuyer_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void Start_Click(object sender, EventArgs e)
        {

        }

        //private List<string> get_profiles(string key)
        //{
        //}

    }
}
