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
    public partial class PartiesConfigControl : Cliver.BotGui.ConfigControl
    {
        override public string Section { get { return "Parties"; } }

        public PartiesConfigControl()
        {
            InitializeComponent();
        }

        override protected void SetToolTip()
        {
            //toolTip1.SetToolTip(this.DbConnectionString, "Database connection string");
        }

        override protected void Set()
        {
            PartyProfiles.Add = PartyProfiles_Add;
            PartyProfiles.Select = PartyProfiles_Select;
            PartyProfiles.Delete = () => { Settings.Parties.PartyProfileNames2PartyProfile.Remove(PartyProfiles.Names.Text); };
            foreach (string name in Settings.Parties.PartyProfileNames2PartyProfile.Keys)
                PartyProfiles.Names.Items.Add(name);

            BuyerProfiles.Add = BuyerProfiles_Add;
            BuyerProfiles.Select = BuyerProfiles_Select;
            BuyerProfiles.Delete = () => { Settings.Parties.BuyerProfileNames2BuyerProfile.Remove(BuyerProfiles.Names.Text); };
            foreach (string name in Settings.Parties.BuyerProfileNames2BuyerProfile.Keys)
                BuyerProfiles.Names.Items.Add(name);

            BrokerProfiles.Add = BrokerProfiles_Add;
            BrokerProfiles.Select = BrokerProfiles_Select;
            BrokerProfiles.Delete = () => { Settings.Parties.BrokerProfileNames2BrokerProfile.Remove(BrokerProfiles.Names.Text); };
            foreach (string name in Settings.Parties.BrokerProfileNames2BrokerProfile.Keys)
                BrokerProfiles.Names.Items.Add(name);

            AgentProfiles.Add = AgentProfiles_Add;
            AgentProfiles.Select = AgentProfiles_Select;
            AgentProfiles.Delete = () => { Settings.Parties.AgentProfileNames2AgentProfile.Remove(AgentProfiles.Names.Text); };
            foreach (string name in Settings.Parties.AgentProfileNames2AgentProfile.Keys)
                AgentProfiles.Names.Items.Add(name);

            EscrowProfiles.Add = EscrowProfiles_Add;
            EscrowProfiles.Select = EscrowProfiles_Select;
            EscrowProfiles.Delete = () => { Settings.Parties.EscrowProfileNames2EscrowProfile.Remove(EscrowProfiles.Names.Text); };
            foreach (string name in Settings.Parties.EscrowProfileNames2EscrowProfile.Keys)
                EscrowProfiles.Names.Items.Add(name);

            PartyProfiles.Names.SelectedItem = Settings.Parties.PartyProfileName;
            BuyerProfiles.Names.SelectedItem = Settings.Parties.BuyerProfile._ProfileName;
            BrokerProfiles.Names.SelectedItem = Settings.Parties.BrokerProfile._ProfileName;
            AgentProfiles.Names.SelectedItem = Settings.Parties.AgentProfile._ProfileName;
            EscrowProfiles.Names.SelectedItem = Settings.Parties.EscrowProfile._ProfileName;
        }

        private void EscrowProfiles_Select()
        {
            if (EscrowProfiles.Names.SelectedItem == null)
            {
                EscrowOfficer.Text = "";
                EscrowTitleCompany.Text = "";
                return;
            }
            Settings.EscrowProfile p = Settings.Parties.EscrowProfileNames2EscrowProfile[(string)EscrowProfiles.Names.SelectedItem];
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
            Settings.AgentProfile p = Settings.Parties.AgentProfileNames2AgentProfile[(string)AgentProfiles.Names.SelectedItem];
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
            Settings.BrokerProfile p = Settings.Parties.BrokerProfileNames2BrokerProfile[(string)BrokerProfiles.Names.SelectedItem];
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
            Settings.BuyerProfile p = Settings.Parties.BuyerProfileNames2BuyerProfile[(string)BuyerProfiles.Names.SelectedItem];
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
            Settings.PartyProfile p = Settings.Parties.PartyProfileNames2PartyProfile[(string)PartyProfiles.Names.SelectedItem];
            AgentProfiles.Names.SelectedItem = p.AgentProfileName;
            BrokerProfiles.Names.SelectedItem = p.BrokerProfileName;
            BuyerProfiles.Names.SelectedItem = p.BuyerProfileName;
            EscrowProfiles.Names.SelectedItem = p.EscrowProfileName;
        }

        private bool EscrowProfiles_Add()
        {
            Settings.EscrowProfile p = new Settings.EscrowProfile();

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

            Settings.Parties.EscrowProfileNames2EscrowProfile[EscrowProfiles.Names.Text] = p;
            return true;
        }

        private bool AgentProfiles_Add()
        {
            Settings.AgentProfile p = new Settings.AgentProfile();

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

            Settings.Parties.AgentProfileNames2AgentProfile[AgentProfiles.Names.Text] = p;
            return true;
        }

        private bool BrokerProfiles_Add()
        {
            Settings.BrokerProfile p = new Settings.BrokerProfile();

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

            Settings.Parties.BrokerProfileNames2BrokerProfile[BrokerProfiles.Names.Text] = p;
            return true;
        }

        private bool BuyerProfiles_Add()
        {
            Settings.BuyerProfile p = new Settings.BuyerProfile();

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

            Settings.Parties.BuyerProfileNames2BuyerProfile[BuyerProfiles.Names.Text] = p;
            return true;
        }

        private bool PartyProfiles_Add()
        {
            Settings.PartyProfile p = new Settings.PartyProfile();

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

            Settings.Parties.PartyProfileNames2PartyProfile[PartyProfiles.Names.Text] = p;
            Settings.Parties.PartyProfileName = PartyProfiles.Names.Text;
            return true;
        }

        private void selectImage(PictureBox pictureBox)
        {
            OpenFileDialog d = new OpenFileDialog();
            d.Title = "Pick an image file";
            //d.Filter = "Filter tree files (*." + Settings.FilterTreeFileExtension + ")|*." + Settings.FilterTreeFileExtension + "|All files (*.*)|*.*";
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

        override protected bool Get()
        {
            string m1 = "";
            string m2 = " is not set.";

            if (!PartyProfiles_Add())
                return false;

            if (string.IsNullOrWhiteSpace(PartyProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "PartyProfileName" + m2);
                return false;
            }

            if (string.IsNullOrWhiteSpace(BuyerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "BrokerProfileName" + m2);
                return false;
            }

            if (string.IsNullOrWhiteSpace(BrokerProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "BuyerProfileName" + m2);
                return false;
            }

            if (string.IsNullOrWhiteSpace(AgentProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "AgentProfileName" + m2);
                return false;
            }

            if (string.IsNullOrWhiteSpace(EscrowProfiles.Names.Text))
            {
                Message.Exclaim(m1 + "EscrowProfileName" + m2);
                return false;
            }
            
            return true;
        }
    }
}

