//********************************************************************************************
//Author: Sergey Stoyan, CliverSoft.com
//        http://cliversoft.com
//        stoyan@cliversoft.com
//        sergey.stoyan@gmail.com
//        27 February 2007
//Copyright: (C) 2007, Sergey Stoyan
//********************************************************************************************

using System;
using System.Linq;
using System.Net;
using System.Text;
using System.IO;
using System.Threading;
using System.Text.RegularExpressions;
using System.Web;
using System.Data;
using System.Web.Script.Serialization;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Net.Mail;
using Cliver;
using System.Configuration;
using System.Windows.Forms;
using Cliver.Bot;
using Cliver.BotGui;
using Microsoft.Win32;
using System.Reflection;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Cliver.PdfMailer2
{
    public class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Cliver.Config.Initialize(new string[] { "Parties", "Offer", "Email", "Engine", "Input", "Log"});
                Cliver.BotGui.BotGui.ConfigControlSections=new string[] { "Parties", "Offer", "Email", "Engine", "Input", /*"Output", "Web", "Browser", "Spider", "Proxy",*/ "Log", };
                //Cliver.Bot.Program.Run();//It is the entry when the app runs as a console app.
                Cliver.BotGui.Program.Run();//It is the entry when the app uses the default GUI.
            }
            catch(Exception e)
            {
                LogMessage.Error(e);
            }
        }
    }

    public class CustomBot : Cliver.Bot.Bot
    {
        new static public string GetAbout()
        {
            return @"PDF MAILER
Created: " + Cliver.Bot.Program.GetCustomizationCompiledTime().ToString() + @"
Developed by: www.cliversoft.com";
        }
        
        new static public void SessionCreating()
        {
            InternetDateTime.CHECK_TEST_PERIOD_VALIDITY(2016, 12, 9);

            Cliver.BotGui.Program.BindProgressBar2InputItemQueue<EmailItem>();
            BotCycle.TreatExceptionAsFatal = true;
            Session.GetInputItemQueue<EmailItem>().PickNext = pick_next_PdfItem;
        }

        static InputItem pick_next_PdfItem(System.Collections.IEnumerator items_ennumerator)
        {
            lock (emails2sent_time)
            {
                int delay = Settings.Email.MinRandomDelayMss + (int)((float)(Settings.Email.MaxRandomDelayMss - Settings.Email.MinRandomDelayMss) * random.NextDouble());
                TimeSpan min_wait_period = TimeSpan.MaxValue;
                EmailItem min_wait_period_ei = null;
                for (items_ennumerator.Reset(); items_ennumerator.MoveNext();)
                {
                    EmailItem ei = ((EmailItem)items_ennumerator.Current);
                    string to_email = ei.Data.ListAgentEmail;
                    DateTime sent_time;
                    if (!emails2sent_time.TryGetValue(to_email, out sent_time))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return ei;
                    }
                    TimeSpan wait_period = sent_time.AddMilliseconds(delay) - DateTime.Now;
                    if (wait_period < TimeSpan.FromSeconds(0))
                    {
                        emails2sent_time[to_email] = DateTime.Now;
                        return ei;
                    }
                    if (wait_period > min_wait_period)
                    {
                        min_wait_period = wait_period;
                        min_wait_period_ei = ei;
                    }
                }
                if (min_wait_period_ei != null && Session.GetInputItemQueue<DataItem>().CountOfNew < 1)
                {
                    Thread.Sleep(min_wait_period);
                    emails2sent_time[min_wait_period_ei.Data.ListAgentEmail] = DateTime.Now;
                    return min_wait_period_ei;
                }
                return null;
            }
        }
        static Random random = new Random();
        static Dictionary<string, DateTime> emails2sent_time = new Dictionary<string, DateTime>();

        new static public void SessionClosing()
        {
        }

        override public void CycleStarting()
        {
        }

        public class DataItem : InputItem
        {
            readonly public string Status;
            readonly public string ML_Id;
            readonly public string Address;
            readonly public string UnitNumber;
            readonly public string City;
            readonly public string ZipCode;
            readonly public string ParcelNumber;
            readonly public string BldgDes;
            readonly public string SubdivisionName;
            readonly public string Sub;
            readonly public string ShortSale;
            readonly public string ApproxLivArea;
            readonly public string BedsTotal;
            readonly public string BathsTotal;
            readonly public string Garage;
            readonly public string PvPool;
            readonly public string Spa;
            readonly public string YearBuilt;
            readonly public string LotSqft;
            readonly public string CurrentPrice;
            readonly public string OfferAmt;
            readonly public string AdditionalTerms;
            readonly public string Agent2AgentRemarks;
            readonly public string ListAgentFullName;
            readonly public string ListAgentEmail;
            readonly public string ListOfficeName;
            readonly public string ActualCloseDate;
            readonly public string ListAgentDirectWorkPhone;
            readonly public string DOM;
            readonly public string OfferSentDate;
            
            override public void PROCESSOR(BotCycle bc)
            {
                // throw new Session.FatalException("fdsfsfs");
                CustomBot cb = (CustomBot)bc.Bot;

                string d = Log.OutputDir + "\\" + Log.This.Id + "_" + DateTime.Now.GetSecondsSinceUnixEpoch();
                Directory.CreateDirectory(d);

                string address = new System.Globalization.CultureInfo("en-US", false).TextInfo.ToTitleCase(Address.ToLower());
                
                string output_addendum1_pdf = null;
                if (Settings.Offer.OtherAddendum1)
                {
                    output_addendum1_pdf = d + "\\" + Regex.Replace("duites " + address + ".pdf", @"\s+", " ");
                    PdfReader.unethicalreading = true;
                    PdfReader pr = new PdfReader(template_addendum1_pdf);
                    PdfStamper ps = new PdfStamper(pr, new FileStream(output_addendum1_pdf, FileMode.Create, FileAccess.Write, FileShare.None));

                    //string fs = "";
                    //foreach (KeyValuePair<string, AcroFields.Item> kvp in ps.AcroFields.Fields)
                    //    fs += "\n{\"" + kvp.Key + "\", \"\"},";

                    set_field(ps.AcroFields, "AgentName", Settings.Parties.AgentProfile.Name);
                    set_field(ps.AcroFields, "Agents License", Settings.Parties.AgentProfile.LicenseNo);
                    set_field(ps.AcroFields, "Buyer Name", Settings.Parties.BuyerProfile.Name);
                    set_field(ps.AcroFields, "Co Buyer Name", Settings.Parties.BuyerProfile.CoBuyerName);
                    set_field(ps.AcroFields, "Buyer Broker", Settings.Parties.BrokerProfile.Name);
                    set_field(ps.AcroFields, "Company Name", Settings.Parties.BrokerProfile.Company);
                    set_field(ps.AcroFields, "Date_3", DateTime.Now.ToShortDateString());
                    set_field(ps.AcroFields, "Time_3", DateTime.Now.ToShortTimeString());
                    set_field(ps.AcroFields, "Date_4", DateTime.Now.ToShortDateString());
                    set_field(ps.AcroFields, "Time_4", DateTime.Now.ToShortTimeString());

                    ps.FormFlattening = true;

                    {
                        var pcb = ps.GetOverContent(1);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.InitialFile), new System.Drawing.Point(70, 190));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(140, 190));
                    }
                    {
                        var pcb = ps.GetOverContent(1);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.SignatureFile), new System.Drawing.Point(110, 85));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerSignatureFile), new System.Drawing.Point(110, 85));
                    }

                    ps.Close();
                    pr.Close();
                }

                string output_pdf = d + "\\" + Regex.Replace(address + " RPA.pdf", @"\s+", " ");
                {
                    //lock (template_pdf)
                    //{
                    //    File.Copy(template_pdf, pdf);
                    //}

                    PdfReader.unethicalreading = true;
                    PdfReader pr = new PdfReader(template_pdf);
                    //pr.RemoveUsageRights();
                    //pr.SelectPages("7,8");
                    PdfStamper ps = new PdfStamper(pr, new FileStream(output_pdf, FileMode.Create, FileAccess.Write, FileShare.None));

                    string fs = "";
                    foreach (KeyValuePair<string, AcroFields.Item> kvp in ps.AcroFields.Fields)
                        fs += "\n{\"" + kvp.Key + "\", \"\"},";

                    set_field(ps.AcroFields, "Todays Date", DateTime.Today.ToShortDateString());
                    set_field(ps.AcroFields, "Buyer Name", Settings.Parties.BuyerProfile.Name);
                    set_field(ps.AcroFields, "Address and Unit Number", address + " " + UnitNumber);
                    set_field(ps.AcroFields, "City/Town", City);
                    //set_field(ps.AcroFields, "CLARK", );
                    set_field(ps.AcroFields, "Zip", ZipCode);
                    set_field(ps.AcroFields, "PARCEL NUMBER", ParcelNumber);
                    set_field(ps.AcroFields, "OfferAmt", OfferAmt);
                    string oa = Regex.Replace(OfferAmt, @"[^\d]", "");
                    if (oa.Length > 0)
                        set_field(ps.AcroFields, "OfferAmt in words", ConvertionRoutines.NumberToWords(int.Parse(oa)).ToUpper());
                    set_field(ps.AcroFields, "EMD", Settings.Offer.Emd);
                    //set_field(ps.AcroFields, "Check Box1", );
                    //set_field(ps.AcroFields, "Balance", );
                    set_field(ps.AcroFields, "Co Buyer Name", Settings.Parties.BuyerProfile.CoBuyerName);
                    set_field(ps.AcroFields, "<address> <UnitNumber> <City/town> NV <ZIP Code>", address + " " + UnitNumber + ", " + City + " NV " + ZipCode);
                    set_field(ps.AcroFields, "ML#", "ALL PER ML# " + ML_Id);
                    set_field(ps.AcroFields, "Title Company", Settings.Parties.EscrowProfile.TitleCompany);
                    set_field(ps.AcroFields, "Escrow Officer", Settings.Parties.EscrowProfile.Officer);
                    set_field(ps.AcroFields, "Close of Escrow", Settings.Offer.CloseOfEscrow.ToShortDateString());

                    //                    string AdditionalTerms = @"This form is available for use by the real estate industry. It is not intended to identify the user as a REALTOR®.
                    //8 REALTOR® is a registered collective membership mark which may be used only by members of the NATIONAL
                    //9 ASSOCIATION OF REALTORS® who subscribe to its Code.";
                    string s = AdditionalTerms;
                    s = fill_field_by_words(ps.AcroFields, "AdditionalTerms1", s);
                    s = fill_field_by_words(ps.AcroFields, "AdditionalTerms2", s);
                    s = fill_field_by_words(ps.AcroFields, "AdditionalTerms3", s);
                    if (s.Length > 0)
                    {
                        s = AdditionalTerms;
                        s = fill_field_by_chars(ps.AcroFields, "AdditionalTerms1", s);
                        s = fill_field_by_chars(ps.AcroFields, "AdditionalTerms2", s);
                        set_field(ps.AcroFields, "AdditionalTerms3", s);
                    }

                    set_field(ps.AcroFields, "Buyer Broker", Settings.Parties.BrokerProfile.Name);
                    set_field(ps.AcroFields, "Agent Name", Settings.Parties.AgentProfile.Name);
                    set_field(ps.AcroFields, "Company Name", Settings.Parties.BrokerProfile.Company);
                    set_field(ps.AcroFields, "Agents License", Settings.Parties.AgentProfile.LicenseNo);
                    set_field(ps.AcroFields, "Brokers License", Settings.Parties.BrokerProfile.LicenseNo);
                    set_field(ps.AcroFields, "Office Address", Settings.Parties.BrokerProfile.Address);
                    set_field(ps.AcroFields, "Office Phone", Settings.Parties.BrokerProfile.Phone);
                    set_field(ps.AcroFields, "City State Zip", Settings.Parties.BrokerProfile.City + " " + Settings.Parties.BrokerProfile.State + " " + Settings.Parties.BrokerProfile.Zip);
                    set_field(ps.AcroFields, "Agent Email", Settings.Parties.AgentProfile.Email);
                    //set_field(ps.AcroFields, "Response Month", );
                    //set_field(ps.AcroFields, "Response day", );
                    //set_field(ps.AcroFields, "year", );
                    //set_field(ps.AcroFields,  "Date_2", );
                    //set_field(ps.AcroFields, "Date_3", );
                    //set_field(ps.AcroFields, "Sellers Broker", );
                    set_field(ps.AcroFields, "Agents Name_2", Settings.Parties.AgentProfile.Name);
                    set_field(ps.AcroFields, "Company Name_2", Settings.Parties.BrokerProfile.Company);
                    set_field(ps.AcroFields, "Agents License Number_2", Settings.Parties.AgentProfile.LicenseNo);
                    set_field(ps.AcroFields, "Brokers License Number_2", Settings.Parties.BrokerProfile.LicenseNo);
                    set_field(ps.AcroFields, "Office Address_2", Settings.Parties.BrokerProfile.Address);
                    set_field(ps.AcroFields, "Phone_2", Settings.Parties.BrokerProfile.Phone);
                    set_field(ps.AcroFields, "City State Zip_2", Settings.Parties.BrokerProfile.Zip);
                    set_field(ps.AcroFields, "Email_2", Settings.Parties.AgentProfile.Email);

                    ps.FormFlattening = true;

                    for (int i = 1; i <= pr.NumberOfPages; i++)
                    {
                        var pcb = ps.GetOverContent(i);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.InitialFile), new System.Drawing.Point(497, 67));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(536, 67));
                    }
                    {
                        var pcb = ps.GetOverContent(3);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.InitialFile), new System.Drawing.Point(140, 103));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(280, 103));
                    }
                    {
                        var pcb = ps.GetOverContent(9);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.SignatureFile), new System.Drawing.Point(60, 190));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerSignatureFile), new System.Drawing.Point(60, 155));
                    }

                    ps.Close();
                    pr.Close();
                }

                string output_addendum_pdf = null;
                if (Settings.Offer.ShortSaleAddendum)
                {
                    output_addendum_pdf = d + "\\" + Regex.Replace(address + " SS Addendum.pdf", @"\s+", " ");

                    PdfReader.unethicalreading = true;
                    PdfReader pr = new PdfReader(template_addendum_pdf);
                    PdfStamper ps = new PdfStamper(pr, new FileStream(output_addendum_pdf, FileMode.Create, FileAccess.Write, FileShare.None));

                    //string fs = "";
                    //foreach (KeyValuePair<string, AcroFields.Item> kvp in ps.AcroFields.Fields)
                    //    fs += "\n{\"" + kvp.Key + "\", \"\"},";

                    if (Settings.Parties.BuyerProfile.UseCoBuyer)
                        set_field(ps.AcroFields, "<Buyer Name> and <Co Buyer Name>", Settings.Parties.BuyerProfile.Name + " and " + Settings.Parties.BuyerProfile.CoBuyerName);
                    else
                        set_field(ps.AcroFields, "<Buyer Name> and <Co Buyer Name>", Settings.Parties.BuyerProfile.Name);
                    set_field(ps.AcroFields, "Todays Date", DateTime.Today.ToShortDateString());
                    set_field(ps.AcroFields, "<Address> <UnitNumber> <City/Town> NV <Zip Code>", address + " " + UnitNumber + ", " + City + " NV " + ZipCode);
                    set_field(ps.AcroFields, "Agent Name", Settings.Parties.AgentProfile.Name);
                    //set_field(ps.AcroFields, "Agent Phone", );

                    ps.FormFlattening = true;

                    for (int i = 1; i <= pr.NumberOfPages; i++)
                    {
                        var pcb = ps.GetOverContent(i);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.InitialFile), new System.Drawing.Point(120, 70));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(157, 70));
                    }
                    {
                        var pcb = ps.GetOverContent(1);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.InitialFile), new System.Drawing.Point(140, 400));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerInitialFile), new System.Drawing.Point(197, 400));
                    }
                    {
                        var pcb = ps.GetOverContent(3);
                        add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.SignatureFile), new System.Drawing.Point(150, 547));
                        if (Settings.Parties.BuyerProfile.UseCoBuyer)
                            add_image(pcb, System.Drawing.Image.FromFile(Settings.Parties.BuyerProfile.CoBuyerSignatureFile), new System.Drawing.Point(150, 504));
                    }

                    ps.Close();
                    pr.Close();
                }

                bc.Add(new EmailItem(output_pdf, output_addendum_pdf, output_addendum1_pdf));
            }
            static readonly string template_pdf = Log.AppDir + "\\RPA.pdf";
            static readonly string template_addendum_pdf = Log.AppDir + "\\addendum.pdf";
            static readonly string template_addendum1_pdf = Log.AppDir + "\\addendum1.pdf";

            static void set_field(AcroFields form, string field_key, string value)
            {
                //BaseFont bf = BaseFont.createFont(FONT, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, false, null, null, false);
                if (baseFont != null)
                    form.SetFieldProperty(field_key, "textfont", baseFont, null);
                if (fontSize != null)
                    form.SetFieldProperty(field_key, "textsize", fontSize, null);
                form.SetField(field_key, value);
            }
            static string fontName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "cour.ttf");
            static BaseFont baseFont = BaseFont.CreateFont(fontName, BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
            static System.Single? fontSize = 8.0f;

            string fill_field_by_words(AcroFields form, string field_key, string text)
            {
                IList<AcroFields.FieldPosition> p = form.GetFieldPositions(field_key);
                float width = p[0].position.Width;

                text = FieldPreparation.Prepare(text);
                int end = 0;
                for ( Match m = Regex.Match(text, @".+?([\-\,\:\.]+|(?=\s)|$)"); m.Success; m = m.NextMatch())
                {
                    int e = m.Index + m.Length;
                    float w = baseFont.GetWidthPoint(text.Substring(0, e).Trim(), (float)fontSize);
                    if (w > width)
                        break;
                    end = e;
                }
                set_field(form, field_key, text.Substring(0, end).Trim());
                return text.Substring(end);
            }

            string fill_field_by_chars(AcroFields form, string field_key, string text)
            {
                IList<AcroFields.FieldPosition> p = form.GetFieldPositions(field_key);
                float width = p[0].position.Width;

                text = FieldPreparation.Prepare(text);
                int end = 0;
                for (int e = 1; e <= text.Length; e++)
                {
                    float w = baseFont.GetWidthPoint(text.Substring(0, e).Trim(), (float)fontSize);
                    if (w > width)
                        break;
                    end = e;
                }
                set_field(form, field_key, text.Substring(0, end).Trim());
                return text.Substring(end);
            }

            static void add_image(PdfContentByte pcb, System.Drawing.Image image, System.Drawing.Point point)
            {
                image = ImageRoutines.GetCroppedByColor(image, System.Drawing.Color.Transparent);
                Image i = Image.GetInstance(image, (BaseColor)null);
                var ratio = Math.Min((float)image_max_size.Width / image.Width, (float)image_max_size.Height / image.Height);
                i.ScalePercent(ratio * 100);
                i.SetAbsolutePosition(point.X, point.Y);
                pcb.AddImage(i);
            }
            static System.Drawing.Size image_max_size = new System.Drawing.Size(105, 22);
        }

        class EmailItem : InputItem
        {
            public DataItem Data { get { return (DataItem)__ParentItem; } }
            public readonly string Pdf;
            public readonly string Addendum;
            public readonly string Addendum1;

            internal EmailItem(string pdf, string addendum, string addendum1)
            {
                Pdf = pdf;
                Addendum = addendum;
                Addendum1 = addendum1;
            }

            override public void PROCESSOR(BotCycle bc)
            {
                CustomBot cb = (CustomBot)bc.Bot;

                cb.send(Data, Pdf, Addendum, Addendum1);
            }
        }

        void send(DataItem data, params string[] attachments)
        {
            MailMessage mm = new MailMessage(Settings.Email.EmailServerProfile.SenderEmail, data.ListAgentEmail);
            mm.Subject = Settings.Offer.EmailTemplateProfile.Subject;
            string afn = Regex.Replace(data.ListAgentFullName.Trim(), @"\s.*", "");
            mm.Body = Regex.Replace(Settings.Offer.EmailTemplateProfile.Body, "<AgentFirstName>", afn, RegexOptions.Singleline);
            mm.From = new MailAddress(Settings.Email.EmailServerProfile.SenderEmail);
            foreach (string a in attachments)
            {
                if (a != null)
                    mm.Attachments.Add(new Attachment(a));
            }
            foreach (int i in Settings.Offer.SelectedAttachmentIds)
            {
                mm.Attachments.Add(new Attachment(Settings.Offer.AttachmentFiles[i]));
            }
            Log.Write("Emailing to " + mm.To + ": " + mm.Subject);
            try
            {
#if !DEBUG
                smtp_client.Send(mm);
#endif
            }
            catch (Exception e)
            {
                Log.Write("Host: " + smtp_client.Host
                    + "\r\nPort: " + smtp_client.Port
                    + "\r\nEnableSsl: " + smtp_client.EnableSsl
                    + "\r\nDeliveryMethod: " + smtp_client.DeliveryMethod
                    + "\r\nUserName: " + Settings.Email.EmailServerProfile.SenderEmail
                    + "\r\nSmtpPassword: " + Settings.Email.EmailServerProfile.SmtpPassword
                    + "\r\nFrom: " + mm.From.Address
                    + "\r\nTo: " + mm.To
                    );
                throw new Session.FatalException("SMTP error.", e);
            }
        }

        SmtpClient smtp_client = new SmtpClient
        {
            Host = Settings.Email.EmailServerProfile.SmtpHost,
            Port = Settings.Email.EmailServerProfile.SmtpPort,
            EnableSsl = true,
            DeliveryMethod = SmtpDeliveryMethod.Network,
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(
                 Settings.Email.EmailServerProfile.SenderEmail,
                 Settings.Email.EmailServerProfile.SmtpPassword
                 )
        };
    }
}