using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Configuration;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;

namespace Cliver.PdfMailer2
{
    public partial class Settings
    {
        public static readonly OfferClass Offer;

        public class OfferClass : Cliver.Settings
        {
            public Dictionary<string, EmailTemplateProfile> EmailTemplateProfileNames2EmailTemplateProfileProfile = new Dictionary<string, EmailTemplateProfile>();
            public string EmailTemplateProfileName;
            public DateTime CloseOfEscrow = DateTime.Now;
            public string Emd;
            public bool ShortSaleAddendum;
            public bool OtherAddendum1;
            public bool OtherAddendum2;
            public string[] AttachmentFiles = new string[0];
            public int[] SelectedAttachmentIds = new int[0];
            [ScriptIgnore]
            public EmailTemplateProfile EmailTemplateProfile
            {
                get
                {
                    return EmailTemplateProfileNames2EmailTemplateProfileProfile[EmailTemplateProfileName];
                }
            }
        }

        public class EmailTemplateProfile : Profile
        {
            public string Subject;
            public string Body;
        }
    }
}