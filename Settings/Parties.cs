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
        public static readonly PartiesClass Parties;

        public class PartiesClass : Cliver.Settings
        {
            public Dictionary<string, PartyProfile> PartyProfileNames2PartyProfile = new Dictionary<string, PartyProfile>();
            public Dictionary<string, BuyerProfile> BuyerProfileNames2BuyerProfile = new Dictionary<string, BuyerProfile>();
            public Dictionary<string, BrokerProfile> BrokerProfileNames2BrokerProfile = new Dictionary<string, BrokerProfile>();
            public Dictionary<string, AgentProfile> AgentProfileNames2AgentProfile = new Dictionary<string, AgentProfile>();
            public Dictionary<string, EscrowProfile> EscrowProfileNames2EscrowProfile = new Dictionary<string, EscrowProfile>();
            public string PartyProfileName;
            [ScriptIgnore]
            public PartyProfile PartyProfile
            {
                get
                {
                    return PartyProfileNames2PartyProfile[PartyProfileName];
                }
            }
            [ScriptIgnore]
            public BuyerProfile BuyerProfile
            {
                get
                {
                    return BuyerProfileNames2BuyerProfile[PartyProfile.BuyerProfileName];
                }
            }
            [ScriptIgnore]
            public BrokerProfile BrokerProfile
            {
                get
                {
                    return BrokerProfileNames2BrokerProfile[PartyProfile.BrokerProfileName];
                }
            }
            [ScriptIgnore]
            public AgentProfile AgentProfile
            {
                get
                {
                    return AgentProfileNames2AgentProfile[PartyProfile.AgentProfileName];
                }
            }
            [ScriptIgnore]
            public EscrowProfile EscrowProfile
            {
                get
                {
                    return EscrowProfileNames2EscrowProfile[PartyProfile.EscrowProfileName];
                }
            }
        }

        public abstract class Profile
        {
            public string _ProfileName;
        }

        public class PartyProfile : Profile
        {
            public string BuyerProfileName;
            public string BrokerProfileName;
            public string AgentProfileName;
            public string EscrowProfileName;
        }

        public class BuyerProfile : Profile
        {
            public string Name;
            public string InitialFile;
            public string SignatureFile;
            public bool UseCoBuyer;
            public string CoBuyerName;
            public string CoBuyerInitialFile;
            public string CoBuyerSignatureFile;
        }

        public class AgentProfile : Profile
        {
            public string Name;
            public string LicenseNo;
            public string Email;
            public string InitialFile;
            public string SignatureFile;
        }

        public class BrokerProfile : Profile
        {
            public string Name;
            public string LicenseNo;
            public string Company;
            public string Address;
            public string City;
            public string State;
            public string Zip;
            public string Phone;
        }

        public class EscrowProfile : Profile
        {
            public string TitleCompany;
            public string Officer;
        }
    }
}