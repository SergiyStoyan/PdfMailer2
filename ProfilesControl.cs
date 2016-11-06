using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cliver.PdfMailer2
{
    public partial class ProfilesControl : UserControl
    {
        public ProfilesControl()
        {
            InitializeComponent();
        }

        private void bSave_Click(object sender, EventArgs e)
        {
            if (Save != null)
                Save();
        }
        public delegate bool OnSave();
        public OnSave Save = null;

        private void bDelete_Click(object sender, EventArgs e)
        {
            if (Delete != null)
                Delete();
        }
        public delegate bool OnDelete();
        public OnDelete Delete = null;
    }
}
