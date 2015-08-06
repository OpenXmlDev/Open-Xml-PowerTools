using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

namespace MarkupSimplifierApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnApply_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            DialogResult dr = ofd.ShowDialog();
            foreach (var item in ofd.FileNames)
            {
                using (WordprocessingDocument doc =
                    WordprocessingDocument.Open(item, true))
                {
                    SimplifyMarkupSettings settings = new SimplifyMarkupSettings
                    {
                        RemoveContentControls = cbRemoveContentControls.Checked,
                        RemoveSmartTags = cbRemoveSmartTags.Checked,
                        RemoveRsidInfo = cbRemoveRsidInfo.Checked,
                        RemoveComments = cbRemoveComments.Checked,
                        RemoveEndAndFootNotes = cbRemoveEndAndFootNotes.Checked,
                        ReplaceTabsWithSpaces = cbReplaceTabsWithSpaces.Checked,
                        RemoveFieldCodes = cbRemoveFieldCodes.Checked,
                        RemovePermissions = cbRemovePermissions.Checked,
                        RemoveProof = cbRemoveProof.Checked,
                        RemoveSoftHyphens = cbRemoveSoftHyphens.Checked,
                        RemoveLastRenderedPageBreak = cbRemoveLastRenderedPageBreak.Checked,
                        RemoveBookmarks = cbRemoveBookmarks.Checked,
                        RemoveWebHidden = cbRemoveWebHidden.Checked,
                        NormalizeXml = cbNormalize.Checked,
                    };
                    OpenXmlPowerTools.MarkupSimplifier.SimplifyMarkup(doc, settings);
                }
            }
        }
    }
}
