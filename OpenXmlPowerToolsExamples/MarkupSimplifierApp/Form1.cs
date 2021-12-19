﻿// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
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
            var ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            DialogResult dr = ofd.ShowDialog();

            foreach (string item in ofd.FileNames)
            {
                using (WordprocessingDocument doc =
                       WordprocessingDocument.Open(item, true))
                {
                    var settings = new SimplifyMarkupSettings
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

                    MarkupSimplifier.SimplifyMarkup(doc, settings);
                }
            }
        }
    }
}
