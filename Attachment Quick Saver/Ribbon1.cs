using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;

namespace Attachment_Quick_Saver
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void saveAttachmentAsFile(Attachment attachment, string f)
        {
            string directory = System.IO.Path.GetDirectoryName(f); // full directory
            string filename = System.IO.Path.GetFileNameWithoutExtension(f);
            string extension = System.IO.Path.GetExtension(f); //.ext
            
            int i = 1;
            while (System.IO.File.Exists(f))
            {
                string newfilename = string.Format("{0}-{1}{2}", filename, i, extension);
                f = System.IO.Path.Combine(directory, newfilename);
                i++;
            }

            try
            {
                attachment.SaveAsFile(f);
            }
            catch (System.Exception e)
            {
                MessageBox.Show("Error saving attachment to " + f + Environment.NewLine + e.Message);
            }
        }

        private void saveAttachmentsToDirectory(RibbonControlEventArgs e, string directory)
        {
            var explorer = (e.Control.Context as Explorer);
            var selection = explorer.Selection;
            if(selection.Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selObject = selection.Application.ActiveExplorer().Selection[1];
                if(selObject is Microsoft.Office.Interop.Outlook.MailItem)
                {
                    Microsoft.Office.Interop.Outlook.MailItem mailitem = (selObject as Microsoft.Office.Interop.Outlook.MailItem);
                    if (mailitem.Attachments.Count > 0)
                    {
                        foreach (Attachment attachment in mailitem.Attachments)
                        {
                            saveAttachmentAsFile(attachment, directory + mailitem.SentOn.ToString("yyyy-MM-dd") + ".pdf");
                        }
                    }
                }
            }
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            saveAttachmentsToDirectory(e, Properties.Settings.Default.button1_path);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            saveAttachmentsToDirectory(e, Properties.Settings.Default.button2_path);
        }
    }
}
