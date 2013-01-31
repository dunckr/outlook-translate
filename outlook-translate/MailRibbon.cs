using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

using System.Windows.Forms;

namespace outlook_translate
{
    public partial class MailRibbon
    {
        const string key = "";

        const string lang = "en"; // required language

        Outlook.Inspector inspector;
        Word.Document document;

        private void MailRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            inspector = Globals.ThisAddIn.Application.ActiveInspector(); ;
        }

        private void translateButton_Click(object sender, RibbonControlEventArgs e)
        {
            document = (Word.Document)inspector.WordEditor;
            if (document != null)
            {
                string selected = document.Application.Selection.Text;
                if (selected.Length > 0)
                {
                    document.Application.Selection.Text = translate(selected);
                }
            }
        }

        private string translate(string toTranslate)
        {
            string translatedText = toTranslate;
            try
            {
                Translate.LanguageServiceClient client = new Translate.LanguageServiceClient();
                client = new Translate.LanguageServiceClient();
                
                translatedText = client.Translate(key, toTranslate, "", lang);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }
            return translatedText;
        }
    }
}
