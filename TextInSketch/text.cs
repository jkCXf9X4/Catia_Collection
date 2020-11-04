using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using System.IO;
using System.Runtime.InteropServices;

using Catia =  Dassault.Catia.R24;

using Dassault.Catia.R24.INFITF;
using Dassault.Catia.R24.MECMOD;
using Dassault.Catia.R24.DRAFTINGITF;


namespace InsertText
{
    public class Text
    {
        public Text(string text, double size = 6)
        {
            Application catia;
            try
            {
                catia = (Application)Marshal.GetActiveObject("Catia.Application");
            }
            catch
            {
                Debug.WriteLine("Start catia");
                throw new ArgumentNullException();
            }

            catia.HSOSynchronized = false;
            catia.RefreshDisplay = false;
            catia.DisplayFileAlerts = false;

            var activeDocument = catia.ActiveDocument;

            var DrwDoc = catia.Documents.Add("Drawing") as DrawingDocument;

            var DrwText = DrwDoc.Sheets.Item(1).Views.Item(1).Texts.Add(text, 0, 0) as DrawingText;

            DrwText.SetFontSize(0, 0, size);

            DrwText.SetFontName(0, 0, "Monospac821 BT");

            string tempFolderPath = Path.GetTempPath() + "Text" + DateTime.Now.ToLongDateString() + ".ig2";

            DrwDoc.ExportData(tempFolderPath, "ig2");
            DrwDoc.Close();

            var Ig2Doc = catia.Documents.Open(tempFolderPath) as Document;

            var Ig2Selection = catia.ActiveDocument.Selection as Selection;

            Ig2Selection.Search("Drafting.Polyline,all");

            Ig2Selection.Copy();

            activeDocument.Activate();

            catia.HSOSynchronized = true;
            catia.RefreshDisplay = true;
            catia.DisplayFileAlerts = true;
        }
    }
}
