using Dassault.Catia.R24.INFITF;
using Dassault.Catia.R24.DRAFTINGITF;

using System.Diagnostics;
using System.Runtime.InteropServices;

namespace RenameInstance
{
    class Program
    {
        static void Main(string[] args)
        {
            Application catia;

            try
            {
                catia = (Application)Marshal.GetActiveObject("Catia.Application");
            }
            catch
            {
                Debug.WriteLine("Start catia");
                return;
            }

            try
            {
                catia.HSOSynchronized = false;
                catia.RefreshDisplay = false;
                catia.DisplayFileAlerts = false;
            }
            catch
            {
                Debug.WriteLine("Connection error");
                return;
            }

            var activeDocument = catia.ActiveDocument;

            DrawingDocument drawDoc;

            try
            {
                drawDoc = (DrawingDocument)activeDocument;
            }
            catch
            {

                Debug.WriteLine("Not a drawing, exiting");
                return;
            }

            DrawingSheet sheet = drawDoc.Sheets.ActiveSheet;

            var Selection = catia.ActiveDocument.Selection as Selection;

            Selection.Search("Drafting.'Area Fill',sel");

            Selection.Delete();


            catia.HSOSynchronized = true;
            catia.RefreshDisplay = true;
            catia.DisplayFileAlerts = true;

            catia = null;
            
        }
    }
}

