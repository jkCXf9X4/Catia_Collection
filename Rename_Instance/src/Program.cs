using Dassault.Catia.R24.INFITF;
using Dassault.Catia.R24.MECMOD;
using Dassault.Catia.R24.PARTITF;
using Dassault.Catia.R24.ProductStructureTypeLib;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace RenameInstance
{
    class Program
    {
        static void Main(string[] args)
        {
            // Setup
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

            // catia.HSOSynchronized = false;
            catia.RefreshDisplay = false;
            catia.DisplayFileAlerts = false;
            // catia.Interactive = false;

            if(catia.Windows.Count == 0)
            {
                Console.WriteLine("No window is open");
            }

            Document activeDocument = (Document)catia.ActiveDocument;

            if(activeDocument.TypeName() == "ProductDocument")
            {
                var productDocument = (ProductDocument)activeDocument;
                Console.WriteLine("Its a productDocument");

                var product = (Product)productDocument.Product;

                product.ApplyWorkMode(CatWorkModeType.DEFAULT_MODE);

                var rootProducts = (Products)product.Products;

                Rename(rootProducts);
            }
        }

        static void Rename(Products products)
        {
            int count = products.Count;

            Rename(products.AsList(), count+1);

            Rename(products.AsList(), 1);
        }

        static void Rename(List<Product> products, int start)
        {
            foreach (var product in products)
            {
                var strPNum = product.get_PartNumber();
                var strPName = product.get_Name();

                product.set_Name( (start++).ToString());
            }
        }
    }

    public static class Misc
    {
        public static string TypeName(this AnyObject i)
        {
            return Microsoft.VisualBasic.Information.TypeName(i);
        }
        public static string TypeName(this Collection i)
        {
            return Microsoft.VisualBasic.Information.TypeName(i);
        }

        public static List<Product> AsList(this Products products)
        {
            List<Product> list = new List<Product>();
            for( int i = 1; i <= products.Count; i++)
            {
                list.Add(products.Item(i));
            }
            return list;
        }
    }
}

