using Dassault.Catia.R24.INFITF;
using Dassault.Catia.R24.MECMOD;
using Dassault.Catia.R24.PARTITF;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Add_bodies
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

            catia.HSOSynchronized = false;
            catia.RefreshDisplay = false;
            catia.DisplayFileAlerts = false;

            var activeDocument = catia.ActiveDocument;

            var partDocument = activeDocument as PartDocument;

            var part = partDocument.Part;

            // Logic

            DeleteEmptyBodies(part, partDocument.Selection);

            Console.WriteLine("Total bodies:" + part.Bodies.Count);

            var inWorkBody = GetInWorkBody(part);

            if (inWorkBody == null)
            {
                Console.WriteLine("Select a body, restart");
                Console.ReadLine();
                return;
            }

            MoveBodiesToInWorkBody(part, inWorkBody);

            Console.WriteLine("Press update");

            //part.Update();

            catia.HSOSynchronized = true;
            catia.RefreshDisplay = true;
            catia.DisplayFileAlerts = true;

            Console.WriteLine("Finished");
            Console.ReadLine();
        }

        private static void DeleteEmptyBodies(Part part, Selection selection)
        {

            var partBodies = part.Bodies as Bodies;

            var countDeletedBodies = 0;

            Console.WriteLine("Deleting empty");
            for (int i = 1; i <= partBodies.Count; i++)
            {
                var tempBody = partBodies.Item(i);

                var y = tempBody.Shapes.Count + tempBody.HybridBodies.Count + tempBody.Sketches.Count;

                if (y == 0 && tempBody != part.MainBody)
                {
                    selection.Add(tempBody);
                    countDeletedBodies += 1;
                }
            }
            selection.Delete();
            Console.WriteLine($"Deleted {countDeletedBodies} bodies");
        }

        private static Body GetInWorkBody(Part part)
        {
            try
            {
                var inWorkObj = part.InWorkObject;

                var tempBodies = inWorkObj as Body;

                var name = tempBodies.get_Name();

                Console.WriteLine(name);

                return part.FindObjectByName(name) as Body;
            }
            catch
            {
                return null;
            }
        }

        private static void MoveBodiesToInWorkBody(Part part, Body inWorkBody)
        {
            var partBodies = part.Bodies as Bodies;

            var bodies = new List<Body>();

            var bar = new ProgressBar(partBodies.Count, 10);

            Console.WriteLine($"Checking {partBodies.Count} bodies");
            
            for (int i = 1; i <= partBodies.Count; i++)
            {
                var tempBody = partBodies.Item(i);
                if (tempBody.get_Name() != inWorkBody.get_Name())
                {
                    bar.Update(i);
                    bodies.Add(tempBody);
                }
            }

            var shapeFactory1 = (ShapeFactory)part.ShapeFactory;

            Console.WriteLine($"Adding {bodies.Count} bodies");
            bar = new ProgressBar(bodies.Count, 10);
            foreach (var body in bodies)
            {
                shapeFactory1.AddNewAssemble(body);
                bar.Update(null);
            }
        }


    }
}
