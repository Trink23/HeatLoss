using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Reflection;
using System.Linq.Expressions;
using Autodesk.AutoCAD.DatabaseServices;
using System.Windows.Forms;
using Autodesk.AutoCAD.Geometry;

namespace HeatLoss
{
   
    public class CommandsAutocad
    {
        Document activeDocument;

        static HashSet<Point3d> points;

        public CommandsAutocad()
        {
            points = new HashSet<Point3d>();
        }


        
        public double GetPointsFromUserAndReturnArea()
        {

            // Get the current database and start the Transaction Manager
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;



            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            //points = new HashSet<Point3d68; 13

            // Create and add our message filter
            MyMessageFilter filter = new MyMessageFilter();
            System.Windows.Forms.Application.AddMessageFilter(filter);

            while (!filter.bCanceled)
            {
                // Prompt for get points
                pPtOpts.Message = "\nPress Y to finalize the command";
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                PopulatePoint(pPtRes.Value);

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == PromptStatus.Cancel) break;


            }

            points.Remove(points.ElementAt(points.Count - 1));

            System.Windows.Forms.Application.RemoveMessageFilter(filter);

            foreach (Point3d a in points)
            {
                ed.WriteMessage(a.ToString());
            }
            ed.WriteMessage("\n" + getArea(points));

            return getArea(points);
        }

        public double GetPointsFromUserAndReturnLine()
        {

            // Get the current database and start the Transaction Manager
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;

            int noPoints = 0;
            
            


            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            //points = new HashSet<Point3d68; 13

            // Create and add our message filter
            MyMessageFilter filter = new MyMessageFilter();
            System.Windows.Forms.Application.AddMessageFilter(filter);

            while (!filter.bCanceled || noPoints <= 1)
            {
                // Prompt for get points
                pPtOpts.Message = "\nPress Y to finalize the command";
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                PopulatePoint(pPtRes.Value);

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == PromptStatus.Cancel) break;

                noPoints++;
                
                
            }

            points.Remove(points.ElementAt(points.Count - 1));

            System.Windows.Forms.Application.RemoveMessageFilter(filter);

            foreach (Point3d a in points)
            {
                ed.WriteMessage(a.ToString());
            }
            ed.WriteMessage("\n" + getLine(points));

            return getLine(points);
        }


        private static double getLine(HashSet<Point3d> aPoint)
        {
            double resultX = 0;
            double resultY = 0;
            double total = 0;          
            
            
            
            total = aPoint.ElementAt(0).DistanceTo(aPoint.ElementAt(1));      
              
           
                

          
            
            return total;
        }

        private static double getArea(HashSet<Point3d> aPoint)
        {
            double resultX = 0;
            double resultY = 0;
            double total = 0;
            for (int i = 0; i < aPoint.Count; i++)
            {
                if ((i + 2) > aPoint.Count)
                {
                    resultX = resultX + (aPoint.ElementAt(i).X * aPoint.ElementAt(0).Y);
                    resultY = resultY + (aPoint.ElementAt(i).Y * aPoint.ElementAt(0).X);
                }
                else
                {
                    resultX = resultX + (aPoint.ElementAt(i).X * aPoint.ElementAt(i + 1).Y);
                    resultY = resultY + (aPoint.ElementAt(i).Y * aPoint.ElementAt(i + 1).X);
                }
            }

            total = (resultX - resultY) / 2;

            if (total < 0)
            {
                total = total * (-1);
            }

            return total / 100;
        }

        public static void PopulatePoint(Point3d apoint)
        {
            points.Add(apoint);

        }


        public class MyMessageFilter : IMessageFilter

        {

            public const int WM_KEYDOWN = 0x0100;


            public bool bCanceled = false;


            public bool PreFilterMessage(ref Message m)

            {

                if (m.Msg == WM_KEYDOWN)

                {

                    // Check for the Escape keypress

                    Keys kc = (Keys)(int)m.WParam & Keys.KeyCode;

                    if (m.Msg == WM_KEYDOWN && kc == Keys.Y)

                    {

                        bCanceled = true;

                    }

                    // Return true to filter all keypresses

                    return true;

                }

                // Return false to let other messages through

                return false;

            }

        }

    }
}
    

