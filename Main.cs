using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System.Windows.Forms;


namespace HeatLoss
{
    public static class Heatloss
    {
        static UserInterface aUserInterface;

        [CommandMethod("Heatloss")]
        public static void Heatloss1()
        {
            Editor ed = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor;
            
            System.Windows.Forms.Application.EnableVisualStyles();

            aUserInterface = new UserInterface();

            System.Windows.Forms.Application.Run(aUserInterface);

           
        }

        public static UserInterface GetUserInterface()
        {
            return aUserInterface;
        }
        
    }
   

}

