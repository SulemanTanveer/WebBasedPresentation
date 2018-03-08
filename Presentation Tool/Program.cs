using Presentation_Tool.DAL;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace Presentation_Tool
{
   [XmlRootAttribute("Presentation", Namespace = "Presentation_Tool",IsNullable = false)]
    public  class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        /// 
        //public Presentation p;

       
        // The XmlArrayAttribute changes the XML element name
        // from the default of "OrderedItems" to "Items".
     


        [STAThread]
        static void Main()
        {
            Presentation p = new Presentation();
            
           // p.createP();
            //p.readP("presentation.xml");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
        
    }
}
