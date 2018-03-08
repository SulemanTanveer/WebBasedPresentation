using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
using System.IO;

namespace Presentation_Tool.DAL
{
   [XmlRoot ("Presentation")]
    public class Presentation
    {
         [XmlElement]
//         [XmlArrayAttribute("Slide")]
         public string Name { get; set; }
        public string BGColor { get; set; }
        public string Theme { get; set; }

       
        public List<Slide> Slides { get; set; }
        //public string filename = "Presentation";
        
        public Presentation()
        {
            Slides = new List<Slide>();
        }
       
       
       
    }
}
