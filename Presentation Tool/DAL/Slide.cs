
using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Serialization;
namespace Presentation_Tool.DAL
{

    public class Slide
    {
    [XmlElement]
        public int ID { get; set; }
    public bool animated { get; set; }

        public string Heading { get; set; }
        public string Text { get; set; }

        [XmlIgnore]
        public string HText { get {
           // return MarkupConverter.RtfToHtmlConverter.ConvertRtfToHtml(Text);
            string htm = RTFConverter.ConvertRtf2Html(Text,ID);
            if (animated)
            {
                HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                doc.LoadHtml("<tool>" + htm + "</tool>");

                var litems = doc.DocumentNode.Descendants("li");
                foreach (HtmlNode item in litems)
                {
                    item.Attributes.Add("class", "tool-class");
                    string s = item.InnerText;
                    if (s.StartsWith("#"))
                    {
                        int ei = s.IndexOf("#", 1);
                        if (ei > 0)
                        {
                            string sorder = s.Substring(1, ei-1);
                            int order;
                            if (int.TryParse(sorder, out order))
                            {
                            item.Attributes.Add("data-order", order.ToString());
                           // item.Attributes.Add("data-slide", ID.ToString());
                            
                            }
                        }
                    }
                }

                htm = doc.DocumentNode.FirstChild.InnerHtml;
            }
           htm = htm.Replace("#1#", "");
           htm = htm.Replace("#2#", "");
           htm = htm.Replace("#3#", "");
           htm = htm.Replace("#4#", "");
           htm = htm.Replace("#4#", "");
           htm = htm.Replace("#5#", "");
           htm = htm.Replace("#6#", "");
           htm = htm.Replace("#7#", "");
           htm = htm.Replace("#8#", "");
           htm = htm.Replace("#9#", "");
            return htm;
            //return RTFConverter.ConvertRtf2Html(Text, ID);
        }
           
        }
        [XmlIgnore]
        public string HHeading
        {
            get
            {
             //   return MarkupConverter.RtfToHtmlConverter.ConvertRtfToHtml(Heading);
                return RTFConverter.ConvertRtf2Html(Heading,ID);
            }
        }
    }
}
