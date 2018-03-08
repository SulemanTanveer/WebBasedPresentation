using System.Drawing.Imaging;
using System.IO;
using Itenso.Rtf;
using Itenso.Rtf.Converter.Html;
using Itenso.Rtf.Converter.Image;
using Itenso.Rtf.Converter.Text;
using Itenso.Rtf.Interpreter;
using Itenso.Rtf.Parser;
using Itenso.Rtf.Support;
using System;

namespace Presentation_Tool
{
    class RTFConverter
    {
        public static string ConvertRtf2Html(string text, int slideID )
        {
            // parser
         text= text.Replace("\\tab","$$$$$$$$");
         text = text.Replace("li240", "$$$$$$$$$$$$$$$$");
         text = text.Replace("li480", "$$$$$$$$$$$");
         text = text.Replace("ri480", "$$$$$$$$$$$");
         text = text.Replace("li720", "$$$$$$$$$$$$$$");
         


            IRtfGroup rtfStructure = ParseRtf(text);
            if (rtfStructure == null)
            {
                return string.Empty;
            }

            // image converter
          
            RtfVisualImageAdapter imageAdapter = new RtfVisualImageAdapter(
                                Path.GetFileNameWithoutExtension(Form1.currentFile)+"\\Slide" + slideID + "_{0}{1}", ImageFormat.Jpeg);
            RtfImageConvertSettings imageConvertSettings = new RtfImageConvertSettings(imageAdapter);
           
            imageConvertSettings.ScaleImage = true; // scale images' 
            
            string filePath = Path.GetDirectoryName(Form1.currentFile);
            string folder = Path.GetFileNameWithoutExtension(Form1.currentFile);
            string directoryPath = filePath + "\\" + folder;
            //string directoryPath = @"C:\Users\Knwal\Documents\Presentation Tool\Presentation Tool\Preview\";
            
            imageConvertSettings.ImagesPath = directoryPath ;
            RtfImageConverter imageConverter = new RtfImageConverter(imageConvertSettings);
           Stream st = GenerateStreamFromString(text);
           RtfInterpreterSettings interpreterSettings = new RtfInterpreterSettings();
           IRtfDocument rtfDocument = RtfInterpreterTool.BuildDoc(rtfStructure, interpreterSettings, null, imageConverter);

            // html converter

            RtfHtmlConvertSettings htmlConvertSettings = new RtfHtmlConvertSettings(imageAdapter);
            htmlConvertSettings.StyleSheetLinks.Add("default.css");
            htmlConvertSettings.ConvertScope = RtfHtmlConvertScope.Content;
            //htmlConvertSettings.GetImageUrl();
            
            RtfHtmlConverter htmlConverter = new RtfHtmlConverter(rtfDocument, htmlConvertSettings);
            string str = htmlConverter.Convert();
            str = str.Replace("  ", "&nbsp;&nbsp;");
          str = str.Replace("$$$$$$$$", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;");
          str = str.Replace("$$$", "&nbsp;&nbsp;&nbsp;");
            //str= str.Replace("#1#","");
            //str = str.Replace("#2#", "");
            //str = str.Replace("#3#", "");
            //str = str.Replace("#4#", "");
            //str = str.Replace("#5#", "");
            //str = str.Replace("#6#", "");
            //str = str.Replace("#7#", "");
            //str = str.Replace("#8#", "");
            //str = str.Replace("#9#", "");
          
          
            return str;
        } // ConvertRtf2Html

        private static IRtfGroup ParseRtf(string text)
        {
            IRtfGroup rtfStructure;
            
            using (Stream stream = GenerateStreamFromString(text))
            {
                RtfParserListenerStructureBuilder structureBuilder = new RtfParserListenerStructureBuilder();
                RtfParser parser = new RtfParser(structureBuilder);
                parser.IgnoreContentAfterRootGroup = true; // support WordPad documents
                
                parser.Parse(new RtfSource(stream));
                rtfStructure = structureBuilder.StructureRoot;
            }
            return rtfStructure;
        } // ParseRtf

        public static Stream GenerateStreamFromString(string s)
        {
           
            MemoryStream stream = new MemoryStream();
            StreamWriter writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }

    }
}
