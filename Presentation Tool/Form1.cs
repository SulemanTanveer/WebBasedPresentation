using Presentation_Tool.DAL;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using System.Runtime.InteropServices;
using System.Drawing.Imaging;
using System.Reflection;
using Lyx;
using System.Diagnostics;

namespace Presentation_Tool
{
    public partial class Form1 : Form
    {
        [DllImport("USER32.dll")]
        private static extern Int32 SendMessage(IntPtr hWnd, int msg, int wParam, IntPtr lParam);
        private const int WM_USER = 0x400;
        private const int EM_FORMATRANGE = WM_USER + 57;
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct CHARRANGE
        {
            public int cpMin;
            public int cpMax;
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct FORMATRANGE
        {
            public IntPtr hdc;
            public IntPtr hdcTarget;
            public RECT rc;
            public RECT rcPage;
            public CHARRANGE chrg;
        }
        private const double inch = 14.4;
       
        private Rectangle contentRectangle;
       
        private void RtbToBitmap(RichTextBox rtb,Rectangle rectangle, string filename)
        {
            Bitmap bmp = new Bitmap(rtb.Width,rtb.Height);
            using (Graphics gr = Graphics.FromImage(bmp))
            {
               
                IntPtr hDC = gr.GetHdc();
                FORMATRANGE fmtRange;
                RECT rect;
                int fromAPI;
                rect.Top = 0; rect.Left = 0;
                rect.Bottom = (int)(bmp.Height + (bmp.Height * (bmp.HorizontalResolution / 100)*inch));
                rect.Right = (int)(bmp.Width + (bmp.Width * (bmp.VerticalResolution / 100)) * inch);
                fmtRange.chrg.cpMin = 0;
                fmtRange.chrg.cpMax = -1;
                fmtRange.hdc = hDC;
                fmtRange.hdcTarget = hDC;
                fmtRange.rc = rect;
                fmtRange.rcPage = rect;
                int wParam = 1;
                IntPtr lParam = Marshal.AllocCoTaskMem(Marshal.SizeOf(fmtRange));
                Marshal.StructureToPtr(fmtRange, lParam, false);
                fromAPI = SendMessage(rtb.Handle, EM_FORMATRANGE, wParam, lParam);
                Marshal.FreeCoTaskMem(lParam);
                fromAPI = SendMessage(rtb.Handle, EM_FORMATRANGE, wParam, new IntPtr(0));
                gr.ReleaseHdc(hDC);
            }
            bmp.Save(filename);
            bmp.Dispose();
        }
        private void txtContents_ContentsResized(object sender, ContentsResizedEventArgs e)
        {
            contentRectangle = e.NewRectangle;
        }
      
        public Form1()
        {
            InitializeComponent();
            currentFile = "";
            this.Text = "QAU Web Presentation Tool: New Document";
           // generateImage(selSlide);
            preview.Enabled = false;
        }
        bool change=true;
        private void Form1_Load(object sender, EventArgs e)
        {
            myPresentation = new Presentation();
            currentFile = "";
            this.Text = "QAU Web Presentation Tool: New Document";
            Slide s = new Slide();
            s.ID = myPresentation.Slides.Count + 1;
            s.Heading = "";
            s.Text = "";
            change = false;
            txtContents.Rtf = s.Text;
            txtHeading.Rtf = s.Heading;
            selSlide = s;
            Animation.Checked = false;
            myPresentation.Slides.Add(s);
          
            string thumbnailPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)+"\\thumbnail";
            DirectoryInfo directory = new DirectoryInfo(thumbnailPath);
            deldirectory(directory);
           generateImage(selSlide);
            loadList(true);
            change = true;
           // Animation.Enabled = false;


           
}

        public static   string currentFile;
        FontDialog fontDlg = new FontDialog();
        Presentation myPresentation;
        Slide selSlide;
        ColorDialog colorDlg = new ColorDialog();
        System.Drawing.Font currentFont;
        System.Drawing.FontStyle newFontStyle;
           
        private void txtContents_TextChanged(object sender, EventArgs e)
        {  
            try
            {
                if (change)
                {
                    selSlide.Text = txtContents.Rtf;
                    generateImage(selSlide);
                    loadList(false);
                    change = true;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Insert Slide First");
            }
        }

        private void cutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currControl.Cut();
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currControl.Copy();
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currControl.Paste();
        }

        private void selectAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            currControl.SelectAll();
        }
    
        private void backColor_Click(object sender, EventArgs e)
        {
            //ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                //txtContents.BackColor = colorDlg.Color;
                //backColor.BackColor = colorDlg.Color;
               currControl.SelectionBackColor = colorDlg.Color;
            }
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Text files (*.wpt)|*.wpt";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                preview.Enabled = true;
                string thumbnailPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)+"\\thumbnail";
            DirectoryInfo directory = new DirectoryInfo(thumbnailPath);
           
                try
                {
                    if (txtContents.Modified == true || txtHeading.Modified == true)
                    {
                       //
                        System.Windows.Forms.DialogResult answer;
                        answer = MessageBox.Show("Save this document before closing?", "Unsaved Document", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (answer == System.Windows.Forms.DialogResult.Yes)
                        {
                            SaveToolStripButton_Click(this, e);

                            read(ofd.FileName);
                            currentFile = ofd.FileName;
                            txtContents.Modified = false;
                            txtHeading.Modified = false;
                            this.Text = "QAU Web Presentation Tool " + currentFile.ToString();
                            deldirectory(directory);
                            foreach (Slide s in myPresentation.Slides)
                            {
                                change = false;
                                txtHeading.Rtf = s.Heading;
                                change = false;
                                txtContents.Rtf = s.Text;
                                generateImage(s);
                            }
                           loadList(true);
                           change = true;
                            //tSlides.SelectedNode = tSlides.SelectedNode.FirstNode;
                            //Form1_Load(this, e);
                        }
                        else if (answer == System.Windows.Forms.DialogResult.No)
                        {
                            read(ofd.FileName);
                    
                            currentFile = ofd.FileName;
                            txtContents.Modified = false;
                            txtHeading.Modified = false;
                            this.Text = "QAU Web Presentation Tool " + currentFile.ToString();
                            deldirectory(directory);
                            foreach (Slide s in myPresentation.Slides)
                            {
                                change = false;
                                txtHeading.Rtf = s.Heading;
                                change = false;
                                txtContents.Rtf = s.Text;
                                generateImage(s);
                            }
                            loadList(true);
                            change = true;
                        }
                        else
                        {

                        }
                    }
                    else
                    {
                        
                        read(ofd.FileName);
                        currentFile = ofd.FileName;
                        txtContents.Modified = false;
                        txtHeading.Modified = false;
                        this.Text = "QAU Web Presentation Tool " + currentFile.ToString();
                        
                        foreach (Slide s in myPresentation.Slides)
                        {
                            change = false;
                            txtHeading.Rtf = s.Heading;
                            change = false;
                            txtContents.Rtf = s.Text;
                            generateImage(s);
                        }
                       loadList(true);
                       change = true;
                        // Form1_Load(this, e);
                        // Application.Exit();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Error");
                }
            }

        }
        
        //private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        //{
        //   // Application.Exit();
        //}

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
                {
                    if (currentFile == string.Empty)
                    {
                        saveAs();
                    }
                    else
                    {
                        SavePresentation(currentFile);
                        preview.Enabled = true;
                    }
         }
            catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Error");
                }
        }

        public void SavePresentation(string filename)
        {
            XmlSerializer serialize = new XmlSerializer(typeof(Presentation));
            TextWriter writer = null;
            try
            {
                writer = new StreamWriter(filename);
                serialize.Serialize(writer, myPresentation);
                //writer.Close();
            }
            finally
            {
                if (writer != null)
                    writer.Close();
            }
            
        } 
       
        
    
        private void read(string filename)
        {
            XmlSerializer sr = new XmlSerializer(typeof(Presentation));
            using (var stream = File.Open(filename, FileMode.Open))
            {
                myPresentation = (Presentation)sr.Deserialize(stream);
            }
            //FileStream stream = new FileStream(filename, FileMode.Open);
           
            
            
        }
      

        private void txtContents_SelectionChanged(object sender, EventArgs e)
        {

            txtContents.SelectionFont = new System.Drawing.Font(fontDialog1.Font.Name, fontDialog1.Font.SizeInPoints);
            //txtContents.SelectionFont = tempFont;
            
            BoldToolStripButton.Checked = txtContents.SelectionFont.Bold;
            UnderlineToolStripButton.Checked = txtContents.SelectionFont.Underline;
            ItalicToolStripButton.Checked = txtContents.SelectionFont.Italic;
            LeftToolStripButton.Checked = txtContents.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Left ? true : false;
            CenterToolStripButton.Checked = txtContents.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Center ? true : false;
            RightToolStripButton.Checked = txtContents.SelectionAlignment == System.Windows.Forms.HorizontalAlignment.Right ? true : false;
            BulletsToolStripButton.Checked = txtContents.SelectionBullet;

        }
    
        private void u_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (currControl == null)
            {
                currControl = txtContents;
            }

            if (e.ClickedItem.Name == "BoldToolStripButton")
            {
                currentFont = currControl.SelectionFont;
                if (currControl.SelectionFont.Bold == true)
                {
                    newFontStyle = (int)currentFont.Style - System.Drawing.FontStyle.Bold;
                }
                else
                {
                    newFontStyle = (int)currentFont.Style + System.Drawing.FontStyle.Bold; 
                }
                currControl.SelectionFont = new Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
                
            }

            if (e.ClickedItem.Name == "UnderlineToolStripButton")
            {
                currentFont = txtContents.SelectionFont;
                newFontStyle = default(System.Drawing.FontStyle);
                if (currControl.SelectionFont.Underline == true)
                {
                    newFontStyle = (int) currentFont.Style - System.Drawing.FontStyle.Underline;
                }
                else
                {
                    newFontStyle = (int)currentFont.Style + System.Drawing.FontStyle.Underline;
                }

               currControl.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
             
               
            }

         
            if (e.ClickedItem.Name == "ItalicToolStripButton")
            {
                currentFont = txtContents.SelectionFont;
                newFontStyle = default(System.Drawing.FontStyle);
                if (currControl.SelectionFont.Italic == true)
                {
                    newFontStyle = (int)currentFont.Style - System.Drawing.FontStyle.Italic;
                }
                else
                {
                    newFontStyle = (int)currentFont.Style + System.Drawing.FontStyle.Italic;
                }
                currControl.SelectionFont = new System.Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle);
               
              }
          
            if (e.ClickedItem.Name == "RightToolStripButton")
            {
                RightToolStripButton_Click(this, e);
               
            }
            if (e.ClickedItem.Name == "LeftToolStripButton")
            {
                LeftToolStripButton_Click(this, e);

            }
            if (e.ClickedItem.Name == "CenterToolStripButton")
            {
                CenterToolStripButton_Click(this, e);
            }
            if (e.ClickedItem.Name == "BulletsToolStripButton")
            {
                try
                {
                    currControl.SelectionBullet = !currControl.SelectionBullet;
                   // BulletsToolStripButton.Checked = currControl.SelectionBullet;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString(), "Error");
                }
                
            }
            if (e.ClickedItem.Name =="UppercaseToolStripButton")
            {
                currControl.SelectedText = currControl.SelectedText.ToUpper();
               
            }
            if (e.ClickedItem.Name == "LowercaseToolStripButton")
            {
                currControl.SelectedRtf = currControl.SelectedRtf.ToLower();
               
            }
            if (e.ClickedItem.Name == "lIndentToolStripButton")
            {
                lIndentToolStripButton_Click(this,e);
            }
            if (e.ClickedItem.Name == "rIndentToolStripButton")
            {
                rIndentToolStripButton_Click(this, e);
            }
         
        }

        private void LeftToolStripButton_Click(object sender, EventArgs e)
        {
           currControl.SelectionAlignment = HorizontalAlignment.Left;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
        string font;
        private void FontToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(currControl.SelectionFont == null))
                {
                    fontDialog1.Font = currControl.SelectionFont;
                    font = fontDialog1.Font.Name;
                }
                else
                {
                    fontDialog1.Font = null;
                }
                fontDialog1.ShowApply = true;
                if (fontDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    currControl.SelectionFont = fontDialog1.Font;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }
        
        private void BackToolStripButton_Click(object sender, EventArgs e)
        {

            try
            {
                backColor_Click(this, e);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void FontColorToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                //colorDlg.Color = txtContents.ForeColor;
                if (colorDlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    currControl.SelectionColor = colorDlg.Color;
                   
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void ImageToolStripButton_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "RTE - Insert Image File";
            openFileDialog1.DefaultExt = "rtf";
            openFileDialog1.Filter = "JPEG Files|*.jpg|GIF Files|*.gif|PNG Files|*.png|Bitmap Files|*.bmp";
            //openFileDialog1.FilterIndex = 1;
            openFileDialog1.ShowDialog();

            if (openFileDialog1.FileName == "")
            {
                return;
            }

            try
            {
                string strImagePath = openFileDialog1.FileName;
                Image img;
                img = Image.FromFile(strImagePath);
                Clipboard.SetDataObject(img);
                DataFormats.Format df;
                df = DataFormats.GetFormat(DataFormats.Bitmap);
                if (this.txtContents.CanPaste(df))
                {
                    this.txtContents.Paste(df);
                }
            }
            catch
            {
                MessageBox.Show("Unable to insert image format selected.", "RTE - Paste", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void rIndentToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                currControl.SelectionIndent += 8;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void lIndentToolStripButton_Click(object sender, EventArgs e)
        {
            try
            {
                currControl.SelectionIndent -= 8;
               
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void top_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            //if (e.ClickedItem.Name == "NewToolStripButton")
            //{
            //    newToolStripMenuItem_Click(this, e);
            //}

            //if(e.ClickedItem.Name=="OpenToolStripButton")
            //{
            //    openToolStripMenuItem_Click(this, e);
            //}

            //if (e.ClickedItem.Name == "SaveToolStripButton")
            //{
            //    saveToolStripMenuItem_Click(this, e);
            //}
        }

        private void txtHeading_TextChanged(object sender, EventArgs e)
        {
            if (change)
            {
                selSlide.Heading = txtHeading.Rtf;
                generateImage(selSlide);
                loadList(false);
                change = true;
                //loadList();
            }
        }
        bool flag= false;

        private void exitToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            try
            {
                if (txtContents.Modified == true || txtHeading.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save this document before closing?", "Unsaved Document", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.Yes)
                    {
                        saveToolStripMenuItem_Click(this, e);
                        flag = true;
                        Application.Exit();
                    }
                    else if (answer == System.Windows.Forms.DialogResult.No)
                    {
                        txtContents.Modified = false;
                        txtHeading.Modified = false;
                        flag = true;
                        Application.Exit();
                    }
                    else
                    {

                    }
                }
                else
                {
                    txtContents.Modified = false;
                    txtHeading.Modified = false;
                    flag = true;
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void newToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                //string thumbnailPath = Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\thumbnail";
                //DirectoryInfo directory = new DirectoryInfo(thumbnailPath);
           
                if (txtContents.Modified == true || txtHeading.Modified == true)
                {
                    System.Windows.Forms.DialogResult answer;
                    answer = MessageBox.Show("Save this document before closing?", "Unsaved Document", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.Yes)
                    {
                        SaveToolStripButton_Click(this, e);
                 //       deldirectory(directory);
                        Form1_Load(this, e);
                    }
                    else if(answer == System.Windows.Forms.DialogResult.No)
                    {
                        txtContents.Modified = false;
                        txtHeading.Modified = false;
                   //     deldirectory(directory);
                        Form1_Load(this, e);

                    }
                    else
                    {

                    }

                }
                else
                {
                    txtContents.Modified = false;
                    txtHeading.Modified = false;
                    //deldirectory(directory);
                    Form1_Load(this, e);

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString(), "Error");
            }
        }

        private void SaveToolStripButton_Click(object sender, EventArgs e)
        {
            saveToolStripMenuItem_Click(this, e);
        }

       
        private void saveAs()
        {
            SaveFileDialog svd = new SaveFileDialog();
            svd.Filter = "Text Files(*.wpt)|*.wpt";
            if (svd.ShowDialog() == DialogResult.OK)
            {
               SavePresentation(svd.FileName);
               currentFile = svd.FileName;
               this.Text = "QAU Web Presentation Tool: " + currentFile.ToString();
               txtContents.Modified = false;
               txtHeading.Modified = false;
               MessageBox.Show(currentFile.ToString() + " saved.", "File Save");
               preview.Enabled = true;
            }
        }
        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveAs();
            string pattern = @"<li><div class='header'><p>{0}</p> </div>
				<div class='content'>	<p>{1} </p></div></li>";
            string html = string.Empty;
            foreach (var slide in myPresentation.Slides)
            {
                html += String.Format(pattern, slide.HHeading, slide.HText);
            }
            string pAtH = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string htmlDoc = File.ReadAllText("index.html");
            htmlDoc = htmlDoc.Replace("##RTF##", html);
            string file = Path.GetFileNameWithoutExtension(currentFile) + ".html";
            if (File.Exists(currentFile))
            {
                filePath = Path.GetDirectoryName(currentFile);
            }
            string folder = Path.GetFileNameWithoutExtension(currentFile);
            if (!Directory.Exists(currentFile))
            {
                System.IO.Directory.CreateDirectory(filePath + "\\" + folder);
            }
            string a = filePath + "\\" + folder;
            System.IO.Directory.CreateDirectory(a + "\\js");
            System.IO.Directory.CreateDirectory(a + "\\img");
            System.IO.Directory.CreateDirectory(a + "\\css");
            File.WriteAllText(@filePath + "\\" + folder + "\\" + file, htmlDoc);
            copyDirectory(pAtH + "\\js", a + "\\js");
            copyDirectory(pAtH + "\\img", a + "\\img");
            copyDirectory(pAtH + "\\css", a + "\\css");
        }
        string filePath;
      //  bool copy = false;
        private void preview_Click(object sender, EventArgs e)
        {
           // SavePresentation(currentFile);
            string pattern = @"<li><div class='header'><p>{0}</p> </div>
				<div class='content'>	<p>{1} </p></div></li>";
            string html = string.Empty;
            foreach (var slide in myPresentation.Slides)
            {
                html += String.Format(pattern, slide.HHeading, slide.HText);
            }
            string pAtH = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location);
            string htmlDoc = File.ReadAllText("index.html");
            htmlDoc = htmlDoc.Replace("##RTF##", html);
            string file = Path.GetFileNameWithoutExtension(currentFile) + ".html";
            if (File.Exists(currentFile))
            {
                filePath = Path.GetDirectoryName(currentFile);
            }
            string folder = Path.GetFileNameWithoutExtension(currentFile);
            string a = filePath + "\\" + folder;
            
            if (!Directory.Exists(a))
            {
                System.IO.Directory.CreateDirectory(filePath + "\\" + folder);
                System.IO.Directory.CreateDirectory(a + "\\js");
                System.IO.Directory.CreateDirectory(a + "\\img");
                System.IO.Directory.CreateDirectory(a + "\\css");
                copyDirectory(pAtH + "\\js", a + "\\js");
                copyDirectory(pAtH + "\\img", a + "\\img");
                copyDirectory(pAtH + "\\css", a + "\\css");
            }
           
          
            File.WriteAllText(@filePath + "\\" + folder + "\\" + file, htmlDoc);
            System.Diagnostics.Process.Start(a+"\\"+file);
        }

        public static void copyDirectory(string Src, string Dst)
        {
            String[] Files;

            if (Dst[Dst.Length - 1] != Path.DirectorySeparatorChar)
                Dst += Path.DirectorySeparatorChar;
            if (!Directory.Exists(Dst)) Directory.CreateDirectory(Dst);
            Files = Directory.GetFileSystemEntries(Src);
            foreach (string Element in Files)
            {
                // Sub directories
                if (Directory.Exists(Element))
                    copyDirectory(Element, Dst + Path.GetFileName(Element));
                // Files in directory
                else
                    File.Copy(Element, Dst + Path.GetFileName(Element), true);
            }
        }
       
         protected override void OnFormClosing(FormClosingEventArgs e)
         {
             if (flag == false)
                 exitToolStripMenuItem_Click_1(this, e);
             else
                 Application.Exit();
         }

         private void charactergrowfont_Click(object sender, EventArgs e)
         {
             float NewFontSize = currControl.SelectionFont.SizeInPoints + 2;

             Font NewSize = new Font(currControl.SelectionFont.Name, NewFontSize, currControl.SelectionFont.Style);

             currControl.SelectionFont = NewSize;
         }

         private void charactershrinkfont_Click(object sender, EventArgs e)
         {
             float NewFontSize = currControl.SelectionFont.SizeInPoints - 2;

             Font NewSize = new Font(currControl.SelectionFont.Name, NewFontSize, currControl.SelectionFont.Style);

             currControl.SelectionFont = NewSize;
         }

         private void addSlide_Click(object sender, System.EventArgs e)
         {
            // string f = currControl.Font.Name;
             Slide s = new Slide();
             s.ID = myPresentation.Slides.Count+1;
             selSlide = s;
             myPresentation.Slides.Add(s);
             
             selSlide.Heading = "";
             selSlide.Text = "";
             change = false;
             txtHeading.Rtf = selSlide.Heading;
             txtContents.Rtf = selSlide.Text;
             change = true;
             generateImage(selSlide);
             //imageListView.Focus[s.ID-1]= true;
         loadList(false);
        
          }
        /* private void LoadTree()
         {
             TreeNode pn = new TreeNode("Presentation");
             var sn = new TreeNode();
             foreach (var s in myPresentation.Slides)
             {
                 sn = new TreeNode(s.ID.ToString());
                 pn.Nodes.Add(sn);
             }

             tSlides.ContextMenuStrip = contextMenuStrip1;
             foreach (TreeNode ChildNode in tSlides.Nodes)
             {
                 ChildNode.ContextMenuStrip = contextMenuStrip1;
             }
             
             tSlides.Nodes.Clear();
             tSlides.Nodes.Add(pn);
             tSlides.ExpandAll();
             tSlides.SelectedNode = sn;
         }*/
         private static ImageList imageList = new ImageList();
        
        //void loadList(bool reload)
        // {
        //     string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thumbnail\\";
        //     if (reload)
        //     {
        //         this.imageListView.Items.Clear();
        //         imageList.Images.Clear();
        //         foreach (Slide s in myPresentation.Slides)
        //         {
        //             try
        //             {

        //                 string imgpath = folder + s.ID + ".jpg";

        //                 if (File.Exists(imgpath))
        //                 {

        //                     Image img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
        //                     imageList.Images.Add(img);
        //                 }
        //                 else
        //                 {
        //                     imgpath = folder + "empty.jpg";
        //                     Image img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
        //                     imageList.Images.Add(img);
        //                 }
        //                 //imageListView.FocusedItem.Selected = true;
        //                 //imageListView.Select();
        //             }
        //             catch
        //             {
        //                 Console.WriteLine("This is not an image file");
        //             }
        //         }
        //         imageList.ImageSize = new Size(100, 75);

        //         this.imageListView.View = View.LargeIcon;

        //         int counter = 0;
        //         foreach (Slide s in myPresentation.Slides)
        //         {

        //             ListViewItem item = new ListViewItem { ImageIndex = counter, Text = (counter + 1).ToString(), Name = s.ID.ToString() };
        //             item.ImageIndex = counter;

        //             this.imageListView.Items.Add(item);
        //             counter++;
        //             imageListView.EnsureVisible(item.Index);
        //             item.Selected = true;
        //             item.Focused = true;
        //         }
        //         this.imageListView.LargeImageList = imageList;
        //         imageListView.ContextMenuStrip = contextMenuStrip1;
        //     }
        //     else
        //     {
        //         imageList.ImageSize = new Size(100, 75);

        //         this.imageListView.View = View.LargeIcon;

        //         string imgpath = folder + selSlide.ID + ".jpg";
        //         ListViewItem selitem = null;
        //         if (imageListView.SelectedItems.Count <= 1)
        //         {
        //             foreach (ListViewItem item in imageListView.Items)
        //             {
        //                 if (item.Name == selSlide.ID.ToString())
        //                 {
        //                     selitem = item;
        //                     break;

        //                 }
        //             }
        //             //if (selitem == null) // new slide case
        //             //{
        //             //    Image image = null;
        //             //    if (File.Exists(imgpath))
        //             //    {
        //             //        image = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
        //             //        imageList.Images.Add(image);

        //             //    }
        //             //    else
        //             //    {
        //             //        imgpath = folder + "empty.jpg";
        //             //        image = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
        //             //        imageList.Images.Add(image);

        //             //    }
        //             //    selitem = new ListViewItem {
        //             //        ImageIndex = imageList.Images.Count+1,
        //             //        Text = (imageList.Images.Count).ToString(),
        //             //        Name = selSlide.ID.ToString()
        //             //    };
        //             //    selitem.ImageIndex = imageList.Images.Count;
        //             //    this.imageListView.Items.Add(selitem);
        //             //    imageListView.Refresh();
        //             //    imageListView.EnsureVisible(selitem.Index);
        //             //    selitem.Selected = true;
        //             //    selitem.Focused = true;
        //             //    this.imageListView.LargeImageList = imageList;
        //             //    imageListView.ContextMenuStrip = contextMenuStrip1;
                         
        //             //}
                     
                     
        //         }
        //         else
        //         {
        //             selitem = imageListView.SelectedItems[0];
        //         }

        //         Image img;
        //         if (File.Exists(imgpath))
        //         {

        //             img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
        //         }
        //         else
        //         {
        //             imgpath = folder + "empty.jpg";
        //             img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));

        //         }
        //         imageList.Images[selitem.ImageIndex] = img;

        //         imageListView.Refresh();
        //         imageListView.EnsureVisible(selitem.ImageIndex);
        //         selitem.Selected = true;
        //         selitem.Focused = true;


                     
        //     }

        //     }

         void loadList(bool reload)
         {
             string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thumbnail\\";
             if (reload)
             {
                 this.imageListView.Items.Clear();
                 imageList.Images.Clear();
                 foreach (Slide s in myPresentation.Slides)
                 {
                     try
                     {

                         string imgpath = folder + s.ID + ".jpg";

                         if (File.Exists(imgpath))
                         {

                             Image img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
                             imageList.Images.Add(img);
                         }
                         else
                         {
                             imgpath = folder + "empty.jpg";
                             Image img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
                             imageList.Images.Add(img);
                         }
                         //imageListView.FocusedItem.Selected = true;
                         //imageListView.Select();
                     }
                     catch
                     {
                         Console.WriteLine("This is not an image file");
                     }
                 }
                 imageList.ImageSize = new Size(100, 75);

                 this.imageListView.View = View.LargeIcon;

                 int counter = 0;
                 foreach (Slide s in myPresentation.Slides)
                 {

                     ListViewItem item = new ListViewItem { ImageIndex = counter, Text = (counter + 1).ToString(), Name = s.ID.ToString() };
                     item.ImageIndex = counter;

                     this.imageListView.Items.Add(item);
                     counter++;
                     imageListView.EnsureVisible(item.Index);
                     item.Selected = true;
                     item.Focused = true;
                 }
                 this.imageListView.LargeImageList = imageList;
                 imageListView.ContextMenuStrip = contextMenuStrip1;
             }
             else
             {
                 imageList.ImageSize = new Size(100, 75);

                 this.imageListView.View = View.LargeIcon;

                 string imgpath = folder + selSlide.ID + ".jpg";
                 ListViewItem selitem = null;
                 if (imageListView.SelectedItems.Count <= 1)
                 {
                     foreach (ListViewItem item in imageListView.Items)
                     {
                         if (item.Name == selSlide.ID.ToString())
                         {
                             selitem = item;
                             break;

                         }
                     }
                     if (selitem == null) // new slide case
                     {
                         Image image = null;
                         if (File.Exists(imgpath))
                         {
                             image = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
                             imageList.Images.Add(image);

                         }
                         else
                         {
                             imgpath = folder + "empty.jpg";
                             image = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
                             imageList.Images.Add(image);

                         }
                         selitem = new ListViewItem
                         {
                             Text = (imageList.Images.Count).ToString(),
                             Name = selSlide.ID.ToString()
                         };
                         selitem.ImageIndex = imageList.Images.Count - 1;
                         this.imageListView.Items.Add(selitem);
                         imageListView.Refresh();
                         imageListView.EnsureVisible(selitem.Index);
                         selitem.Selected = true;
                         selitem.Focused = true;
                         this.imageListView.LargeImageList = imageList;
                         imageListView.ContextMenuStrip = contextMenuStrip1;

                     }
                     else
                     {
                         Image img;
                         if (File.Exists(imgpath))
                         {

                             img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));
                         }
                         else
                         {
                             imgpath = folder + "empty.jpg";
                             img = Image.FromStream(new MemoryStream(File.ReadAllBytes(imgpath)));

                         }
                         imageList.Images[selitem.ImageIndex] = img;

                         imageListView.Refresh();
                         imageListView.EnsureVisible(selitem.ImageIndex);
                         selitem.Selected = true;
                         selitem.Focused = true;


                     }

                 }
                 else
                 {
                     selitem = imageListView.SelectedItems[0];
                 }
             }

         }

         private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
         {
             ListViewItem l = new ListViewItem();
             l = imageListView.SelectedItems[0];
             
             imageListView.Items.Remove(l);
             string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thumbnail\\";
            

             DirectoryInfo d = new DirectoryInfo(folder);

             FileInfo[] infos = d.GetFiles("*.jpg");
             foreach (FileInfo f in infos)
             {
                 string g = l.Text.ToString() + ".jpg";
                 string fi = f.Name;
                 if (fi == g)
                 {
                     f.Delete();
                     break;
                  }
             }
             int cont = Int32.Parse(l.Text.ToString());
             myPresentation.Slides.RemoveAt(cont-1);
            //tSlides.SelectedNode.Remove();
             reArrange();
            // LoadTree();
         loadList(false);
            }

        void reArrange()
         {
             for (var s = 1; s <= myPresentation.Slides.Count; s++)
             {
                 myPresentation.Slides[s - 1].ID = s;
             }
             string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thumbnail\\";
             DirectoryInfo d = new DirectoryInfo(folder);
             FileInfo[] infos = d.GetFiles("*.jpg");
            
             List<FileInfo> temp= new List<FileInfo>();
            foreach (FileInfo f in infos)
            {
                if (f.Name != "empty.jpg")
                    temp.Add(f);
            }


             FileInfo[] sortedFiles = temp.OrderBy(r =>int.Parse(Path.GetFileNameWithoutExtension(r.Name))).ToArray();
            int k = 1;
          
           foreach (FileInfo f in sortedFiles)
            {
                var fi = f.Name;
                fi = Path.GetFileNameWithoutExtension(fi);
                if (fi != k.ToString() && fi!="empty")
                {
                    File.Move(f.FullName, folder + "\\" + k.ToString() + ".jpg");
                   // break;
                }
                if (k <= imageListView.Items.Count)
                k++;
             }
            }
            
         private void NewToolStripButton_Click(object sender, EventArgs e)
         {
             newToolStripMenuItem_Click(this, e);
         }

         private void OpenToolStripButton_Click(object sender, EventArgs e)
         {
             openToolStripMenuItem_Click(this, e);
         }

        
         private void txtContents_AcceptsTabChanged(object sender, EventArgs e)
         {
             currControl.Text = "        ";
         }

         private void Animation_Click(object sender, EventArgs e)
         {
             if (Animation.Checked == true)
                 selSlide.animated = false;
             else
                 selSlide.animated = true;
         }

       public void generateImage(Slide slide)
         {
             string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\thumbnail\\";
           
             string img1 = folder + "tmp_head_" + slide.ID + ".jpg";
             string img2 = folder + "temp_content_" + slide.ID + ".jpg";
             File.Delete(img1);
             File.Delete(img2);
             RtbToBitmap(txtHeading,contentRectangle, img1);
             RtbToBitmap(txtContents, contentRectangle, img2);

             CombineImages(new FileInfo(img1), new FileInfo(img2), slide.ID.ToString());
         }
    
        void deldirectory( DirectoryInfo d)
         {
             if (d != null)
             {
                 FileInfo[] files = d.GetFiles("*.jpg");
                 foreach (FileInfo file in files)
                 {
                     if(file.Name!="empty.jpg")
                     file.Delete();
                 }
             }
         }
        
         private void CombineImages(FileInfo file1, FileInfo file2, string id  )
         {
             //change the location to store the final image.

             List<FileInfo> files = new List<FileInfo>();
             files.Add(file1);
             files.Add(file2);

             string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                 System.IO.Directory.CreateDirectory(folder+"\\thumbnail");
             
             List<int> imageWidths = new List <int> ();
             int nIndex = 0;
             int height = 0;
             foreach (FileInfo file in files)
             {
                 Image img = Image.FromFile(file.FullName);
                 imageWidths.Add(img.Width);
                 height += img.Height;
                 img.Dispose();
             }
             imageWidths.Sort();
             int width = imageWidths[(imageWidths.Count-1)];
             Bitmap img3 = new Bitmap(width, height);
             Graphics g = Graphics.FromImage(img3);
             g.Clear(SystemColors.Window);
             
             foreach (FileInfo file in files)
             {
                 Image img = Image.FromFile(file.FullName);
                 if (nIndex == 0)
                 {
                     g.DrawImage(img, new Point(0, 0));
                     nIndex++;
                     height = img.Height;
                 }
                 else
                 {
                     g.DrawImage(img, new Point(0,height));
                     height += img.Height;
                 }
                img.Dispose();
             }
                 string finalImage = folder + "\\thumbnail\\" + id + ".jpg";
                 File.Delete(finalImage);
                 img3.Save(finalImage, ImageFormat.Jpeg);
                 g.Dispose();
                 img3.Dispose();
                 Bitmap img4 = new Bitmap(CreateThumbnail(finalImage, 100, 100));
                 img4.Save(finalImage, System.Drawing.Imaging.ImageFormat.Jpeg);
                 img4.Dispose();
                 file1.Delete();
                 file2.Delete();
         }
   
          RichTextBox currControl;
         private void txtContents_Enter(object sender, EventArgs e)
         {
             currControl = sender as RichTextBox;
          }

         private void txtHeading_Enter(object sender, EventArgs e)
         {
             currControl = sender as RichTextBox;
         }

         private void CenterToolStripButton_Click(object sender, EventArgs e)
         {
             currControl.SelectionAlignment = HorizontalAlignment.Center;
         }

         private void RightToolStripButton_Click(object sender, EventArgs e)
         {
             currControl.SelectionAlignment = HorizontalAlignment.Right;
         }

         public static Bitmap CreateThumbnail(String lcFilename, int lnWidth, int lnHeight)
         {
             System.Drawing.Bitmap bmpOut = null;
             try
             {
                 Bitmap loBMP = new Bitmap(lcFilename);
                 ImageFormat loFormat = loBMP.RawFormat;

                 decimal lnRatio;
                 int lnNewWidth = 0;
                 int lnNewHeight = 0;

                 //*** If the image is smaller than a thumbnail just return it
                 if (loBMP.Width < lnWidth && loBMP.Height < lnHeight)
                     return loBMP;

                 if (loBMP.Width > loBMP.Height)
                 {
                     lnRatio = (decimal)lnWidth / loBMP.Width;
                     lnNewWidth = lnWidth;
                     decimal lnTemp = loBMP.Height * lnRatio;
                     lnNewHeight = (int)lnTemp;
                 }
                 else
                 {
                     lnRatio = (decimal)lnHeight / loBMP.Height;
                     lnNewHeight = lnHeight;
                     decimal lnTemp = loBMP.Width * lnRatio;
                     lnNewWidth = (int)lnTemp;
                 }
                 bmpOut = new Bitmap(lnNewWidth, lnNewHeight);
                 Graphics g = Graphics.FromImage(bmpOut);
                 
                 g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                 g.FillRectangle(Brushes.White, 0, 0, lnNewWidth, lnNewHeight);
                 g.DrawImage(loBMP, 0, 0, lnNewWidth, lnNewHeight);
                 ControlPaint.DrawBorder(g, new Rectangle(0, 0, lnNewWidth, lnNewHeight), Color.Black, ButtonBorderStyle.Solid);
                 loBMP.Dispose();
             }
             catch
             {
                 return null;
             }

             return bmpOut;
         }

         private void imageListView_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
         {
             if (imageListView.SelectedItems.Count > 0)
             {
                 ListViewItem item = imageListView.SelectedItems[0];
                 foreach (var s in myPresentation.Slides)
                 {
                    if (s.ID.ToString() == item.Name)
                         {
                             selSlide = s;
                             change = false;
                             txtContents.Rtf = s.Text;
                             txtHeading.Rtf = s.Heading;
                             Animation.Checked = s.animated;
                             break;
                         }
                 }
              change = true;
             }
         }

       

         

        } 
}


