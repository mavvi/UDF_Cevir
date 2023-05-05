using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO.Compression;
using System.Drawing;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace UdfCevir
{
    public partial class Form1 : Form
    {
        List<Paragraf> paragraf = new List<Paragraf>();
        List<int> paraCharCounts = new List<int>();
        List<int> hizalama = new List<int>();

        string plainText;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            string rtfFilePath = OpenFile();


            richTextBox.LoadFile(rtfFilePath);

            using (StreamReader reader = new StreamReader(rtfFilePath))
            {
                // RTF dosyasını metne dönüştürme
                string rtfText = reader.ReadToEnd();
                System.Windows.Forms.RichTextBox rtBox = new System.Windows.Forms.RichTextBox();
                rtBox.Rtf = rtfText;
                plainText = rtBox.Text.Replace("\t", ""); ;

                // Her paragrafın karakter sayısını hesaplamak için:

                int yaslama = 0;

                foreach (var paraText in rtBox.Text.Replace("\t", "").Split(new string[] { "\n" },
                             StringSplitOptions.None))
                {
                    int charCount = paraText.Length;
                    paraCharCounts.Add(charCount);

                    // Paragraf hizalama kontrolü
                    if (!string.IsNullOrEmpty(paraText))
                    {
                        rtBox.SelectionStart = rtBox.Find(paraText);
                        rtBox.SelectionLength = paraText.Length;

                        switch (rtBox.SelectionAlignment)
                        {
                            case HorizontalAlignment.Left:
                                yaslama = 0;
                                break;
                            case HorizontalAlignment.Right:
                                yaslama = 2;
                                break;
                            case HorizontalAlignment.Center:
                                yaslama = 1;
                                break;
                            default:
                                yaslama = 3;
                                break;
                        }
                    }

                    // Yazı Stilini alma

                    Dictionary<string, bool> styleProps = new Dictionary<string, bool>();
                    for (int i = rtBox.Text.IndexOf(paraText); i < rtBox.Text.IndexOf(paraText) + paraText.Length; i++)
                    {
                        Font font = rtBox.SelectionFont ?? rtBox.Font;
                        styleProps["Bold"] = font.Bold;
                        styleProps["Italic"] = font.Italic;
                        styleProps["Underline"] = rtBox.SelectionFont?.Underline ?? false;
                    }
                    
                    paragraf.Add(new Paragraf { KarakterSayisi = paraText.Length, Alignment = yaslama , Bold = styleProps["Bold"], Italic = styleProps["Italic"], Underline = styleProps["Underline"] });
                }
                
            }
        }


        private string OpenFile()
        {
            // Create an OpenFileDialog to request a file to open.
            OpenFileDialog openFile1 = new OpenFileDialog();

            // Initialize the OpenFileDialog to look for RTF files.
            openFile1.DefaultExt = "*.rtf";
            openFile1.Filter = "RTF Files|*.rtf";

            // Determine whether the user selected a file from the OpenFileDialog.
            if (openFile1.ShowDialog() == System.Windows.Forms.DialogResult.OK &&
                openFile1.FileName.Length > 0)
            {
                // Load the contents of the file into the RichTextBox.
                // richTextBox.LoadFile(openFile1.FileName);

                return openFile1.FileName;
            }

            return null;

            // richTextBox.LoadFile(openFile1.FileName, RichTextBoxStreamType.PlainText);
        }

        public class Paragraf
        {
            public int KarakterSayisi { get; set; }
            public int Alignment { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
            public bool Underline { get; set; }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.OverwritePrompt = true;
            save.CreatePrompt = false;
            save.InitialDirectory = @"%USERPROFILE%\Desktop\";
            save.Title = "UDF Dosyaları";
            save.DefaultExt = "UDF";
            save.Filter = "UDF Dosyaları (*.UDF)|*.UDF|Tüm Dosyalar(*.*)|*.*";
            if (save.ShowDialog() == DialogResult.OK)
            {
                string xmlContent = null;

                xmlContent =
                    "<?xml version=\"1.0\" encoding=\"UTF-8\" ?> \r\n\r\n<template format_id=\"1.8\" >\r\n<content><![CDATA[";
                xmlContent += plainText;
                xmlContent +=
                    "\n]]></content><properties><pageFormat mediaSizeName=\"1\" leftMargin=\"42.525000000000006\" rightMargin=\"42.525000000000006\" topMargin=\"42.525000000000006\" bottomMargin=\"42.52500000000006\" paperOrientation=\"1\" headerFOffset=\"20.0\" footerFOffset=\"20.0\" /></properties>\r\n<elements resolver=\"hvl-default\" >";

                int startOffset = 0;

                //Paragraflar
                foreach (var p in paragraf)
                {
                    xmlContent +=
                        $"<paragraph Alignment=\"{p.Alignment}\"><content ";

                    if (p.Bold)
                        xmlContent += " bold=\"true\"";
                    if (p.Italic)
                        xmlContent += " italic=\"true\"";
                    if (p.Underline)
                        xmlContent += " underline=\"true\"";

                    xmlContent += $" startOffset=\"{startOffset}\" length=\"{p.KarakterSayisi + 1}\" /></paragraph>";
                    startOffset = startOffset + p.KarakterSayisi + 1;
                }


                //Belgeyi Kapat
                xmlContent +=
                    "</elements>\r\n<styles><style name=\"default\" description=\"Geçerli\" family=\"Dialog\" size=\"12\" bold=\"false\" italic=\"false\" FONT_ATTRIBUTE_KEY=\"javax.swing.plaf.FontUIResource[family=Dialog,name=Dialog,style=plain,size=12]\" foreground=\"-13421773\" /><style name=\"hvl-default\" family=\"Times New Roman\" size=\"12\" description=\"Gövde\" /></styles>\r\n</template>";
                // 'udfFilePath' adında bir dosya oluşturun

                // MessageBox.Show(xmlContent);

                string udfFilePath = save.FileName;

                // ZipArchive oluşturun
                using (FileStream zipStream = new FileStream(udfFilePath, FileMode.Create))
                using (ZipArchive archive = new ZipArchive(zipStream, ZipArchiveMode.Create))
                {
                    // content.xml dosyasını ZipArchive'e ekleyin
                    ZipArchiveEntry contentXmlEntry = archive.CreateEntry("content.xml");
                    using (StreamWriter writer = new StreamWriter(contentXmlEntry.Open()))
                    {
                        writer.Write(xmlContent);
                    }
                }

                // dosyayı kaydet
                File.WriteAllBytes(udfFilePath, File.ReadAllBytes(udfFilePath));
            }
        }
    }
}
