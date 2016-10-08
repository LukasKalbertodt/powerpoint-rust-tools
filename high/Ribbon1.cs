using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Core;

namespace high
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            // PowerPoint.SlideRange r = Globals.ThisAddIn.range;
            PowerPoint.ShapeRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (sr != null)
            {
                foreach(PowerPoint.Shape shape in sr)
                {
                    String s = shape.TextFrame.TextRange.Text;

                    /*string mydocpath =
                        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    using (StreamWriter outputFile = new StreamWriter(mydocpath + @"\WriteLines.txt"))
                    {
                        outputFile.WriteLine(s);
                    }*/

                    Process Program = new Process();
                    Program.StartInfo.FileName = "C:\\Users\\Lukas Kalbertodt\\AppData\\Local\\Programs\\Python\\Python35\\Scripts\\pygmentize.exe";
                    // Program.StartInfo.FileName = "C:\\Users\\Lukas Kalbertodt\\echoargs.exe";
                    Program.StartInfo.Arguments = "-l rust  -f rtf -O style=solarizedlight -O encoding=utf8";
                    // Program.StartInfo.WorkingDirectory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "/Build." + this.name;
                    Program.StartInfo.RedirectStandardOutput = true;
                    Program.StartInfo.RedirectStandardInput = true;
                    Program.StartInfo.RedirectStandardError = true;
                    Program.StartInfo.CreateNoWindow = true;
                    Program.StartInfo.UseShellExecute = false;
                    Program.Start();
                    StreamWriter utf8Writer = new StreamWriter(Program.StandardInput.BaseStream, Encoding.UTF8);
                    utf8Writer.Write(s);
                    utf8Writer.Flush();
                    utf8Writer.Close();
                    String rtfString = Program.StandardOutput.ReadToEnd();
                    Program.WaitForExit(1000);
                    
                    /*using (StreamWriter outputFile = new StreamWriter(mydocpath + @"\WriteLines.txt"))
                    {
                        outputFile.WriteLine(rtfString);
                    }*/

                    String err = Program.StandardError.ReadToEnd();

                    System.Diagnostics.Debug.WriteLine("output: " + rtfString);
                    System.Diagnostics.Debug.WriteLine("err: " + err);

                    // Get the current cliboard content, save it for later and empty the clipboard
                    var backupText = Clipboard.GetText().Clone();
                    Clipboard.Clear();

                    // Prepare and set the new clipboard contents
                    DataObject data = new DataObject();
                    data.SetData(DataFormats.Rtf, rtfString);
                    Clipboard.SetDataObject(data, false);
                    
                    // Paste clipboard contents into textbox
                    if (Clipboard.ContainsText(TextDataFormat.Rtf) && shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        shape.TextFrame.TextRange.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteRTF);
                    }

                    // Restore clipboard content
                    if (backupText != null)
                    {
                        Clipboard.SetText((string)backupText);
                    }
                    

                    // Set useful textbox properties
                    shape.TextEffect.FontName = "Fira Code Retina";
                    shape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                    shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0;
                    shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0;
                    shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3f;
                }
            }
        }
    }
}
