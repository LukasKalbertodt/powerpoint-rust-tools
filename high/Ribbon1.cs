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
using System.ComponentModel;

namespace high
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            //MessageBoxButtons buttonx = MessageBoxButtons.OK;
            //MessageBoxIcon iconx = MessageBoxIcon.Error;
            //DialogResult res = MessageBox.Show("hiiii", "caption", buttonx, iconx);

            // PowerPoint.SlideRange r = Globals.ThisAddIn.range;
            PowerPoint.ShapeRange sr = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (sr != null)
            {
                try
                {
                    foreach (PowerPoint.Shape shape in sr)
                    {
                        String s = shape.TextFrame.TextRange.Text;

                        Process Program = new Process();
                        Program.StartInfo.FileName = "pygmentize.exe";
                        String colorScheme;
                        if (this.colorSchemeBox.Text == "")
                        {
                            colorScheme = "";
                        }
                        else
                        {
                            colorScheme = "-O style=" + this.colorSchemeBox.Text;
                        }
                        Program.StartInfo.Arguments = "-l rust -f rtf " + colorScheme + " -O encoding=utf8";
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

                        String err = Program.StandardError.ReadToEnd();

                        System.Diagnostics.Debug.WriteLine("output: " + rtfString);
                        System.Diagnostics.Debug.WriteLine("err: " + err);

                        if (err != "")
                        {
                            throw new Exception("pygmentize errored: " + err);
                        }

                        // Get the current cliboard content, save it for later and empty the clipboard
                        var backupText = (string)Clipboard.GetText().Clone();
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
                        if (backupText != null && backupText != "")
                        {
                            Clipboard.SetText(backupText);
                        }


                        // Set useful textbox properties
                        if (this.fontBox.Text != "")
                        {
                            shape.TextEffect.FontName = this.fontBox.Text;
                        }
                        shape.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNone;
                        shape.TextFrame.TextRange.ParagraphFormat.SpaceAfter = 0;
                        shape.TextFrame.TextRange.ParagraphFormat.SpaceBefore = 0;
                        shape.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.3f;
                        shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    }
                }
                catch (Win32Exception ex)
                {
                    string caption = "Rust Addin: oopsie woopsie, we made a fucky wucky!!";
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBoxIcon icon = MessageBoxIcon.Error;
                    string body = "I think I wasn't able to start 'pygmentize'. Are you sure it's in your PATH? \n" + ex;
                    MessageBox.Show(body, caption, button, icon);
                }
                catch (Exception ex)
                {
                    string caption = "Rust Addin: oopsie woopsie, we made a fucky wucky!!";
                    MessageBoxButtons button = MessageBoxButtons.OK;
                    MessageBoxIcon icon = MessageBoxIcon.Error;
                    MessageBox.Show(ex.ToString(), caption, button, icon);
                }
            }
        }
    }
}
