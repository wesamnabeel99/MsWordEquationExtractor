using DocumentFormat.OpenXml.ExtendedProperties;
using Microsoft.Office.Interop.Word;
using System;
using Application = Microsoft.Office.Interop.Word.Application;

namespace PageSizeAdjustment
{
    public class DocxEditor
    {
        private Application app = new Application();
        private Document doc;

        public DocxEditor(String inputFilePath)
        {
            doc = app.Documents.Open(inputFilePath);
        }

        public void deleteParagraphs()
        {
            app.Visible = true;

            int paraCount = doc.Paragraphs.Count;
            int deleted = 0;
            int equationsNumber = 0;
            int failed = 0;

            foreach (Paragraph paragraph in doc.Paragraphs)
            {
                OMaths equations = paragraph.Range.OMaths;

                if (equations.Count == 0)
                {
                    try
                    {
                        deleted++;
                        paragraph.Range.Delete();
                        paragraph.TextboxTightWrap = WdTextboxTightWrap.wdTightAll;


                    }
                    catch (Exception e)
                    {
                        failed++;
                    }
                }
                else
                {
                    equationsNumber++;
                }


            }

            removeMargins(doc);

            Console.Beep();
            Console.WriteLine("deleted " + deleted + " paragraphs and failed to delete " + failed + "\nkept " + equationsNumber + " equations" + "\ntotal number of paragraphs in file is:" + paraCount);

        }

        public void SaveDocument()
        {
            doc.SaveAs2(FileName: "NewFile", FileFormat: WdExportFormat.wdExportFormatPDF);

        }

        public void CloseDocument()
        {
            doc.Close();
        }

        public void CloseApp()
        {
            app.Quit();
        }
        public void removeMargins(Document document)
        {
            document.PageSetup.LeftMargin = 0;
            document.PageSetup.RightMargin = 0;
            document.PageSetup.TopMargin = 0;
            document.PageSetup.BottomMargin = 0;
        }


        public void findAndReplace(String find, String replace)
        {
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Text = find;

            app.Selection.Find.Replacement.ClearFormatting();
            app.Selection.Find.Replacement.Text = replace;
            app.Selection.Find.Execute(Replace: WdReplace.wdReplaceAll);
        }

    }
}