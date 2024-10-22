using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;

namespace Excel2PPT
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {

                // Step 1: Open Excel and load the workbook
                Excel.Application excelApp =new Excel.Application();
                //Console.WriteLine("Digite o caminho do Arquivo da Planilha");
                string excelPath = args[0];// Console.ReadLine();
                Console.WriteLine($@"Confirma {excelPath}");
                Console.ReadLine();
                Excel.Workbook workbook = excelApp.Workbooks.Open(excelPath);
                Excel._Worksheet worksheet = (Excel._Worksheet)workbook.Sheets[1];
                Excel.Range range = worksheet.UsedRange;

                // Step 2: Create a new PowerPoint presentation
                PowerPoint.Application pptApp = new PowerPoint.Application();
                PowerPoint.Presentation presentation = pptApp.Presentations.Add(Office.MsoTriState.msoTrue);

                // Step 3: Loop through the rows of the Excel sheet
                for (int i = 1; i <= range.Rows.Count; i++)
                {
                    // Extract the content of the first column to use as the slide title
                    string slideTitle = Convert.ToString((range.Cells[i, 1] as Excel.Range).Value2);

                    // Add a new slide to the PowerPoint presentation
                    PowerPoint.Slide slide = presentation.Slides.Add(i, PowerPoint.PpSlideLayout.ppLayoutText);
                    slide.Shapes[1].TextFrame.TextRange.Text = slideTitle;

                    // Add content from other columns as the body of the slide
                    for (int j = 2; j <= range.Columns.Count; j++)
                    {
                        string slideContent = Convert.ToString((range.Cells[i, j] as Excel.Range).Value2);
                        slide.Shapes[2].TextFrame.TextRange.Text += slideContent + Environment.NewLine;
                    }
                }

                // Step 4: Save the PowerPoint presentation
                Console.WriteLine("Digite o nome e caminho do PPT");
                string pptPath = args[1];//  Console.ReadLine();
                presentation.SaveAs(pptPath);
                presentation.Close();
                pptApp.Quit();

                // Step 5: Clean up Excel
                workbook.Close(false);
                excelApp.Quit();
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(excelApp);
            }


            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
            
            }
    }
}
