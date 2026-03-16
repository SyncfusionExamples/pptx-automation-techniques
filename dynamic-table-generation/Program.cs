
using Syncfusion.Presentation;


namespace dynamic_table_generation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load the existing PowerPoint presentation from Data folder
            using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read);
            IPresentation presentation = Presentation.Open(inputStream);

            //Inject financial data from Excel into the presentation slides
            ExcelToPresentationHelper helper = new ExcelToPresentationHelper();
            helper.AddFinancialDataToPresentation(presentation);

            //Save the presentation to disk and release resources
            FileStream outputStream = new FileStream(Path.GetFullPath(@"Output.pptx"), FileMode.Create);
            presentation.Save(outputStream);
            outputStream.Dispose();

            presentation.Close();

        }
    }
}
