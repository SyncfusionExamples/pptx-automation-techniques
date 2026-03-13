using Syncfusion.Licensing;
using Syncfusion.Presentation;
using System.Diagnostics;

namespace Create_Table_with_Dynamic_Data
{
    class Program
    {
        static void Main(string[] args)
        {
            //Register Syncfusion license
            SyncfusionLicenseProvider.RegisterLicense("Ngo9BigBOggjHTQxAR8/V1JGaF1cXmhNYVBpR2NbeU54flVPallWVAciSV9jS3hTdUVnWXdfcHVXQmNZUk91XQ==");

            //Load the existing PowerPoint presentation from Data folder
            using FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read);
            IPresentation presentation = Presentation.Open(inputStream);

            //Inject financial data from Excel into the presentation slides
            ExcelToPresentationHelper helper = new ExcelToPresentationHelper();
            helper.AddFinancialDataToPresentation(presentation);

            //Save the presentation to disk and release resources
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create))
            {
                presentation.Save(outputStream);
            }
            presentation.Close();

        }
    }
}
