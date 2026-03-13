using Syncfusion.Presentation;

class Program
{
    static void Main()
    {
        using (FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read))
        using (IPresentation presentation = Presentation.Open(inputStream))
        {
            // Add title
            var titleBox = presentation.Slides[0].Shapes.AddTextBox(48f, 30f, 864f, 80f);
            var titleParagraph = titleBox.TextBody.AddParagraph("Financial Operations Lifecycle");
            titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
            titleParagraph.TextParts[0].Font.FontName = "Calibri";
            titleParagraph.TextParts[0].Font.FontSize = 36f;
            titleParagraph.TextParts[0].Font.Color = ColorObject.FromArgb(16, 66, 96);

            // Read workflow items
            var workflowItems = File.ReadAllLines(Path.GetFullPath(@"Data/workflow.txt"))
                .Select(line => line.Trim())
                .Where(line => !string.IsNullOrEmpty(line))
                .ToList();

            // Add SmartArt with workflow data
            ISmartArt smartArt = presentation.Slides[0].Shapes.AddSmartArt(SmartArtType.BasicCycle, 100, 150, 750, 350);

            for (int i = 0; i < Math.Min(smartArt.Nodes.Count, workflowItems.Count); i++)
                smartArt.Nodes[i].TextBody.AddParagraph(workflowItems[i]);

            // Save the presentation
            using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create))
            {
                presentation.Save(outputStream);
            }
            presentation.Close();
        }
    }
}
