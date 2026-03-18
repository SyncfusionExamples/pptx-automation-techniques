using Syncfusion.Presentation;

// Create a new PowerPoint presentation
using (IPresentation presentation = Presentation.Create())
{
    // Add a TitleOnly custom layout to the first master slide
    ILayoutSlide layoutSlide = presentation.Masters[0].LayoutSlides.Add(SlideLayoutType.TitleOnly, "CustomLayout");

    // Set layout background (pale cream) so all slides using this layout inherit it
    layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(252, 244, 240);

    // Add a thin terracotta rule under the title area
    layoutSlide.Shapes.AddShape(AutoShapeType.Rectangle, 48f, 120f, 864f, 6f).Fill.SolidFill.Color = ColorObject.FromArgb(215, 100, 67);

    // Add one slide using the custom layout
    ISlide slide1 = presentation.Slides.Add(layoutSlide);

    // Populate the title placeholder and apply basic formatting
    IShape? titleShape = slide1.Shapes[0] as IShape;
    var titleParagraph = titleShape.TextBody.AddParagraph("Financial Report \u2014 FY 2024\u20132025");
    titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
    titleParagraph.TextParts[0].Font.FontName = "Calibri";
    titleParagraph.TextParts[0].Font.FontSize = 48f;
    titleParagraph.TextParts[0].Font.Color = ColorObject.FromArgb(16, 66, 96);

    // Add a descriptive text box below the title
    IShape descriptionShape = slide1.AddTextBox(50.22f, 140f, 874.19f, 120f);
    descriptionShape.TextBody.Text = "This report presents a consolidated view of the company's financial performance across FY 2024–2025. It highlights key trends in revenue, expenses, and growth to support informed strategic decisions.";

    //Add image into the slides
    FileStream pictureStream = new FileStream("data/Image.png", FileMode.Open);
    slide1.Shapes.AddPicture(pictureStream, 450, 210, 420, 300);

    //Add second slide using the same custom layout - it automatically inherits the background from layout slide
    ISlide slide2 = presentation.Slides.Add(layoutSlide);
    ISlide slide3 = presentation.Slides.Add(layoutSlide);
    ISlide slide4 = presentation.Slides.Add(layoutSlide);

    // Save the presentation
    FileStream outputStream = new FileStream(Path.GetFullPath(@"Output.pptx"), FileMode.Create);
    presentation.Save(outputStream);
    outputStream.Dispose();
    // Close the presentation
    presentation.Close();

}


