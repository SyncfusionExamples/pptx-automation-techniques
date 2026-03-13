using Syncfusion.Presentation;


// Create a new PowerPoint presentation
IPresentation presentation = Presentation.Create();

// Add a TitleOnly custom layout to the first master slide
ILayoutSlide layoutSlide = presentation.Masters[0].LayoutSlides.Add(SlideLayoutType.TitleOnly, "CustomLayout");

// Set layout background (pale cream)
layoutSlide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(252, 244, 240);

// Add a thin terracotta rule under the title area
layoutSlide.Shapes.AddShape(AutoShapeType.Rectangle, 48f, 120f, 864f, 6f).Fill.SolidFill.Color = ColorObject.FromArgb(215, 100, 67);

// Add one slide using the custom layout
ISlide slide1 = presentation.Slides.Add(layoutSlide);

// Populate the title placeholder
IShape? titleShape = slide1.Shapes[0] as IShape;

//Add a Text Body to the shape
var titleParagraph = titleShape.TextBody.AddParagraph("Financial Report \u2014 FY 2024\u20132025");
titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
titleParagraph.TextParts[0].Font.FontName = "Calibri";
titleParagraph.TextParts[0].Font.FontSize = 48f;
titleParagraph.TextParts[0].Font.Color = ColorObject.FromArgb(16, 66, 96);

// Add a descriptive text box below the title
IShape descriptionShape = slide1.AddTextBox(53.22f, 140f, 874.19f, 120f);
descriptionShape.TextBody.Text = "This report presents a consolidated view of the company's financial performance across FY 2024–2025. It highlights key trends in revenue, expenses, and growth to support informed strategic decisions.";

// Insert an image into the slide
FileStream imageStream = new FileStream("Data/Image.png", FileMode.Open, FileAccess.Read);
slide1.Shapes.AddPicture(imageStream, 450f, 210f, 420f, 300f);

// Save the presentation
using (FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create))
{
    presentation.Save(outputStream);
}
//Release all resources of the stream
outputStream.Dispose();
presentation.Close();
