using Syncfusion.Presentation;

// Load the existing PowerPoint presentation
FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read);
IPresentation presentation = Presentation.Open(inputStream);

// Add title text box with centered heading
var titleBox = presentation.Slides[0].Shapes.AddTextBox(48f, 30f, 864f, 80f);
var titleParagraph = titleBox.TextBody.AddParagraph("Key Strategies to Improve Business Profitability");
titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
titleParagraph.TextParts[0].Font.FontName = "Calibri";
titleParagraph.TextParts[0].Font.FontSize = 36f;
titleParagraph.TextParts[0].Font.Color = ColorObject.FromArgb(16, 66, 96);

// Add picture to slide
FileStream pictureStream = new FileStream(Path.GetFullPath(@"Data/Image.png"), FileMode.Open);
IPicture picture = presentation.Slides[0].Pictures.AddPicture(pictureStream, 150, 150, 650, 350);

// Save the PowerPoint Presentation
FileStream outputStream = new FileStream(Path.GetFullPath(@"Output.pptx"), FileMode.Create);
presentation.Save(outputStream);

// Dispose the image stream
inputStream.Dispose();
pictureStream.Dispose();
outputStream.Dispose();

// Closes the Presentation
presentation.Close();
