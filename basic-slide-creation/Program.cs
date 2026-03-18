using Syncfusion.Presentation;


// Create a new PowerPoint presentation
IPresentation presentation = Presentation.Create();

//Add a new slide to file and apply background color
ISlide slide = presentation.Slides.Add(SlideLayoutType.TitleOnly);

//Specify the fill type and fill color for the slide background 
slide.Background.Fill.FillType = FillType.Solid;
slide.Background.Fill.SolidFill.Color = ColorObject.FromArgb(252, 244, 240);

//Add title content to the slide by accessing the title placeholder of the TitleOnly layout-slide
IShape titleShape = slide.Shapes[0] as IShape;
titleShape.TextBody.AddParagraph("Financial Report \u2014 FY 2024\u20132025").HorizontalAlignment = HorizontalAlignmentType.Center;

//Add description content to the slide by adding a new TextBox
IShape descriptionShape = slide.AddTextBox(53.22, 141.73, 874.19, 77.70);
descriptionShape.TextBody.Text = "This report presents a consolidated view of the company's financial performance across FY 2024–2025. It highlights key trends in revenue, expenses, and growth to support informed strategic decisions.";

FileStream outputStream = new FileStream(Path.GetFullPath(@"Output.pptx"), FileMode.Create);
presentation.Save(outputStream);
//Release all resources of the stream
outputStream.Dispose();
presentation.Close();
