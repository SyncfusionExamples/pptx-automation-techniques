using Syncfusion.Drawing;
using Syncfusion.Presentation;

// Load the existing PowerPoint presentation
FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read);
IPresentation presentation = Presentation.Open(inputStream);

// Add title text box
var titleBox = presentation.Slides[0].Shapes.AddTextBox(48f, 30f, 864f, 80f);
var titleParagraph = titleBox.TextBody.AddParagraph("Financial Year Profit Visuals — FY 2024–2025");
titleParagraph.HorizontalAlignment = HorizontalAlignmentType.Center;
titleParagraph.TextParts[0].Font.FontName = "Calibri";
titleParagraph.TextParts[0].Font.FontSize = 36f;
titleParagraph.TextParts[0].Font.Color = ColorObject.FromArgb(16, 66, 96);

// Add chart from Excel (A1:B13 = Month, Profit in Lakhs)
FileStream excelStream = new FileStream(Path.GetFullPath(@"Data/Book1.xlsx"), FileMode.Open);
IPresentationChart chart = presentation.Slides[0].Charts.AddChart(excelStream, 1, "A1:B13", new RectangleF(90, 150, 800, 380));
chart.ChartTitle = "Financial Year Profit Visuals";
chart.PrimaryCategoryAxis.Title = "Month";
chart.PrimaryValueAxis.Title = "Profit (Lakhs)";
chart.HasLegend = false;
excelStream.Dispose();
inputStream.Dispose();

// Save the PowerPoint Presentation
FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create);
presentation.Save(outputStream);
outputStream.Dispose();
presentation.Close();
