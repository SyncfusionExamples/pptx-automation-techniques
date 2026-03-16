using Syncfusion.Presentation;

// Open the destination presentation (table.pptx) as the base
IPresentation destinationPresentation = Presentation.Open(new FileStream(Path.GetFullPath(@"Data/table.pptx"), FileMode.Open, FileAccess.Read));

// Open the source presentation (chart.pptx)
IPresentation sourcePresentation = Presentation.Open(new FileStream(Path.GetFullPath(@"Data/chart.pptx"), FileMode.Open, FileAccess.Read));

// Clone and merge all slides from chart presentation to table presentation
for (int i = 0; i < sourcePresentation.Slides.Count; i++)
{
    // Clone the slide from source presentation
    ISlide clonedSlide = sourcePresentation.Slides[i].Clone();

    // Add the cloned slide to destination presentation with destination theme
    destinationPresentation.Slides.Add(clonedSlide, PasteOptions.UseDestinationTheme);
}

// Save the merged presentation
FileStream outputStream = new FileStream(Path.GetFullPath(@"Output.pptx"), FileMode.Create);
destinationPresentation.Save(outputStream);
outputStream.Dispose();

// Close all presentations
sourcePresentation.Close();
destinationPresentation.Close();

