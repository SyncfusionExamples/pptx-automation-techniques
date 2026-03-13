using Syncfusion.Presentation;

// Load the existing PowerPoint presentation
FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/Input.pptx"), FileMode.Open, FileAccess.Read);
IPresentation presentation = Presentation.Open(inputStream);
// Get the first slide
ISlide slide = presentation.Slides[0];
// Find the title shape from the slide
IShape? titleShape = null;
foreach (IShape shape in slide.Shapes)
{
    if (shape.TextBody != null && shape.TextBody.Text.Length > 0)
    {
        titleShape = shape;
        break;
    }
}
// Add fly animation to the title (from left to center)
if (titleShape != null)
{
    ISequence sequence = slide.Timeline.MainSequence;
    IEffect effectLeft = sequence.AddEffect(titleShape, EffectType.Fly, EffectSubtype.Left, EffectTriggerType.OnClick);
}

// Save the PowerPoint presentation with animation
FileStream outputStream = new FileStream(Path.GetFullPath(@"Output/Output.pptx"), FileMode.Create);
presentation.Save(outputStream);

// Dispose streams
inputStream.Dispose();
outputStream.Dispose();

// Close the presentation
presentation.Close();
