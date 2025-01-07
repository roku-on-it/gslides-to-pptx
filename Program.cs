using ShapeCrawler;

if (args.Length < 2)
{
    Console.WriteLine("Usage: dotnet run <input.pptx> <output.pptx>");
    return;
}

string inputPath = args[0];
string outputPath = args[1];

var pres = new Presentation(inputPath);

var http = new HttpClient();

foreach (var slide in pres.Slides)
{
    foreach (var shape in slide.Shapes)
    {
        var url = shape.AltText;
        if (url.Contains("video"))
        {
            using (var httpStream = http.GetStreamAsync(url).Result)
            using (var memoryStream = new MemoryStream())
            {
                httpStream.CopyTo(memoryStream);
                memoryStream.Position = 0;

                slide.Shapes.AddVideo(703, 76, memoryStream);

                var lastShape = slide.Shapes.Last();
                lastShape.Height = 435;
                lastShape.Width = 245;
            }
        }
    }
}

pres.SaveAs(outputPath);