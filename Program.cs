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
        if (string.IsNullOrEmpty(url) || url.StartsWith("http://"))
        {
            continue;
        }

        try
        {
            var request = new HttpRequestMessage(HttpMethod.Head, url);
            var response = await http.SendAsync(request);
            if (response.IsSuccessStatusCode)
            {
                if (response.Content.Headers.ContentType?.MediaType?.StartsWith("video") == true)
                {
                    using (var videoStream = await http.GetStreamAsync(url))
                    using (var videoMemoryStream = new MemoryStream())
                    { 
                        await videoStream.CopyToAsync(videoMemoryStream);
                        videoMemoryStream.Position = 0;

                        slide.Shapes.AddVideo(Convert.ToInt32(shape.X), Convert.ToInt32(shape.Y), videoMemoryStream);
                        var addedVideo = slide.Shapes.Last();
                        addedVideo.Width = shape.Width;
                        addedVideo.Height = shape.Height;
                    }
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing shape with URL {url}: {ex.Message}");
        }
    }
}

pres.Save(outputPath);