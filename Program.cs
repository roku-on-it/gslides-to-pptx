using System.IO.Compression;
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
var images = new List<byte[]>();

foreach (var slide in pres.Slides)
{
    var shapesToProcess = new List<(IShape Shape, string Url)>();

    for (int i = 0; i < slide.Shapes.Count; i++)
    {
        var shape = slide.Shapes[i];
        var url = shape.AltText;
        if (!string.IsNullOrEmpty(url) && !url.StartsWith("http://"))
        {
            var imgBytes = ((shape as IPicture)!).Image!.AsByteArray();
            images.Add(imgBytes);
            shapesToProcess.Add((shape, url));
        }
    }

    foreach (var item in shapesToProcess)
    {
        var shape = item.Shape;
        var url = item.Url;

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

            shape.Remove();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing shape with URL {url}: {ex.Message}");
        }
    }
}

pres.Save(outputPath);

// Releasing resources
pres.Dispose();

using var zip = ZipFile.Open(outputPath, ZipArchiveMode.Update);

foreach (var entry in zip.Entries)
{
    if (entry.FullName.StartsWith("ppt/media/") && entry.FullName.EndsWith(".png") && entry.Crc32 == 859486013)
    {
        await using (var stream = entry.Open())
        {
            stream.SetLength(0);
            stream.Write(images[0], 0, images[0].Length);
        }

        images.RemoveAt(0);
    }
}