# Excel2Object
This library convert an Excel file to an object list and also convert a list to an Excel File

## Usage
- Add nuget package:
```csharp
dotnet add package Excel2Object.Extensions --version 1.0.1
```

- For converting a list to an Excel file:
```csharp
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", 
            "Mild", "Warm", "Balmy", "Hot", "Sweltering", 
            "Scorching"
        };


        [HttpGet(Name = "ExportWeatherAsExcelFile")]
        public IActionResult Get()
        {
            var myList = Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToList();

            // <YourList>.ToExcelFile();
            var fileInfo = myList.ToExcelFile();

            return File(fileInfo.File, fileInfo.ContentType, fileInfo.FileName);
        }

```

- For converting an Excel file to an Object list:

```csharp

        [HttpPost(Name = "ImportWeatherAsExcelFile")]
        public IActionResult Post(IFormFile file)
        {

            var fileStream = file.OpenReadStream();

            // <YourStream>.ToList<YourModel>();
            var objectList = fileStream.ToList<WeatherForecast>();

            return Ok(objectList);
        }
```

 