using Microsoft.AspNetCore.Mvc;
using Excel2Object.Extensions;

namespace Excel2Object.Example.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

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

            var fileInfo = myList.ToExcelFile();

            return File(fileInfo.File, fileInfo.ContentType, fileInfo.FileName);
        }


        [HttpPost(Name = "ImportWeatherAsExcelFile")]
        public IActionResult Post(IFormFile file)
        {

            var fileInfo = file.OpenReadStream().ToList<WeatherForecast>();

            return Ok(fileInfo);
        }
    }
}