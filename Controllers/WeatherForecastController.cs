﻿using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;

namespace ClosedXML_Practice.Controllers;

[ApiController]
[Route("[controller]/[Action]")]
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

    [HttpGet]
    public IEnumerable<WeatherForecast> Get()
    {
      return Enumerable.Range(1, 5).Select(index => new WeatherForecast
        {
            Date = DateTime.Now.AddDays(index),
            TemperatureC = Random.Shared.Next(-20, 55),
            Summary = Summaries[Random.Shared.Next(Summaries.Length)]
        })
      .ToArray();
    }

    [HttpGet]
    public FileContentResult Download()
    {
      var wb = BuildExcelFile(10);
      var memoryStream = new MemoryStream();
      wb.SaveAs(memoryStream);
      memoryStream.Seek(0, SeekOrigin.Begin);

      byte[] data = memoryStream.ToArray();
      string fileName = "test.xlsx";
      return File(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
    }

    private XLWorkbook BuildExcelFile(int id)
    {
      //Creating the workbook
      var wb = new XLWorkbook();
      var ws = wb.AddWorksheet("Sheet1");
      ws.FirstCell().SetValue(id);
      return wb;
    }
}
