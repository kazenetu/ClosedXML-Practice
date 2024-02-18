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
    public HttpResponseMessage Download()
     {
        var wb = BuildExcelFile(10);
        var memoryStream = new MemoryStream();
        wb.SaveAs(memoryStream);
        memoryStream.Seek(0, SeekOrigin.Begin);

        var message = new HttpResponseMessage(HttpStatusCode.OK)
          {
           Content = new StreamContent(memoryStream)
          };
        var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
        message.Content.Headers.ContentType = new MediaTypeHeaderValue(contentType);
            message.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment")
            {
                FileName = "test.xlsx"
            };

         return message;
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
