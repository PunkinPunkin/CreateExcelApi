using System.Text.RegularExpressions;
using CreateExcelApi.Models;
using Microsoft.AspNetCore.Mvc;

namespace CreateExcelApi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public IActionResult Post([FromBody] ExcelInfo excel)
        {
            string fileName = Request.Query.ContainsKey("fileName") ? Regex.Replace(Request.Query["fileName"], @"[\W_]+", string.Empty) : string.Empty;
            if (string.IsNullOrWhiteSpace(fileName) || fileName.Length > 200)
            {
                var st = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                var t = DateTime.Now.ToUniversalTime() - st;
                fileName = $"{t.TotalMilliseconds:0}";
            }

            try
            {
                return File(excel.GetStream().ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"{fileName}.xlsx");
            }
            catch (Exception e)
            {
                _logger.LogError("Exception occurred during generating excel.", e);
                return BadRequest();
            }
            finally
            {
                excel.Dispose();
            }
        }
    }
}