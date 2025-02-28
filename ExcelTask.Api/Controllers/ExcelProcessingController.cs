using ExcelTask.Api.Services.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace ExcelTask.Api.Controllers;

[Route("api/excel")]
[ApiController]
public class ExcelProcessingController(IExcelProcessingService excelProcessingService) : ControllerBase
{
    [HttpPost("upload")]
    public async Task<IActionResult> UploadExcel(IFormFile? recordLayoutFile, IFormFile? controlFigureFile, IFormFile? datasetFile)
    {
        if (recordLayoutFile == null || recordLayoutFile.Length == 0)
        {
            return BadRequest("Record Layout file is missing or empty.");
        }

        if (controlFigureFile == null || controlFigureFile.Length == 0)
        {
            return BadRequest("Control Figure file is missing or empty.");
        }

        if (datasetFile == null || datasetFile.Length == 0)
        {
            return BadRequest("Dataset file is missing or empty.");
        }

        try
        {
            // Save Record Layout File
            string recordLayoutPath = Path.GetTempFileName();
            using (var stream = new FileStream(recordLayoutPath, FileMode.Create))
            {
                await recordLayoutFile.CopyToAsync(stream);
            }

            // Save Control Figure File
            string controlFigurePath = Path.GetTempFileName();
            using (var stream = new FileStream(controlFigurePath, FileMode.Create))
            {
                await controlFigureFile.CopyToAsync(stream);
            }

            // Save Dataset File
            string datasetPath = Path.GetTempFileName();
            using (var stream = new FileStream(datasetPath, FileMode.Create))
            {
                await datasetFile.CopyToAsync(stream);
            }

            // Process all three files
            await excelProcessingService.ProcessExcelFilesAsync(recordLayoutPath, controlFigurePath, datasetPath);

            return Ok("Excel files processed successfully.");
        }
        catch (Exception ex)
        {
            return StatusCode(500, $"An error occurred while processing the files: {ex.Message}");
        }
    }
}