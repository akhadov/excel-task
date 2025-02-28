namespace ExcelTask.Api.Services.Interfaces;

public interface IExcelProcessingService
{
    Task ProcessExcelFilesAsync(string recordLayoutPath, string controlFigurePath, string datasetPath);
}
