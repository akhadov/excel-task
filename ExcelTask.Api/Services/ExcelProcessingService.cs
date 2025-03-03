using Dapper;
using ExcelDataReader;
using ExcelTask.Api.Data;
using ExcelTask.Api.Services.Interfaces;
using System.Data;
using System.Data.Common;

namespace ExcelTask.Api.Services;

public class ExcelProcessingService(IDbConnectionFactory dbConnectionFactory) : IExcelProcessingService
{
    public async Task ProcessExcelFilesAsync(string recordLayoutPath, string controlFigurePath, string datasetPath)
    {
        await ProcessRecordLayoutAsync(recordLayoutPath);
        await ProcessControlFigureAsync(controlFigurePath);
        await ProcessDatasetAsync(datasetPath);
    }

    private async Task ProcessRecordLayoutAsync(string recordLayoutPath)
    {
        await ProcessExcelFileAsync(recordLayoutPath, "record_layout");
    }

    private async Task ProcessControlFigureAsync(string controlFigurePath)
    {
        await ProcessExcelFileAsync(controlFigurePath, "control_figure");
    }

    private async Task ProcessDatasetAsync(string datasetPath)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(datasetPath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read()) return;

        var datasetColumns = new List<string>();

        for (int i = 0; i < reader.FieldCount; i++)
        {
            string columnName = reader.GetString(i)?.Trim();
            if (!string.IsNullOrWhiteSpace(columnName))
            {
                datasetColumns.Add(SanitizeColumnName(columnName));
            }
        }

        int actualDatasetRowCount = 0;
        while (reader.Read())
        {
            bool isEmptyRow = true;

            for (int i = 0; i < reader.FieldCount; i++)
            {
                object cellValue = reader.GetValue(i);
                if (cellValue != null && cellValue != DBNull.Value && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    isEmptyRow = false;
                    break;
                }
            }

            if (!isEmptyRow)
            {
                actualDatasetRowCount++;
            }
        }

        if (!await ValidateDataset(datasetColumns, actualDatasetRowCount))
        {
            throw new Exception("Dataset validation failed: Column count or row count does not match control_figure.");
        }

        stream.Position = 0;
        using var newReader = ExcelReaderFactory.CreateReader(stream);
        newReader.Read();

        var dataRows = new List<Dictionary<string, object>>();

        while (newReader.Read())
        {
            var values = new Dictionary<string, object>();
            bool isEmptyRow = true;

            for (int i = 0; i < datasetColumns.Count; i++)
            {
                object cellValue = newReader.GetValue(i);
                if (cellValue != null && cellValue != DBNull.Value && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    isEmptyRow = false;
                }

                values[datasetColumns[i]] = cellValue is DBNull ? null : cellValue;
            }

            if (!isEmptyRow)
            {
                dataRows.Add(values);
            }
        }

        if (dataRows.Count > 0)
        {
            await CreateTableAsync("dataset", datasetColumns);
            await InsertDataAsync("dataset", datasetColumns, dataRows);
        }
    }


    private async Task<bool> ValidateDataset(List<string> datasetColumns, int actualDatasetRowCount)
    {
        await using DbConnection connection = await dbConnectionFactory.OpenConnectionAsync();

        var recordLayoutColumns = (await connection.QueryAsync<string>("SELECT Column_Name FROM record_layout")).ToList();
        recordLayoutColumns = recordLayoutColumns.Select(SanitizeColumnName).ToList();

        var controlFigureData = await connection.QueryAsync<(string Column_Name, double Total)>(
            "SELECT Column_Name, Total FROM control_figure");

        double expectedColumnCount = controlFigureData.FirstOrDefault(x => x.Column_Name == "Column Count").Total;
        double expectedRowCount = controlFigureData.FirstOrDefault(x => x.Column_Name == "Row Count").Total;

        bool columnMatch = datasetColumns.SequenceEqual(recordLayoutColumns);
        bool columnCountMatch = datasetColumns.Count == expectedColumnCount;
        bool rowCountMatch = actualDatasetRowCount == expectedRowCount;

        string validationStatus = (columnMatch && columnCountMatch && rowCountMatch) ? "Success" : "Fail";
        string description = validationStatus == "Success"
            ? "Dataset matches control figures."
            : "Mismatch in column count, row count, or dataset structure.";

        await CreateCompareResultsTableAsync();
        await InsertCompareResultsAsync(expectedColumnCount, datasetColumns.Count, expectedRowCount, actualDatasetRowCount, validationStatus, description);

        return columnMatch && columnCountMatch && rowCountMatch;
    }

    private async Task ProcessExcelFileAsync(string filePath, string tableName)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using var stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read()) return;

        var originalColumnNames = new List<string>();
        var sanitizedColumnNames = new List<string>();

        for (int i = 0; i < reader.FieldCount; i++)
        {
            string originalName = reader.GetString(i)?.Trim();
            if (string.IsNullOrWhiteSpace(originalName)) continue;

            string sanitized = SanitizeColumnName(originalName);
            originalColumnNames.Add(originalName);
            sanitizedColumnNames.Add(sanitized);
        }

        await CreateTableAsync(tableName, sanitizedColumnNames);

        var dataRows = new List<Dictionary<string, object>>();

        while (reader.Read())
        {
            var values = new Dictionary<string, object>();
            bool isEmptyRow = true;

            for (int i = 0; i < originalColumnNames.Count; i++)
            {
                object cellValue = reader.GetValue(i);
                if (cellValue != null && cellValue != DBNull.Value && !string.IsNullOrWhiteSpace(cellValue.ToString()))
                {
                    isEmptyRow = false;
                }

                values[sanitizedColumnNames[i]] = cellValue is DBNull ? null : cellValue;
            }

            if (!isEmptyRow)
            {
                dataRows.Add(values);
            }
        }

        if (dataRows.Count > 0)
        {
            await InsertDataAsync(tableName, sanitizedColumnNames, dataRows);
        }
    }

    private async Task CreateTableAsync(string tableName, List<string> columnNames)
    {
        await using DbConnection connection = await dbConnectionFactory.OpenConnectionAsync();

        var columnsSql = columnNames.Select(col => $"[{col}] TEXT").ToArray();

        string createTableQuery = $@"
            CREATE TABLE IF NOT EXISTS {tableName} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT, 
                {string.Join(", ", columnsSql)}
            );";

        await connection.ExecuteAsync(createTableQuery);
    }

    private async Task InsertDataAsync(string tableName, List<string> columnNames, List<Dictionary<string, object>> dataRows)
    {
        await using DbConnection connection = await dbConnectionFactory.OpenConnectionAsync();

        var sanitizedColumnNames = columnNames.Select(SanitizeColumnName).ToList();

        string columnList = string.Join(", ", sanitizedColumnNames.Select(c => $"[{c}]"));
        string parameterList = string.Join(", ", sanitizedColumnNames.Select(c => $"@{c}"));

        string insertQuery = $@"
        INSERT INTO {tableName} ({columnList}) 
        VALUES ({parameterList});";

        foreach (var row in dataRows)
        {
            bool isEmptyRow = row.Values.All(value =>
                value == null || value is DBNull || (value is string str && string.IsNullOrWhiteSpace(str)));

            if (isEmptyRow)
            {
                continue;
            }

            var sanitizedRow = new Dictionary<string, object>();

            foreach (var col in columnNames)
            {
                string sanitizedCol = SanitizeColumnName(col);
                sanitizedRow[sanitizedCol] = row[sanitizedCol] is DBNull ? null : row[sanitizedCol];
            }

            await connection.ExecuteAsync(insertQuery, sanitizedRow);
        }
    }

    private string SanitizeColumnName(string columnName)
    {
        return columnName
            .Replace(" ", "_")
            .Replace("-", "_")
            .Replace(".", "")
            .Replace("/", "_")
            .Replace("\\", "_")
            .Replace("(", "")
            .Replace(")", "")
            .Replace("[", "")
            .Replace("]", "");
    }

    private async Task CreateCompareResultsTableAsync()
    {
        await using DbConnection connection = await dbConnectionFactory.OpenConnectionAsync();

        string createTableQuery = @"
            CREATE TABLE IF NOT EXISTS compare_results (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Expected_Column_Count INTEGER,
                Actual_Column_Count INTEGER,
                Expected_Row_Count INTEGER,
                Actual_Row_Count INTEGER,
                Validation_Status TEXT,
                Description TEXT,
                Timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            );";

        await connection.ExecuteAsync(createTableQuery);
    }

    private async Task InsertCompareResultsAsync(double expectedColumnCount, int actualColumnCount,
        double expectedRowCount, int actualRowCount, string validationStatus, string description)
    {
        await using DbConnection connection = await dbConnectionFactory.OpenConnectionAsync();

        string insertQuery = @"
            INSERT INTO compare_results (Expected_Column_Count, Actual_Column_Count, Expected_Row_Count, Actual_Row_Count, Validation_Status, Description)
            VALUES (@ExpectedColumnCount, @ActualColumnCount, @ExpectedRowCount, @ActualRowCount, @ValidationStatus, @Description);";

        await connection.ExecuteAsync(insertQuery, new
        {
            ExpectedColumnCount = expectedColumnCount,
            ActualColumnCount = actualColumnCount,
            ExpectedRowCount = expectedRowCount,
            ActualRowCount = actualRowCount,
            ValidationStatus = validationStatus,
            Description = description
        });
    }
}