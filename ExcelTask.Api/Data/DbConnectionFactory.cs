using Microsoft.Data.Sqlite;
using System.Data.Common;

namespace ExcelTask.Api.Data;

public class DbConnectionFactory(string connectionString) : IDbConnectionFactory
{
    private readonly string _connectionString = connectionString;

    public async ValueTask<DbConnection> OpenConnectionAsync()
    {
        var connection = new SqliteConnection(_connectionString);
        await connection.OpenAsync();
        return connection;
    }
}
