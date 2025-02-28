using System.Data.Common;

namespace ExcelTask.Api.Data;

public interface IDbConnectionFactory
{
    ValueTask<DbConnection> OpenConnectionAsync();
}
