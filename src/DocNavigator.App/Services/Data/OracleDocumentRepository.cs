using System;
using System.Data;
using System.Data.Common;
using Dapper;
using DocNavigator.App.Models;
using Oracle.ManagedDataAccess.Client;

namespace DocNavigator.App.Services.Data;

public class OracleDocumentRepository : IDocumentRepository
{
    private readonly DbProfile _profile;
    private readonly string _schema;

    public OracleDocumentRepository(DbProfile profile)
    {
        _profile = profile;
        _schema = string.IsNullOrWhiteSpace(profile.Schema) ? profile.Name : profile.Schema!;
    }

    public async Task<DocRow?> FindDocByGuidAsync(Guid globalDocId, CancellationToken ct)
    {
        await using var conn = (DbConnection)DbConnectionFactory.Create(_profile);
        var sql = $@"SELECT docid, globaldocid, doctypeid, createdate
FROM {_schema}.doc WHERE globaldocid = :g FETCH FIRST 1 ROWS ONLY";
        var row = await conn.QuerySingleOrDefaultAsync(sql, new { g = globalDocId.ToString() });
        if (row == null) return null;
        return new DocRow(Convert.ToInt64(row.DOCID), Guid.Parse((string)row.GLOBALDOCID), Convert.ToInt64(row.DOCTYPEID), (DateTime)row.CREATEDATE);
    }

    public async Task<DocTypeRow?> FindDocTypeAsync(long doctypeId, CancellationToken ct)
    {
        await using var conn = (DbConnection)DbConnectionFactory.Create(_profile);
        var sql = $@"SELECT doctypeid, systemname FROM {_schema}.doctype WHERE doctypeid = :id";
        var row = await conn.QuerySingleOrDefaultAsync(sql, new { id = doctypeId });
        if (row == null) return null;
        return new DocTypeRow(Convert.ToInt64(row.DOCTYPEID), (string)row.SYSTEMNAME);
    }

    public async Task<bool> HasRowsAsync(string tableName, long docId, CancellationToken ct)
    {
        await using var conn = (DbConnection)DbConnectionFactory.Create(_profile);
        var sql = $@"SELECT 1 FROM {_schema}.{tableName} WHERE docid=:id AND ROWNUM = 1";
        var exists = await conn.ExecuteScalarAsync<int?>(sql, new { id = docId });
        return exists.HasValue;
    }

    public async Task<DataTable> ReadTableAsync(string tableName, long docId, CancellationToken ct)
    {
        await using var conn = (DbConnection)DbConnectionFactory.Create(_profile);
        var sql = $@"SELECT * FROM {_schema}.{tableName} WHERE docid=:id ORDER BY 1";
        using var cmd = conn.CreateCommand();
        cmd.CommandText = sql;
        var p = cmd.CreateParameter();
        p.ParameterName = "id";
        p.Value = docId;
        cmd.Parameters.Add(p);
        await conn.OpenAsync(ct);
        var adapter = new OracleDataAdapter((OracleCommand)cmd);
        var dt = new DataTable(tableName);
        adapter.Fill(dt);
        return dt;
    }

    
    public Task<(string DoctypeCode, string ServiceCode)?> FindDoctypeAndServiceAsync(long doctypeId, CancellationToken ct)
{
    using var conn = DbConnectionFactory.Create(_profile);
    string sql =
        $"SELECT dt.systemname as doctype_code, ds.systemname as service_code " +
        $"FROM {_schema}.doctype dt " +
        $"JOIN {_schema}.docservice ds ON ds.docserviceid = dt.docserviceid " +
        $"WHERE dt.doctypeid = :id";

    using var cmd = conn.CreateCommand();
    cmd.CommandText = sql;

    var p = cmd.CreateParameter();
    p.ParameterName = "id"; // имя параметра в коллекции — без двоеточия
    p.Value = doctypeId;
    cmd.Parameters.Add(p);

    conn.Open(); // синхронно
    using var reader = cmd.ExecuteReader(); // синхронно
    if (reader.Read())
    {
        var dt = reader.GetString(0);
        var svc = reader.GetString(1);
        return Task.FromResult< (string DoctypeCode, string ServiceCode)? >((dt, svc));
    }
    return Task.FromResult< (string DoctypeCode, string ServiceCode)? >(null);
}

}
