using System.Data;
using DocNavigator.App.Models;

namespace DocNavigator.App.Services.Data;

public interface IDocumentRepository
{
    Task<(string DoctypeCode, string ServiceCode)?> FindDoctypeAndServiceAsync(long doctypeId, CancellationToken ct);
    Task<DocRow?> FindDocByGuidAsync(Guid globalDocId, CancellationToken ct);
    Task<DocTypeRow?> FindDocTypeAsync(long doctypeId, CancellationToken ct);
    Task<bool> HasRowsAsync(string tableName, long docId, CancellationToken ct);
    Task<DataTable> ReadTableAsync(string tableName, long docId, CancellationToken ct);
}
