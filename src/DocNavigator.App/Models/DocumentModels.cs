using System;
namespace DocNavigator.App.Models;
public record DocRow(long DocId, Guid GlobalDocId, long DocTypeId, DateTime CreatedOn);
public record DocTypeRow(long DocTypeId, string SystemName);
