using System.Data;
using DocNavigator.App.Models;
using Npgsql;
using Oracle.ManagedDataAccess.Client;

namespace DocNavigator.App.Services.Data;

public static class DbConnectionFactory
{
    public static IDbConnection Create(DbProfile profile)
        => profile.Kind.ToLowerInvariant() switch
        {
            "postgres" => new NpgsqlConnection(profile.ConnectionString),
            "oracle"   => new OracleConnection(profile.ConnectionString),
            _ => throw new NotSupportedException($"Unknown profile kind: {profile.Kind}")
        };
}
