namespace DocNavigator.App.Models;

public class DbProfile
{
    public string Name { get; set; } = "Default";
    public string Kind { get; set; } = "postgres";
    public string ConnectionString { get; set; } = string.Empty;
    public string? Schema { get; set; } = null;

    // Remote .desc source settings
    public string DescBaseUrl { get; set; } = "http://fk-eb-arp-demo-ufos:18080";
    public string DescUrlTemplate { get; set; } = "/ARP/static/resources/forms/services/{service}/{doctype}/{version}/{doctype}.desc";
    public string DescVersion { get; set; } = "1.0";
}
