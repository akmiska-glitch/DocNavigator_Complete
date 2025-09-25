namespace DocNavigator.App.Models
{
    public class FieldMeta
    {
        public string SystemName { get; }
        public string? RussianName { get; }
        public string? DataType { get; }

        public FieldMeta(string sys, string? ru, string? dt)
        {
            SystemName = sys;
            RussianName = ru;
            DataType = dt;
        }
    }
}
