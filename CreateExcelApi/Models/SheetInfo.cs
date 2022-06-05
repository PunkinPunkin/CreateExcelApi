namespace CreateExcelApi.Models
{
    public class SheetInfo
    {
        private string? _title;
        public string Name { get; set; } = "Sheet1";
        public string Title { get => _title ?? Name; set => _title = value; }
        public string? SubTitle { get; set; }
        public string? SearchCondition { get; set; }
        public ICollection<CellInfo> TableHeaders { get; set; } = Array.Empty<CellInfo>();
        public ICollection<ICollection<string>> Data { get; set; } = Array.Empty<ICollection<string>>();
    }
}
