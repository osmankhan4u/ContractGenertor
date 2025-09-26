namespace ContractGeneratorBlazor.Models
{
    public class ContractTemplate
    {
        public string Type { get; set; } = string.Empty;
        public Dictionary<string, string> Placeholders { get; set; } = new();
        public string TemplatePath { get; set; } = string.Empty;
    }
}
