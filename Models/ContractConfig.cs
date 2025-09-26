namespace ContractGeneratorBlazor.Models
{
    public class ContractTemplateConfig
    {
        public string Type { get; set; } = string.Empty;
        public string TemplatePath { get; set; } = string.Empty;
        public Dictionary<string, string> Placeholders { get; set; } = new();
    }

    public class ContractConfig
    {
        public List<ContractTemplateConfig> Contracts { get; set; } = new();
    }
}
