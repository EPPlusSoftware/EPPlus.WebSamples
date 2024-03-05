namespace EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew
{
    public class InputData
    {
        //public static InputData Default { get; }
        //TODO: Output data as well?
        public string[]? SelectedEnums;
        public string[] Formulas;
        public FormatSettings Settings;
        public bool? CheckBox;

        public InputData(string[]? selectedEnums = null, string[]? formulas = null, FormatSettings? settings = null, bool? checkbox = null)
        {
            CheckBox = checkbox;
            SelectedEnums = selectedEnums;
            Formulas = formulas == null ? new string[1] { "" } : formulas;
            Settings = settings == null ?
                new FormatSettings(CFColor.Blue) :
                settings;
        }
    }
}
