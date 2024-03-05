using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System;
using EPPlus.WebSampleMvc.NetCore.HelperClasses;
using System.Text.Json;
using EPPlus.WebSampleMvc.NetCore.HelperClasses.ConditionalFormattingNew;
using Newtonsoft.Json.Linq;

namespace EPPlus.WebSampleMvc.NetCore.Models.HtmlExport
{
    public class ExportRanges8Model
    {
        public void SetupSampleData()
        {
            using (var package = new ExcelPackage())
            {
                var collection = new RuleTypeCollection();

                var options = new JsonSerializerOptions()
                {
                    IncludeFields = true,
                };
                //types.ActiveRule.Settings.SetColor(ActiveColor);
                var JsonData = JsonSerializer.Serialize(collection, options);

                var jsonObject = JObject.Parse(JsonData);

                var AllCells = jsonObject[0];
                var CellContains = jsonObject[1];
                var Ranked = jsonObject[2];
                var Average = jsonObject[3];
                var UniqueDuplicates = jsonObject[4];
                var CustomExpression = jsonObject[5];

                var ws = package.Workbook.Worksheets.Add("NewWorksheet");

                var exporter = ws.Cells["A1:E11"].CreateHtmlExporter();

                var settings = exporter.Settings;
                settings.Pictures.Include = ePictureInclude.Include;
                //settings.Pictures.KeepOriginalSize = true;
                settings.Minify = false;
                settings.SetColumnWidth = true;
                settings.SetRowHeight = true;
                settings.Pictures.AddNameAsId = true;

                ws.Calculate();

                // export css and html
                Css = exporter.GetCssString();
                Html = exporter.GetHtmlString();
            }
        }

        public string Css { get; set; }
        public string Html { get; set; }

        public string JsonData = "";

        public CFRuleType CurrentRuleType { get; set; } = CFRuleType.CellContains;

        public string[] GetEnumValues()
        {
            return Enum.GetNames(typeof(eCellContains));
        }

        public string[] GetEnumValuesDependent()
        {
            return Enum.GetNames(typeof(CellValueCondition));
        }

        public string Formula1 { get; set; } = "";
        public string Formula2 { get; set; } = "";

        public Array GetColValues()
        {
            return Enum.GetValues(typeof(CFColor));
        }

        public CFColor ActiveColor { get; set; }

        public bool? Checkbox { get; set; }

        public string CurrentRuleTypeStr { get; set; }

        //public void SetupSampleData()
        //{
        //    using (var package = new ExcelPackage())
        //    {



        //        if (Formula1 != null)
        //        {
        //            Formulas[0] = Formula1;
        //        }
        //        if (Formula2 != null)
        //        {
        //            Formulas[1] = Formula2;
        //        }

        //        CurrentRuleType = (CFRuleType)Enum.Parse(typeof(CFRuleType), CurrentRuleTypeStr);

        //        var types = new RuleTypes(CurrentRuleType, SelectedEnums, Formulas, Settings, Checkbox);
        //        ActiveIndex = types.ActiveIndex;
        //        CFRules = types.Types;

        //        var options = new JsonSerializerOptions()
        //        {
        //            IncludeFields = true,
        //        };
        //        types.ActiveRule.Settings.SetColor(ActiveColor);
        //        JsonData = JsonSerializer.Serialize(CFRules, options);

        //        var ws = package.Workbook.Worksheets.Add("CF_ws");

        //        ws.Cells["A1:A11"].Formula = "ROW()-5";
        //        ws.Cells["B1:B11"].Formula = "\"Row number \" & ROW()";
        //        ws.Cells["C1:C11"].Formula = "TODAY() -2 + ROW()";
        //        ws.Cells["D1"].Formula = "TODAY() -6";
        //        ws.Cells["D2"].Formula = "TODAY()";
        //        ws.Cells["D3"].Formula = "TODAY() + 7";
        //        ws.Cells["D5"].Formula = "TODAY() -31";
        //        ws.Cells["D6"].Formula = "TODAY()";
        //        ws.Cells["D7"].Formula = "TODAY() + 31";

        //        //Add in to test errors
        //        //ws.Cells["E1:E5"].Formula = "E10-E11+NULL";

        //        ws.Cells["C1:D11"].Style.Numberformat.Format = "mm-dd-yy";

        //        ws.Cells["A1:E11"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        ws.Cells["A1:E11"].Style.Fill.BackgroundColor.SetColor(Color.AliceBlue);


        //        var cs = ws.Cells["Z1:Z10"].ConditionalFormatting.AddDatabar(Color.AliceBlue);
        //        string testJson = JsonSerializer.Serialize(cs);

        //        var cf = types.ActiveRule.GetCFForRange(ws.Cells["A1:E11"].ConditionalFormatting);

        //        cf.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
        //        cf.Style.Fill.BackgroundColor.Color = types.ActiveRule.Settings.Color;

        //        if (((int)cf.Type < 10) && ((int)cf.Type > 3))
        //        {
        //            if (int.TryParse(Formula1, out int newRank))
        //            {
        //                cf.Rank = (ushort)newRank;
        //            }
        //            else
        //            {
        //                //Faulty rank input becomes 1
        //                cf.Rank = 1;
        //            }
        //        }

        //        cf.Formula = Formula1 == null ? "" : Formula1;
        //        if (cf.Type == eExcelConditionalFormattingRuleType.Between || cf.Type == eExcelConditionalFormattingRuleType.NotBetween)
        //        {
        //            cf.Formula2 = Formula2 == null ? "" : Formula2;
        //        }

        //        var exporter = ws.Cells["A1:E11"].CreateHtmlExporter();

        //        var settings = exporter.Settings;
        //        settings.Pictures.Include = ePictureInclude.Include;
        //        //settings.Pictures.KeepOriginalSize = true;
        //        settings.Minify = false;
        //        settings.SetColumnWidth = true;
        //        settings.SetRowHeight = true;
        //        settings.Pictures.AddNameAsId = true;

        //        ws.Calculate();

        //        // export css and html
        //        Css = exporter.GetCssString();
        //        Html = exporter.GetHtmlString();

        //        //try
        //        //{
        //        //    ws.Calculate();

        //        //    // export css and html
        //        //    Css = exporter.GetCssString();
        //        //    Html = exporter.GetHtmlString();
        //        //}
        //        //catch (Exception e)
        //        //{
        //        //    ErrorMessage = e.Message;
        //        //}
        //    }
        //}

        //public CFRuleType CurrentRuleType { get; set; } = CFRuleType.CellContains;

        //public string ErrorMessage { get; set; } = "";

        //public bool? Checkbox { get; set; }

        //public List<IExcelConditionalFormattingHTMLRuleGroup> CFRules = new List<IExcelConditionalFormattingHTMLRuleGroup>();

        //public string CurrentRuleTypeStr { get; set; }

        //public int ActiveIndex { get; set; }

        //public string SelectedKey { get; set; } = null;
        //public string SelectedValue { get; set; } = null;

        //public string[] SelectedEnums { get; set; } = new string[5];

        //public string[] Formulas { get; set; } = new string[2];

        //public string Formula1 { get; set; } = "";
        //public string Formula2 { get; set; } = "";

        //public FormatSettings Settings { get; set; }

        //public string[] GetEnumValues()
        //{
        //    return Enum.GetNames(typeof(CellContains));
        //}

        //public string[] GetEnumValuesDependent()
        //{
        //    return Enum.GetNames(typeof(CellValueCondition));
        //}

        //public string JsonData = "";

        //public string Css { get; set; }

        //public string Html { get; set; }

        //public CFColor ActiveColor { get; set; }

        //public Array GetColValues()
        //{
        //    return Enum.GetValues(typeof(CFColor));
        //}
    }
}
