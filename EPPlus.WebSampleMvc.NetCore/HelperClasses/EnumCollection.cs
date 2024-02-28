using System.Collections.Generic;
using System;
using System.Linq;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
{
    public class EnumCollection
    {
        public Dictionary<string, string[]> Dict = new Dictionary<string, string?[]>();

        public string SelectedKey { get; set; }
        public string? SelectedValue { get; set; }

        public EnumCollection(Type parentEnum, Type[] enums, string[] selectedEnums = null, int maxSelected = 5, int minSelected = 1)
        {
            var arr = Enum.GetNames(parentEnum);

            for (int i = 0; i < arr.Length; i++)
            {
                string[] names = new string[0];
                if (i < enums.Length)
                {
                    names = Enum.GetNames(enums[i]);
                }
                Dict.Add(arr[i], names);
            }
            SelectedKey = "";
            SelectedValue = null;
            DetermineSelected(selectedEnums, maxSelected, minSelected);
            //var first = dict.First();
            //if (SelectedKey == null)
            //{
            //    SelectedKey = first.Key;
            //}
            //if(SelectedValue == null) 
            //{
            //    SelectedValue = first.Value[0];
            //}
        }

        private void DetermineSelected(string[] selectedEnums = null, int? maxSelected = null, int? minSelected = null)
        {
            if (selectedEnums != null && selectedEnums[0] != null)
            {
                if (selectedEnums.Length < minSelected)
                {
                    throw new ArgumentException($"EnumCollection requires at least {minSelected} enums");
                }

                if (selectedEnums.Length > maxSelected)
                {
                    throw new ArgumentException($"Selected Enums cannot be larger than {maxSelected}");
                }

                //for(int i = 0; i < selectedEnums.Length; i++)
                //{

                //}

                SelectedKey = selectedEnums[0];

                if (selectedEnums.Length > minSelected)
                {
                    SelectedValue = selectedEnums[1];
                }
                else
                {
                    SelectedValue = "";
                }
            }
            else
            {
                var first = Dict.First();
                SelectedKey = first.Key;
                if (first.Value.Length > 0)
                {
                    SelectedValue = first.Value[0];
                }
            }
        }
    }
}
