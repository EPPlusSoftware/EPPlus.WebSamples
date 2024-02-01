using System.Collections.Generic;
using System;

namespace EPPlus.WebSampleMvc.NetCore.HelperClasses
{
    public class EnumCollection<T> where T : struct
    {
        public Dictionary<string, string[]> dict = new Dictionary<string, string[]>();

        public EnumCollection(Type[] enums)
        {
            var arr = Enum.GetNames(typeof(T));

            for (int i = 0; i < enums.Length; i++)
            {
                var names = Enum.GetNames(enums[i]);
                dict.Add(arr[i], names);
            }
        }
    }
}
