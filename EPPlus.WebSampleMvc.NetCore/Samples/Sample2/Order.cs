using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Samples.Sample2
{
    /// <summary>
    /// A simple class that represents an order.
    /// The <see cref="DisplayNameAttribute"></see> is used to set the
    /// headers when calling the LoadFromCollection method
    /// </summary>
    public class Order
    {
        [DisplayName("Order Number")]
        public int OrderId { get; set; }

        [DisplayName("Customer")]
        public string CustomerName { get; set; }

        [DisplayName("Order date")]
        public DateTime OrderDate { get; set; }

        [DisplayName("Order value")]
        public decimal OrderValue { get; set; }

        [DisplayName("Currency")]
        public string Currency { get; set; }
    }
}
