using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace EPPlus.WebSampleMvc.NetCore.Samples.Sample2
{
    /// <summary>
    /// This class generates 1000
    /// </summary>
    public class OrderRepository
    {
        private string[] _customers = new string[]
        {
            "Customer 1",
            "Customer 2",
            "Customer 3",
            "Customer 4",
            "Customer 5",
            "Customer 6",
            "Customer 7",
            "Customer 8",
            "Customer 9",
            "Customer 10",
            "Customer 11",
            "Customer 12"
        };

        private readonly List<Order> _orders = new List<Order>();
        private readonly Random _random = new Random(DateTime.Now.Millisecond);
        private const int NumberOfOrders = 1000;

        private string GetNewCustomer()
        {
            var ix = _random.Next(0, _customers.Length);
            return _customers[ix];
        }

        private DateTime GetNewOrderDate()
        {
            var minDate = DateTime.Now.AddDays(-150d).ToOADate();
            var maxDate = DateTime.Now.AddDays(-2d).ToOADate();
            var date = _random.Next((int)minDate, (int)maxDate);
            return DateTime.FromOADate(date);
        }

        private int GetNewOrderValue()
        {
            return _random.Next(100, 1500);
        }

        public IEnumerable<Order> GetAllOrders()
        {
            if (_orders.Count > 0) return _orders;
            for(var i = 1; i <= NumberOfOrders; i++)
            {
                _orders.Add(new Order
                {
                    OrderId = i,
                    CustomerName = GetNewCustomer(),
                    OrderDate = GetNewOrderDate(),
                    OrderValue = GetNewOrderValue(),
                    Currency = "USD"
                }); ;
            }
            return _orders.OrderBy(x => x.OrderDate);
        }
    }
}
