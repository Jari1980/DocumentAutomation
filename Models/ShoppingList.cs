using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections.Generic;
using System.Text;

namespace DocumentAutomation.Models
{
    public class Item
    {
        public string Name { get; set; } = "";
        public int Quantity { get; set; }
    }
    public class ShoppingList
    {
        public string CustomerName { get; set; } = "";
        public DateTime Date { get; set; } = DateTime.Today;
        public List<Item> Items { get; set; } = new();
    }
}
