using System;
using System.Collections.Generic;
using System.Text;

namespace ProductSaveLoad
{
    public class Product
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public double Price { get; set; }
        public string Category { get; set; }

        public Product(int id, string name, double price, string category)
        {
            Id = id;
            Name = name;
            Price = price;
            Category = category;
        }

        public override string ToString()
        {
            return $"id = {Id}, name = {Name}, price = {Price}, category = {Category}";
        }
    }
}
