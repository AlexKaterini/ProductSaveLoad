using System;

namespace ProductSaveLoad
{
    class Program
    {
        static void Main(string[] args)
        {
            Basket basket = new Basket();

            basket.FillWithDummyData();

            basket.SaveText("list.txt");
            basket.SaveJson("jsondata.json");
            basket.SaveExcel("ExcelFileName.xlsx");
            basket.LoadText("list.txt");
            basket.LoadJson("jsondata.json");
            basket.LoadExcel("ExcelFileName.xlsx");
        }
    }
}
