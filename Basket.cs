using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ProductSaveLoad
{
    class Basket
    {
        public List<Product> grocery = new List<Product>();

        public void FillWithDummyData()
        {
            grocery.Add(new Product(1, "Spaghetti", 2.5, "Food"));
            grocery.Add(new Product(2, "Tommato", 3.87, "Vegetables"));
            grocery.Add(new Product(3, "Coffee", 6.55, "Drinks"));
        }

        public bool SaveExcel(string ExcelFileName)
        {
            XSSFWorkbook wb = new XSSFWorkbook();
            ISheet sheet = wb.CreateSheet("My Sheet");

            var r = sheet.CreateRow(0);
            r.CreateCell(0).SetCellValue("Id");
            r.CreateCell(1).SetCellValue("Name");
            r.CreateCell(2).SetCellValue("Price");
            r.CreateCell(3).SetCellValue("Category");

            for (int i = 0; i < grocery.Count; i++)
            {
                r = sheet.CreateRow(i + 1);
                r.CreateCell(0).SetCellValue(grocery[i].Id);
                r.CreateCell(1).SetCellValue(grocery[i].Name);
                r.CreateCell(2).SetCellValue(grocery[i].Price);
                r.CreateCell(3).SetCellValue(grocery[i].Category);
            }
            using (var fs = new FileStream(ExcelFileName, FileMode.Create, FileAccess.Write))
            {
                wb.Write(fs);
            }

            return true;
        }

        public bool SaveJson(string JasonFileName)
        {
            string jsonData = JsonConvert.SerializeObject(grocery);
            File.WriteAllText(JasonFileName, jsonData);

            return true;
        }

        public void SaveText(string TextFileName)
        {
            using (StreamWriter file = new StreamWriter(TextFileName))
            {
                foreach (Product p in grocery)
                {
                    file.WriteLine(p);
                }
            }
        }

        public bool LoadExcel(string ExcelFileName)
        {
            grocery.Clear();
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(ExcelFileName, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);
            }
            ISheet sheet = hssfwb.GetSheet("My Sheet");

            //first line (zero line) contains headers
            for (int row = 1; row <= sheet.LastRowNum; row++)
            {
                //null is when the row only conatins empty cells
                if (sheet.GetRow(row) != null)
                {
                    int Id = int.Parse(sheet.GetRow(row).GetCell(0).ToString());
                    string Name = sheet.GetRow(row).GetCell(1).ToString();
                    double Price = double.Parse(sheet.GetRow(row).GetCell(2).ToString());
                    string Category = sheet.GetRow(row).GetCell(3).ToString();

                    Product p = new Product(Id, Name, Price, Category);
                    grocery.Add(p);
                }
            }
            return true;
        }

        public void LoadJson(string JasonFileName)
        {
            string jsondata = File.ReadAllText(JasonFileName);
            grocery = JsonConvert.DeserializeObject<List<Product>>(jsondata);
        }

        public bool LoadText(string TextFileName)
        {
            try
            {

                var list = new List<string>();
                var fileStream = new FileStream(TextFileName, FileMode.Open, FileAccess.Read);
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8))
                {
                    string line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        list.Add(line);
                    }
                }
                foreach (String l in list)
                {
                    Console.WriteLine("....................");
                    Console.WriteLine(l);
                }

            }

            catch (Exception)
            {
                return false;

            }
            return true;
        }
    }
}