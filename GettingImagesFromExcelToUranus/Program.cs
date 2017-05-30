using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;

namespace GettingImagesFromExcelToUranus
{
    class Program
    {
        static void Main(string[] args)
        {
            string brand = "sloan";

            var table = new LinqToExcel.ExcelQueryFactory(@"Z:\_Product Data\Sloan\SloanImages.xlsx");
            var query =
                from row in table.Worksheet("Product Data")
                let item = new
                {
                    InternalID = row["Internal ID"].Cast<string>(),
                    Sku = row["SKU"].Cast<string>(),
                    Image = row["Product Image"].Cast<string>()
                }
                where item.Image != null
                select item;

            int count = 1;
            string targetFilePath = @"Z:\_Product Data\Sloan\TEST\";  
            foreach (var row in query)
            {
                string assetPhotoFile = brand + "-" + row.Sku.ToLower() + ".jpg";
                string targetFile = Path.Combine(targetFilePath, assetPhotoFile);

                try
                {
                    using (WebClient client = new WebClient())
                    {
                        client.DownloadFile(row.Image, targetFile);
                    }
                }

                catch (Exception ex) 
                {
                    Console.Write("\nCaught exception at count: " + count + "\nDetails: " + ex.GetBaseException());
                }

                Console.Write("\nImage number " + count + " added to TEST FOLDER: " + assetPhotoFile + @" successfully");
                count++;
            }
         }
    }
}
