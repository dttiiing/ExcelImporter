using NPOI.HPSF;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Diagnostics;
using System.IO;

namespace ExcelImporter
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"E:\MyProjects\ExcelImporter\excel\User.xlsx";

            ReadExcel(fileName);
        }

        public static void ReadExcel(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"{filePath} is not exit");
                return;
            }

            string extension = Path.GetExtension(filePath);
            FileStream fileStream = File.OpenRead(filePath);

            IWorkbook workbook = extension.Equals(".xls") ? new HSSFWorkbook(fileStream) : new XSSFWorkbook(fileStream);
            fileStream.Close();

            ISheet sheet = workbook.GetSheetAt(0);
            IRow row = null;

            for (int i = 0; i < sheet.LastRowNum + 1; i++)
            {
                row = sheet.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < row.LastCellNum; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null)
                        {
                            Console.Write(cell.ToString() + " ");
                        }
                    }
                    Console.WriteLine("\n");
                }
            }
        }
    }
}
