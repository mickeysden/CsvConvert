using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using CsvConvert.Helpers;
using CsvHelper;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace CsvConvert.Services
{
    public interface IXlsService
    {
        Task<string> ToXlsx(string csvFilename);
    }

    public class XlsService : IXlsService
    {
        public async Task<string> ToXlsx(string csvFilename)
        {
            SimpleLogger.Log("processing " + csvFilename);
            string resultFile = String.Empty;

            try
            {
                using (var reader = new StreamReader(csvFilename))
                using (var csv = new CsvReader(reader))
                {
                    var records = csv.GetRecords<dynamic>();
                    resultFile = await WriteXlsx(records, csvFilename.Split("/")[csvFilename.Split("/").Length-1]);
                }
                File.Delete(csvFilename);
            }
            catch (Exception ex)
            {
                SimpleLogger.Log(ex);
            }

            return resultFile;
        }

        private async Task<string> WriteXlsx(IEnumerable<dynamic> records, string originalFilename)
        {
            string path = "Resources/Xlsx/" + Guid.NewGuid().ToString() + "_";
            string fileName = path + originalFilename + "_" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx";

            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet sheet1 = await CreateSheet(records, workbook, "Sheet1");
                workbook.Write(fs);
            }
            return fileName;
        }

        private async Task<ISheet> CreateSheet(IEnumerable<dynamic> records, IWorkbook workbook, string sheetName)
        {
            await Task.Delay(1);
            List<Dictionary<string,string>> data = new List<Dictionary<string, string>>();

            foreach (dynamic s in records)
            {
                var dynamicDictionary = s as IDictionary<string, object>;
                //data.Add(s);

                var columnData = new Dictionary<string,string>();
                foreach (KeyValuePair<string, object> property in dynamicDictionary)
                {
                    columnData.Add(property.Key,property.Value.ToString());
                }
                data.Add(columnData);
            }

            var style = workbook.CreateCellStyle();
            style.Alignment = HorizontalAlignment.Center;

            ISheet sheet1 = workbook.CreateSheet(sheetName);
            IRow row1 = sheet1.CreateRow(0);

            //int funnelStepCounter = 0;
            int cellCounter = 0;

            //header row
            int rowCounter = 0;
            foreach (var i in data)
            {
                foreach (var key in i.Keys)
                {
                    SimpleLogger.Log(++rowCounter + " " + key);
                    row1.CreateCell(cellCounter++).SetCellValue(key);
                }
                break;
            }

            int counter = 0;
            foreach (var i in data)
            {
                //foreach row
                IRow row = sheet1.CreateRow(++counter);
                cellCounter = 0;
                foreach (var key in i.Keys) row.CreateCell(cellCounter++).SetCellValue(i[key]);
            }

            for (int i = 0; i <= cellCounter; i++) sheet1.AutoSizeColumn(i);
            return sheet1;
        }
    }
}