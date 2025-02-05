using Google.Apis.Calendar.v3.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

class ExcelExporter
{
    public static void ExportToExcel(List<Event> events, List<string> selectedColumns, string filePath)
    {
        using (var package = new OfficeOpenXml.ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets["Calendar Events"];
            if (worksheet != null)
            {
                package.Workbook.Worksheets.Delete(worksheet);
            }
            worksheet = package.Workbook.Worksheets.Add("Calendar Events");

            // 設置表頭，根據使用者選擇的欄位
            int columnIndex = 1;
            foreach (var column in selectedColumns)
            {
                worksheet.Cells[1, columnIndex].Value = column;
                columnIndex++;
            }

            // 填入資料，根據選擇的欄位順序
            int rowIndex = 2;
            foreach (var eventItem in events)
            {
                columnIndex = 1;
                foreach (var column in selectedColumns)
                {
                    switch (column)
                    {
                        case "事件名稱":
                            worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Summary;
                            break;
                        case "開始時間":
                            // 如果是全天事件，使用日期
                            if (eventItem.Start.Date != null)
                            {
                                worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Start.Date.ToString();
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Start?.DateTime.ToString();
                            }
                            break;
                        case "結束時間":
                            // 如果是全天事件，使用日期
                            if (eventItem.End.Date != null)
                            {
                                worksheet.Cells[rowIndex, columnIndex].Value = eventItem.End.Date.ToString();
                            }
                            else
                            {
                                worksheet.Cells[rowIndex, columnIndex].Value = eventItem.End?.DateTime.ToString();
                            }
                            break;
                        case "地點":
                            worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Location;
                            break;
                        case "參與者":
                            worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Attendees != null ? string.Join(", ", eventItem.Attendees.Select(a => a.Email)) : "None";
                            break;
                        case "活動說明":
                            worksheet.Cells[rowIndex, columnIndex].Value = eventItem.Description;
                            break;
                    }
                    columnIndex++;
                }
                rowIndex++;
            }

            package.Save();
        }
    }
}
