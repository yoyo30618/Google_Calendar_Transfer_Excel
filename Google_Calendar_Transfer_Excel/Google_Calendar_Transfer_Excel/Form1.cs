using Google.Apis.Calendar.v3.Data;
using Google.Apis.Calendar.v3;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Windows.Forms;
using static GoogleCalendarHelper;
using OfficeOpenXml;
using static Microsoft.IO.RecyclableMemoryStreamManager;
using System.IO.Packaging;
using System.Runtime.CompilerServices;

namespace Google_Calendar_Transfer_Excel
{
    public partial class Form1 : Form
    {
        private List<string> selectedColumns = new List<string>();
        public Form1()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            InitializeComponent();

            Chk_DefaultRules.Items.Add("事件名稱");   // "Event Name"
            Chk_DefaultRules.Items.Add("開始時間");   // "Start Time"
            Chk_DefaultRules.Items.Add("結束時間");   // "End Time"
            Chk_DefaultRules.Items.Add("地點");       // "Location"
            Chk_DefaultRules.Items.Add("活動說明");     // "Attendees"
            Chk_DefaultRules.SetItemChecked(0, true); // 預設選擇 "事件名稱"
            Chk_DefaultRules.SetItemChecked(1, true); // 預設選擇 "開始時間"
            Chk_DefaultRules.SetItemChecked(2, true); // 預設選擇 "開始時間"
        }
        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (comboBoxCalendars.SelectedItem == null)
            {
                MessageBox.Show("請選擇一個日曆。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CalendarItem selectedCalendar = (CalendarItem)comboBoxCalendars.SelectedItem;
            string selectedCalendarId = selectedCalendar.CalendarId;
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.FileName = $"CalendarEvents_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";  // 加上時間戳
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    btnExport.Enabled = false;
                    btnExport.Text = "匯出中...";
                    DateTime startDate = dateTimePickerStart.Value;
                    DateTime endDate = dateTimePickerEnd.Value.AddDays(1);
                    List<Event> events = await GoogleCalendarHelper.GetCalendarEventsForSpecificCalendar(selectedCalendarId, startDate, endDate);
                    if (events.Count > 0)
                    {
                        string filePath = saveFileDialog.FileName; // 使用者選擇的檔案路徑
                        if (Rbt_Default.Checked)
                        {
                            selectedColumns.Clear();
                            foreach (var item in Chk_DefaultRules.CheckedItems)
                            {
                                selectedColumns.Add(item.ToString());
                            }
                            ExcelExporter.ExportToExcel(events, selectedColumns, filePath);
                            MessageBox.Show($"成功匯出 {events.Count} 筆事件至 {filePath}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (Rbt_Spec.Checked)
                        {
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                var worksheet = package.Workbook.Worksheets.Add("Calendar Events");

                                // 解析特殊規則
                                List<KeyValuePair<string, string>> customColumns = ParseSpecialRules(Tbx_SpecialRule.Text);

                                // 設置標題
                                int columnIndex = 1;
                                foreach (var column in customColumns)
                                {
                                    worksheet.Cells[1, columnIndex].Value = column.Key;
                                    columnIndex++;
                                }

                                // 插入每一筆事件資料
                                int rowIndex = 2;
                                int SpcCount = 0;
                                foreach (var eventItem in events)
                                {
                                    columnIndex = 1;

                                    // 插入使用者選擇的欄位
                                    foreach (var column in customColumns)
                                    {
                                        string value = string.Empty;
                                        if (column.Value[0] == '"' && column.Value[column.Value.Length - 1] == '"')//雙引號包覆的，代表他是純字串
                                            value = column.Value.Substring(1, column.Value.Length - 2);
                                        else if (column.Value[0] == '[' && column.Value[column.Value.Length - 1] == ']')//中瓜號包覆的，代表他是切割字串
                                        {
                                            string specstr = column.Value.Substring(1, column.Value.Length - 2);
                                            var parts = specstr.Split(new[] { "," }, StringSplitOptions.None);
                                            int x = -1;
                                            try
                                            {
                                                x = int.Parse(parts[2]);
                                            }
                                            catch
                                            {
                                                MessageBox.Show("欄位" + column.Key + "_" + column.Value + "有問題，請檢查");
                                                break;
                                            }
                                            if (parts.Length != 3)
                                            {
                                                MessageBox.Show("欄位" + column.Key + "_" + column.Value + "有問題，請檢查");
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    switch (parts[0])
                                                    {
                                                        case "名稱":
                                                            value = eventItem.Summary;
                                                            break;
                                                        case "開始時間":
                                                            value = eventItem.Start.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.Start.Date;
                                                            break;
                                                        case "結束時間":
                                                            value = eventItem.End.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.End.Date;
                                                            break;
                                                        case "地點":
                                                            value = eventItem.Location;
                                                            break;
                                                        case "活動說明":
                                                            value = eventItem.Description;
                                                            break;
                                                        default: value = ""; break;
                                                    }
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("欄位" + column.Key + "_" + column.Value + "有問題，請檢查");
                                                }
                                                var ans = value.Split(new[] { parts[1] }, StringSplitOptions.None);
                                                if (ans.Length <= x)
                                                {
                                                    MessageBox.Show("欄位" + column.Key + "_" + column.Value + "有問題，請檢查");
                                                }
                                                else
                                                {
                                                    value = ans[x];
                                                }
                                            }
                                        }
                                        else//保留字
                                        {
                                            switch (column.Value)
                                            {
                                                case "名稱":
                                                    value = eventItem.Summary;
                                                    break;
                                                case "開始時間":
                                                    value = eventItem.Start.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.Start.Date;
                                                    break;
                                                case "結束時間":
                                                    value = eventItem.End.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.End.Date;
                                                    break;
                                                case "地點":
                                                    value = eventItem.Location;
                                                    break;
                                                case "活動說明":
                                                    value = eventItem.Description;
                                                    break;
                                                default: value = ""; break;
                                            }

                                        }
                                        worksheet.Cells[rowIndex, columnIndex].Value = value;
                                        columnIndex++;
                                    }
                                        SpcCount++;

                                    rowIndex++;
                                }

                                package.Save();
                                MessageBox.Show($"成功匯出 {SpcCount} 筆事件至 {filePath}", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("沒有找到符合條件的事件。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    btnExport.Enabled = true;
                    btnExport.Text = "匯出 Excel";
                }
            }
        }
        private async void Btn_Login_Click(object sender, EventArgs e)
        {
            try
            {
                List<CalendarListEntry> calendars = await GoogleCalendarHelper.GetAllCalendarsAsync();

                if (calendars.Count > 0)
                {
                    // 清空 ComboBox 並添加新日曆
                    comboBoxCalendars.Items.Clear();
                    foreach (var calendar in calendars)
                    {
                        // 將日曆名稱和 ID 包裝為 CalendarItem 物件
                        comboBoxCalendars.Items.Add(new CalendarItem(calendar.Summary, calendar.Id));
                    }

                    if (comboBoxCalendars.Items.Count > 0)
                        comboBoxCalendars.SelectedIndex = 0; // 預設選中第一個
                    comboBoxCalendars.Enabled = true;
                    dateTimePickerStart.Enabled = true;
                    dateTimePickerEnd.Enabled = true;
                    Rbt_Default.Enabled = true;
                    Rbt_Spec.Enabled = true;
                    Chk_DefaultRules.Enabled = true;
                    //Tbx_SpecialRule.Enabled = true;
                    btnExport.Enabled = true;
                    button1.Enabled = true;
                    MessageBox.Show($"成功載入 {calendars.Count} 個日曆", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Btn_Login.Enabled = false;
                }
                else
                {
                    MessageBox.Show("該帳戶沒有其他日曆。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"載入日曆時發生錯誤: {ex.Message}", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 解析特殊規則
        private static List<KeyValuePair<string, string>> ParseSpecialRules(string specialRules)
        {
            var customColumns = new List<KeyValuePair<string, string>>();
            var lines = specialRules.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                var parts = line.Split(new[] { "_" }, StringSplitOptions.None);
                if (parts.Length == 2)
                {
                    string columnName = parts[0].Trim();
                    string rule = parts[1].Trim();
                    customColumns.Add(new KeyValuePair<string, string> ( columnName, rule ));
                }
            }

            return customColumns;
        }
        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            ChangeRadioButton();
        }
        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            ChangeRadioButton();
        }
        private void ChangeRadioButton()
        {
            if (Rbt_Default.Checked)
            {
                Tbx_SpecialRule.Enabled = false;
                Chk_DefaultRules.Enabled = true;
            }
            else if (Rbt_Spec.Checked)
            {
                Tbx_SpecialRule.Enabled = true;
                Chk_DefaultRules.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string credPath = "token.json"; // OAuth Token 儲存
            if (Directory.Exists(credPath))
            {
                Directory.Delete(credPath, true); // true 代表遞迴刪除
            }
            comboBoxCalendars.Enabled = false;
            dateTimePickerStart.Enabled = false;
            dateTimePickerEnd.Enabled = false;
            Rbt_Default.Enabled = false;
            Rbt_Spec.Enabled = false;
            Chk_DefaultRules.Enabled = false;
            Tbx_SpecialRule.Enabled = true;
            btnExport.Enabled = false;
            button1.Enabled=false;
            Tbx_SpecialRule.Enabled=false;
            Btn_Login.Enabled = true;
        }
    }
}
