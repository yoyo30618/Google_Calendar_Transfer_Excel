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

            Chk_DefaultRules.Items.Add("�ƥ�W��");   // "Event Name"
            Chk_DefaultRules.Items.Add("�}�l�ɶ�");   // "Start Time"
            Chk_DefaultRules.Items.Add("�����ɶ�");   // "End Time"
            Chk_DefaultRules.Items.Add("�a�I");       // "Location"
            Chk_DefaultRules.Items.Add("���ʻ���");     // "Attendees"
            Chk_DefaultRules.SetItemChecked(0, true); // �w�]��� "�ƥ�W��"
            Chk_DefaultRules.SetItemChecked(1, true); // �w�]��� "�}�l�ɶ�"
            Chk_DefaultRules.SetItemChecked(2, true); // �w�]��� "�}�l�ɶ�"
        }
        private async void btnExport_Click(object sender, EventArgs e)
        {
            if (comboBoxCalendars.SelectedItem == null)
            {
                MessageBox.Show("�п�ܤ@�Ӥ��C", "����", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            CalendarItem selectedCalendar = (CalendarItem)comboBoxCalendars.SelectedItem;
            string selectedCalendarId = selectedCalendar.CalendarId;
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                saveFileDialog.DefaultExt = "xlsx";
                saveFileDialog.FileName = $"CalendarEvents_{DateTime.Now.ToString("yyyyMMdd_HHmmss")}.xlsx";  // �[�W�ɶ��W
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    btnExport.Enabled = false;
                    btnExport.Text = "�ץX��...";
                    DateTime startDate = dateTimePickerStart.Value;
                    DateTime endDate = dateTimePickerEnd.Value.AddDays(1);
                    List<Event> events = await GoogleCalendarHelper.GetCalendarEventsForSpecificCalendar(selectedCalendarId, startDate, endDate);
                    if (events.Count > 0)
                    {
                        string filePath = saveFileDialog.FileName; // �ϥΪ̿�ܪ��ɮ׸��|
                        if (Rbt_Default.Checked)
                        {
                            selectedColumns.Clear();
                            foreach (var item in Chk_DefaultRules.CheckedItems)
                            {
                                selectedColumns.Add(item.ToString());
                            }
                            ExcelExporter.ExportToExcel(events, selectedColumns, filePath);
                            MessageBox.Show($"���\�ץX {events.Count} ���ƥ�� {filePath}", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else if (Rbt_Spec.Checked)
                        {
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                var worksheet = package.Workbook.Worksheets.Add("Calendar Events");

                                // �ѪR�S��W�h
                                List<KeyValuePair<string, string>> customColumns = ParseSpecialRules(Tbx_SpecialRule.Text);

                                // �]�m���D
                                int columnIndex = 1;
                                foreach (var column in customColumns)
                                {
                                    worksheet.Cells[1, columnIndex].Value = column.Key;
                                    columnIndex++;
                                }

                                // ���J�C�@���ƥ���
                                int rowIndex = 2;
                                int SpcCount = 0;
                                foreach (var eventItem in events)
                                {
                                    columnIndex = 1;

                                    // ���J�ϥΪ̿�ܪ����
                                    foreach (var column in customColumns)
                                    {
                                        string value = string.Empty;
                                        if (column.Value[0] == '"' && column.Value[column.Value.Length - 1] == '"')//���޸��]�Ъ��A�N��L�O�¦r��
                                            value = column.Value.Substring(1, column.Value.Length - 2);
                                        else if (column.Value[0] == '[' && column.Value[column.Value.Length - 1] == ']')//���ʸ��]�Ъ��A�N��L�O���Φr��
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
                                                MessageBox.Show("���" + column.Key + "_" + column.Value + "�����D�A���ˬd");
                                                break;
                                            }
                                            if (parts.Length != 3)
                                            {
                                                MessageBox.Show("���" + column.Key + "_" + column.Value + "�����D�A���ˬd");
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    switch (parts[0])
                                                    {
                                                        case "�W��":
                                                            value = eventItem.Summary;
                                                            break;
                                                        case "�}�l�ɶ�":
                                                            value = eventItem.Start.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.Start.Date;
                                                            break;
                                                        case "�����ɶ�":
                                                            value = eventItem.End.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.End.Date;
                                                            break;
                                                        case "�a�I":
                                                            value = eventItem.Location;
                                                            break;
                                                        case "���ʻ���":
                                                            value = eventItem.Description;
                                                            break;
                                                        default: value = ""; break;
                                                    }
                                                }
                                                catch
                                                {
                                                    MessageBox.Show("���" + column.Key + "_" + column.Value + "�����D�A���ˬd");
                                                }
                                                var ans = value.Split(new[] { parts[1] }, StringSplitOptions.None);
                                                if (ans.Length <= x)
                                                {
                                                    MessageBox.Show("���" + column.Key + "_" + column.Value + "�����D�A���ˬd");
                                                }
                                                else
                                                {
                                                    value = ans[x];
                                                }
                                            }
                                        }
                                        else//�O�d�r
                                        {
                                            switch (column.Value)
                                            {
                                                case "�W��":
                                                    value = eventItem.Summary;
                                                    break;
                                                case "�}�l�ɶ�":
                                                    value = eventItem.Start.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.Start.Date;
                                                    break;
                                                case "�����ɶ�":
                                                    value = eventItem.End.DateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? eventItem.End.Date;
                                                    break;
                                                case "�a�I":
                                                    value = eventItem.Location;
                                                    break;
                                                case "���ʻ���":
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
                                MessageBox.Show($"���\�ץX {SpcCount} ���ƥ�� {filePath}", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("�S�����ŦX���󪺨ƥ�C", "����", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    btnExport.Enabled = true;
                    btnExport.Text = "�ץX Excel";
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
                    // �M�� ComboBox �òK�[�s���
                    comboBoxCalendars.Items.Clear();
                    foreach (var calendar in calendars)
                    {
                        // �N���W�٩M ID �]�ˬ� CalendarItem ����
                        comboBoxCalendars.Items.Add(new CalendarItem(calendar.Summary, calendar.Id));
                    }

                    if (comboBoxCalendars.Items.Count > 0)
                        comboBoxCalendars.SelectedIndex = 0; // �w�]�襤�Ĥ@��
                    comboBoxCalendars.Enabled = true;
                    dateTimePickerStart.Enabled = true;
                    dateTimePickerEnd.Enabled = true;
                    Rbt_Default.Enabled = true;
                    Rbt_Spec.Enabled = true;
                    Chk_DefaultRules.Enabled = true;
                    //Tbx_SpecialRule.Enabled = true;
                    btnExport.Enabled = true;
                    button1.Enabled = true;
                    MessageBox.Show($"���\���J {calendars.Count} �Ӥ��", "���\", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Btn_Login.Enabled = false;
                }
                else
                {
                    MessageBox.Show("�ӱb��S����L���C", "����", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"���J���ɵo�Ϳ��~: {ex.Message}", "���~", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // �ѪR�S��W�h
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
            string credPath = "token.json"; // OAuth Token �x�s
            if (Directory.Exists(credPath))
            {
                Directory.Delete(credPath, true); // true �N���j�R��
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
