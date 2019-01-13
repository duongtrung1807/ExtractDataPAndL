using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExportDataForPAndL02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filename = "";
        List<string> listFromSheet = new List<string>();
        List<string> listToSheet = new List<string>();
        List<string> listchannel = new List<string>();
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                listFromSheet.Clear();
                ccbListMonth.Items.Clear();

                if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    filename = openFileDialog1.FileName;
                    textBox1.Text = filename;
                    //Create a test file
                    var fi = new FileInfo(filename);
                    using (var package = new ExcelPackage(fi))
                    {
                        var workbook = package.Workbook;
                        listFromSheet.Clear();
                        foreach (var item in workbook.Worksheets)
                        {
                            listFromSheet.Add(item.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        private List<string> months = new List<string>()
        {
            "Jan",
            "Feb",
            "Mar",
            "Apr",
            "May",
            "Jun",
            "Jul",
            "Aug",
            "Sep",
            "Oct",
            "Nov",
            "Dec"
        };

        public class MonthModel
        {
            public string Month { get; set; }
            public int Column { get; set; }
        }
        public class PSIModel
        {
            public string Month { get; set; }
            public string Channel { get; set; }
            public string ModelNo { get; set; }
            public string Status { get; set; }
            public int Qty { get; set; }
        }
        string filename2 = "";
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                listToSheet.Clear();
                ccbListMonth.Items.Clear();
                if (openFileDialog2.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    filename2 = openFileDialog2.FileName;
                    textBox2.Text = filename2;
                    //Create a test file
                    var fi = new FileInfo(filename2);
                    using (var package = new ExcelPackage(fi))
                    {
                        var workbook = package.Workbook;
                        listToSheet.Clear();
                        foreach (var item in workbook.Worksheets)
                        {
                            listToSheet.Add(item.Name);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error");
            }
        }

        public class ExportModel
        {
            public int row { get; set; }
            public string model { get; set; }
            public string channel { get; set; }
            public string status { get; set; }
        }

        public class ListMonth
        {
            public int col { get; set; }
            public string month { get; set; }
            public string monthshort { get; set; }
        }

        public class FromMonthModel
        {
            public int Column { get; set; }
            public string Month { get; set; }
            public string Status { get; set; }
        }

        private List<FromMonthModel> listFromMonth = new List<FromMonthModel>();
        private List<ExportModel> listExportModel = new List<ExportModel>();
        private List<ListMonth> listMonth = new List<ListMonth>();
        private void button3_Click(object sender, EventArgs e)
        {
            ccbListMonth.Items.Clear();
            if (listFromSheet.Count > -1 && listToSheet.Count > -1)
            {
                //get PSI Data
                var fi2 = new FileInfo(filename2);
                using (var package = new ExcelPackage(fi2))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[1];
                    var row = 4;
                    var col = 8;
                    while (worksheet.Cells[row, col].Text.Trim() != "")
                    {
                        listMonth.Add(new ListMonth()
                        {
                            col = col,
                            month = worksheet.Cells[row, col].Text.Trim(),
                            monthshort = worksheet.Cells[row, col].Text.Trim().Substring(0, 3).ToLower()
                        });
                        col += 1;
                    }
                    row = 5;
                    col = 5;
                    while (worksheet.Cells[row, col].Text.Trim() != "")
                    {
                        listExportModel.Add(new ExportModel()
                        {
                            row = row,
                            model = worksheet.Cells[row, col].Text.ToLower().Trim(),
                            channel = worksheet.Cells[row, col + 1].Text.ToLower().Trim(),
                            status = worksheet.Cells[row, col + 2].Text.ToLower().Trim(),
                        });
                        row += 1;
                    }
                }
                ccbListMonth.Items.Clear();
                foreach (var item in listMonth)
                {
                    ComboboxItem item2 = new ComboboxItem();
                    item2.Text = item.month;
                    item2.Value = item.col;
                    ccbListMonth.Items.Add(item2);
                }
                listchannel.Clear();
                listchannel = listExportModel.Select(o => o.channel.ToLower().Trim()).Distinct().ToList();
                MessageBox.Show("Processing Completed!", "Information");
            }
            else
            {
                MessageBox.Show("File has not Sheet!", "Error");
            }
        }

        public class ComboboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }
            public override string ToString()
            {
                return Text;
            }
        }

        private void ccbListMonth_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        List<Model1> listmatch = new List<Model1>()
        {
            new Model1(){SheetName = "North KK", Channel ="KK North"},
            new Model1(){SheetName = "DMX", Channel ="DMX South"},
            new Model1(){SheetName = "266", Channel ="WS South"},
            new Model1(){SheetName = "North MM", Channel ="M/M North"},
            new Model1(){SheetName = "North WS", Channel ="WS North"},
            new Model1(){SheetName = "South MM Other", Channel ="M/M South"},
            new Model1(){SheetName = "South KK", Channel ="KK South"}
        };
        List<PSIModel> listPSIData = new List<PSIModel>();
        private void button4_Click(object sender, EventArgs e)
        {
            var check = false;
            var firstsheet = "";
            foreach (var item in listFromSheet)
            {
                foreach (var item2 in listmatch)
                {
                    if (item.ToLower().Trim() == item2.SheetName.ToLower().Trim())
                    {
                        check = true;
                        firstsheet = item2.SheetName;
                        break;
                    }
                }
                if (check) break;
            }
            if (check && ccbListMonth.SelectedIndex > 0)
            {
                var columnstart = Convert.ToInt32((ccbListMonth.SelectedItem as ComboboxItem).Value);
                var strmonth = (ccbListMonth.SelectedItem as ComboboxItem).Text.ToString().Trim();
                var substr = ccbListMonth.SelectedItem.ToString().Substring(0, 3);
                //Lay du lieu PSI

                listPSIData.Clear();
                var fi1 = new FileInfo(filename);
                //Lay du lieu file nguon
                listFromMonth.Clear();
                var fi = new FileInfo(filename);
                using (var package = new ExcelPackage(fi))
                {
                    var workbook = package.Workbook;
                    var worksheet = workbook.Worksheets[firstsheet];
                    var j = 7;
                    var tempmonth = worksheet.Cells[9, j].Text;
                    while (months.Any(l => l.ToLower().Trim() == tempmonth.ToLower().Trim()))
                    {
                        tempmonth = worksheet.Cells[9, j].Text;
                        var tempstatus = worksheet.Cells[10, j].Text;
                        if (tempstatus == "P" || tempstatus == "S")
                        {
                            if (!listFromMonth.Contains(new FromMonthModel() { Month = tempmonth, Status = tempstatus, Column = j }))
                                listFromMonth.Add(new FromMonthModel() { Month = tempmonth, Status = tempstatus, Column = j });
                        }
                        j += 1;
                    }
                }

                var monthstart = 0;
                while (monthstart <= listFromMonth.Count() && substr != listFromMonth[monthstart].Month) monthstart += 1;
                using (var package = new ExcelPackage(fi1))
                {
                    foreach (var sheet in listFromSheet)
                    {
                        foreach (var item in listmatch)
                        {
                            if (sheet.ToLower().Trim() == item.SheetName.ToLower().Trim())
                            {
                                var workbook = package.Workbook;
                                var worksheet = workbook.Worksheets[sheet];
                                for (int ii = monthstart; ii < listFromMonth.Count; ii++)
                                {
                                    var k = 11;
                                    var column = listFromMonth[ii].Column;
                                    while (worksheet.Cells[k, 1].Text.ToString() != "")
                                    {
                                        var tempStr = worksheet.Cells[k, column].Text.ToString().Replace(",", "").Replace("-", "").Replace(" ", "");
                                        var tempQty = tempStr == "" ? 0 : Convert.ToInt32(tempStr);
                                        listPSIData.Add(new PSIModel()
                                        {
                                            Month = listFromMonth[ii].Month.ToLower().Trim(),
                                            Status = listFromMonth[ii].Status.ToLower().Trim(),
                                            Channel = listmatch.FirstOrDefault(p => p.SheetName == sheet).Channel.ToLower().Trim(),
                                            ModelNo = worksheet.Cells[k, 4].Text.ToString().ToLower().Trim(),
                                            Qty = tempQty
                                        });
                                        k += 1;
                                    }
                                }
                            }
                        }
                    }
                }

                try
                {
                    if (months.Where(o => string.Equals(substr, o, StringComparison.OrdinalIgnoreCase)).Any())
                    {
                        var month = months.IndexOf(substr) + 1;
                        if (month > 0)
                        {
                            var fi2 = new FileInfo(filename2);
                            var listmodelno = listPSIData.Select(o => o.ModelNo).Distinct().ToList();
                            using (var package = new ExcelPackage(fi2))
                            {
                                var workbook = package.Workbook;
                                var worksheet = workbook.Worksheets[1];
                                foreach (var item in listExportModel)
                                {
                                    if (listmatch.Any(l => l.Channel.ToLower().Trim() == item.channel.ToLower().Trim()) && listmodelno.Contains(item.model))
                                    {
                                        var lastcolumn = listMonth[listMonth.Count - 1].col;
                                        for (int i = columnstart; i < lastcolumn; i++)
                                        {
                                            var tempmonth = listMonth.FirstOrDefault(o => o.col == i).monthshort;
                                            if (listPSIData.Exists(o => o.Channel == item.channel && o.ModelNo == item.model && o.Status == item.status && o.Month == tempmonth))
                                            {
                                                var tempQty = listPSIData.FirstOrDefault(o => o.Channel == item.channel && o.ModelNo == item.model && o.Status == item.status && o.Month == tempmonth).Qty;
                                                worksheet.Cells[item.row, i].Value = tempQty;
                                            }
                                            else
                                            {
                                                worksheet.Cells[item.row, i].Value = 0;
                                            }
                                        }
                                    }
                                }
                                package.Save();
                            }
                            ccbListMonth.Items.Clear();
                            MessageBox.Show("Transport Data Completed!", "Information");
                        }
                        else
                        {
                            MessageBox.Show("Month is not number!", "Error");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error");
                }
            }
            else
            {
                MessageBox.Show("Data are not matched between two file!", "Error");
            }
        }
    }
}
