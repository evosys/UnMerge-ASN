using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Win32;

/* To work with EPPlus library */
using OfficeOpenXml;
using OfficeOpenXml.Drawing;


namespace UnMerge_ASN_Report
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            Title = "UnMerge ASN Report";

            // Init value of control
            App.Current.Properties["pathFile"] = null;
            App.Current.Properties["nameFile"] = null;
            App.Current.Properties["destFile"] = null;
        }

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {

            OpenFileDialog openFileDialogOP = new OpenFileDialog()
            {
                Filter = "XLSX Files|*.xlsx",
                Title = "Select a XLSX File",
                FilterIndex = 1,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            };

            // Show the Dialog..  
            if (openFileDialogOP.ShowDialog() == true)
            {
                App.Current.Properties["pathFile"] = openFileDialogOP.FileName;
                App.Current.Properties["nameFile"] = openFileDialogOP.SafeFileName;

                if (App.Current.Properties["pathFile"] != null)
                {
                    string curnameFile = App.Current.Properties["nameFile"].ToString();
                    string curpathFile = System.IO.Path.GetDirectoryName(App.Current.Properties["pathFile"].ToString());
                    txtOpen.Text = curnameFile;
                    // txtDest.Text = curpathFile;
                    txtOpen.Foreground = new SolidColorBrush(Colors.Blue);

                    App.Current.Properties["destFile"] = curpathFile + "\\out-" + curnameFile;
                    txtDest.Text = App.Current.Properties["destFile"].ToString();
                }
            }

        }
        private void BtnDest_Click(object sender, RoutedEventArgs e)
        {
            string fileName;

            if (App.Current.Properties["nameFile"] != null)
            {
                fileName = "out-" + App.Current.Properties["nameFile"].ToString();
            }
            else
            {
                fileName = "";
            }

            SaveFileDialog saveFileDialogDT = new SaveFileDialog()
            {
                Filter = "XLSX Files|*.xlsx",
                Title = "Save a XLSX File",
                FileName = fileName
            };

            // Show the Dialog..  
            if (saveFileDialogDT.ShowDialog() == true)
            {
                // string TMPcurpathFile = System.IO.Path.GetDirectoryName(saveFileDialogDT.SafeFileName);
                string TMPcurnameFile = saveFileDialogDT.FileName;

                App.Current.Properties["destFile"] = TMPcurnameFile;
                txtDest.Text = App.Current.Properties["destFile"].ToString();
            }
        }

        private void BtnProc_Click(object sender, RoutedEventArgs e)
        {

            if(String.IsNullOrEmpty(txtOpen.Text))
            {
                MessageBox.Show("Please select file first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
            }
            else
            {
                // ExcelAddress opened file path
                var FilePath = App.Current.Properties["pathFile"].ToString();

                List<string> tabHeadList = new List<string>();
                List<string> ret = new List<string>();

                var headDict = new Dictionary<string, string>();

                if (txtOpen.Text != "")
                {
                    if (txtDest.Text != "")
                    {
                        var OutPath = txtDest.Text;

                        // delete file if exist
                        if (File.Exists(OutPath))
                        {
                            if (!FileInUse(OutPath))
                            {
                                File.Delete(OutPath);
                            }
                            else
                            {
                                MessageBox.Show("File: " + OutPath + System.Environment.NewLine +"Is used by another process, please close before continue.", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
                                return;
                            }

                        }

                        if (cbUnMer.IsChecked ?? true)
                        {
                            ProgressIndicator.IsBusy = true;

                            Task.Factory.StartNew(() =>
                            {
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Please wait..");
                                    }
                                ));

                                Thread.Sleep(1000);

                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Get header information..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            // Variabel for header
                                            // Get header info
                                            headDict["CompanyName"] = ws.Cells[2, 2].Value.ToString();
                                            headDict["ReportName"] = ws.Cells[4, 2].Value.ToString();
                                            headDict["PrintDate"] = ws.Cells[6, 3].Value.ToString();
                                            headDict["PrintDateVal"] = ws.Cells[6, 6].Value.ToString();
                                            // formating datetime
                                            double pdv = double.Parse(headDict["PrintDateVal"]);
                                            headDict["PrintDateVal"] = DateTime.FromOADate(pdv).ToString("dd/MM/yyyy hh:mm tt");
                                            headDict["PageNo"] = ws.Cells[6, 9].Value.ToString();
                                            headDict["PageNoVal"] = ws.Cells[6, 12].Value.ToString();
                                            headDict["FromDate"] = ws.Cells[8, 3].Value.ToString();
                                            headDict["FromDateVal"] = ws.Cells[8, 6].Value.ToString();
                                            // formating datetime
                                            double fdv = double.Parse(headDict["FromDateVal"]);
                                            headDict["FromDateVal"] = DateTime.FromOADate(fdv).ToString("dd/MM/yyyy");
                                            headDict["ToDate"] = ws.Cells[8, 9].Value.ToString();
                                            headDict["ToDateVal"] = ws.Cells[8, 12].Value.ToString();
                                            // formating datetime
                                            double tdv = double.Parse(headDict["ToDateVal"]);
                                            headDict["ToDateVal"] = DateTime.FromOADate(tdv).ToString("dd/MM/yyyy");
                                            headDict["ShipmentNo"] = ws.Cells[8, 17].Value.ToString();
                                            headDict["ShipmentNoVal"] = ws.Cells[8, 20].Value.ToString();
                                            headDict["SONumber"] = ws.Cells[7, 25].Value.ToString();
                                            headDict["SONumberVal"] = ws.Cells[7, 29].Value.ToString();
                                            headDict["InvNumber"] = ws.Cells[7, 32].Value.ToString();
                                            headDict["InvNumberVal"] = ws.Cells[7, 36].Value.ToString();
                                            headDict["Status"] = ws.Cells[7, 40].Value.ToString();
                                            headDict["StatusVal"] = ws.Cells[7, 43].Value.ToString();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing header..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {

                                            var wb = package.Workbook;
                                            // ExcelWorksheet ws = wb.Worksheets.Add("UID ASN Report");

                                            //Check if worksheet with name "Content" exists and retrieve that instance or null if it doesn't exist       
                                            ExcelWorksheet ws = wb.Worksheets.FirstOrDefault(x => x.Name == "UID ASN Report");

                                            //If worksheet "Content" was not found, add it
                                            if (ws == null)
                                            {
                                                ws = wb.Worksheets.Add("UID ASN Report");
                                            }
                                            else
                                            {
                                                ws = wb.Worksheets["UID ASN Report"];
                                            }

                                            ws.View.ShowGridLines = false;

                                            // write header
                                            ws.Cells["A2:M2"].Merge = true;
                                            ws.Cells["A2"].Style.Font.Bold = true;
                                            ws.Cells["A2"].Style.Font.Size = 20;
                                            ws.Cells["A2"].Style.Indent = 3;
                                            ws.Cells["A2"].Value = headDict["CompanyName"];

                                            // report name
                                            ws.Cells["A3:C3"].Merge = true;
                                            ws.Cells["A3"].Style.Font.Bold = true;
                                            ws.Cells["A3"].Style.Font.Size = 12;
                                            ws.Cells["A3"].Style.Indent = 3;
                                            ws.Cells["A3"].Value = headDict["ReportName"];

                                            // additional info

                                            // Print Date
                                            ws.Cells["A5:D5"].Merge = true;
                                            ws.Cells["A5:D5"].Style.Indent = 3;
                                            ws.Cells["A5:D5"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText PrintDate = ws.Cells["A5"].RichText.Add(headDict["PrintDate"]);
                                            PrintDate.Bold = true;
                                            PrintDate.Size = 9;

                                            PrintDate = ws.Cells["A5"].RichText.Add(" " + headDict["PrintDateVal"]);
                                            PrintDate.Bold = false;

                                            // Page No
                                            ws.Cells["E5:F5"].Merge = true;
                                            ws.Cells["E5"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText PageNo = ws.Cells["E5"].RichText.Add(headDict["PageNo"]);
                                            PageNo.Bold = true;
                                            PageNo.Size = 9;

                                            PageNo = ws.Cells["E5"].RichText.Add(" " + headDict["PageNoVal"]);
                                            PageNo.Bold = false;

                                            // From Date
                                            ws.Cells["A6:D6"].Merge = true;
                                            ws.Cells["A6:D6"].Style.Indent = 3;
                                            ws.Cells["A6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText FormDate = ws.Cells["A6"].RichText.Add(headDict["FromDate"]);
                                            FormDate.Bold = true;
                                            FormDate.Size = 9;

                                            FormDate = ws.Cells["A6"].RichText.Add(" " + headDict["FromDateVal"]);
                                            FormDate.Bold = false;

                                            // To Date
                                            ws.Cells["E6:F6"].Merge = true;
                                            ws.Cells["E6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText ToDate = ws.Cells["E6"].RichText.Add(headDict["ToDate"]);
                                            ToDate.Bold = true;
                                            ToDate.Size = 9;

                                            ToDate = ws.Cells["E6"].RichText.Add(" " + headDict["ToDateVal"]);
                                            ToDate.Bold = false;

                                            // Shipment
                                            ws.Cells["H6:J6"].Merge = true;
                                            ws.Cells["H6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText Shipment = ws.Cells["H6"].RichText.Add(headDict["ShipmentNo"]);
                                            Shipment.Bold = true;
                                            Shipment.Size = 9;

                                            Shipment = ws.Cells["H6"].RichText.Add(" " + headDict["ShipmentNoVal"]);
                                            Shipment.Bold = false;

                                            // SO Number
                                            ws.Cells["L6:N6"].Merge = true;
                                            ws.Cells["L6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText SONumber = ws.Cells["L6"].RichText.Add(headDict["SONumber"]);
                                            SONumber.Bold = true;
                                            SONumber.Size = 9;

                                            SONumber = ws.Cells["L6"].RichText.Add(" " + headDict["SONumberVal"]);
                                            SONumber.Bold = false;

                                            // Invoice Number
                                            ws.Cells["P6:R6"].Merge = true;
                                            ws.Cells["P6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText InvNumber = ws.Cells["P6"].RichText.Add(headDict["InvNumber"]);
                                            InvNumber.Bold = true;
                                            InvNumber.Size = 9;

                                            InvNumber = ws.Cells["P6"].RichText.Add(" " + headDict["InvNumberVal"]);
                                            InvNumber.Bold = false;

                                            // Status
                                            ws.Cells["S6:U6"].Merge = true;
                                            ws.Cells["S6"].IsRichText = true;
                                            OfficeOpenXml.Style.ExcelRichText Status = ws.Cells["S6"].RichText.Add(headDict["Status"]);
                                            Status.Bold = true;
                                            Status.Size = 9;

                                            Status = ws.Cells["S6"].RichText.Add(" " + headDict["StatusVal"]);
                                            Status.Bold = false;

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                        {
                                            ProgressIndicator.BusyContent = string.Format("Get table information..");

                                            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
                                            {
                                                var wb = package.Workbook;
                                                ExcelWorksheet ws = wb.Worksheets[1];

                                                var RowCount = ws.Dimension.End.Row;
                                                var ColCount = ws.Dimension.End.Column;

                                                for (int col = 1; col <= ColCount; col++)
                                                {
                                                    string cellValue = ws.Cells[11, col].Text; // This got me the actual value I needed.
                                                    tabHeadList.Add(cellValue);
                                                }

                                                // remove blank value
                                                tabHeadList = tabHeadList.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
                                            }
                                        }
                                ));

                                Thread.Sleep(1000);

                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing header table..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            ws.Cells[8, 1].LoadFromArrays(new List<string[]>(new[] { tabHeadList.ToArray() }));
                                            int countAr = tabHeadList.Count() + 1;
                                            // ws.Cells[8, 1, 8, countAr].Style.Font.Bold = true;
                                            // ws.Cells[8, 1, 8, countAr].Style.Font.Size = 9;
                                            // ws.Cells[8, 1, 8, countAr].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);

                                            using (ExcelRange Rng = ws.Cells[8, 1, 8, countAr])
                                            {
                                                Rng.Style.Font.Size = 9;
                                                Rng.Style.Font.Bold = true;
                                            }

                                            for (int i = 1; i < countAr; i++)
                                            {
                                                ws.Cells[8, i].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.FromArgb(191, 191, 191));
                                            }

                                            ws.Cells[8, countAr, 8, countAr + 2].Merge = true;
                                            ws.Cells[8, countAr, 8, countAr + 2].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.FromArgb(191, 191, 191));

                                            ws.View.FreezePanes(9, 1);
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // kode distributor
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row kode distributor..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 1;

                                            List<String> resVal = GetMergeRangeString(FilePath, 1, false);

                                            int countRow = resVal.Count() + startRow + 2;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;

                                            var modelTable = ws.Cells[startRow, startCol, countRow, 21];

                                            // Assign borders
                                            modelTable.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            modelTable.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            modelTable.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                            modelTable.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                            modelTable.Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(191, 191, 191));
                                            modelTable.Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(191, 191, 191));
                                            modelTable.Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(191, 191, 191));
                                            modelTable.Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(191, 191, 191));

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // no.pengiriman
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row no.pengiriman..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 2;

                                            List<String> resVal = GetMergeRangeString(FilePath, 4, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // no.SO
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row no.SO..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 3;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 7, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // no.DO
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row no.DO..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 4;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 10, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // TGL. Pengiriman
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row tgl.pengiriman..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 5;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 13, false);

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                string tmpRes = resVal[i];
                                                double resDouble = double.Parse(tmpRes);
                                                var tmpDate = DateTime.FromOADate(resDouble).ToString("dd/MM/yyyy");

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = tmpDate;
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // TGL.Perkiraan
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row tgl.perkiraan..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 6;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 15, false);

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                string tmpRes = resVal[i];
                                                double resDouble = double.Parse(tmpRes);
                                                var tmpDate = DateTime.FromOADate(resDouble).ToString("dd/MM/yyyy");

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = tmpDate;
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // TGL.Terima
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row tgl.terima..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 7;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 18, false);

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                string tmpRes = resVal[i];
                                                double resDouble = double.Parse(tmpRes);
                                                var tmpDate = DateTime.FromOADate(resDouble).ToString("dd/MM/yyyy");

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = tmpDate;
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Lead Time
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row lead time..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 8;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 21, false);
                                            List<Double> intList = resVal.Select(s => Convert.ToDouble(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = intList[i];
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Faktur
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row faktur..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 9;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 23, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // PCode
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row pcode..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 10;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 26, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // nama barang
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row nama barang..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 11;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 27, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Batch
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row batch..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 12;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 30, false);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Unit
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row unit..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 13;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 33, true);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Price
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row price..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 14;

                                            var lastRow = ws.Dimension.End.Row - 3;


                                            List<String> resVal = GetMergeRangeString(FilePath, 34, false);
                                            List<Double> intList = resVal.Select(s => Convert.ToDouble(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            // ws.Cells[startRow, startCol, countRow, startCol].Style.Numberformat.Format = "#.##0,0";
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                if (intList[i] == 0)
                                                {
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "0;-0;-;@";
                                                }
                                                else
                                                {
                                                    // write to cell
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "#,##0.00";
                                                }
                                            }

                                            package.Save();

                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Dispatch QTY
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row dispatch qty..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 15;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 37, true);
                                            List<int> intList = resVal.Select(s => Convert.ToInt32(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = intList[i];
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Shipped QTY
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row shipped qty..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 16;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 38, true);
                                            List<int> intList = resVal.Select(s => Convert.ToInt32(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = intList[i];
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Loss QTY
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row lost qty..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 17;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 41, true);
                                            List<int> intList = resVal.Select(s => Convert.ToInt32(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = intList[i];
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Received QTY
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row received qty..");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 18;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 44, true);
                                            List<int> intList = resVal.Select(s => Convert.ToInt32(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                // write to cell
                                                ws.Cells[startRow, startCol].Value = intList[i];
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Finishing..
                                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row data...");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 19;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 45, true);

                                            int countRow = resVal.Count() + startRow;

                                            ws.Cells[startRow, startCol].LoadFromCollection(resVal);
                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Finishing..
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row data...");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 20;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 46, true);
                                            List<Double> intList = resVal.Select(s => Convert.ToDouble(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                if (intList[i] < 0)
                                                {
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "[$-10809]#,##0.00;(#,##0.00);\"-\"";
                                                }
                                                else
                                                {
                                                    // write to cell
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "#,##0.00";
                                                }
                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Finishing..
                                    Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row data...");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            int startRow = 9;
                                            int startCol = 21;

                                            var lastRow = ws.Dimension.End.Row - 3;

                                            List<String> resVal = GetMergeRangeString(FilePath, 47, true);
                                            List<Double> intList = resVal.Select(s => Convert.ToDouble(s)).ToList();

                                            int countRow = resVal.Count() + startRow;
                                            int indList = resVal.Count();

                                            ws.Cells[startRow, startCol, countRow, startCol].Style.Font.Size = 9;

                                            for (int i = 0; i < indList; i++)
                                            {
                                                startRow = 9 + i;

                                                if (intList[i] < 0)
                                                {
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "#,##0;(#,##0.00)";
                                                }
                                                else
                                                {
                                                    // write to cell
                                                    ws.Cells[startRow, startCol].Value = intList[i];
                                                    ws.Cells[startRow, startCol].Style.Numberformat.Format = "#,##0.00";
                                                }

                                            }

                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                                // Formating total..
                                Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                    {
                                        ProgressIndicator.BusyContent = string.Format("Writing row total...");

                                        using (ExcelPackage package = new ExcelPackage(new FileInfo(OutPath)))
                                        {
                                            var wb = package.Workbook;
                                            ExcelWorksheet ws = wb.Worksheets[1];

                                            List<String> resVal = GetTotal(FilePath);

                                            int lastRow = ws.Dimension.End.Row - 3;

                                            var numAlpha = new Regex("(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)");
                                            var match = numAlpha.Match(resVal[1]);

                                            string sth = String.Concat("A", lastRow);
                                            string eth = String.Concat("L", lastRow + 2);
                                            string bth = String.Concat("U", lastRow + 2);
                                            string MrgRng = sth + ":" + eth;
                                            string BoldRng = sth + ":" + bth;

                                            var alpha = match.Groups["Alpha"].Value;
                                            var num = Int32.Parse(match.Groups["Numeric"].Value);

                                            // Write total
                                            ws.Cells[MrgRng].Merge = true;
                                            ws.Cells[sth].Value = resVal[0];
                                            ws.Cells[BoldRng].Style.Font.Bold = true;
                                            ws.Cells[sth].Style.Font.Size = 9;
                                            ws.Cells[MrgRng].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin, System.Drawing.Color.FromArgb(191, 191, 191));
                                            ws.Cells[MrgRng].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                            ws.Cells.AutoFitColumns();
                                            package.Save();
                                        }
                                    }
                                ));

                                Thread.Sleep(1000);

                            }
                            ).ContinueWith((task) =>
                            {
                                ProgressIndicator.IsBusy = false;

                                MessageBox.Show("Process success", "Success", MessageBoxButton.OK, MessageBoxImage.Information, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);

                                if (File.Exists(OutPath))
                                {
                                    System.Diagnostics.Process.Start("explorer.exe", "/select, " + OutPath);
                                }

                            }, TaskScheduler.FromCurrentSynchronizationContext());


                        }
                        else
                        {
                            MessageBox.Show("Please select options first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please choose destination of file", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                else
                {
                    MessageBox.Show("Please select file first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error, MessageBoxResult.None, MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            

        }

        private void BtnProc_Click1(object sender, RoutedEventArgs e)
        {
            // ExcelAddress opened file path
            var FilePath = App.Current.Properties["pathFile"].ToString();

            List<string> tabHeadList = new List<string>();
            List<string> ret = new List<string>();

            var headDict = new Dictionary<string, string>();


            if (txtOpen.Text != "")
            {
                if (txtDest.Text != "")
                {
                    var OutPath = txtDest.Text;

                    // delete file if exist
                    if (File.Exists(OutPath))
                    {
                        File.Delete(OutPath);
                    }

                    if (cbUnMer.IsChecked ?? true)
                    {
                        ProgressIndicator.IsBusy = true;

                        Task.Factory.StartNew(() =>
                        {
                            Dispatcher.Invoke(DispatcherPriority.Normal, new Action(() =>
                                {
                                    ProgressIndicator.BusyContent = string.Format("Please wait..");

                                    using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
                                    {
                                        var wb = package.Workbook;
                                        ExcelWorksheet ws = wb.Worksheets[1];

                                        // int startRow = 9;
                                        // int startCol = 14;



                                        // List<String> resVal = GetMergeRangeString(FilePath, 34);

                                        List<String> resVal = GetTotal(FilePath);

                                        int lastRow = ws.Dimension.End.Row + 3;

                                        var numAlpha = new Regex("(?<Alpha>[a-zA-Z]*)(?<Numeric>[0-9]*)");
                                        var match = numAlpha.Match(resVal[1]);

                                        var alpha = match.Groups["Alpha"].Value;
                                        var num = Int32.Parse(match.Groups["Numeric"].Value);

                                        Console.WriteLine(num);
                                        Console.WriteLine(resVal[0]);
                                        Console.WriteLine(resVal[1]);
                                    }
                                }
                            )
                        );

                            Thread.Sleep(1000);

                        }
                        ).ContinueWith((task) =>
                            {
                                ProgressIndicator.IsBusy = false;
                            }, TaskScheduler.FromCurrentSynchronizationContext()
                        );
                    }
                    else
                    {
                        MessageBox.Show("Please select options first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    }
                }
                else
                {
                    MessageBox.Show("Please choose destination of file");
                }
            }
            else
            {
                MessageBox.Show("Please select file first!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }

        }

        private bool UnMerge()
        {
            return true;
        }



        // get range of string row
        static List<String> GetMergeRangeString(string FilePath, int col, bool fullRow)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                var result = new List<String>();

                int lastRow;

                var wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[1];

                if (fullRow)
                {
                    lastRow = ws.Dimension.End.Row;
                }
                else
                {

                    lastRow = ws.Dimension.End.Row - 3;
                }

                for (int i = 12; i <= lastRow; i++)
                {
                    var myCell = ws.MergedCells[i, col];

                    if (string.IsNullOrEmpty(myCell))
                    {
                        string tmpRes = ws.Cells[i, col].Value.ToString();
                        result.Add(tmpRes);
                    }
                    else
                    {
                        if (myCell.Contains(':'))
                        {
                            string tmpRes = myCell.Substring(0, myCell.LastIndexOf(':'));
                            tmpRes = ws.Cells[tmpRes].Value.ToString();
                            result.Add(tmpRes);
                        }

                    }

                }

                return result;
            }

        }

        // get range of string row
        static List<String> GetTotal(string FilePath)
        {
            var result = new List<String>();

            using (ExcelPackage package = new ExcelPackage(new FileInfo(FilePath)))
            {
                var wb = package.Workbook;
                ExcelWorksheet ws = wb.Worksheets[1];

                int lastRow = ws.Dimension.End.Row - 2;

                var myCell = ws.MergedCells[lastRow, 1];

                result.Add(ws.Cells[lastRow, 1].Value.ToString());

                if (myCell.Contains(':'))
                {
                    string tmpRes = myCell.Substring(0, myCell.LastIndexOf(':'));
                    result.Add(tmpRes);
                }
            }

            return result;
        }

        // check file if opened
        static bool FileInUse(string path)
        {
            try
            {
                using (FileStream fs = new FileStream(path, FileMode.OpenOrCreate))
                {
                    return false;
                }
                //return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        // check variabel null or empty
        static bool IsEmpty<T>(List<T> @this)
        {
            return @this == null || @this.Count == 0;
        }
    }
}