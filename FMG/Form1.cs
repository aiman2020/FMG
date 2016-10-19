using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FMG
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Excel.Application xlApp;
        Excel.Workbook xlWorkBookCurrent;
        Excel.Worksheet xlWorkSheetCurrent;
        Excel.Range range;

        string workbookPath;
        string fileName = string.Empty; //@"\AmmunitionChart.xlsx";
        string fileName2 = @"\myFile.xlsx";

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName;
            fileName = openFileDialog1.FileName;
            progressBar1.Value = 0;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnProcess_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "")
                MessageBox.Show("الرجاء اختيار الملف");
            else
            {
                try
                {
                    xlApp = new Excel.Application();
                    if (xlApp == null)
                    {
                        MessageBox.Show("Excel is not properly installed!!");
                        return;
                    }
                    workbookPath = System.Windows.Forms.Application.StartupPath + fileName;
                    xlWorkBookCurrent = xlApp.Workbooks.Open(fileName, 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    xlWorkSheetCurrent = (Excel.Worksheet)xlWorkBookCurrent.Worksheets.get_Item(1);
                    range = xlWorkSheetCurrent.UsedRange;
                    progressBar1.Maximum = range.Rows.Count - 1;
                    //generateExcel(xlWorkBookCurrent, xlWorkSheetCurrent);
                    SummaryOfOvertime();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }


            }
        }

        //private void generateExcel(Excel.Workbook xlWorkBookCurrent, Excel.Worksheet xlWorkSheetCurrent)
        //{
        //    //string workbookPath = System.Windows.Forms.Application.StartupPath + fileName;

        //    try
        //    {
        //        //Excel.Application xlApp = new Excel.Application();
        //        //if (xlApp == null)
        //        //{
        //        //    MessageBox.Show("Excel is not properly installed!!");
        //        //    return;
        //        //}

        //        //Excel.Workbook xlWorkBookCurrent;
        //        //Excel.Worksheet xlWorkSheetCurrent;
        //        //Excel.Range range;




        //        Excel.Workbook xlWorkBookNew;
        //        Excel.Worksheet xlWorkSheetNew;
        //        object misValue = System.Reflection.Missing.Value;

        //        xlWorkBookNew = xlApp.Workbooks.Add(misValue);
        //        //xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
        //        xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.Add(misValue, misValue, misValue, misValue);
        //        xlWorkSheetCurrent.UsedRange.Copy(Type.Missing);
        //        xlWorkSheetNew.Paste(Type.Missing, Type.Missing);

        //        Clipboard.Clear();
        //        //xlWorkSheetNew = xlWorkSheetCurrent;
        //        //xlWorkSheet2.Name = "OT";
        //        Excel.Range formatRange;
        //        formatRange = xlWorkSheetNew.Cells;
        //        formatRange.NumberFormat = "@";
        //        xlWorkSheetNew.Cells[1, "AD"].Value2 = "OT Corrected";
        //        range = xlWorkSheetNew.UsedRange;
        //        TimeSpan BaseTime = new TimeSpan(8, 0, 0);
        //        TimeSpan WorkingHours;
        //        TimeSpan OT;
        //        TimeSpan CorrectOT;
        //        DateTime WorkingHours2;
        //        DateTime OT2;

        //        for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
        //        {
        //            string s1 = xlWorkSheetNew.Cells[rCnt, "R"].Value2;
        //            string s2 = xlWorkSheetNew.Cells[rCnt, "Q"].Value2;

        //            if (s1 != "" && s2 != "")
        //            {
        //                WorkingHours2 = Convert.ToDateTime(s1);
        //                OT2 = Convert.ToDateTime(s2);
        //                WorkingHours = new TimeSpan(WorkingHours2.Hour, WorkingHours2.Minute, 0);
        //                OT = new TimeSpan(OT2.Hour, OT2.Minute, 0);
        //                CorrectOT = OT - (BaseTime - WorkingHours);
        //                DateTime dt = new DateTime(CorrectOT.Ticks);
        //                string s = dt.ToString("HH:mm");
        //                //string s = CorrectOT.Hours.ToString() + ":" + CorrectOT.Minutes.ToString();
        //                xlWorkSheetNew.Cells[rCnt, "AD"].Value2 = s;
        //            }
        //        }

        //        xlWorkBookCurrent.Close(false, null, null);
        //        saveFileDialog1.ShowDialog();
        //        xlWorkBookNew.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
        //        xlWorkBookNew.Close(true, null, null);
        //        xlApp.Quit();

        //        releaseObject(xlWorkSheetCurrent);
        //        releaseObject(xlWorkBookCurrent);
        //        releaseObject(xlWorkSheetNew);
        //        releaseObject(xlWorkBookNew);
        //        releaseObject(xlApp);
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }
        //}

        private void SummaryOfOvertime()
        {
            try
            {
                Excel.Workbook xlWorkBookNew;
                Excel.Worksheet xlWorkSheetNew;
                Excel.Workbook xlWorkBookSummaryOvertime;
                Excel.Worksheet xlWorkSheetSummaryOvertime;
                Excel.Workbook xlWorkBookDailyOvertime;
                Excel.Worksheet xlWorkSheetDailyOvertime;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBookNew = xlApp.Workbooks.Add(misValue);
                //xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.Add(misValue, misValue, misValue, misValue);
                xlWorkSheetCurrent.UsedRange.Copy(Type.Missing);
                xlWorkSheetNew.Paste(Type.Missing, Type.Missing);

                xlWorkBookSummaryOvertime = xlApp.Workbooks.Add(misValue);
                xlWorkSheetSummaryOvertime = (Excel.Worksheet)xlWorkBookSummaryOvertime.Worksheets.Add(misValue, misValue, misValue, misValue);

                xlWorkBookDailyOvertime = xlApp.Workbooks.Add(misValue);
                xlWorkSheetDailyOvertime = (Excel.Worksheet)xlWorkBookDailyOvertime.Worksheets.Add(misValue, misValue, misValue, misValue);

                //xlWorkSheet2.Name = "OT";
                Excel.Range formatRange;
                formatRange = xlWorkSheetNew.Cells;
                formatRange.NumberFormat = "@";


                xlWorkSheetNew.Cells[1, "AD"].Value2 = "OT Corrected";
                xlWorkSheetNew.Cells[1, "AE"].Value2 = "Total Early";
                xlWorkSheetNew.Cells[1, "AF"].Value2 = "Total Late";

                //range = xlWorkSheetNew.UsedRange;

                //DateTime WorkingHours2;
                //DateTime OT2;
                //WorkingHours2 = Convert.ToDateTime(WorkTime).Hour;
                //OT2 = Convert.ToDateTime(OT_Time);

                string Ac_No = string.Empty;
                string Name = string.Empty;
                //TimeSpan BaseTime = TimeSpan.Zero;
                TimeSpan WorkingHours = TimeSpan.Zero;
                TimeSpan OT = TimeSpan.Zero;
                //TimeSpan CorrectOT = TimeSpan.Zero;
                TimeSpan Late = TimeSpan.Zero;
                TimeSpan Early = TimeSpan.Zero;
                TimeSpan TotalLate = TimeSpan.Zero;
                TimeSpan TotalEarly = TimeSpan.Zero;
                double Normal_OT_Total = 0;
                double Weekend_OT_Total = 0;
                double Holiday_OT_Total = 0;
                int index = 2;
                int cCnt = 2;
                double RegularOT = 0;
                double WeekEnd_OT = 0;
                double Holiday_OT = 0;

                // SummaryOvertime Header
                #region Summary Overtime Header
                xlWorkSheetSummaryOvertime.Cells[1, "A"].value2 = "AC_No";
                xlWorkSheetSummaryOvertime.Cells[1, "B"].value2 = "Name";
                xlWorkSheetSummaryOvertime.Cells[1, "C"].value2 = "NDays_OT";
                xlWorkSheetSummaryOvertime.Cells[1, "D"].value2 = "Weekend_OT";
                xlWorkSheetSummaryOvertime.Cells[1, "E"].value2 = "Holiday_OT";
                xlWorkSheetSummaryOvertime.Cells[1, "F"].value2 = "Total Weekend+Holiday OT"; 
                #endregion

                // DailyOvertime Header
                #region Daily Overtime Header
                xlWorkSheetDailyOvertime.Cells[1, "A"].value2 = "AC_No";
                xlWorkSheetDailyOvertime.Cells[1, "B"].value2 = "Name";
                xlWorkSheetDailyOvertime.Cells[1, "C"].value2 = "01";
                xlWorkSheetDailyOvertime.Cells[1, "D"].value2 = "02";
                xlWorkSheetDailyOvertime.Cells[1, "E"].value2 = "03";
                xlWorkSheetDailyOvertime.Cells[1, "F"].value2 = "04";
                xlWorkSheetDailyOvertime.Cells[1, "G"].value2 = "05";
                xlWorkSheetDailyOvertime.Cells[1, "H"].value2 = "06";
                xlWorkSheetDailyOvertime.Cells[1, "I"].value2 = "07";
                xlWorkSheetDailyOvertime.Cells[1, "J"].value2 = "08";
                xlWorkSheetDailyOvertime.Cells[1, "K"].value2 = "09";
                xlWorkSheetDailyOvertime.Cells[1, "L"].value2 = "10";
                xlWorkSheetDailyOvertime.Cells[1, "M"].value2 = "11";
                xlWorkSheetDailyOvertime.Cells[1, "N"].value2 = "12";
                xlWorkSheetDailyOvertime.Cells[1, "O"].value2 = "13";
                xlWorkSheetDailyOvertime.Cells[1, "P"].value2 = "14";
                xlWorkSheetDailyOvertime.Cells[1, "Q"].value2 = "15";
                xlWorkSheetDailyOvertime.Cells[1, "R"].value2 = "16";
                xlWorkSheetDailyOvertime.Cells[1, "S"].value2 = "17";
                xlWorkSheetDailyOvertime.Cells[1, "T"].value2 = "18";
                xlWorkSheetDailyOvertime.Cells[1, "U"].value2 = "19";
                xlWorkSheetDailyOvertime.Cells[1, "V"].value2 = "20";
                xlWorkSheetDailyOvertime.Cells[1, "W"].value2 = "21";
                xlWorkSheetDailyOvertime.Cells[1, "X"].value2 = "22";
                xlWorkSheetDailyOvertime.Cells[1, "Y"].value2 = "23";
                xlWorkSheetDailyOvertime.Cells[1, "Z"].value2 = "24";
                xlWorkSheetDailyOvertime.Cells[1, "AA"].value2 = "25";
                xlWorkSheetDailyOvertime.Cells[1, "AB"].value2 = "26";
                xlWorkSheetDailyOvertime.Cells[1, "AC"].value2 = "27";
                xlWorkSheetDailyOvertime.Cells[1, "AD"].value2 = "28";
                xlWorkSheetDailyOvertime.Cells[1, "AE"].value2 = "29";
                xlWorkSheetDailyOvertime.Cells[1, "AF"].value2 = "30";
                xlWorkSheetDailyOvertime.Cells[1, "AG"].value2 = "31";
                xlWorkSheetDailyOvertime.Cells[1, "AH"].value2 = "NDays_OT";
                xlWorkSheetDailyOvertime.Cells[1, "AI"].value2 = "Weekend_OT";
                xlWorkSheetDailyOvertime.Cells[1, "AJ"].value2 = "Holiday_OT";
                xlWorkSheetDailyOvertime.Cells[1, "AK"].value2 = "Total Weekend+Holiday OT"; 
                #endregion

                Ac_No = xlWorkSheetNew.Cells[2, "B"].Value2;
                Name = xlWorkSheetNew.Cells[2, "D"].Value2;
                xlWorkSheetSummaryOvertime.Cells[2, "A"].value2 = Ac_No;
                xlWorkSheetSummaryOvertime.Cells[2, "B"].value2 = Name;
                xlWorkSheetDailyOvertime.Cells[2, "A"].value2 = Ac_No;
                xlWorkSheetDailyOvertime.Cells[2, "B"].value2 = Name;

                for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
                {
                    string WorkTime = xlWorkSheetNew.Cells[rCnt, "R"].Value2;
                    string OT_Time = xlWorkSheetNew.Cells[rCnt, "Q"].Value2;
                    string Weekend = xlWorkSheetNew.Cells[rCnt, "AB"].Value2;
                    string Holiday = xlWorkSheetNew.Cells[rCnt, "AC"].Value2;
                    string FlagNormal = xlWorkSheetNew.Cells[rCnt, "W"].Value2;
                    string FlagWeekend = xlWorkSheetNew.Cells[rCnt, "X"].Value2;
                    string FlagHoliday = xlWorkSheetNew.Cells[rCnt, "Y"].Value2;
                    string Att_Time = xlWorkSheetNew.Cells[rCnt, "Z"].Value2;
                    //string OnDuty = xlWorkSheetNew.Cells[rCnt, "H"].Value2;
                    //string OffDuty = xlWorkSheetNew.Cells[rCnt, "I"].Value2;
                    string Late_Time = xlWorkSheetNew.Cells[rCnt, "N"].Value2;
                    string Early_Time = xlWorkSheetNew.Cells[rCnt, "O"].Value2;
                    string NDays_OT = xlWorkSheetNew.Cells[rCnt, "AA"].Value2;

                    //BaseTime = (new TimeSpan(Convert.ToDateTime(OffDuty).Hour, Convert.ToDateTime(OffDuty).Minute, 0)) - (new TimeSpan(Convert.ToDateTime(OnDuty).Hour, Convert.ToDateTime(OnDuty).Minute, 0));

                    if (Late_Time != "")
                    {
                        Late = new TimeSpan(Convert.ToDateTime(Late_Time).Hour, Convert.ToDateTime(Late_Time).Minute, 0);
                    }
                    else
                    {
                        Late = TimeSpan.Zero;
                    }

                    if (Early_Time != "")
                    {
                        Early = new TimeSpan(Convert.ToDateTime(Early_Time).Hour, Convert.ToDateTime(Early_Time).Minute, 0);
                    }
                    else
                    {
                        Early = TimeSpan.Zero;
                    }


                    if (FlagNormal == "1")
                    {
                        if (NDays_OT != "")
                        {
                            RegularOT = Convert.ToDouble(NDays_OT) - Late.TotalMinutes;
                            if (RegularOT < 0) RegularOT = 0;
                            xlWorkSheetNew.Cells[rCnt, "AD"].Value2 = RegularOT;
                        }
                        else
                        {
                            RegularOT = 0;
                        }
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = RegularOT / 60.0;
                        //if (OT_Time != "")
                        //{
                        //    //WorkingHours = new TimeSpan(Convert.ToDateTime(WorkTime).Hour, Convert.ToDateTime(WorkTime).Minute, 0);
                        //    OT = new TimeSpan(Convert.ToDateTime(OT_Time).Hour, Convert.ToDateTime(OT_Time).Minute, 0);
                        //    if (OT - Late <= TimeSpan.Zero)
                        //    {
                        //        CorrectOT = TimeSpan.Zero;
                        //    }
                        //    else
                        //    {
                        //        CorrectOT = OT - Late;
                        //    }
                        //    xlWorkSheetNew.Cells[rCnt, "AD"].Value2 = GetTotalAsString(CorrectOT);
                        //}
                        //else
                        //{
                        //    CorrectOT = TimeSpan.Zero;
                        //}
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = (double) (CorrectOT.TotalMinutes / 60.0);
                    }
                    else if (FlagWeekend == "1")
                    {
                        WeekEnd_OT = Convert.ToDouble(Weekend);
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = WeekEnd_OT / 60.0;
                        Holiday_OT = 0;
                        //Weekend_OT = (int)new TimeSpan(Convert.ToDateTime(Att_Time).Hour, Convert.ToDateTime(Att_Time).Minute, 0).TotalMinutes;
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = (double)(Weekend_OT / 60.0);
                    }
                    else if (FlagHoliday == "1")
                    {
                        Holiday_OT = Convert.ToDouble(Holiday);
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = Holiday_OT / 60.0;
                        WeekEnd_OT = 0;
                        //Holiday_OT = WorkingHours + OT;
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = (double)(Holiday_OT.TotalMinutes / 60.0);
                    }
                    else
                    {
                        //xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = 0;
                        RegularOT = 0;
                        Holiday_OT = 0;
                        WeekEnd_OT = 0;
                    }

                    if (Ac_No == xlWorkSheetNew.Cells[rCnt, "B"].Value2)
                    {
                        //Same Emp
                        Normal_OT_Total += RegularOT;
                        Weekend_OT_Total += WeekEnd_OT;
                        Holiday_OT_Total += Holiday_OT;
                        TotalLate += Late;
                        TotalEarly += Early;
                        cCnt++;
                        xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = (RegularOT + WeekEnd_OT + Holiday_OT) / 60.0;

                        RegularOT = 0;
                        WeekEnd_OT = 0;
                        Holiday_OT = 0;
                        
                        

                        if (rCnt == range.Rows.Count)
                        {
                            xlWorkSheetSummaryOvertime.Cells[index, "C"].value2 = Normal_OT_Total / 60.0;
                            xlWorkSheetDailyOvertime.Cells[index, "AH"].value2 = Normal_OT_Total / 60.0;
                            xlWorkSheetSummaryOvertime.Cells[index, "D"].value2 = Weekend_OT_Total / 60.0;
                            xlWorkSheetDailyOvertime.Cells[index, "AI"].value2 = Weekend_OT_Total / 60.0;
                            xlWorkSheetSummaryOvertime.Cells[index, "E"].value2 = Holiday_OT_Total / 60.0;
                            xlWorkSheetDailyOvertime.Cells[index, "AJ"].value2 = Holiday_OT_Total / 60.0;
                            xlWorkSheetSummaryOvertime.Cells[index, "F"].value2 = ((Weekend_OT_Total + Holiday_OT_Total) / 60.0);//NDays_OT + Weekend_OT_Total + Holiday_OT_Total;
                            xlWorkSheetDailyOvertime.Cells[index, "AK"].value2 = ((Weekend_OT_Total + Holiday_OT_Total) / 60.0);//NDays_OT + Weekend_OT_Total + Holiday_OT_Total;
                            xlWorkSheetNew.Cells[rCnt, "AE"].Value2 = GetTotalAsString(TotalEarly);
                            xlWorkSheetNew.Cells[rCnt, "AF"].Value2 = GetTotalAsString(TotalLate);
                            
                        }
                    }
                    else
                    {
                        // Another emp
                        xlWorkSheetSummaryOvertime.Cells[index, "C"].value2 = Normal_OT_Total / 60.0;
                        xlWorkSheetDailyOvertime.Cells[index, "AH"].value2 = Normal_OT_Total / 60.0;
                        xlWorkSheetSummaryOvertime.Cells[index, "D"].value2 = Weekend_OT_Total / 60.0;
                        xlWorkSheetDailyOvertime.Cells[index, "AI"].value2 = Weekend_OT_Total / 60.0;
                        xlWorkSheetSummaryOvertime.Cells[index, "E"].value2 = Holiday_OT_Total / 60.0;
                        xlWorkSheetDailyOvertime.Cells[index, "AJ"].value2 = Holiday_OT_Total / 60.0;
                        xlWorkSheetSummaryOvertime.Cells[index, "F"].value2 = ((Weekend_OT_Total + Holiday_OT_Total) / 60.0);//NDays_OT + Weekend_OT_Total + Holiday_OT_Total;
                        xlWorkSheetDailyOvertime.Cells[index, "AK"].value2 = ((Weekend_OT_Total + Holiday_OT_Total) / 60.0);//NDays_OT + Weekend_OT_Total + Holiday_OT_Total;
                        xlWorkSheetNew.Cells[rCnt - 1, "AE"].Value2 = GetTotalAsString(TotalEarly);
                        xlWorkSheetNew.Cells[rCnt - 1, "AF"].Value2 = GetTotalAsString(TotalLate);
                        index++;
                        cCnt = 3;
                        xlWorkSheetDailyOvertime.Cells[index, cCnt].value2 = (RegularOT + WeekEnd_OT + Holiday_OT) / 60.0;



                        Normal_OT_Total = 0;
                        Weekend_OT_Total = 0;
                        Holiday_OT_Total = 0;
                        TotalEarly = TimeSpan.Zero;
                        TotalLate = TimeSpan.Zero;


                        Normal_OT_Total += RegularOT;
                        Weekend_OT_Total += WeekEnd_OT;
                        Holiday_OT_Total += Holiday_OT;
                        TotalLate += Late;
                        TotalEarly += Early;

                        RegularOT = 0;
                        WeekEnd_OT = 0;
                        Holiday_OT = 0;

                        Ac_No = xlWorkSheetNew.Cells[rCnt, "B"].Value2;
                        Name = xlWorkSheetNew.Cells[rCnt, "D"].Value2;

                        xlWorkSheetSummaryOvertime.Cells[index, "A"].value2 = Ac_No;
                        xlWorkSheetSummaryOvertime.Cells[index, "B"].value2 = Name;
                        xlWorkSheetDailyOvertime.Cells[index, "A"].value2 = Ac_No;
                        xlWorkSheetDailyOvertime.Cells[index, "B"].value2 = Name;
                    }
                    progressBar1.Increment(1);             
                }

                Excel.Range myRange;
                myRange = xlWorkSheetDailyOvertime.UsedRange;
                myRange.EntireColumn.AutoFit();
                myRange.HorizontalAlignment = HorizontalAlignment.Center;
                myRange = xlWorkSheetDailyOvertime.get_Range("b2", "AK2000");
                myRange.NumberFormat = "0.00";

                Excel.Range myRange2;
                myRange2 = xlWorkSheetSummaryOvertime.UsedRange;
                myRange2.EntireColumn.AutoFit();
                myRange2.HorizontalAlignment = HorizontalAlignment.Center;
                myRange2 = xlWorkSheetSummaryOvertime.get_Range("c2", "f2000");
                myRange2.NumberFormat = "0.00";

                xlWorkBookCurrent.Close(false, null, null);
                saveFileDialog1.ShowDialog();
                xlWorkBookNew.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookNew.Close(true, null, null);

                saveFileDialog1.ShowDialog();
                xlWorkBookSummaryOvertime.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookSummaryOvertime.Close(true, null, null);

                saveFileDialog1.ShowDialog();
                xlWorkBookDailyOvertime.SaveAs(saveFileDialog1.FileName, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookDailyOvertime.Close(true, null, null);

                xlApp.Quit();

                releaseObject(xlWorkSheetCurrent);
                releaseObject(xlWorkBookCurrent);
                releaseObject(xlWorkSheetNew);
                releaseObject(xlWorkBookNew);
                releaseObject(xlWorkSheetSummaryOvertime);
                releaseObject(xlWorkBookSummaryOvertime);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private string GetTotalAsString(TimeSpan ts)
        {
            DateTime dt = new DateTime(ts.Ticks);
            string s = dt.ToString("HH:mm");
            return  s.Replace(":", ".");
        }
    }
}
