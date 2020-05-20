using System;
using System.IO;
using System.Windows.Forms;
using Tools;

using Excel = Microsoft.Office.Interop.Excel;

namespace PlateData
{
    class Plate
    {
        // geometry data
        public double dL = 650.0;           // length
        public double dW = 390.0;           // width
        public double L1 = 50.0;            // Line 1
        public double L2 = 70.0;            // Line 2
        public double L3 = 200.0;           // Line 3
        public double L4 = 80.0;            // Line 4
        public double L5 = 100.0;           // Line 5
        public double L6 = 70.0;            // Line 5
        public double L7 = 120.0;           // Line 5
        public double L8 = 70.0;            // Line 5
        public double L9 = 90.0;            // Line 5
        public double R1 = 120.0;           // length
        public double R2 = 40.0;           // length
        public double dT = 10.0;            // thickness

        // drawing parameters
        public double dDimScale = 5.0;      // dimension scaling factor
        public double dDimLineSp = 50.0;    // dimension distance from object
        public double dViewSp = 200.0;      // view spacing

        public int[] nColor = new int[4];   // drawing colors 

        // Konstruktor
        public Plate()
        {
            nColor[0] = 1;                  // contour
            nColor[1] = 2;                  // dimensions
            nColor[2] = 3;                  // hidden lines
            nColor[3] = 4;                  // center lines
        }

        // write plate data into an EXCEL file
        public void Write(String fileName)
        {
            Log.Append(String.Format("> write '{0}'...", fileName));

            // excel object
            Excel.Application xlsApp;
            Excel.Workbook xlsWorkbook;
            Excel.Worksheet xlsSheet;

            // run excel
            xlsApp = new Excel.Application();
            xlsApp.Visible = true;          // -> show excel app
            xlsApp.DisplayAlerts = false;   // -> no popups

            // create workbook (xls file)
            xlsWorkbook = xlsApp.Workbooks.Add();

            // get xls sheet
            xlsSheet = xlsWorkbook.Worksheets[1];   // xls start from 1!!!

            // start writing
            // - title
            xlsSheet.Cells[1, 1] = "PlateApp Version 1.0";

            // geometric data
            int nRow = 3;
            xlsSheet.Cells[nRow, 1] = "Length";
            xlsSheet.Cells[nRow++, 2] = dL;
            xlsSheet.Cells[nRow, 1] = "Width";
            xlsSheet.Cells[nRow++, 2] = dW;
            xlsSheet.Cells[nRow, 1] = "Thickness";
            xlsSheet.Cells[nRow++, 2] = dT;

            // controling data
            nRow++;
            xlsSheet.Cells[nRow, 1] = "Scaling factor";
            xlsSheet.Cells[nRow++, 2] = dDimScale;
            xlsSheet.Cells[nRow, 1] = "Dimension spacing";
            xlsSheet.Cells[nRow++, 2] = dDimLineSp;
            xlsSheet.Cells[nRow, 1] = "View spacing";
            xlsSheet.Cells[nRow++, 2] = dViewSp;

            // color data
            nRow++;
            for (int i = 0; i < nColor.Length; i++)
            {
                xlsSheet.Cells[nRow, 1] = String.Format("Color {0}", i + 1);
                xlsSheet.Cells[nRow++, 2] = nColor[i];
            }

            // save file
            xlsWorkbook.SaveAs(fileName);
            xlsWorkbook.Close();

            // shutdown EXCEL
            xlsApp.Quit();

            // free com object
            ReleaseComObject(xlsApp);

        }

        // free com objects
        private void ReleaseComObject(Object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Log.Append("*** error: releasing com object!");
                Log.Append(ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        // load plate data from EXCEL file
        public Boolean Load()
        {
            String xlsFile;

            // get file name
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.InitialDirectory = Directory.GetCurrentDirectory();
            dlg.Filter = "New EXCEL|*.xlsx|Old EXCEL|*.xls|All Files|*.*";

            // start dialog
            if (dlg.ShowDialog() == DialogResult.OK)
            {
                xlsFile = dlg.FileName;
                Log.Append(String.Format("> load '{0}' ...", xlsFile));     
            }
            // cancel command
            else
            {
                Log.Append("*** Command canceled!");
                return false;
            }

            // read plate data from EXCEL file
            try
            {
                // excel object
                Excel.Application xlsApp;
                Excel.Workbook xlsWorkbook;
                Excel.Worksheet xlsSheet;

                // run excel
                xlsApp = new Excel.Application();
                xlsApp.Visible = true;          // -> show excel app
                xlsApp.DisplayAlerts = false;   // -> no popups

                // open excel file
                xlsWorkbook = xlsApp.Workbooks.Open(xlsFile);
                // get the sheet
                xlsSheet = xlsWorkbook.Worksheets[1];

                // print file version
                Log.Append(String.Format("File version: '{0}'",
                    xlsSheet.Cells[1, 1].value));

                // geometric data
                int nRow = 3;
                dL = xlsSheet.Cells[nRow++, 2].value;
                dW = xlsSheet.Cells[nRow++, 2].value;
                L1 = xlsSheet.Cells[nRow++, 2].value;
                L2 = xlsSheet.Cells[nRow++, 2].value;
                L3 = xlsSheet.Cells[nRow++, 2].value;
                L4 = xlsSheet.Cells[nRow++, 2].value;
                L5 = xlsSheet.Cells[nRow++, 2].value;
                L6 = xlsSheet.Cells[nRow++, 2].value;
                L7 = xlsSheet.Cells[nRow++, 2].value;
                L8 = xlsSheet.Cells[nRow++, 2].value;
                L9 = xlsSheet.Cells[nRow++, 2].value;
                R1 = xlsSheet.Cells[nRow++, 2].value;
                R2 = xlsSheet.Cells[nRow++, 2].value;
                dT = xlsSheet.Cells[nRow++, 2].value;

                // print file version
                Log.Append(String.Format("Length: '{0}'",
                    dL));

                // controling data
                nRow++;
                dDimScale  = xlsSheet.Cells[nRow++, 2].value;
                dDimLineSp = xlsSheet.Cells[nRow++, 2].value;
                dViewSp    = xlsSheet.Cells[nRow++, 2].value;

                // color data
                nRow++;
                for (int i = 0; i < nColor.Length; i++)
                {
                    nColor[i] = (int)xlsSheet.Cells[nRow++, 2].value;
                }

                // shutdown EXCEL
                xlsApp.Quit();

                // free com object
                ReleaseComObject(xlsApp);
            }

            // error handler
            catch (Exception ex)
            {
                Log.Append(String.Format("*** error: loading file '{0}'. Command canceled!",xlsFile));
                Log.Append(ex.ToString());
                return false;
            }

            // list plate data
            ListData();

            return true;
        }

        // check the plate data
        // -> geometric data in [mm]
        public int Check()
        {
            double dEps = 0.1;
            int nErr = 0;

            if (dL < dEps)
            {
                Log.Append(string.Format("*** error: invalid lenght: {0,10:0.00} ", dL));
                nErr++;
            }
            if (dW < dEps)
            {
                Log.Append(string.Format("*** error: invalid widht: {0,10:0.00} ", dW));
                nErr++;
            }
            if (dT < dEps)
            {
                Log.Append(string.Format("*** error: invalid thickness: {0,10:0.00} ", dT));
                nErr++;
            }
            if (R1 < dEps)
            {
                Log.Append(string.Format("*** error: invalid diameter: {0,10:0.00} ", R1));
                nErr++;
            }
            if (R2 < dEps)
            {
                Log.Append(string.Format("*** error: invalid diameter: {0,10:0.00} ", R2));
                nErr++;
            }
            // collision check
            if (R2 > dW)
            {
                Log.Append(string.Format("*** error: invalid diameter: {0,10:0.00} ", R2));
                nErr++;
            }
            if (R2 > dL)
            {
                Log.Append(string.Format("*** error: invalid diameter: {0,10:0.00} ", R2));
                nErr++;
            }
            if (dDimScale <= 0.0)
            {
                Log.Append(string.Format("*** error: invalid dimension scaling: {0,10:0.00} ",
                                         dDimScale));
                nErr++;
            }
            if (L1 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L1: {0,10:0.00} ", L1));
                nErr++;
            }
            if (L2 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L2: {0,10:0.00} ", L2));
                nErr++;
            }
            if (L3 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L3: {0,10:0.00} ", L3));
                nErr++;
            }
            if (L4 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L4: {0,10:0.00} ", L4));
                nErr++;
            }
            if (L5 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L5: {0,10:0.00} ", L5));
                nErr++;
            }
            if (L6 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L6: {0,10:0.00} ", L6));
                nErr++;
            }
            if (L7 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L7: {0,10:0.00} ", L7));
                nErr++;
            }
            if (L8 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L8: {0,10:0.00} ", L8));
                nErr++;
            }
            if (L9 < dEps)
            {
                Log.Append(string.Format("*** error: invalid L9: {0,10:0.00} ", L9));
                nErr++;
            }

            return nErr;
        }

        // List the plate's data
        public void ListData()
        {
            //                                             | Width
            //                                                | Format, 2 digits
            Log.Append(string.Format("Length        = {0,10:0.00}", dL));
            Log.Append(string.Format("Width         = {0,10:0.00}", dW));
            Log.Append(string.Format("Thickness     = {0,10:0.00}", dT));
            Log.Append(string.Format("L1            = {0,10:0.00}", L1));
            Log.Append(string.Format("L2            = {0,10:0.00}", L2));
            Log.Append(string.Format("L3            = {0,10:0.00}", L3));
            Log.Append(string.Format("L4            = {0,10:0.00}", L4));
            Log.Append(string.Format("L5            = {0,10:0.00}", L5));
            Log.Append(string.Format("L6            = {0,10:0.00}", L6));
            Log.Append(string.Format("L7            = {0,10:0.00}", L7));
            Log.Append(string.Format("L8            = {0,10:0.00}", L8));
            Log.Append(string.Format("L9            = {0,10:0.00}", L9));
            Log.Append(string.Format("R1            = {0,10:0.00}", R1));
            Log.Append(string.Format("R2            = {0,10:0.00}", R2));

            Log.Append(string.Format("Dim Scaling   = {0,10:0.00}", dDimScale));
            Log.Append(string.Format("Line Spacing  = {0,10:0.00}", dDimLineSp));
            Log.Append(string.Format("View Spacing  = {0,10:0.00}", dViewSp));
        }

    }
}
