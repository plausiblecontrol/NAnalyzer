using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace sherpanalyzer {
  class Program {
    static void Main(string[] args) {
      List<string> asyncT = new List<string>();
      List<string> errors = new List<string>();  
      Console.WriteLine(System.Environment.NewLine);
      Console.WriteLine("Running automatic full reports on all nmon-based files and blg files in this recursive directory...");
      Console.WriteLine("Review .\\Output.txt when completed for a detailed report.");
      string[] zips = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.zip", SearchOption.AllDirectories).ToArray();
      if (zips.Count() > 0) {
        try {//can you .NET4.5?
        Parallel.ForEach(zips, z => {
          using (ZipArchive archive = ZipFile.OpenRead(z)) {
              foreach (ZipArchiveEntry entry in archive.Entries) {
                string exName = "";
                if (entry.FullName.Contains("nmon")) {
                  try {//can you unzip that file?
                      entry.ExtractToFile(Path.Combine(z.Substring(0, z.LastIndexOf('\\')), entry.Name), true);
                    Console.WriteLine("Unzipped " + entry.Name);
                  } catch {
                    errors.Add("Had trouble unzipping this file: " + z);
                  }
                } else if (entry.Name.Contains(".blg")) {
                  try {//can you unzip that file?
                    //OVERWRITE SAME FOLDER MUHGAWD pls fix
                    string fzName = entry.FullName;
                    int ixf = fzName.IndexOf("/");
                    exName = fzName.Substring(0, ixf);
                    string nfhx = Path.Combine(z.Substring(0, z.LastIndexOf('\\')), exName);
                    //int inc = 1;
                    while (Directory.Exists(nfhx)) {
                      nfhx += "x";
                      }
                    //entry.ExtractToDirectory(nfhx);
                    archive.ExtractToDirectory(nfhx);
                    Console.WriteLine("Unzipped " + exName);
                    break;
                  } catch {
                    errors.Add("Had trouble unzipping " + exName +" please check for completion.");
                  }
                }
              }
            }
          //using (ZipArchive archive = ZipFile.Open(z, ZipArchiveMode.Read)) {
          //  foreach (ZipArchiveEntry entry in archive.Entries) {
          //    if (entry.FullName.Contains("blg")) {
          //      try {//can you unzip that file?
          //        string fzName = entry.FullName;
          //        int ixf = fzName.IndexOf("/");
          //        string exName = fzName.Substring(0, ixf);
          //        string nfhx = Path.Combine(z.Substring(0, z.LastIndexOf('\\')), exName);
          //        //entry.ExtractToDirectory(nfhx);
          //        archive.ExtractToDirectory(nfhx);
          //        Console.WriteLine("Unzipped " + exName);
          //      } catch {
          //        errors.Add("Had trouble unzipping this file: " + z);
          //      }
          //    }
          //  }
          //}
          
        });
        } catch {
          Console.WriteLine("Your System is not updated to support .NET Framework 4.5, please update!!!");
          errors.Add("Your System is not updated to support .NET Framework 4.5, please update!!!");
          errors.Add("Could not unzip and process, exiting with errors.");
        }
      }
      string here = Directory.GetCurrentDirectory();
      string[] nmonfiles = Directory.GetFiles(here, "*.*", SearchOption.AllDirectories).Where(name => name.Substring(name.LastIndexOf('\\'), name.Length - name.LastIndexOf('\\')).Contains("nmon")).Where(name => !name.Contains(".zip")).ToArray();
      string[] blgfiles = Directory.GetFiles(here, "*.blg", SearchOption.AllDirectories);

      #region perfmons
      List<string> csvfiles = new List<string>();
      if (blgfiles.Count() > 0) {
        string filterText = @here + @"\CB745.txt";
        using (StreamWriter sw = File.CreateText(filterText)) {
          sw.WriteLine(@"\PhysicalDisk(_Total)\% Disk Time");
          sw.WriteLine(@"\Memory\Available MBytes");
          sw.WriteLine(@"\Processor(_Total)\% Processor Time ");
          sw.WriteLine(@"\Network Interface(*)\Bytes Total/sec ");
        }
        Console.WriteLine("Relogging Windows PerfMons.");
        System.Diagnostics.Process process = new System.Diagnostics.Process();
        System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo();
        startInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
        startInfo.FileName = "CMD.exe";
        //Parallel.ForEach(blgfiles, blgfile => {
        foreach (string blgfile in blgfiles) { 
          File.Delete(blgfile.Substring(0, blgfile.Length - 4) + ".csv");
          startInfo.Arguments = "/C relog " + blgfile + " -cf " + filterText + " -f csv -o " + blgfile.Substring(0, blgfile.Length - 4) + ".csv";
          csvfiles.Add(blgfile.Substring(0, blgfile.Length - 4) + ".csv");
          process.StartInfo = startInfo;
          process.Start();
          process.WaitForExit();
        //});
        }
        File.Delete(filterText);
        Console.WriteLine("Windows PerfMons can take a few minutes.");
        //string temptest = @"T:\Progs\nmon\sherpanalyzer\sherpanalyzer\bin\x64\Debug\JCRE2_DAC1\000001\PerfMonitor.csv";
        List<csvfile> AllCSVs = new List<csvfile>();
        foreach (string csvf in csvfiles) {
          csvfile csvTemp = new csvfile(csvf);
          AllCSVs.Add(csvTemp);
        }
        string[] TotalHosts = (from x in AllCSVs select x.hostname).Distinct().ToArray();
        var sortedHosts = from x in AllCSVs
                          orderby x.date
                          group x by x.hostname into newgroup
                          orderby newgroup.Key
                          select newgroup;

        List<string> MergedCSVs = new List<string>();
        //List<string> completedFiles = new List<string>();
        foreach (var shost in sortedHosts) {
          List<string> hoststoCombine = new List<string>();
          Console.WriteLine("Sherpalyzing- "+shost.Key);
          foreach (var dda in shost) {
            hoststoCombine.Add(dda.filelocation);
          }
          MergedCSVs.Add(combineCSVs(hoststoCombine, here, shost.Key));
        }
        Parallel.ForEach(MergedCSVs, mcsv=>{
          OutputUnits resultBLG = makePerfMonGraphs(mcsv);
           asyncT.Add(resultBLG.display());
           if (resultBLG.isLeaking) {
             Console.WriteLine("Possible Memory leak found in "+Path.GetFileNameWithoutExtension(mcsv));
           }
           if (resultBLG.highCPU) {
             Console.WriteLine("Possible high CPU usage found in " + Path.GetFileNameWithoutExtension(mcsv));
           }
        });
        foreach (csvfile mcsv in AllCSVs) {
          File.Delete(mcsv.filelocation);
        }
        foreach (string mcsv in MergedCSVs) {
          try { File.Delete(mcsv); } catch { Console.WriteLine("FILE IN USE! Excel frozen?"); }
        }

        //asyncT.Add("Completed Windows PerfMon analysis. "+System.Environment.NewLine);
        
      } else {
        asyncT.Add("No blg (windows) files were found in this directory.");
      }
      #endregion
      #region nmon
      if (nmonfiles.Count() < 1) {
        asyncT.Add("No nmon (redhat) files were found in this directory.");
      } else {
        //asyncT.Add("Starting nmon analysis");
        Parallel.ForEach(nmonfiles, nmons => {
          try {//only try; do not. unhandled crashes; doing will cause.
            asyncT.Add(readSherpa(nmons, true));
          } catch {//when someone makes readSherpa crash, tell them there was a bad file in there
            errors.Add("Had trouble sherpalyzing this file: " + nmons);
          }
        });



        if (errors.Count == 0) {
          asyncT.Add("No errors sherpalyzing!");
        } else {
          foreach (string e in errors) {
            asyncT.Add(e);
          }
        }
//        outputFile(asyncT);
        Console.WriteLine("Finished");
        String graphs = "";
        if (args.Length != 0) {
          Console.WriteLine("Create graphs? (y/n)");
          graphs = Console.ReadLine();
        } else {
          Console.WriteLine("Starting automated graph creation.");
          graphs = "y";
        }

        if (graphs.Length > 0) {
          if (graphs.Substring(0, 1) == "y" || graphs.Substring(0, 1) == "Y") {
            Console.WriteLine("Working with Excel in the background.");
            Console.WriteLine("This may take several minutes...");
            string[] CSVs = Directory.GetFiles(Directory.GetCurrentDirectory(), "*.csv", SearchOption.TopDirectoryOnly).ToArray();
            int cL = CSVs.Length;
            int i = 1;
            Parallel.ForEach(CSVs, csvF => {
              makeGraphs(csvF);
              File.Delete(csvF);
              Console.WriteLine("Completed " + i + " of " + cL);
              i++;
            });
          }
        }
      }
      outputFile(asyncT);
      #endregion
    }

    static double[] ColToDouble(Array items) {
      double[] result = new double[items.Length];
      for (int i = 1; i <= items.Length; i++) {
        result[i - 1] = Convert.ToDouble(items.GetValue(i, 1));
      }
      return result;
    }

    static double Variance(List<double> list) {
      double mean = list.Average();
      double result = list.Sum(number => Math.Pow(number - mean, 2.0));
      return result / list.Count();
    }
    static double StandardDev2(List<double> list) {
      return Math.Sqrt(Variance(list));
    }
    static double StandardDev(List<double> valueList) {
      double M = 0.0;
      double S = 0.0;
      int k = 1;
      foreach (double value in valueList) {
        double tmpM = M;
        M += (value - tmpM) / k;
        S += (value - tmpM) * (value - M);
        k++;
      } 
      return Math.Sqrt(S / (k - 1));
    }

    static double[] LeastSquares(List<double> valueList) {
      int numPoints = 0;
      double sumX = 0;
      double sumY = 0;
      double sumXX = 0;
      double sumXY = 0;
      //foreach (var tuple in valueList) {
        for (int i = 0; i < valueList.Count(); i++) { 
          numPoints++;
          sumX += i;// tuple.Item1;
          sumY += valueList[i];// tuple.Item2;
        sumXX += i*i;//tuple.Item1 * tuple.Item1;
        sumXY += i*valueList[i];//tuple.Item1 * tuple.Item2;
      }

      double b = (-sumX * sumXY + sumXX * sumY) / (numPoints * sumXX - sumX * sumX);
      double m = (-sumX * sumY + numPoints * sumXY) / (numPoints * sumXX - sumX * sumX);

      return new double[] { m, b };
    }

    static double[] SumOfSq(List<double> valueList) {
      double SumXY = 0;
        double SumXX = 0;
          double SumY = 0;
            double SumX = 0;
              //lolTabWat
      double mSlope = 0;
      double bIntercept = 0;
      for (int i = 1; i < valueList.Count; i++) {
        double a = valueList[i - 1];//Convert.ToDouble(availMem.GetValue(i, 1));
        SumXY += i * a;
        SumXX += i * i;
        SumY += a;
        SumX += i;
      }
      mSlope = ((SumXY - SumX * SumY / valueList.Count) / (SumXX - SumX * SumX / valueList.Count));
      bIntercept = (SumXY * SumX - SumY * SumXX) / (SumX * SumX - valueList.Count * SumXX);
      double[] xMX = new double[] {mSlope,bIntercept};
      return xMX;
    }

    static List<double> sigma(List<double> valueList, double StdDev, double[] equation) {
      List<double> memList = new List<double>();      
      for (int i = 1; i <= valueList.Count; i++) {
        double kI = equation[0] * i + equation[1];
        if (valueList[i - 1] <= (kI + StdDev) && valueList[i - 1] >= (kI - StdDev)) {
          memList.Add(valueList[i - 1]);
        } else if (valueList[i - 1] > (kI + StdDev)) {
          memList.Add(kI + StdDev);
        } else {
          memList.Add(kI - StdDev);
        }
      }
      return memList;
    }
    static OutputUnits makePerfMonGraphs(string file) { 
      Excel.Application excelApp = null;
      Excel.Workbook workbook = null;
      Excel.Sheets sheets = null;
      Excel.Worksheet dataSheet = null;
      Excel.Worksheet newSheet = null;
      Excel.ChartObjects xlChart = null;
      Excel.Range dataY = null;
      Excel.Chart memChart = null;
      Excel.Chart diskChart = null;
      Excel.Chart cpuChart = null;
      Excel.Chart netChart = null;
      Excel.Axis xAxis = null;
      OutputUnits csData = null;
      bool leaking = false;
      bool highcpu = false;
      string exitFile = "";
      try {
        excelApp = new Excel.Application();
        string dir = file.Substring(0, file.LastIndexOf("\\") + 1);
        string fm = file.Substring(0, file.Length - 4).Substring(file.LastIndexOf("\\") + 1);
        workbook = excelApp.Workbooks.Open(file, 0, false, 6, Type.Missing, Type.Missing, Type.Missing, XlPlatform.xlWindows, ",",
          true, false, 0, false, false, false);

        sheets = workbook.Sheets;
        dataSheet = sheets[1];
        dataSheet.Name = "data";
        dataSheet.get_Range("A2:A2", Type.Missing).EntireRow.Delete(XlDeleteShiftDirection.xlShiftUp);//garbage row
        newSheet = (Worksheet)sheets.Add(Type.Missing, dataSheet, Type.Missing, Type.Missing);
        newSheet.Name = "results";
        xlChart = (Excel.ChartObjects)newSheet.ChartObjects(Type.Missing);
        
        memChart = xlChart.Add(20, 100, 450, 175).Chart;
        diskChart = xlChart.Add(20, 280, 450, 175).Chart;
        cpuChart = xlChart.Add(500, 100, 450, 175).Chart;
        netChart = xlChart.Add(500, 280, 450, 175).Chart;
        int rowTotal = dataSheet.UsedRange.Rows.Count;
        int colTotal = dataSheet.UsedRange.Columns.Count;
        dataSheet.get_Range("A2", "A" + rowTotal).NumberFormat ="m/d/yyyy h:mm";
        string ttime = dataSheet.Cells[2, 1].Value.ToString();
        Array availMem = (System.Array)dataSheet.get_Range("C2", "C" + rowTotal).Value;
        Array cpuTotal = (System.Array)dataSheet.get_Range("D2", "D" + rowTotal).Value;
        Array diskTotal = (System.Array)dataSheet.get_Range("B2", "B" + rowTotal).Value;

        dataSheet.Cells[1, colTotal + 1] = "Total LAN (Bytes Total/Sec)";
        double[] netties = new double[rowTotal-1];
        for (int i = 2; i <= rowTotal; i++) {
          if (colTotal > 5) { 
            Array netLine = (System.Array)dataSheet.get_Range(xlStr(5)+ i, xlStr(colTotal)+ i).Value;
            double netLineTotal = 0;
            for(int j=1;j<=netLine.Length;j++){
              netLineTotal += Convert.ToDouble(netLine.GetValue(1,j));
            }
            netties[i - 2] = netLineTotal;
            dataSheet.Cells[i, colTotal + 1] = netLineTotal;
          }else{
            dataSheet.Cells[i, colTotal + 1] = "0";
          }
        }

        #region BuildCounters
        double[] mems = ColToDouble(availMem);
        double[] cpus = ColToDouble(cpuTotal);
        double[] disks = ColToDouble(diskTotal);
        //netties[]
        double avgCPUs = cpus.Average();
        PCounter CPU = new PCounter(avgCPUs, cpus.Max(), cpus.Min());
        PCounter MEM = new PCounter(mems.Average(), mems.Max(), mems.Min());
        PCounter DISK = new PCounter(disks.Average(), disks.Max(), disks.Min());
        PCounter NETS = new PCounter(netties.Average(), netties.Max(), netties.Min());
        if (avgCPUs > 40) {
          highcpu = true;
        }
        #endregion

        #region leakCheck
        double[] eqMB = new double[2];
        List<double> memList = new List<double>();
        int cX = availMem.Length;
        for (int i = 1; i < rowTotal - 1; i++) {
          memList.Add(Convert.ToDouble(availMem.GetValue(i, 1)));
        }
        eqMB = LeastSquares(memList);
        double stdD1 = StandardDev(memList);
        List<double> memList2 = sigma(memList, stdD1, eqMB);
        cX = memList2.Count();
        eqMB = LeastSquares(memList2);
        double stdD2 = StandardDev(memList2)*1.2;
        List<double> memList3 = sigma(memList2, stdD2, eqMB);
        eqMB = LeastSquares(memList3);

        if (eqMB[0] < 0) {
          leaking = true;
          newSheet.get_Range("E4", Type.Missing).Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Tomato);
        }
        #endregion

        #region formatting
        string lan = xlStr(colTotal + 1);
        newSheet.get_Range("A1", Type.Missing).EntireColumn.ColumnWidth = 12;
        newSheet.get_Range("A1", Type.Missing).EntireColumn.HorizontalAlignment = XlHAlign.xlHAlignRight;
        newSheet.get_Range("A2", Type.Missing).EntireRow.HorizontalAlignment = XlHAlign.xlHAlignCenter;
        newSheet.Cells[4, 5] = eqMB[0];
        newSheet.Cells[2, 2] = "Avg";
        newSheet.Cells[2, 3] = "Min";
        newSheet.Cells[2, 4] = "Max";
        newSheet.Cells[2, 5] = "Slope(3Sigma)";
        newSheet.Cells[3, 1] = "CPU";
        newSheet.Cells[3, 2] = "=AVERAGE(data!D:D)";
        newSheet.Cells[3, 3] = "=MIN(data!D:D)";
        newSheet.Cells[3, 4] = "=MAX(data!D:D)";
        newSheet.Cells[4, 1] = "Avail.RAM";
        newSheet.Cells[4, 2] = "=AVERAGE(data!C:C)";
        newSheet.Cells[4, 3] = "=MIN(data!C:C)";
        newSheet.Cells[4, 4] = "=MAX(data!C:C)";
        newSheet.Cells[5, 1] = "LAN Usage";
        newSheet.Cells[5, 2] = "=AVERAGE(data!"+lan+":"+lan+")";
        newSheet.Cells[5, 3] = "=MIN(data!"+lan+":"+lan+")";
        newSheet.Cells[5, 4] = "=MAX(data!" + lan + ":" + lan + ")";
        newSheet.Cells[6, 1] = "Disk Usage";
        newSheet.Cells[6, 2] = "=AVERAGE(data!B:B)";
        newSheet.Cells[6, 3] = "=MIN(data!B:B)";
        newSheet.Cells[6, 4] = "=MAX(data!B:B)";

        #endregion

        #region memChart
        dataY = dataSheet.Range["C1", "C" + rowTotal];
        memChart.SetSourceData(dataY, Type.Missing);
        memChart.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
        memChart.HasLegend = false;
        xAxis = (Axis)memChart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.MaximumScaleIsAuto = false;
        xAxis.MaximumScale = rowTotal + 1;
        xAxis.MinimumScaleIsAuto = false;
        xAxis.MinimumScale = 0;
        #endregion

        #region diskChart
        dataY = dataSheet.Range["B1", "B" + rowTotal];
        diskChart.SetSourceData(dataY, Type.Missing);
        diskChart.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
        diskChart.HasLegend = false;
        xAxis = (Axis)diskChart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.MaximumScaleIsAuto = false;
        xAxis.MaximumScale = rowTotal + 1;
        xAxis.MinimumScaleIsAuto = false;
        xAxis.MinimumScale = 0;
        #endregion

        #region cpuChart
        dataY = dataSheet.Range["D1", "D" + rowTotal];
        cpuChart.SetSourceData(dataY, Type.Missing);
        cpuChart.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
        cpuChart.HasLegend = false;
        xAxis = (Axis)cpuChart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.MaximumScaleIsAuto = false;
        xAxis.MaximumScale = rowTotal + 1;
        xAxis.MinimumScaleIsAuto = false;
        xAxis.MinimumScale = 0;
        #endregion

        #region netChart
        dataY = dataSheet.Range[xlStr(colTotal + 1)+"1", xlStr(colTotal + 1) + rowTotal];
        netChart.SetSourceData(dataY, Type.Missing);
        netChart.ChartType = XlChartType.xlXYScatterLinesNoMarkers;
        netChart.HasLegend = false;
        xAxis = (Axis)netChart.Axes(XlAxisType.xlCategory, XlAxisGroup.xlPrimary);
        xAxis.MaximumScaleIsAuto = false;
        xAxis.MaximumScale = rowTotal + 1;
        xAxis.MinimumScaleIsAuto = false;
        xAxis.MinimumScale = 0;
        #endregion

        string host = Path.GetFileNameWithoutExtension(dir + fm);
        csData = new OutputUnits(host, ttime+" time chunks: "+(rowTotal-1), CPU, MEM, NETS, DISK, leaking, highcpu);
        exitFile = dir + fm;
        excelApp.DisplayAlerts = false;
        workbook.SaveAs(@exitFile, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        workbook.Close(true, Type.Missing, Type.Missing);
        excelApp.Quit();
        
        //releaseObject(sC);
        //releaseObject(myChart);
      } catch {
        Console.WriteLine("Had issues interacting with your Excel installation...maybe try a restart?");
        //using (StreamWriter outfile = File.AppendText("output.txt")) {
        //  outfile.WriteLine("Did have issues interacting with Excel on " + file);
        //}
      } finally {
        releaseObject(xAxis);
        releaseObject(dataY);
        releaseObject(diskChart);
        releaseObject(memChart);
        releaseObject(cpuChart);
        releaseObject(netChart);
        releaseObject(xlChart);
        releaseObject(newSheet);
        releaseObject(dataSheet);
        releaseObject(sheets);
        releaseObject(workbook);
        releaseObject(excelApp);
      }
      return csData;
    }
    static string xlStr(int column) {
      string[] alphabet = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"};
      string a = alphabet[column-1];
      return a;
    }
    static string combineCSVs(List<string> OrderedFileLocations, string path, string host) {
      string fn = host + ".csv";
      string combinedFile = @path + @"\"+fn;
      File.Delete(combinedFile);
      File.Copy(OrderedFileLocations[0], combinedFile);
      if (OrderedFileLocations.Count > 1) {
        for(int i =1;i<OrderedFileLocations.Count();i++){
          List<string> lines = new List<string>();
          using(StreamReader sr = new StreamReader(OrderedFileLocations[i])){
            sr.ReadLine();//strip header
            sr.ReadLine();//strip garbage line
            string line;
            while ((line = sr.ReadLine()) != null) {
              lines.Add(line);
            }
          }
          File.AppendAllLines(combinedFile, lines);
        }

      }
      //using (StreamWriter sw = File.CreateText(combinedFile)) {
      //  sw.WriteLine(@"\PhysicalDisk(_Total)\% Disk Time");
      //  sw.WriteLine(@"\Memory\Available MBytes");
      //  sw.WriteLine(@"\Processor(_Total)\% Processor Time ");
      //  sw.WriteLine(@"\Network Interface(*)\Bytes Total/sec ");
      //}
      return combinedFile;
    }

    static void makeGraphs(string file) {
      Excel.Application excelApp = null;
      Excel.Workbook workbook = null;
      Excel.Sheets sheets = null;
      Excel.Worksheet dataSheet = null;
      Excel.Worksheet newSheet = null;
      Excel.Worksheet chartSheet = null;
      Excel.Range range = null;
      Excel.Range dataR = null;
      int rowC = 0;
      try {
        excelApp = new Excel.Application();
        string dir = file.Substring(0, file.LastIndexOf("\\") + 1);
        string fm = file.Substring(0, file.Length - 4).Substring(file.LastIndexOf("\\") + 1);
        workbook = excelApp.Workbooks.Open(file, 0, false, 6, Type.Missing, Type.Missing, Type.Missing, XlPlatform.xlWindows, ",",
          true, false, 0, false, false, false);

        sheets = workbook.Sheets;
        dataSheet = sheets[1];
        dataSheet.Name = "data";
        newSheet = (Worksheet)sheets.Add(Type.Missing, dataSheet, Type.Missing, Type.Missing);
        newSheet.Name = "table";
        chartSheet = (Worksheet)sheets.Add(Type.Missing, dataSheet, Type.Missing, Type.Missing);
        chartSheet.Name = "graph";
        Excel.ChartObjects xlChart = (Excel.ChartObjects)chartSheet.ChartObjects(Type.Missing);
        dataR = dataSheet.UsedRange;
        rowC = dataR.Rows.Count;

        range = newSheet.get_Range("A1");
        PivotCaches pCs = workbook.PivotCaches();
        PivotCache pC = pCs.Create(XlPivotTableSourceType.xlDatabase, dataR, Type.Missing);
        PivotTable pT = pC.CreatePivotTable(TableDestination: range, TableName: "PivotTable1");
        PivotField fA = pT.PivotFields("Time");
        PivotField fB = pT.PivotFields("Command");
        fA.Orientation = XlPivotFieldOrientation.xlRowField;
        fA.Position = 1;
        fB.Orientation = XlPivotFieldOrientation.xlColumnField;
        fB.Position = 1;
        pT.AddDataField(pT.PivotFields("%CPU"), "Sum of %CPU", XlConsolidationFunction.xlSum);

        ChartObject pChart = (Excel.ChartObject)xlChart.Add(0, 0, 650, 450);
        Chart chartP = pChart.Chart;
        chartP.SetSourceData(pT.TableRange1, Type.Missing);
        chartP.ChartType = XlChartType.xlLine;
        excelApp.DisplayAlerts = false;
        workbook.SaveAs(@dir + fm, XlFileFormat.xlOpenXMLWorkbook, Type.Missing, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing);
        workbook.Close(true, Type.Missing, Type.Missing);
        excelApp.Quit();


      } catch {
        Console.WriteLine("Had issues interacting with your Excel installation...maybe try a restart?");
        using (StreamWriter outfile = File.AppendText("output.txt")) {
          outfile.WriteLine("Did have issues interacting with Excel on " + file);
        }
      } finally {
        /*Excel.Application excelApp = null;
      Excel.Workbook workbook = null;
      Excel.Sheets sheets = null;
      Excel.Worksheet dataSheet = null;
      Excel.Worksheet newSheet = null;
      Excel.Worksheet chartSheet = null;
      Excel.Range range = null;
      Excel.Range dataR = null;*/
        releaseObject(dataR);
        releaseObject(range);
        releaseObject(chartSheet);
        releaseObject(newSheet);
        releaseObject(dataSheet);
        releaseObject(sheets);
        releaseObject(workbook);
        releaseObject(excelApp);
      }
    }

    static void releaseObject(object obj) {
      try {
        System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
        obj = null;
      } catch {
        obj = null;
        Console.WriteLine("Exception Occured while releasing Excel objects - Bad files, or corrupted Excel.");
        Console.WriteLine("Ensure Excel.exe gets terminated from Task Manager.");
        Console.ReadKey();// + ex.ToString());
      } finally {
        GC.Collect();
        GC.WaitForPendingFinalizers();
      }
    }

    static void outputFile(List<string> outputs) {
      using (StreamWriter file = new StreamWriter("output.txt")) {
        file.WriteLine("Sherpalyzer - A.Pedersen 2/1/15" + System.Environment.NewLine);
        if (outputs.Count() > 0) {
          foreach (string e in outputs) {
            file.WriteLine(e);
          }
        }       
      }
    }

    static string readSherpa(string filename, bool print) {
      string line;
      string host = "new";
      int cores = 0;
      List<string> summary = new List<string>();
      List<string> times = new List<string>();
      List<double> cpu_all_usr = new List<double>();
      List<double> memsum = new List<double>();
      List<string> disklabels = new List<string>();
      List<double> disksizes = new List<double>();
      List<string> disksizesb = new List<string>();
      List<string> netwuts = new List<string>();
      List<List<string>> topList = new List<List<string>>();
      List<double>[] diskbusy;
      List<double>[] netties;      
      string[] dix;

      string datime = "";
      string ddate = "";
      string ttime = "";
      string warnings = "";
     
      using (StreamReader reader = new StreamReader(filename)) {

        /*read in each line of text then do stuff with it*/
        //small while loop only does maybe 50lines before breaking
        while ((line = reader.ReadLine()) != null) {//this is the prelim loop to make the primary loop go quicker
          summary.Add(line);
          string[] values = line.Split(',');
          if (values[1] == "time") {
            if (values[0] == "AAA")
              ttime = values[2];
              datime = String.Join("",values[2].Split(new[] { ':', '.' }));
          }
          if (values[1] == "date") {
            if (values[0] == "AAA")
              ddate = values[2];
              datime = String.Join("", values[2].Split('-')) + "_"+datime;
          }
          if (values[1] == "host")
            host = values[2];
          if (values[1] == "cpus")
            cores = Convert.ToInt32(values[2]);
          if (values[0] == "NET") {//first line of NET data from the file
            foreach (string nets in values.Skip(2)) { //for all the nets presented on this line (skipping the first 2 garbage lines)
              if(nets != "") netwuts.Add(nets);//all the things, each iface, each bond, eths, los..  everything from the ifconfig
            }
          }
          if (values[0] == "DISKBUSY") {//first line of DISKBUSY holds disk names
            foreach (string diskN in values.Skip(2)) { //for all the disk labels presented on this line (skipping the first 2 garbage lines)
              if(diskN != "") disklabels.Add(diskN);//all sd and dm partitions, just keep it all in there
            }
          }
          if (values[0] == "BBBP"){
            if (values[2] == "/proc/partitions") {
              try {
                dix = values[3].Split(new[] { ' ', '\"' }, StringSplitOptions.RemoveEmptyEntries);
                if (dix[0] != "major") {
                  disksizes.Add(Convert.ToDouble(dix[2])/1000);
                  disksizesb.Add(dix[3]);
                }
              } catch { }
            } else if (values[2] == "/proc/1/stat")
              break;
          }
        }//some background info was gathered from AAA


        netties = new List<double>[netwuts.Count()];
        for (int i = 0; i < netties.Count(); i++) {
          netties[i] = new List<double>();//so many I dont even
        }//we now have netwuts.count netties[]s; each netties is a double list we can add each(EVERY SINGLE) line nmon records
        
        diskbusy = new List<double>[disklabels.Count()];
        for (int i = 0; i < disklabels.Count(); i++) {
          diskbusy[i] = new List<double>();//almost as many I dont even
        }//we now have disklabels.count diskbusy[]s; each diskbusy is a double list we can add each(EVERY SINGLE) line nmon records
        List<process> processes = new List<process>();
          while ((line = reader.ReadLine()) != null) { //Got all the prelim done, now do the rest of the file
            string[] values = line.Split(',');
            /*switch was faster than an if block*/
            try{
            switch (values[0]) {
              case "ZZZZ":
                  times.Add(values[2] + " " + values[3]);
                break;
              case "TOP":
                List<string> topstuff = new List<string>();
                //TOP,+PID,Time,%CPU,%Usr,%Sys,Size,ResSet,ResText,ResData,ShdLib,MajorFault,MinorFault,Command
                //TOP,0031885,T0050,92.9,89.2,3.7,1416768,1105388,144,0,143692,34,0,osii_dbms_adapt  
                topstuff.Add(values[2].Substring(1, values[2].Length - 1));//time in front of topstuff
                for (int i = 1; i < values.Count(); i++) {
                  if (i != 2) {//skip time
                    topstuff.Add(values[i]);//add each value starting from 1 (skipping 2)
                    //Time,+PID,%CPU,%Usr,%Sys,Size,ResSet,ResText,ResData,ShdLib,MajorFault,MinorFault,Command
                    //0050,0031885,92.9,89.2,3.7,1416768,1105388,144,0,143692,34,0,osii_dbms_adapt
                  }                  
                }
                topList.Add(topstuff);
                string TOPpid = topstuff[1];
                    int TPID = Convert.ToInt32(TOPpid);
                    string TOPres = topstuff[6];
                    string TOPcpu = topstuff[2];
                    bool found = false;
                     
                    foreach (process x in processes) {
                      if (x.pid == TPID) {
                        x.add(TOPres, TOPcpu);
                        found = true;
                        break;
                      }
                    }
                    if (!found) {
                      process t = new process(TOPpid, TOPres, TOPcpu,topstuff[12]);
                      processes.Add(t);
                    }
                  break;
              case "CPU_ALL":
                if (values[2] != "User%") {
                    cpu_all_usr.Add((Convert.ToDouble(values[2]) + Convert.ToDouble(values[3])));
                }
                break;
              case "MEM":
                if (values[2] != "memtotal") {
                  memsum.Add(100.0 * (1 - ((Convert.ToDouble(values[6]) + Convert.ToDouble(values[11]) + Convert.ToDouble(values[14])) / Convert.ToDouble(values[2]))));
                  
                }
                break;
              case "NET":
                Parallel.ForEach(values.Skip(2), (nets, y, i) => {
                  if (nets != "") netties[i].Add(Convert.ToDouble(nets));
                });
                break;
              case "DISKBUSY":
                Parallel.ForEach(values.Skip(2), (disk, y, i) => {
                  diskbusy[i].Add(Convert.ToDouble(disk));
                });
                break;
              //etc
              default: //poison buckets barf pile
                break;
            }//end switch
          }catch(Exception e){
            string m = e.Message;
          }
          }//end while

        //eqMB = LeastSquares(memList);
        //double stdD1 = StandardDev(memList);
        //List<double> memList2 = sigma(memList, stdD1, eqMB);
        //cX = memList2.Count();
        //eqMB = LeastSquares(memList2);
        //double stdD2 = StandardDev(memList2)*1.2;
        //List<double> memList3 = sigma(memList2, stdD2, eqMB);
        //eqMB = LeastSquares(memList3);

        //if (eqMB[0] < 0) {
        //  leaking = true;

          double leakthreshold = 0.1;
          List<string> procXS = new List<string>();
          Parallel.ForEach(processes, x => {
            double leaks = x.isLeaking();
            if (leaks > leakthreshold && x.total > (times.Count*0.25)) {
              procXS.Add(x.pid + " (" + x.command+") rate of: "+Convert.ToDouble(leaks));
            }
          });
          if (procXS.Count > 0) {
            warnings += host+" has potential memory leaks in:" + System.Environment.NewLine;
            Console.WriteLine("Recommended graph investigation - possible leaks!");
            foreach (string l in procXS) {
              warnings += l + System.Environment.NewLine;
            }
          }
          List<string> cpuXS = new List<string>();
          Parallel.ForEach(processes, x => {            
            double cpuU = x.isHigh();
            if (cpuU>40) {
              cpuXS.Add(x.pid + " (" + x.command + ") averaging %"+Convert.ToInt32(cpuU));
              
            }
          });
          if (cpuXS.Count > 0) {
            warnings += host+" has potential high CPU usage in:" + System.Environment.NewLine;
            Console.WriteLine("Recommended graph investigation - found high CPU!");
            foreach (string l in cpuXS) {
              warnings += l + System.Environment.NewLine;
            }
          }
        
      }//done file handling
      
	  //inframortions
       //feels like a bad way to do this, but worked well
      string dump = "";
      dump += host + " from "+ttime+" "+ddate+" time chunks: " + times.Count + System.Environment.NewLine;
      
	  //CPU
      dump += host + " CPU(%) average: " + cpu_all_usr.Average() + System.Environment.NewLine;
      dump += host + " CPU(%) max: " + cpu_all_usr.Max() + System.Environment.NewLine;
      
	  //MEM
      dump += host+ " MEM(%) average: " + memsum.Average()+System.Environment.NewLine;
      dump += host + " MEM(%) max: " + memsum.Max() + System.Environment.NewLine;
      
	  //DISKBUSY
      for(int i=0;i<disklabels.Count;i++){
        if(disklabels[i].Substring(0,1)!="d"){
          dump += host + " DISKBUSY(%) avg for " + disklabels[i] + ": " + diskbusy[i].Average() + System.Environment.NewLine;
          dump += host + " DISKBUSY(%) max for " + disklabels[i] + ": " + diskbusy[i].Max() + System.Environment.NewLine;
        }
      }
      
	  //DISKBUSY weights
      double sdSum = 0.0;
      double diskweight = 0.0;
      double diskmaxes = 0.0;
      for (int i = 0; i < disksizesb.Count; i++) {
        if (disksizesb[i].Substring(0, 1) == "s") {
          sdSum += disksizes[i];
        }
      }
      for (int i = 0; i < disklabels.Count; i++) {  
        if (disklabels[i].Substring(0, 1) == "s") {
          for (int j = 0; j < disksizesb.Count; j++) {
            if (disksizesb[j] == disklabels[i]) {
              diskweight += (diskbusy[i].Average() * disksizes[j]);
              diskmaxes += (diskbusy[i].Max() * disksizes[j]);
            }
          }
        }
      }
      dump += host + " weighted DISKBUSY(%) avg: " + diskweight / sdSum + System.Environment.NewLine;
      dump += host + " weighted DISKBUSY(%) max: " + diskmaxes / sdSum + System.Environment.NewLine;
      
	  //NET
      for (int i = 0; i < netwuts.Count; i++) {
        if (netwuts[i].Substring(0, 2) != "lo") {//we dont need to see the loopback
          dump += host + " NET average for " + netwuts[i] + ": " + netties[i].Average() + System.Environment.NewLine;
        }
      }

      //warnings!!!
      if (warnings != "") {
        dump += warnings;
      }
      
      //CSV file stuff
      if (print) {
        string topTitle = "Time,PID,%CPU,%Usr,%Sys,Size,ResSet,ResText,ResData,ShdLib,MajorFault,MinorFault,Command";
        using (StreamWriter file = new StreamWriter(host +"_"+ datime+"_TOP.csv")) {
          file.WriteLine(topTitle);
          for (int i = 0; i < topList.Count; i++) {
            try {
              file.WriteLine(times[Convert.ToInt16(topList[i][0])-1] + "," + string.Join(",", topList[i].Skip(1)));//wat, dats right
            } catch {
              // *shrug do nothing
            }
          }
        }
      }
      Console.WriteLine("Finishing " + host +" from "+datime);
      return (dump);
    } 
  }
}
