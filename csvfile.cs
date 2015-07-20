using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace sherpanalyzer {
  public class csvfile {
    public string hostname;
    public DateTime date;
    public string filelocation;
    public csvfile(string filePath) {
      filelocation = filePath;
      using(StreamReader sr = new StreamReader(filePath)){
        string[] header = sr.ReadLine().Split(new[] { ',', '"' }, StringSplitOptions.RemoveEmptyEntries);
        hostname = header[1].Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries)[0];
        string empty = sr.ReadLine();
        string ddate = sr.ReadLine().Split(new[] { ',', '"' }, StringSplitOptions.RemoveEmptyEntries)[0];
        date = DateTime.Parse(ddate);        
      }
     
      //file.hostname //"(PDH-CSV 4.0) (Central Standard Time)(360)","\\DAC1\PhysicalDisk(_Total)\% Disk Time",
      //file.date (skip a couple down) //"06/30/2014 15:57:44.904"," 


    }
  }
}
