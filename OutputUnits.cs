using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sherpanalyzer {
  public class OutputUnits {
    public string hostname;
    public string timestamp;
    public PCounter cpu;
    public PCounter mem;
    public PCounter net;
    public PCounter disk;
    public bool isLeaking;
    public bool highCPU;
    public OutputUnits(string host, string time, PCounter CPUstats, PCounter MEMstats, PCounter NETstats, PCounter DISKstats, bool leaking, bool cpuUsage){
      hostname = host;
      cpu = CPUstats;
      mem = MEMstats;
      net = NETstats;
      disk = DISKstats;
      timestamp = time;
      isLeaking = leaking;
      highCPU = cpuUsage;
  }

    public string display() {
      string result = "";
      result += hostname + " from " + timestamp + Environment.NewLine;
      result += hostname + " CPU(%) average: " + cpu.avg + Environment.NewLine;
      result += hostname + " CPU(%) max: " + cpu.max + Environment.NewLine;      
      result += hostname + " MEM(MB) average: " + mem.avg + Environment.NewLine;
      result += hostname + " MEM(MB) min: " + mem.min + Environment.NewLine;      
      result += hostname + " DISKTIME(%) average: " + disk.avg + Environment.NewLine;
      result += hostname + " DISKTIME(%) max: " + disk.max + Environment.NewLine;
      result += hostname + " NET Usage(B/sec) average: " + net.avg + Environment.NewLine;
      result += hostname + " NET Usage(B/sec) max: " + net.max + Environment.NewLine;
      result += hostname + " NET Usage(KB/sec) average: " + (net.avg/1000) + Environment.NewLine;
      result += hostname + " NET Usage(KB/sec) max: " + (net.max/1000) + Environment.NewLine;
      result += hostname + " NET Usage(Mbps) average: " + (net.avg * 0.000008) + Environment.NewLine;
      result += hostname + " NET Usage(Mbps) max: " + (net.max * 0.000008) + Environment.NewLine;
      if (isLeaking) {
        result += hostname + " has potential memory leaks!" + Environment.NewLine;
      }
      if (highCPU) {
        result += hostname + " has potential high CPU usage!" + Environment.NewLine;
      }
      return result;
    }
  }
}
