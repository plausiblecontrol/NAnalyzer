using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sherpanalyzer {
  public class PCounter {
    public double avg;
    public double max;
    public double min;
    public PCounter(double average, double maximum, double minimum){
      avg = average;
      max = maximum;
      min = minimum;
  }
  }
}
