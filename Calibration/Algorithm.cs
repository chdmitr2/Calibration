using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;



namespace Calibration
{
    public class Algorithm
    {
        public List<int> SubBand { get; }
        public List<double> TunedFrequency { get; }
        public List<double> NotchRejection { get; }

        Excel excel = new Excel("C:\\Users\\chdmi\\OneDrive\\Desktop\\Full_Table2.xlsx", 1);
        Excel excelBest = new Excel("C:\\Users\\chdmi\\OneDrive\\Desktop\\BestNotch.xlsx", 1);

        public Algorithm()
        {
            int i = 1;
            string subband = "subband";
            List<int> SubBand = new List<int>();
            OpenFile();
            while(subband != string.Empty)
            {
                subband = excel.ReadCell(i, 3);
                if (subband == string.Empty) continue;
                else
                {
                    excelBest.WriteToCell(i, 3, subband);


                   // excel.SaveAs("C:\\Users\\chdmi\\OneDrive\\Desktop\\BestNotch.xlsx");
                   // excelBest.Save();
                  
                  
               
                   // excel.Close();
                    i++;
                    int numOfSubband = Int32.Parse(subband);
                    SubBand.Add(numOfSubband);
                }
                
            }
        }

        public Algorithm(List<int> subBand, List<double> tunedFrequency, List<double> notchRejection)
        {

        }

       

        public void OpenFile()
        {
            
           // MessageBox.Show(excel.ReadCell(1, 3));
           // excel.Close();

        }
    }
}
