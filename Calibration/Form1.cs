using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calibration
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Algorithm StartCalibration = new Algorithm();
            //OpenFile();
            //WriteData();
        }

       

      //  public void WriteData()
        //{
        //    Excel excel = new Excel("C:\\Users\\chdmi\\OneDrive\\Desktop\\BestNotch.xlsx", 1);
        //    excel.WriteToCell(0, 0, "Test2");
         //   excel.Save();
         //   excel.SaveAs("C:\\Users\\chdmi\\OneDrive\\Desktop\\BestNotch.xlsx");
         //   excel.Close();
       // }
        
    }
}
