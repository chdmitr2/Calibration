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
        public List<int> BinaryWord { get; }
        public List<double> TunedFrequency { get; }
        public List<double> NotchRejection { get; }
        public List<double> BandWidth { get; }

        Excel excel = new Excel("C:\\Users\\chdmi\\OneDrive\\Desktop\\Full_Table.xlsx", 1);
        Excel excelBest = new Excel("C:\\Users\\chdmi\\OneDrive\\Desktop\\BestNotch.xlsx", 1);

        public Algorithm()
        {
            List<int> SubBand = new List<int>();
            List<int> BinaryWord = new List<int>();
            List<double> TunedFrequency = new List<double>();
            List<double> NotchRejection = new List<double>();
            List<double> BandWidth = new List<double>();

            

            BinaryWord = GetData(BinaryWord, 2, 2);
            SubBand = GetData(SubBand, 2, 3);
            TunedFrequency = GetData(TunedFrequency, 2, 4);
            NotchRejection = GetData(NotchRejection, 2, 6);
            BandWidth = GetDataSplit(BandWidth, 2, 7);

            filterCal(SubBand, TunedFrequency, BinaryWord, NotchRejection, BandWidth);
            //printTable(SubBand, TunedFrequency, BinaryWord, NotchRejection, BandWidth);
            //printBestNotchTable(Output_SubBand, Output_TunedFrequency, Output_BinaryWord, Output_NotchRejection, Output_BandWidth);

            Console.WriteLine("END");
        }

            public List<int> GetData(List<int> list,int initRow,int initColumn)
            {
                string temp = "temp";
                while (temp != string.Empty)
                {
                    temp = excel.ReadCell(initRow, initColumn);
                    if (temp == string.Empty) continue;
                    else
                    {

                        initRow++;
                        int number = Int32.Parse(temp);
                        list.Add(number);
                    }

                }

                return list;
            }

            public List<double> GetData(List<double> list, int initRow, int initColumn)
            {
                string temp = "temp";
                while (temp != string.Empty)
                {
                    temp = excel.ReadCell(initRow, initColumn);
                    if (temp == string.Empty) continue;
                    else
                    {

                        initRow++;
                        double number = Math.Round(Convert.ToDouble(temp),3);
                        list.Add(number);
                    }

                }

                return list;
            }

        public List<double> GetDataSplit(List<double> list, int initRow, int initColumn)
        {
            string temp = "temp";
            while (temp != string.Empty)
            {
                temp = excel.ReadCell(initRow, initColumn);
                if (temp == string.Empty) continue;
                else
                {
                    string[] words = temp.Split(';');
                    initRow++;
                    double number = Math.Round(Convert.ToDouble(words[1]), 3);
                    list.Add(number);
                }

            }

            return list;
        }


       public void filterCal(List<int> subband,List<double> frequency,List<int> words,List<double> rejection,List<double> bandwidth)
       {
            List<int> Temp_SubBand = new List<int>();
            List<int> Temp_BinaryWord = new List<int>();
            List<double> Temp_TunedFrequency = new List<double>();
            List<double> Temp_NotchRejection = new List<double>();
            List<double> Temp_BandWidth = new List<double>();

            int buffer_subband = subband[15];
            int buffer_words = words[15];
            double buffer_frequency = frequency[15];
            double buffer_rejection = rejection[15];
            double buffer_bandwidth = bandwidth[15];

            List<int> Output_SubBand = new List<int>();
            List<int> Output_BinaryWord = new List<int>();
            List<double> Output_TunedFrequency = new List<double>();
            List<double> Output_NotchRejection = new List<double>();
            List<double> Output_BandWidth = new List<double>();

            //Console.WriteLine("          " + buffer_subband + "             " + buffer_frequency + "      " + buffer_words + "      " + buffer_rejection + "            " + buffer_bandwidth);

            int[] friquencies = Enumerable.Range(225, 288).ToArray();
            Console.WriteLine(" LNA1 SubBand      Frequency[MHz]  BinaryWord   Rejection[dB]   Gain Variation ");
            foreach (int freq in friquencies)
            {
                int row = 2;
                int index = 0;
                
                //Console.WriteLine("\n");
                //WriteLine(" Frequency " + freq + " MHz :\n");
                foreach (double fr in frequency)
                {
                    if (fr > freq - 1 && fr < freq + 1 && rejection[index] >= 20.5 && bandwidth[index] < 12)
                    {
                        Temp_TunedFrequency.Add(frequency[index]);
                        Temp_SubBand.Add(subband[index]);
                        Temp_BinaryWord.Add(words[index]);
                        Temp_NotchRejection.Add(rejection[index]);
                        Temp_BandWidth.Add(bandwidth[index]);
                    }
                    index++;
                }
                if (Temp_TunedFrequency.Count != 0)
                {
                        int minIndex = IndexOfMin(Temp_BandWidth);
                        Output_TunedFrequency.Add(Temp_TunedFrequency[minIndex]);
                        Output_SubBand.Add(Temp_SubBand[minIndex]);
                        Output_BinaryWord.Add(Temp_BinaryWord[minIndex]);
                        Output_NotchRejection.Add(Temp_NotchRejection[minIndex]);
                        Output_BandWidth.Add(Temp_BandWidth[minIndex]);

                        buffer_subband = Output_SubBand[0];
                        buffer_words = Output_BinaryWord[0];
                        buffer_frequency = Output_TunedFrequency[0];
                        buffer_rejection = Output_NotchRejection[0];
                        buffer_bandwidth = Output_BandWidth[0];


                }
                else
                {
                        Output_TunedFrequency.Add(buffer_frequency);
                        Output_SubBand.Add(buffer_subband);
                        Output_BinaryWord.Add(buffer_words);
                        Output_NotchRejection.Add(buffer_rejection);
                        Output_BandWidth.Add(buffer_bandwidth);
                }


                string bestSubband = Output_SubBand[0].ToString();
                string bestBinaryWord = Output_BinaryWord[0].ToString();
                string bestFrequency = freq.ToString();
                string bestRejection = Output_NotchRejection[0].ToString();
                string bestBandwidth = Output_BandWidth[0].ToString();

              /*  SetData(bestBinaryWord, row, 2);
                SetData(bestSubband, row, 3);
                SetData(bestFrequency, row, 4);
                SetData(bestRejection, row, 6);
                SetData(bestBandwidth, row, 7);
              */  row++;

                printTable(Output_SubBand, Output_TunedFrequency, Output_BinaryWord, Output_NotchRejection, Output_BandWidth);
                //printTable(Temp_SubBand, Temp_TunedFrequency, Temp_BinaryWord, Temp_NotchRejection, Temp_BandWidth);
                Temp_TunedFrequency.Clear();
                Temp_SubBand.Clear();
                Temp_BinaryWord.Clear();
                Temp_NotchRejection.Clear();
                Temp_BandWidth.Clear();
                Output_TunedFrequency.Clear();
                Output_SubBand.Clear();
                Output_BinaryWord.Clear();
                Output_NotchRejection.Clear();
                Output_BandWidth.Clear();
            }
       }


        public void SetData(string parameter, int Row, int Column)
        {
            excelBest.Close();
            
            excelBest.WriteToCell(Row, Column, parameter);
            excelBest.Save();
            excelBest.Close();
        }

        public static int IndexOfMin (List<double> list)
        {
            double min = list[0];
            int minIndex = 0;

            for (int i = 1; i < list.Count; ++i)
            {
                if(list [i] < min)
                {
                    min = list[i];
                    minIndex = i;
                }    
            }

            return minIndex;

        }
        public void printTable(List<int> subband, List<double> frequency, List<int> words, List<double> rejection, List<double> bandwidth)
        {
            int count = subband.Count();
            int[] friquencies = Enumerable.Range(0, count).ToArray();
           // Console.WriteLine(" LNA1 SubBand      Frequency[MHz]  BinaryWord   Rejection[dB]   Gain Variation ");
            foreach(int freq in friquencies)
            {
                Console.WriteLine("          " + subband[freq] + "             " + frequency[freq] + "      " + words[freq] + "      " + rejection[freq] + "            " + bandwidth[freq] );
            }

        }

    }       
}

