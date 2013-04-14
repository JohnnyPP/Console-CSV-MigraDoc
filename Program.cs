using System;
using System.Collections.Generic;
using System.IO;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LumenWorks.Framework.IO.Csv;
using System.Globalization;
using MathNet.Numerics.Statistics;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;

namespace ConsoleApplicationCSV
{



    class Program
    {
        #region ListWithMeasurements

        private static List<double> FW = new List<double>();
        private static List<double> FWN = new List<double>();
        private static List<double> HWL = new List<double>();
        private static List<double> HWNL = new List<double>();
        private static List<double> HWR = new List<double>();
        private static List<double> HWNR = new List<double>();

        #endregion



        static void Main(string[] args)
        {


            using (CachedCsvReader csv = new CachedCsvReader(new StreamReader(@"e:\1.txt"), false, '\t'))       //http://www.codeproject.com/Articles/9258/A-Fast-CSV-Reader?msg=1294084#xx1294084xx
            {                                                                                                   //CachedCsvReader is used for MoveTo() method     

                # region PDFinitialization

                Document document = new Document();
                Style style = document.Styles["Normal"];
                style.Font.Name = "Tahoma";
                style.Font.Size = 8;

                Section section = document.AddSection();
                //section.AddParagraph("Hello, World!");
               
                section.AddParagraph();

                Paragraph paragraph = section.AddParagraph();
                paragraph.Format.Font.Color = Color.FromCmyk(100, 30, 20, 50);
                //paragraph.AddFormattedText("Hello, World!", TextFormat.Underline);
                //AddFormattedText also returns an object:
                //FormattedText ft = paragraph.AddFormattedText("Small text",
                //  TextFormat.Bold);
                //ft.Font.Size = 6;

                csv.MoveTo(1);
                document.LastSection.AddParagraph("Sample: " + csv[1] + ", Date and time: " + csv[2], "Heading2");
                document.LastSection.AddParagraph();

                Table table = new Table();
                table.Borders.Width = 0.75;

                Column column = table.AddColumn(Unit.FromCentimeter(3));    //1st column
                column.Format.Alignment = ParagraphAlignment.Center;
                
                column.Table.AddColumn(Unit.FromCentimeter(2));             //2nd column
                column.Format.Alignment = ParagraphAlignment.Right;
                
                column.Table.AddColumn(Unit.FromCentimeter(2));             //3rd column
                column.Format.Alignment = ParagraphAlignment.Center;

                column.Table.AddColumn(Unit.FromCentimeter(2));             //4th column
                column.Format.Alignment = ParagraphAlignment.Center;

                column.Table.AddColumn(Unit.FromCentimeter(2));             //5th column
                column.Format.Alignment = ParagraphAlignment.Center;

                column.Table.AddColumn(Unit.FromCentimeter(2));             //6th column
                column.Format.Alignment = ParagraphAlignment.Center;

                column.Table.AddColumn(Unit.FromCentimeter(2));             //7th column
                column.Format.Alignment = ParagraphAlignment.Center;

                Row row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                Cell cell = row.Cells[0];

                cell.AddParagraph("Measurement no.");

                row.Cells[1].AddParagraph("FW");
                row.Cells[2].AddParagraph("FWN");
                row.Cells[3].AddParagraph("HWL");
                row.Cells[4].AddParagraph("HWNL");
                row.Cells[5].AddParagraph("HWR");
                row.Cells[6].AddParagraph("HWNR");

                # endregion PDFinitialization


              


                
                #region loops
                
                    for (int j = 1; j < 10; j++)
                    {


                        csv.MoveTo(j);          //row
                        //Console.Write(string.Format("{0} ", csv[i]));

                        FW.Add(double.Parse(csv[3], CultureInfo.InvariantCulture));
                        FWN.Add(double.Parse(csv[4], CultureInfo.InvariantCulture));
                        HWL.Add(double.Parse(csv[5], CultureInfo.InvariantCulture));
                        HWNL.Add(double.Parse(csv[6], CultureInfo.InvariantCulture));
                        HWR.Add(double.Parse(csv[7], CultureInfo.InvariantCulture));
                        HWNR.Add(double.Parse(csv[8], CultureInfo.InvariantCulture));
                            
                        table.AddRow();
                            
                            table.Rows[j].Cells[0].AddParagraph(j.ToString());
                            table.Rows[j].Cells[1].AddParagraph(csv[3]);
                            table.Rows[j].Cells[2].AddParagraph(csv[4]);
                            table.Rows[j].Cells[3].AddParagraph(csv[5]);
                            table.Rows[j].Cells[4].AddParagraph(csv[6]);
                            table.Rows[j].Cells[5].AddParagraph(csv[7]);
                            table.Rows[j].Cells[6].AddParagraph(csv[8]);
      

                    }

              
                #endregion

                var descrStatFW = new DescriptiveStatistics(FW);
                double KurtosisHW = descrStatFW.Kurtosis;
                double SkewnessHW = descrStatFW.Skewness;

                var descrStatFWN = new DescriptiveStatistics(FWN);
                double KurtosisFWN = descrStatFWN.Kurtosis;
                double SkewnessFWN = descrStatFWN.Skewness;

                var descrStatHWL = new DescriptiveStatistics(HWL);
                double KurtosisHWL = descrStatHWL.Kurtosis;
                double SkewnessHWL = descrStatHWL.Skewness;

                var descrStatHWNL = new DescriptiveStatistics(HWNL);
                double KurtosisHWNL = descrStatHWNL.Kurtosis;
                double SkewnessHWNL = descrStatHWNL.Skewness;

                var descrStatHWR = new DescriptiveStatistics(HWR);
                double KurtosisHWR = descrStatHWR.Kurtosis;
                double SkewnessHWR = descrStatHWR.Skewness;

                var descrStatHWNR = new DescriptiveStatistics(HWNR);
                double KurtosisHWNR = descrStatHWNR.Kurtosis;
                double SkewnessHWNR = descrStatHWNR.Skewness;


                var histogram = new Histogram(FW, 4);
                //double bucket3count = histogram[3].Count;
                //double bucket3count = histogram.UpperBound;
                //double bucket3count = histogram.DataCount;
                //double bucket3count = histogram.BucketCount;
                //double bucket3count = histogram[2].Width;                   //width of the bin
                double bucket3count = histogram[0].LowerBound;                //histogram[0].LowerBound=histogram.LowerBound

                Console.WriteLine();
                Console.WriteLine(string.Format("{0} ", FW.Mean()));
                Console.WriteLine();
                Console.WriteLine(string.Format("{0} ", FW.StandardDeviation()));
                
                Console.WriteLine();
                Console.WriteLine(string.Format("{0} ", KurtosisHW));
                Console.WriteLine();
                Console.WriteLine(string.Format("{0} ", SkewnessHW));
                Console.WriteLine();
                Console.WriteLine(string.Format("{0} ", bucket3count));


                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Mean");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(FW.Mean(),2).ToString(CultureInfo.InvariantCulture));
                table.Rows[10].Cells[2].AddParagraph(Math.Round(FWN.Mean(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[10].Cells[3].AddParagraph(Math.Round(HWL.Mean(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[10].Cells[4].AddParagraph(Math.Round(HWNL.Mean(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[10].Cells[5].AddParagraph(Math.Round(HWR.Mean(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[10].Cells[6].AddParagraph(Math.Round(HWNR.Mean(), 2).ToString(CultureInfo.InvariantCulture));
              

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Standard dev.");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(FW.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[11].Cells[2].AddParagraph(Math.Round(FWN.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[11].Cells[3].AddParagraph(Math.Round(HWL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[11].Cells[4].AddParagraph(Math.Round(HWNL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[11].Cells[5].AddParagraph(Math.Round(HWR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[11].Cells[6].AddParagraph(Math.Round(HWNR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));



                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("2x Standard dev.");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(2 * FW.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[12].Cells[2].AddParagraph(Math.Round(2 * FWN.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[12].Cells[3].AddParagraph(Math.Round(2 * HWL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[12].Cells[4].AddParagraph(Math.Round(2 * HWNL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[12].Cells[5].AddParagraph(Math.Round(2 * HWR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[12].Cells[6].AddParagraph(Math.Round(2 * HWNR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("3x Standard dev.");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(3 * FW.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[13].Cells[2].AddParagraph(Math.Round(3 * FWN.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[13].Cells[3].AddParagraph(Math.Round(3 * HWL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[13].Cells[4].AddParagraph(Math.Round(3 * HWNL.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[13].Cells[5].AddParagraph(Math.Round(3 * HWR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[13].Cells[6].AddParagraph(Math.Round(3 * HWNR.StandardDeviation(), 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Min");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(FW.Min(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[14].Cells[2].AddParagraph(Math.Round(FWN.Min(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[14].Cells[3].AddParagraph(Math.Round(HWL.Min(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[14].Cells[4].AddParagraph(Math.Round(HWNL.Min(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[14].Cells[5].AddParagraph(Math.Round(HWR.Min(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[14].Cells[6].AddParagraph(Math.Round(HWNR.Min(), 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Max");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(FW.Max(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[15].Cells[2].AddParagraph(Math.Round(FWN.Max(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[15].Cells[3].AddParagraph(Math.Round(HWL.Max(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[15].Cells[4].AddParagraph(Math.Round(HWNL.Max(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[15].Cells[5].AddParagraph(Math.Round(HWR.Max(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[15].Cells[6].AddParagraph(Math.Round(HWNR.Max(), 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Median");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(FW.Median() , 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[16].Cells[2].AddParagraph(Math.Round(FWN.Median(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[16].Cells[3].AddParagraph(Math.Round(HWL.Median(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[16].Cells[4].AddParagraph(Math.Round(HWNL.Median(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[16].Cells[5].AddParagraph(Math.Round(HWR.Median(), 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[16].Cells[6].AddParagraph(Math.Round(HWNR.Median(), 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Kurtosis");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(KurtosisHW, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[17].Cells[2].AddParagraph(Math.Round(KurtosisFWN, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[17].Cells[3].AddParagraph(Math.Round(KurtosisHWL, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[17].Cells[4].AddParagraph(Math.Round(KurtosisHWNL, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[17].Cells[5].AddParagraph(Math.Round(KurtosisHWR, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[17].Cells[6].AddParagraph(Math.Round(KurtosisHWNR, 2).ToString(CultureInfo.InvariantCulture));

                row = table.AddRow();
                row.Shading.Color = Colors.PaleGoldenrod;
                cell = row.Cells[0];
                cell.AddParagraph("Skewness");
                cell = row.Cells[1];
                cell.AddParagraph(Math.Round(SkewnessHW, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[18].Cells[2].AddParagraph(Math.Round(SkewnessFWN, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[18].Cells[3].AddParagraph(Math.Round(SkewnessHWL, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[18].Cells[4].AddParagraph(Math.Round(SkewnessHWNL, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[18].Cells[5].AddParagraph(Math.Round(SkewnessHWR, 2).ToString(CultureInfo.InvariantCulture));
                table.Rows[18].Cells[6].AddParagraph(Math.Round(SkewnessHWNR, 2).ToString(CultureInfo.InvariantCulture));


               

                //table.SetEdge(0, 0, 2, 3, Edge.Box, BorderStyle.Single, 1.5, Colors.Black);

                document.LastSection.Add(table);


                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(false,
                PdfFontEmbedding.Always);
                pdfRenderer.Document = document;
                pdfRenderer.RenderDocument();
                string filename = "HelloWorld5.pdf";
                pdfRenderer.PdfDocument.Save(filename);

            }

           

        }
    }
}
