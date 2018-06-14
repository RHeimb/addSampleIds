using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace addSampleID
{
    public class AddSampleIdToBlvSheet
    {
        //COM Application Object instanzieren
        public static Excel.Application xApp = new Excel.Application
        {
            ScreenUpdating = false
        };
        public static void AlterSheet(string filename, string laborSetId, string lNumber = "_", string date = "_")
        {

            //string path = "C:\\Users\\robert\\Documents\\excelTests\\";
            //Workbooks.open Method only works with an absolute path!
            string path = filename;
            Excel.Workbook WB = AddSampleIdToBlvSheet.xApp.Workbooks.Open(Filename: path, UpdateLinks: 0, ReadOnly: false, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "\t", Editable: true, Notify: false, Converter: 0, AddToMru: true, Local: 1, CorruptLoad: 0);
            Excel.Worksheet WS = WB.Worksheets[1];
            Excel.Range userRange = WS.UsedRange;
            int recordCount = userRange.Rows.Count;
            Console.WriteLine("konkatiniere...");

            //Iff there are more than one entry for a specific blvId, then use the date to distiguish between V1 and V2
            //If WS.Cells !contains like laborSetId???
            if (date != "_")
            {
                int i;
                for (i = 2; i <= recordCount; i++)
                {
                    string dataset = Convert.ToString(WS.Cells[i, "C"].Value); //important to use .Value since you get the com_object referance if you don't
                    if (dataset.Contains(date))
                    {
                        if (dataset.Contains(laborSetId.Substring(0, 6)) != true)
                        {
                            WS.Cells[i, "C"].Value = laborSetId + " " + dataset;
                        }
                    }
                }
            }
            else
            {
                int i;
                for (i = 2; i <= recordCount; i++)
                {
                    string dataset = Convert.ToString(WS.Cells[i, "C"].Value);
                    if (dataset.Contains(laborSetId) != true)
                    {
                        WS.Cells[i, "C"].Value = laborSetId + " " + dataset;
                    }
                }
            }
            WB.Save();
            WB.Close();
            Marshal.ReleaseComObject(WS);
            Marshal.ReleaseComObject(WB);
        }

        public static void Main()
        {
            //1.create list of 5-tuples from labvantage output
            //2.search the directory for files that contain the number part of the blvId
            //3.For each tuple there are 3 possible events
            //  3.1 there are 2 files which contain the blvID -> iterate over the array of filenames and pass date to alterSheet()
            //  3.2 there are no files which contain the blvID -> repeat step 2 with lNumber instead of blvID (yet it's to decide which part of the lNumber)
            //  3.3 there are 1 file which contains the blvID -> run alterSheet()

            //Workbooks.open Method only works with an absolute path!

            #region openExcelAndCreateTupleList_template
            /////////////////////Tests//////////////////////////////////////////////
            /*Excel.Application xApp = new Excel.Application();
            Excel.Workbook sourceWB = xApp.Workbooks.Open(@"C:\\Users\\robert\\Documents\\excelTests\\test.xlsx", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Excel.Worksheet sourceWS = sourceWB.Worksheets[1];
            Excel.Range userRange = sourceWS.UsedRange;
            int recordCount = userRange.Rows.Count;

            int i;
            for (i=1; i<= recordCount; i++)
            {
                string x = Convert.ToString(sourceWS.Cells[i, "A"].Value); //important to use .Value since you get the com_object name if you don't
                string y = Convert.ToString(sourceWS.Cells[i, "B"].Value);
                blvSampleID.Add((x,y));

            }
            sourceWB.Close(true, null, null);
            xApp.Quit();

            Marshal.ReleaseComObject(sourceWS);
            Marshal.ReleaseComObject(sourceWB);
            Marshal.ReleaseComObject(xApp);
            */
            //////////////////////////////////////////////////////////////////////
            #endregion

            Console.WriteLine("Instanziere COM-Objekt...");

            Console.WriteLine("Bitte Dateipfad zu durch Labvantage generierten .csv Datei angeben.");
            string sourcePath = Console.ReadLine();
            Console.WriteLine("Bitte Dateipfad angeben, der zu bearbeitende Excel-Dateien enthält");
            string workingPath = Console.ReadLine();

            var blvSampleID = new List<List<string>>();
            using (var reader = new StreamReader(sourcePath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    string[] values = line.Split(';');
                    //Datum Format t.m.xxxx -> xxxx-m-t
                    string date = values[3];
                    string[] dateValues = date.Split('.');
                    if (date != "NULL") {date = dateValues[2] + "-" + dateValues[1] + "-" + dateValues[0];}
                    

                    blvSampleID.Add(new List<string> { values[0], values[2], date, values[4] }); //0=S-Belov; 1:L-0; 2=date; 3=BlvID
                }
            }

            int i;
            for (i = 0; i <= blvSampleID.Count() - 2; i++) //-2 wegen der i+1 notlösung
            {
                if (i == 0 || blvSampleID[i][3] != blvSampleID[i - 1][3]) //Doppelte BLV-IDs aus dem Labvantage File werden übersprungen !Geht nur, wenn source absteigendnach L-Nummer sortiert wird! 
                {
                    string fn = blvSampleID[i][3].Substring(blvSampleID[i][3].Length - 6); //Benutze nur die Laufnummer der BLV-ID
                    string[] filesWithBlvId = Directory.GetFiles(workingPath, "*" + fn + "*");

                    if (filesWithBlvId.Length == 2)
                    {
                        int item;
                        for (item = 0; item <= filesWithBlvId.Length - 1; item++)
                        {
                            AlterSheet(filesWithBlvId[item], blvSampleID[i+1][0],"_", blvSampleID[i][2]); //i+1 nur notlösung, da es im Ordner nur CEC-EPC gibt
                        }
                    }
                    else if (filesWithBlvId.Length == 0)
                    {
                        string altFn = blvSampleID[i][1];
                        //vielleicht contains() auf directory files?
                        string[] filesWithLNumber = Directory.GetFiles(workingPath, "*" + altFn + "*");

                        int t;
                        for (t = 0; t <= filesWithLNumber.Length - 1; t++)
                        {
                            Console.WriteLine(filesWithLNumber[t]);//L-Nummer Tests fehlen
                        }

                        if (filesWithLNumber.Length != 0)
                        {
                            AlterSheet(filesWithLNumber[0], blvSampleID[i+1][0]);
                        }
                    }
                    else if (filesWithBlvId.Length == 1)
                    {
                        AlterSheet(filesWithBlvId[0], blvSampleID[i+1][0]);
                    }
                }
            }

            //tidy up. Kill every used Excel process
            Console.WriteLine("Fertig. Räume auf...");
            xApp.ScreenUpdating = true;
            xApp.Quit();
            Marshal.ReleaseComObject(xApp);
        }
    }
}
