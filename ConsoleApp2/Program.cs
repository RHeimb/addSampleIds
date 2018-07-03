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

        public static List<string> AlterSheet(string filename, string laborSetId, string lNumber = "_", string date = "_")
        {

            //string path = "C:\\Users\\robert\\Documents\\excelTests\\";
            //Workbooks.open Method only works with an absolute path!
            string path = filename;
            Excel.Workbook WB = AddSampleIdToBlvSheet.xApp.Workbooks.Open(Filename: path, UpdateLinks: 0, ReadOnly: false, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "\t", Editable: true, Notify: false, Converter: 0, AddToMru: true, Local: 1, CorruptLoad: 0);
            Excel.Worksheet WS = WB.Worksheets[1];
            Excel.Range userRange = WS.UsedRange;
            int recordCount = userRange.Rows.Count;
            Console.WriteLine("konkatiniere..." + path);

            List<string> suspiciousDatasets = new List<string>();

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
                        if (dataset.Contains(laborSetId.Substring(0, 7)) == false)
                        {
                            WS.Cells[i, "C"].Value = laborSetId + " " + dataset;
                        }
                        else
                        {                           
                            dataset = dataset.Replace(dataset.Substring(0, 14), laborSetId);    
                            WS.Cells[i, "C"].Value = dataset;

                        }
                    }
                }
            }
            else if (date == "_")
            {
                for (int i = 2; i <= recordCount; i++)
                {
                    string dataset = Convert.ToString(WS.Cells[i, "C"].Value);       
                    if (dataset.Contains(laborSetId.Substring(0, 7)) == false)
                    {
                        WS.Cells[i, "C"].Value = laborSetId + " " + dataset;
                    }
                    else
                    {
                        dataset = dataset.Replace(dataset.Substring(0, 14), laborSetId);
                        WS.Cells[i, "C"].Value = dataset;
                    }
                }
            }
            else
            {
                suspiciousDatasets.Add(Convert.ToString(WS.Cells[4, "C"].Value) + "--------------Pruefen!");
            }
            string check = Convert.ToString(WS.Cells[recordCount - 6, "C"].Value);
            if (check.Contains("S-BeLOV") == false) { suspiciousDatasets.Add(Convert.ToString(WS.Cells[recordCount - 6, "C"].Value) + "--------------LEER"); }

            WB.Save();
            WB.Close();
            Marshal.ReleaseComObject(WS);
            Marshal.ReleaseComObject(WB);
            return suspiciousDatasets;
        }
        public static void WriteToCsv(params object[] ListOfDatasetsAndCsvFile)
        {
            if (ListOfDatasetsAndCsvFile.Length == 2) {
                //add stringbuilder and write first entry
                var csv = new StringBuilder();
                string entry = ListOfDatasetsAndCsvFile[0].ToString();
                csv.AppendLine(entry);
                File.WriteAllText(ListOfDatasetsAndCsvFile[2].ToString(), csv.ToString());
            }
            else
            {
                //append to entry to given csv 
                StringBuilder csv = ListOfDatasetsAndCsvFile[1] as StringBuilder;
                string entry = ListOfDatasetsAndCsvFile[0].ToString();
                csv.AppendLine(entry);
                File.WriteAllText(ListOfDatasetsAndCsvFile[2].ToString(), csv.ToString());
            }
        }
        public static List<List<String>> GetListWithParamsForEachBlvId(string sourcePath) {
            var blvSampleID = new List<List<string>>();
            using (var reader = new StreamReader(sourcePath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    string[] values = line.Split(';');
                    //Datum Format t.m.xxxx -> xxxx-m-t
                    string date = values[3];
                    if (date.Contains("."))
                    {
                        string[] dateValues = date.Split('.');
                        if (date != "NULL") { date = dateValues[2] + "-" + dateValues[1] + "-" + dateValues[0]; }
                    }

                    blvSampleID.Add(new List<string> { values[0], values[2], date, values[4] }); //0=S-Belov; 1:L-0; 2=date; 3=BlvID
                }
            }
            return blvSampleID;
        }
        public static Dictionary<string,string> BuildDict(string sourcePath, string keyValue = "blv")
        {
            Dictionary<string, string> dict = new Dictionary<string, string>();

            using (var reader = new StreamReader(sourcePath))
            {
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    string[] values = line.Split(';');

                    if (keyValue == "blv")
                    {
                        if (!dict.ContainsKey(values[0])) { dict.Add(values[0], values[1]); } //BlvID:analyse-basic SampleID                       
                    }
                    if (keyValue == "sample")
                    {
                        if (!dict.ContainsKey(values[1])) { dict.Add(values[1], values[0]); } //analyse-basic:BlvID
                        if (!dict.ContainsKey(values[2])) { dict.Add(values[2], values[0]); } //CEC:BlvID
                        if (!dict.ContainsKey(values[3])) { dict.Add(values[3], values[0]); } //cytokines:BlvID
                    }
                }
            }
            return dict;
        }
        public static List<string> AlterSheetWithDict(string filename, Dictionary<string, string> blvDict, Dictionary<string,string> sampleDict)
        {
            string path = filename;
            Excel.Workbook WB = AddSampleIdToBlvSheet.xApp.Workbooks.Open(Filename: path, UpdateLinks: 0, ReadOnly: false, Format: 5, Password: "", WriteResPassword: "", IgnoreReadOnlyRecommended: true, Origin: Excel.XlPlatform.xlWindows, Delimiter: "\t", Editable: true, Notify: false, Converter: 0, AddToMru: true, Local: 1, CorruptLoad: 0);
            Excel.Worksheet WS = WB.Worksheets[1];
            Excel.Range userRange = WS.UsedRange;
            int recordCount = userRange.Rows.Count;
            Console.WriteLine("konkatiniere...");

            List<string> suspiciousDatasets = new List<string>();

            for (int i = 4; i <= recordCount; i++)
            {
                string dataset = Convert.ToString(WS.Cells[i, "D"].Value);
                if (dataset.Contains("BLV-"))
                {
                    string blvID = Convert.ToString(WS.Cells[i, "D"].Value);
                    blvID = blvID.Replace(" ", "");
                    blvID.ToLower();
                    if (blvDict.ContainsKey(blvID)) { WS.Cells[i, "C"].Value = blvDict[blvID]; }
                    else if (blvID.Length >= 14)
                    {
                        if (blvDict.ContainsKey(blvID.Substring(0, 14))) { WS.Cells[i, "C"].Value = blvDict[blvID.Substring(0, 14)]; }
                        else { suspiciousDatasets.Add(Convert.ToString(WS.Cells[i, "D"].Value)); }
                    }
                    else 
                    {
                        suspiciousDatasets.Add(Convert.ToString(WS.Cells[i, "D"].Value));
                    }
    }
                else if (dataset.Contains("S-BeLOV-")) 
                {
                    string sampleID = Convert.ToString(WS.Cells[i, "D"].Value);
                    sampleID = sampleID.Replace(" ", "");
                    sampleID.ToLower();
                    if (sampleDict.ContainsKey(sampleID))
                    {
                        string blvID = sampleDict[sampleID];
                        WS.Cells[i, "C"].Value = blvDict[blvID];
                    }
                    else { suspiciousDatasets.Add(sampleID); }
                }
                else
                {
                    suspiciousDatasets.Add(Convert.ToString(WS.Cells[i, "D"].Value));
                }
            }
            WB.Save();
            WB.Close();
            Marshal.ReleaseComObject(WS);
            Marshal.ReleaseComObject(WB);
            return suspiciousDatasets;
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
            Console.WriteLine("s = standard, h = haema");
            string choosed = Console.ReadLine();
            Console.WriteLine("Bitte Dateipfad zu durch Labvantage generierten .csv Datei angeben.");
            string sourcePath = Console.ReadLine();
            Console.WriteLine("Bitte Dateipfad angeben, der zu bearbeitende Excel-Dateien enthält");
            string workingPath = Console.ReadLine();
            Console.WriteLine("Bitte Dateipfad angeben, in dem die Fehlerzusammenfassung erstellt werden soll");
            string summaryOutPath = Console.ReadLine();

            List<string> Errors = new List<string>();

            if (choosed == "s")
            {
                List<List<string>> blvSampleID = GetListWithParamsForEachBlvId(sourcePath);

                for (int i = 0; i <= blvSampleID.Count() - 1; i++)
                {
                    List<string> potentialErrors = new List<string>();
                    string fn = blvSampleID[i][3].Substring(blvSampleID[i][3].Length - 6); //Benutze nur die Laufnummer der BLV-ID
                    string[] filesWithBlvId = Directory.GetFiles(workingPath, "*" + fn + "*");

                    if (filesWithBlvId.Length == 2)
                    {
                        int item;
                        for (item = 0; item <= filesWithBlvId.Length - 1; item++)
                        {
                            potentialErrors = AlterSheet(filesWithBlvId[item], blvSampleID[i][0], "_", blvSampleID[i][2]); //i+1 nur notlösung, da es im Ordner nur CEC-EPC gibt
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
                            potentialErrors = AlterSheet(filesWithLNumber[0], blvSampleID[i][0]);
                        }
                    }
                    else if (filesWithBlvId.Length == 1)
                    {
                        potentialErrors = AlterSheet(filesWithBlvId[0], blvSampleID[i][0]);
                    }

                    foreach (string entry in potentialErrors) { Errors.Add(entry); }
                }
            }

            else if (choosed == "h")
            {
                    Dictionary<string, string> dictBlvSample = BuildDict(sourcePath, "blv");
                    Dictionary<string, string> dictSampleBlv = BuildDict(sourcePath, "sample");
                    List<string> potentialErrors = new List<string>();

                    string[] tables = Directory.GetFiles(workingPath);

                    for (int i = 0; i < tables.Length; i++)
                    {
                        if (tables[i].Contains(".xl")) //um weitere filenames ergänzen
                        {
                            potentialErrors = AlterSheetWithDict(tables[i], dictBlvSample, dictSampleBlv);
                        }
                    }
                foreach (string entry in potentialErrors) { Errors.Add(entry); }
            }

            //write log
            StringBuilder csv = new StringBuilder();
            foreach (string entry in Errors) { WriteToCsv(entry, csv, summaryOutPath); }

            //tidy up. Kill every used Excel process
            Console.WriteLine("Fertig. Räume auf...");
            xApp.ScreenUpdating = true;
            xApp.Quit();
            Marshal.ReleaseComObject(xApp);
        }
    }
}
