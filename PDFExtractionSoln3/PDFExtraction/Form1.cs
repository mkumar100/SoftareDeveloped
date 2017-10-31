using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using System.IO;
using iTextSharp.text.pdf;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Configuration;
  
    
namespace PDFExtraction
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DateTime now = DateTime.Now;
        string xmlfilepath = ConfigurationManager.AppSettings["xmlfilepath"];
        //this.xmlfilepath = this.xmlfilepath 
        string logfilepath = ConfigurationManager.AppSettings["logfilepath"];
        string logfilename = null;
        string xmlfilename = null;

        private void button1_Click(object sender, EventArgs e)
        {
            //read vendor id, pdf names from excel
            //run loop
            //extract values from pdf
            //appraiseers
            //insert values in xml
            ///appraiseers
            ///
            //string excelfilename = @"C:\inputother\Vendor Document Location.xlsx";
            string excelfilename = ConfigurationManager.AppSettings["excelfilename"];
            string[] vendorIdArr = new string[] { };
            vendorIdArr = null;
            string[] fileNameArr = new string[] { };
            fileNameArr = null;
            int rowsCount = 0;
            int colsCount = 0;
            //string xmlfileName = "XML.xml";
            //string xmlfilepath = Path.Combine(@"C:\output\xml\", xmlfileName);
            //string xmlfilepath = @"C:\output\xml\XML.xml";
            //string xmlfilepath = ConfigurationManager.AppSettings["xmlfilepath"];//"C:\output\xml\" 
            xmlfilename = xmlfilepath + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xml";
            logfilename = logfilepath + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".log";

            string vendorId = null;
            string fileName = null;
            int code = 0;
            TakeValuesFromExcel(excelfilename, out vendorIdArr, out fileNameArr, out rowsCount, out colsCount);
            
            try
            {
                File.AppendAllText(xmlfilename, "<appraisers>" + Environment.NewLine);
                //Run(excelfilename, vendorIdArr, fileNameArr, out vendorId, out fileName);
                for (int i = 1; i < vendorIdArr.Count(); i++)
                {
                    fileName = fileNameArr[i].ToString();
                    vendorId = vendorIdArr[i].ToString();
                    ProcessPdf(vendorId, fileName, code, xmlfilename);
                    //File.AppendAllText(xmlfilepath, ProcessPdf(vendorId, fileName, code, sb).ToString());
                }
            }
            catch(Exception ex)
            {
                //File.AppendAllText(xmlfilepath, "</appraisers>" + Environment.NewLine);
                string logmsg = null;
                logmsg = vendorId + "," + fileName + "," + ex.Message;
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                string logfilePath = ConfigurationManager.AppSettings["logfilePath"];
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                File.AppendAllText(logfilename, logmsg + Environment.NewLine);
            }
            finally
            {
                File.AppendAllText(xmlfilename, "</appraisers>" + Environment.NewLine);
            }
            //TakeValuesFromExcel(excelfilename, out vendorIdArr, out fileNameArr, out rowsCount, out colsCount)
        }

        //private void ExtractInfoFromExcel(out string vendorId, out string fileName)
        //{
        //    //string excelfilename = @"C:\inputother\Vendor Document Location.xlsx";
        //    string excelfilename = ConfigurationManager.AppSettings["excelfilename"];

        //    string[] vendorIdArr = new string[] { };
        //    vendorIdArr = null;
        //    string[] fileNameArr = new string[] { };
        //    fileNameArr = null;
        //    vendorId = null;

        //    fileName = null;
        //    StringBuilder sb = new StringBuilder();
        //    try
        //    {
        //        //sb.Append("<appraisers>" + Environment.NewLine);
        //        //void Run(string excelfilename, string[] vendorIdArr, string[] fileNameArr, out string vendorId, out string fileName)
        //        Run(excelfilename, vendorIdArr, fileNameArr, out vendorId, out fileName);
        //        //sb.Append("</appraisers>" + Environment.NewLine);
        //    }
        //    catch
        //    {
        //        throw;
        //        //Run(excelfilename, vendorIdArr, fileNameArr,  rowsCount, colsCount);//run from the point at which exception was raised
        //        //sb.Append("</appraisers>" + Environment.NewLine);
        //    }
        //    finally
        //    {
                
        //    }
        //}

        private void Run(string excelfilename, string[] vendorIdArr, string[] fileNameArr, out string vendorId, out string fileName)
        {
            //string xmlfileName = "XML.xml";
            //string xmlfilepath = Path.Combine(@"C:\output\xml\", xmlfileName);
            //string xmlfilepath = ConfigurationManager.AppSettings["xmlfilepath"];

            int i = 0;

            fileName = null;
            vendorId = null;
            try
            {
                int code = 0;
                StringBuilder sb = new StringBuilder();
                //File.AppendAllText(xmlfilepath, "<appraisers>" + Environment.NewLine);
                for (i = 1; i < vendorIdArr.Count(); i++)
                {
                    fileName = fileNameArr[i].ToString();
                    vendorId = vendorIdArr[i].ToString();
                    ProcessPdf(vendorId, fileName, code, xmlfilename);
                    //File.AppendAllText(xmlfilepath, ProcessPdf(vendorId, fileName, code, sb).ToString());
                }
                //File.AppendAllText(xmlfilepath, "</appraisers>" + Environment.NewLine);
            }
            catch
            {
                //int index = i;
                //string logmsg = null;
                //logmsg = vendorId + "," + fileName + "," + ex.Message;
                ////string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //string logfilePath = ConfigurationManager.AppSettings["logfilePath"];
                ////string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //File.AppendAllText(logfilePath, logmsg + Environment.NewLine);
                RunRemaining(xmlfilename, excelfilename, vendorIdArr, fileNameArr, i+1);
                //throw;
                //string logmsg = null;
                //logmsg = vendorId + "," + fileName + "," + ex.Message;
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //File.AppendAllText(logfilePath, logmsg + Environment.NewLine);
                //File.AppendAllText(xmlfilepath, "</appraisers>" + Environment.NewLine);
                //Run(excelfilename, vendorIdArr, fileNameArr, rowsCount, columnsCount);
            }
        }

        private void RunRemaining(string xmlfilename, string excelfilename, string[] vendorIdArr, string[] fileNameArr, int index)
        {
            //string xmlfileName = "XML.xml";
            //string xmlfilepath = Path.Combine(@"C:\output\xml\", xmlfileName);

            string fileName = null;
            string vendorId = null;
            int i = 1;
            try
            {
                int code = 0;
                StringBuilder sb = new StringBuilder();
                //File.AppendAllText(xmlfilepath, "<appraisers>" + Environment.NewLine);
                for (i = index; i < vendorIdArr.Count(); i++)
                {
                    fileName = fileNameArr[i].ToString();
                    vendorId = vendorIdArr[i].ToString();
                    ProcessPdf(vendorId, fileName, code, xmlfilename);
                    //File.AppendAllText(xmlfilepath, ProcessPdf(vendorId, fileName, code, sb).ToString());
                }
                //File.AppendAllText(xmlfilepath, "</appraisers>" + Environment.NewLine);
            }
            catch(Exception ex)
            {
                string logmsg = null;
                logmsg = vendorId + "," + fileName + "," + ex.Message;
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //string logfilePath = ConfigurationManager.AppSettings["logfilePath"];
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //logfilename = logfilepath + now.ToString("yyyy-MM-dd_HH-mm-ss") + ".txt";
                File.AppendAllText(logfilename, logmsg + Environment.NewLine);
                RunRemaining(xmlfilename, excelfilename, vendorIdArr, fileNameArr, i+1);
                //throw;
                //string logmsg = null;
                //logmsg = vendorId + "," + fileName + "," + ex.Message;
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //File.AppendAllText(logfilePath, logmsg + Environment.NewLine);
                //File.AppendAllText(xmlfilepath, "</appraisers>" + Environment.NewLine);
                //Run(excelfilename, vendorIdArr, fileNameArr, rowsCount, columnsCount);
            }
        }

        private void TakeValuesFromExcel(string excelfilename, out string[] firstarr, out string[] otherarr, out int rowsCount, out int columnsCount)
        {
            Microsoft.Office.Interop.Excel.Application xlsApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlsApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
            }

            Excel.Workbook wb = xlsApp.Workbooks.Open(excelfilename,
                0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true);
            Excel.Sheets sheets = wb.Worksheets;
            Excel.Worksheet ws = (Excel.Worksheet)sheets.get_Item(1);

            Excel.Range firstColumnRange = ws.UsedRange.Columns[1];
            Excel.Range otherColumnRange = ws.UsedRange.Columns[4];
            try
            {
                rowsCount = ws.UsedRange.Rows.Count;
                columnsCount = ws.UsedRange.Columns.Count;
                System.Array myvalues = (System.Array)firstColumnRange.Cells.Value;
                System.Array othervalues = (System.Array)otherColumnRange.Cells.Value;
                firstarr = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
                otherarr = othervalues.OfType<object>().Select(o => o.ToString()).ToArray();
            }
            //catch
            //{
            //    firstarr = null;
            //    otherarr = null;
            //    rowsCount = 0;
            //    columnsCount = 0;
            //}
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                //Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(ws);
                Marshal.ReleaseComObject(sheets);
                //close and release
                wb.Close();
                Marshal.ReleaseComObject(wb);

                //quit and release
                xlsApp.Quit();
                Marshal.ReleaseComObject(xlsApp);
            }
        }

        //private StringBuilder ProcessPdf(string vendorId, string fileName, int code, StringBuilder sb)
        private void ProcessPdf(string vendorId, string fileName, int code, string xmlfilename)
        {
            string initialpdfpath = ConfigurationManager.AppSettings["initialpdfpath"];//C:\input\2013-7\
            //string completepdffileName = @"C:\input\2013-7\" + fileName;
            string completepdfpath = initialpdfpath + fileName;
            List<Appraiser> appraiserList = GetListAppraiser(completepdfpath, vendorId, out code);
            //sb.Append("<appraisers>" + Environment.NewLine);
            //StringBuilder sb2 = new StringBuilder();
            //try
            //{
            if(appraiserList!=null)
            foreach (Appraiser ap in appraiserList)
            {
                //sb.Append("<appraiser>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<appraiser>" + Environment.NewLine);
                //sb.Append("<vendorId>").Append(vendorId).Append("</vendorId>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<vendorId>" + vendorId + "</vendorId>" + Environment.NewLine);
                //sb.Append("<AppraiserName>").Append(ap.AppraiserName).Append("</AppraiserName>" + Environment.NewLine);//appraisernae
                File.AppendAllText(xmlfilename, "<AppraiserName>" + ap.AppraiserName + "</AppraiserName>" + Environment.NewLine);
                //sb.Append("<BusinessAddress>").Append(ap.BusinessAddress).Append("</BusinessAddress>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<BusinessAddress>" + ap.BusinessAddress + "</BusinessAddress>" + Environment.NewLine);
                //sb.Append("<CityStateZip>").Append(ap.CityStateZip).Append("</CityStateZip>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<CityStateZip>" + ap.CityStateZip + "</CityStateZip>" + Environment.NewLine);
                //sb.Append("<Email>"+ap.Email+"</Email>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<Email>" + ap.Email + "</Email>" + Environment.NewLine);
                //sb.Append("<LicenseState1>"+ap._203k+"</LicenseState1>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<LicenseState1>" + ap._203k + "</LicenseState1>" + Environment.NewLine);
                //sb.Append("<Numbe1>"+ap.Numbe1+"</Numbe1>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<Numbe1>" + ap.Numbe1 + "</Numbe1>" + Environment.NewLine);
                //sb.Append("<Exp1>"+ap.Exp1+"</Exp1>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<Exp1>" + ap.Exp1 + "</Exp1>" + Environment.NewLine);
                //sb.Append("<HUD_FHA>"+ap.HUD_FHA+"</HUD_FHA>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<HUD_FHA>" + ap.HUD_FHA + "</HUD_FHA>" + Environment.NewLine);
                //sb.Append("<HUD_FHAAverage>"+ap.HUD_FHA__Average+"</HUD_FHAAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<HUD_FHAAverage>" + ap.HUD_FHA__Average + "</HUD_FHAAverage>" + Environment.NewLine);
                //sb.Append("<_203k>"+ap._203k+"</_203k>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<_203k>" + ap._203k + "</_203k>" + Environment.NewLine);
                //sb.Append("<_203kAverage>"+ap._203k__Average+"</_203kAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<_203kAverage>" + ap._203k__Average + "</_203kAverage>" + Environment.NewLine);
                //sb.Append("<VA>"+ap.VA+"</VA>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<VA>" + ap.VA + "</VA>" + Environment.NewLine);
                //sb.Append("<VAAverage>"+ap.VAAverage+"</VAAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<VAAverage>" + ap.VAAverage + "</VAAverage>" + Environment.NewLine);
                //sb.Append("<USDA>"+ap.USDA+"</USDA>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<USDA>" + ap.USDA + "</USDA>" + Environment.NewLine);
                //sb.Append("<USDAAverage>"+ap.USDAAverage+"</USDAAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<USDAAverage>" + ap.USDAAverage + "</USDAAverage>" + Environment.NewLine);
                //sb.Append("<REO>"+ap.REO+"</REO>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<REO>" + ap.REO + "</REO>" + Environment.NewLine);
                //sb.Append("<REOAverage>"+ap.REOAverage+"</REOAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<REOAverage>" + ap.REOAverage + "</REOAverage>" + Environment.NewLine);
                File.AppendAllText(xmlfilename, "<ERC>" + ap.ERC + "</ERC>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<ERCAverage>" + ap.ERCAverage + "</ERCAverage>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<Luxury>" + ap.Luxury + "</Luxury>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<LuxuryAverage>" + ap.LuxuryAverage + "</LuxuryAverage>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<Manufactured>" + ap.Manufactured + "</Manufactured>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<ManufacturedAverage>" + ap.ManufacturedAverage + "</ManufacturedAverage>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<Green>" + ap.Green + "</Green>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<GreenAverage>" + ap.GreenAverage + "</GreenAverage>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<PlansSpecs>" + ap.PlansSpecs + "</PlansSpecs>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<PlansSpecsAverage>" + ap.Plans__Specs__Average + "</PlansSpecsAverage>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "<Date>" + ap.SignatureDate + "</Date>" + Environment.NewLine);

                File.AppendAllText(xmlfilename, "</appraiser>" + Environment.NewLine);

                //}

                //sb2 = sb;
                fileName = initialpdfpath + fileName;
                string msg = null;
                switch (code)
                {
                    case 1:
                        msg = vendorId + ", " + fileName + ", " + "Full" + ", " + "Complete";
                        //sb2 = sb;
                        break;
                    case 2:
                        msg = vendorId + ", " + fileName + ", " + "Partial" + ", " + "Incomplete";
                        //sb2 = sb;
                        break;
                    case 3:
                        msg = vendorId + ", " + fileName + ", " + "Skipped" + ", " + "Unreadable";
                        break;
                }
                //string logfilePath = Path.Combine(@"C:\output\log\", "log.txt");
                //string logfilePath = ConfigurationManager.AppSettings["logfilePath"];
                File.AppendAllText(logfilename, msg + Environment.NewLine);
                //}
                //catch
                //{
                //throw;
                //}
                //finally
                //{

                //}
                //return sb2;
            }
        }

        public static class Utility
        {
            public static List<Appraiser> GetAppraiserList(string appraiserName, string businessAddress, string cityStateZip, string email, string licenseState1, string numbe1,
            string exp1, string hUD_FHA, string hUD_FHA__Average, string p203k, string p203k__Average, string vA, string vAAverage, string uSDA, string uSDAAverage, string rEO,
            string rEOAverage, string eRC, string eRCAverage, string luxury, string luxuryAverage, string manufactured, string manufacturedAverage, string green, string greenAverage,
            string plansSpecs, string plans__Specs__Average, string signatureDate)
            {
                List<Appraiser> appraisers = new List<Appraiser>();
                Appraiser appraiser = new Appraiser
                {
                    AppraiserName = appraiserName,
                    BusinessAddress = businessAddress,
                    CityStateZip = cityStateZip,
                    Email = email,
                    LicenseState1 = licenseState1,
                    Numbe1 = numbe1,
                    Exp1 = exp1,
                    HUD_FHA = hUD_FHA,
                    HUD_FHA__Average = hUD_FHA,
                    _203k = p203k,
                    _203k__Average = p203k__Average,
                    VA = vA,
                    VAAverage = vAAverage,
                    USDA = uSDA,
                    USDAAverage = uSDAAverage,
                    REO = rEO,
                    REOAverage = rEOAverage,
                    ERC = eRC,
                    ERCAverage = eRCAverage,
                    Luxury = luxury,
                    LuxuryAverage = luxuryAverage,
                    Manufactured = manufactured,
                    ManufacturedAverage = manufacturedAverage,
                    Green = green,
                    GreenAverage = greenAverage,
                    PlansSpecs = plansSpecs,
                    Plans__Specs__Average = plans__Specs__Average,
                    SignatureDate = signatureDate
                };

                appraisers.Add(appraiser);
                return appraisers;
            }
        }

        public class Appraiser
        {
            public string VendorId { get; set; }
            public string AppraiserName { get; set; }
            public string BusinessAddress { get; set; }
            public string CityStateZip { get; set; }
            public string Email { get; set; }
            public string LicenseState1 { get; set; }
            public string Numbe1 { get; set; }
            public string Exp1 { get; set; }
            public string HUD_FHA { get; set; }
            public string HUD_FHA__Average { get; set; }
            public string _203k { get; set; }
            public string _203k__Average { get; set; }
            public string VA { get; set; }
            public string VAAverage { get; set; }
            public string USDA { get; set; }
            public string USDAAverage { get; set; }
            public string REO { get; set; }
            public string REOAverage { get; set; }
            public string ERC { get; set; }
            public string ERCAverage { get; set; }
            public string Luxury { get; set; }
            public string LuxuryAverage { get; set; }
            public string Manufactured { get; set; }
            public string ManufacturedAverage { get; set; }
            public string Green { get; set; }
            public string GreenAverage { get; set; }
            public string PlansSpecs { get; set; }
            public string Plans__Specs__Average { get; set; }
            public string SignatureDate { get; set; }
        }

        private List<Appraiser> GetListAppraiser(string fileName, string vendorId, out int code)
        {
            //PdfReader reader = new PdfReader(@"C:\Manish\PDFextraction\FilesReferenced\FromClient\fwpdfdataextraction\Valuation Connect Appraiser Application cab 20170907.pdf");
            PdfReader reader = null; //new PdfReader(fileName);
            AcroFields form = null;// reader.AcroFields;
            List<string> listFieldNames = new List<string>{"Appraiser Name", "Business Address", "CityStateZip", "Email", "LicenseState 1",
                "Numbe 1", "Exp 1", "HUD/FHA",
                "HUD/FHA Average", "203k", "203k Average", "VA", "VAAverage", "USDA", "USDAAverage", "REO", "REOAverage", "ERC", "ERCAverage", "Luxury", "LuxuryAverage",
                "Manufactured", "ManufacturedAverage", "Green", "GreenAverage", "PlansSpecs", "PlansSpecs Average", "Date"};
            List<Appraiser> listAppraisers = null;
            string joinedListstring = string.Join(",", listFieldNames);
            Dictionary<string, string> dict = new Dictionary<string, string>();
            StringBuilder sb = new StringBuilder();
            StringBuilder sbKeyWords = new StringBuilder();
            //string tmp = null;
            List<string> listAccrued = new List<string>();
            code = 0;
            FileInfo fileInfo = new FileInfo(fileName);
            if (fileInfo.Exists)
            {
                reader = new PdfReader(fileName);
                form = reader.AcroFields;
                for (int page = 1; page <= 1; page++)
                {
                    foreach (KeyValuePair<string, AcroFields.Item> kvp in form.Fields)
                    {
                        switch (form.GetFieldType(kvp.Key))
                        {
                            case AcroFields.FIELD_TYPE_CHECKBOX:
                            case AcroFields.FIELD_TYPE_COMBO:
                            case AcroFields.FIELD_TYPE_LIST:
                            case AcroFields.FIELD_TYPE_RADIOBUTTON:
                            case AcroFields.FIELD_TYPE_NONE:
                            case AcroFields.FIELD_TYPE_PUSHBUTTON:
                            case AcroFields.FIELD_TYPE_SIGNATURE:
                            case AcroFields.FIELD_TYPE_TEXT:
                                int fileType = form.GetFieldType(kvp.Key);
                                string translatedFileName = form.GetTranslatedFieldName(kvp.Key);
                                string fieldValue = form.GetField(kvp.Key);
                                bool chkKeyWord = joinedListstring.Contains(translatedFileName);
                                if (chkKeyWord)
                                {
                                    listAccrued.Add(translatedFileName);
                                    dict.Add(translatedFileName, fieldValue);
                                    sbKeyWords.Append(translatedFileName).Append(",");
                                    //tmp = translatedFileName + ", " + fieldValue;
                                    //sb.Append(tmp);
                                    //sb.Append(System.Environment.NewLine);
                                }
                                //else
                                //{
                                //    dict.Add(translatedFileName, null);
                                //}
                                //foreach (string str in listString)
                                //{
                                //    if (translatedFileName == str)
                                //        dict.Add(translatedFileName, fieldValue);
                                //    else
                                //        dict.Add(str, null);
                                //}
                                break;
                        }
                    }
                }
                string strKeyWords = sbKeyWords.ToString();
                List<string> result = listFieldNames.Except(listAccrued).ToList();

                if (listAccrued.Count == listFieldNames.Count)
                    code = 1;
                //msg = vendorId + string.Format(", Full, Complete{0}", Environment.NewLine);
                else if (listAccrued.Count < listFieldNames.Count && listAccrued.Count != 0)
                    code = 2;
                //msg = vendorId + string.Format(", Partial, Incomplete{0}", Environment.NewLine);
                else if (listAccrued.Count == 0)
                    code = 3;
                //msg = vendorId + string.Format(", Skipped, Unreadable{0}", Environment.NewLine);

                foreach (string str in result)
                {
                    dict.Add(str, null);
                }
                //listAccrued.Union(result).ToList();
                listAppraisers = Utility.GetAppraiserList(HandleNull(dict["Appraiser Name"]), HandleNull(dict["Business Address"]), HandleNull(dict["CityStateZip"]),
                                        HandleNull(dict["Email"]), HandleNull(dict["LicenseState 1"]), HandleNull(dict["Numbe 1"]), HandleNull(dict["Exp 1"]),
                                        HandleNull(dict["HUD/FHA"]), HandleNull(dict["HUD/FHA Average"]), HandleNull(dict["203k"]), HandleNull(dict["203k Average"]),
                                        HandleNull(dict["VA"]), HandleNull(dict["VAAverage"]), HandleNull(dict["USDA"]), HandleNull(dict["USDAAverage"]),
                                        HandleNull(dict["REO"]), HandleNull(dict["REOAverage"]), HandleNull(dict["ERC"]), HandleNull(dict["ERCAverage"]),
                                        HandleNull(dict["Luxury"]), HandleNull(dict["LuxuryAverage"]), HandleNull(dict["Manufactured"]), HandleNull(dict["ManufacturedAverage"]),
                                        HandleNull(dict["Green"]), HandleNull(dict["GreenAverage"]), HandleNull(dict["PlansSpecs"]), HandleNull(dict["PlansSpecs Average"]),
                                        HandleNull(dict["Date"]));
                //}
                //catch//(Exception ex)
                //{
                //throw;// MessageBox.Show(ex.Message)
                //}
                //finally
                //{
                reader.Close();
                //}
                return listAppraisers;
            }
            else
                return null;
        }

        private string HandleNull(string str)
        {
            string result =   str == null ? null : str;
            return result;
        }
    }
}
