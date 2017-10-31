using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using iTextSharp.text.pdf;
using System.Configuration;

namespace ConsoleApplicationReadFromFolder
{
    class Program
    {
        static string folder = ConfigurationManager.AppSettings["folder"];
        static string now = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        static string logfilepath = ConfigurationManager.AppSettings["logfilepath"];
        //string logfilename = logfilepath + folder + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + "_log.txt";
        //"C:\output\log\2017-10_time_log.txt"
        static string logfilename = logfilepath + folder + "_" + now + "_log.txt";
        static void Main(string[] args)
        {
            //string[] filePaths = Directory.GetFiles(@"C:\input\2017-10", "*.pdf");
            
            string pdffilepath = ConfigurationManager.AppSettings["pdffilepath"] + folder + @"\";//"C:\input\"
            string[] pdffilenames = Directory.GetFiles(pdffilepath, "*.pdf");
 
            //<add key="xmlfilepath" value="C:\output\xml\"/>
            
            string xmlfilepath = ConfigurationManager.AppSettings["xmlfilepath"];// +folder + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xml";
            
            //string xmlfilename = @"C:\output\xml\2017-10\" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xml";

            //string xmlfilename = xmlfilepath + folder + "_" + DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss") + ".xml";
            string xmlfilename = xmlfilepath + folder + "_" + now + ".xml";
            //"C:\output\log\"
            
            string pdffilename = null;
            try
            {
                File.AppendAllText(xmlfilename, "<appraisers>" + Environment.NewLine);
            
                foreach (string pdf in pdffilenames)
                {
                    //Console.WriteLine(str);
                    pdffilename = pdf;
                    ProcessPdf(pdffilename, xmlfilename);
                }

            }
            catch (Exception ex)
            {
                string msg = pdffilename + ", " + ex.Message;
                //string logfilename = logfilepath + 
                File.AppendAllText(logfilename, msg + Environment.NewLine);
            }
            finally
            {
                File.AppendAllText(xmlfilename, "</appraisers>" + Environment.NewLine);
            }
            
        }

        static void ProcessPdf(string pdffilename, string xmlfilename)
        {
            List<Appraiser> appraiserList = null;// GetListAppraiser(completepdfpath);
            int code = 0;
            string msg = null;
            try
            {
                GetListAppraiser(pdffilename, out code, out appraiserList);

                foreach (Appraiser ap in appraiserList)
                {
                    if (ap.BusinessAddress != null && ap.CityStateZip != null && ap.Email != null)
                    {
                        //sb.Append("<appraiser>" + Environment.NewLine);
                        File.AppendAllText(xmlfilename, "<appraiser>" + Environment.NewLine);
                        //sb.Append("<vendorId>").Append(vendorId).Append("</vendorId>" + Environment.NewLine);
                        //File.AppendAllText(xmlfilename, "<vendorId>" + vendorId + "</vendorId>" + Environment.NewLine);
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
                    }

                    switch (code)
                    {
                        case 1:
                            //msg = vendorId + "," + fileName + "," + "Full" + "," + "Complete";
                            msg = pdffilename + "," + "Full" + "," + "Complete";
                            //sb2 = sb;
                            break;
                        case 2:
                            //msg = vendorId + "," + fileName + "," + "Partial" + "," + "Incomplete";
                            msg = pdffilename + "," + "Partial" + "," + "Incomplete";
                            //sb2 = sb;
                            break;
                        case 3:
                            msg = pdffilename + "," + "Skipped" + "," + "Unreadable";
                            break;
                    }
                    //string logfilepath = Path.Combine(@"C:\output\log\", "log.txt");
                    //string logfilename = @"C:\output\log\log.txt";
                    //logfilename = ConfigurationManager.AppSettings["logfilename"];
                    //string logfilename = logfilepath + 
                    File.AppendAllText(logfilename, msg + Environment.NewLine);

                }
            }
            catch
            {
                throw;
            }
            finally
            {

            }
            //return sb2;
        }

        static void GetListAppraiser(string fileName, out int code, out List<Appraiser> listAppraisers)
        {
            //PdfReader reader = new PdfReader(@"C:\Manish\PDFextraction\FilesReferenced\FromClient\fwpdfdataextraction\Valuation Connect Appraiser Application cab 20170907.pdf");
            PdfReader reader = null; //new PdfReader(fileName);
            AcroFields form = null;// reader.AcroFields;
            List<string> listFieldNames = new List<string>{"Appraiser Name", "Business Address", "CityStateZip", "Email", "LicenseState 1",
                "Numbe 1", "Exp 1", "HUD/FHA",
                "HUD/FHA Average", "203k", "203k Average", "VA", "VAAverage", "USDA", "USDAAverage", "REO", "REOAverage", "ERC", "ERCAverage", "Luxury", "LuxuryAverage",
                "Manufactured", "ManufacturedAverage", "Green", "GreenAverage", "PlansSpecs", "PlansSpecs Average", "Date"};
            listAppraisers = null;
            code = 0;
            string joinedListstring = string.Join(",", listFieldNames);
            Dictionary<string, string> dict = new Dictionary<string, string>();
            StringBuilder sb = new StringBuilder();
            StringBuilder sbKeyWords = new StringBuilder();
            //string tmp = null;
            List<string> listAccrued = new List<string>();
            //StringBuilder sb = new StringBuilder();
            //code = 0;
            try
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
                                string keyWord = form.GetTranslatedFieldName(kvp.Key);
                                sb.Append(keyWord + ", ");
                                string keyValue = form.GetField(kvp.Key);
                                sb.Append(keyValue + "; ");
                                bool chkKeyWord = joinedListstring.Contains(keyWord);
                                if (chkKeyWord)
                                {
                                    listAccrued.Add(keyWord);
                                    dict.Add(keyWord, keyValue);
                                    sbKeyWords.Append(keyWord).Append(",");
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
                string strn = sb.ToString();
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
                listAppraisers = GetAppraiserList(HandleNull(dict["Appraiser Name"]), HandleNull(dict["Business Address"]), HandleNull(dict["CityStateZip"]),
                                        HandleNull(dict["Email"]), HandleNull(dict["LicenseState 1"]), HandleNull(dict["Numbe 1"]), HandleNull(dict["Exp 1"]),
                                        HandleNull(dict["HUD/FHA"]), HandleNull(dict["HUD/FHA Average"]), HandleNull(dict["203k"]), HandleNull(dict["203k Average"]),
                                        HandleNull(dict["VA"]), HandleNull(dict["VAAverage"]), HandleNull(dict["USDA"]), HandleNull(dict["USDAAverage"]),
                                        HandleNull(dict["REO"]), HandleNull(dict["REOAverage"]), HandleNull(dict["ERC"]), HandleNull(dict["ERCAverage"]),
                                        HandleNull(dict["Luxury"]), HandleNull(dict["LuxuryAverage"]), HandleNull(dict["Manufactured"]), HandleNull(dict["ManufacturedAverage"]),
                                        HandleNull(dict["Green"]), HandleNull(dict["GreenAverage"]), HandleNull(dict["PlansSpecs"]), HandleNull(dict["PlansSpecs Average"]),
                                        HandleNull(dict["Date"]));
            }
            catch//(Exception ex)
            {
                listAppraisers = null;
                code = 0;
                throw;// MessageBox.Show(ex.Message)
            }
            finally
            {
                reader.Close();
            }
            //return listAppraisers;
        }

        static private string HandleNull(string str)
        {
            string result = str == null ? null : str;
            return result;
        }



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


        public class Appraiser
        {
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
    }
}
