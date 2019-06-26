using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Net.Http;
using System.Net;
using Newtonsoft.Json.Linq;
using System.Data.OleDb;
using System.IO;
//using ImportTestCvs;
using System.Data;

namespace AllisonCoffmanConsoleAdaptor
{
    class AllisonCoffmanAdaptor
    {
        public static string GetProductDataService = ConfigurationManager.AppSettings["getProductDataService"];
        public static string apiUrl = string.Empty;
        public static string userName = string.Empty;
        public static string password = string.Empty;
        public static string referenceType = ConfigurationManager.AppSettings["sourceRefrenceType"];
        public static string id = string.Empty;

        #region Main Method
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Retriving Data From Web Api... Please Wait....");
                Factory objFactory = new Factory();
                DataSet ds = objFactory.GetRetailerName(3);
                foreach (DataTable table in ds.Tables)
                {
                    foreach (DataRow dr in table.Rows)
                    {
                        apiUrl = Convert.ToString(dr["FtpPath"]);
                        id = Convert.ToString(dr["RetailerId"]);
                        userName = Convert.ToString(dr["FtpUserName"]);
                        password = Convert.ToString(dr["FtpPassword"]);
                        //authHeadersToken = Convert.ToString(dr["Token"]);

                    }
                }

                string timeStamp = DateTime.Now.ToString("dd-MM-yyyy-HH-mm-ss");
                string projectLocation = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
                string directoryName = @"" + projectLocation + "/BulkUpload/" + referenceType + "/" + id + "/";
                string newDirectoryName = @"" + projectLocation + "/BulkUpload/" + referenceType + "/" + id + "/" + timeStamp + "/";
                //CreateDataExcel(newDirectoryName);
                CreateDataCSV(newDirectoryName);

                string responseData = GetProductDataByService(apiUrl, userName, password);

                JObject jsonObjectResponse = JObject.Parse(responseData);

                for (int i = 0; i < jsonObjectResponse["meta"].Children().LongCount(); i++)
                {
                    int total = (int)jsonObjectResponse["meta"]["total"];

                    Console.WriteLine("Total Product: " + total);
                    int lastPage = (int)jsonObjectResponse["meta"]["last_page"];
                    Console.WriteLine("Total Pages: " + lastPage);
                    Console.WriteLine("Product Per Page: 32");
                    for (int j = 0; j < lastPage; j++)
                    {
                        //Console.WriteLine("Retriving Data By Page Number..");
                        Console.WriteLine("Retriving Data For Page Number " + (j + 1));
                        string pageResponseData = GetProductDataByService(apiUrl, userName, password);
                        if (pageResponseData != "" && pageResponseData != null)
                        {
                            JObject jsonObjectPageResponse = JObject.Parse(pageResponseData);
                            ConvertResponseDataToAvalonData(jsonObjectPageResponse, newDirectoryName);
                        }
                        break;
                    }
                    break;
                }
                Console.WriteLine("Retriving Data From Web Api Completed.");
                //ImportUtility objImportUtility = new ImportUtility(newDirectoryName + "/" + "SimonGDataSheet.csv", "", Convert.ToInt32(id), "J", referenceType, "", "", timeStamp);
                //objImportUtility.INSERTCSV("D:/BulkUpload/Retailer/BulkUpload.xls");
                //objImportUtility.INSERTCSV(newDirectoryName + "/" + "SimonGDataSheet.csv");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                //throw ex;
                Console.ReadLine();
            }

        }
        #endregion

        #region GetDataByService
        public static string GetProductDataByService(string apiUrl,string userName, string password)
        {
            // string token = GetTokenForRequest();
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(apiUrl);
            request.Method = "GET";
            request.ContentType = "application/json";
            // request.Headers["Authorization"] = Token;
            request.Credentials = new System.Net.NetworkCredential(userName, password);


            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            WebHeaderCollection header = response.Headers;

            var encoding = ASCIIEncoding.ASCII;
            string responseText = string.Empty;
            using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
            {
                responseText = reader.ReadToEnd();
            }

            return responseText;
        }
        #endregion

        public static string GetTokenForRequest(int i)
        {

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(GetProductDataService + "&page=" + 1);
            request.Method = "GET";
            request.ContentType = "application/json";
            //request.Headers["Authorization"] = authHeadersToken;
            request.Credentials = new System.Net.NetworkCredential(userName, password);
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            WebHeaderCollection header = response.Headers;

            var encoding = ASCIIEncoding.ASCII;
            string responseText = string.Empty;
            using (var reader = new System.IO.StreamReader(response.GetResponseStream(), encoding))
            {
                responseText = reader.ReadToEnd();
            }

            return responseText;
        }

        #region Convert Service Response Data To Avalon Data Format
        public static void ConvertResponseDataToAvalonData(JObject jsonObjectResponse, string newDirectoryName)
        {
            int count = (int)jsonObjectResponse["data"].Children().LongCount();
            List<Model.ReponseDataToAvalonData> objResponseDataToAvalonDataList = new List<Model.ReponseDataToAvalonData>();

            for (int i = 0; i < jsonObjectResponse["data"].Children().LongCount(); i++)
            {
                for (int j = 0; j < jsonObjectResponse["data"][i]["available_metal"].Children().LongCount(); j++)
                {
                    Model.ReponseDataToAvalonData objResponseDataToAvalonData = new Model.ReponseDataToAvalonData();

                    string Collection_Name = ((string)(jsonObjectResponse["data"][i]["collection"])).Trim();
                    if (Collection_Name.ToLower() == "men ring")
                    {
                        Collection_Name = "MEN";
                    }
                    else if (Collection_Name.ToLower() == "nocturnal sophistication")
                    {
                        Collection_Name = "NOCTURNAL";
                    }

                    string Metal_Karat = ((string)jsonObjectResponse["data"][i]["variants"][j]["metal"]).Trim();
                    if (Metal_Karat.ToLower() == "14k")
                    {
                        Metal_Karat = "14 KT";
                    }
                    else if (Metal_Karat.ToLower() == "18k")
                    {
                        Metal_Karat = "18 KT";
                    }
                    else if (Metal_Karat.ToLower() == "plat")
                    {
                        Metal_Karat = "Platinum";
                    }

                    string Metal_Color = ((string)jsonObjectResponse["data"][i]["variants"][j]["metal_color"]).Trim();
                    if (Metal_Color.ToUpper() == "2T")
                    {
                        Metal_Color = "Yellow & White";
                    }
                    else if (Metal_Color.ToUpper() == "3T")
                    {
                        Metal_Color = "White & Yellow & Rose";
                    }
                    else if (Metal_Color.ToUpper() == "ROSE")
                    {
                        Metal_Color = "Rose";
                    }
                    else if (Metal_Color.ToUpper() == "WHITE")
                    {
                        Metal_Color = "White";
                    }
                    else if (Metal_Color.ToUpper() == "WHITE-BLACK" || Metal_Color.ToUpper() == "WHITE-BLAC")
                    {
                        Metal_Color = "White & Black";
                    }
                    else if (Metal_Color.ToUpper() == "WHITE-BROWN" || Metal_Color.ToUpper() == "WHITE-BROW")
                    {
                        Metal_Color = "White & Brown";
                    }
                    else if (Metal_Color.ToUpper() == "WHITE-ROSE")
                    {
                        Metal_Color = "White & Rose";
                    }
                    else if (Metal_Color.ToUpper() == "YELLOW")
                    {
                        Metal_Color = "Yellow";
                    }
                    if (jsonObjectResponse["data"][i]["available_metal"].Children().LongCount() > 1)
                    {
                        objResponseDataToAvalonData.PRODUCT_STYLE_ID = (string)jsonObjectResponse["data"][i]["id"] + "-" + Metal_Karat.Replace(" ", "").ToUpper();
                    }
                    else
                    {
                        objResponseDataToAvalonData.PRODUCT_STYLE_ID = (string)jsonObjectResponse["data"][i]["id"];
                    }
                    objResponseDataToAvalonData.SKU_ID = (string)jsonObjectResponse["data"][i]["id"];
                    objResponseDataToAvalonData.PRODUCT_NAME = (string)jsonObjectResponse["data"][i]["name"];
                    objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION = (string)jsonObjectResponse["data"][i]["description"];
                    objResponseDataToAvalonData.CATEGORY_INFO = "Jewelry > Engagement";
                    objResponseDataToAvalonData.DESIGNER_COLLECTION = ("Simon G > " + Collection_Name.Replace(" ", "-")).ToUpper();
                    objResponseDataToAvalonData.COST_PRICE = ((string)jsonObjectResponse["data"][i]["variants"][j]["price"]).Replace(",", "");
                    objResponseDataToAvalonData.SELL_PRICE = ((string)jsonObjectResponse["data"][i]["variants"][j]["price"]).Replace(",", "");
                    objResponseDataToAvalonData.METAL_KARAT = Metal_Karat;
                    objResponseDataToAvalonData.METAL_COLOR = Metal_Color;
                    objResponseDataToAvalonData.STATUS = (string)jsonObjectResponse["data"][i]["published"];
                    for (int k = 0; k < jsonObjectResponse["data"][i]["images"].Children().LongCount(); k++)
                    {
                        if (jsonObjectResponse["data"][i]["available_metal"].Children().LongCount() > 1)
                        {
                            if (k == 0)
                            {
                                if (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString() != null)
                                {
                                    string[] imageArray = (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString().Replace("\r\n", "").Replace("]", "").Replace("{", "").Replace("}", "").Replace("\"", "")).Split('[');
                                    if (imageArray.Length > 1)
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_1 = imageArray[1].Trim();
                                    }
                                    else
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_1 = string.Empty;
                                    }
                                }
                                else
                                {
                                    objResponseDataToAvalonData.IMAGE_NAME_1 = string.Empty;
                                }
                            }
                            else if (k == 1)
                            {
                                if (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString() != null)
                                {
                                    string[] imageArray = (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString().Replace("\r\n", "").Replace("]", "").Replace("{", "").Replace("}", "").Replace("\"", "")).Split('[');
                                    if (imageArray.Length > 1)
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_2 = imageArray[1].Trim();
                                    }
                                    else
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_2 = string.Empty;
                                    }
                                }
                                else
                                {
                                    objResponseDataToAvalonData.IMAGE_NAME_2 = string.Empty;
                                }
                            }
                            else if (k == 2)
                            {
                                if (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString() != null)
                                {
                                    string[] imageArray = (jsonObjectResponse["data"][i]["variants"][j]["images"].ToString().Replace("\r\n", "").Replace("]", "").Replace("{", "").Replace("}", "").Replace("\"", "")).Split('[');
                                    if (imageArray.Length > 1)
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_3 = imageArray[1].Trim();
                                    }
                                    else
                                    {
                                        objResponseDataToAvalonData.IMAGE_NAME_3 = string.Empty;
                                    }
                                }
                                else
                                {
                                    objResponseDataToAvalonData.IMAGE_NAME_3 = string.Empty;
                                }
                            }
                        }
                        else
                        {
                            if (k == 0)
                            {
                                objResponseDataToAvalonData.IMAGE_NAME_1 = (string)jsonObjectResponse["data"][i]["images"][k];
                            }
                            else if (k == 1)
                            {
                                objResponseDataToAvalonData.IMAGE_NAME_2 = (string)jsonObjectResponse["data"][i]["images"][k];
                            }
                            else if (k == 2)
                            {
                                objResponseDataToAvalonData.IMAGE_NAME_3 = (string)jsonObjectResponse["data"][i]["images"][k];
                            }
                        }
                    }
                    objResponseDataToAvalonDataList.Add(objResponseDataToAvalonData);
                }
            }
            //InsertDataIntoAvalonExcel(objResponseDataToAvalonDataList, newDirectoryName + "/" + "SimonGDataSheet.xls");
            InsertDataIntoAvalonCSV(objResponseDataToAvalonDataList, newDirectoryName + "/" + "SimonGDataSheet.csv");
        }
        #endregion

        #region Create CSV For Avalon Bulk Upload
        private static void CreateDataCSV(string path)
        {
            string directory = path;
            if (!string.IsNullOrEmpty("SimonGDataSheet.csv"))
            {
                directory = path.Replace("SimonGDataSheet.csv", "");
            }
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }
            string LogFileName = path + "SimonGDataSheet.csv";
            if (File.Exists(LogFileName) == false)
            {
                FileStream fs = File.Create(LogFileName);
                fs.Close();
            }
            string delimiter = ",";
            string createText = "PRODUCT_STYLE_ID" + delimiter + "PRODUCT_NAME" + delimiter + "WEB_SHORT_DESCRIPTION" + delimiter + "CATEGORY_INFO" + delimiter + "DESIGNER_COLLECTION" + delimiter + "IMAGE_NAME_1" + delimiter + "IMAGE_NAME_2" + delimiter + "IMAGE_NAME_3" + delimiter + "COST_PRICE" + delimiter + "SELL_PRICE" + delimiter + "METAL_KARAT" + delimiter + "METAL_COLOR" + delimiter + "PUBLISH" + delimiter + "PRICE_DISPLAY" + delimiter + "IF_NO_THEN_MESSAGE" + delimiter + "SKU_ID" + delimiter + Environment.NewLine;
            File.WriteAllText(LogFileName, createText);
        }
        #endregion

        #region Insert Into Avalon CSV Column Row By Row
        private static void InsertDataIntoAvalonCSV(List<Model.ReponseDataToAvalonData> objResponseDataToAvalonDataList, string path)
        {
            try
            {
                string delimiter = ",";
                foreach (Model.ReponseDataToAvalonData objResponseDataToAvalonData in objResponseDataToAvalonDataList)
                {
                    Model.ReponseDataToAvalonData objResponseData = ReplaceSpecialCharacter(objResponseDataToAvalonData);

                    string createText = objResponseData.PRODUCT_STYLE_ID + delimiter + objResponseData.PRODUCT_NAME + delimiter + objResponseData.WEB_SHORT_DESCRIPTION + delimiter + objResponseData.CATEGORY_INFO + delimiter + objResponseData.DESIGNER_COLLECTION + delimiter + objResponseData.IMAGE_NAME_1 + delimiter + objResponseData.IMAGE_NAME_2 + delimiter + objResponseData.IMAGE_NAME_3 + delimiter + objResponseData.COST_PRICE.Replace("$", "") + delimiter + objResponseData.SELL_PRICE.Replace("$", "") + delimiter + objResponseData.METAL_KARAT + delimiter + objResponseData.METAL_COLOR + delimiter + objResponseDataToAvalonData.STATUS + delimiter + "Yes" + delimiter + "Call Store for Price" + delimiter + objResponseData.SKU_ID + delimiter + Environment.NewLine;
                    File.AppendAllText(path, createText);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
        #endregion

        #region Create Excel For Avalon Bulk Upload
        private static void CreateDataExcel(string path)
        {
            FileInfo file = new FileInfo(path);
            string replaceFileName = file.Name;

            string directory = path;
            if (!string.IsNullOrEmpty(replaceFileName))
            {
                directory = path.Replace(replaceFileName, "");
            }

            //string LogFileName = "LogUploadedCsvTemp.xls";
            string LogFileName = "SimonGDataSheet.xls";
            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string LogFilePath = directory + @"\" + LogFileName;
            String strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + LogFilePath + ";" + "Extended Properties='Excel 8.0;HDR=Yes'";

            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();
            cmdExcel.Connection = connExcel;

            if (File.Exists(LogFilePath))
            {
                File.Delete(LogFilePath);
            }
            //******************Accessing Sheets*********************//

            connExcel.Open();

            //dt = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            connExcel.Close();

            //****************Create a new sheet*********************//
            string sheetname = "DataBulkUpload";
            try
            {
                //PRODUCT_STYLE_ID,SKU_ID,PRODUCT_NAME,DESIGNER_COLLECTION,CATEGORY_INFO,IMAGE_NAME_1,IMAGE_NAME_2,IMAGE_NAME_3,PRICE_DISPLAY,IF_NO_THEN_MESSAGE,COST_PRICE,SELL_PRICE,SALE_PRICE,WEB_SHORT_DESCRIPTION,JEWELRY_TYPE,PUBLISH

                cmdExcel.CommandText = "CREATE TABLE " + sheetname + "" + "(PRODUCT_STYLE_ID varchar(22),SKU_ID varchar(22),PRODUCT_NAME varchar(200),DESIGNER_COLLECTION varchar(100),CATEGORY_INFO_1 varchar(200),CATEGORY_INFO_2 varchar(200),CATEGORY_INFO_3 varchar(200),CATEGORY_INFO_4 varchar(200),CATEGORY_INFO_5 varchar(200),IMAGE_NAME_1 varchar(100),IMAGE_NAME_2 varchar(100),IMAGE_NAME_3 varchar(100),PRICE_DISPLAY varchar(10),IF_NO_THEN_MESSAGE varchar(50),COST_PRICE varchar(50),SELL_PRICE varchar(50),SALE_PRICE varchar(50),SALE_PRICE_START_DATE varchar(100),SALE_PRICE_END_DATE varchar(100),WEB_SHORT_DESCRIPTION_1 varchar(200),WEB_SHORT_DESCRIPTION_2 varchar(200),JEWELRY_TYPE varchar(50),PUBLISH varchar(1));";
                connExcel.Open();
                cmdExcel.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                string exe = "Please Enter Sheet Name In Order To Create Sheet In Excel...Thank You!";
                Console.WriteLine(ex);
                Console.ReadLine();
            }
            connExcel.Close();
        }
        #endregion

        #region Insert Into Avalon Excel Column Row By Row
        private static void InsertDataIntoAvalonExcel(List<Model.ReponseDataToAvalonData> objResponseDataToAvalonDataList, string path)
        {
            string LogFilePath = path;
            OleDbConnection MyConnection;
            OleDbCommand myCommand = new OleDbCommand();
            string sql = null;
            MyConnection = new OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + LogFilePath.ToUpper().Replace(".CSV", ".xls") + ";Extended Properties=Excel 8.0;");
            MyConnection.Open();
            myCommand.Connection = MyConnection;
            try
            {
                foreach (Model.ReponseDataToAvalonData objResponseDataToAvalonData in objResponseDataToAvalonDataList)
                {
                    Model.ReponseDataToAvalonData objResponseData = ReplaceSpecialCharacter(objResponseDataToAvalonData);
                    //PRODUCT_STYLE_ID varchar(22),PRODUCT_NAME varchar(100),DESIGNER_COLLECTION varchar(100),CATEGORY_INFO varchar(200),IMAGE_NAME_1 varchar(100),IMAGE_NAME_2 varchar(100),IMAGE_NAME_3 varchar(100),COST_PRICE varchar(10),SELL_PRICE varchar(10),WEB_SHORT_DESCRIPTION varchar(200),Publish varchar(1)
                    string WEB_SHORT_DESCRIPTION_1 = string.Empty;
                    string WEB_SHORT_DESCRIPTION_2 = string.Empty;
                    if (objResponseData.WEB_SHORT_DESCRIPTION != null && objResponseData.WEB_SHORT_DESCRIPTION.Length >= 300)
                    {
                        WEB_SHORT_DESCRIPTION_1 = objResponseData.WEB_SHORT_DESCRIPTION.Substring(0, 199);
                        WEB_SHORT_DESCRIPTION_2 = objResponseData.WEB_SHORT_DESCRIPTION.Substring(199, 101);
                    }
                    else if (objResponseData.WEB_SHORT_DESCRIPTION != null && objResponseData.WEB_SHORT_DESCRIPTION.Length < 199)
                    {
                        WEB_SHORT_DESCRIPTION_1 = objResponseData.WEB_SHORT_DESCRIPTION;
                    }
                    else if (objResponseData.WEB_SHORT_DESCRIPTION != null && objResponseData.WEB_SHORT_DESCRIPTION.Length < 300 && objResponseData.WEB_SHORT_DESCRIPTION.Length > 199)
                    {
                        WEB_SHORT_DESCRIPTION_1 = objResponseData.WEB_SHORT_DESCRIPTION.Substring(0, 199);
                        WEB_SHORT_DESCRIPTION_2 = objResponseData.WEB_SHORT_DESCRIPTION.Substring(199, objResponseData.WEB_SHORT_DESCRIPTION.Length - 199);
                    }

                    string CATEGORY_INFO_1 = string.Empty;
                    string CATEGORY_INFO_2 = string.Empty;
                    string CATEGORY_INFO_3 = string.Empty;
                    string CATEGORY_INFO_4 = string.Empty;
                    string CATEGORY_INFO_5 = string.Empty;
                    if (objResponseData.CATEGORY_INFO.Length < 0)
                    {
                        if (objResponseData.CATEGORY_INFO.Length >= 1000)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO.Substring(0, 200);
                            CATEGORY_INFO_2 = objResponseData.CATEGORY_INFO.Substring(200, 200);
                            CATEGORY_INFO_3 = objResponseData.CATEGORY_INFO.Substring(400, 200);
                            CATEGORY_INFO_4 = objResponseData.CATEGORY_INFO.Substring(600, 200);
                            CATEGORY_INFO_5 = objResponseData.CATEGORY_INFO.Substring(800, 200);
                        }
                        else if (objResponseData.CATEGORY_INFO.Length <= 200)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO;
                        }
                        else if (objResponseData.CATEGORY_INFO.Length >= 201 && objResponseData.CATEGORY_INFO.Length <= 400)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO.Substring(0, 200);
                            CATEGORY_INFO_2 = objResponseData.CATEGORY_INFO.Substring(200, objResponseData.CATEGORY_INFO.Length - 200);
                        }
                        else if (objResponseData.CATEGORY_INFO.Length >= 401 && objResponseData.CATEGORY_INFO.Length <= 600)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO.Substring(0, 200);
                            CATEGORY_INFO_2 = objResponseData.CATEGORY_INFO.Substring(200, 200);
                            CATEGORY_INFO_3 = objResponseData.CATEGORY_INFO.Substring(400, objResponseData.CATEGORY_INFO.Length - 400);
                        }
                        else if (objResponseData.CATEGORY_INFO.Length >= 601 && objResponseData.CATEGORY_INFO.Length <= 800)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO.Substring(0, 200);
                            CATEGORY_INFO_2 = objResponseData.CATEGORY_INFO.Substring(200, 200);
                            CATEGORY_INFO_3 = objResponseData.CATEGORY_INFO.Substring(400, 200);
                            CATEGORY_INFO_4 = objResponseData.CATEGORY_INFO.Substring(600, objResponseData.CATEGORY_INFO.Length - 600);
                        }
                        else if (objResponseData.CATEGORY_INFO.Length >= 801 && objResponseData.CATEGORY_INFO.Length <= 1000)
                        {
                            CATEGORY_INFO_1 = objResponseData.CATEGORY_INFO.Substring(0, 200);
                            CATEGORY_INFO_2 = objResponseData.CATEGORY_INFO.Substring(200, 200);
                            CATEGORY_INFO_3 = objResponseData.CATEGORY_INFO.Substring(400, 200);
                            CATEGORY_INFO_4 = objResponseData.CATEGORY_INFO.Substring(600, 200);
                            CATEGORY_INFO_5 = objResponseData.CATEGORY_INFO.Substring(800, objResponseData.CATEGORY_INFO.Length - 800);
                        }
                    }

                    sql = "Insert into [DataBulkUpload$] (PRODUCT_STYLE_ID,SKU_ID,PRODUCT_NAME,DESIGNER_COLLECTION,CATEGORY_INFO_1,CATEGORY_INFO_2,CATEGORY_INFO_3,CATEGORY_INFO_4,CATEGORY_INFO_5,IMAGE_NAME_1,IMAGE_NAME_2,IMAGE_NAME_3,PRICE_DISPLAY,IF_NO_THEN_MESSAGE,COST_PRICE,SELL_PRICE,SALE_PRICE,SALE_PRICE_START_DATE,SALE_PRICE_END_DATE,WEB_SHORT_DESCRIPTION_1,WEB_SHORT_DESCRIPTION_2,JEWELRY_TYPE,PUBLISH) values('" + objResponseData.PRODUCT_STYLE_ID + "','" + objResponseData.SKU_ID + "','" + objResponseData.PRODUCT_NAME + "','" + objResponseData.DESIGNER_COLLECTION + "','" + CATEGORY_INFO_1 + "','" + CATEGORY_INFO_2 + "','" + CATEGORY_INFO_3 + "','" + CATEGORY_INFO_4 + "','" + CATEGORY_INFO_5 + "','" + objResponseData.IMAGE_NAME_1 + "','" + objResponseData.IMAGE_NAME_2 + "','" + objResponseData.IMAGE_NAME_3 + "','" + "Yes" + "','" + "Call Store for Price" + "','" + objResponseData.COST_PRICE + "','" + objResponseData.SELL_PRICE + "','" + objResponseData.SALE_PRICE + "','" + objResponseData.SALE_PRICE_START_DATE + "','" + objResponseData.SALE_PRICE_END_DATE + "','" + WEB_SHORT_DESCRIPTION_1 + "','" + WEB_SHORT_DESCRIPTION_2 + "','" + objResponseData.JEWELRY_TYPE + "','" + objResponseData.STATUS + "')";
                    myCommand.CommandText = sql;
                    myCommand.ExecuteNonQuery();
                }
                //  file = null;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                MyConnection.Close();
            }
        }
        #endregion

        #region Replace Special Characters
        public static Model.ReponseDataToAvalonData ReplaceSpecialCharacter(Model.ReponseDataToAvalonData objResponseDataToAvalonData)
        {
            objResponseDataToAvalonData.PRODUCT_STYLE_ID = objResponseDataToAvalonData.PRODUCT_STYLE_ID == null ? string.Empty : objResponseDataToAvalonData.PRODUCT_STYLE_ID;
            objResponseDataToAvalonData.PRODUCT_NAME = objResponseDataToAvalonData.PRODUCT_NAME == null ? string.Empty : objResponseDataToAvalonData.PRODUCT_NAME;
            objResponseDataToAvalonData.CATEGORY_INFO = objResponseDataToAvalonData.CATEGORY_INFO == null ? string.Empty : objResponseDataToAvalonData.CATEGORY_INFO;
            objResponseDataToAvalonData.DESIGNER_COLLECTION = objResponseDataToAvalonData.DESIGNER_COLLECTION == null ? string.Empty : objResponseDataToAvalonData.DESIGNER_COLLECTION;
            objResponseDataToAvalonData.IMAGE_NAME_1 = objResponseDataToAvalonData.IMAGE_NAME_1 == null ? string.Empty : objResponseDataToAvalonData.IMAGE_NAME_1;
            objResponseDataToAvalonData.IMAGE_NAME_2 = objResponseDataToAvalonData.IMAGE_NAME_2 == null ? string.Empty : objResponseDataToAvalonData.IMAGE_NAME_2;
            objResponseDataToAvalonData.IMAGE_NAME_3 = objResponseDataToAvalonData.IMAGE_NAME_3 == null ? string.Empty : objResponseDataToAvalonData.IMAGE_NAME_3;
            objResponseDataToAvalonData.CATEGORY_INFO = objResponseDataToAvalonData.CATEGORY_INFO == null ? string.Empty : objResponseDataToAvalonData.CATEGORY_INFO;
            objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION = objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION == null ? string.Empty : objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION;
            objResponseDataToAvalonData.COST_PRICE = objResponseDataToAvalonData.COST_PRICE == null ? string.Empty : objResponseDataToAvalonData.COST_PRICE;
            objResponseDataToAvalonData.SELL_PRICE = objResponseDataToAvalonData.SELL_PRICE == null ? string.Empty : objResponseDataToAvalonData.SELL_PRICE;
            objResponseDataToAvalonData.METAL_KARAT = objResponseDataToAvalonData.METAL_KARAT == null ? string.Empty : objResponseDataToAvalonData.METAL_KARAT;
            objResponseDataToAvalonData.METAL_COLOR = objResponseDataToAvalonData.METAL_COLOR == null ? string.Empty : objResponseDataToAvalonData.METAL_COLOR;

            objResponseDataToAvalonData.PRODUCT_STYLE_ID = objResponseDataToAvalonData.PRODUCT_STYLE_ID.Replace(",", "`!^");
            objResponseDataToAvalonData.PRODUCT_NAME = objResponseDataToAvalonData.PRODUCT_NAME.Replace(",", "`!^");
            objResponseDataToAvalonData.CATEGORY_INFO = objResponseDataToAvalonData.CATEGORY_INFO.Replace(",", "`!^");
            objResponseDataToAvalonData.DESIGNER_COLLECTION = objResponseDataToAvalonData.DESIGNER_COLLECTION.Replace(",", "`!^");
            objResponseDataToAvalonData.IMAGE_NAME_1 = objResponseDataToAvalonData.IMAGE_NAME_1.Replace(",", "`!^");
            objResponseDataToAvalonData.IMAGE_NAME_2 = objResponseDataToAvalonData.IMAGE_NAME_2.Replace(",", "`!^");
            objResponseDataToAvalonData.IMAGE_NAME_3 = objResponseDataToAvalonData.IMAGE_NAME_3.Replace(",", "`!^");
            objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION = objResponseDataToAvalonData.WEB_SHORT_DESCRIPTION.Replace("'", "&#39").Replace(",", "`!^").Replace("\n", "").Replace("\n", "");
            objResponseDataToAvalonData.CATEGORY_INFO = objResponseDataToAvalonData.CATEGORY_INFO.Replace("'", "").Replace(",", "`!^");
            objResponseDataToAvalonData.COST_PRICE = objResponseDataToAvalonData.COST_PRICE.Replace("'", "").Replace(",", "`!^");
            objResponseDataToAvalonData.SELL_PRICE = objResponseDataToAvalonData.SELL_PRICE.Replace("'", "").Replace(",", "`!^");
            objResponseDataToAvalonData.METAL_KARAT = objResponseDataToAvalonData.METAL_KARAT.Replace("'", "").Replace(",", "`!^");
            objResponseDataToAvalonData.METAL_COLOR = objResponseDataToAvalonData.METAL_COLOR.Replace("'", "").Replace(",", "`!^");
            return objResponseDataToAvalonData;
        }
        #endregion
    }
}
