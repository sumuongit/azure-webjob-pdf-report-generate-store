using System;
using System.IO;
using System.Linq;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Configuration;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Fonts;
using PdfSharp.Drawing;
using TheArtOfDev.HtmlRenderer.PdfSharp;
using Zen.Barcode;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;

namespace ProductReportGenerator
{
    class Program
    {
        static readonly bool DATA_TEST = true;
        static readonly bool DATA_WRITE_TO_FILE = false;
        static readonly bool DATA_READ_FROM_FILE = false;
        static readonly bool SAVE_TO_AZURE = false;
        static readonly string SAVE_TO_PC_FOLDER_PATH = "C:/Users/Thinkpad/source/";
        static void Main(string[] args)
        {
            if (DATA_TEST)
                //args = new string[] { "eb53e5fe-3573-4276-80f5-4a1da8dcaae5", "2020-01-08", "2020-02-29" };
                args = new string[] {  };

            //Console.WriteLine(JsonConvert.SerializeObject(args));

            if (args.Length == 0)
            {
                var uid = "";
                var fromDate = "";
                var toDate = "";

                for (int i = 0; i < args.Length; i++)
                {
                    if (i == 0)
                        uid = args[i];
                    if (i == 1)
                        fromDate = args[i];
                    if (i == 2)
                        toDate = args[i];
                }

                if (uid != "" || fromDate != "" || toDate != "")
                {
                    //Console.WriteLine("ERROR: No args found.");
                }
                else
                {
                    CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

                    JObject jsonObj__Products;

                    if (!DATA_READ_FROM_FILE)
                        jsonObj__Products = GetProducts();
                    else
                        jsonObj__Products = JObject.Parse(File.ReadAllText(SAVE_TO_PC_FOLDER_PATH + "jsonObj__Products_res.json"));

                    //Console.WriteLine("START Configurate HTML");

                    string CSS = @"
                        <style>
                            div.body { font-size:32px; counter-reset:page; }
                            table { border-collapse:collapse; border:1px solid #ccc; width:100%; page-break-inside:avoid; }
		                    th,td {border:1px solid #ccc; vertical-align:top; padding:2px 5px; }
                            th { font-style:italic; text-align:center; font-weight:normal; }
                            p { margin:0; font-size:48px; font-weight:bold; margin-top:50px; margin-bottom:10px; }
	                    </style>";

                    string HTML = CSS + @"
	                    <div class='body'>
		                    <div style='text-align:center; font-weight:bold'>			                    
			                    <p style='font-size:54px; margin:0'>Product Report___" + fromDate + "___" + toDate + @"</p>
		                    </div> <br> <br>";

                    string HTML_table1 = "";

                    if (jsonObj__Products["GetProducts"].ToObject<JObject>().Count > 0)
                    {
                        foreach (JToken jToken__Products in jsonObj__Products["GetProducts"].Children())
                        {
                            foreach (JToken grandChild__Product in jToken__Products)
                            {
                                HTML_table1 += @"
                                    <tr>
                                        <td style='text-align:right'>" + grandChild__Product["Product_Code"] + @"</td>
                                        <td style='text-align:left'>" + grandChild__Product["Product_Name"] + @"</td>
                                        <td style='text-align:right'>" + grandChild__Product["Product_Qty"] + @"</td>
                                    </tr>";
                            }
                        }
                    }

                    HTML += @"
                        <div style='text-align:center; border-bottom:5px solid; padding-bottom:25px; margin-bottom:25px'>                           
                            <table style='width:60%; margin:0 auto; margin-left:340px'>
                                <colgroup>
                                    <col span='1' style='width:25%'>
                                    <col span='1' style='width:auto'>
                                    <col span='1' style='width:25%'>
                                </colgroup>
                                <tr>
                                    <th>Product Code</th>
                                    <th>Product Name</th>
                                    <th>Product Qty.</th>
                                </tr>"
                                 + HTML_table1 + @"
                            </table>
                        </div>";

                    //Console.WriteLine("END Configurate HTML");

                    //Console.WriteLine("START RenderHtmlAsPdf");

                    EZFontResolver fontResolver = EZFontResolver.Get;
                    GlobalFontSettings.FontResolver = fontResolver;
                    fontResolver.AddFont("Ubuntu", XFontStyle.Regular, @"fonts\ubuntu\ubuntu-R.ttf");
                    fontResolver.AddFont("Ubuntu", XFontStyle.Italic, @"fonts\ubuntu\ubuntu-RI.ttf");
                    fontResolver.AddFont("Ubuntu", XFontStyle.Bold, @"fonts\ubuntu\ubuntu-B.ttf");
                    fontResolver.AddFont("Ubuntu", XFontStyle.BoldItalic, @"fonts\ubuntu\ubuntu-BI.ttf");
                    XFont font = new XFont("Ubuntu", 18, XFontStyle.Regular);

                    var configPdf = new PdfGenerateConfig
                    {
                        ManualPageSize = new XSize(1240, 1750),
                        PageOrientation = PageOrientation.Landscape,
                        MarginTop = 25,
                        MarginRight = 25,
                        MarginBottom = 50,
                        MarginLeft = 25
                    };

                    PdfDocument pdf = PdfGenerator.GeneratePdf(HTML, configPdf);

                    int countPages = pdf.PageCount;
                    int numberPage = 0;
                    foreach (PdfPage page in pdf.Pages)
                    {
                        XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
                        gfx.DrawString("Page" + (++numberPage) + " of " + countPages, font, XBrushes.Black, 1620, 1210, XStringFormats.Default);
                    }

                    //Console.WriteLine("END RenderHtmlAsPdf");

                    if (SAVE_TO_AZURE)
                    {
                        Stream stream = new MemoryStream();
                        pdf.Save(stream);

                        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(ConfigurationManager.ConnectionStrings["StorageConnectionString"].ConnectionString);
                        CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
                        CloudBlobContainer container = blobClient.GetContainerReference("erp");
                        container.CreateIfNotExists();

                        BlobContainerPermissions permissions = container.GetPermissions();
                        permissions.PublicAccess = BlobContainerPublicAccessType.Container;
                        container.SetPermissions(permissions);

                        //Console.WriteLine("START UploadFromStream");

                        CloudBlockBlob report = container.GetBlockBlobReference("Reports/Product report/Product_report_" + fromDate + "_" + toDate + ".pdf");
                        report.UploadFromStream(stream);

                        //Console.WriteLine("END UploadFromStream");
                    }
                    else
                        pdf.Save(SAVE_TO_PC_FOLDER_PATH + "Product_report_" + fromDate + "_" + toDate + ".pdf");
                }
            }
            else
            {
                //Console.WriteLine("ERROR: No args found.");
            }
        }

       static JObject GetProducts()
        {
            //Console.WriteLine("START GetProducts");

            SqlConnection conn;
            using (conn = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlConnectionString"].ConnectionString))
            {
                conn.Open();

                JObject jsonObj__GetProducts = new JObject();
                string json = "{'GetProducts' : {";

                SqlCommand sqlCmd__GetProducts = new SqlCommand("SELECT * FROM Product", conn);

                //SqlParameter sqlParam__GetProducts_uuid = new SqlParameter("@uuid", uuid);
                //SqlParameter sqlParam__GetProducts_fromDate = new SqlParameter("@fromDate", fromDate);
                //SqlParameter sqlParam__GetProducts_toDate = new SqlParameter("@toDate", toDate);

                //sqlCmd__GetProducts.Parameters.Add(sqlParam__GetProducts_uuid);
                //sqlCmd__GetProducts.Parameters.Add(sqlParam__GetProducts_fromDate);
                //sqlCmd__GetProducts.Parameters.Add(sqlParam__GetProducts_toDate);

                SqlDataReader reader__GetProducts = sqlCmd__GetProducts.ExecuteReader();

                if (reader__GetProducts.HasRows)
                {
                    if (!DATA_READ_FROM_FILE)
                    {
                        DataTable dt__GetProducts = new DataTable();
                        dt__GetProducts.Load(reader__GetProducts);

                        JArray jsonAr__GetProducts_tmp = JArray.Parse(JsonConvert.SerializeObject(dt__GetProducts, Formatting.Indented));

                        int key = 0;

                        foreach(JObject jsonObj__GetProduct in jsonAr__GetProducts_tmp)
                        {
                            json += @"
                                    '" + jsonAr__GetProducts_tmp[key]["Product_Id"].ToObject<int>() + @"' : {
                                        'Product_Code' : '" + jsonAr__GetProducts_tmp[key]["Product_Code"] + @"',
                                        'Product_Name' : '" + jsonAr__GetProducts_tmp[key]["Product_Name"] + @"',
                                        'Product_Qty' : '" + jsonAr__GetProducts_tmp[key]["Product_Qty"] + @"'
                                       },";

                            key++;
                        }
                       
                        json += "} }";

                        jsonObj__GetProducts = JObject.Parse(json);

                        if (DATA_WRITE_TO_FILE)
                            File.WriteAllText(SAVE_TO_PC_FOLDER_PATH + "jsonObj__GetProducts.json", JsonConvert.SerializeObject(jsonObj__GetProducts));
                    }
                    else
                        jsonObj__GetProducts = JObject.Parse(File.ReadAllText(SAVE_TO_PC_FOLDER_PATH + "jsonObj__GetProducts.json"));

                    return jsonObj__GetProducts;
                }
                else
                {
                    //Console.WriteLine("ERROR: No rows found in sqlCmd__GetProduct.");

                    conn.Close();
                    //Console.WriteLine("END GetProducts");

                    return jsonObj__GetProducts;
                }
            }
        }
    }     
 }
