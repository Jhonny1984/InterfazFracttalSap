using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Security.Cryptography;
using RestSharp;
using Newtonsoft.Json;
using System.Data.SqlClient;
using System.Data;
using System.ComponentModel;
using HawkNet;
using System.Web.UI;
using System.Web;
using System.Security.Permissions;
using MSScriptControl;
using System.Threading;
using System.Net;
using System.Diagnostics;
using System.IO;
using System.Web;
using RestSharp.Extensions;

namespace ProInterfaceFracttal
{
    public static class Program
    {
        static HawkCredential _credential;

        private const string IdFracttal = "TLFmgX1Kuef4rsaNxk9z";//// Id entregado por fracttal
        private const string key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU"; ////Key entregada por Fracttal

        private static readonly Encoding encoding = Encoding.UTF8;

        public static string nonce;
        public static string jsonRequis;
        public static string jsonOrdene;
        public static string itemsmov;
        public static DataTable tabla = new DataTable();

        /*to sql server conexions string use te global variables string */
        public static string IpServerSql = "192.168.0.4";
        public static string ServerSqlUser = "sa";
        public static string ServerSqlPass = "Ceo2015*";
        public static string ServerSqlPort = "1433";
        public static string ServerSqlDBTran = "TRANSFLESA91";
        public static string ServerSqlDBMaci = "INVERFFACSA91";
        public static string ServerSqlDBSant = "SBO_SANTAINES_2";
        public static string ServerSqlDBINTER = "DB_INTERFACE";






        static void Main(string[] args)
        {
            //EnviaCorreo.EnviaCorreos("E001","123456","1");
            // PutSiteMacizo();
            //PutiteMacizo();
            //PutiteSantaInes();
          //  PutiteTransflesa();

            //PutCargaFracttalInventories();

            //TestBatches.ConectandoDB();
            ////TestBatches.BatchUpdate();
            ////L////
            //POSTItemsNuevos();
            //Mantenimiento industrial
            //POSTItemsNuevos2();


            //PutDatosFracttal();

            // PutDatosFracttal2();

            ////funcional 04072018 no tocar
            //RequisFracttal();  


            //PutCargaFracttal();

            //GetCargaFracttalMonitores();

             //  GetCargaFracttal();
            /* LOGISTICA */
            AddOrderToDatabase();
            //AddComplementsToDatabase();

          //  DraftToCocument();

        ///    AddPurchaseOrder();
         
            ///*MANTENIMIENTO
          // AddOrderToDatabase2();
            //AddComplementsToDatabase2();
            //AddPurchaseOrder2();
            ///////// SANTA INES 
            //AddOrderToDatabase3();
            //AddComplementsToDatabase3();
            //AddPurchaseOrder3();




        }


        //public static SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*");

        public static void POSTItemsNuevos() {


            DataTable tabla = new DataTable();

            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBTran + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))


            //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                var ArticulosMovimiento = "select top 200 * from v_ArticulosNuevos order by 4 desc";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(tabla);
            }

            string data = string.Empty;
            StringBuilder sb = new StringBuilder();

            if (null != tabla && null != tabla.Rows)
            {
                foreach (DataRow dataRow in tabla.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        sb.Append(item);
                        sb.Append(',');
                    }
                    sb.AppendLine();
                }

                data = sb.ToString();
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("REGISTROS PARA CREAR NUEVOS ITEMS");
            Console.WriteLine("-----------LOGISTICA-------------");
            Console.ForegroundColor = ConsoleColor.White;
            //Console.WriteLine(data);

            //Int32 ts = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds; ///Obtiene el timestamp en unix
            //string macfinal = getmac(ts);
            //macfinal = macfinal.Replace("+", "-").Replace("/", "_");///Base64 nos devuelve unos caracteres que pueden dañar la llamada

            foreach (DataRow dataRow in tabla.Rows)
            {

                //foreach (var item in dataRow.ItemArray)
                //{
                //array[i] = tabla.Rows[i][0].ToString();
                //try
                //{
                var proc3 = new System.Diagnostics.Process();
                proc3.StartInfo.CreateNoWindow = true;
                proc3.StartInfo.RedirectStandardInput = true;
                proc3.StartInfo.RedirectStandardOutput = true;
                proc3.StartInfo.UseShellExecute = false;
                proc3.StartInfo.RedirectStandardError = true;
                proc3.StartInfo.FileName = "node.exe";
                proc3.StartInfo.Arguments = "-i";
                proc3.Start();
                proc3.BeginOutputReadLine();

                //proc.StandardInput.WriteLine("2 + 2;");
                proc3.StandardInput.WriteLine("var request = require(\"request\");" +
                    "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
                    "var options = { method: 'POST'," +
                    "url: 'https://app.fracttal.com/api/items/', " +
                    "headers:" +
                        "{ } ," +
                        "hawk: {" +
                                   "credentials: {" +
                                     "id: id," +
                                     "key: key," +
                                     "algorithm: 'sha256'" +
                                   "}" +
                                "}," +
                    "formData:" +
                      "{ id_type_item: '4'," +
                        "code: '" + dataRow[0] + "'," +
                        "field_1: '" + dataRow[1] + "'," +
                        "field_2: '" + dataRow[0] + "'," +
                        "unit_code: '12'," +
                        "field_6: '" + dataRow[2] + "'," +
                        "description: '' } };" +
                    "request(options, function(error, response, body) {" +
                    "if (error) throw new Error(error);" +
                    "console.log(body);" +
                    "});");

                proc3.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 1000).suppressOut;");
                proc3.OutputDataReceived += proc_OutputDataReceivedPut;

                proc3.WaitForExit();
                //}


                using (var progress = new ProgressBar())
                {
                    for (int h = 0; h <= tabla.Rows.Count; h++)
                    {
                        progress.Report((double)h / tabla.Rows.Count);
                        System.Threading.Thread.Sleep(50);
                    }
                }


            }

        }
        /*-----------------------------------------------------------------*/


        public static void POSTItemsNuevos2()
        {


            DataTable tabla = new DataTable();

            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBMaci + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))


            //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                var ArticulosMovimiento = "select top 200 * from v_ArticulosNuevos order by 4 desc";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(tabla);
            }

            string data = string.Empty;
            StringBuilder sb = new StringBuilder();

            if (null != tabla && null != tabla.Rows)
            {
                foreach (DataRow dataRow in tabla.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        sb.Append(item);
                        sb.Append(',');
                    }
                    sb.AppendLine();
                }

                data = sb.ToString();
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("REGISTROS PARA CREAR NUEVOS ITEMS");
            Console.WriteLine("----MANTENIMIENTO INDUSTRIAL----");
            Console.ForegroundColor = ConsoleColor.White;
            //Console.WriteLine(data);

            //Int32 ts = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds; ///Obtiene el timestamp en unix
            //string macfinal = getmac(ts);
            //macfinal = macfinal.Replace("+", "-").Replace("/", "_");///Base64 nos devuelve unos caracteres que pueden dañar la llamada

            foreach (DataRow dataRow in tabla.Rows)
            {

                //foreach (var item in dataRow.ItemArray)
                //{
                //array[i] = tabla.Rows[i][0].ToString();
                //try
                //{
                var proc3 = new System.Diagnostics.Process();
                proc3.StartInfo.CreateNoWindow = true;
                proc3.StartInfo.RedirectStandardInput = true;
                proc3.StartInfo.RedirectStandardOutput = true;
                proc3.StartInfo.UseShellExecute = false;
                proc3.StartInfo.RedirectStandardError = true;
                proc3.StartInfo.FileName = "node.exe";
                proc3.StartInfo.Arguments = "-i";
                proc3.Start();
                proc3.BeginOutputReadLine();

                //proc.StandardInput.WriteLine("2 + 2;");
                proc3.StandardInput.WriteLine("var request = require(\"request\");" +
                    "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
                    "var options = { method: 'POST'," +
                    "url: 'https://app.fracttal.com/api/items/', " +
                    "headers:" +
                        "{ } ," +
                        "hawk: {" +
                                   "credentials: {" +
                                     "id: id," +
                                     "key: key," +
                                     "algorithm: 'sha256'" +
                                   "}" +
                                "}," +
                    "formData:" +
                      "{ id_type_item: '4'," +
                        "code: '" + dataRow[0] + "'," +
                        "field_1: '" + dataRow[1] + "'," +
                        "field_2: '" + dataRow[0] + "'," +
                        "unit_code: '12'," +
                        "field_6: '" + dataRow[2] + "'," +
                        "description: '' } };" +
                    "request(options, function(error, response, body) {" +
                    "if (error) throw new Error(error);" +
                    "console.log(body);" +
                    "});");

                proc3.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 1000).suppressOut;");
                proc3.OutputDataReceived += proc_OutputDataReceivedPut;

                proc3.WaitForExit();
                //}


                using (var progress = new ProgressBar())
                {
                    for (int h = 0; h <= tabla.Rows.Count; h++)
                    {
                        progress.Report((double)h / tabla.Rows.Count);
                        System.Threading.Thread.Sleep(50);
                    }
                }


            }

        }





        public static void PutDatosFracttal() {


            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBTran + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //  using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                var ArticulosMovimiento = "select * from V_ArticulosMovimiento order by 1";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(tabla);
            }

            string data = string.Empty;
            StringBuilder sb = new StringBuilder();

            if (null != tabla && null != tabla.Rows)
            {
                foreach (DataRow dataRow in tabla.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        sb.Append(item);
                        sb.Append(',');
                    }
                    sb.AppendLine();
                }

                data = sb.ToString();
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("REGISTROS PARA ACTUALIZAR ENCONTRADOS");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(data);
            //Console.ReadKey();

            //Int32 ts = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds; ///Obtiene el timestamp en unix
            //string macfinal = getmac(ts);
            //macfinal = macfinal.Replace("+", "-").Replace("/", "_");///Base64 nos devuelve unos caracteres que pueden dañar la llamada
            // foreach (DataRow dataRow in tabla.Rows)

            string[] array = new string[tabla.Rows.Count];
            //for (int i = 0; i < tabla.Rows.Count; i++)


            foreach (DataRow dataRow in tabla.Rows)
            {

                //foreach (var item in dataRow.ItemArray)
                //{
                //array[i] = tabla.Rows[i][0].ToString();
                //try
                //{
                var proc2 = new System.Diagnostics.Process();
                proc2.StartInfo.CreateNoWindow = true;
                proc2.StartInfo.RedirectStandardInput = true;
                proc2.StartInfo.RedirectStandardOutput = true;
                proc2.StartInfo.UseShellExecute = false;
                proc2.StartInfo.RedirectStandardError = true;
                proc2.StartInfo.FileName = "node.exe";
                proc2.StartInfo.Arguments = "-i";
                proc2.Start();
                proc2.BeginOutputReadLine();

                //proc.StandardInput.WriteLine("2 + 2;");
                proc2.StandardInput.WriteLine("var request = require(\"request\");" +
                    "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
                    "var options = { method: 'PUT'," +
                    "url: 'https://app.fracttal.com/api/inventories/" + dataRow[0] + "', " +
                    "headers:" +
                        "{ } ," +
                        "hawk: {" +
                                   "credentials: {" +
                                     "id: id," +
                                     "key: key," +
                                     "algorithm: 'sha256'" +
                                   "}" +
                                "}," +
                    "formData:" +
                      "{ field_1: '" + dataRow[1] + "'," +
                        "id_warehouse: '11'," +
                        "unit_cost_stock: '" + dataRow[3] + "'," +
                        "stock: '" + dataRow[2] + "'," +
                        "min_stock_level: '" + dataRow[4] + "'," +
                        "max_stock_level: '" + dataRow[5] + "' } };" +
                    "request(options, function(error, response, body) {" +
                    "if (error) throw new Error(error);" +
                    "console.log(body);" +
                    "});");

                proc2.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 1000).suppressOut;");
                proc2.OutputDataReceived += proc_OutputDataReceivedPut;

                proc2.WaitForExit();
                //}


                //catch (WebException e)
                //{
                //    Console.WriteLine("This program is expected to throw WebException on successful run." +
                //    "\n\nException Message :" + e.Message);
                //    if (e.Status == WebExceptionStatus.ProtocolError)
                //    {
                //        Console.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                //        Console.WriteLine("Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                //    }
                //}
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.Message);
                //}
            }
            //} 

        }




        public static void PutDatosFracttal2()
        {


            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBMaci + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //  using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                var ArticulosMovimiento = "select * from V_ArticulosMovimiento order by 1";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(tabla);
            }

            string data = string.Empty;
            StringBuilder sb = new StringBuilder();

            if (null != tabla && null != tabla.Rows)
            {
                foreach (DataRow dataRow in tabla.Rows)
                {
                    foreach (var item in dataRow.ItemArray)
                    {
                        sb.Append(item);
                        sb.Append(',');
                    }
                    sb.AppendLine();
                }

                data = sb.ToString();
            }
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("REGISTROS PARA ACTUALIZAR ENCONTRADOS");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(data);
            //Console.ReadKey();

            //Int32 ts = (Int32)(DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1))).TotalSeconds; ///Obtiene el timestamp en unix
            //string macfinal = getmac(ts);
            //macfinal = macfinal.Replace("+", "-").Replace("/", "_");///Base64 nos devuelve unos caracteres que pueden dañar la llamada
            // foreach (DataRow dataRow in tabla.Rows)

            string[] array = new string[tabla.Rows.Count];
            //for (int i = 0; i < tabla.Rows.Count; i++)


            foreach (DataRow dataRow in tabla.Rows)
            {

                //foreach (var item in dataRow.ItemArray)
                //{
                //array[i] = tabla.Rows[i][0].ToString();
                //try
                //{
                var proc2 = new System.Diagnostics.Process();
                proc2.StartInfo.CreateNoWindow = true;
                proc2.StartInfo.RedirectStandardInput = true;
                proc2.StartInfo.RedirectStandardOutput = true;
                proc2.StartInfo.UseShellExecute = false;
                proc2.StartInfo.RedirectStandardError = true;
                proc2.StartInfo.FileName = "node.exe";
                proc2.StartInfo.Arguments = "-i";
                proc2.Start();
                proc2.BeginOutputReadLine();

                //proc.StandardInput.WriteLine("2 + 2;");
                proc2.StandardInput.WriteLine("var request = require(\"request\");" +
                    "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
                    "var options = { method: 'PUT'," +
                    "url: 'https://app.fracttal.com/api/inventories/" + dataRow[0] + "', " +
                    "headers:" +
                        "{ } ," +
                        "hawk: {" +
                                   "credentials: {" +
                                     "id: id," +
                                     "key: key," +
                                     "algorithm: 'sha256'" +
                                   "}" +
                                "}," +
                    "formData:" +
                      "{ field_1: '" + dataRow[1] + "'," +
                        "id_warehouse: '17'," +
                        "unit_cost_stock: '" + dataRow[3] + "'," +
                        "stock: '" + dataRow[2] + "'," +
                        "min_stock_level: '" + dataRow[4] + "'," +
                        "max_stock_level: '" + dataRow[5] + "' } };" +
                    "request(options, function(error, response, body) {" +
                    "if (error) throw new Error(error);" +
                    "console.log(body);" +
                    "});");

                proc2.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 1000).suppressOut;");
                proc2.OutputDataReceived += proc_OutputDataReceivedPut;

                proc2.WaitForExit();
                //}


                //catch (WebException e)
                //{
                //    Console.WriteLine("This program is expected to throw WebException on successful run." +
                //    "\n\nException Message :" + e.Message);
                //    if (e.Status == WebExceptionStatus.ProtocolError)
                //    {
                //        Console.WriteLine("Status Code : {0}", ((HttpWebResponse)e.Response).StatusCode);
                //        Console.WriteLine("Status Description : {0}", ((HttpWebResponse)e.Response).StatusDescription);
                //    }
                //}
                //catch (Exception e)
                //{
                //    Console.WriteLine(e.Message);
                //}
            }
            //} 

        }

        public static void proc_OutputDataReceivedPut(object sender, System.Diagnostics.DataReceivedEventArgs e)
        {
            //foreach (DataRow dataRow in tabla.Rows)
            //{ 
            //for (int i = 0; i < tabla.Rows.Count; i++)
            //{    //Console.ReadKey();
            string json = @"" + e.Data + "";
            itemsmov = json.Trim(new Char[] { ' ', '>' });
            //Console.WriteLine(jsonRequis.Length);
            //Console.WriteLine(jsonRequis);
            //if (itemsmov.Length > 3654)


            itemsmov = itemsmov.Substring(0, itemsmov.Length);

            Console.ForegroundColor = System.ConsoleColor.Green;
            Console.WriteLine(itemsmov);



            //using (var progress = new ProgressBar())
            //{
            //    for (int h = 0; h <= tabla.Rows.Count; h++)
            //    {
            //        progress.Report((double)h / 100);
            //        System.Threading.Thread.Sleep(tabla.Rows.Count);
            //    }
            //}



            //OrdenesFracttal();
            //}


            //Deserealizacion del jason de requisiciones 
            //Requis Requi = JsonConvert.DeserializeObject<Requis>(jsonFinal);

        }




        //public static void RequisFracttal()
        //{
        //    ////Setiamos el dia para cargar ordenes
        //    DateTime hoy = DateTime.Today;
        //    string dia = hoy.ToString("dd");
        //    string diac = Convert.ToString((Convert.ToInt32(dia) + 1)); ;
        //    string diai = Convert.ToString((Convert.ToInt32(diac) - 1));
        //    //Cargamos el mes actual
        //    string mes = hoy.ToString("MM");
        //    string año = hoy.ToString("yyyy");


        //    try
        //    {

        //        var proc = new System.Diagnostics.Process();
        //        proc.StartInfo.CreateNoWindow = true;
        //        proc.StartInfo.RedirectStandardInput = true;
        //        proc.StartInfo.RedirectStandardOutput = true;
        //        proc.StartInfo.UseShellExecute = false;
        //        proc.StartInfo.RedirectStandardError = true;
        //        proc.StartInfo.FileName = "node.exe";
        //        proc.StartInfo.Arguments = "-i";
        //        proc.Start();
        //        proc.BeginOutputReadLine();

        //        //proc.StandardInput.WriteLine("2 + 2;");
        //        proc.StandardInput.WriteLine("var request = require(\"request\");" +
        //            "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
        //            "var options = { method: 'GET'," +
        //            "url: 'https://app.fracttal.com/api/work_orders_movements/', " +//+ "" + "'2018-" + mes + "-" + diai + "T00:00:00-00'&until='2018-" + mes +"-" + diac + "T00:00:00-00'" +
        //            "qs: { since: '20180801' ," +
        //                  "until: '20180802'} ," +
        //            "headers:" +

        //            "{ } ," +
        //            "hawk:" + "{" +
        //                        "credentials:" + "{" +
        //                        "id: id," +
        //                        "key: key," +
        //                        "algorithm: 'sha256'" +
        //                        "}" +
        //                      "}" +
        //                   "};" +
        //            "request(options, function(error, response, body) {" +
        //            "if (error) throw new Error(error);" +
        //            "console.log(body);" +
        //            "});");

        //        //proc.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 10000).suppressOut;");
        //        proc.OutputDataReceived += proc_OutputDataReceived;

        //        proc.WaitForExit();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Execeptio " + ex);
        //        Console.ReadKey();
        //    }
        //}




        //public static void proc_OutputDataReceived(object sender, System.Diagnostics.DataReceivedEventArgs e)
        //{

        //    //Console.ReadKey();
        //    string json = @"" + e.Data + "";
        //    jsonRequis = json.Trim(new Char[] { ' ', '>' });
        //    //Console.WriteLine(jsonRequis.Length);
        //    // Console.WriteLine(jsonRequis);
        //    if (jsonRequis.Length > 3654)
        //    {

        //        jsonRequis = jsonRequis.Substring(0, jsonRequis.Length);

        //        Console.WriteLine(jsonRequis);
        //        OrdenesFracttal();
        //    }


        //    //Deserealizacion del jason de requisiciones 
        //    //Requis Requi = JsonConvert.DeserializeObject<Requis>(jsonFinal);

        //}


        //public static void OrdenesFracttal()
        //{

        //    ////Setiamos el dia para cargar ordenes
        //    DateTime hoy = DateTime.Today;
        //    string dia = hoy.ToString("dd");
        //    string diac = Convert.ToString((Convert.ToInt32(dia) + 0)); ;
        //    string diai = Convert.ToString((Convert.ToInt32(diac) - 1));
        //    //Cargamos el mes actual
        //    string mes = hoy.ToString("MM");
        //    string año = hoy.ToString("yyyy");


        //    try
        //    {

        //        var proc = new System.Diagnostics.Process();
        //        proc.StartInfo.CreateNoWindow = true;
        //        proc.StartInfo.RedirectStandardInput = true;
        //        proc.StartInfo.RedirectStandardOutput = true;
        //        proc.StartInfo.UseShellExecute = false;
        //        proc.StartInfo.RedirectStandardError = true;
        //        proc.StartInfo.FileName = "node.exe";
        //        proc.StartInfo.Arguments = "-i";
        //        proc.Start();
        //        proc.BeginOutputReadLine();

        //        //proc.StandardInput.WriteLine("2 + 2;");
        //        proc.StandardInput.WriteLine("var request = require(\"request\");" +
        //            "id = 'TLFmgX1Kuef4rsaNxk9z',key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
        //            "var options = { method: 'GET'," +                              //'2016-07-12T20:00:00-03'
        //            "url: 'https://app.fracttal.com/api/work_orders/'," + //+ "" + "'2018-" + mes +"-" + diai + "T00:00:00-00'&until='2018-" + mes + "-"  + diac + "T00:00:00-00'" +
        //            "qs:  { since: '" + año + mes + dia + "' ," +
        //            "       until: '" + año + mes + diac + "'} ," +
        //            "headers:" +
        //            "{ } ," +
        //            "hawk:" + "{" +
        //                        "credentials:" + "{" +
        //                        "id: id," +
        //                        "key: key," +
        //                        "algorithm: 'sha256'" +
        //                        "}" +
        //                      "}" +
        //                   "};" +
        //            "request(options, function(error, response, body) {" +
        //            "if (error) throw new Error(error);" +
        //            "console.log(body);" +
        //            "});");

        //        //proc.StandardInput.WriteLine("setTimeout(function(){ process.exit();}, 10000).suppressOut;");
        //        proc.OutputDataReceived += proc_OutputDataReceivedW;
        //        proc.WaitForExit();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Execeptio " + ex);
        //    }
        //}




        //public static void proc_OutputDataReceivedW(object sender, System.Diagnostics.DataReceivedEventArgs e)
        //{

        //    //Console.ReadKey();
        //    string json = @"" + e.Data + "";
        //    jsonOrdene = json.Trim(new Char[] { ' ', '>' });
        //    // Console.WriteLine(jsonOrdene.Length);
        //    if (jsonOrdene.Length > 3645)
        //    {

        //        jsonOrdene = jsonOrdene.Substring(0, jsonOrdene.Length);

        //        Console.WriteLine(jsonOrdene);
        //        GetCargaFracttal();
        //    }


        //    //Deserealizacion del jason de requisiciones 
        //    //Requis Requi = JsonConvert.DeserializeObject<Requis>(jsonFinal);

        //}

        //MSScriptControl.ScriptControl js = new MSScriptControl.ScriptControl();

        //js.AllowUI = false;
        //js.Language = "JScript";
        //js.Reset();

        //js.AddCode(@"var request = require('\u0022'request'\u0022');" +
        //    "id = 'TLFmgX1Kuef4rsaNxk9z'," +
        //    "key = '0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU';" +
        //    "var options = { method: 'GET'," +
        //    "url: 'https://app.fracttal.com/api/work_orders/'," +
        //    "headers:" +
        //    "{ } ," +
        //    "hawk:" + "{" +
        //                "credentials:" + "{" +
        //                    "id: id," +
        //                    "key: key," +
        //                    "algorithm: 'sha256'" +
        //             "  }" +
        //            "}" +
        //   "};" +
        //    "request(options, function (error, response, body) {" +
        //    "if (error) throw new Error(error);" +
        //    //"console.log(body);" +
        //    "});");
        ////object[] parms = new object[] { 11 };
        //string result = (string)js.Run("test"/*/, ref parms*/);
        //Console.WriteLine(result);
        //Console.ReadKey();


        static public void Authenticate(IRestClient client, IRestRequest request)
        {
            var uri = client.BuildUri(request);
            var portSuffix = uri.Port != 80 ? ":" + uri.Port : "";
            var host = uri.Host + portSuffix;
            var method = request.Method.ToString();

            var header = Hawk.GetAuthorizationHeader(host, method, uri, _credential);

            request.AddHeader("Authorization", "Hawk " + header);
        }









        public static DataTable tabla2 = new DataTable();



        public static void GetCargaFracttal() {


            using (
               var progress = new ProgressBar())
            {
                for (int h = 0; h <= /*TableC.Rows.Count*/100; h++)
                {
                    progress.Report((double)h / 100);
                    System.Threading.Thread.Sleep(50);
                }
            }


            ////Setiamos el dia para cargar ordenes
            ///
            ///CAMBIAR DIA
            DateTime hoy = DateTime.Today;
            DateTime diaini = (hoy.AddDays(0));
            DateTime diafin = (hoy.AddDays(1));


            DateTime DiaOrd = diaini.AddDays(0);

            string diaorden = DiaOrd.ToString("dd");

            string añoini = diaini.ToString("yyyy");
            string mesini = diaini.ToString("MM");
            string diai = diaini.ToString("dd");

            string añofin = diafin.ToString("yyyy");
            string mesfin = diafin.ToString("MM");
            string diaf = diafin.ToString("dd");
            string OrdenTra = "";



            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/work_orders_movements/?since=" + añoini + "-" + mesini + "-" + diai + "T00:00:00-00&until=" + añofin + "-" + mesfin + "-" + diaf + "T00:00:00-00");
            //    var client = new RestClient("https://app.fracttal.com/api/work_orders_movements/032627");

            var request = new RestRequest(Method.GET);
            Authenticate(client, request);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            //Console.WriteLine(jsonResponse);
            //Console.ReadKey();
            
            ///
            //var cliento= new RestClient("https://app.fracttal.com/api/work_orders/?since=" + añoini + "-" + mesini + "-" + diaorden + "T00:00:00-00&until=" + añofin + "-" + mesfin + "-" + diaf + "T00:00:00-00");
            //para cargar orden de forma directa
            ////var cliento = new RestClient("https://app.fracttal.com/api/work_orders/");

            //var requesto = new RestRequest(Method.GET);
            //Authenticate(cliento, requesto);
            //IRestResponse responseo = cliento.Execute(requesto);
            //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
            //Console.WriteLine(jsonResponseo);





            try
            {

                //cambio en recibir datos de fracttal 01/08/2018

                ////Deserealizacion del jason de requisiciones 
                Requis Requi = JsonConvert.DeserializeObject<Requis>(response.Content);
                //inicializa desearealzacion de ordenes de trabajo con procedimiento restsharp <>
                //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);







                /*Tabla para los requerimientos de materiales*/
                DataTable tabla = new DataTable();
                tabla.Columns.Add("id");
                tabla.Columns.Add("warehouses_source_description");
                tabla.Columns.Add("document");
                tabla.Columns.Add("date");
                tabla.Columns.Add("code");
                tabla.Columns.Add("folio_source");
                tabla.Columns.Add("document1");
                tabla.Columns.Add("Itemcode");
                tabla.Columns.Add("items_description");
                tabla.Columns.Add("qty");
                tabla.Columns.Add("qty_pending");
                tabla.Columns.Add("unit_cost_company");
                tabla.Columns.Add("movements_states_description");

                /*Tabla para las ordenes de trabajo*/
                tabla2.Columns.Add("id_work_order");
                tabla2.Columns.Add("id_work_orders_tasks");
                tabla2.Columns.Add("id_status_work_order");
                tabla2.Columns.Add("wo_folio");
                tabla2.Columns.Add("creation_date");
                tabla2.Columns.Add("duration");
                tabla2.Columns.Add("initial_date");
                tabla2.Columns.Add("id_priorities");
                tabla2.Columns.Add("final_date");
                tabla2.Columns.Add("completed_percentage");
                tabla2.Columns.Add("created_by");
                tabla2.Columns.Add("personnel_description");
                tabla2.Columns.Add("code1");
                tabla2.Columns.Add("items_log_description");
                tabla2.Columns.Add("note");
                tabla2.Columns.Add("tasks_log_task_type_main");
                tabla2.Columns.Add("parent_description");
                tabla2.Columns.Add("user_assigned");
                tabla2.Columns.Add("description");
                tabla2.Columns.Add("detection_method_description");
                tabla2.Columns.Add("tasks_log_types_2_description");
                tabla2.Columns.Add("costs_center_description");
                tabla2.Columns.Add("Department1");
                tabla2.Columns.Add("id_request");
                tabla2.Columns.Add("groups_2_description");
                



                Console.ForegroundColor = ConsoleColor.Blue;
                Console.WriteLine("[-----------------Requisiciones------------------]");
                //Console.WriteLine("Requisiciones Encontradas");
                Console.ForegroundColor = ConsoleColor.Yellow;
                if (Requi.Data.Count.Equals(0))
                {

                    Console.WriteLine("[LA ORDEN NO TIENE REQUESICIONES DE MATERIALES]");


                }
                else {

                    Console.WriteLine(Requi.Data.Count);

                    if (Requi.Data.Count > 0)
                    {


                        for (int i = 0; i < Requi.Data.Count; i++)
                        {
                             OrdenTra = Requi.Data[i].document.ToString();
                            Console.WriteLine("\n");
                            Console.WriteLine("************************************************************************************************************");
                            Console.ForegroundColor = ConsoleColor.Cyan;
                            Console.WriteLine("\n");
                            Console.WriteLine(Requi.Data[i].id.ToString() + "" +
                                ",\t" + Requi.Data[i].warehouses_source_description.ToString() + "" +
                                ",\t" + Requi.Data[i].document.ToString() + "" +
                                ",\t" + Requi.Data[i].date + "" +
                                ",\t" + Requi.Data[i].code + "" +
                                ",\t" + Requi.Data[i].folio_source + "" +
                                ",\t" + Requi.Data[i].movements_states_description);
                            /*CARGA ORDENS DE FRACTTAL */
                            Cargarordenesfracttal(OrdenTra);

                            Console.ForegroundColor = ConsoleColor.Green;
                            for (int n = 0; n < Requi.Data[i].list_items.Count; n++)
                            {
                                Console.WriteLine(Requi.Data[i].document + ",\t" + Requi.Data[i].list_items[n].code + ",\t" + Requi.Data[i].list_items[n].items_description + ",\t" + Requi.Data[i].list_items[n].qty + ",\t" + Requi.Data[i].list_items[n].qty_pending + ",\t" + Requi.Data[i].list_items[n].unit_cost_company);

                                DataRow fila = tabla.NewRow();

                                fila["id"] = Requi.Data[i].id.ToString();
                                fila["warehouses_source_description"] = Requi.Data[i].warehouses_source_description.ToString();
                                fila["document"] = Requi.Data[i].document.ToString();
                                fila["date"] = Requi.Data[i].date;
                                fila["code"] = Requi.Data[i].code;
                                fila["folio_source"] = Requi.Data[i].folio_source;
                                fila["document1"] = Requi.Data[i].document;
                                fila["ItemCode"] = Requi.Data[i].list_items[n].code;
                                fila["items_description"] = Requi.Data[i].list_items[n].items_description;
                                fila["qty"] = Requi.Data[i].list_items[n].qty;
                                fila["qty_pending"] = Requi.Data[i].list_items[n].qty_pending;
                                fila["unit_cost_company"] = Requi.Data[i].list_items[n].unit_cost_company;
                                fila["movements_states_description"] = Requi.Data[i].movements_states_description;

                                tabla.Rows.Add(fila);
                            }


                            //var cliento = new RestClient("https://app.fracttal.com/api/work_orders/" + OrdenTra);

                            //var requesto = new RestRequest(Method.GET);
                            //Authenticate(cliento, requesto);
                            //IRestResponse responseo = cliento.Execute(requesto);
                            //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
                            //Console.WriteLine(jsonResponseo);

                            //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);
                            //Console.WriteLine("Lineas Encontradas " + workOrdenes.Data.Count);


                        }

                    }
                    else {

                        Console.WriteLine("[LA ORDEN NO CONTIENE REQUISICIONES DE MATERIALES ]");

                    }
                }
                DataTable dtResult = new DataTable();



                dtResult.Columns.Add("id");
                dtResult.Columns.Add("warehouses_source_description");
                dtResult.Columns.Add("document");
                dtResult.Columns.Add("date");
                dtResult.Columns.Add("code");
                dtResult.Columns.Add("folio_source");
                dtResult.Columns.Add("document1");
                dtResult.Columns.Add("ItemCode");
                dtResult.Columns.Add("items_description");
                dtResult.Columns.Add("qty");
                dtResult.Columns.Add("qty_pending");
                dtResult.Columns.Add("unit_cost_company");
                dtResult.Columns.Add("id_work_order");
                dtResult.Columns.Add("id_work_orders_tasks");
                dtResult.Columns.Add("id_status_work_order");
                dtResult.Columns.Add("wo_folio");
                dtResult.Columns.Add("creation_date");
                dtResult.Columns.Add("duration");
                dtResult.Columns.Add("initial_date");
                dtResult.Columns.Add("id_priorities");
                dtResult.Columns.Add("final_date");
                dtResult.Columns.Add("completed_percentage");
                dtResult.Columns.Add("created_by");
                dtResult.Columns.Add("personnel_description");
                dtResult.Columns.Add("code1");
                dtResult.Columns.Add("items_log_description");
                dtResult.Columns.Add("note");
                dtResult.Columns.Add("tasks_log_task_type_main");
                dtResult.Columns.Add("movements_states_description");
                dtResult.Columns.Add("parent_description");
                dtResult.Columns.Add("tasks_log_types_2_description");
                dtResult.Columns.Add("costs_center_description");
                dtResult.Columns.Add("id_request");
                dtResult.Columns.Add("groups_2_description");

                /*Se unen las dos tablas */
                var dataRows = from dataRows1 in tabla.AsEnumerable().Distinct().ToList()
                               join dataRows2 in tabla2.AsEnumerable().Distinct().ToList()
                               on dataRows1.Field<string>("document").ToString() equals dataRows2.Field<string>("wo_folio").ToString()
                               // where dataRows2.Field<string>("completed_percentage") != "100" 

                               orderby dataRows1.Field<string>("document") descending
                               select dtResult.LoadDataRow(new object[]
                                {
                                    dataRows1.Field<string>("id"),
                                    dataRows1.Field<string>("warehouses_source_description"),
                                    dataRows1.Field<string>("document"),
                                    dataRows1.Field<string>("date"),
                                    dataRows1.Field<string>("code"),
                                    dataRows1.Field<string>("folio_source"),
                                    dataRows1.Field<string>("document1"),
                                    dataRows1.Field<string>("ItemCode"),
                                    dataRows1.Field<string>("items_description"),
                                    dataRows1.Field<string>("qty"),
                                    dataRows1.Field<string>("qty_pending"),
                                    dataRows1.Field<string>("unit_cost_company"),
                                    dataRows2.Field<string>("id_work_order"),
                                    dataRows2.Field<string>("id_work_orders_tasks"),
                                    dataRows2.Field<string>("id_status_work_order"),
                                    dataRows2.Field<string>("wo_folio"),
                                    Convert.ToDateTime( dataRows2.Field<string>("creation_date")).ToString(format: "dd/MM/yyyy"),
                                    dataRows2.Field<string>("duration"),
                                    Convert.ToDateTime( dataRows2.Field<string>("initial_date")).ToString(format: "dd/MM/yyyy"),
                                    dataRows2.Field<string>("id_priorities"),
                                    Convert.ToDateTime(dataRows2.Field<string>("final_date")).ToString(format: "dd/MM/yyyy"),
                                    dataRows2.Field<string>("completed_percentage"),
                                    dataRows2.Field<string>("created_by"),
                                    dataRows2.Field<string>("personnel_description"),
                                    dataRows2.Field<string>("code1"),
                                    dataRows2.Field<string>("items_log_description"),
                                    dataRows2.Field<string>("note"),
                                    dataRows2.Field<string>("tasks_log_task_type_main"),
                                    dataRows1.Field<string>("movements_states_description"),
                                    dataRows2.Field<string>("parent_description"),
                                    dataRows2.Field<string>("tasks_log_types_2_description"),
                                    dataRows2.Field<string>("costs_center_description"),
                                    dataRows2.Field<string>("id_request"),
                                    dataRows2.Field<string>("groups_2_description"),


                                }, false);




                DataTable dataTableLinqJoined = dataRows.CopyToDataTable();

                DataView vista = new DataView(dataTableLinqJoined);
                DataTable dstabla = vista.ToTable(true
                    , "warehouses_source_description"
                    , "document"
                    , "folio_source"
                    , "ItemCode"
                    , "items_description"
                    , "qty"
                    , "qty_pending"
                    , "unit_cost_company"
                    , "id_status_work_order"
                    , "creation_date"
                    , "initial_date"
                    , "final_date"
                    , "completed_percentage"
                    , "created_by"
                    , "personnel_description"
                    , "code1"
                    , "items_log_description"
                    , "note"
                    , "parent_description" 
                    , "tasks_log_types_2_description"
                    , "costs_center_description"
                    , "id_request" 
                    , "groups_2_description");



                //DataTable distinctTable = dataTableLinqJoined.AsEnumerable().Distinct() 
                //            .OrderBy(r => r.Field<string>("document"))
                //            .GroupBy(r =>  r.Field<string>("folio_source"))                           
                //            .Select(g => g.First())
                //            .CopyToDataTable();



                //var json = new JavaScriptSerializer().Serialize(distinctTable);


                DataTable datadistinct = new DataTable();

                datadistinct.Columns.Add("folio_source");
                datadistinct.Columns.Add("date");
                datadistinct.Columns.Add("document1");
                //datadistinct.Columns.Add("ItemsLists");
                datadistinct.Columns.Add("ItemCode");
                datadistinct.Columns.Add("items_description");
                datadistinct.Columns.Add("qty");
                datadistinct.Columns.Add("qty_pending");
                datadistinct.Columns.Add("unit_cost_company");
                datadistinct.Columns.Add("completed_percentage");
                datadistinct.Columns.Add("code1");
                datadistinct.Columns.Add("personnel_description");
                datadistinct.Columns.Add("note");
                datadistinct.Columns.Add("id_work_orders_tasks");
                datadistinct.Columns.Add("tasks_log_task_type_main");
                datadistinct.Columns.Add("movements_states_description");
                datadistinct.Columns.Add("parent_description");
                datadistinct.Columns.Add("tasks_log_types_2_description");
                datadistinct.Columns.Add("costs_center_description");
                datadistinct.Columns.Add("id_request");
                datadistinct.Columns.Add("groups_2_description");


                var distinctRows = (from /*DataRow*/ dRow in dataTableLinqJoined.AsEnumerable().Distinct().ToList()
                                    orderby dRow.Field<string>("document1") descending

                                    select datadistinct.LoadDataRow(new object[]
                                    { dRow.Field<string>("document1"), //0
                                      //dRow.Field<string>("ItemsLists"),
                                      Convert.ToDateTime(dRow.Field<string>("date")).ToString(format: "dd/MM/yyyy"), //1
                                      dRow.Field<string>("folio_source"), //2
                                        dRow.Field<string>("ItemCode"),//3
                                        dRow.Field<string>("items_description"),//4
                                        dRow.Field<string>("qty"),//5
                                        dRow.Field<string>("qty_pending"),//6
                                        dRow.Field<string>("unit_cost_company"),//7
                                        dRow.Field<string>("completed_percentage"),//8
                                        dRow.Field<string>("code1"),//9
                                        dRow.Field<string>("personnel_description"),//10
                                        dRow.Field<string>("note"),//11
                                        dRow.Field<string>("id_work_orders_tasks"),//12
                                        dRow.Field<string>("tasks_log_task_type_main"),//13
                                        dRow.Field<string>("movements_states_description"),//14
                                        dRow.Field<string>("parent_description"),//15
                                        dRow.Field<string>("tasks_log_types_2_description"),
                                        dRow.Field<string>("costs_center_description"), //16
                                        dRow.Field<string>("id_request"),  //17
                                        dRow.Field<string>("groups_2_description")  //18
                                    }, false)).Distinct().ToList()  ;


                 DataTable datadistintos = distinctRows.AsEnumerable().Distinct().ToList().CopyToDataTable();
                //Console.WriteLine(datadistintos.Select("parent_description like '%MACIZO%'"));


                DataTable tbcomparaenca = new DataTable();

                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
                {
                    //var ArticulosMovimiento = " select distinct * from (" +
                    //    " select top 1000  t0.U_folio_source,t0.U_document1 as Document1   from " +
                    //    " [@LINEASOFRACTTAL] t0 " +
                    //    " (nolock) " +
                    //    " inner join " +
                    //    " [@ORDENESFRACTTAL] t1 " +
                    //    " (nolock) " +
                    //    " on t0.code = t1.code " +
                    //    " where U_parent_description like '%// MACIZO MANTENIMIENTO/%' " +
                    //    " order by  convert(int, t0.U_folio_source) desc  )rr " +
                    //    " order by rr.Document1 desc "; 

                    var ArticulosMovimiento = "SELECT distinct [Code],[U_document1]  as  Document1  FROM [DB_INTERFACE].[dbo].[@LINEASOFRACTTAL] order by code desc";

                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(tbcomparaenca);
                }


                ///revisando ordenes ya registradas
                var idsNotInB = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("Document1"))).Except(tbcomparaenca.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("Document1"))));
                if ((idsNotInB.Count()) > 0)
                {
                    DataTable TableC = (from row in datadistintos.AsEnumerable()
                                        where row["Document1"] != System.DBNull.Value
                                        join id in idsNotInB
                                        on row.Field<string>("Document1") equals id
                                        select row).CopyToDataTable();

                   
                 
                    //Console.Write("Verficando Ordenes Registradas... ");


                    /**/

                    //DataTable TableResultado = new DataTable();

                    //TableResultado = tbcomparaenca.AsEnumerable()

                    //    .Where(r =>

                    //        distinctRows.AsEnumerable().Any(w =>

                    //            w.Field<string>("Document1") == r.Field<string>("Document1")))

                    //        .CopyToDataTable<DataRow>();

                    /**/


                    /**/

                    // Fill the DataSet.
                    //DataSet ds = new DataSet();
                    //ds.Locale = CultureInfo.InvariantCulture;
                    //FillDataSet(ds);

                    //DataTable contactTable = ds.Tables["Contact"];

                    // Create two tables.
                    //IEnumerable<DataRow> query1 = from Document1 in tbcomparaenca.AsEnumerable()
                    //                              where Document1.Field<string>("Document1") == "Ms."
                    //                              select Document1;

                    //IEnumerable<DataRow> query2 = from Document1 in distinctRows.AsEnumerable()
                    //                              where Document1.Field<string>("Document1") == "Sandra"
                    //                              select Document1;

                    //DataTable contacts1 = query1.CopyToDataTable();
                    //DataTable contacts2 = query2.CopyToDataTable();

                    //// Find the intersection of the two tables.
                    //var contacts = contacts1.AsEnumerable().Intersect(contacts2.AsEnumerable(),
                    //                                                    DataRowComparer.Default);







                    /**/






                    //using (
                    //    var progress = new ProgressBar())
                    //{
                    //    for (int h = 0; h <= TableC.Rows.Count; h++)
                    //    {
                    //        progress.Report((double)h / 100);
                    //        System.Threading.Thread.Sleep(20);
                    //    }
                    //}
                   
                    //string sortOrder;
                   // sortOrder = "parent_description desc";

                    DataTable dtDistinct = new DataTable();

                    string[] sColumnas = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description", "parent_description", /*"tasks_log_types_2_description",*/ "costs_center_description", "id_request"/*, "groups_2_description"*/ };
                    dtDistinct = TableC.DefaultView.ToTable(true, sColumnas);


                    //Console.BackgroundColor = ConsoleColor.Magenta;
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    //Console.WriteLine("BUSCANDO OPERACIONES DE MACIZO");
                    Console.WriteLine("BUSCANDO OPERACIONES DE LOGISTICA");


                    DataTable dtdfiltrado = new DataTable();
                    DataTable dtOrdenesfracttal = new DataTable();

                    if ((dtDistinct.Rows.Count) > 0)
                    {

                        DataRow[] foundDT = dtDistinct.Select("parent_description like '%MACIZO%'");


                        /*PARA FILTRAR OPERACIONES DEPENDIENDO LA UNIDAD QUE LAS GENERA*/
                        //DataRow[] foundDT = dtDistinct.Select("parent_description like '%LOGISTICA%'");



                        if (foundDT.Count() > 0)
                        {

                            dtdfiltrado = foundDT.CopyToDataTable();

                            dtOrdenesfracttal = dtdfiltrado.AsEnumerable().GroupBy(x => x.Field<string>("Document1")).Select(g => g.First()).CopyToDataTable(); ;


                            foreach (DataRow row in foundDT)
                            {
                                Console.WriteLine("{0}, {1}", row[0], row[15]);
                            }

                        }
                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("NO SE ENCOTRARON DATOS PARA FILTRAR !!");
                        

                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.DarkGreen;
                        Console.WriteLine("NO HAY OT NUEVAS PARA AGREGAR !!");
                    }

                    // dtDistinct.Select(expression);

                    //gvDatos.DataSource = TableC.DefaultView;
                    //gvDatos.ForeColor = System.Drawing.Color.Red;

                    /*Insertadon en @ORDENESFRACTTAL*/
                    if (dtOrdenesfracttal.Rows.Count > 0)
                    {
                        InserOrdenesFracttal(dtOrdenesfracttal);
                        //Guarda las ordenes 
                        //GuardarOrden();
                    }

                    ///*Insertando en @LINEASFRACTTAL*/
                    if (dtdfiltrado.Rows.Count > 0)
                    {
                        InserOrdenesLineasFracttal(dtdfiltrado);
                        //Guarda las ordenes 
                        //GuardarOrden();
                    }




                    //}

                    ///Agrega los Complementos o requisiciones que se crearon despues de la carga de las ordenes 
                    //DataTable TbtComparaRequis = new DataTable();


                    //using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                    ////using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
                    //{
                    //    var Complementos = "select distinct * from (" +
                    //    " select top 1000  U_folio_source,U_document1 as Document1   from " +
                    //    " [@LINEASOFRACTTAL](nolock) order by  convert(int, U_folio_source) desc )rr";
                    //    conexion.Open();
                    //    SqlDataAdapter adaptador = new SqlDataAdapter(Complementos, conexion);
                    //    adaptador.Fill(TbtComparaRequis);
                    //}
                    //var idsNotInBR = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("document1"))).Except(TbtComparaRequis.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("U_document1"))));
                    //if (idsNotInBR.Count() > 0)
                    //{
                    //    DataTable TableCR = (from row in datadistintos.AsEnumerable()
                    //                         where row["document1"] != System.DBNull.Value
                    //                         join id in idsNotInBR
                    //                         on row.Field<string>("document1") equals id
                    //                         select row).CopyToDataTable();

                    //    DataTable dtDistinctCR = new DataTable();
                    //    string[] sColumnasCR = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description" };
                    //    dtDistinctCR = TableCR.DefaultView.ToTable(true, sColumnas);

                    //    //Console.ForegroundColor = ConsoleColor.DarkGreen;
                    //    //Console.BackgroundColor = ConsoleColor.Gray;
                    //    Console.Write("Verficando Complementos de las Ordenes Registradas... ");
                    //    using (var progress = new ProgressBar())
                    //    {
                    //        for (int h = 0; h <= TableCR.Rows.Count; h++)
                    //        {
                    //            progress.Report((double)h / 100);
                    //            System.Threading.Thread.Sleep(20);
                    //        }
                    //    }
                    //    Console.WriteLine("Done.");

                    //    //gvDatos.DataSource = TableC.DefaultView;
                    //    //gvDatos.ForeColor = System.Drawing.Color.Red;
                    //    if (TableCR.Rows.Count > 0)
                    //    {
                    //        InserOrdenes(dtDistinctCR);
                    //    }
                    //}

                    //Console.WriteLine("No se Encontraron Ordenes Nuevas");
                    // para carga en sap
                    //ConectandoDB();
                    //AddOrderToDatabase();
                    // DraftToCocument();

                }

               

            }

  


            catch (Exception ex)
            {

              
                Console.WriteLine("Error al conectar" + ex);

            }

            using (
                var progress = new ProgressBar())
            {
                for (int h = 0; h <= /*TableC.Rows.Count*/100; h++)
                {
                    progress.Report((double)h / 100);
                    System.Threading.Thread.Sleep(50);
                }
            }


            var info = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
            System.Diagnostics.Process.Start(info);


            // LoopGetfracttal();

            Environment.Exit(0);


        }


        public static void LoopGetfracttal() {



           
            GetCargaFracttal();
            Environment.Exit(0);

        }
              



    

        /*Lista de Medidores*/

        public static void GetCargaFracttalMonitores()
        {

            ////Setiamos el dia para cargar ordenes
            ///
            ///CAMBIAR DIA
            DateTime hoy = DateTime.Today;
            DateTime diaini = (hoy.AddDays(-5));
            DateTime diafin = (hoy.AddDays(1));


            DateTime DiaOrd = diaini.AddDays(0);

            string diaorden = DiaOrd.ToString("dd");

            string añoini = diaini.ToString("yyyy");
            string mesini = diaini.ToString("MM");
            string diai = diaini.ToString("dd");

            string añofin = diafin.ToString("yyyy");
            string mesfin = diafin.ToString("MM");
            string diaf = diafin.ToString("dd");



            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/meters_list/");

            var request = new RestRequest(Method.GET);
            Authenticate(client, request);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            //Console.WriteLine(jsonResponse);
            //Console.ReadKey();

            ///
            //var cliento= new RestClient("https://app.fracttal.com/api/work_orders/?since=" + añoini + "-" + mesini + "-" + diaorden + "T00:00:00-00&until=" + añofin + "-" + mesfin + "-" + diaf + "T00:00:00-00");
            //para cargar orden de forma directa
            ////var cliento = new RestClient("https://app.fracttal.com/api/work_orders/");

            //var requesto = new RestRequest(Method.GET);
            //Authenticate(cliento, requesto);
            //IRestResponse responseo = cliento.Execute(requesto);
            //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
            //Console.WriteLine(jsonResponseo);


            DataTable dtResult = new DataTable();
            dtResult.Columns.Add("Code");


            try
            {

                //cambio en recibir datos de fracttal 01/08/2018

                ////Deserealizacion del jason de requisiciones 
                Requis Requi = JsonConvert.DeserializeObject<Requis>(response.Content);
                //inicializa desearealzacion de ordenes de trabajo con procedimiento restsharp <>
                //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);

                //Crear tabla para ingresar la data deserializada 
                DataTable tabla = new DataTable();
                tabla.Columns.Add("description");
                tabla.Columns.Add("serial");
                tabla.Columns.Add("is_counter");
                tabla.Columns.Add("counter_offset_value");
                tabla.Columns.Add("date");
                tabla.Columns.Add("value");
                tabla.Columns.Add("accumulated_value");
                tabla.Columns.Add("monthly_average_data");
                tabla.Columns.Add("units_description");
                tabla.Columns.Add("units_code");
                tabla.Columns.Add("items_description");
                tabla.Columns.Add("parent_description");
                tabla.Columns.Add("monthly_average_auto");
                tabla.Columns.Add("min_value");
                tabla.Columns.Add("max_value");
                tabla.Columns.Add("code");
                //tabla.Columns.Add("id_company");

                //tabla2.Columns.Add("id_work_order");
                //tabla2.Columns.Add("id_work_orders_tasks");
                //tabla2.Columns.Add("id_status_work_order");
                //tabla2.Columns.Add("wo_folio");
                //tabla2.Columns.Add("creation_date");
                //tabla2.Columns.Add("duration");
                //tabla2.Columns.Add("initial_date");
                //tabla2.Columns.Add("id_priorities");
                //tabla2.Columns.Add("final_date");
                //tabla2.Columns.Add("completed_percentage");
                //tabla2.Columns.Add("created_by");
                //tabla2.Columns.Add("personnel_description");
                //tabla2.Columns.Add("code1");
                //tabla2.Columns.Add("items_log_description");
                //tabla2.Columns.Add("note");
                //tabla2.Columns.Add("tasks_log_task_type_main");
                //tabla2.Columns.Add("parent_description");
                //tabla2.Columns.Add("user_assigned");
                //tabla2.Columns.Add("description");
                //tabla2.Columns.Add("detection_method_description");



                //Console.ForegroundColor = ConsoleColor.Blue;
                //Console.WriteLine("[-----------------Requisiciones------------------]");
                ////Console.WriteLine("Requisiciones Encontradas");
                //Console.ForegroundColor = ConsoleColor.Yellow;
                //Console.WriteLine(Requi.Data.Count);
                for (int i = 0; i < Requi.Data.Count; i++)
                {
                    //    string OrdenTra = Requi.Data[i].document.ToString();
                    //    Console.WriteLine("\n");
                    //    Console.WriteLine("*****************************************************************************************************************************");
                    //    Console.ForegroundColor = ConsoleColor.Cyan;
                    //    Console.WriteLine("\n");
                    //    Console.WriteLine(Requi.Data[i].id.ToString() + "" +
                    //        ",\t" + Requi.Data[i].warehouses_source_description.ToString() + "" +
                    //        ",\t" + Requi.Data[i].document.ToString() + "" +
                    //        ",\t" + Requi.Data[i].date + "" +
                    //        ",\t" + Requi.Data[i].code + "" +
                    //        ",\t" + Requi.Data[i].folio_source + "" +
                    //        ",\t" + Requi.Data[i].movements_states_description);

                    //    Cargarordenesfracttal(OrdenTra);

                    Console.ForegroundColor = ConsoleColor.Green;
                    for (int n = 0; n < Requi.Data.Count; n++)
                    {
                        //Console.WriteLine(Requi.Data[i].document + ",\t" + Requi.Data[i].list_items[n].code + ",\t" + Requi.Data[i].list_items[n].items_description + ",\t" + Requi.Data[i].list_items[n].qty + ",\t" + Requi.Data[i].list_items[n].qty_pending + ",\t" + Requi.Data[i].list_items[n].unit_cost_company);

                        DataRow fila = tabla.NewRow();

                        fila["Code"] = Requi.Data[i].code;
                        //fila["warehouses_source_description"] = Requi.Data[i].code;
                        //fila["document"] = Requi.Data[i].docum;
                        //fila["date"] = Requi.Data[i].date;
                        //fila["code"] = Requi.Data[i].code;
                        //fila["folio_source"] = Requi.Data[i].folio_source;
                        //fila["document1"] = Requi.Data[i].document;
                        //fila["ItemCode"] = Requi.Data[i].list_items[n].code;
                        //fila["items_description"] = Requi.Data[i].list_items[n].items_description;
                        //fila["qty"] = Requi.Data[i].list_items[n].qty;
                        //fila["qty_pending"] = Requi.Data[i].list_items[n].qty_pending;
                        //fila["unit_cost_company"] = Requi.Data[i].list_items[n].unit_cost_company;
                        //fila["movements_states_description"] = Requi.Data[i].movements_states_description;


                        tabla.Rows.Add(fila);


                        var dataRows = from dataRows1 in tabla.AsEnumerable().Distinct()
                                           //                   join dataRows2 in tabla2.AsEnumerable().Distinct()
                                           //                   on dataRows1.Field<string>("document").ToString() equals dataRows2.Field<string>("wo_folio").ToString()
                                           //                   // where dataRows2.Field<string>("completed_percentage") != "100" 

                                       orderby dataRows1.Field<string>("Code") descending
                                       select dtResult.LoadDataRow(new object[]
                                     {
             //dataRows1.Field<string>("id"),
                                                                   dataRows1.Field<string>("Code")//,
                                                                                                  //dataRows1.Field<string>("document"),
                                                                                                  //dataRows1.Field<string>("date"),
                                                                                                  //dataRows1.Field<string>("code"),
                                                                                                  //dataRows1.Field<string>("folio_source"),
                                                                                                  //dataRows1.Field<string>("document1"),
                                                                                                  //dataRows1.Field<string>("ItemCode"),
                                                                                                  //dataRows1.Field<string>("items_description"),
                                                                                                  //dataRows1.Field<string>("qty"),
                                                                                                  //dataRows1.Field<string>("qty_pending"),
                                                                                                  //dataRows1.Field<string>("unit_cost_company"),
                                                                                                  //dataRows2.Field<string>("id_work_order"),
                                                                                                  //dataRows2.Field<string>("id_work_orders_tasks"),
                                                                                                  //dataRows2.Field<string>("id_status_work_order"),
                                                                                                  //dataRows2.Field<string>("wo_folio"),
                                                                                                  //Convert.ToDateTime(dataRows2.Field<string>("creation_date")).ToString(format: "dd/MM/yyyy"),
                                                                                                  //dataRows2.Field<string>("duration"),
                                                                                                  //Convert.ToDateTime(dataRows2.Field<string>("initial_date")).ToString(format: "dd/MM/yyyy"),
                                                                                                  //dataRows2.Field<string>("id_priorities"),
                                                                                                  //Convert.ToDateTime(dataRows2.Field<string>("final_date")).ToString(format: "dd/MM/yyyy"),
                                                                                                  //dataRows2.Field<string>("completed_percentage"),
                                                                                                  //dataRows2.Field<string>("created_by"),
                                                                                                  //dataRows2.Field<string>("personnel_description"),
                                                                                                  //dataRows2.Field<string>("code1"),
                                                                                                  //dataRows2.Field<string>("items_log_description"),
                                                                                                  //dataRows2.Field<string>("note"),
                                                                                                  //dataRows2.Field<string>("tasks_log_task_type_main"),
                                                                                                  //dataRows1.Field<string>("movements_states_description"),
                                                                                                  //dataRows2.Field<string>("parent_description"),
                                     }, false);




                        //DataTable distinctTable = dataTableLinqJoined.AsEnumerable()
                        //                    .GroupBy(r => r.Field<string>("document1"))
                        //                    .Select(g => g.First())
                        //                    .CopyToDataTable();


                        //    Console.WriteLine(dataRows[i].code );

                    }


                    //var cliento = new RestClient("https://app.fracttal.com/api/work_orders/" + OrdenTra);

                    //var requesto = new RestRequest(Method.GET);
                    //Authenticate(cliento, requesto);
                    //IRestResponse responseo = cliento.Execute(requesto);
                    //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
                    //Console.WriteLine(jsonResponseo);

                    //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);
                    //Console.WriteLine("Lineas Encontradas " + workOrdenes.Data.Count);


                }

            }

            //    DataTable dtResult = new DataTable();



            //    dtResult.Columns.Add("id");
            //    dtResult.Columns.Add("warehouses_source_description");
            //    dtResult.Columns.Add("document");
            //    dtResult.Columns.Add("date");
            //    dtResult.Columns.Add("code");
            //    dtResult.Columns.Add("folio_source");
            //    dtResult.Columns.Add("document1");
            //    dtResult.Columns.Add("ItemCode");
            //    dtResult.Columns.Add("items_description");
            //    dtResult.Columns.Add("qty");
            //    dtResult.Columns.Add("qty_pending");
            //    dtResult.Columns.Add("unit_cost_company");
            //    dtResult.Columns.Add("id_work_order");
            //    dtResult.Columns.Add("id_work_orders_tasks");
            //    dtResult.Columns.Add("id_status_work_order");
            //    dtResult.Columns.Add("wo_folio");
            //    dtResult.Columns.Add("creation_date");
            //    dtResult.Columns.Add("duration");
            //    dtResult.Columns.Add("initial_date");
            //    dtResult.Columns.Add("id_priorities");
            //    dtResult.Columns.Add("final_date");
            //    dtResult.Columns.Add("completed_percentage");
            //    dtResult.Columns.Add("created_by");
            //    dtResult.Columns.Add("personnel_description");
            //    dtResult.Columns.Add("code1");
            //    dtResult.Columns.Add("items_log_description");
            //    dtResult.Columns.Add("note");
            //    dtResult.Columns.Add("tasks_log_task_type_main");
            //    dtResult.Columns.Add("movements_states_description");
            //    dtResult.Columns.Add("parent_description");


            //    var dataRows = from dataRows1 in tabla.AsEnumerable().Distinct()
            //                   join dataRows2 in tabla2.AsEnumerable().Distinct()
            //                   on dataRows1.Field<string>("document").ToString() equals dataRows2.Field<string>("wo_folio").ToString()
            //                   // where dataRows2.Field<string>("completed_percentage") != "100" 

            //                   orderby dataRows1.Field<string>("document") descending
            //                   select dtResult.LoadDataRow(new object[]
            //                    {
            //                        dataRows1.Field<string>("id"),
            //                        dataRows1.Field<string>("warehouses_source_description"),
            //                        dataRows1.Field<string>("document"),
            //                        dataRows1.Field<string>("date"),
            //                        dataRows1.Field<string>("code"),
            //                        dataRows1.Field<string>("folio_source"),
            //                        dataRows1.Field<string>("document1"),
            //                        dataRows1.Field<string>("ItemCode"),
            //                        dataRows1.Field<string>("items_description"),
            //                        dataRows1.Field<string>("qty"),
            //                        dataRows1.Field<string>("qty_pending"),
            //                        dataRows1.Field<string>("unit_cost_company"),
            //                        dataRows2.Field<string>("id_work_order"),
            //                        dataRows2.Field<string>("id_work_orders_tasks"),
            //                        dataRows2.Field<string>("id_status_work_order"),
            //                        dataRows2.Field<string>("wo_folio"),
            //                        Convert.ToDateTime( dataRows2.Field<string>("creation_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("duration"),
            //                        Convert.ToDateTime( dataRows2.Field<string>("initial_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("id_priorities"),
            //                        Convert.ToDateTime(dataRows2.Field<string>("final_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("completed_percentage"),
            //                        dataRows2.Field<string>("created_by"),
            //                        dataRows2.Field<string>("personnel_description"),
            //                        dataRows2.Field<string>("code1"),
            //                        dataRows2.Field<string>("items_log_description"),
            //                        dataRows2.Field<string>("note"),
            //                        dataRows2.Field<string>("tasks_log_task_type_main"),
            //                        dataRows1.Field<string>("movements_states_description"),
            //                        dataRows2.Field<string>("parent_description"),
            //                    }, false);




            //    DataTable dataTableLinqJoined = dataRows.CopyToDataTable();

            //    DataTable distinctTable = dataTableLinqJoined.AsEnumerable()
            //                .GroupBy(r => r.Field<string>("document1"))
            //                .Select(g => g.First())
            //                .CopyToDataTable();



            //    //var json = new JavaScriptSerializer().Serialize(distinctTable);


            //    DataTable datadistinct = new DataTable();

            //    datadistinct.Columns.Add("folio_source");
            //    datadistinct.Columns.Add("date");
            //    datadistinct.Columns.Add("document1");
            //    //datadistinct.Columns.Add("ItemsLists");
            //    datadistinct.Columns.Add("ItemCode");
            //    datadistinct.Columns.Add("items_description");
            //    datadistinct.Columns.Add("qty");
            //    datadistinct.Columns.Add("qty_pending");
            //    datadistinct.Columns.Add("unit_cost_company");
            //    datadistinct.Columns.Add("completed_percentage");
            //    datadistinct.Columns.Add("code1");
            //    datadistinct.Columns.Add("personnel_description");
            //    datadistinct.Columns.Add("note");
            //    datadistinct.Columns.Add("id_work_orders_tasks");
            //    datadistinct.Columns.Add("tasks_log_task_type_main");
            //    datadistinct.Columns.Add("movements_states_description");
            //    datadistinct.Columns.Add("parent_description");


            //    var distinctRows = (from /*DataRow*/ dRow in distinctTable.AsEnumerable().Distinct()
            //                        orderby dRow.Field<string>("document1") descending

            //                        select datadistinct.LoadDataRow(new object[]
            //                        { dRow.Field<string>("document1"), //0
            //                          //dRow.Field<string>("ItemsLists"),
            //                          Convert.ToDateTime(dRow.Field<string>("date")).ToString(format: "dd/MM/yyyy"), //1
            //                          dRow.Field<string>("folio_source"), //2
            //                            dRow.Field<string>("ItemCode"),//3
            //                            dRow.Field<string>("items_description"),//4
            //                            dRow.Field<string>("qty"),//5
            //                            dRow.Field<string>("qty_pending"),//6
            //                            dRow.Field<string>("unit_cost_company"),//7
            //                            dRow.Field<string>("completed_percentage"),//8
            //                            dRow.Field<string>("code1"),//9
            //                            dRow.Field<string>("personnel_description"),//10
            //                            dRow.Field<string>("note"),//11
            //                            dRow.Field<string>("id_work_orders_tasks"),//12
            //                            dRow.Field<string>("tasks_log_task_type_main"),//13
            //                            dRow.Field<string>("movements_states_description"),//14
            //                            dRow.Field<string>("parent_description")//15
            //                        }, false)).Distinct();


            //    DataTable datadistintos = distinctRows.AsEnumerable().Distinct().CopyToDataTable();
            //    Console.WriteLine(datadistintos);
            //    DataTable tbcomparaenca = new DataTable();

            //    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //    // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            //    {
            //        var ArticulosMovimiento = " select distinct [Code],U_Document1 from [@LINEASOFRACTTAL] order by [Code] desc";
            //        conexion.Open();
            //        SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
            //        adaptador.Fill(tbcomparaenca);
            //    }


            //    ///revisando ordenes ya registradas
            //    var idsNotInB = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("Document1"))).Except(tbcomparaenca.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("U_Document1"))));
            //    if (idsNotInB.Count() > 0)
            //    {
            //        DataTable TableC = (from row in datadistintos.AsEnumerable()
            //                            where row["Document1"] != System.DBNull.Value
            //                            join id in idsNotInB
            //                            on row.Field<string>("Document1") equals id
            //                            select row).CopyToDataTable();

            //        //Console.ForegroundColor = ConsoleColor.DarkGreen;
            //        //Console.BackgroundColor = ConsoleColor.Gray;
            //        Console.Write("Verficando Ordenes Registradas... ");
            //        using (var progress = new ProgressBar())
            //        {
            //            for (int h = 0; h <= TableC.Rows.Count; h++)
            //            {
            //                progress.Report((double)h / 100);
            //                System.Threading.Thread.Sleep(20);
            //            }
            //        }

            //        DataTable dtDistinct = new DataTable();

            //        string[] sColumnas = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description", "parent_description" };
            //        dtDistinct = TableC.DefaultView.ToTable(true, sColumnas);

            //        Console.WriteLine("Done.");

            //        //gvDatos.DataSource = TableC.DefaultView;
            //        //gvDatos.ForeColor = System.Drawing.Color.Red;
            //        if (TableC.Rows.Count > 0)
            //        {
            //            InserOrdenes(dtDistinct);
            //            //Guarda las ordenes 
            //            //GuardarOrden();
            //        }
            //        //}

            //        ///Agrega los Complementos o requisiciones que se crearon despues de la carga de las ordenes 
            //        DataTable TbtComparaRequis = new DataTable();


            //        using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //        //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            //        {
            //            var Complementos = "select distinct [U_document1]  from [@LINEASOFRACTTAL] order by 1 desc";
            //            conexion.Open();
            //            SqlDataAdapter adaptador = new SqlDataAdapter(Complementos, conexion);
            //            adaptador.Fill(TbtComparaRequis);
            //        }
            //        var idsNotInBR = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("document1"))).Except(TbtComparaRequis.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("U_document1"))));
            //        if (idsNotInBR.Count() > 0)
            //        {
            //            DataTable TableCR = (from row in datadistintos.AsEnumerable()
            //                                 where row["document1"] != System.DBNull.Value
            //                                 join id in idsNotInBR
            //                                 on row.Field<string>("document1") equals id
            //                                 select row).CopyToDataTable();

            //            DataTable dtDistinctCR = new DataTable();
            //            string[] sColumnasCR = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description" };
            //            dtDistinctCR = TableCR.DefaultView.ToTable(true, sColumnas);

            //            //Console.ForegroundColor = ConsoleColor.DarkGreen;
            //            //Console.BackgroundColor = ConsoleColor.Gray;
            //            Console.Write("Verficando Complementos de las Ordenes Registradas... ");
            //            using (var progress = new ProgressBar())
            //            {
            //                for (int h = 0; h <= TableCR.Rows.Count; h++)
            //                {
            //                    progress.Report((double)h / 100);
            //                    System.Threading.Thread.Sleep(20);
            //                }
            //            }
            //            Console.WriteLine("Done.");

            //            //gvDatos.DataSource = TableC.DefaultView;
            //            //gvDatos.ForeColor = System.Drawing.Color.Red;
            //            if (TableCR.Rows.Count > 0)
            //            {
            //                InserOrdenes(dtDistinctCR);
            //            }
            //        }

            //        //Console.WriteLine("No se Encontraron Ordenes Nuevas");
            //        // para carga en sap
            //        //ConectandoDB();
            //        //AddOrderToDatabase();
            //        // DraftToCocument();

            //    }


            //}

            //}//pruebas

            catch (Exception ex)
            {

                Console.WriteLine("Error al conectar" + ex);

            }

            //Console.WriteLine("Press any key to exit.");
            //Console.ReadKey();
        }



        //---------------------------------PRUEBA DE PUT REST------------------------------------------------------


        public static void PutCargaFracttal()
        {

            ////Setiamos el dia para cargar ordenes
            ///
            ///CAMBIAR DIA
            //DateTime hoy = DateTime.Today;
            //DateTime diaini = (hoy.AddDays(-2));
            //DateTime diafin = (hoy.AddDays(1));


            //DateTime DiaOrd = diaini.AddDays(0);

            //string diaorden = DiaOrd.ToString("dd");

            //string añoini = diaini.ToString("yyyy");
            //string mesini = diaini.ToString("MM");
            //string diai = diaini.ToString("dd");

            //string añofin = diafin.ToString("yyyy");
            //string mesfin = diafin.ToString("MM");
            //string diaf = diafin.ToString("dd");



            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/meter_reading/IT001");

            var request = new RestRequest(Method.PUT);

            Authenticate(client, request);

            request.AddHeader("Content-Type", "application/json");
            request.AddParameter("undefined", "{\n    \"date\": \"2019-07-22T21:00:00-03\",\n    \"value\": 1,\n    \"serial\": \"100\"\n}", ParameterType.RequestBody);
            //            request.AddParameter("\"" + "Date\":" + "\"" + "2019-07-06T00:00:00-05\""+"," + "\"" + "value\":"+"2"+"," + "\"" + "serial\":"+"\""+"100\"" , ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            Console.WriteLine(jsonResponse);
            Console.ReadKey();







            //try
            //{

            //    //cambio en recibir datos de fracttal 01/08/2018

            //    ////Deserealizacion del jason de requisiciones 
            //    Requis Requi = JsonConvert.DeserializeObject<Requis>(response.Content);
            //    //inicializa desearealzacion de ordenes de trabajo con procedimiento restsharp <>
            //    //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);


            //    DataTable tabla = new DataTable();
            //    tabla.Columns.Add("id");
            //    tabla.Columns.Add("warehouses_source_description");
            //    tabla.Columns.Add("document");
            //    tabla.Columns.Add("date");
            //    tabla.Columns.Add("code");
            //    tabla.Columns.Add("folio_source");
            //    tabla.Columns.Add("document1");
            //    tabla.Columns.Add("Itemcode");
            //    tabla.Columns.Add("items_description");
            //    tabla.Columns.Add("qty");
            //    tabla.Columns.Add("qty_pending");
            //    tabla.Columns.Add("unit_cost_company");
            //    tabla.Columns.Add("movements_states_description");


            //    tabla2.Columns.Add("id_work_order");
            //    tabla2.Columns.Add("id_work_orders_tasks");
            //    tabla2.Columns.Add("id_status_work_order");
            //    tabla2.Columns.Add("wo_folio");
            //    tabla2.Columns.Add("creation_date");
            //    tabla2.Columns.Add("duration");
            //    tabla2.Columns.Add("initial_date");
            //    tabla2.Columns.Add("id_priorities");
            //    tabla2.Columns.Add("final_date");
            //    tabla2.Columns.Add("completed_percentage");
            //    tabla2.Columns.Add("created_by");
            //    tabla2.Columns.Add("personnel_description");
            //    tabla2.Columns.Add("code1");
            //    tabla2.Columns.Add("items_log_description");
            //    tabla2.Columns.Add("note");
            //    tabla2.Columns.Add("tasks_log_task_type_main");
            //    tabla2.Columns.Add("parent_description");

            //    Console.ForegroundColor = ConsoleColor.Blue;
            //    Console.WriteLine("[-----------------Requisiciones------------------]");
            //    //Console.WriteLine("Requisiciones Encontradas");
            //    Console.ForegroundColor = ConsoleColor.Yellow;
            //    Console.WriteLine(Requi.Data.Count);
            //    for (int i = 0; i < Requi.Data.Count; i++)
            //    {
            //        string OrdenTra = Requi.Data[i].document.ToString();
            //        Console.WriteLine("\n");
            //        Console.WriteLine("*****************************************************************************************************************************");
            //        Console.ForegroundColor = ConsoleColor.Cyan;
            //        Console.WriteLine("\n");
            //        Console.WriteLine(Requi.Data[i].id.ToString() + "" +
            //            ",\t" + Requi.Data[i].warehouses_source_description.ToString() + "" +
            //            ",\t" + Requi.Data[i].document.ToString() + "" +
            //            ",\t" + Requi.Data[i].date + "" +
            //            ",\t" + Requi.Data[i].code + "" +
            //            ",\t" + Requi.Data[i].folio_source + "" +
            //            ",\t" + Requi.Data[i].movements_states_description);

            //        Cargarordenesfracttal(OrdenTra);

            //        Console.ForegroundColor = ConsoleColor.Green;
            //        for (int n = 0; n < Requi.Data[i].list_items.Count; n++)
            //        {
            //            Console.WriteLine(Requi.Data[i].document + ",\t" + Requi.Data[i].list_items[n].code + ",\t" + Requi.Data[i].list_items[n].items_description + ",\t" + Requi.Data[i].list_items[n].qty + ",\t" + Requi.Data[i].list_items[n].qty_pending + ",\t" + Requi.Data[i].list_items[n].unit_cost_company);

            //            DataRow fila = tabla.NewRow();

            //            fila["id"] = Requi.Data[i].id.ToString();
            //            fila["warehouses_source_description"] = Requi.Data[i].warehouses_source_description.ToString();
            //            fila["document"] = Requi.Data[i].document.ToString();
            //            fila["date"] = Requi.Data[i].date;
            //            fila["code"] = Requi.Data[i].code;
            //            fila["folio_source"] = Requi.Data[i].folio_source;
            //            fila["document1"] = Requi.Data[i].document;
            //            fila["ItemCode"] = Requi.Data[i].list_items[n].code;
            //            fila["items_description"] = Requi.Data[i].list_items[n].items_description;
            //            fila["qty"] = Requi.Data[i].list_items[n].qty;
            //            fila["qty_pending"] = Requi.Data[i].list_items[n].qty_pending;
            //            fila["unit_cost_company"] = Requi.Data[i].list_items[n].unit_cost_company;
            //            fila["movements_states_description"] = Requi.Data[i].movements_states_description;

            //            tabla.Rows.Add(fila);
            //        }


            //        //var cliento = new RestClient("https://app.fracttal.com/api/work_orders/" + OrdenTra);

            //        //var requesto = new RestRequest(Method.GET);
            //        //Authenticate(cliento, requesto);
            //        //IRestResponse responseo = cliento.Execute(requesto);
            //        //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
            //        //Console.WriteLine(jsonResponseo);

            //        //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);
            //        //Console.WriteLine("Lineas Encontradas " + workOrdenes.Data.Count);


            //    }



            //    DataTable dtResult = new DataTable();



            //    dtResult.Columns.Add("id");
            //    dtResult.Columns.Add("warehouses_source_description");
            //    dtResult.Columns.Add("document");
            //    dtResult.Columns.Add("date");
            //    dtResult.Columns.Add("code");
            //    dtResult.Columns.Add("folio_source");
            //    dtResult.Columns.Add("document1");
            //    dtResult.Columns.Add("ItemCode");
            //    dtResult.Columns.Add("items_description");
            //    dtResult.Columns.Add("qty");
            //    dtResult.Columns.Add("qty_pending");
            //    dtResult.Columns.Add("unit_cost_company");
            //    dtResult.Columns.Add("id_work_order");
            //    dtResult.Columns.Add("id_work_orders_tasks");
            //    dtResult.Columns.Add("id_status_work_order");
            //    dtResult.Columns.Add("wo_folio");
            //    dtResult.Columns.Add("creation_date");
            //    dtResult.Columns.Add("duration");
            //    dtResult.Columns.Add("initial_date");
            //    dtResult.Columns.Add("id_priorities");
            //    dtResult.Columns.Add("final_date");
            //    dtResult.Columns.Add("completed_percentage");
            //    dtResult.Columns.Add("created_by");
            //    dtResult.Columns.Add("personnel_description");
            //    dtResult.Columns.Add("code1");
            //    dtResult.Columns.Add("items_log_description");
            //    dtResult.Columns.Add("note");
            //    dtResult.Columns.Add("tasks_log_task_type_main");
            //    dtResult.Columns.Add("movements_states_description");
            //    dtResult.Columns.Add("parent_description");


            //    var dataRows = from dataRows1 in tabla.AsEnumerable()
            //                   join dataRows2 in tabla2.AsEnumerable()
            //                   on dataRows1.Field<string>("document").ToString() equals dataRows2.Field<string>("wo_folio").ToString()
            //                   // where dataRows2.Field<string>("completed_percentage") != "100" 

            //                   orderby dataRows1.Field<string>("document") descending
            //                   select dtResult.LoadDataRow(new object[]
            //                    {
            //                        dataRows1.Field<string>("id"),
            //                        dataRows1.Field<string>("warehouses_source_description"),
            //                        dataRows1.Field<string>("document"),
            //                        dataRows1.Field<string>("date"),
            //                        dataRows1.Field<string>("code"),
            //                        dataRows1.Field<string>("folio_source"),
            //                        dataRows1.Field<string>("document1"),
            //                        dataRows1.Field<string>("ItemCode"),
            //                        dataRows1.Field<string>("items_description"),
            //                        dataRows1.Field<string>("qty"),
            //                        dataRows1.Field<string>("qty_pending"),
            //                        dataRows1.Field<string>("unit_cost_company"),
            //                        dataRows2.Field<string>("id_work_order"),
            //                        dataRows2.Field<string>("id_work_orders_tasks"),
            //                        dataRows2.Field<string>("id_status_work_order"),
            //                        dataRows2.Field<string>("wo_folio"),
            //                        Convert.ToDateTime( dataRows2.Field<string>("creation_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("duration"),
            //                        Convert.ToDateTime( dataRows2.Field<string>("initial_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("id_priorities"),
            //                        Convert.ToDateTime(dataRows2.Field<string>("final_date")).ToString(format: "dd/MM/yyyy"),
            //                        dataRows2.Field<string>("completed_percentage"),
            //                        dataRows2.Field<string>("created_by"),
            //                        dataRows2.Field<string>("personnel_description"),
            //                        dataRows2.Field<string>("code1"),
            //                        dataRows2.Field<string>("items_log_description"),
            //                        dataRows2.Field<string>("note"),
            //                        dataRows2.Field<string>("tasks_log_task_type_main"),
            //                        dataRows1.Field<string>("movements_states_description"),
            //                        dataRows2.Field<string>("parent_description"),
            //                    }, false);

            //    DataTable dataTableLinqJoined = dataRows.CopyToDataTable();


            //    DataTable datadistinct = new DataTable();

            //    datadistinct.Columns.Add("folio_source");
            //    datadistinct.Columns.Add("date");
            //    datadistinct.Columns.Add("document1");
            //    //datadistinct.Columns.Add("ItemsLists");
            //    datadistinct.Columns.Add("ItemCode");
            //    datadistinct.Columns.Add("items_description");
            //    datadistinct.Columns.Add("qty");
            //    datadistinct.Columns.Add("qty_pending");
            //    datadistinct.Columns.Add("unit_cost_company");
            //    datadistinct.Columns.Add("completed_percentage");
            //    datadistinct.Columns.Add("code1");
            //    datadistinct.Columns.Add("personnel_description");
            //    datadistinct.Columns.Add("note");
            //    datadistinct.Columns.Add("id_work_orders_tasks");
            //    datadistinct.Columns.Add("tasks_log_task_type_main");
            //    datadistinct.Columns.Add("movements_states_description");
            //    datadistinct.Columns.Add("parent_description");


            //    var distinctRows = (from /*DataRow*/ dRow in dataTableLinqJoined.AsEnumerable().Distinct()
            //                        orderby dRow.Field<string>("document1") descending

            //                        select datadistinct.LoadDataRow(new object[]
            //                        { dRow.Field<string>("document1"), //0
            //                          //dRow.Field<string>("ItemsLists"),
            //                          Convert.ToDateTime(dRow.Field<string>("date")).ToString(format: "dd/MM/yyyy"), //1
            //                          dRow.Field<string>("folio_source"), //2
            //                            dRow.Field<string>("ItemCode"),//3
            //                            dRow.Field<string>("items_description"),//4
            //                            dRow.Field<string>("qty"),//5
            //                            dRow.Field<string>("qty_pending"),//6
            //                            dRow.Field<string>("unit_cost_company"),//7
            //                            dRow.Field<string>("completed_percentage"),//8
            //                            dRow.Field<string>("code1"),//9
            //                            dRow.Field<string>("personnel_description"),//10
            //                            dRow.Field<string>("note"),//11
            //                            dRow.Field<string>("id_work_orders_tasks"),//12
            //                            dRow.Field<string>("tasks_log_task_type_main"),//13
            //                            dRow.Field<string>("movements_states_description"),//14
            //                            dRow.Field<string>("parent_description")//15
            //                        }, false)).Distinct();


            //    DataTable datadistintos = distinctRows.AsEnumerable().Distinct().CopyToDataTable();

            //    DataTable tbcomparaenca = new DataTable();

            //    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //    // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            //    {
            //        var ArticulosMovimiento = " select distinct [Code],U_Document1 from [@LINEASOFRACTTAL] order by [Code] desc";
            //        conexion.Open();
            //        SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
            //        adaptador.Fill(tbcomparaenca);
            //    }


            //    ///revisando ordenes ya registradas
            //    var idsNotInB = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("Document1"))).Except(tbcomparaenca.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("U_Document1"))));
            //    if (idsNotInB.Count() > 0)
            //    {
            //        DataTable TableC = (from row in datadistintos.AsEnumerable()
            //                            where row["Document1"] != System.DBNull.Value
            //                            join id in idsNotInB
            //                            on row.Field<string>("Document1") equals id
            //                            select row).CopyToDataTable();

            //        //Console.ForegroundColor = ConsoleColor.DarkGreen;
            //        //Console.BackgroundColor = ConsoleColor.Gray;
            //        Console.Write("Verficando Ordenes Registradas... ");
            //        using (var progress = new ProgressBar())
            //        {
            //            for (int h = 0; h <= TableC.Rows.Count; h++)
            //            {
            //                progress.Report((double)h / 100);
            //                System.Threading.Thread.Sleep(20);
            //            }
            //        }

            //        DataTable dtDistinct = new DataTable();

            //        string[] sColumnas = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description", "parent_description" };
            //        dtDistinct = TableC.DefaultView.ToTable(true, sColumnas);

            //        Console.WriteLine("Done.");

            //        //gvDatos.DataSource = TableC.DefaultView;
            //        //gvDatos.ForeColor = System.Drawing.Color.Red;
            //        if (TableC.Rows.Count > 0)
            //        {
            //            InserOrdenes(dtDistinct);
            //            //Guarda las ordenes 
            //            //GuardarOrden();
            //        }
            //        //}

            //        ///Agrega los Complementos o requisiciones que se crearon despues de la carga de las ordenes 
            //        DataTable TbtComparaRequis = new DataTable();


            //        using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //        //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            //        {
            //            var Complementos = "select distinct [U_document1]  from [@LINEASOFRACTTAL] order by 1 desc";
            //            conexion.Open();
            //            SqlDataAdapter adaptador = new SqlDataAdapter(Complementos, conexion);
            //            adaptador.Fill(TbtComparaRequis);
            //        }
            //        var idsNotInBR = distinctRows.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("document1"))).Except(TbtComparaRequis.AsEnumerable().Select(r => Convert.ToString(r.Field<string>("U_document1"))));
            //        if (idsNotInBR.Count() > 0)
            //        {
            //            DataTable TableCR = (from row in datadistintos.AsEnumerable()
            //                                 where row["document1"] != System.DBNull.Value
            //                                 join id in idsNotInBR
            //                                 on row.Field<string>("document1") equals id
            //                                 select row).CopyToDataTable();

            //            DataTable dtDistinctCR = new DataTable();
            //            string[] sColumnasCR = { "folio_source", "date", "document1", "ItemCode", "items_description", "qty", "qty_pending", "unit_cost_company", "Completed_percentage", "code1", "personnel_description", "note", "tasks_log_task_type_main", "movements_states_description" };
            //            dtDistinctCR = TableCR.DefaultView.ToTable(true, sColumnas);

            //            //Console.ForegroundColor = ConsoleColor.DarkGreen;
            //            //Console.BackgroundColor = ConsoleColor.Gray;
            //            Console.Write("Verficando Complementos de las Ordenes Registradas... ");
            //            using (var progress = new ProgressBar())
            //            {
            //                for (int h = 0; h <= TableCR.Rows.Count; h++)
            //                {
            //                    progress.Report((double)h / 100);
            //                    System.Threading.Thread.Sleep(20);
            //                }
            //            }
            //            Console.WriteLine("Done.");

            //            //gvDatos.DataSource = TableC.DefaultView;
            //            //gvDatos.ForeColor = System.Drawing.Color.Red;
            //            if (TableCR.Rows.Count > 0)
            //            {
            //                InserOrdenes(dtDistinctCR);
            //            }
            //        }

            //        //Console.WriteLine("No se Encontraron Ordenes Nuevas");
            //        // para carga en sap
            //        //ConectandoDB();
            //        //AddOrderToDatabase();
            //        // DraftToCocument();

            //    }


            //}

            ////}//pruebas

            //catch (Exception ex)
            //{

            //    Console.WriteLine("Error al conectar" + ex);

            //}

            //Console.WriteLine("Press any key to exit.");
            //Console.ReadKey();
        }


        //--------------------------------------------------------------------------------------

        // public static void PutiteMacizo()
        public static void PutSiteMacizo()
        {

            DataTable dataitems = new DataTable();
            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBTran + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                // var ArticulosMovimiento = "select  * from INVENTARIOS where Codigo between 'VARIOS1831' and 'VARIOS1906'";
                var ArticulosMovimiento = "SELECT [ItemCode],[ItemName],[ItmsGrpNam],[CreateDate],[LastPurPrc],[InvntryUom] FROM [dbo].[V_ArticulosNuevos]";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(dataitems);
            }

            Console.WriteLine("INICIANDO CARGA .");


            if (dataitems.Rows.Count > 0)
                Console.WriteLine("DATOS ENCONTRADOS");

            
                   
            {
                for (int f = 0; f < dataitems.Rows.Count; f++)
                {
                    using (var progress = new ProgressBar())
                    {
                        for (int i = 0; i <= Convert.ToInt32(dataitems.Rows.Count); i++)
                        {
                            progress.Report((double)i / 100);
                            System.Threading.Thread.Sleep(1);
                        }



                        //dt.Rows[i][0].ToString().Trim()
                        string ITEM = dataitems.Rows[f][0].ToString();
                        string DESC = dataitems.Rows[f][1].ToString();
                        string LOCA = dataitems.Rows[f][2].ToString();
                        string UNI = dataitems.Rows[f][5].ToString();
                        double LASTPUR = Convert.ToDouble(dataitems.Rows[f][4].ToString());





                        PostCargaFracttalInventories(ITEM, DESC, LOCA, UNI, LASTPUR);




                    }
                }

            }

        }


        public static void PostCargaFracttalInventories(string Item, string Desccripcion, string Localidad, string Unidad, double Costo)
        { 
            Console.WriteLine("CREANDO ARTICULOS  UNIDAD LOGISTICA");
            string itemcode = Item;
            string descripcion = Desccripcion;
            string localidad = "// LOGISTICA/"; //"// MANTENIMIENTO INDUSTRIAL/
            string unida = Unidad;
            double costo = Costo;




            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            //var client = new RestClient("https://app.fracttal.com/api/inventories/"+itemcode);
            var client = new RestClient("https://app.fracttal.com/api/inventories");

            //var request = new RestRequest(Method.POS);
            var request = new RestRequest(Method.POST);


            Authenticate(client, request);

            request.AddHeader("Content-Type", "application/json");
            // request.AddParameter("undefined", "{\n    \"date\": \"2019-07-22T21:00:00-03\",\n    \"value\": 1,\n    \"serial\": \"100\"\n}", ParameterType.RequestBody);
            request.AddParameter("undefined", "{\n \"code\": \"" + itemcode + "\",  \"field_1\": \"" + descripcion + "\",\n  \"field_2\": \"" + itemcode + "\",\n  \"field_3\": \"\",\n  \"field_4\": \"\",\n  \"field_5\": \"\",\n  \"field_6\": \"\",\n  \"id_warehouse\": 11,\n  \"location\": \"// LOGISTICA/\",\n  \"max_stock_level\": 2,\n  \"min_stock_level\": 0,\n  \"reorder_level\": 1,\n  \"stock\": 0,\n  \"stock_temp\": 0,\n  \"unit_cost_stock\": " + costo + "\n, \"unit_code\": 12\n}", ParameterType.RequestBody);


            //            request.AddParameter("\"" + "Date\":" + "\"" + "2019-07-06T00:00:00-05\""+"," + "\"" + "value\":"+"2"+"," + "\"" + "serial\":"+"\""+"100\"" , ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            Console.ForegroundColor = ConsoleColor.Green;
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine(jsonResponse);
            Console.ResetColor();
            // Console.ReadKey();

            

        }


        //-----------------------Actualiza ITEMS Logistica---------------------------------------------------------------


        public static void PutiteMacizo()
        //public static void PutSiteMacizo()
        {

            DataTable dataitems = new DataTable();
            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBMaci + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                // var ArticulosMovimiento = "select  * from INVENTARIOS where Codigo between 'VARIOS1831' and 'VARIOS1906'";
                var ArticulosMovimiento = "SELECT [ItemCode],[ItemName],[ItmsGrpNam],[CreateDate],[LastPurPrc],[InvntryUom],[OnHand] FROM [dbo].[V_ArticulosNuevos]";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(dataitems);
            }

            Console.WriteLine("DATOS ENCONTRADOS PARA ACTUALIZAR.");


            if (dataitems.Rows.Count > 0)
            {
                for (int f = 0; f < dataitems.Rows.Count; f++)
                {
                    //dt.Rows[i][0].ToString().Trim()
                    string ITEM = dataitems.Rows[f][0].ToString();
                    string DESC = dataitems.Rows[f][1].ToString();
                    string LOCA = dataitems.Rows[f][2].ToString();
                    double LASTPUR = Convert.ToDouble(dataitems.Rows[f][4].ToString());
                    string UNI = dataitems.Rows[f][5].ToString();
                    double STOCK = Convert.ToDouble(dataitems.Rows[f][6].ToString());

                    PutCargaFracttalInventoriesMacizo(ITEM, DESC, LOCA, UNI, LASTPUR, STOCK);


                }

            }

        }

        public static void PutCargaFracttalInventoriesMacizo(string Item, string Desccripcion, string Localidad, string Unidad, double Costo, double Stock)
        {
            Console.WriteLine("ACTUALIZANDO ARTICULOS  UNIDAD MANTENIMIENTO INDUSTRIAL MACIZO");
            string itemcode = Item;
            string descripcion = Desccripcion;
            string localidad = "// MANTENIMIENTO PLANTAS/"; //tiene que ser cambiado por // MACIZO MANTENIMIENTO/
            string unida = Unidad;
            double costo = Costo;
            double stock = Stock;
            string alamcen = "17";//almacen actual tiene que ser cambiado por el 1296


            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/inventories/" + itemcode);
            //var client = new RestClient("https://app.fracttal.com/api/inventories");

            //var request = new RestRequest(Method.POS);
            var request = new RestRequest(Method.PUT);


            Authenticate(client, request);

            request.AddHeader("Content-Type", "application/json");
            // request.AddParameter("undefined", "{\n    \"date\": \"2019-07-22T21:00:00-03\",\n    \"value\": 1,\n    \"serial\": \"100\"\n}", ParameterType.RequestBody);
            request.AddParameter("undefined", "{\n \"code\": \"" + itemcode + "\",  \"field_1\": \"" + descripcion + "\",\n  \"field_2\": \"" + itemcode + "\",\n  \"field_3\": \"\",\n  \"field_4\": \"\",\n  \"field_5\": \"\",\n  \"field_6\": \"\",\n  \"id_warehouse\": \"" + alamcen + "\",\n  \"location\": \"" + localidad + "\",\n  \"max_stock_level\": 2,\n  \"min_stock_level\": 0,\n  \"reorder_level\": 1,\n  \"stock\": \"" + stock + "\",\n  \"stock_temp\": 0,\n  \"unit_cost_stock\": " + costo + "\n, \"unit_code\": 12\n}", ParameterType.RequestBody);


            //            request.AddParameter("\"" + "Date\":" + "\"" + "2019-07-06T00:00:00-05\""+"," + "\"" + "value\":"+"2"+"," + "\"" + "serial\":"+"\""+"100\"" , ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(jsonResponse);
            Console.ResetColor();
            //Console.ReadKey();


        }

        public static void PutiteSantaInes()
        //public static void PutSiteMacizo()
        {

            DataTable dataitems = new DataTable();
            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBSant + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                // var ArticulosMovimiento = "select  * from INVENTARIOS where Codigo between 'VARIOS1831' and 'VARIOS1906'";
                var ArticulosMovimiento = "SELECT [ItemCode],[ItemName],[ItmsGrpNam],[CreateDate],[LastPurPrc],[InvntryUom],[OnHand] FROM [dbo].[V_ArticulosNuevos]";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(dataitems);
            }

            Console.WriteLine("DATOS ENCONTRADOS PARA ACTUALIZAR.");


            if (dataitems.Rows.Count > 0)
            {
                for (int f = 0; f < dataitems.Rows.Count; f++)
                {
                    //dt.Rows[i][0].ToString().Trim()
                    string ITEM = dataitems.Rows[f][0].ToString();
                    string DESC = dataitems.Rows[f][1].ToString();
                    string LOCA = dataitems.Rows[f][2].ToString();
                    double LASTPUR = Convert.ToDouble(dataitems.Rows[f][4].ToString());
                    string UNI = dataitems.Rows[f][5].ToString();
                    double STOCK = Convert.ToDouble(dataitems.Rows[f][6].ToString());

                    PutCargaFracttalInventoriesSantaInes(ITEM, DESC, LOCA, UNI, LASTPUR, STOCK);

                }

            }

        }


        public static void PutCargaFracttalInventoriesSantaInes(string Item, string Desccripcion, string Localidad, string Unidad, double Costo, double Stock)
        {
            Console.WriteLine("ACTUALIZANDO ARTICULOS  UNIDAD MANTENIMIENTO INDUSTRIAL SANTA INES");
            string itemcode = Item;
            string descripcion = Desccripcion;
            string localidad = "// MANTENIMIENTO INDUSTRIAL/ "; //tiene que ser cambiado por // MACIZO MANTENIMIENTO/
            string unida = Unidad;
            double costo = Costo;
            double stock = Stock;
            string alamcen = "17";//almacen actual tiene que ser cambiado por el 1296




            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/inventories/" + itemcode);
            //var client = new RestClient("https://app.fracttal.com/api/inventories");

            //var request = new RestRequest(Method.POS);
            var request = new RestRequest(Method.PUT);


            Authenticate(client, request);

            request.AddHeader("Content-Type", "application/json");
            // request.AddParameter("undefined", "{\n    \"date\": \"2019-07-22T21:00:00-03\",\n    \"value\": 1,\n    \"serial\": \"100\"\n}", ParameterType.RequestBody);
            request.AddParameter("undefined", "{\n \"code\": \"" + itemcode + "\",  \"field_1\": \"" + descripcion + "\",\n  \"field_2\": \"" + itemcode + "\",\n  \"field_3\": \"\",\n  \"field_4\": \"\",\n  \"field_5\": \"\",\n  \"field_6\": \"\",\n  \"id_warehouse\": \"" + alamcen + "\",\n  \"location\": \"" + localidad + "\",\n  \"max_stock_level\": 2,\n  \"min_stock_level\": 0,\n  \"reorder_level\": 1,\n  \"stock\": \"" + stock + "\",\n  \"stock_temp\": 0,\n  \"unit_cost_stock\": " + costo + "\n, \"unit_code\": 12\n}", ParameterType.RequestBody);


            //            request.AddParameter("\"" + "Date\":" + "\"" + "2019-07-06T00:00:00-05\""+"," + "\"" + "value\":"+"2"+"," + "\"" + "serial\":"+"\""+"100\"" , ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(jsonResponse);
            Console.ResetColor();
            //Console.ReadKey();


        }


        public static void PutiteTransflesa()
        //public static void PutSiteMacizo()
        {

            DataTable dataitems = new DataTable();
            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBTran + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                // var ArticulosMovimiento = "select  * from INVENTARIOS where Codigo between 'VARIOS1831' and 'VARIOS1906'";
                var ArticulosMovimiento = "SELECT [ItemCode],[ItemName],[ItmsGrpNam],[CreateDate],[LastPurPrc],[InvntryUom],[OnHand] FROM [dbo].[V_ArticulosNuevos]";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(dataitems);
            }

            Console.WriteLine("DATOS ENCONTRADOS PARA ACTUALIZAR.");


            if (dataitems.Rows.Count > 0)
            {
                for (int f = 0; f < dataitems.Rows.Count; f++)
                {
                    //dt.Rows[i][0].ToString().Trim()
                    string ITEM = dataitems.Rows[f][0].ToString();
                    string DESC = dataitems.Rows[f][1].ToString();
                    string LOCA = dataitems.Rows[f][2].ToString();
                    double LASTPUR = Convert.ToDouble(dataitems.Rows[f][4].ToString());
                    string UNI = dataitems.Rows[f][5].ToString();
                    double STOCK = Convert.ToDouble(dataitems.Rows[f][6].ToString());
                    

                    PutCargaFracttalInventories(ITEM, DESC, LOCA, UNI, LASTPUR,STOCK);

                }

            }

        }



        public static void PutCargaFracttalInventories(string Item, string Desccripcion, string Localidad, string Unidad, double Costo, double Stock)
        {
            Console.WriteLine("ACTUALIZANDO ARTICULOS  UNIDAD LOGISTICA");
            string itemcode = Item;
            string descripcion = Desccripcion;
            string localidad = "// LOGISTICA/"; //"// MANTENIMIENTO INDUSTRIAL/
            string unida = Unidad;
            double costo = Costo;
            double stock = Stock;




            HawkCredential credential = new HawkCredential
            {
                Id = "TLFmgX1Kuef4rsaNxk9z",
                Key = "0EIAQTJhtUvNqBAXkCserYjjRL7P6HP7WhIxBcf67aUynfWXrPCjyxU",
                Algorithm = "sha256"
            };

            _credential = credential;
            //requisiciones
            var client = new RestClient("https://app.fracttal.com/api/inventories/"+itemcode);
            //var client = new RestClient("https://app.fracttal.com/api/inventories");

            //var request = new RestRequest(Method.POS);
            var request = new RestRequest(Method.PUT);


            Authenticate(client, request);

            request.AddHeader("Content-Type", "application/json");
            // request.AddParameter("undefined", "{\n    \"date\": \"2019-07-22T21:00:00-03\",\n    \"value\": 1,\n    \"serial\": \"100\"\n}", ParameterType.RequestBody);
            request.AddParameter("undefined", "{\n \"code\": \"" + itemcode + "\",  \"field_1\": \"" + descripcion + "\",\n  \"field_2\": \"" + itemcode + "\",\n  \"field_3\": \"\",\n  \"field_4\": \"\",\n  \"field_5\": \"\",\n  \"field_6\": \"\",\n  \"id_warehouse\": 11,\n  \"location\": \"// LOGISTICA/\",\n  \"max_stock_level\": 2,\n  \"min_stock_level\": 0,\n  \"reorder_level\": 1,\n  \"stock\": \"" + stock + "\",\n  \"stock_temp\": 0,\n  \"unit_cost_stock\": " + costo + "\n, \"unit_code\": 12\n}", ParameterType.RequestBody);


            //            request.AddParameter("\"" + "Date\":" + "\"" + "2019-07-06T00:00:00-05\""+"," + "\"" + "value\":"+"2"+"," + "\"" + "serial\":"+"\""+"100\"" , ParameterType.RequestBody);

            IRestResponse response = client.Execute(request);
            var jsonResponse = JsonConvert.DeserializeObject(response.Content);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.ResetColor();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(jsonResponse);
            Console.ResetColor();
            //Console.ReadKey();


        }




        //---------------------------------------------------------------------------------------

        public static void Cargarordenesfracttal(string OrdenTra)
        {

            //Console.ForegroundColor = ConsoleColor.Blue;
            //Console.WriteLine("[-------------------Ordenes de Trabajo---------------------]");

            var cliento = new RestClient("https://app.fracttal.com/api/work_orders/" + OrdenTra);


            var requesto = new RestRequest(Method.GET);

            Authenticate(cliento, requesto);

            IRestResponse responseo = cliento.Execute(requesto);

            var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);

            Console.WriteLine();
            WorkOrden workOrdenes = NewMethod(responseo);
            Console.WriteLine("Lineas Encontradas " + workOrdenes.Data.Count);




            //tabla2.Columns.Add("id_work_order");
            //tabla2.Columns.Add("id_work_orders_tasks");
            //tabla2.Columns.Add("id_status_work_order");
            //tabla2.Columns.Add("wo_folio");
            //tabla2.Columns.Add("creation_date");
            //tabla2.Columns.Add("duration");
            //tabla2.Columns.Add("initial_date");
            //tabla2.Columns.Add("id_priorities");
            //tabla2.Columns.Add("final_date");
            //tabla2.Columns.Add("completed_percentage");
            //tabla2.Columns.Add("created_by");
            //tabla2.Columns.Add("personnel_description");
            //tabla2.Columns.Add("code1");
            //tabla2.Columns.Add("items_log_description");
            //tabla2.Columns.Add("note");
            //tabla2.Columns.Add("tasks_log_task_type_main");


            //Console.WriteLine("Ordenes de Trabajo Encontradas en Fracttal");
            // Console.ForegroundColor = ConsoleColor.Yellow;

            //for (int i = 0; i < Requi.Data.Count; i++)
            //{

            //var cliento = new RestClient("https://app.fracttal.com/api/work_orders/" + OrdenTra);

            //var requesto = new RestRequest(Method.GET);
            //Authenticate(cliento, requesto);
            //IRestResponse responseo = cliento.Execute(requesto);
            //var jsonResponseo = JsonConvert.DeserializeObject(responseo.Content);
            //Console.WriteLine( );

            //WorkOrden workOrdenes = JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);
            //Console.WriteLine("Lineas Encontradas " + workOrdenes.Data.Count);




            for (int f = 0; f < workOrdenes.Data.Count; f++)

            {
                //Console.Write("Performing some task... ");
                //using (var progress = new ProgressBar())
                //{
                //    for (int h = 0; h <= workOrdenes.Data.Count; h++)
                //    {
                //        progress.Report((double)h / 100);
                //        System.Threading.Thread.Sleep(20);
                //    }
                //}
                Console.WriteLine("Done.");
                Console.ForegroundColor = ConsoleColor.White;

                Console.WriteLine(workOrdenes.Data[f].id_work_order + "\n" +
                "Tarea---------: " + workOrdenes.Data[f].id_work_orders_tasks + "\n" +
                "Estado--------: " + workOrdenes.Data[f].id_status_work_order + "\n" +
                "Orden---------: " + workOrdenes.Data[f].wo_folio + "\n" +
                "Fecha---------: " + workOrdenes.Data[f].creation_date + "\n" +
                "Duracion------: " + workOrdenes.Data[f].duration + "\n" +
                "Fecha Inicial-: " + workOrdenes.Data[f].initial_date + "\n" +
                "Prioridad-----: " + workOrdenes.Data[f].id_priorities + "\n" +
                "Fecha Final---: " + workOrdenes.Data[f].final_date + "\n" +
                "Porcentaje----: " + workOrdenes.Data[f].completed_percentage + "\n" +
                "Personal------: " + workOrdenes.Data[f].personnel_description + "\n" +
                "Activo--------: " + workOrdenes.Data[f].code + "\n" +
                "Descripcion---: " + workOrdenes.Data[f].items_log_description + "\n" +
                "Nota----------: " + workOrdenes.Data[f].note + "\n" +
                "Tipo----------: " + workOrdenes.Data[f].tasks_log_task_type_main + "\n" + /*Tipo de mantenimiento C - Correctivo cuenta 540101070101 P - Preventivo Cuenta 110901340101*/
                "Localizacion--: " + workOrdenes.Data[f].parent_description + "\n" +
                "Referencias---: " + workOrdenes.Data[f].tasks_log_types_2_description + "\n" +
                "Localidad-----: " + workOrdenes.Data[f].costs_center_description + "\n" +
                "Departamento--: " + "14" /*Siempre el departamento sera el 14*/ + "\n" +
                "Departamento2--: " + workOrdenes.Data[f].tasks_log_types_2_description + "\n" +
                "Solicitud-----: " + workOrdenes.Data[f].id_request
                );

            



                DataRow fila2 = tabla2.NewRow();

                fila2["id_work_order"] = workOrdenes.Data[f].id_work_order;
                fila2["id_work_orders_tasks"] = workOrdenes.Data[f].id_work_orders_tasks;
                fila2["id_status_work_order"] = workOrdenes.Data[f].id_status_work_order;
                fila2["wo_folio"] = workOrdenes.Data[f].wo_folio;
                fila2["creation_date"] = workOrdenes.Data[f].creation_date.ToString();
                fila2["duration"] = workOrdenes.Data[f].duration;
                fila2["initial_date"] = workOrdenes.Data[f].initial_date;
                fila2["id_priorities"] = workOrdenes.Data[f].id_priorities;
                fila2["final_date"] = workOrdenes.Data[f].final_date;
                fila2["completed_percentage"] = workOrdenes.Data[f].completed_percentage;
                fila2["personnel_description"] = workOrdenes.Data[f].personnel_description;
                fila2["code1"] = workOrdenes.Data[f].code;
                fila2["items_log_description"] = workOrdenes.Data[f].items_log_description;
                fila2["note"] = workOrdenes.Data[f].note;
                fila2["tasks_log_task_type_main"] = workOrdenes.Data[f].tasks_log_task_type_main;
                fila2["parent_description"] = workOrdenes.Data[f].parent_description;
                fila2["created_by"] = workOrdenes.Data[f].created_by;
                fila2["description"] = workOrdenes.Data[f].description;
                fila2["tasks_log_types_2_description"] = workOrdenes.Data[f].tasks_log_types_2_description;
                fila2["costs_center_description"] = workOrdenes.Data[f].costs_center_description;
                fila2["Department1"] = "14" ;
                fila2["id_request"] = workOrdenes.Data[f].id_request;
                fila2["groups_2_description"] = workOrdenes.Data[f].tasks_log_types_2_description;

                 tabla2.Rows.Add(fila2);

            }
        }

        private static WorkOrden NewMethod(IRestResponse responseo)
        {
            return JsonConvert.DeserializeObject<WorkOrden>(responseo.Content);
        }

        public static string CrearPassword(int longitud)
        {
            string caracteres = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890";
            StringBuilder res = new StringBuilder();
            Random rnd = new Random();
            while (0 < longitud--)
            {
                res.Append(caracteres[rnd.Next(caracteres.Length)]);
            }
            return res.ToString();
        }

        public static DataTable InserOrdenesLineasFracttal(DataTable dt) {


            //MyDBDataContext sqlObj = new MyDBDataContext();
            //DataTable datas = new DataTable();
            //datas = dt.AsEnumerable().CopyToDataTable();

            //var vardata = (from dRowss in dt.AsEnumerable()
            //               group dRowss by new

            //               {
            //                   dRowss.Field<string>("folio_source")data,

            //               } into newgrp
            //               orderby newgrp.Key
            //               select newgrp.Key
            //    );

            //string Departamento = "14";

            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            ////using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                conexion.Open();
            //    String query = "INSERT INTO [dbo].[@ORDENESFRACTTAL]" +
            //        "   ([Code], " +
            //        "   [Name], " +
            //        "   [U_Document1], " +
            //        "   [U_Date], " +
            //        "   [U_type_main], " +
            //        "   [U_parent_description], " +
            //        "   [Activo], " +
            //        "   [U_costs_center_description], " +
            //        "   [U_Department1]  " +
            //      //  "   [U_id_request] " +
            //       // "   ,[U_groups_2_description] " +
            //        " ) " +
            //        "   VALUES( @Code, " +
            //        "   @Name, " +
            //        "   @U_Document1, " +
            //        "   @U_Date, " +
            //        "   @U_type_main, " +
            //        "   @U_parent_description, " +
            //        "   @Activo, " +
            //        "   @costs_center_description, " +
            //        "   @Department1 " +
            //    //    "   @id_request " +
            //      //  "   , @U_groups_2_description " +
            //        " ) ";
            //    int contador = 0;
            //    using (SqlCommand command = new SqlCommand(query, conexion))

            //        for (int i = 0; i < dt.Rows.Count  ; i++)
            //        {
            //            contador++;

            //            command.Parameters.Clear();
            //            command.Parameters.AddWithValue("@Code", dt.Rows[i][0].ToString().Trim());
            //            command.Parameters.AddWithValue("@Name", (CrearPassword(10).ToString() + contador));
            //            command.Parameters.AddWithValue("@U_Document1", dt.Rows[i][2].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_Date", dt.Rows[i][1].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_type_main", dt.Rows[i][12].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_parent_description", dt.Rows[i][14].ToString().Trim());
            //            command.Parameters.AddWithValue("@Activo" , dt.Rows[i][9].ToString().Trim());
            //            command.Parameters.AddWithValue("@costs_center_description" , dt.Rows[i][16].ToString().Trim());
            //            command.Parameters.AddWithValue("@Department1", /*dt.Rows[i][17].ToString().Trim()*/Departamento.ToString().Trim());
            //         //   command.Parameters.AddWithValue("@id_request", dt.Rows[i][17].ToString().Trim());
            //            //command.Parameters.AddWithValue("@U_groups_2_description", dt.Rows[i][18].ToString().Trim());



            //            int result = command.ExecuteNonQuery();



            //            // Check Error
            //            if (result < 0)
            //            {
            //                Console.WriteLine("Error al Insertar en la Base de Datos ORDENES DE TRABAJO");
            //            }
            //        }




                ////Insertando Lineas 
                String query2 = "INSERT INTO [dbo].[@LINEASOFRACTTAL]([Code]," +
                    "[Name]," +
                    "[U_folio_source]," +
                    "[U_document1]," +
                    "[U_ItemCode]," +
                    "[U_items_description]," +
                    "[U_qty]," +
                    "[U_qty_pending]," +
                    "[U_unit_cost_company]," +
                    "[U_completed_percenta]," +
                    "[U_code1]," +
                    "[U_personnel_descript]," +
                    "[U_note], " +
                    "[U_movements_states_description] " +
                    ")" +
                    "VALUES" +
                    "(@Code" +
                    ",@Name" +
                    ",@U_folio_source" +
                    ",@U_document1" +
                    ",@U_ItemCode" +
                    ",@U_items_description" +
                    ",@U_qty" +
                    ",@U_qty_pending" +
                    ",@U_unit_cost_company" +
                    ",@U_completed_percenta" +
                    ",@U_code1" +
                    ",@U_personnel_descript" +
                    ",@U_note" +
                    ",@U_movements_states_description )";
                int contador2 = 0;
                using (SqlCommand command = new SqlCommand(query2, conexion))

                    for (int i = 0; i < (dt.Rows.Count); i++)
                    {
                        contador2++;

                        command.Parameters.Clear();
                        command.Parameters.AddWithValue("@Code", dt.Rows[i][0].ToString().Trim());
                        command.Parameters.AddWithValue("@Name", (CrearPassword(10).ToString() + contador2).ToString().Trim());
                        command.Parameters.AddWithValue("@U_folio_source", dt.Rows[i][0].ToString().Trim());
                        command.Parameters.AddWithValue("@U_document1", dt.Rows[i][2].ToString().Trim());
                        command.Parameters.AddWithValue("@U_ItemCode", dt.Rows[i][3].ToString().Trim());
                        command.Parameters.AddWithValue("@U_items_description", dt.Rows[i][4].ToString().Trim());
                        command.Parameters.AddWithValue("@U_qty", Convert.ToDecimal(dt.Rows[i][5]).ToString().Trim());
                        command.Parameters.AddWithValue("@U_qty_pending", Convert.ToDecimal(dt.Rows[i][6]).ToString().Trim());
                        command.Parameters.AddWithValue("@U_unit_cost_company", Convert.ToDecimal(dt.Rows[i][7].ToString().Trim()));
                        command.Parameters.AddWithValue("@U_completed_percenta", dt.Rows[i][8].ToString().Trim());
                        command.Parameters.AddWithValue("@U_code1", dt.Rows[i][9].ToString().Trim());
                        command.Parameters.AddWithValue("@U_personnel_descript", dt.Rows[i][10].ToString().Trim());
                        command.Parameters.AddWithValue("@U_note", dt.Rows[i][11].ToString().Trim());
                        command.Parameters.AddWithValue("@U_movements_states_description", dt.Rows[i][13].ToString().Trim());
                        //command.Parameters.AddWithValue("@parent_description", dt.Rows[i][14].ToString().Trim());

                        int result = command.ExecuteNonQuery();

                        command.Parameters.Clear();
                        // Check Error
                        if (result < 0)
                            Console.WriteLine("Error! al incertar datos a la base de datos! REQUISICIONES DE MATERIALES ");

                    }

            }

            Console.Write("REQUISICIONES ENCONTRADAS... ");
            using (var progress = new ProgressBar())
            {
                for (int i = 0; i <= Convert.ToInt32(dt.Rows.Count); i++)
                {
                    progress.Report((double)i / 100);
                    System.Threading.Thread.Sleep(20);
                }
            }
            Console.WriteLine("Done.");
            Console.WriteLine(" !OTS INGRESADA CORRECTAMENTE!");

            Console.WriteLine("ORDENES NUEVAS ALAMCENADAS " + dt.Rows.Count);



            // Console.WriteLine("Iniciando Carga en SAP");

            //Iniciando carga en sap

            return dt;

        }



        public static DataTable InserOrdenesFracttal(DataTable dt)
        {


            //MyDBDataContext sqlObj = new MyDBDataContext();
            //DataTable datas = new DataTable();
            //datas = dt.AsEnumerable().CopyToDataTable();

            //var vardata = (from dRowss in dt.AsEnumerable()
            //               group dRowss by new

            //               {
            //                   dRowss.Field<string>("folio_source")data,

            //               } into newgrp
            //               orderby newgrp.Key
            //               select newgrp.Key
            //    );

            string Departamento = "14";

            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {
                conexion.Open();
                String query = "INSERT INTO [dbo].[@ORDENESFRACTTAL]" +
                    "   ([Code], " +
                    "   [Name], " +
                    "   [U_Document1], " +
                    "   [U_Date], " +
                    "   [U_type_main], " +
                    "   [U_parent_description], " +
                    "   [Activo], " +
                    "   [U_costs_center_description], " +
                    "   [U_Department1]  " +
                    //  "   [U_id_request] " +
                    // "   ,[U_groups_2_description] " +
                    " ) " +
                    "   VALUES( @Code, " +
                    "   @Name, " +
                    "   @U_Document1, " +
                    "   @U_Date, " +
                    "   @U_type_main, " +
                    "   @U_parent_description, " +
                    "   @Activo, " +
                    "   @costs_center_description, " +
                    "   @Department1 " +
                    //    "   @id_request " +
                    //  "   , @U_groups_2_description " +
                    " ) ";
                int contador = 0;
                using (SqlCommand command = new SqlCommand(query, conexion))

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        contador++;

                        command.Parameters.Clear();
                        command.Parameters.AddWithValue("@Code", dt.Rows[i][0].ToString().Trim());
                        command.Parameters.AddWithValue("@Name", (CrearPassword(10).ToString() + contador));
                        command.Parameters.AddWithValue("@U_Document1", dt.Rows[i][2].ToString().Trim());
                        command.Parameters.AddWithValue("@U_Date", dt.Rows[i][1].ToString().Trim());
                        command.Parameters.AddWithValue("@U_type_main", dt.Rows[i][12].ToString().Trim());
                        command.Parameters.AddWithValue("@U_parent_description", dt.Rows[i][14].ToString().Trim());
                        command.Parameters.AddWithValue("@Activo", dt.Rows[i][9].ToString().Trim());
                        command.Parameters.AddWithValue("@costs_center_description", dt.Rows[i][16].ToString().Trim());
                        command.Parameters.AddWithValue("@Department1", /*dt.Rows[i][17].ToString().Trim()*/Departamento.ToString().Trim());
                        //   command.Parameters.AddWithValue("@id_request", dt.Rows[i][17].ToString().Trim());
                        //command.Parameters.AddWithValue("@U_groups_2_description", dt.Rows[i][18].ToString().Trim());



                        int result = command.ExecuteNonQuery();



                        // Check Error
                        if (result < 0)
                        {
                            Console.WriteLine("Error al Insertar en la Base de Datos ORDENES DE TRABAJO");
                        }
                    }




                ////Insertando Lineas 
            //    String query2 = "INSERT INTO [dbo].[@LINEASOFRACTTAL]([Code]," +
            //        "[Name]," +
            //        "[U_folio_source]," +
            //        "[U_document1]," +
            //        "[U_ItemCode]," +
            //        "[U_items_description]," +
            //        "[U_qty]," +
            //        "[U_qty_pending]," +
            //        "[U_unit_cost_company]," +
            //        "[U_completed_percenta]," +
            //        "[U_code1]," +
            //        "[U_personnel_descript]," +
            //        "[U_note], " +
            //        "[U_movements_states_description] " +
            //        ")" +
            //        "VALUES" +
            //        "(@Code" +
            //        ",@Name" +
            //        ",@U_folio_source" +
            //        ",@U_document1" +
            //        ",@U_ItemCode" +
            //        ",@U_items_description" +
            //        ",@U_qty" +
            //        ",@U_qty_pending" +
            //        ",@U_unit_cost_company" +
            //        ",@U_completed_percenta" +
            //        ",@U_code1" +
            //        ",@U_personnel_descript" +
            //        ",@U_note" +
            //        ",@U_movements_states_description )";
            //    int contador2 = 0;
            //    using (SqlCommand command = new SqlCommand(query2, conexion))

            //        for (int i = 0; i < (dt.Rows.Count); i++)
            //        {
            //            contador2++;

            //            command.Parameters.Clear();
            //            command.Parameters.AddWithValue("@Code", dt.Rows[i][0].ToString().Trim());
            //            command.Parameters.AddWithValue("@Name", (CrearPassword(10).ToString() + contador2).ToString().Trim());
            //            command.Parameters.AddWithValue("@U_folio_source", dt.Rows[i][0].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_document1", dt.Rows[i][2].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_ItemCode", dt.Rows[i][3].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_items_description", dt.Rows[i][4].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_qty", Convert.ToDecimal(dt.Rows[i][5]).ToString().Trim());
            //            command.Parameters.AddWithValue("@U_qty_pending", Convert.ToDecimal(dt.Rows[i][6]).ToString().Trim());
            //            command.Parameters.AddWithValue("@U_unit_cost_company", Convert.ToDecimal(dt.Rows[i][7].ToString().Trim()));
            //            command.Parameters.AddWithValue("@U_completed_percenta", dt.Rows[i][8].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_code1", dt.Rows[i][9].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_personnel_descript", dt.Rows[i][10].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_note", dt.Rows[i][11].ToString().Trim());
            //            command.Parameters.AddWithValue("@U_movements_states_description", dt.Rows[i][13].ToString().Trim());
            //            //command.Parameters.AddWithValue("@parent_description", dt.Rows[i][14].ToString().Trim());

            //            int result = command.ExecuteNonQuery();

            //            command.Parameters.Clear();
            //            // Check Error
            //            if (result < 0)
            //                Console.WriteLine("Error! al incertar datos a la base de datos! REQUISICIONES DE MATERIALES ");

            //        }

            }

            //Console.Write("REQUISICIONES ENCONTRADAS... ");
            //using (var progress = new ProgressBar())
            //{
            //    for (int i = 0; i <= Convert.ToInt32(dt.Rows.Count); i++)
            //    {
            //        progress.Report((double)i / 100);
            //        System.Threading.Thread.Sleep(20);
            //    }
            //}
            //Console.WriteLine("Done.");
            //Console.WriteLine(" !OTS INGRESADA CORRECTAMENTE!");

            //Console.WriteLine("ORDENES NUEVAS ALAMCENADAS " + dt.Rows.Count);



            // Console.WriteLine("Iniciando Carga en SAP");

            //Iniciando carga en sap

            return dt;

        }



        //public static string DataTableToJSONWithJSONNet(DataTable table)
        //{
        //    string JSONString = string.Empty;
        //    JSONString = JsonConvert.SerializeObject(table);
        //    return JSONString;

        //}

        //public static string DataTableToJSONWithStringBuilder(DataTable table)
        //{
        //    var JSONString = new StringBuilder();
        //    if (table.Rows.Count > 0)
        //    {
        //        JSONString.Append("header"+":[");
        //        for (int i = 0; i < table.Rows.Count; i++)
        //        {
        //            JSONString.Append("{");
        //            for (int j = 0; j < table.Columns.Count; j++)
        //            {
        //                if (j < table.Columns.Count - 1)
        //                {
        //                    JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\",");
        //                }
        //                else if (j == table.Columns.Count - 1)
        //                {
        //                    JSONString.Append("\"" + table.Columns[j].ColumnName.ToString() + "\":" + "\"" + table.Rows[i][j].ToString() + "\"");
        //                }
        //            }
        //            if (i == table.Rows.Count - 1)
        //            {
        //                JSONString.Append("}");
        //            }
        //            else
        //            {
        //                JSONString.Append("},");
        //            }
        //        }
        //        JSONString.Append("]");


        //    }
        //    return JSONString.ToString();

        //}
        //Console.WriteLine(JSONString.ToString());


        public static void ConectandoDB()
        {
            OrdenApp.oCompany = new SAPbobsCOM.Company();

            // Init Connection Properties
            OrdenApp.oCompany.DbServerType = (SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014) /*(cmbDBType.SelectedIndex + 1)*/;
            OrdenApp.oCompany.Server = "128.0.0.4"; // change to your company server
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";
            OrdenApp.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La; // change to your language
            OrdenApp.oCompany.UseTrusted = false;
            OrdenApp.oCompany.DbUserName = /*txtDBUser.Text.ToString()*/"sa";
            OrdenApp.oCompany.DbPassword = /*txtDBPass.Text.ToString()*/"Ceo2015*";

            //Me.Show() ' shows the form while it's loaded...

            //Create a list of companies...
            try
            {
                OrdenApp.oRecordSet = OrdenApp.oCompany.GetCompanyList(); // get the company list
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "Error 40001");
                //MessageBox.Show(ex.Message + " Error 40001");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                return;
            }

            int temp_int = lErrCode;
            string temp_string = sErrMsg;
            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);

            if (lErrCode != 0)
            {
                Console.WriteLine(sErrMsg + "Error 5000");
                //MessageBox.Show(sErrMsg + " Error 5000");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

            }
            else
            {
                //while (!(OrdenApp.oRecordSet.EoF == true))
                //{
                //    cmbCompany.Items.Add(OrdenApp.oRecordSet.Fields.Item(0).Value);
                //    OrdenApp.oRecordSet.MoveNext();
                //}

                //Disable Controls
                //cmdGetCompanyList.Enabled = false;
                //cmbDBType.Enabled = false;

                ////Enable Controls
                //txtUSer.Enabled = true;
                //txtPassword.Enabled = true;
                //cmdConnect.Enabled = true;
                //cmbCompany.Enabled = true;
                // Console.WriteLine("conectando");
                Connect();
            }

            //Select the first company as default
            //if (cmbCompany.Items.Count > 0)
            //{
            //    cmbCompany.SelectedIndex = 0;
            //}
            //else
            //{
            //    Interaction.MsgBox("There was no Database Found...", 0, "Database not found...");
            //    ProjectData.EndApp(); // Terminate Application...
            //}

            if (OrdenApp.oCompany.Connected) // if already connected
            {
                //Console.WriteLine("Conectado");
                //this.Text = this.Text + ": Connected";
                // Remove the following 2 remark lines if you want to try to connect automatically
                //Else
                //Connect();
            }
            //CreateLinesTable();

        }
        /*-----------------------------------------*/
        public static void ConectandoDB2()
        {
            OrdenApp.oCompany = new SAPbobsCOM.Company();

            // Init Connection Properties
            OrdenApp.oCompany.DbServerType = (SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014) /*(cmbDBType.SelectedIndex + 1)*/;
            OrdenApp.oCompany.Server = "128.0.0.4"; // change to your company server
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";
            OrdenApp.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La; // change to your language
            OrdenApp.oCompany.UseTrusted = false;
            OrdenApp.oCompany.DbUserName = /*txtDBUser.Text.ToString()*/"sa";
            OrdenApp.oCompany.DbPassword = /*txtDBPass.Text.ToString()*/"Ceo2015*";

            //Me.Show() ' shows the form while it's loaded...

            //Create a list of companies...
            try
            {
                OrdenApp.oRecordSet = OrdenApp.oCompany.GetCompanyList(); // get the company list
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "Error 40001");
                //MessageBox.Show(ex.Message + " Error 40001");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                return;
            }

            int temp_int = lErrCode;
            string temp_string = sErrMsg;
            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);

            if (lErrCode != 0)
            {
                Console.WriteLine(sErrMsg + "Error 5000");
                //MessageBox.Show(sErrMsg + " Error 5000");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

            }
            else
            {
                //while (!(OrdenApp.oRecordSet.EoF == true))
                //{
                //    cmbCompany.Items.Add(OrdenApp.oRecordSet.Fields.Item(0).Value);
                //    OrdenApp.oRecordSet.MoveNext();
                //}

                //Disable Controls
                //cmdGetCompanyList.Enabled = false;
                //cmbDBType.Enabled = false;

                ////Enable Controls
                //txtUSer.Enabled = true;
                //txtPassword.Enabled = true;
                //cmdConnect.Enabled = true;
                //cmbCompany.Enabled = true;
                // Console.WriteLine("conectando");
                Connect2();
            }

            //Select the first company as default
            //if (cmbCompany.Items.Count > 0)
            //{
            //    cmbCompany.SelectedIndex = 0;
            //}
            //else
            //{
            //    Interaction.MsgBox("There was no Database Found...", 0, "Database not found...");
            //    ProjectData.EndApp(); // Terminate Application...
            //}

            if (OrdenApp.oCompany.Connected) // if already connected
            {
                //Console.WriteLine("Conectado");
                //this.Text = this.Text + ": Connected";
                // Remove the following 2 remark lines if you want to try to connect automatically
                //Else
                //Connect();
            }
            //CreateLinesTable();

        }


        public static void ConectandoDB3()
        {
            OrdenApp.oCompany = new SAPbobsCOM.Company();

            // Init Connection Properties
            OrdenApp.oCompany.DbServerType = (SAPbobsCOM.BoDataServerTypes.dst_MSSQL2014) /*(cmbDBType.SelectedIndex + 1)*/;
            OrdenApp.oCompany.Server = "128.0.0.4"; // change to your company server
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";
            OrdenApp.oCompany.language = SAPbobsCOM.BoSuppLangs.ln_Spanish_La; // change to your language
            OrdenApp.oCompany.UseTrusted = false;
            OrdenApp.oCompany.DbUserName = /*txtDBUser.Text.ToString()*/"sa";
            OrdenApp.oCompany.DbPassword = /*txtDBPass.Text.ToString()*/"Ceo2015*";

            //Me.Show() ' shows the form while it's loaded...

            //Create a list of companies...
            try
            {
                OrdenApp.oRecordSet = OrdenApp.oCompany.GetCompanyList(); // get the company list
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message + "Error 40001");
                //MessageBox.Show(ex.Message + " Error 40001");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                return;
            }

            int temp_int = lErrCode;
            string temp_string = sErrMsg;
            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);

            if (lErrCode != 0)
            {
                Console.WriteLine(sErrMsg + "Error 5000");
                //MessageBox.Show(sErrMsg + " Error 5000");
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

            }
            else
            {
                //while (!(OrdenApp.oRecordSet.EoF == true))
                //{
                //    cmbCompany.Items.Add(OrdenApp.oRecordSet.Fields.Item(0).Value);
                //    OrdenApp.oRecordSet.MoveNext();
                //}

                //Disable Controls
                //cmdGetCompanyList.Enabled = false;
                //cmbDBType.Enabled = false;

                ////Enable Controls
                //txtUSer.Enabled = true;
                //txtPassword.Enabled = true;
                //cmdConnect.Enabled = true;
                //cmbCompany.Enabled = true;
                // Console.WriteLine("conectando");
                Connect3();
            }

            //Select the first company as default
            //if (cmbCompany.Items.Count > 0)
            //{
            //    cmbCompany.SelectedIndex = 0;
            //}
            //else
            //{
            //    Interaction.MsgBox("There was no Database Found...", 0, "Database not found...");
            //    ProjectData.EndApp(); // Terminate Application...
            //}

            if (OrdenApp.oCompany.Connected) // if already connected
            {
                //Console.WriteLine("Conectado");
                //this.Text = this.Text + ": Connected";
                // Remove the following 2 remark lines if you want to try to connect automatically
                //Else
                //Connect();
            }
            //CreateLinesTable();

        }





        public static void Connect()
        {
            //Cursor = System.Windows.Forms.Cursors.WaitCursor; //Change mouse cursor

            // Set connection properties

            OrdenApp.oCompany.CompanyDB = ServerSqlDBTran.ToString(); /*"TRANSFLESA91".ToString()*//*cmbCompany.Text*/
            OrdenApp.oCompany.UserName = "pro3"/*txtUSer.Text*/;
            OrdenApp.oCompany.Password = "123456"/*txtPassword.Text*/;
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";

            //Try to connect
            lRetCode = OrdenApp.oCompany.Connect();

            if (lRetCode != 0) // if the connection failed
            {
                int temp_int = lErrCode;
                string temp_string = sErrMsg;
                OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                Console.WriteLine("Error al Conectar: " + sErrMsg);
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                // Interaction.MsgBox("Connection Failed - " + sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed");
            }
            if (OrdenApp.oCompany.Connected) // if connected
            {
                //Console.WriteLine("Conectando...: " + OrdenApp.oCompany.CompanyDB);
                //this.Text = this.Text + " - Connected to " + OrdenApp.oCompany.CompanyDB;
                //grpConn.Enabled = false;
                //grpOrder.Enabled = true;
                //LoadGui(); // Load data for UI elements like combo boxes

                //  Console.WriteLine("Conectado !!! " + OrdenApp.oCompany.UserSignature.ToString());




            }
            //AddOrderToDatabase();
            //Cursor = System.Windows.Forms.Cursors.Default; //Change mouse cursor
        }


        /*------------------------------------------------------------------*/


        public static void Connect2()
        {
            //Cursor = System.Windows.Forms.Cursors.WaitCursor; //Change mouse cursor

            // Set connection properties

            OrdenApp.oCompany.CompanyDB = ServerSqlDBMaci.ToString()/*cmbCompany.Text*/;
            OrdenApp.oCompany.UserName = "manager"/*txtUSer.Text*/;
            OrdenApp.oCompany.Password = "maciz0"/*txtPassword.Text*/;
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";

            //Try to connect
            lRetCode = OrdenApp.oCompany.Connect();

            if (lRetCode != 0) // if the connection failed
            {
                int temp_int = lErrCode;
                string temp_string = sErrMsg;
                OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                Console.ForegroundColor = System.ConsoleColor.Red;
                Console.WriteLine("Error al Conectar: " + sErrMsg);
                Console.WriteLine("Problemas en la conexion de SAP  :(   PRECIONE UNA TECLA PARA CONTINUAR");
                Console.ReadKey();

                // Interaction.MsgBox("Connection Failed - " + sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed");
            }
            if (OrdenApp.oCompany.Connected) // if connected
            {
               // Console.WriteLine("Conectando...: " + OrdenApp.oCompany.CompanyDB);
                //this.Text = this.Text + " - Connected to " + OrdenApp.oCompany.CompanyDB;
                //grpConn.Enabled = false;
                //grpOrder.Enabled = true;
                //LoadGui(); // Load data for UI elements like combo boxes

                //  Console.WriteLine("Conectado !!! " + OrdenApp.oCompany.UserSignature.ToString());




            }
            //AddOrderToDatabase();
            //Cursor = System.Windows.Forms.Cursors.Default; //Change mouse cursor
        }


        public static void Connect3()
        {
            //Cursor = System.Windows.Forms.Cursors.WaitCursor; //Change mouse cursor

            // Set connection properties

            OrdenApp.oCompany.CompanyDB = ServerSqlDBSant.ToString()/*cmbCompany.Text*/;
            OrdenApp.oCompany.UserName = "manager"/*txtUSer.Text*/;
            OrdenApp.oCompany.Password = "maciz0"/*txtPassword.Text*/;
            OrdenApp.oCompany.LicenseServer = "128.0.0.10:30000";

            //Try to connect
            lRetCode = OrdenApp.oCompany.Connect();

            if (lRetCode != 0) // if the connection failed
            {
                int temp_int = lErrCode;
                string temp_string = sErrMsg;
                OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                Console.WriteLine("Error al Conectar: " + sErrMsg);
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                // Interaction.MsgBox("Connection Failed - " + sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed");
            }
            if (OrdenApp.oCompany.Connected) // if connected
            {
                // Console.WriteLine("Conectando...: " + OrdenApp.oCompany.CompanyDB);
                //this.Text = this.Text + " - Connected to " + OrdenApp.oCompany.CompanyDB;
                //grpConn.Enabled = false;
                //grpOrder.Enabled = true;
                //LoadGui(); // Load data for UI elements like combo boxes

                //  Console.WriteLine("Conectado !!! " + OrdenApp.oCompany.UserSignature.ToString());




            }
            //AddOrderToDatabase();
            //Cursor = System.Windows.Forms.Cursors.Default; //Change mouse cursor
        }



        //Error handling variables
        public static string sErrMsg;
        public static int lErrCode;
        public static int lRetCode;

        //****************************************************************************
        // This function adds an order to the database
        //****************************************************************************
        public static string DocNumOrders = "";
        public static string Requerimiento = "";
        public static string Departamento2 = "";
        public static string centroCosto = "";
        public static string Activo = "";
        public static string Solicitud = "";
        public static string Solicitante = "";
        public static string Notas = "";


        public static string TipoMantenimiento = "";
        public static void AddOrderToDatabase()
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("-----------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE REQUISICIONES DE INSUMOS--");
            Console.WriteLine("-----------------LOGISTICA---------------------");
            Console.WriteLine("-----------------------------------------------");
            Console.ForegroundColor = ConsoleColor.Green;
            
           
                DocNumOrders = "";
                Requerimiento = "";

                DataTable solicitudes = new DataTable();
                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ";Connection Timeout=100"))


               /*FILTRA LAS REQUISICIONES DE INSUMOS ENVIADAS DESDE FRACTTAL */
                {
                    var ArticulosMovimiento = "SELECT DISTINCT top 100  convert(nvarchar, [Code]) as code,U_Document1 ,[U_type_main], U_Date,U_parent_description  FROM[dbo].[@ORDENESFRACTTAL] (nolock)  " +
                    " where[U_Document1] not in (Select U_Numero_Doc COLLATE SQL_Latin1_General_CP1_CI_AS FROM ODRF23L(nolock) WHERE U_Numero_Doc IS NOT NULL and U_Numero_Doc <> ''" +
                    " and U_Numero_Doc = convert(nvarchar(20), U_Document1) COLLATE SQL_Latin1_General_CP1_CI_AS  )  " +
                    " and U_parent_description like '%LOGISTICA%'" +
                    " group by[code], U_Document1,[U_type_main], U_date, U_parent_description  order by U_Date desc";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(solicitudes);


                // set the CommandTimeout
                //adaptador.SelectCommand.CommandTimeout = 180;  // seconds



                

            }





            for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
                {

                    ConectandoDB();
                    // Connect();

                    // Init the Order object
                    OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                    SAPbobsCOM.SBObob oBob;
                    //set properties of the Order object
                    OrdenApp.oOrder.CardCode = "C00003";
                    OrdenApp.oOrder.CardName = "TRANSFLESA, S.A.";


                    DocNumOrders = "";
                    Requerimiento = "";
                    //for (int k = 0; k < solicitudes.Columns.Count; k++)
                    //{


                    DateTime fechaHoy = DateTime.Now;

                    string fecha = fechaHoy.ToString("dd/MM/yyyy");

                    DocNumOrders = Convert.ToString(solicitudes.Rows[x][0]);
                    Requerimiento = Convert.ToString(solicitudes.Rows[x][1]);

                    //TipoMantenimiento =  Convert.ToString(solicitudes.Rows[x][1]);

                    //string TipoManto = TipoMantenimiento.Substring(0, 1).Replace("M", "X").Replace("NULL", "-").Trim() ;




                    //if (TipoMantenimiento == "NULL")
                    //{

                    //    TipoMantenimiento = "-";

                    //}



                    OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrders);
                    OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrders);
                    OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocCurrency = "QTZ";
                    //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                    OrdenApp.oOrder.Comments = "Cotizacion Realizada desde Interface: " + DocNumOrders;
                    OrdenApp.oOrder.Reference1 = DocNumOrders;
                    OrdenApp.oOrder.Reference2 = DocNumOrders;



                    //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                    DataTable Olineas = new DataTable();
                    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                    //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                    {
                        var LineasOrdenes = "SELECT DISTINCT  [Code]" +
                            ",'' AS NAME" +
                            //",[Name]" + NO VAN POR QUE DUPLICAN LINEAS
                            ",[U_folio_source]" +
                            ",[U_document1]" +
                            ",[U_ItemCode]" +
                            ",SUBSTRING ([U_items_description],0,99) as [U_items_description]" +
                            ",[U_qty]" +
                            ",[U_qty_pending]" +
                            ",[U_unit_cost_company]" +
                            ",[U_completed_percenta]" +
                            ",[U_code1]" +
                            ",[U_personnel_descript]" +
                            ",[U_note]" +
                            ",'' AS U_id_work_orders_tas " +
                            ",isnull(REPLACE(SUBSTRING([U_type_main],1,1),'M','X'),'-')" +  //TIPO DE MANTENIMIENTO
                            "FROM [dbo].[ORDENESFRACTTAL] (nolock)  where [U_Itemcode] is not null and [U_parent_description] like '%LOGISTICA%' and [U_document1] = " + Requerimiento.ToString();
                        conexion.Open();
                        SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenes, conexion);
                        adaptador.Fill(Olineas);
                    //Aumentamos a 2 minutos el tiempo de espera de la ejecución
                  
                }

                    //for (int y = 0; y <= Olineas.Rows.Count; y++)
                    //{

                    if (Olineas.Rows.Count > 0)
                    {

                        // Add lines to the Orer object from the table
                        int i;
                        i = 0;

                        do
                        {
                            OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][4].ToString();
                            OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][5].ToString();
                            OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][6]);
                            OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][8]);
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Vehiculo").Value = Olineas.Rows[i][10];
                            OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                  //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                            OrdenApp.oOrder.Lines.WarehouseCode = "401";//OrdenApp.TableLines.Rows[i][6].ToString();
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][3];
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][14];
                            i += 1;
                            if (i != Olineas.Rows.Count)
                            {
                                OrdenApp.oOrder.Lines.Add();

                            }
                        } while (!(i == Olineas.Rows.Count));
                        OrdenApp.oOrder.Comments = "Orden Realizada desde Interface: " + DocNumOrders;

                        //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                        OrdenApp.oOrder.GetApprovalTemplates();
                        if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                        {
                            //Your document will fire an approval procedure
                            //If you want to add some remarks to your approval you can do this
                            OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                            OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Orden Fracttal: " + DocNumOrders;
                        }


                        Console.ForegroundColor = ConsoleColor.Red;
                        if (OrdenApp.oOrder.Add() != 0) {
                            //                        MessageBox.Show(oCompany.GetLastErrorDescription());

                            Console.WriteLine("Error en la Orden: " + DocNumOrders);
                            Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                            string err = OrdenApp.oCompany.GetLastErrorDescription();
                            EnviaCorreo.EnviaCorreos(err, DocNumOrders, "" );


                        }
                        else { 
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());
                            Console.WriteLine("Cotizacion Enviada a Autorizacion " + DocNumOrders);

                            string succ = OrdenApp.oCompany.GetLastErrorDescription();
                            EnviaCorreo.EnviaCorreos(succ, DocNumOrders, "Orden Enviada a SAP");

                        }


                        OrdenApp.oOrder.Lines.Delete();
                        OrdenApp.oCompany.Disconnect();
                        //using (var progress = new ProgressBar())
                        //{
                        //    for (i = 0; i <= Olineas.Rows.Count; i++)
                        //    {
                        //        progress.Report((double)i / 100);
                        //        System.Threading.Thread.Sleep(Olineas.Rows.Count);
                        //    }
                        //}
                        //Olineas.Clear();
                        //Console.WriteLine("Done.");

                        Console.ForegroundColor = ConsoleColor.Green;

                        



                        if (lRetCode != 0)
                        {
                            int temp_int = lErrCode;
                            string temp_string = sErrMsg;
                            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                            if (lErrCode != -4006) // Incase adding an order failed
                            {

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(lErrCode + "Error " + temp_string);




                            var infoO = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
                            System.Diagnostics.Process.Start(infoO);

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine("--------CERRANDO CARGA DE REQUERIMIENTOS SE ENCONTRO UN PROBLEMA EN LA CARGA A SAP-------");


                            Environment.Exit(0);



                            ///
                            /// Console.ReadKey();

                            //MessageBox.Show(lErrCode + " " + sErrMsg); // Display error message
                        }

                        }

                        //else
                        //{


                        //    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
                        //    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");


                        //}


                    }



                    OrdenApp.oCompany.Disconnect();
                }
            



            var info = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
            System.Diagnostics.Process.Start(info);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("--------CERRANDO CARGA DE REQUERIMIENTOS-------");


            Environment.Exit(0);


         

        }
        /*-------------------------------------------------------------------------------*/
        //Error handling variables
        public static string sErrMsg2;
        public static int lErrCode2;
        public static int lRetCode2;

        //****************************************************************************
        // This function adds an order to the database
        //****************************************************************************
        public static string DocNumOrders2 = "";
        public static string TipoMantenimiento2 = "";
        public static void AddOrderToDatabase2()
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE COTIZACION DE VENTA--");
            Console.WriteLine("---------MANTENIMIENTO INDUSTRIAL----------");
            Console.WriteLine("-------------------MACIZO------------------");

            try
            {
                //ConectandoDB();
                //// Connect();

                //// Init the Order object
                //OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                //SAPbobsCOM.SBObob oBob;

                // set properties of the Order object
                //OrdenApp.oOrder.CardCode = "C00003";
                //OrdenApp.oOrder.CardName = "TRANSFLESA, S.A.";


                //Consultando las ordenes registradas para iniciar en enviarlas a SAP

                DataTable solicitudes = new DataTable();

                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

               // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {

                    var ArticulosMovimiento = " SELECT DISTINCT top 10 convert(nvarchar, [Code]) as code" + //0
                                              ",[U_type_main]" + //1
                                              ",[U_costs_center_description]" + //2
                                              ",[U_personnel_descript]" + //3
                                              ",[U_document1] " +//4
                                              ",[U_department1]" +//5
                                              ",[U_groups_2_description]" +//6
                                              ",[Activo]" + //7
                                              ",[U_id_Request]" +//8
                                              ",[U_note] "+ //9
                                              " FROM[DB_INTERFACE].[dbo].[V_ORDENESFRACTTAL]" +
                                              " WHERE[U_parent_description]" +
                                              " LIKE ltrim(rtrim('%// MACIZO MANTENIMIENTO%') ) " +
                                              " AND[U_parent_description] NOT LIKE ltrim(rtrim('%SANTA INES%'))" +
                                              " AND[U_ItemCode] != ''" +
                                              " AND [Code] not in ('035803','035804','035805','035806','035461','035662','035777','035787','035789')" +
                                             // " AND [Code] = '035638' " +
                                              " order by 1 desc";

                    //var ArticulosMovimiento = " SELECT DISTINCT  convert(nvarchar, [Code]) as code ,[U_type_main]  FROM[dbo].[@ORDENESFRACTTAL] where[Code] not in (Select NumAtCard COLLATE SQL_Latin1_General_CP1_CI_AS FROM ODRF23 WHERE NumAtCard IS NOT NULL and NumAtCard <> '' )   group by [code],[U_type_main]  order by [code] desc";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(solicitudes);
                }

                for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
                {

                    ConectandoDB2();
                    // Connect();

                    // Init the Order object
                    OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                    SAPbobsCOM.SBObob oBob;
                    //set properties of the Order object
                    OrdenApp.oOrder.CardCode = "C00010";
                    OrdenApp.oOrder.CardName = "MACIZO, S.A.";


                    DocNumOrders2 = "";
                    Departamento2 = "";
                    centroCosto = "";
                    Activo = "";
                    Solicitud = "";
                    Notas = "";
                    Solicitante = "";



                    //for (int k = 0; k < solicitudes.Columns.Count; k++)
                    //{


                    DateTime fechaHoy = DateTime.Now;

                    string fecha = fechaHoy.ToString("dd/MM/yyyy");

                    DocNumOrders2 =  Convert.ToString(solicitudes.Rows[x][0]).Trim();
                    Departamento2 = Convert.ToString(solicitudes.Rows[x][5]).Trim();
                    centroCosto = Convert.ToString(solicitudes.Rows[x][2]).Trim();
                    Solicitante = Convert.ToString(solicitudes.Rows[x][3]).Trim();
                    Activo = Convert.ToString(solicitudes.Rows[x][7]).Trim();
                    Solicitud = Convert.ToString(solicitudes.Rows[x][4]).Trim();
                    Notas = Convert.ToString(solicitudes.Rows[x][9]).Trim();


                    



                    OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrders2);
                    OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrders2);
                    OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocCurrency = "QTZ";
                    //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                    OrdenApp.oOrder.Comments = "Cotización Realizada desde interfaz: " + DocNumOrders2 +" "+ Departamento2 +" "+ centroCosto + " "+ Activo;
                    OrdenApp.oOrder.Reference1 = DocNumOrders2;
                    OrdenApp.oOrder.Reference2 = DocNumOrders2;



                    //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                    DataTable Olineas = new DataTable();


                    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

//                    using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                    {
                        var LineasOrdenes = " SELECT  [Code] " +
                           // "--,[NAME]" +
                            ",[U_folio_source]" +
                            ",[U_document1]" +
                            ",SUBSTRING( [U_ItemCode],2,25) AS U_ItemCode" +
                            ",[U_items_description]" +
                            ",[U_qty]" +
                            ",[U_qty_pending]" +
                            ",[U_unit_cost_company]" +
                            ",[U_completed_percenta]" +
                            ",[U_code1]" +
                            ",[U_personnel_descript]" +
                            ",[U_note]" +
                            ",[U_id_work_orders_tas]" +
                            ",[Expr1]" +
                            ",[U_parent_description]" +
                            ",[U_type_main]" +
                            ",[Cuenta]" +
                            ",[ACtivo]" +
                            ",SUBSTRING(U_costs_center_description,1,3) AS U_costs_center_description" +
                            ",[U_department1]" +
                            ",[U_id_Request] " +
                            ",[U_groups_2_description]  " +
                            ",case when [U_groups_2_description] = 'ELECTRICO MANTENIMIENTO' " +
                            "   THEN '015' WHEN[U_groups_2_description] = 'SOLDADURA PLANTAS'" +
                            "   THEN '016' WHEN[U_groups_2_description] = 'SOLDADURA PROYECTOS'" +
                            "   THEN '017' WHEN[U_groups_2_description] = 'SUPERVICION DE PROYECTOS'" +
                            "   THEN '018' WHEN[U_groups_2_description] = 'TORNO MANTENIMIENTO'" +
                            "   THEN '013' WHEN[U_groups_2_description] = 'MECANICA INDUSTRIAL'" +
                            "   THEN '014' END AS U_Department2   " +
                            " FROM [DB_INTERFACE].[dbo].[V_CARGAMANTINDUSTRIALMACIZO] WHERE [Code] = " + DocNumOrders2.ToString();
                        conexion.Open();
                        SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenes, conexion);
                        adaptador.Fill(Olineas);
                    }

                    //for (int y = 0; y <= Olineas.Rows.Count; y++)
                    //{

                    if (Olineas.Rows.Count > 0)
                    {

                        // Add lines to the Orer object from the table
                        int i;
                        i = 0;

                        do
                        {
                            OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][3].ToString();
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Referencia").Value = Olineas.Rows[i][1].ToString();
                           // OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Referencia")
                            OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][4].ToString();
                            OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][5]);
                            OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][7]);
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Activo").Value = Olineas.Rows[i][9];
                            OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                  //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                            OrdenApp.oOrder.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][2];
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][13].ToString().Trim();
                            OrdenApp.oOrder.Lines.CostingCode = Olineas.Rows[i][18].ToString();
                            OrdenApp.oOrder.Lines.CostingCode3 = Olineas.Rows[i][19].ToString().Trim();
                            OrdenApp.oOrder.Lines.CostingCode4 = Olineas.Rows[i][22].ToString().Trim();
                            OrdenApp.oOrder.Lines.COGSAccountCode = Olineas.Rows[i][16].ToString().Trim();


                            i += 1;
                            if (i != Olineas.Rows.Count)
                            {
                                OrdenApp.oOrder.Lines.Add();


                            }
                        } while (!(i == Olineas.Rows.Count));                                            
                        OrdenApp.oOrder.Comments = "    OTS : " + DocNumOrders2 + "" +
                                                   "    SOLICITUD: " +Solicitud +
                                                   "    Solicitante: " +Solicitante+
                                                   "    LOCALIDAD: " + centroCosto + 
                                                   "    DEPARTAMENTO2: "+Departamento2 + 
                                                   "    ACTIVO: " +Activo+ 
                                                   "    NOTAS: " +Notas ;

                        //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                        OrdenApp.oOrder.GetApprovalTemplates();
                        if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                        {
                            //Your document will fire an approval procedure
                            //If you want to add some remarks to your approval you can do this
                            OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                            OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Orden Fracttal: " + DocNumOrders2;
                        }

                        if (OrdenApp.oOrder.Add() != 0)
                        {
                            //                        MessageBox.Show(oCompany.GetLastErrorDescription());
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Green;

                            Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                            Console.WriteLine("Cotizacion Enviada a Autorizacion " + DocNumOrders2);

                            

                        }

                        OrdenApp.oOrder.Lines.Delete();
                        OrdenApp.oCompany.Disconnect();
                        
                        Olineas.Clear();
                        Console.WriteLine("Done.");

                        //Console.ForegroundColor = ConsoleColor.Red;





                        if (lRetCode != 0)
                        {
                            int temp_int = lErrCode;
                            string temp_string = sErrMsg;
                            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                            if (lErrCode != -4006) // Incase adding an order failed
                            {

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(lErrCode + "Error " + temp_string);


                                
                            }

                        }

                     
                       


                    }



                    OrdenApp.oCompany.Disconnect();

                   // AddComplementsToDatabase2();

                }
            }
            catch (Exception ex)
            {


                Console.WriteLine("ERROR: "+ex);

                //ELog.save(OrdenApp.oOrder, ex);

            }

            //Environment.Exit(0);

            //por el momento desabilitado por cambio de base 
            //AddComplementsToDatabase();


            //AddPurchaseOrder();
        
            //AddComplementsToDatabase2();
        }
        /*-----------------------------------------------------------------------------------------------------------------------------------------------------------*/
        //Error handling variables
        public static string sErrMsg3;
        public static int lErrCode3;
        public static int lRetCode3;

        //****************************************************************************
        // This function adds an order to the database
        //****************************************************************************
        public static string DocNumOrders3 = "";
        public static string TipoMantenimiento3 = "";
        public static void AddOrderToDatabase3()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE COTIZACION DE VENTA--");
            Console.WriteLine("---------MANTENIMIENTO INDUSTRIAL----------");
            Console.WriteLine("---------------SANTA INES------------------");

            try
            {
                //ConectandoDB();
                //// Connect();

                //// Init the Order object
                //OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                //SAPbobsCOM.SBObob oBob;

                // set properties of the Order object
                //OrdenApp.oOrder.CardCode = "C00003";
                //OrdenApp.oOrder.CardName = "TRANSFLESA, S.A.";


                //Consultando las ordenes registradas para iniciar en enviarlas a SAP

                DataTable solicitudes = new DataTable();

                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {

                    var ArticulosMovimiento = "SELECT  DISTINCT  convert(nvarchar, [Code]) as code " +
                        ",[U_type_main] FROM [DB_INTERFACE].[dbo].[ORDENESFRACTTAL] " +
                        "WHERE[U_parent_description]  LIKE ltrim(rtrim('%MANTENIMIENTO INDUSTRIAL%') ) " +
                        "AND [U_parent_description]  LIKE ltrim(rtrim('%SANTA INES%') ) " +
                        "AND [U_ItemCode] != ''    order by 1 desc";

                    //var ArticulosMovimiento = " SELECT DISTINCT  convert(nvarchar, [Code]) as code ,[U_type_main]  FROM[dbo].[@ORDENESFRACTTAL] where[Code] not in (Select NumAtCard COLLATE SQL_Latin1_General_CP1_CI_AS FROM ODRF23 WHERE NumAtCard IS NOT NULL and NumAtCard <> '' )   group by [code],[U_type_main]  order by [code] desc";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(solicitudes);
                }

                for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
                {

                    ConectandoDB3();
                    // Connect();

                    // Init the Order object
                    OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                    SAPbobsCOM.SBObob oBob;
                    //set properties of the Order object
                    OrdenApp.oOrder.CardCode = "C00009";
                    OrdenApp.oOrder.CardName = "AGREGADOS SANTA INES, S.A.";


                    DocNumOrders3 = "";
                    //for (int k = 0; k < solicitudes.Columns.Count; k++)
                    //{


                    DateTime fechaHoy = DateTime.Now;

                    string fecha = fechaHoy.ToString("dd/MM/yyyy");

                    DocNumOrders3 = Convert.ToString(solicitudes.Rows[x][0]);


                    //TipoMantenimiento =  Convert.ToString(solicitudes.Rows[x][1]);

                    //string TipoManto = TipoMantenimiento.Substring(0, 1).Replace("M", "X").Replace("NULL", "-").Trim() ;




                    //if (TipoMantenimiento == "NULL")
                    //{

                    //    TipoMantenimiento = "-";

                    //}



                    OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrders3);
                    OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrders3);
                    OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                    OrdenApp.oOrder.DocCurrency = "QTZ";
                    //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                    OrdenApp.oOrder.Comments = "Cotizacion Realizada desde Interface: " + DocNumOrders3;
                    OrdenApp.oOrder.Reference1 = DocNumOrders3;
                    OrdenApp.oOrder.Reference2 = DocNumOrders3;



                    //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                    DataTable Olineas = new DataTable();


                    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                    //                    using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                    {
                        var LineasOrdenes = " SELECT  [Code] " +
                            // "--,[NAME]" +
                            ",[U_folio_source]" +
                            ",[U_document1]" +
                            ",SUBSTRING([U_ItemCode],2,15)" +
                            ",[U_items_description]" +
                            ",[U_qty]" +
                            ",[U_qty_pending]" +
                            ",[U_unit_cost_company]" +
                            ",[U_completed_percenta]" +
                            ",[U_code1]" +
                            ",[U_personnel_descript]" +
                            ",[U_note]" +
                            ",[U_id_work_orders_tas]" +
                            ",[Expr1]" +
                            "FROM[DB_INTERFACE].[dbo].[V_CARGAMANTINDUSTRIALSANTA] WHERE [Code] = " + DocNumOrders3.ToString();
                        conexion.Open();
                        SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenes, conexion);
                        adaptador.Fill(Olineas);
                    }

                    //for (int y = 0; y <= Olineas.Rows.Count; y++)
                    //{

                    if (Olineas.Rows.Count > 0)
                    {

                        // Add lines to the Orer object from the table
                        int i;
                        i = 0;

                        do
                        {
                            OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][3].ToString();
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Referencia").Value = Olineas.Rows[i][1].ToString();

                            OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][4].ToString();
                            OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][5]);
                            OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][7]);
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Activo").Value = Olineas.Rows[i][9];
                            OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                  //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                            OrdenApp.oOrder.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][2];
                            OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][13];
                            i += 1;
                            if (i != Olineas.Rows.Count)
                            {
                                OrdenApp.oOrder.Lines.Add();


                            }
                        } while (!(i == Olineas.Rows.Count));
                        OrdenApp.oOrder.Comments = "Orden Realizada desde Interface: " + DocNumOrders3;

                        //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                        OrdenApp.oOrder.GetApprovalTemplates();
                        if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                        {
                            //Your document will fire an approval procedure
                            //If you want to add some remarks to your approval you can do this
                            OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                            OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Orden Fracttal: " + DocNumOrders3;
                        }

                        if (OrdenApp.oOrder.Add() != 0)
                        {
                            //                        MessageBox.Show(oCompany.GetLastErrorDescription());
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Green;

                            Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                            Console.WriteLine("Cotizacion Enviada a Autorizacion " + DocNumOrders3);

                        }

                        OrdenApp.oOrder.Lines.Delete();
                        OrdenApp.oCompany.Disconnect();

                        Olineas.Clear();
                        Console.WriteLine("Done.");

                        //Console.ForegroundColor = ConsoleColor.Red;





                        if (lRetCode != 0)
                        {
                            int temp_int = lErrCode;
                            string temp_string = sErrMsg;
                            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                            if (lErrCode != -4006) // Incase adding an order failed
                            {

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(lErrCode + "Error " + temp_string);


                                ///
                                /// Console.ReadKey();

                                //MessageBox.Show(lErrCode + " " + sErrMsg); // Display error message
                            }

                        }


                        //else
                        //{


                        //    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
                        //    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");


                        //}


                    }



                    OrdenApp.oCompany.Disconnect();

                    //AddPurchaseOrder3();
                }
            }
            catch (Exception ex)
            {


                Console.WriteLine("ERROR: " + ex);

                //ELog.save(OrdenApp.oOrder, ex);

            }

            //Environment.Exit(0);

            //por el momento desabilitado por cambio de base 
            //AddComplementsToDatabase();


            //AddPurchaseOrder();

        }





        /*-----------------------------------------------------------------------------------------------------------------------------------------------------------*/


        public static string DocNumOrdersC = "";
        public static string NumeroRequi = "";
        public static string TipoMantenimientoC = "";
        public static void AddComplementsToDatabase()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE COTIZACIONES COMPLEMENTOS--");
            Console.WriteLine("-------------------LOGISTICA--------------------");
            Console.WriteLine("------------------------------------------------");

            //Consultando las ordenes registradas para iniciar en enviarlas a SAP

            DataTable solicitudes = new DataTable();


            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

//            using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

            {
                var ArticulosMovimiento = " select * from (SELECT distinct  t0.[U_folio_source]" +
                    ", cast(t0.[U_document1] as Int) as U_document1 " +
                    " FROM [dbo].[@LINEASOFRACTTAL] t0 " +
                    " inner join[dbo].[COMPLEMENTOS] t1 " +
                    " on t0.U_folio_source = t1.U_folio_source " +
                    " where t0.[U_document1] not in (select U_Numero_Doc  COLLATE Modern_Spanish_CI_AS from  DRF1   ) and t1.[U_parent_description]  LIKE ltrim(rtrim('%LOGISTICA%'))    and t1.U_type_main in ('P - Preventivo','C - Correctivo','X - No Aplica','D - Predictivo') " +
                    " )tx order by convert(int, tx.[U_folio_source]) desc ";

                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(solicitudes);
            }

            for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
            {

                ConectandoDB();
                // Connect();

                // Init the Order object
                OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                SAPbobsCOM.SBObob oBob;
                //set properties of the Order object
                OrdenApp.oOrder.CardCode = "C00003";
                OrdenApp.oOrder.CardName = "TRANSFLESA, S.A.";


                DocNumOrdersC = "";
                //for (int k = 0; k < solicitudes.Columns.Count; k++)
                //{



                DateTime fechaHoy = DateTime.Now;

                string fecha = fechaHoy.ToString("dd/MM/yyyy");

                DocNumOrdersC = Convert.ToString(solicitudes.Rows[x][0]);
                NumeroRequi = Convert.ToString(solicitudes.Rows[x][1]);


                OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrdersC);
                OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrdersC);
                OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocCurrency = "QTZ";
                //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                OrdenApp.oOrder.Comments = "Cotizacion Realizada desde Interface: " + DocNumOrdersC;
                OrdenApp.oOrder.Reference1 = DocNumOrdersC;
                OrdenApp.oOrder.Reference2 = DocNumOrdersC;





                DataTable Olineas = new DataTable();


                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

//                using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var LineasOrdenesC = "SELECT DISTINCT  [Code]" +
                        //",'' AS NAME" +
                        //",[Name]" + NO VAN POR QUE DUPLICAN LINEAS
                        ",[U_folio_source]" +
                        ",[U_document1]" +
                        ",[U_ItemCode]" +
                        ",SUBSTRING ([U_items_description],0,99) as [U_items_description]" +
                        ",[U_qty]" +
                        ",[U_qty_pending]" +
                        ",[U_unit_cost_company]" +
                        ",[U_completed_percenta]" +
                        ",[U_code1]" +
                        ",[U_personnel_descript]" +
                        ",[U_note]" +
                        //",[U_type_main]" +
                        ",case when [U_type_main] = ('Mantenimiento Preventivo') then 'P'  when [U_type_main] = ('Mantenimiento Correctivo') then 'C' when " +
                        "[U_type_main] = ('Mantenimiento Predictivo') then 'D'" +
                        "when[U_type_main] = ('No Aplica') then 'X'" +
                        "Else isnull(REPLACE(SUBSTRING([U_type_main],1,1),'M','X'),'-') end " +  //TIPO DE MANTENIMIENTO
                        "FROM [dbo].[COMPLEMENTOS] where /*U_type_main not in ('Mantenimiento Preventivo') and*/ [Code] =  "+ DocNumOrdersC.ToString();
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenesC, conexion);

                    adaptador.Fill(Olineas);
                }



                if (Olineas.Rows.Count > 0)
                {

                    // Add lines to the Orer object from the table
                    int i;
                    i = 0;

                    do
                    {
                        OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][3].ToString();
                        OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][4].ToString();
                        OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][5]);
                        OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][6]);
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Vehiculo").Value = Olineas.Rows[i][9];
                        OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                              //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                        OrdenApp.oOrder.Lines.WarehouseCode = "401";//OrdenApp.TableLines.Rows[i][6].ToString();
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][2];
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][12];
                        i += 1;
                        if (i != Olineas.Rows.Count)
                        {
                            OrdenApp.oOrder.Lines.Add();

                        }
                    } while (!(i == Olineas.Rows.Count));
                    OrdenApp.oOrder.Comments = "Cotizacion Complemento Realizada desde Interface: " + DocNumOrdersC + "-" + NumeroRequi;

                    //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                    OrdenApp.oOrder.GetApprovalTemplates();
                    if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                    {
                        //Your document will fire an approval procedure
                        //If you want to add some remarks to your approval you can do this
                        OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                        OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Cotizacion Complemento Fracttal: " + DocNumOrdersC;
                    }
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    if (OrdenApp.oOrder.Add() != 0) {
                        //                        MessageBox.Show(oCompany.GetLastErrorDescription());
                        Console.WriteLine("Error en la Orden: " + DocNumOrdersC);
                        Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                        string err = OrdenApp.oCompany.GetLastErrorDescription();
                        EnviaCorreo.EnviaCorreos(err, DocNumOrders, "");
                    }
                    else { 
                        Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                        Console.WriteLine("Cotizacion Complemento Enviada a Autorizacion " + DocNumOrders + "-R-" + NumeroRequi);

                        string succ = OrdenApp.oCompany.GetLastErrorDescription();
                        EnviaCorreo.EnviaCorreos(succ, DocNumOrders, NumeroRequi);
                    }


                    OrdenApp.oOrder.Lines.Delete();
                    OrdenApp.oCompany.Disconnect();
                    Console.WriteLine("Desconectado.....");
                    Olineas.Clear();
                    Console.WriteLine("Done.");

                    Console.ForegroundColor = ConsoleColor.Gray;





                    if (lRetCode != 0)
                    {
                        int temp_int = lErrCode;
                        string temp_string = sErrMsg;
                        OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                        if (lErrCode != -4006) // Incase adding an order failed
                        {

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(lErrCode + "Error " + temp_string);

                        }


                    }




                }




            }

            //Environment.Exit(0);


            // AddPurchaseOrder();

        }



        /// <summary>
        /// complementos de mantenimiento industrial
        /// 
        /// </summary>


        public static string DocNumOrdersCM = "";
        public static string NumeroRequiM = "";
        public static string TipoMantenimientoCM = "";
        public static void AddComplementsToDatabase2()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE COTIZACIONES COMPLEMENTOS--");
            Console.WriteLine("----------MANTENIMIENTO INDUSTRIAL--------------");
            Console.WriteLine("------------------------------------------------");

            //Consultando las ordenes registradas para iniciar en enviarlas a SAP

            DataTable solicitudes = new DataTable();


            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //            using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

            {
                var ArticulosMovimiento = "SELECT distinct  t0.[U_folio_source]" +
                    ", cast(t0.[U_document1] as Int) as U_document1" +
                    " FROM [dbo].[@LINEASOFRACTTAL] t0" +
                    " inner join[dbo].[COMPLEMENTOS] t1" +
                    " on t0.U_folio_source = t1.U_folio_source" +
                    " where t0.[U_document1] not in (select U_Numero_Doc  COLLATE Modern_Spanish_CI_AS from  QUT1   )and t1.[U_parent_description]  LIKE ltrim(rtrim('%MANTENIMIENTO INDUSTRIAL%'))    and t1.U_type_main in ('P - Preventivo','C - Correctivo','X - No Aplica','D - Predictivo')" +
                    " order by 1 desc";

                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(solicitudes);
            }

            for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
            {

                ConectandoDB2();
                // Connect();

                // Init the Order object
                OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                SAPbobsCOM.SBObob oBob;
                //set properties of the Order object
                OrdenApp.oOrder.CardCode = "C00010";
                OrdenApp.oOrder.CardName = "MACIZO, S.A.";


                DocNumOrdersC = "";
                //for (int k = 0; k < solicitudes.Columns.Count; k++)
                //{



                DateTime fechaHoy = DateTime.Now;

                string fecha = fechaHoy.ToString("dd/MM/yyyy");

                DocNumOrdersC = Convert.ToString(solicitudes.Rows[x][0]);
                NumeroRequi = Convert.ToString(solicitudes.Rows[x][1]);


                OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrdersC);
                OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrdersC);
                OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocCurrency = "QTZ";
                //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                OrdenApp.oOrder.Comments = "Cotizacion Realizada desde Interface: " + DocNumOrdersC;
                OrdenApp.oOrder.Reference1 = DocNumOrdersC;
                OrdenApp.oOrder.Reference2 = DocNumOrdersC;





                DataTable Olineas = new DataTable();


                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                //                using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var LineasOrdenesC = "SELECT DISTINCT  [Code]" +
                        //",'' AS NAME" +
                        //",[Name]" + NO VAN POR QUE DUPLICAN LINEAS
                        ",[U_folio_source]" +
                        ",[U_document1]" +
                        ",[U_ItemCode]" +
                        ",SUBSTRING ([U_items_description],0,99) as [U_items_description]" +
                        ",[U_qty]" +
                        ",[U_qty_pending]" +
                        ",[U_unit_cost_company]" +
                        ",[U_completed_percenta]" +
                        ",[U_code1]" +
                        ",[U_personnel_descript]" +
                        ",[U_note]" +
                        //",[U_type_main]" +
                        ",case when [U_type_main] = ('Mantenimiento Preventivo') then 'P'  when [U_type_main] = ('Mantenimiento Correctivo') then 'C' when " +
                        "[U_type_main] = ('Mantenimiento Predictivo') then 'D'" +
                        "when[U_type_main] = ('No Aplica') then 'X'" +
                        "Else isnull(REPLACE(SUBSTRING([U_type_main],1,1),'M','X'),'-') end " +  //TIPO DE MANTENIMIENTO
                        "FROM [dbo].[COMPLEMENTOSMACIZO] where /*U_type_main not in ('Mantenimiento Preventivo') and*/ [Code] =  " + DocNumOrdersC.ToString();
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenesC, conexion);

                    adaptador.Fill(Olineas);
                }



                if (Olineas.Rows.Count > 0)
                {

                    // Add lines to the Orer object from the table
                    int i;
                    i = 0;

                    do
                    {
                        OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][3].ToString();
                        OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][4].ToString();
                        OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][5]);
                        OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][6]);
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Activo").Value = Olineas.Rows[i][9];
                        OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                              //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                        OrdenApp.oOrder.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][2];
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][12];
                        i += 1;
                        if (i != Olineas.Rows.Count)
                        {
                            OrdenApp.oOrder.Lines.Add();

                        }
                    } while (!(i == Olineas.Rows.Count));
                    OrdenApp.oOrder.Comments = "Cotizacion Complemento Realizada desde Interface: " + DocNumOrdersC + "-" + NumeroRequi;

                    //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                    OrdenApp.oOrder.GetApprovalTemplates();
                    if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                    {
                        //Your document will fire an approval procedure
                        //If you want to add some remarks to your approval you can do this
                        OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                        OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Cotizacion Complemento Fracttal: " + DocNumOrdersC;
                    }
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    if (OrdenApp.oOrder.Add() != 0) { 
                        //                        MessageBox.Show(oCompany.GetLastErrorDescription());

                        Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                    }
                    else { 
                        Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                    Console.WriteLine("Cotizacion Complemento Enviada a Autorizacion " + DocNumOrders + "-R-" + NumeroRequi);
                    }


                    OrdenApp.oOrder.Lines.Delete();
                    OrdenApp.oCompany.Disconnect();
                    Console.WriteLine("Desconectado.....");
                    Olineas.Clear();
                    Console.WriteLine("Done.");

                    Console.ForegroundColor = ConsoleColor.Gray;





                    if (lRetCode != 0)
                    {
                        int temp_int = lErrCode;
                        string temp_string = sErrMsg;
                        OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                        if (lErrCode != -4006) // Incase adding an order failed
                        {

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(lErrCode + "Error " + temp_string);

                        }


                    }




                }




            }

            //Environment.Exit(0);


          //  AddPurchaseOrder2();

        }

        /*
         Complememtos santa  ines 
             ------------------------------------------------------------------------------------
             */

        public static string DocNumOrdersCS = "";
        public static string NumeroRequiS = "";
        public static string TipoMantenimientoCS = "";
        public static void AddComplementsToDatabase3()
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("------------------------------------------------");
            Console.WriteLine("--INICIANDO CARGA DE COTIZACIONES COMPLEMENTOS--");
            Console.WriteLine("----------MANTENIMIENTO INDUSTRIAL--------------");
            Console.WriteLine("------------------SANTA INES--------------------");

            //Consultando las ordenes registradas para iniciar en enviarlas a SAP

            DataTable solicitudes = new DataTable();


            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

            //            using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

            {
                var ArticulosMovimiento = "SELECT distinct  t0.[U_folio_source]" +
                    ", cast(t0.[U_document1] as Int) as U_document1" +
                    " FROM [dbo].[@LINEASOFRACTTAL] t0" +
                    " inner join[dbo].[COMPLEMENTOSSANTAINES] t1" +
                    " on t0.U_folio_source = t1.U_folio_source" +
                    " where t0.[U_document1] not in (select cast(U_Numero_Doc as nvarchar(25))  COLLATE Modern_Spanish_CI_AS from  QUT1_SANTAINES   ) " +
                    " and t1.[U_parent_description]  LIKE ltrim(rtrim('%MANTENIMIENTO INDUSTRIAL%')) " +
                    " and t1.[U_parent_description]  LIKE ltrim(rtrim('%SANTA INES%'))   " +
                    " and t1.U_type_main in ('P - Preventivo','C - Correctivo','X - No Aplica','D - Predictivo')" +
                    " order by 1 desc";

                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(solicitudes);
            }

            for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
            {

                ConectandoDB3();
                // Connect();

                // Init the Order object
                OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQuotations);

                SAPbobsCOM.SBObob oBob;
                //set properties of the Order object
                OrdenApp.oOrder.CardCode = "C00009";
                OrdenApp.oOrder.CardName = "AGREGADOS SANTA INES, S.A.";


                DocNumOrdersCS = "";
                //for (int k = 0; k < solicitudes.Columns.Count; k++)
                //{



                DateTime fechaHoy = DateTime.Now;

                string fecha = fechaHoy.ToString("dd/MM/yyyy");

                DocNumOrdersCS = Convert.ToString(solicitudes.Rows[x][0]);
                NumeroRequiS = Convert.ToString(solicitudes.Rows[x][1]);


                OrdenApp.oOrder.DocNum = Convert.ToInt32(DocNumOrdersCS);
                OrdenApp.oOrder.NumAtCard = Convert.ToString(DocNumOrdersCS);
                OrdenApp.oOrder.DocDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocDueDate = Convert.ToDateTime(fecha);
                OrdenApp.oOrder.DocCurrency = "QTZ";
                //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                OrdenApp.oOrder.Comments = "Cotizacion Realizada desde Interface: " + DocNumOrdersCS;
                OrdenApp.oOrder.Reference1 = DocNumOrdersCS;
                OrdenApp.oOrder.Reference2 = DocNumOrdersCS;





                DataTable Olineas = new DataTable();


                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                //                using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var LineasOrdenesC = "SELECT DISTINCT  [Code]" +
                        //",'' AS NAME" +
                        //",[Name]" + NO VAN POR QUE DUPLICAN LINEAS
                        ",[U_folio_source]" +
                        ",[U_document1]" +
                        ",[U_ItemCode]" +
                        ",SUBSTRING ([U_items_description],0,99) as [U_items_description]" +
                        ",[U_qty]" +
                        ",[U_qty_pending]" +
                        ",[U_unit_cost_company]" +
                        ",[U_completed_percenta]" +
                        ",[U_code1]" +
                        ",[U_personnel_descript]" +
                        ",[U_note]" +
                        //",[U_type_main]" +
                        ",case when [U_type_main] = ('Mantenimiento Preventivo') then 'P'  when [U_type_main] = ('Mantenimiento Correctivo') then 'C' when " +
                        "[U_type_main] = ('Mantenimiento Predictivo') then 'D'" +
                        "when[U_type_main] = ('No Aplica') then 'X'" +
                        "Else isnull(REPLACE(SUBSTRING([U_type_main],1,1),'M','X'),'-') end " +  //TIPO DE MANTENIMIENTO
                        "FROM [dbo].[COMPLEMENTOSSANTAINES] where /*U_type_main not in ('Mantenimiento Preventivo') and*/ [Code] =  " + DocNumOrdersCS.ToString();
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(LineasOrdenesC, conexion);

                    adaptador.Fill(Olineas);
                }



                if (Olineas.Rows.Count > 0)
                {

                    // Add lines to the Orer object from the table
                    int i;
                    i = 0;

                    do
                    {
                        OrdenApp.oOrder.Lines.ItemCode = Olineas.Rows[i][3].ToString();
                        OrdenApp.oOrder.Lines.ItemDescription = Olineas.Rows[i][4].ToString();
                        OrdenApp.oOrder.Lines.Quantity = System.Convert.ToDouble(Olineas.Rows[i][5]);
                        OrdenApp.oOrder.Lines.UnitPrice = System.Convert.ToDouble(Olineas.Rows[i][6]);
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Activo").Value = Olineas.Rows[i][9];
                        OrdenApp.oOrder.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                              //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                        OrdenApp.oOrder.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = Olineas.Rows[i][2];
                        OrdenApp.oOrder.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = Olineas.Rows[i][12];
                        i += 1;
                        if (i != Olineas.Rows.Count)
                        {
                            OrdenApp.oOrder.Lines.Add();

                        }
                    } while (!(i == Olineas.Rows.Count));
                    OrdenApp.oOrder.Comments = "Cotizacion Complemento Realizada desde Interface: " + DocNumOrdersCS + "-" + NumeroRequiS;

                    //Comentamos para probar que solo cree una ves el documeto //lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
                    OrdenApp.oOrder.GetApprovalTemplates();
                    if (OrdenApp.oOrder.Document_ApprovalRequests.Count > 0 && OrdenApp.oOrder.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                    {
                        //Your document will fire an approval procedure
                        //If you want to add some remarks to your approval you can do this
                        OrdenApp.oOrder.Document_ApprovalRequests.SetCurrentLine(0);
                        OrdenApp.oOrder.Document_ApprovalRequests.Remarks = "Numero de Cotizacion Complemento Fracttal: " + DocNumOrdersCS;
                    }
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    if (OrdenApp.oOrder.Add() != 0) { 
                        //                        MessageBox.Show(oCompany.GetLastErrorDescription());

                        Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                    }
                    else { 
                        Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                    Console.WriteLine("Cotizacion Complemento Enviada a Autorizacion " + DocNumOrdersCS + "-R-" + NumeroRequiS);
                    }


                    OrdenApp.oOrder.Lines.Delete();
                    OrdenApp.oCompany.Disconnect();
                    Console.WriteLine("Desconectado.....");
                    Olineas.Clear();
                    Console.WriteLine("Done.");

                    Console.ForegroundColor = ConsoleColor.Gray;





                    if (lRetCode != 0)
                    {
                        int temp_int = lErrCode;
                        string temp_string = sErrMsg;
                        OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                        if (lErrCode != -4006) // Incase adding an order failed
                        {

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(lErrCode + "Error " + temp_string);

                        }


                    }




                }




            }

            //Environment.Exit(0);


            //  AddPurchaseOrder2();

        }





        public static void AddPurchaseOrder()
        {

            // Console.BackgroundColor = System.ConsoleColor.DarkMagenta;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("-INICIANDO CARGA DE SOLICITUDES DE COMPRA-");
            Console.WriteLine("----------------LOGISTICA-----------------");
            Console.WriteLine("------------------------------------------");

            try
            {

                //ConectandoDB();

                //Consultando las ordenes registradas para iniciar en enviarlas a SAP

                DataTable solicitudes = new DataTable();

                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))


                //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var ArticulosMovimiento = "Select  Distinct top 15 [NumAtCard]" +
                        "   , convert(int, U_Numero_Doc) as U_Numero_Doc" +
                        "   , [Docdate]" +
                        "   ,   [U_Vehiculo]" +
                        "   ,    [Comments]" +
                        "   ,   [U_Mecanico]" +
                        "   ,   [U_JefeMecanica]" +
                        "       FROM QUTATION_Y" +
                        "   order by  convert(int, U_Numero_Doc) desc";
                    //var ArticulosMovimiento = "SELECT 08269";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(solicitudes);

                    //conexion.Close();
                }

                for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
                {
                    ConectandoDB();





                    // Init the Order object
                    OrdenApp.ORequest = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);

                    SAPbobsCOM.SBObob oBob;

                    // set properties of the Order object
                    //OrdenApp.ORequest.CardCode = "C00003";
                    //OrdenApp.ORequest.CardName = "TRANSFLESA, S.A.";
                    //for (int k = 0; k < solicitudes.Columns.Count; k++)
                    //{

                    DateTime fechaHoy = DateTime.Now;

                    string fecha = fechaHoy.ToString("dd/MM/yyyy");

                    DocNumOrders = Convert.ToString(solicitudes.Rows[x][0]);

                    string solicitud = Convert.ToString(solicitudes.Rows[x][1]);


                    string vehiculo = Convert.ToString(solicitudes.Rows[x][3]);


                    string comments = Convert.ToString(solicitudes.Rows[x][4]); //ACTUALIZAR LINEA ***MR
                    string mecanico = Convert.ToString(solicitudes.Rows[x][5]); //ACTUALIZAR LINEA ***MR
                    string jefeMecanica = Convert.ToString(solicitudes.Rows[x][6]); //jefe de mecanica

                    OrdenApp.ORequest.UserFields.Fields.Item("U_Referencia").Value = Convert.ToString(DocNumOrders);
                    OrdenApp.ORequest.UserFields.Fields.Item("U_Vehiculo").Value = vehiculo; //ACTUALIZAR LINEA ***MR
                    OrdenApp.ORequest.UserFields.Fields.Item("U_Mecanico").Value = mecanico; //ACTUALIZAR LINEA ***MR
                    OrdenApp.ORequest.UserFields.Fields.Item("U_JefeMecanica").Value = jefeMecanica; // ingresa jefe de mecanica


                    OrdenApp.ORequest.UserFields.Fields.Item("U_Referencia").Value = Convert.ToString(DocNumOrders);
                    //OrdenApp.ORequest.NumAtCard = Convert.ToString(DocNumOrders);
                    OrdenApp.ORequest.DocDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.DocDueDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.RequriedDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.DocCurrency = "QTZ";
                    //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                    OrdenApp.ORequest.Comments = comments; //"Requisicion Realizada desde la Interface: " + DocNumOrders;
                    OrdenApp.ORequest.RequesterDepartment = 1;
                    OrdenApp.ORequest.UserFields.Fields.Item("U_Categoria").Value = "1";


                    //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                    DataTable OlineasR = new DataTable();
                    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                    // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                    {
                        var LineasRequis = "SELECT tx.[Code]" +
                            //",[Name]" +
                            ",tx.[U_folio_source]" +
                            ",tx.[U_document1]" +
                            ",tx.[U_ItemCode]" +
                            ",SUBSTRING(tx.[U_items_description],0,99)" +
                            ",tx.[U_qty]" +
                            ",tx.[U_qty_pending]" +
                            ",tx.[U_unit_cost_company]" +
                            ",tx.[U_completed_percenta]" +
                            ",tx.[U_code1]" +
                            ",tx.[U_personnel_descript]" +
                            ",tx.[U_note]" +
                            ",isnull(REPLACE(SUBSTRING(tx.[U_type_main],1,1),'M','X'),'-')" +
                            "FROM [dbo].[V_LineasNoexistencia] tx where tx.[U_document1] = " + solicitud.ToString() +
                            " and tx.[U_document1] not in ( SELECT TOP 500 [SOLICITUD] FROM [DB_INTERFACE].[dbo].[VERI_COMPRAS]  where [SOLICITUD] = tx.U_document1  ) order by 1 desc ";
                        conexion.Open();
                        SqlDataAdapter adaptador = new SqlDataAdapter(LineasRequis, conexion);
                        adaptador.Fill(OlineasR);
                        conexion.Close();
                    }

                    //for (int y = 0; y <= Olineas.Rows.Count; y++)
                    //{




                    if (OlineasR.Rows.Count > 0)
                    {

                        // Add lines to the Orer object from the table
                        int i;
                        i = 0;

                        do
                        {
                            OrdenApp.ORequest.Lines.ItemCode = OlineasR.Rows[i][3].ToString();
                            OrdenApp.ORequest.Lines.ItemDescription = OlineasR.Rows[i][4].ToString();
                            OrdenApp.ORequest.Lines.Quantity = System.Convert.ToDouble(OlineasR.Rows[i][5]);
                            OrdenApp.ORequest.Lines.UnitPrice = System.Convert.ToDouble(OlineasR.Rows[i][7]);
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Vehiculo").Value = OlineasR.Rows[i][9];
                            OrdenApp.ORequest.Lines.RequiredDate = Convert.ToDateTime(fecha);
                            OrdenApp.ORequest.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                    //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                            OrdenApp.ORequest.Lines.WarehouseCode = "401";//OrdenApp.TableLines.Rows[i][6].ToString();
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = OlineasR.Rows[i][2];
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = OlineasR.Rows[i][12];
                            i += 1;
                            if (i != OlineasR.Rows.Count)
                            {
                                OrdenApp.ORequest.Lines.Add();

                            }
                        } while (!(i == OlineasR.Rows.Count));

                        //-----------------------------------------------------------------------------
                        //Envia el Documento a Proceso de autorizacion SAP // Try to add the orer to the database
                        OrdenApp.ORequest.GetApprovalTemplates();
                        if (OrdenApp.ORequest.Document_ApprovalRequests.Count > 0 && OrdenApp.ORequest.Document_ApprovalRequests.ApprovalTemplatesID > 0)
                        {
                            //Your document will fire an approval procedure
                            //If you want to add some remarks to your approval you can do this
                            OrdenApp.ORequest.Document_ApprovalRequests.SetCurrentLine(0);
                            OrdenApp.ORequest.Document_ApprovalRequests.Remarks = "Numero de Orden Fracttal: " + DocNumOrders;
                        }

                        //------------------------------------------------------------------------------

                        lRetCode = OrdenApp.ORequest.Add(); // Try to add the orer to the database

                        Console.ForegroundColor = ConsoleColor.Green;

                        if (OrdenApp.oOrder.Add() != 0)
                        {

                            Console.WriteLine("Error en la Orden: " + DocNumOrdersCS);
                            Console.WriteLine(OrdenApp.oCompany.GetLastErrorDescription());
                            string err = OrdenApp.oCompany.GetLastErrorDescription();
                            EnviaCorreo.EnviaCorreos(err, DocNumOrders, "Error al Enviar la Solicitude de Compra");
                        }
                        else
                        {
                            //Console.ForegroundColor = ConsoleColor.Green;
                            //Console.WriteLine(OrdenApp.oCompany.GetNewObjectKey());

                            //using (var progress = new ProgressBar())
                            //{
                            //    for (i = 0; i <= solicitudes.Rows.Count; i++)
                            //    {

                            //        progress.Report((double)i / solicitudes.Rows.Count);
                            //        System.Threading.Thread.Sleep(50);


                            //    }
                               
                            //    Console.WriteLine("SOLICITUDES DE COMPRA CREADA  " + DocNumOrders + " SOLICITUD FRACTTAL  " + solicitud);

                            //}


                            Console.WriteLine("Solicitud de Compra Enviada a Autorizacion " + DocNumOrdersCS + "-R-" + NumeroRequiS + "Vehiculo: " + vehiculo);

                            string mensaje = "Solicitud de Compra Enviada a Autorizacion " + DocNumOrdersCS + "-R-" + NumeroRequiS + "Vehiculo: " + vehiculo;


                            Console.ForegroundColor = ConsoleColor.Green;

                            Console.WriteLine("Solicitud de Compra Enviada a Autorizacion " + DocNumOrders);

                            string succ = OrdenApp.oCompany.GetLastErrorDescription();
                            EnviaCorreo.EnviaCorreos(succ, DocNumOrders, mensaje);

                        }

                        OrdenApp.ORequest.Lines.Delete();
                        

                    }
                    OrdenApp.oCompany.Disconnect();

                    Console.ForegroundColor = ConsoleColor.DarkYellow;




                    int l;
                    l = 0;


                    using (var progress = new ProgressBar())
                    {
                        for (l = 0; l <= solicitudes.Rows.Count; l++)
                        {
                            
                            progress.Report((double)l / solicitudes.Rows.Count);
                            System.Threading.Thread.Sleep(50);

                          
                        }

                        Console.WriteLine("!NO HAY SOLICITUDES DE COMPRA  " + DocNumOrders + " SOLICITUD FRACTTAL  " + solicitud);

                    }


                   

                    

                }



                
            }
            catch (Exception e)
            {
                Console.WriteLine("Error en ingreso de Cotizacion de Compra " + e);

            }

            var info = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
            System.Diagnostics.Process.Start(info);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("------CERRANDO CARGA DE SOLICITUDES DE COMPRA-------");


            Environment.Exit(0);



        }

        /*------------------------------------------------------------------------------------------*/
        public static void AddPurchaseOrder2()
        {

            // Console.BackgroundColor = System.ConsoleColor.DarkMagenta;
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("-INICIANDO CARGA DE SOLICITUDES DE COMPRA-");
            Console.WriteLine("---------MANTENIMIENTO INDUSTRIAL----------");
            Console.WriteLine("------------------MACIZO------------------");

            //try
            //{

              //  ConectandoDB2();

                //Consultando las ordenes registradas para iniciar en enviarlas a SAP

                DataTable solicitudes = new DataTable();

                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))


                //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var ArticulosMovimiento = "SELECT distinct t0.[Code] , t0.U_code1 FROM[DB_INTERFACE].[dbo].[V_CARGACOMPRASMACIZO] t0 where t0.U_Date BETWEEN convert(nvarchar, DATEADD(DAY,-60, GETDATE())) and convert(nvarchar,GETDATE())  order by [Code] desc ";
                    // var ArticulosMovimiento = "SELECT 01990";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                    adaptador.Fill(solicitudes);
                }

                for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
                {
                    ConectandoDB2();

                    //using (var progress = new ProgressBar())
                    //{
                    //    for (int i = 0; i <= solicitudes.Rows.Count; i++)
                    //    {
                    //        progress.Report((double)i / solicitudes.Rows.Count);
                    //        System.Threading.Thread.Sleep(solicitudes.Rows.Count);
                    //    }
                    //}



                    // Init the Order object
                    OrdenApp.ORequest = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);

                    SAPbobsCOM.SBObob oBob;

                    // set properties of the Order object
                    //OrdenApp.ORequest.CardCode = "C00003";
                    //OrdenApp.ORequest.CardName = "TRANSFLESA, S.A.";
                    //for (int k = 0; k < solicitudes.Columns.Count; k++)
                    //{

                    DateTime fechaHoy = DateTime.Now;

                    string fecha = fechaHoy.ToString("dd/MM/yyyy");

                    DocNumOrders = Convert.ToString(solicitudes.Rows[x][0]);

                    string vehiculo = Convert.ToString(solicitudes.Rows[x][1]);

                    OrdenApp.ORequest.UserFields.Fields.Item("U_Referencia").Value = Convert.ToString(DocNumOrders);
                    //OrdenApp.ORequest.NumAtCard = Convert.ToString(DocNumOrders);
                    OrdenApp.ORequest.DocDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.DocDueDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.RequriedDate = Convert.ToDateTime(fecha);
                    OrdenApp.ORequest.DocCurrency = "QTZ";
                    //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                    OrdenApp.ORequest.Comments = "Requisicion Realizada desde la Interface: " + DocNumOrders;


                    //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                    DataTable OlineasR = new DataTable();
                    using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                    // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                    {
                        var LineasRequis = "SELECT tx.[Code]" +
                            //",[Name]" +
                            ",tx.[U_folio_source]" +
                            ",tx.[U_document1]" +
                            ",tx.[U_ItemCode]" +
                            ",SUBSTRING(tx.[U_items_description],0,99)" +
                            ",tx.[U_qty]" +
                            ",tx.[U_qty_pending]" +
                            ",tx.[U_unit_cost_company]" +
                            ",tx.[U_completed_percenta]" +
                            ",tx.[U_code1]" +
                            ",tx.[U_personnel_descript]" +
                            ",tx.[U_note]" +
                            ",isnull(REPLACE(SUBSTRING(tx.[U_type_main],1,1),'M','X'),'-')" +
                            "FROM [dbo].[V_LineasNoexistencia2] tx where tx.[Code] = " + DocNumOrders.ToString() +
                            " and tx.[U_document1] not in (select  X.U_Numero_Doc COLLATE SQL_Latin1_General_CP1_CI_AS " +
                            " from INVERFFACSA91.dbo.PRQ1 X where x.U_Numero_Doc  COLLATE SQL_Latin1_General_CP1_CI_AS = tx.[U_document1] ) order by 1 desc ";
                        conexion.Open();
                        SqlDataAdapter adaptador = new SqlDataAdapter(LineasRequis, conexion);
                        adaptador.Fill(OlineasR);
                    }

                    //for (int y = 0; y <= Olineas.Rows.Count; y++)
                    //{

                    if (OlineasR.Rows.Count > 0)
                    {

                        // Add lines to the Orer object from the table
                        int i;
                        i = 0;

                        do
                        {
                            OrdenApp.ORequest.Lines.ItemCode = OlineasR.Rows[i][3].ToString();
                            OrdenApp.ORequest.Lines.ItemDescription = OlineasR.Rows[i][4].ToString();
                            OrdenApp.ORequest.Lines.Quantity = System.Convert.ToDouble(OlineasR.Rows[i][5]);
                            OrdenApp.ORequest.Lines.UnitPrice = System.Convert.ToDouble(OlineasR.Rows[i][7]);
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Activo").Value = OlineasR.Rows[i][9];
                            OrdenApp.ORequest.Lines.RequiredDate = Convert.ToDateTime(fecha);
                            OrdenApp.ORequest.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                    //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                            OrdenApp.ORequest.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = OlineasR.Rows[i][2];
                            OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = OlineasR.Rows[i][12];
                            i += 1;
                            if (i != OlineasR.Rows.Count)
                            {
                                OrdenApp.ORequest.Lines.Add();

                            }
                        } while (!(i == OlineasR.Rows.Count));


                        lRetCode = OrdenApp.ORequest.Add(); // Try to add the orer to the database

                        OrdenApp.ORequest.Lines.Delete();

                        OrdenApp.oCompany.Disconnect();


                        using (var progress = new ProgressBar())
                        {
                            for (i = 0; i <= OlineasR.Rows.Count; i++)
                            {
                                progress.Report((double)i / 100);
                                System.Threading.Thread.Sleep(OlineasR.Rows.Count);
                            }
                        }

                        OlineasR.Clear();
                        Console.WriteLine("Done.");

                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Solicitud de Compra Creada en SAP " + DocNumOrders + "-" + vehiculo);


                        if (lRetCode != 0)
                        {
                            int temp_int = lErrCode;
                            string temp_string = sErrMsg;
                            OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                            if (lErrCode != -4006) // Incase adding an order failed
                            {

                                Console.ForegroundColor = ConsoleColor.Red;
                                Console.WriteLine(lErrCode + "Error " + temp_string);

                                string path = HttpContext.Current.Request.MapPath("~");
                                Log oLog = new Log("Logs/", path);
                                oLog.Add("Hola paso algo" + temp_string);


                            }


                        }

                    //else
                    //{


                    //    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
                    //    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");
                    // Environment.Exit(0);
                    // DraftToCocument();
                    //}

                    OrdenApp.oCompany.Disconnect();
                }

              //  AddOrderToDatabase3();
                }

            //}
            //catch (Exception ex)
            //{

            //    Console.WriteLine("Hay un problema :" + ex);

            //}

             Environment.Exit(0);

            // DraftToCocument();

            
        }
        /// <summary>
        /// ------------------------------------------------------------------------------------------------------------------------------------------------
        /// </summary>
        /// 


        public static void AddPurchaseOrder3()
        {

            // Console.BackgroundColor = System.ConsoleColor.DarkMagenta;
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("------------------------------------------");
            Console.WriteLine("-INICIANDO CARGA DE SOLICITUDES DE COMPRA-");
            Console.WriteLine("---------MANTENIMIENTO INDUSTRIAL----------");
            Console.WriteLine("----------------SANTA INES-----------------");

            //try
            //{

            ConectandoDB2();

            //Consultando las ordenes registradas para iniciar en enviarlas a SAP

            DataTable solicitudes = new DataTable();

            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))


            //using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

            {
                var ArticulosMovimiento = "SELECT t0.[Code], t0.U_code1 FROM[DB_INTERFACE].[dbo].[V_CARGAMANTINDUSTRIALSANTA] t0 order by[Code] desc";
                // var ArticulosMovimiento = "SELECT 01990";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
                adaptador.Fill(solicitudes);
            }

            for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
            {
                ConectandoDB3();

                //using (var progress = new ProgressBar())
                //{
                //    for (int i = 0; i <= solicitudes.Rows.Count; i++)
                //    {
                //        progress.Report((double)i / solicitudes.Rows.Count);
                //        System.Threading.Thread.Sleep(solicitudes.Rows.Count);
                //    }
                //}



                // Init the Order object
                OrdenApp.ORequest = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);

                SAPbobsCOM.SBObob oBob;

                // set properties of the Order object
                //OrdenApp.ORequest.CardCode = "C00003";
                //OrdenApp.ORequest.CardName = "TRANSFLESA, S.A.";
                //for (int k = 0; k < solicitudes.Columns.Count; k++)
                //{

                DateTime fechaHoy = DateTime.Now;

                string fecha = fechaHoy.ToString("dd/MM/yyyy");

                DocNumOrders = Convert.ToString(solicitudes.Rows[x][0]);

                string vehiculo = Convert.ToString(solicitudes.Rows[x][1]);

                OrdenApp.ORequest.UserFields.Fields.Item("U_Referencia").Value = Convert.ToString(DocNumOrders);
                //OrdenApp.ORequest.NumAtCard = Convert.ToString(DocNumOrders);
                OrdenApp.ORequest.DocDate = Convert.ToDateTime(fecha);
                OrdenApp.ORequest.DocDueDate = Convert.ToDateTime(fecha);
                OrdenApp.ORequest.RequriedDate = Convert.ToDateTime(fecha);
                OrdenApp.ORequest.DocCurrency = "QTZ";
                //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
                OrdenApp.ORequest.Comments = "Requisicion Realizada desde la Interface: " + DocNumOrders;


                //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


                DataTable OlineasR = new DataTable();
                using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBINTER + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

                // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

                {
                    var LineasRequis = "SELECT tx.[Code]" +
                        //",[Name]" +
                        ",tx.[U_folio_source]" +
                        ",tx.[U_document1]" +
                        ",tx.[U_ItemCode]" +
                        ",SUBSTRING(tx.[U_items_description],0,99)" +
                        ",tx.[U_qty]" +
                        ",tx.[U_qty_pending]" +
                        ",tx.[U_unit_cost_company]" +
                        ",tx.[U_completed_percenta]" +
                        ",tx.[U_code1]" +
                        ",tx.[U_personnel_descript]" +
                        ",tx.[U_note]" +
                        ",isnull(REPLACE(SUBSTRING(tx.[U_type_main],1,1),'M','X'),'-')" +
                        "FROM [dbo].[V_LineasNoexistencia3] tx where tx.[Code] = " + DocNumOrders.ToString() +
                        " and tx.[U_document1] not in (select  X.U_Numero_Doc COLLATE SQL_Latin1_General_CP1_CI_AS " +
                        " from SBO_SANTAINES.dbo.PRQ1 X where x.U_Numero_Doc  COLLATE SQL_Latin1_General_CP1_CI_AS = tx.[U_document1] ) order by 1 desc ";
                    conexion.Open();
                    SqlDataAdapter adaptador = new SqlDataAdapter(LineasRequis, conexion);
                    adaptador.Fill(OlineasR);
                }

                //for (int y = 0; y <= Olineas.Rows.Count; y++)
                //{

                if (OlineasR.Rows.Count > 0)
                {

                    // Add lines to the Orer object from the table
                    int i;
                    i = 0;

                    do
                    {
                        OrdenApp.ORequest.Lines.ItemCode = OlineasR.Rows[i][3].ToString();
                        OrdenApp.ORequest.Lines.ItemDescription = OlineasR.Rows[i][4].ToString();
                        OrdenApp.ORequest.Lines.Quantity = System.Convert.ToDouble(OlineasR.Rows[i][5]);
                        OrdenApp.ORequest.Lines.UnitPrice = System.Convert.ToDouble(OlineasR.Rows[i][7]);
                        OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Activo").Value = OlineasR.Rows[i][9];
                        OrdenApp.ORequest.Lines.RequiredDate = Convert.ToDateTime(fecha);
                        OrdenApp.ORequest.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
                                                                //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
                        OrdenApp.ORequest.Lines.WarehouseCode = "ZR01";//OrdenApp.TableLines.Rows[i][6].ToString();
                        OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = OlineasR.Rows[i][2];
                        OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = OlineasR.Rows[i][12];
                        i += 1;
                        if (i != OlineasR.Rows.Count)
                        {
                            OrdenApp.ORequest.Lines.Add();

                        }
                    } while (!(i == OlineasR.Rows.Count));


                    lRetCode = OrdenApp.ORequest.Add(); // Try to add the orer to the database

                    OrdenApp.ORequest.Lines.Delete();

                    OrdenApp.oCompany.Disconnect();


                    using (var progress = new ProgressBar())
                    {
                        for (i = 0; i <= OlineasR.Rows.Count; i++)
                        {
                            progress.Report((double)i / 100);
                            System.Threading.Thread.Sleep(OlineasR.Rows.Count);
                        }
                    }

                    OlineasR.Clear();
                    Console.WriteLine("Done.");

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("Solicitud de Compra Creada en SAP " + DocNumOrders + "-" + vehiculo);


                    if (lRetCode != 0)
                    {
                        int temp_int = lErrCode;
                        string temp_string = sErrMsg;
                        OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
                        if (lErrCode != -4006) // Incase adding an order failed
                        {

                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(lErrCode + "Error " + temp_string);

                            string path = HttpContext.Current.Request.MapPath("~");
                            Log oLog = new Log("Logs/", path);
                            oLog.Add("Hola paso algo" + temp_string);


                        }


                    }

                    //else
                    //{


                    //    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
                    //    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");
                    // Environment.Exit(0);
                    // DraftToCocument();
                    //}


                }


            }

            //}
            //catch (Exception ex)
            //{

            //    Console.WriteLine("Hay un problema :" + ex);

            //}

            Environment.Exit(0);

            // DraftToCocument();


        }




        //public static void AddPurchaseOrderComplementos()
        //{

        //    // Console.BackgroundColor = System.ConsoleColor.DarkMagenta;
        //    Console.ForegroundColor = ConsoleColor.Green;
        //    Console.WriteLine("------------------------------------------");
        //    Console.WriteLine("INICIANDO CARGA DE SOLICITUDES DE COMPRA DE COMPLEMENTOS");
        //    Console.WriteLine("------------------------------------------");


        //    try
        //    {


        //        //Consultando las ordenes registradas para iniciar en enviarlas a SAP

        //        DataTable solicitudes = new DataTable();
        //        using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

        //        {
        //            var ArticulosMovimiento = "SELECT DISTINCT [Code] FROM [DB_INTERFACE].[dbo].[COMPLEMENTOS]";
        //            // var ArticulosMovimiento = "SELECT 01990";
        //            conexion.Open();
        //            SqlDataAdapter adaptador = new SqlDataAdapter(ArticulosMovimiento, conexion);
        //            adaptador.Fill(solicitudes);
        //        }

        //        for (int x = 0; x <= solicitudes.Rows.Count - 1; x++)
        //        {
        //            ConectandoDB();

        //            // Init the Order object
        //            OrdenApp.ORequest = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseRequest);

        //            SAPbobsCOM.SBObob oBob;

        //            // set properties of the Order object
        //            //OrdenApp.ORequest.CardCode = "C00003";
        //            //OrdenApp.ORequest.CardName = "TRANSFLESA, S.A.";
        //            //for (int k = 0; k < solicitudes.Columns.Count; k++)
        //            //{

        //            DateTime fechaHoy = DateTime.Now;

        //            string fecha = fechaHoy.ToString("dd/MM/yyyy");

        //            DocNumOrders = Convert.ToString(solicitudes.Rows[x][0]);

        //            OrdenApp.ORequest.UserFields.Fields.Item("U_Referencia").Value = Convert.ToString(DocNumOrders);
        //            //OrdenApp.ORequest.NumAtCard = Convert.ToString(DocNumOrders);
        //            OrdenApp.ORequest.DocDate = Convert.ToDateTime(fecha);
        //            OrdenApp.ORequest.DocDueDate = Convert.ToDateTime(fecha);
        //            OrdenApp.ORequest.RequriedDate = Convert.ToDateTime(fecha);
        //            OrdenApp.ORequest.DocCurrency = "QTZ";
        //            //OrdenApp.oOrder.AuthorizationStatus.Equals("W");
        //            OrdenApp.ORequest.Comments = "Requisicion Realizada desde la Interface: " + DocNumOrders;


        //            //OrdenApp.TableLines.AcceptChanges(); // Update the lines table


        //            DataTable OlineasR = new DataTable();

        //            using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=DB_INTERFACE;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))

        //            {
        //                var LineasRequis = "SELECT tx.[Code]" +
        //                    //",[Name]" +
        //                    ",tx.[U_folio_source]" +
        //                    ",tx.[U_document1]" +
        //                    ",tx.[U_ItemCode]" +
        //                    ", SUBSTRING(tx.[U_items_description],0,100)" +
        //                    ",tx.[U_qty]" +
        //                    ",tx.[U_qty_pending]" +
        //                    ",tx.[U_unit_cost_company]" +
        //                    ",tx.[U_completed_percenta]" +
        //                    ",tx.[U_code1]" +
        //                    ",tx.[U_personnel_descript]" +
        //                    ",tx.[U_note]" +
        //                    ",isnull(REPLACE(SUBSTRING(tx.[U_type_main],1,1),'M','X'),'-')" +
        //                    "FROM [dbo].[V_LineasNoexistencia] tx where tx.[Code] = " + DocNumOrders.ToString() +
        //                    " and tx.[U_document1] not in (select  X.U_Numero_Doc COLLATE SQL_Latin1_General_CP1_CI_AS " +
        //                    " from TRANSFLESA91.dbo.PRQ1 X where x.U_Numero_Doc  COLLATE SQL_Latin1_General_CP1_CI_AS = tx.[U_document1] ) and [U_type_main] not in ('Mantenimiento Preventivo') ";
        //                conexion.Open();
        //                SqlDataAdapter adaptador = new SqlDataAdapter(LineasRequis, conexion);
        //                adaptador.Fill(OlineasR);
        //            }

        //            //for (int y = 0; y <= Olineas.Rows.Count; y++)
        //            //{

        //            if (OlineasR.Rows.Count > 0)
        //            {

        //                // Add lines to the Orer object from the table
        //                int i;
        //                i = 0;

        //                do
        //                {
        //                    OrdenApp.ORequest.Lines.ItemCode = OlineasR.Rows[i][3].ToString();
        //                    OrdenApp.ORequest.Lines.ItemDescription = OlineasR.Rows[i][4].ToString();
        //                    OrdenApp.ORequest.Lines.Quantity = System.Convert.ToDouble(OlineasR.Rows[i][5]);
        //                    OrdenApp.ORequest.Lines.UnitPrice = System.Convert.ToDouble(OlineasR.Rows[i][7]);
        //                    OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Vehiculo").Value = OlineasR.Rows[i][9];
        //                    OrdenApp.ORequest.Lines.RequiredDate = Convert.ToDateTime(fecha);
        //                    OrdenApp.ORequest.Lines.TaxCode = "IVA";//OrdenApp.TableLines.Rows[i][4].ToString();
        //                                                            //OrdenApp.oOrder = 0.50;//System.Convert.ToDouble(OrdenApp.TableLines.Rows[i][5]);
        //                    OrdenApp.ORequest.Lines.WarehouseCode = "401";//OrdenApp.TableLines.Rows[i][6].ToString();
        //                    OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_Numero_Doc").Value = OlineasR.Rows[i][2];
        //                    OrdenApp.ORequest.Lines.UserFields.Fields.Item("U_TipoMantenimiento").Value = OlineasR.Rows[i][12];
        //                    i += 1;
        //                    if (i != OlineasR.Rows.Count)
        //                    {
        //                        OrdenApp.ORequest.Lines.Add();

        //                    }
        //                } while (!(i == OlineasR.Rows.Count));


        //                lRetCode = OrdenApp.ORequest.Add(); // Try to add the orer to the database

        //                OrdenApp.ORequest.Lines.Delete();

        //                OrdenApp.oCompany.Disconnect();


        //                //using (var progress = new ProgressBar())
        //                //{
        //                //    for (i = 0; i <= 100; i++)
        //                //    {
        //                //        progress.Report((double)i / 100);
        //                //        System.Threading.Thread.Sleep(20);
        //                //    }
        //                //}

        //                OlineasR.Clear();
        //                Console.WriteLine("Done.");

        //                Console.ForegroundColor = ConsoleColor.Green;
        //                Console.WriteLine("Solicitud de Compra Creada en SAP " + DocNumOrders);


        //                if (lRetCode != 0)
        //                {
        //                    int temp_int = lErrCode;
        //                    string temp_string = sErrMsg;
        //                    OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
        //                    if (lErrCode != -4006) // Incase adding an order failed
        //                    {

        //                        Console.ForegroundColor = ConsoleColor.Red;
        //                        Console.WriteLine(lErrCode + "Error " + temp_string);

        //                        string path = HttpContext.Current.Request.MapPath("~");
        //                        Log oLog = new Log("Logs/", path);
        //                        oLog.Add("Hola paso algo" + temp_string);


        //                        //MessageBox.Show(lErrCode + " " + sErrMsg); // Display error message
        //                    }
        //                    //else // If the currency Exchange is not set
        //                    //{
        //                    //    double dCur;
        //                    //    object sCur;
        //                    //    sCur = Interaction.InputBox("Currency Exchange - exchange rate has not been set for today. set the exchange rate", "Currency Exchange Setting", "", -1, -1);
        //                    //    if (Information.IsNumeric(sCur))
        //                    //    {
        //                    //        dCur = System.Convert.ToDouble(sCur);
        //                    //        oBob = (SAPbobsCOM.SBObob)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
        //                    //        //Update Currency rate
        //                    //        oBob.SetCurrencyRate(cmbCurrency.Text, DateAndTime.Today, dCur, false);
        //                    //    }
        //                    //    else
        //                    //    {
        //                    //        Interaction.MsgBox("Invalid Value to Currency Exchange", MsgBoxStyle.Exclamation, "Invalid Value");
        //                    //    }
        //                    //}

        //                }

        //                //else
        //                //{


        //                //    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
        //                //    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");
        //                // Environment.Exit(0);
        //                // DraftToCocument();
        //                //}


        //            }


        //        }

        //    }
        //    catch (Exception ex)
        //    {

        //        Console.WriteLine("Hay un problema :" + ex);

        //    }

        //    // Environment.Exit(0);

        //    DraftToCocument();


        //}






        public static void DraftToCocument() {

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("--------------------------------------------");
            Console.WriteLine("CONVIRTIENDO BORRADORES EN DOCUMENTOS REALES");
            Console.WriteLine("----------------LOGISTICA-----------------");
            Console.WriteLine("--------------------------------------------");
            DataTable tabla = new DataTable();
            using (SqlConnection conexion = new SqlConnection("Data Source=" + IpServerSql/*128.0.0.4*/+ ";Initial Catalog=" +/*DB_INTERFACE*/ServerSqlDBTran + ";Persist Security Info=True;User ID=" +/*sa-*/ServerSqlUser + ";Password=" + ServerSqlPass + ""))

           // using (SqlConnection conexion = new SqlConnection("Data Source=128.0.0.4;Initial Catalog=TRANSFLESA91;Persist Security Info=True;User ID=sa;Password=Ceo2015*"))
            {



                var DrftsAutorizadas = "select distinct T0.DocEntry, T0.DocDate, DocNum, T0.NumatCard,  T0.WddStatus from ODRF T0 where T0.DocStatus = 'O' and T0.ObjType = '23' and T0.WddStatus in ('Y') and (T0.DocDate >= CONVERT(nvarchar, DATEADD(day, - 20, GETDATE()))) ";
                conexion.Open();
                SqlDataAdapter adaptador = new SqlDataAdapter(DrftsAutorizadas, conexion);
                adaptador.Fill(tabla);
            }


            for (int i = 0; i < tabla.Rows.Count; i++)
            {
                ConectandoDB();

                DataRow rowdocentry = tabla.Rows[i];

                string NumAtCard = Convert.ToString(rowdocentry["NumatCard"]);

                int DocNum = Convert.ToInt32(rowdocentry["DocNum"]);

                int docentry = Convert.ToInt32(rowdocentry["DocEntry"]);


                DataRow rowfecha = tabla.Rows[i];

                DateTime Hoy = DateTime.Today;


                String docdate = Hoy.ToString("yyyy-MM-dd");  //Convert.ToDateTime(rowfecha["DocDate"]).ToString("yyyy-MM-dd");




                OrdenApp.oDraf = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);

                OrdenApp.oDraf.DocObjectCode = SAPbobsCOM.BoObjectTypes.oQuotations;


                OrdenApp.oDraf.GetByKey(docentry);
                //OrdenApp.oDraf.DocDate = Convert.ToDateTime(docdate);
                //OrdenApp.oDraf.DocDueDate = Convert.ToDateTime(docdate);//Convert.ToDateTime(docdate.ToString()); // pass due date "2018-05-21"
                //OrdenApp.oDraf.TaxDate = Convert.ToDateTime(docdate);
                int result = OrdenApp.oDraf.Update();
                if (result == 0)
                {


                    result = OrdenApp.oDraf.SaveDraftToDocument();
                    //Console.BackgroundColor = ConsoleColor.DarkBlue;
                    if (result == 0)
                    { 
                    Console.WriteLine("Solicitud Creada Con Exito !! "+ DocNum +" "+ NumAtCard);

                    //Console.WriteLine("Desconectado......");
                    //using (var progress = new ProgressBar())
                    //{
                    //    for (i = 0; i <= 100; i++)
                    //    {
                    //        progress.Report((double)i / 100);
                    //        System.Threading.Thread.Sleep(20);
                    //    }
                    //}
                    Console.WriteLine("ok.");

                        OrdenApp.oCompany.Disconnect();
                        Console.WriteLine("Desconectado......");
                    }
                    //OrdenApp.oCompany.Disconnect();
                    //Console.WriteLine("Desconectado......");


                    //else if (result != 0)
                    //{
                    //    int a;
                    //    string err = string.Empty;



                    //    OrdenApp.oCompany.GetLastError(out a, out err);

                    //    Console.BackgroundColor = ConsoleColor.Black;
                    //    Console.ForegroundColor = ConsoleColor.Red;
                    //    Console.WriteLine("Hay un problema respuesta: " + err + " " + a);
                    //    OrdenApp.oCompany.Disconnect();
                    //    Console.WriteLine("Desconectado......");
                    //    //AddPurchaseOrder();
                    //}

                }
                else if (result != 0)
                {
                    int a;
                    string err = string.Empty;



                    OrdenApp.oCompany.GetLastError(out a, out err);

                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Hay un problema respuesta  : " + err + " " + a + " Orden: "  + NumAtCard);
                    OrdenApp.oCompany.Disconnect();
                    Console.WriteLine("Desconectado......");
                    //AddPurchaseOrder();
                }







                //OrdenApp.oCompany.Disconnect();
                //Console.WriteLine("Desconectado......");

            }

            int l;
            l = 0;

            using (var progress = new ProgressBar())
            {
                for (l = 0; l <= 100; l++)
                {
                    progress.Report((double)l / 100);
                    System.Threading.Thread.Sleep(20);
                }
            }


            var info = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
            System.Diagnostics.Process.Start(info);

            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("------CERRANDO CARGA DE SOLICITUDES DE COMPRA-------");


            Environment.Exit(0);





        }


        //public static void AddDraftoOrderToDatabase()
        //{
        //    // Connect();

        //    // Init the Order object
        //    OrdenApp.oDraf = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts);
        //    OrdenApp.oRecordSet = (SAPbobsCOM.Recordset)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);



        //        SAPbobsCOM.SBObob oBob;

        //        string sqlStr = "select Docentry from ODRF " +
        //            "where DocStatus = 'O' " +
        //            "and ObjType = 17 " +
        //            "and WddStatus = 'W'";

        //        OrdenApp.oRecordSet.DoQuery(sqlStr);

        //        while (OrdenApp.oRecordSet.EoF)

        //            if (OrdenApp.oDraf.GetByKey(Convert.ToInt32(OrdenApp.oRecordSet.Fields.Item(0).Value)))
        //            {


        //                // Init the Order object
        //                OrdenApp.oOrder = (SAPbobsCOM.Documents)OrdenApp.oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);

        //                OrdenApp.oOrder.DocDate = OrdenApp.oDraf.DocDate;
        //                OrdenApp.oOrder.DocDueDate = OrdenApp.oDraf.DocDueDate;
        //                OrdenApp.oOrder.DiscountPercent = OrdenApp.oDraf.DiscountPercent;
        //                OrdenApp.oOrder.CardCode = OrdenApp.oDraf.CardCode;
        //                for (int i = 0; i < OrdenApp.oDraf.Lines.Count; i++)
        //                {

        //                    OrdenApp.oDraf.Lines.SetCurrentLine(i);
        //                    OrdenApp.oOrder.Lines.BaseType = OrdenApp.oDraf.Lines.BaseType;
        //                    OrdenApp.oOrder.Lines.BaseEntry = OrdenApp.oDraf.Lines.BaseEntry;
        //                    OrdenApp.oOrder.Lines.BaseLine = OrdenApp.oDraf.Lines.BaseLine;



        //                    if (i < OrdenApp.oDraf.Lines.Count - 1)
        //                    {

        //                        OrdenApp.oOrder.Lines.Add();

        //                       // Console.WriteLine("Orden Auorizada Creada");

        //                    }
        //                }


        //                lRetCode = OrdenApp.oOrder.Add(); // Try to add the orer to the database
        //                if (lRetCode != 0)
        //                {
        //                    int temp_int = lErrCode;
        //                    string temp_string = sErrMsg;
        //                    OrdenApp.oCompany.GetLastError(out temp_int, out temp_string);
        //                    if (lErrCode != -4006) // Incase adding an order failed
        //                    {
        //                        Console.WriteLine(lErrCode + "Error " + temp_string);
        //                        //MessageBox.Show(lErrCode + " " + sErrMsg); // Display error message
        //                    }

        //                }
        //                else
        //                {
        //                    Console.WriteLine("Orden Creada en base de datos");
        //                    //cmdInvoice.Enabled = true; // Enable the "Make Invoice On Order" button
        //                    //Interaction.MsgBox("Order Added to DataBase", MsgBoxStyle.Information, "Order Added");


        //                }


        //            }


        //}




        public static  DataTable ToDataTable<T>(this IList<T> data)
        {
            PropertyDescriptorCollection props =
            TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            for (int i = 0; i < props.Count; i++)
            {
                PropertyDescriptor prop = props[i];
                table.Columns.Add(prop.Name, prop.PropertyType);
            }
            object[] values = new object[props.Count];
            foreach (T item in data)
            {
                for (int i = 0; i < values.Length; i++)
                {
                    values[i] = props[i].GetValue(item);
                }
                table.Rows.Add(values);
            }

            DataRow[] currentRows = table.Select(
         null, null, DataViewRowState.CurrentRows);

            if (currentRows.Length < 1)
                Console.WriteLine("No Current Rows Found");
            else
            {
                foreach (DataColumn column in table.Columns)
                    Console.Write("\t{0}", column.ColumnName);

                Console.WriteLine("\tRowState");

                foreach (DataRow row in currentRows)
                {
                    foreach (DataColumn column in table.Columns)
                        Console.Write("\t{0}", row[column]);

                    Console.WriteLine("\t" + row.RowState);
                }

            }

                return table;

        }






        static string getmac(Int32 ts)
    {

    nonce = "ab" + ts + "cd"; ///Crear un texto unico por llamada, en este caso es el timestamp mas text, puede ser cualquier texto pero debe ser unico por llamada
    string mac = "hawk.1.header\n"; /// Parte de autorizacion de fracttal

            /*Inicia contruccion del string mac*/
            mac = mac + ts + "\n";
            mac = mac + nonce + "\n";
            mac = mac + "GET\n";////Metodo de llamada del API en este ejemplo se llama por GET pero también puede ser POST o PUT
            mac = mac + "https://app.fracttal.com/api/work_orders_movements/";////// URL del API en fractal
            mac = mac + "https://app.fracttal.com/api/"; /// Pide una direccion, este campo no se puede dejar en blanco pero no es verificado
            mac = mac + "443";///Puerto abierto de la direccion en la opción de arriba

            var keyByte = encoding.GetBytes(key);  ////Codificación de KeyByte para la Llave de fractal
    using (var hmacsha256 = new HMACSHA256(keyByte))
    {
        hmacsha256.ComputeHash(encoding.GetBytes(mac)); //// Uszndo la llame se encripta el mac

        return Base64Encode(ByteToString(hmacsha256.Hash)); ////El mac debe estar en Base64 y se devuelve como string
    }
    return "";


}





/***************Metodos utilizados*************************/

public static string Base64Encode(string plainText)
{
    var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
    return System.Convert.ToBase64String(plainTextBytes);
}

static string ByteToString(byte[] buff)
{
    string sbinary = "";
    for (int i = 0; i < buff.Length; i++)
        sbinary += buff[i].ToString("X2"); /* hex format */
    return sbinary;
 }

 }





}



