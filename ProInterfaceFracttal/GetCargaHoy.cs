using HawkNet;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;

namespace ProInterfaceFracttal
{
    class GetCargaHoy
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

        public static DataTable tabla2 = new DataTable();

        public static void GetCargaFracttal()
        {
            Console.ForegroundColor = ConsoleColor.Magenta;


            Console.WriteLine("[INICIA LA CARGAR DE OTS DE DIA DE HOY]");

            using (
               var progress = new ProgressBar())
            {
                for (int h = 0; h <= /*TableC.Rows.Count*/100; h++)
                {
                    progress.Report((double)h / 100);
                    System.Threading.Thread.Sleep(60);
                }
            }


            ////Setiamos el dia para cargar ordenes
            ///
            ///CAMBIAR DIA
            DateTime hoy = DateTime.Today;
            DateTime diaini = (hoy.AddDays(-1));
            DateTime diafin = (hoy.AddDays(0));


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
                else
                {

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
                    else
                    {

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
                                    }, false)).Distinct().ToList();


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
                    Console.WriteLine("BUSCANDO OPERACIONES DE MACIZO");


                    DataTable dtdfiltrado = new DataTable();
                    DataTable dtOrdenesfracttal = new DataTable();

                    if ((dtDistinct.Rows.Count) > 0)
                    {

                        //DataRow[] foundDT = dtDistinct.Select("parent_description like '%MACIZO%'");


                        /*PARA FILTRAR OPERACIONES DEPENDIENDO LA UNIDAD QUE LAS GENERA*/
                        DataRow[] foundDT = dtDistinct.Select("parent_description like '%LOGISTICA%'");



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
                    System.Threading.Thread.Sleep(60);
                }
            }


            //var info = new System.Diagnostics.ProcessStartInfo(Environment.GetCommandLineArgs()[0]);
            //System.Diagnostics.Process.Start(info);


            // LoopGetfracttal();

            // Environment.Exit(0);


            Program.GetCargaFracttal();

        }

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
                fila2["Department1"] = "14";
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

        public static DataTable InserOrdenesLineasFracttal(DataTable dt)
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
                    ",SUBSTRING(@U_note,1, 150)" + // Rocortado a 150 caracteres 16092021
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
                        command.Parameters.AddWithValue("@costs_center_description", dt.Rows[i][15].ToString().Trim());
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



        static public void Authenticate(IRestClient client, IRestRequest request)
        {
            var uri = client.BuildUri(request);
            var portSuffix = uri.Port != 80 ? ":" + uri.Port : "";
            var host = uri.Host + portSuffix;
            var method = request.Method.ToString();

            var header = Hawk.GetAuthorizationHeader(host, method, uri, _credential);

            request.AddHeader("Authorization", "Hawk " + header);
        }



    }
}
