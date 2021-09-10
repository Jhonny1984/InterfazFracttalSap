using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ProInterfaceFracttal
{
   public static class TestBatches
    {

        static void main()
        {
            ConectandoDB();


            BatchUpdate();



        }


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
                    Console.WriteLine("conectando");
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
                    Console.WriteLine("Conectado");
                    //this.Text = this.Text + ": Connected";
                    // Remove the following 2 remark lines if you want to try to connect automatically
                    //Else
                    //Connect();
                }
                //CreateLinesTable();

            }

        //Error handling variables
        public static string sErrMsg;
        public static int lErrCode;
        public static int lRetCode;


        public static void Connect()
        {
            //Cursor = System.Windows.Forms.Cursors.WaitCursor; //Change mouse cursor

            // Set connection properties

            OrdenApp.oCompany.CompanyDB = "PRUEBAS_INVERFACSA91".ToString()/*cmbCompany.Text*/;
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
                Console.WriteLine("Error al Conectar: " + temp_string);
                Console.WriteLine("Problemas en la conexion presiones una tecla");
                Console.ReadKey();

                // Interaction.MsgBox("Connection Failed - " + sErrMsg, MsgBoxStyle.Exclamation, "Default Connection Failed");
            }
            if (OrdenApp.oCompany.Connected) // if connected
            {
                Console.WriteLine("Conectando...: " + OrdenApp.oCompany.CompanyDB);
                //this.Text = this.Text + " - Connected to " + OrdenApp.oCompany.CompanyDB;
                //grpConn.Enabled = false;
                //grpOrder.Enabled = true;
                //LoadGui(); // Load data for UI elements like combo boxes

                Console.WriteLine("Conectado !!! " + OrdenApp.oCompany.UserSignature.ToString());




            }
            //AddOrderToDatabase();
            //Cursor = System.Windows.Forms.Cursors.Default; //Change mouse cursor
        }



        public static void BatchUpdate()

        {

            SAPbobsCOM.CompanyService oCompanyService;

            oCompanyService = OrdenApp.oCompany.GetCompanyService();

            SAPbobsCOM.BatchNumberDetailsService oBatchNumbersService;

            oBatchNumbersService = oCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.BatchNumberDetailsService);

            SAPbobsCOM.BatchNumberDetailParams oBatchNumberDetailParams;

            oBatchNumberDetailParams = oBatchNumbersService.GetDataInterface(SAPbobsCOM.BatchNumberDetailsServiceDataInterfaces.bndsBatchNumberDetailParams);



            try

            {

                int docentry = 6157; //Put here the actual AbsEntry of the Batchnumber record from table OBTN

                oBatchNumberDetailParams.DocEntry = docentry;

                SAPbobsCOM.BatchNumberDetail oBatchNumberDetail;

                oBatchNumberDetail = oBatchNumbersService.Get(oBatchNumberDetailParams);

                oBatchNumberDetail.BatchAttribute1 = "Ok ya esta ";

                oBatchNumberDetail.Status = BoDefaultBatchStatus.dbs_Released; //put here the status you want

                //oBatchNumberDetail.UserFields.Item(“U_PalletID”).Value = “45109 - 07”;

                oBatchNumbersService.Update(oBatchNumberDetail);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBatchNumberDetail);



                Console.WriteLine("Lote Liberado :)");

            }

            catch (Exception e)

            {


                Console.WriteLine("Error al Liberar el Lote "+e);
                //gApp.MessageBox(“An unknown error occurred in method BatchUpdate.\n\n” +e.ToString(), 1, “Ok”, “”, “”);

            }

            finally

            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBatchNumbersService);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBatchNumberDetailParams);



            }

        }

    }
  
}
