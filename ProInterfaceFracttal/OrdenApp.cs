using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProInterfaceFracttal
{
    class OrdenApp
    {
        public static DataTable TableLines; // the datasource of the datagrid called DataLines.
        public static SAPbobsCOM.Documents oOrder; // Order object
        public static SAPbobsCOM.Documents ORequest; //solicitud de compra
        public static SAPbobsCOM.Documents oDraf; // Documentos preliminares
        public static SAPbobsCOM.Documents oInvoice; // Invoice Object
        public static SAPbobsCOM.Documents oBatches; // Lotes 
        public static SAPbobsCOM.Recordset oRecordSet; // A recordset object
        public static SAPbobsCOM.Company oCompany; // The company object

    }
}
