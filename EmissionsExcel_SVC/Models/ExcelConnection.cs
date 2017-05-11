using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace EmissionsExcel_SVC.Models {

   class ExcelConnection {
      
      public DataSet load(string sBook, string sSheet) {
         string sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + sBook + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
         
         try {
            if (File.Exists(sBook)) {
               OleDbDataAdapter oDataAdapter = new OleDbDataAdapter("select * from [" + sSheet + "$]", sConnectionString);
               DataSet oInformation = new DataSet();

               oDataAdapter.Fill(oInformation);
               oDataAdapter.Dispose();

               oDataAdapter = null;

               return oInformation;
            }
            else {
               return null;
            }
         }
         catch(Exception) {
            return null;
         }
      }
   }
}
