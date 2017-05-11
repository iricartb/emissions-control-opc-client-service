using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using EmissionsExcel;
using System.IO;

namespace EmissionsExcel_SVC {

    static class Program {

       /* === Parámetros principales a modificar para parametrizar la aplicación en otra estación de trabajo === 
        *    
        *    - oTagsNames: Nombre de los tags a leer del servidor OPC. 
        *                  Total = (Nº Contaminantes * 9 Tags/Contaminante * Nº Líneas) + (2 {Tags de Arranque y Parada} * Nº Líneas)
        *                 
        *      Ejemplo: Nº Contaminantes: 18, Nº Líneas: 2  
        *               Total = (18 * 9 * 2) + (2 * 2) = 328 Tags 
        *               Resumen para 18 contaminantes (múltiplos de 164) = 164 * Nº Líneas
        *                 
        *    - nNumLines: Número de líneas de proceso en la planta. 
        *                 Máximo 3 líneas, si existen más líneas => Asignar mas memoria a los Arrays del fichero ClientOPC (ahora 512).
        *
        *      Ejemplo: Memoria asignada 512 posiciones (podemos almacenar 3 líneas) 
        *               164 * 3 = 492 < 512 => OK
        *               164 * 4 = 656 > 512 => NO_OK
        *    
        *    - sMethod: Política (REALTIME, BEST_EFFORT)
        *               Como tracta los datos leidos del servidor OPC.
        *
        *      REALTIME: Lee los datos del servidor OPC, en caso de error el minuto queda vacio
        *      BEST_EFFORT: Lee los datos del servidor OPC, en caso de error el minuto copia el minuto anterior
        *                 
        */
       public static string[] oTagsNames = new string[] { "StgL11_ValC#A", "StgL11_ValC#P",
                                                          "StgL11_CO_Raw", "StgL11_CO_EmmO2Val", "StgL11_CO_ValC#V", "StgL11_CO_ValC#C", "StgL11_CO_ValC#N", "StgL11_CO_QAL2RMin", "StgL11_CO_QAL2RMax", "StgL11_CO_QAL2A", "StgL11_CO_QAL2B", 
                                                          "StgL11_Dust_Raw", "StgL11_Dust_EmmO2Val", "StgL11_Dust_ValC#V", "StgL11_Dust_ValC#C", "StgL11_Dust_ValC#N", "StgL11_Dust_QAL2RMin", "StgL11_Dust_QAL2RMax", "StgL11_Dust_QAL2A", "StgL11_Dust_QAL2B",
                                                          "StgL11_TOC_Raw", "StgL11_TOC_EmmO2Val", "StgL11_TOC_ValC#V", "StgL11_TOC_ValC#C", "StgL11_TOC_ValC#N", "StgL11_TOC_QAL2RMin", "StgL11_TOC_QAL2RMax", "StgL11_TOC_QAL2A", "StgL11_TOC_QAL2B",
                                                          "StgL11_HCl_Raw", "StgL11_HCl_EmmO2Val", "StgL11_HCl_ValC#V", "StgL11_HCl_ValC#C", "StgL11_HCl_ValC#N", "StgL11_HCl_QAL2RMin", "StgL11_HCl_QAL2RMax", "StgL11_HCl_QAL2A", "StgL11_HCl_QAL2B",
                                                          "StgL11_HF_Raw", "StgL11_HF_EmmO2Val", "StgL11_HF_ValC#V", "StgL11_HF_ValC#C", "StgL11_HF_ValC#N", "StgL11_HF_QAL2RMin", "StgL11_HF_QAL2RMax", "StgL11_HF_QAL2A", "StgL11_HF_QAL2B",
                                                          "StgL11_SO2_Raw", "StgL11_SO2_EmmO2Val", "StgL11_SO2_ValC#V", "StgL11_SO2_ValC#C", "StgL11_SO2_ValC#N", "StgL11_SO2_QAL2RMin", "StgL11_SO2_QAL2RMax", "StgL11_SO2_QAL2A", "StgL11_SO2_QAL2B",
                                                          "StgL11_NOx_Raw", "StgL11_NOx_EmmO2Val", "StgL11_NOx_ValC#V", "StgL11_NOx_ValC#C", "StgL11_NOx_ValC#N", "StgL11_NOx_QAL2RMin", "StgL11_NOx_QAL2RMax", "StgL11_NOx_QAL2A", "StgL11_NOx_QAL2B",
                                                          "StgL11_NO2_Raw", "StgL11_NO2_EmmO2Val", "StgL11_NO2_ValC#V", "StgL11_NO2_ValC#C", "StgL11_NO2_ValC#N", "StgL11_NO2_QAL2RMin", "StgL11_NO2_QAL2RMax", "StgL11_NO2_QAL2A", "StgL11_NO2_QAL2B",
                                                          "StgL11_NO_Raw", "StgL11_NO_EmmO2Val", "StgL11_NO_ValC#V", "StgL11_NO_ValC#C", "StgL11_NO_ValC#N", "StgL11_NO_QAL2RMin", "StgL11_NO_QAL2RMax", "StgL11_NO_QAL2A", "StgL11_NO_QAL2B",
                                                          "StgL11_Hg_Raw", "StgL11_Hg_EmmO2Val", "StgL11_Hg_ValC#V", "StgL11_Hg_ValC#C", "StgL11_Hg_ValC#N", "StgL11_Hg_QAL2RMin", "StgL11_Hg_QAL2RMax", "StgL11_Hg_QAL2A", "StgL11_Hg_QAL2B",
                                                          "StgL11_H2O_Raw", "StgL11_H2O_EmmO2Val", "StgL11_H2O_ValC#V", "StgL11_H2O_ValC#C", "StgL11_H2O_ValC#N", "StgL11_H2O_QAL2RMin", "StgL11_H2O_QAL2RMax", "StgL11_H2O_QAL2A", "StgL11_H2O_QAL2B",
                                                          "StgL11_Press_Raw", "StgL11_Press_EmmO2Val", "StgL11_Press_ValC#V", "StgL11_Press_ValC#C", "StgL11_Press_ValC#N", "StgL11_Press_QAL2RMin", "StgL11_Press_QAL2RMax", "StgL11_Press_QAL2A", "StgL11_Press_QAL2B",
                                                          "StgL11_Temp_Raw", "StgL11_Temp_EmmO2Val", "StgL11_Temp_ValC#V", "StgL11_Temp_ValC#C", "StgL11_Temp_ValC#N", "StgL11_Temp_QAL2RMin", "StgL11_Temp_QAL2RMax", "StgL11_Temp_QAL2A", "StgL11_Temp_QAL2B",
                                                          "StgL11_O2_Raw", "StgL11_O2_EmmO2Val", "StgL11_O2_ValC#V", "StgL11_O2_ValC#C", "StgL11_O2_ValC#N", "StgL11_O2_QAL2RMin", "StgL11_O2_QAL2RMax", "StgL11_O2_QAL2A", "StgL11_O2_QAL2B",
                                                          "StgL11_CO2_Raw", "StgL11_CO2_EmmO2Val", "StgL11_CO2_ValC#V", "StgL11_CO2_ValC#C", "StgL11_CO2_ValC#N", "StgL11_CO2_QAL2RMin", "StgL11_CO2_QAL2RMax", "StgL11_CO2_QAL2A", "StgL11_CO2_QAL2B",
                                                          "StgL11_NH3_Raw", "StgL11_NH3_EmmO2Val", "StgL11_NH3_ValC#V", "StgL11_NH3_ValC#C", "StgL11_NH3_ValC#N", "StgL11_NH3_QAL2RMin", "StgL11_NH3_QAL2RMax", "StgL11_NH3_QAL2A", "StgL11_NH3_QAL2B",
                                                          "StgL11_Flow_Raw", "StgL11_Flow_EmmO2Val", "StgL11_Flow_ValC#V", "StgL11_Flow_ValC#C", "StgL11_Flow_ValC#N", "StgL11_Flow_QAL2RMin", "StgL11_Flow_QAL2RMax", "StgL11_Flow_QAL2A", "StgL11_Flow_QAL2B",
                                                          "StgL11_TempHogar_Raw", "StgL11_TempHogar_Raw", "TEMP_HOGAR_I_MIN", "TEMP_HOGAR_I_MIN", "TEMP_HOGAR_I_MIN", "StgL11_Temp_QAL2RMin", "StgL11_Temp_QAL2RMax", "StgL11_Temp_QAL2A", "StgL11_Temp_QAL2B",

                                                          "StgL21_ValC#A", "StgL21_ValC#P",
                                                          "StgL21_CO_Raw", "StgL21_CO_EmmO2Val", "StgL21_CO_ValC#V", "StgL21_CO_ValC#C", "StgL21_CO_ValC#N", "StgL21_CO_QAL2RMin", "StgL21_CO_QAL2RMax", "StgL21_CO_QAL2A", "StgL21_CO_QAL2B", 
                                                          "StgL21_Dust_Raw", "StgL21_Dust_EmmO2Val", "StgL21_Dust_ValC#V", "StgL21_Dust_ValC#C", "StgL21_Dust_ValC#N", "StgL21_Dust_QAL2RMin", "StgL21_Dust_QAL2RMax", "StgL21_Dust_QAL2A", "StgL21_Dust_QAL2B",
                                                          "StgL21_TOC_Raw", "StgL21_TOC_EmmO2Val", "StgL21_TOC_ValC#V", "StgL21_TOC_ValC#C", "StgL21_TOC_ValC#N", "StgL21_TOC_QAL2RMin", "StgL21_TOC_QAL2RMax", "StgL21_TOC_QAL2A", "StgL21_TOC_QAL2B",
                                                          "StgL21_HCl_Raw", "StgL21_HCl_EmmO2Val", "StgL21_HCl_ValC#V", "StgL21_HCl_ValC#C", "StgL21_HCl_ValC#N", "StgL21_HCl_QAL2RMin", "StgL21_HCl_QAL2RMax", "StgL21_HCl_QAL2A", "StgL21_HCl_QAL2B",
                                                          "StgL21_HF_Raw", "StgL21_HF_EmmO2Val", "StgL21_HF_ValC#V", "StgL21_HF_ValC#C", "StgL21_HF_ValC#N", "StgL21_HF_QAL2RMin", "StgL21_HF_QAL2RMax", "StgL21_HF_QAL2A", "StgL21_HF_QAL2B",
                                                          "StgL21_SO2_Raw", "StgL21_SO2_EmmO2Val", "StgL21_SO2_ValC#V", "StgL21_SO2_ValC#C", "StgL21_SO2_ValC#N", "StgL21_SO2_QAL2RMin", "StgL21_SO2_QAL2RMax", "StgL21_SO2_QAL2A", "StgL21_SO2_QAL2B",
                                                          "StgL21_NOx_Raw", "StgL21_NOx_EmmO2Val", "StgL21_NOx_ValC#V", "StgL21_NOx_ValC#C", "StgL21_NOx_ValC#N", "StgL21_NOx_QAL2RMin", "StgL21_NOx_QAL2RMax", "StgL21_NOx_QAL2A", "StgL21_NOx_QAL2B",
                                                          "StgL21_NO2_Raw", "StgL21_NO2_EmmO2Val", "StgL21_NO2_ValC#V", "StgL21_NO2_ValC#C", "StgL21_NO2_ValC#N", "StgL21_NO2_QAL2RMin", "StgL21_NO2_QAL2RMax", "StgL21_NO2_QAL2A", "StgL21_NO2_QAL2B",
                                                          "StgL21_NO_Raw", "StgL21_NO_EmmO2Val", "StgL21_NO_ValC#V", "StgL21_NO_ValC#C", "StgL21_NO_ValC#N", "StgL21_NO_QAL2RMin", "StgL21_NO_QAL2RMax", "StgL21_NO_QAL2A", "StgL21_NO_QAL2B",
                                                          "StgL21_Hg_Raw", "StgL21_Hg_EmmO2Val", "StgL21_Hg_ValC#V", "StgL21_Hg_ValC#C", "StgL21_Hg_ValC#N", "StgL21_Hg_QAL2RMin", "StgL21_Hg_QAL2RMax", "StgL21_Hg_QAL2A", "StgL21_Hg_QAL2B",
                                                          "StgL21_H2O_Raw", "StgL21_H2O_EmmO2Val", "StgL21_H2O_ValC#V", "StgL21_H2O_ValC#C", "StgL21_H2O_ValC#N", "StgL21_H2O_QAL2RMin", "StgL21_H2O_QAL2RMax", "StgL21_H2O_QAL2A", "StgL21_H2O_QAL2B",
                                                          "StgL21_Press_Raw", "StgL21_Press_EmmO2Val", "StgL21_Press_ValC#V", "StgL21_Press_ValC#C", "StgL21_Press_ValC#N", "StgL21_Press_QAL2RMin", "StgL21_Press_QAL2RMax", "StgL21_Press_QAL2A", "StgL21_Press_QAL2B",
                                                          "StgL21_Temp_Raw", "StgL21_Temp_EmmO2Val", "StgL21_Temp_ValC#V", "StgL21_Temp_ValC#C", "StgL21_Temp_ValC#N", "StgL21_Temp_QAL2RMin", "StgL21_Temp_QAL2RMax", "StgL21_Temp_QAL2A", "StgL21_Temp_QAL2B",
                                                          "StgL21_O2_Raw", "StgL21_O2_EmmO2Val", "StgL21_O2_ValC#V", "StgL21_O2_ValC#C", "StgL21_O2_ValC#N", "StgL21_O2_QAL2RMin", "StgL21_O2_QAL2RMax", "StgL21_O2_QAL2A", "StgL21_O2_QAL2B",
                                                          "StgL21_CO2_Raw", "StgL21_CO2_EmmO2Val", "StgL21_CO2_ValC#V", "StgL21_CO2_ValC#C", "StgL21_CO2_ValC#N", "StgL21_CO2_QAL2RMin", "StgL21_CO2_QAL2RMax", "StgL21_CO2_QAL2A", "StgL21_CO2_QAL2B",
                                                          "StgL21_NH3_Raw", "StgL21_NH3_EmmO2Val", "StgL21_NH3_ValC#V", "StgL21_NH3_ValC#C", "StgL21_NH3_ValC#N", "StgL21_NH3_QAL2RMin", "StgL21_NH3_QAL2RMax", "StgL21_NH3_QAL2A", "StgL21_NH3_QAL2B",
                                                          "StgL21_Flow_Raw", "StgL21_Flow_EmmO2Val", "StgL21_Flow_ValC#V", "StgL21_Flow_ValC#C", "StgL21_Flow_ValC#N", "StgL21_Flow_QAL2RMin", "StgL21_Flow_QAL2RMax", "StgL21_Flow_QAL2A", "StgL21_Flow_QAL2B",
                                                          "StgL21_TempHogar_Raw", "StgL21_TempHogar_Raw", "TEMP_HOGAR_II_MIN", "TEMP_HOGAR_II_MIN", "TEMP_HOGAR_II_MIN", "StgL21_Temp_QAL2RMin", "StgL21_Temp_QAL2RMax", "StgL21_Temp_QAL2A", "StgL21_Temp_QAL2B",
      }; 

      public static int nNumLines = 2;
      public static string sMethod = "BEST_EFFORT";
      /* === Fin parámetros principales a modificar === */

      public const string FOLDER_DESTINY = "C:\\xeac\\excels\\online\\";
      public const string FOLDER_DESTINY_COPY = "C:\\xeac\\excels\\copy\\";
      public const string FOLDER_DESTINY_XEAC = "C:\\xeac\\ftp\\online\\";
      public const string FOLDER_DESTINY_XEAC_COPY = "C:\\xeac\\ftp\\copy\\";
      public const string FOLDER_DESTINY_XEAC_REPORTS = "C:\\xeac\\ftp\\reports\\";
      public const string FOLDER_FILE_LOG_HISTORY = "C:\\xeac\\logs\\";
      public const string FOLDER_FILE_LOG = "C:\\xeac\\";
      public const string FOLDER_TEMP = "C:\\xeac\\temp\\";

      public const string FILE_EXCEL_TOKEN_NAME = "Stg";

      public static int nNumElements = 18;
      public static int nTimeoutOPCRead = 6000;
      public static bool bShowLog = true;
      public static bool bShowVerboseLog = true;

      public static int nLastMinute = -1;
      public static ClientOPC oClientOPC;
        
      public static ElementValue[] oBeforeElementsValues = new ElementValue[] { 
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1"),
         new ElementValue("0", "0", "V", "0", "100", "0", "1")
      };

      public static StreamWriter oLogFile = null;

      /// <summary>
      /// Punto de entrada principal para la aplicación.
      /// </summary>
      static void Main() {
         ServiceBase[] ServicesToRun;
         ServicesToRun = new ServiceBase[] { 
		      new ServiceEmissionsExcel() 
	      };
         ServiceBase.Run(ServicesToRun);
      }

      public static void writeLineInLogFile(TextWriter oLogWriter, string sLine, bool bInsertNewLine) {
         try {
            if (oLogWriter != null) {
               DateTime oDateNowCurrent = DateTime.Now;

               string sNewLine;
               if (bInsertNewLine) sNewLine = "\r\n";
               else sNewLine = "";

               oLogWriter.WriteLine(sNewLine + oDateNowCurrent.Day.ToString().PadLeft(2, '0') + "/" + oDateNowCurrent.Month.ToString().PadLeft(2, '0') + "/" + oDateNowCurrent.Year.ToString() + " " + oDateNowCurrent.Hour.ToString().PadLeft(2, '0') + "h " + oDateNowCurrent.Minute.ToString().PadLeft(2, '0') + "m: " + sLine);     
               oLogWriter.Flush();
            }
         } catch(Exception) { }
      }
   }
}
