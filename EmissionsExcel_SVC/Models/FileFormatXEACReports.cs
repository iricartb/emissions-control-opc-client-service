using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;

namespace EmissionsExcel_SVC.Models {

   class FileFormatXEACReports {
      public const string NOT_EMISSION_VALUE = "9999";
      public const string NOT_EMISSION_CHARACTER = "N";

      private static string[] oCodes = new string[] { "109", "110", "112", "130", "142", "212", "220", "221", "223", 
                                               "234", "241", "252", "261", "262", "270", "350", "280", "362", 
                                               "902", "903", "904"
      };

      private static string[] oCleans = new string[] { "112", "142", "212", "220", "234", "252", "261", "270", "350", "280", "362" };
      private static string[] oPollutants = new string[] { "Press", "Temp", "Polvo", "H2O", "Hg", "SO2", "NOx", "NO", "NH3", "HCl", "NO2", "O2", "CO", "CO2", "TOC", "Flow", "HF", "TempHogar" };
   
      private string sFocusClean;
      private string sFocusDirty;

      public bool convertExcelFile(string sExcelFile, int nTotalMinutes) {
         bool bLoadExcelFile = false;
         
         try {
            if (Directory.Exists(Program.FOLDER_TEMP)) {
               string[] oDestinyFiles = getDestinyFiles(sExcelFile);

               string sUniqueExcelPathTmp = Program.FOLDER_TEMP + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xls";
               File.Copy(Program.FOLDER_DESTINY + sExcelFile, sUniqueExcelPathTmp);

               if (File.Exists(sUniqueExcelPathTmp)) {
                  if ((oDestinyFiles[0] != null) && (oDestinyFiles[1] != null)) {
                     if (createFileFormatXEAC(sUniqueExcelPathTmp, oDestinyFiles, nTotalMinutes)) {
                        bLoadExcelFile = true;   
                     }

                     if (File.Exists(sUniqueExcelPathTmp)) File.Delete(sUniqueExcelPathTmp); 
                  }
                  else {
                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} No se han podido devolver los directorios de destino de la XEAC", false);
                  }
               }
               else {
                  if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} No se ha podido copiar el Fichero Excel '" + Program.FOLDER_DESTINY + sExcelFile + "' a la carpeta temporal", false);
               }

               oDestinyFiles = null;
            }
         } catch(Exception e) {
            if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error durante la carga del Fichero Excel para proceder con su conversión a Fichero XEAC: " + e.ToString(), false);
            bLoadExcelFile = false;
         }

         return bLoadExcelFile;
      }

      private string[] getDestinyFiles(string sExcelFile) {
         try {
            if ((File.Exists(Program.FOLDER_DESTINY + sExcelFile)) && (sExcelFile.Length > 25)) {
               string sYear = sExcelFile.Substring(16, 2);
               string sMonth = sExcelFile.Substring(18, 2);
               string sDay = sExcelFile.Substring(20, 2);
               string sHour = sExcelFile.Substring(23, 2);

               string sStartDate = sDay + "/" + sMonth + "/" + sYear;
               int nLine = Int32.Parse(sExcelFile.Substring(9 + Program.FILE_EXCEL_TOKEN_NAME.Length, 1));

               sFocusClean = "F" + Convert.ToChar(66 + nLine);
               sFocusDirty = "J" + Convert.ToChar(66 + nLine);
            
               string sNameClean = sFocusClean + sDay + sMonth + sYear + ".M" + sHour;
               string sNameDirty = sFocusDirty + sDay + sMonth + sYear + ".M" + sHour;

               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusDirty);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean + "\\" + sYear)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean + "\\" + sYear);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusDirty + "\\" + sYear);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean + "\\" + sYear + "\\" + sMonth)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean + "\\" + sYear + "\\" + sMonth);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusDirty + "\\" + sYear + "\\" + sMonth);
               }

               return new string[] { Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusClean + "\\" + sYear + "\\" + sMonth + "\\" + sNameClean, Program.FOLDER_DESTINY_XEAC_REPORTS + sFocusDirty + "\\" + sYear + "\\" + sMonth + "\\" + sNameDirty, sFocusClean, sFocusDirty, sYear, sMonth, sNameClean, sNameDirty }; 
            }
            else {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error al devolver los directorios de destino de la XEAC: No existe el Fichero " + Program.FOLDER_DESTINY + sExcelFile + " o la longitud del nombre es menor a 25", false);

               return new string[] { null, null, null, null, null, null, null, null };
            }
         } catch (Exception e) {
            if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error al devolver los directorios de destino de la XEAC: " + e.ToString(), false);

            return new string[] { null, null, null, null, null, null, null, null };
         }
      }

      private bool createFileFormatXEAC(string sExcelFilePath, string[] oDestinyFiles, int nTotalMinutes) {
         bool bCreateFileFormatXEAC = false;

         StreamWriter oCleanStreamWriter = null;
         StreamWriter oDirtyStreamWriter = null;
         DataSet oSheetDataSet = null;
         ExcelConnection oExcelConnection = null;

         try {
            if (oDestinyFiles.Length == 8) {
               oExcelConnection = new ExcelConnection();
               oCleanStreamWriter = new StreamWriter(oDestinyFiles[0]);
               oDirtyStreamWriter = new StreamWriter(oDestinyFiles[1]);

               for(int nSheet = 0; nSheet < Program.nNumElements; nSheet++) {
                  int nCurrentMinute = 0;

                  oDirtyStreamWriter.Write(FileFormatXEACReports.oCodes[nSheet]);
                  oCleanStreamWriter.Write(FileFormatXEACReports.oCodes[nSheet]); 

                  oSheetDataSet = oExcelConnection.load(sExcelFilePath, FileFormatXEACReports.oPollutants[nSheet]);
                  if (oSheetDataSet != null) {
                     if (oSheetDataSet.Tables[0].Columns.Count == 8) {
                        string sCleanValue = "";
                        string sDirtyValue = "";
                        string sTypeValue = "";

                        for(;nCurrentMinute < nTotalMinutes;) {
                           if ((4 + nCurrentMinute) < oSheetDataSet.Tables[0].Rows.Count) {
                              DataRow oRow = oSheetDataSet.Tables[0].Rows[4 + nCurrentMinute];

                              if (oRow[2].ToString().Trim().Length > 0) sCleanValue = getDigitFormat(oRow[2].ToString().Trim());
                              else sCleanValue = FileFormatXEACReports.NOT_EMISSION_VALUE;
                              
                              if (oRow[1].ToString().Trim().Length > 0) sDirtyValue = getDigitFormat(oRow[1].ToString().Trim());
                              else sDirtyValue = FileFormatXEACReports.NOT_EMISSION_VALUE;

                              if (oRow[3].ToString().Trim().Length > 0) sTypeValue = getEmissionTypeValue(oRow[3].ToString().Trim());
                              else sTypeValue = FileFormatXEACReports.NOT_EMISSION_CHARACTER;

                              if (sTypeValue == "S") {
                                 if (FileFormatXEACReports.oPollutants[nSheet] == "H2O") {
                                    if (sCleanValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sCleanValue = "0015";
                                    if (sDirtyValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sDirtyValue = "0015";
                                 }
                                 else if (FileFormatXEACReports.oPollutants[nSheet] == "O2") {
                                    if (sCleanValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sCleanValue = "0011";
                                    if (sDirtyValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sDirtyValue = "0011";
                                 }
                                 else if (FileFormatXEACReports.oPollutants[nSheet] == "Temp") {
                                    if (sCleanValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sCleanValue = "0132";
                                    if (sDirtyValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sDirtyValue = "0132";
                                 }
                                 else if (FileFormatXEACReports.oPollutants[nSheet] == "Press") {
                                    if (sCleanValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sCleanValue = "1013";
                                    if (sDirtyValue != FileFormatXEACReports.NOT_EMISSION_VALUE) sDirtyValue = "1013";
                                 }
                              }

                              if ((oRow[1].ToString().Trim().Length == 0) || (oRow[2].ToString().Trim().Length == 0)) {
                                 if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEACReports.oPollutants[nSheet] + "' del minuto " + nCurrentMinute + ", se pondra el dato de 'SIN MEDIDA' -> " + FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER, false);   
                              }

                              oDirtyStreamWriter.Write(sDirtyValue + sTypeValue);
                              oCleanStreamWriter.Write(sCleanValue + sTypeValue);
   
                              oRow = null;
                           }
                           else {
                              sCleanValue = FileFormatXEACReports.NOT_EMISSION_VALUE;
                              sDirtyValue = FileFormatXEACReports.NOT_EMISSION_VALUE;
                              sTypeValue = FileFormatXEACReports.NOT_EMISSION_CHARACTER;

                              if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEACReports.oPollutants[nSheet] + "' del minuto " + nCurrentMinute + ", se pondra el dato de 'SIN MEDIDA' -> " + FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER, false);
                              
                              oDirtyStreamWriter.Write(sDirtyValue + sTypeValue);
                              oCleanStreamWriter.Write(sCleanValue + sTypeValue);
                           }

                           nCurrentMinute++;
                        }
                     }
                     else {
                        if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido las 8 columnas del Fichero Excel sobre el contaminante '" + FileFormatXEACReports.oPollutants[nSheet] + "', se pondran datos de 'SIN MEDIDA' -> " + FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER + " en todos los minutos (" + nTotalMinutes + ")", false);   
                  
                        for(nCurrentMinute = 0; nCurrentMinute < nTotalMinutes; nCurrentMinute++) {
                           oDirtyStreamWriter.Write(FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER);

                           oCleanStreamWriter.Write(FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER); 
                        }
                     }

                     oSheetDataSet.Tables.Clear();
                     oSheetDataSet.Dispose();
                     oSheetDataSet = null;
                  }
                  else {
                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEACReports.oPollutants[nSheet] + "', se pondran datos de 'SIN MEDIDA' -> " + FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER + " en todos los minutos (" + nTotalMinutes + ")", false);   
                  
                     for(nCurrentMinute = 0; nCurrentMinute < nTotalMinutes; nCurrentMinute++) {
                        oDirtyStreamWriter.Write(FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER);

                        oCleanStreamWriter.Write(FileFormatXEACReports.NOT_EMISSION_VALUE + FileFormatXEACReports.NOT_EMISSION_CHARACTER); 
                     } 
                  }

                  if (nSheet != 17) {
                     oDirtyStreamWriter.Write(Environment.NewLine);
                     oCleanStreamWriter.Write(Environment.NewLine); 
                  }
               }

               oCleanStreamWriter.Flush();
               oDirtyStreamWriter.Flush();

               oCleanStreamWriter.Close();
               oDirtyStreamWriter.Close();

               oSheetDataSet = null;
               bCreateFileFormatXEAC = true;
            }
         } catch (Exception e) {
            if (oCleanStreamWriter != null) oCleanStreamWriter.Close();
            if (oDirtyStreamWriter != null) oDirtyStreamWriter.Close();

            if (File.Exists(oDestinyFiles[0])) File.Delete(oDestinyFiles[0]);
            if (File.Exists(oDestinyFiles[1])) File.Delete(oDestinyFiles[1]);

            if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error durante la transformación del Fichero Excel al formato del Fichero XEAC: " + e.ToString(), false); 
         }

         oCleanStreamWriter = null;
         oDirtyStreamWriter = null;
         oSheetDataSet = null;
         oExcelConnection = null;

         return bCreateFileFormatXEAC;
      }

      private string getEmissionTypeValue(string sEmissionTypeValue) {
         if (sEmissionTypeValue.Trim().Length > 0) return sEmissionTypeValue.Trim().Substring(sEmissionTypeValue.Trim().Length - 1, 1);
         else return "V";
      }

      private string getDigitFormat(string sValue) {
         Double nValueDouble;

         try {
            if (sValue.Length > 4) {
               if (sValue.Contains(".")) {
                  if ((sValue.Substring(0, sValue.LastIndexOf(".")).Length == 3) || (sValue.Substring(0, sValue.LastIndexOf(".")).Length == 4)) {
                     sValue = sValue.Substring(0, sValue.LastIndexOf("."));
                     nValueDouble = Double.Parse(sValue);
                     sValue = Math.Round(nValueDouble, 0).ToString();
                  }
                  else {
                     if (sValue.Substring(0, sValue.LastIndexOf(".")).Length == 2) {
                        nValueDouble = Double.Parse(sValue);
                        sValue = Math.Round(nValueDouble, 1).ToString();
                     }
                     else if (sValue.Substring(0, sValue.LastIndexOf(".")).Length == 1) {
                        nValueDouble = Double.Parse(sValue);
                        sValue = Math.Round(nValueDouble, 2).ToString();
                     }
                     else {
                        nValueDouble = Double.Parse(sValue);
                        nValueDouble = nValueDouble / 1000;
                        sValue = getDigitFormat(nValueDouble.ToString());
                     }
                  }
               }
               else {
                  nValueDouble = Double.Parse(sValue);
                  nValueDouble = nValueDouble / 1000;
                  sValue = getDigitFormat(nValueDouble.ToString());
               }
            }
         
            if (sValue.Length < 4) {
               if (sValue.Contains(".")) {
                  if (sValue.Length == 3) sValue = sValue + "0"; 
               }
               else {
                  if (sValue.Length == 1) sValue = sValue + ".00";
                  else if (sValue.Length == 2) sValue = sValue + ".0";
                  else if (sValue.Length == 3) {
                     if (sValue.StartsWith("-")) {
                        sValue = sValue.Replace("-", "0");
                        sValue = "-" + sValue;
                     }
                     else {
                        sValue = "0" + sValue;
                     }
                  }
               }
            }
         } catch(Exception) { sValue = "0.00"; }

         return sValue;
      }
   }
}
