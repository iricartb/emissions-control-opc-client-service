using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;

namespace EmissionsExcel_SVC.Models {

   class FileFormatXEAC {
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

               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC + sFocusClean)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusClean);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusDirty);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC + sFocusClean + "\\" + sYear)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusClean + "\\" + sYear);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusDirty + "\\" + sYear);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC + sFocusClean + "\\" + sYear + "\\" + sMonth)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusClean + "\\" + sYear + "\\" + sMonth);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC + sFocusDirty + "\\" + sYear + "\\" + sMonth);
               }

               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusDirty);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean + "\\" + sYear)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean + "\\" + sYear);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusDirty + "\\" + sYear);
               }
               if (!Directory.Exists(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean + "\\" + sYear + "\\" + sMonth)) {
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusClean + "\\" + sYear + "\\" + sMonth);
                  Directory.CreateDirectory(Program.FOLDER_DESTINY_XEAC_COPY + sFocusDirty + "\\" + sYear + "\\" + sMonth);
               }

               return new string[] { Program.FOLDER_DESTINY_XEAC + sFocusClean + "\\" + sYear + "\\" + sMonth + "\\" + sNameClean, Program.FOLDER_DESTINY_XEAC + sFocusDirty + "\\" + sYear + "\\" + sMonth + "\\" + sNameDirty, sFocusClean, sFocusDirty, sYear, sMonth, sNameClean, sNameDirty }; 
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
         string[] oPendings = new string[] { FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER };
         string[] oIndependents = new string[] { FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER };
         string[] oRanges = new string[] { FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER };
         bool bFinishDocuments = false;

         try {
            if (oDestinyFiles.Length == 8) {
               oExcelConnection = new ExcelConnection();
               oCleanStreamWriter = new StreamWriter(oDestinyFiles[0]);
               oDirtyStreamWriter = new StreamWriter(oDestinyFiles[1]);

               for(int nSheet = 0; nSheet < Program.nNumElements; nSheet++) {
                  int nCurrentMinute = 0;

                  oDirtyStreamWriter.Write(FileFormatXEAC.oCodes[nSheet]);
                  if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(FileFormatXEAC.oCodes[nSheet]); 

                  oSheetDataSet = oExcelConnection.load(sExcelFilePath, FileFormatXEAC.oPollutants[nSheet]);
                  if (oSheetDataSet != null) {
                     if (oSheetDataSet.Tables[0].Columns.Count == 8) {
                        string sCleanValue = "";
                        string sDirtyValue = "";
                        string sTypeValue = "";

                        for(;nCurrentMinute < nTotalMinutes;) {
                           if ((4 + nCurrentMinute) < oSheetDataSet.Tables[0].Rows.Count) {
                              DataRow oRow = oSheetDataSet.Tables[0].Rows[4 + nCurrentMinute];

                              if (oRow[2].ToString().Trim().Length > 0) sCleanValue = getDigitFormat(oRow[2].ToString().Trim());
                              else sCleanValue = FileFormatXEAC.NOT_EMISSION_VALUE;
                              
                              if (oRow[1].ToString().Trim().Length > 0) sDirtyValue = getDigitFormat(oRow[1].ToString().Trim());
                              else sDirtyValue = FileFormatXEAC.NOT_EMISSION_VALUE;

                              if (oRow[3].ToString().Trim().Length > 0) sTypeValue = getEmissionTypeValue(oRow[3].ToString().Trim());
                              else sTypeValue = FileFormatXEAC.NOT_EMISSION_CHARACTER;

                              if (sTypeValue == "S") {
                                 if (FileFormatXEAC.oPollutants[nSheet] == "H2O") {
                                    if (sCleanValue != FileFormatXEAC.NOT_EMISSION_VALUE) sCleanValue = "0015";
                                    if (sDirtyValue != FileFormatXEAC.NOT_EMISSION_VALUE) sDirtyValue = "0015";
                                 }
                                 else if (FileFormatXEAC.oPollutants[nSheet] == "O2") {
                                    if (sCleanValue != FileFormatXEAC.NOT_EMISSION_VALUE) sCleanValue = "0011";
                                    if (sDirtyValue != FileFormatXEAC.NOT_EMISSION_VALUE) sDirtyValue = "0011";
                                 }
                                 else if (FileFormatXEAC.oPollutants[nSheet] == "Temp") {
                                    if (sCleanValue != FileFormatXEAC.NOT_EMISSION_VALUE) sCleanValue = "0132";
                                    if (sDirtyValue != FileFormatXEAC.NOT_EMISSION_VALUE) sDirtyValue = "0132";
                                 }
                                 else if (FileFormatXEAC.oPollutants[nSheet] == "Press") {
                                    if (sCleanValue != FileFormatXEAC.NOT_EMISSION_VALUE) sCleanValue = "1013";
                                    if (sDirtyValue != FileFormatXEAC.NOT_EMISSION_VALUE) sDirtyValue = "1013";
                                 }
                              }

                              if ((oRow[1].ToString().Trim().Length == 0) || (oRow[2].ToString().Trim().Length == 0)) {
                                 if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEAC.oPollutants[nSheet] + "' del minuto " + nCurrentMinute + ", se pondra el dato de 'SIN MEDIDA' -> " + FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, false);   
                              }

                              oDirtyStreamWriter.Write(sDirtyValue + sTypeValue);
                              if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(sCleanValue + sTypeValue); 

                              if (nCurrentMinute == (nTotalMinutes - 1)) {
                                 if (!((nSheet == 0) || (nSheet == 1) || (nSheet == 7) || (nSheet == 8) || (nSheet == 10) || (nSheet == 11) || (nSheet == 12) || (nSheet == 15))) {
                                    if (oRow[7].ToString().Trim().Length > 0) oPendings[nSheet] = getDigitFormat(oRow[7].ToString().Trim()) + sTypeValue;
                                    else oPendings[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 

                                    if (oRow[6].ToString().Trim().Length > 0) oIndependents[nSheet] = getDigitFormat(oRow[6].ToString().Trim()) + sTypeValue;
                                    else oIndependents[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 

                                    if (oRow[5].ToString().Trim().Length > 0) oRanges[nSheet] = getDigitFormat(oRow[5].ToString().Trim()) + sTypeValue;
                                    else oRanges[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 
                                 }
                                 else {
                                    if (oRow[7].ToString().Trim().Length > 0) oPendings[nSheet] = "1.00" + sTypeValue;
                                    else oPendings[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;

                                    if (oRow[6].ToString().Trim().Length > 0) oIndependents[nSheet] = "0.00" + sTypeValue;
                                    else oIndependents[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;

                                    if (oRow[5].ToString().Trim().Length > 0) oRanges[nSheet] = "0.00" + sTypeValue;
                                    else oRanges[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 
                                 }
                              }
                              
                              if (nCurrentMinute == 59) bFinishDocuments = true;
    
                              oRow = null;
                           }
                           else {
                              sCleanValue = FileFormatXEAC.NOT_EMISSION_VALUE;
                              sDirtyValue = FileFormatXEAC.NOT_EMISSION_VALUE;
                              sTypeValue = FileFormatXEAC.NOT_EMISSION_CHARACTER;

                              if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEAC.oPollutants[nSheet] + "' del minuto " + nCurrentMinute + ", se pondra el dato de 'SIN MEDIDA' -> " + FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER, false);
                              
                              oDirtyStreamWriter.Write(sDirtyValue + sTypeValue);
                              if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(sCleanValue + sTypeValue);

                              if (nCurrentMinute == (nTotalMinutes - 1)) {
                                 oPendings[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                                 oIndependents[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                                 oRanges[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 
                              }

                              if (nCurrentMinute == 59) bFinishDocuments = true;
                           }

                           nCurrentMinute++;
                        }
                     }
                     else {
                        if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido las 8 columnas del Fichero Excel sobre el contaminante '" + FileFormatXEAC.oPollutants[nSheet] + "', se pondran datos de 'SIN MEDIDA' -> " + FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER + " en todos los minutos (" + nTotalMinutes + ")", false);   
                  
                        for(nCurrentMinute = 0; nCurrentMinute < nTotalMinutes; nCurrentMinute++) {
                           oDirtyStreamWriter.Write(FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER);

                           if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER); 
                        }

                        oPendings[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                        oIndependents[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                        oRanges[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 

                        if (nTotalMinutes == 60) bFinishDocuments = true;
                     }

                     oSheetDataSet.Tables.Clear();
                     oSheetDataSet.Dispose();
                     oSheetDataSet = null;
                  }
                  else {
                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han obtenido datos del Fichero Excel sobre el contaminante '" + FileFormatXEAC.oPollutants[nSheet] + "', se pondran datos de 'SIN MEDIDA' -> " + FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER + " en todos los minutos (" + nTotalMinutes + ")", false);   
                  
                     for(nCurrentMinute = 0; nCurrentMinute < nTotalMinutes; nCurrentMinute++) {
                        oDirtyStreamWriter.Write(FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER);

                        if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER); 
                     }

                     oPendings[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                     oIndependents[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER;
                     oRanges[nSheet] = FileFormatXEAC.NOT_EMISSION_VALUE + FileFormatXEAC.NOT_EMISSION_CHARACTER; 

                     if (nTotalMinutes == 60) bFinishDocuments = true;
                  }

                  if (nSheet != 17) {
                     oDirtyStreamWriter.Write(Environment.NewLine);
                     if (Array.IndexOf(FileFormatXEAC.oCleans, FileFormatXEAC.oCodes[nSheet]) != -1) oCleanStreamWriter.Write(Environment.NewLine); 
                  }
               }

               oDirtyStreamWriter.Write(Environment.NewLine);
               oDirtyStreamWriter.Write(FileFormatXEAC.oCodes[18]);
               for(int j = 0; j < Program.nNumElements; j++) {
                  oDirtyStreamWriter.Write(oPendings[j]);
               }

               oDirtyStreamWriter.Write(Environment.NewLine);
               oDirtyStreamWriter.Write(FileFormatXEAC.oCodes[19]);
               for(int j = 0; j < Program.nNumElements; j++) {
                  oDirtyStreamWriter.Write(oIndependents[j]);
               }

               oDirtyStreamWriter.Write(Environment.NewLine);
               oDirtyStreamWriter.Write(FileFormatXEAC.oCodes[20]);
               for(int j = 0; j < Program.nNumElements; j++) {
                  oDirtyStreamWriter.Write(oRanges[j]);
               }

               oCleanStreamWriter.Flush();
               oDirtyStreamWriter.Flush();

               oCleanStreamWriter.Close();
               oDirtyStreamWriter.Close();

               if (bFinishDocuments) {
                  try {
                     File.Copy(oDestinyFiles[0], Program.FOLDER_DESTINY_XEAC_COPY + oDestinyFiles[2] + "\\" + oDestinyFiles[4] + "\\" + oDestinyFiles[5] + "\\" + oDestinyFiles[6]);
                     File.Copy(oDestinyFiles[1], Program.FOLDER_DESTINY_XEAC_COPY + oDestinyFiles[3] + "\\" + oDestinyFiles[4] + "\\" + oDestinyFiles[5] + "\\" + oDestinyFiles[7]);
                  }
                  catch(Exception) { }
               }

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
         oPendings = null;
         oIndependents = null;
         oRanges = null;

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
