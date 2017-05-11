using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using EmissionsExcel;
using EmissionsExcel_SVC.Models;

namespace EmissionsExcel_SVC {

   public partial class ServiceEmissionsExcel : ServiceBase {
      private System.Object oThreadLockResource = new System.Object();

      public ServiceEmissionsExcel() {
         InitializeComponent();
      }

      protected override void OnStart(string[] args) {
         DateTime oDateNow = DateTime.Now;

         if (existsBaseDirectories()) {
            Program.oClientOPC = new ClientOPC("E1390SVR1", "SdrPointSvr30.EcsOpcServer.1");
         
            if (Program.bShowLog) Program.oLogFile = File.AppendText(Program.FOLDER_FILE_LOG + "EmissionsExcel_log_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + ".txt");

            if (Program.oClientOPC.connect()) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Comunicación establecida con el servidor {E1390SVR1:SdrPointSvr30.EcsOpcServer.1}", false);

               Program.oClientOPC.setTags(Program.oTagsNames);
        
               System.Timers.Timer oTimer = new System.Timers.Timer(30000);
            
               oTimer.Interval = 30000;
               oTimer.Elapsed += new ElapsedEventHandler(timerHandler);
               oTimer.Enabled = true;

               timerHandler(null, null);
            }
            else { 
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Comunicación no establecida con el servidor {E1390SVR1:SdrPointSvr30.EcsOpcServer.1}", false); 
               this.Stop(); 
            }
         }
         else {
            if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Debe crear de forma manual los directorios base de la aplicación (" + Program.FOLDER_DESTINY + "," + Program.FOLDER_DESTINY_COPY + "," + Program.FOLDER_DESTINY_XEAC + "," + Program.FOLDER_DESTINY_XEAC_COPY + "," + Program.FOLDER_FILE_LOG + "," + Program.FOLDER_FILE_LOG_HISTORY + ")", false); 
            this.Stop();
         }
      }

      protected override void OnStop() {
         Program.oClientOPC.disconnect();

         if (Program.oLogFile != null) {
            Program.oLogFile.Close();
         }
      }

      private bool existsBaseDirectories() {
         bool bExistsDirectories = false;

         try {
            if ((Directory.Exists(Program.FOLDER_DESTINY)) && (Directory.Exists(Program.FOLDER_DESTINY_COPY)) && (Directory.Exists(Program.FOLDER_DESTINY_XEAC)) && (Directory.Exists(Program.FOLDER_DESTINY_XEAC_COPY)) && (Directory.Exists(Program.FOLDER_DESTINY_XEAC_REPORTS)) && (Directory.Exists(Program.FOLDER_FILE_LOG)) && (Directory.Exists(Program.FOLDER_FILE_LOG_HISTORY))) {
               if (!Directory.Exists(Program.FOLDER_TEMP)) Directory.CreateDirectory(Program.FOLDER_TEMP);

               bExistsDirectories = true;
            }
         } catch(Exception) { bExistsDirectories = false; }
 
         return bExistsDirectories;
      }

      private void timerHandler(object source, ElapsedEventArgs e) {
         DateTime oDateNow = DateTime.Now;
         int nMinute = oDateNow.Minute;

         lock (oThreadLockResource) {
            Microsoft.Office.Interop.Excel.Application oExcelApplication = null;
            Microsoft.Office.Interop.Excel.Workbook oExcelWorkBook = null;
            HashSet<TagItem> oHashSetTagsValues = null;

            try {
               try {
                  if (nMinute == 0) {
                     string sLogFile = Program.FOLDER_FILE_LOG + "EmissionsExcel_log_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + ".txt";

                     if (!File.Exists(sLogFile)) {
                        if (Program.bShowLog) {
                           if (Program.oLogFile != null) {
                              Program.oLogFile.Flush();
                              Program.oLogFile.Close();
                           }

                           Program.oLogFile = File.AppendText(sLogFile);

                           DateTime oDateTimeBefore = oDateNow.AddDays(-1);
                           string sLogFileBefore = Program.FOLDER_FILE_LOG + "EmissionsExcel_log_" + oDateTimeBefore.Year.ToString() + oDateTimeBefore.Month.ToString().PadLeft(2, '0') + oDateTimeBefore.Day.ToString().PadLeft(2, '0') + ".txt";

                           if (File.Exists(sLogFileBefore)) {
                              File.Move(sLogFileBefore, Program.FOLDER_FILE_LOG_HISTORY + "EmissionsExcel_log_" + oDateTimeBefore.Year.ToString() + oDateTimeBefore.Month.ToString().PadLeft(2, '0') + oDateTimeBefore.Day.ToString().PadLeft(2, '0') + ".txt");
                           }
                        }
                     }
                  }
               } catch(Exception) { }

               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{TIMERHANDLER_START:OK} Llamada al evento {timerHandler} al segundo " + oDateNow.Second, true);

               if (nMinute != Program.nLastMinute) {
                  Program.nLastMinute = nMinute;
                  if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Creando valores de los contaminantes sobre el minuto " + oDateNow.Minute, false);

                  for(int nLine = 1; nLine <= Program.nNumLines; nLine++) {
                     string sPathFilename = Program.FOLDER_DESTINY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
                     if ((!File.Exists(sPathFilename)) || ((File.Exists(sPathFilename)) && (nMinute == 0))) {
                        if ((File.Exists(sPathFilename)) && (nMinute == 0)) File.Delete(sPathFilename);

                        oExcelApplication = new Microsoft.Office.Interop.Excel.Application();
                        oExcelApplication.Visible = false;

                        Microsoft.Office.Interop.Excel.Workbooks oExcelWorkBooks = oExcelApplication.Workbooks;
                        oExcelWorkBook = oExcelWorkBooks.Add();
                        createSheets(oExcelApplication, oDateNow);
                        removeDefaultSheets(oExcelApplication);
               
                        oExcelWorkBook.SaveAs(sPathFilename, XlFileFormat.xlExcel5);
                        oExcelWorkBook.Close(false);
                        nullExcelObject(oExcelWorkBook);

                        oExcelWorkBooks.Close();
                        nullExcelObject(oExcelWorkBooks);

                        oExcelApplication.Application.Quit();
                        nullExcelObject(oExcelApplication);

                        killExcelProcesses();
                     }
                  }

                  oHashSetTagsValues = TryExecuteSyncRead(Program.nTimeoutOPCRead); 

                  if ((oHashSetTagsValues != null) && (oHashSetTagsValues.Count == (((Program.nNumElements * 9) + 2) * Program.nNumLines))) {
                     DateTime oDateNowCurrent = DateTime.Now;
                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Se han devuelto todos los valores del OPC de forma correcta sobre el minuto " + oDateNow.Minute + " a la hora: " + oDateNowCurrent.Hour + "h " + oDateNowCurrent.Minute + "m " + oDateNowCurrent.Second + "s", false);

                     for(int nLine = 1; nLine <= Program.nNumLines; nLine++) {
                        if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Creando valores de los contaminantes de la Línea " + nLine + " sobre el minuto " + oDateNow.Minute, false);

                        Array oArrayTagsValues = oHashSetTagsValues.ToArray<TagItem>();
                        string sPathFilename = Program.FOLDER_DESTINY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
                        string sPathFilenameCopy = Program.FOLDER_DESTINY_COPY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
              
                        if (File.Exists(sPathFilename)) {
                           oExcelApplication = new Microsoft.Office.Interop.Excel.Application();
                           oExcelApplication.Visible = false;

                           Microsoft.Office.Interop.Excel.Workbooks oExcelWorkBooks = oExcelApplication.Workbooks;
                           oExcelWorkBook = oExcelWorkBooks.Open(sPathFilename);

                           if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel de la línea " + nLine + " abierto correctamente", false);

                           TagItem oTagItem1, oTagItem2, oTagItem3, oTagItem4, oTagItem5, oTagItem6, oTagItem7, oTagItem8, oTagItem9, oTagItem10, oTagItem11; 
                           for(int i = 1; i <= Program.nNumElements; i++) {
                              string sStatus = "V";

                              int nIndexBase = (((Program.nNumElements * 9) + 2) * (nLine - 1)) + ((i - 1) * 9) + 2;

                              oTagItem1 = (TagItem) oArrayTagsValues.GetValue((((Program.nNumElements * 9) + 2) * (nLine - 1)));
                              oTagItem2 = (TagItem) oArrayTagsValues.GetValue((((Program.nNumElements * 9) + 2) * (nLine - 1)) + 1);

                              oTagItem3 = (TagItem) oArrayTagsValues.GetValue(nIndexBase);
                              oTagItem4 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 1);
                              oTagItem5 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 2);
                              oTagItem6 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 3);
                              oTagItem7 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 4);

                              oTagItem8 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 5);
                              oTagItem9 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 6);
                              oTagItem10 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 7);
                              oTagItem11 = (TagItem) oArrayTagsValues.GetValue(nIndexBase + 8);

                              if (oTagItem1.getValue() == "1") sStatus = "A";
                              else if (oTagItem2.getValue() == "1") sStatus = "P";
                              else if (oTagItem5.getValue() == "1") sStatus = "V";
                              else if (oTagItem6.getValue() == "1") sStatus = "C";
                              else if (oTagItem7.getValue() == "1") sStatus = "N";

                              if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Volcando datos al Fichero Excel del contaminante " + i.ToString() + " de la línea " + nLine + " {" + oTagItem3.getName().ToString() + ":" + oTagItem3.getValue().ToString() + ", " + oTagItem4.getName().ToString() + ":" + oTagItem4.getValue().ToString() + ", Status: " + sStatus + "}", false);
                              
                              ElementValue oElementValue = new ElementValue(oTagItem3.getValue(), oTagItem4.getValue(), sStatus, oTagItem8.getValue(), oTagItem9.getValue(), oTagItem10.getValue(), oTagItem11.getValue());
                              Microsoft.Office.Interop.Excel.Worksheet oWorkSheet = oExcelWorkBook.Worksheets[(i + 1)];
                              flushMinuteValue(oWorkSheet, oDateNow, nMinute, oElementValue, (i - 1));
                              oElementValue = null;
    
                              oTagItem1 = null; oTagItem2 = null; oTagItem3 = null; oTagItem4 = null; oTagItem5 = null; 
                              oTagItem6 = null; oTagItem7 = null; oTagItem8 = null; oTagItem9 = null; oTagItem10 = null; 
                              oTagItem11 = null; 

                              nullExcelObject(oWorkSheet);

                              if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Datos volcados al Fichero Excel del contaminante " + i.ToString() + " de la línea " + nLine + " de forma correcta", false);
                           }

                           try {
                              oExcelWorkBook.Save();
                              Program.writeLineInLogFile(Program.oLogFile, "{OK} Datos salvados al Fichero Excel de la línea " + nLine + " de forma correcta", false);

                              if (nMinute == 59) {
                                 if (File.Exists(sPathFilenameCopy)) File.Delete(sPathFilenameCopy);
                                 oExcelWorkBook.SaveAs(sPathFilenameCopy, XlFileFormat.xlExcel5);

                                 if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Copia del Fichero Excel de la línea " + nLine + " realizada correctamente", false);
                              }
                           } catch(Exception) {
                              if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error salvando los datos del Fichero Excel de la línea " + nLine, false);
                           }
                           
                           oExcelWorkBook.Close(false);
                           nullExcelObject(oExcelWorkBook);

                           oExcelWorkBooks.Close();
                           nullExcelObject(oExcelWorkBooks);

                           oExcelApplication.Application.Quit();
                           nullExcelObject(oExcelApplication);

                           killExcelProcesses();

                           FileFormatXEAC oFileFormatXEAC = new FileFormatXEAC();
                           if (oFileFormatXEAC.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
                              if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC", false);   
                           }
                           oFileFormatXEAC = null;

                           FileFormatXEACReports oFileFormatXEACReports = new FileFormatXEACReports();
                           if (oFileFormatXEACReports.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
                              if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC - Reports", false);   
                           }
                           oFileFormatXEACReports = null;
                        }
                        else {
                           if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} No existe el Fichero Excel de la Línea " + nLine, false);   
                        }

                        oArrayTagsValues = null;
                     }
                  } else {
                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} No se han devuelto todos los valores del OPC sobre el evento {timerHandler}", false);

                     killExcelProcesses();

                     if (Program.sMethod == "BEST_EFFORT") {
                        if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} Creando valores anteriores de los contaminantes sobre el minuto " + oDateNow.Minute, false);  
                     } else {
                        if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} Creando líneas vacias de los contaminantes sobre el minuto " + oDateNow.Minute, false);  
                     }

                     for(int nLine = 1; nLine <= Program.nNumLines; nLine++) {
                        if (Program.sMethod == "BEST_EFFORT") {
                           cleanOPCReaderValues(oDateNow, nMinute, nLine);
                        }
                        else {
                           flushEmptyMinute(oDateNow, nMinute, nLine);
                        }
                     }

                     if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} Fin de la creación de valores anteriores de los contaminantes", false);
                  }
               }
            } catch (Exception err) {
                if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{EXCEPTION} Se ha producido una excepción sobre el evento {timerHandler:" + err.ToString() + "}", false);

                killExcelProcesses();
                
                if (Program.sMethod == "BEST_EFFORT") {
                   if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{EXCEPTION} Creando valores anteriores de los contaminantes sobre el minuto " + oDateNow.Minute, false);
                } else {
                   if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{EXCEPTION} Creando líneas vacias de los contaminantes sobre el minuto " + oDateNow.Minute, false);
                }

                for(int nLine = 1; nLine <= Program.nNumLines; nLine++) {
                   if (Program.sMethod == "BEST_EFFORT") {
                      cleanOPCReaderValues(oDateNow, nMinute, nLine);
                   }
                   else {
                      flushEmptyMinute(oDateNow, nMinute, nLine);
                   }
                }

                if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{EXCEPTION} Fin de la creación de valores anteriores de los contaminantes", false);
            }

            // Clean memory
            oHashSetTagsValues = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{TIMERHANDLER_END:OK} Fin de la llamada al evento {timerHandler}", false);
         }
      }

      private void flushEmptyMinute(DateTime oDateNow, int nMinute, int nLine) {
         string sPathFilename = Program.FOLDER_DESTINY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
         string sPathFilenameCopy = Program.FOLDER_DESTINY_COPY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
         int nRowStart = 4;

         if (File.Exists(sPathFilename)) {
            Microsoft.Office.Interop.Excel.Application oExcelApplication = new Microsoft.Office.Interop.Excel.Application();
            oExcelApplication.Visible = false;

            Microsoft.Office.Interop.Excel.Workbooks oExcelWorkBooks = oExcelApplication.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook oExcelWorkBook = oExcelWorkBooks.Open(sPathFilename);
                  
            for(int i = 1; i <= Program.nNumElements; i++) {
               if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Volcando la fecha y la hora al Fichero Excel del contaminante " + i.ToString() + " la línea " + nLine, false);

               Microsoft.Office.Interop.Excel.Worksheet oWorkSheet = oExcelWorkBook.Worksheets[(i + 1)];
               oWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "A"] = "'" + oDateNow.Day.ToString().PadLeft(2, '0') + "/" + oDateNow.Month.ToString().PadLeft(2, '0') + "/" + oDateNow.Year + " " + oDateNow.Hour.ToString().PadLeft(2, '0') + ":" + nMinute.ToString().PadLeft(2, '0') + ":00";
               nullExcelObject(oWorkSheet);

               if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fecha y hora volcandos al Fichero Excel del contaminante " + i.ToString() + " de la línea " + nLine + " de forma correcta", false);
            }

            try {
               oExcelWorkBook.Save();
               Program.writeLineInLogFile(Program.oLogFile, "{OK} Datos salvados al Fichero Excel de la línea " + nLine + " de forma correcta", false);

               if (nMinute == 59) {
                  if (File.Exists(sPathFilenameCopy)) File.Delete(sPathFilenameCopy);
                  oExcelWorkBook.SaveAs(sPathFilenameCopy, XlFileFormat.xlExcel5);

                  if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Copia del Fichero Excel de la línea " + nLine + " realizada correctamente", false);
               }
            } catch(Exception) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error salvando los datos del Fichero Excel de la línea " + nLine, false);   
            }

            oExcelWorkBook.Close(false);
            nullExcelObject(oExcelWorkBook);

            oExcelWorkBooks.Close();
            nullExcelObject(oExcelWorkBooks);

            oExcelApplication.Application.Quit();
            nullExcelObject(oExcelApplication);

            killExcelProcesses();

            FileFormatXEAC oFileFormatXEAC = new FileFormatXEAC();
            if (oFileFormatXEAC.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC", false);  
            }
            oFileFormatXEAC = null;

            FileFormatXEACReports oFileFormatXEACReports = new FileFormatXEACReports();
            if (oFileFormatXEACReports.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC - Reports", false);  
            }
            oFileFormatXEACReports = null;
         }
      }

      private void cleanOPCReaderValues(DateTime oDateNow, int nMinute, int nLine) {
         string sPathFilename = Program.FOLDER_DESTINY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";
         string sPathFilenameCopy = Program.FOLDER_DESTINY_COPY + Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls";

         if (File.Exists(sPathFilename)) {
            Microsoft.Office.Interop.Excel.Application oExcelApplication = new Microsoft.Office.Interop.Excel.Application();
            oExcelApplication.Visible = false;

            Microsoft.Office.Interop.Excel.Workbooks oExcelWorkBooks = oExcelApplication.Workbooks;
            Microsoft.Office.Interop.Excel.Workbook oExcelWorkBook = oExcelWorkBooks.Open(sPathFilename);
                  
            for(int i = 1; i <= Program.nNumElements; i++) {
               if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Volcando datos al Fichero Excel del contaminante " + i.ToString() + " la línea " + nLine + " {" + Program.oBeforeElementsValues[i - 1].getDirtyValue().ToString() + ", " + Program.oBeforeElementsValues[i - 1].getRealValue().ToString() + ", " + Program.oBeforeElementsValues[i - 1].getType() + "}", false);

               ElementValue oElementValue = new ElementValue(Program.oBeforeElementsValues[i - 1].getDirtyValue(), Program.oBeforeElementsValues[i - 1].getRealValue(), Program.oBeforeElementsValues[i - 1].getType(), Program.oBeforeElementsValues[i - 1].getQAL2Min(), Program.oBeforeElementsValues[i - 1].getQAL2Max(), Program.oBeforeElementsValues[i - 1].getQAL2A(), Program.oBeforeElementsValues[i - 1].getQAL2B());
               Microsoft.Office.Interop.Excel.Worksheet oWorkSheet = oExcelWorkBook.Worksheets[(i + 1)];
               flushMinuteValue(oWorkSheet, oDateNow, nMinute, oElementValue, (i - 1));
               oElementValue = null;
               nullExcelObject(oWorkSheet);

               if ((Program.bShowLog) && (Program.bShowVerboseLog)) Program.writeLineInLogFile(Program.oLogFile, "{OK} Datos volcados al Fichero Excel del contaminante " + i.ToString() + " de la línea " + nLine + " de forma correcta", false);
            }

            try {
               oExcelWorkBook.Save();
               Program.writeLineInLogFile(Program.oLogFile, "{OK} Datos salvados al Fichero Excel de la línea " + nLine + " de forma correcta", false);

               if (nMinute == 59) {
                  if (File.Exists(sPathFilenameCopy)) File.Delete(sPathFilenameCopy);
                  oExcelWorkBook.SaveAs(sPathFilenameCopy, XlFileFormat.xlExcel5);

                  if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Copia del Fichero Excel de la línea " + nLine + " realizada correctamente", false);
               }
            } catch(Exception) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{ERROR} Error salvando los datos del Fichero Excel de la línea " + nLine, false);   
            }

            oExcelWorkBook.Close(false);
            nullExcelObject(oExcelWorkBook);

            oExcelWorkBooks.Close();
            nullExcelObject(oExcelWorkBooks);

            oExcelApplication.Application.Quit();
            nullExcelObject(oExcelApplication);

            killExcelProcesses();

            FileFormatXEAC oFileFormatXEAC = new FileFormatXEAC();
            if (oFileFormatXEAC.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC", false);  
            }
            oFileFormatXEAC = null;

            FileFormatXEACReports oFileFormatXEACReports = new FileFormatXEACReports();
            if (oFileFormatXEACReports.convertExcelFile(Program.FILE_EXCEL_TOKEN_NAME + "DayExcelL" + nLine + "_" + oDateNow.Year.ToString() + oDateNow.Month.ToString().PadLeft(2, '0') + oDateNow.Day.ToString().PadLeft(2, '0') + "_" + oDateNow.Hour.ToString().PadLeft(2, '0') + "00_A_Hour_Hourly.xls", (nMinute + 1))) {
               if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Fichero Excel convertido al formato de la XEAC - Reports", false);  
            }
            oFileFormatXEACReports = null;
         }    
      }

      public HashSet<TagItem> TryExecuteSyncRead(int nTimeout) {
          HashSet<TagItem> oResult = new HashSet<TagItem>();
          DateTime oDateNow = DateTime.Now;
          Thread oThread = null;

          if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{OK} Llamada al método SyncRead a la hora: " + oDateNow.Hour + "h " + oDateNow.Minute + "m " + oDateNow.Second + "s", false);

          try {
             oThread = new Thread(() => { oResult = Program.oClientOPC.getTagsValues(); });
             oThread.IsBackground = true;
             oThread.Start();

             Thread.Sleep(nTimeout);
             if (oThread.IsAlive) {
                oDateNow = DateTime.Now;

                if (Program.bShowLog) Program.writeLineInLogFile(Program.oLogFile, "{WARNING} Ha expirado el tiempo de la llamada al método SyncRead (" + nTimeout + "ms) para devolver los resultados del OPC a la hora: " + oDateNow.Hour + "h " + oDateNow.Minute + "m " + oDateNow.Second + "s", false);
                oThread.Abort();
                oResult.Clear();
             }
          }
          catch (Exception) { oResult.Clear(); }

          oThread = null; 

          return oResult;
      }

      private void createSheets(Microsoft.Office.Interop.Excel.Application oExcelApplication, DateTime oDateNow) {
         int nMinute = oDateNow.Minute;

         Microsoft.Office.Interop.Excel.Worksheet oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "TempHogar";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Flow";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "NH3";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "CO2";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "O2";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Temp";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Press";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "H2O";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Hg";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "NO";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "NO2";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "NOx";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "SO2";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "HF";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "HCl";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "TOC";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Polvo";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "CO";
         createHeaderSheet(oExcelWorkSheet);
         nullExcelObject(oExcelWorkSheet);

         oExcelWorkSheet = oExcelApplication.Worksheets.Add();
         oExcelWorkSheet.Name = "Not Used";
         nullExcelObject(oExcelWorkSheet);
      }

      private void createHeaderSheet(Microsoft.Office.Interop.Excel.Worksheet oExcelWorkSheet) {
         oExcelWorkSheet.Columns["A"].ColumnWidth = 20;
         oExcelWorkSheet.Columns["B"].ColumnWidth = 15;
         oExcelWorkSheet.Columns["C"].ColumnWidth = 15;
         oExcelWorkSheet.Columns["D"].ColumnWidth = 10;
         oExcelWorkSheet.Columns["E"].ColumnWidth = 10;
         oExcelWorkSheet.Columns["F"].ColumnWidth = 10;
         oExcelWorkSheet.Columns["G"].ColumnWidth = 10;
         oExcelWorkSheet.Columns["H"].ColumnWidth = 10;

         oExcelWorkSheet.Range["A1:H4"].HorizontalAlignment = -4108;

         oExcelWorkSheet.Cells["1", "B"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "C"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "D"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "E"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "F"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "G"] = "'" + oExcelWorkSheet.Name;
         oExcelWorkSheet.Cells["1", "H"] = "'" + oExcelWorkSheet.Name;

         oExcelWorkSheet.Cells["2", "B"] = "'Average";
         oExcelWorkSheet.Cells["2", "C"] = "'Average";
         oExcelWorkSheet.Cells["2", "D"] = "'Character";
         oExcelWorkSheet.Cells["2", "E"] = "'QAL2";
         oExcelWorkSheet.Cells["2", "F"] = "'QAL2";
         oExcelWorkSheet.Cells["2", "G"] = "'QAL2";
         oExcelWorkSheet.Cells["2", "H"] = "'QAL2";

         oExcelWorkSheet.Cells["3", "B"] = "'mg/m3";
         oExcelWorkSheet.Cells["3", "C"] = "'mg/Nm3";
         oExcelWorkSheet.Cells["3", "D"] = "'Validation";
         oExcelWorkSheet.Cells["3", "E"] = "'RANGO";
         oExcelWorkSheet.Cells["3", "F"] = "'RANGO";
         oExcelWorkSheet.Cells["3", "G"] = "'A";
         oExcelWorkSheet.Cells["3", "H"] = "'B";

         oExcelWorkSheet.Cells["4", "B"] = "'1m";
         oExcelWorkSheet.Cells["4", "C"] = "'1m";
         oExcelWorkSheet.Cells["4", "D"] = "'1m";
         oExcelWorkSheet.Cells["4", "E"] = "'Min";
         oExcelWorkSheet.Cells["4", "F"] = "'Max";
         oExcelWorkSheet.Cells["4", "G"] = "'+";
         oExcelWorkSheet.Cells["4", "H"] = "'*";
      }

      private void removeDefaultSheets(Microsoft.Office.Interop.Excel.Application oExcelApplication) {
         for (int i = oExcelApplication.ActiveWorkbook.Worksheets.Count; i > 0 ; i--) {
             Worksheet oExcelWorkSheet = (Worksheet) oExcelApplication.ActiveWorkbook.Worksheets[i];
             
             if ((oExcelWorkSheet.Name == "Hoja1") || (oExcelWorkSheet.Name == "Hoja2") || (oExcelWorkSheet.Name == "Hoja3") || (oExcelWorkSheet.Name == "Hoja4")) {
                oExcelWorkSheet.Delete();
             }

             nullExcelObject(oExcelWorkSheet);
         }
      }

      private void flushMinuteValue(Microsoft.Office.Interop.Excel.Worksheet oExcelWorkSheet, DateTime oDate, int nMinute, ElementValue nElement, int nIndex) {
         int nRowStart = 4;

         Program.oBeforeElementsValues[nIndex].setDirtyValue(nElement.getDirtyValue());
         Program.oBeforeElementsValues[nIndex].setRealValue(nElement.getRealValue());
         Program.oBeforeElementsValues[nIndex].setType(nElement.getType());

         Program.oBeforeElementsValues[nIndex].setQAL2Min(nElement.getQAL2Min());
         Program.oBeforeElementsValues[nIndex].setQAL2Max(nElement.getQAL2Max());
         Program.oBeforeElementsValues[nIndex].setQAL2A(nElement.getQAL2A());
         Program.oBeforeElementsValues[nIndex].setQAL2B(nElement.getQAL2B());
         
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "A"] = "'" + oDate.Day.ToString().PadLeft(2, '0') + "/" + oDate.Month.ToString().PadLeft(2, '0') + "/" + oDate.Year + " " + oDate.Hour.ToString().PadLeft(2, '0') + ":" + nMinute.ToString().PadLeft(2, '0') + ":00";
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "B"] = nElement.getDirtyValue().Replace(',', '.');
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "C"] = nElement.getRealValue().Replace(',', '.');
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "D"] = "'" + nElement.getType();

         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "E"] = nElement.getQAL2Min().Replace(',', '.');
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "F"] = nElement.getQAL2Max().Replace(',', '.');
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "G"] = nElement.getQAL2A().Replace(',', '.');
         oExcelWorkSheet.Cells[(nMinute + nRowStart + 1).ToString(), "H"] = nElement.getQAL2B().Replace(',', '.');

         nullExcelObject(oExcelWorkSheet);
      }
          
      private void nullExcelObject(Object oObject) {
         try {
            for(;System.Runtime.InteropServices.Marshal.ReleaseComObject(oObject) > 0;);
         } catch(Exception) { }

         oObject = null;

         GC.Collect();
         GC.WaitForPendingFinalizers();
      }

      private void killExcelProcesses() {
         try {
            System.Diagnostics.Process[] oAllProcesses = System.Diagnostics.Process.GetProcessesByName("excel");

            foreach (System.Diagnostics.Process oExcelProcess in oAllProcesses) {
              oExcelProcess.Kill();
            }
         } catch(Exception) { }
       }
   }
}
