using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OPCAutomation;

namespace EmissionsExcel {
   
   class ClientOPC {
      string sOPCServer;
      string sServerHost;
      bool bConnect;
      OPCServer oOPCServer;
      Array oOPCItemIDs = Array.CreateInstance(typeof(string), 512);
      Array oItemServerHandles = Array.CreateInstance(typeof(Int32), 512);
      Array oItemServerErrors = Array.CreateInstance(typeof(Int32), 512);
      Array oClientHandles = Array.CreateInstance(typeof(Int32), 512);
      Array oRequestedDataTypes = Array.CreateInstance(typeof(Int16), 512);
      Array oAccessPaths = Array.CreateInstance(typeof(string), 512);
      Array oItemServerValues = Array.CreateInstance(typeof(string), 512);
      HashSet<TagItem> oHashSetTagsValues;
      OPCGroup oOpcGroupNames;
      int nOPCItemsCount = 0; 

      public ClientOPC(string sServerHost, string sOPCServer) {
         this.sServerHost = sServerHost;
         this.sOPCServer = sOPCServer;

         oHashSetTagsValues = new HashSet<TagItem>();
          
         this.bConnect = false;
      }

      public bool connect() {
         try {
            oOPCServer = new OPCServer();
            oOPCServer.Connect(this.sOPCServer, this.sServerHost);

            this.bConnect = true;
         }
         catch (Exception) {
            this.bConnect = false;
         }
         return this.bConnect;
      }
      
      public void setTags(string[] oTagsNames) {
         this.nOPCItemsCount = 0;
         
         try {
            if ((this.bConnect) && (oTagsNames.Length > 0)) {
               oOpcGroupNames = oOPCServer.OPCGroups.Add("Group01");
               oOpcGroupNames.DeadBand = 0;
               oOpcGroupNames.UpdateRate = 100;
               oOpcGroupNames.IsSubscribed = true;
               oOpcGroupNames.IsActive = true;

               foreach(string oTagName in oTagsNames) {
                  oOPCItemIDs.SetValue(oTagName, (this.nOPCItemsCount + 1));

                  this.nOPCItemsCount++;
               }

               oOpcGroupNames.OPCItems.AddItems(this.nOPCItemsCount, ref oOPCItemIDs, ref oClientHandles, out oItemServerHandles, out oItemServerErrors, oRequestedDataTypes, oAccessPaths);
            }
         } catch(Exception) {

         }
      }

      public HashSet<TagItem> getTagsValues() {
         object oObject1, oObject2;
         oHashSetTagsValues.Clear();
         bool bError = false;

         try {
            if ((this.bConnect) && (this.nOPCItemsCount > 0)) {
            
               oOpcGroupNames.SyncRead((short) OPCAutomation.OPCDataSource.OPCDevice, this.nOPCItemsCount, ref oItemServerHandles, out oItemServerValues, out oItemServerErrors, out oObject1, out oObject2);
               
               for(int i = 1; ((i <= this.nOPCItemsCount) && (!bError)); i++) {
                  if (((int) oItemServerErrors.GetValue(i)) != 0) bError = true;
               }

               if (!bError) {
                  for(int i = 1; i <= this.nOPCItemsCount; i++) {
                     TagItem oTagItem = new TagItem((string) oOPCItemIDs.GetValue(i));
                     oTagItem.setValue(oItemServerValues.GetValue(i).ToString());

                     oHashSetTagsValues.Add(oTagItem);
                  }
               }
            }
         } catch(Exception) {
            oHashSetTagsValues.Clear();
         }

         return oHashSetTagsValues;
      }

      public void disconnect() {
         if (this.bConnect) oOPCServer.Disconnect(); 
      }
   }
}
