using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmissionsExcel {

   class ElementValue {
      private string sDirtyValue;
      private string sRealValue;
      private string sType;
      private string sQAL2Min;
      private string sQAL2Max;
      private string sQAL2A;
      private string sQAL2B;

      public ElementValue(string sDirtyValue, string sRealValue, string sType, string sQAL2Min, string sQAL2Max, string sQAL2A, string sQAL2B) {
         this.sDirtyValue = sDirtyValue;
         this.sRealValue = sRealValue;
         this.sType = sType;
         this.sQAL2Min = sQAL2Min;
         this.sQAL2Max = sQAL2Max;
         this.sQAL2A = sQAL2A;
         this.sQAL2B = sQAL2B;
      }

      public void setDirtyValue(string sDirtyValue) { this.sDirtyValue = sDirtyValue; }
      public void setRealValue(string sRealValue) { this.sRealValue = sRealValue; }
      public void setType(string sType) { this.sType = sType; }
      public void setQAL2Min(string sQAL2Min) { this.sQAL2Min = sQAL2Min; }
      public void setQAL2Max(string sQAL2Max) { this.sQAL2Max = sQAL2Max; }
      public void setQAL2A(string sQAL2A) { this.sQAL2A = sQAL2A; }
      public void setQAL2B(string sQAL2B) { this.sQAL2B = sQAL2B; }

      public string getDirtyValue() { return this.sDirtyValue; }
      public string getRealValue() { return this.sRealValue; }
      public string getType() { return this.sType; }
      public string getQAL2Min() { return this.sQAL2Min; }
      public string getQAL2Max() { return this.sQAL2Max; }
      public string getQAL2A() { return this.sQAL2A; }
      public string getQAL2B() { return this.sQAL2B; }
   }
}
