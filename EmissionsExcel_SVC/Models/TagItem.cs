using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EmissionsExcel {

    public class TagItem {
       private string sName;
       private string sValue;

       public TagItem(string sName) {
          this.sName = sName;
       }

       public void setValue(string sValue) { this.sValue = sValue; }

       public string getName() { return this.sName; }
       public string getValue() { return this.sValue; }
    }
}
