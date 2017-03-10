using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TableParagraph
{
    class myData
    {
       public string title;
       public  string remark;
        public DataTable dt;
        public myData(string t,string r,DataTable dt)
        {
            title = t;
            remark = r;
            this.dt = dt;
        }
        public myData()
        {

        }
    }
}
