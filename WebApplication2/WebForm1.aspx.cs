using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace WebApplication2
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.TableName = "TblService";

            dt.Columns.Add("Service_Id");
            dt.Columns.Add("Service_Name");

            dt.Rows.Add("1", "AAA");
            dt.Rows.Add("2", "BBB");

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);

            DataTable dt1 = new DataTable();
            dt1.TableName = "Service_Method";

            dt1.Columns.Add("Id");
            dt1.Columns.Add("Service_Id");
            dt1.Columns.Add("Method_Name");

            //dt1.Rows.Add("", "", ""); If I use this I get 2 tables serialized. But I want without using this step.

            ds.Tables.Add(dt1);

           // var json = JsonConvert.SerializeObject(ds, Newtonsoft.Json.Formatting.Indented);
        }
    }
}