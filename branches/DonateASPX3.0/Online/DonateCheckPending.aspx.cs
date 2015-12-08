using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;

public partial class Online_DonateCheckPending : System.Web.UI.Page
{
    string strStoreid;
    string strOrderId;
    string strAmount;
    string strDate;

    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            strOrderId = Request.Params["orderid"].ToString();
            //strStoreid = Request.Params["storeid"].ToString();
            //strAmount = Request.Params["amount"].ToString();
            strDate = Request.Params["date"].ToString();

            Dictionary<string, object> dict = new Dictionary<string, object>();
            String strSql = "select * from DONATE_IEPAY where orderid = '" + strOrderId + "'";
            DataTable dt = NpoDB.GetDataTableS(strSql, dict);
            bool flag = false;
            if (dt.Rows.Count > 0)
            {
                try
                {
                    string strSql2 = "update DONATE_IEPAY set IEPAY_returnOK = '" + strDate + "' where orderid = '" + strOrderId + "' and IEPAY_returnOK is NULL";
                    int intsqlcnt = NpoDB.ExecuteSQLS(strSql2, dict);
                    NpoDB.ExecuteSQLS(strSql2, dict);
                    if (intsqlcnt > 0)
                    {
                        flag = true;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                if (flag == true)
                {
                    Response.Write("Y");
                }
                else
                {
                    Response.Write("N");
                }
            }
        }

    }
}