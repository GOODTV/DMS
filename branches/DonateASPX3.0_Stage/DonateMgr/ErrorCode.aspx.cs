using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class DonateMgr_ErrorCode : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        Print();
    }
    private void Print()
    {
        string strSql;
        DataTable dt;
        strSql = @" select ErrorCode , DefinedType , Note FROM Pledge_Status";

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dt = NpoDB.GetDataTableS(strSql, dict);
        GridList.Text = "";
        if (dt.Rows.Count == 0)
        {
            GridList.Text = "** 沒有符合條件的資料 **";
        }
        else
        {
            string strBody = "<table class='table_h' width='100%'>";
            strBody += @"<TR><TH noWrap>ErrorCode</SPAN></TH>
                            <TH noWrap><SPAN>DefinedType</SPAN></TH>
                            <TH noWrap><SPAN>Note</SPAN></TH>";
            strBody += "</TR>";
            foreach (DataRow dr in dt.Rows)
            {

                strBody += "<TR>";
                strBody += "<TD noWrap><SPAN>" + dr["ErrorCode"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["DefinedType"].ToString() + "</SPAN></TD>";
                strBody += "<TD noWrap><SPAN>" + dr["Note"].ToString() + "&nbsp;</SPAN></TD>";
                strBody += "</TR>";
            }

            strBody += "</table>";
            GridList.Text = strBody;
        }
    }
}