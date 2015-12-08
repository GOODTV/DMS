using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Text;
using System.IO;

public partial class Online_DonateOTTFailRetURL : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        //Get OrderInfo Result
        string strOrderInfo = Request["strOrderInfo"].ToString();
        XmlDocument XmlDoc = new XmlDocument();
        XmlDoc.LoadXml(strOrderInfo);
        //訂單編號
        XmlNode ordernumber = XmlDoc.SelectSingleNode("/CUBXML/ORDERINFO/ORDERNUMBER");
        string OrderNumber = ordernumber.InnerText;

        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("od_sob", OrderNumber);
        string Amoumt = NpoDB.GetScalarS("select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amount),1),'.00','') from Donate_Web where od_sob = @od_sob", dict);

        lblTitle.Text += "<br/><br/>交易序號：" + OrderNumber;
        lblTitle.Text += "<br/><br/><font color='red'>授權狀態：失敗</font>";
        lblTitle.Text += "<br/><br/>您的奉獻金額：NT$" + Amoumt + "元";
        lblTitle.Text += "<br/><br/>付款方式：信用卡";
    }
}