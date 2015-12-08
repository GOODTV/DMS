using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Xml;
using System.Text;
using System.IO;

public partial class Online_DonateOTTRetURL : System.Web.UI.Page
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
        string Donor_Id = NpoDB.GetScalarS("select Donor_Id from Donate_Web where od_sob = @od_sob", dict);
        string Amoumt = NpoDB.GetScalarS("select REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,Donate_Amount),1),'.00','') from Donate_Web where od_sob = @od_sob", dict);

        lblTitle.Text += "<br/><br/>交易序號 ：" + OrderNumber;
        lblTitle.Text += "<br/><br/><font color='red'>授權狀態 ：成功</font>"; //信用卡
        lblTitle.Text += "<br/><br/>您的奉獻金額：NT$" + Amoumt + "元";
        lblTitle.Text += "<br/><br/>付款方式：信用卡 ";

        //20151014增加填寫問卷
        //ClientScript.RegisterStartupScript(this.GetType(), "s", "<script>if (confirm('感謝您的奉獻！您的寶貴意見將可以協助我們提供更優質的捐款服務。您願意現在填寫奉獻問卷調查嗎？')){window.open('OnlineQuestionnaire.aspx?OrderNumber=" + OrderNumber + "&Donor_Id=" + Donor_Id + "&Device=3', 'window','HEIGHT=570,WIDTH=550,top=50,left=50,toolbar=yes,scrollbars=yes,resizable=yes'); }else {alert('非常感謝您的奉獻！'); };</script>");
    }
}