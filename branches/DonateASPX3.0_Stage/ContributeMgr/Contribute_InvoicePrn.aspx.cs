using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Text;
using System.Web.UI.HtmlControls;
using System.IO;

public partial class ContributeMgr_Contribute_InvoicePrn : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            string Contribute_Id = Util.GetQueryString("Contribute_Id");
            if (Contribute_Id != "")
            {
                Print(Contribute_Id);
                Update(Contribute_Id);
            }
            else
            {
                GridList.Text += ("empty string");
            }
        }
    }
    public void Print(string Contribute_Id)
    {
        //抓DEPT欄位
        string strSql_Dept = "select * from Dept where DeptID='" + SessionInfo.DeptID + "'";
        DataTable dt_Dept = NpoDB.GetDataTableS(strSql_Dept, null);
        DataRow dr_Dept = dt_Dept.Rows[0];

        //抓Organization欄位
        string strSql_Organization = "select * from Organization where OrgID='" + SessionInfo.OrgID + "'";
        DataTable dt_Organization = NpoDB.GetDataTableS(strSql_Organization, null);
        DataRow dr_Organization = dt_Organization.Rows[0];

        int Row = 1;
        Dictionary<string, object> dict = new Dictionary<string, object>();
        string strSql = @"select (Case When D.NickName<>'' Then D.Donor_Name+'('+D.NickName+')' Else D.Donor_Name End)as [捐贈者], 
                                 D.IDNo as [身份證/統編], CONVERT(nvarchar,C.Contribute_Date,111) as [捐贈日期], C.Invoice_Pre + C.Invoice_No as [收據編號],
                                (Case When D.Invoice_City='' Then D.Invoice_Address Else Case When A.mValue<>B.mValue Then A.mValue+D.Invoice_ZipCode+B.mValue+D.Invoice_Address Else A.mValue+D.Invoice_ZipCode+D.Address End End) as [收據地址],
                                 left(C.Item,len(C.Item)-1)+'。' as [捐贈內容] , REPLACE(CONVERT(VARCHAR,CONVERT(MONEY,C.Contribute_Amt),1),'.00','') as [折合現金], Invoice_PrintComment as [備註]
                          from    
	                             (select CC.Donor_Id,CC.Contribute_Id,CC.Contribute_Date,CC.Invoice_Print,CC.Invoice_PrintComment,CC.Contribute_Amt, CC.Invoice_Pre, CC.Invoice_No,
		                         (SELECT cast(CD.Goods_Name+'(' + (CONVERT(VARCHAR,CD.Goods_Qty))+CD.Goods_Unit+')' AS NVARCHAR ) + '、'
		                         from Contribute CCC left join ContributeData CD on CCC.Contribute_Id = CD.Contribute_Id 
		                              right join Goods G on CD.Goods_Id = G.Goods_Id
		                              where CC.Contribute_Id  = CCC.Contribute_Id FOR XML PATH('')) as [Item]
	                                        from Contribute CC left join Donor D on CC.Donor_Id = D.Donor_Id )C 
	                     Left Join Donor D on C.Donor_Id = D.Donor_Id   
                         Left Join CODECITY As A On D.Invoice_City=A.mCode 
                         Left Join CODECITY As B On D.Invoice_Area=B.mCode
                         where Contribute_Id = @Contribute_Id";
        dict.Add("Contribute_Id", Contribute_Id);
        DataTable dt = NpoDB.GetDataTableS(strSql, dict);
        DataRow dr;
        if (dt.Rows.Count != 0)
        {
            dr = dt.Rows[0];
            bool end = false;
            while (end == false)
            {
                GridList.Text += "<table id='grid' border='0' cellpadding='0' cellspacing='0' class='TableGrid'>";
                for (int i = 0; i < 3; i++)
                {
                    GridList.Text += "<tr>";
                    GridList.Text += "  <td>";
                    GridList.Text += "    <table id='grid1' border='0' cellpadding='0' cellspacing='0' class='TableGrid" + i + "'>";
                    GridList.Text += "      <tr>";
                    GridList.Text += "        <td align='center' valign='top' style='height:20mm' colspan='3' class='Comp_Name'><u>" + SessionInfo.OrgName + "<br />捐贈收據</u></td>";
                    GridList.Text += "      </tr>";
                    GridList.Text += "      <tr>";
                    GridList.Text += "        <td align='left' valign='top' class='Invoice_No'>捐贈日期:" + dr["捐贈日期"].ToString() + "</td>";
                    GridList.Text += "        <td align='right' valign='top' class='Invoice_No'>收據編號:" + dr["收據編號"].ToString() + "</td>";
                    GridList.Text += "        <td align='center' valign='top' class='Invoice_No'></td>";
                    GridList.Text += "      </tr>";
                    GridList.Text += "      <tr>";
                    GridList.Text += "        <td align='center' valign='top' style='height:40mm' colspan='2'>";
                    GridList.Text += "          <table width='100%' border='1' cellpadding='0' cellspacing='0' bordercolor='#000000' style='border-collapse: collapse'>";
                    GridList.Text += "            <tr>";
                    GridList.Text += "              <td style='width:25mm;height:8mm' class='Donate_Desc' align='center'>捐贈者</td>";
                    GridList.Text += "              <td style='width:110mm;height:8mm' class='Donate_Desc' align='left'>&nbsp;" + dr["捐贈者"].ToString() + "</td>";
                    GridList.Text += "              <td style='width:30mm;height:8mm' class='Donate_Desc' align='center'>身份證/統編</td>";
                    GridList.Text += "              <td style='width:29mm;height:8mm' class='Donate_Desc'>&nbsp;" + dr["身份證/統編"].ToString() + "</td>";
                    GridList.Text += "            </tr>";
                    GridList.Text += "            <tr>";
                    GridList.Text += "              <td style='width:25mm;height:8mm' align='center' class='Donate_Desc'>收據地址</td>";
                    GridList.Text += "              <td style='width:169mm;height:8mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;" + dr["收據地址"].ToString() + "</td>";
                    GridList.Text += "            </tr>";
                    GridList.Text += "            <tr>";
                    GridList.Text += "              <td style='width:25mm;height:16mm' align='center' class='Donate_Desc'>捐贈內容</td>";
                    GridList.Text += "              <td style='width:169mm;height:16mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;" + dr["捐贈內容"].ToString() + "</td>";
                    GridList.Text += "            </tr>";
                    if (dr["折合現金"].ToString() != "0")
                    {
                        string Contribute_Amt = GetMoneyUpper(Decimal.Parse(dr["折合現金"].ToString().Replace(",", "") + ".00"));
                        GridList.Text += "            <tr>";
                        GridList.Text += "              <td style='width:25mm;height:8mm' align='center' class='Donate_Desc'>折合現金</td>";
                        GridList.Text += "              <td style='width:169mm;height:8mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;新台幣" + Contribute_Amt + "&nbsp;(NT$" + dr["折合現金"].ToString() + ")</td>";
                        GridList.Text += "            </tr>";
                    }
                    GridList.Text += "            <tr>";
                    GridList.Text += "              <td style='width:25mm;height:8mm' align='center' class='Donate_Desc'>備註</td>";
                    GridList.Text += "              <td style='width:169mm;height:8mm' align='left' class='Donate_Desc' colspan='3'>&nbsp;" + dr["備註"].ToString() + "</td>";
                    GridList.Text += "            </tr>";
                    GridList.Text += "          </table>";
                    GridList.Text += "        </td>";
                    if (i == 0)
                    {
                        GridList.Text += "      <td class='Donate_Right' align='center' valign='center'>第<br />一<br />聯<br />：<br>捐<br>助<br>人<br>留<br>存</td>";
                    }
                    else if (i == 1)
                    {
                        GridList.Text += "      <td class='Donate_Right' align='center' valign='center'>第<br />二<br />聯<br />：<br>會<br>計<br>聯</td>";
                    }
                    else if (i == 2)
                    {
                        GridList.Text += "      <td class='Donate_Right' align='center' valign='center'>第<br />三<br />聯<br />：<br>存<br>根<br>聯</td>";
                    }
                    GridList.Text += "      </tr>";
                    GridList.Text += "      <tr>";
                    GridList.Text += "        <td align='center' valign='top' style='height:10mm' colspan='3'>";
                    GridList.Text += "          <table width='100%' border='0' cellpadding='0' cellspacing='0'>";
                    GridList.Text += "            <tr>";
                    GridList.Text += "              <td style='width:16mm;height:10mm' align='left' class='Donate_Seal'>理事長:</td>";
                    GridList.Text += "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'></td>";
                    GridList.Text += "              <td style='width:12mm;height:10mm' align='left' class='Donate_Seal'>覆核:</td>";
                    GridList.Text += "              <td style='width:50mm;height:10mm' align='left' class='Donate_Seal'></td>";
                    GridList.Text += "              <td style='width:19mm;height:10mm' align='left' class='Donate_Seal'>經手人:</td>";
                    GridList.Text += "              <td style='width:47mm;height:10mm' align='left' class='Donate_Seal'></td>";
                    GridList.Text += "            </tr>";
                    GridList.Text += "          </table>";
                    GridList.Text += "        </td>";
                    GridList.Text += "      </tr>";
                    if (i == 0)
                    {
                        GridList.Text += "    <tr>";
                        GridList.Text += "      <td align='center' valign='top' style='height:18mm' colspan='3'>";
                        GridList.Text += "        <table width='100%' border='0' cellpadding='0' cellspacing='0'>";
                        GridList.Text += "          <tr>";
                        GridList.Text += "            <td style='width:64mm;height:8mm' align='left' class='Donate_Account'>劃撥帳號:" + dr_Organization["Account"].ToString() + "</td>";
                        GridList.Text += "            <td style='width:130mm;height:8mm' align='left' class='Donate_Account' colspan='2'>戶名:" + dr_Organization["OrgName"].ToString() + "</td>";
                        GridList.Text += "          </tr>";
                        GridList.Text += "          <tr>";
                        GridList.Text += "            <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>電話:" + dr_Organization["Tel"].ToString() + "</td>";
                        GridList.Text += "            <td style='width:55mm;height:5mm' align='left' class='Donate_Foot'>傳真:" + dr_Organization["Fax"].ToString() + "</td>";
                        GridList.Text += "            <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>地址:" + dr_Organization["Address"].ToString() + "</td>";
                        GridList.Text += "          </tr>";
                        GridList.Text += "          <tr>";
                        GridList.Text += "            <td style='width:64mm;height:5mm' align='left' class='Donate_Foot'>EMail:" + dr_Organization["Email"].ToString() + "</td>";
                        GridList.Text += "            <td style='width:55mm;height:5mm' align='left' class='Donate_Foot'>統一編號:" + dr_Organization["Uniform_No"].ToString() + "</td>";
                        GridList.Text += "            <td style='width:75mm;height:5mm' align='left' class='Donate_Foot'>立案字號:北府社老字第1001825306號</td>";
                        GridList.Text += "          </tr>";
                        GridList.Text += "        </table>";
                        GridList.Text += "      </td>";
                        GridList.Text += "    </tr>";
                    }
                    GridList.Text += "    </table>";
                    GridList.Text += "  </td>";
                    GridList.Text += "</tr>";
                    if (i < 2)
                    {
                        GridList.Text += "<tr>";
                        GridList.Text += "  <td class='Line'> </td>";
                        GridList.Text += "</tr>";
                        GridList.Text += "<tr>";
                        GridList.Text += "  <td class='CellMiddle'> </td>";
                        GridList.Text += "</tr>";
                    }
                }
                GridList.Text += "</table>";
                if (Row - 1 < dt.Rows.Count)
                {
                    GridList.Text += "<div class='pagebreak'>&nbsp;</div>";
                    Row += 1;
                }
                if (Row - 1 == dt.Rows.Count)
                {
                    end = true;
                }
            }
        }
        else
        {
            GridList.Text = "** 無符合的資料 **";
        }
    }
    public void Update(string Contribute_Id)
    {
        Dictionary<string, object> dict = new Dictionary<string, object>();
        //****設定SQL指令****//
        string strSql = " update Contribute set ";
        strSql += " Invoice_Print= @Invoice_Print";
        strSql += " where Contribute_Id = @Contribute_Id";
        dict.Add("Invoice_Print", "1");
        dict.Add("Contribute_Id", Contribute_Id);
        NpoDB.ExecuteSQLS(strSql, dict);
    }
    public static string GetMoneyUpper(decimal d)
    {
        string o = d.ToString();
        string dq = "", dh = "";
        if (o.Contains("."))
        {
            dq = o.Split('.')[0];
            dh = o.Split('.')[1];
        }
        else
        {
            dq = o;
        }
        string ret = GetMoney(dq, true, "圓") + GetMoney(dh, false, "");
        if (ret.Contains("厘") || ret.Contains("分"))
            return ret;
        return ret + "整";
    }
    private static string GetMoney(string number, bool left, string lastdw)
    {
        string[] NTD = new string[10] { "零", "壹", "貳", "參", "肆", "伍", "陸", "柒", "捌", "玖" };
        string[] DW = new string[8] { "厘", "分", "角", "", "拾", "佰", "仟", "萬" };
        int first = 4;
        string str = "";
        if (!left)
        {
            first = 1;
            if (number.Length == 1)
            {
                number += "00";
            }
            else if (number.Length == 2)
            {
                number += "0";
            }
            else number = number.Substring(0, 3);

        }
        else
        {
            if (number.Length >= 9)
            {
                return GetMoney(number.Substring(0, number.Length - 8), true, "億") + GetMoney(number.Substring(number.Length - 8, 8), true, "圓");
            }
            if (number.Length >= 5)
            {
                return GetMoney(number.Substring(0, number.Length - 4), true, "萬") + GetMoney(number.Substring(number.Length - 4, 4), true, "圓");
            }
        }
        bool has0 = false;
        for (int i = 0; i < number.Length; ++i)
        {
            int w = number.Length - i + first - 2;
            if (int.Parse(number[i].ToString()) == 0)
            {
                has0 = true;
                continue;
            }
            else
            {
                if (has0)
                {
                    str += "零";
                    has0 = false;
                }
            }
            str += NTD[int.Parse(number[i].ToString())];
            str += DW[w];
        }
        if (left)
            str += lastdw;
        return str;
    }
}