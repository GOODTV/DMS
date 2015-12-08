using System;
using System.Web.UI;
using System.Data;
using System.Text;
using System.Collections.Generic;
using System.IO ;
using System.Web;
using System.Web.UI.WebControls;


public partial class Function_image_left_list : BasePage
{
    protected void Page_Load(object sender, EventArgs e)
    {
        ShowSysMsg();
        //LTL_Button.Text = "<a href target=\"parent.parent.parent.parent\" style=\"text-decoration: none\" OnClick=\"location.href='../content/content.aspx';\"><button style=\"cursor:hand;\" class=\"addbutton\">回目錄選單頁</button></a>";
        LTL_Button.Text += "<a href style=\"text-decoration: none\" onClick=\"window.open('image_left_upload.aspx?dept_id=" + Util.GetQueryString("dept_id") + "&item=" + Util.GetQueryString("item") + "&ser_no=" + Util.GetQueryString("ser_no") +"&code_id=" + Util.GetQueryString("code_id") + "&id=" + Util.GetQueryString("id") + "','','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=170')\"><button style=\"cursor:hand;\" class=\"addbutton\">圖文繞圖檔上傳</button></a>";

        string sqlstr = "select uid,圖片=UploadFileURL from UpLoad where Apname=@item and attachtype='img'";
        Dictionary<string, object> dict = new Dictionary<string, object>();
        dict.Add("item",  Util.GetQueryString("item"));        
        DataTable dt = NpoDB.GetDataTableS(sqlstr, dict);

        if (dt.Rows.Count>0)
        {
            ltlGridView.Text = "<br/><table id=grid border=0 cellspacing='0' cellpadding='3' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='110%' align='center'>";
            ltlGridView.Text += "<tr>";
            ltlGridView.Text += "    <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>圖片</span></font></td>";
            ltlGridView.Text += "    <td bgcolor='#FFE1AF'><span style='font-size: 9pt; font-family: 新細明體'></span></td>";
            ltlGridView.Text += "</tr>";
            foreach(DataRow dr in dt.Rows)
            {
                ltlGridView.Text += "<tr>";
                ltlGridView.Text += "<td bgcolor='#FFFFFF' align='center'><img src='../upload/" + dr["圖片"].ToString() + "' border=0 width='90' height='67' onClick=\"window.open(\'image_left_show.aspx?imgfile=" + dr["圖片"].ToString() + "','','scrollbars=yes,resizable=yes,width=800,height=600')\"></td>";
                ltlGridView.Text += "<td valign='bottom'><a href=\"JavaScript:if(confirm('是否確定要刪除圖片 ?')){window.location.href='image_left_list.aspx?act=idel&dept_id="+SessionInfo.DeptID+"&id="+ Util.GetQueryString("id") +"&code_id="+ Util.GetQueryString("code_id") +"&item="+ Util.GetQueryString("item") +"&ser_no="+ dr["uid"].ToString() +"';}\" target=\"_self\"><img src='../images/x1.gif' border=0 width='16' height='14' alt='刪除'></a></td>";
                ltlGridView.Text += "</tr>";
            }
            ltlGridView.Text += "</table>";
        }

        if (Util.GetQueryString("act") == "idel")
        { 
            string Object_ID="";
            string Upload_FileURL="";
            //string Upload_FileURL_Old="";
            string folderPath = Util.GetAppSetting("UploadPath");

            //sqlstr = "Select * From UPLOAD Where ser_no=@ser_no";
            sqlstr = "Select * From UPLOAD Where uid=@ser_no";
            dict.Clear();
            dict.Add("ser_no",  Util.GetQueryString("ser_no"));
            dt = NpoDB.GetDataTableS(sqlstr, dict);
            if (dt.Rows.Count>0)
            {
                Object_ID=dt.Rows[0]["ObjectID"].ToString();
                Upload_FileURL=dt.Rows[0]["UploadFileURL"].ToString();
                //Upload_FileURL_Old=dt.Rows[0]["Upload_FileURL_Old"].ToString();
            }
            
            //sqlstr = "Delete From UPLOAD Where ser_no=@ser_no";
            sqlstr = "Delete From UPLOAD Where uid=@ser_no";
            dict.Clear();
            dict.Add("ser_no",  Util.GetQueryString("ser_no"));
            NpoDB.ExecuteSQLS(sqlstr,dict);
            
            if (Upload_FileURL.Trim() != "")
            {
                string DataFile = Server.MapPath("~" + folderPath) + @"/" + Upload_FileURL;
                if (File.Exists(DataFile))
                {
                    File.Delete(DataFile);
                }                
            }

            //if (Upload_FileURL_Old.Trim() != "")
            //{
            //    string DataFile = Server.MapPath("~" + folderPath) + @"/" + Upload_FileURL_Old;
            //    if (File.Exists(DataFile))
            //    {
            //        File.Delete(DataFile);
            //    }  
            //}
  
            Session["Msg"]="圖檔刪除成功 ！";
            Response.Redirect(Util.RedirectByTime("image_left_list.aspx","code_id="+ Util.GetQueryString("code_id") +"&item="+ Util.GetQueryString("item") +"&ser_no="+ Object_ID));

        }

        if (!IsPostBack)
        {
        }        
    }
    //------------------------------------------------------------------------------     
}
