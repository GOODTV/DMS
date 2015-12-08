<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  SQL="CODEFILE"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS.Addnew
  RS("CodeType")=request("CodeType")
  RS("CodeName")=request("CodeName")
  RS("Category")=request("Category")
  RS("Menu_Url")=request("Menu_Url")
  RS("Image_Url")=request("Image_Url")
  If request("Image_Width")<>"" Then  
    RS("Image_Width")=request("Image_Width")
  Else
    RS("Image_Width")=null
  End If  
  RS("Image_Sub_Url")=request("Image_Sub_Url")
  If request("Image_Sub_Width")<>"" Then  
    RS("Image_Sub_Width")=request("Image_Sub_Width")
  Else
    RS("Image_Sub_Width")=null
  End If
  If request("Menu_Width")<>"" Then
    RS("Menu_Width")=request("Menu_Width")
  Else
    RS("Menu_Width")="0"
  End If
  If request("Menu_Seq")<>"" Then  
    RS("Menu_Seq")=request("Menu_Seq")
  Else
    RS("Menu_Seq")=null
  End If     
  RS("CodeTable")=request("CodeTable")
  RS("CodeFile")=request("CodeFile")
  RS("CodeWhere")=request("CodeWhere")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="代碼主檔新增成功，\n\n請繼續新增『 代碼選項 』 ！"
  response.redirect "codefile.asp?codetype="&request("CodeType") 
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>新增代碼主檔</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr> 
                <td width="13%" height=22 align="right" noWrap><font color="#000080">代碼：</font></td>
                <td width="28%" height=22><input type="text" name="CodeType" size="16" class="font9" maxlength="20"></td>
                <td width="19%" height=22 align="right" noWrap><font color="#000080">代碼名稱：</font></td>
                <td width="40%" height=22><input type="text" name="CodeName" size="25" class="font9" maxlength="40"></td>
              </tr>
              <tr> 
                <td noWrap align="right"><font color="#000080">類別：</font></td>
                <td> 
                  <select size="1" name="Category">
		                <option value="類別選項">類別選項</option>
		                <option value="單一網頁">單一網頁</option>
		                <option value="條列網頁">條列網頁</option>
		              </select>
		            </td>
		            <td noWrap align="right"><font color="#000080">連結頁面：</font></td>
		            <td><input type="text" name="Menu_Url" size="25" class="font9" maxlength="40"></td>
              </tr>
              <tr> 
                <td align="right"><font color="#000080">主選單圖檔</font></td>
                <td><input type="text" name="Image_Url" size="16" class="font9" maxlength="200"></td>
                <td align="right"><font color="#000080">主選單圖寬：</font></td>
                <td><input type="text" name="Image_Width" size="25" class="font9" maxlength="200"></td>
              </tr>
              <tr> 
                <td align="right"><font color="#000080">次選單圖檔</font></td>
                <td><input type="text" name="Image_Sub_Url" size="16" class="font9" maxlength="200"></td>
                <td align="right"><font color="#000080">次選單圖寬：</font></td>
                <td><input type="text" name="Image_Sub_Width" size="25" class="font9" maxlength="200"></td>
              </tr>                
              <tr> 
                <td align="right"><font color="#000080">下拉選單寬</font></td>
                <td><input type="text" name="Menu_Width" size="16" class="font9" maxlength="3"></td>
                <td align="right"><font color="#000080">排序：</font></td>
                <td><input type="text" name="Menu_Seq" size="25" class="font9" maxlength="3"></td>
              </tr>      
              <%If Session("Menu_Edit")="Y" Then%>
              <tr> 
                <td align="right"><font color="#000080">資料表<span lang="en-us">：</span><br>(TableName)</font></td>
                <td colspan="3"><input type="text" name="CodeTable" size="58" class="font9" maxlength="200"></td>
              </tr>
              <tr> 
                <td align="right"><font color="#000080">欄位<span lang="en-us">：</span><br>(FileName)</font></td>
                <td colspan="3"><input type="text" name="CodeFile" size="58" class="font9" maxlength="200"></td>
              </tr>
              <tr> 
                <td align="right"><font color="#000080">條件<span lang="en-us">：</span><br>(Where)</font></td>
                <td colspan="3"><input type="text" name="CodeWhere" size="58" class="font9" maxlength="200"></td>
              </tr>
              <tr> 
                <td align="right"></td>
                <td colspan="3"><font color="#ff0000">(二個以上資料表請用&nbsp;/&nbsp;區隔)</font></td>
              </tr>
              <%End If%>                    
              <tr>
                <td width="100%" height=15 colspan="4" align="center">
                  <button type='button' id='save' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='Save_OnClick()'> <img src='../images/save.gif' width='19' height='20' align='absmiddle'> 存檔</button>&nbsp;
                  <button type='button' id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='window.close();'> <img src='../images/icon6.gif' width='20' height='15' align='absmiddle'> 離開</button>
                </td>
              </tr>   
            </table>
          </td>
        </tr>
      </table>
    </center></div>
  </form>
  <%Message()%> 
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("CodeType","代碼")%>
  <%call ChecklenJ("CodeType",20,"代碼")%>
  <%call CheckStringJ("CodeName","代碼名稱")%>
  <%call ChecklenJ("CodeName",40,"代碼名稱")%>
  <%call ChecklenJ("Menu_Url",40,"連結頁面")%>
    <%call ChecklenJ("Image_Url",200,"主選單圖檔")%>
  if(document.form.Image_Width.value!=''){
    <%call CheckNumberJ("Image_Width","主選單圖寬")%>
  }  
  <%call ChecklenJ("Image_Sub_Url",200,"次選單圖檔")%>  
  if(document.form.Image_Sub_Width.value!=''){
    <%call CheckNumberJ("Image_Sub_Width","次選單圖寬")%>
  }
  if(document.form.Menu_Width.value!=''){
    <%call CheckNumberJ("Menu_Width","下選單寬(div)")%>
  }
  if(document.form.Menu_Seq.value!=''){
    <%call CheckNumberJ("Menu_Seq","排序")%>
  }
  <%call ChecklenJ("CodeTable",200,"資料表")%>
  <%call ChecklenJ("CodeFile",200,"欄位")%>
  <%call ChecklenJ("CodeWhere",200,"條件")%>
  <%call SubmitJ("save")%>
  window.close();
}
--></script>