<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  If request("item")="album_photo" Then
    SQL="ALBUM_PHOTO"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Album_Id")=request("ser_no")
    RS("Album_Date")=request("albumdate")
    RS("Album_Desc")=request("Upload_FileName")
    RS("Album_URL")=""
    RS("Album_FileSize")="0"
    RS("Album_URL_Old")=""
    RS("Album_FileSize_Old")="0"
    RS("Album_Location")=request("Upload_FileURL")
    RS("Album_Type")=request("AVType")
    RS("Album_Seq")=request("albumseq")
    RS.Update
    RS.Close
    Set RS=Nothing
    session("errnumber")=1
    session("msg")="影音連結成功 !"
    response.redirect request("item")&".asp?code_id="&request("ser_no")&"&ser_no="&request("ser_no")
  ElseIf request("item")="ads" Then
    SQL="Select * From ADS Where ser_no='"&request("ser_no")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS("Ads_TitleImg")=request("Upload_FileURL")
    RS.Update
    RS.Close
    Set RS=Nothing
    session("errnumber")=1
    session("msg")="影音連結成功 !"
    response.redirect request("item")&"_edit.asp?code_id="&request("ser_no")&"&ser_no="&request("ser_no")
  Else
    SQL="UPLOAD"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("Object_ID")=request("ser_no")
    RS("Ap_Name")=request("item")
    RS("Attach_Type")=request("AVType")
    RS("Upload_FileName")=request("Upload_FileName")
    RS("Upload_FileURL")=request("Upload_FileURL")
    RS("Upload_FileSize")="0"
    RS.Update
    RS.Close
    Set RS=Nothing
    session("errnumber")=1
    session("msg")="影音連結成功 !"
    response.redirect request("item")&"_edit.asp?code_id="&request("ser_no")&"&ser_no="&request("ser_no")
  End If
End If

album_seq=1
If request("item")="album_photo" Then
  SQL="Select IsNull(Max(Album_Seq),0) As Album_Seq From ALBUM_PHOTO Where Album_Id='"&request("ser_no")&"'"
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    album_seq=Cstr(Cint(RS("Album_Seq"))+1)
  End If
End If    
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>影音連結</title>
</head>
<body class=gray>
  <div align="center"><center>
    <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
      <tr>
        <td width="100%">
          <form name="form" action="" method="post" target="main">
            <input type="hidden" name="action">
            <input type="hidden" name="item" value="<%=request("item")%>">
            <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
            <input type="hidden" name="code_id" value="<%=request("code_id")%>">	
            <div align="center"><center>
              <table width="100%" border=1 cellspacing="0" style="border-collapse: collapse" cellpadding="2">
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">標題：</font></td>
                  <td width="80%" height=18><input type="text" name="Subject" size="60" class="font9t" readonly value="<%=request("Subject")%>"></td>
                </tr>
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">類型：</font></td>
                  <td width="80%" height=18> 
                  <%SQL="Select AVType=CodeDesc From CaseCode Where CodeType='AVtype' Order By seq"
                    FName="AVType"
                    Listfield="AVType"
                    BoundColumn=request("AVType")
                    NoChecked=1
                    call RadioBoxList (SQL,FName,Listfield,BoundColumn,NoChecked)
                  %>
                  </td>
                </tr>
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">連結網址：</font></td>
                  <td width="80%" height=18> <input type="text" name="Upload_FileURL" size="60" class="font9" maxlength="1000"></td>
                </tr>
                <%If request("item")<>"ads" Then%> 
                <tr> 
                  <td align=right width="20%" height=18><font color="#000080">連結說明：</font></td>
                  <td width="80%" height=18> <input type="text" name="Upload_FileName" size="60" class="font9" maxlength="200"></td>
                </tr>
                <%Else%>
                <input type="hidden" name="Upload_FileName">
                <%End If%>
                <%If request("item")="album_photo" Then%> 
                <tr> 
                  <td align=right height=18><font color="#000080">影音日期：</font></td>
                  <td height=18>
                    <%call Calendar("albumdate",date())%>
                    &nbsp;&nbsp;&nbsp;&nbsp;<font color="#000080">排序：</font>
                    <input type="text" name="albumseq" size="5" class="font9" maxlength="5" value="<%=album_seq%>">
                  </td>
                </tr>
                <!--#include file="../include/calendar2.asp"-->
                <%Else%>
                <input type="hidden" name="albumdate" value="<%=date()%>">
                <input type="hidden" name="albumseq" value="0">
                <%End If%>
                <tr>
                  <td width="100%" height=15 class="font9" colspan="2" align="center">
                    <input type="button" value=" 存 檔 " name="save" class="cbutton" onclick="Save_OnClick()">&nbsp;
                    <input type="button" value=" 取 消 " name="cancel" class="cbutton" onclick="window.close();">
                  </td>
                </tr>
              </table>
            </center></div>
          </form>
        </td>
      </tr>
    </table>
  </center></div>
  <%message%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("Upload_FileURL","連結網址")%>
  <%call ChecklenJ("Upload_FileURL",1000,"連結網址")%>
  <%call ChecklenJ("Upload_FileName",200,"連結說明")%>
  <%call CheckStringJ("albumseq","排序")%>
  <%call CheckNumberJ("albumseq","排序")%>  
  <%call SubmitJ("save")%>
  window.close();
}
--></script>