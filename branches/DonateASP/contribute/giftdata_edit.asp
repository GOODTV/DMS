<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL1="Select * From GIFTDATA Where ser_no='"&request("ser_no")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3  
  RS1("Goods_Qty")=request("Goods_Qty")   
  RS1("Goods_Comment")=request("Goods_Comment")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
    
  session("errnumber")=1
  session("msg")="修改成功 ！" 
 
  response.redirect "gift_edit.asp?ctype="&request("ctype")&"&gift_id="&request("gift_id")&""
End If

SQL="Select * From GIFTDATA Where Ser_No="&request("ser_no")&""
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="ctype" value="<%=request("ctype")%>">
    <input type="hidden" name="gift_id" value="<%=request("gift_id")%>">
    <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
    <input type="hidden" name="Donor_Id" value="<%=RS("Donor_Id")%>">	
    <input type="hidden" name="Goods_Id" value="<%=RS("Goods_Id")%>">
    <input type="hidden" name="Goods_Qty_Old" value="<%=RS("Goods_Qty")%>">	    
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td align="right"><font color="#000080">物品名稱<span lang="en-us">：</span></font></td>
                <td colspan="3">
                	<input type="text" name="Goods_Name" size="45" maxlength="50" value="<%=RS("Goods_Name")%>" readonly class="font9t">                  
                </td>                
              </tr>
              <tr>
                <td width="18%" align="right"><font color="#000080">數量<span lang="en-us">：</span></font></td>
                <td width="32%"><input type="text" name="Goods_Qty" size="13" maxlength="7" value="<%=RS("Goods_Qty")%>" class="font9"></td>                
              </tr>              
              <tr>
                <td align="right"><font color="#000080">備註<span lang="en-us">：</span></font></td>
                <td colspan="3"><input type="text" name="Goods_Comment" size="45" maxlength="100" value="<%=RS("Goods_Comment")%>" class="font9"></td>                
              </tr>
              <!--#include file="../include/calendar2.asp"-->
              <tr>
                <td width="100%" height=15 align="center" colspan="4">
                  <button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='Update_OnClick()'> <img src='../images/update.gif' width='19' height='20' align='absmiddle'> 修改</button>&nbsp;
                  <button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onclick='window.close();'> <img src='../images/icon6.gif' width='20' height='15' align='absmiddle'> 離開</button>
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
function Update_OnClick(){
  <%call CheckStringJ("Goods_Name","物品名稱")%>
  <%call ChecklenJ("Goods_Name",50,"物品名稱")%>
  <%call CheckStringJ("Goods_Qty","數量")%>
  <%call CheckNumberJ("Goods_Qty","數量")%>
  <%call ChecklenJ("Goods_Qty",7,"數量")%>  
  <%call ChecklenJ("Goods_Comment",100,"備註")%>      
  <%call SubmitJ("update")%>
  window.close();
}
--></script>