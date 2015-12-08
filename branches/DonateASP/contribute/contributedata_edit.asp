<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  Check_Update=True
  If request("Contribute_IsStock")="Y" Then
    SQL1="Select Goods_Qty,Goods_IsStock From GOODS Where Goods_Id='"&request("Goods_Id")&"' And Goods_IsStock='Y'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    If Not RS1.EOF Then
      If Cdbl(request("Goods_Qty"))>Cdbl(request("Goods_Qty_Old")) Then
        RS1("Goods_Qty")=Cdbl(RS1("Goods_Qty"))+Cdbl(Cdbl(request("Goods_Qty"))-Cdbl(request("Goods_Qty_Old")))
        RS1.Update
      ElseIf Cdbl(request("Goods_Qty_Old"))>Cdbl(request("Goods_Qty")) Then
        If Cdbl(Cdbl(request("Goods_Qty_Old"))-Cdbl(request("Goods_Qty")))<Cdbl(RS1("Goods_Qty")) Then
          RS1("Goods_Qty")=Cdbl(RS1("Goods_Qty"))-Cdbl(Cdbl(request("Goods_Qty_Old"))-Cdbl(request("Goods_Qty")))
          RS1.Update
        Else
          Check_Update=False
          session("errnumber")=1
          session("msg")="『"&request("Goods_Name")&"』 庫存量不足數量無法減少 ！"
        End If
      End If
    End If
    RS1.Close
    Set RS1=Nothing
  End If

  If Check_Update Then
    SQL1="Select * From CONTRIBUTEDATA Where ser_no='"&request("ser_no")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    If request("Contribute_IsStock")<>"Y" Then
      RS1("Goods_Name")=request("Goods_Name")
      RS1("Goods_Unit")=request("Goods_Unit")
    End If
    RS1("Goods_Qty")=request("Goods_Qty")
    If request("Goods_Amt")<>"" Then
      RS1("Goods_Amt")=request("Goods_Amt")
    Else
      RS1("Goods_Amt")="0"
    End If
    If request("Goods_DueDate")<>"" Then
      RS1("Goods_DueDate")=request("Goods_DueDate")
    Else
      RS1("Goods_DueDate")=null
    End If
    RS1("Goods_Comment")=request("Goods_Comment")
    RS1("Export")="N"
    RS1.Update
    RS1.Close
    Set RS1=Nothing
    
    SQL="Update CONTRIBUTE Set Contribute_Amt=(Select Isnull(Sum(Goods_Amt),0) From CONTRIBUTEDATA Where Contribute_Id='"&request("contribute_id")&"') Where Contribute_Id='"&request("contribute_id")&"'"
    Set RS=Conn.Execute(SQL)

    '修改捐贈人捐物紀錄
    call Declare_DonorId (request("donor_id"))
      
    session("errnumber")=1
    session("msg")="修改成功 ！" 
  End If
  response.redirect "contribute_edit.asp?ctype="&request("ctype")&"&contribute_id="&request("contribute_id")&""
End If

SQL="Select * From CONTRIBUTEDATA Where Ser_No="&request("ser_no")&""
Call QuerySQL(SQL,RS)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>捐贈物品修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=gray>
  <form name="form" method="post" action="" target="main">
    <input type="hidden" name="action">
    <input type="hidden" name="ctype" value="<%=request("ctype")%>">
    <input type="hidden" name="contribute_id" value="<%=request("contribute_id")%>">
    <input type="hidden" name="ser_no" value="<%=request("ser_no")%>">
    <input type="hidden" name="Donor_Id" value="<%=RS("Donor_Id")%>">	
    <input type="hidden" name="Goods_Id" value="<%=RS("Goods_Id")%>">
    <input type="hidden" name="Goods_Qty_Old" value="<%=RS("Goods_Qty")%>">	
    <input type="hidden" name="Contribute_IsStock" value="<%=RS("Contribute_IsStock")%>">	
    <div align="center"><center>
      <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
          <td width="100%">
            <table width="100%" border=1 cellspacing="0" cellpadding="2">
              <tr>
                <td align="right"><font color="#000080">物品名稱<span lang="en-us">：</span></font></td>
                <td colspan="3">
                	<input type="text" name="Goods_Name" size="45" maxlength="50" value="<%=RS("Goods_Name")%>" <%If RS("Contribute_IsStock")="Y" Then Response.Write "readonly class='font9t'" Else Response.Write "class='font9'" End If %> >
                  <%If RS("Contribute_IsStock")="Y" Then%><br /><font color="#ff0000">庫存物品不可更改物品名稱 / 單位</font><%End If %>
                </td>                
              </tr>
              <tr>
                <td width="18%" align="right"><font color="#000080">數量<span lang="en-us">：</span></font></td>
                <td width="32%"><input type="text" name="Goods_Qty" size="13" maxlength="7" value="<%=RS("Goods_Qty")%>" class="font9"></td>
                <td width="18%" align="right"><font color="#000080">單位<span lang="en-us">：</span></font></td>
                <td width="32%"><input type="text" name="Goods_Unit" size="13" maxlength="10" value="<%=RS("Goods_Unit")%>" <%If RS("Contribute_IsStock")="Y" Then Response.Write "readonly class='font9t'" Else Response.Write "class='font9'" End If %> ></td>	              
              </tr>
              <tr>
                <td align="right"><font color="#000080">折合現金<span lang="en-us">：</span></font></td>
                <td><input type="text" name="Goods_Amt" size="13" maxlength="7" value="<%=RS("Goods_Amt")%>" class="font9"></td>                	
                <td align="right"><font color="#000080">保存期限<span lang="en-us">：</span></font></td>
                <td><%call Calendar("Goods_DueDate",RS("Goods_DueDate"))%></td>  
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
  <%call CheckStringJ("Goods_Unit","單位")%>
  <%call ChecklenJ("Goods_Unit",10,"單位")%>
  if(document.form.Goods_Amt.value==''){
    document.form.Goods_Amt.value.value='0';
  }
  <%call CheckNumberJ("Goods_Amt","折合現金")%>
  <%call ChecklenJ("Goods_Amt",7,"折合現金")%>
  <%call ChecklenJ("Goods_Comment",100,"備註")%>      
  <%call SubmitJ("update")%>
  window.close();
}
--></script>