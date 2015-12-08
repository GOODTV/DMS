<!--#include file="../include/dbfunction.asp"-->
<%
SQL="select issue.*,issuedetail.goods_id,goods_name,goods_unit,issue_qty,Issue_Purpose,issue_Comment,issue_org from issue left join issuedetail on issuedetail.issue_id=issue.issue_id"_
+"       where issue.issue_id='"&request("issue_id")&"' order by ser_no"
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL,Conn,1,1
if Not RS1.EOF then
   Processor=RS1("Processor")
   issue_date=RS1("issue_date")
end if
i = 0
k = 0
pageno=1
call heading()
%>
<html>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 6.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>物品出貨單</title>
<link REL="stylesheet" type="text/css" HREF="../include/dms.css">
<style>
.fonts
{
    FONT-FAMILY: 標楷體;
    FONT-SIZE: 12pt;
}
.fonts1
{
    FONT-FAMILY: 標楷體;
    FONT-SIZE: 12pt;
}
.PageBreak { page-break-after:always; }
</style>
</head>

<body class=tool>
<%sub heading()%>
<div align="center">
  <center>
&nbsp;<div align="center">
	<table border="0" width="720" id="table2" cellspacing="0" cellpadding="4">
		<tr>
			<td align="center" colspan="2">
  <span style="font-size: 14pt;font-family: 標楷體"><%=session("comp_shortname")%> 
	物品出貨單</span></td>
		</tr>
		<tr>
			<td width="523" class="fonts">領用編號：<%=request("issue_id")%></td>
			<td width="403" class="fonts">領用日期：<%=year(issue_date)%>年<%=month(issue_date)%>月<%=day(issue_date)%>日
  			</td>
		</tr>
		<tr>
			<td width="523" class="fonts">用途：<%=rs1("issue_purpose")%></td>
			<td width="403" class="fonts">出貨單位：<%=rs1("issue_org")%></td>
		</tr>
	</table>
</div>
  <div align="center">
  <table border="2" bordercolor="#111111" cellspacing="0" cellpadding="1">
  <tr>
  <td>
  <table border="1" width="720" id="table1" bordercolor="#808080" cellspacing="0" cellpadding="2" class="fonts">
	<tr>
		<td width="72" align="center" height="26" class="fonts" bgcolor="#E5ECFF">
		編號</td>
		<td width="296" height="26" class="fonts" bgcolor="#E5ECFF">
		物品名稱</td>
		<td align="center" height="26" width="127" class="fonts" bgcolor="#E5ECFF">
		數量</td>
		<td align="center" height="26" width="194" class="fonts" bgcolor="#E5ECFF">
		備註</td>
	</tr>
	<%end sub%>
	<%While Not rs1.EOF
	  i = i + 1
	  k = k + 1
	  %>
	<tr>
		<td width="72" align="center" height="26" class="fonts1"><%=rs1("goods_id")%>　</td>
		<td width="296" height="26" class="fonts1"><%=rs1("goods_name")%></td>
		<td height="26" width="127" align="center" class="fonts1"><%=rs1("issue_qty")%> <%=rs1("goods_unit")%></td>
		<td height="26" width="194" class="fonts1"><%=rs1("issue_Comment")%></td>
	</tr>
	<%rs1.MoveNext
	  if k = 30 and Not rs1.EOF then
	     k = 0
	     pageno=pageno + 1%>
   </table>
	</td>
	</tr>	     
   </table>
	<div align="center" class="pagebreak">
	</div>  
	 <%call heading%>
     <%end if
	  Wend%>
	<%if i < 30 then
	     for j = 1 to 30 - i%>  
	<tr>
		<td width="72" align="center" height="26" class="fonts1"></td>
		<td width="296" height="26" class="fonts1"></td>
		<td height="26" width="127" class="fonts1"></td>
		<td height="26" width="194" class="fonts1"></td>
	</tr>
	<%  Next
	  end if%>
	</table>
	</td>
	</tr>
	</table>
	<table border="0" width="720" id="table2" cellspacing="0" cellpadding="4">
		<tr>
			<td width="60%">
  			<span style="font-size: 12pt;font-family: 標楷體">取貨人：</span></td>
  			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="60%">
  			<span style="font-size: 12pt;font-family: 標楷體">領用人：</span></td>
  			<td>&nbsp;</td>
		</tr>
		<tr>
			<td width="60%">
  			<span style="font-size: 12pt;font-family: 標楷體">主管：</span></td>
  			<td>&nbsp;</td>
		</tr>
	</table>
</div>
</body>

</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->