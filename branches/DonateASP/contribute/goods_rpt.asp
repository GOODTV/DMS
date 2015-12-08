<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=物品主檔維護.xls"
End If

Function DonateList (SQL,ReportName)
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=5 Then
        Response.Write "<td bgcolor='#FFFFFF' align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(i),0) & "</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
	    End If
	  Next
	  Response.Flush
    Response.Clear
    RS1.MoveNext
	  Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function 
%>
<%Prog_Id="goods"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL="Select Goods_Id,物品代號=Goods_Id,物品名稱=Goods_Name,物品性質=Goods_Property,物品類別=Goods_Type,現有庫存量=Goods_Qty,庫存單位=Goods_Unit,庫存管理=(Case When Goods_IsStock='Y' Then '是' Else '否' End) " & _
      "From GOODS Join LINKED On GOODS.Goods_Property=LINKED.Linked_Name Join LINKED2 On GOODS.Goods_Type=LINKED2.Linked2_Name Where LINKED.Linked_Type='contribute' "
  If request("Goods_Id")<>"" Then SQL=SQL&"And Goods_Id Like '%"&request("Goods_Id")&"%' "
  If request("Goods_Name")<>"" Then SQL=SQL&"And Goods_Name Like '%"&request("Goods_Name")&"%' "
  If request("Goods_Property")<>"" Then SQL=SQL&"And Goods_Property Like '%"&request("Goods_Property")&"%' "
  If request("Goods_Type")<>"" Then SQL=SQL&"And Goods_Type Like '%"&request("Goods_Type")&"%' "
  If request("Goods_IsStock")<>"" Then SQL=SQL&"And Goods_IsStock='"&request("Goods_IsStock")&"' "
  If request("Goods_Qty")<>"" Then SQL=SQL&"And Goods_Qty>0 "
  SQL=SQL&"Order By LINKED.Linked_Seq,Linked2_Seq,Goods_Id Desc"
  call DonateList (SQL,Prog_Desc)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->