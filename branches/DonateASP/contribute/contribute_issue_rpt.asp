<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=contribute_issue.xls"
End If

Function ContributeIssueList (SQL,ReportName)
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
    If I=4 Then Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>領用內容</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=4 Then
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(i) & "</span></td>"
        Goods_Name=""
        SQL2="Select Goods_Name,Goods_Qty,Goods_Unit From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&RS1(0)&"' Order By Ser_No"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,1
        If Not RS2.EOF Then
          While Not RS2.EOF
            If Goods_Name="" Then
              Goods_Name=RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            Else
              Goods_Name=Goods_Name&"、"&RS2("Goods_Name")&"("&FormatNumber(RS2("Goods_Qty"),0)&RS2("Goods_Unit")&")"
            End If
            RS2.MoveNext
          Wend
          Goods_Name=Goods_Name&"。"  
        End If
        RS2.Close
        Set RS2=Nothing
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Goods_Name & "</span></td>"
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
<%Prog_Id="issue"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  SQL="Select Issue_Id,領取人=Issue_Processor,領取日期=CONVERT(nvarchar,Issue_Date,111),領取用途=Issue_Purpose,領取編號=Issue_Pre+Issue_No,列印=(Case When Issue_Print='1' Then 'V' Else '' End),狀態=(Case When Issue_Type='M' Then '手開' Else Case When Issue_Type='D' Then '作廢' Else '' End End),經手人=Create_User From CONTRIBUTE_ISSUE Where 1=1 "
  If request("Dept_Id")<>"" Then SQL=SQL&"And Dept_Id = '"&request("Dept_Id")&"' "
  If request("Issue_Processor")<>"" Then SQL=SQL&"And (Issue_Processor Like '%"&request("Issue_Processor")&"%') "
  If request("Issue_Purpose")<>"" Then SQL=SQL&"And Issue_Purpose = '"&request("Issue_Purpose")&"' "
  If request("Create_User")<>"" Then SQL=SQL&"And Create_User = '"&request("Create_User")&"' "
  If request("Issue_Date_B")<>"" Then SQL=SQL&"And Issue_Date >= '"&request("Issue_Date_B")&"' "
  If request("Issue_Date_E")<>"" Then SQL=SQL&"And Issue_Date <= '"&request("Issue_Date_E")&"' "
  If request("Issue_No_B")<>"" Then SQL=SQL&"And Issue_No >= '"&request("Issue_No_B")&"' "
  If request("Issue_No_E")<>"" Then SQL=SQL&"And Issue_No <= '"&request("Issue_No_E")&"' "
  If request("Issue_TypeM")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeM")&"' "
  If request("Issue_TypeD")<>"" Then SQL=SQL&"And Issue_Type = '"&request("Issue_TypeD")&"' "  
  SQL=SQL&"Order By CONVERT(VarChar,Issue_Date,111) Desc,Issue_Id Desc"
  ReportName="物品領用明細表"
  call ContributeIssueList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->