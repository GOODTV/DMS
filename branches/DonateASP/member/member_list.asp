<!--#include file="../include/dbfunctionJ.asp"-->
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>讀者資料維護</title>
  <%If request("action")<>"export" Then%><link REL="stylesheet" type="text/css" HREF="../include/dms.css"><%End If%>
</head>
<body class=tool>
<%
If request("action")="" and request("SQL") ="" Then
	Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
  Response.Write "  <tr>"
  Response.Write "    <td width='100%' align='center' style='color:#ff0000'>** 請先輸入查詢條件 **</td>"	  
  Response.Write "  </tr>"
  Response.Write "</table>" 
ElseIf request("SQL")="" Then 
  SQL="Select Distinct Donor_Id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End),讀者姓名=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),聯絡電話=(Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End),通訊地址=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End),電子報=(Case When IsSendEpaper='Y' Then 'V' Else '' End) " & _
      "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember='Y' "
  If request("Donor_Name")<>"" Then SQL=SQL & "And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Member_No")<>"" Then SQL=SQL & "And Donor_Id Like '%"&request("Member_No")&"%' "
  If request("Tel_Office")<>"" Then SQL=SQL & "And (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
  If request("Category")<>"" Then SQL=SQL & "And Category = '"&request("Category")&"' "
  If request("Donor_Type")<>"" Then SQL=SQL & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
  If request("Member_Type")<>"" Then SQL=SQL & "And Member_Type='"&request("Member_Type")&"' "
  If request("Member_Status")<>"" Then SQL=SQL & "And Member_Status='"&request("Member_Status")&"' "
  If request("City")<>"" Then SQL=SQL & "And (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
  If request("Area")<>"" Then SQL=SQL & "And (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
  If request("Address")<>"" Then SQL=SQL & "And (Address Like '%"&request("Address")&"%' Or Invoice_Address Like '%"&request("Address")&"%') "
  If request("IsAbroad")<>"" Then SQL=SQL & "And (IsAbroad = 'Y' Or IsAbroad_Invoice = 'Y') "
  If request("IsSendNews")<>"" Then SQL=SQL & "And IsSendNews = 'Y' "
  If request("IsSendEpaper")<>"" Then SQL=SQL & "And IsSendEpaper = 'Y' "
  If request("IsSendYNews")<>"" Then SQL=SQL & "And IsSendYNews = 'Y' "
  If request("IsBirthday")<>"" Then SQL=SQL & "And IsBirthday = 'Y' "
  If request("IsXmas")<>"" Then SQL=SQL & "And IsXmas = 'Y' "
  SQL=SQL&"Order By (Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End) Desc,Donor_Id Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="member_list"
HLink="member_edit.asp?donor_id="
LinkParam="donor_id"
LinkTarget="main"
AddLink="member_add.asp?donor_name2="&request("Donor_Name")&"&tel_office2="&request("Tel_Office")&"&category2="&request("Category")&"&donor_type2="&request("Donor_Type")&"&idno2="&request("IDNo")&"&city2="&request("City")&"&area2="&request("Area")&"&address2="&request("Address")&"&isabroad2="&request("IsAbroad")&"&issendnews2="&request("IsSendNews")&"&issendepaper2="&request("IsSendEpaper")&"&issendynews2="&request("IsSendYNews")&""
If request("action")="stop" Then
  call GridList_S (AddLink)
Else
  If request("action")="report" Or request("action")="export" Then Server.Transfer "member_rpt.asp"
  If request("action")<>"" Or SQL<>"" Then call GridList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->