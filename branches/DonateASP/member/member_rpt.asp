<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "big5"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=會員資料.xls"
End If 
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>會員資料維護</title>
</head>
<body class=tool>
<%
  SQL="Select Distinct Donor_Id,編號=(Case When Member_No<>'' Then Member_No Else CONVERT(nvarchar(20),Donor_Id) End),會員姓名=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),會員別=Member_Type,會員別=Volunteer_Status,聯絡電話=(Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End),通訊地址=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End),會訊=(Case When IsSendNews='Y' Then 'V' Else '' End),電子報=(Case When IsSendEpaper='Y' Then 'V' Else '' End),年報=(Case When IsSendYNews='Y' Then 'V' Else '' End),生日卡=(Case When IsBirthday='Y' Then 'V' Else '' End),賀卡=(Case When IsXmas='Y' Then 'V' Else '' End),最近捐款日期=CONVERT(VarChar,Last_DonateDate,111) " & _
      "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember='Y' "
  If request("Donor_Name")<>"" Then SQL=SQL & "And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Tel_Office")<>"" Then SQL=SQL & "And (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
  If request("Category")<>"" Then SQL=SQL & "And Category = '"&request("Category")&"' "
  If request("Donor_Type")<>"" Then SQL=SQL & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
  If request("Member_Type")<>"" Then SQL=SQL & "And Member_Type='"&request("Member_Type")&"' "
  If request("Member_Status")<>"" Then SQL=SQL & "And Member_Status='"&request("Member_Status")&"' "
  If request("City")<>"" Then SQL=SQL & "And (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
  If request("Area")<>"" Then SQL=SQL & "And (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
  If request("Address")<>"" Then SQL=SQL & "And (Address Like '%"&request("Address")&"%' Or Invoice_Address Like '%"&request("Address")&"%') "
  If request("IsSendNews")<>"" Then SQL=SQL & "And IsSendNews = 'Y' "
  If request("IsSendEpaper")<>"" Then SQL=SQL & "And IsSendEpaper = 'Y' "
  If request("IsSendYNews")<>"" Then SQL=SQL & "And IsSendYNews = 'Y' "  
  SQL=SQL&"Order By Donor_Id Desc"
  ReportName="會員資料明細表"
  call ReportList (SQL,ReportName)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->