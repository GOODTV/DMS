<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="export" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=donor.xls"
ElseIf request("action")="mobile" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=mobile.xls"
ElseIf request("action")="email" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-excel"
  Response.AddHeader "content-disposition", "attachment;filename=email.xls"
End If

Function MobileList (SQL,ReportName)
  Response.Write "<center><span style='font-size: 12pt; font-family: 標楷體'>"&ReportName&"</span></center>"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table x:str id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(I).Name&"</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr>"
	  For I = 1 To FieldsCount
      If I=FieldsCount Then
        Response.Write "<td x:str bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(I)&"</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(I))&"</span></td>"
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

Function EMailList (SQL,ReportName)
  Str_EMail=""
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  Dim I
  Response.Write "<table x:str id=grid border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='98%' align='center'>"
  Response.Write "<tr>"
  For I = 1 To FieldsCount
	  Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  While Not RS1.EOF
	  If Str_EMail<>"" Then
	    Str_EMail=Str_EMail&";"&RS1(1)
	  Else
	    Str_EMail=RS1(1)
	  End If
	  Response.Flush
    Response.Clear
    RS1.MoveNext
  Wend
  Response.Write "<tr>"
    Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>" & Str_EMail & "</span></td>"
  Response.Write "</tr>"
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

Function ReportList2 (SQL,ReportName)
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
      If i=5 Then
        Response.Write "<td x:str bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(I)&"</span></td>"
      Else
        Response.Write "<td bgcolor='#FFFFFF'><span style='font-size: 9pt; font-family: 新細明體'>"&Data_Minus(RS1(I))&"</span></td>"
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
<%Prog_Id="donor"%>
<!--#include file="../include/head_rpt.asp"-->
<body class=tool>
<%
  '20130913Modify by GoodTV-Tanya:Add IsMember<>'Y'只顯示天使的資料
  If request("action")="mobile" Then
    SQL="Select Distinct Donor_Id,編號=Donor_Id,捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),手機=Cellular_Phone " & _
        "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember<>'Y' And Cellular_Phone<>'' "
  ElseIf request("action")="email" Then
    SQL="Select Distinct Donor_Id,Email " & _
        "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember<>'Y' And Email<>'' "
  Else
    SQL="Select Distinct Donor_Id,編號=Donor_Id,捐款人=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),類別=Category,身份別=Donor_Type,聯絡電話=(Case When Cellular_Phone<>'' Then Cellular_Phone Else Case When Tel_Office<>'' Then Tel_Office Else Tel_Home End End),通訊地址=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End),會訊=(Case When IsSendNews='Y' Then 'V' Else '' End),電子報=(Case When IsSendEpaper='Y' Then 'V' Else '' End),年報=(Case When IsSendYNews='Y' Then 'V' Else '' End),生日卡=(Case When IsBirthday='Y' Then 'V' Else '' End),賀卡=(Case When IsXmas='Y' Then 'V' Else '' End),最近捐款日期=CONVERT(VarChar,Last_DonateDate,111) " & _
        "From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where IsMember<>'Y' "
  End If
  If request("Donor_Name")<>"" Then SQL=SQL & "And (Donor_Name Like '%"&request("Donor_Name")&"%' Or NickName Like '%"&request("Donor_Name")&"%' Or Contactor Like '%"&request("Donor_Name")&"%' Or Invoice_Title Like '%"&request("Donor_Name")&"%') "
  If request("Member_No")<>"" Then SQL=SQL & "And Member_No Like '%"&request("Member_No")&"%' "
  If request("Tel_Office")<>"" Then SQL=SQL & "And (Tel_Office Like '%"&request("Tel_Office")&"%' Or Tel_Home Like '%"&request("Tel_Office")&"%' Or Cellular_Phone Like '%"&request("Tel_Office")&"%') "
  If request("Category")<>"" Then SQL=SQL & "And Category = '"&request("Category")&"' "
  If request("Donor_Type")<>"" Then SQL=SQL & "And Donor_Type Like '%"&request("Donor_Type")&"%' "
  If request("IDNo")<>"" Then SQL=SQL & "And IDNo = '"&request("IDNo")&"' "
  If request("Invoice_Title")<>"" Then SQL=SQL & "And Invoice_Title Like '%"&request("Invoice_Title")&"%' "
  If request("City")<>"" Then SQL=SQL & "And (City = '"&request("City")&"' Or Invoice_City = '"&request("City")&"') "
  If request("Area")<>"" Then SQL=SQL & "And (Area = '"&request("Area")&"' Or Invoice_Area = '"&request("Area")&"') "
  If request("Address")<>"" Then SQL=SQL & "And (Address Like '%"&request("Address")&"%' Or Invoice_Address Like '%"&request("Address")&"%') "
  If request("IsSendNews")<>"" Then SQL=SQL & "And IsSendNews = 'Y' "
  If request("IsSendEpaper")<>"" Then SQL=SQL & "And IsSendEpaper = 'Y' "
  If request("IsSendYNews")<>"" Then SQL=SQL & "And IsSendYNews = 'Y' "
  If request("IsBirthday")<>"" Then SQL=SQL & "And IsBirthday = 'Y' "
  If request("IsXmas")<>"" Then SQL=SQL & "And IsXmas = 'Y' "
  SQL=SQL&"Order By Donor_Id Desc"
  If request("action")="mobile" Then
    ReportName="捐款人手機名單"
    call MobileList (SQL,ReportName)
  ElseIf request("action")="email" Then
    ReportName="捐款人EMail名單"
    call EMailList (SQL,ReportName)
  Else
    call ReportList2 (SQL,Prog_Desc)
  End If
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->