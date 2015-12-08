<!--#include file="../include/dbfunctionJ.asp"-->
<%
If Session("action")="txt" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-txt"  
  Response.AddHeader "content-disposition", "attachment;filename=madazine.txt"
End If

If Session("Name_Type")="1" Then
  InvoiceTitle="(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End)"
ElseIf Session("Name_Type")="2" Then
  InvoiceTitle="(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End End Else Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End End)"
ElseIf Session("Name_Type")="3" Then
  InvoiceTitle="(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End)"
ElseIf Session("Name_Type")="4" Then
  InvoiceTitle="(Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End)"
End If
    
SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' Order By Seq"
Set RS = Server.CreateObject("ADODB.RecordSet")
RS.Open SQL,Conn,1,1
While Not RS.EOF
  Row=0
  SQL1="Select Donate.Donor_Id,"&InvoiceTitle&",Donate_Total=Sum(Donate_Amt) From DONATE Join Donor On Donate.Donor_Id=Donor.Donor_Id Where Donate_Purpose='"&RS("Donate_Purpose")&"' And Donate_Amt>0 And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" Group By Donate.Donor_Id,"&InvoiceTitle&" Order By Donate_Total "&Session("Seq_Type")&""
    Response.Write SQL1&vbcrlf
    Response.End
  SQL1="Select Distinct Donate_Total=Sum(Donate_Amt) From DONATE Join Donor On Donate.Donor_Id=Donor.Donor_Id Where Donate_Purpose='"&RS("Donate_Purpose")&"' And Donate_Amt>0 And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" Group By Donate.Donor_Id,"&InvoiceTitle&" Order By Donate_Total "&Session("Seq_Type")&""
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  While Not RS1.EOF
    Row=Row+1
    '捐款用途
    If Row=1 Then Response.Write RS("Donate_Purpose")&vbcrlf
    '捐款金額
    If Cdbl(RS1("Donate_Total"))<1000 Then
      Response.Write Space(Cint(Session("maxlen"))+1-Len(RS1("Donate_Total")))&FormatNumber(RS1("Donate_Total"),0)&vbcrlf
    Else
      Response.Write Space(Cint(Session("maxlen"))-Len(RS1("Donate_Total")))&FormatNumber(RS1("Donate_Total"),0)&vbcrlf
    End If
    '捐款明細
    SQL2="Select Donor.Donor_Id,Donate_Total=CONVERT(numeric,A.Donate_Total) From Donor Join (Select Donate.Donor_Id,"&InvoiceTitle&",Donate_Total=Sum(Donate_Amt) From DONATE Join Donor On Donate.Donor_Id=Donor.Donor_Id Where Donate_Purpose='"&RS("Donate_Purpose")&"' And Donate_Amt>0 And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" Group By Donate.Donor_Id,"&InvoiceTitle&") As A On Donor.Donor_Id=A.Donor_Id And CONVERT(int,A.Donate_Total)='"&RS1("Donate_Total")&"'"
    Response.Write SQL2&vbcrlf
    Response.End
    '     "Order By {fn LENGTH("&InvoiceTitle&")},"&InvoiceTitle&",DONATE.Donor_Id Desc "
    'SQL2="Select Donate.Donor_Id,"&InvoiceTitle&",Donate_Total=Sum(Donate_Amt) From DONATE Join Donor On Donor.Donor_Id=Donate.Donor_Id Where Donate_Purpose='"&RS("Donate_Purpose")&"' And CONVERT(int,Donate_Total)='"&RS1("Donate_Total")&"' And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" Group By Donate.Donor_Id,"&InvoiceTitle&" Order By Donate_Total "&Session("Seq_Type")&""
    RS1.MoveNext
  Wend
  RS1.Close
  Set RS1=Nothing
  RS.MoveNext
Wend
RS.Close
Set RS=Nothing
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("action")
Session.Contents.Remove("Col")
Session.Contents.Remove("keyword")
Session.Contents.Remove("maxlen")
Session.Contents.Remove("Dept_Id")
Session.Contents.Remove("Donate_Date_Begin")
Session.Contents.Remove("Donate_Date_End")
Session.Contents.Remove("Act_Id")
Session.Contents.Remove("Name_Type")
Session.Contents.Remove("Amt_Type")
Session.Contents.Remove("Purpose_Type")
Session.Contents.Remove("Seq_Type")
Session.Contents.Remove("SQL_Where") 
%>