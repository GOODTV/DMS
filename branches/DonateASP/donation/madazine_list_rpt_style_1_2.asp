<!--#include file="../include/dbfunctionJ.asp"-->
<%
If Session("action")="txt" Then
  Response.buffer = true
  Response.Charset = "utf-8"
  Response.ContentType = "application/vnd.ms-txt"  
  Response.AddHeader "content-disposition", "attachment;filename=madazine.txt"
End If

Session("maxlen")=Cint(Session("maxlen"))-1
If Session("Name_Type")="1" Then
  InvoiceTitle="(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End Else Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End End)"
ElseIf Session("Name_Type")="2" Then
  InvoiceTitle="(Case When IsAnonymous='Y' Then Case When NickName<>'' Then NickName Else Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End End Else Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End End)"
ElseIf Session("Name_Type")="3" Then
  InvoiceTitle="(Case When Donate.Invoice_Title<>'' Then Donate.Invoice_Title Else Donor_Name End)"
ElseIf Session("Name_Type")="4" Then
  InvoiceTitle="(Case When Donor_Name<>'' Then Donor_Name Else Donate.Invoice_Title End)"
End If

SQL1="Select Distinct Donate_Amt From DONATE Join Donor On Donor.Donor_Id=Donate.Donor_Id Where Donate_Amt>0 And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" Order By Donate_Amt "&Session("Seq_Type")&""
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
While Not RS1.EOF
  Row=Row+1
  '捐款金額
  If Cdbl(RS1("Donate_Amt"))<1000 Then
    Response.Write Space(Cint(Session("maxlen"))+1-Len(RS1("Donate_Amt")))&FormatNumber(RS1("Donate_Amt"),0)
  Else
    Response.Write Space(Cint(Session("maxlen"))-Len(RS1("Donate_Amt")))&FormatNumber(RS1("Donate_Amt"),0)
  End If
  '捐款明細
  Row2=0
  Same_Amt=1
  Donor_Id=""
  Invoice_Title=""
  SQL2="Select Donate_Amt,Invoice_Title="&InvoiceTitle&",Donate_Date,Donate.Donor_Id From DONATE Join Donor On Donor.Donor_Id=Donate.Donor_Id Where CONVERT(int,Donate_Amt)='"&RS1("Donate_Amt")&"' And (Issue_Type Is Null Or Issue_Type='' Or Issue_Type='M') "&Session("SQL_Where")&" " & _
       "Order By {fn LENGTH("&InvoiceTitle&")},"&InvoiceTitle&",Donate_Date,DONATE.Donor_Id Desc "
  Set RS2 = Server.CreateObject("ADODB.RecordSet")
  RS2.Open SQL2,Conn,1,1
  While Not RS2.EOF
    If Cstr(RS2("Donor_Id"))<>Donor_Id And Cstr(RS2("Invoice_Title"))<>Invoice_Title Then
      Row2=Row2+1
      Donor_Id=Cstr(RS2("Donor_Id"))
      Invoice_Title=Cstr(RS2("Invoice_Title"))
      If Instr(Invoice_Title,Session("keyword"))>0 Then Invoice_Title="("&Invoice_Title&")"
      If Same_Amt>1 Then Invoice_Title=Invoice_Title&"*"&Same_Amt
        
      If (Row2+Cint(Session("Col"))) Mod Cint(Session("Col")) = 1 Then
        Response.Write Space(1)
        If Row2>1 Then Response.Write Space(Cint(Session("maxlen"))+1)
      Else
        Response.Write Session("keyword")
      End If
      Response.Write Invoice_Title
        
      If Row2>1 Then
        If (Row2+Cint(Session("Col"))) Mod Cint(Session("Col")) = 0 Then Response.Write vbcrlf
      End If
      Same_Amt=1
    Else
      Same_Amt=Same_Amt+1
    End If
    Response.Flush
    Response.Clear
    RS2.MoveNext
  Wend
  RS2.Close
  Set RS2=Nothing
  If (Row2+Cint(Session("Col"))) Mod Cint(Session("Col"))<>0 Then Response.Write vbcrlf
  RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing
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