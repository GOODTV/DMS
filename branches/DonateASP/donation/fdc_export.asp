<%Response.ContentType="text/html; charset=utf-8"%>
<%
Response.Buffer =true
Response.Expires=-1
If session("user_id")="" Then
  Session.Abandon
  Response.Redirect("../../sysmgr/timeout.asp")
End If
session.Timeout=60
Server.ScriptTimeout=600
Set Conn=server.createobject("ADODB.Connection")
Conn.Connectiontimeout=600
Conn.commandtimeout=600
Conn.Provider="sqloledb"
Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")

Transfer_FileName=Cstr(Cint(Session("Donate_Year")-1911))
If Len(Transfer_FileName)=2 Then Transfer_FileName="0"&Transfer_FileName
Transfer_FileName=Transfer_FileName&".31."&Cstr(Session("Uniform_No"))&".txt"

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'捐贈年度
Donate_Year=Cstr(Cint(Session("Donate_Year")-1911))
If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
'受捐者單位統編
Uniform_No=Cstr(Session("Uniform_No"))
Uniform_No=Uniform_No&Space(10-Len(Uniform_No))
'受捐者單位名稱
SQL="Select Comp_Name From DEPT Where Dept_Id='"&Session("DeptId")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,1
For I = 1 To Len(RS("Comp_Name"))
  If Asc(Mid(RS("Comp_Name"),I,1))<0 Then
    CompNameLen=CompNameLen+2
  Else
    CompNameLen=CompNameLen+1
  End If
Next
Comp_Name=RS("Comp_Name")&Space(40-CompNameLen)
RS.Close
Set RS=Nothing
'許可文號
For I = 1 To Len(Session("Licence"))
  If Asc(Mid(Session("Licence"),I,1))<0 Then
    LicenceLen=LicenceLen+2
  Else
    LicenceLen=LicenceLen+1
  End If
Next
Licence=Session("Licence")&Space(60-LicenceLen)
Row=1    
SQL="Select Distinct Donor.Donor_Id,Invoice_Title,Invoice_IDNo,Donate_No=A.Donate_No,Donate_Total=CONVERT(numeric,A.Donate_Total) " & _
    "From Donor Join (Select Donor_Id,Donate_No=Count(*),Donate_Total=Sum(Donate_Amt) From Donate " & _
    "Where Year(Donate_Date)='"&Session("Donate_Year")&"' And Issue_Type<>'D' And Donate_Amt>0 "
If Session("DeptId")<>"" Then
  SQL=SQL&"And Dept_Id='"&Session("DeptId")&"' "
Else
  SQL=SQL&"And Dept_Id In ("&Session("all_dept_type")&") "
End If
If Session("Act_Id")<>"" Then SQL=SQL&"And Act_Id='"&Session("Act_Id")&"' "
SQL=SQL&"Group By Donor_Id) As A On Donor.Donor_Id=A.Donor_Id Where IsFdc='Y' And ({ fn LENGTH(DONOR.Invoice_IDNo) }=8 Or { fn LENGTH(DONOR.Invoice_IDNo) }=10) Order By Donor.Donor_Id"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,1
While Not RS.EOF
  Check_IdNo=True
  For I = 1 To Len(RS("Invoice_IDNo"))
   If Asc(Mid(RS("Invoice_IDNo"),I,1))<0 Then
     Check_IdNo=False
     Exit For
   End If
  Next
  If Check_IdNo Then
    '捐贈年度
    Response.Write Donate_Year
    '捐贈者身分證/統編
    Invoice_IDNo=UCase(RS("Invoice_IDNo"))
    IDNoLen=10
    If Len(Invoice_IDNo)>=IDNoLen Then
      Response.Write Left(Invoice_IDNo,IDNoLen)
    Else
      Response.Write Invoice_IDNo&Space(IDNoLen-Len(Invoice_IDNo))
    End If
    '捐贈者姓名
    NamekLen=12
    InvoiceTitle=""
    If Instr(RS("Invoice_Title"),"&#")>0 Then
      For I=1 To Len(RS("Invoice_Title"))
        If Mid(RS("Invoice_Title"),I,1)="&" And I<Len(RS("Invoice_Title")) Then
          If Mid(RS("Invoice_Title"),I+1,1)="#" Then
            InvoiceTitle=InvoiceTitle&"*"
            I=I+7
          Else
            InvoiceTitle=InvoiceTitle&Mid(RS("Invoice_Title"),I,1)
          End If
        Else
          InvoiceTitle=InvoiceTitle&Mid(RS("Invoice_Title"),I,1)
        End If
      Next
    Else
      InvoiceTitle=RS("Invoice_Title")
    End If
    Invoice_Title=""
    TitleLen=0
    For I = 1 To Len(InvoiceTitle)
      If Asc(Mid(InvoiceTitle,I,1))<0 Then
        TitleLen=TitleLen+2
      Else
        TitleLen=TitleLen+1
      End If
      If TitleLen<=12 Then
        Invoice_Title=Invoice_Title&Mid(InvoiceTitle,I,1)
      Else
        If Asc(Mid(InvoiceTitle,I,1))<0 Then
          TitleLen=TitleLen-2
        Else
          TitleLen=TitleLen-1
        End If 
        Exit For
      End If
    Next
    Response.Write Invoice_Title
    If TitleLen<12 Then Response.Write Space(12-TitleLen)
    '捐贈金額
    TotalLen=10
    If Len(RS("Donate_Total"))>=TotalLen Then
      Response.Write Right(RS("Donate_Total"),TotalLen)
    Else
      Response.Write Left("0000000000",10-Len(RS("Donate_Total")))&RS("Donate_Total")
    End If
    '受捐者單位統編
    Response.Write Uniform_No
    '捐贈別
    Response.Write "00"
    '受捐者單位名稱
    Response.Write Comp_Name
    '許可文號
    Response.Write Licence
    '換行
    If Row<RS.Recordcount Then Response.Write vbcrlf
  End If
  Row=Row+1
  Response.Flush
  Response.Clear
  RS.MoveNext
Wend
RS.Close
Set RS=Nothing
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("DeptId")
Session.Contents.Remove("Donate_Year")
Session.Contents.Remove("Act_Id")
Session.Contents.Remove("Uniform_No")
Session.Contents.Remove("Licence")
%>
