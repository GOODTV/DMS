<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL="Select DONATE.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,IDNo=DONOR.IDNo, " & _
    "Address2=A.mValue+DONOR.ZipCode+B.mValue+DONOR.Address,Invoice_Address2=C.mValue+DONOR.Invoice_ZipCode+D.mValue+DONOR.Invoice_Address, " & _
    "DEPT.Comp_ShortName,ACT.Act_ShortName,DEPT.Invoice_Prog " &_
    "From DONATE " & _
    "Join DONOR On DONATE.Donor_Id=DONOR.Donor_Id " & _
    "Left Join Dept On DONATE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join ACT On DONATE.Act_Id=ACT.Act_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donate_Id='"&request("donate_id")&"' "
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL,Conn,1,3	
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>收據列印</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class="tool" onload="print();">
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br><br><br><font size=3>沒有符合條件的資料可以列印!!</font>"
      Response.End
    Else
      While Not RS1.EOF
        Response.Write "<br><br><br><font size=3>捐款收據套印中</font>"
        If Not RS1.EOF Then
          Response.Write "<div align='center' class='pagebreak'>&nbsp;</div>"
        End If        
        RS1("Invoice_Print")="1"
        RS1.Update
        Response.Flush
        Response.Clear        
        RS1.MoveNext
      Wend
    End If
    RS1.Close
    Set RS1=Nothing
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->