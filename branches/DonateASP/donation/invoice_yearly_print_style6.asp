<!--#include file="../include/dbfunctionJ.asp"-->
<%
'邊界:左:8 右:8 上:8 下:8
SQL1=Session("SQL1")
Donate_Where=Session("Donate_Where")
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open SQL1,Conn,1,1
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=ProgDesc("invoice_yearly_print")%></title>
  <object id="factory" viewastext style="display:none" classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="http://mis.npois.com.tw/smsx.cab"></object>
  <script language="javascript">
    function window.onload(){
      factory.printing.header='' ;          //頁首
      factory.printing.footer='' ;          //頁尾
      factory.printing.portrait = true ;    //直印(true),橫印(false)
      factory.printing.leftMargin = 8.0 ;   //左邊界
      factory.printing.topMargin = 8.0 ;    //上邊界
      factory.printing.rightMargin = 8.0 ;  //右邊界
      factory.printing.bottomMargin = 8.0 ; //下邊界
      factory.printing.print();
    }
  </Script>
  <style>
  <!--
    table.TableGrid{
	    font-size: 12px;
	    width:194mm;
    }
    .Comp_Name{ 
	    width:194mm;
	    height:12mm;
      font-size:20;
      font-family:"標楷體";
    }
    .Invoice_No{
      height:5mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Desc{
      height:8mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Right{
      font-size:12;
      font-family:"標楷體";
    }
    .Donate_Seal{
      height:10mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Account{
      height:8mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Foot{
      height:5mm;
      font-size:12;
      font-family:"標楷體";
    }
    .Donate_Seal{
      height:10mm;
      font-size:15;
      font-family:"標楷體";
    }
    .Donate_Amt{
      height:9mm;
      font-size:13;
      font-family:"標楷體";
    }      
    .PageBreak {
      page-break-after:always;
    }
  -->
  </style>
</head>
<body class=tool <%If Not RS1.EOF And Session("action")="report" Then%>onload='print();'<%End If%>>	
  <p><div align="center"><center>
  <%
    If RS1.EOF Then
      Response.Write "<br /><br /><br /><font size=3>沒有符合條件的資料可以列印!!</font>"
    Else
      Response.Write "<br /><br /><br /><font size=3>請貴會提供樣板格式，謝謝!!</font>"
      Response.End
      Row=1
      Max_Page=3
      SQL2="Select *,Dept_Address=Zip_Code+A.mValue+B.mValue+Address,Dept_Address2=Zip_Code2+C.mValue+D.mValue+Address2 From Dept " & _
           "Left Join CodeCity As A On Dept.City_Code=A.mCode " & _
           "Left Join CodeCity As B On Dept.Area_Code=B.mCode " & _
           "Left Join CodeCity As C On Dept.City_Code2=C.mCode " & _
           "Left Join CodeCity As D On Dept.Area_Code2=D.mCode Where Dept_Id='"&Session("DeptId")&"' "
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      If Not RS2.EOF Then
        Invoice_Pre=RS2("Invoice_Pre")
        Invoice_Rule_Type=RS2("Invoice_Rule_Type")
        Invoice_Rule_YMD=RS2("Invoice_Rule_YMD")
        Invoice_Rule_Len=RS2("Invoice_Rule_Len")
        Invoice_Rule_Pub=RS2("Invoice_Rule_Pub")
        Invoice_Pre2=RS2("Invoice_Pre2")
        Invoice_Rule_Type2=RS2("Invoice_Rule_Type2")
        Invoice_Rule_YMD2=RS2("Invoice_Rule_YMD2")
        Invoice_Rule_Len2=RS2("Invoice_Rule_Len2")
        Invoice_Rule_Pub2=RS2("Invoice_Rule_Pub2")
        Donate_Invoice=RS2("Donate_Invoice")
        
        Comp_Name=RS2("Comp_Name")
        Account=RS2("Account")
        Uniform_No=RS2("Uniform_No")
        Tel=RS2("Tel")
        Fax=RS2("Fax")
        EMail=RS2("EMail")
        Fax=RS2("Fax")
        Dept_Address=RS2("Dept_Address")
      End If
      RS2.Close
      Set RS2=Nothing

      '電子印章
      SEAL_1="../images/blank.gif"
      SEAL_2="../images/blank.gif"
      SEAL_3="../images/blank.gif"
      SEAL_D="../images/blank.gif"
      SQL2="Select * From SEAL Where (Ser_No<=4 Or Seal_Flg='Y' Or Seal_Subject='"&session("user_name")&"')"
      call QuerySQL(SQL2,RS2)
      While Not RS2.EOF
        If RS2("Seal_TitleImg")<>"" Then
          If RS2("Ser_No")="1" Then 
            SEAL_1="../upload/"&RS2("Seal_TitleImg")
          ElseIf RS2("Ser_No")="2" Then 
            SEAL_2="../upload/"&RS2("Seal_TitleImg")
          ElseIf RS2("Ser_No")="3" Then 
            SEAL_3="../upload/"&RS2("Seal_TitleImg")        
          Else
            If RS2("Seal_Flg")="Y" Then
              If RS2("Seal_TitleImg")<>"" Then SEAL_D="../upload/"&RS2("Seal_TitleImg")
            Else
              If RS2("Seal_TitleImg")<>"" Then SEAL_4="../upload/"&RS2("Seal_TitleImg")
            End If
          End If
        End If
        RS2.MoveNext
      Wend
      RS2.Close
      Set RS2=Nothing
      If SEAL_4="" Then SEAL_4=SEAL_D
                
      While Not RS1.EOF
        '取捐款編號
        If RS1("Invoice_No")="" Then
          Invoice_Pre=Invoice_Pre
          Invoice_No=Get_Invoice_NoY("1",Session("DeptId"),Session("Donate_Year")&"/12/31",Invoice_Rule_Type2,Session("DonateDesc"),Invoice_Rule_Type,Invoice_Rule_YMD,Invoice_Rule_Len,Invoice_Rule_Pub)
        Else
          Invoice_Pre=RS1("Invoice_Pre")
          Invoice_No=RS1("Invoice_No") 
        End If

        '收據抬頭改印
        If Session("Invoice_Title_New")<>"" Then
          Donor_Name=Session("Invoice_Title_New")
        Else
          Donor_Name=RS1("Donor_Name")
        End If
        
        SQL2="Select Donate_Date=CONVERT(VarChar,Donate_Date,111),Donate_Amt=Isnull(Donate_Amt,0),Donate_Payment,Donate_Purpose,Invoice_Pre,Invoice_No,Comment=CONVERT(nvarchar(100), Comment),Invoice_No,Invoice_Print,Invoice_Print_Yearly From Donate Join DONOR On Donate.Donor_Id=DONOR.Donor_Id "&Donate_Where&" And Donate.Donor_Id='"&RS1("Donor_Id")&"' And Donate.Invoice_Title='"&RS1("Donor_Name")&"' Order By Donate_Date,Donate_Id"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
        While Not RS2.EOF
          If RS2("Invoice_No")="" Then
            RS2("Invoice_Pre")=Invoice_Pre
            RS2("Invoice_No")=Invoice_No
          End If
          RS2("Invoice_Print")="1"
          RS2("Invoice_Print_Yearly")="1"
          RS2.Update
          RS2.MoveNext
        Wend
        RS2.Close
        Set RS2=Nothing

        If Row<RS1.Recordcount Then Response.Write "<div class='pagebreak'>&nbsp;</div>"  
        Row=Row+1
        Response.Flush
        Response.Clear
        RS1.MoveNext
      Wend
      RS1.Close
      Set RS1=Nothing
    End If
    Session.Contents.Remove("DeptId")
    Session.Contents.Remove("DonateDesc")
    Session.Contents.Remove("Rept_Licence")
    Session.Contents.Remove("Print_Desc")
    Session.Contents.Remove("SQL1")
    Session.Contents.Remove("Donate_Where")
    Session.Contents.Remove("Donate_Year")
    Session.Contents.Remove("action")
    Session.Contents.Remove("Invoice_Title_New")
  %>
  </center></div></p>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->