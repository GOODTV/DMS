<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<%
  If Request.Form("action")="donationa" Then
    Dept_Id=Request.Form("Donate_DeptId")
    Donate_Act_Id=Request.Form("Donate_ActId")
    Donate_Type=Request.Form("Donate_Type")
    Donate_Name=Request.Form("Donate_Name")
    Donate_IDNO=Request.Form("Donate_IDNO")
    Donate_Amount=Request.Form("Donate_Amount")
    Donate_CellPhone=Request.Form("Donate_CellPhone")
    Donate_ZipCode=Request.Form("Donate_ZipCode")
    Donate_Address=Request.Form("Donate_Address")
    Donate_Email=Request.Form("Donate_Email")
    Check_Request_Form("b")
    
    '捐款人資料(DONOR)
    SQL1="Select * From DONOR Where Donor_Name='"&Request.Form("Donate_Name")&"' "
    WhereSQL=""
    If Request.Form("Donate_IDNO")<>"" Then
      If WhereSQL="" Then
        WhereSQL=WhereSQL&"And (IDNo='"&Request.Form("Donate_IDNO")&"' "
      Else
         WhereSQL=WhereSQL&"Or IDNo='"&Request.Form("Donate_IDNO")&"' "
      End If
    End If
    If Request.Form("Donate_Address")<>"" Then
      If WhereSQL="" Then
        WhereSQL=WhereSQL&"And (Address='"&Request.Form("Donate_Address")&"' Or Invoice_Address='"&Request.Form("Donate_Address")&"' "
      Else
         WhereSQL=WhereSQL&"Or Address='"&Request.Form("Donate_Address")&"' Or Invoice_Address='"&Request.Form("Donate_Address")&"' "
      End If
    End If
    If Request.Form("Donate_CellPhone")<>"" Then
      If WhereSQL="" Then
        WhereSQL=WhereSQL&"And (Cellular_Phone='"&Request.Form("Donate_CellPhone")&"' "
      Else
         WhereSQL=WhereSQL&"Or Cellular_Phone='"&Request.Form("Donate_CellPhone")&"' "
      End If
    End If
    If Request.Form("Donate_TelOffice")<>"" Then
      If WhereSQL="" Then
        WhereSQL=WhereSQL&"And (Tel_Office='"&Request.Form("Donate_TelOffice")&"' Or Tel_Home='"&Request.Form("Donate_TelOffice")&"' "
      Else
         WhereSQL=WhereSQL&"Or Tel_Office='"&Request.Form("Donate_TelOffice")&"' Or Tel_Home='"&Request.Form("Donate_TelOffice")&"' "
      End If
    End If
    If Request.Form("Donate_Email")<>"" Then
      If WhereSQL="" Then
        WhereSQL=WhereSQL&"And (Email='"&Request.Form("Donate_Email")&"' "
      Else
         WhereSQL=WhereSQL&"Or Email='"&Request.Form("Donate_Email")&"' "
      End If
    End If
    If WhereSQL<>"" Then WhereSQL=WhereSQL&")"
    SQL1=SQL1&WhereSQL

    Donor_Id=0
    Set RS1=Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    If RS1.EOF Then
      SQL2="DONOR"
      Set RS2=Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,3
      RS2.Addnew
      RS2("Dept_Id")=Dept_Id
      RS2("Donor_Name")=Request.Form("Donate_Name")
      RS2("Sex")=Request.Form("Donate_Sex")
      RS2("Title")="君"
      If Request.Form("Donate_IDNO")<>"" Then
        If Len(Request.Form("Donate_IDNO"))=8 Then
          RS2("Category")="團體"
        Else
          RS2("Category")="個人"
        End If
      Else
        RS2("Category")="個人"
      End If  
      RS2("Donor_type")=""
      RS2("IDNo")=Request.Form("Donate_IDNO")
      If Request.Form("Donate_Birthday")<>"" Then RS2("Birthday")=Request.Form("Donate_Birthday")
      RS2("Education")=Request.Form("Donate_Education")
      RS2("Occupation")=Request.Form("Donate_Occupation")
      RS2("Marriage")=Request.Form("Donate_Marriage")
      RS2("Religion")=""
      RS2("ReligionName")=""
      RS2("Cellular_Phone")=Request.Form("Donate_CellPhone")
      Tel_Office=Request.Form("Donate_TelOffice")
      If Request.Form("Donate_TelOffice_Region")<>"" Then Tel_Office="("&Request.Form("Donate_TelOffice_Region")&")"&Tel_Office
      If Request.Form("Donate_TelOffice_Ext")<>"" Then Tel_Office=Tel_Office&"#"&Request.Form("Donate_TelOffice_Ext")
      RS2("Tel_Office")=Tel_Office
      RS2("Tel_Home")=""
      RS2("Fax")=""
      RS2("Email")=Request.Form("Donate_Email")
      RS2("Contactor")=""
      RS2("OrgName")=""
      RS2("JobTitle")=""
      RS2("IsAbroad")="N"
      RS2("ZipCode")=Request.Form("Donate_ZipCode")  
      RS2("City")=Request.Form("Donate_CityCode")
      RS2("Area")=Request.Form("Donate_AreaCode")
      RS2("Address")=Request.Form("Donate_Address")
      RS2("IsSendNews")="N"
      RS2("IsSendEpaper")="N"
      RS2("IsSendYNews")="N"
      RS2("IsBirthday")="N"
      RS2("IsXmas")="N"
      RS2("Invoice_type")=Request.Form("Donate_Invoice_Type")
      RS2("IsAnonymous")="N"
      RS2("NickName")=""
      If Request.Form("Donate_Invoice_Title")<>"" Then
        RS2("Invoice_Title")=Request.Form("Donate_Invoice_Title")
      Else
        RS2("Invoice_Title")=Request.Form("Donate_Name")
      End If
      If Request.Form("Donate_Invoice_IDNo")<>"" Then
        RS2("Invoice_IDNo")=Request.Form("Donate_Invoice_IDNo")
      Else
        RS2("Invoice_IDNo")=Request.Form("Donate_IDNO")
      End If
      If Request.Form("Donate_Invoice_ZipCode")<>"" Then
        RS2("Invoice_ZipCode")=Request.Form("Donate_Invoice_ZipCode")
        RS2("Invoice_City")=Request.Form("Donate_Invoice_CityCode")
        RS2("Invoice_Area")=Request.Form("Donate_Invoice_AreaCode")
        RS2("Invoice_Address")=Request.Form("Donate_Invoice_Address")
      Else
        RS2("Invoice_ZipCode")=Request.Form("Donate_ZipCode")
        RS2("Invoice_City")=Request.Form("Donate_CityCode")
        RS2("Invoice_Area")=Request.Form("Donate_AreaCode")
        RS2("Invoice_Address")=Request.Form("Donate_Address")
      End If
      RS2("Remark")=""
      RS2("IsThanks")="0"
      RS2("IsThanks_Add")="0"
      RS2("Donate_No")="0"
      RS2("Donate_Total")="0"
      RS2("Donate_NoD")="0"
      RS2("Donate_TotalD")="0"
      RS2("Donate_NoC")="0"
      RS2("Donate_TotalC")="0"
      RS2("Donate_NoM")="0"
      RS2("Donate_TotalM")="0"
      RS2("Donate_NoS")="0"
      RS2("Donate_TotalS")="0"
      RS2("Donate_NoA")="0"
      RS2("Donate_TotalA")="0"
      RS2("Donate_NoND")="0"
      RS2("Donate_TotalND")="0"
      RS2("Create_Date")=Date()
      RS2("Create_DateTime")=Now()
      RS2("Create_IP")=Request.ServerVariables("REMOTE_ADDR")
      RS2("Create_User")="線上金流"
      RS2("IsMember")="N"
      RS2("Member_No")=""
      RS2("Member_Type")=""
      RS2("Member_Status")=""
      RS2("IsVolunteer")="N"
      RS2("Volunteer_Type")=""
      RS2("Volunteer_Status")=""
      RS2("IsGroup")="N"
      RS2("Group_No")=""
      RS2("Group_Licence")=""
      RS2("Group_CreateDate")=null
      RS2("Group_Person")=""
      RS2("Group_Person_JobTitle")=""
      RS2("Group_WebUrll")=""
      RS2("Group_Mission")=""
      RS2("Group_Service")=""
      RS2("Group_Other")=""
      RS2.Update
      RS2.Close
      Set RS2=Nothing
      SQL2="Select @@IDENTITY As Donor_Id"
      Set RS2=Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,1
      If Not RS2.EOF Then Donor_Id=RS2("Donor_Id")
      RS2.Close
      Set RS2=Nothing
    Else
      Donor_Id=RS1("Donor_Id")
    End If
    RS1.Close
    Set RS1=Nothing
    
    '產生交易單號od_sob
    If session("get_od_sob")="" Then
      od_sob=get_od_sob("D")
    
      '新增交易資料(DONATE_WEB)
      SQL1="DONATE_WEB"
      Set RS1=Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      RS1.Addnew
      RS1("Donate_Type")="barcode"
      RS1("od_sob")=od_sob
      RS1("Dept_Id")=Dept_Id
      RS1("Donor_Id")=Donor_Id
      RS1("Donate_CreateDate")=Date()
      RS1("Donate_CreateDateTime")=Now()
      RS1("Donate_CreateIP")=Request.ServerVariables("REMOTE_ADDR")
      RS1("Donate_Amount")=Request.Form("Donate_Amount")
      RS1("Donate_DonorName")=Request.Form("Donate_Name")
      RS1("Donate_Sex")=Request.Form("Donate_Sex")
      RS1("Donate_IDNO")=Request.Form("Donate_IDNO")
      If Request.Form("Donate_Birthday")<>"" Then RS1("Donate_Birthday")=Request.Form("Donate_Birthday")
      RS1("Donate_Education")=Request.Form("Donate_Education")
      RS1("Donate_Occupation")=Request.Form("Donate_Occupation")
      RS1("Donate_Marriage")=Request.Form("Donate_Marriage")
      RS1("Donate_CellPhone")=Request.Form("Donate_CellPhone")
      RS1("Donate_TelOffice_Region")=Request.Form("Donate_TelOffice_Region")
      RS1("Donate_TelOffice")=Request.Form("Donate_TelOffice")
      RS1("Donate_TelOffice_Ext")=Request.Form("Donate_TelOffice_Ext")
      RS1("Donate_TelHome_Region")=Request.Form("Donate_TelHome_Region")
      RS1("Donate_TelHome")=Request.Form("Donate_TelHome")
      RS1("Donate_TelHome_Ext")=Request.Form("Donate_TelHome_Ext")
      RS1("Donate_ZipCode")=Request.Form("Donate_ZipCode")
      RS1("Donate_CityCode")=Request.Form("Donate_CityCode")
      RS1("Donate_AreaCode")=Request.Form("Donate_AreaCode")
      RS1("Donate_Address")=Request.Form("Donate_Address")
      RS1("Donate_Email")=Request.Form("Donate_Email")
      RS1("Donate_Purpose")=Request.Form("Donate_Purpose")
      RS1("Donate_Invoice_Type")=Request.Form("Donate_Invoice_Type")
      If Request.Form("Donate_Invoice_Title")<>"" Then
        RS1("Donate_Invoice_Title")=Request.Form("Donate_Invoice_Title")
      Else
        RS1("Donate_Invoice_Title")=Request.Form("Donate_DonorName")
      End If
      If Request.Form("Donate_Invoice_IDNo")<>"" Then
        RS1("Donate_Invoice_IDNo")=Request.Form("Donate_Invoice_IDNo")
      Else
        RS1("Donate_Invoice_IDNo")=Request.Form("Donate_IDNO")
      End If
      If Request.Form("Donate_Invoice_ZipCode")<>"" Then
        RS1("Donate_Invoice_ZipCode")=Request.Form("Donate_Invoice_ZipCode")
        RS1("Donate_Invoice_CityCode")=Request.Form("Donate_Invoice_CityCode")
        RS1("Donate_Invoice_AreaCode")=Request.Form("Donate_Invoice_AreaCode")
        RS1("Donate_Invoice_Address")=Request.Form("Donate_Invoice_Address")
      Else
        RS1("Donate_Invoice_ZipCode")=Request.Form("Donate_ZipCode")
        RS1("Donate_Invoice_CityCode")=Request.Form("Donate_CityCode")
        RS1("Donate_Invoice_AreaCode")=Request.Form("Donate_AreaCode")
        RS1("Donate_Invoice_Address")=Request.Form("Donate_Address")
      End If
      RS1("Donate_Update")="N"
      RS1("Donate_Export")="N"
      If Donate_Act_Id<>"" Then RS1("Donate_Act_Id")=Donate_Act_Id
      RS1.Update
      RS1.Close
      Set RS1=Nothing
      session("get_od_sob")=od_sob
    Else
      od_sob=session("get_od_sob")
    End If
    
    '收據寄送地址
    Donate_Invoice_Address=""
    If Request.Form("Donate_AreaCode")<>"" Then
      SQL1="Select Invoice_Address=B.mValue+Left(A.mCode,3)+A.mValue From CODECITY A Join CODECITY B On Substring(A.codeMetaID,7,1)=B.mCode Where A.mCode='"&Request.Form("Donate_AreaCode")&"'"
      Set RS1=Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      If Not RS1.EOF Then Donate_Invoice_Address=RS1("Invoice_Address")&Request.Form("Donate_Address")
      RS1.Close
      Set RS1=Nothing
    Else
      Donate_Invoice_Address=Request.Form("Donate_Invoice_Address")
    End If
  Else
    session("errnumber")=1
    session("msg")="很抱歉您輸入的訊息不足，誠摯地請您再試一次！"
    Response.Redirect "ecpay.asp"
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <link href="include/donate.css" rel="stylesheet" type="text/css">
  <script src="js/npois.js" type="text/javascript"></script>
  <script src="js/ecpayb.js" type="text/javascript"></script>
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->  
  <title><%=Comp_Name%></title>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" />  </div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 列印捐款單</div>
    <div id="mid">
      <div id="mid-1">
        <table width="95%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <form action="ecbank_barcode_print.asp" method="post" name="ecbank">
          	<input type="hidden" name="action" value="">
            <input type="hidden" name="amt" value="<%=Request.Form("Donate_Amount")%>">
            <input type="hidden" name="od_sob" value="<%=od_sob%>">
            <input type="hidden" name="payment_type" value="<%=Request.Form("payment_type")%>">	
            <tr>
              <td align="center" colspan="2"><span class="msgshow">◎您即將透過綠界科技便利商店付費系統，列印捐款單<%=FormatNumber(Request.Form("Donate_Amount"),0)%>元，是否確定？</span></td>
            </tr>
            <tr>
              <td width="120" align="right" class="contents"><span class="mustcolumn">◎捐款人資料</span></td>
              <td width="530" class="contents">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">捐款金額：&nbsp;</td>
              <td class="contents"><%=FormatNumber(Request.Form("Donate_Amount"),0)%>元(超商條碼捐款)&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">捐款人姓名：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_Name")%>&nbsp;</td>
            </tr>
            <tr>
              <td width="180" align="right" class="contents">交易序號：&nbsp;</td>
              <td width="470" class="contents"><%=od_sob%>&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">聯絡電話：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_CellPhone")%>&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">電子郵件：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_Email")%>&nbsp;</td>
            </tr>
            <%If Request.Form("Donate_Invoice_Type")<>"不寄收據" Then%>
            <tr>
              <td width="120" align="right" class="contents"><span class="mustcolumn">◎捐款收據資料</span></td>
              <td width="530" class="contents">&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">寄發方式：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_Invoice_Type")%></td>
            </tr>
            <tr>
              <td align="right" class="contents">收據抬頭：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_Invoice_Title")%></td>
            </tr>
            <tr>
              <td align="right" class="contents">收據身分：&nbsp;</td>
              <td class="contents"><%=Request.Form("Donate_Invoice_IDNO")%>&nbsp;</td>
            </tr>
            <tr>
              <td align="right" class="contents">收據寄送地址：&nbsp;</td>
              <td class="contents"><%=Donate_Invoice_Address%></td>
            </tr>
            <%End If%>
            <tr>
              <td align="center" colspan="2"><input type="button" value=" 列印捐款單 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:BarCode_OnClick();"></td>
            </tr>
          </form>
        </table>     
      </div>
    </div>
    <div id="bottom"></div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->