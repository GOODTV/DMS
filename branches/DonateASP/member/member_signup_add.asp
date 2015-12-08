<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  '新增報名資料
  SQL1="MEMBER_SIGNUP"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1.Addnew
  RS1("Member_Act_Id")=request("member_act_id")
  RS1("Donor_Id")=request("Donor_Id")
  RS1("Member_Signup_Date")=Date()
  RS1("Member_Signup_DateTime")=Now()
  RS1("Member_Signup_IP")=Request.ServerVariables("REMOTE_HOST")
  RS1("Member_Signup_Price")="0"
  RS1("Member_Signup_Food")=request("Member_Signup_Food")
  RS1("Member_Signup_Status")="已報名"
  RS1("od_sob")=""
  RS1("Create_Date")=Date()
  RS1("Create_DateTime")=Now()
  RS1("Create_User")=session("user_name")
  RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="活動報名資料新增成功 ！"
  Response.Redirect "member_signup.asp?member_act_id="&request("member_act_id")
End If

SQL1="Select * From MEMBER_SIGNUP Where Member_Act_Id='"&request("member_act_id")&"' And Donor_Id='"&request("donor_id")&"' And Member_Signup_Status In ('已繳費','已報名')"
call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then
  session("errnumber")=1
  session("msg")="您選擇的報名人資料\n系統已有報名紀錄請確認 ！"
  Response.Redirect "member_signup.asp?member_act_id="&request("member_act_id")
End If
RS1.Close
Set RS1=Nothing

SQL1="Select Member_Act_Name,Member_Act_IsFood,Member_Act_PriceD,Member_Act_PriceM From MEMBER_ACT Where Member_Act_Id='"&request("member_act_id")&"'"
call QuerySQL(SQL1,RS1)
Member_Act_Name=RS1("Member_Act_Name")
Member_Act_IsFood=RS1("Member_Act_IsFood")
Member_Act_PriceD=RS1("Member_Act_PriceD")
Member_Act_PriceM=RS1("Member_Act_PriceM")
RS1.Close
Set RS1=Nothing
  
SQL="Select *,Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Invoice_Address Else Case When C.mValue<>B.mValue Then C.mValue+DONOR.Invoice_ZipCode+D.mValue+DONOR.Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End) From DONOR " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _ 
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donor_Id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>活動報名資料新增</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="member_act_id" value="<%=request("member_act_id")%>">
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">	
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"> </td>
                <td width="95%">
  		            <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                <tr>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                      <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                      <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                    </tr>
                    <tr>
                      <td class="table62-bg">　</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">活動報名資料新增</td>
                      <td class="table63-bg">　</td>
                    </tr>  
    	            </table>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td> 
	          <table width="780" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">　</td>
                <td valign="top">
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">姓名：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=RS("Donor_Name")%>&nbsp;<%=RS("Title")%>">
                              &nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">
                              &nbsp;會員別：
                              <input type="text" name="Member_Type" size="40" class="font9t" readonly value="<%=RS("Member_Type")%>">
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=RS("Invoice_Address2")%>">
                              &nbsp;收據開立：
                              <input type="text" name="InvoiceType" size="10" class="font9t" readonly value="<%=RS("Invoice_Type")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;最近捐款日：
                              <input type="text" name="Last_DonateDate" size="10" class="font9t" readonly value="<%=RS("Last_DonateDate")%>">	
                            </td>
                          </tr>
                          <tr>
                            <td align="right">活動名稱：</td>
                            <td align="left" colspan="7"><input type="text" name="Member_Act_Name" size="104" class="font9t" readonly value="<%=Member_Act_Name%>"></td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">報名日期：</td>
                            <%If Member_Act_IsFood="Y" Then%>
                            <td width="15% align="left"><%call Calendar("Donate_Date",Date())%></td>
                            <td width="10%" align="right">飲食習慣：</td>
                            <td width="65%">
                              <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food1" value="葷" checked >葷
                      	      <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food2" value="素">素
                      	      <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food3" value="不需要">不需要	
                            </td>
                            <%Else%>
                            <td width="90% align="left"><%call Calendar("Donate_Date",Date())%></td>
                            <input type="hidden" name="Member_Signup_Food">
                            <%End If%>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 存 檔 " name="save" class="cbutton" style="cursor:hand" onClick="Save_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 取 消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">　</td>
              </tr>
              <tr>
                <td style="background-color:#EEEEE3"><img src="../images/table06_06.gif" width="10" height="10"></td>
                <td class="table64-bg"><img src="../images/table06_07.gif" width="1" height="10"></td>
                <td style="background-color:#EEEEE3"><img src="../images/table06_08.gif" width="10" height="10"></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Save_OnClick(){
  <%call CheckStringJ("Donate_Date","報名日期")%>
  <%call CheckDateJ("Donate_Date","報名日期")%>
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='member_signup.asp?member_act_id='+document.form.member_act_id.value+'';
}
--></script>