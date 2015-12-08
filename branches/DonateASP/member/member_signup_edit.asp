<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  '修改報名資料
  SQL1="Select * From MEMBER_SIGNUP Where ser_no='"&request("Ser_No")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1("Member_Signup_Status")=request("Member_Signup_Status")
  RS1("Member_Signup_Food")=request("Member_Signup_Food")
  RS1("LastUpdate_Date")=Date()
  RS1("LastUpdate_DateTime")=Now()
  RS1("LastUpdate_User")=session("user_name")
  RS1("LastUpdate_IP")=Request.ServerVariables("REMOTE_HOST")
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  session("errnumber")=1
  session("msg")="活動報名資料修改成功 ！"
  'Response.Redirect "member_signup.asp?member_act_id="&request("member_act_id")
  If request("ctype")="signup_data" Then
    Response.Redirect "../donation/signup_data.asp?donor_id="&request("donor_id")  
  ElseIf request("ctype")="member_signup_data" Then
    Response.Redirect "member_signup_data.asp?donor_id="&request("donor_id")
  Else
    Response.Redirect "member_signup.asp?member_act_id="&request("member_act_id")
  End If
End If

If request("action")="delete" Then
  SQL="Delete From MEMBER_SIGNUP Where ser_no='"&request("Ser_No")&"'"
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="活動報名資料刪除成功 ！"
  If request("ctype")="signup_data" Then
    Response.Redirect "../donation/signup_data.asp?donor_id="&request("donor_id")  
  ElseIf request("ctype")="member_signup_data" Then
    Response.Redirect "member_signup_data.asp?donor_id="&request("donor_id")
  Else
    Response.Redirect "member_signup.asp?member_act_id="&request("member_act_id")
  End If 
End If

SQL="Select Member_Act_Name,Member_Act_IsPrice,MEMBER_SIGNUP.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,InvoiceType=DONOR.Invoice_Type,IDNo=DONOR.IDNo,Member_Type,Invoice_Type,Last_DonateDate, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+DONOR.Address Else A.mValue+DONOR.ZipCode+DONOR.Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>D.mValue Then C.mValue+DONOR.Invoice_ZipCode+D.mValue+Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End) " & _
    "From MEMBER_SIGNUP " & _
    "Join DONOR On MEMBER_SIGNUP.Donor_Id=DONOR.Donor_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Join MEMBER_ACT On MEMBER_SIGNUP.Member_Act_Id=MEMBER_ACT.Member_Act_Id " & _
    "Where MEMBER_SIGNUP.ser_no='"&request("ser_no")&"' "
Call QuerySQL(SQL,RS)
Member_Act_IsPrice=RS("Member_Act_IsPrice")
%>
<html>
<head>
  <meta name="GENERATOR" content="Microsoft FrontPage 6.0">
  <meta name="ProgId" content="FrontPage.Editor.Document">
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>活動報名資料修改</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
<p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ser_no" value="<%=request("Ser_No")%>">	
      <input type="hidden" name="member_act_id" value="<%=RS("member_act_id")%>">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="donor_id" value="<%=RS("donor_id")%>">	
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">活動報名資料修改</td>
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
                            <td align="left" colspan="7"><input type="text" name="Member_Act_Name" size="104" class="font9t" readonly value="<%=RS("Member_Act_Name")%>"></td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">報名狀態：</td>
                            <td width="15%" align="left">
                            <%
                              If Member_Act_IsPrice="N" Then
                                SQL="Select Member_Signup_Status=CodeDesc From CASECODE Where codetype='SignupStatus' And CodeDesc Not In ('已繳費','已退費') Order By Seq"
                              Else
                                SQL="Select Member_Signup_Status=CodeDesc From CASECODE Where codetype='SignupStatus' Order By Seq"
                              End If
                              FName="Member_Signup_Status"
                              Listfield="Member_Signup_Status"
                              menusize="1"
                              BoundColumn=RS("Member_Signup_Status")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                            <%If Member_Act_IsFood="Y" Then%>
                            <td width="10%" align="right">飲食習慣：</td>
                            <td width="65%">
                              <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food1" value="葷" <%If RS("Member_Signup_Food")="葷" Then Response.Write "checked" End If%> >葷
                      	      <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food2" value="素" <%If RS("Member_Signup_Food")="素" Then Response.Write "checked" End If%> >素
                      	      <input type="radio" name="Member_Signup_Food" id="Member_Signup_Food3" value="不需要" <%If RS("Member_Signup_Food")="不需要" Then Response.Write "checked" End If%> >不需要	
                            </td>
                            <%Else%>
                            <td width="75%" align="right"> </td>
                            <input type="hidden" name="Member_Signup_Food">
                            <%End If%>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="update" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
                               &nbsp;&nbsp;
                               <input type="button" value=" 刪 除 " name="save" class="cbutton" style="cursor:hand" onClick="Delete_OnClick()">	
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
function Update_OnClick(){
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='signup_data'){
    location.href='../donation/signup_data.asp?donor_id='+document.form.donor_id.value+'';
  }else if(document.form.ctype.value=='member_signup_data'){
    location.href='member_signup_data.asp?donor_id='+document.form.donor_id.value+'';   
  }else{
    location.href='member_signup.asp?member_act_id='+document.form.member_act_id.value+'';
  }
}
--></script>