<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where Dept_Id='"&request("dept_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Comp_Name")=request("Comp_Name")
  RS("Comp_ShortName")=request("Comp_ShortName")
  RS("Dept_Desc")=request("Dept_Desc")
  RS("Comp_Slogan")=request("Comp_Slogan")
  RS("DEPT_LOGO")=""
  RS("Server_Url")=request("Server_Url")
  RS("Sys_Name")=request("Sys_Name")
  RS("Mail_Url")=request("Mail_Url")
  RS("Mail_SendType")=request("Mail_SendType")
  RS("ContaCtor")=request("ContaCtor")
  RS("Tel")=request("Tel")
  RS("Fax")=request("Fax")
  RS("EMail")=request("EMail")
  RS("Zip_Code")=request("Zip_Code")
  RS("City_Code")=request("City_Code")
  RS("Area_Code")=request("Area_Code")
  RS("Address")=request("Address")
  'RS("Zip_Code2")=request("Zip_Code2")
  'RS("City_Code2")=request("City_Code2")
  'RS("Area_Code2")=request("Area_Code2")
  'RS("Address2")=request("Address2")
  RS("Uniform_No")=request("Uniform_No")
  RS("Account")=request("Account")
  If request("CreateDate")<>"" Then
    RS("CreateDate")=request("CreateDate")
  Else
    RS("CreateDate")=null
  End If
  RS("Licence")=request("Licence")
  RS("Password_Day")=request("Password_Day")
  RS("Rept_Licence")=request("Rept_Licence")
  RS.Update
  RS.Close
  Set RS=Nothing
  session("errnumber")=1
  session("msg")="資料修改成功 !"
  'response.redirect "dept.asp"  
End If

If request("action")="delete" Then
  SQL="Delete From DEPT Where DEPT_ID='"&request("dept_id")&"' "
  Set RS=Conn.Execute(SQL)
  session("errnumber")=1
  session("msg")="資料刪除成功 !"
  Response.Redirect "dept.asp"
End If

'SQL1="Select * From DEPT_HIT"
'call QuerySQL(SQL1,RS1) 
'HitNo=RS1("HitNo")
'RS1.Close
'Set RS1=Nothing

SQL="Select * From DEPT Where DEPT_ID='"&request("dept_id")&"'"
call QuerySQL(SQL,RS) 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>機構部門管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" action="" method="post">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
                <tr>
                  <td width="5%"></td>
                  <td width="95%">
  		              <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                  <tr>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                        <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                      </tr>
                      <tr>
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">參數設定</td>
                        <td class="table63-bg"></td>
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
                  <td height="445" valign="top" class="table62-bg">&nbsp;</td>
                  <td height="445" valign="top">              
                    <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                      <%If Session("user_id")="npois" Then%>
                      <tr>
                        <td class="td02-c" width="100%" valign="top" colspan="3">
                          <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                            <tr>
                              <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool"><font color="#800000"><b>機構資訊設定</b></font></a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_system_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">網站參數設定</a></td>
                              <%If Session("user_id")="npois" And (Request.ServerVariables("REMOTE_HOST")="60.250.147.33" Or Request.ServerVariables("REMOTE_HOST")="127.0.0.1" Or Request.ServerVariables("REMOTE_HOST")="114.34.186.238") Then%>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_donate_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款參數設定</a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款機制說明</a></td>
                              <%Else%>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款機制說明</a></td>
                              <%End If%>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <%End If%>
                      <tr>
                        <td valign="top">&nbsp;</td>
                      </tr>                  
                      <tr>
                        <td class="td02-c" width="100%">
                          <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">                        
                            <tr>
                              <td width="12%" align="right"><font color="#000080">機構代碼<span lang="en-us">:</span></font></td>
                              <td width="38%" ><input type="text" name="Dept_Id" size="10" class="font9t" maxlength="10" readonly value="<%=RS("Dept_Id")%>"></td>
                              <td width="12%" align="right"> </td>
                              <td width="38%"> </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">機構名稱<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Comp_Name" size="40" class="font9" maxlength="60" value="<%=RS("Comp_Name")%>"></td>
                              <td align="right"><font color="#000080">機構Slogan<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Comp_Slogan" size="40" class="font9" maxlength="60" value="<%=RS("Comp_Slogan")%>"></td>	
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">機構簡稱<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Comp_ShortName" size="40" class="font9" maxlength="30" value="<%=RS("Comp_ShortName")%>"></td>
                              <td align="right"><font color="#000080">部門名稱<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Dept_Desc" size="40" class="font9" maxlength="20" value="<%=RS("Dept_Desc")%>"></td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">網站網址<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Server_Url" size="40" class="font9" maxlength="80" value="<%=RS("Server_Url")%>"></td>                        
                              <td align="right"><font color="#000080">網站名稱<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Sys_Name" size="40" class="font9" maxlength="60" value="<%=RS("Sys_Name")%>"></td>
                            </tr>                        
                            <tr>
                              <td align="right"><font color="#000080">寄件E-Mail<span lang="en-us">:</span></font></td>
                              <td><input type="text" name="Mail_Url" size="40" class="font9" maxlength="80" value="<%=RS("Mail_Url")%>"></td>
                              <td align="right"><font color="#000080">郵件寄送方式<span lang="en-us">:</span></font></td>
                              <td>
                              <%
                                SQL="Select Mail_SendType=CodeDesc From CASECODE Where CodeType='SendType' Order By Seq"
                                FName="Mail_SendType"
                                Listfield="Mail_SendType"
                                BoundColumn=RS("Mail_SendType")
                                menusize="1"
                                call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                              %>
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">聯絡人<span lang="en-us">:</span></font></td>
                              <td colspan="3">
                                <input type="text" name="ContaCtor" size="10" class="font9" maxlength="20" value="<%=RS("ContaCtor")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">電話<span lang="en-us">:</span></font>
                                <input type="text" name="Tel" size="15" class="font9" maxlength="20" value="<%=RS("Tel")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">傳真<span lang="en-us">:</span></font>
                                <input type="text" name="Fax" size="15" class="font9" maxlength="20" value="<%=RS("Fax")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">E-Mail<span lang="en-us">:</span></font>
                                <input type="text" name="EMail" size="34" class="font9" maxlength="60" value="<%=RS("EMail")%>">
                              </td>
                            </tr>
                            <!--
                            <tr>
                              <td align="right"><font color="#000080">會址<span lang="en-us">:</span></font></td>
                              <td colspan="3"><%call CodeCity ("form","Zip_Code2",RS("Zip_Code2"),"City_Code2",RS("City_Code2"),"Area_Code2",RS("Area_Code2"),"Address2",RS("Address2"),"74","N")%></td>
                            </tr>
                            -->
                            <tr>
                              <td align="right"><font color="#000080">聯絡地址<span lang="en-us">:</span></font></td>
                              <td colspan="3"><%call CodeCity ("form","Zip_Code",RS("Zip_Code"),"City_Code",RS("City_Code"),"Area_Code",RS("Area_Code"),"Address",RS("Address"),"74","Y")%></td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">統一編號<span lang="en-us">:</span></font></td>
                              <td colspan="3">
                                <input type="text" name="Uniform_No" size="10" class="font9" maxlength="8" value="<%=RS("Uniform_No")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">劃撥帳號<span lang="en-us">:</span></font>
                                <input type="text" name="Account" size="10" class="font9" maxlength="20" value="<%=RS("Account")%>">&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">成立日期<span lang="en-us">:</span></font>
                                <%call Calendar("CreateDate",RS("CreateDate"))%>&nbsp;&nbsp;&nbsp;&nbsp;
                                <font color="#000080">立案字號<span lang="en-us">:</span></font>
                                <input type="text" name="Licence" size="30" class="font9" maxlength="80" value="<%=RS("Licence")%>">
                              </td>
                            </tr>
                            <!--#include file="../include/calendar2.asp"-->                            
                            <tr>
                              <td align="right"><font color="#000080">密碼使用天數<span lang="en-us">:</span></font></td>
                              <td colspan="3">
                                <input type="text" name="Password_Day" size="5" class="font9" value="<%=RS("Password_Day")%>">&nbsp;&nbsp;(<font color="#FF0000">&nbsp;若設為&nbsp;0&nbsp;則密碼可無限期使用&nbsp;</font>)&nbsp;&nbsp;
                                <font color="#000080">勸募許可文號<span lang="en-us">:</span></font>
                                <%
                                  SQL1="Select Licence=CodeDesc From CASECODE Where CodeType='Licence' Order By Seq"
                                  Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                  RS1.Open SQL1,Conn,1,1
                                  Response.Write "<SELECT Name='Rept_Licence' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                                  Response.Write "<OPTION>" & " " & "</OPTION>"
                                  While Not RS1.EOF
                                    If Cstr(RS1("Licence"))=Cstr(RS("Rept_Licence")) Then
                                      Response.Write "<OPTION selected value='"&RS1("Licence")&"'>"&RS1("Licence")&"</OPTION>"
                                    Else
                                      Response.Write "<OPTION value='"&RS1("Licence")&"'>"&RS1("Licence")&"</OPTION>"
                                    End If
                                    RS1.MoveNext
                                  Wend
                                  Response.Write "</SELECT>"
                                  RS1.Close
                                  Set RS1=Nothing
                                %>
                                <!--<font color="#000080">訪客人數<span lang="en-us">:</span></font>
                                <input type="text" name="HitNo" size="13" readonly class="font9t" value="<%=FormatNumber(HitNo,0)%>">-->
                              </td>
                            </tr>                          
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td class="td02-c" width="100%" align="center"><%EditButton%></td>
                      </tr>
                    </table>
                  </td>
                  <td height="445" class="table63-bg">&nbsp;</td>
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
      </center></div>   
    </form>
  </center></div>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  <%call CheckStringJ("Dept_Id","機構代碼")%>
  <%call ChecklenJ("Dept_Id",10,"機構代碼")%>
  <%call CheckStringJ("Comp_Name","機構名稱")%>
  <%call ChecklenJ("Comp_Name",60,"機構名稱")%>
  <%call ChecklenJ("Comp_Slogan",60,"機構Slogan")%>
  <%call CheckStringJ("Comp_ShortName","機構簡稱")%>
  <%call ChecklenJ("Comp_ShortName",30,"機構代碼")%>
  <%call CheckStringJ("Dept_Desc","部門名稱")%>
  <%call ChecklenJ("Dept_Desc",20,"部門名稱")%>
  <%call ChecklenJ("Server_Url",80,"網站網址")%>
  <%call CheckStringJ("Sys_Name","網站名稱")%>
  <%call ChecklenJ("Sys_Name",60,"網站名稱")%>
  if(document.form.Mail_Url.value!=''){
    <%call CheckEmailJ("Mail_Url","寄件E-Mail")%>
    <%call ChecklenJ("Mail_Url",80,"寄件E-Mail")%>
  }
  <%call ChecklenJ("Mail_SendType",60,"郵件寄送方式")%>
  <%call ChecklenJ("ContaCtor",20,"聯絡人")%>
  <%call ChecklenJ("Tel",20,"電話")%>
  <%call ChecklenJ("Fax",20,"傳真")%>
  if(document.form.EMail.value!=''){
    <%call CheckEmailJ("EMail","E-Mail")%>
    <%call ChecklenJ("EMail",60,"E-Mail")%>
  }
  <%call ChecklenJ("Address",80,"路段號")%>
  if(document.form.Uniform_No.value!=''){
    <%call CheckNumberJ("Uniform_No","統一編號")%>
    if (document.form.Uniform_No.value.length!=8){
      alert('統一編號 欄位資料錯誤 ！');
      document.form.Uniform_No.focus();
      return;
    }
  }
  <%call ChecklenJ("Account",20,"劃撥帳號")%>
  <%call ChecklenJ("Licence",80,"立案字號")%>
  <%call CheckStringJ("Password_Day","密碼使用天數")%>
  <%call CheckNumberJ("Password_Day","密碼使用天數")%>
  <%call SubmitJ("update")%>
}

function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}

function Cancel_OnClick(){
  location.href="dept.asp"
}
--></Script>