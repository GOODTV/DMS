<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("Donor_Id_Begin")<>"" Then
  Donor_Id_Begin=request("Donor_Id_Begin")
Else
  SQL1="Select Donor_Id_Begin=Min(Donor_Id) From DONOR"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Donor_Id_Begin=RS1("Donor_Id_Begin")
  RS1.Close
  Set RS1=Nothing
End If

If request("Donor_Id_End")<>"" Then
  Donor_Id_End=request("Donor_Id_End")
Else
  SQL1="Select Donor_Id_End=Max(Donor_Id) From DONOR"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  Donor_Id_End=RS1("Donor_Id_End")
  RS1.Close
  Set RS1=Nothing
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>同地址過濾</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="check_address_rpt.asp">
    	<input type="hidden" name="action">
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
                      <td class="table62-bg">&nbsp;</td>
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">同地址過濾</td>
                      <td class="table63-bg">&nbsp;</td>
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
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="1" cellspacing="1">
                    <tr>
                      <td class="td02-c" width="100" align="right">捐款人編號：
                      <td class="td02-c" width="680"><input type="text" name="Donor_Id_Begin" size="9" class="font9" value="<%=Donor_Id_Begin%>"> ~ <input type="text" name="Donor_Id_End" size="9" class="font9" value="<%=Donor_Id_End%>"></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">過濾項目：</td>
                      <td class="td02-c">
                        <input type="radio" name="Check_Type" id="Check_Type1" value="IsDonor">捐款人同名同地址
                      	<input type="radio" name="Check_Type" id="Check_Type2" value="IsSendNews" checked >會訊同地址
                      	<input type="radio" name="Check_Type" id="Check_Type3" value="IsSendYNews">年報同地址
                      	<input type="radio" name="Check_Type" id="Check_Type4" value="IsBirthday">生日卡同名同地址
                      	<input type="radio" name="Check_Type" id="Check_Type5" value="IsXmas">賀卡同名同地址
                      </td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">過濾地址：</td>
                      <td class="td02-c">
                        <input type="radio" name="Address_Type" id="Address_Type1" value="Address" checked >通訊地址&nbsp;
                        <input type="radio" name="Address_Type" id="Address_Type2" value="Invoice_Address">收據地址	
                      </td>
                    </tr>
                    <tr>
                      <td width="100%" colspan="8" bgcolor="#C0C0C0" height="1"> </td>
                    </tr>
                    <tr>
                      <td colspan="8" width="100%">
                      <%
                        Response.Write "<button id='query' style='position:relative;left:30;width:45;height:40;font-size:9pt' class='font9' style='cursor:hand' onClick='Check_OnClick()'><img src='../images/search.gif' width='20' height='20'><br>過濾</button>&nbsp;"
    			              Response.Write "<span id='movebar' style='display:none'><img border='0' src='../images/movebar.gif'></span>"
    			            %>
              　      </td>
                    </tr>
                  </table>
                </td>
                <td class="table63-bg">&nbsp;</td>
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
function Check_OnClick(){
  <%call CheckStringJ("Donor_Id_Begin","捐款人編號啟")%>
  <%call CheckNumberJ("Donor_Id_Begin","捐款人編號啟")%>
  <%call CheckStringJ("Donor_Id_End","捐款人編號迄")%>
  <%call CheckNumberJ("Donor_Id_End","捐款人編號迄")%>   
  if(confirm('您是否確定要執行過濾？')){
    document.form.action.value="check"
    document.form.submit();
  }
}
--></script>	