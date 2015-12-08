<!--#include file="../include/dbfunctionJ.asp"-->
<%
SQL_IS="Select * From DEPT Where Dept_Id='"&Session("dept_id")&"' "
Set RS_IS=Server.CreateObject("ADODB.Recordset")
RS_IS.Open SQL_IS,Conn,1,1
IsEpaper=RS_IS("IsMember")
IsDonation=RS_IS("IsDonation")
IsContribute=RS_IS("IsContribute")
IsMember=RS_IS("IsMember")
IsShopping=RS_IS("IsShopping")
RS_IS.Close
Set RS_IS=Nothing
Category2="'sys'"
If IsEpaper="Y" Then Category2=Category2&",'epaper'"
If IsDonation="Y" Then Category2=Category2&",'donation'"
If IsContribute="Y" Then Category2=Category2&",'contribute'"
If IsMember="Y" Then Category2=Category2&",'member'"
If IsShopping="Y" Then Category2=Category2&",'shopping'"

If request("codetype")<>"" Then
  codetype=request("codetype")
Else
  SQL="Select *,'1' As Seq From CODEFILE Where Category<>'類別選項' " & _
      "Union " & _
      "Select *,'2' As Seq From CODEFILE Where Category='類別選項' And Category2 In ("&Category2&") " & _
      "Order By Seq,Menu_Seq,Category,CodeType"
  Call QuerySQL(SQL,RS)
  If Not RS.EOF Then codetype=RS("CodeType")
  RS.Close
  Set RS=Nothing
End If      
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>代碼主檔列表</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="760" border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">代碼主檔維護</td>
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
	            <table width="820" border="0" cellspacing="0" cellpadding="0" align="center">
                <tr>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                  <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                  <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                </tr>
                <tr>
                  <td class="table62-bg">&nbsp;</td>
                  <td valign="top">
                    <table width="100%"  border="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111">
                      <tr>
                        <td class="td02-c" width="400"><iframe name="left" src="codefile_list.asp" height="455" width="100%" frameborder="0" scrolling="auto"></iframe>　</td>
                        <td class="td02-c" width="440"><iframe name="detail" src="casecode_add.asp?codetype=<%=codetype%>" height="400" width="100%" frameborder="0" scrolling="auto"></iframe>　</td>
                      </tr>
                      <tr>
                        <td width="702" colspan="2" bgcolor="#C0C0C0" height="1"> </td>
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
      </center></div>
    </form>
  </center></div>
  <%Message()%>
</body>
</html>