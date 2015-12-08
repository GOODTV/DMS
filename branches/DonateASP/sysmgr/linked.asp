<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("linked_id")<>"" Then
  linked_id=request("linked_id")
Else
  linked_id="0"
  If request("linked_type")<>"" Then
    SQL="Select TOP 1 Ser_No From LINKED Where Linked_Type='"&request("linked_type")&"' Order By Linked_Seq"
  Else
    SQL="Select TOP 1 Ser_No From LINKED Order By Linked_Seq"
  End If
  
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then linked_id=RS("Ser_No")
End If      
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title><%=request("subject")%>連動選單維護</title>
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
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=request("subject")%>連動選單維護</td>
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
                        <td class="td02-c" width="400"><iframe name="left" src="linked_list.asp?linked_type=<%=request("linked_type")%>&subject=<%=request("subject")%>" height="455" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                        <%If Cint(linked_id)>0 Then%>
                        <td class="td02-c" width="440"><iframe name="detail" src="linked2_add.asp?linked_type=<%=request("linked_type")%>&subject=<%=request("subject")%>&linked_id=<%=linked_id%>" height="400" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                        <%Else%>
                        <td class="td02-c" width="440"> </td>
                        <%End If%>
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