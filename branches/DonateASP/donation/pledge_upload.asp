<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="pledge_upload"%>
<!--#include file="../include/head.asp"-->
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="pledge_upload_add.asp?MaxFileSize=10" method="post" enctype="multipart/form-data">
      <input type="hidden" name="action">
      <input type="hidden" name="Object_ID" value="0">	
      <input type="hidden" name="Ap_Name" value="pledge">
      <input type="hidden" name="Upload_FileName" value="轉帳授權書下載">		
      <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%></td>
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
                      <td class="td02-c" width="120" align="right">轉帳授權書上傳：</td>
                      <td class="td02-c" width="660">
                      	<input type="file" class="font9" size="50" name="Upload_FileURL">
                        <input type="button" value=" 重新上傳 " name="save" class="cbutton" style="cursor:hand" OnClick="Upload_OnClick()">&nbsp;&nbsp;
                        <%
                          Download_FileURL="../upload/定額捐款授權4合一表.doc"
                          Download_FileName="定額捐款授權4合一表.doc"
                          SQL1="Select TOP 1 * From UPLOAD Where Object_ID='0' And Ap_Name='pledge' And Attach_Type='doc' Order By Ser_No Desc"
                          Set RS1 = Server.CreateObject("ADODB.RecordSet")
                          RS1.Open SQL1,Conn,1,1
                          If Not RS1.EOF Then
                            Download_FileURL="../upload/"&RS1("Upload_FileURL")
                            Download_FileName=RS1("Upload_FileURL")
                          End If
                          RS1.Close
                          Set RS1=Nothing
                        %>
                        <input type="button" value="轉帳授權書下載" name="download" class="addbutton" style="cursor:hand" OnClick="MM_openBrWindow('../upload/<%=Download_FileURL%>','download','toolbar=yes,location=yes,status=yes,menubar=yes,scrollbars=yes,resizable=yes,top=84,left=112,width=1024,height=768')">
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
function Upload_OnClick(){
  <%call CheckStringJ("Upload_FileURL","轉帳授權書")%>
  <%call CheckExtJ("Upload_FileURL","doc","轉帳授權書")%>
  <%call SubmitJ("upload")%>
}
--></script>