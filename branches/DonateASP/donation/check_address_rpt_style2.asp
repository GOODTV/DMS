<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then  
  Ary_Donor_Id=Split(request("all_donor_id"),",")
  For I = 0 To UBound(Ary_Donor_Id)
    SQL="Update DONOR Set "&request("Check_Type")&"='N' Where Donor_Id='"&Trim(Ary_Donor_Id(I))&"'"
    Set RS=Conn.Execute(SQL)
  Next
  
  Ary_Donor_Id=Split(request("donor_id"),",")
  For I = 0 To UBound(Ary_Donor_Id)
    SQL="Update DONOR Set "&request("Check_Type")&"='Y' Where Donor_Id='"&Trim(Ary_Donor_Id(I))&"'"
    Set RS=Conn.Execute(SQL)
  Next

  session("errnumber")=1
  session("msg")=request("Check_Name")&"訂閱名單修改成功 ！"
  Response.Redirect "check_address.asp"
End If

SQL1=Session("SQL1")
Call QuerySQL(SQL1,RS1)
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
    <form name="form" method="post" action="">
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
              <%
                Total_Row=0
                All_Donor_Id=""
                bgcolor="#FFFF99"
                While Not RS1.EOF
                  SQL2="Select Donor_Id,Donor_Name,Invoice_Title From DONOR Where CONVERT(int,DONOR.Donor_Id)<'"&RS1("Donor_Id")&"' "
                  If Session("Check_Type")="IsSendNews" Then
                    SQL2=SQL2&"And IsSendNews='Y' "
                    Check_Name="會訊"
                  End If
                  If Session("Check_Type")="IsSendYNews" Then
                    SQL2=SQL2&"And IsSendYNews='Y' "
                    Check_Name="年報"
                  End If
                  If Session("Check_Type")="IsBirthday" Then
                    SQL2=SQL2&"And IsBirthday='Y' "
                    Check_Name="生日卡"
                  End If
                  If Session("Check_Type")="IsXmas" Then
                    SQL2=SQL2&"And IsXmas='Y' "
                    Check_Name="賀年(耶誕)卡"
                  End If
                  If Session("Address_Type")="Address" Then
                    SQL2=SQL2&"And City='"&RS1("City")&"' And Area='"&RS1("Area")&"' And Address='"&RS1("Address")&"' "
                  Else
                    SQL2=SQL2&"And Invoice_City='"&RS1("Invoice_City")&"' And Invoice_Area='"&RS1("Invoice_Area")&"' And Invoice_Address='"&RS1("Invoice_Address")&"' "
                  End If
                  SQL2=SQL2&"Order By Donor_Id Desc"
                  Call QuerySQL(SQL2,RS2)
                  If Not RS2.EOF Then
                    If bgcolor="#FFFF99" Then
                      bgcolor="#FFFFFF"
                    Else
                      bgcolor="#FFFF99"
                    End If
                    Total_Row=Total_Row+1
                    If All_Donor_Id="" Then
                      All_Donor_Id=RS1("Donor_Id")
                    Else
                      All_Donor_Id=All_Donor_Id&","&RS1("Donor_Id")
                    End If
                    If Total_Row=1 Then
                      Response.Write "<tr><td class='table62-bg'> </td><td valign='top' align='left'>&nbsp;&nbsp;&nbsp;<input type=""button"" value=""修改"&Check_Name&"訂閱名單"" name=""update"" class=""addbutton"" style=""cursor:hand"" onClick=""Update_OnClick()"">&nbsp;&nbsp;&nbsp;<input type=""button"" value=""離  開"" name=""cancel"" class=""cbutton"" style=""cursor:hand"" onClick=""Cancel_OnClick()""></td><td class='table63-bg'>　</td></tr>"&vbcrlf
                      Response.Write "<tr><td class='table62-bg'> </td><td valign='top' align='left' style='color:#ff0000'><br />※&nbsp;系統僅勾選最早訂閱之捐款人，其餘未勾選之捐款人經重新<font color='#3366CC'>【修改後將視同不再訂閱】</font>，如捐款人須繼續訂閱請打""V""。</td><td class='table63-bg'>　</td></tr>"&vbcrlf
                      Response.Write "<tr>"&vbcrlf
                      Response.Write "  <td class='table62-bg'>　</td>"&vbcrlf
                      Response.Write "  <td valign='top'>"&vbcrlf
                      Response.Write "    <table border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"&vbcrlf
                      Response.Write "      <tr>"&vbcrlf
                      Response.Write "        <td bgcolor='#FFE1AF'><input type='checkbox' name='donor_id' id='donor_id_0' value='0' OnClick='DonorId_OnClick()'></td>"&vbcrlf
                      Response.Write "        <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編號</span></font></td>"&vbcrlf
                      Response.Write "        <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>捐款人</span></font></td>"&vbcrlf
                      Response.Write "        <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>收據抬頭</span></font></td>"&vbcrlf
                      If Session("Address_Type")="Address" Then
                        Response.Write "      <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>通訊地址</span></font></td>"&vbcrlf
                      Else
                        Response.Write "      <td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>收據地址</span></font></td>"&vbcrlf
                      End If
                      Response.Write "      </tr>"&vbcrlf
                    End If
                    Response.Write "<tr bgcolor='"&bgcolor&"' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor="""&bgcolor&"""'>"&vbcrlf
                    Response.Write "  <td><input type='checkbox' name='donor_id' id='donor_id_"&Total_Row&"' value='"&RS1("Donor_Id")&"'></td>"&vbcrlf
                    Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Donor_Id")&"</span></td>"&vbcrlf
                    Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Donor_Name")&"</span></td>"&vbcrlf
                    Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Invoice_Title")&"</span></td>"&vbcrlf
                    If Session("Address_Type")="Address" Then
                      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Address2")&"</span></td>"&vbcrlf
                    Else
                      Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Invoice_Address2")&"</span></td>"&vbcrlf
                    End If
                    Response.Write "</tr>"&vbcrlf
                    Row2=0
                    While Not RS2.EOF
                      Total_Row=Total_Row+1
                      Row2=Row2+1
                      All_Donor_Id=All_Donor_Id&","&RS2("Donor_Id")
                      Response.Write "<tr bgcolor='"&bgcolor&"' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor="""&bgcolor&"""'>"&vbcrlf
                      If Row2=RS2.Recordcount Then
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='donor_id' id='donor_id_"&Total_Row&"' value='"&RS2("Donor_Id")&"' checked ></span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體; font-weight: bold; color: #3366CC'>"&RS2("Donor_Id")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體; font-weight: bold; color: #3366CC'>"&RS2("Donor_Name")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體; font-weight: bold; color: #3366CC'>"&RS2("Invoice_Title")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體; font-weight: bold; color: #3366CC'>同上</span></td>"&vbcrlf
                      Else
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='donor_id' id='donor_id_"&Total_Row&"' value='"&RS2("Donor_Id")&"'></span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS2("Donor_Id")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS2("Donor_Name")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS2("Invoice_Title")&"</span></td>"&vbcrlf
                        Response.Write "  <td><span style='font-size: 9pt; font-family: 新細明體'>同上</span></td>"&vbcrlf
                      End If
                      Response.Write "</tr>"&vbcrlf
                      RS2.MoveNext
		                Wend
                  End If
                  RS2.Close
		              Set RS2=Nothing
                  Response.Flush
                  Response.Clear
                  RS1.MoveNext
                Wend
		            RS1.Close
		            Set RS1=Nothing
		            Response.Write "<input type='hidden' name='Check_Type' value='"&Session("Check_Type")&"'>"&vbcrlf
		            Response.Write "<input type='hidden' name='Check_Name' value='"&Check_Name&"'>"&vbcrlf
		            Response.Write "<input type='hidden' name='Total_Row' value='"&Total_Row&"'>"&vbcrlf
                Response.Write "<input type='hidden' name='All_Donor_Id' value='"&All_Donor_Id&"'>"&vbcrlf
                If Total_Row>0 Then 
                  Response.Write "    </table>"&vbcrlf
                  Response.Write "  </td>"&vbcrlf
                  Response.Write "  <td class='table63-bg'>　</td>"&vbcrlf
                  Response.Write "</tr>"&vbcrlf
                Else
                  Response.Write "<tr><td class='table62-bg'> </td><td valign='top' align='center' style='color:#ff0000'><br />** 沒有符合條件的資料 **<br /><br /><input type=""button"" value=""離  開"" name=""cancel"" class=""cbutton"" style=""cursor:hand"" onClick=""Cancel_OnClick()""></td><td class='table63-bg'>　</td></tr>"&vbcrlf
		            End If
		            Session.Contents.Remove("Check_Type")
                Session.Contents.Remove("Address_Type")
                Session.Contents.Remove("SQL1")
              %>
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
function DonorId_OnClick(){
  if(document.form.donor_id[0].checked){
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.donor_id[i].checked=true;
    }
  }else{
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.donor_id[i].checked=false;
    }
  }
}
function Update_OnClick(){
  if(confirm('您是否確定要修改'+document.form.Check_Name.value+'訂閱名單？')){
    document.form.action.value="update"
    document.form.submit();
  }
}
function Cancel_OnClick(){
  location.href='check_address.asp';
}
--></script>