<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function GiftDataGrid (SQL)
	Row=0
	Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  FieldsCount = RS1.Fields.Count-1
  
  Dim I
  Response.Write "<table id=grid border='1' cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
  Response.Write "<tr>"
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>序號</span></font></td>"
  For I = 1 To FieldsCount  	
	  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
  Next
  Response.Write "</tr>"
  
  While Not RS1.EOF
  	Response.Write "<tr>"
  	For I = 1 To FieldsCount
  		If I=1 Then
        Row=Row+1
        Response.Write "<td align=""center""><span style='font-size: 9pt; font-family: 新細明體'>" & Row & "</span></td>"
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
			ElseIf I=2 Then
        Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"   
      Else
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"     
      End If
  	Next
  	RS1.MoveNext
  	Response.Write "</tr>"
	Wend
	  	
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

SQL="Select GIFT.*,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Donor_Type=DONOR.Donor_Type,IDNo=DONOR.IDNo, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address Else DONOR.ZipCode+A.mValue+DONOR.Address End End) " & _    
    "From GIFT " & _
    "Join DONOR On GIFT.Donor_Id=DONOR.Donor_Id " & _   
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Where Gift_Id='"&request("gift_id")&"' "
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="gift"%>
<!--#include file="../include/head.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
    	<input type="hidden" name="action" value="">
    	<input type="hidden" name="ctype" value="<%=request("ctype")%>">
    	<input type="hidden" name="Donor_Id" value="<%=RS("Donor_Id")%>">
      <input type="hidden" name="gift_id" value="<%=RS("Gift_Id")%>">     
      
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">公關贈品維護</td>
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
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
                          <tr>
                            <td align="right">捐款人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">                              
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="42" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>                                                
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>                          
                          <tr>
                            <td width="10%" align="right">贈送日期：</td>
                            <td width="16%" align="left" colspan="3"><input type="text" name="Gift_Date" size="12" class="font9t" readonly value="<%=RS("Gift_Date")%>"></td>                            
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">備註：</td>
                            <td align="left" colspan="5"><textarea rows="3" name="Comment" cols="42" class="font9t" readonly ><%=RS("Comment")%></textarea></td>                            
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                          	<td align="right">贈品內容：</td>
                            <td align="left" colspan="7">
                            <%
                              SQL="Select Ser_No,物品名稱=Goods_Name,數量=Goods_Qty,備註=Goods_Comment " & _
                                  "From GIFTDATA Where Gift_Id='"&RS("gift_id")&"' Order By Ser_No "
                              HLink="contribute_data_edit.asp?gift_id="&RS("Gift_id")&"&ser_no="
                              LinkParam="ser_no"
                              LinkTarget="main"
                              LinkType="window"
                              Call GiftDataGrid (SQL)
                            %>
                            </td>                            
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                              <%
                                Response.Write"<input type=""button"" value=""修改公關品資料"" name=""Update"" class=""cbutton"" style=""cursor:hand"" onClick=""Update_OnClick()"">&nbsp;&nbsp;"                              
                              %>
				                      <input type="button" value=" 取消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">&nbsp;&nbsp;				                      
                            </td>
                          </tr>
                        </table>
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
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  location.href='gift_edit.asp?ctype='+document.form.ctype.value+'&gift_id='+document.form.gift_id.value+'';
}
function Cancel_OnClick(){	
    location.href='contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';
}
--></script>	