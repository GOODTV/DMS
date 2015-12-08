<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function IssueDataGrid (SQL)
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
      ElseIf I=5 Then
        Response.Write "<td align=""center""><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I) & "</span></td>"
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

If request("action")="void" Then
  SQL1="Select * From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  If Not RS1.EOF Then
    While Not RS1.EOF
      Goods_Qty=CSng(RS1("Goods_Qty"))
      If RS1("Contribute_IsStock")="Y" And RS1("Goods_Id")<>"" Then
        SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
        Set RS2 = Server.CreateObject("ADODB.RecordSet")
        RS2.Open SQL2,Conn,1,3
        If Not RS2.EOF Then
          RS2("Goods_Qty")=CSng(RS2("Goods_Qty"))+CSng(RS1("Goods_Qty"))
          RS2.Update
        End If
        RS2.Close
        Set RS2=Nothing
      End If
      RS1("Goods_Qty")="0"
      RS1("Goods_Qty_D")=Goods_Qty
      RS1.MoveNext
    Wend
  End If
  RS1.Close
  Set RS1=Nothing
    
  SQL1="Select * From CONTRIBUTE_ISSUE Where Issue_Id='"&request("issue_id")&"'"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1("Issue_Type")="D"
  RS1.Update
  RS1.Close
  Set RS1=Nothing
  
  session("errnumber")=1
  session("msg")="領取單已作廢 ！"
End If

If request("action")="return" Then
  Check_Return=True
  Goods_Name=""
  SQL1="Select CONTRIBUTE_ISSUEDATA.*,Storage_Qty=Isnull(GOODS.Goods_Qty,0),Goods_IsStock=GOODS.Goods_IsStock From CONTRIBUTE_ISSUEDATA Left Join GOODS On CONTRIBUTE_ISSUEDATA.Goods_Id=GOODS.Goods_Id Where CONTRIBUTE_ISSUEDATA.Issue_Id='"&request("issue_id")&"'"
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    While Not RS1.EOF
      If RS1("Contribute_IsStock")="Y" And RS1("Goods_IsStock")="Y" And CSng(RS1("Goods_Qty_D"))>CSng(RS1("Storage_Qty")) Then
        Check_Return=False
        If Goods_Name="" Then
          Goods_Name=RS1("Goods_Name")
        Else
          Goods_Name=Goods_Name&"、"&RS1("Goods_Name")
        End If
      End If
      RS1.MoveNext
    Wend
    If Check_Returns=False Then
      session("errnumber")=1
      session("msg")="『"&Goods_Name&"』 \n\n庫存量不足領取單無法還原 ！"
    End If
  End If
  RS1.Close
  Set RS1=Nothing
  
  If Check_Return Then
    SQL1="Select * From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    If Not RS1.EOF Then
      While Not RS1.EOF
        Goods_Qty=CSng(RS1("Goods_Qty_D"))
        If RS1("Contribute_IsStock")="Y" And RS1("Goods_Id")<>"" Then
          SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,3
          If Not RS2.EOF Then
            RS2("Goods_Qty")=CSng(RS2("Goods_Qty"))-CSng(RS1("Goods_Qty_D"))
            RS2.Update
          End If
          RS2.Close
          Set RS2=Nothing
        End If
        RS1("Goods_Qty")=Goods_Qty
        RS1("Goods_Qty_D")="0"
        RS1.MoveNext
      Wend
    End If
    RS1.Close
    Set RS1=Nothing
    
    SQL1="Select * From CONTRIBUTE_ISSUE Where Issue_Id='"&request("issue_id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    RS1("Issue_Type")=""
    RS1.Update
    RS1.Close
    Set RS1=Nothing  
  
    session("errnumber")=1
    session("msg")="領取單已還原 ！"
  End If
End If

SQL="Select *,Comp_ShortName From CONTRIBUTE_ISSUE Join Dept On CONTRIBUTE_ISSUE.Dept_Id=DEPT.Dept_Id  Where Issue_Id='"&request("issue_id")&"' "
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="issue"%>
<!--#include file="../include/head.asp"-->
<body>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
    	<input type="hidden" name="action" value="">
      <input type="hidden" name="issue_id" value="<%=RS("Issue_Id")%>">
      <%If RS("Issue_Print")="1" Then%>	
      <input type="hidden" name="RePrint" value="Y">
      <%Else%>
      <input type="hidden" name="RePrint" value="">
      <%End If%>
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
                  <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                    <tr>
                      <td class="td02-c" width="100%">
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
                          <tr>
                            <td width="10%" align="right">機構名稱：</td>
                            <td width="15% align="left"><input type="text" name="Comp_ShortName" size="15" class="font9t" readonly value="<%=RS("Comp_ShortName")%>"></td>
                            <td width="8%" align="right">領取人：</td>
                            <td width="17%" align="left"><input type="text" name="Issue_Processor" size="17" maxlength="20" class="font9t" readonly value="<%=RS("Issue_Processor")%>"></td>
                            <td width="10%" align="right">領用日期：</td>
                            <td width="15% align="left"><input type="text" name="Comp_ShortName" size="12" class="font9t" readonly value="<%=RS("Issue_Date")%>"></td>
                            <td width="10%" align="right">領用用途：</td>
                            <td width="15%" align="left"><input type="text" name="Issue_Purpose" size="15" class="font9t" readonly value="<%=RS("Issue_Purpose")%>"></td>
                          </tr>
                          <tr>
                            <td align="right">出貨單位：</td>
                            <td><input type="text" name="Issue_Org" size="15" maxlength="40" class="font9t" readonly value="<%=RS("Issue_Org")%>"></td>
                            <td align="right">備註：</td>
                            <td colspan="2"><input type="text" name="Issue_Comment" size="30" maxlength="100" class="font9t" readonly value="<%=RS("Issue_Comment")%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" <%If RS("Issue_No")="M" Then Response.Write "checked" End If%> disabled >手開領用編號</td>
                            <td align="right">領用編號：</td>
                            <td><input type="text" name="Issue_No" size="15" class="font9t" readonly maxlength="20" value="<%=RS("Issue_Pre")&RS("Issue_No")%>"></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">領用內容：</td>
                            <td align="left" colspan="7">
                            <%
                              SQL="Select ser_no,物品名稱=Goods_Name,數量=Goods_Qty,單位=Goods_Unit,備註=Goods_Comment,庫存品=(Case When Contribute_IsStock='Y' Then 'V' Else '' End) " & _
                                  "From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&RS("Issue_Id")&"' Order By Ser_No "
                              HLink="contribute_issue_edit.asp.asp?issue_id="&RS("Issue_Id")&"&ser_no="
                              LinkParam="ser_no"
                              LinkTarget="main"
                              LinkType="window"
                              Call IssueDataGrid (SQL)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                              <%
                                If RS("Issue_Type")<>"D" Then
                                  Response.Write"<input type=""button"" value=""修改領用資料"" name=""Update"" class=""cbutton"" style=""cursor:hand"" onClick=""Update_OnClick()"">&nbsp;&nbsp;"
                                  If RS("Issue_Print")="1" Then
                                    Response.Write"<input type=""button"" value=""單據補印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  Else
                                    Response.Write"<input type=""button"" value=""單據列印"" name=""Invoice"" class=""cbutton"" style=""cursor:hand"" onClick=""Invoice_OnClick()"">&nbsp;&nbsp;"
                                  End If
                                  Response.Write"<input type=""button"" value=""單據作廢"" name=""Void"" class=""cbutton"" style=""cursor:hand"" onClick=""Void_OnClick()"">&nbsp;&nbsp;"
                                Else
                                  Response.Write"<input type=""button"" value=""還原作廢單據"" name=""Return"" class=""cbutton"" style=""cursor:hand"" onClick=""Return_OnClick()"">&nbsp;&nbsp;"
                                End If
                              %>
				                      <input type="button" value=" 取消 " name="cancel" class="cbutton" style="cursor:hand" onClick="Cancel_OnClick()">
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
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){
  location.href='contribute_issue_edit.asp?issue_id='+document.form.issue_id.value+'';
}
function Invoice_OnClick(){
  if(confirm('您是否確定要列印單據？')){
    window.open('contribute_issue_print.asp?issue_id='+document.form.issue_id.value+'','issue_rpt','toolbar=no,location=no,status=no,menubar=yes,scrollbars=yes,resizable=yes,top=0,left=0,width=800,height=600');
  }
}
function Void_OnClick(){
  if(document.form.RePrint.value==''){
    alert('作廢單據前請先列印！');
    return;
  }
  if(confirm('您是否確定要作廢單據？')){
    document.form.action.value='void';
    document.form.submit();
  }
}
function Return_OnClick(){
  if(confirm('您是否確定要還原作廢單據？')){
    document.form.action.value='return';
    document.form.submit();
  }
}
function Cancel_OnClick(){
  location.href='contribute_issue.asp';
}
--></script>	