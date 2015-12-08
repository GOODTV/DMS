<!--#include file="../include/dbfunctionJ.asp"-->
<%
Transfer_Row=0
If request("action")="query" Then
  SQL1="Select Donor_Name=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),Address=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where Donor_Id='"&request("From_Donor_Id")&"'"
  Call QuerySQL(SQL1,RS1)
  If RS1.EOF Then
    session("errnumber")=1
    session("msg")="您輸入的『轉出』捐款人編號不存在，請您重新確認 ！"
    response.redirect "donate_merge.asp"
  Else
    SQL2="Select Transfer_Row=Count(*) From DONATE Where Donor_Id='"&request("From_Donor_Id")&"'"
    Call QuerySQL(SQL2,RS2)
    Transfer_Row=RS2("Transfer_Row")
    If Transfer_Row=0 Then
      session("errnumber")=1
      session("msg")="您輸入的『轉出』捐款人編號無捐款紀錄，請您重新確認 ！"
    End If
    RS2.Close
    Set RS2=Nothing
  End If
  
  SQL2="Select Donor_Name=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),Address=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where Donor_Id='"&request("To_Donor_Id")&"'"
  Call QuerySQL(SQL2,RS2)
  If RS2.EOF Then
    session("errnumber")=1
    session("msg")="您輸入的『轉入』捐款人編號不存在，請您重新確認 ！"
    response.redirect "donate_merge.asp"
  End If
End If

If request("action")="transfer" Then
  If Trim(request("donate_id"))<>"" Then
    Row=0
    Ary_Donate_Id=Split(Trim(request("donate_id")),",")
    For I = 0 To UBound(Ary_Donate_Id)
      SQL1="Select Donor_Id,Donate_Merge,Invoice_Title From DONATE Where Donate_Id='"&Trim(Ary_Donate_Id(I))&"'"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      If Not RS1.EOF Then
        RS1("Donor_Id")=request("To_Donor_Id")
        RS1("Donate_Merge")=request("From_Donor_Id")
        RS1("Invoice_Title")=request("Donor_Name")
        RS1.Update
        Row=Row+1
      End If
      RS1.Close
      Set RS1=Nothing 
    Next
    
    '2014/6/19 取消修改授權書紀錄
    '修改授權書紀錄
    'SQL1="Update PLEDGE Set Donor_Id='"&request("To_Donor_Id")&"' Where Donor_Id='"&request("From_Donor_Id")&"' "
    'Set RS1=Conn.Execute(SQL1)
        
    '修改捐款人捐款紀錄
    call Declare_DonorId (request("From_Donor_Id"))
    call Declare_DonorId (request("To_Donor_Id"))
                
    SQL1="Select Transfer_Row=Count(*) From DONATE Where Donor_Id='"&request("From_Donor_Id")&"'"
    Call QuerySQL(SQL1,RS1)
    Transfer_Row=RS1("Transfer_Row")
    RS1.Close
    Set RS1=Nothing
            
    SQL1="Select Donor_Name=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),Address=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where Donor_Id='"&request("From_Donor_Id")&"'"
    Call QuerySQL(SQL1,RS1)
    SQL2="Select Donor_Name=(Case When NickName<>'' Then Donor_Name+'('+NickName+')' Else Donor_Name End),Address=(Case When DONOR.City='' Then Address Else Case When A.mValue<>B.mValue Then A.mValue+DONOR.ZipCode+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End) From DONOR Left Join CODECITY As A On DONOR.City=A.mCode Left Join CODECITY As B On DONOR.Area=B.mCode Where Donor_Id='"&request("To_Donor_Id")&"'"
    Call QuerySQL(SQL2,RS2)
    session("errnumber")=1
    session("msg")="捐款資料已合併成功，共計"&FormatNumber(Row,0)&"筆！"
  Else
    session("errnumber")=1
    session("msg")="您尚未勾選任何轉出資料 ！" 
  End If
End If
%>
<%Prog_Id="donate_merge"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <table width="700" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
        <tr>
          <td>
            <table width="100%"  border="0" cellspacing="0" cellpadding="0" align="center" style="background-color:#EEEEE3">
              <tr>
                <td width="5%"></td>
                <td width="95%">
  		            <table width="60%"  border="0" cellspacing="0" cellpadding="0">
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
	          <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                <td width="10" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
              </tr>
              <tr>
                <td class="table62-bg">&nbsp;</td>
                <td valign="top">
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right" width="90">捐款人編號：</td>
                      <td class="td02-c" width="270">
                        <input type="text" name="From_Donor_Id" size="10" class="font9" value="<%=request("From_Donor_Id")%>" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'> 轉出 
				                --&gt;
                        <input type="text" name="To_Donor_Id" size="10" class="font9" value="<%=request("To_Donor_Id")%>" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'> 轉入
                      </td>
                      <td class="td02-c" width="420">
                        <input type="button" value=" 查 詢 " name="query" class="cbutton" style="cursor:hand" onClick="Query_OnClick()">&nbsp;
                        <%If Cint(Transfer_Row)>0 Then%><input type="button" value="資料合併" name="transfer" class="addbutton" style="cursor:hand" onClick="Transfer_OnClick()"><%End If%>
                      </td>
                    </tr>
                  </table>
                  <%If request("action")="query" Or request("action")="transfer" Then%>
                  <table width="100%" border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111">
                    <tr>
                      <td class="td02-c" align="right" width="90">轉出捐款人：</td>
                      <td class="td02-c" width="295"><%=RS1("Donor_Name")%></td>
                      <td class="td02-c" width="10"> </td>
                      <td class="td02-c" align="right" width="90">轉入捐款人：</td>
                      <td class="td02-c" width="295"><%=RS2("Donor_Name")%>
                          <input type="hidden" id="Donor_Name" name="Donor_Name" value="<%=RS2("Donor_Name")%>"/></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">通訊地址：</td>
                      <td class="td02-c"><%=RS1("Address")%></td>
                      <td class="td02-c" width="10"> </td>
                      <td class="td02-c" align="right">通訊地址：</td>
                      <td class="td02-c"><%=RS2("Address")%></td>
                    </tr>
                    <tr>
                      <td class="td02-c" align="right">欲轉出資料：</td>
                      <td class="td02-c"><font color="#FF0000">(&nbsp;請先勾選然後按下<font color="#008700">『&nbsp;資料合併&nbsp;』</font>按鈕)</font>&nbsp;</td>
                      <td class="td02-c" width="10"> </td>
                      <td class="td02-c" align="right">欲轉入資料：</td>
                      <td class="td02-c"> </td>
                    </tr>
                    <tr>
                      <td class="td02-c" colspan="2" valign="top">
                      <%
                        RS1.Close
                        Set RS1=Nothing
                        RS2.Close
                        Set RS2=Nothing
                        '20131105 Modify by GoodTV Tanya:修正收據編號未顯示問題
                        SQL1="Select Donate_Id,捐款日期=Donate_Date,捐款方式=Donate_Payment,捐款金額=Donate_Amt,機構=(Select Comp_ShortName From DEPT Where DEPT.Dept_Id=DONATE.Dept_Id),收據編號=IsNull(Invoice_Pre,'')+IsNull(Invoice_No,'') " & _
                             "From DONATE Where Donor_Id='"&request("From_Donor_Id")&"' Order By Donate_Date Desc,Donate_Id Desc"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        FieldsCount = RS1.Fields.Count-1
                        Dim I, J
                        Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
                        Response.Write "<tr>"
                        If RS1.Recordcount>0 Then
                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='donate_id' id='donate_id_0' value='0' checked OnClick='DonateId_OnClick()'></span></font></td>"
                        Else
                          Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>選擇</span></font></td>"
                        End If
                        For I = 1 To FieldsCount
	                        Response.Write "<td bgcolor='#FFE1AF'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
                        Next
                        Response.Write "</tr>"
                        
                        StrChecked=""
                        If request("action")="query" Then StrChecked="checked"
                        Row=0
                        While Not RS1.EOF
                          Row=Row+1
                          Response.Write "<tr>"
                          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'><input type='checkbox' name='donate_id' id='donate_id_"&Row&"' value='"&RS1(0)&"' "&StrChecked&" >" & "</span></td>"
                          For J = 1 To FieldsCount
                            If J=3 Then
                              If RS1(J)<>"" Then
                                Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
                              Else
                                Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
                              End If
                            Else
                              Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
                            End If
                          Next
                          Response.Write "</tr>"
                          Response.Flush
                          Response.Clear
                          RS1.MoveNext
                        Wend  
                        Response.Write "</table>"
                        RS1.Close
                        Set RS1=Nothing
                      %>
                      </td>
                      <input type="hidden" name="Total_Row" value="<%=Row%>">
                      <td class="td02-c" width="10"> </td>
                      <td class="td02-c" colspan="2" valign="top">
                      <%
                        SQL1="Select Donate_Id,捐款日期=Donate_Date,捐款方式=Donate_Payment,捐款金額=Donate_Amt,機構=(Select Comp_ShortName From DEPT Where DEPT.Dept_Id=DONATE.Dept_Id),收據編號=IsNull(Invoice_Pre,'')+IsNull(Invoice_No,'') " & _
                             "From DONATE Where Donor_Id='"&request("To_Donor_Id")&"' Order By Donate_Date Desc,Donate_Id Desc"
                        Set RS1 = Server.CreateObject("ADODB.RecordSet")
                        RS1.Open SQL1,Conn,1,1
                        FieldsCount = RS1.Fields.Count-1
                        Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;background-color: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
                        Response.Write "<tr>"
                        For I = 1 To FieldsCount
	                        Response.Write "<td bgcolor='#FFFF00'><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(I).Name & "</span></font></td>"
                        Next
                        Response.Write "</tr>"
                        
                        Row=0
                        While Not RS1.EOF
                          Row=Row+1
                          Response.Write "<tr>"
                          For J = 1 To FieldsCount
                            If J=3 Then
                              If RS1(J)<>"" Then
                                Response.Write "<td align='right' height='25'><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
                              Else
                                Response.Write "<td> height='25'<span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
                              End If
                            Else
                              Response.Write "<td height='25'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
                            End If
                          Next
                          Response.Write "</tr>"
                          Response.Flush
                          Response.Clear
                          RS1.MoveNext
                        Wend  
                        Response.Write "</table>"
                        RS1.Close
                        Set RS1=Nothing
                      %>
                      </td>
                    </tr>
                  </table>
                  <%End If%>
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
function Window_OnLoad(){
  document.form.From_Donor_Id.focus();
}	
function Query_OnClick(){
  <%call CheckStringJ("From_Donor_Id","欲轉出的捐款人編號")%>
  <%call CheckNumberJ("From_Donor_Id","欲轉出的捐款人編號")%>
  <%call CheckStringJ("To_Donor_Id","欲轉入的捐款人編號")%>
  <%call CheckNumberJ("To_Donor_Id","欲轉入的捐款人編號")%>
  document.form.action.value='query';
  document.form.submit();
}
function Transfer_OnClick(){
  if(confirm('您是否確定要合併捐款資料？')){
    document.form.action.value='transfer';
    document.form.submit()
  }
}
function DonateId_OnClick(){
  if(document.form.donate_id[0].checked){
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.donate_id[i].checked=true;
    }
  }else{
    for(var i=1;i<=Number(document.form.Total_Row.value);i++){
      document.form.donate_id[i].checked=false;
    }
  }
}
--></script>