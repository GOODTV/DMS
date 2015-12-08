<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function IssueDataGrid (SQL,HLink,HLink2,LinkParam)
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
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></font></td>"
  Response.Write "<td align=""center"" bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
  Response.Write "</tr>"
  While Not RS1.EOF
	  Response.Write "<tr style='cursor:hand' bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
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
    Response.Write "<td align=""center""><a href='#' onclick=""window.open('"&HLink&RS1(LinkParam)&"','','status=no,scrollbars=no,top=100,left=120,width=400,height=170')""><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></a></td>"
    Response.Write "<td align=""center""><a href='JavaScript:if(confirm(""是否確定要刪除『"&RS1(1)&"』 ?"")){window.location.href="""&HLink2&RS1(LinkParam)&""";}' target='main'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>" 
    RS1.MoveNext
    Response.Write "</tr>"
  Wend
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="update" Then
  Ckeck_GoodID=True
  For I=1 To Cint(request("Goods_Max"))
    If request("Goods_Name_"&I)<>"" Then
      If request("Contribute_IsStock_"&I)="Y" Then
        SQL1="Select * From GOODS Where Goods_Id='"&request("Goods_Id_"&I)&"' "
        Call QuerySQL(SQL1,RS1)
        If RS1.EOF Then
          Ckeck_GoodID=False
          session("errnumber")=1
          session("msg")="您輸入的『 "&request("Goods_Name_"&I)&" 』 該物品代號不存在"
          Exit For
        Else
          If RS1("Goods_IsStock")="Y" And CSng(RS1("Goods_Qty"))<CSng(request("Goods_Qty_"&I)) Then
            Ckeck_GoodID=False
            session("errnumber")=1
            session("msg")="您輸入的『 "&request("Goods_Name_"&I)&" 』 該物品庫存量不足無法領用"
            Exit For
          End If
        End If
        RS1.Close
        Set RS1=Nothing
      End If
      
      SQL1="Select * From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&" And Goods_Id='"&request("Goods_Id_"&I)&"' And Goods_Name='"&request("Goods_Name_"&I)&"' And Goods_Unit='"&request("Goods_Unit"&I)&"' "
      Call QuerySQL(SQL1,RS1)
      If Not RS1.EOF Then
        Ckeck_GoodID=False
        session("errnumber")=1
        session("msg")="您輸入的『 "&request("Goods_Name_"&I)&" 』 物品名稱重覆出現"
        Exit For
      End If
      RS1.Close
      Set RS1=Nothing
    End If
  Next
  
  If Ckeck_GoodID Then
    '修改領用資料
    SQL1="Select * From CONTRIBUTE_ISSUE Where Issue_Id='"&request("Issue_Id")&"'"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    RS1("Dept_Id")=request("Dept_Id")
    RS1("Issue_Date")=request("Issue_Date")
    RS1("Issue_Processor")=request("Issue_Processor")
    RS1("Issue_Purpose")=request("Issue_Purpose")
    RS1("Issue_Org")=request("Issue_Org")
    RS1("Issue_Comment")=request("Issue_Comment")
    RS1("Issue_Type")=request("Issue_Type")
    If request("Issue_Type")<>"" Then
      RS1("Issue_Type")=request("Issue_Type")
      RS1("Issue_Type_Keep")=request("Issue_Type")
      If Trim(request("Invoice_No"))<>"" Then RS1("Invoice_No")=Trim(request("Invoice_No"))
    Else
      RS1("Issue_Type")=""
      RS1("Issue_Type_Keep")=""
    End If
    RS1.Update
    RS1.Close
    Set RS1=Nothing
      
    '新增領用明細
    For I=1 To Cint(request("Goods_Max"))
      If request("Goods_Name_"&I)<>"" Then
        SQL1="CONTRIBUTE_ISSUEDATA"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,3
        RS1.Addnew
        RS1("Issue_Id")=request("Issue_Id")
        RS1("Goods_Id")=request("Goods_Id_"&I)
        RS1("Goods_Name")=request("Goods_Name_"&I)
        If request("Goods_Qty_"&I)<>"" Then
          RS1("Goods_Qty")=request("Goods_Qty_"&I)
        Else
          RS1("Goods_Qty")="0"
        End If
        RS1("Goods_Unit")=request("Goods_Unit_"&I)
        RS1("Goods_Comment")=request("Goods_Comment_"&I)
        If request("Contribute_IsStock_"&I)<>"" Then
          RS1("Contribute_IsStock")="Y"
        Else
          RS1("Contribute_IsStock")="N"
        End If
        RS1("Create_Date")=Date()
        RS1("Create_User")=session("user_name")
        RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
        RS1.Update
        RS1.Close
        Set RS1=Nothing
        
        '減少庫存量
        If request("Contribute_IsStock_"&I)<>"" Then
          SQL1="Select * From GOODS Where Goods_Id='"&request("Goods_Id_"&I)&"' "
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,3
          If Not RS1.EOF And request("Goods_Qty_"&I)<>"" Then
            If RS1("Goods_IsStock")="Y" Then
              RS1("Goods_Qty")=CSng(RS1("Goods_Qty"))-CSng(request("Goods_Qty_"&I))
           	  RS1.Update
           	End If
          End If
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        End If
      End If
    Next
    session("errnumber")=1
    session("msg")="領用資料修改成功 ！"
    Response.Redirect "contribute_issue_detail.asp?issue_id="&request("issue_id")
  End If
End If

If request("action")="delete" Then
  SQL1="Select * From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&"' And Contribute_IsStock='Y' And Goods_Id<>'' "
  Call QuerySQL(SQL1,RS1)
  If Not RS1.EOF Then
    While Not RS1.EOF
      SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
      Set RS2 = Server.CreateObject("ADODB.RecordSet")
      RS2.Open SQL2,Conn,1,3
      If Not RS2.EOF Then
        RS2("Goods_Qty")=CSng(RS2("Goods_Qty"))+CSng(RS1("Goods_Qty"))
        RS2.Update
      End If
      RS2.Close
      Set RS2=Nothing
      RS1.MoveNext
    Wend
  End If
  RS1.Close
  Set RS1=Nothing
    
  SQL="Delete From CONTRIBUTE_ISSUE Where Issue_Id='"&request("issue_id")&"' " & _
      "Delete From CONTRIBUTE_ISSUEDATA Where Issue_Id='"&request("issue_id")&"' "
  Set RS=Conn.Execute(SQL)
  
  session("errnumber")=1
  session("msg")="領用資料刪除成功 ！"
  Response.Redirect "contribute_issue.asp"
End If

SQL="Select * From CONTRIBUTE_ISSUE Where Issue_Id='"&request("issue_id")&"' "
Call QuerySQL(SQL,RS)

SQL1="Select * From DEPT Where Dept_Id='"&RS("Dept_Id")&"'"
Call QuerySQL(SQL1,RS1)
Goods_Max=RS1("Goods_Max")
IsStock=RS1("IsStock")
RS1.Close
Set RS1=Nothing
%>
<%Prog_Id="issue"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="Goods_Max" value="<%=Goods_Max%>">
      <input type="hidden" name="IsStock" value="<%=IsStock%>">
      <input type="hidden" name="Issue_Id" value="<%=RS("issue_id")%>">	
      <input type="hidden" name="IssueNo" value="<%=RS("Issue_Pre")&RS("Issue_No")%>">
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【修改】</td>
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
                        <table border="0" cellpadding="3" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td width="10%" align="right">機構名稱：</td>
                            <td width="15% align="left">
                            <%
                              If Session("comp_label")="1" Then
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where Comp_Label>='"&Session("comp_label")&"' And Dept_Id In ("&Session("all_dept_type")&") Order By Comp_Label,Dept_Id"
                              Else
                                SQL="Select Dept_Id,Comp_ShortName From DEPT Where (Dept_Id In ("&Session("all_dept_type")&") Or Comp_Label>'"&Session("comp_label")&"') Order By Comp_Label,Dept_Id"
                              End If
                              FName="Dept_Id"
                              Listfield="Comp_ShortName"
                              menusize="1"
                              BoundColumn=RS("Dept_Id")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                  　        </td>
                            <td width="8%" align="right">領取人：</td>
                            <td width="17%"><input type="text" name="Issue_Processor" size="17" maxlength="20" class="font9" value="<%=RS("Issue_Processor")%>"></td>
                            <td width="10%" align="right">領用日期：</td>
                            <td width="15% align="left"><%call Calendar("Issue_Date",RS("Issue_Date"))%></td>
                            <td width="10%" align="right">領用用途：</td>
                            <td width="15%">
                            <%
                              SQL="Select Issue_Purpose=CodeDesc From CASECODE Where CodeType='Purpose3' Order By Seq"
                              FName="Issue_Purpose"
                              Listfield="Issue_Purpose"
                              menusize="1"
                              BoundColumn=RS("Issue_Purpose")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">出貨單位：</td>
                            <td><input type="text" name="Issue_Org" size="15" maxlength="40" class="font9" value="<%=RS("Issue_Org")%>"></td>
                            <td align="right">備註：</td>
                            <td colspan="2"><input type="text" name="Issue_Comment" size="30" maxlength="100" class="font9" value="<%=RS("Issue_Comment")%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" OnClick="Issue_Type_OnClick()" <%If RS("Issue_Type")="M" Then Response.Write "checked" End If%>>手開領用編號</td>
                            <td align="right">領用編號：</td>
                            <td><input type="text" name="Issue_No" size="15" class="font9t" maxlength="20" <%If RS("Issue_Type")="M" Then Response.Write "class=""font9""" Else Response.Write "class=""font9t"" readonly " End If%> maxlength="20" value="<%=RS("Issue_Pre")&RS("Issue_No")%>"></td>
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
                              HLink="contributeissuedata_edit.asp?issue_id="&RS("Issue_Id")&"&ser_no="
                              HLink2="contributeissuedata_delete.asp?issue_id="&RS("Issue_Id")&"&ser_no="
                              LinkParam="ser_no"
                              Call IssueDataGrid (SQL,HLink,HLink2,LinkParam)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="8"> 
                              <table border="0" width="100%" cellpadding="2" cellspacing="2">
                                <tr>
                                	<td width="6%" align="center">庫存品</td>
                                	<td width="20%">物品名稱</td>
                                	<td width="12%">數量</td>
                                	<td width="12%">單位</td>
                                	<td width="25%">備註</td>
                                	<td width="5%" align="center">清除</td>
                                	<td width="40%"></td>
                                </tr>
                                <%For I=1 To Goods_Max%>
                                <tr>
                                	<input type="hidden" name="Goods_Id_<%=I%>" value="<%=request("Goods_Id_"&I)%>">
                                	<td align="center"><input type="checkbox" name="Contribute_IsStock_<%=I%>" value="Y" checked OnClick="Contribute_IsStock_OnClick(<%=I%>)"></td>
                                	<td><input type="text" name="Goods_Name_<%=I%>" size="18" maxlength="50" class="font9t" value="<%=request("Goods_Name_"&I)%>">&nbsp;<a href onclick="window.open('goods_show.asp?LinkId=Goods_Id_<%=I%>&LinkName=Goods_Name_<%=I%>&LinkUnit=Goods_Unit_<%=I%>','','status=no,scrollbars=yes,top=100,left=120,width=500,height=450')" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a></td>
                                	<td><input type="text" name="Goods_Qty_<%=I%>" size="12" maxlength="7" class="font9" value="<%=request("Goods_Qty_"&I)%>"></td>
                                	<td><input type="text" name="Goods_Unit_<%=I%>" size="12" maxlength="10" class="font9t" value="<%=request("Goods_Unit_"&I)%>"></td>
                                	<td><input type="text" name="Goods_Comment_<%=I%>" size="26" maxlength="100" class="font9" value="<%=request("Goods_Comment_"&I)%>"></td>
                                  <td align="center"><img border="0" src="../images/toobar_cancle.gif" width="20" onClick="Contribute_Cancel_OnClick(<%=I%>)" style="cursor:hand"></td>
                                  <td> </td>
                                </tr>
                                <%Next%>
                              </table>	
                            </td>
                          </tr>
                          <!--#include file="../include/calendar2.asp"-->
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="center" colspan="8">
                               <input type="button" value=" 修 改 " name="save" class="cbutton" style="cursor:hand" onClick="Update_OnClick()">
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
  for(i=1;i<=Number(document.form.Goods_Max.value);i++){
    if(document.form.IsStock.value=='Y'){
      document.getElementById('Contribute_IsStock_'+i).checked=true;
      document.getElementById('Goods_Name_'+i).style.backgroundColor='#ffffcc';
      document.getElementById('Goods_Name_'+i).readOnly=true;
      document.getElementById('Goods_Unit_'+i).style.backgroundColor='#ffffcc';
      document.getElementById('Goods_Unit_'+i).readOnly=true;    
    }else{
      document.getElementById('Contribute_IsStock_'+i).checked=false;
      document.getElementById('Goods_Name_'+i).style.backgroundColor='#ffffff';
      document.getElementById('Goods_Name_'+i).readOnly=false;
      document.getElementById('Goods_Unit_'+i).style.backgroundColor='#ffffff';
      document.getElementById('Goods_Unit_'+i).readOnly=false;
    }
  }
  document.form.Issue_Processor.focus();
}
function Issue_Type_OnClick(){
  if(document.form.Issue_Type.checked){
    document.form.Issue_No.style.backgroundColor='#ffffff';
    document.form.Issue_No.readOnly=false;
    document.form.Issue_No.value='';
    document.form.Issue_No.focus();
  }else{
    document.form.Issue_No.style.backgroundColor='#ffffcc';
    document.form.Issue_No.readOnly=true;
    document.form.Issue_No.value='自動編號';
  }
}
function Contribute_IsStock_OnClick(i){
  if(document.getElementById('Contribute_IsStock_'+i).checked){
    document.getElementById('Goods_Name_'+i).style.backgroundColor='#ffffcc';
    document.getElementById('Goods_Name_'+i).readOnly=true;
    document.getElementById('Goods_Unit_'+i).style.backgroundColor='#ffffcc';
    document.getElementById('Goods_Unit_'+i).readOnly=true;
  }else{
    document.getElementById('Goods_Name_'+i).style.backgroundColor='#ffffff';
    document.getElementById('Goods_Name_'+i).readOnly=false;
    document.getElementById('Goods_Unit_'+i).style.backgroundColor='#ffffff';
    document.getElementById('Goods_Unit_'+i).readOnly=false;
  }
  document.getElementById('Goods_Id_'+i).value='';
  document.getElementById('Goods_Name_'+i).value='';
  document.getElementById('Goods_Qty_'+i).value='';
  document.getElementById('Goods_Unit_'+i).value='';
  document.getElementById('Goods_Comment_'+i).value='';
}
function Contribute_Cancel_OnClick(i){
  if(document.getElementById('Goods_Name_'+i).value!=''){
    if(confirm('您是否確定要清除『 '+document.getElementById('Goods_Name_'+i).value+' 』？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';
      document.getElementById('Goods_Unit_'+i).value='';
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }else{
    if(confirm('您是否確定要清除？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';
      document.getElementById('Goods_Unit_'+i).value='';
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }
}
function Update_OnClick(){
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call CheckStringJ("Issue_Processor","領取人")%>
  <%call ChecklenJ("Issue_Processor",20,"領取人")%>
  <%call CheckStringJ("Issue_Date","領用日期")%>
  <%call CheckDateJ("Issue_Date","領用日期")%>
  <%call CheckStringJ("Issue_Purpose","領用用途")%>
  <%call ChecklenJ("Issue_Purpose",20,"領用用途")%>
  <%call ChecklenJ("Issue_Org",40,"出貨單位")%>
  <%call ChecklenJ("Issue_Comment",100,"備註")%>
  if(document.form.Issue_Type.checked){
    <%call CheckStringJ("Issue_No","收據編號")%>
    <%call ChecklenJ("Issue_No",20,"收據編號")%>
  }else{
    document.form.Issue_No.value='';
  }
  for(i=1;i<=Number(document.form.Goods_Max.value);i++){
    if(document.getElementById('Goods_Name_'+i).value!=''){
      if(document.getElementById('Contribute_IsStock_'+i).checked&&document.getElementById('Goods_Id_'+i).value==''){
        alert('您輸入的『'+document.getElementById('Goods_Name_'+i).value+'』該商品代號不存在！');
        return;
      }
      var cnt=0;
      var sName=document.getElementById('Goods_Name_'+i).value
      for(var j=0;j<sName.length;j++ ){
        if(escape(sName.charAt(j)).length>=4) cnt+=2;
        else cnt++;
      }
      if(cnt>50){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品名稱  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_Qty_'+i).value==''){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』數量  欄位不可為空白！');
        document.getElementById('Goods_Qty_'+i).focus();
        return;
      }
      if(isNaN(Number(document.getElementById('Goods_Qty_'+i).value))==true){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』數量  欄位必須為數字！');
        document.getElementById('Goods_Qty_'+i).focus();
        return;
      }
      var cnt=0;
      var sName=document.getElementById('Goods_Qty_'+i).value
      for(var j=0;j<sName.length;j++ ){
        if(escape(sName.charAt(i)).length>=4) cnt+=2;
        else cnt++;
      }
      if(cnt>7){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』數量  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_Unit_'+i).value==''){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』單位  欄位不可為空白！');
        document.getElementById('Goods_Unit_'+i).focus();
        return;
      }
      var cnt=0;
      var sName=document.getElementById('Goods_Unit_'+i).value
      for(var j=0;j<sName.length;j++ ){
        if(escape(sName.charAt(i)).length>=4) cnt+=2;
        else cnt++;
      }
      if(cnt>10){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』單位  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_Comment_'+i).value!=''){
        var cnt=0;
        var sName=document.getElementById('Goods_Comment_'+i).value
        for(var j=0;j<sName.length;j++ ){
          if(escape(sName.charAt(j)).length>=4) cnt+=2;
          else cnt++;
        }
        if(cnt>100){
          alert('『'+document.getElementById('Goods_Name_'+i).value+'』備註  欄位長度超過限制！');
          return;
        }
      }
      for(j=i+1;j<Number(document.form.Goods_Max.value);j++){
        if(document.getElementById('Goods_Id_'+i).value==document.getElementById('Goods_Id_'+j).value&&document.getElementById('Goods_Name_'+i).value==document.getElementById('Goods_Name_'+j).value&&document.getElementById('Goods_Unit_'+i).value==document.getElementById('Goods_Unit_'+j).value){
          alert('『'+document.getElementById('Goods_Name_'+j).value+'』物品名稱重覆出現！');
          return;
        }
      }
    } 
  }
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  location.href='contribute_issue_detail.asp?issue_id='+document.form.Issue_Id.value+'';
}
--></script>