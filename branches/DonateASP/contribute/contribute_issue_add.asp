<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  Ckeck_GoodID=True
  For I=1 To Cint(request("Goods_Max"))
    If request("Goods_Name_"&I)<>"" And request("Contribute_IsStock_"&I)="Y" Then
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
  Next
       
  If Ckeck_GoodID Then
    '取領用編號
    Issue_Pre=""
    Issue_No=""
    If request("Issue_Type")="M" And request("Issue_No")<>"" Then
      Issue_No=Trim(request("Issue_No"))
    Else
      IssueNo=Get_Invoice_No2("3",Cstr(request("Dept_Id")),Cstr(request("Issue_Date")),"","")
      If IssueNo<>"" Then
        Issue_Pre=Split(IssueNo,"/")(0)
        Issue_No=Split(IssueNo,"/")(1)
      End If
    End If

    '新增領用資料
    SQL1="CONTRIBUTE_ISSUE"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,3
    RS1.Addnew
    RS1("Dept_Id")=request("Dept_Id")
    RS1("Issue_Date")=request("Issue_Date")
    RS1("Issue_Processor")=request("Issue_Processor")
    RS1("Issue_Purpose")=request("Issue_Purpose")
    RS1("Issue_Org")=request("Issue_Org")
    RS1("Issue_Comment")=request("Issue_Comment")
    If request("Issue_Type")<>"" Then 
      RS1("Issue_Type")=request("Issue_Type")
      RS1("Issue_Type_Keep")=request("Issue_Type_Keep")
    Else
      RS1("Issue_Type")=""
      RS1("Issue_Type_Keep")=""
    End If 
    RS1("Issue_Pre")=Issue_Pre
    RS1("Issue_No")=Issue_No
    RS1("Issue_Print")="0"
    RS1("Create_Date")=Date()
    RS1("Create_User")=session("user_name")
    RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
    RS1.Update
    RS1.Close
    Set RS1=Nothing
    
    SQL1="Select @@IDENTITY As Issue_Id"
    Set RS1 = Server.CreateObject("ADODB.RecordSet")
    RS1.Open SQL1,Conn,1,1
    Issue_Id=RS1("Issue_Id")
    RS1.Close
    Set RS1=Nothing
       
    '新增領用明細
    For I=1 To Cint(request("Goods_Max"))
      If request("Goods_Name_"&I)<>"" Then
        SQL1="CONTRIBUTE_ISSUEDATA"
        Set RS1 = Server.CreateObject("ADODB.RecordSet")
        RS1.Open SQL1,Conn,1,3
        RS1.Addnew
        RS1("Issue_Id")=Issue_Id
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

    '確認領用編號無重覆
    If request("Issue_Type")="" And Issue_No<>"" Then
      Issue_Pre_Old=Issue_Pre
      Issue_No_Old=Issue_No
      Check_InvoiceNo=False
      While Check_InvoiceNo=False
        Check_Issue=False
        SQL1="Select * From CONTRIBUTE_ISSUE Where Issue_Pre='"&Issue_Pre_Old&"' And Issue_No='"&Issue_No_Old&"' And Issue_Id<>'"&Issue_Id&"'"
        Call QuerySQL(SQL1,RS1)
        If RS1.EOF Then Check_Issue=True
        RS1.Close
        Set RS1=Nothing
        If Check_Issue Then
          Check_InvoiceNo=True
        Else
          IssueNo=Get_Invoice_No2("3",Cstr(request("Dept_Id")),Cstr(request("Issue_Date")),"","")
          If IssueNo<>"" Then
            Issue_Pre_Old=Split(IssueNo,"/")(0)
            Issue_No_Old=Split(IssueNo,"/")(1)
            SQL="Update CONTRIBUTE_ISSUE Set Issue_Pre='"&Issue_Pre_Old&"',Issue_No='"&Issue_No_Old&"' Where Issue_Id='"&Issue_Id&"' "
            Set RS=Conn.Execute(SQL)
          End If
        End If
      Wend
    End If
  
    session("errnumber")=1
    session("msg")="領用資料新增成功 ！"
    Response.Redirect "contribute_issue_detail.asp?issue_id="&Issue_Id
  End If
End If

SQL1="Select * From DEPT Where Dept_Id='"&session("dept_id")&"'"
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
                      <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;"><%=Prog_Desc%>【新增】</td>
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
                              If request("Dept_Id")<>"" Then
                                BoundColumn=request("Dept_Id")
                              Else
                                BoundColumn=Session("dept_id")
                              End If
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                  　        </td>
                            <td width="8%" align="right">領取人：</td>
                            <td width="17%"><input type="text" name="Issue_Processor" size="17" maxlength="20" class="font9" value="<%If request("Issue_Processor")<>"" Then Response.Write request("Issue_Processor") Else Response.Write Session("user_name") End If%>"></td>
                            <td width="10%" align="right">領用日期：</td>
                            <td width="15% align="left"><%call Calendar("Issue_Date",Date())%></td>
                            <td width="10%" align="right">領用用途：</td>
                            <td width="15%">
                            <%
                              SQL="Select Issue_Purpose=CodeDesc From CASECODE Where CodeType='Purpose3' Order By Seq"
                              FName="Issue_Purpose"
                              Listfield="Issue_Purpose"
                              menusize="1"
                              If request("Issue_Purpose")<>"" Then
                                BoundColumn=request("Issue_Purpose")
                              Else
                                BoundColumn="一般領用"
                              End If
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">出貨單位：</td>
                            <td><input type="text" name="Issue_Org" size="15" maxlength="40" class="font9" value="<%=request("Issue_Org")%>"></td>
                            <td align="right">備註：</td>
                            <td colspan="2"><input type="text" name="Issue_Comment" size="30" maxlength="100" class="font9" value="<%=request("Issue_Comment")%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" OnClick="Issue_Type_OnClick()">手開領用編號</td>
                            <td align="right">領用編號：</td>
                            <td><input type="text" name="Issue_No" size="15" class="font9t" maxlength="20" readonly value="自動編號"></td>
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
                               <input type="button" value=" 存 檔 " name="save" class="cbutton" style="cursor:hand" onClick="Save_OnClick()">
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
function Save_OnClick(){
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
  ck_name=new Boolean(false);
  for(i=1;i<=Number(document.form.Goods_Max.value);i++){
    if(document.getElementById('Goods_Name_'+i).value!=''){
      ck_name=true;
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
  if(ck_name==false){
    alert('物品名稱  欄位不可為空白！');
    return;
  }
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  location.href='contribute_issue.asp';
}
--></script>