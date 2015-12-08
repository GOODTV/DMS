<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="save" Then
  Check_Close=Get_Close("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(Session("user_id")))
  If Check_Close Then
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
        End If
        RS1.Close
        Set RS1=Nothing
      End If
    Next
  
    If Ckeck_GoodID Then
      '取捐物編號
      Invoice_Pre=""
      Invoice_No=""
      If request("Issue_Type")="M" And request("Invoice_No")<>"" Then
        Invoice_No=Trim(request("Invoice_No"))
      Else
        Act_id=""
        If request("Act_Id")<>"" Then Act_id=Cstr(request("Act_Id"))
        InvoiceNo=Get_Invoice_No2("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(request("Invoice_Type")),Act_id)
        If InvoiceNo<>"" Then
          Invoice_Pre=Split(InvoiceNo,"/")(0)
          Invoice_No=Split(InvoiceNo,"/")(1)
        End If
      End If
    
      '新增捐物資料
      SQL1="CONTRIBUTE"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      RS1.Addnew
      RS1("Donor_id")=request("Donor_Id")
      RS1("Contribute_Date")=request("Contribute_Date")
      RS1("Contribute_Payment")=request("Contribute_Payment")
      RS1("Contribute_Purpose")=request("Contribute_Purpose")
      If request("Contribute_Amt")<>"" Then
        RS1("Contribute_Amt")=request("Contribute_Amt")
      Else
        RS1("Contribute_Amt")="0"
      End If
      RS1("Dept_Id")=request("Dept_Id")
      RS1("Invoice_Title")=Data_Plus(request("Invoice_Title"))
      RS1("Invoice_Pre")=Invoice_Pre
      RS1("Invoice_No")=Invoice_No
      RS1("Invoice_Print")="0"
      RS1("Invoice_Print_Add")="0"
      RS1("Invoice_Print_Yearly")="0"
      RS1("Invoice_Print_Yearly_Add")="0"
      If request("Accoun_Date")<>"" Then
        RS1("Accoun_Date")=request("Accoun_Date")
      Else
        RS1("Accoun_Date")=null
      End If
      RS1("Invoice_Type")=request("Invoice_Type")
      RS1("Accounting_Title")=request("Accounting_Title")        
      If request("Act_id")<>"" Then
       RS1("Act_id")=request("Act_id")
      Else
        RS1("Act_id")=null
      End If
      RS1("Comment")=request("Comment")
      RS1("Invoice_PrintComment")=request("Invoice_PrintComment")
      If request("Issue_Type")<>"" Then
        RS1("Issue_Type")=request("Issue_Type")
        RS1("Issue_Type_Keep")=request("Issue_Type_Keep")
      Else
        RS1("Issue_Type")=""
        RS1("Issue_Type_Keep")=""
      End If
      RS1("Export")="N"
      RS1("Create_Date")=Date()
      RS1("Create_User")=session("user_name")
      RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
      RS1.Update
      RS1.Close
      Set RS1=Nothing
    
      SQL1="Select @@IDENTITY As Contribute_Id"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,1
      Contribute_id=RS1("Contribute_Id")
      RS1.Close
      Set RS1=Nothing
       
      '新增捐物明細
      For I=1 To Cint(request("Goods_Max"))
        If request("Goods_Name_"&I)<>"" Then
          SQL1="CONTRIBUTEDATA"
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,3
          RS1.Addnew
          RS1("Contribute_Id")=Contribute_Id
          RS1("Donor_id")=request("Donor_Id")
          RS1("Goods_Id")=request("Goods_Id_"&I)
          RS1("Goods_Name")=request("Goods_Name_"&I)
          If request("Goods_Qty_"&I)<>"" Then
            RS1("Goods_Qty")=request("Goods_Qty_"&I)
          Else
            RS1("Goods_Qty")="0"
          End If
          RS1("Goods_Unit")=request("Goods_Unit_"&I)
          If request("Goods_Amt_"&I)<>"" Then
            RS1("Goods_Amt")=request("Goods_Amt_"&I)
          Else
            RS1("Goods_Amt")="0"
          End If
          If request("Goods_DueDate_"&I)<>"" Then
            RS1("Goods_DueDate")=request("Goods_DueDate_"&I)
          Else
            RS1("Goods_DueDate")=null
          End If
          RS1("Goods_Comment")=request("Goods_Comment_"&I)
          If request("Contribute_IsStock_"&I)<>"" Then
            RS1("Contribute_IsStock")="Y"
          Else
            RS1("Contribute_IsStock")="N"
          End If
          RS1("Export")="N"
          RS1("Create_Date")=Date()
          RS1("Create_User")=session("user_name")
          RS1("Create_IP")=Request.ServerVariables("REMOTE_HOST")
          RS1.Update
          RS1.Close
          Set RS1=Nothing
        
          '寫入庫存量
          If request("Contribute_IsStock_"&I)<>"" Then
            SQL1="Select * From GOODS Where Goods_Id='"&request("Goods_Id_"&I)&"' "
            Set RS1 = Server.CreateObject("ADODB.RecordSet")
            RS1.Open SQL1,Conn,1,3
            If Not RS1.EOF And request("Goods_Qty_"&I)<>"" Then
              If RS1("Goods_IsStock")="Y" Then
                RS1("Goods_Qty")=Cdbl(RS1("Goods_Qty"))+Cdbl(request("Goods_Qty_"&I))
           	    RS1.Update
           	  End If
            End If
            RS1.Update
            RS1.Close
            Set RS1=Nothing
          End If
        End If
      Next

      '確認收據編號無重覆
      If request("Issue_Type")="" And Invoice_No<>"" Then
        Invoice_Pre_Old=Invoice_Pre
        Invoice_No_Old=Invoice_No
        Check_InvoiceNo=False
        While Check_InvoiceNo=False
          Check_Contribute=False
          SQL1="Select * From CONTRIBUTE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"' And Contribute_Id<>'"&contribute_id&"'"
          Call QuerySQL(SQL1,RS1)
          If RS1.EOF Then Check_Contribute=True
          RS1.Close
          Set RS1=Nothing
        
          Check_Donate=False
          If Check_Contribute Then
            SQL1="Select * From DONATE Where Invoice_Pre='"&Invoice_Pre_Old&"' And Invoice_No='"&Invoice_No_Old&"'"
            Call QuerySQL(SQL1,RS1)
            If RS1.EOF Then Check_Donate=True
            RS1.Close
            Set RS1=Nothing
          End If

          If Check_Contribute And Check_Donate Then
            Check_InvoiceNo=True
          Else
            Act_id=""
            If request("Act_Id")<>"" Then Act_id=Cstr(request("Act_Id"))
            InvoiceNo=Get_Invoice_No2("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(request("Invoice_Type")),Act_id)
            If InvoiceNo<>"" Then
              Invoice_Pre_Old=Split(InvoiceNo,"/")(0)
              Invoice_No_Old=Split(InvoiceNo,"/")(1)
              SQL="Update CONTRIBUTE Set Invoice_Pre='"&Invoice_Pre_Old&"',Invoice_No='"&Invoice_No_Old&"' Where Contribute_Id='"&contribute_id&"' "
              Set RS=Conn.Execute(SQL)
            End If
          End If
        Wend
      End If

      '修改捐贈人捐物紀錄
      call Declare_DonorId (request("donor_id"))
      
      session("errnumber")=1
      session("msg")="捐物資料新增成功 ！"
      If request("ctype")="contribute_data" Then
        Response.Redirect "contribute_data.asp?donor_id="&request("donor_id")
      ElseIf request("ctype")="member_contribute_data" Then
        Response.Redirect "member_contribute_data.asp?donor_id="&request("donor_id")  
      Else
        Response.Redirect "contribute_detail.asp?contribute_id="&contribute_id&"&ctype="&request("ctype")
      End If
    End If
  Else
    session("errnumber")=1
    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Contribute_Date"))&"』 已關帳無法新增 ！"
  End If  
End If

InvoiceTypeY="年度匯整"
SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
RS1.Close
Set RS1=Nothing

SQL1="Select * From DEPT Where Dept_Id='"&session("dept_id")&"'"
Call QuerySQL(SQL1,RS1)
Goods_Max=RS1("Goods_Max")
IsStock=RS1("IsStock")
RS1.Close
Set RS1=Nothing

SQL="Select *,Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+Address Else A.mValue+DONOR.ZipCode+Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then DONOR.Invoice_Address Else Case When C.mValue<>B.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+DONOR.Invoice_Address Else C.mValue+DONOR.Invoice_ZipCode+DONOR.Invoice_Address End End) From DONOR " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _ 
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Donor_Id='"&request("donor_id")&"'"
Call QuerySQL(SQL,RS)
%>
<%Prog_Id="contribute"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="post" action="">
      <input type="hidden" name="action">
      <input type="hidden" name="ctype" value="<%=request("ctype")%>">
      <input type="hidden" name="Goods_Max" value="<%=Goods_Max%>">
      <input type="hidden" name="IsStock" value="<%=IsStock%>">	
      <input type="hidden" name="Contribute_Amt" value="0">
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">
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
                        <table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">捐贈人：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Donor_Name" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Donor_Name"))%>&nbsp;<%=RS("Title")%>">
                              &nbsp;類&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：
                              <input type="text" name="Category" size="10" class="font9t" readonly value="<%=RS("Category")%>">
                              &nbsp;身份別：
                              <input type="text" name="Donor_Type" size="40" class="font9t" readonly value="<%=RS("Donor_Type")%>">
                            </td>
                          </tr>
                          <tr>
                            <td align="right">收據地址：</td>
                            <td align="left" colspan="7">
                              <input type="text" name="Invoice_Address2" size="30" class="font9t" readonly value="<%=Data_Minus(RS("Invoice_Address2"))%>">
                              &nbsp;收據開立：
                              <input type="text" name="InvoiceType" size="10" class="font9t" readonly value="<%=RS("Invoice_Type")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                              &nbsp;最近捐贈日：
                              <input type="text" name="Last_DonateDateC" size="10" class="font9t" readonly value="<%=RS("Last_DonateDateC")%>">	
                            </td>
                          </tr>
                          <tr>
                            <td align="right">捐贈人備註：</td>
                            <td align="left" colspan="7"><textarea rows="2" name="Remark" cols="51" readonly class="font9t"><%=RS("Remark")%></textarea></td> </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="12%" align="right">捐贈日期：</td>
                            <td width="16% align="left"><%call Calendar("Contribute_Date",request("Contribute_Date"))%></td>
                            <td width="10%" align="right">捐贈方式：</td>
                            <td width="16%">
                            <%
                              SQL="Select Contribute_Payment=CodeDesc From CASECODE Where CodeType='Payment2' Order By Seq"
                              FName="Contribute_Payment"
                              Listfield="Contribute_Payment"
                              menusize="1"
                              If request("Contribute_Payment")<>"" Then
                                BoundColumn=request("Contribute_Payment")
                              Else
                                BoundColumn="物品捐贈"
                              End If
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td width="10%" align="right">捐贈用途：</td>
                            <td width="14%">
                            <%
                              SQL="Select Contribute_Purpose=CodeDesc From CASECODE Where CodeType='Purpose2' Order By Seq"
                              FName="Contribute_Purpose"
                              Listfield="Contribute_Purpose"
                              menusize="1"
                              If request("Contribute_Purpose")<>"" Then
                                BoundColumn=request("Contribute_Purpose")
                              Else
                                BoundColumn="一般捐物"
                              End If
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td width="10%" align="right">收據開立：</td>
                            <td width="12%">
                            <%
                              SQL="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Not In ('"&InvoiceTypeY&"') Order By Seq"
                              FName="Invoice_Type"
                              Listfield="Invoice_Type"
                              menusize="1"
                              If request("Invoice_Type")<>"" Then
                                BoundColumn=request("Invoice_Type")
                              Else
                                If Cstr(InvoiceTypeY)<>Cstr(RS("Invoice_Type")) Then
                                  BoundColumn=RS("Invoice_Type")
                                Else
                                  BoundColumn="單次收據"
                                End If  
                              End If
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td>
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
                            <td align="right">收據抬頭：</td>
                            <td colspan="2"><input type="text" name="Invoice_Title" size="30" class="font9" maxlength="80" value="<%If request("Invoice_Title")<>"" Then Response.Write Data_Minus(request("Invoice_Title")) Else Response.Write Data_Minus(RS("Invoice_Title")) End If%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" OnClick="Issue_Type_OnClick()">手開收據</td>
                            <td align="right">收據編號：</td>
                            <td><input type="text" name="Invoice_No" size="12" class="font9t" maxlength="20" readonly value="自動編號"></td>
                          </tr> 
                          <tr>
                            <td align="right">沖帳日期：</td>
                            <td align="left"><%call Calendar("Accoun_Date",request("Accoun_Date"))%></td>
                            <td align="right">會計科目：</td>
                            <td align="left">
                            <%
                              SQL="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun2' Order By Seq"
                              FName="Accounting_Title"
                              Listfield="Accounting_Title"
                              menusize="1"
                              BoundColumn=request("Accounting_Title")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td align="right">募款活動：</td>
                            <td align="left" colspan="3">
                            <%
                              SQL="Select Act_Id,Act_ShortName From ACT Order By Act_id Desc"
                              FName="Act_Id"
                              Listfield="Act_ShortName"
                              menusize="1"
                              BoundColumn=request("Act_Id")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">捐贈備註：</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Comment" cols="45" class="font9"><%=request("Comment")%></textarea></td>
                            <td align="right">收據備註：<br />(列印收據用)</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Invoice_PrintComment" cols="40" class="font9"><%=request("Invoice_PrintComment")%></textarea></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="100%" colspan="8"> 
                              <table border="0" width="100%" cellpadding="2" cellspacing="2">
                                <tr>
                                	<td width="5%" align="center">寫入庫存</td>
                                	<td width="20%">物品名稱</td>
                                	<td width="11%">物品數量</td>
                                	<td width="11%">物品單位</td>
                                	<td width="11%">折合現金</td>
                                	<td width="13%">物品保存期限</td>
                                	<td width="24%">備註</td>
                                	<td width="5%" align="center">清除</td>
                                </tr>
                                <%For I=1 To Goods_Max%>
                                <tr>
                                	<input type="hidden" name="Goods_Id_<%=I%>" value="<%=request("Goods_Id_"&I)%>">
                                	<td align="center"><input type="checkbox" name="Contribute_IsStock_<%=I%>" value="Y" checked OnClick="Contribute_IsStock_OnClick(<%=I%>)"></td>	
                                	<!--20131004 Modify by GoodTV Tanya:修改開啟視窗位置-->
                                	<td><input type="text" name="Goods_Name_<%=I%>" size="18" maxlength="50" class="font9" value="<%=request("Goods_Name_"&I)%>">&nbsp;<a href onclick="window.open('goods_show.asp?LinkId=Goods_Id_<%=I%>&LinkName=Goods_Name_<%=I%>&LinkUnit=Goods_Unit_<%=I%>','','status=no,scrollbars=yes,top=100,left=600,width=500,height=450')" style="cursor:hand"><img border="0" src="../images/toolbar_search.gif" width="17"></a></td>
                                	<td><input type="text" name="Goods_Qty_<%=I%>" size="10" maxlength="7" class="font9" value="<%=request("Goods_Qty_"&I)%>"></td>
                                	<td><input type="text" name="Goods_Unit_<%=I%>" size="10" maxlength="10" class="font9" value="<%=request("Goods_Unit_"&I)%>"></td>
                                	<td><input type="text" name="Goods_Amt_<%=I%>" size="10" maxlength="7" class="font9" value="<%If request("Goods_Amt_"&I)<>"" Then Response.Write request("Goods_Amt_"&I) Else Response.Write "0" End If%>"></td>
                                	<td><%call Calendar("Goods_DueDate_"&I&"",request("Goods_DueDate_"&I))%></td>
                                	<td><input type="text" name="Goods_Comment_<%=I%>" size="26" maxlength="100" class="font9" value="<%=request("Goods_Comment_"&I)%>"></td>
                                  <td align="center"><img border="0" src="../images/toobar_cancle.gif" width="20" onClick="Contribute_Cancel_OnClick(<%=I%>)" style="cursor:hand"></td>
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
  document.form.Contribute_Date.focus();
}
function Issue_Type_OnClick(){
  if(document.form.Issue_Type.checked){
    document.form.Invoice_No.style.backgroundColor='#ffffff';
    document.form.Invoice_No.readOnly=false;
    document.form.Invoice_No.value='';
    document.form.Invoice_No.focus();
  }else{
    document.form.Invoice_No.style.backgroundColor='#ffffcc';
    document.form.Invoice_No.readOnly=true;
    document.form.Invoice_No.value='自動編號';
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
  document.getElementById('Goods_Amt_'+i).value='0';
  document.getElementById('Goods_DueDate_'+i).value='';
  document.getElementById('Goods_Comment_'+i).value='';
}
function Contribute_Cancel_OnClick(i){
  if(document.getElementById('Goods_Name_'+i).value!=''){
    if(confirm('您是否確定要清除『 '+document.getElementById('Goods_Name_'+i).value+' 』？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';
      document.getElementById('Goods_Unit_'+i).value='';
      document.getElementById('Goods_Amt_'+i).value='0';
      document.getElementById('Goods_DueDate_'+i).value='';
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }else{
    if(confirm('您是否確定要清除？')){
      document.getElementById('Goods_Id_'+i).value='';
      document.getElementById('Goods_Name_'+i).value='';
      document.getElementById('Goods_Qty_'+i).value='';
      document.getElementById('Goods_Unit_'+i).value='';
      document.getElementById('Goods_Amt_'+i).value='0';
      document.getElementById('Goods_DueDate_'+i).value='';
      document.getElementById('Goods_Comment_'+i).value='';
    }
  }
}
function Save_OnClick(){
  <%call CheckStringJ("Contribute_Date","捐贈日期")%>
  <%call CheckDateJ("Contribute_Date","捐贈日期")%>
  <%call CheckStringJ("Contribute_Payment","捐贈方式")%>
  <%call ChecklenJ("Contribute_Payment",20,"捐贈方式")%>
  <%call CheckStringJ("Contribute_Purpose","捐贈用途")%>
  <%call ChecklenJ("Contribute_Purpose",20,"捐贈用途")%>
  <%call CheckStringJ("Invoice_Type","收據開立")%>
  <%call ChecklenJ("Invoice_Type",20,"收據開立")%>
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call ChecklenJ("Invoice_Title",80,"收據抬頭")%>
  if(document.form.Issue_Type.checked){
    <%call CheckStringJ("Invoice_No","收據編號")%>
    <%call ChecklenJ("Invoice_No",20,"收據編號")%>
  }else{
    document.form.Invoice_No.value='';
  }
  if(document.form.Accoun_Date.value!=''){
    <%call CheckDateJ("Accoun_Date","沖帳日期")%>
  }
  <%call ChecklenJ("Accounting_Title",30,"會計科目")%>
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
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品數量  欄位不可為空白！');
        document.getElementById('Goods_Qty_'+i).focus();
        return;
      }
      if(isNaN(Number(document.getElementById('Goods_Qty_'+i).value))==true){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品數量  欄位必須為數字！');
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
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品數量  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_Unit_'+i).value==''){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品單位  欄位不可為空白！');
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
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』物品單位  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_Amt_'+i).value==''){
        document.getElementById('Goods_Amt_'+i).value='0';
      }
      if(isNaN(Number(document.getElementById('Goods_Amt_'+i).value))==true){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』折合現金  欄位必須為數字！');
        document.getElementById('Goods_Amt_'+i).focus();
        return;
      }
      var cnt=0;
      var sName=document.getElementById('Goods_Amt_'+i).value
      for(var j=0;j<sName.length;j++ ){
        if(escape(sName.charAt(j)).length>=4) cnt+=2;
        else cnt++;
      }
      if(cnt>7){
        alert('『'+document.getElementById('Goods_Name_'+i).value+'』折合現金  欄位長度超過限制！');
        return;
      }
      if(document.getElementById('Goods_DueDate_'+i).value!=''){
        if(document.getElementById('Goods_DueDate_'+i).value.indexOf("/")==-1||document.getElementById('Goods_DueDate_'+i).value.indexOf("/")==1||document.getElementById('Goods_DueDate_'+i).value.indexOf("/")==document.getElementById('Goods_DueDate_'+i).length){
          alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
          document.getElementById('Goods_DueDate_'+i).focus();
          return;
        }
        Ary_Date=document.getElementById('Goods_DueDate_'+i).value.split("/");
        if(Ary_Date.length!=3) return false;
        for(j=0;j<3;j++){
          if(isNaN(Number(Ary_Date[j]))==true){
            alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
            document.getElementById('Goods_DueDate_'+i).focus();
            return false;
          }
        }
        if(Ary_Date[0].length!=4){
          alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
          document.getElementById('Goods_DueDate_'+i).focus();
          return false;
        }
        if(parseInt(Number(Ary_Date[0]))<1000){
          alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
          document.getElementById('Goods_DueDate_'+i).focus();
          return false;
        }
        if(parseInt(Number(Ary_Date[1]))<1||parseInt(Number(Ary_Date[1]))>12){
          alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
          document.getElementById('Goods_DueDate_'+i).focus();
          return false;
        }
        var YYYY=parseInt(Number(Ary_Date[0]))
        var MM=parseInt(Number(Ary_Date[1]))
        var DD=parseInt(Number(Ary_Date[2]))
        if(MM==1||MM==3||MM==5||MM==7||MM==8||MM==10||MM==12){
          if(DD<1||DD>31) return false;
        }else if(MM==4||MM==6||MM==9||MM==11){
          if(DD<1||DD>30){
            alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
            document.getElementById('Goods_DueDate_'+i).focus();
            return false;
          }
        }else if(MM==2){
          if(leapyear(YYYY)){
            if(DD<1||DD>29){
              alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
              document.getElementById('Goods_DueDate_'+i).focus();
              return false;
            }
          }else{
            if(DD<1||DD>28){
              alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
              document.getElementById('Goods_DueDate_'+i).focus();
              return false;
            }
          }
        }else{
          alert('物品保存期限  欄位格式錯誤！(西元年/月/日)');
          document.getElementById('Goods_DueDate_'+i).focus();
          return false;
        }
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
      var ContributeAmt=Number(document.form.Contribute_Amt.value)+Number(document.getElementById('Goods_Amt_'+i).value);
      document.form.Contribute_Amt.value=ContributeAmt;
    }
  }
  if(ck_name==false){
    alert('物品名稱  欄位不可為空白！');
    return;
  }
  <%call SubmitJ("save")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='contribute_data'){
    location.href='contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';
  }else if(document.form.ctype.value=='member_contribute_data'){
    location.href='member_contribute_data.asp?donor_id='+document.form.Donor_Id.value+'';    
  }else{
    location.href='contribute.asp';
  }
}
--></script>