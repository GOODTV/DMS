<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function ContributeDataGrid (SQL,HLink,HLink2,LinkParam)
  Goods_Amt=0
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
      ElseIf I=4 Then
        Goods_Amt=Goods_Amt+Cdbl(RS1(I))
        Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(I),0) & "</span></td>"
      ElseIf I=5 Or I=7 Then
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
  Response.Write "<tr><td colspan=""4"" align=""right""><b>折合現金合計</b></td><td align=""right""><b>"&FormatNumber(Goods_Amt,0)&"</b></td><td colspan=""5""></td></tr>"
  Response.Write "</table>"
  RS1.Close
  Set RS1=Nothing
End Function

If request("action")="update" Then
  Check_Close=Get_Close("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(Session("user_id")))
  If Check_Close Then
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
          End If
          RS1.Close
          Set RS1=Nothing
        End If
      
        SQL1="Select * From CONTRIBUTEDATA Where Contribute_Id='"&request("Contribute_Id")&"' And Goods_Id='"&request("Goods_Id_"&I)&"' And Goods_Name='"&request("Goods_Name_"&I)&"' And Goods_Unit='"&request("Goods_Unit"&I)&"' "
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
      '修改捐物資料
      SQL1="Select * From CONTRIBUTE Where Contribute_Id='"&request("Contribute_Id")&"'"
      Set RS1 = Server.CreateObject("ADODB.RecordSet")
      RS1.Open SQL1,Conn,1,3
      RS1("Donor_id")=request("Donor_Id")
      RS1("Contribute_Date")=request("Contribute_Date")
      RS1("Contribute_Payment")=request("Contribute_Payment")
      RS1("Contribute_Purpose")=request("Contribute_Purpose")
      If request("Contribute_Amt")<>"" Then RS1("Contribute_Amt")=Cdbl(RS1("Contribute_Amt"))+Cdbl(request("Contribute_Amt"))
      'RS1("Dept_Id")=request("Dept_Id")
      RS1("Invoice_Title")=Data_Plus(request("Invoice_Title"))
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
        If Trim(request("Invoice_No"))<>"" Then RS1("Invoice_No")=Trim(request("Invoice_No"))
      Else
        RS1("Issue_Type")=""
        RS1("Issue_Type_Keep")=""
      End If
      RS1("Export")="N"
      RS1.Update
      RS1.Close
      Set RS1=Nothing
      
      '新增捐物明細
      For I=1 To Cint(request("Goods_Max"))
        If request("Goods_Name_"&I)<>"" Then
          SQL1="CONTRIBUTEDATA"
          Set RS1 = Server.CreateObject("ADODB.RecordSet")
          RS1.Open SQL1,Conn,1,3
          RS1.Addnew
          RS1("Contribute_Id")=request("contribute_id")
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

      '修改捐款人捐款紀錄
      call Declare_DonorId (request("donor_id"))
    
      session("errnumber")=1
      session("msg")="捐贈資料修改成功 ！"
      If request("ctype")="contribute_data" Then
        Response.Redirect "contribute_data.asp?donor_id="&request("donor_id")
      ElseIf request("ctype")="member_contribute_data" Then
        Response.Redirect "member_contribute_data.asp?donor_id="&request("donor_id")          
      Else
        Response.Redirect "contribute_detail.asp?contribute_id="&request("contribute_id")
      End If 
    End If
  Else
    session("errnumber")=1
    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法修改 ！"
  End If  
End If

If request("action")="delete" Then
  Check_Close=Get_Close("2",Cstr(request("Dept_Id")),Cstr(request("Contribute_Date")),Cstr(Session("user_id")))
  If Check_Close Then
    Check_Delete=True
    Goods_Name=""
    SQL1="Select CONTRIBUTEDATA.*,Storage_Qty=Isnull(GOODS.Goods_Qty,0),Goods_IsStock=GOODS.Goods_IsStock From CONTRIBUTEDATA Left Join GOODS On CONTRIBUTEDATA.Goods_Id=GOODS.Goods_Id Where CONTRIBUTEDATA.Contribute_Id='"&request("Contribute_Id")&"'"
    Call QuerySQL(SQL1,RS1)
    If Not RS1.EOF Then
      While Not RS1.EOF
        If RS1("Contribute_IsStock")="Y" And RS1("Goods_IsStock")="Y" And Cdbl(RS1("Goods_Qty"))>Cdbl(RS1("Storage_Qty")) Then
          Check_Delete=False
          If Goods_Name="" Then
            Goods_Name=RS1("Goods_Name")
          Else
            Goods_Name=Goods_Name&"、"&RS1("Goods_Name")
          End If
        End If
        RS1.MoveNext
      Wend
      If Check_Delete=False Then
        session("errnumber")=1
        session("msg")="『"&Goods_Name&"』 \n\n庫存量不足無法刪除 ！"
      End If
    End If
    RS1.Close
    Set RS1=Nothing
  
    If Check_Delete Then
      SQL1="Select * From CONTRIBUTEDATA Where Contribute_Id='"&request("Contribute_Id")&"' And Contribute_IsStock='Y' And Goods_Id<>'' "
      Call QuerySQL(SQL1,RS1)
      If Not RS1.EOF Then
        While Not RS1.EOF
          SQL2="Select Goods_Qty From GOODS Where Goods_Id='"&RS1("Goods_Id")&"' And Goods_IsStock='Y'"
          Set RS2 = Server.CreateObject("ADODB.RecordSet")
          RS2.Open SQL2,Conn,1,3
          If Not RS2.EOF Then
            RS2("Goods_Qty")=Cdbl(RS2("Goods_Qty"))-Cdbl(RS1("Goods_Qty"))
            RS2.Update
          End If
          RS2.Close
          Set RS2=Nothing
          RS1.MoveNext
        Wend
      End If
      RS1.Close
      Set RS1=Nothing
    
      SQL="Delete From CONTRIBUTE Where Contribute_Id='"&request("Contribute_Id")&"' " & _
          "Delete From CONTRIBUTEDATA Where Contribute_Id='"&request("Contribute_Id")&"' "
      Set RS=Conn.Execute(SQL)
  
      '修改捐款人捐款紀錄
      call Declare_DonorId (request("donor_id"))

      session("errnumber")=1
      session("msg")="捐贈資料刪除成功 ！"
      If request("ctype")="contribute_data" Then
        Response.Redirect "contribute_data.asp?donor_id="&request("donor_id")
      ElseIf request("ctype")="member_contribute_data" Then
        Response.Redirect "member_contribute_data.asp?donor_id="&request("donor_id")  
      Else
        Response.Redirect "contribute.asp"
      End If
    End If
  Else
    session("errnumber")=1
    session("msg")="您輸入的捐贈日期『 "&Cstr(request("Donate_Date"))&"』 已關帳無法刪除 ！"
  End If  
End If

InvoiceTypeY="年度匯整"
SQL1="Select InvoiceTypeY=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Like '%年%' Order By Seq"
Call QuerySQL(SQL1,RS1)
If Not RS1.EOF Then InvoiceTypeY=RS1("InvoiceTypeY")
RS1.Close
Set RS1=Nothing

SQL="Select CONTRIBUTE.*,Comp_ShortName,Donor_Name=DONOR.Donor_Name,Title=DONOR.Title,Category=DONOR.Category,Donor_Type=DONOR.Donor_Type,InvoiceType=DONOR.Invoice_Type,IDNo=DONOR.IDNo, " & _
    "Address2=(Case When DONOR.City='' Then DONOR.Address Else Case When A.mValue<>B.mValue Then DONOR.ZipCode+A.mValue+B.mValue+DONOR.Address Else DONOR.ZipCode+A.mValue+DONOR.Address End End),Invoice_Address2=(Case When DONOR.Invoice_City='' Then Invoice_Address Else Case When C.mValue<>D.mValue Then DONOR.Invoice_ZipCode+C.mValue+D.mValue+Invoice_Address Else DONOR.Invoice_ZipCode+C.mValue+DONOR.Invoice_Address End End) " & _
    "From CONTRIBUTE " & _
    "Join DONOR On CONTRIBUTE.Donor_Id=DONOR.Donor_Id " & _
    "Join Dept On CONTRIBUTE.Dept_Id=DEPT.Dept_Id " & _
    "Left Join CodeCity As A On DONOR.City=A.mCode " & _
    "Left Join CodeCity As B On DONOR.Area=B.mCode " & _
    "Left Join CodeCity As C On DONOR.Invoice_City=C.mCode " & _
    "Left Join CodeCity As D On DONOR.Invoice_Area=D.mCode " & _
    "Where Contribute_Id='"&request("contribute_id")&"' "
Call QuerySQL(SQL,RS)

SQL1="Select * From DEPT Where Dept_Id='"&RS("Dept_Id")&"'"
Call QuerySQL(SQL1,RS1)
Goods_Max=RS1("Goods_Max")
IsStock=RS1("IsStock")
RS1.Close
Set RS1=Nothing
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
      <input type="hidden" name="Dept_Id" value="<%=RS("Dept_Id")%>">		
      <input type="hidden" name="Contribute_Id" value="<%=RS("Contribute_Id")%>">	
      <input type="hidden" name="Donor_Id" value="<%=RS("donor_id")%>">
      <input type="hidden" name="DonorName" value="<%=RS("Donor_Name")%>">
      <input type="hidden" name="DonorIDNo" value="<%=RS("IDNo")%>">
      <input type="hidden" name="InvoiceNo" value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>">
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
                        <table border="0" cellpadding="2" cellspacing="2" style="border-collapse: collapse" bordercolor="#111111" width="100%">
                          <tr>
                            <td align="right">捐款人：</td>
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
                              <input type="text" name="InvoiceType" size="10" class="font9t" readonly value="<%=RS("InvoiceType")%>">
                              &nbsp;身分證/統編：
                              <input type="text" name="IDNo" size="10" class="font9t" readonly value="<%=RS("IDNo")%>">
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td width="10%" align="right">捐贈日期：</td>
                            <td width="16% align="left"><%call Calendar("Contribute_Date",RS("Contribute_Date"))%></td>
                            <td width="10%" align="right">捐贈方式：</td>
                            <td width="16%">
                            <%
                              SQL1="Select Contribute_Payment=CodeDesc From CASECODE Where CodeType='Payment2' Order By Seq"
                              Set RS1 = Server.CreateObject("ADODB.RecordSet")
                              RS1.Open SQL1,Conn,1,1
                              Response.Write "<SELECT Name='Contribute_Payment' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                              Response.Write "<OPTION>" & " " & "</OPTION>"
                              While Not RS1.EOF
                                If Cstr(RS("Contribute_Payment"))=Cstr(RS1("Contribute_Payment")) Then
                                  Response.Write "<OPTION value='"&RS1("Contribute_Payment")&"' selected >"&RS1("Contribute_Payment")&"</OPTION>"
                                Else
                                  Response.Write "<OPTION value='"&RS1("Contribute_Payment")&"'>"&RS1("Contribute_Payment")&"</OPTION>"
                                End If
                                RS1.MoveNext
                              Wend
                              Response.Write "</SELECT>"
                              RS1.Close
                              Set RS1=Nothing
                            %>
                            </td>
                            <td width="10%" align="right">捐贈用途：</td>
                            <td width="14%">
                            <%
                              SQL="Select Contribute_Purpose=CodeDesc From CASECODE Where CodeType='Purpose2' Order By Seq"
                              FName="Contribute_Purpose"
                              Listfield="Contribute_Purpose"
                              menusize="1"
                              BoundColumn=RS("Contribute_Purpose")
                              call OptionList2 (SQL,FName,Listfield,BoundColumn,menusize,"N")
                            %>
                            </td>
                            <td width="10%" align="right">收據開立：</td>
                            <td width="14%">
                            <%
                              SQL="Select Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' And CodeDesc Not In ('"&InvoiceTypeY&"') Order By Seq"
                              FName="Invoice_Type"
                              Listfield="Invoice_Type"
                              menusize="1"
                              BoundColumn=RS("Invoice_Type")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">機構名稱：</td>
                            <td><input type="text" name="Comp_ShortName" size="12" class="font9t" readonly value="<%=RS("Comp_ShortName")%>"></td>
                            <td align="right">收據抬頭：</td>
                            <td colspan="2"><input type="text" name="Invoice_Title" size="30" class="font9" maxlength="80" value="<%=Data_Minus(RS("Invoice_Title"))%>"></td>
                            <td><input type="checkbox" name="Issue_Type" value="M" OnClick="Issue_Type_OnClick()" <%If RS("Issue_Type")="M" Then Response.Write "checked" End If%> >手開收據</td>
                            <td align="right">收據編號：</td>
                            <td><input type="text" name="Invoice_No" size="12" <%If RS("Issue_Type")="M" Then Response.Write "class=""font9""" Else Response.Write "class=""font9t"" readonly " End If%> maxlength="20" value="<%=RS("Invoice_Pre")&RS("Invoice_No")%>"></td>
                          </tr> 
                          <tr>
                            <td align="right">沖帳日期：</td>
                            <td align="left"><%call Calendar("Accoun_Date",RS("Accoun_Date"))%></td>
                            <td align="right">會計科目：</td>
                            <td align="left">
                            <%
                              SQL="Select Accounting_Title=CodeDesc From CASECODE Where CodeType='Accoun2' Order By Seq"
                              FName="Accounting_Title"
                              Listfield="Accounting_Title"
                              menusize="1"
                              BoundColumn=RS("Accounting_Title")
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
                              BoundColumn=RS("Act_Id")
                              call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                            %>
                            </td>
                          </tr>
                          <tr>
                            <td align="right">捐贈備註：</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Comment" cols="45" class="font9"><%=RS("Comment")%></textarea></td>
                            <td align="right">收據備註：<br />(列印收據用)</td>
                            <td align="left" colspan="3"><textarea rows="3" name="Invoice_PrintComment" cols="40" class="font9"><%=RS("Invoice_PrintComment")%></textarea></td>
                          </tr>
                          <tr>
                            <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                          </tr>
                          <tr>
                            <td align="right">捐贈內容：</td>
                            <td align="left" colspan="7">
                            <%
                              SQL="Select ser_no,物品名稱=Goods_Name,數量=Goods_Qty,單位=Goods_Unit,折合現金=Goods_Amt,保存期限=CONVERT(VarChar,Goods_DueDate,111),備註=Goods_Comment,寫入庫存=(Case When Contribute_IsStock='Y' Then 'V' Else '' End) " & _
                                  "From CONTRIBUTEDATA Where Contribute_Id='"&RS("contribute_id")&"' Order By Ser_No "
                              HLink="contributedata_edit.asp?ctype="&request("ctype")&"&contribute_id="&RS("contribute_id")&"&ser_no="
                              HLink2="contributedata_delete.asp?ctype="&request("ctype")&"&contribute_id="&RS("contribute_id")&"&ser_no="
                              LinkParam="ser_no"
                              Call ContributeDataGrid (SQL,HLink,HLink2,LinkParam)
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
                                	<td width="5%" align="center">寫入庫存</td>
                                	<td width="20%"><font color="#ff0000">續增捐贈物品</font></td>
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
    document.form.Invoice_No.value=document.form.InvoiceNo.value;
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
function Update_OnClick(){
  <%call CheckStringJ("Contribute_Date","捐贈日期")%>
  <%call CheckDateJ("Contribute_Date","捐贈日期")%>
  <%call CheckStringJ("Contribute_Payment","捐款方式")%>
  <%call ChecklenJ("Contribute_Payment",20,"捐款方式")%>
  <%call CheckStringJ("Contribute_Purpose","捐款用途")%>
  <%call ChecklenJ("Contribute_Purpose",20,"捐款用途")%>
  <%call CheckStringJ("Invoice_Type","收據開立")%>
  <%call ChecklenJ("Invoice_Type",20,"收據開立")%>
  <%call CheckStringJ("Dept_Id","機構名稱")%>
  <%call ChecklenJ("Invoice_Title",80,"收據抬頭")%>
  if(document.form.Issue_Type.checked){
    <%call CheckStringJ("Invoice_No","收據編號")%>
    <%call ChecklenJ("Invoice_No",20,"收據編號")%>
  }
  if(document.form.Accoun_Date.value!=''){
    <%call CheckDateJ("Accoun_Date","沖帳日期")%>
  }  
  <%call ChecklenJ("Accounting_Title",30,"會計科目")%>
  for(i=1;i<=Number(document.form.Goods_Max.value);i++){
    if(document.getElementById('Goods_Name_'+i).value!=''){
      if(document.getElementById('Contribute_IsStock_'+i).checked&&document.getElementById('Goods_Id_'+i).value==''){
        alert('您輸入的『'+document.getElementById('Goods_Name_'+i).value+'』該商品代號不存在！');
        return;
      }
      for(j=i+1;j<Number(document.form.Goods_Max.value);j++){
        if(document.getElementById('Goods_Id_'+i).value==document.getElementById('Goods_Id_'+j).value&&document.getElementById('Goods_Name_'+i).value==document.getElementById('Goods_Name_'+j).value){
          alert('『'+document.getElementById('Goods_Name_'+j).value+'』物品名稱重覆出現！');
          return;
        }
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
      var ContributeAmt=Number(document.form.Contribute_Amt.value)+Number(document.getElementById('Goods_Amt_'+i).value);
      document.form.Contribute_Amt.value=ContributeAmt;
    } 
  }
  <%call SubmitJ("update")%>
}
function Delete_OnClick(){
  <%call SubmitJ("delete")%>
}
function Cancel_OnClick(){
  if(document.form.ctype.value=='member_contribute_data'){
    location.href='member_contribute_detail.asp?ctype='+document.form.ctype.value+'&contribute_id='+document.form.Contribute_Id.value+'';
  }else{
    location.href='contribute_detail.asp?ctype='+document.form.ctype.value+'&contribute_id='+document.form.Contribute_Id.value+'';
  }
}
--></script>