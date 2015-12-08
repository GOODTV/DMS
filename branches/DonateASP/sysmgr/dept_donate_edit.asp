<!--#include file="../include/dbfunctionJ.asp"-->
<%
If request("action")="update" Then
  SQL="Select * From DEPT Where Dept_Id='"&request("dept_id")&"'"
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,3
  RS("Transfer_Date")=request("Transfer_Date")
  RS("Invoice_Pre")=request("Invoice_Pre")
  RS("Invoice_Rule_Type")=request("Invoice_Rule_Type")
  RS("Invoice_Rule_YMD")=request("Invoice_Rule_YMD")
  RS("Invoice_Rule_Len")=request("Invoice_Rule_Len")
  RS("Invoice_Rule_Pub")=request("Invoice_Rule_Pub")   
  RS("Invoice_Pre")=request("Invoice_Pre")
  RS("Invoice_Rule_Type")=request("Invoice_Rule_Type")
  RS("Invoice_Rule_YMD")=request("Invoice_Rule_YMD")
  RS("Invoice_Rule_Len")=request("Invoice_Rule_Len")
  RS("Invoice_Rule_Pub")=request("Invoice_Rule_Pub")
  RS("Invoice_Pre2")=request("Invoice_Pre2")
  RS("Invoice_Rule_Type2")=request("Invoice_Rule_Type2")
  RS("Invoice_Rule_YMD2")=request("Invoice_Rule_YMD2")
  RS("Invoice_Rule_Len2")=request("Invoice_Rule_Len2")
  RS("Invoice_Rule_Pub2")=request("Invoice_Rule_Pub2")
  RS("Invoice_Pre3")=request("Invoice_Pre3")
  RS("Invoice_Rule_Type3")=request("Invoice_Rule_Type3")
  RS("Invoice_Rule_YMD3")=request("Invoice_Rule_YMD3")
  RS("Invoice_Rule_Len3")=request("Invoice_Rule_Len3")
  RS("Invoice_Rule_Pub3")=request("Invoice_Rule_Pub3")
  RS("Invoice_Prog")=request("Invoice_Prog")
  RS("Invoice_Prog2")=request("Invoice_Prog2")
  RS("Invoice_Prog3")=request("Invoice_Prog3")
  RS("Label")=request("Label")
  RS("Comp_Label")=request("Comp_Label")
  'RS("Rept_Licence")=request("Rept_Licence")
  If request("Goods_Max")<>"" Then
    RS("Goods_Max")=request("Goods_Max")
  Else
    RS("Goods_Max")="5"
  End If
  If request("IsStock")<>"" Then
    RS("IsStock")=request("IsStock")
  Else
    RS("IsStock")="N"
  End If
  If request("IsAct")<>"" Then
    RS("IsAct")=request("IsAct")
  Else
    RS("IsAct")="N"
  End If  
  RS("Donate_PayName")=request("Donate_PayName")
  RS("Donate_Invoice")=request("Donate_Invoice")
  RS("Donate_AddPrint")=request("Donate_AddPrint")
  RS("Check_ID_Type")=request("Check_ID_Type")
  If request("Card_Code")<>"" Then
    RS("Card_Code")=request("Card_Code")
  Else
    RS("Card_Code")="3"
  End If
  If request("Card_Verify")<>"" Then
    RS("Card_Verify")=request("Card_Verify")
  Else
    RS("Card_Verify")="0"
  End If
  RS("MD5_Code")=request("MD5_Code")
  RS("MD5_rCode")=request("MD5_rCode")
  If request("IsCard")<>"" Then
    RS("IsCard")="Y"
  Else
    RS("IsCard")="N"
  End If
  If request("IsDoenLoad")<>"" Then
    RS("IsDoenLoad")="Y"
  Else
    RS("IsDoenLoad")="N"
  End If
  If request("Atm_Code")<>"" Then
    RS("Atm_Code")=request("Atm_Code")
  Else
    RS("Atm_Code")="1111"
  End If
  If request("Atm_Verify")<>"" Then
    RS("Atm_Verify")=request("Atm_Verify")
  Else
    RS("Atm_Verify")="0"
  End If 
  If request("IsAtm")<>"" Then
    RS("IsAtm")="Y"
  Else
    RS("IsAtm")="N"
  End If 
  If request("IsCsv")<>"" Then
    RS("IsCsv")="Y"
  Else
    RS("IsCsv")="N"
  End If 
  If request("IsIbon")<>"" Then
    RS("IsIbon")="Y"
  Else
    RS("IsIbon")="N"
  End If
  If request("IsFamiPort")<>"" Then
    RS("IsFamiPort")="Y"
  Else
    RS("IsFamiPort")="N"
  End If    
  If request("IsLifeEt")<>"" Then
    RS("IsLifeEt")="Y"
  Else
    RS("IsLifeEt")="N"
  End If     
  If request("od_dt")<>"" Then
    RS("od_dt")=request("od_dt")
  Else
    RS("od_dt")="7"
  End If
  If request("IsNews")<>"" Then
    RS("IsNews")="Y"
  Else
    RS("IsNews")="N"
  End If
  If request("IsEpaper")<>"" Then
    RS("IsEpaper")="Y"
  Else
    RS("IsEpaper")="N"
  End If
  If request("IsBooty")<>"" Then
    RS("IsBooty")="Y"
  Else
    RS("IsBooty")="N"
  End If
  RS.Update
  RS.Close
  Set RS=Nothing
  
  If request("IsAct")<>"" Then
    SQL="Update ACT Set Act_Flg='"&request("IsAct")&"',Act_Flg2='N' Where DEPT_ID='"&request("dept_id")&"'"
  Else
    SQL="Update ACT Set Act_Flg='N',Act_Flg2='N' Where DEPT_ID='"&request("dept_id")&"'"
  End If
  Set RS=Conn.Execute(SQL)
  
  session("errnumber")=1
  session("msg")="資料修改成功 !"
End If
		                                
SQL="Select * From DEPT Where DEPT_ID='"&request("dept_id")&"'"
call QuerySQL(SQL,RS) 
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>機構部門管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body>
  <p><div align="center"><center>
    <form name="form" action="" method="post">
    	<input type="hidden" name="Dept_Id" value="<%=request("dept_id")%>">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table width="100%" border="0" cellspacing="0" cellpadding="0" style="background-color:#EEEEE3">
          <tr>
            <td>
              <table width="780"  border="0" cellspacing="0" cellpadding="0"  align="center" style="background-color:#EEEEE3">
                <tr>
                  <td width="5%"></td>
                  <td width="95%">
  		              <table width="40%"  border="0" cellspacing="0" cellpadding="0">
		                  <tr>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_01.gif" width="10" height="10"></td>
                        <td class="table61-bg"><img src="../images/table06_02.gif" width="1" height="10"></td>
                        <td width="10" valign="top" style="background-color:#EEEEE3"><img src="../images/table06_03.gif" width="10" height="10"></td>
                      </tr>
                      <tr>
                        <td class="table62-bg"></td>
                        <td style="text-align:center;color:#660000;FONT-SIZE: 11pt;letter-spacing: 1;">機構參數設定</td>
                        <td class="table63-bg"></td>
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
                  <td height="445" valign="top" class="table62-bg">&nbsp;</td>
                  <td height="445" valign="top">              
                    <table width="100%" border="0" cellspacing="0" style="border-collapse: collapse">
                      <tr>
                        <td class="td02-c" width="100%" valign="top" colspan="3">
                          <table border="1" cellpadding="2" style="border-collapse: collapse" width="100%" height="25" cellspacing="1">
                            <tr>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">機構資訊設定</a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_system_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">網站參數設定</a></td>
                              <td class="button2-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_donate_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool"><font color="#800000"><b>捐款參數設定</b></font></a></td>
                              <td class="button-bg"><img border="0" src="../images/red_arrow.gIf" align="texttop"><a href="dept_content_edit.asp?dept_id=<%=RS("dept_id")%>" class="tool">捐款機制說明</a></td>
                            </tr>
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td valign="top">&nbsp;</td>
                      </tr>                  
                      <tr>
                        <td class="td02-c" width="100%">
                          <table border="0" cellpadding="2" cellspacing="1" style="border-collapse: collapse" bordercolor="#111111" width="100%">                        
                            <tr>
                              <td width="14%" align="right"><font color="#000080">機構簡稱<span lang="en-us">:</span></font></td>
                              <td width="86%">
                              	<input type="text" name="Comp_ShortName" size="21" class="font9t" readonly maxlength="30" value="<%=RS("Comp_ShortName")%>">
		                            &nbsp;&nbsp;&nbsp;
		                            <font color="#000080">機構主從<span lang="en-us">:</span></font>
		                            <select size="1" name="Comp_Label">
		                              <option value=""> </option>
		                              <option value="1" <%If RS("Comp_Label")="1" Then%>selected<%End If%> >第一層</option>
		                              <option value="2" <%If RS("Comp_Label")="2" Then%>selected<%End If%> >第二層</option>
		                              <option value="3" <%If RS("Comp_Label")="3" Then%>selected<%End If%> >第三層</option>
		                            </select>
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">最近轉帳日<span lang="en-us">:</span></font>
                                <input type="text" name="LastTransfer_Date" size="10" class="font9t" readonly value="<%=RS("LastTransfer_Date")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">固定轉帳日<span lang="en-us">:</span></font>
                                <input type="text" name="Transfer_Date" size="5" class="font9" value="<%=RS("Transfer_Date")%>">	
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">捐款前置碼<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="text" name="Invoice_Pre" size="5" class="font9" maxlength="5" value="<%=RS("Invoice_Pre")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">收據編碼規則<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Type">
                                  <option value=""> </option>
		                              <option value="R" <%If RS("Invoice_Rule_Type")="R" Then%>selected<%End If%> >民國年</option>
		                              <option value="A" <%If RS("Invoice_Rule_Type")="A" Then%>selected<%End If%> >西元年</option>
		                              <option value="S" <%If RS("Invoice_Rule_Type")="S" Then%>selected<%End If%> >流水號</option>
		                              <option value="O" <%If RS("Invoice_Rule_Type")="O" Then%>selected<%End If%> >客制化</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
                                <select size="1" name="Invoice_Rule_YMD">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_YMD")="Y" Then%>selected<%End If%> >年序號</option>
		                              <option value="M" <%If RS("Invoice_Rule_YMD")="M" Then%>selected<%End If%> >月序號</option>
		                              <option value="D" <%If RS("Invoice_Rule_YMD")="D" Then%>selected<%End If%> >日序號</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
		                            <font color="#000080">流水號長度<span lang="en-us">:</span></font>
		                            <input type="text" name="Invoice_Rule_Len" size="5" class="font9" maxlength="3" value="<%=RS("Invoice_Rule_Len")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">多單位流水號取號<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Pub">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_Pub")="Y" Then%>selected<%End If%> >共用</option>
		                              <option value="N" <%If RS("Invoice_Rule_Pub")="N" Then%>selected<%End If%> >獨立</option>
		                            </select> 
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">捐款收據報表<span lang="en-us">:</span></font></td>
                              <td>
                                <%
                                  SQL1="Select Seq,Invoice_Prog=CodeDesc From CASECODE Where CodeType='Invoice' Order By Seq"
                                  Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                  RS1.Open SQL1,Conn,1,1
                                  Response.Write "<SELECT Name='Invoice_Prog' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                                  Response.Write "<OPTION>" & " " & "</OPTION>"
                                  While Not RS1.EOF
                                    If Cstr(RS1("Seq"))=Cstr(RS("Invoice_Prog")) Then
                                      Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog")&"</OPTION>"
                                    Else
                                      Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog")&"</OPTION>"
                                    End If
                                    RS1.MoveNext
                                  Wend
                                  Response.Write "</SELECT>"
                                  RS1.Close
                                  Set RS1=Nothing
                                %>
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">年度收據報表<span lang="en-us">:</span></font>
                                <%
                                  SQL1="Select Seq,Invoice_Prog2=CodeDesc From CASECODE Where CodeType='InvoiceY' Order By Seq"
                                  Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                  RS1.Open SQL1,Conn,1,1
                                  Response.Write "<SELECT Name='Invoice_Prog2' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                                  Response.Write "<OPTION>" & " " & "</OPTION>"
                                  While Not RS1.EOF
                                    If Cstr(RS1("Seq"))=Cstr(RS("Invoice_Prog2")) Then
                                      Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog2")&"</OPTION>"
                                    Else
                                      Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog2")&"</OPTION>"
                                    End If
                                    RS1.MoveNext
                                  Wend
                                  Response.Write "</SELECT>"
                                  RS1.Close
                                  Set RS1=Nothing
                                %>
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">年度收據編號產生<span lang="en-us">:</span></font>
                                <select size="1" name="Donate_Invoice">
		                              <option value=""> </option>
		                              <option value="Y" <%If RS("Donate_Invoice")="Y" Then%>selected<%End If%> >逐筆</option>
		                              <option value="N" <%If RS("Donate_Invoice")="N" Then%>selected<%End If%> >年度</option>
		                            </select>
		                            &nbsp;&nbsp;&nbsp;
                                <font color="#000080">收據名條列印<span lang="en-us">:</span></font>
                                <select size="1" name="Donate_AddPrint">
		                              <option value=""> </option>
		                              <option value="Y" <%If RS("Donate_AddPrint")="Y" Then%>selected<%End If%> >開啟</option>
		                              <option value="N" <%If RS("Donate_AddPrint")="N" Then%>selected<%End If%> >關閉</option>
		                            </select>
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">預設名條貼紙<span lang="en-us">:</span></font></td>
                              <td>
                                <%
                                  SQL1="Select Seq,Label=CodeDesc From CASECODE Where CodeType='Label' Order By Seq"
                                  Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                  RS1.Open SQL1,Conn,1,1
                                  Response.Write "<SELECT Name='Label' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                                  Response.Write "<OPTION>" & " " & "</OPTION>"
                                  While Not RS1.EOF
                                    If Cstr(RS1("Seq"))=Cstr(RS("Label")) Then
                                      Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                                    Else
                                      Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Label")&"</OPTION>"
                                    End If
                                    RS1.MoveNext
                                  Wend
                                  Response.Write "</SELECT>"
                                  RS1.Close
                                  Set RS1=Nothing
                                %>
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">勸募活動自訂收據編號<span lang="en-us">:</span></font>	
                                <select size="1" name="IsAct">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("IsAct")="Y" Then%>selected<%End If%> >開啟</option>
		                              <option value="N" <%If RS("IsAct")="N" Then%>selected<%End If%> >關閉</option>
		                            </select>
                              </td>
                            </tr>
                            <tr>
                              <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">物品捐贈前置碼<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="text" name="Invoice_Pre2" size="5" class="font9" maxlength="5" value="<%=RS("Invoice_Pre2")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">收據編碼規則<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Type2">
                                  <option value=""> </option>
		                              <option value="R" <%If RS("Invoice_Rule_Type2")="R" Then%>selected<%End If%> >民國年</option>
		                              <option value="A" <%If RS("Invoice_Rule_Type2")="A" Then%>selected<%End If%> >西元年</option>
		                              <option value="S" <%If RS("Invoice_Rule_Type2")="S" Then%>selected<%End If%> >流水號</option>
		                              <option value="O" <%If RS("Invoice_Rule_Type2")="O" Then%>selected<%End If%> >客制化</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
                                <select size="1" name="Invoice_Rule_YMD2">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_YMD2")="Y" Then%>selected<%End If%> >年序號</option>
		                              <option value="M" <%If RS("Invoice_Rule_YMD2")="M" Then%>selected<%End If%> >月序號</option>
		                              <option value="D" <%If RS("Invoice_Rule_YMD2")="D" Then%>selected<%End If%> >日序號</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
		                            <font color="#000080">流水號長度<span lang="en-us">:</span></font>
		                            <input type="text" name="Invoice_Rule_Len2" size="5" class="font9" maxlength="3" value="<%=RS("Invoice_Rule_Len2")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">多單位流水號取號<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Pub2">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_Pub2")="Y" Then%>selected<%End If%> >共用</option>
		                              <option value="N" <%If RS("Invoice_Rule_Pub2")="N" Then%>selected<%End If%> >獨立</option>
		                            </select>
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">物品捐贈收據報表<span lang="en-us">:</span></font></td>
                              <td>
                                <%
                                  SQL1="Select Seq,Invoice_Prog3=CodeDesc From CASECODE Where CodeType='Invoice2' Order By Seq"
                                  Set RS1 = Server.CreateObject("ADODB.RecordSet")
                                  RS1.Open SQL1,Conn,1,1
                                  Response.Write "<SELECT Name='Invoice_Prog3' size='1' style='font-size: 9pt; font-family: 新細明體'>" 
                                  Response.Write "<OPTION>" & " " & "</OPTION>"
                                  While Not RS1.EOF
                                    If Cstr(RS1("Seq"))=Cstr(RS("Invoice_Prog3")) Then
                                      Response.Write "<OPTION selected value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog3")&"</OPTION>"
                                    Else
                                      Response.Write "<OPTION value='"&RS1("Seq")&"'>"&RS1("Invoice_Prog3")&"</OPTION>"
                                    End If
                                    RS1.MoveNext
                                  Wend
                                  Response.Write "</SELECT>"
                                  RS1.Close
                                  Set RS1=Nothing
                                %>
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">物品捐贈輸入上限<span lang="en-us">:</span></font>
                                <input type="text" name="Goods_Max" size="5" class="font9" maxlength="5" value="<%=RS("Goods_Max")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">預設庫存管理功能<span lang="en-us">:</span></font>
                                <select size="1" name="IsStock">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("IsStock")="Y" Then%>selected<%End If%> >開啟</option>
		                              <option value="N" <%If RS("IsStock")="N" Then%>selected<%End If%> >關閉</option>
		                            </select>
                              </td>
                            </tr>
                            <tr>
                              <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">物品領用前置碼<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="text" name="Invoice_Pre3" size="5" class="font9" maxlength="5" value="<%=RS("Invoice_Pre3")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">領用編碼規則<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Type3">
                                  <option value=""> </option>
		                              <option value="R" <%If RS("Invoice_Rule_Type3")="R" Then%>selected<%End If%> >民國年</option>
		                              <option value="A" <%If RS("Invoice_Rule_Type3")="A" Then%>selected<%End If%> >西元年</option>
		                              <option value="S" <%If RS("Invoice_Rule_Type3")="S" Then%>selected<%End If%> >流水號</option>
		                              <option value="O" <%If RS("Invoice_Rule_Type3")="O" Then%>selected<%End If%> >客制化</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
                                <select size="1" name="Invoice_Rule_YMD3">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_YMD3")="Y" Then%>selected<%End If%> >年序號</option>
		                              <option value="M" <%If RS("Invoice_Rule_YMD3")="M" Then%>selected<%End If%> >月序號</option>
		                              <option value="D" <%If RS("Invoice_Rule_YMD3")="D" Then%>selected<%End If%> >日序號</option>
		                            </select>
		                            <font color="#000080"><span lang="en-us">+</span></font>
		                            <font color="#000080">流水號長度<span lang="en-us">:</span></font>
		                            <input type="text" name="Invoice_Rule_Len3" size="5" class="font9" maxlength="3" value="<%=RS("Invoice_Rule_Len3")%>">
                                &nbsp;&nbsp;&nbsp;
                                <font color="#000080">多單位流水號取號<span lang="en-us">:</span></font>
                                <select size="1" name="Invoice_Rule_Pub3">
                                  <option value=""> </option>
		                              <option value="Y" <%If RS("Invoice_Rule_Pub3")="Y" Then%>selected<%End If%> >共用</option>
		                              <option value="N" <%If RS("Invoice_Rule_Pub3")="N" Then%>selected<%End If%> >獨立</option>
		                            </select> 
                              </td>
                            </tr>
                            <!--<tr>
                              <td width="99%" bgcolor="#C0C0C0" height="1" colspan="8"> </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">金流付款閘道<span lang="en-us">:</span></font></td>
                              <td>
		                            <font color="#000080">金流付款閘道<span lang="en-us">:</span></font>
                                <select size="1" name="Donate_PayName">
		                              <option value=""> </option>
		                              <option value="ecpay" <%If RS("Donate_PayName")="ecpay" Then%>selected<%End If%> >ecpay(綠界聯信)</option>
		                              <option value="gwpay" <%If RS("Donate_PayName")="gwpay" Then%>selected<%End If%> >gwpay(綠界中信)</option>
		                              <option value="maple" <%If RS("Donate_PayName")="maple" Then%>selected<%End If%> >maple(藍新測試)</option>
		                              <option value="mpp" <%If RS("Donate_PayName")="mpp" Then%>selected<%End If%> >mpp(藍新正式)</option>
		                              <option value="hotrusttest" <%If RS("Donate_PayName")="hotrusttest" Then%>selected<%End If%> >hotrust(網際威信測試)</option>
		                              <option value="hotrust" <%If RS("Donate_PayName")="hotrust" Then%>selected<%End If%> >hotrust(網際威信正式)</option>
		                            </select>
		                            &nbsp;&nbsp;&nbsp;
		                            <font color="#000080">金流身份判斷<span lang="en-us">:</span></font>
                                <select size="1" name="Check_ID_Type">
		                              <option value=""> </option>
		                              <option value="1" <%If RS("Check_ID_Type")="1" Then%>selected<%End If%> >身分證</option>
		                              <option value="2" <%If RS("Check_ID_Type")="2" Then%>selected<%End If%> >姓名+通訊地址</option>
		                            </select>
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">金流商店代碼<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="text" name="Card_Code" size="5" class="font9" maxlength="10" value=<%=RS("Card_Code")%>>&nbsp;
                                <font color="#000080">驗證碼<span lang="en-us">:</span></font>
                                <input type="text" name="Card_Verify" size="5" class="font9" maxlength="10" value=<%=RS("Card_Verify")%>>&nbsp;
                                <font color="#000080">便利付商店代碼<span lang="en-us">:</span></font>
                                <input type="text" name="Atm_Code" size="5" class="font9" maxlength="10" value=<%=RS("Atm_Code")%>>&nbsp;
                                <font color="#000080">驗證碼<span lang="en-us">:</span></font>
                                <input type="text" name="Atm_Verify" size="5" class="font9" maxlength="10" value=<%=RS("Atm_Verify")%>>&nbsp;
                                <font color="#000080">MD5 R_CODE<span lang="en-us">:</span></font>
                                <input type="text" name="MD5_Code" size="8" class="font9" maxlength="10" value=<%=RS("MD5_Code")%>>&nbsp;
                                <font color="#000080">MD5 R_CODE<span lang="en-us">:</span></font>
                                <input type="text" name="MD5_rCode" size="8" class="font9" maxlength="10" value=<%=RS("MD5_rCode")%>>&nbsp;	
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">金流啟用<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="checkbox" class="font9a" name="IsCard" value="Y" <%If RS("IsCard")="Y" Then%>checked<%End If%> ><font color="#000080">信用卡</font>
                                <input type="checkbox" class="font9a" name="IsAtm" value="Y" <%If RS("IsAtm")="Y" Then%>checked<%End If%> ><font color="#000080">WEB-ATM</font>
                                <input type="checkbox" class="font9a" name="IsCsv" value="Y" <%If RS("IsCsv")="Y" Then%>checked<%End If%> ><font color="#000080">超商代收繳費期限</font>
                                <input type="text" name="od_dt" size="3" class="font9" maxlength="2" value=<%=RS("od_dt")%>><font color="#000080">(天)</font>
                                <input type="checkbox" class="font9a" name="IsIbon" value="Y" <%If RS("IsIbon")="Y" Then%>checked<%End If%> ><font color="#000080">7-11&nbsp;ibon</font>
                                <input type="checkbox" class="font9a" name="IsFamiPort" value="Y" <%If RS("IsFamiPort")="Y" Then%>checked<%End If%> ><font color="#000080">全家&nbsp;FamiPort</font>
                                <input type="checkbox" class="font9a" name="IsLifeEt" value="Y" <%If RS("IsLifeEt")="Y" Then%>checked<%End If%> ><font color="#000080">萊爾富LiftET</font>
                                <input type="checkbox" class="font9a" name="IsDoenLoad" value="Y" <%If RS("IsDoenLoad")="Y" Then%>checked<%End If%> ><font color="#000080">授權書下載</font>
                              </td>
                            </tr>
                            <tr>
                              <td align="right"><font color="#000080">捐款免費贈閱<span lang="en-us">:</span></font></td>
                              <td>
                                <input type="checkbox" class="font9a" name="IsNews" value="Y" <%If RS("IsNews")="Y" Then%>checked<%End If%> ><font color="#000080">會訊</font>
                                <input type="checkbox" class="font9a" name="IsEpaper" value="Y" <%If RS("IsEpaper")="Y" Then%>checked<%End If%> ><font color="#000080">電子報</font>
                                <input type="checkbox" class="font9a" name="IsBooty" value="Y" <%If RS("IsBooty")="Y" Then%>checked<%End If%> ><font color="#000080">贈品發放</font>
                              </td>
                            </tr>-->                                                                                                                                                                                    
                          </table>
                        </td>
                      </tr>
                      <tr>
                        <td class="td02-c" width="100%" align="center"><button id='update' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Update_OnClick()'> <img src='../images/update.gIf' width='20' height='20' align='absmiddle'> 修改</button>&nbsp;<button id='cancel' style='position:relative;left:0;width:75;height:25;font-size:9pt' class='button3-bg' onClick='Cancel_OnClick()'> <img src='../images/icon6.gIf' width='20' height='15' align='absmiddle'> 離開</button>&nbsp;</td>
                      </tr>
                    </table>
                  </td>
                  <td height="445" class="table63-bg">&nbsp;</td>
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
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Update_OnClick(){          
  <%call SubmitJ("update")%>
}

function Cancel_OnClick(){
  location.href="dept.asp"
}
--></Script>