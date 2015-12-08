<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<%
  If Request.Form("action")="donation" Then  
    'Submit連結網址
    If Request.Form("Donate_Type")="creditcard" Then
      Donate_Desc="網路信用卡"
      SubmitPage="ecpayb.asp"
      payment_type="creditcard"
    ElseIf Request.Form("Donate_Type")="webatm" Then
      Donate_Desc="Web-ATM"
      SubmitPage="ecbank_atm.asp"
      payment_type="web_atm"
    ElseIf Request.Form("Donate_Type")="vacc" Then
      Donate_Desc="銀行虛擬帳號"
      SubmitPage="ecbank_vacc.asp"
      payment_type="vacc"      
    ElseIf Request.Form("Donate_Type")="barcode" Then
      Donate_Desc="超商條碼"
      SubmitPage="ecbank_barcode.asp"
      payment_type="barcode"
    ElseIf Request.Form("Donate_Type")="ibon" Then
      Donate_Desc="7-11 ibon"
      SubmitPage="ecbank_4in1.asp"
      payment_type="ibon"
    ElseIf Request.Form("Donate_Type")="famiport" Then
      Donate_Desc="全家 Famiport"
      SubmitPage="ecbank_4in1.asp"
      payment_type="cvs"
    ElseIf Request.Form("Donate_Type")="lifeet" Then
      Donate_Desc="萊爾富 Life-Et"
      SubmitPage="ecbank_4in1.asp"
      payment_type="cvs"
    ElseIf Request.Form("Donate_Type")="okgo" Then
      Donate_Desc="OK超商 OK-GO"
      SubmitPage="ecbank_4in1.asp"
      payment_type="cvs"      
    End If
  Else
    session("errnumber")=1
    session("msg")="很抱歉您輸入的訊息不足，誠摯地請您再試一次！"
    Response.Redirect "ecpay.asp"
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <link href="include/donate.css" rel="stylesheet" type="text/css">
  <script src="js/npois.js" type="text/javascript"></script>
  <script src="js/ecpaya.js" type="text/javascript"></script>
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->  
  <title><%=Comp_Name%></title>
</head>
<body onload="Window_OnLoad()" oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()" onbeforecopy="return false">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" />  </div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 捐款資料填寫</div>
    <div id="mid">
      <div id="mid-1">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" class="page_form">
          <form name="form_donate" method="POST" action="<%=SubmitPage%>">
            <input type="hidden" name="action">
            <input type="hidden" name="Donate_Type" value="<%=Request.Form("Donate_Type")%>">
            <input type="hidden" name="Donate_Desc" value="<%=Donate_Desc%>">
            <input type="hidden" name="payment_type" value="<%=payment_type%>">
            <input type="hidden" name="Donate_Birthday" value="">
            <input type="hidden" name="Donate_ZipCode" value="">
            <input type="hidden" name="Donate_Invoice_ZipCode" value="">            	
            <%
              SQL1="Select Dept_Id,Comp_ShortName From DEPT Order By Comp_Label,Dept_Id"
              Set RS1 = Server.CreateObject("ADODB.RecordSet")
              RS1.Open SQL1,Conn,1,1
              If RS1.Recordcount>1 Then
                Response.Write "<tr>"&vbcrlf 
                Response.Write "  <td align=""left"" class=""contents"" colspan=""4""><span class=""mustcolumn"">◎捐款指定機構(*為必填欄位)</span></td>"&vbcrlf
                Response.Write "</tr>"&vbcrlf
                Response.Write "<tr>"&vbcrlf
                Response.Write "  <td align=""right"" class=""contents"">機構名稱：<span class=""mustcolumn"">*</span></td>"&vbcrlf
                Response.Write "  <td class=""contents"" colspan=""3"">"&vbcrlf
                Response.Write "    <SELECT Name=""Donate_DeptId"" size=""1"" style=""font-size: 9pt; font-family: 新細明體"">"&vbcrlf
                While Not RS1.EOF
                  Response.Write "    <OPTION value="""&RS1("Dept_Id")&""">"&RS1("Comp_ShortName")&"</OPTION>"&vbcrlf
                  RS1.MoveNext
                Wend
                Response.Write "    </SELECT>"&vbcrlf
                Response.Write "  </td>"&vbcrlf
                Response.Write "</tr>"&vbcrlf
                Response.Write "<tr>"&vbcrlf
                Response.Write "  <td align=""left"" class=""contents"" colspan=""4""><span class=""mustcolumn"">◎捐款人資料</span></td>"&vbcrlf
                Response.Write "</tr>"&vbcrlf
              Else
                Response.Write "<input type=""hidden"" name=""Donate_DeptId"" value="""&RS1("Dept_Id")&""">"&vbcrlf
                Response.Write "<tr>"&vbcrlf
                Response.Write "  <td align=""left"" class=""contents"" colspan=""4""><span class=""mustcolumn"">◎捐款人資料(*為必填欄位)</span></td>"&vbcrlf
                Response.Write "</tr>"&vbcrlf
              End If
              RS1.Close
              Set RS1=Nothing
            %>
            <tr>
              <td width="160" align="right" class="contents">捐款金額：<span class="mustcolumn">*</span></td>
              <td width="160" class="contents"><input name="Donate_Amount" type="text" size="16" maxlength="6" /></td>
              <td width="100" align="right" class="contents">捐款方式：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td width="250" class="contents"><%=Donate_Desc%> </td>
            </tr>
           <tr>
              <td align="right" class="contents">真實姓名：<span class="mustcolumn">*</span></td>
              <td class="contents"><input name="Donate_Name" type="text" size="16" maxlength="20" value="<%=Donate_Name%>" /></td>
              <td align="right" class="contents">性&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;別：<span class="mustcolumn">*</span></td>
              <td class="contents">
                <input type="radio" name="Donate_Sex" value="男" <%If Donate_Sex<>"女" Then%>checked<%End If%> >男&nbsp;
                <input type="radio" name="Donate_Sex" value="女" <%If Donate_Sex="女" Then%>checked<%End If%> >女
              </td>
            </tr>
            <tr>
              <td align="right" class="contents">身分證：<span class="mustcolumn">&nbsp;</span></td>
              <td class="contents"><input name="Donate_IDNO" type="text" size="16" maxlength="10" value="<%=Donate_IDNO%>" onkeyup="javascript:UCaseIDNO();" /></td>
              <td align="right" class="contents">出生日期：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents">
              <%
                If Donate_Birthday<>"" Then
                  Ary_Birthday=Split(Donate_Birthday,"/")
                  Birthday_Year=Ary_Birthday(0)
                  Birthday_Month=Ary_Birthday(1)
                  Birthday_Day=Ary_Birthday(2)
                Else
                  Birthday_Year=""
                  Birthday_Month=""
                  Birthday_Day=""
                End If
                Response.Write"<select name='Donate_Birthday_Year' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                Response.Write"<option  value=''></option>"
                For I=Year(Date)-15 To Year(Date)-100 step -1
                  If Cstr(I)=Cstr(Birthday_Year) Then
                    Response.Write"<option  value='"&I&"' selected >"&I&"</option>"
                  Else
                    Response.Write"<option  value='"&I&"'>"&I&"</option>"
                  End If
                Next
                Response.Write "</select>年&nbsp;"
                Response.Write"<select name='Donate_Birthday_Month' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                Response.Write"<option  value=''></option>"
                For I=1 To 12
                  If Cstr(I)=Cstr(Birthday_Month) Then
                    Response.Write"<option  value='"&I&"' selected >"&I&"</option>"
                  Else
                    Response.Write"<option  value='"&I&"'>"&I&"</option>"
                  End If
                Next
                Response.Write "</select>月&nbsp;"
                Response.Write"<select name='Donate_Birthday_Day' size='1' style='font-size: 9pt; font-family: 新細明體'>"
                Response.Write"<option  value=''></option>"
                For I=1 To 31
                  If Cstr(I)=Cstr(Birthday_Day) Then
                    Response.Write"<option  value='"&I&"' selected >"&I&"</option>"
                  Else
                    Response.Write"<option  value='"&I&"'>"&I&"</option>"
                  End If
                Next
                Response.Write "</select>日"
              %>	
              </td>
            </tr>
            <tr>
              <td align="right" class="contents">教育程度：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents">
              <%
                SQL="Select Donate_Education=CodeDesc From CASECODE Where CodeType='Education' Order By Seq"
                FName="Donate_Education"
                Listfield="Donate_Education"
                menusize="1"
                BoundColumn=Donate_Education
                call OptionList (SQL,FName,Listfield,BoundColumn,NoChecked)
              %>
              </td>
              <td align="right" class="contents">職&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;業：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents">
              <%
                SQL="Select Donate_Occupation=CodeDesc From CASECODE Where CodeType='Occupation' Order By Seq"
                FName="Donate_Occupation"
                Listfield="Donate_Occupation"
                menusize="1"
                BoundColumn=Donate_Occupation
                call OptionList (SQL,FName,Listfield,BoundColumn,NoChecked)
              %>
              </td>
            </tr>
            <tr>
              <td align="right" class="contents">手&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;機：<span class="mustcolumn">*</span></td>
              <td class="contents" colspan="3"><input name="Donate_CellPhone" type="text" size="16" maxlength="10" value="<%=Donate_CellPhone%>" />&nbsp;(格式:0985885885)</td>
            </tr>
            <tr>
              <td align="right" class="contents">聯絡電話：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents" colspan="3"><input name="Donate_TelOffice" type="text" size="16" maxlength="20" value="<%=Donate_TelOffice%>" /></td>
            </tr>
            <tr>
              <td align="right" class="contents">通訊地址：<span class="mustcolumn">*</span></td>
              <td colspan="3" class="contents"><%call CodeCity2 ("form_donate","Donate_ZipCode",Donate_ZipCode,"Donate_CityCode",Donate_CityCode,"Donate_AreaCode",Donate_AreaCode,"Donate_Address",Donate_Address,"35","N")%><br /><span class="mustcolumn">&nbsp;郵遞區號不必填寫</span></td>
            </tr>
            <tr>
              <td align="right" class="contents">電子郵件：<span class="mustcolumn">*</span></td>
              <td colspan="3" class="contents"><input name="Donate_Email" type="text" size="40" maxlength="60" value="<%=Donate_Email%>" /></td>
            </tr>
            <tr>
              <td align="left" class="contents" colspan="4"><span class="mustcolumn">◎捐款用途</span></td>
            </tr>
            <tr>
              <td align="right" class="contents">捐款用途：<span class="mustcolumn">*</span></td>
              <td class="contents">
              <%
                SQL="Select Donate_Purpose=CodeDesc From CASECODE Where CodeType='Purpose' And CodeKind='D' Order By Seq"
                FName="Donate_Purpose"
                Listfield="Donate_Purpose"
                menusize="1"
                BoundColumn=Donate_Purpose
                call OptionList (SQL,FName,Listfield,BoundColumn,NoChecked)
              %>
              </td>
              <%
                SQL1="Select Act_Id,Act_ShortName From ACT Where ('"&Date()&"' Between Act_BeginDate And Act_EndDate) Order By Act_id Desc"
                Set RS1 = Server.CreateObject("ADODB.RecordSet")
                RS1.Open SQL1,Conn,1,1
                If RS1.Recordcount>0 Then
                  Response.Write "<td align=""right"" class=""contents"">勸募活動：<span class=""mustcolumn"">&nbsp;&nbsp;</span></td>"&vbcrlf
                  Response.Write "<td class=""contents"">"&vbcrlf
                  Response.Write "  <SELECT Name=""Donate_ActId"" size=""1"" style=""font-size: 9pt; font-family: 新細明體"">"&vbcrlf
                  Response.Write "    <OPTION value=""""></OPTION>"&vbcrlf
                  While Not RS1.EOF
                    Response.Write "  <OPTION value="""&RS1("Act_Id")&""">"&RS1("Act_ShortName")&"</OPTION>"&vbcrlf
                    RS1.MoveNext
                  Wend
                  Response.Write "  </SELECT>"&vbcrlf
                  Response.Write "</td>"&vbcrlf
                Else
                  Response.Write "<input type=""hidden"" name=""Donate_ActId"" value="""">"&vbcrlf
                  Response.Write "<td align=""right"" class=""contents"">&nbsp;</td>"&vbcrlf
                  Response.Write "<td class=""contents"">&nbsp;</td>"&vbcrlf
                End If
                RS1.Close
                Set RS1=Nothing
              %>
            </tr>
            <tr>
              <td align="left" class="contents" colspan="4"><span class="mustcolumn">◎捐款收據資料</span></td>
            </tr>
            <tr>
              <td align="right" class="contents">收據寄發方式：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td colspan="3" class="contents">
			        <%
			          Row=1
			          SQL="Select Donate_Invoice_Type=CodeDesc From CASECODE Where CodeType='InvoiceType' Order By Seq"
                Set RS1 = Server.CreateObject("ADODB.RecordSet")
                RS1.Open SQL,Conn,1,1
                While Not RS1.EOF
                  If Cstr(RS1("Donate_Invoice_Type"))=Cstr(Donate_Invoice_Type) Then
                    StrChecked="checked"
                  Else
                    If Row=1 Then
                      StrChecked="checked"
                    Else
                      StrChecked=""
                    End If
                  End If 
                  Response.Write "<input type='radio' " & StrChecked & " name='Donate_Invoice_Type' id='Donate_Invoice_Type"&Row&"' value='"&RS1("Donate_Invoice_Type")&"'>&nbsp;"&RS1("Donate_Invoice_Type")
                  If Instr(RS1("Donate_Invoice_Type"),"年")>0 Then  Response.Write "(次年四月底前陸續寄發)"
                  Row=Row+1
                  RS1.MoveNext
                Wend
                RS1.Close
                Set RS1=Nothing
              %>
              </td>
            </tr>
            <tr>
              <td align="right" class="contents">收據抬頭：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents" colspan="3">
                <input name="Donate_Invoice_Title" type="text" size="16" maxlength="20" value="<%=Donate_Invoice_Title%>" />
                <span class="mustcolumn"><input type="checkbox" name="IsSameAddress" value="Y" OnClick="SameAddress_OnClick()"><font color="#ff0000">收據資料同上</font></span>
              </td>
            </tr>
            <tr>
              <td align="right" class="contents">收據身分證：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents"><input name="Donate_Invoice_IDNO" type="text" size="16" maxlength="10" value="<%=Donate_Invoice_IDNO%>" onkeyup="javascript:UCaseInvoiceIDNO();" /></td>
            </tr>
            <tr>
              <td align="right" class="contents">收據地址：<span class="mustcolumn">&nbsp;&nbsp;</span></td>
              <td class="contents" colspan="3"><%call CodeCity2 ("form_donate","Donate_Invoice_ZipCode",Donate_Invoice_ZipCode,"Donate_Invoice_CityCode",Donate_Invoice_CityCode,"Donate_Invoice_AreaCode",Donate_Invoice_AreaCode,"Donate_Invoice_Address",Donate_Invoice_Address,"35","N")%><br /><span class="mustcolumn">&nbsp;郵遞區號不必填寫</span></td>
            </tr>
            <tr>
              <td align="center" colspan="4"><input type="button" value=" 下一步 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:Next_OnClick();"></td>
            </tr>
          </form>
        </table>      
      </div>
    </div>
    <div id="bottom"></div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->