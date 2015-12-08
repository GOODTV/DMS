<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function DonateWebList (SQL1,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,1
  If Not RS1.EOF Then 
    FieldsCount = RS1.Fields.Count-1
    totRec=RS1.Recordcount
    If totRec>0 Then 
      RS1.PageSize=PageSize
      If nowPage="" or nowPage=0 Then
        nowPage=1
      ElseIf cint(nowPage) > RS1.PageCount Then 
        nowPage=RS1.PageCount 
      End If
      session("nowPage")=nowPage
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      SQL=server.URLEncode(SQL)
    End If
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'> </td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & FormatNumber(totRec,0) & "筆&nbsp;&nbsp;</span>"
    If cint(nowPage) <>1 Then
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    End If
    If cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount Then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    End If
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體' onchange='GoPage_OnChange(this.value)'>"
    For iPage=1 to totPage
      If iPage=cint(nowPage) Then
        strSelected = "selected"
      Else
	      strSelected = ""
      End If
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"
    Next
    Response.Write "</select>頁</span></td>" 
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If   
    Response.Write "</tr></table>"
    Dim I
    Dim J
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For J = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J).Name & "</span></font></td>"
    Next
    Response.Write "</tr>"
    I = 1
    While Not RS1.EOF And I <= RS1.PageSize
      If Hlink<>"" Then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      End If
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For J = 1 To FieldsCount
        If J=2 Then
          If RS1(2)="webatm" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>Web-ATM</span></td>" 
          ElseIf RS1(2)="vacc" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>銀行虛擬帳號</span></td>" 
          ElseIf RS1(2)="barcode" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>超商條碼</span></td>" 
          ElseIf RS1(2)="ibon" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>7-11 ibon</span></td>" 
          ElseIf RS1(2)="famiport" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>全家 Famiport</span></td>" 
          ElseIf RS1(2)="lifeet" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>萊爾富 Life-Et</span></td>"
          ElseIf RS1(2)="okgo" Then
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>OK超商 OK-GO</span></td>"           
          Else
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
          End If
        ElseIf J=5 Then
          If RS1(J)<>"" Then
            Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & FormatNumber(RS1(J),0) & "</span></td>"
          Else
            Response.Write "<td align=""right""><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
          End If
        ElseIf J=6 Then
          If RS1(6)<>"" Then
            expire_datetime=Cstr(RS1(6))
            expire_datetime=Left(Replace(Replace(Replace(expire_datetime,"-","")," ",""),":",""),12)
            expire_date_time=""
            If expire_datetime<>"" Then expire_date_time=Left(expire_datetime,4)&"/"&Mid(expire_datetime,5,2)&"/"&Mid(expire_datetime,7,2)&" "&Mid(expire_datetime,9,2)&":"&Mid(expire_datetime,11,2)
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & expire_date_time & "</span></td>"
          Else
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>&nbsp;</span></td>"
          End If
        ElseIf J=7 Then
          If RS1(6)="" Then
            If Donate_Type="barcode" Or Donate_Type="ibon" Or Donate_Type="famiport" Or Donate_Type="lifeet" Or Donate_Type="okgo" Then
              Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>交易失敗</span></td>"
            Else
              Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
            End If
          Else
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
          End If
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
        End If
      Next
      I = I + 1
	    Response.Flush
      Response.Clear      
      RS1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  Else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    If AddLink <> "" Then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    End If
    Response.Write "</table>"
  End If
  RS1.Close
  Set RS1=Nothing
End Function
%>
<%Prog_Id="ecbank_qry"%>
<!--#include file="../include/head.asp"-->
<body class=tool>
<%
If request("SQL")="" Then
  SQL="Select A.Ser_No,捐款人=A.Donate_DonorName,交易方式=A.Donate_Type,交易日期=A.Donate_CreateDateTime,交易編號=A.od_sob,交易金額=(Case When CONVERT(numeric,B.amt)>0 Then CONVERT(numeric,B.amt) Else CONVERT(numeric,A.Donate_Amount) End)," & _
      "繳費截止日期=(Case When A.Donate_Type='webatm' Then CONVERT(nvarchar,Donate_CreateDateTime,120) Else Case When B.get_expire_date+B.get_expire_time<>'' Then B.get_expire_date+B.get_expire_time Else '' End End)," & _
      "繳費狀態=(Case When B.close_type='ok' Then '付款完成' Else '' End) " & _
      "From DONATE_WEB A Left Join DONATE_ECBANK As B On A.od_sob=B.od_sob Where A.Donate_Type<>'creditcard' "
  If Request("Donate_DonorName")<>"" Then SQL=SQL&"And Donate_DonorName Like '%"&Request("Donate_DonorName")&"%' "
  If Request("Donate_Purpose")<>"" Then SQL=SQL&"And Donate_Purpose = '"&Request("Donate_Purpose")&"' "
  If Request("Donate_Invoice_Type")<>"" Then SQL=SQL&"And Donate_Invoice_Type = '"&Request("Donate_Invoice_Type")&"' "
  If Request("Donate_CreateDate_B")<>"" Then SQL=SQL&"And Donate_CreateDate >= '"&Request("Donate_CreateDate_B")&"' "
  If Request("Donate_CreateDate_E")<>"" Then SQL=SQL&"And Donate_CreateDate <= '"&Request("Donate_CreateDate_E")&"' "
  If Request("Donate_Type")<>"" Then SQL=SQL&"And Donate_Type = '"&Request("Donate_Type")&"' "
  If Request("Donate_Export")<>"" Then SQL=SQL&"And Donate_Export = '"&Request("Donate_Export")&"' "
  If Request("Close_Type")="L" Then SQL=SQL&"And (Case When A.Donate_Type='webatm' Then CONVERT(nvarchar,Donate_CreateDateTime,120) Else Case When B.get_expire_date+B.get_expire_time<>'' Then B.get_expire_date+B.get_expire_time Else '' End End)='' "
  If Request("Close_Type")="N" Then SQL=SQL&"And (Close_Type Is null Or Close_Type='') "
  If Request("Close_Type")="Y" Then SQL=SQL&"And (Close_Type = 'ok') "
  SQL=SQL&"Order By Donate_CreateDateTime Desc,A.Ser_No Desc"
Else
  SQL=request("SQL")
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="ecbank_qry_list"
HLink=""
LinkParam="ser_no"
LinkTarget="main"
AddLink=""
If request("action")="report" Or request("action")="export" Then Server.Transfer "ecbank_qry_rpt.asp"
call DonateWebList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->