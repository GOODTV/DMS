<!--#include file="../include/dbfunctionJ.asp"-->
<%
Function DeptList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
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
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & totRec & "筆&nbsp;&nbsp;</span>"
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
    For J = 1 To FieldsCount-1
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
      For J = 1 To FieldsCount-1
        If J = FieldsCount-1 Then
          If RS1(FieldsCount)<>"" Then
            SQL2="Select 縣市鄉鎮=B.mValue+A.mValue From CODECITY A, CODECITY B Where A.mCode='"&RS1(FieldsCount)&"' And B.mCode=Right(A.codeMetaID,1)"
            Call QuerySQL(SQL2,RS2)
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1(FieldsCount)&RS2("縣市鄉鎮")&RS1(J)&"</span></td>"
            RS2.Close
            Set RS2=Nothing
          Else
            Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
          End If
        Else
          Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(J) & "</span></td>"
        End If
      Next
      I = I + 1
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
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>機構部門列表</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
<%
  SQL="select Dept_Id,機構簡稱=Comp_ShortName,網站名稱=Sys_Name,部門代號=Dept_Id,部門名稱=Dept_Desc,聯絡人=Contactor,電話=Tel,傳真=Fax,地址=Address,Zip_Code From Dept Order By Dept_Id"
  PageSize=20
  If request("nowPage")="" Then
    nowPage=1
  Else
    nowPage=request("nowPage")
  End If
  ProgID="dept_list"
  HLink="dept_edit.asp?dept_id="
  LinkParam="dept_id"
  LinkTarget="main"
  addLink=""
  If Session("user_id")="npois" Then addLink="dept_add.asp"
  call DeptList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->