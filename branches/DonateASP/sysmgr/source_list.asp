<!--#include file="../include/dbfunction.asp"-->
<%
Function SourceList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,DelLink,DelParam)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  if Not RS1.EOF then 
    FieldsCount = rs1.Fields.Count-1
    totRec=RS1.Recordcount
    if totRec>0 then 
      RS1.PageSize=PageSize
      if nowPage="" or nowPage=0 then 
        nowPage=1
      elseif cint(nowPage) > RS1.PageCount then 
        nowPage=RS1.PageCount 
      end if    
      session("nowPage")=nowPage        	
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount
      Sql=server.URLEncode(Sql)
    end if    
    Response.Write "<Form action='' id=form1 name=form1>" 
    Response.Write "<table border=0 cellspacing='0' cellpadding='1' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='30%'></td>"
    Response.Write "<td width='55%'><span style='font-size: 9pt; font-family: 新細明體'>第" & nowPage & "/" & totPage & "頁&nbsp;&nbsp;</span>"
    Response.Write "<span style='font-size: 9pt; font-family: 新細明體'>共" & totRec & "筆&nbsp;&nbsp;</span>"
    if cint(nowPage) <>1 then             
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage-1)&"&SQL="&SQL&"'>上一頁</a>" 
    end if      
    if cint(nowPage)<>RS1.PageCount and cint(nowPage)<RS1.PageCount then 
      Response.Write " | <a href='"&ProgID&".asp?nowPage="&(nowPage+1)&"&SQL="&SQL&"'>下一頁</a>" 
    end if
    Response.Write " |&nbsp;<span style='font-size: 9pt; font-family: 新細明體'> 跳至第<select name=GoPage size='1' style='font-size: 9pt; font-family: 新細明體'>"
    For iPage=1 to totPage
      if iPage=cint(nowPage) then
        strSelected = "selected"
      else
	strSelected = "" 
      end if   
      Response.Write "<option value='"&iPage&"'" & strSelected & ">" & iPage & "</option>"          
    Next   
    Response.Write "</select>頁</span></td>" 
    if AddLink <> "" then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    end if   
    Response.Write "</tr></table>"
    Dim i
    Dim j
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For j = 1 To FieldsCount
      if j=3 then
        Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & rs1(j).Name & "( KB )</span></font></td>"
      else
        Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & rs1(j).Name & "</span></font></td>"
      end if
    Next
    Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
    Response.Write "</tr>"
    i = 1
    While Not rs1.EOF And i <= rs1.PageSize
      if Hlink<>"" then
        Response.Write "<a href='" & HLink & RS1(LinkParam) & "' target='" & LinkTarget &"'>"
        showhand="style='cursor:hand'"
      end if   
      Response.Write "<tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For j = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(j) & "</span></td>"
      Next
      Response.Write "<td><a href='JavaScript:if(confirm(""是否確定要刪除【"&RS1(2)&"】這個檔案\n\n這會可能會造成系統不完整 ？"")){window.location.href=""" & DelLink & RS1(DelParam) & """;}' target='" & DataTarget &"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>"
      i = i + 1
      rs1.MoveNext
      Response.Write "</tr>"
      Response.Write "</a>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>"
    if AddLink <> "" then
      Response.Write "<td align='right'><span style='font-size: 9pt; font-family: 新細明體'><img border='0' src='../images/DIR_tri.gif' align='absmiddle'> <a href='" & AddLink & "' target='main'>新增資料</a></span></td>"
    end if  
    Response.Write "</table>"
  end if 
  rs1.close
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>檔案上傳管理</title>
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
</head>
<body class=tool>
<%
If request("SQL")="" then
  sys_date=date()
  SQL="Select Ser_No,檔案路徑=Source_RootType,檔案名稱=Source_FileURL,檔案大小=Source_FileSize,上傳日期=Source_RegDate From SOURCE "
  If request("keyword")<>"" Then SQL=SQL&" Where Source_FileURL Like '%"&request("keyword")&"%' "   
  SQL=SQL&"Order By Ser_No Desc"
Else
  SQL=request("SQL")   
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="source_list"
HLink=""
LinkParam="ser_no"
LinkTarget="main"
AddLink="source_add.asp"
DelLink="source_delete.asp?ser_no="
DelParam="ser_no"
call SourceList (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,DelLink,DelParam)
%>
<%message%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->