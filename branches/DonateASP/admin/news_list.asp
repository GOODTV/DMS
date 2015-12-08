<!--#include file="../include/dbfunctionJ.asp"-->
<%
News_Type="訊息內容"
If request("code_id")<>"" Then
  SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And Seq='"&request("code_id")&"' " 
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then News_Type=RS("News_Type")
End If  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=News_Type%>管理</title>
</head>
<body class=tool>
<%
If request("SQL")="" then
  If request("action")="query" Then
    SQL="Select NEWS.Ser_No,訊息類別=News_Type,訊息標題=News_Subject,首頁=News_ShowHome,次頁=News_ShowSubPage,頭條=News_Ad,排序=News_Seq,發佈日期=News_RegDate,刊登起日=News_BeginDate,刊登迄日=News_EndDate " & _
        "From NEWS Join CASECODE On NEWS.News_Type=CASECODE.CodeDesc Where CASECODE.CodeType='NewsType' "
    If request("code_id")<>"" Then
      SQL=SQL&"And CASECODE.Seq='"&request("code_id")&"' And NEWS.News_EndDate >= '"&Date()&"' "
    Else
      SQL=SQL&"And NEWS.News_EndDate >= '"&Date()&"'"
    End If
  Else
    SQL="Select NEWS.Ser_No,訊息類別=News_Type,訊息標題=News_Subject,首頁=News_ShowHome,次頁=News_ShowSubPage,頭條=News_Ad,排序=News_Seq,發佈日期=News_RegDate,刊登起日=News_BeginDate,刊登迄日=News_EndDate " & _
        "From NEWS Join CASECODE On NEWS.News_Type=CASECODE.CodeDesc Where CASECODE.CodeType='NewsType' "
    If request("code_id")<>"" Then
      SQL=SQL&"And CASECODE.Seq='"&request("code_id")&"' And NEWS.News_EndDate < '"&Date()&"'"
    Else
      SQL=SQL&"And NEWS.News_EndDate < '"&Date()&"'"
    End If
  End If
  If Session("user_group")<>"系統管理員" And Session("BranchType")<>"" Then SQL=SQL & "And BranchType = '"&Session("BranchType")&"' "
  If request("News_Type")<>"" Then SQL=SQL & "And News_Type = '"&request("News_Type")&"' "
  If request("KeyWord")<>"" Then SQL=SQL & "And (News_Subject Like '%"&request("KeyWord")&"%' Or News_Brief Like '%"&request("KeyWord")&"%' Or News_Desc Like '%"&request("KeyWord")&"%') "   
  SQL=SQL&" Order By CASECODE.Seq,News_Ad Desc,News_Seq,News_BeginDate Desc,News_EndDate Desc,News_RegDate Desc,NEWS.Ser_No Desc"
Else
  SQL=request("SQL")   
End If
PageSize=20
If request("nowPage")="" Then
  nowPage=1
Else
  nowPage=request("nowPage")
End If
ProgID="news_list"
HLink="news_edit.asp?code_id="&request("code_id")&"&ser_no="
LinkParam="ser_no"
LinkTarget="main"
AddLink="news_add.asp?code_id="&request("code_id")
Hit_Nemu="news"
call GridListHit (SQL,PageSize,nowPage,ProgID,HLink,LinkParam,LinkTarget,AddLink,Hit_Nemu)
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/javafunction.inc"-->