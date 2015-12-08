<!--#include file="../include/dbfunction.asp"-->
<%
Function ShowGrid2 (SQL,PageSize,nowPage,ProgID)
  Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL,Conn,1,1
  if Not RS1.EOF then 
    FieldsCount = rs1.Fields.Count-1
    totRec=RS1.Recordcount         '總筆數
    if totRec>0 then 
      RS1.PageSize=PageSize       '每頁筆數
      if nowPage="" or nowPage=0 then 
        nowPage=1
      elseif cint(nowPage) > RS1.PageCount then 
        nowPage=RS1.PageCount 
      end if    
      session("nowPage")=nowPage        	
      RS1.AbsolutePage=nowPage
      totPage=RS1.PageCount       '總頁數
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
    
    Response.Write "</tr></table>"
    Dim i
    Dim j
    Response.Write "<table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
    Response.Write "<tr>"
    For j = 1 To FieldsCount
      Response.Write "<td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>" & rs1(j).Name & "</span></font></td>"
    Next
    Response.Write "<td bgcolor='#FFE1AF' nowrap><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></td>"
    Response.Write "</tr>"
    i = 1
    While Not rs1.EOF And i <= rs1.PageSize 
      Response.Write "<tr bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
      For j = 1 To FieldsCount
        Response.Write "<td><span style='font-size: 9pt; font-family: 新細明體'>" & RS1(j) & "</span></td>"
      Next
      Response.Write "<td><a href='JavaScript:if(confirm(""是否確定要刪除 ?"")){window.location.href=""syslog_delete.asp?ser_no="&RS1(0)&""";}' target='" & LinkTarget &"'><img src='../images/x5.gif' border=0 width='16' height='14' alt='刪除'></a></td>"
      i = i + 1
      rs1.MoveNext
      Response.Write "</tr>"
    Wend
    Response.Write "</table>"
    Response.Write "</form>"
  else
    Response.Write "<table border=0 cellspacing='0' cellpadding='2' style='border-collapse: collapse' width='100%'>"
    Response.Write "<tr><td width='85%' align='center' style='color:#ff0000'>** 沒有符合條件的資料 **</td>" 
    Response.Write "</table>"
  end if 
  rs1.close
End Function
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>系統LOG查詢</title>
</head>
<body class=tool>
<% 
SQL="select ser_no,LOG時間=Log_Time,機構=comp_shortname,帳號=User_ID,帳號名稱=User_Name,LOG內容=Log_Desc,登入IP=IP_Address from sys_log join dept on sys_log.dept_id=dept.dept_id"_ 
+"     where ('"&request("dept_id")&"'='' or sys_log.dept_id='"&request("dept_id")&"') and"_ 
+"     (log_date between '"&request("date1")&"' and '"&request("date2")&"' or '"&request("date1")&"'='') and"_
+"     (user_id='"&request("userid")&"' or '"&request("userid")&"'='') and"_
+"     (log_type='"&request("log_type")&"' or '"&request("log_type")&"'='') order by log_time desc"
PageSize=20
if request("nowPage")="" then
  nowPage=1
else
  nowPage=request("nowPage")
end if
ProgID="syslog_qry_d"
call ShowGrid2(SQL,PageSize,nowPage,ProgID) 
%> 
</body> 
</html> 
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->