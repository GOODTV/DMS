<!--#include file="../include/dbfunction.asp"-->

<%
SQL="delete from sys_log where ser_no='"&request("ser_no")&"'"
call ExecSQL(SQL)
response.redirect "syslog_qry_d.asp?nowPage="&session("nowPage")
%><!--#include file="../include/dbclose.asp"-->