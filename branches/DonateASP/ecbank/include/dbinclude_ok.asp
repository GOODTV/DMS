<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%Session.CodePage=65001%>
<%
session.Timeout=600
Server.ScriptTimeout=600
set conn=server.createobject("ADODB.Connection")
conn.connectiontimeout=600
conn.commandtimeout=600
conn.Provider="sqloledb"
conn.open "server=(local);uid=sa;pwd=kc1997;database=DONATION"
'conn.open "server=HFS23;uid=dbwrite;pwd=dbwrite1679pwd;database=DONATION"
%>
<!--#include file="dbtool.asp"-->