<%
If Session("Web_Type")="Close" Then Response.End
response.buffer =true
response.expires=-1
ERROR_MESSSAGE=""
query_string=Trim(Cstr(LCase(Request.ServerVariables("query_string"))))

KeyInj="bchar|bdocument|bdeclare|between|buser|cast|char|count|declare|document|eval|exists|is_|libdir|phplib|script|select|substring|truncate|url|unicode|varchar|_server|.txt|`|*|'|1=1|a=a|--|'='"
Ary_KeyInj=Split(KeyInj,"|")
If Request.QueryString<>"" Then
  For Each SQL_Get In Request.QueryString
    For I = 0 To Ubound(Ary_KeyInj)
      If instr(Request.QueryString(SQL_Get),Ary_KeyInj(I))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyInj(I)&" "
    Next
  Next 
End If
call MESSSAGE("SQL_Get",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

If Request.Form<>"" Then 
  For Each Sql_Post In Request.Form
    For I = 0 To Ubound(Ary_KeyInj)
      If instr(Request.Form(Sql_Post),Ary_KeyInj(I))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyInj(I)&" "
    Next
  Next 
End If
call MESSSAGE("Sql_Post",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

If Len(query_string)>300 Then ERROR_MESSSAGE=ERROR_MESSSAGE&"length over"&"/"
call MESSSAGE("over",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeySpecial="a href|admin|alert|application|bchar|bdocument|between|buser|cache|cast|char|config|cookie|count|create|delete|document|eval|exists|exec|from|hidden|http|iframe|insert|is_|libdir|manage|master|password|phplib|script|select|substring|server|session|truncate|url|unicode|update|user|`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|;|./|//|'='"
Ary_KeySpecial=Split(KeySpecial,"|")
For I = 0 To Ubound(Ary_KeySpecial)
  If instr(query_string,Ary_KeySpecial(I))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeySpecial(I)&" "
Next
call MESSSAGE("KeySpecial",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

If Len(query_string)>255 Then ERROR_MESSSAGE=ERROR_MESSSAGE&"length over"&" "
If instr(query_string,"|")>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&"|"&" "
KeyWord="`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|application|begin|cast|cache|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|http|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|www.|xtype|%2527|address.tst"
Ary_KeyWord=Split(KeyWord,"|")
For I = 0 To Ubound(Ary_KeyWord)
  If instr(query_string,Ary_KeyWord(I))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyWord(I)&" "
Next
call MESSSAGE("KeyWord",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyHex="`|7e|!|#|$|5e|2a|(|)|2b|5b|5d|7b|7d|5c|3b|3a|'|""|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
Ary_KeyHex=Split(KeyHex,"|")
For I = 0 To Ubound(Ary_KeyHex)
  If instr(query_string,Cstr(Ary_KeyHex(I)))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyHex(I)&" "
Next
call MESSSAGE("KeyHex",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyBase="YA==|~|fg==|Iw==|JA==|Xg==|Kg==|KA==|KQ==|Kw==|Ww==|XQ==|ew==|fQ==|XA==|Ow==|Og==|Jw==|IiI=|PA==|Pg==|LA==|MT0x|YT1h|LS0=|Ow==|Li8=|LmFzcA==|LmJhaw==|LmNmbQ==|LmNzcw==|LmRvcw==|Lmh0bQ==|LmluYw==|LmluaQ==|Lmpz|LnBocA==|LnR4dA==|Lyo=|Ki8=|LkA=|QC4=|QEA=|QEM=|QFQ=|QFA=|YSBocmVm|YWRtaW4=|YWN1bmV0aXg=|YWx0ZXI=|YW5k|YXBwbGljYXRpb24=|YmVnaW4=|Y2FjaGU=|Y2FzdA==|Y29uZmln|Y2hhcg==|Y2hy|Y29va2ll|Y291bnQ=|Y3JlYXRl|Y3Nz|Y3Vyc29y|ZGVhbGxvY2F0ZQ==|ZGVjbGFyZQ==|ZGVsZXRl|ZGly|ZHJvcA==|ZWNobw==|ZW5k|ZXZhbA==|ZXhlYw==|ZXhpc3Rz|ZmV0Y2g=|ZnJvbQ==|aGlkZGVu|aHR0cA==|aWZyYW1l|aW5zZXJ0|aW50bw==|aXNf|am9pbg==|anM=|a2lsbA==|bGVmdA==|bWFuYWdl|bWFzdGVy|bWlk|bmNoYXI=|bnRleHQ=|bnZhcmNoYXI=|b3Blbg==|cGFzc3dvcmQ=|cmlnaHQ=|c2NyaXB0|c2V0|c2VsZWN0|c2Vzc2lvbg==|c3Jj|c3lz|c3lzb2JqZWN0cw==|c3lzY29sdW1ucw==|dGFibGU=|dGV4dA==|dHJ1bmNhdGU=|dXJs|dXNlcg==|dXBkYXRl|dmFyY2hhcg==|d2hlcmU=|d2hpbGU=|d3d3|eHR5cGU=|JTI1Mjc=|YWRkcmVzcy50c3Q="
Ary_KeyBase=Split(KeyBase,"|")
For I = 0 To Ubound(Ary_KeyBase)
  If instr(query_string,Ary_KeyBase(I))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyBase(I)&" "
Next
call MESSSAGE("KeyBase",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyScript="011100110110001101110010011010010111000001110100|736372697074|c2NyaXB0|11599114105112116"
Ary_KeyScript=Split(KeyScript,"|")
For I = 0 To Ubound(Ary_KeyScript)
  If instr(query_string,Cstr(Ary_KeyScript(I)))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyScript(I)&" "
Next
call MESSSAGE("KeyScript",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyIframe="011010010110011001110010011000010110110101100101|696672616d65|aWZyYW1l|10510211497109101"
Ary_KeyIframe=Split(KeyIframe,"|")
For I = 0 To Ubound(Ary_KeyIframe)
  If instr(query_string,Cstr(Ary_KeyIframe(I)))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyIframe(I)&" "
Next
call MESSSAGE("KeyIframe",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyAlert="0110000101101100011001010111001001110100|616c657274|YWxlcnQ=|97108101114116"
Ary_KeyAlert=Split(KeyAlert,"|")
For I = 0 To Ubound(Ary_KeyAlert)
  If instr(query_string,Cstr(Ary_KeyAlert(I)))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyAlert(I)&" "
Next
call MESSSAGE("KeyAlert",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

KeyOther="sample@email.tst|JyI=|acu75e3fa3f006b30025eb226cec0acecf6|PFNjUmlQdD5hbGVydCgneHNzLXRlc3QnKTs8L1NjUmlQdD4=|IiBvbm1vdXNlb3Zlcj0iYWxlcnQoJ3hzcy10ZXN0Jyk=|dir|..|boot.ini|268435455|-268435455|0x7fffffff|0x80000000|0x3fffffff|0xffffffff|/some_inexistent_file_with_long_name|www.haesung.hs.kr"
Ary_KeyOther=Split(KeyOther,"|")
For I = 0 To Ubound(Ary_KeyOther)
  If instr(query_string,Cstr(Ary_KeyOther(I)))>0 Then ERROR_MESSSAGE=ERROR_MESSSAGE&Ary_KeyOther(I)&" "
Next
call MESSSAGE("KeyOther",ERROR_MESSSAGE)
ERROR_MESSSAGE=""

Function MESSSAGE (KEY_NAME,ERROR_QUERY_STRING)
  If ERROR_QUERY_STRING<>"" Then 
    session.Timeout=600
    Server.ScriptTimeout=600
    Set Conn=server.createobject("ADODB.Connection")
    Conn.connectiontimeout=600
    Conn.commandtimeout=600
    Conn.Provider="sqloledb"
    Conn.open "server=(local);uid=sa;pwd=kc1997;database=DONATION"
    'Conn.open "server=HFS23;uid=dbread;pwd=dbread1679pwd;database=DONATION"
    SQL="QUERY_STRING"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,3
    RS.Addnew
    RS("QUERY_TYPE")=KEY_NAME
    RS("QUERY_STRING_DATE")=Date()
    RS("QUERY_STRING_DATETIME")=Now()
    RS("REMOTE_HOST")=Request.ServerVariables("REMOTE_HOST")
    RS("REMOTE_ADDR")=Request.serverVariables("REMOTE_ADDR")
    RS("SERVER_NAME_URL")=Left(Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL"),100)
    RS("HTTP_ACCEPT_LANGUAGE")=Left(Request.ServerVariables("HTTP_ACCEPT_LANGUAGE"),50)
    RS("HTTP_USER_AGENT")=Left(Request.ServerVariables("HTTP_USER_AGENT"),300)
    RS("HTTP_REFERER")=Left(Request.ServerVariables("HTTP_REFERER"),300)
    RS("SCRIPT_NAME")=Left(Request.ServerVariables("SCRIPT_NAME"),300)
    RS("QUERY_STRING")=Left(Request.ServerVariables("QUERY_STRING"),300)
    RS("ERROR_MESSSAGE")=Left(ERROR_QUERY_STRING,300)
    RS.Update
    RS.Close
    Set RS=Nothing

    If KEY_NAME="SQL_Get" Or KEY_NAME="Sql_Post" Then Session("Web_Type")="Close"
    response.redirect "index.asp"
    
    Total1=0
    SQL="Select Total=Count(*) From QUERY_STRING Where QUERY_STRING_DATE='"&Date()&"' And REMOTE_ADDR='"&Request.serverVariables("REMOTE_ADDR")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      Total1=Cdbl(RS("Total"))
    End If
    RS.Close
    Set RS=Nothing

    StrDate=DateAdd("D",-3,date())
    Total2=0
    SQL="Select Total=Count(*) From QUERY_STRING Where QUERY_STRING_DATE>='"&StrDate&"' And REMOTE_ADDR='"&Request.serverVariables("REMOTE_ADDR")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      Total2=Cdbl(RS("Total"))
    End If
    RS.Close
    Set RS=Nothing

    StrDate=DateAdd("D",-7,date())
    Total3=0
    SQL="Select Total=Count(*) From QUERY_STRING Where QUERY_STRING_DATE>='"&StrDate&"' And REMOTE_ADDR='"&Request.serverVariables("REMOTE_ADDR")&"'"
    Set RS = Server.CreateObject("ADODB.RecordSet")
    RS.Open SQL,Conn,1,1
    If Not RS.EOF Then
      Total3=Cdbl(RS("Total"))
    End If
    RS.Close
    Set RS=Nothing
    Conn.Close
    Set Conn=Nothing

    If Total1>=10 Or Total2>=20 Or Total3>=30 Then Session("Web_Type")="Close"
    response.redirect "index.asp"
  End If
End Function
%>