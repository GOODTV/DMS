<%
response.buffer =true
response.expires=-1

Function check_post()
  check_post = True
  URL = Cstr(Request.ServerVariables("URL"))
  URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
  URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  If Instr(URL1,"https://")>0 Then
    UrlLen=9
    If URL<>"" Then
      Ary_Url=Split(URL,"/")
      For I=1 To UBound(Ary_Url)-1
        If Cstr(Ary_Url(I))<>"admin" Then
          URL2=URL2&"/"&Ary_Url(I)
        Else
          I=UBound(Ary_Url)
        End If
      Next    
    End If
  Else
    UrlLen=8
  End If
  If Mid(URL1,UrlLen,Len(URL2))<>URL2 Then check_post = False
End Function

Function check_post2()
  check_post = True
  URL1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
  URL2 = Cstr(Request.ServerVariables("SERVER_NAME"))
  If Mid(URL1,8,Len(URL2))<>URL2 Then check_post = False
End Function

Function check_keyword(query_string)
  If query_string<>"" Then 
    call check_query_string(query_string,"")
  End If
End Function

Function check_string(query_string,qs_space,qs_type,qs_len)
  If query_string="" Then
    If qs_space="Y" Then response.redirect "/index.asp"
  Else
    If Cint(Len(query_string))>Cint(qs_len) Then response.redirect "/index.asp"
    If qs_type="email" Then
      If instr(query_string,"@")=0 Or instr(query_string,".")=0 or left(query_string,1)="@" Or right(query_string,1)="@" Or left(query_string,1)="." Or right(query_string,1)="." Then response.redirect "/index.asp"
    ElseIf qs_type="number" Then
      If IsNumeric(query_string) = False Then response.redirect "/index.asp"
    ElseIf qs_type="url" Then
      If instr(query_string,".")=0 Or left(query_string,1)="." Or right(query_string,1)="." Then response.redirect "/index.asp"
    End If
    call check_query_string(query_string,qs_type)
  End If
End Function

Function check_query_string(query_string,qs_type)
  If query_string<>"" Then
    If instr(query_string,"|")>0 Then response.redirect "/index.asp"
    If qs_type="content" Then
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|1=1|a=a|--|./|'='"
    ElseIf qs_type="tel" Then
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|`|~|!|$|^|*|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|./|//|'='"
    ElseIf qs_type="url" Then
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|'|""|<|>|,|1=1|a=a|--|./|//|'='"
    ElseIf qs_type="checkbox" Then
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|1=1|a=a|--|./|//|'='"
    ElseIf qs_type="money" Then
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|1=1|a=a|--|./|//|'='"
    Else
      KeySpecial="a href|admin|alert|and|application|cache|cast|config|cookie|char|eval|exists|hidden|http|iframe|is_|manage|password|script|select|session|url|user|`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|./|//|'='"
    End If
    Ary_KeySpecial=Split(KeySpecial,"|")
    For I = 0 To Ubound(Ary_KeySpecial)
      If instr(query_string,Ary_KeySpecial(I))>0 Then response.redirect "/index.asp"
    Next
    
    If qs_type="content" Then
      KeyWord=".asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|xtype|%2527|address.tst"
    ElseIf qs_type="tel" Then
      KeyWord="`|~|!|$|^|*|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|http|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|www|xtype|%2527|address.tst"
    ElseIf qs_type="url" Then
      KeyWord="`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|'|""|<|>|,|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|xtype|%2527|address.tst"
    ElseIf qs_type="checkbox" Then
      KeyWord="`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|http|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|www|xtype|%2527|address.tst"
    ElseIf qs_type="money" Then
      KeyWord="`|~|!|#|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|http|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|www|xtype|%2527|address.tst"
    Else
      KeyWord="`|~|!|#|$|^|*|(|)|+|[|]|{|}|\|;|:|'|""|<|>|,|1=1|a=a|--|;|./|.asp|.bak|.cfm|.css|.dos|.htm|.inc|.ini|.js|.php|.txt|/*|*/|.@|@.|@@|@C|@T|@P|a href|admin|acunetix|alter|and|application|begin|cache|cast|config|char|chr|cookie|count|create|css|cursor|deallocate|declare|delete|dir|drop|echo|end|eval|exec|exists|fetch|from|hidden|http|iframe|insert|into|is_|join|js|kill|left|manage|master|mid|nchar|ntext|nvarchar|open|password|right|script|set|select|session|src|sys|sysobjects|syscolumns|table|text|truncate|url|user|update|varchar|where|while|www|xtype|%2527|address.tst"
    End If
    Ary_KeyWord=Split(KeyWord,"|")
    For I = 0 To Ubound(Ary_KeyWord)
      If instr(query_string,Ary_KeyWord(I))>0 Then response.redirect "/index.asp"
    Next
    
    If qs_type="content" Then
      KeyHex="7e|5e|2a|2b|5b|5d|7b|7d|5c|3b|3a|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
    ElseIf qs_type="tel" Then
      KeyHex="`|7e|!|$|5e|2a|2b|5b|5d|7b|7d|5c|3b|3a|'|""|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
    ElseIf qs_type="url" Then
      KeyHex="`|7e|!|#|$|5e|2a|(|)|2b|5b|5d|7b|7d|5c|3b|3a|'|""|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
    ElseIf qs_type="money" Then
      KeyHex="`|7e|!|#|5e|2a|(|)|2b|5b|5d|7b|7d|5c|3b|3a|'|""|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
    Else
      KeyHex="`|7e|!|#|$|5e|2a|(|)|2b|5b|5d|7b|7d|5c|3b|3a|'|""|3c|3e|2c|313d31|613d61|2d2d|3b|2e2f|2e617370|2e62616b|2e63666d|2e637373|2e646f73|2e68746d|2e696e63|2e696e69|2e6a73|2e706870|2e747874|2f2a|2a2f|2e40|402e|4040|4043|4054|4050|612068726566|61646d696e|6163756e65746978|616c746572|616e64|6170706c69636174696f6e|626567696e|6361636865|63617374|636f6e666967|63686172|636872|636f6f6b6965|636f756e74|637265617465|637373|637572736f72|6465616c6c6f63617465|6465636c617265|64656c657465|646972|64726f70|6563686f|656e64|6576616c|65786563|657869737473|6665746368|66726f6d|68696464656e|68747470|696672616d65|696e73657274|696e746f|69735f|6a6f696e|6a73|6b696c6c|6c656674|6d616e616765|6d6173746572|6d6964|6e63686172|6e74657874|6e76617263686172|6f70656e|70617373776f7264|7269676874|736372697074|736574|73656c656374|73657373696f6e|737263|737973|7379736f626a65637473|737973636f6c756d6e73|7461626c65|74657874|7472756e63617465|75726c|75736572|757064617465|76617263686172|7768657265|7768696c65|777777|7874797065|2532353237|616464726573732e747374"
    End If
    Ary_KeyHex=Split(KeyHex,"|")
    For I = 0 To Ubound(Ary_KeyHex)
      If instr(query_string,Cstr(Ary_KeyHex(I)))>0 Then response.redirect "/index.asp"
    Next
    
    KeyBase="YA==|~|fg==|Iw==|JA==|Xg==|Kg==|KA==|KQ==|Kw==|Ww==|XQ==|ew==|fQ==|XA==|Ow==|Og==|Jw==|IiI=|PA==|Pg==|LA==|MT0x|YT1h|LS0=|Ow==|Li8=|LmFzcA==|LmJhaw==|LmNmbQ==|LmNzcw==|LmRvcw==|Lmh0bQ==|LmluYw==|LmluaQ==|Lmpz|LnBocA==|LnR4dA==|Lyo=|Ki8=|LkA=|QC4=|QEA=|QEM=|QFQ=|QFA=|YSBocmVm|YWRtaW4=|YWN1bmV0aXg=|YWx0ZXI=|YW5k|YXBwbGljYXRpb24=|YmVnaW4=|Y2FjaGU=|Y2FzdA==|Y29uZmln|Y2hhcg==|Y2hy|Y29va2ll|Y291bnQ=|Y3JlYXRl|Y3Nz|Y3Vyc29y|ZGVhbGxvY2F0ZQ==|ZGVjbGFyZQ==|ZGVsZXRl|ZGly|ZHJvcA==|ZWNobw==|ZW5k|ZXZhbA==|ZXhlYw==|ZXhpc3Rz|ZmV0Y2g=|ZnJvbQ==|aGlkZGVu|aHR0cA==|aWZyYW1l|aW5zZXJ0|aW50bw==|aXNf|am9pbg==|anM=|a2lsbA==|bGVmdA==|bWFuYWdl|bWFzdGVy|bWlk|bmNoYXI=|bnRleHQ=|bnZhcmNoYXI=|b3Blbg==|cGFzc3dvcmQ=|cmlnaHQ=|c2NyaXB0|c2V0|c2VsZWN0|c2Vzc2lvbg==|c3Jj|c3lz|c3lzb2JqZWN0cw==|c3lzY29sdW1ucw==|dGFibGU=|dGV4dA==|dHJ1bmNhdGU=|dXJs|dXNlcg==|dXBkYXRl|dmFyY2hhcg==|d2hlcmU=|d2hpbGU=|d3d3|eHR5cGU=|JTI1Mjc=|YWRkcmVzcy50c3Q="
    Ary_KeyBase=Split(KeyBase,"|")
    For I = 0 To Ubound(Ary_KeyBase)
      If instr(query_string,Ary_KeyBase(I))>0 Then response.redirect "/index.asp"
    Next
    
    KeyScript="011100110110001101110010011010010111000001110100|736372697074|c2NyaXB0|11599114105112116"
    Ary_KeyScript=Split(KeyScript,"|")
    For I = 0 To Ubound(Ary_KeyScript)
      If instr(query_string,Cstr(Ary_KeyScript(I)))>0 Then response.redirect "/index.asp"
    Next

    KeyAlert="0110000101101100011001010111001001110100|616c657274|YWxlcnQ=|97108101114116"
    Ary_KeyAlert=Split(KeyAlert,"|")
    For I = 0 To Ubound(Ary_KeyAlert)
      If instr(query_string,Cstr(Ary_KeyAlert(I)))>0 Then response.redirect "/index.asp"
    Next

    KeyIframe="011010010110011001110010011000010110110101100101|696672616d65|aWZyYW1l|10510211497109101"
    Ary_KeyIframe=Split(KeyIframe,"|")
    For I = 0 To Ubound(Ary_KeyIframe)
      If instr(query_string,Cstr(Ary_KeyIframe(I)))>0 Then response.redirect "/index.asp"
    Next

    KeyOther="sample@email.tst|%27|JyI=|acu75e3fa3f006b30025eb226cec0acecf6|PFNjUmlQdD5hbGVydCgneHNzLXRlc3QnKTs8L1NjUmlQdD4=|IiBvbm1vdXNlb3Zlcj0iYWxlcnQoJ3hzcy10ZXN0Jyk=|dir|..|boot.ini|268435455|-268435455|0x7fffffff|0x80000000|0x3fffffff|0xffffffff|/some_inexistent_file_with_long_name|www.haesung.hs.kr"
    Ary_KeyOther=Split(KeyOther,"|")
    For I = 0 To Ubound(Ary_KeyOther)
      If instr(query_string,Cstr(Ary_KeyOther(I)))>0 Then response.redirect "/index.asp"
    Next
  End If
End Function
%>