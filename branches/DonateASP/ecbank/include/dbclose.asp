<%
If Isobject(cmd)=true Then
  Set cmd=Nothing
End If

If Isobject(RS)=true Then
  RS.Close
  Set RS=Nothing
End If

If Isobject(Conn)=true Then
  Conn.Close
  Set Conn=Nothing
End If
%>