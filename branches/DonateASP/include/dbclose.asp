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

If UrlName<>"rss.asp" And UrlName<>"pledge_transfer.asp" And UrlName<>"epaper_preview.asp" And UrlName<>"epaper_preview2.asp" Then
  Response.Write "<script language=""JavaScript""><!--"&vbcrlf
  Response.Write "  function MM_openBrWindow(theURL,winName,features) {"&vbcrlf
  Response.Write "    window.open(theURL,winName,features);"&vbcrlf
  Response.Write "  }"&vbcrlf
  Response.Write "function leapyear(year){"&vbcrlf
  Response.Write " if(parseInt(year)%4==0){"&vbcrlf
  Response.Write "  if(parseInt(year)%100==0){"&vbcrlf
  Response.Write "   if(parseInt(year)%400==0) return true;"&vbcrlf
  Response.Write "   	return false;	"&vbcrlf
  Response.Write "   }else return true;"&vbcrlf
  Response.Write " }else return false;"&vbcrlf
  Response.Write "}"&vbcrlf
  Response.Write "--></Script>"&vbcrlf
End If
%>