<script language="VBScript"><!--

sub ConfirmUpdate()
  DIM answer
  answer = window.confirm ( "您是否確定要修改 ?")
  if answer = true then
    form.action.value="update"
    form.submit
  end if
end sub

sub delete_OnClick
  DIM answer
  answer = window.confirm ( "您是否確定要刪除 ?")
  if answer = true then
    form.action.value="delete"
    form.submit
  end if
end sub

sub cancel_OnClick
  history.back
end sub

sub GoPage_OnChange      
  newPage=form1.GoPage.value     
  document.location.href="<%=ProgID%>.asp?nowPage=" & newPage & "&SQL=<%=SQL%>"
end sub

Function blen(xs)
  Dim i
  xl = len(xs)
  for i=1 to len(xs)
    if asc(mid(xs,i,1))<0 then xl = xl + 1
  next
  blen = xl
End function
--></script>