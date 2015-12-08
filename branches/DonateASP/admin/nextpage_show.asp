                      <%
                        SQL2="Select * From NEXTPAGE Where Page_type='"&Item&"' And ser_no='"&request("ser_no")&"'"
                        Call QuerySQL(SQL2,RS2)
                        If Not RS2.EOF Then
                      %>
                      <tr> 
                        <td align=right height=22><font color="#000080">分頁內容：</font></td>
                        <td height=22 colspan="7">
                        <%While Not RS2.EOF%>
                        &nbsp;<a href="#" 
                        OnClick="MM_openBrWindow('nextpage_edit.asp?code_id=<%=request("code_id")%>&item=<%=Item%>&page_id=<%=RS2("Page_Id")%>','nextpage','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=no,top=220,left=170,width=640,height=320')">第&nbsp;<%=RS2("Page_No")%>&nbsp;頁</a>&nbsp;&nbsp;
                        <%
                            RS2.MoveNext
                          Wend
                          RS2.Close
                          Set RS2=Nothing
                        %></td></tr><%End If%>