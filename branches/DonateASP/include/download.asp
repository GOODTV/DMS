	        <%
            If Not RS12.EOF Then
          %>
	        <br><table border="0" width="100%" cellspacing="0" cellpadding="0">
	        <%
	          While Not RS12.EOF
	            If RS12("Upload_FileName")<>"" Then
	              Upload_FileName=RS12("Upload_FileName")
	            Else
	              Upload_FileName=RS12("Upload_FileURL")
	            End If
	        %>
	          <tr>
              <td class="font10"><img border="0" src="http://<%=Server_Line%>/image/icon/37.gif" width="9" height="12" />&nbsp;&nbsp;<a href="http://<%=Server_Line%>/upload/<%=RS12("Upload_FileURL")%>" class="menu6" target="_blank"><%=Upload_FileName%></a></td>
            </tr>
	        <%
	            RS12.MoveNext
            Wend
            RS12.close
		        Set RS12=nothing
          %>
	        </table>
	        <%
	          End If
            If ImageShow_Type="album" Then
          %>
          <br><table width="100%" border="0" valign="top" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="100%" valign="top" align="center">
                <iframe name="album_photo" src="http://<%=server_line%>/include/album_photo.asp?ap_name=<%=Ap_Name%>&ser_no=<%=Object_ID%>" height="0" width="100%" frameborder="0" scrolling="no"></iframe>
					    </td>
            </tr>
          </table>
          <%
            End If
          %>