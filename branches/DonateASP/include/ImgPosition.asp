                    <%
                      If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then 
                        Server_Line="localhost/mustard"
                      Else
                        Server_Line=Cstr(Request.ServerVariables("SERVER_NAME"))
                      End If
                      If ImgPosition="bottom" Then Response.Write content_desc
                      
                      If Not RS11.EOF Then
                    %>
                    <table cellSpacing="1" cellPadding="6" width="100%" align="<%=ImgPosition%>" border="0">
                      <%While Not RS11.EOF%>
                        <tr> 
                          <td align=right width="12%" height=22></td>
                          <td valign=center width="88%" height=22 colspan="4"> 
                            <%If RS11("attach_type")="doc" Then%>
                              <a href="http://<%=server_line%>/upload/<%=RS11("Upload_fileURL")%>" target="_blank"><img border="0" src="../images/save.GIF" align="absmiddle"><%=RS11("upload_fileURL")%></a>
                            <%ElseIf RS11("attach_type")="image" Then%>
                              <%If ImageShow_Type<>"album" Then%>
                              <img border="0" src="http://<%=server_line%>/upload/<%=RS11("upload_fileurl")%>" alt="<%=RS11("Upload_FileName")%>">
                              <%End If%>
                            <%ElseIf RS11("attach_type")="flash" Then%>
                              <OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
				                        codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"WIDTH=611 HEIGHT=82>
				                        <PARAM NAME=movie VALUE="http://<%=server_line%>/upload/<%=RS11("upload_fileurl")%>">
				                        <PARAM NAME=quality VALUE=high>
				                        <PARAM NAME=bgcolor VALUE=#FFFFFF>
				                        <EMBED src="http://<%=server_line%>/upload/<%=RS11("upload_fileurl")%>" quality=high bgcolor=#FFFFFF  WIDTH=611 HEIGHT=82 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
			                        </OBJECT>
                            <%ElseIf RS11("attach_type")="av" or RS11("attach_type")="WMV" Then
                               If RS11("attach_type")="WMV" Then
                                 data_url=RS11("upload_fileurl")
                               Else
                                 data_url="http://"&server_line&"/upload/"&RS11("upload_fileurl")
                               End If%>
			                        <object id='MediaPlayer' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT>
			                          <param name='AllowChangeDisplaySize' value='1'>
			                          <PARAM name='autoStart' value='false'>
			                          <param name='AutoSize' value='0'>
			                          <param name='AnimationAtStart' value='1'>
			                          <param name='ClickToPlay' value='1'>
			                          <param name='EnableContextMenu' value='0'>
			                          <param name='EnablePositionControls' value='1'>
			                          <param name='EnableFullScreenControls' value='1'>
			                          <param name='URL' value='<%=data_url%>'>
			                          <param name='ShowControls' value='1'>
			                          <param name='ShowAudioControls' value='1'>
			                          <param name='ShowDisplay' value='0'>
			                          <param name='ShowGotoBar' value='0'>
			                          <param name='ShowPositionControls' value='1'>
			                          <param name='ShowStatusBar' value='1'>
			                          <param name='ShowTracker' value='1'>
			                          <embed src='<%=data_url%>' 
			                            type='video/x-ms-wmv' 
			                            width='1400' height='1050' 
			                            autoStart='1' showControls='0'
			                            AutoSize='0'
			                            AnimationAtStart='1'
			                            ClickToPlay='1'
			                            EnableContextMenu='0'
			                            EnablePositionControls='1'
			                            EnableFullScreenControls='1'
			                            ShowControls='1'
			                            ShowAudioControls='1'
			                            ShowDisplay='0'
			                            ShowGotoBar='0'
			                            ShowPositionControls='1'
			                            ShowStatusBar='1'
			                            ShowTracker='1'>
			                          </embed>
			                        </object>
                            <%ElseIf  RS11("attach_type")="YouTube" Then
                                response.write RS11("upload_fileurl")
                      	      End If
                      	    %> 
                        </td>
                      </tr>
                      <%  
                          RS11.MoveNext
                        Wend
                        RS11.Close
		                    Set RS11=Nothing
		                  %>
                    </table>
                    <%
                      End If
                      If ImgPosition<>"bottom" Then Response.write content_desc
                    %>