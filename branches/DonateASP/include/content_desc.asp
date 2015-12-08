                    <%
                      If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then 
                        Server_Line="localhost/"&Split(Request.ServerVariables("URL"),"/")(1)
                      ElseIf Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
                        server_line=replace("http://"&Cstr(Request.ServerVariables("SERVER_NAME"))&Request.ServerVariables("URL"),"epaper_preview.asp","")
                      Else
                        Server_Line=Cstr(Request.ServerVariables("SERVER_NAME"))
                      End If
                      If ImgPosition="bottom" Then Response.Write Desc
                      
                      '附加圖檔
                      If ImageShow_Type<>"album" Then
                        SQL="Select * From UPLOAD Where Ap_Name='"&Ap_Name&"' And Attach_Type NOT IN ('doc','img') And Object_ID='"&Object_ID&"'"
                      Else
                        SQL="Select * From UPLOAD Where Ap_Name='"&Ap_Name&"' And Attach_Type NOT IN ('doc','img','image') And Object_ID='"&Object_ID&"'"
                      End If
                      Call QuerySQL(SQL,RS11)
                      If Not RS11.EOF Then
                    %>
                    <table cellSpacing="1" cellPadding="6" width="100%" align="<%=ImgPosition%>" border="0">
                      <%While Not RS11.EOF%>
                        <tr> 
                          <td align=right width="12%" height=22></td>
                          <td valign=center width="88%" height=22 colspan="4"> 
                            <%If RS11("attach_type")="doc" Then%>
                              <img border="0" src="http://<%=Server_Line%>/images/save.GIF" align="absmiddle"><a href="http://<%=Server_Line%>/upload/<%=RS11("Upload_fileURL")%>" target="_blank"><%=RS11("upload_fileURL")%></a>
                            <%ElseIf RS11("attach_type")="image" Then%>
                              <%If ImageShow_Type<>"album" Then%>
                              <img border="0" src="http://<%=Server_Line%>/upload/<%=RS11("upload_fileurl")%>" alt="<%=RS11("Upload_FileName")%>">
                              <%End If%>
                            <%ElseIf RS11("attach_type")="flash" Then%>
                              <OBJECT classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000"
				                        codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0"WIDTH=611 HEIGHT=82>
				                        <PARAM NAME=movie VALUE="http://<%=Server_Line%>/upload/<%=RS11("upload_fileurl")%>">
				                        <PARAM NAME=quality VALUE=high>
				                        <PARAM NAME=bgcolor VALUE=#FFFFFF>
				                        <EMBED src="http://<%=Server_Line%>/upload/<%=RS11("upload_fileurl")%>" quality=high bgcolor=#FFFFFF  WIDTH=611 HEIGHT=82 TYPE="application/x-shockwave-flash" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash"></EMBED>
			                        </OBJECT>
                            <%ElseIf RS11("attach_type")="av" or RS11("attach_type")="WMV" Then
                               If RS11("attach_type")="WMV" Then
                                 data_url=RS11("upload_fileurl")
                               Else
                                 data_url="http://"&Server_Line&"/upload/"&RS11("upload_fileurl")
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
                      If ImgPosition<>"bottom" Then Response.write Desc
                      
                      '檔案下載
                      SQL="Select * From UPLOAD Where Ap_Name='"&Ap_Name&"' And Attach_Type IN ('doc') And Object_ID='"&Object_ID&"'"
                      Call QuerySQL(SQL,RS12)
                      If Not RS12.EOF Then
                    %>
	                  <table border="0" width="100%" cellspacing="0" cellpadding="0">
	                  <%
	                    While Not RS12.EOF
	                      If RS12("Upload_FileName")<>"" Then
	                        Upload_FileName=RS12("Upload_FileName")
	                      Else
	                        Upload_FileName=RS12("Upload_FileURL")
	                      End If
	                      icon="blank.gif"
	                      If Right(RS12("Upload_FileURL"),3)="pdf" Then
	                        icon="icon_PDF.gif"
	                      ElseIf Right(RS12("Upload_FileURL"),3)="doc" Then
	                        icon="icon_DOC.gif"
	                      ElseIf Right(RS12("Upload_FileURL"),3)="xls" Then  
	                        icon="icon_XLS.gif"
	                      End If
	                  %>
	                  <tr>
                      <td class="font10" valign="top"><img border="0" src="http://<%=Server_Line%>/images/<%=icon%>" width="16" height="16" />&nbsp;&nbsp;<a href="http://<%=Server_Line%>/upload/<%=RS12("Upload_FileURL")%>" class="linka" target="_blank"><%=Upload_FileName%></a></td>
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
                        SQL13="Select * From UPLOAD Where Ap_Name='"&Ap_Name&"' And Attach_Type IN ('image') And Object_ID='"&Object_ID&"'"
                        Call QuerySQL(SQL13,RS13)
                        If Not RS13.EOF Then
                          bg_color="FFFFFF"
                          DataFile="public/image/album_bg.gif"
                          Set objFS = Server.CreateObject("Scripting.FileSystemObject")
                          If objFS.FileExists(Server.MapPath(DataFile)) And UrlName<>"epaper_preview.asp" Then 
                            Response.Write "<table border=""0"" valign=""top"" align=""center"" cellpadding=""0"" cellspacing=""0"" background="""&DataFile&""">"& vbCRLF
                            Response.Write "  <tr>"& vbCRLF
                            Response.Write "    <td valign=""top"" align=""center"" background="""&DataFile&""">"& vbCRLF
                            Response.Write "      <iframe name=""album_photo"" src=""http://"&Server_Line&"/include/album_photo.asp?bg_image="&DataFile&"&ap_name="&Ap_Name&"&ser_no="&Object_ID&""" width=""550"" height=""485"" frameborder=""0"" scrolling=""no""></iframe>"& vbCRLF
					                  Response.Write "    </td>"& vbCRLF
                            Response.Write "  </tr>"& vbCRLF
                            Response.Write "</table>"& vbCRLF
                          Else
                            Response.Write "<table border=""0"" valign=""top"" align=""center"" cellpadding=""0"" cellspacing=""0"" bgcolor=""#"&bg_color&""">"& vbCRLF
                            Response.Write "  <tr>"& vbCRLF
                            Response.Write "    <td valign=""top"" align=""center"">"& vbCRLF
                            Response.Write "      <iframe name=""album_photo"" src=""http://"&Server_Line&"/include/album_photo.asp?bg_color="&bg_color&"&ap_name="&Ap_Name&"&ser_no="&Object_ID&""" width=""550"" height=""485"" frameborder=""0"" scrolling=""no""></iframe>"& vbCRLF
					                  Response.Write "    </td>"& vbCRLF
                            Response.Write "  </tr>"& vbCRLF
                            Response.Write "</table>"& vbCRLF
                          End If
                          Set objFS = Nothing
                        End If
                        RS13.Close
                        Set RS13=Nothing
                      End If
                    %>