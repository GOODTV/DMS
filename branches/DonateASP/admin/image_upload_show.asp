                      <%If Item<>"periodical" And Item<>"magazine" And Item<>"emailmgr" Then%>
                      <tr> 
                        <td align="right"><font color="#000080">圖檔排列：</font></td>
                        <td colspan="7"> 
			                    <%If Item="epaper_basic" Then%>
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Auto" value="barrier" <%If RS("epaper_ImageShow_Type")="barrier" Then%>checked<%End If%> >無障礙相簿
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Album" value="album" <%If RS("epaper_ImageShow_Type")="album" Then%>checked<%End If%> >網軟相簿
			                    <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Album" value="flickr" <%If RS("epaper_ImageShow_Type")="flickr" Then%>checked<%End If%> >Flickr相簿<br><font color="#FF0000">(&nbsp;縮圖高度請設定：80pixal&nbsp;)</font>
			                    <%Else%>
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Auto" value="barrier" <%If RS(""&Item&"_ImageShow_Type")="barrier" Then%>checked<%End If%> >無障礙相簿	
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Album" value="album" <%If RS(""&Item&"_ImageShow_Type")="album" Then%>checked<%End If%> >網軟相簿
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Auto" value="flickr" <%If RS(""&Item&"_ImageShow_Type")="flickr" Then%>checked<%End If%> >Flickr相簿<br><font color="#FF0000">(&nbsp;縮圖高度請設定：80pixal&nbsp;)</font>
                          <%End If%>
                        </td>
                      </tr>
                      <%End If%>
                      <tr> 
                        <td align="right"><font color="#000080">附加檔案：</font></td>
                        <td colspan="7">
                        	<%If Item="periodical" Then%>
                        	<input type="button" value="文件" name="image_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                        	<%ElseIf Item="magazine" Then%>
                        	<input type="button" value="PDF檔案" name="image_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_word_upload.asp?dept_id=<%=session("dept_id")%>&item=magazine&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                        	<%ElseIf Item="album" Then%> 
                          <input type="button" value="單一圖檔" name="image_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                          <input type="button" value="多筆圖檔" name="image10_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image10_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=84,left=170,width=500,height=600')">
                          <input type="button" value="Flickr相簿" name="image_flickr_link" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_flickr_link.asp?item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                        	<%ElseIf Item="emailmgr" Then%>
                        	<input type="button" value="新增文件" name="image_word_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_word_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                        	<%Else%>
                        	<%If Item="activity" Then%><input type="button" value="報名簡章" name="image_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_word_upload.asp?dept_id=<%=session("dept_id")%>&item=signup&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')"><%End If%>
                          <input type="button" value="新增文件" name="image_word_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_word_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                          <input type="button" value="圖檔&nbsp;/&nbsp;動畫" name="image_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                          <input type="button" value="多筆圖檔" name="image10_upload_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image10_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&img_width=&img_height=80&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=84,left=170,width=500,height=600')">
                          <input type="button" value="影音檔案" name="image_av_add" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_av_upload.asp?dept_id=<%=session("dept_id")%>&item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                          <input type="button" value="影音連結" name="image_av_link" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_av_link.asp?item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')">
                          <input type="button" value="Flickr相簿" name="image_flickr_link" class="cbutton" style="cursor:hand" OnClick="MM_openBrWindow('image_flickr_link.asp?item=<%=Item%>&ser_no=<%=RS("ser_no")%>&subject=<%=Subject%>&code_id=<%=request("code_id")%>','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=500,height=160')"> 	 
                          <%End If%></td></tr><%
                        If Item="activity" Then
                          SQL1="Select * From UPLOAD Where Ap_Name In ('activity','signup') And Attach_Type<>'img' And Object_ID='"&request("ser_no")&"' "
                        Else
                          SQL1="Select * From UPLOAD Where Ap_Name='"&Item&"' And Attach_Type<>'img' And Object_ID='"&request("ser_no")&"' "
                        End If
                        call QuerySQL(SQL1,RS1)
                        If Not RS.EOF Then
                          FieldsCount = RS1.Fields.Count-1
                          Response.Write "<tr>"
                          Response.Write "  <td valign=""center"" colspan=""8"">"
                          Response.Write "    <table border=1 cellspacing='0' cellpadding='2' style='border-collapse: collapse;BACKGROUND-COLOR: #ffffff' bordercolor='#C0C0C0' width='100%' align='center'>"
                          Response.Write "      <tr>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>檔案類型</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>檔案名稱</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>檔案大小(K)</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>檔案說明</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>瀏覽</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>編輯</span></font></td>"
                          Response.Write "        <td bgcolor='#FFE1AF' nowrap><font color='#111111'><span style='font-size: 9pt; font-family: 新細明體'>刪除</span></font></td>"
                          Response.Write "      </tr>"
                          While Not RS1.EOF
                            Response.Write "    <tr "&showhand&" bgcolor='#FFFFFF' onmouseover='this.bgColor=""#E2F1FF""' onmouseout='this.bgColor=""#FFFFFF""'>"
                            If RS1("Attach_Type")="doc" Then
                              If RS1("Ap_Name")="signup" Then
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>報名簡章</span></td>"  
                              Else
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>文件</span></td>"  
                              End If
                            ElseIf RS1("Attach_Type")="image" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>圖檔</span></td>"
                            ElseIf RS1("Attach_Type")="flash" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>動畫</span></td>"
                            ElseIf RS1("Attach_Type")="WMV" Or RS1("attach_type")="flv" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>影音檔案</span></td>"
                            ElseIf RS1("Attach_Type")="av" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>影音檔案</span></td>"
                            ElseIf RS1("attach_type")="YouTube" Or RS1("attach_type")="無名小站" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>影音連結</span></td>"
                            ElseIf RS1("attach_type")="Flickr" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>相簿連結</span></td>"  
                            End If                          
                            If RS1("Upload_FileURL_Old")<>"" Then
                              If Left(RS1("Upload_FileURL_Old"),1)="{" Then
                                Upload_FileURL_Old=""
                                Ary_FileURL_Old=Split(RS1("Upload_FileURL_Old"),"_")
                                For I = 1 to UBound(Ary_FileURL_Old)
                                  If Upload_FileURL_Old="" Then
                                    Upload_FileURL_Old=Ary_FileURL_Old(I)
                                  Else
                                    Upload_FileURL_Old=Upload_FileURL_Old&"_"&Ary_FileURL_Old(I)
                                  End If
                                Next
                              Else
                                Upload_FileURL_Old=RS1("Upload_FileURL_Old")
                              End If
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&Upload_FileURL_Old&"</span></td>"
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Upload_FileSize_Old")&"</span></td>"
                            Else
                              If RS1("attach_type")="YouTube" Or RS1("attach_type")="無名小站" Or RS1("attach_type")="Flickr" Then
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("attach_type")&"</span></td>"
                              Else
                                If Left(RS1("Upload_FileURL"),1)="{" Then
                                  Upload_FileURL=""
                                  Ary_FileURL=Split(RS1("Upload_FileURL"),"_")
                                  For I = 1 to UBound(Ary_FileURL)
                                    If Upload_FileURL="" Then
                                      Upload_FileURL=Ary_FileURL(I)
                                    Else
                                      Upload_FileURL=Upload_FileURL&"_"&Ary_FileURL(I)
                                    End If
                                  Next
                                Else
                                  Upload_FileURL=RS1("Upload_FileURL")
                                End If
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&Upload_FileURL&"</span></td>"
                              End If
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Upload_FileSize")&"</span></td>"
                            End If
                            Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'>"&RS1("Upload_FileName")&"</span></td>"
                            If RS1("Attach_Type")="doc" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('../upload/"&RS1("Upload_FileURL")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=800,height=600')"">瀏覽</a></span></td>"
                            ElseIf RS1("Attach_Type")="image" Then
                              If RS1("Upload_FileURL_Old")<>"" Then
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('../upload/"&RS1("Upload_FileURL_Old")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=800,height=600')"">瀏覽</a></span></td>"
                              Else
                                Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('../upload/"&RS1("Upload_FileURL")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=800,height=600')"">瀏覽</a></span></td>"
                              End If
                            ElseIf RS1("Attach_Type")="flash" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_flv_show.asp?filename=../upload/"&RS1("Upload_FileURL")&"&filetype="&RS1("attach_type")&"&xwidth=800','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=800,height=600')"">瀏覽</a></span></td>"
                            ElseIf RS1("Attach_Type")="WMV" Or RS1("attach_type")="flv" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_flv_show.asp?filename=../upload/"&RS1("Upload_FileURL")&"&filetype="&RS1("attach_type")&"&xwidth=240','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=260,height=200')"">瀏覽</a></span></td>"
                            ElseIf RS1("Attach_Type")="av" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_flv_show.asp?filename=../upload/"&RS1("Upload_FileURL")&"&filetype="&RS1("attach_type")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=340,height=310')"">瀏覽</a></span></td>"
                            ElseIf RS1("attach_type")="YouTube" Or RS1("attach_type")="無名小站" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_flv_show.asp?ser_no="&RS1("ser_no")&"&filetype="&RS1("attach_type")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=450,height=360')"">瀏覽</a></span></td>"
                            ElseIf RS1("attach_type")="Flickr" Then
                              Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_flv_show.asp?ser_no="&RS1("ser_no")&"&filetype="&RS1("attach_type")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=yes,resizable=yes,top=84,left=112,width=345,height=245')"">瀏覽</a></span></td>"  
                            End If
                            Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick=""MM_openBrWindow('image_filename_update.asp?item="&Item&"&ser_no="&RS1("ser_no")&"&subject="&Subject&"&code_id="&request("code_id")&"','upload','toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no,top=220,left=170,width=450,height=50')"">編輯</a></span></td>"
                            Response.Write "    <td><span style='font-size: 9pt; font-family: 新細明體'><a href=""#"" OnClick='JavaScript:if(confirm(""是否確定要刪除附加檔案 ?"")){window.location.href=""image_upload_delete.asp?item="&Item&"&ser_no="&RS1("ser_no")&"&code_id="&request("code_id")&""";}'>刪除</a></span></td>"
                            Response.Write "    </tr>" 
                            RS1.MoveNext
                          Wend
                          Response.Write "    </table>"	
                          Response.Write "  </td>"
                          Response.Write "<tr>"
                        End If
                        RS1.Close
                        Set RS1=Nothing
                      %>