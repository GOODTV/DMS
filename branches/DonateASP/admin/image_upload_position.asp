                      <tr> 
                        <td align="right"><font color="#000080">圖檔位置：</font></td>
                        <td colspan="7"> 
			                    <%If Item="epaper_basic" Then%>
			                    <input type="radio" name="epaper_ImgPosition" id="PosRight" value="right" <%If RS("epaper_ImgPosition")="right" Then%>checked<%End If%> >靠右 
                          <input type="radio" name="epaper_ImgPosition" id="PosLeft" value="left" <%If RS("epaper_ImgPosition")="left" Then%>checked<%End If%> >靠左 
                          <input type="radio" name="epaper_ImgPosition" id="PosTop" value="top" <%If RS("epaper_ImgPosition")="top" Then%>checked<%End If%> >上面 
                          <input type="radio" name="epaper_ImgPosition" id="PosBottom" value="bottom" <%If RS("epaper_ImgPosition")="bottom" Then%>checked<%End If%> >下面
                          &nbsp;&nbsp;&nbsp;&nbsp;
                          <font color="#000080">圖檔排列方式：</font>
                          <input type="radio" name="ePaper_ImageShow_Type" id="ImageShow_Type_Auto" value="auto" <%If RS("ePaper_ImageShow_Type")="auto" Then%>checked<%End If%> >自動
                          <input type="radio" name="ePaper_ImageShow_Type" id="ImageShow_Type_Album" value="album" <%If RS("ePaper_ImageShow_Type")="album" Then%>checked<%End If%> >相簿&nbsp;(&nbsp;縮圖高度請設定：80pixal&nbsp;) 	
			                    <%Else%>
			                    <input type="radio" name="<%=Item%>_ImgPosition" id="PosRight" value="right" <%If RS(""&Item&"_ImgPosition")="right" Then%>checked<%End If%> >靠右 
                          <input type="radio" name="<%=Item%>_ImgPosition" id="PosLeft" value="left" <%If RS(""&Item&"_ImgPosition")="left" Then%>checked<%End If%> >靠左 
                          <input type="radio" name="<%=Item%>_ImgPosition" id="PosTop" value="top" <%If RS(""&Item&"_ImgPosition")="top" Then%>checked<%End If%> >上面 
                          <input type="radio" name="<%=Item%>_ImgPosition" id="PosBottom" value="bottom" <%If RS(""&Item&"_ImgPosition")="bottom" Then%>checked<%End If%> >下面
                          &nbsp;&nbsp;&nbsp;&nbsp;
                          <font color="#000080">圖檔排列方式：</font>
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Auto" value="auto" <%If RS(""&Item&"_ImageShow_Type")="auto" Then%>checked<%End If%> >自動
                          <input type="radio" name="<%=Item%>_ImageShow_Type" id="ImageShow_Type_Album" value="album" <%If RS(""&Item&"_ImageShow_Type")="album" Then%>checked<%End If%> >相簿&nbsp;(&nbsp;縮圖高度請設定：80pixal&nbsp;) 
                          <%End If%></td></tr>