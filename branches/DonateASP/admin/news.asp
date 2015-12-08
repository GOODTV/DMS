<!--#include file="../include/dbfunctionJ.asp"-->
<%
News_Type="訊息內容"
If request("code_id")<>"" Then
  SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' And Seq='"&request("code_id")&"' " 
  Set RS = Server.CreateObject("ADODB.RecordSet")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then News_Type=RS("News_Type")
End If  
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title><%=News_Type%>管理</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" method="POST" action="news_list.asp?&code_id=<%=request("code_id")%>" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table border="1" width="830" cellspacing="0">
          <tr>
            <td>   
              <table border="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
	                <td width="100%" class="font11" height="25" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><b><font color="#FFFFFF">&nbsp;&nbsp;<%=News_Type%>管理</font></b></td>
                </tr>
                <tr>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="keyword" size="20" class="font9">&nbsp;&nbsp;&nbsp;&nbsp;
                    <%If request("code_id")="" Then%>
                    訊息類別<span lang="en-us"> :</span>
                    <%
                      SQL="Select News_Type=CodeDesc From CASECODE Where CodeType='NewsType' Order By Seq"
                      FName="News_Type"
                      Listfield="News_Type"
                      menusize="1"
                      BoundColumn=""
                      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                    %>&nbsp;&nbsp;&nbsp;&nbsp;
                    <%End If%>
                    <input type="button" value=" 查 詢 " name="query" class="cbutton" onClick="Query_OnClick()">&nbsp;&nbsp;
                    <input type="button" value="歷史資料查詢" name="history" class="cbutton" onClick="History_OnClick()">
                  </td>
                </tr>
                <tr>
                  <td width="100%"><iframe name="detail" src="news_list.asp?action=query&code_id=<%=request("code_id")%>" height="465" width="100%" frameborder="0" scrolling="auto"></iframe></td>
                </tr>
              </table>
            </td>
          </tr>  
        </table>
      </center></div>
    </form>
  </center></div></p>
  <%Message()%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->
<script language="JavaScript"><!--
function Query_OnClick(){
  <%call SubmitJ("query")%>
}
function History_OnClick(){
  <%call SubmitJ("history")%>
}
--></script>