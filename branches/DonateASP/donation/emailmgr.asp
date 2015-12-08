<!--#include file="../include/dbfunctionJ.asp"-->
<%Prog_Id="emailmgr"%>
<!--#include file="../include/head.asp"-->
<body OnLoad='Window_OnLoad()'>
  <p><div align="center"><center>
    <form name="form" method="POST" action="news_list.asp?&code_id=<%=request("code_id")%>" target="detail">
      <input type="hidden" name="action">
      <div align="center"><center>
        <table border="1" width="830" cellspacing="0">
          <tr>
            <td>   
              <table border="0" width="100%" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
                <tr>
	                <td width="100%" class="font11" height="25" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)"><b><font color="#FFFFFF">&nbsp;&nbsp;<%=Prog_Desc%></td></font></b></td>
                </tr>
                <tr>
                  <td align="center">
                    關鍵字<span lang="en-us"> :</span>
                    <input type="text" name="keyword" size="20" class="font9" onKeypress='if(event.keyCode==13){JavaScript:Query_OnClick();}'>&nbsp;&nbsp;&nbsp;&nbsp;
                    <%If request("code_id")="" Then%>
                    郵件類別<span lang="en-us"> :</span>
                    <%
                      SQL="Select EmailMgr_Type=CodeDesc From CASECODE Where CodeType='EmailMgrType' Order By Seq"
                      FName="EmailMgr_Type"
                      Listfield="EmailMgr_Type"
                      menusize="1"
                      BoundColumn=""
                      call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                    %>&nbsp;&nbsp;&nbsp;&nbsp;
                    <%End If%>
                    <input type="button" value=" 查 詢 " name="query" class="cbutton" onClick="Query_OnClick()">
                  </td>
                </tr>
                <tr>
                  <td width="100%"><iframe name="detail" src="emailmgr_list.asp?action=query" height="465" width="100%" frameborder="0" scrolling="auto"></iframe></td>
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
function Window_OnLoad(){
  document.form.keyword.focus();
}
function Query_OnClick(){
  <%call SubmitJ("query")%>
}
--></script>