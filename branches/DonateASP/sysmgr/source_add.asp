<!--#include file="../include/dbfunction.asp"-->
<%
  SQL="Select MaxFileSize,MaxUploadSize From dept Where dept_id='"&session("dept_id")&"'" 
  call QuerySQL(SQL,RS)
  If Not RS.EOF Then
    If RS("MaxFileSize")<>"" Then
      MaxFileSize=RS("MaxUploadSize")
    Else
      MaxFileSize=10
    End If   
  Else
    MaxFileSize=10
  End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <link REL="stylesheet" type="text/css" HREF="../include/dms.css">
  <title>檔案上傳管理 【新增資料】</title>
</head>
<body class=gray>
  <p><div align="center"><center>
    <form name="form" action="source_add_upload.asp?MaxFileSize=<%=MaxFileSize%>" method="post" enctype="multipart/form-data">
      <table border="0" width="830" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
        <tr>
	  <td width="100%" class="font11" style="background-color: rgb(55,111,135); color: rgb(255,255,255); filter:Alpha(Opacity=100, style=1)" height="25" colspan="2"><font color="#FFFFFF"><b>&nbsp;&nbsp;檔案上傳 </b> <font size="2">【新增資料】</font></font></td>
        </tr>
        <tr>
	  <td width="100%" class="font11" valign="top">
            <div align="center"><center>
              <table border="0" width="100%" cellspacing="0">
   		<tr>
   		  <td>
                    <table width="100%" border="1" cellspacing="0" style="border-collapse: collapse" cellpadding="2" bordercolor="#E1DFCE">
                      <tr>
                        <td noWrap align="right" width="10%"><font color="#000080">上傳路徑：</font></td>
                        <td width="90%">
                        <%
                          SQL="Select Source_RootType=MainName From CODEMAIN Where CodeId=(Select CodeId From CODEFILE Where CodeType='RootType' And CodeCategory='類別') Order By MainSeq"
                          FName="Source_RootType"
                          Listfield="Source_RootType"
                          BoundColumn=""
                          menusize="1"
                          call OptionList (SQL,FName,Listfield,BoundColumn,menusize)
                        %>&nbsp;&nbsp;<font color="FF0000">(&nbsp;不選擇則為根目錄&nbsp;)</font>                     
                        </td>
                      </tr>
                      <%For I= 1 To 10%>                   
                      <tr> 
                        <td noWrap align=right><font color="#000080">上傳檔案<%=I%>：</font></td>
                        <td><input type="file" class="font9" size="60" name="Upload_FileURL<%=I%>" maxlength="120">
                          <%If I=1 Then%><font color="FF0000">&nbsp;&nbsp;檔案大小不可為&nbsp;0<%End If%>
                          <%If I=2 Then%><font color="FF0000">&nbsp;&nbsp;單一檔案大小限制&nbsp;<%=MaxFileSize%>&nbsp;M<%End If%>
                        </td>
                      </tr>
                      <%Next%>                                        
                      <tr>
                        <td width="100%" bgcolor=#FFFCE1 class="font9" colspan="2"><%SaveButton%></td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </center></div>
          </td>
        </tr>
      </table>
    </form>
  </center></div>
  <%message%>
</body>
<!--#include file="../include/dbclose.asp"-->
<!--#include file="../include/vbfunction.inc"-->
<script language="VBScript"><!--

sub save_OnClick           
  form.submit   
end sub

sub cancel_OnClick
  history.back
end sub

sub query_OnClick
  location.href="source.asp"
end sub  

--></script>