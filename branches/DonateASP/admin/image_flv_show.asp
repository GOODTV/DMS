<!--#include file="../include/dbfunctionJ.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
	<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" /> 
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>影音/相簿播放</title>
</head>
<body topmargin="0" leftmargin="0" rightmargin="0" bottommargin="0">
<%
  If Request("FileType")="flash" Or Request("FileType")="WMV" Or Request("FileType")="flv" Then
    call ShowFlv(Request("FileName"),Request("FileType"),Request("xWidth"),Request("xHeight"))
  ElseIf Request("FileType")="av" Then
%>
	  <object id='MediaPlayer' classid='CLSID:6BF52A52-394A-11d3-B153-00C04F79FAA6' VIEWASTEXT>
		  <param name='AllowChangeDisplaySize' value='1'>
			<param name='autoStart' value='true'>
			<param name='AutoSize' value='0'>
			<param name='AnimationAtStart' value='1'>
			<param name='ClickToPlay' value='1'>
			<param name='EnableContextMenu' value='0'>
			<param name='EnablePositionControls' value='1'>
			<param name='EnableFullScreenControls' value='1'>
			<param name='URL' value='<%=Request("FileName")%>'>
			<param name='ShowControls' value='1'>
			<param name='ShowAudioControls' value='1'>
			<param name='ShowDisplay' value='0'>
			<param name='ShowGotoBar' value='0'>
			<param name='ShowPositionControls' value='1'>
			<param name='ShowStatusBar' value='1'>
			<param name='ShowTracker' value='1'>
			<embed src='<%=Request("FileName")%>' 
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
<%  
  ElseIf Request("FileType")="YouTube" Or Request("FileType")="無名小站" Or Request("FileType")="Flickr" Then
    SQL1="Select Upload_FileURL From UPLOAD Where Ser_No='"&Request("Ser_No")&"' "
    call QuerySQL(SQL1,RS1)
    If Not RS1.EOF Then
      Response.Write RS1("Upload_FileURL")
    End If
    RS1.Close
    Set RS1=Nothing
  End If  
%>
</body>
</html>
<!--#include file="../include/dbclose.asp"-->