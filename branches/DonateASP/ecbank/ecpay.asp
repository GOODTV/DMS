<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="include/dbinclude_d.asp"-->
<%Session.Contents.Remove("get_od_sob")%>
<%
If Download_FileURL<>"" Then
  If Cstr(Request.ServerVariables("SERVER_NAME"))="localhost" Or Cstr(Request.ServerVariables("SERVER_NAME"))="127.0.0.1" Then
    download_url="定額捐款授權4合一表.doc"
  ElseIf Cstr(Request.ServerVariables("SERVER_NAME"))="web.npois.com.tw" Then
    download_url="https://web.npois.com.tw/donation/"&Split(Request.ServerVariables("URL"),"/")(2)&Download_FileURL
  Else
    download_url="http://donation.npois.com.tw/"&Split(Request.ServerVariables("URL"),"/")(1)&Download_FileURL
  End If
Else
  download_url="定額捐款授權4合一表.doc"
End If
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE8" />
  <link href="include/donate.css" rel="stylesheet" type="text/css">
  <script src="js/npois.js" type="text/javascript"></script>
  <script src="js/ecpay.js" type="text/javascript"></script>
  <!--[if lt IE 8]>  <script src="js/IE8.js"></script>  <![endif]-->
  <title><%=Comp_Name%></title>
</head>
<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false" onbeforecopy="return false" onselect="document.selection.empty()" oncopy="document.selection.empty()">
  <noscript><iframe src=*.html></iframe></noscript>	
  <div id="wrapper">
    <div id="header"><img src="image/banner.jpg" alt="<%=Comp_Name%>" /></div>
    <div id="top"><a href="ecpay.asp">首頁</a> / 線上捐款 / 捐款方式(交易方式)選擇</div>
    <div id="mid">
      <div id="mid-1">
        <table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
          <form name="form_donate" method="POST" action="ecpaya.asp">
						<input type="hidden" name="action">
            <%If ECBank_WebAtm="N" And ECBank_Vacc="N" And ECBank_Barcode="N" And ECBank_Ibon="N" And ECBank_FamiPort="N" And ECBank_LiftET="N" And ECBank_OKGo="N" Then%>
            <input type="hidden" name="Donate_Type" value="creditcard">
            <%Else%>
            <tr>
              <td class="contents">
              	<table width="100%" border="0" align="left" cellpadding="0" cellspacing="0">
              	  <tr>
              	  	<td width="95">捐款方式選擇：</td>
              	  	<td>
                      <input type="radio" name="Donate_Type" value="creditcard" checked >信用卡
                      <%If ECBank_WebAtm="Y" Then%><input type="radio" name="Donate_Type" value="webatm">Web-ATM<%End If%>
                      <%If ECBank_Vacc="Y" Then%><input type="radio" name="Donate_Type" value="vacc">銀行虛擬帳號<%End If%>	
                      <%If ECBank_Barcode="Y" Then%><input type="radio" name="Donate_Type" value="barcode">超商條碼<%End If%>
                      <%If ECBank_Ibon="Y" Then%><input type="radio" name="Donate_Type" value="ibon">7-11 ibon<%End If%>
                      <%If ECBank_FamiPort="Y" Then%><input type="radio" name="Donate_Type" value="famiport">全家 Famiport<%End If%><br />
                      <%If ECBank_LiftET="Y" Then%><input type="radio" name="Donate_Type" value="lifeet" >萊爾富 Life-Et<%End If%>
                      <%If ECBank_OKGo="Y" Then%><input type="radio" name="Donate_Type" value="okgo" >OK超商 OK-GO<%End If%>
                      <input type="button" value=" 下一步 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:Next_OnClick();">	
              	  	</td>
              	  </tr>
              	</table>
              </td>
            </tr>
            <%End If%>
          </form>
        </table><br />
        <table width="193" height="30" border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="143" align="center" class="contents"><font color="#4e6118" size="3"><strong>捐款機制說明</strong></font></td>
          </tr>
        </table>
        <table width="100%" border="0" align="center" cellpadding="3" cellspacing="3">
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">本公益勸募資料傳送採用SSL(Secure Socket Layer)傳輸加密並使用綠界金流ECPAY付款閘道。</td>
          </tr>
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">您的重要資料皆用複雜的數學運算方式進行亂數編碼再進行資料傳送，且取得網際威信的伺服器數位憑證（HiTrust，如下圖示），讓您能夠安全安心地進行線上捐款動作。</td>
          </tr>
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">您捐贈的每一筆金額，皆可開立捐款收據！「<%=Comp_Name%>」不會將您的個人資料透漏給第三者，請放心填寫之！</td>
          </tr>
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">提醒您，在網頁下方的狀態列中會有一個上鎖的U型鎖匙(或在網頁頁右鍵按選內容可檢視您目前連線方式及加密強度)，表示您的重要資料是以安全有保障的方式在網路上傳送。</td>
          </tr>
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">捐款方式選擇：</td>
          </tr>
          <tr>
            <td valign="top" align="left"> </td>
            <td align="left" class="contents">
              <table width="100%" border="0" align="left" cellpadding="1" cellspacing="1">
                <tr>
                  <td class="contents"><font color="#d51017">〈一〉</font></td>
                  <td class="contents"><font color="#d51017"><b>信用卡</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請先於本網頁上方選擇『信用卡』捐款方式，並於送出信用卡資料後，靜候5～10秒的授權處理時間。若連續授權失敗，可能是付款伺服器過於忙碌所致，請您稍候再嘗試捐款。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈二〉</font></td>
                  <td class="contents"><font color="#d51017"><b>Web-ATM</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents"><font color="#000099">須使用讀卡機</font>，請先於本網頁上方選擇『Web-ATM』捐款方式，任何銀行發行的晶片提款卡都可以使用Web-ATM捐款，不受跨行限制。<br /><font color="#000099">重要提醒</font>:非捐款合作銀行(&nbsp;元大、玉山、台新&nbsp;)之金融卡轉帳時銀行將<font color="#000099">加收跨行轉帳手續費</font>，此費用因非由本會收取因此將不會在捐款收據上呈現。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈三〉</font></td>
                  <td class="contents"><font color="#d51017"><b>銀行虛擬帳號</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請先於本網頁上方選擇『銀行虛擬帳號』捐款方式，<font color="#000099">並請牢記轉帳帳號</font>，至任何一台提款機轉帳即可。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈四〉</font></td>
                  <td class="contents"><font color="#d51017"><b>超商條碼</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請先於本網頁上方選擇『超商條碼』捐款方式，並列印捐款單據(<font color="#000099">&nbsp;請使用雷射印表機列印，一般噴墨印表機列印之品質，很可能導致超商無法讀取條碼&nbsp;</font>)，持單據可至任何一家超商繳款。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈五〉</font></td>
                  <td class="contents"><font color="#d51017"><b>7-11 ibon、全家 Famiport、萊爾富 Life-et、OK超商 OK-GO</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請先於本網頁上方選擇『7-11 ibon、全家 Famiport、萊爾富 Life-et、OK超商 OK-GO』任一捐款方式，<font color="#000099">並請牢記收費代碼至超商機器輸入代碼</font>，列印捐款單至櫃檯繳款。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈六〉</font></td>
                  <td class="contents"><font color="#d51017"><b>郵政劃撥</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">關懷帳戶：<%=Comp_Name%><br />關懷帳號：<%=Account%><br />請您至郵局索取劃撥單，並請詳細填寫捐款人聯絡資料，完成劃撥。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈七〉</font></td>
                  <td class="contents"><font color="#d51017"><b>信用卡自動扣款</b></font>(&nbsp;<a href="<%=download_url%>"><font color="#4e6118"><b>授權書下載</b></font></a>&nbsp;)</td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請下載委託轉帳代繳慈善捐款授權書，填妥資料用印後，請用掛號將正本寄至<br /><%=Dept_Address%><br /><%=Comp_Name%>&nbsp;&nbsp;會計室收</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈八〉</font></td>
                  <td class="contents"><font color="#d51017"><b>電匯</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">關懷帳戶：<%=Comp_Name%><br />匯款銀行：華南銀行淡水分行(銀行代號008)<br />匯款帳號：167200033988<br />請您至任何一家銀行匯款，匯款完成後請務必來電告知以便本校開立收據。</td>
                </tr>
                <tr>
                  <td class="contents"><font color="#d51017">〈九〉</font></td>
                  <td class="contents"><font color="#d51017"><b>支票捐款</b></font></td>
                </tr>
                <tr>
                  <td class="contents"></td>
                  <td class="contents">請您將支票寄至<br /><%=Dept_Address%><br /><%=Comp_Name%>&nbsp;會計室收<br />並於背面註明禁止背書轉讓。</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <td valign="top" align="left"><span lang="EN-US" style="font-size: 10.5pt; font-family: Webdings; mso-ascii-font-family: 'Times New Roman'; mso-hansi-font-family: 'Times New Roman'; mso-char-type: symbol; mso-symbol-font-family: Webdings"><span style="mso-char-type: symbol; mso-symbol-font-family: Webdings"><font color="#d51017">Y</font></span></span><span lang="EN-US" style="font-size: 10.5pt"><font face="Times New Roman"></font></span></td>
            <td align="left" class="contents">如對線上<font color="#000099">捐款機制</font>有疑問，請洽 <font color="#000099"><%=Dept_Tel%> 會計室。</td>
          </tr>
        </table>
        <table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center" valign="center">
              <a href="http://www.greenworld.com.tw/" target="_blank"><img alt="連線至綠界科技(另開視窗)" border="0" src="image/ecpay.gif" /></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
              <a href="http://www.hitrust.com.tw/" target="_blank"><img alt="連線至網際威信(另開視窗)" border="0" src="image/htrust.gif" /></a>
              <a href="http://www.ecbank.com.tw/" target="_blank"><img alt="連線至綠界科技ECBank(另開視窗)" border="0" src="image/cvs.gif" /></a>
            </td>
          </tr>
        </table>
        <%If ECBank_WebAtm="N" And ECBank_Vacc="N" And ECBank_Barcode="N" And ECBank_Ibon="N" And ECBank_FamiPort="N" And ECBank_LiftET="N" And ECBank_OKGo="N" Then%>
        <br /><table border="0" width="100%" cellspacing="0" cellpadding="0">
          <tr>
            <td align="center" valign="center"><input type="button" value=" 我要捐款 " name="But_Dn" class="cbutton" style="cursor:hand;" onClick="javascript:Next_OnClick();"></td>
          </tr>
        </table>
        <%End If%>
      </div>
      <div id="bottom"></div>
    </div>
  </div>
  <%Message()%>
</body>
</html>
<!--#include file="include/dbclose.asp"-->