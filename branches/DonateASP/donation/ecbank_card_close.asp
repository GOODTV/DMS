<!--#include file="../include/dbfunctionJ.asp"-->
<!--#include file="../include/email.inc"-->
<%
Function get_close_url (Client_Card,PayName,Card_Verify,gwsr,amount)
  get_close_url=""
  PayName=Card_PayName
  Client_Card=Trim(Card_Code)
  If Len(Client_Card)>6 Then Client_Card=Len(Client_Card)
  CkeckD=Card_Verify
  If Csng(Client_Card)<>3 Then
    '商店代碼Card_Code(6碼,靠右,不足左補0)  
    Client_Card=Left("000000",6-Len(Cstr(Client_Card)))&Cstr(Client_Card)
    '授權單號gwsrcode(10碼,靠右,不足左捕0)
    gwsrcode=Left("0000000000",10-Len(Cstr(gwsr))) & Cstr(gwsr)
    '金額money(8碼,靠右,不足左捕0)
    money=Left("00000000",8-Len(Cstr(amount))) & Cstr(amount)
    '檢核碼CkeckE(9碼,靠右,不足左捕0)
    Ckeck=Client_Card & gwsrcode & money
    CkeckA=""
    CkeckB=""
    For I = 1 To Len(Ckeck)
      If I>=5 Then
        If I Mod 5 = 0 Then
          If CkeckA<>"" Then
            CkeckA=CkeckA & Mid(Ckeck,I,1)
          Else
            CkeckA=Mid(Ckeck,I,1)
          End If
        End If  
      End If
      If I>=3 Then
        If I Mod 3 = 0 Then
          If CkeckB<>"" Then
            CkeckB=CkeckB & Mid(Ckeck,I,1)
          Else
            CkeckB=Mid(Ckeck,I,1)
          End If
        End If
      End If
    Next
    '亂數碼CkeckC
    For I = 1 to 3
      Randomize
      CkeckC = CkeckC & Chr(int(10*Rnd)+48)
    Next
    CkeckE_Mod=(CkeckA*CkeckC+CkeckB) Mod CkeckD
    CkeckE=Left("000000000",9-Len(Cstr(CkeckE_Mod))) & Cstr(CkeckE_Mod)
    '關帳URL
    If PayName="ecpay" Then
      get_close_url="https://ecpay.com.tw/g_get.php?" & "C" & Client_Card & gwsrcode & money & CkeckE & CkeckC
    ElseIf PayName="gwpay" Then
      get_close_url="https://gwpay.com.tw/close_get.php?" & "C" & Client_Card & gwsrcode & money & CkeckE & CkeckC
    End If
  Else
    get_close_url="3"
  End If
End Function

WhereSQL=""
If Request("Donate_DonorName")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_DonorName="&Request("Donate_DonorName")
  Else
    WhereSQL=WhereSQL&"&Donate_DonorName="&Request("Donate_DonorName")
  End If  
End If
If Request("Donate_Purpose")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_Purpose="&Request("Donate_Purpose")
  Else
    WhereSQL=WhereSQL&"&Donate_Purpose="&Request("Donate_Purpose")
  End If   
End If
If Request("Donate_Invoice_Type")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_Invoice_Type="&Request("Donate_Invoice_Type")
  Else
    WhereSQL=WhereSQL&"&Donate_Invoice_Type="&Request("Donate_Invoice_Type")
  End If
End If
If Request("Donate_CreateDate_B")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_CreateDate_B="&Request("Donate_CreateDate_B")
  Else
    WhereSQL=WhereSQL&"&Donate_CreateDate_B="&Request("Donate_CreateDate_B")
  End If
End If
If Request("Donate_CreateDate_E")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_CreateDate_E="&Request("Donate_CreateDate_E")
  Else
    WhereSQL=WhereSQL&"&Donate_CreateDate_E="&Request("Donate_CreateDate_E")
  End If
End If
If Request("Donate_Type")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_Type="&Request("Donate_Type")
  Else
    WhereSQL=WhereSQL&"&Donate_Type="&Request("Donate_Type")
  End If
End If
If Request("Close_Type")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Close_Type="&Request("Close_Type")
  Else
    WhereSQL=WhereSQL&"&Close_Type="&Request("Close_Type")
  End If
End If
If Request("Donate_Export")<>"" Then
  If WhereSQL="" Then
    WhereSQL="?Donate_Export="&Request("Donate_Export")
  Else
    WhereSQL=WhereSQL&"&Donate_Export="&Request("Donate_Export")
  End If
End If

'捐款參數
SQL1="Select * From DONATE_SYSTEM Where System_Key In ('Card_Code','Card_PayName','Card_Verify') Order By System_Key"
Set RS1=Server.CreateObject("ADODB.Recordset")
RS1.Open SQL1,conn,1,1
While Not RS1.EOF
  If RS1("System_Key")="Card_Code"    Then Card_Code=RS1("System_Value")
  If RS1("System_Key")="Card_PayName" Then Card_PayName=RS1("System_Value")
  If RS1("System_Key")="Card_Verify"  Then Card_Verify=RS1("System_Value")
  RS1.MoveNext
Wend
RS1.Close
Set RS1=Nothing

SQL1="Select gwsr,amount From DONATE_ECPAY Where od_sob='"&Request("od_sob")&"' And succ='1' And gwsr<>'' And CONVERT(numeric,IsNull(amount,0))>0 "
Set RS1=Server.CreateObject("ADODB.Recordset")
RS1.Open SQL1,conn,1,1
If Not RS1.EOF Then
  gwsr=RS1("gwsr")
  amount=RS1("amount")
Else
  session("errnumber")=1
  session("msg")="觸發請款失敗 !"
  response.redirect "ecbank_card_qry_list.asp"&WhereSQL
End If
RS1.Close
Set RS1=Nothing

'交易成功觸發請款並關帳
close_url=get_close_url(Card_Code,Card_PayName,Card_Verify,gwsr,amount)
If close_url<>"3" Then
  For I=1 To 10
    close_type=WinHttp(close_url,"POST")
    If close_type="ok" Then Exit For
  Next
Else
  close_type="ok"
End If
Set RS1 = Server.CreateObject("ADODB.RecordSet")
RS1.Open "Select * From DONATE_ECPAY Where od_sob = '"&Request("od_sob")&"'", Conn, 1, 3
If Not RS1.EOF Then
  RS1("close_type") = close_type
  RS1("close_url") = close_url
  RS1.Update
Else
  session("errnumber")=1
  session("msg")="觸發請款失敗 !"
  response.redirect "ecbank_card_qry_list.asp"&WhereSQL 
End If
RS1.Close
Set RS1 = Nothing          
session("errnumber")=1
session("msg")="觸發請款成功 !"
response.redirect "ecbank_card_qry_list.asp"&WhereSQL
%>