<%Response.ContentType="text/html; charset=utf-8"%>
<%
Response.Buffer =true
Response.Expires=-1
If session("user_id")="" Then
  Session.Abandon
  Response.Redirect("../../sysmgr/timeout.asp")
End If
session.Timeout=60
Server.ScriptTimeout=600
Set Conn=server.createobject("ADODB.Connection")
Conn.Connectiontimeout=600
Conn.commandtimeout=600
Conn.Provider="sqloledb"
Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")

SQL="Select * From DONATE_TRANSFER Where Transfer_AspName='"&Session("Transfer_AspName")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,1
'商店代號
Transfer_StoreCode=RS("Transfer_StoreCode")
StoreLen=15
If Len(Transfer_StoreCode)>=StoreLen Then
  Transfer_StoreCode=Left(Transfer_StoreCode,StoreLen)
Else
  Transfer_StoreCode=Transfer_StoreCode&Space(StoreLen-Len(Transfer_StoreCode))
End If
'終端機號碼   
Transfer_ServerCode=RS("Transfer_ServerCode")
ServereLen=8
If Len(Transfer_ServerCode)>=ServereLen Then
  Transfer_ServerCode=Left(Transfer_ServerCode,ServereLen)
Else
  Transfer_ServerCode=Transfer_ServerCode&Space(ServereLen-Len(Transfer_ServerCode))
End If
'檔案名稱
Transfer_FileName=""&RS("Transfer_StoreCode")&"."&Right(Year(Date()),2)&Right("0"&Month(Date()),2)&Right("0"&Day(Date()),2)&".01.in"
RS.Close
Set RS=Nothing

Response.Charset ="BIG5"
Response.ContentType = "Content-Language;content=zh-tw" 
Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "text/html"

'授權筆數
Row=0
'授權金額
Danate_Total=0
'交易日期
Donate_Year_AD=Cstr(Year(Date()))
Donate_Year=Cstr(Year(Date())-1911)
If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
Donate_Month=Cstr(Month(Date()))
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
Donate_Day=Cstr(Day(Date()))
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
Transfer_Date=Right(Donate_Year_AD,2)&Donate_Month&Donate_Day

'查詢時必要條件(卡號不可為空白)
WhereSQL="And Account_No<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End) From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    Danate_Total=Csng(Danate_Total)+Csng(RS("Donate_Amt"))
    Row=Row+1
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing
Next

'產生抬頭資料
'1.記錄識別碼
Response.Write "H"
'終端機號碼
Response.Write Transfer_ServerCode
'商店代號
Response.Write Transfer_StoreCode
'幣別
Response.Write "T"
'交易總筆數
Response.Write Left("000000",6-Len(Row))&Cstr(Row)
'授權不清算交易筆數(A)
Response.Write "000000"
'授權不清算交易金額(A)
Response.Write "000000000000"
'授權後清算交易
Response.Write Left("000000",6-Len(Row))&Cstr(Row)
'授權後清算交易(S)金額
Response.Write Left("0000000000",10-Len(Danate_Total))&Cstr(Danate_Total)&"00"
'補登清算交易筆數(O)
Response.Write "000000"
'補登清算交易金額(O)
Response.Write "000000000000"
'退貨清算交易筆數(R)
Response.Write "000000"
'退貨清算交易金額(R)
Response.Write "000000000000"
'保留欄位
RemarkLen=151
Response.Write Space(RemarkLen)
Response.Write vbcrlf

'產生內容資料
Row2=0
Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,PLEDGE.Donor_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    Row2=Row2+1
    If Row2>1 Then Response.Write vbcrlf
	  '1.記錄識別碼(1)
	  Response.Write "D"
	  '2.自訂交易序號
	  Response.Write Left("000000",6-Len(Row2))&Cstr(Row2)
    '3.終端機號碼
    Response.Write Transfer_ServerCode
    '4.商店代號
    Response.Write Transfer_StoreCode
	  '5.交易日期
    Response.Write Transfer_Date
    '6.授權清算方式 A:只授權不清算 S:授權後清算 O:補登清算 R:退貨清算
    Response.Write"S"
    '7.卡號
    CardLen=19
    account_no=RS("account_no")
    If Len(account_no)>=CardLen Then
      Response.Write Left(account_no,CardLen)
    Else
      Response.Write account_no&Space(CardLen-Len(account_no))
    End If
	  '8.信用卡有效月年
    ValidDateLen=4
    Valid_Date=Right(RS("Valid_Date"),2)&Left(RS("Valid_Date"),2)
    If Len(Valid_Date)>=ValidDateLen Then
      Response.Write Left(Valid_Date,ValidDateLen)
    Else
      Response.Write Valid_Date&Space(ValidDateLen-Len(Valid_Date))
    End If 
    '9.CVC2
    CVC2Len=3
    Response.Write Space(CVC2Len)
    '10.交易金額
    Response.Write Left("0000000000",10-Len(RS("Donate_Amt")))&Cstr(RS("Donate_Amt"))&"00"
    '11.核准號碼(未授權填入999999已取得授權碼則直接填入授權碼)
    Response.Write "999999"
    '12.回復碼(未授權填入99已取得授權碼則直接填入授權碼) 
    Response.Write "99"
    '13.持卡人帳單列印資料
    PrintLen=20
    Response.Write Space(PrintLen)
    '14.商店保留欄位
    StoreLen=20
    Response.Write Space(StoreLen)
    '15.備註(商家自用)
    RemarkLen=40
    Donor_Remark=""
    Donor_Name=""
    If Instr(RS("Donor_Name"),"&#")>0 Then
      For J=1 To Len(RS("Donor_Name"))
        If Mid(RS("Donor_Name"),J,1)="&" And J<Len(RS("Donor_Name")) Then
          If Mid(RS("Donor_Name"),J+1,1)="#" Then
            Donor_Name=Donor_Name&"*"
            J=J+7
          Else
            Donor_Name=Donor_Name&Mid(RS("Donor_Name"),J,1)
          End If
        Else
          Donor_Name=Donor_Name&Mid(RS("Donor_Name"),J,1)
        End If
      Next
    Else
      Donor_Name=RS("Donor_Name")
    End If
    DonorRemark=Left("00000",5-Len(RS("Pledge_Id")))&RS("Pledge_Id")&Donor_Name&RS("Donor_Id")
    iLen=0
    For J = 1 To Len(DonorRemark)
      If Asc(Mid(DonorRemark,J,1))<0 Then
    	  sLen=2
      Else
        sLen=1
      End If
      If sLen+iLen<=RemarkLen Then
        Donor_Remark=Donor_Remark&Mid(DonorRemark,J,1)
        iLen=iLen+sLen
      Else
        Exit For
      End If
    Next
    Response.Write Donor_Remark
	  If iLen<RemarkLen Then Response.Write Space(RemarkLen-iLen)
    '15.分期註記
    Response.Write Space(1)
    '16.產品代碼
    Response.Write Space(10)
	  '17.期數
    Response.Write Space(3)
	  '18.每期金額
    Response.Write Space(9)
    '19.第一期金額
    Response.Write Space(9)
    '20.手續費率
    Response.Write Space(7)
    '21.晶片碼(TC)
    Response.Write Space(16)
    '22.保留欄位
    Response.Write Space(36)
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing
Next
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
