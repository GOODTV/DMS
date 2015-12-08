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
Transfer_StoreCode=RS("Transfer_StoreCode")
Transfer_ServerCode=RS("Transfer_ServerCode")
Transfer_FileName=RS("Transfer_FileName")
RS.Close
Set RS=Nothing

'交易日期
Donate_Year_AD=Cstr(Year(Date()))
If Session("Donate_Date")="" Then 
	Donate_Year=Cstr(Year(Date()))
	Donate_Month=Cstr(Month(Date()))
	Donate_Day=Cstr(Day(Date()))
Else 
	Donate_Year=Cstr(Year(CDate(Session("Donate_Date")))) 
	Donate_Month=Cstr(Month(CDate(Session("Donate_Date"))))
	Donate_Day=Cstr(Day(CDate(Session("Donate_Date"))))
End If

If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
	
Donate_Hour=Cstr(Hour(Now()))
If Len(Donate_Hour)=1 Then Donate_Hour="0"&Donate_Hour
Donate_Minute=Cstr(Minute(Now()))
If Len(Donate_Minute)=1 Then Donate_Minute="0"&Donate_Minute
Donate_Second=Cstr(Second(Now()))
If Len(Donate_Second)=1 Then Donate_Second="0"&Donate_Second

'Get當日檔案序號
SQL="Select * From DONATE_TRANSFER_ORDER Where Transfer_AspName='"&Session("Transfer_AspName")&"' AND Donate_Date='"&Donate_Year&Donate_Month&Donate_Day&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,1
If Not RS.Eof Then
	Donate_OrderNo = Cstr(CInt(RS("Donate_OrderNo"))+1)
Else
	Donate_OrderNo = "1"
End If
RS.Close
Set RS=Nothing
If Len(Donate_OrderNo)=1 Then Donate_OrderNo="0"&Donate_OrderNo

'Update DONATE_TRANSFER_ORDER
If Donate_OrderNo = "01" Then
	SQL1="DONATE_TRANSFER_ORDER"
  Set RS1 = Server.CreateObject("ADODB.RecordSet")  
  RS1.Open SQL1,Conn,1,3  
  RS1.Addnew  
  RS1("Transfer_AspName")=Session("Transfer_AspName")  
  RS1("Donate_Date")=Donate_Year&Donate_Month&Donate_Day  
	RS1("Donate_OrderNo")=1	
	RS1.Update	
  RS1.Close
  Set RS1=Nothing
Else
	SQL1="Select * From DONATE_TRANSFER_ORDER Where Transfer_AspName='"&Session("Transfer_AspName")&"' AND Donate_Date='"&Donate_Year&Donate_Month&Donate_Day&"'"
	Set RS1 = Server.CreateObject("ADODB.RecordSet")
  RS1.Open SQL1,Conn,1,3
  RS1("Donate_OrderNo")=CInt(Donate_OrderNo)
  RS1.Update
  RS1.Close
  Set RS1=Nothing 
End If


Transfer_FileName=Transfer_FileName&Donate_Month&Donate_Day&"."&Donate_OrderNo

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'授權筆數
Row=0
'授權金額
Donate_Total=0

'必要條件(卡號不可為空白)
WhereSQL="And Account_No<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select DONOR.Donor_Id,Pledge_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Status='授權中' And Pledge_Id='"&Ary_Pledge_Id(I)&"' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then    
    '1.商店代號(銀行別004+統一編號81582054+分店別9001)
    StoreLen=15
    If Len(Transfer_StoreCode)>=StoreLen Then
      Response.Write Left(Transfer_StoreCode,StoreLen)
    Else
      Response.Write Transfer_StoreCode
    End If
    '2.商店行業代號
    OrderLen=4
    Response.Write Transfer_ServerCode
    
    '3.卡號
    CardLen=19
    Account_No=RS("Account_No")
    If Len(Account_No)>=CardLen Then
      Response.Write Left(Account_No,CardLen)
    Else
      Response.Write Account_No&Space(CardLen-Len(Account_No))
    End If
    '4.信用卡CSV2
    AuthorizeLen=3
    Authorize=RS("Authorize")
		If isnull(RS("Authorize")) Then Authorize="   "
    If Len(Authorize)>=AuthorizeLen Then
      Response.Write Left(Authorize,AuthorizeLen)
    Else
      Response.Write Authorize&Space(AuthorizeLen-Len(Authorize))
    End If    
    '5.有效年月(年YY月MM)
    ValidDateLen=4
    Valid_Date=RS("Valid_Date")
    Response.Write Valid_Date
    
    '6.交易金額
    AmtLen=12
    Donate_Amt=RS("Donate_Amt")
    For J= Len(Donate_Amt) To AmtLen-3
      Donate_Amt="0"&Donate_Amt
    Next
    Response.Write Donate_Amt&"00"
    
    '7.回覆碼(空白)
    Response.Write Space(2)
    '8.授權碼(空白)
    Response.Write Space(6)
    '9.保留欄位(空白)
    Response.Write Space(3)
    
    '10.備註
    RemarkLen=30
    Donor_Id=RS("Donor_Id")
    For J= Len(Donor_Id) To 8-1
    	Donor_Id="0"&Donor_Id
    Next
    Pledge_Id=RS("Pledge_Id")
    For J= Len(Pledge_Id) To 8-1
    	Pledge_Id="0"&Pledge_Id
    Next
    Remark=Donor_Id&Pledge_Id
    Response.Write Remark&Space(RemarkLen-Len(Remark))
    
    '跳行符號
    Response.Write vbcrlf
    
    Row=Row+1    
    Donate_Total=Cdbl(Donate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing 
Next
Conn.Close
Set Conn=Nothing

'Tailer
	'1.檔尾
	Response.Write "T"
	
	'2.資料筆數
	RowLen=6
	For J= Len(Row) To 6-1
    	Row="0"&Row
  Next
	Response.Write Row
	
	'3.交易總額
	AmtTotalLen=12
	For J= Len(Donate_Total) To AmtTotalLen-3
      Donate_Total="0"&Donate_Total
  Next
  Response.Write Donate_Total&"00"
    
	'4.交易日期(YYYYMMDD)
	Response.Write Donate_Year&Donate_Month&Donate_Day
	
	'5.交易時間(HHMMSS)
	Response.Write Donate_Hour&Donate_Minute&Donate_Second
	
	'6.保留欄位
	Response.Write Space(35)
	'7.備註
	Response.Write Space(30)
	
	'跳行符號
  Response.Write vbcrlf

Session.Contents.Remove("Transfer_AspName")
%>
