<%Response.ContentType="text/html; charset=Big5"%>
<%
Response.Buffer =true
Response.Expires=-1
If session("user_id")="" Then
  Session.Abandon
  Response.Redirect("../sysmgr/timeout.asp")
End If
session.Timeout=60
Server.ScriptTimeout=600
Set Conn=server.createobject("ADODB.Connection")
Conn.Connectiontimeout=600
Conn.commandtimeout=600
Conn.Provider="sqloledb"
Conn.open "server="&session("server")&";uid="&session("uid")&";pwd="&session("pwd")&";database="&session("database")

P_PCLNO="00066001039995"
P_CID="81582054"

'授權筆數
Row=0
'授權金額
Danate_Total=0
'20131216 Modify by GoodTV Tanya:交易日期為畫面上的「扣款日期」
'交易日期
If Session("Donate_Date")="" Then 
	Donate_Year_AD=Cstr(Year(Now())-1911)
	Donate_Month=Cstr(Month(Now()))	
	Donate_Day=Cstr(Day(Now()))
Else 
	Donate_Year_AD=Cstr(Year(CDate(Session("Donate_Date")))-1911) 
	Donate_Month=Cstr(Month(CDate(Session("Donate_Date"))))
	Donate_Day=Cstr(Day(CDate(Session("Donate_Date"))))
End If

If Len(Donate_Year_AD)=3 Then Donate_Year_AD="0"&Donate_Year_AD
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
	
Donate_Hour=Cstr(Hour(Now()))
If Len(Donate_Hour)=1 Then Donate_Hour="0"&Donate_Hour
Donate_Minute=Cstr(Minute(Now()))
If Len(Donate_Minute)=1 Then Donate_Minute="0"&Donate_Minute
Donate_Second=Cstr(Second(Now()))
If Len(Donate_Second)=1 Then Donate_Second="0"&Donate_Second

SQL="Select * From DONATE_TRANSFER Where Transfer_AspName='"&Session("Transfer_AspName")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,2
Transfer_StoreCode=RS("Transfer_StoreCode")
Transfer_ServerCode=RS("Transfer_ServerCode")
OrderNo=""
If RS("Transfer_OrderNo")<>"" Then
  If Instr(RS("Transfer_OrderNo"),Donate_Year_AD&Donate_Month&Donate_Day)>0 Then
    OrderNo="ACH02_"&Donate_Year_AD&Donate_Month&Donate_Day&Left("00",2-Len(Cint(Right(RS("Transfer_OrderNo"),2))+1))&Cint(Right(RS("Transfer_OrderNo"),2))+1
  Else
    OrderNo="ACH02_"&Donate_Year_AD&Donate_Month&Donate_Day&"01"
  End If
Else
  OrderNo="ACH02_"&Donate_Year_AD&Donate_Month&Donate_Day&"01"
End If
Transfer_FileName=OrderNo&".txt"
RS("Transfer_OrderNo")=OrderNo
RS.Update
RS.Close
Set RS=Nothing

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"
'20131220 Modify by GoodTV Tanya:修改匯出資料格式
'ACH控制首錄
'1.首錄別
Response.Write "BOF"
'2.資料代號
Response.Write "ACHP02"
'3.處理日期
Response.Write Donate_Year_AD&Donate_Month&Donate_Day
''4.處理時間
'Response.Write Donate_Hour&Donate_Minute&Donate_Second
'5.發送單位代號
Response.Write Transfer_StoreCode
''6.接收單位代號
'Response.Write "9990250"
'7.備用
Response.Write Space(96)

'必要條件(收受行代號、收受者帳號、收受者身分證/統編不可為空白)
'20131113 Modify by GoodTV Tanya:扣款金額一律抓Donate_Amt(因為沒有分期付款的首次扣款金額)
WhereSQL="And P_BANK<>'' And P_RCLNO<>'' And P_PID<>''"
Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,Donor_Id=Isnull(PLEDGE.Donor_Id_Old,PLEDGE.Donor_Id),Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,P_BANK,P_RCLNO,P_PID,Member_No,Donate_Amt,DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  'response.write sql&"<br>"
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    'ACH明細資料
    Response.Write vbcrlf
		'1.交易序號(6位流水號)
    Row=Row+1
    Response.Write Left("000000",6-Len(Row))&Row
    '2.交易代號(529公益捐款) 20131113 Modify by GoodTV Tanya
    Response.Write "529"
    '3.發動者統一編號
    P_CIDLen=10
    If Len(P_CID)>=P_CIDLen Then
      Response.Write Left(P_CID,P_CIDLen)
    Else
      Response.Write P_CID&Space(P_CIDLen-Len(P_CID))
    End If  
    '4.收受行代號
    P_BANKLen=7
    P_BANK=RS("P_BANK")
    If Len(P_BANK)>=P_BANKLen Then
      Response.Write Left(P_BANK,P_BANKLen)
    Else
      Response.Write P_BANK&Space(P_BANKLen-Len(P_BANK))
    End If  
    '5.收受者帳號
    Response.Write Left("00000000000000",14-Len(RS("P_RCLNO")))&RS("P_RCLNO")
    '6.收受者統一編號
    P_PIDLen=10
    P_PID=RS("P_PID")
    If Len(P_PID)>=P_PIDLen Then
      Response.Write Left(P_PID,P_PIDLen)
    Else
      Response.Write P_PID&Space(P_PIDLen-Len(P_PID))
    End If
    '7.授權編號
    '20140113 Modify by GoodTV Tanya:授權編號為Donor_Id
    Member_NoLen=20
    Member_No=RS("Donor_Id")
    If Len(Member_No)>=Member_NoLen Then
      Response.Write Left(Member_No,Member_NoLen)
    Else
      Response.Write Member_No&Space(Member_NoLen-Len(Member_No))
    End If   
    '8.提示交換次序
    Response.Write "A"
    '9.交易日期
		Response.Write Donate_Year_AD&Donate_Month&Donate_Day		
    '10提出行代號
    Response.Write Transfer_ServerCode
    '11發動者專用區
    Response.Write Space(20)
		'12交易型態
    Response.Write "N"   
    Response.Write Space(1)    
    '13金額
    Response.Write Left("0000000000",8-Len(RS("Donate_Amt")))&RS("Donate_Amt")    
    '14備用
    Response.Write Space(4)

    Donate_Total=Cdbl(Donate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing 
Next

'ACH控制尾錄
Response.Write vbcrlf
'1.尾錄別
Response.Write "EOF"
''2.資料代號
'Response.Write "ACHP01"
''3.處理日期
'Response.Write Donate_Year_AD&Donate_Month&Donate_Day
''4.發送單位代號
'Response.Write Transfer_StoreCode
''5.接收單位代號
'Response.Write "9990250"
'6.總筆數
Response.Write Left("00000000",8-Len(Row))&Row
''7.總金額
'Response.Write Left("0000000000000000",16-Len(Donate_Total))&Donate_Total
'8.備用
Response.Write Space(109)
Response.Write vbcrlf

Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
