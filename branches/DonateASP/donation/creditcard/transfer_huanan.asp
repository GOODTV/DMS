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

'交易日期
Donate_Year_AD=Cstr(Year(Date()))
Donate_Year=Cstr(Year(Date())-1911)
If Len(Donate_Year)=2 Then Donate_Year="0"&Donate_Year
Donate_Month=Cstr(Month(Date()))
If Len(Donate_Month)=1 Then Donate_Month="0"&Donate_Month
Donate_Day=Cstr(Day(Date()))
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
Donate_Hour=Cstr(Hour(Now()))
If Len(Donate_Day)=1 Then Donate_Day="0"&Donate_Day
Donate_Minute=Cstr(Minute(Now()))
If Len(Donate_Minute)=1 Then Donate_Minute="0"&Donate_Minute
Donate_Second=Cstr(Second(Now()))
If Len(Donate_Second)=1 Then Donate_Second="0"&Donate_Second

SQL="Select * From DONATE_TRANSFER Where Transfer_AspName='"&Session("Transfer_AspName")&"'"
Set RS=Server.CreateObject("ADODB.Recordset")
RS.Open SQL,Conn,1,3
Transfer_StoreCode=RS("Transfer_StoreCode")
Transfer_ServerCode=RS("Transfer_ServerCode")
Transfer_UniformNo=RS("Transfer_UniformNo")
Transfer_FileName=RS("Transfer_FileName")
OrderNo=""
If RS("Transfer_OrderNo")<>"" Then
  If Instr(RS("Transfer_OrderNo"),Donate_Year_AD&Donate_Month&Donate_Day)>0 Then
    OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&Left("00",2-Len(Cint(Right(RS("Transfer_OrderNo"),2))+1))&Cint(Right(RS("Transfer_OrderNo"),2))+1
  Else
    OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&"01"
  End If
Else
  OrderNo=Donate_Year_AD&Donate_Month&Donate_Day&"01"
End If
Transfer_FileName=Transfer_UniformNo&Right(Transfer_StoreCode,4)&OrderNo&""
RS("Transfer_OrderNo")=OrderNo
RS.Update
RS.Close
Set RS=Nothing

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'授權筆數
Row=0
'授權金額
Danate_Total=0

'必要條件(卡號不可為空白)
WhereSQL="And Account_No<>''"

Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,PLEDGE.Donor_Id,Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,Donate_Amt=(Case When Donate_FirstAmt>0 Then Donate_FirstAmt Else Donate_Amt End),DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='授權中' "&WhereSQL&" "
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    If Row>0 Then Response.Write vbcrlf
    '1.商店代號
    StoreLen=15
    If Len(Transfer_StoreCode)>=StoreLen Then
      Response.Write Left(Transfer_StoreCode,StoreLen)
    Else
      Response.Write Transfer_StoreCode&Space(StoreLen-Len(Transfer_StoreCode))
    End If
    '2.行業代號(MCC Code)
    ServereLen=4
    If Len(Transfer_ServerCode)>=ServereLen Then
      Response.Write Left(Transfer_ServerCode,ServereLen)
    Else
      Response.Write Transfer_ServerCode&Space(ServereLen-Len(Transfer_ServerCode))
    End If
    '3.卡號
    CardLen=19
    account_no= RS("account_no")
    If Len(account_no)>=CardLen Then
      Response.Write Left(account_no,CardLen)
    Else
      Response.Write account_no&Space(CardLen-Len(account_no))
    End If
    '4.授權碼
    Csv2Len=3
    Response.Write Space(Csv2Len)
	  '5.信用卡有效年月
    ValidDateLen=4
    Valid_Date=RS("Valid_Date")
    If Len(Valid_Date)>=ValidDateLen Then
      Response.Write Left(Valid_Date,ValidDateLen)
    Else
      Response.Write Valid_Date&Space(ValidDateLen-Len(Valid_Date))
    End If
    '6.金額
    AmtLen=12
    Donate_Amt=RS("Donate_Amt")&"00"
    For J= Len(Donate_Amt) To AmtLen-1
      Donate_Amt="0"&Donate_Amt
    Next
    Response.Write Donate_Amt
    '7.回應碼
    SpaceLen=2
    Response.Write Space(SpaceLen)
    '8.授權碼
    SpaceLen=6
    Response.Write Space(SpaceLen)
    '9.備註Filler
    SpaceLen=3
    Response.Write Space(SpaceLen)
    '10.備註Remark(商家自用)
    RemarkLen=29
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
    '11.備註Line Feed
    SpaceLen=2
    Response.Write Space(SpaceLen) 
    Danate_Total=Danate_Total+Cdbl(RS("Donate_Amt"))
    Row=Row+1
  End If
  RS.Close
  Set RS=Nothing
  Response.Flush
  Response.Clear  
Next
Response.Write vbcrlf
'8.末筆代號
Response.Write "T"
'9.總件數
DonateRowLen=6
If Len(Cstr(Row))>=DonateRowLen Then
  Response.Write Left(Cstr(Row),DonateRowLen)
Else
  StrDonateRow=""
  For J=1 To DonateRowLen
    StrDonateRow=StrDonateRow&"0"
  Next
  Response.Write Left(StrDonateRow,DonateRowLen-Len(Cstr(Row)))&Row
End If
'10.總金額
TotalLen=12
DanateTotal=Cstr(Danate_Total)&"00"
For J= Len(DanateTotal) To TotalLen-1
  DanateTotal="0"&DanateTotal
Next
Response.Write DanateTotal
'11.交易日/時間
Response.Write Donate_Year_AD&Donate_Month&Donate_Day&Donate_Hour&Donate_Minute&Donate_Second
'12.備註Filler
SpaceLen=35
Response.Write Space(SpaceLen)
'13.備註Remark
SpaceLen=29
Response.Write Space(SpaceLen)
'14.備註Line Feed
SpaceLen=2
Response.Write Space(SpaceLen)
Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
