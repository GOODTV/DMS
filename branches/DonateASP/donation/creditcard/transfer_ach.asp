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

'���v����
Row=0
'���v���B
Danate_Total=0
'20131216 Modify by GoodTV Tanya:���������e���W���u���ڤ���v
'������
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
    OrderNo="ACH"&Donate_Year_AD&Donate_Month&Donate_Day&Left("00",2-Len(Cint(Right(RS("Transfer_OrderNo"),2))+1))&Cint(Right(RS("Transfer_OrderNo"),2))+1
  Else
    OrderNo="ACH"&Donate_Year_AD&Donate_Month&Donate_Day&"01"
  End If
Else
  OrderNo="ACH"&Donate_Year_AD&Donate_Month&Donate_Day&"01"
End If
Transfer_FileName=OrderNo&".txt"
RS("Transfer_OrderNo")=OrderNo
RS.Update
RS.Close
Set RS=Nothing

Response.AddHeader "Content-disposition","attachment; filename="&Transfer_FileName&""
Response.ContentType = "application/vnd.ms-txt"

'ACH�����
'1.�����O
Response.Write "BOF"
'2.��ƥN��
Response.Write "ACHP01"
'3.�B�z���
Response.Write Donate_Year_AD&Donate_Month&Donate_Day
'4.�B�z�ɶ�
Response.Write Donate_Hour&Donate_Minute&Donate_Second
'5.�o�e���N��
Response.Write Transfer_StoreCode
'6.�������N��
Response.Write "9990250"
'7.�ƥ�
Response.Write Space(123)

'���n����(������N���B�����̱b���B�����̨�����/�νs���i���ť�)
'20131113 Modify by GoodTV Tanya:���ڪ��B�@�ߧ�Donate_Amt(�]���S�������I�ڪ��������ڪ��B)
WhereSQL="And P_BANK<>'' And P_RCLNO<>'' And P_PID<>''"
Ary_Pledge_Id=Split(request("pledge_id"),",")
For I = 0 To UBound(Ary_Pledge_Id)
  SQL="Select Pledge_Id,Donor_Id=Isnull(PLEDGE.Donor_Id_Old,PLEDGE.Donor_Id),Card_Bank,Card_Type,Account_No,Valid_Date=(Case When Valid_Date<>'' Then Right(Valid_Date,2)+Left(Valid_Date,2) Else '' End),Card_Owner,Owner_IDNo,Authorize, " & _
      "Post_SavingsNo,Post_AccountNo,Post_IDNo,P_BANK,P_RCLNO,P_PID,Member_No,Donate_Amt,DONOR.Donor_Name,Donate_ToDate " & _
      "From PLEDGE Join DONOR On PLEDGE.Donor_Id=DONOR.Donor_Id Where Pledge_Id='"&Trim(Ary_Pledge_Id(I))&"' And Status='���v��' "&WhereSQL&" "
  'response.write sql&"<br>"
  Set RS=Server.CreateObject("ADODB.Recordset")
  RS.Open SQL,Conn,1,1
  If Not RS.EOF Then
    'ACH���Ӹ��
    Response.Write vbcrlf
    '1.������A
    Response.Write "N"
    '2.������O
    Response.Write "SD"
    '3.����N��(529���q����) 20131113 Modify by GoodTV Tanya
    Response.Write "529"
    '4.����Ǹ�(6��y����)
    Row=Row+1
    Response.Write Left("000000",6-Len(Row))&Row
    '5.���X��N��
    Response.Write Transfer_ServerCode
    '6.�o�ʪ̱b��
    Response.Write P_PCLNO
    '7.���^��N��
    P_BANKLen=7
    P_BANK=RS("P_BANK")
    If Len(P_BANK)>=P_BANKLen Then
      Response.Write Left(P_BANK,P_BANKLen)
    Else
      Response.Write P_BANK&Space(P_BANKLen-Len(P_BANK))
    End If
    '8�����̱b��
    Response.Write Left("00000000000000",14-Len(RS("P_RCLNO")))&RS("P_RCLNO")
    '9���B
    Response.Write Left("0000000000",10-Len(RS("Donate_Amt")))&RS("Donate_Amt")
    '10�h��z�ѥN��
    Response.Write "00"
    '11���ܥ洫����
    Response.Write "B"
    '12�o�ʪ̲Τ@�s��
    P_CIDLen=10
    If Len(P_CID)>=P_CIDLen Then
      Response.Write Left(P_CID,P_CIDLen)
    Else
      Response.Write P_CID&Space(P_CIDLen-Len(P_CID))
    End If
    '13�����̲Τ@�s��
    P_PIDLen=10
    P_PID=RS("P_PID")
    If Len(P_PID)>=P_PIDLen Then
      Response.Write Left(P_PID,P_PIDLen)
    Else
      Response.Write P_PID&Space(P_PIDLen-Len(P_PID))
    End If
    '14�W���W�d���q�N��
    Response.Write Space(6)
    '15�촣�ܥ�����
    Response.Write "00000000"
    '16�촣�ܥ���Ǹ�
    Response.Write "000000"
    '17�촣�ܥ洫����
    Response.Write Space(1)
    '18�Τḹ�X
    Member_NoLen=20
    Member_No=RS("Donor_Id")
    If Len(Member_No)>=Member_NoLen Then
      Response.Write Left(Member_No,Member_NoLen)
    Else
      Response.Write Member_No&Space(Member_NoLen-Len(Member_No))
    End If    
    '19�o�ʪ̱M�ΰ�
    Response.Write Space(20)
    '20�s�P�K�n
    Response.Write Space(10)
    '21�ƥ�
    Response.Write Space(2)

    Donate_Total=Cdbl(Donate_Total)+Cdbl(RS("Donate_Amt"))
    Response.Flush
    Response.Clear
  End If
  RS.Close
  Set RS=Nothing 
Next

'ACH�������
Response.Write vbcrlf
'1.�����O
Response.Write "EOF"
'2.��ƥN��
Response.Write "ACHP01"
'3.�B�z���
Response.Write Donate_Year_AD&Donate_Month&Donate_Day
'4.�o�e���N��
Response.Write Transfer_StoreCode
'5.�������N��
Response.Write "9990250"
'6.�`����
Response.Write Left("00000000",8-Len(Row))&Row
'7.�`���B
Response.Write Left("0000000000000000",16-Len(Donate_Total))&Donate_Total
'8.�ƥ�
Response.Write Space(105)
Response.Write vbcrlf

Conn.Close
Set Conn=Nothing
Session.Contents.Remove("Transfer_AspName")
%>
