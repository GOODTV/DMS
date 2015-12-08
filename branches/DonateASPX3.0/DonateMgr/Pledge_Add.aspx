<%@ Page Language="C#" AutoEventWireup="true" CodeFile="Pledge_Add.aspx.cs" Inherits="DonateMgr_Pledge_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>固定轉帳授權書【新增】</title>
    <link rel="stylesheet" type="text/css" href="../include/calendar-win2k-cold-1.css" />
    <link href="../include/main.css" rel="stylesheet" type="text/css" />
    <link href="../include/table.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="../include/calendar.js"></script>
    <script type="text/javascript" src="../include/calendar-big5.js"></script>
    <script type="text/javascript" src="../include/jquery-1.7.js"></script>
    <script type="text/javascript" src="../include/jquery.metadata.min.js"></script>
    <script type="text/javascript" src="../include/jquery.swapimage.min.js"></script>
    <script type="text/javascript" src="../include/common.js"></script>
    <script type="text/javascript">
        // menu控制
        $(document).ready(function () {
            InitMenu();
        });
    </script>
    <script type="text/javascript">
        window.onload = initCalendar;
        function initCalendar() {
            Calendar.setup({
                inputField: "tbxDonate_FromDate",   // id of the input field
                button: "imgDonate_FromDate"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxDonate_ToDate",   // id of the input field
                button: "imgDonate_ToDate"     // 與觸發動作的物件ID相同
            });
            Calendar.setup({
                inputField: "tbxNext_DonateDate",   // id of the input field
                button: "imgNext_DonateDate"     // 與觸發動作的物件ID相同
            });
        }
        function Account_No_Keyup(AccountNo) {
            var Account_No1 = document.getElementById('tbxAccount_No1');
            var Account_No2 = document.getElementById('tbxAccount_No2');
            var Account_No3 = document.getElementById('tbxAccount_No3');
            var Account_No4 = document.getElementById('tbxAccount_No4');
            if (AccountNo == 1 && Account_No1.value.length == 4) {
                Account_No2.focus();
            } else if (AccountNo == 2 && Account_No2.value.length == 4) {
                Account_No3.focus();
            } else if (AccountNo == 3 && Account_No3.value.length == 4) {
                Account_No4.focus();
            }
        }
        function Same_Owner_OnClick() {
            var Same_Owner = document.getElementById('cbxSame_Owner');
            var Card_Owner = document.getElementById('tbxCard_Owner');
            var Donor_Name = document.getElementById('tbxDonor_Name');
            var Owner_IDNo = document.getElementById('tbxOwner_IDNo');
            var DonorIDNo = document.getElementById('tbxIDNo');
            var Relation = document.getElementById('tbxRelation');

            if (Same_Owner.checked) {
                Card_Owner.value = Donor_Name.value;
                Owner_IDNo.value = DonorIDNo.value;
                Relation.value = '本人';
            }
        }
        function Same_Owner2_OnClick() {
            var Same_Owner2 = document.getElementById('cbxSame_Owner2');

            var Post_Name = document.getElementById('tbxPost_Name');
            var Donor_Name = document.getElementById('tbxDonor_Name');
            var Post_IDNo = document.getElementById('tbxPost_IDNo');
            var DonorIDNo = document.getElementById('tbxIDNo');

            if (Same_Owner2.checked) {
                Post_Name.value = Donor_Name.value;
                Post_IDNo.value = DonorIDNo.value;
            }
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt=0;
            var sName = "";
            var ddlDonate_Payment = document.getElementById('ddlDonate_Payment');
            var ddlDonate_Purpose = document.getElementById('ddlDonate_Purpose');
            var tbxDonate_Amt = document.getElementById('tbxDonate_Amt');
            var tbxDonate_FromDate = document.getElementById('tbxDonate_FromDate');
            var tbxDonate_ToDate = document.getElementById('tbxDonate_ToDate');
            var tbxCard_Bank = document.getElementById('tbxCard_Bank');
            var ddlCard_Type = document.getElementById('ddlCard_Type'); 
            var tbxAccount_No1 = document.getElementById('tbxAccount_No1');
            var tbxAccount_No2 = document.getElementById('tbxAccount_No2');
            var tbxAccount_No3 = document.getElementById('tbxAccount_No3');
            var tbxAccount_No4 = document.getElementById('tbxAccount_No4');
            var ddlMonth_Valid_Date = document.getElementById('ddlMonth_Valid_Date');
            var ddlYear_Valid_Date = document.getElementById('ddlYear_Valid_Date');
            var tbxCard_Owner = document.getElementById('tbxCard_Owner');
            var tbxOwner_IDNo = document.getElementById('tbxOwner_IDNo');
            var tbxRelation = document.getElementById('tbxRelation');
            var tbxAuthorize = document.getElementById('tbxAuthorize');
            var tbxPost_Name = document.getElementById('tbxPost_Name');
            var tbxPost_IDNo = document.getElementById('tbxPost_IDNo');
            var tbxPost_SavingsNo = document.getElementById('tbxPost_SavingsNo');
            var tbxPost_AccountNo = document.getElementById('tbxPost_AccountNo');
            var ddlDept = document.getElementById('ddlDept');
            var tbxInvoice_Title = document.getElementById('tbxInvoice_Title');
            var ddlInvoice_Type = document.getElementById('ddlInvoice_Type');
            var tbxAccoun_Bank = document.getElementById('tbxAccoun_Bank'); 
            var ddlAccounting_Title = document.getElementById('ddlAccounting_Title');  
            var tbxInvoice_Type = document.getElementById('tbxInvoice_Type');   
            var tbxP_BANK = document.getElementById('tbxP_BANK');
            var tbxP_RCLNO = document.getElementById('tbxP_RCLNO');
            var tbxP_PID = document.getElementById('tbxP_PID');   
            if ($('#ddlDonate_Payment option:selected').text() == "") {
                strRet += "授權方式 ";
            }
            if ($('#ddlDonate_Purpose option:selected').text() == "") {
                strRet += "指定用途 ";
            }
            if ($('#tbxDonate_Amt').val() == "") {
                strRet += "每次扣款 ";
            }
            if ($('#tbxDonate_ToDate').val() == "") {
                strRet += "授權迄日 ";
            }
            if (ddlDept.value == "") {
                strRet += "機構名稱 ";
            }
            if (ddlInvoice_Type.value == "單次收據" || ddlInvoice_Type.value == "年度證明") {
                if (tbxInvoice_Title.value == "") {
                    strRet += "收據抬頭 ";
                }
            }
            if (ddlInvoice_Type.value == "") {
                strRet += "收據開立 ";
            }
            if (strRet != "") {
                strRet += "欄位不可為空白！";
                alert(strRet)
                return false;
            }
            cnt = 0;
            sName = ddlDonate_Payment.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('授權方式 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlDonate_Purpose.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('指定用途 欄位長度超過限制！');
                return false;
            }
            if (isNaN(Number(tbxDonate_Amt.value)) == true) {
                alert('每次扣款 欄位必須為數字！');
                tbxDonate_Amt.focus();
                return false;
            }
            var begindate = new Date(tbxDonate_FromDate.value)
            var enddate = new Date(tbxDonate_ToDate.value)
            var diffdate = (Date.parse(begindate.toString()) - Date.parse(enddate.toString())) / (1000 * 60 * 60 * 24)
            if (parseInt(diffdate) > 0) {
                alert('授權起日不可大於迄日！');
                tbxDonate_FromDate.focus();
                return false;
            }
            //20140509 修改 by Ian_Kao 
            if ($('#ddlDonate_Payment option:selected').text() == "信用卡授權書(一般)" || $('#ddlDonate_Payment option:selected').text() == "信用卡授權書(聯信)") {
                /*if (tbxCard_Bank.value == '') {
                    strRet += "銀行別 ";
                }*/
                if (ddlCard_Type.value == '') {
                    strRet += "信用卡別 ";
                }
                if ($('#ddlDonate_Payment option:selected').text() == "信用卡授權書(聯信)") {
                    if (ddlCard_Type.value != 'AE') {
                        strRet += "信用卡別需是AE";
                    }
                }
                if (tbxAccount_No1.value == "" || tbxAccount_No2.value == "" || tbxAccount_No3.value == "" || tbxAccount_No4.value == "") {
                    strRet += "信用卡號 ";
                }
                if (ddlMonth_Valid_Date.value == '') {
                    strRet += "有效月 ";
                }
                if (ddlYear_Valid_Date.value == '') {
                    strRet += "有效年 ";
                }
                /*if (tbxCard_Owner.value == '') {
                    strRet += "持卡人 ";
                }
                if (tbxOwner_IDNo.value == '') {
                    strRet += "持卡人身分證 ";
                }
                if (tbxRelation.value == '') {
                    strRet += "與捐款人關係 ";
                }
                if (tbxAuthorize.value == '') {
                    strRet += "授權碼 ";
                }*/
                if (strRet != "") {
                    strRet += "欄位不可為空白！";
                    alert(strRet)
                    return false;
                }
                if ($('#ddlDonate_Payment option:selected').text() == "信用卡授權書(一般)") {
                    if (ddlCard_Type.value == 'AE') {
                        strRet += "若授權方式為「信用卡授權書(一般)」，信用卡別不可為AE！";
                        alert(strRet)
                        return false;
                    }
                }
                cnt = 0;
                sName = tbxCard_Bank.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 20) {
                    alert('銀行別 欄位長度超過限制！');
                    return false;
                }
                cnt = 0;
                sName = ddlCard_Type.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('信用卡別 欄位長度超過限制！');
                    return false;
                }
                if (isNaN(Number(tbxAccount_No1.value)) == true) {
                    alert('信用卡號 欄位必須為數字！');
                    tbxAccount_No1.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxAccount_No2.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 4) {
                    alert('信用卡號 欄位長度超過限制！');
                    tbxAccount_No1.focus();
                    return false;
                }
                if (isNaN(Number(tbxAccount_No2.value)) == true) {
                    alert('信用卡號 欄位必須為數字！');
                    tbxAccount_No2.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxAccount_No2.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4 ) cnt += 2;
                    else cnt++;
                }
                if (cnt > 4) {
                    alert('信用卡號 欄位長度超過限制！');
                    tbxAccount_No2.focus();
                    return false;
                }
                if (isNaN(Number(tbxAccount_No3.value)) == true) {
                    alert('信用卡號 欄位必須為數字！');
                    tbxAccount_No3.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxAccount_No2.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 4) {
                    alert('信用卡號 欄位長度超過限制！');
                    tbxAccount_No3.focus();
                    return false;
                }

                if (isNaN(Number(tbxAccount_No4.value)) == true) {
                    alert('信用卡號 欄位必須為數字！');
                    tbxAccount_No4.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxAccount_No4.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 4) {
                    alert('信用卡號 欄位長度超過限制！');
                    tbxAccount_No4.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxCard_Owner.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 40) {
                    alert('持卡人 欄位長度超過限制！');
                    return false;
                }
                cnt = 0;
                sName = tbxOwner_IDNo.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('持卡人身分證 欄位長度超過限制！');
                    return false;
                }
                cnt = 0;
                sName = tbxRelation.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('與捐款人關係 欄位長度超過限制！');
                    return false;
                }
            }
            if ($('#ddlDonate_Payment option:selected').text() == "郵局帳戶轉帳授權書") {
                if (tbxPost_Name.value == "") {
                    strRet += "存簿戶名 ";
                }
                if (tbxPost_IDNo.value == "") {
                    strRet += "持有人身分證 ";
                }
                if (tbxPost_SavingsNo.value == "") {
                    strRet += "存簿局號 ";
                }
                if (tbxPost_AccountNo.value == "") {
                    strRet += "存簿帳號 ";
                }
                if (strRet != "") {
                    strRet += "欄位不可為空白！";
                    alert(strRet)
                    return false;
                }
                cnt = 0;
                sName = tbxPost_Name.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 40) {
                    alert('存簿戶名 欄位長度超過限制！');
                    return false;
                }
                cnt = 0;
                sName = tbxPost_IDNo.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('持有人身分證 欄位長度超過限制！');
                    return false;
                }
                if (tbxPost_SavingsNo.value == '') {
                    alert('存簿局號 欄位不可為空白！');
                    tbxPost_SavingsNo.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxPost_SavingsNo.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('存簿局號 欄位長度超過限制！');
                    return false;
                }
                if (isNaN(Number(tbxPost_SavingsNo.value)) == true) {
                    alert('存簿局號 欄位必須為數字！');
                    tbxPost_SavingsNo.focus();
                    return false;
                }
                if (tbxPost_AccountNo.value == '') {
                    alert('存簿帳號 欄位不可為空白！');
                    tbxPost_AccountNo.focus();
                    return false;
                }
                cnt = 0;
                sName = tbxPost_AccountNo.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('存簿帳號 欄位長度超過限制！');
                    return false;
                }
                if (isNaN(Number(tbxPost_AccountNo.value)) == true) {
                    alert('存簿帳號 欄位必須為數字！');
                    tbxPost_AccountNo.focus();
                    return false;
                }
            }
            if ($('#ddlDonate_Payment option:selected').text() == "ACH轉帳授權書") {
                if (tbxP_BANK.value == "") {
                    strRet += "收受行代號 ";
                }
                if (tbxP_RCLNO.value == "") {
                    strRet += "收受者帳號 ";
                }
                if (tbxP_PID.value == "") {
                    strRet += "收受者身分證/統編 ";
                }
                if (strRet != "") {
                    strRet += "欄位不可為空白！";
                    alert(strRet)
                    return false;
                }
                cnt = 0;
                sName = tbxP_BANK.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 7) {
                    alert('收受行代號 欄位長度超過限制！');
                    return false;
                }
                cnt = 0;
                sName = tbxP_RCLNO.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 14) {
                    alert('收受者帳號 欄位長度超過限制！');
                    return false;
                } 
                cnt = 0;
                sName = tbxP_PID.value;
                for (var i = 0; i < sName.length; i++) {
                    if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                    else cnt++;
                }
                if (cnt > 10) {
                    alert('收受者身分證/統編 欄位長度超過限制！');
                    return false;
                }
                if (isNaN(Number(tbxP_BANK.value)) == true) {
                    alert('收受行代號 欄位必須為數字！');
                    tbxP_BANK.focus();
                    return false;
                }
                if (isNaN(Number(tbxP_RCLNO.value)) == true) {
                    alert('收受者帳號 欄位必須為數字！');
                    tbxP_RCLNO.focus();
                    return false;
                }
            }
            cnt = 0;
            sName = tbxInvoice_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('收據抬頭 欄位長度超過限制！');
                return false;
            }
            cnt = 0; 
            sName = tbxAccoun_Bank.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('入帳銀行 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlAccounting_Title.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 30) {
                alert('會計科目 欄位長度超過限制！');
                return false;
            }
            return true;
        }
        //20140523 新增by Ian 增加信用卡 卡別和卡號的檢核機制
        function CheckCreditCardNumJ() {
            var Card_Type = document.getElementById('ddlCard_Type').value;
            var Account_No1 = document.getElementById('tbxAccount_No1').value;
            var Account_No2 = document.getElementById('tbxAccount_No2').value;
            var Account_No3 = document.getElementById('tbxAccount_No3').value;
            var Account_No4 = document.getElementById('tbxAccount_No4').value;
            if (Card_Type == '銀聯' || Card_Type == '聯合信用卡') {
                return true;
            }
            var num = Account_No1 + Account_No2 + Account_No3 + Account_No4;
            var len = num.length;
            var sum = 0;
            var mul = 1;
            for (i = (len - 1); i >= 0; --i) {
                numAtI = parseInt(num.charAt(i));
                tempSum = numAtI * mul;
                if (tempSum > 9) {
                    first = tempSum % 10;
                    tempSum = 1 + first;
                }
                sum += tempSum;
                if (mul == 1) mul++;
                else mul--;
            }
            if (sum % 10) {
                if (len == 15) {
                    txtcard = CheckFor15(num, 0);
                    if (txtcard == 'invalid') {
                        alert('此卡號無效!!');
                        return false;
                    }
                }
                else {
                    alert('此卡號無效!!');
                    return false;
                }
            }
            switch (len) {
                case 13:
                    txtcard = CheckFor13(num);
                    break;
                case 14:
                    txtcard = CheckFor14(num);
                    break;
                case 15:
                    txtcard = CheckFor15(num, 1);
                    break;
                case 16:
                    txtcard = CheckFor16(num);
                    break;
                default:
                    txtcard = 'invalid';
            }
            if (txtcard == 'invalid') {
                alert('此卡號無效!!');
                return false;
            }
            else {
                if (Card_Type == txtcard) {
                    alert('此卡別及卡號正確!!');
                    return true;
                }
                else {
                    alert('此卡號正確!!但卡片應該是：' + txtcard);
                    return false;
                }
            }
        }
        //20140523 新增by Ian Checks for Visa
        function CheckFor13(number) {
            if (parseInt(number.charAt(0)) == 4)
                return 'VISA';
            return 'invalid';
        }
        //20140523 新增by Ian Check for Diners Club
	    function CheckFor14(number){
		    if( parseInt(number.charAt(0)) != 3 )		
			    return 'invalid';		
		    if( (parseInt(number.charAt(1)) == 6) || (parseInt(number.charAt(1)) == 8) )		
			    return 'Diners Club/Carte Blanche';	
		    if( !(parseInt(number.charAt(1)) ))		
			    if( (parseInt(number.charAt(2)) >= 0) && (parseInt(number.charAt(2)) <= 5))	
				    return 'Diners Club/Carte Blanche';
		    return 'invalid';		
	    }
	    //20140523 新增by Ian Check for American Express, enRoute, JCB
	    function CheckFor15(number, chec) {
	        if ((number.charAt(0) == 3) && ((number.charAt(1) == 4) || (number.charAt(1) == 7)) && chec)
	            return 'AE';
	        var FirstFour = parseInt(number.charAt(0)) * 1000;
	        FirstFour += parseInt(number.charAt(1)) * 100;
	        FirstFour += parseInt(number.charAt(2)) * 10;
	        FirstFour += parseInt(number.charAt(3));
	        if ((FirstFour == 2014) || (FirstFour == 2149))
	            return 'enRoute';
	        if (((FirstFour == 2131) || (FirstFour == 1800)) && chec)
	            return 'JCB';
	        return 'invalid';
	    }
        //20140523 新增by Ian Check for Visa, MasterCard, Discover or JCB
	    function CheckFor16(number) {
	        var a = parseInt(number.charAt(0));
	        var b = parseInt(number.charAt(1));
	        switch (a) {
	            case 5:
	                if ((b > 0) && (b < 6))
	                    return 'MASTER';
	                else
	                    return 'invalid';
	                break;
	            case 4:
	                return 'VISA';
	                break;
	            case 6:
	                if ((b == 0) && (parseInt(number.charAt(2)) == 1) && (parseInt(number.charAt(3)) == 1))
	                    return 'Discover';
	                else
	                    return 'invalid';
	                break;
	            case 3:
	                return 'JCB';
	                break;
	            default:
	                return 'invalid';
	                break;
	        }
	    }
        //20140807 首次扣款日與授權期間起日連動  by 詩儀
	    function SameDate() {
	        var tbxDonate_FromDate = document.getElementById('tbxDonate_FromDate');
	        var tbxNext_DonateDate = document.getElementById('tbxNext_DonateDate');
	        tbxNext_DonateDate.value = tbxDonate_FromDate.value;
	    }
    </script>
    </head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <asp:HiddenField runat="server" ID="HFD_Create_DateTime" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        固定轉帳授權書【新增】 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1">
                捐款人：
            </th>
            <td align="left" colspan="1" >
                <asp:TextBox runat="server" ID="tbxDonor_Name" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right"colspan="1" >
                捐款人編號：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Id" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox runat="server" ID="tbxAddress" CssClass="font9"  Width="250px" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分別：
            </th>
            <td align="left" colspan="1">
               <asp:TextBox runat="server" ID="tbxDonor_Type" CssClass="font9"
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_Type" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" BackColor="#FFE1AF" 
                    ReadOnly="True"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                最近捐款日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxLast_DonateDate" CssClass="font9" 
                    BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td>
        </tr>
        
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權方式：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Payment" CssClass="font9" 
                    AutoPostBack="True" 
                    onselectedindexchanged="ddlDonate_Payment_SelectedIndexChanged"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                指定用途：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Purpose" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                每次扣款：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxDonate_Amt" CssClass="font9"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                授權期間：
            </th>
            <td align="left" colspan="3">
                <asp:TextBox ID="tbxDonate_FromDate" runat="server" onchange="CheckDateFormat(this, '授權期間');" OnPropertyChange="SameDate()"></asp:TextBox>
                <img id="imgDonate_FromDate" alt="" src="../images/date.gif" />
                ~
                <asp:TextBox ID="tbxDonate_ToDate" runat="server" onchange="CheckDateFormat(this, '授權期間');" AutoPostBack="true" OnTextChanged="tbxDonate_ToDate_Changed"></asp:TextBox>
                <img id="imgDonate_ToDate" alt="" src="../images/date.gif" />
            </td>
            <th align="right" colspan="1">
                轉帳週期：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDonate_Period" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                首次扣款日：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox ID="tbxNext_DonateDate" runat="server" onchange="CheckDateFormat(this, '首次扣款日');"></asp:TextBox>
                <img id="imgNext_DonateDate" alt="" src="../images/date.gif" />
            </td> 
            
        </tr>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <asp:Panel ID="PanelCreditCard" runat="server" Visible ="false">
        <tr>
            <th align="right" colspan="1">
                銀行別：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Bank" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                信用卡別：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlCard_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                信用卡號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccount_No1" CssClass="font9" Width="28px"  OnKeyup="Account_No_Keyup('1');" MaxLength="4"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No2" CssClass="font9" Width="28px"  OnKeyup="Account_No_Keyup('2');" MaxLength="4"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No3" CssClass="font9" Width="28px"  OnKeyup="Account_No_Keyup('3');" MaxLength="4"></asp:TextBox>-
                <asp:TextBox runat="server" ID="tbxAccount_No4" CssClass="font9" Width="28px" onblur="CheckCreditCardNumJ();" MaxLength="4"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                有效月年：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlMonth_Valid_Date" CssClass="font9"></asp:dropdownlist> / 
                <asp:dropdownlist runat="server" ID="ddlYear_Valid_Date" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                持卡人：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxCard_Owner" CssClass="font9" ></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner" Text="同捐款人" CssClass="font9"  OnClick="Same_Owner_OnClick()"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持卡人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxOwner_IDNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                與捐款人關係：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxRelation" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                末三碼(CVV)：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAuthorize" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelAccount" runat="server"  Visible ="false">
        <tr>
             <th align="right" colspan="1">
                存簿戶名：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_Name" CssClass="font9" ></asp:TextBox>
                <asp:checkbox runat="server" ID="cbxSame_Owner2" Text="同捐款人" CssClass="font9"  OnClick="Same_Owner2_OnClick()"></asp:checkbox>
            </td>
            <th align="right" colspan="1">
                持有人身分證：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_IDNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿局號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_SavingsNo" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                存簿帳號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxPost_AccountNo" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <asp:Panel ID="PanelACH" runat="server"  Visible ="false">
        <tr>
            <th align="right" colspan="1">
                收受行代號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxP_BANK" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收受者帳號：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxP_RCLNO" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                收受者身分證/統編：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxP_PID" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        </asp:Panel>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right"colspan="1" >
                機構：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlDept" CssClass="font9"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1">
                收據開立：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlInvoice_Type" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                入帳銀行：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxAccoun_Bank" CssClass="font9" BackColor="#FFE1AF" ReadOnly="True"></asp:TextBox>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                會計科目：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlAccounting_Title" CssClass="font9"></asp:dropdownlist>
            </td> 
            <th align="right" colspan="1">
                募款活動：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlAct_Id" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th colspan="1" align="right" valign="top">
                備註：
            </th>
            <td colspan="3" valign="top">
                <asp:Textbox runat="server" ID="tbxComment" CssClass="font9" 
                    TextMode="MultiLine" Width="450px" Height="100px"></asp:Textbox> 
            </td>
            <td colspan="4">
               <table width="530" border="0" cellpadding="0" cellspacing="1" class="table_v">
                    <tr>
                        <th colspan="3"><font color="blue">奉獻動機與收看管道調查</font></th>
                    </tr>
                    <tr>
                        <th align="center" style="width:150px;">奉獻動機(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="DonateMotive1" runat="server" Text="支持媒體宣教大平台，可廣傳福音" /><br />
                            <asp:CheckBox ID="DonateMotive2" runat="server" Text="個人靈命得造就" /><br />
                            <asp:CheckBox ID="DonateMotive3" runat="server" Text="支持優質節目製作" /><br />
                            <asp:CheckBox ID="DonateMotive4" runat="server" Text="支持GOOD TV家庭事工" /><br />
                            <%--<asp:CheckBox ID="DonateMotive5" runat="server" Text="抵扣稅額" /><br />--%>
                            <asp:CheckBox ID="DonateMotive5" runat="server" Text="感恩奉獻" /><br />
                            <asp:CheckBox ID="DonateMotive6" runat="server" Text="其他" />(如有請寫入〝想對GOODTV說的話〞)
                        </td>
                        </tr>
                        <tr>
                        <th align="center" style="width:150px;">收看管道(可複選)：</th>
                        <td>
                            <asp:CheckBox ID="WatchMode1" runat="server" Text="GOOD TV電視頻道" />
                            <br />
                            <asp:CheckBox ID="WatchMode9" runat="server" Text="報章雜誌" />
                            <br />新媒體平台：
                            <asp:CheckBox ID="WatchMode2" runat="server" Text="官網" />
                            <asp:CheckBox ID="WatchMode3" runat="server" Text="Facebook" />
                            <asp:CheckBox ID="WatchMode4" runat="server" Text="Youtube" />
                            <br />平面：　　　
                            <asp:CheckBox ID="WatchMode5" runat="server" Text="好消息月刊" />
                            <asp:CheckBox ID="WatchMode6" runat="server" Text="GOOD TV簡介刊物" />
                            <br />親友推薦：　
                            <asp:CheckBox ID="WatchMode7" runat="server" Text="教會牧者" />
                            <asp:CheckBox ID="WatchMode8" runat="server" Text="親友" />
                        </td>
                    </tr>
                   <tr>
                        <th align="right" nowrap="nowrap">想對GOOD TV說的話：
                        </th>
                        <td>
                            <asp:TextBox ID="txtToGoodTV" runat="server" Width="90%" TextMode="MultiLine" Rows="5" Font-Size="16px"></asp:TextBox>
                        </td>
                    </tr>
                </table>         
            </td>
        </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" 
            Text="存檔" OnClientClick= "return CheckFieldMustFillBasic();" onclick="btnAdd_Click"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="取消" onclick="btnExit_Click"/>
    </div>
    </form>
</body>
</html>
