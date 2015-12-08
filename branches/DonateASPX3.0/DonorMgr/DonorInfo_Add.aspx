<%@ Page Language="C#" AutoEventWireup="true" CodeFile="DonorInfo_Add.aspx.cs" Inherits="DonorMgr_DonorInfo_Add" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">    
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"/>
    <title>捐款人基本資料維護【新增】</title>
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
            //20151029新增　文宣品　全選/全清除
            $('#cbxAll').bind('change', function (e) {
                if ($('#cbxAll').prop('checked') == true) {
                    $("#cbxIsDVD").attr("checked", "checked");
                    $("#cbxIsSendEpaper").attr("checked", "checked");
                    $("#cbxIsGift").attr("checked", "checked");
                    $("#cbxIsBigAmtThank").attr("checked", "checked");
                    $("#cbxIsPost").attr("checked", "checked");
                    $("#cbxClear").removeAttr("checked");
                }
            });
            $('#cbxClear').bind('change', function (e) {
                if ($('#cbxClear').prop('checked') == true) {
                    $("#cbxIsDVD").removeAttr("checked");
                    $("#cbxIsSendEpaper").removeAttr("checked");
                    $("#cbxIsGift").removeAttr("checked");
                    $("#cbxIsBigAmtThank").removeAttr("checked");
                    $("#cbxIsPost").removeAttr("checked");
                    $("#cbxAll").removeAttr("checked");
                }
            });
        });
    </script>
    <script type="text/javascript">
        window.onload = initCalendar;

        var ALP_STR = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
        var NUM_STR = "0123456789";
        var SEX_STR = "12";
        var MAX_COUNT = 999;

        function initCalendar() {
            Calendar.setup({
                inputField: "txtBirthday",   // id of the input field
                button: "imgBirthday"     // 與觸發動作的物件ID相同
            });
        }
        function CBL_SingleChoice(sender)
        {
            var container = sender.parentNode;
            if (container.tagName.toUpperCase() == "TD")
            { // table 布局，否則為span布局
                container = container.parentNode.parentNode; // 層次: <table><tr><td><input />
            }
            var chkList = container.getElementsByTagName("input");
            var senderState = sender.checked;
            for (var i = 0; i < chkList.length; i++)
            {
                chkList[i].checked = false;
            }
            sender.checked = senderState;
        }
        function CheckFieldMustFillBasic() {
            var strRet = "";
            var cnt = 0;
            var sName = "";
            var txtDonor_Name = document.getElementById('txtDonor_Name');
            //var cblDonor_Type = document.getElementById('cblDonor_Type');
            var ddlInvoice_Type = document.getElementById('ddlInvoice_Type');
            var ddlSex = document.getElementById('ddlSex'); ddlTitle
            var ddlTitle = document.getElementById('ddlTitle');
            var tbxIDNo = document.getElementById('tbxIDNo');
            var tbxCellular_Phone = document.getElementById('tbxCellular_Phone');
            var tbxTel_Office = document.getElementById('tbxTel_Office');
            var tbxFax = document.getElementById('tbxFax');
            var tbxEMail = document.getElementById('tbxEMail');
            var tbxContactor = document.getElementById('tbxContactor');
            //var tbxJobTitle = document.getElementById('tbxJobTitle');
            var cbxIsLocal = document.getElementById('cbxIsLocal');
            var ddlCity = document.getElementById('ddlCity');
            //var ddlArea = document.getElementById('ddlArea');
            var tbxStreet = document.getElementById('tbxStreet');
            var tbxLane = document.getElementById('tbxLane');
            var tbxAlley = document.getElementById('tbxAlley');
            var tbxNo1 = document.getElementById('tbxNo1');
            var tbxNo2 = document.getElementById('tbxNo2');
            var tbxFloor1 = document.getElementById('tbxFloor1');
            var tbxFloor2 = document.getElementById('tbxFloor2');
            var tbxRoom = document.getElementById('tbxRoom');
            var cbxIsAbroad = document.getElementById('cbxIsAbroad');
            var ddlOverseasCountry = document.getElementById('ddlOverseasCountry');
            var tbxOverseasAddress = document.getElementById('tbxOverseasAddress');
            var tbxInvoice_Street = document.getElementById('tbxInvoice_Street');
            var tbxInvoice_Lane = document.getElementById('tbxInvoice_Lane');
            var tbxInvoice_Alley0 = document.getElementById('tbxInvoice_Alley0');
            var tbxInvoice_No1 = document.getElementById('tbxInvoice_No1');
            var tbxInvoice_No2 = document.getElementById('tbxInvoice_No2');
            var tbxInvoice_Floor1 = document.getElementById('tbxInvoice_Floor1');
            var tbxInvoice_Floor2 = document.getElementById('tbxInvoice_Floor2');
            var tbxInvoice_Room = document.getElementById('tbxInvoice_Room');
            var ddlInvoice_OverseasCountry = document.getElementById('ddlInvoice_OverseasCountry');
            var tbxInvoice_OverseasAddress = document.getElementById('tbxInvoice_OverseasAddress');
            var tbxInvoice_Title = document.getElementById('tbxInvoice_Title');
            var tbxInvoice_IDNo = document.getElementById('tbxInvoice_IDNo');
            var tbxIsSendNewsNum = document.getElementById('tbxIsSendNewsNum');
           
            if (txtDonor_Name.value == '') {
                strRet += "捐款人 ";
            }
//            if (cblDonor_Type.value == '') {
//                strRet += "身份別 ";
//            }
  
            if (ddlInvoice_Type.value == '') {
                strRet += "收據開立 ";
            }
            // 2014/4/9 修改成與APS 2.0版一樣，收據抬頭無輸入資料以捐款人資料存入
            /*
            if (ddlInvoice_Type.value != '1' && tbxInvoice_Title.value == '') {
                strRet += "收據抬頭 ";
            }
            */
            if (cbxIsLocal.checked) {
                if (ddlCity.value == '縣 市') {
                    strRet += "通訊地址-台灣本島(縣市) ";
                }
                //if (ddlArea.value == '鄉鎮市區') {
                //    strRet += "'通訊地址-台灣本島(鄉鎮市區) ";
                //}
                if (tbxStreet.value == '') {
                    strRet += "通訊地址-台灣本島(大道/路/街/部落) ";
                }
            }
            if (cbxIsAbroad.checked) {
                if (ddlOverseasCountry.value == '') {
                    strRet += "通訊地址-海外地址(國家/省城市/區) ";
                }
                if (tbxOverseasAddress.value == '') {
                    strRet += "通訊地址-海外地址(地址) ";
                }
            }
            if (cbxIsLocal.checked == false && cbxIsAbroad.checked == false) {
                strRet += "請勾選通訊地址 - 台灣本島或是海外地址 ";
            }
            if (strRet != "") {
                strRet += "欄位不得為空白！"
                alert(strRet);
                return false;
            }


            cnt = 0;
            sName = txtDonor_Name.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('捐款人 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlSex.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 2) {
                alert('性別 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = ddlTitle.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('稱謂 欄位長度超過限制！');
                return false;
            }
            // 2014/6/25 判斷身份別 by Samson
            var resGN_DonorType = false;
            $("input[name='GN_DonorType']").each(function() {
                if ($(this).prop("checked")) resGN_DonorType = true;
            });
            if (!resGN_DonorType)
            {
                alert("身分別 欄位至少需勾選一個！");
                return false;
            }
            cnt = 0;
            sName = tbxIDNo.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('身分證/統編 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxCellular_Phone.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('手機 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxTel_Office.value;
            for ( i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('電話(日) 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxFax.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('傳真 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxEMail.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 80) {
                alert('E-Mail 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxContactor.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('聯絡人 欄位長度超過限制！');
                return false;
            }
//            cnt = 0;
//            sName = tbxJobTitle.value;
//            for (var i = 0; i <sName.length; i++) {
//                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
//                else cnt++;
//            }
//            if (cnt > 40) {
//                alert('職稱 欄位長度超過限制！');
//                return false;
//            }
             cnt = 0;
            sName = tbxStreet.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('大道/路/街/部落 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxLane.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('巷 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxAlley.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('弄/衖 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxNo1.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxNo2.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號之 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxFloor1.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxFloor2.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓之 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxRoom.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('室 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
             sName = ddlOverseasCountry.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('通訊地址國家/省城市/區 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxOverseasAddress.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('通訊地址 欄位長度超過限制！');
                return false;
            }
           
            cnt = 0;
            sName = ddlInvoice_Type.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 20) {
                alert('收據開立 欄位長度超過限制！');
                return false;
            }       
            cnt = 0;
            sName = tbxInvoice_Street.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 40) {
                alert('大道/路/街/部落 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_Lane.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('巷 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxInvoice_Alley0.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('弄/衖 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_No1.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_No2.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('號之 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxInvoice_Floor1.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
             sName = tbxInvoice_Floor2.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('樓之 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
            sName = tbxInvoice_Room.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('室 欄位長度超過限制！');
                return false;
            }
             cnt = 0;
             sName = ddlInvoice_OverseasCountry.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 50) {
                alert('收據地址國家/省城市/區 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_OverseasAddress.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('收據地址 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_Title.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 100) {
                alert('收據抬頭 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxInvoice_IDNo.value;
            for (var i = 0; i <sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 10) {
                alert('收據身分證/統編 欄位長度超過限制！');
                return false;
            }
            cnt = 0;
            sName = tbxIsSendNewsNum.value;
            for (var i = 0; i < sName.length; i++) {
                if (escape(sName.charAt(i)).length >= 4) cnt += 2;
                else cnt++;
            }
            if (cnt > 3) {
                alert('紙本月刊本數 欄位長度不得超過3位數！');
                return false;
            }

            if (isNaN(Number(tbxCellular_Phone.value)) == true) {
                alert('手機 欄位必須為數字！');
                tbxCellular_Phone.focus();
                return false;
            }
            if (isNaN(Number(tbxTel_Office.value)) == true) {
                alert('電話 欄位必須為數字！');
                tbxTel_Office.focus();
                return false;
            }
            if (isNaN(Number(tbxFax.value)) == true) {
                alert('傳真 欄位必須為數字！');
                tbxFax.focus();
                return false;
            }
            if (isNaN(Number(tbxIsSendNewsNum.value)) == true) {
                alert('紙本月刊本數 欄位必須為數字！');
                tbxIsSendNewsNum.focus();
                return false;
            }
            //-----------------
            //2014/08/01 修改 by Samson 身分證或是統編驗證
            var IDNo = document.getElementById('tbxIDNo').value;
            if (IDNo != "" && IDNo.length == 8) {//統編規則
                var cx = new Array;
                cx[0] = 1;
                cx[1] = 2;
                cx[2] = 1;
                cx[3] = 2;
                cx[4] = 1;
                cx[5] = 2;
                cx[6] = 4;
                cx[7] = 1;
                var SUM = 0;
                var cnum = IDNo.split("");
                for (i = 0; i <= 7; i++) {
                    if (IDNo.charCodeAt() < 48 || IDNo.charCodeAt() > 57) {
                        alert("統一編號，要有 8 個 0-9 數字組合");
                        return false;
                    }
                    SUM += cc(cnum[i] * cx[i]);
                }
                if (SUM % 10 != 0 && cnum[6] != 7 && (SUM + 1) % 10 != 0) {
                    alert("統一編號：" + IDNo + " 錯誤!");
                    return false;
                }
            }
            else if (IDNo != "" && IDNo.length == 10) {//身分證規則

                sPID = IDNo.toUpperCase();
                var iPIDLen = String(sPID).length;
                var sChk = ALP_STR + NUM_STR;
                for (i = 0; i < iPIDLen; i++) {
                    if (sChk.indexOf(sPID.substr(i, 1)) < 0) {
                        alert("這個身分證字號含有不正確的字元！");
                        return false;
                    }
                }

                if (ALP_STR.indexOf(sPID.substr(0, 1)) < 0) {
                    alert("身分證字號第 1 碼應為英文字母(A~Z)。");
                    return false;
                } else if ((sPID.substr(1, 1) != "1") && (sPID.substr(1, 1) != "2")) {
                    alert("身分證字號第 2 碼應為數字(1~2)。");
                    return false;
                } else {
                    for (var i = 2; i < iPIDLen; i++) {
                        if (NUM_STR.indexOf(sPID.substr(i, 1)) < 0) {
                            alert("第 " + (i + 1) + " 碼應為數字(0~9)。");
                            return false;
                        }
                    }
                }

                var iChkNum = getPID_SUM(sPID);

                if (iChkNum % 10 != 0) {
                    var iLastNum = sPID.substr(9, 1) * 1;
                    for (i = 0; i < 10; i++) {
                        var xRightAlpNum = iChkNum - iLastNum + i;
                        if ((xRightAlpNum % 10) == 0) {
                            alert("身分證號碼錯誤！");
                            return false;
                        }
                    }
                }

            }
            else if (IDNo != "" && (IDNo.length != 8 || IDNo.length != 10)) {
                alert("請填入八碼統一編號或是十碼身分證號碼");
                return false;
            }
            else {
                //目前非必填，不需驗證 
            }
            
            //身份證字號檢查器 - 累加檢查碼
            function getPID_SUM(sPID) {
                var iChkNum = 0;

                // 第 1 碼
                iChkNum = ALP_STR.indexOf(sPID.substr(0, 1)) + 10;
                iChkNum = Math.floor(iChkNum / 10) + (iChkNum % 10 * 9);

                // 第 2 - 9 碼
                for (var i = 1; i < sPID.length - 1; i++) {
                    iChkNum += sPID.substr(i, 1) * (9 - i);
                }

                // 第 10 碼
                iChkNum += sPID.substr(9, 1) * 1;

                return iChkNum;
            }

            // 2014/6/25 增加檢查相同通訊地址
            var resAddress;
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=8' + '&IsAbroad=' + $('#cbxIsLocal').prop("checked") + '&City=' + $('#ddlCity').val() + '&Street=' + $('#tbxStreet').val()
                 + '&Section=' + $('#ddlSection').val() + '&Lane=' + $('#tbxLane').val() + '&Alley=' + $('#tbxAlley').val()
                 + '&HouseNo=' + $('#tbxNo1').val() + '&HouseNoSub=' + $('#tbxNo2').val() + '&Floor=' + $('#tbxFloor1').val()
                 + '&FloorSub=' + $('#tbxFloor2').val() + '&Room=' + $('#tbxRoom').val() //+ '&Attn=' + $('#tbxAttn').val()
                 + '&OverseasCountry=' + $('#ddlOverseasCountry').val() + '&OverseasAddress=' + $('#tbxOverseasAddress').val(),
                async: false,
                success: function (result) {
                    resAddress = result;
                },
                error: function () { alert('ajax failed'); }
            });

            if (resAddress == "Y") {
                //if (!confirm('已有相同通訊地址存在！\n您確定要繼續嗎？') )
                resAddress = "已有相同通訊地址存在！\n";
            }
            else if (resAddress == "X") {
                alert('程式異常請聯絡資訊組人員');
                return false;
            }

            // 2014/6/25 增加檢查相同收據地址
            if (($('#cbxInvoice_IsLocal').prop("checked") && $('#ddlInvoice_City').val() != "縣 市" && $('#tbxInvoice_Street').val() != "")
                || (!$('#cbxInvoice_IsLocal').prop("checked") && ($('#tbxInvoice_OverseasCountry').val() != "" || $('#tbxInvoice_OverseasAddress').val() != ""))){
                var resInvoiceAddress;
                $.ajax({
                    type: 'post',
                    url: "../common/ajax.aspx",
                    data: 'Type=9' + '&IsAbroad=' + $('#cbxInvoice_IsLocal').prop("checked") + '&City=' + $('#ddlInvoice_City').val() + '&Street=' + $('#tbxInvoice_Street').val()
                     + '&Section=' + $('#ddlInvoice_Section').val() + '&Lane=' + $('#tbxInvoice_Lane').val() + '&Alley=' + $('#tbxInvoice_Alley0').val()
                     + '&HouseNo=' + $('#tbxInvoice_No1').val() + '&HouseNoSub=' + $('#tbxInvoice_No2').val() + '&Floor=' + $('#tbxInvoice_Floor1').val()
                     + '&FloorSub=' + $('#tbxInvoice_Floor2').val() + '&Room=' + $('#tbxInvoice_Room').val() //+ '&Attn=' + $('#tbxInvoice_Attn').val()
                     + '&OverseasCountry=' + $('#ddlInvoice_OverseasCountry').val() + '&OverseasAddress=' + $('#tbxInvoice_OverseasAddress').val(),
                    async: false,
                    success: function (result) {
                        resInvoiceAddress = result;
                    },
                    error: function () { alert('ajax failed'); }
                });

                if (resInvoiceAddress == "Y") {
                    //if (!confirm('已有相同收據地址存在！\n您確定要繼續嗎？') )
                    resAddress += "已有相同收據地址存在！\n";
                }
                else if (resInvoiceAddress == "X") {
                    alert('程式異常請聯絡資訊組人員');
                    return false;
                }
            }
            if (resAddress != "") {
                if (!confirm(resAddress+"您確定要新增嗎？") )
                return false;
            }

        }

        function cc(n) {
            if (n > 9) {
                var s = n + "";
                n1 = s.substring(0, 1) * 1;
                n2 = s.substring(1, 2) * 1;
                n = n1 + n2;
            }
            return n;
        }
        //------------------
        function CheckDonorNameLen(Type, objValue) {
            var regx = /^[u4E00-u9FA5]+$/;
            if (Type == '1') {
                if (!regx.test(objValue)) {
                    if (objValue.length > 11)
                        alert("「捐款人」中文字超過11字，列印收據會被截斷!");
                }
                else {
                    if (objValue.length > 22)
                        alert("「捐款人」英數字超過22字，列印收據會被截斷!");
                }
            }
            else if (Type == '2') {
                if (!regx.test(objValue)) {
                    if (objValue.length > 20)
                        alert("「收據抬頭」中文字超過20字，列印收據會被截斷!");
                }
                else {
                    if (objValue.length > 40)
                        alert("「收據抬頭」英數字超過40字，列印收據會被截斷!");
                }
            }
        }
        //20140812 身份證與收據身份證連動  by 詩儀
        function SameID() {
            var tbxIDNo = document.getElementById('tbxIDNo');
            var tbxInvoice_IDNo = document.getElementById('tbxInvoice_IDNo');
            tbxInvoice_IDNo.value = tbxIDNo.value;
        }

    </script>
    <style type="text/css">
        .style1
        {
            width: 35px;
        }
    </style>
    </head>
<body class="body">
    <form id="Form1" runat="server">
    <div id="menucontrol">
        <a href="#"><img id="menu_hide" alt="" src="../images/menu_hide.gif" width="15" height="60" class="swapImage {src:'../images/menu_hide_over.gif'}" onclick="SwitchMenu();return false;" /></a>
        <a href="#"><img id="menu_show" alt="" src="../images/menu_show.gif" width="15" height="60" class="swapImage {src:'../images/menu_show_over.gif'}" onclick="SwitchMenu();return false;"/></a>
    </div>
    <div id="container">
    <asp:HiddenField runat="server" ID="HFD_Uid" />
    <h1 style="	padding-bottom:0px;">
        <img src="../images/h1_arr.gif" alt="" width="10" height="10" />
        捐款人基本資料維護【新增】 
    </h1>
    <table width="100%" border="0" align="center" cellpadding="0" cellspacing="1" class="table_v"> 
        <tr>
            <th align="right" colspan="1"><font color="red">*</font>
                捐款人：
            </th>
            <td align="left" colspan="2" class="style1">
                <asp:TextBox runat="server" ID="txtDonor_Name" CssClass="font9" Width="300px"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                性別：
            </th>
            <td align="left" colspan="2" >
                <asp:dropdownlist runat="server" ID="ddlSex" CssClass="font9" 
                    AutoPostBack="True" onselectedindexchanged="ddlSex_SelectedIndexChanged"></asp:dropdownlist>
            </td>
            <th align="right" colspan="1" >
                稱謂：
            </th>
            <td align="left" colspan="1">
                <asp:dropdownlist runat="server" ID="ddlTitle" CssClass="font9"></asp:dropdownlist>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1"><font color="red">*</font>
                身分別：
            </th>
            <td align="left" colspan="2">
                <%--<asp:CheckBoxList ID="cblDonor_Type" runat="server" RepeatDirection="Horizontal">
                </asp:CheckBoxList>--%>
                <%--20140509 新增 by Ian_Kao 可以多選的CheckBoxList--%>
                <asp:Label ID="lblCheckBoxList" runat="server"></asp:Label>
                <%--20140509 修改 by Ian_Kao--%>
                <%--20140425 新增 by Ian_Kao 另外新增驗證項來驗證是否身分別有填入--%>
                <%--<asp:CustomValidator ID="CustomValidator1" runat="server" 
                    ErrorMessage="" onservervalidate="CustomValidator1_ServerValidate"></asp:CustomValidator>--%>
            </td>
            <th align="right" colspan="1">
                其他：
            </th>
            <td colspan="4">
                <asp:checkbox runat="server" ID="cbxIsContact" Text="不主動聯絡" CssClass="font9"></asp:checkbox>　　
                <asp:checkbox runat="server" ID="cbxIsErrAddress" Text="地址錯誤" CssClass="font9"></asp:checkbox>　　
                <asp:CheckBox runat="server" ID="cbxIsAnonymous" Text="不刊登徵信錄" CssClass="font9" Checked="true"></asp:CheckBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                身分證/統編：
            </th>
            <td align="left" colspan="2" class="style1">
                <asp:TextBox runat="server" ID="tbxIDNo" CssClass="font9" OnPropertyChange="SameID()"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                出生日期：
            </th>
            <td align="left" colspan="4">
                <asp:TextBox ID="txtBirthday" runat="server" onchange="CheckDateFormat(this, '出生日期');"></asp:TextBox>
                <img id="imgBirthday" alt="" src="../images/date.gif" />
            </td>
            
            
        </tr>
        <tr>
            <th align="right" colspan="1">
                手機：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxCellular_Phone" CssClass="font9"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                電話：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxTel_Office_Loc" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxTel_Office" CssClass="font9"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxTel_Office_Ext" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                傳真：
            </th>
            <td align="left" colspan="1"">
                <asp:TextBox runat="server" ID="tbxFax_Loc" CssClass="font9" 
                    MaxLength="5" Width="35px"></asp:TextBox>&nbsp;-
                <asp:TextBox runat="server" ID="tbxFax" CssClass="font9"></asp:TextBox>
            </td>
        </tr>
        <tr>
            <th align="right" colspan="1">
                E-Mail：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxEMail" CssClass="font9" Width="300px"></asp:TextBox>
            </td>
            <th align="right" colspan="1">
                聯絡人：
            </th>
            <td align="left" colspan="4">
                <asp:TextBox runat="server" ID="tbxContactor" CssClass="font9" Width="200px"></asp:TextBox>
            </td>
            <%--20140127
            <th align="right" colspan="1">
                職稱：
            </th>
            <td align="left" colspan="1">
                <asp:TextBox runat="server" ID="tbxJobTitle" CssClass="font9"></asp:TextBox>
            </td>--%>
        </tr>
        <tr>
            <th align="right" colspan="1"><font color="red">*</font>
                通訊地址：
            </th>
             <td align="left" colspan="7" style="LINE-HEIGHT:30px;">
                <asp:CheckBox ID="cbxIsLocal" runat="server" text="台灣本島" Checked="true" 
                     AutoPostBack="True" oncheckedchanged="cbxIsLocal_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxZipCode" CssClass="font9" Width="60px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlCity" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlCity_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlArea" CssClass="font9" 
                     AutoPostBack="True" onselectedindexchanged="ddlArea_SelectedIndexChanged"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxStreet" CssClass="font9" Width="100px"></asp:TextBox>大道/路/街/部落
                <asp:dropdownlist runat="server" ID="ddlSection" CssClass="font9"></asp:dropdownlist>段
                <asp:TextBox runat="server" ID="tbxLane" CssClass="font9" Width="40px"></asp:TextBox>巷<asp:TextBox runat="server" ID="tbxAlley" CssClass="font9" Width="40px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlAlley" CssClass="font9"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxNo1" CssClass="font9" Width="55px"></asp:TextBox>號之
                <asp:TextBox runat="server" ID="tbxNo2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                <asp:TextBox runat="server" ID="tbxFloor1" CssClass="font9" Width="40px"></asp:TextBox>樓之
                <asp:TextBox runat="server" ID="tbxFloor2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                <asp:TextBox runat="server" ID="tbxRoom" CssClass="font9" Width="40px"></asp:TextBox>室
                Attn：<asp:TextBox runat="server" ID="tbxAttn" CssClass="font9" Width="180px"></asp:TextBox>
                <br />
                <asp:CheckBox ID="cbxIsAbroad" runat="server" text="海外地址" 
                     AutoPostBack="True" oncheckedchanged="cbxIsAbroad_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxOverseasAddress" CssClass="font9" Width="500px"></asp:TextBox>國家/省城市/區 
                <%--<asp:TextBox runat="server" ID="tbxOverseasCountry" CssClass="font9" Width="150px"></asp:TextBox>--%><asp:DropDownList ID="ddlOverseasCountry" runat="server"></asp:DropDownList>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                文宣品寄送：
            </th>
            <td align="left" colspan="3">
                紙本月刊本數<asp:Textbox runat="server" ID="tbxIsSendNewsNum" CssClass="font9" 
                    Width="50px" Text="1"></asp:Textbox>
                <asp:checkbox runat="server" ID="cbxIsDVD" Text="DVD" CssClass="font9" Checked="True"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsSendEpaper" Text="電子文宣" CssClass="font9" Checked="True"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsGift" Text="公關贈品" CssClass="font9" Checked="True"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsBigAmtThank" Text="大額謝卡" CssClass="font9"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxIsPost" Text="郵簡" CssClass="font9" Checked="True"></asp:checkbox>
            </td>
            <td colspan="4"><asp:checkbox runat="server" ID="cbxAll" Text="全選" CssClass="font9"></asp:checkbox>
                <asp:checkbox runat="server" ID="cbxClear" Text="全清除" CssClass="font9"></asp:checkbox>
            </td>
        </tr>
        <tr>
            <td colspan="8">

            </td>
        </tr>
        <tr>
            <th align="right" colspan="1"><font color="red">*</font>
                收據開立：
            </th>
            <td align="left" colspan="7">
                <asp:dropdownlist runat="server" ID="ddlInvoice_Type" CssClass="font9"></asp:dropdownlist>
                <asp:CheckBox ID="cbxInvoice_same" runat="server" CssClass="font9" 
                    AutoPostBack="True" Text="收據資料同上" ForeColor="Red" 
                    oncheckedchanged="cbxInvoice_same_CheckedChanged" />
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                收據抬頭：
            </th>
            <td align="left" colspan="2">
                <!--20140114 增加「字數」限制提醒-->
                <asp:TextBox runat="server" ID="tbxInvoice_Title" CssClass="font9" Width="300px" onblur="javascript:CheckDonorNameLen('2',this.value);"></asp:TextBox>
            </td> 
            <th align="right" colspan="1">
                收據稱謂：
            </th>
            <td align="left" colspan="4">
                <asp:dropdownlist runat="server" ID="ddlTitle2" CssClass="font9"></asp:dropdownlist>
            </td> 
        </tr>
        <tr>
            <th align="right" colspan="1">
                收據身分證/統編：
            </th>
            <td align="left" colspan="2">
                <asp:TextBox runat="server" ID="tbxInvoice_IDNo" CssClass="font9"></asp:TextBox>
            </td>
            <td colspan="5">
                <asp:CheckBox runat="server" ID="cbxIsFdc" CssClass="font9" Text="本捐款人願意將捐款資料上傳至國稅局以供報稅用" ForeColor="Red"></asp:CheckBox> ( 請填寫捐款人收據身分證/統編 )
            </td>
            <!--th align="right" colspan="1">
                徵信錄原則：
            </th>
            <td align="left" colspan="3">
                <asp:dropdownlist runat="server" ID="ddlIsAnonymous" CssClass="font9"></asp:dropdownlist>
            </td--> 
         </tr>
         <tr>
            <th align="right" colspan="1">
                收據地址：
            </th>
            <td align="left" colspan="7" style="LINE-HEIGHT:30px;">

                <asp:CheckBox ID="cbxInvoice_IsLocal" runat="server" text="台灣本島" Checked="true" 
                     AutoPostBack="True" oncheckedchanged="cbxInvoice_IsLocal_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxInvoice_ZipCode" CssClass="font9" Width="60px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlInvoice_City" CssClass="font9" 
                     AutoPostBack="True" 
                    onselectedindexchanged="ddlInvoice_City_SelectedIndexChanged"></asp:dropdownlist>
                <asp:dropdownlist runat="server" ID="ddlInvoice_Area" CssClass="font9" 
                     AutoPostBack="True" 
                    onselectedindexchanged="ddlInvoice_Area_SelectedIndexChanged"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxInvoice_Street" CssClass="font9" Width="100px"></asp:TextBox>大道/路/街/部落
                <asp:dropdownlist runat="server" ID="ddlInvoice_Section" CssClass="font9"></asp:dropdownlist>段
                <asp:TextBox runat="server" ID="tbxInvoice_Lane" CssClass="font9" Width="40px"></asp:TextBox>巷<asp:TextBox 
                    runat="server" ID="tbxInvoice_Alley0" CssClass="font9" Width="40px"></asp:TextBox>
                <asp:dropdownlist runat="server" ID="ddlInvoice_Alley" CssClass="font9"></asp:dropdownlist>
                <asp:TextBox runat="server" ID="tbxInvoice_No1" CssClass="font9" Width="55px"></asp:TextBox>號之
                <asp:TextBox runat="server" ID="tbxInvoice_No2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                <asp:TextBox runat="server" ID="tbxInvoice_Floor1" CssClass="font9" Width="40px"></asp:TextBox>樓之
                <asp:TextBox runat="server" ID="tbxInvoice_Floor2" CssClass="font9" Width="40px"></asp:TextBox>&nbsp;
                <asp:TextBox runat="server" ID="tbxInvoice_Room" CssClass="font9" Width="40px"></asp:TextBox>室
                Attn：<asp:TextBox runat="server" ID="tbxInvoice_Attn" CssClass="font9" Width="180px"></asp:TextBox>
                <br />
                <asp:CheckBox ID="cbxInvoice_IsAbroad" runat="server" text="海外地址" 
                     AutoPostBack="True" oncheckedchanged="cbxInvoice_IsAbroad_CheckedChanged"></asp:CheckBox>
                <asp:TextBox runat="server" ID="tbxInvoice_OverseasAddress" CssClass="font9" 
                    Width="500px"></asp:TextBox>國家/省城市/區 
                <%--<asp:TextBox runat="server" ID="tbxInvoice_OverseasCountry" CssClass="font9" Width="150px"></asp:TextBox>--%><asp:DropDownList ID="ddlInvoice_OverseasCountry" runat="server"></asp:DropDownList>
            </td>
        </tr>
         <tr>
            <th colspan="1" align="right">
                捐款人備註：
            </th>
            <td colspan="7">
                <asp:Textbox runat="server" ID="tbxRemark" CssClass="font9" 
                    TextMode="MultiLine" Width="600px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
         <tr>
            <th colspan="1" align="right">
                想對GOOD TV說的話：
            </th>
            <td colspan="7">
                <asp:Textbox runat="server" ID="tbxToGOODTV" CssClass="font9" 
                    TextMode="MultiLine" Width="600px" Height="100px"></asp:Textbox> 
            </td>
         </tr>
    </table>
    </div>
    <div class="function">
        <asp:Button ID="btnAdd" class="npoButton npoButton_New" runat="server" 
            Text="新增" onclick="btnAdd_Click" OnClientClick= "return CheckFieldMustFillBasic()"/>
        <asp:Button ID="btnExit" class="npoButton npoButton_Exit" runat="server" 
            Text="離開" onclick="btnExit_Click"/>
    </div>
    </form>
</body>
</html>
