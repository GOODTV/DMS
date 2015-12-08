//客製JS File
//----------------------------------------------------------------------------
//組合奉獻項目及內容
        function GetItemList() {
            var amtTotal = 0;
            if ($('#HFD_chkItem').val() != '') {	
				//Modify by GoodTV Tanya:修正部份IE text()不work問題，不使用text()
                //$('#HFD_chkItem').text($('#HFD_chkItem').val());					
                $('#HFD_ItemList').val($('#HFD_chkItem').val());
                var item = $('#HFD_chkItem').val().split(",");					
                var output = '';
                var period = '奉獻期間：';				
                $.each(item, function (key, value) {
                    var amt = 0;
                    switch (value) {
                        case '單筆奉獻':
                            amt = isNaN(parseInt($('#txtAmountOnce').val()), 10) ? 0 : parseInt($('#txtAmountOnce').val(), 10);
                            $('#txtAmountOnce').val(amt);
                            $('#HFD_ItemList').val($('#HFD_ItemList').val().replace(value, value + '|經常奉獻||' + amt));
                            output += '<tr><td>' + value + '：</td><td></td><td align="right">　　　　　' + amt + '元</td></tr>';
                            break;
                        case '定期定額奉獻':
                            amt = isNaN(parseInt($('#txtAmountPeriod').val()), 10) ? 0 : parseInt($('#txtAmountPeriod').val(), 10);
                            $('#txtAmountPeriod').val(amt);

                            if ($('#HFD_PayType').val() == '月繳') {
                                if ($('#ddlYearMS').val() != '請選擇') {
                                    period += $('#ddlYearMS').val() + '年';
                                }
                                if ($('#ddlMonthS').val() != '請選擇') {
                                    period += $('#ddlMonthS').val() + '月';
                                }
                                period += '～';
                                if ($('#ddlYearME').val() != '請選擇') {
                                    period += $('#ddlYearME').val() + '年';
                                }
                                if ($('#ddlMonthE').val() != '請選擇') {
                                    period += $('#ddlMonthE').val() + '月';
                                }
                                if ($('#HFD_PayType').val() != '') {
                                    period += '；繳費方式：' + $('#HFD_PayType').val();
                                }
                            }
                            if ($('#HFD_PayType').val() == '季繳') {
//                                if ($('#ddlYearQS').val() != '請選擇') {
//                                    period += $('#ddlYearQS').val() + '年';
//                                }
//                                if ($('#ddlQtrS').val() != '請選擇') {
//                                    period += '第' + $('#ddlQtrS').val() + '季';
//                                }
//                                period += '～';
//                                if ($('#ddlYearQE').val() != '請選擇') {
//                                    period += $('#ddlYearQE').val() + '年';
//                                }
//                                if ($('#ddlQtrE').val() != '請選擇') {
//                                    period += '第' + $('#ddlQtrE').val() + '季';
                                //                                }
                                
                                if ($('#HFD_PayType').val() != '') {
                                    period += '；繳費方式：' + $('#HFD_PayType').val();
                                }
                            }
                            if ($('#HFD_PayType').val() == '年繳') {
                                if ($('#ddlYearS').val() != '請選擇') {
                                    period += $('#ddlYearS').val() + '年';
                                }
                                period += '～';
                                if ($('#ddlYearE').val() != '請選擇') {
                                    period += $('#ddlYearE').val() + '年';
                                }
                                if ($('#HFD_PayType').val() != '') {
                                    period += '；繳費方式：' + $('#HFD_PayType').val();
                                }
                            }
                            $('#HFD_ItemList').val($('#HFD_ItemList').val().replace(value, value + '|經常奉獻|' + period + '|' + amt));
                            output += '<tr><td>' + value + '：</td><td>' + period + '</td><td align="right">　　　　　' + amt + '元</td></tr>';
                            break;
                    }
                    amtTotal += amt;
                });

                $('#HFD_ItemList').val($('#HFD_ItemList').val());
                //output += '<tr><td> 　　　 </td><td align="right">　　　　－－－－－－－－－</td></tr>';
                //output += '<tr><td style="font-weight: bold"> 合計： </td><td align="right" style="font-weight: bold">　　新台幣：　' + amtTotal + '元</td></tr>';

                var title = '<tr><td colspan="3" align="center" style="font-weight: bold"> <br/>奉獻項目明細：<br/></td></tr>';

                $('#btnNext').show();
                return '<table align="center" border="0" cellpadding="0" cellspacing="0" width="60%">' + title + output + '</table><br/>';
            }
            else {
                $('#btnNext').hide();
                return '';
            }
        }

//----------------------------------------------------------------------------
//check out 時檢核奉獻資料
        function CheckAmount() {
            if (!$('#rdoPurpose1').prop('checked') && !$('#rdoPurpose2').prop('checked')) {
                alert('請選擇奉獻用途！');
                return false;
            }
//            if ($('#CHK_ItemOnce').attr('checked') && $('#txtAmountOnce').val() <= 0) {
            if ($('#RDO_Item1').attr('checked') && $('#txtAmountOnce').val() < 100) {
//                alert('【' + $('#CHK_ItemOnce').val() + '】金額有誤，請重新輸入！');
                alert('【' + $('#RDO_Item1').val() + '】最低金額為100元，請重新輸入！');
                $('#txtAmountOnce').focus();
                return false;
            }
//            if ($('#CHK_ItemPeriod').attr('checked')) {
            if ($('#RDO_Item2').attr('checked')) {
                if ($('#txtAmountPeriod').val() < 100) {
//                    alert('【' + $('#CHK_ItemPeriod').val() + '】金額有誤，請重新輸入！');
                    alert('【' + $('#RDO_Item2').val() + '】最低金額為100元，請重新輸入！');
                    $('#txtAmountPeriod').focus();
                    return false;
                }
                if ($('#HFD_PayType').val() == '') {
                    alert('請選擇繳費方式！');
                    return false;
                }
                //月繳檢核
                if ($('#HFD_PayType').val() == '月繳') {
                    if ($('#ddlYearMS').val() == '請選擇') {
                        alert('請選擇開始年月！');
                        $('#ddlYearMS').focus();
                        return false;
                    }
                    if ($('#ddlMonthS').val() == '請選擇') {
                        alert('請選擇開始年月！');
                        $('#ddlMonthS').focus();
                        return false;
                    }
                    if ($('#ddlYearME').val() == '請選擇') {
                        alert('請選擇結束年月！');
                        $('#ddlYearME').focus();
                        return false;
                    }
                    if ($('#ddlMonthE').val() == '請選擇') {
                        alert('請選擇結束年月！');
                        $('#ddlMonthE').focus();
                        return false;
                    }
                    if (parseInt($('#ddlYearME').val()) < parseInt($('#ddlYearMS').val()) ||
                        ($('#ddlYearME').val() == $('#ddlYearMS').val() &&
                        parseInt($('#ddlMonthE').val()) < parseInt($('#ddlMonthS').val()))) {
                        alert('結束年月不可小於開始年月！');
                        $('#ddlYearS').focus();
                        return false;
                    }

                    // 檢核月繳開始年月是否小於現在日期 by Hilty 2013/12/19
                    var today = new Date();
                    var tMonth = today.getMonth() + 1;
                    var tYear = today.getFullYear();

                    if ($('#ddlYearMS').val() <= tYear && $('#ddlMonthS').val() < tMonth) {
                        alert('請確認月繳開始年月！');
                        return false;
                    }
                }
                //季繳檢核
                if ($('#HFD_PayType').val() == '季繳') {
//                    if ($('#ddlYearQS').val() == '請選擇') {
//                        alert('請選擇開始年季！');
//                        $('#ddlYearQS').focus();
//                        return false;
//                    }
//                    if ($('#ddlQtrS').val() == '請選擇') {
//                        alert('請選擇開始年季！');
//                        $('#ddlQtrS').focus();
//                        return false;
//                    }
//                    if ($('#ddlYearQE').val() == '請選擇') {
//                        alert('請選擇結束年季！');
//                        $('#ddlYearQE').focus();
//                        return false;
//                    }
//                    if ($('#ddlQtrE').val() == '請選擇') {
//                        alert('請選擇結束年季！');
//                        $('#ddlQtrE').focus();
//                        return false;
//                    }
//                    if (parseInt($('#ddlYearQE').val()) < parseInt($('#ddlYearQS').val()) ||
//                        ($('#ddlYearQE').val() == $('#ddlYearQS').val() &&
//                        parseInt($('#ddlQtrE').val()) < parseInt($('#ddlQtrS').val()))) {
//                        alert('結束年季不可小於開始年季！');
//                        $('#ddlYearQS').focus();
//                        return false;
//                    }
                }
                //年繳檢核
                if ($('#HFD_PayType').val() == '年繳') {
                    if ($('#ddlYearS').val() == '請選擇') {
                        alert('請選擇開始年！');
                        $('#ddlYearS').focus();
                        return false;
                    }
                    if ($('#ddlYearE').val() == '請選擇') {
                        alert('請選擇結束年！');
                        $('#ddlYearE').focus();
                        return false;
                    }
                    if (parseInt($('#ddlYearE').val()) < parseInt($('#ddlYearS').val())) {
                        alert('結束年不可小於開始年！');
                        $('#ddlYearS').focus();
                        return false;
                    }
                }
            }
        }
//----------------------------------------------------------------------------
//奉獻方式更改時呼叫
function PayTypeChange()
{
    $('#HFD_PayType').val($('input:radio[name="PayType"]').getValue());

    $('#divMonth').hide();
    $('#divQtr').hide();
	// 季繳提示訊息 by Hilty 2013/12/17
	$('#divQtrText').hide();
    $('#divYear').hide();
    if ($('#HFD_PayType').val() == '月繳') {
        $('#divMonth').show();
    }
    if ($('#HFD_PayType').val() == '季繳') {
        //$('#divQtr').show();
		// 季繳提示訊息 by Hilty 2013/12/17
		$('#divQtrText').show();
    }
    if ($('#HFD_PayType').val() == '年繳') {
        $('#divYear').show();
    }
    var itemList = GetItemList();
    $('#lblGridList').html(itemList);
}
