(function ($) {
    $.fn.CityAreaZip = function (ddlCity, ddlArea, txtZip, HFD_Area) {
        $('#' + ddlCity).bind('change', function (e) {
            var CityID = $(e.target).val();
            //alert(CityID);
            ChangeArea(e, ddlArea);
        });
        //----------------------------------------------------------------------
        $('#' + ddlArea).bind("change", function (e) {
            var ZipCode = $(this).val();
            if (ZipCode != '0') {
                $('#' + txtZip).val(ZipCode.substring(0, 3));
                $('#' + HFD_Area).val(ZipCode);
            }
        });
        //----------------------------------------------------------------------
        //�H CityID ���o��m��C��
        function ChangeArea(e, ddlArea) {
            var CityID = $(e.target).val();
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=1' + '&CityID=' + CityID,
                success: function (result) {
                    $('select[id$=' + ddlArea + ']').html(result);
                    //
                    $('#' + HFD_Area).val('');
                    $('#' + txtZip).val('');
                },
                error: function () { alert('ajax failed'); }
            })
        } //end of ChangeArea()
        //----------------------------------------------------------------------
        //alert("City=" + City + " Area=" + Area + " Zip=" + Zip);
    } //end of $.fn.CityAreaZip

    //�P���y�a�}
    $.fn.SetSameAddress = function (ddlDestCity, ddlDestArea, txtDestZip, destHFD_Area, txtDestAddress,
                                    ddlSrcCity, ddlSrcArea, txtSrcZip, txtSrcAddress) {
        var CityID = $('#' + ddlSrcCity).val();
        var AreaID = $('#' + ddlSrcArea).val();
        var ZipCode = $('#' + txtSrcZip).val();
        var Address = $('#' + txtSrcAddress).val();
        $('#' + ddlDestCity).val(CityID);
        //�n�� load �m����
        ChangeArea2(CityID, AreaID, ddlDestArea)
        $('#' + txtDestZip).val(ZipCode);
        $('#' + txtDestAddress).val(Address);
        $('#' + destHFD_Area).val(AreaID);

        function ChangeArea2(CityID, AreaID, destAreaID) {
            $.ajax({
                type: 'post',
                url: "../common/ajax.aspx",
                data: 'Type=1' + '&CityID=' + CityID,
                success: function (result) {
                    $('select[id$=' + destAreaID + ']').html(result);
                    $('#' + destAreaID).val(AreaID);
                },
                error: function () { alert('ajax failed'); }
            })
        }
    } //end of $.fn.SetSameAddress

})(jQuery);

