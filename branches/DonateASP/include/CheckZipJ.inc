<script language="JavaScript">
  function chgCity(val,Areaobj,Zipobj){
    Zipobj.value='';
    clearOption(Areaobj);
    Areaobj.options[0] = new Option('鄉鎮市區','');
    var aryArea = new Array(<%=strAll_Area%>);
    for(var i=0;i<=aryArea.length-1;i=i+2){
      if(aryArea[i]==val){
        var theArea = aryArea[i+1];
        var k=1;
        for(var j=0;j<=theArea.length-1;j=j+2){
          Areaobj.options[k] = new Option(theArea[j+1],theArea[j]);
          k++;
        }
      }
    }
  }
  function clearOption(theobj){
    var i;
    for(i=theobj.length-1;i>=0;i--){
      theobj.options[i] = null;
    }
  }
  function chgArea(val,Zipobj){
    Zipobj.value=val;
  }
  
  function ChgAddress(){
    if(form.Address_Type[0].checked == true){
      form.Invoice_Title.value=form.Donor_Name.value;
      form.Invoice_IDNo.value=form.IDNo.value;
      form.Invoice_City.value=form.City.value;
      chgCity(form.City.value,form.Invoice_Area,form.Invoice_ZipCode);
      form.Invoice_ZipCode.value=form.ZipCode.value;
      form.Invoice_Area.value=form.Area.value;
      form.Invoice_Address.value=form.Address.value;  
    }else{
      form.Invoice_Title.value='';
      form.Invoice_IDNo.value='';
      form.Invoice_City.value='';
      clearOption(form.Invoice_Area);
      form.Invoice_Area.options[0] = new Option('鄉鎮市區','');
      form.Invoice_Address.value='';
      form.Invoice_ZipCode.value='';    
    }
  }
  
  function ShowSubMenu(){
    if(form.Donate_Purpose.value=='一般捐款'){
      document.getElementById('SubMenu1').style.display='';
    }else{
      document.getElementById('SubMenu1').style.display='none';
    }
  }     
</script>