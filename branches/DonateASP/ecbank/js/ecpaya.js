function Window_OnLoad(){
  document.form_donate.Donate_Amount.focus();
}
function UCaseIDNO(){
  document.form_donate.Donate_IDNO.value=document.form_donate.Donate_IDNO.value.toUpperCase();
}
function UCaseInvoiceIDNO(){
  document.form_donate.Donate_Invoice_IDNO.value=document.form_donate.Donate_Invoice_IDNO.value.toUpperCase();
}  
function SameAddress_OnClick(){
  if(document.form_donate.IsSameAddress.checked){
    document.form_donate.Donate_Invoice_Title.value=document.form_donate.Donate_Name.value;
    document.form_donate.Donate_Invoice_IDNO.value=document.form_donate.Donate_IDNO.value;
    ChgCity(document.form_donate.Donate_CityCode.value,document.form_donate.Donate_Invoice_AreaCode,document.form_donate.Donate_Invoice_ZipCode);
    document.form_donate.Donate_Invoice_ZipCode.value=document.form_donate.Donate_ZipCode.value;
    document.form_donate.Donate_Invoice_CityCode.value=document.form_donate.Donate_CityCode.value;
    document.form_donate.Donate_Invoice_AreaCode.value=document.form_donate.Donate_AreaCode.value;
    document.form_donate.Donate_Invoice_Address.value=document.form_donate.Donate_Address.value;
  }
}
function Next_OnClick(){
  document.form_donate.Donate_Name.value=document.form_donate.Donate_Name.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_IDNO.value=document.form_donate.Donate_IDNO.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_CellPhone.value=document.form_donate.Donate_CellPhone.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_TelOffice.value=document.form_donate.Donate_TelOffice.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_Address.value=document.form_donate.Donate_Address.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_Email.value=document.form_donate.Donate_Email.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_Invoice_Title.value=document.form_donate.Donate_Invoice_Title.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_Invoice_IDNO.value=document.form_donate.Donate_Invoice_IDNO.value.replace(/(^\s*)|(\s*$)/g,"");
  document.form_donate.Donate_Invoice_Address.value=document.form_donate.Donate_Invoice_Address.value.replace(/(^\s*)|(\s*$)/g,"");
  ck=new Boolean(true);
  if(ck) ck=column(document.form_donate.Donate_DeptId,'string','Y','Y','10','機構名稱');
  if(ck) ck=column(document.form_donate.Donate_Amount,'number','Y','Y','7','捐款金額');
  if(ck){
    if(Number(document.form_donate.Donate_Amount.value)<100){
      alert('捐款金額  欄位請勿小於100元！');
      document.form_donate.Donate_Amount.focus();
      ck==false;
      return;
    }
  }
  if(ck){
    if((document.form_donate.Donate_Type.value=='creditcard')&&Number(document.form_donate.Donate_Amount.value)>100000){
      alert('信用卡捐款金額上限為100,000元！');
      document.form_donate.Donate_Amount.focus();
      ck==false;
      return;
    }
  }
  if(ck){
    if((document.form_donate.Donate_Type.value=='webatm'||document.form_donate.Donate_Type.value=='vacc')&&Number(document.form_donate.Donate_Amount.value)>30000){
      if(document.form_donate.Donate_Type.value=='webatm'){
      	alert('Web-ATM捐款金額上限為30,000元！');
      }else{
        alert('銀行虛擬帳號捐款金額上限為30,000元！');
      }
      document.form_donate.Donate_Amount.focus();
      ck==false;
      return;
    }
  }
  if(ck){
    if((document.form_donate.Donate_Type.value=='barcode'||document.form_donate.Donate_Type.value=='ibon'||document.form_donate.Donate_Type.value=='famiport'||document.form_donate.Donate_Type.value=='lifeet'||document.form_donate.Donate_Type.value=='okgo'||document.form_donate.Donate_Type.value=='cvs')&&Number(document.form_donate.Donate_Amount.value)>20000){
      alert('超商捐款金額上限為20,000元！');
      document.form_donate.Donate_Amount.focus();
      ck==false;
      return;
    }
  }
  if(ck) ck=column(document.form_donate.Donate_Name,'string','Y','Y','40','真實姓名');
  if(ck) ck=column(document.form_donate.Donate_IDNO,'idno','N','Y','10','身分證');
  if(ck) ck=column(document.form_donate.Donate_Birthday_Year,'number','N','Y','4','出生日期(年)');
  if(ck) ck=column(document.form_donate.Donate_Birthday_Month,'number','N','Y','2','出生日期(月)');
  if(ck) ck=column(document.form_donate.Donate_Birthday_Day,'number','N','Y','2','出生日期(日)');
  if(document.form_donate.Donate_Birthday_Year.value!=''||document.form_donate.Donate_Birthday_Month.value!=''||document.form_donate.Donate_Birthday_Day.value!=''){
    document.form_donate.Donate_Birthday.value=document.form_donate.Donate_Birthday_Year.value+'/'+document.form_donate.Donate_Birthday_Month.value+'/'+document.form_donate.Donate_Birthday_Day.value;
  }else{
    document.form_donate.Donate_Birthday.value=''
  }
  if(ck){
    if (form_donate.Donate_Birthday.value!='') ck=column(document.form_donate.Donate_Birthday,'date','N','Y','10','出生日期');
  }
  if(ck) ck=column(document.form_donate.Donate_Education,'string','N','Y','20','教育程度');
  if(ck) ck=column(document.form_donate.Donate_Occupation,'string','N','Y','20','職業');
  if(ck) ck=column(document.form_donate.Donate_CellPhone,'cell','Y','Y','10','手機');
  if(ck) ck=column(document.form_donate.Donate_TelOffice,'string','N','Y','20','聯絡電話');
  if(ck) ck=column(document.form_donate.Donate_CityCode,'string','Y','Y','1','通訊地址縣市');
  if(ck) ck=column(document.form_donate.Donate_AreaCode,'member','Y','Y','5','通訊地址鄉鎮市區');
  if(ck) ck=column(document.form_donate.Donate_Address,'string','Y','Y','100','通訊地址路段號樓');
  if(ck) ck=column(document.form_donate.Donate_Email,'email','Y','N','60','電子郵件');
  if(ck) ck=column(document.form_donate.Donate_Purpose,'string','Y','Y','20','捐款用途'); 
  if(document.form_donate.Donate_Invoice_Type[0].checked||document.form_donate.Donate_Invoice_Type[1].checked){
    if(ck) ck=column(document.form_donate.Donate_Invoice_Title,'string','Y','Y','40','收據抬頭');
    if(ck) ck=column(document.form_donate.Donate_Invoice_CityCode,'string','Y','Y','1','收據地址縣市');
    if(ck) ck=column(document.form_donate.Donate_Invoice_AreaCode,'member','Y','Y','5','收據地址鄉鎮市區');
    if(ck) ck=column(document.form_donate.Donate_Invoice_Address,'string','Y','Y','100','收據地址路段號樓');
  }else{
    if(ck) ck=column(document.form_donate.Donate_Invoice_Title,'string','N','Y','40','收據抬頭');
    if(ck) ck=column(document.form_donate.Donate_Invoice_CityCode,'string','N','Y','1','收據地址縣市');
    if(ck) ck=column(document.form_donate.Donate_Invoice_AreaCode,'member','N','Y','5','收據地址鄉鎮市區');
    if(ck) ck=column(document.form_donate.Donate_Invoice_Address,'string','N','Y','100','收據地址路段號樓');
  }
  if(ck) ck=column(document.form_donate.Donate_Invoice_IDNO,'idno','N','Y','10','收據身分證');
  if(ck) document.form_donate.But_Dn.disabled=true;
  if(ck) act_submit(document.form_donate,'donationa','');
}