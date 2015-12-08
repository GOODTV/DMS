function CreditCard_OnClick(){
  document.ecpay.But_Dn.disabled=true;
  act_submit(document.ecpay,'donationb','');
}
function Atm_OnClick(){
  document.ecbank.But_Dn.disabled=true;
  act_submit(document.ecbank,'donationatm','');
}
function Vacc_OnClick(){
  document.ecbank.But_Dn.disabled=true;
  act_submit(document.ecbank,'donationvacc','');
}
function BarCode_OnClick(){
  document.ecbank.But_Dn.disabled=true;
  act_submit(document.ecbank,'donationbarcode','');
}
function FourInOne_OnClick(){
  document.ecbank.But_Dn.disabled=true;
  act_submit(document.ecbank,'donation4in1','');
}
function Next_OnClick(){
  location.href='ecpay.asp';
}
function Exit_OnClick(){
  location.href='ecpay.asp';
}