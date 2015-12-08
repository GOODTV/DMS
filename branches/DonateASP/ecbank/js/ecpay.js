function Next_OnClick(){
  ck=new Boolean(true);
  if(ck) document.form_donate.But_Dn.disabled=true;
  if(ck) act_submit(document.form_donate,'donation','');
}