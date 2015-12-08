<script language="JavaScript">
  function ChkID(idno){
    if(idno.length!=10){
      return false;
    }
    if(isNaN(Number(idno.substring(0,1)))==false){
      return false;
    }
    if(isNaN(Number(idno.substring(1,10)))==true){
      return false;
    }
    
    var alpha = idno.substring(0,1).toUpperCase();
    var d1 = idno.substring(1,2);
    var d2 = idno.substring(2,3);
    var d3 = idno.substring(3,4);
    var d4 = idno.substring(4,5);
    var d5 = idno.substring(5,6);
    var d6 = idno.substring(6,7);
    var d7 = idno.substring(7,8);
    var d8 = idno.substring(8,9);
    var d9 = idno.substring(9,10);
      
    var acc=0;
    switch(alpha){
      case "A":
        acc=1;
        break;
      case "B":
        acc=10;
        break;
      case "C":
        acc=19;
        break;
      case "D":
        acc=28;
        break;
      case "E":
        acc=37;
        break;
      case "F":
        acc=46;
        break;
      case "G":
        acc=55;
        break;
      case "H":
        acc=64;
        break;
      case "I":
        acc=39;
        break;
      case "J":
        acc=73;
        break;
      case "K":
        acc=82;
        break;
      case "L":
        acc=2;
        break;
      case "M":
        acc=11;
        break;
      case "N":
        acc=20;
        break;
      case "O":
        acc=48;
        break;
      case "P":
        acc=29;
        break;
      case "Q":
        acc=38;
        break;
      case "R":
        acc=47;
        break;
      case "S":
        acc=56;
        break;
      case "T":
        acc=65;
        break; 
      case "U":
        acc=74;
        break;
      case "V":
        acc=83;
        break;
      case "W":
        acc=21;
        break;
      case "X":
        acc=3;
        break;
      case "Y":
        acc=12;
        break;
      case "Z":
        acc=30;
        break;
    }
    
    if(acc=='0'){
      return false;
    }
    var checksum = acc+8*d1+7*d2+6*d3+5*d4+4*d5+3*d6+2*d7+1*d8+1*d9;
    var check1 = parseInt(checksum/10);
    var check2 = checksum/10;
    var check3 = (check2-check1)*10;
    if((checksum==check1*10)||(d9=(10-check3))){
      return true;	
    }       
    return false;
  }
</script>