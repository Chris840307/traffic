function lockString(strObj)
{
	len=strObj.value.length ;
	for(j=0;j<len;j++)
	{
	   acode=strObj.value.charCodeAt(j);	
     if (acode > "47" && acode < "58"){}
     else if (acode > "64" && acode < "91"){}
     //elif (acode > "96" && acode < "123"){}
     else {
     	 alert("此欄位只可輸入英文(大寫)或數字!!")
     	 strObj.value=strObj.value.substring(0,j)+strObj.value.substring(j+1,len)
     	 return false; 
     }	
	}
}

function openAddGetBill(OpenFileStr,frmName)
{
  window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=1000,height=550,resizable=yes,left=0,top=0,status=yes");	
}
function openAddUnitInfo(OpenFileStr,frmName)
{
  window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=550,resizable=yes,left=0,top=0,status=no");	
}
function openAddWindow(OpenFileStr,frmName)
{
  window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=no,width=800,height=550,resizable=yes,left=0,top=0,status=no");	
}

function exportExcel(OpenFileStr,frmName)
{
window.open(OpenFileStr,frmName,"scrollbars=yes,menubar=yes,width=800,height=600,resizable=yes,left=0,top=0,status=no");	
}

function chkBillNumber(formObj,errStr){
   var BillNumber = formObj.value      
   
   if (formObj.value != ""){
      if (BillNumber.length != 9) {
		  if (errStr != "none"){
      		alert(errStr);
		  }
      	  return "N"; 
      }else {
		 acode=BillNumber.charCodeAt(0);
         if ((acode > "64" && acode < "91") || (acode > "96" && acode < "123"))//檢查第一碼是否為英文字母
         { 
         	  return "Y";
         }else{
			  if (errStr != "none"){
				alert(errStr);
			  }
         	  formObj.focus();
         	  formObj.select();
         	  return "N"; 
         } 
      }	
   }
}

function calStr(strObj,maxLen)
{
  var strlen = strObj.value.length;
  var charCount = 0;
  
  if(strlen > 0){
    for(var i=0;i < strlen;i++) {
      c = '';
      c = escape(strObj.value.charAt(i));
      acode=strObj.value.charCodeAt(i);
      if(acode == 10 ) {
      	 alert("請勿輸入斷行標記!!");
      	 strObj.value=strObj.value.substring(0,i-1)+strObj.value.substring(i+1,strlen);
      	 return false;
      }      
      if(acode==34 | acode==39 | acode==38 ) {
      	 alert("請勿輸入' & \" 等字元!!");
      	 strObj.value=strObj.value.substring(0,i)+strObj.value.substring(i+1,strlen);
      	 return false;
      }
      if(c.charAt(0) == '%') {
        cc = c.charAt(1); //IE~u,NS~A
        if(cc == 'A' || cc=='B' || cc =='u' || cc == 'a') {
        	charCount = charCount + 2;
        }else{
        	charCount = charCount + 1;     
        }
      }else{
         charCount = charCount + 1;   
      }
      if ((maxLen - charCount) < 0){
         alert("此欄位可輸入之最大長度為:" + maxLen);	
      	 strObj.value=strObj.value.substring(0,i)+strObj.value.substring(i+1,strlen);
      	 return false;         
         break;
      }
    }
  }    
  document.all.nbchars.value = maxLen - charCount;
}

function lockSpecialCharr(strObj)
{
  var strlen = strObj.value.length;
  
  if(strlen > 0){
    for(var i=0;i < strlen;i++) {
      c = '';
      c = escape(strObj.value.charAt(i));
      acode=strObj.value.charCodeAt(i);
      if(acode == 10 ) {
      	 alert("請勿輸入斷行標記!!");
      	 strObj.value=strObj.value.substring(0,i-1)+strObj.value.substring(i+1,strlen);
      	 return false;
      }      
      if(acode==34 | acode==39 | acode==38 ) {
      	 alert("請勿輸入' & \" 等字元!!");
      	 strObj.value=strObj.value.substring(0,i)+strObj.value.substring(i+1,strlen);
      	 return false;
      }
    }
  }    
}

function lockNum(strObj)
{
  var strlen = strObj.value.length;
  
  if(strlen > 0){
    for(var i=0;i < strlen;i++) {  
      acode=strObj.value.charCodeAt(i);
      if(acode < 48 || acode > 57 ) {
      	 alert("此欄位只可輸入數字!!");      	 
      	 strObj.value=strObj.value.substring(0,i)+strObj.value.substring(i+1,strlen);
      	 strObj.focus();
      	 return false;
      }   
    }
 }

}
//用車牌號碼判斷車種
function chkCarNoFormat(CarNo){
	strHeavy="ABCFGHIJKLMNOPYX";	//重機第一碼
	strSmall="DEQRSTUVWZ";	//輕機第一碼
	if ( (CarNo.indexOf("-",0)) != -1)	{
		CarNoArray=CarNo.split("-");
		if ((CarNoArray[0].length<=3 && CarNoArray[1]=="HHH") || (CarNoArray[1].length<=3 && CarNoArray[0]=="HHH")){
			return 3;
		}else if (CarNoArray[0].length==2 && CarNoArray[1].length==2){
			return 2;
		}else if ((CarNoArray[0].length==2 && CarNoArray[1].length==4) || (CarNoArray[0].length==4 && CarNoArray[1].length==2) || (CarNoArray[0].length==2 && CarNoArray[1].length==3) || (CarNoArray[0].length==3 && CarNoArray[1].length==2)){
			return 1;
		}else if (CarNoArray[0].length==3 && CarNoArray[1].length==3){
			//if ((CarNoArray[0].indexOf("0",0)) != -1){
			//	return 0;
			//}else{
				if ((strHeavy.indexOf(CarNo.charAt(0),0)) != -1){
					if ((CarNoArray[0].indexOf("0",0)) != -1){
						return 0;
					}else{
						//修改開頭二碼為重車或輕機
						if(CarNoArray[1].substr(0,1) =="Q" || CarNo.substr(0,2)=="YN" || CarNo.substr(0,2)=="YO" || CarNo.substr(0,2)=="YP" || CarNo.substr(0,2)=="YQ" || CarNo.substr(0,2)=="YR" || CarNo.substr(0,2)=="YS" || CarNo.substr(0,2)=="YT" || CarNo.substr(0,2)=="YW" || CarNo.substr(0,2)=="YU" || CarNo.substr(0,2)=="YY" || CarNo.substr(0,2)=="YV" || CarNo.substr(0,2)=="YX" || CarNo.substr(0,2)=="YZ")
						{
						return 4;
						}
						else
						{
						return 3;
						}
					}
				}else if((strSmall.indexOf(CarNo.charAt(0),0)) != -1){
					if ((CarNoArray[0].indexOf("0",0)) != -1){
						return 0;
					}else{
						return 4;
					}
				}else{
						if(CarNoArray[1].substr(0,1) =="Q" ||CarNo.substr(0,2)=="YN" || CarNo.substr(0,2)=="YO" || CarNo.substr(0,2)=="YP" || CarNo.substr(0,2)=="YQ" || CarNo.substr(0,2)=="YR" || CarNo.substr(0,2)=="YS" || CarNo.substr(0,2)=="YT" || CarNo.substr(0,2)=="YW" || CarNo.substr(0,2)=="YU" || CarNo.substr(0,2)=="YY" || CarNo.substr(0,2)=="YV"  || CarNo.substr(0,2)=="YX" || CarNo.substr(0,2)=="YZ")
						{
						return 4;
						}
						else
						{
						return 3;
						}
				}
			//}
		}else{
			return 0;
		}
	}else{
		return 0;
	}

}
//身分証檢查
function check_tw_id(sId){
  var LegalID = "0123456789"
  var fResult=true;
  if(sId.length<10)
    fResult=false;
  else{
    if((sId.charAt(0)=='A') || (sId.charAt(0)=='a')) value=10
    else if((sId.charAt(0)=='B') || (sId.charAt(0)=='b')) value=11
    else if((sId.charAt(0)=='C') || (sId.charAt(0)=='c')) value=12
    else if((sId.charAt(0)=='D') || (sId.charAt(0)=='d')) value=13
    else if((sId.charAt(0)=='E') || (sId.charAt(0)=='e')) value=14
    else if((sId.charAt(0)=='F') || (sId.charAt(0)=='f')) value=15
    else if((sId.charAt(0)=='G') || (sId.charAt(0)=='g')) value=16
    else if((sId.charAt(0)=='H') || (sId.charAt(0)=='h')) value=17
    else if((sId.charAt(0)=='J') || (sId.charAt(0)=='j')) value=18
    else if((sId.charAt(0)=='K') || (sId.charAt(0)=='k')) value=19
    else if((sId.charAt(0)=='L') || (sId.charAt(0)=='l')) value=20
    else if((sId.charAt(0)=='M') || (sId.charAt(0)=='m')) value=21
    else if((sId.charAt(0)=='N') || (sId.charAt(0)=='n')) value=22
    else if((sId.charAt(0)=='P') || (sId.charAt(0)=='p')) value=23
    else if((sId.charAt(0)=='Q') || (sId.charAt(0)=='q')) value=24
    else if((sId.charAt(0)=='R') || (sId.charAt(0)=='r')) value=25
    else if((sId.charAt(0)=='S') || (sId.charAt(0)=='s')) value=26
    else if((sId.charAt(0)=='T') || (sId.charAt(0)=='t')) value=27
    else if((sId.charAt(0)=='U') || (sId.charAt(0)=='u')) value=28
    else if((sId.charAt(0)=='V') || (sId.charAt(0)=='v')) value=29
    else if((sId.charAt(0)=='X') || (sId.charAt(0)=='x')) value=30
    else if((sId.charAt(0)=='Y') || (sId.charAt(0)=='y')) value=31
    else if((sId.charAt(0)=='W') || (sId.charAt(0)=='w')) value=32
    else if((sId.charAt(0)=='Z') || (sId.charAt(0)=='z')) value=33
    else if((sId.charAt(0)=='I') || (sId.charAt(0)=='i')) value=34
    else if((sId.charAt(0)=='O') || (sId.charAt(0)=='o')) value=35
    else fResult = false ;
  }
  if(fResult==true){
    value = Math.floor(value/10) + (value%10)*9 + parseInt(sId.charAt(1))*8 +
            parseInt(sId.charAt(2))*7 + parseInt(sId.charAt(3)) * 6 + parseInt(sId.charAt(4)) * 5 +
            parseInt(sId.charAt(5))*4 + parseInt(sId.charAt(6)) * 3+ parseInt(sId.charAt(7)) * 2+
            parseInt(sId.charAt(8)) + parseInt(sId.charAt(9)) ;
    value = value % 10 ;
    if(value!=0) fResult = false ;
    var i;
    var c;
    for (i = 1; i < sId.length; i++){
      c = sId.charAt(i);
      if (LegalID.indexOf(c) == -1) fResult = false;
    }
  }
  if(fResult == false)
    return false;
  else
    return true;
}

//用車速，地點得到違規法條
function getIllegalRule(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;
//		if (Illaddr.indexOf("高速公路",0)!=-1){
//			if (Speed <= 20 && Speed > 0){
//				return "3310101";
//			}else if (Speed > 20 && Speed <= 40){
//				return "3310103";
//			}else if (Speed > 40 && Speed <= 60){
//				return "3310105";
//			}else if (Speed > 60 && Speed <= 80){
//				return "4310210";
//			}else if (Speed > 80 && Speed <= 100){
//				return "4310211";
//			}else if (Speed > 100){
//				return "4310212";
//			}else{
//				return "Null";
//			}
//		}else 
		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310134";
			}else if (Speed > 20 && Speed <= 40){
				return "3310136";
			}else if (Speed > 40 && Speed <= 60){
				return "4310240";
			}else if (Speed > 60 && Speed <= 80){
				return "4310241";
			}else if (Speed > 80){
				return "4310242";
			}else{
				return "Null";
			}
		}else{
			if (Speed <= 20 && Speed > 0){
				return "4000005";
			}else if (Speed > 20 && Speed <= 40){
				return "4000006";
			}else if (Speed > 40 && Speed <= 60){
				return "4310240";
			}else if (Speed > 60 && Speed <= 80){
				return "4310241";
			}else if (Speed > 80){
				return "4310242";
			}else{
				return "Null";
			}	
		}
	}
}

//用車速，地點得到違規法條--違規日1120630帶舊法條
function getIllegalRule_Old1120630(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;

		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310102";
			}else if (Speed > 20 && Speed <= 40){
				return "3310104";
			}else if (Speed > 40 && Speed <= 60){
				return "3310106";
			}else if (Speed > 60 && Speed <= 80){
				return "4310210";
			}else if (Speed > 80 && Speed <= 100){
				return "4310211";
			}else if (Speed > 100){
				return "4310212";
			}else{
				return "Null";
			}
		}else{
			if (Speed <= 20 && Speed > 0){
				return "4000005";
			}else if (Speed > 20 && Speed <= 40){
				return "4000006";
			}else if (Speed > 40 && Speed <= 60){
				return "4000007";
			}else if (Speed > 60 && Speed <= 80){
				return "4310210";
			}else if (Speed > 80 && Speed <= 100){
				return "4310211";
			}else if (Speed > 100){
				return "4310212";
			}else{
				return "Null";
			}	
		}
	}
}

//用車速，地點得到違規法條(台中市)80KM以上帶快速道路法條
function getIllegalRule2(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;
//		if (Illaddr.indexOf("高速公路",0)!=-1){
//			if (Speed <= 20 && Speed > 0){
//				return "3310101";
//			}else if (Speed > 20 && Speed <= 40){
//				return "3310103";
//			}else if (Speed > 40 && Speed <= 60){
//				return "3310105";
//			}else if (Speed > 60 && Speed <= 80){
//				return "4310210";
//			}else if (Speed > 80 && Speed <= 100){
//				return "4310211";
//			}else if (Speed > 100){
//				return "4310212";
//			}else{
//				return "Null";
//			}
//		}else 
		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310134";
			}else if (Speed > 20 && Speed <= 40){
				return "3310136";
			}else if (Speed > 40 && Speed <= 60){
				return "4310240";
			}else if (Speed > 60 && Speed <= 80){
				return "4310241";
			}else if (Speed > 80){
				return "4310242";
			}else{
				return "Null";
			}
		}else{
			if (RuleSpeed>=80){
				if (Speed <= 20 && Speed > 0){
					return "3310134";
				}else if (Speed > 20 && Speed <= 40){
					return "3310136";
				}else if (Speed > 40 && Speed <= 60){
					return "4310240";
				}else if (Speed > 60 && Speed <= 80){
					return "4310241";
				}else if (Speed > 80){
					return "4310242";
				}else{
					return "Null";
				}
			}else{
				if (Speed <= 20 && Speed > 0){
					return "4000005";
				}else if (Speed > 20 && Speed <= 40){
					return "4000006";
				}else if (Speed > 40 && Speed <= 60){
					return "4310240";
				}else if (Speed > 60 && Speed <= 80){
					return "4310241";
				}else if (Speed > 80 ){
					return "4310242";

				}else{
					return "Null";
				}	
			}
		}
	}
}

//用車速，地點得到違規法條(台中市)80KM以上帶快速道路法條--違規日1120630帶舊法條
function getIllegalRule2_Old1120630(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;
//		if (Illaddr.indexOf("高速公路",0)!=-1){
//			if (Speed <= 20 && Speed > 0){
//				return "3310101";
//			}else if (Speed > 20 && Speed <= 40){
//				return "3310103";
//			}else if (Speed > 40 && Speed <= 60){
//				return "3310105";
//			}else if (Speed > 60 && Speed <= 80){
//				return "4310210";
//			}else if (Speed > 80 && Speed <= 100){
//				return "4310211";
//			}else if (Speed > 100){
//				return "4310212";
//			}else{
//				return "Null";
//			}
//		}else 
		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310102";
			}else if (Speed > 20 && Speed <= 40){
				return "3310104";
			}else if (Speed > 40 && Speed <= 60){
				return "3310106";
			}else if (Speed > 60 && Speed <= 80){
				return "4310210";
			}else if (Speed > 80 && Speed <= 100){
				return "4310211";
			}else if (Speed > 100){
				return "4310212";
			}else{
				return "Null";
			}
		}else{
			if (RuleSpeed>=80){
				if (Speed <= 20 && Speed > 0){
					return "3310102";
				}else if (Speed > 20 && Speed <= 40){
					return "3310104";
				}else if (Speed > 40 && Speed <= 60){
					return "3310106";
				}else if (Speed > 60 && Speed <= 80){
					return "4310210";
				}else if (Speed > 80 && Speed <= 100){
					return "4310211";
				}else if (Speed > 100){
					return "4310212";
				}else{
					return "Null";
				}
			}else{
				if (Speed <= 20 && Speed > 0){
					return "4000005";
				}else if (Speed > 20 && Speed <= 40){
					return "4000006";
				}else if (Speed > 40 && Speed <= 60){
					return "4000007";
				}else if (Speed > 60 && Speed <= 80){
					return "4310210";
				}else if (Speed > 80 && Speed <= 100){
					return "4310211";
				}else if (Speed > 100){
					return "4310212";
				}else{
					return "Null";
				}	
			}
		}
	}
}

//用車速，地點得到違規法條(台東縣 4310201 , 4000003)61以上才能開43條
function getIllegalRule3(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad){
	if (ProsecutionTypeID=="R"){
		return "5310001";
	}else{
		Speed=IllSpeed-RuleSpeed;
//		if (Illaddr.indexOf("高速公路",0)!=-1){
//			if (Speed <= 20 && Speed > 0){
//				return "3310101";
//			}else if (Speed > 20 && Speed <= 40){
//				return "3310103";
//			}else if (Speed > 40 && Speed <= 60){
//				return "3310105";
//			}else if (Speed > 60 && Speed <= 80){
//				return "4310210";
//			}else if (Speed > 80 && Speed <= 100){
//				return "4310211";
//			}else if (Speed > 100){
//				return "4310212";
//			}else{
//				return "Null";
//			}
//		}else 
		if ((Illaddr.indexOf("快速道路",0)!=-1) || (Illaddr.indexOf("快速公路",0)!=-1) || (chkHighRoad==true)){
			if (Speed <= 20 && Speed > 0){
				return "3310134";
			}else if (Speed > 20 && Speed <= 40){
				return "3310136";
			}else if (Speed > 40 && Speed <= 60){
				return "4310240";
			}else if (Speed > 60 && Speed <= 80){
				return "4310241";
			}else if (Speed > 80){
				return "4310242";
			}else{
				return "Null";
			}
		}else{
			if (Speed <= 20 && Speed > 0){
				return "4000005";
			}else if (Speed > 20 && Speed <= 40){
				return "4000006";
			}else if (Speed > 40 && Speed <= 60){
				return "4310240";
			}else if (Speed > 60 && Speed <= 80){
				return "4310241";
			}else if (Speed > 80 ){
				return "4310242";
			}else{
				return "Null";
			}	
		}
	}
}

//檢查超速法條是否正確
function chkSpeedRuleIsRight(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad,UserRule,ChkFlag){
	var SpeedRule;
	if (ChkFlag=="1"){
		SpeedRule=getIllegalRule(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad);
	}else if (ChkFlag=="2"){
		SpeedRule=getIllegalRule2(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad);
	}else if (ChkFlag=="3"){
		SpeedRule=getIllegalRule3(Illaddr,RuleSpeed,IllSpeed,ProsecutionTypeID,chkHighRoad);
	}
	if (SpeedRule==UserRule){
		return true; 
	}else if (SpeedRule=="4000005" && (UserRule=="4000001" || UserRule=="4000005" || UserRule=="4000011" || UserRule=="4000014")){
		return true; 
	}else if (SpeedRule=="4000006" && (UserRule=="4000002" || UserRule=="4000006" || UserRule=="4000012" || UserRule=="4000015")){
		return true; 
	}else if (SpeedRule=="4000007" && (UserRule=="4000003" || UserRule=="4000007" || UserRule=="4000013" || UserRule=="4000016")){
		return true; 
	}else if (SpeedRule=="4310210" && (UserRule=="4310210" || UserRule=="4310201" || UserRule=="4310213" || UserRule=="4310216" || UserRule=="4310219" || UserRule=="4310221" || UserRule=="4310223" || UserRule=="4310225")){
		return true; 
	}else if (SpeedRule=="4310211" && (UserRule=="4310211" || UserRule=="4310202" || UserRule=="4310214" || UserRule=="4310217" || UserRule=="4310220" || UserRule=="4310222" || UserRule=="4310224" || UserRule=="4310226")){
		return true; 
	}else if (SpeedRule=="4310212" && (UserRule=="4310212" || UserRule=="4310203" || UserRule=="4310215" || UserRule=="4310218" || UserRule=="4310227")){
		return true; 
	}else if (SpeedRule=="3310102" && (UserRule=="3310102" || UserRule=="3310112" || UserRule=="3310118" || UserRule=="3310128")){
		return true; 
	}else if (SpeedRule=="3310104" && (UserRule=="3310104" || UserRule=="3310114" || UserRule=="3310120" || UserRule=="3310130")){
		return true; 
	}else if (SpeedRule=="3310106" && (UserRule=="3310106" || UserRule=="3310116" || UserRule=="3310122" || UserRule=="3310132")){
		return true; 
	//new
	}else if (SpeedRule=="4310240" && (UserRule=="4310240" || UserRule=="4000007" || UserRule=="4000013" || UserRule=="4000016" || UserRule=="3310106" || UserRule=="3310116" || UserRule=="3310122" || UserRule=="3310132" || UserRule=="4310243" || UserRule=="4310246" || UserRule=="4310249" || UserRule=="4310251" || UserRule=="4310253" || UserRule=="4310255" || UserRule=="4310258" || UserRule=="4310261" || UserRule=="4310264" || UserRule=="4310266" || UserRule=="4310268")){
		return true; 
	}else if (SpeedRule=="4310241" && (UserRule=="4310241" || UserRule=="4310210" || UserRule=="4310225" || UserRule=="4310244" || UserRule=="4310247" || UserRule=="4310250" || UserRule=="4310252" || UserRule=="4310254" || UserRule=="4310256" || UserRule=="4310259" || UserRule=="4310262" || UserRule=="4310265" || UserRule=="4310267" || UserRule=="4310269")){
		return true; 
	}else if (SpeedRule=="4310242" && (UserRule=="4310242" || UserRule=="4310211" || UserRule=="4310212" || UserRule=="4310226" || UserRule=="4310245" || UserRule=="4310248" || UserRule=="4310257" || UserRule=="4310260" || UserRule=="4310263")){
		return true; 
	}else if (SpeedRule=="3310134" && (UserRule=="3310134" || UserRule=="3310102" || UserRule=="3310118" || UserRule=="3310142" || UserRule=="3310146" || UserRule=="3310154")){
		return true; 
	}else if (SpeedRule=="3310136" && (UserRule=="3310136" || UserRule=="3310104" || UserRule=="3310120" || UserRule=="3310144" || UserRule=="3310148" || UserRule=="3310156")){
		return true; 
	}else{
		return false;
	}
}


//檢查居留證號
function CheckResidenceID(ResidenceID){
	var s ='';
	var i =0;

	// 格式錯誤
	s = ResidenceID.substr(0, 1);
	if (ResidenceID.length != 10 || s < 'A' || s > 'Z') {
		return false;
	}

	// 2-10內有非1-9的資料
	for (i=1; i<10;i++ )
	{
		if ((ResidenceID.substr(i, 1))>'9' || (ResidenceID.substr(i, 1))<'0')
		{
			return false;
		}
	}
       
    //統號1~10碼各自乘以(1987654321)然後個位數相加
    //証號第十碼為檢查碼=(10-統號個位數相加後之個位數)

    //計算結果(已計算個位數相加)：居留證首字(統號前2碼)
    var N1;
	if (s == 'A'){
		N1 = 1;
	}else if (s == 'B')
	{
		N1 = 10;
	}else if (s == 'C')
	{
		N1 = 9;
	}else if (s == 'D')
	{	
		N1 = 8;
	}else if (s == 'E')
	{
		N1 = 7;
	}else if (s == 'F')
	{
		N1 = 6;
	}else if (s == 'G')
	{
		N1 = 5;
	}else if (s == 'H')
	{
		N1 = 4;
	}else if (s == 'J')
	{
		N1 = 3;
	}else if (s == 'K')
	{
		N1 = 2;	
	}else if (s == 'L')
	{
		N1 = 2;
	}else if (s == 'M')
	{
		N1 = 11;
	}else if (s == 'N')
	{
		N1 = 10;
	}else if (s == 'P')
	{
		N1 = 9;
	}else if ( s == 'Q')
	{
		N1 = 8;
	}else if (s == 'R')
	{
		N1 = 7;
	}else if (s == 'S')
	{
		N1 = 6;
	}else if (s == 'T')
	{
		N1 = 5;
	}else if (s == 'U')
	{
		N1 = 4;
	}else if (s == 'V')
	{
		N1 = 3;
	}else if (s == 'X')
	{
		N1 = 3;
	}else if (s == 'Y')
	{
		N1 = 12;
	}else if (s == 'W')
	{
		N1 = 11;
	}else if (s == 'Z')
	{
		N1 = 10;
	}else if (s == 'I')
	{
		N1 = 9;
	}else if (s == 'O')
	{
		N1 = 8;
	}

	//証號2~9碼=統號3~10碼，依序乘以87654321，然後個位數相加
    var N  = 0;
	for (i=1; i<10;i++ )
	{
		N = N + (( ResidenceID.substr(i, 1) * (9 - i) ) % 10);
	}

	//檢查証號第10碼是否=(10-相加後個位數)
	//若相加後個位數=0，檢查碼=0
	var ChkNum ;
	if (((N + N1) % 10) == 0)
	{
		ChkNum = 0;
	}else{
		ChkNum = (10 - ((N + N1) % 10));
	}

	if ( ResidenceID.substr(9, 1) == ChkNum)
	{
		return true;
	}else{
		return false;
	}
}