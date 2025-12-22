var focusY;var focusX;var ArrText;
function MoveTextVar(textObj){
	var space="||";
	var temp_Arr=textObj.split(space);
	ArrText=new Array(temp_Arr.length);
	for (i=0;i<temp_Arr.length;i++){
		ArrText[i]=temp_Arr[i].split(',');
	}
}

function cleanSpelchar(th){
	if(/[\&\<\>\=\"\']/g.test(th.value)){
		th.value=th.value.replace(/[\&\<\>\=\"\']/g,'');
		alert("不可輸入 & < > = ' \u0022 等符號！");
	}
}

function focuslocation(objname){
	focusY="";focusX="";
	for(i=0;i<ArrText.length;i++){
		for(j=0;j<ArrText[i].length;j++){
			if(objname==ArrText[i][j]){
				focusY=i;focusX=j;
				break;
			}
		}
	}
}
function CodeEnter(objname){
	focuslocation(objname);
	var tmpObj;
	if(ArrText[focusY].length-1==focusX){
		if(ArrText.length-1>focusY){
			tmpObj=eval("myForm."+ArrText[focusY+1][0]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}
	}else if(ArrText[focusY].length-1>focusX){
		tmpObj=eval("myForm."+ArrText[focusY][focusX+1]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}
}
function CodeMoveLeft(objname){
	focuslocation(objname);
	var tmpObj;
	if(focusX==0&&focusY==0){
		tmpObj=eval("myForm."+ArrText[ArrText.length-1][ArrText[ArrText.length-1].length-1]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}else if(focusX==0){
		tmpObj=eval("myForm."+ArrText[focusY-1][ArrText[focusY-1].length-1]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}else{
		tmpObj=eval("myForm."+ArrText[focusY][focusX-1]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}
}
function CodeMoveUp(objname){
	focuslocation(objname);
	var tmpObj;
	if(focusY==0){
		if(focusX>ArrText[ArrText.length-1].length-1){
			tmpObj=eval("myForm."+ArrText[ArrText.length-1][ArrText[ArrText.length-1].length-1]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}else{
			tmpObj=eval("myForm."+ArrText[ArrText.length-1][focusX]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}
	}else{
		if(focusX>ArrText[focusY-1].length-1){
			tmpObj=eval("myForm."+ArrText[focusY-1][ArrText[focusY-1].length-1]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}else{
			tmpObj=eval("myForm."+ArrText[focusY-1][focusX]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}
	}
}
function CodeMoveRight(objname){
	focuslocation(objname);
	var tmpObj;
	if(focusY==ArrText.length-1&&focusX==ArrText[ArrText.length-1].length-1){
		tmpObj=eval("myForm."+ArrText[0][0]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}else if(focusX==ArrText[focusY].length-1){
		tmpObj=eval("myForm."+ArrText[focusY+1][0]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}else{
		tmpObj=eval("myForm."+ArrText[focusY][focusX+1]);
		tmpObj.focus();
		if(tmpObj.tagName=='INPUT'){tmpObj.select();}
	}
}
function CodeMoveDown(objname){
	focuslocation(objname);
	var tmpObj;
	if(focusY==ArrText.length-1){
		if(focusX>ArrText[0].length-1){
			tmpObj=eval("myForm."+ArrText[0][ArrText[0].length-1]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}else{
			tmpObj=eval("myForm."+ArrText[0][focusX]);
			tmpObjfocus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}
	}else{
		if(focusX>ArrText[focusY+1].length-1){
			tmpObj=eval("myForm."+ArrText[focusY+1][ArrText[focusY+1].length-1]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}else{
			tmpObj=eval("myForm."+ArrText[focusY+1][focusX]);
			tmpObj.focus();
			if(tmpObj.tagName=='INPUT'){tmpObj.select();}
		}
	}
}
function atutoEnter(objname){
	if(document.all[objname].getAttribute('maxlength')<1000){
		if(document.all[objname].value.length==document.all[objname].getAttribute('maxlength')){
			CodeEnter(objname);
		}
	}
}

function autoKeyEnd(){
	var e = event.srcElement;
	var r =e.createTextRange();
	r.moveStart("character",e.value.length);
	r.collapse(true);
	r.select();
}

function ListItem(UnitListName,MemListName){
	runServerScript("/traffic/Common/ListItem.asp?LoginID="+document.all["chekChMemID"].value+"&UnitListName="+UnitListName+"&MemListName="+MemListName);
}

function ListItemLaver(UnitListName,MemListName){
	runServerScript("/traffic/Common/LaverListItem.asp?LoginID="+document.all["chekLaverChMemID"].value+"&UnitListName="+UnitListName+"&MemListName="+MemListName);
}

function OpenWindow(str) {
// CDATEID
   var today;
   var date1;
   today = document.all[str].value;

   if (today!=''&&today.length>5){
		today = eval(document.all[str].value.substr(0,eval(document.all[str].value.length)-4))+1911+'/'+document.all[str].value.substr(eval(document.all[str].value.length)-4,2)+'/'+document.all[str].value.substr(eval(document.all[str].value.length)-2,2);
   }else{
	today='';
   }
   //For LNP 日期格式
   if (str == "TRANSFER_ACCEPT_DATE" || str == "NP_CHG_CONTRACT_DATE"){
     today = "";
   }
   date1=new Date(today);
   if(today.length != 0) 
     str = "/traffic/Common/date.asp?d=" +  date1.getDate(today) + "&m="  + (date1.getMonth(today)+1) + "&y=" + date1.getFullYear(today) + "&ClickName=" + str ;
   else
     str = "/traffic/Common/date.asp" + "?ClickName=" + str ; 
   
   
   sWindow=window.open(str,"awindow","scrollbars=no,left=300,top=280,status=yes,toolbar=no,width=280,height=240,resizable=no,menubar=no");
   sWindow.focus();
   sWindow.opener = self; 
}
function popup(indx,url){
	if(indx=='1'){
		pObj=window.createPopup();
		popObj=pObj.document.body;
		popObj.innerHTML="<img src='"+url+"' width='200' height='100'>";
		//popObj.style.backgroundColor="gray";
		pObj.show(event.screenX,event.screenY,200,100);
	}else{
		pObj.hide();
	}
}
function chkBillNo(item){
	if(myForm.item[item-1].value.length==9){
		return true;
	}else{
		return false;
	}
}
function setDate(D){
     var f = document.forms["search"];
     var Show_Month;
	   var Show_Day;
      m = f.elements['month_h'].value*1;
      y = f.elements['Year_h'].value*1;
      ClickName = f.elements['ClickName1'].value;
	  //----------將個位數補成十位數---------
	  Show_Month=new String(m);
      if (Show_Month.length< 2 ){
	      Show_Month='0'+Show_Month;
	  }
	  Show_Day=new String(D);
      if (Show_Day.length< 2 ){
	      Show_Day='0'+Show_Day;
	  }
	  //------------------------------------
     if (ClickName == "TRANSFER_ACCEPT_DATE" || ClickName == "NP_CHG_CONTRACT_DATE" ){
     window.opener.document.all(f.elements['ClickName1'].value).value  =  (y-1911) + Show_Month +  Show_Day ;
	 }else{
	   window.opener.document.all(f.elements['ClickName1'].value).value  = (y-1911) + Show_Month + Show_Day ;
	 }
	 window.close(); 
  }
function runServerScript(url){
	// Create new JS element 
	var js = document.createElement('SCRIPT');
	js.type = 'text/javascript';
	js.src = url;
	// Append JS element (therefore executing the 'AJAX' call)
	document.body.appendChild (js);

	return true;
}
function DateAdd(timeU,byMany,dateObj) {   
	var millisecond=1;   
	var second=millisecond*1000;   
	var minute=second*60;   
	var hour=minute*60;   
	var day=hour*24;   
	var year=day*365;  
	var newDate;
	var dVal=dateObj.valueOf();   
	var date=new Date(dateObj);  // For months
	switch(timeU) {   
		case "ms": newDate=new Date(dVal+millisecond*byMany); break;   
		case "s": newDate=new Date(dVal+second*byMany); break;   
		case "mi": newDate=new Date(dVal+minute*byMany); break;   
		case "h": newDate=new Date(dVal+hour*byMany); break;   
		case "d": newDate=new Date(dVal+day*byMany); break;   
		case "m": newDate=new Date(date.setMonth(date.getMonth() + byMany)); break;   //月
		case "y": newDate=new Date(dVal+year*byMany); break;   
	}   
	return newDate;   
}   

function DateAdd2(interval, number, date) {
  switch (interval.toLowerCase()) {
  case "y":
    date.setFullYear(date.getFullYear() + number);
    break;
  case "m":
    date.setMonth(date.getMonth() + number);
    break;
  case "d":
    date.setDate(date.getDate() + number);
    break;
  case "h":
    date.setHours(date.getHours() + number);
    break;
  case "n":
    date.setMinutes(date.getMinutes() + number);
    break;
  case "s":
    date.setSeconds(date.getSeconds() + number);
    break;
  default:
  }
  return date
}

function dateCheck(Sys_date){
	var error=0;
	var error2=1;
	var date_y=0,date_m=0,date_y=0;
	if(Sys_date.length>5){
		date_y=eval(Sys_date.substr(0,eval(Sys_date.length)-4))+1911;
		date_m=Sys_date.substr(eval(Sys_date.length)-4,2);
		date_d=Sys_date.substr(eval(Sys_date.length)-2,2);
		if (date_m > 0 && date_m < 13){
			if (date_d > 0 && date_d <32){
				error2=0;
				if (date_m == 4 || date_m == 6 || date_m == 9 || date_m == 11){
					if (date_d > 30){
						error=1;
					}
				}
				if(date_m == 2){
					if ((date_y) % 4 == 0){
						if (date_d > 29) error=1;
					}else{
						if (date_d > 28) error=1;
					}
				}
				if(error==0){
					return true;
				}
			}
		}
		if(error2==1){
			return false;
		}
	}else{
		return false;
	}
}
//違規事實
function setLawDetail(RuleOrder,LawDetail,level1)
{
	if (RuleOrder=="1"){
		if (LawDetail != ""){
			Layer1.innerHTML=LawDetail;
			document.myForm.ForFeit1.value=level1;
			TDLawErrorLog1=0;
		}else{
			Layer1.innerHTML=" ";
			document.myForm.ForFeit1.value="";
			TDLawErrorLog1=1;
		}
	}else if (RuleOrder=="2"){
		if (LawDetail != ""){
			Layer2.innerHTML=LawDetail;
			document.myForm.ForFeit2.value=level1;
			TDLawErrorLog2=0;
		}else{
			Layer2.innerHTML=" ";
			document.myForm.ForFeit2.value="";
			TDLawErrorLog2=1;
		}
	}else if (RuleOrder=="3"){
		if (LawDetail != ""){
			Layer3.innerHTML=LawDetail;
			document.myForm.ForFeit3.value=level1;
			TDLawErrorLog3=0;
		}else{
			Layer3.innerHTML=" ";
			document.myForm.ForFeit3.value="";
			TDLawErrorLog3=1;
		}
	}else if (RuleOrder=="4"){
		if (LawDetail != ""){
			Layer4.innerHTML=LawDetail;
			document.myForm.ForFeit4.value=level1;
			TDLawErrorLog4=0;
		}else{
			Layer4.innerHTML=" ";
			document.myForm.ForFeit4.value="";
			TDLawErrorLog4=1;
		}
	}
}
//到案處所
function setStationName(StationName){
	Layer5.innerHTML=StationName;
	if (StationName==""){
		TDStationErrorLog=1;
	}else{
		TDStationErrorLog=0;
	}
}
//舉發單位
function setUnitName(UnitName){
	Layer6.innerHTML=UnitName;
	if (UnitName==""){
		TDUnitErrorLog=1;
	}else{
		TDUnitErrorLog=0;
	}
}
//保管物品
function setFastenerName(FastenerOrder,FastenerName){
	if (FastenerOrder=="1"){
		Layer8.innerHTML=FastenerName;
		document.myForm.Fastener1Val.value=FastenerName;
		if (FastenerName==""){
			TDFastenerErrorLog1=1;
		}else{
			TDFastenerErrorLog1=0;
		}
	}else if (FastenerOrder=="2"){
		Layer9.innerHTML=FastenerName;
		document.myForm.Fastener2Val.value=FastenerName;
		if (FastenerName==""){
			TDFastenerErrorLog2=1;
		}else{
			TDFastenerErrorLog2=0;
		}
	}else if (FastenerOrder=="3"){
		Layer10.innerHTML=FastenerName;
		document.myForm.Fastener3Val.value=FastenerName;
		if (FastenerName==""){
			TDFastenerErrorLog3=1;
		}else{
			TDFastenerErrorLog3=0;
		}
	}
}
//違規地點
function setIllStreetName(AddressName){
	if (AddressName!=""){
		myForm.IllegalAddress.value=AddressName;
	}else{
		myForm.IllegalAddress.value="";
	}
}
//舉發人臂章號碼
function setMemName(MType,MemOrder,MemName,MemID,UnitID,UnitName,UTypeFlag){
	if (UTypeFlag=='1' && MemName!=""){
		alert("舉發人 " + MemName + " 隸屬於其他分局，請至『人員管理系統』，檢查員警資料是否正確後再做建檔!!");
	}
	if (MType=="Car")
	{
		if (MemOrder=="1"){
			Layer12.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID1.value="";
				document.myForm.BillMemName1.value="";
				TDMemErrorLog1=1;
			}else{
				document.myForm.BillMemID1.value=MemID;
				document.myForm.BillMemName1.value=MemName;
				TDMemErrorLog1=0;
				//if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				//}
			}
		}else if (MemOrder=="2"){
			Layer13.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID2.value="";
				document.myForm.BillMemName2.value="";
				TDMemErrorLog2=1;
			}else{
				document.myForm.BillMemID2.value=MemID;
				document.myForm.BillMemName2.value=MemName;
				TDMemErrorLog2=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
			}
		}else if (MemOrder=="3"){
			Layer14.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID3.value="";
				document.myForm.BillMemName3.value="";
				TDMemErrorLog3=1;
			}else{
				document.myForm.BillMemID3.value=MemID;
				document.myForm.BillMemName3.value=MemName;
				TDMemErrorLog3=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
			}
		}else if (MemOrder=="4"){
			Layer17.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID4.value="";
				document.myForm.BillMemName4.value="";
				TDMemErrorLog4=1;
			}else{
				document.myForm.BillMemID4.value=MemID;
				document.myForm.BillMemName4.value=MemName;
				TDMemErrorLog4=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
			}
		}
	}else{
		if (MemOrder=="1"){
			Layer12.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID1.value="";
				document.myForm.BillMemName1.value="";
				TDMemErrorLog1=1;
			}else{
				document.myForm.BillMemID1.value=MemID;
				document.myForm.BillMemName1.value=MemName;
				TDMemErrorLog1=0;
				//if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				//}
				/*
				if (document.myForm.MemberStation.value==""){
					document.myForm.MemberStation.value=UnitID;
					Layer5.innerHTML=UnitName;
					TDStationErrorLog=0;
				}
				*/
			}
		}else if (MemOrder=="2"){
			Layer13.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID2.value="";
				document.myForm.BillMemName2.value="";
				TDMemErrorLog2=1;
			}else{
				document.myForm.BillMemID2.value=MemID;
				document.myForm.BillMemName2.value=MemName;
				TDMemErrorLog2=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
				/*
				if (document.myForm.MemberStation.value==""){
					document.myForm.MemberStation.value=UnitID;
					Layer5.innerHTML=UnitName;
					TDStationErrorLog=0;
				}
				*/
			}
		}else if (MemOrder=="3"){
			Layer14.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID3.value="";
				document.myForm.BillMemName3.value="";
				TDMemErrorLog3=1;
			}else{
				document.myForm.BillMemID3.value=MemID;
				document.myForm.BillMemName3.value=MemName;
				TDMemErrorLog3=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
				/*
				if (document.myForm.MemberStation.value==""){
					document.myForm.MemberStation.value=UnitID;
					Layer5.innerHTML=UnitName;
					TDStationErrorLog=0;
				}
				*/
			}
		}else if (MemOrder=="4"){
			Layer17.innerHTML=MemName;
			if (MemName==""){
				document.myForm.BillMemID4.value="";
				document.myForm.BillMemName4.value="";
				TDMemErrorLog4=1;
			}else{
				document.myForm.BillMemID4.value=MemID;
				document.myForm.BillMemName4.value=MemName;
				TDMemErrorLog4=0;
				if (document.myForm.BillUnitID.value==""){
					document.myForm.BillUnitID.value=UnitID;
					Layer6.innerHTML=UnitName;
					TDUnitErrorLog=0;
				}
				/*
				if (document.myForm.MemberStation.value==""){
					document.myForm.MemberStation.value=UnitID;
					Layer5.innerHTML=UnitName;
					TDStationErrorLog=0;
				}
				*/
			}
		}
	}
}

function setPeoPleMemName(MType,MemOrder,MemName,MemID,LoginID,UnitID,UnitName){
	if (MemOrder=="1"){
		Layer12.innerHTML=LoginID;
		if (MemName==""){
			document.myForm.BillMemID1.value="";
			document.myForm.BillMemName1.value="";
			TDMemErrorLog1=1;
		}else{
			document.myForm.BillMemID1.value=MemID;
			document.myForm.BillMemName1.value=MemName;
			TDMemErrorLog1=0;
			//if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			//}
			/*
			if (document.myForm.MemberStation.value==""){
				document.myForm.MemberStation.value=UnitID;
				Layer5.innerHTML=UnitName;
				TDStationErrorLog=0;
			}
			*/
		}
	}else if (MemOrder=="2"){
		Layer13.innerHTML=LoginID;
		if (MemName==""){
			document.myForm.BillMemID2.value="";
			document.myForm.BillMemName2.value="";
			TDMemErrorLog2=1;
		}else{
			document.myForm.BillMemID2.value=MemID;
			document.myForm.BillMemName2.value=MemName;
			TDMemErrorLog2=0;
			//if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			//}
			/*
			if (document.myForm.MemberStation.value==""){
				document.myForm.MemberStation.value=UnitID;
				Layer5.innerHTML=UnitName;
				TDStationErrorLog=0;
			}
			*/
		}
	}else if (MemOrder=="3"){
		Layer14.innerHTML=LoginID;
		if (MemName==""){
			document.myForm.BillMemID3.value="";
			document.myForm.BillMemName3.value="";
			TDMemErrorLog3=1;
		}else{
			document.myForm.BillMemID3.value=MemID;
			document.myForm.BillMemName3.value=MemName;
			TDMemErrorLog3=0;
			//if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			//}
			/*
			if (document.myForm.MemberStation.value==""){
				document.myForm.MemberStation.value=UnitID;
				Layer5.innerHTML=UnitName;
				TDStationErrorLog=0;
			}
			*/
		}
	}else if (MemOrder=="4"){
		Layer17.innerHTML=LoginID;
		if (MemName==""){
			document.myForm.BillMemID4.value="";
			document.myForm.BillMemName4.value="";
			TDMemErrorLog4=1;
		}else{
			document.myForm.BillMemID4.value=MemID;
			document.myForm.BillMemName4.value=MemName;
			TDMemErrorLog4=0;
			//if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			//}
			/*
			if (document.myForm.MemberStation.value==""){
				document.myForm.MemberStation.value=UnitID;
				Layer5.innerHTML=UnitName;
				TDStationErrorLog=0;
			}
			*/
		}
	}
}

//舉發人臂章號碼(判斷舉發人是否同分局)
function setMemName2(MType,MemOrder,MemName,MemID,UnitID,UnitName,UnitTypeID,UTypeFlag){
	if (UTypeFlag=='1' && MemName!=""){
		alert("舉發人 " + MemName + " 隸屬於其他分局，請至『人員管理系統』，檢查員警資料是否正確後再做建檔!!");
	}
	if (MemOrder=="1"){
		Layer12.innerHTML=MemName;
		if (MemName==""){
			document.myForm.BillMemID1.value="";
			document.myForm.BillMemName1.value="";
			document.myForm.BillUnitTypeID1.value="";
			TDMemErrorLog1=1;
		}else{
			document.myForm.BillMemID1.value=MemID;
			document.myForm.BillMemName1.value=MemName;
			TDMemErrorLog1=0;
			//if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				document.myForm.BillUnitTypeID1.value=UnitTypeID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			//}
			if (document.myForm.BillUnitTypeID2.value!="" && document.myForm.BillUnitTypeID1.value!=document.myForm.BillUnitTypeID2.value){
				alert("舉發人1與舉發人2屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID3.value!="" && document.myForm.BillUnitTypeID1.value!=document.myForm.BillUnitTypeID3.value){
				alert("舉發人1與舉發人3屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID4.value!="" && document.myForm.BillUnitTypeID1.value!=document.myForm.BillUnitTypeID4.value){
				alert("舉發人1與舉發人4屬於不同分局!!");
			}
		}
	}else if (MemOrder=="2"){
		Layer13.innerHTML=MemName;
		if (MemName==""){
			document.myForm.BillMemID2.value="";
			document.myForm.BillMemName2.value="";
			document.myForm.BillUnitTypeID2.value="";
			TDMemErrorLog2=1;
		}else{
			document.myForm.BillMemID2.value=MemID;
			document.myForm.BillMemName2.value=MemName;
			document.myForm.BillUnitTypeID2.value=UnitTypeID;
			TDMemErrorLog2=0;
			if (document.myForm.BillUnitTypeID1.value!="" && document.myForm.BillUnitTypeID2.value!=document.myForm.BillUnitTypeID1.value){
				alert("舉發人2與舉發人1屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID3.value!="" && document.myForm.BillUnitTypeID2.value!=document.myForm.BillUnitTypeID3.value){
				alert("舉發人2與舉發人3屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID4.value!="" && document.myForm.BillUnitTypeID2.value!=document.myForm.BillUnitTypeID4.value){
				alert("舉發人2與舉發人4屬於不同分局!!");
			}
			if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			}
		}
	}else if (MemOrder=="3"){
		Layer14.innerHTML=MemName;
		if (MemName==""){
			document.myForm.BillMemID3.value="";
			document.myForm.BillMemName3.value="";
			document.myForm.BillUnitTypeID3.value="";
			TDMemErrorLog3=1;
		}else{
			document.myForm.BillMemID3.value=MemID;
			document.myForm.BillMemName3.value=MemName;
			document.myForm.BillUnitTypeID3.value=UnitTypeID;
			TDMemErrorLog3=0;
			if (document.myForm.BillUnitTypeID1.value!="" && document.myForm.BillUnitTypeID3.value!=document.myForm.BillUnitTypeID1.value){
				alert("舉發人3與舉發人1屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID2.value!="" && document.myForm.BillUnitTypeID3.value!=document.myForm.BillUnitTypeID2.value){
				alert("舉發人3與舉發人2屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID4.value!="" && document.myForm.BillUnitTypeID3.value!=document.myForm.BillUnitTypeID4.value){
				alert("舉發人3與舉發人4屬於不同分局!!");
			}
			if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			}
		}
	}else if (MemOrder=="4"){
		Layer17.innerHTML=MemName;
		if (MemName==""){
			document.myForm.BillMemID4.value="";
			document.myForm.BillMemName4.value="";

document.myForm.BillUnitTypeID4.value="";
			TDMemErrorLog4=1;
		}else{
			document.myForm.BillMemID4.value=MemID;
			document.myForm.BillMemName4.value=MemName;
			document.myForm.BillUnitTypeID4.value=UnitTypeID;
			TDMemErrorLog4=0;
			if (document.myForm.BillUnitTypeID1.value!="" && document.myForm.BillUnitTypeID4.value!=document.myForm.BillUnitTypeID1.value){
				alert("舉發人4與舉發人1屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID2.value!="" && document.myForm.BillUnitTypeID4.value!=document.myForm.BillUnitTypeID2.value){
				alert("舉發人4與舉發人2屬於不同分局!!");
			}else if (document.myForm.BillUnitTypeID3.value!="" && document.myForm.BillUnitTypeID4.value!=document.myForm.BillUnitTypeID3.value){
				alert("舉發人4與舉發人3屬於不同分局!!");
			}
			if (document.myForm.BillUnitID.value==""){
				document.myForm.BillUnitID.value=UnitID;
				Layer6.innerHTML=UnitName;
				TDUnitErrorLog=0;
			}
		}
	}
}
//用固定桿編號抓出違規地點
function setFixIDAddress(FixNum,FixAddr,FixStreetID)
{
	document.myForm.FixID.value=FixNum;
	if (document.myForm.IllegalAddress.value=="")
	{
		document.myForm.IllegalAddressID.value=FixStreetID;
		document.myForm.IllegalAddress.value=FixAddr;
	}
	//document.myForm.BillMemID2.value=FixNum;
}
//檢查違規日期是否為三個月前(90天)
function ChkIllegalDate(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-90,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//檢查違規日期是否為二個月前(60天) 109/12/1逕舉改60天
function ChkIllegalDate60_109(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-60,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

//檢查違規日期是否為二個月前(高雄) 
function ChkIllegalDate2M_KS(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-1,(DateAdd("m",2,IFillDate)));
	if (thisDay > OverDate){
		return false;
	}else{
		return true;
	}
}

//檢查違規日期是否為30天前(30天)
function ChkIllegalDate30(IllDate){
	Iyear=parseInt(IllDate.substr(0,IllDate.length-4))+1911;
	Imonth=IllDate.substr(IllDate.length-4,2);
	Iday=IllDate.substr(IllDate.length-2,2);
	var IFillDate=new Date(Iyear,Imonth-1,Iday);
	var thisDay=new Date((new Date()).getFullYear(),(new Date()).getMonth(),(new Date()).getDate());
	var OverDate=new Date();
	OverDate=DateAdd("d",-30,thisDay);
	if (OverDate > IFillDate){
		return false;
	}else{
		return true;
	}
}

function UnitMan(UnitName,UnitMem,tmpMemID){
	var now = new Date();

	if(document.all[UnitName].value!=''){
		runServerScript("/traffic/Common/UnitAdd.asp?UnitID="+document.all[UnitName].value+"&UnitMem="+UnitMem+"&MemberID="+tmpMemID+"&nowdate="+now);
	}else{
		document.all[UnitMem].options[1]=new Option('所有人','');
		document.all[UnitMem].length=1;
	}
}

function UnitLaverMan(UnitName,UnitMem,tmpMemID){
	var now = new Date();

	if(document.all[UnitName].value!=''){
		runServerScript("/traffic/Common/UnitLaverMan.asp?UnitID="+document.all[UnitName].value+"&UnitMem="+UnitMem+"&MemberID="+tmpMemID+"&nowdate="+now);
	}else{
		document.all[UnitMem].options[1]=new Option('所有人','');
		document.all[UnitMem].length=1;
	}
}

function printWindow(printStyle,printleft,printtop,printright,printbot) {
	factory.printing.header = "";
	factory.printing.footer = "";
	factory.printing.portrait = printStyle;
	factory.printing.leftMargin = printleft;
	factory.printing.topMargin = printtop;
	factory.printing.rightMargin = printright;
	factory.printing.bottomMargin = printbot;
	factory.printing.Print(false);
	return true;
}

function chknumber(obj){
	var tmpvalue=obj.value.replace(/[^\d]/g,'');
	if (obj.value!=tmpvalue){
		obj.value=tmpvalue;
	}
}

//轉換為西元年
function transferAD(obj)
{
	Iyear = parseInt(obj.substr(0, obj.length - 4)) + 1911;
	Imonth = obj.substr(obj.length - 4, 2);
	Iday = obj.substr(obj.length - 2, 2);
	var transDate = new Date(Iyear, Imonth - 1, Iday);
	return transDate;
}



//判斷空值邏輯
//true 為空值、false為有值
function chkNull(value) {
	return value === null || value === undefined || value === '';
}