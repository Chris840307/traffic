// JScript 檔
//日期格式判斷
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

function printWindow(portraitVal,leftMarginVal,topMarginVal,rightMarginVal,bottomMargin) {
    factory.printing.header = "";
    factory.printing.footer = "";
    factory.printing.portrait = portraitVal;
    factory.printing.leftMargin = leftMarginVal;
    factory.printing.topMargin = topMarginVal;
    factory.printing.rightMargin = rightMarginVal;
    factory.printing.bottomMargin = bottomMargin;
    factory.printing.Print(false);
    return true;
}