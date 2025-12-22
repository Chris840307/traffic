function ValidateBillNumbers(startObj,endObj,isCount){
  var startValue = startObj.value;
  var endValue = endObj.value;
  var tmpCnt=null;

  for(i=startValue.length;i>=0;i--){
	  if(isNaN(startValue.substr(i,1))){
		  tmpCnt=i;
		  break;
	  }
  }

  var startHead = startValue.substring(0,(tmpCnt+1));
  var startTail = startValue.substring((tmpCnt+1));
  var endHead = endValue.substring(0,(tmpCnt+1));
  var endTail = endValue.substring((tmpCnt+1));
  var numStart = Number(startTail);
  var intStart = Number(numStart);
  var numEnd = Number(endTail);
  var intEnd = Number(numEnd);
  if (startHead != endHead){
  	 return 1;
  }else if (intStart > intEnd){
  	 return 2;
  }else if (isCount=='Y'){
    document.all.BillCount.value = eval(intEnd - intStart);
  }
	return 0;
}