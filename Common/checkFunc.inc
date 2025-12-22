<Script Language="VBScript">
Function ValidateBillNumbers(startObj,endObj,isCount)
  startValue = startObj.value
  endValue = endObj.value
  startHead = Mid(startValue,1,2)
  startTail = Mid(startValue,3,7)
  endHead = Mid(endValue,1,2)
  endTail = Mid(endValue,3,7)
  numStart = FormatNumber(startTail,0)
  intStart = Int(numStart)
  numEnd = FormatNumber(endTail,0)
  intEnd = Int(numEnd)  
  ValidateBillNumbers = 0
  
  if (startHead <> endHead) then
  	 ValidateBillNumbers = 1
  	 Exit function
  end if
  if (intStart > intEnd) then
  	 ValidateBillNumbers = 2
  	 
  	 Exit function  	
  end if
  if (isCount="Y") then
    document.all.BillCount.value = (intEnd - intStart) + 1
  end if
End Function
</script>	