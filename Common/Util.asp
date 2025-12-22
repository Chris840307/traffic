<%
Function GetDciTypeById(Id)
	   Select Case Cstr(Id)
	   	  Case "1" :
	   	    GetDciTypeById = "刪除原因"
	   	  Case "2" :
	   	    GetDciTypeById = "舉發單類型"
	   	  Case "4" :
	   	    GetDciTypeById = "車輛顏色"
	   	  Case "5" :
	   	    GetDciTypeById = "DCI車種代號"
	   	  Case "6" :
	   	    GetDciTypeById = "扣件物品"
	   	  Case "7" :
	   	    GetDciTypeById = "退件原因"
	   	  Case "8" :
	   	    GetDciTypeById = "是否有保險證"  	
	   End Select   
End Function

Function GetDCIActionNameById(Id)
	   Select Case Cstr(Id)
	   	  Case "A" :
	   	    GetDCIActionNameById = "查詢車籍"
	   	  Case "W" :
	   	    GetDCIActionNameById = "入案"
	   	  Case "WE" :
	   	    GetDCIActionNameById = "入案錯誤"
	   	  Case "N" :
	   	    GetDCIActionNameById = "送達註記"
	   	  Case "E" :
	   	    GetDCIActionNameById = "刪除資料"
	   End Select   
End Function

Function ReturnPermission(bolAudit)
     If Not bolAudit Then
     	  ReturnPermission = " disabled "
     Else
     	  ReturnPermission = ""
     End If  
End Function
%>