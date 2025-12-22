<%
    Set Codetab = CreateObject("Scripting.Dictionary") '檢查碼對照表 
	Sub CodeMapping()    '建立檢查碼對照表	   
	    Codetab.RemoveAll()
        Dim i, j
        j = 1
        For i = 0 To 25
            If Chr(65 + i) = "J" Then j = 1
            If Chr(65 + i) = "S" Then j = 2
            Codetab.Add Chr(65 + i) , j
            j = j + 1
        Next
        Codetab.Add "+","1"
        Codetab.Add "%","2"
        Codetab.Add "-","6"
        Codetab.Add ".","7"
        Codetab.Add " ","8"
        Codetab.Add "$","9"
        Codetab.Add "/","0"
    End Sub
   
    Function CreateChkCode(deadline,item,BillNo,paydate,money)
         CodeMapping()
        Barcode1 = deadline & item
        Barcode2 = BillNo
        Barcode3 = paydate & money
        '愚蠢的方法
        sum1 = cint(GetNumber(mid(Barcode1,1,1))) + GetNumber(mid(BarCode1,3,1))+ GetNumber(mid(BarCode1,5,1)) + GetNumber(mid(BarCode1,7,1)) + GetNumber(mid(BarCode1,9,1)) +_
                  GetNumber(mid(Barcode2,1,1)) + GetNumber(mid(BarCode2,3,1)) + GetNumber(mid(BarCode2,5,1)) + GetNumber(mid(BarCode2,7,1)) + GetNumber(mid(BarCode2,9,1)) +_
                  GetNumber(mid(Barcode2,11,1)) + GetNumber(mid(BarCode2,13,1)) + GetNumber(mid(BarCode2,15,1))+_
                  GetNumber(mid(Barcode3,1,1)) + GetNumber(mid(BarCode3,3,1)) + GetNumber(mid(BarCode3,5,1)) + GetNumber(mid(BarCode3,7,1)) + GetNumber(mid(BarCode3,9,1)) +_                  
                  GetNumber(mid(Barcode3,11,1)) + GetNumber(mid(BarCode3,13,1))
        sum2 = cint(GetNumber(mid(Barcode1,2,1))) + GetNumber(mid(BarCode1,4,1))+ GetNumber(mid(BarCode1,6,1)) + GetNumber(mid(BarCode1,8,1)) +_
                  GetNumber(mid(Barcode2,2,1)) + GetNumber(mid(BarCode2,4,1)) + GetNumber(mid(BarCode2,6,1)) + GetNumber(mid(BarCode2,8,1)) + GetNumber(mid(BarCode2,10,1)) +_
                  GetNumber(mid(Barcode2,12,1)) + GetNumber(mid(BarCode2,14,1)) + GetNumber(mid(BarCode2,16,1))+_
                  GetNumber(mid(Barcode3,2,1)) + GetNumber(mid(BarCode3,4,1)) + GetNumber(mid(BarCode3,6,1)) + GetNumber(mid(BarCode3,8,1)) + GetNumber(mid(BarCode3,10,1)) +_                  
                  GetNumber(mid(Barcode3,12,1))
                  
        CreateChkCode = GetChkCode(sum1,1) & GetChkCode(sum2,0)
    End Function  
   
  Function GetNumber(str)
                If Not IsNumeric(str) Then
                   GetNumber = Codetab.Item(str)
                Else
                    GetNumber = str
                End If
  end Function  
     
    Function GetChkCode(sumvalue, state)
    If state = 1 Then  '算奇數位的總合並除以11求餘數
        If sumvalue Mod 11 = 10 Then
            GetChkCode = "B"
        ElseIf sumvalue Mod 11 = 0 Then
            GetChkCode = "A"
        Else
            GetChkCode = CStr(sumvalue Mod 11)
        End If
    Else    '算偶數位的總合並除以11求餘數
        If sumvalue Mod 11 = 10 Then
            GetChkCode = "Y"
        ElseIf sumvalue Mod 11 = 0 Then
            GetChkCode = "X"
        Else
            GetChkCode = CStr(sumvalue Mod 11)
        End If
    End If
    End Function  
   
   'CodeMapping() 
   'for i=0 to 400
    '    response.Write(CreateChkCode("960604","241","0000000000050593","0522","000000054"))
        'response.Write(CreateChkCode("991231","Y01","ABCDEFGHIKLMNPQR","1234","000007890"))
   'next 
 %>
