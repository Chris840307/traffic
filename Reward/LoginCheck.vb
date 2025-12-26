    Sub LoginCheck()
        If (Request.Cookies("UserFunction") IsNot Nothing) Then
            Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
            If Trim(FuncCookie.Values("FuncID")) = "" Then
            Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
            End If
        Else
        Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
        End If
    End Sub
    Public Sub AuthorityCheck(ByVal FID)
        Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
        Dim FuncIDtemp As String = Trim(FuncCookie.Values("FuncID"))

        Dim FunctionTemp = Split(FuncIDtemp, "@@")
        Dim FuncStatus As Integer = 0
        Dim FunctionCountValue As Integer
        Dim ATemp
        For FunctionCountValue = 0 To UBound(FunctionTemp)
            ATemp = Split(Trim(FunctionTemp(FunctionCountValue)), ",")
            'Response.Write(FID & ATemp(0) & "," & FuncStatus & "<br>")
            If Trim(ATemp(0)) = Trim(FID) Then
                FuncStatus = 1
                Exit For
                'Response.Write(FID & ATemp(0) & "<br>")
            End If
        Next
        If FuncStatus = 0 Then
        Response.Redirect("/traffic/Reward/Login.aspx?ErrMsg=1")
        End If
    End Sub
    '檢查是否有查詢新增等權限
    Public Function CheckPermission(ByVal FunctionID, ByVal ActionID) As Boolean
        'ActionID 查詢:1
        '		  新增:2
        '		  修改:3
        '		  刪除:4
        Dim FuncCookie As HttpCookie = Request.Cookies("UserFunction")
        Dim FuncIDtemp As String = Trim(FuncCookie.Values("FuncID"))
        Dim FunctionTemp = Split(FuncIDtemp, "@@")
        Dim FuncStatus As Integer = 0
        Dim FunctionCountValue As Integer
        Dim ATemp
        For FunctionCountValue = 0 To UBound(FunctionTemp)
            ATemp = Split(Trim(FunctionTemp(FunctionCountValue)), ",")
            If Trim(ATemp(0)) = Trim(FunctionID) Then
                If ATemp(Trim(ActionID)) = "1" Then
                    CheckPermission = True
                Else
                    CheckPermission = False
                End If
            End If
        Next
    End Function
