
' VBScript 檔
'將西元年轉換為民國yymmdd
Public Function gInitDT(ByVal iDate)
    Dim DatetTemp As String
    If iDate IsNot DBNull.Value Then
        DatetTemp = year(iDate) - 1911 & right("0" & month(iDate), 2) & right("0" & day(iDate), 2)
        gInitDT = DatetTemp
    Else
        gInitDT = ""
    End If
End Function
