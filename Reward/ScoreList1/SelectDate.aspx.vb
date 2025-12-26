
Partial Class ScoreList_SelectDate
    Inherits System.Web.UI.Page

    Protected Sub Calendar1_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar1.SelectionChanged
        Dim DateTemp As String = ""
        DateTemp = gInitDT(Calendar1.SelectedDate)
        Response.Write("<script language=""JavaScript"">")
        Response.Write("opener.form1." & Trim(Request("tag")) & ".value=""" & DateTemp & """;")
        Response.Write("window.close();")
        Response.Write("</script>")

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()
        If Trim(Request("InitDate")) <> "" And Len(Trim(Request("InitDate"))) > 5 Then
            Calendar1.VisibleDate = "#" & gOutDT(Trim(Request("InitDate"))) & "#"
        End If
    End Sub

    '將西元年轉換為民國yymmdd
    Public Function gInitDT(ByVal iDate)
        Dim DatetTemp As String
        If iDate IsNot DBNull.Value Then
            DatetTemp = Year(iDate) - 1911 & Right("0" & Month(iDate), 2) & Right("0" & Day(iDate), 2)
            gInitDT = DatetTemp
        Else
            gInitDT = ""
        End If
    End Function

    '將民國yymmdd轉換為yyyy/mm/dd
    Public Function gOutDT(ByVal iDate)
        Dim DatetTemp As String
        If iDate IsNot DBNull.Value Then
            DatetTemp = DateSerial(Left(iDate, Len(iDate) - 4) + 1911, Mid(iDate, Len(iDate) - 3, 2), Right(iDate, 2))
            gOutDT = DatetTemp
        Else
            gOutDT = ""
        End If
    End Function

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
End Class
