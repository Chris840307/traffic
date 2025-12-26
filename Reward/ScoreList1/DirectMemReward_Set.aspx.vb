
Partial Class ScoreList1_DirectMemReward_Set
    Inherits System.Web.UI.Page
    Public strUnit As String = "select * from UnitInfo order by UnitID"
    Public checkFlag As Char = ""
    Public ErrorCode As String = ""
    Public UnitFlag As String = ""
    Public sys_City As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()


    End Sub

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
