
Partial Class ScoreList_getScoreList_SubUnit_set
    Inherits System.Web.UI.Page

    Public strUnit As String = "select * from UnitInfo order by UnitID"
    Public checkFlag As Char = ""
    Public ErrorCode As String = ""
    Public UnitFlag As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()
        Dim UserCookie As HttpCookie = Request.Cookies("RewardUser")

        'LEVEL1能看全部單位  LEVEL2能看所屬單位 LEVEL3只能看自己
        If Trim(UserCookie.Values("UnitLevelID")) = 1 Then
            strUnit = "select UnitID,UnitName from UnitInfo order by UnitID"
        ElseIf Trim(UserCookie.Values("UnitLevelID")) = 2 Then
            strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' or UnitTypeID='" & Trim(UserCookie.Values("UnitID")) & "' order by UnitID"
        Else
            strUnit = "select UnitID,UnitName from UnitInfo where UnitID='" & Trim(UserCookie.Values("UnitID")) & "' order by UnitID"
        End If

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
