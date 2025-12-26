
Partial Class LawScore_LawScoreSetAll
    Inherits System.Web.UI.Page
    Public sys_City As String

    Sub LawScoreUpdate(ByVal sender As Object, ByVal e As System.EventArgs)
        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()

        If Trim(Request("SecoreType")) = "" Then
            Response.Write("<script language=""JavaScript"">")
            Response.Write("alert(""請輸入法條類別!!"");")
            Response.Write("</script>")
        ElseIf Trim(Request("tScore")) = "" Then
            Response.Write("<script language=""JavaScript"">")
            Response.Write("alert(""請輸入配分!!"");")
            Response.Write("</script>")
        Else
            Dim strSType As String = ""
            'If Trim(Request("SecoreType")) = "" Then
            'strSType = " BillType1Score='" & Trim(Request("tScore")) & "',BillType2Score='" & Trim(Request("tScore")) & "'"
            'strSType = strSType & ",A1Score='" & Trim(Request("tScore")) & "',A2Score='" & Trim(Request("tScore")) & "'"
            'strSType = strSType & ",A3Score='" & Trim(Request("tScore")) & "'"

            If InStr(Trim(Request("SecoreType")), "1") <> 0 Then
                strSType = " BillType1Score='" & Trim(Request("tScore")) & "'"
            End If
            If InStr(Trim(Request("SecoreType")), "2") <> 0 Then
                If strSType <> "" Then
                    strSType = strSType & ", BillType2Score='" & Trim(Request("tScore")) & "'"
                Else
                    strSType = " BillType2Score='" & Trim(Request("tScore")) & "'"
                End If
            End If
            If InStr(Trim(Request("SecoreType")), "3") <> 0 Then
                If strSType <> "" Then
                    strSType = strSType & ", A1Score='" & Trim(Request("tScore")) & "'"
                Else
                    strSType = " A1Score='" & Trim(Request("tScore")) & "'"
                End If
            End If
            If InStr(Trim(Request("SecoreType")), "4") <> 0 Then
                If strSType <> "" Then
                    strSType = strSType & ", A2Score='" & Trim(Request("tScore")) & "'"
                Else
                    strSType = " A2Score='" & Trim(Request("tScore")) & "'"
                End If
            End If
            If InStr(Trim(Request("SecoreType")), "5") <> 0 Then
                If strSType <> "" Then
                    strSType = strSType & ", A3Score='" & Trim(Request("tScore")) & "'"
                Else
                    strSType = " A3Score='" & Trim(Request("tScore")) & "'"
                End If
            End If
            If sys_City = "雲林縣" Or sys_City = "宜蘭縣" Or sys_City = "台東縣" Or sys_City = "嘉義市" Then
                If InStr(Trim(Request("SecoreType")), "6") <> 0 Then
                    If strSType <> "" Then
                        strSType = strSType & ", Other1='" & Trim(Request("tScore")) & "'"
                    Else
                        strSType = " Other1='" & Trim(Request("tScore")) & "'"
                    End If
                End If
            End If

            Dim strSqlPlus As String = ""
            '抓現行法條代碼
            Dim strLawVer = "select Value from ApConfigure where ID=3"
            Dim CmdLawVer As New Data.OracleClient.OracleCommand(strLawVer, conn)
            Dim rdLawVer As Data.OracleClient.OracleDataReader = CmdLawVer.ExecuteReader()
            If rdLawVer.HasRows Then
                rdLawVer.Read()
                strSqlPlus = " where LawVersion='" & Trim(rdLawVer("Value")) & "'"

            End If
            rdLawVer.Close()

            If Trim(Request("sCountyOrNpa1")) <> "n" Then
                If strSqlPlus <> "" Then
                    strSqlPlus = strSqlPlus & " and CountyOrNpa=" & Trim(Request("sCountyOrNpa1"))
                Else
                    strSqlPlus = " where CountyOrNpa=" & Trim(Request("sCountyOrNpa1"))
                End If
            End If

            If Trim(Request("LawRange")) = "1" Then
                strSqlPlus = strSqlPlus
            ElseIf Trim(Request("LawRange")) = "2" Then
                If strSqlPlus <> "" Then
                    strSqlPlus = strSqlPlus & " and substr(LawItem,1,2) between '1' and '68'"
                Else
                    strSqlPlus = " where substr(LawItem,1,2) between '1' and '68'"
                End If
            ElseIf Trim(Request("LawRange")) = "3" Then
                If strSqlPlus <> "" Then
                    strSqlPlus = strSqlPlus & " and substr(LawItem,1,2) > '68'"
                Else
                    strSqlPlus = " where substr(LawItem,1,2) > '68'"
                End If
            End If

            Dim strSql As String = "Update LawScore set " & strSType & strSqlPlus
            Dim cmdUpd As New Data.OracleClient.OracleCommand()
            cmdUpd.CommandText = strSql
            cmdUpd.Connection = conn
            cmdUpd.ExecuteNonQuery()

            Response.Write("<script language=""JavaScript"">")
            Response.Write("alert(""儲存成功!!"");")
            Response.Write("</script>")
            'Response.Write(strSql)
            End If

            conn.Close()
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoginCheck()

        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")
        '建立 Connection 物件
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()
        sys_City = ""
        Dim strCity = "select Value from ApConfigure where ID=31"
        Dim CmdCity As New Data.OracleClient.OracleCommand(strCity, conn)
        Dim rdCity As Data.OracleClient.OracleDataReader = CmdCity.ExecuteReader()
        If rdCity.HasRows Then
            rdCity.Read()
            sys_City = Trim(rdCity("Value"))

        End If
        rdCity.Close()

        conn.Close()
    End Sub
End Class
