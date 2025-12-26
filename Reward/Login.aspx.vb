

Partial Class _Default
    Inherits System.Web.UI.Page

Public Function Decrypt(text As String) As String
    If String.IsNullOrWhiteSpace(text) Then
        If text Is Nothing Then
            Return Nothing
        Else
            Return text.Trim()
        End If
    Else
        Const key As String = "HD"
        Dim str As String = key & text
        Dim str1 As New System.Text.StringBuilder()

        For i As Integer = 0 To str.Length - 1
            Dim str2 As String = AscW(str(i)).ToString("X4") ' 4位16進制，不足補0
            str1.Append(str2)
        Next

        ' 反轉字符串
        Dim arr = str1.ToString().ToCharArray()
        Array.Reverse(arr)
        Return New String(arr)
    End If
End Function


    Sub User_Check(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim UserCreditID, UserPW As String

        '------------建立cookie-------------
        '使用者資料
        Response.Cookies("RewardUser")("CreditID") = ""
        Response.Cookies("RewardUser")("Password") = ""
        Response.Cookies("RewardUser")("CreditID") = ""
        Response.Cookies("RewardUser")("LoginID") = ""
        Response.Cookies("RewardUser")("UnitID") = ""
        Response.Cookies("RewardUser")("ChName") = ""
        Response.Cookies("RewardUser")("GroupRoleID") = ""
        Response.Cookies("RewardUser")("UnitLevelID") = ""
        Response.Cookies("RewardUser")("ManagerPower") = ""
        Response.Cookies("RewardUser")("DCIwindowName") = ""
        Response.Cookies("RewardUser")("DoubleCheck") = ""
        '使用者權限
        Response.Cookies("UserFunction")("FuncID") = ""
        '--------------------------------------

        If Trim(Request("UserID")) = "" Or Trim(Request("UserPW")) = "" Then
            Response.Write("<script language=""JavaScript"">")
            Response.Write("alert(""請輸入帳號密碼!!"");")
            Response.Write("</script>")
        Else
            ErrorMsg.Text = ""
            UserCreditID = Replace(Request("UserID"), "'", "")
            '取得 Web.config 檔的資料連接設定
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("DB_Orcl")

            '建立 Connection 物件
            Dim conn As New Data.OracleClient.OracleConnection
            conn.ConnectionString = setting.ConnectionString
            '開啟資料連接
            conn.Open()
            '建立 Command 物件
            Dim str1 = "select * from MemberData where CreditID='" & Trim(UserCreditID) & "' and AccountStateID=0 and RecordstateID=0"
            Dim Cmd1 As New Data.OracleClient.OracleCommand(str1, conn)

            '建立 DataReader 物件
            Dim rd1 As Data.OracleClient.OracleDataReader = Cmd1.ExecuteReader()
            If rd1.HasRows Then
                rd1.Read()
                If Trim(rd1("PassWord")) <> Trim(Request("UserPW")) Then
                    Response.Write("<script language=""JavaScript"">")
                    Response.Write("alert(""使用者密碼錯誤!!"");")
                    Response.Write("</script>")
                Else
                    '驗證正確，開始抓使用者資料寫入cookies
                    Dim FuncTemp As String = ""
                    Dim strFunc = "select * from functionDataDetail where GroupID='" & Trim(rd1("GrouproleID")) & "'"
                    Dim CmdFunc As New Data.OracleClient.OracleCommand(strFunc, conn)
                    Dim rdFunc As Data.OracleClient.OracleDataReader = CmdFunc.ExecuteReader()
                    While rdFunc.Read()
                        If FuncTemp = "" Then
                            FuncTemp = Trim(rdFunc("systemID")) & "," & Trim(rdFunc("SelectFlag")) & "," & Trim(rdFunc("InsertFlag")) & "," & Trim(rdFunc("UpdateFlag")) & "," & Trim(rdFunc("DeleteFlag"))
                        Else
                            FuncTemp = FuncTemp & "@@" & Trim(rdFunc("systemID")) & "," & Trim(rdFunc("SelectFlag")) & "," & Trim(rdFunc("InsertFlag")) & "," & Trim(rdFunc("UpdateFlag")) & "," & Trim(rdFunc("DeleteFlag"))
                        End If
                    End While
                    rdFunc.Close()
                    Response.Cookies("UserFunction")("FuncID") = FuncTemp
                    Response.Cookies("RewardUser")("CreditID") = Trim(rd1("CreditID"))
                    Response.Cookies("RewardUser")("Password") = Trim(rd1("Password"))
                    Response.Cookies("RewardUser")("MemberID") = Trim(rd1("MemberID"))
                    Response.Cookies("RewardUser")("LoginID") = Trim(rd1("LoginID"))
                    Response.Cookies("RewardUser")("UnitID") = Trim(rd1("UnitID"))
                    Response.Cookies("RewardUser")("ChName") = HttpUtility.UrlEncode(Trim(rd1("ChName")))
                    Response.Cookies("RewardUser")("GroupRoleID") = Trim(rd1("GroupRoleID"))
                    Response.Cookies("RewardUser")("ManagerPower") = Trim(rd1("ManagerPower"))

                    '單位等級
                    Dim strUnit = "select UnitLevelID,DCIwindowName from UnitInfo where UnitID='" & Trim(rd1("UnitID")) & "'"
                    Dim CmdUnit As New Data.OracleClient.OracleCommand(strUnit, conn)
                    Dim rdUnit As Data.OracleClient.OracleDataReader = CmdUnit.ExecuteReader()
                    If rdUnit.HasRows Then
                        rdUnit.Read()
                        Response.Cookies("RewardUser")("UnitLevelID") = Trim(rdUnit("UnitLevelID"))
                        'Response.Cookies("RewardUser")("DCIwindowName") = Trim(rdUnit("DCIwindowName"))
                    End If
                    rdUnit.Close()

                    '一打一驗判斷
                    Dim DoubleChk As String = "0"
                    Dim strDbChk = "select Value from Apconfigure where ID=38"
                    Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
                    Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
                    If rdDbChk.HasRows Then
                        rdDbChk.Read()
                        Response.Cookies("RewardUser")("DoubleCheck") = Trim(rdDbChk("Value"))
                    End If
                    rdDbChk.Close()
                    rd1.Close()
                    conn.Close()
                    Response.Redirect("Main.aspx")
                    'Response.Write(FuncTemp & "<br>")
                    'Dim acok As HttpCookie = Request.Cookies("UserFunction")
                    'Response.Write(acok.Values("FuncID"))
                End If
            Else
                Response.Write("<script language=""JavaScript"">")
                Response.Write("alert(""查無此使用者身分證帳號!!"");")
                Response.Write("</script>")
            End If
            'Response.Write(str1)

            rd1.Close()
            conn.Close()
        End If

    End Sub
End Class
