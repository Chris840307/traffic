
Partial Class AddData
    Inherits System.Web.UI.Page

    Protected Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click

        '取得 Web.config 檔的資料連接設定
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        '建立 Connection 物件
        Dim ILLEGALITEMLIST As String = ""
        Dim Usertype As String = ""
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        '開啟資料連接
        conn.Open()

        '是否為新增狀態
        If txtNo.Enabled = True Then
            Dim strClose As String = ""
            Dim strDbChk = "select No from BILLEXHORTRECORD where NO='" & txtNo.Text & "' and RecordStateID=0 "
            Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
            Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
            If rdDbChk.HasRows Then
                rdDbChk.Read()
                strClose = Trim(rdDbChk("No"))
            End If
            rdDbChk.Close()

            If strClose <> "" Then
                Response.Write("<script language=""JavaScript"">")
                Response.Write("alert(""該編號已使用過!!"");")
                Response.Write("</script>")
                conn.Close()
                Exit Sub
            End If
        End If

        If cbxIllTrue1.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "1,"
        End If
        If cbxIllTrue2.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "2,"
        End If
        If cbxIllTrue3.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "3,"
        End If
        If cbxIllTrue4.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "4,"
        End If
        If cbxIllTrue5.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "5,"
        End If
        If cbxIllTrue6.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "6,"
        End If
        If cbxIllTrue7.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "7,"
        End If
        If cbxIllTrue8.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "8,"
        End If
        If cbxIllTrue9.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "9,"
        End If
        If cbxIllTrue10.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "10,"
        End If
        If cbxIllTrue11.Checked Then
            ILLEGALITEMLIST = ILLEGALITEMLIST & "11,"
        End If

        If RBUserName.Checked Then
            Usertype = "0"
        Else
            Usertype = "1"
        End If

        Dim cmd As New Data.OracleClient.OracleCommand()


        '先刪除後新增
        Dim strSql = "Delete BILLEXHORTRECORD where NO='" & Trim(txtNo.Text) & "' and RECORDSTATEID=0"

        cmd.CommandText = strSql
        cmd.Connection = conn
        cmd.ExecuteNonQuery()

        strSql = "Insert into BILLEXHORTRECORD(NO,USERNAME,USERID,USERBIRTH,CARNO,USERADDRESS,USERTYPEID,ILLEGALDATETIME,ILLEGALADDRESS,ILLEGALITEMLIST,OTHERILLEGALITEM,BILLFILLERMEMBERID,EXHORTUNITID,RECORDDATE,RECORDMEMBERID,RECORDSTATEID) "
        strSql = strSql & " Values('" & txtNo.Text & "','" & txtUserName.Text & "','" & txtID.Text & "'," & funGetDate(gOutDT(txtBirthYear.Text & txtBirthMonth.Text & txtBirthDay.Text), 0) & ",'" & txtCar_No.Text & "','" & txtAddress.Text & "'," & Usertype & "," & funGetDate(gOutDT(txtIllYear.Text & txtIllMonth.Text & txtIllDay.Text), 1) & ",'" & txtIllAddress.Text & "','" & ILLEGALITEMLIST & "','" & txtOTHERILLEGALITEM.Text & "'," & DDLFillMemID.SelectedValue & ",'" & DDLUnit.SelectedValue & "',sysdate," & Request.Cookies("RewardUser")("MemberID") & ",0)"
        'strSql = strSql & " Values('" & Trim(txtNo.Text) & "','" & txtUserName.Text & "','" & txtID.Text & "'," & funGetDate(gOutDT(txtBirthYear.Text & txtBirthMonth.Text & txtBirthDay.Text), 0) & ",'" & txtCar_No.Text & "','" & txtAddress.Text & "'," & Usertype & "," & funGetDate(gOutDT(txtIllYear.Text & txtIllMonth.Text & txtIllDay.Text), 1) & ",'" & txtIllAddress.Text & "','" & ILLEGALITEMLIST & "','" & txtOTHERILLEGALITEM.Text & "'," & DDLFillMemID.SelectedValue & ",'" & DDLUnit.SelectedValue & "',sysdate,6000,0)"

        cmd.CommandText = strSql
        cmd.Connection = conn
        cmd.ExecuteNonQuery()

        conn.Close()

        Response.Write("<SCRIPT>alert(""儲存完畢"");")
        Response.Write("</script>")
        '修改結束後即關閉，新增結束則繼續畫面
        If txtNo.Enabled = False Then
            Response.Write("<script language=""javascript"">window.close()</script>")
        Else
            Response.Redirect("AddData.aspx?DDLUnit=" & DDLUnit.SelectedValue & "&DDLFillMemID=" & DDLFillMemID.SelectedValue & "&txtMemID=" & txtMemID.Text & "&YN=Y")
        End If
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
        Dim conn As New Data.OracleClient.OracleConnection()
        conn.ConnectionString = setting.ConnectionString
        conn.Open()
        Dim strUNITLEVELID As String = ""
        Dim strDbChk = "select UNITLEVELID from UNITINFO where UNITID='" & Request.Cookies("RewardUser")("UnitID") & "'"
        Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
        Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
        If rdDbChk.HasRows Then
            rdDbChk.Read()
            strUNITLEVELID = Trim(rdDbChk("UNITLEVELID"))
        End If
        rdDbChk.Close()

        If strUNITLEVELID = "1" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO order by showorder"
        ElseIf strUNITLEVELID = "2" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        ElseIf strUNITLEVELID = "3" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UnitID='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        Else
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        End If

        If Not Page.IsPostBack Then

            Dim strCheck
            Dim i As Integer
            Dim strDbChk2 = "select * from BILLEXHORTRECORD where NO='" & Trim(Request("NO")) & "' and RecordStateID=0 "
            Dim CmdDbChk2 As New Data.OracleClient.OracleCommand(strDbChk2, conn)
            Dim rdDbChk2 As Data.OracleClient.OracleDataReader = CmdDbChk2.ExecuteReader()
            txtNo.Enabled = True
            '修改畫面抓出所有的值


            If rdDbChk2.HasRows Then
                rdDbChk2.Read()
                txtNo.Text = Trim(rdDbChk2("NO"))
                txtNo.Enabled = False
                txtUserName.Text = Trim(rdDbChk2("UserName"))
                txtID.Text = Trim(rdDbChk2("UserID"))
                txtBirthYear.Text = Year((rdDbChk2("UserBirth"))) - 1911
                txtBirthMonth.Text = Format(Month(rdDbChk2("UserBirth")), "00")
                txtBirthDay.Text = Format(Day(rdDbChk2("UserBirth")), "00")
                txtCar_No.Text = Trim(rdDbChk2("CarNO"))
                txtAddress.Text = Trim(rdDbChk2("UserAddress"))

                If Trim(rdDbChk2("UserTypeID")) = "0" Then
                    RBUserName.Checked = True
                    RBUserAddress.Checked = True
                Else
                    RBUserName.Checked = False
                    RBUserAddress.Checked = False
                End If

                txtIllYear.Text = Year(rdDbChk2("ILLEGALDATETIME")) - 1911
                txtIllMonth.Text = Format(Month(rdDbChk2("ILLEGALDATETIME")), "00")
                txtIllDay.Text = Format(Day(rdDbChk2("ILLEGALDATETIME")), "00")
                txtIllHour.Text = Format(Hour(rdDbChk2("ILLEGALDATETIME")), "00")
                txtIllMin.Text = Format(Minute(rdDbChk2("ILLEGALDATETIME")), "00")

                txtIllAddress.Text = Trim(rdDbChk2("ILLEGALAddress"))
                txtOTHERILLEGALITEM.Text = Trim(rdDbChk2("OTHERILLEGALITEM") & "")

                DDLUnit.SelectedValue = Trim(rdDbChk2("EXHORTUNITID"))
                DDLUnit.DataBind()
                DDLFillMemID.SelectedValue = Trim(rdDbChk2("BILLFILLERMEMBERID"))

                strCheck = Split(Trim(rdDbChk2("ILLEGALITEMLIST") & ""), ",")
                If Trim(strCheck(0)) <> "" Then
                    For i = 0 To UBound(strCheck)
                        If strCheck(i) = "1" Then
                            cbxIllTrue1.Checked = True
                        End If

                        If strCheck(i) = "2" Then
                            cbxIllTrue2.Checked = True
                        End If

                        If strCheck(i) = "3" Then
                            cbxIllTrue3.Checked = True
                        End If

                        If strCheck(i) = "4" Then
                            cbxIllTrue4.Checked = True
                        End If

                        If strCheck(i) = "5" Then
                            cbxIllTrue5.Checked = True
                        End If

                        If strCheck(i) = "6" Then
                            cbxIllTrue6.Checked = True
                        End If

                        If strCheck(i) = "7" Then
                            cbxIllTrue7.Checked = True
                        End If

                        If strCheck(i) = "8" Then
                            cbxIllTrue8.Checked = True
                        End If

                        If strCheck(i) = "9" Then
                            cbxIllTrue9.Checked = True
                        End If

                        If strCheck(i) = "10" Then
                            cbxIllTrue10.Checked = True
                        End If

                        If strCheck(i) = "11" Then
                            cbxIllTrue11.Checked = True
                        End If

                    Next
                End If

            End If
            rdDbChk2.Close()
        End If
        If Request("YN") = "Y" Then
            DDLUnit.SelectedValue = Request("DDLUnit")
            DDLUnit.DataBind()
            DDLFillMemID.SelectedValue = Request("DDLFillMemID")

            txtMemID.Text = Request("txtMemID")
        End If
    End Sub

    Public Function gOutDT(ByVal iDate)
        Dim tTemp As String
        If iDate IsNot DBNull.Value Then
            tTemp = DateSerial(Left(iDate, Len(iDate) - 4) + 1911, Mid(iDate, Len(iDate) - 3, 2), Right(iDate, 2))
            gOutDT = tTemp
        Else
            gOutDT = ""
        End If
    End Function

    Public Function funGetDate(ByVal strDay, ByVal indx)
        Dim dbDay As String = ""

        If Trim(strDay) IsNot DBNull.Value Then
            If indx = 1 Then
                dbDay = "TO_DATE('" & FormatDateTime(strDay, 2) & " " & txtIllHour.Text & ":" & txtIllMin.Text & ":00','YYYY/MM/DD HH24:MI:SS')"
            Else
                dbDay = "TO_DATE('" & strDay & "','YYYY/MM/DD')"
            End If
            funGetDate = dbDay
        End If
    End Function

    Protected Sub txtNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtNo.TextChanged

    End Sub

    Protected Sub DDLUnit_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLUnit.SelectedIndexChanged
        DDLUnit.SelectedValue = DDLUnit.SelectedValue
    End Sub

    Protected Sub DDLUnit_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLUnit.DataBinding

    End Sub

    Protected Sub DDLUnit_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLUnit.DataBound

    End Sub

    Protected Sub DDLUnit_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles DDLUnit.PreRender

    End Sub

    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete

    End Sub

    Protected Sub txtMemID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMemID.TextChanged

        Try
            Dim setting As ConnectionStringSettings = ConfigurationManager.ConnectionStrings("ConnectionString")
            Dim conn As New Data.OracleClient.OracleConnection()
            conn.ConnectionString = setting.ConnectionString
            conn.Open()
            Dim UnitID As String = ""
            Dim strDbChk = "select UnitID from memberdata where recordstateid=0 and memberid=" & txtMemID.Text

            Dim CmdDbChk As New Data.OracleClient.OracleCommand(strDbChk, conn)
            Dim rdDbChk As Data.OracleClient.OracleDataReader = CmdDbChk.ExecuteReader()
            If rdDbChk.HasRows Then
                rdDbChk.Read()
                UnitID = Trim(rdDbChk("UnitID"))

                DDLUnit.SelectedValue = UnitID
                DDLUnit.DataBind()
                DDLFillMemID.DataBind()
                DDLFillMemID.SelectedValue = txtMemID.Text
                DDLFillMemID.DataBind()
                txtMemID.Focus()
            Else
                Response.Write("<SCRIPT>alert(""查無資料"");")
                Response.Write("</script>")
            End If
            rdDbChk.Close()


            conn.Close()
        Catch
            Response.Write("<SCRIPT>alert(""輸入錯誤"");")
            Response.Write("</script>")
        End Try
    End Sub

    Protected Sub txtMemID_TextChanged1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMemID.TextChanged
    End Sub
End Class
