
Partial Class _Default
    Inherits System.Web.UI.Page

    Public Overloads Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)

        'Confirms that an HtmlForm control is rendered for the specified ASP.NET server control at run time.

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        GirdViewBind()

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
        '權限設定可以看到的單位()
        If strUNITLEVELID = "1" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO order by showorder"
        ElseIf strUNITLEVELID = "2" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        ElseIf strUNITLEVELID = "3" Then
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UnitID='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        Else
            DSUnit.SelectCommand = "SELECT UNITID,UNITNAME,UNITLEVELID FROM UNITINFO where UNITTYPEID='" & Request.Cookies("RewardUser")("UnitID") & "' or unitid='" & Request.Cookies("RewardUser")("UnitID") & "' order by showorder"
        End If
        conn.Close()
    End Sub

    Protected Sub GridView1_DataBinding(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBinding
        Label8.Text = GridView1.PageIndex + 1
        Label9.Text = GridView1.PageCount
    End Sub

    Protected Sub GridView1_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.DataBound

    End Sub

    Protected Sub GridView1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.Load

    End Sub

    Protected Sub btnExcel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnExcel.Click
        '轉出為Excel
        GridView1.AllowPaging = False
        Response.Clear()
        Response.AddHeader("content-disposition", "attachment;filename=" & Server.UrlPathEncode(Format(Now(), "yyyyMMdd") & "_交通違規勸導單.xls"))
        Response.Charset = "big5"
        Response.ContentType = "application/vnd.xls"
        GridView1.Columns.Item(6).Visible = False

        Dim stringWrite As System.IO.StringWriter = New System.IO.StringWriter
        Dim htmlWrite As System.Web.UI.HtmlTextWriter = New HtmlTextWriter(stringWrite)
        GridView1.DataBind()
        GridView1.RenderControl(htmlWrite)

        Response.Write(stringWrite.ToString)
        Response.End()
        GridView1.Columns.Item(6).Visible = True
        GridView1.AllowPaging = True
    End Sub

    Protected Sub btnQry_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnQry.Click
        Dim strSql As String = ""

        If Trim(txtNo.Text) <> "" Then
            strSql = strSql & " and No like '%" & txtNo.Text & "%'"
        End If

        If Trim(txtID.Text) <> "" Then
            strSql = strSql & " and USERID like '%" & txtID.Text & "%'"
        End If

        If Trim(txtCarNo.Text) <> "" Then
            strSql = strSql & " and CarNo like '%" & txtCarNo.Text & "%'"
        End If

        If DDLUnit.SelectedValue <> "所有單位" Then
            strSql = strSql & " and EXHORTUNITID = '" & DDLUnit.SelectedValue & "'"
        End If

        If DDLFillMemID.SelectedValue <> "所有人員" Then
            strSql = strSql & " and BILLFILLERMEMBERID = '" & DDLFillMemID.SelectedValue & "'"
        End If

        If Trim(txtIllDate.Text) <> "" And Trim(txtIllDate2.Text) <> "" Then
            strSql = strSql & " and IllegalDateTime between to_date('" & txtIllDate.Text & " 00:00:00','yyyy/MM/dd HH24:MI:SS') and to_date('" & txtIllDate2.Text & " 23:59:59','yyyy/MM/dd HH24:MI:SS')"
        End If

        SQLDSOrcl.SelectCommand = "SELECT * FROM BILLEXHORTRECORD where RECORDSTATEID =0 " & strSql & " order by no"
        GridView1.DataBind()
        '重整筆數       
        GirdViewBind()

    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click
        '彈出新增畫面
        Response.Write("<script language=""javascript"">")
        Response.Write("window.open(""AddData.aspx"",""tmpWindow"",""width=730,height=655,left=150,top=0,resizable=yes,scrollbars=yes"");")
        Response.Write("</script>")

    End Sub

    Protected Sub btnEdit_Click(ByVal sender As Object, ByVal e As System.EventArgs)


    End Sub
    Protected Sub GirdViewBind()
        Dim dv As System.Data.DataView = CType(SQLDSOrcl.Select(DataSourceSelectArguments.Empty), System.Data.DataView)
        Dim pagerRow As GridViewRow = GridView1.BottomPagerRow

        Label7.Text = dv.Count

        Label8.Text = GridView1.PageIndex + 1
        Label9.Text = GridView1.PageCount

    End Sub

    Protected Sub DropDownList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DropDownList1.SelectedIndexChanged
        '設定每頁幾筆
        GridView1.PageSize = DropDownList1.SelectedValue
    End Sub

    Protected Sub SQLDSOrcl_Selecting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceSelectingEventArgs) Handles SQLDSOrcl.Selecting

    End Sub

    Protected Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        txtNo.Text = ""
        txtIllDate.Text = ""
        txtIllDate2.Text = ""
        txtID.Text = ""
        txtCarNo.Text = ""
        DDLUnit.SelectedIndex = 0
        DDLFillMemID.SelectedIndex = 0
    End Sub

    Protected Sub First_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub Prev_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub Next_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub Last_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub GridView1_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs) Handles GridView1.RowCommand
        '抓該欄位值
        '彈出修改畫面
        If e.CommandName = "EditData" Then '當有CommandName屬性為Detail的按鈕被按下時
            Dim id As String = e.CommandArgument
            Response.Write("<script language=""javascript"">")
            Response.Write("window.open(""AddData.aspx?NO=" & Server.UrlPathEncode(id) & """,""tmpWindow"",""width=730,height=655,left=150,top=0,resizable=yes,scrollbars=yes"");")
            Response.Write("</script>")
            
        End If
        If e.CommandName = "Delete" Then '當有CommandName屬性為Detail的按鈕被按下時
            Dim id As String = e.CommandArgument
            Response.Write("<script language=""javascript"">")
            Response.Write("window.open(""AddData.aspx?NO=" & Server.UrlPathEncode(id) & """,""tmpWindow"",""width=730,height=655,left=150,top=0,resizable=yes,scrollbars=yes"");")
            Response.Write("</script>")

        End If
    End Sub

    Protected Sub GridView1_RowCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowCreated
        '隱藏紀錄人列
        If e.Row.RowType = DataControlRowType.DataRow Or e.Row.RowType = DataControlRowType.Header Then
            e.Row.Cells(7).Visible = False
        End If
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound
        '只針對資料列做處理

        If e.Row.RowType = DataControlRowType.DataRow Then
            '滑鼠移至資料列上的顏色
            e.Row.Attributes.Add("onmouseover", "this.style.backgroundColor='SkyBlue'")
            '滑鼠離開資料列上的顏色
            e.Row.Attributes.Add("onmouseout", "this.style.backgroundColor='#FFFFC0'")

            '非記錄人就不顯示修改、刪除按鈕
            If Trim(e.Row.Cells(7).Text) <> Request.Cookies("RewardUser")("MemberID") Then
                e.Row.Cells(6).Visible = False
            End If

        End If

    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub btnFirst_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFirst.Click
        '第一筆
        GridView1.PageIndex = 0
    End Sub

    Protected Sub btnPre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPre.Click
        '上一筆
        If GridView1.PageIndex <> 0 Then
            GridView1.PageIndex = GridView1.PageIndex - 1
        End If
    End Sub

    Protected Sub btnNext_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNext.Click
        '下一筆
        If GridView1.PageIndex <> GridView1.PageCount - 1 Then
            GridView1.PageIndex = GridView1.PageIndex + 1
        End If
    End Sub

    Protected Sub btnLast_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLast.Click
        '最後一筆
        GridView1.PageIndex = GridView1.PageCount - 1
    End Sub

    Protected Sub btnDel_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub txtNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub
End Class
