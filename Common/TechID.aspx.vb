Partial Class TechID
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'Request("RID")
        'Request("CTName")
        'Request("type") 'Addx(產投任課教師) 'Addy(產投助教)
        Me.modifytype.Value = Me.Request("type")
        If Not IsPostBack Then Create()
        Close.Attributes("onclick") = "OpenProMenu(0);"
        Open.Attributes("onclick") = "OpenProMenu(1);"
        Close.Style("CURSOR") = "hand"
        Open.Style("CURSOR") = "hand"
        If State.Value = "0" Then
            Page.RegisterStartupScript("loading", "<script>OpenProMenu(0);</script>")
        Else
            Page.RegisterStartupScript("loading", "<script>OpenProMenu(1);</script>")
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Button2.Visible = True
        Else
            Button2.Visible = False
        End If
    End Sub

    Sub Create()
        Dim sqlTechType As String = ""
        ''啟動年度2013
        'If sm.UserInfo.Years >= 2013 Then
        'End If
        'localhost:12986/Common/TechID.aspx?type=Addx&RID=F3962&TextField=TeacherName1_2&ValueField=NewData11_1&CTName=
        Select Case Me.modifytype.Value 'Addx/Addy
            Case "Addx" '產投任課教師
                sqlTechType = " AND TechType1='Y'" '& vbCrLf
            Case "Addy" '產投助教
                sqlTechType = " AND TechType2='Y'" '& vbCrLf
        End Select

        Dim sql As String
        Dim dt As DataTable
        TeachID.Value = ""
        TeachName.Value = ""

        sql = "" & vbCrLf
        sql &= " SELECT * FROM Teach_TeacherInfo " & vbCrLf
        sql &= " WHERE WorkStatus = '1' " & vbCrLf
        sql += sqlTechType & vbCrLf
        sql &= "  AND RID = '" & Request("RID") & "' " & vbCrLf
        sql &= "  AND KindEngage LIKE '" & KindEngage.SelectedValue & "' " & vbCrLf
        sql &= " ORDER BY KindEngage ,TeachCName ,TeacherID " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub KindEngage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindEngage.SelectedIndexChanged
        Call Create()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim drv As DataRowView = e.Item.DataItem
                Radio1.Value = drv("TechID").ToString
                Checkbox1.Value = drv("TechID").ToString
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Radio1.Visible = False
                    Checkbox1.Visible = True
                    For i As Integer = 0 To Split(Request("CTName"), ",").Length - 1
                        If drv("TechID").ToString = Split(Request("CTName"), ",")(i) Then
                            Checkbox1.Checked = True
                            If TeachID.Value <> "" Then TeachID.Value &= ","
                            TeachID.Value &= TIMS.ClearSQM(drv("TechID"))
                            If TeachName.Value <> "" Then TeachName.Value &= ","
                            TeachName.Value &= TIMS.ClearSQM(drv("TeachCName"))

                            'If TeachID.Value = "" Then
                            '    TeachID.Value = drv("TechID").ToString
                            '    TeachName.Value = drv("TeachCName").ToString
                            'Else
                            '    TeachID.Value += "," & drv("TechID").ToString
                            '    TeachName.Value += "," & drv("TeachCName").ToString
                            'End If
                        End If
                    Next
                Else
                    Radio1.Visible = True
                    Checkbox1.Visible = False
                End If
                Radio1.Attributes("onclick") = "ReturnTechID('" & drv("TechID") & "','" & drv("TeachCName") & "')"
                Checkbox1.Attributes("onclick") = "SelectTechID(this.checked,'" & drv("TechID") & "','" & drv("TeachCName") & "');"
                e.Item.Cells(1).Text = If(drv("KindEngage").ToString = "1", "內聘", "外聘")

            Case ListItemType.Footer
                If DataGrid1.Items.Count = 0 Then
                    DataGrid1.ShowFooter = True
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(New TableCell)
                    e.Item.Cells(0).ColumnSpan = DataGrid1.Columns.Count
                    e.Item.Cells(0).Text = "查無資料!"
                    e.Item.Cells(0).HorizontalAlign = HorizontalAlign.Center
                Else
                    DataGrid1.ShowFooter = False
                End If
        End Select
    End Sub

    '進階搜尋 (查詢)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim sqlTechType As String = ""
        '啟動年度2013
        'If sm.UserInfo.Years >= 2013 Then
        'End If

        TeachCName.Text = TIMS.ClearSQM(TeachCName.Text)
        TeacherID.Text = TIMS.ClearSQM(TeacherID.Text)
        Dim v_KindEngage1 As String = TIMS.GetListValue(KindEngage1)

        Dim sql As String = ""
        Dim dt As DataTable
        'Dim SearchStr As String
        TeachID.Value = ""
        TeachName.Value = ""
        sql = "" & vbCrLf
        sql &= " SELECT * FROM Teach_TeacherInfo " & vbCrLf
        sql &= " WHERE WorkStatus = '1' " & vbCrLf
        'sql += sqlTechType & vbCrLf
        Select Case Me.modifytype.Value
            Case "Addx" '產投任課教師
                sql &= " AND TechType1='Y'" '& vbCrLf
            Case "Addy" '產投助教
                sql &= " AND TechType2='Y'" '& vbCrLf
        End Select
        sql &= " AND RID='" & Request("RID") & "' " & vbCrLf
        sql &= " AND (TeachCName LIKE '%" & Replace(TeachCName.Text, " ", "%") & "%' OR TeachEName LIKE '%" & Replace(TeachCName.Text, " ", "%") & "%') " & vbCrLf
        sql &= " AND TeacherID LIKE '%" & Replace(TeacherID.Text, " ", "%") & "%' " & vbCrLf
        sql &= " AND KindEngage LIKE '" & v_KindEngage1 & "' " & vbCrLf
        sql &= " ORDER BY KindEngage ,TeachCName ,TeacherID " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub
End Class