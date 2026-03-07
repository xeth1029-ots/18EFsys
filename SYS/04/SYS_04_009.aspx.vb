Partial Class SYS_04_009
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Create()
        End If
    End Sub

    Sub Create()
        Dim sql As String
        Dim dt As DataTable

        sql = ""
        sql += " SELECT a.TPlanID,a.PlanName,b.CancelID "
        sql += " FROM Key_Plan a "
        sql += " LEFT JOIN Plan_BudgetCan b ON a.TPlanID=b.TPlanID "
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "TPlanID"
            DataGrid1.DataBind()
        Else
            Common.MessageBox(Me, "資料異常，請連絡系統管理者!!")
            'Exit Sub
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim CancelID1 As HtmlInputRadioButton = e.Item.FindControl("CancelID1")
                Dim CancelID2 As HtmlInputRadioButton = e.Item.FindControl("CancelID2")
                Dim CancelID3 As HtmlInputRadioButton = e.Item.FindControl("CancelID3")
                Dim CancelID4 As HtmlInputRadioButton = e.Item.FindControl("CancelID4")
                Dim CancelID5 As HtmlInputRadioButton = e.Item.FindControl("CancelID5")
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.CssClass = ""
                If Not IsDBNull(drv("CancelID")) Then
                    Select Case Val(drv("CancelID"))
                        Case 1
                            CancelID1.Checked = True
                        Case 2
                            CancelID2.Checked = True
                        Case 3
                            CancelID3.Checked = True
                        Case 4
                            CancelID4.Checked = True
                        Case 5
                            CancelID5.Checked = True
                    End Select
                End If
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow
        'Dim conn As SqlConnection
        '2006/03/28 add conn by matt
        'conn = DbAccess.GetConnection

        sql = "SELECT * FROM Plan_BudgetCan"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each item As DataGridItem In DataGrid1.Items
            Dim CancelID1 As HtmlInputRadioButton = item.FindControl("CancelID1")
            Dim CancelID2 As HtmlInputRadioButton = item.FindControl("CancelID2")
            Dim CancelID3 As HtmlInputRadioButton = item.FindControl("CancelID3")
            Dim CancelID4 As HtmlInputRadioButton = item.FindControl("CancelID4")
            Dim CancelID5 As HtmlInputRadioButton = item.FindControl("CancelID5")
            Dim CancelID As Integer = 0

            CancelID = 0
            Select Case True
                Case CancelID1.Checked
                    CancelID = CancelID1.Value
                Case CancelID2.Checked
                    CancelID = CancelID2.Value
                Case CancelID3.Checked
                    CancelID = CancelID3.Value
                Case CancelID4.Checked
                    CancelID = CancelID4.Value
                Case CancelID5.Checked
                    CancelID = CancelID5.Value
            End Select
            'If CancelID1.Checked = True Then
            '    CancelID = CancelID1.Value
            'End If
            'If CancelID2.Checked = True Then
            '    CancelID = CancelID2.Value
            'End If
            'If CancelID3.Checked = True Then
            '    CancelID = CancelID3.Value
            'End If
            'If CancelID4.Checked = True Then
            '    CancelID = CancelID4.Value
            'End If
            'If CancelID5.Checked = True Then
            '    CancelID = CancelID5.Value
            'End If

            If CancelID <> 0 Then
                If dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)

                    dr("TPlanID") = DataGrid1.DataKeys(item.ItemIndex)
                Else
                    dr = dt.Select("TPlanID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0)
                End If
                dr("CancelID") = CancelID
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
            End If
        Next

        DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功")
    End Sub
End Class
