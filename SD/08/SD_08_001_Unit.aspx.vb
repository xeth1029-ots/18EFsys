Partial Class SD_08_001_Unit
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        '檢查Session是否存在 End

        If Not IsPostBack Then
            'Create()
        End If
        Button1.Style("display") = "none"
    End Sub

#Region "NO USE"
    'Function Create()
    '    Dim sql As String
    '    Dim dt As DataTable
    '    Dim dr As DataRow

    '    sql = "SELECT LUID,LUID+'：'+LUName as LUName FROM Key_LapmUnit"
    '    dt = DbAccess.GetDataTable(sql, objconn)
    '    DataGrid1.DataSource = dt
    '    DataGrid1.DataBind()
    'End Function
#End Region

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Attributes("onclick") = "ReturnMyValue('" & drv("LUID") & "','" & drv("LUName") & "');"
                e.Item.Style("Cursor") = "hand"
                e.Item.Attributes("onmouseover") = "this.style.backgroundColor='#EFEFEF'"
                e.Item.Attributes("onmouseout") = "this.style.backgroundColor=''"
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

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        sql = "SELECT LUID,LUID+'：'+LUName as LUName  FROM Key_LapmUnit WHERE LUID like '%" & Replace(NowValue.Value, " ", "%") & "%' or LUName like '%" & Replace(NowValue.Value, " ", "%") & "%' ORDER BY LUID"
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub
End Class
