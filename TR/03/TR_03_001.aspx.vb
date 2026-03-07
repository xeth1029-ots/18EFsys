Partial Class TR_03_001
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
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            CreateItem()
            DataGridTable.Visible = False
        End If

        If SearchState.Value = "1" Then
            SearchTable.Style("display") = "inline"
            StateButton.Text = "關閉查詢條件"
        Else
            SearchTable.Style("display") = "none"
            StateButton.Text = "開啟查詢條件"
        End If
        StateButton.Attributes("onclick") = "HidTable();return false;"
        Button3.Attributes("onclick") = "ReportPrint();return false;"
    End Sub

    Sub CreateItem()
        KEID = TIMS.Get_KeyControl(KEID, "Key_Emp", "KENAME", "KEID", objconn)
        SCTID = TIMS.Get_CityName(SCTID, TIMS.dtNothing)
        SCTID.Items.Insert(0, New ListItem("全選", ""))
        SCTID.Attributes("onclick") = "SelectAll();"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        'Dim NewSql As String
        DataGrid1.CurrentPageIndex = 0
        BDID.Value = ""

        Dim KEIDStr As String = ""
        If KEID.SelectedIndex <> -1 AndAlso KEID.SelectedValue <> "" Then
            KEIDStr = KEID.SelectedValue
        End If

        Dim SCTIDStr As String = ""
        For i As Integer = 1 To SCTID.Items.Count - 1
            If SCTID.Items(i).Selected = True AndAlso SCTID.Items(i).Value <> "" Then
                If SCTIDStr <> "" Then SCTIDStr &= ","
                SCTIDStr &= SCTID.Items(i).Value
            End If
        Next

        sql = ""
        sql &= " SELECT * FROM Bus_BasicData"
        sql &= " WHERE 1=1"
        sql &= " and Uname Like '%" & Uname.Text & "%'"
        If KEIDStr <> "" Then
            sql &= " and KEID='" & KEID.SelectedValue & "'"
        End If
        If SCTIDStr <> "" Then
            sql &= " and Zip IN (SELECT ZipCode FROM ID_ZIP WHERE CTID IN (" & SCTIDStr & "))"
        End If
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "BDID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Checkbox1.Attributes("onclick") = "SetBDID(this.checked,'" & drv("BDID").ToString & "')"

                Checkbox1.Value = drv("BDID").ToString
                Dim MyArray As Array = Split(BDID.Value, ",")
                For i As Integer = 0 To MyArray.Length - 1
                    If drv("BDID").ToString = MyArray(i).ToString Then
                        Checkbox1.Checked = True
                    End If
                Next

                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e)
        End Select
    End Sub

End Class
