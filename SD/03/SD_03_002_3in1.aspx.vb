Partial Class SD_03_002_3in1
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            AddItem(1)
            DataGridTable.Visible = False
            If Not Session("_SearchStr") Is Nothing Then
                Me.ViewState("_SearchStr") = Session("_SearchStr")
                Session("_SearchStr") = Nothing
            End If
        End If

        DataGridPage1.MyDataGrid = Datagrid1
        msg.Text = ""
    End Sub

    Sub AddItem(ByVal num As Integer)
        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow

        Select Case num
            Case 1
                sql = "select * from Adp_WorkStation where Station_ID=000 and Station_Showable=1 and Station_Del=0"
                station.Visible = False
                tai.Visible = False
            Case 2
                station.Items.Clear()
                tai.Items.Clear()
                station.Visible = True
                sql = "select * from Adp_WorkStation where Station_Scheme_ID='" & Left(center.SelectedValue, 1) & "' and Station_Unit_ID='" & Mid(center.SelectedValue, 2, 2) & "' and Station_Showable=1 and Station_Del=0 and Station_ID<>000 and Station_ID like '%0'"
            Case 3
                tai.Items.Clear()
                tai.Visible = True
                sql = "select * from Adp_WorkStation where Station_Scheme_ID='" & Left(station.SelectedValue, 1) & "' and Station_Unit_ID='" & Mid(station.SelectedValue, 2, 2) & "' and Station_Showable=1 and Station_Del=0 and Station_ID<>000 and Station_ID not like '%0' and Station_ID like '" & Mid(station.SelectedValue, 4, 2) & "%'"
        End Select
        dt = DbAccess.GetDataTable(sql, objconn)

        Select Case num
            Case 1
                center.Items.Add(New ListItem("全部範圍", "%"))
                For Each dr In dt.Rows
                    center.Items.Add(New ListItem(dr("Station_Name"), dr("Station_Scheme_ID") & dr("Station_Unit_ID") & dr("Station_ID")))
                Next
            Case 2
                station.Items.Add(New ListItem("全部範圍", center.SelectedValue))
                For Each dr In dt.Rows
                    station.Items.Add(New ListItem(dr("Station_Name"), dr("Station_Scheme_ID") & dr("Station_Unit_ID") & dr("Station_ID")))
                Next
            Case 3
                tai.Items.Add(New ListItem("全部範圍", station.SelectedValue))
                For Each dr In dt.Rows
                    tai.Items.Add(New ListItem(dr("Station_Name"), dr("Station_Scheme_ID") & dr("Station_Unit_ID") & dr("Station_ID")))
                Next
        End Select
    End Sub

    Private Sub center_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles center.SelectedIndexChanged
        If center.SelectedIndex = 0 Then
            station.Visible = False
            tai.Visible = False
        Else
            AddItem(2)
        End If
    End Sub

    Private Sub station_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles station.SelectedIndexChanged
        If station.SelectedIndex = 0 Then
            tai.Visible = False
        Else
            AddItem(3)
        End If
    End Sub

    Private Sub Datagrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Radio1.Attributes("onclick") = "checkRadio('DataGrid1','" & e.Item.ItemIndex + 1 & "');"
                Radio1.Value = drv("TICKET_NO").ToString
                If Not IsDBNull(drv("TICKET_TYPE")) Then
                    If Mid(drv("TICKET_TYPE"), 1, 1) = 1 Then e.Item.Cells(6).Text = "電腦基本操作"
                    If Mid(drv("TICKET_TYPE"), 2, 1) = 1 Then
                        If e.Item.Cells(6).Text = "" Then
                            e.Item.Cells(6).Text = "文書處裡與問題練習"
                        Else
                            e.Item.Cells(6).Text = e.Item.Cells(6).Text + "<br>文書處裡與問題練習"
                        End If
                    End If
                    If Mid(drv("TICKET_TYPE"), 3, 1) = 1 Then
                        If e.Item.Cells(6).Text = "" Then
                            e.Item.Cells(6).Text = "網際網路應用與問題練習"
                        Else
                            e.Item.Cells(6).Text = e.Item.Cells(6).Text + "<br>網際網路應用與問題練習"
                        End If
                    End If
                End If
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
            Me.TxtPageSize.Text = 10
        End If
        If Me.TxtPageSize.Text <> Me.Datagrid1.PageSize Then Me.Datagrid1.PageSize = Me.TxtPageSize.Text

        Dim sql As String
        Dim DataStr As String = ""
        Dim dt As DataTable
        'Dim NewSql As String
        If start_date.Text <> "" Then DataStr = " AND APPLY_DATE >= " & TIMS.to_date(start_date.Text)
        If end_date.Text <> "" Then DataStr = " AND APPLY_DATE <= " & TIMS.to_date(end_date.Text)

        Dim Reg As String
        If center.SelectedValue = "%" Then
            Reg = "%"
        ElseIf center.SelectedValue <> "%" And center.SelectedValue = station.SelectedValue Then
            Reg = Mid(center.SelectedValue, 1, 3) & "%"
        ElseIf center.SelectedValue <> station.SelectedValue And tai.SelectedValue = station.SelectedValue Then
            Reg = Mid(station.SelectedValue, 1, 5) & "%"
        Else
            Reg = tai.SelectedValue
        End If

        sql = "" & vbCrLf
        sql += " SELECT a.IDNO " & vbCrLf
        sql += "        ,a.TICKET_NO " & vbCrLf
        sql += "        ,a.APPLY_DATE " & vbCrLf
        sql += "        ,b.Name " & vbCrLf
        sql += "        ,c.Share_Name " & vbCrLf
        sql += "        ,a.TICKET_TYPE " & vbCrLf
        sql += "        ,d.Station_Name " & vbCrLf
        sql += " FROM (" & vbCrLf
        sql += "    SELECT * FROM Adp_DGTRNData " & vbCrLf
        sql += "    WHERE 1=1 " & vbCrLf
        sql += "       AND UPPER(IDNO) LIKE '%" & IDNO.Text & "%' " & vbCrLf
        sql += "       AND TICKET_STATE = '1' " & vbCrLf
        sql += "       AND CREATE_RGSTN LIKE '" & Reg & "' " & vbCrLf
        sql += "       AND TransToTIMS = 'N' " & vbCrLf
        sql += DataStr
        sql += " ) a " & vbCrLf
        sql += " LEFT JOIN Adp_StdData b ON a.IDNO = b.IDNO " & vbCrLf
        sql += " LEFT JOIN Adp_ShareSource c ON a.OBJECT_TYPE = c.Share_ID AND c.Share_Type = '301' " & vbCrLf
        sql += " LEFT JOIN Adp_WorkStation d ON a.CREATE_RGSTN = d.Station_Scheme_ID + d.Station_Unit_ID + d.Station_ID " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!"
        DataGridTable.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable.Visible = True
            Datagrid1.DataSource = dt
            Datagrid1.DataBind()
            '分頁用--------------------------------------------Start
            DataGridPage1.MyDataTable = dt
            DataGridPage1.FirstTime()
            '分頁用--------------------------------------------End
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        TIMS.Utl_Redirect1(Me, "SD_03_002.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        For Each item As DataGridItem In Datagrid1.Items
            Dim myradio As HtmlInputRadioButton = item.FindControl("Radio1")
            If myradio.Checked = True Then TIMS.Utl_Redirect1(Me, "SD_03_002_add.aspx?ID=" & Request("ID") & "&OCID=" & Request("OCID") & "&TICKET_NO=" & myradio.Value)
        Next
    End Sub
End Class