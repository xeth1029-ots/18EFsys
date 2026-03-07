Partial Class SD_03_006_3in1
    Inherits System.Web.UI.Page

    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在---------------------------End

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

    Sub AddItem(ByVal iNum As Integer)
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Select Case iNum
            Case 1
                sql = "select * from Adp_WorkStation where Station_ID='000' and Station_Showable='1' and Station_Del='0'"
                station.Visible = False
                tai.Visible = False
            Case 2
                station.Items.Clear()
                tai.Items.Clear()
                station.Visible = True
                sql = " select * from Adp_WorkStation where Station_Showable='1' and Station_Del='0' and Station_ID<>'000' and Station_ID like '%0'"
                sql += " and Station_Scheme_ID='" & Left(center.SelectedValue, 1) & "'"
                sql += " and Station_Unit_ID='" & Mid(center.SelectedValue, 2, 2) & "'"
            Case 3
                tai.Items.Clear()
                tai.Visible = True
                sql = " select * from Adp_WorkStation where Station_Showable='1' and Station_Del='0' and Station_ID<>'000' and Station_ID like '%0'"
                sql += " and Station_Scheme_ID='" & Left(center.SelectedValue, 1) & "'"
                sql += " and Station_Unit_ID='" & Mid(center.SelectedValue, 2, 2) & "'"
                sql += " and Station_ID like '" & Mid(station.SelectedValue, 4, 2) & "%'"
        End Select
        dt = DbAccess.GetDataTable(sql, objconn)

        Select Case iNum
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
                    If Mid(drv("TICKET_TYPE"), 1, 1) = 1 Then
                        e.Item.Cells(6).Text = "電腦基本操作"
                    End If
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
        Dim sql As String
        Dim DataStr As String = ""
        Dim dt As DataTable
        'Dim NewSql As String
        If start_date.Text <> "" Then
            DataStr = " and APPLY_DATE>=" & TIMS.to_date(start_date.Text)
        End If
        If end_date.Text <> "" Then
            DataStr = " and APPLY_DATE<=" & TIMS.to_date(end_date.Text)
        End If

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

        'sql = "SELECT * FROM  "
        'sql += "(SELECT * FROM Adp_DGTRNData WHERE IDNO like '%" & IDNO.Text & "%' and TICKET_STATE='1' and CREATE_RGSTN like '" & Reg & "' and TransToTIMS='N'" & DataStr & ") a "
        'sql += "LEFT join Adp_StdData b on a.IDNO=b.IDNO "
        'sql += "LEFT JOIN (SELECT * FROM Adp_ShareSource WHERE Share_Type='301') c ON a.OBJECT_TYPE=c.Share_ID "
        'sql += "LEFT JOIN (SELECT * FROM Adp_WorkStation) d ON a.CREATE_RGSTN=d.Station_Scheme_ID+d.Station_Unit_ID+d.Station_ID "

        sql = "" & vbCrLf
        sql += " SELECT a.IDNO" & vbCrLf
        sql += " ,a.TICKET_NO" & vbCrLf
        sql += " ,a.APPLY_DATE" & vbCrLf
        sql += " ,a.TICKET_TYPE" & vbCrLf
        sql += " ,b.NAME" & vbCrLf
        sql += " ,c.Share_Name" & vbCrLf
        sql += " ,d.Station_Name" & vbCrLf
        sql += " FROM (" & vbCrLf
        sql += "   SELECT IDNO,TICKET_NO,APPLY_DATE,TICKET_TYPE,OBJECT_TYPE,CREATE_RGSTN FROM Adp_DGTRNData " & vbCrLf
        sql += "   WHERE IDNO like '%" & IDNO.Text & "%' and TICKET_STATE='1' " & vbCrLf
        sql += "   and CREATE_RGSTN like '" & Reg & "' and TransToTIMS='N'" & DataStr & "  ) a" & vbCrLf
        sql += " LEFT join Adp_StdData b on a.IDNO=b.IDNO" & vbCrLf
        sql += " LEFT JOIN Adp_ShareSource c ON a.OBJECT_TYPE=c.Share_ID and c.Share_Type='301'" & vbCrLf
        sql += " LEFT JOIN Adp_WorkStation d ON a.CREATE_RGSTN=d.Station_Scheme_ID+d.Station_Unit_ID+d.Station_ID" & vbCrLf
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
        TIMS.Utl_Redirect1(Me, "SD_03_006.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        For Each item As DataGridItem In Datagrid1.Items
            Dim myradio As HtmlInputRadioButton = item.FindControl("Radio1")
            If myradio.Checked = True Then
                TIMS.Utl_Redirect1(Me, "SD_03_006_add.aspx?ID=" & Request("ID") & "&OCID=" & Request("OCID") & "&TICKET_NO=" & myradio.Value)
            End If
        Next
    End Sub
End Class
