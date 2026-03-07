Partial Class SYS_03_021
    Inherits AuthBasePage

    Dim trID As Integer = 0
    Dim sql As String = ""
    Dim intCntSch As Int16 = 0 '判斷是否為新查詢(0=>是;1=>否)
    Dim chkEnableAll As CheckBox

    Dim dtAPF As DataTable

    'arrFun = FunSort.Split(",")
    Const cst_tab2 As String = "&nbsp;&nbsp;&nbsp;　"
    Dim arrFun As String()  '= {"TC", "SD", "CP", "TR", "CM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"} 'fun排列順序
    'Dim FunSort As String = System.Configuration.ConfigurationSettings.AppSettings("FunSort")

#Region "FUNC1"

    '查詢
    Private Sub search()
        sql = " SELECT * FROM AUTH_PLANFUNCTION WHERE TPLANID = @TPLANID "
        Call TIMS.OpenDbConn(objconn)
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = ddlTPlan.SelectedValue
            dtAPF = New DataTable
            dtAPF.Load(.ExecuteReader())
        End With

        Dim dt_kind As DataTable = Nothing
        'Dim dr As DataRow = Nothing

        '組清單用DataTable
        Dim dt_fun As New DataTable
        With dt_fun
            .Columns.Add(New DataColumn("funid"))
            .Columns.Add(New DataColumn("name"))
            .Columns.Add(New DataColumn("spage"))
            .Columns.Add(New DataColumn("kind"))
            .Columns.Add(New DataColumn("levels"))
            .Columns.Add(New DataColumn("parent"))
            .Columns.Add(New DataColumn("sort"))
            .Columns.Add(New DataColumn("memo"))
            .Columns.Add(New DataColumn("subs"))
        End With

        Try
            Dim sdaSelect As New SqlDataAdapter
            Dim ds As New DataSet

            'conn.Open()
            For i As Integer = 0 To arrFun.Length - 1
                '取得功能清單
                sql = ""
                sql &= " SELECT a.funid ,a.name ,a.spage ,a.kind ,a.levels" & vbCrLf
                sql &= " ,(CASE a.Levels WHEN '0' THEN CONVERT(VARCHAR(5),a.FunID) ELSE a.Parent END) AS Parent" & vbCrLf
                sql &= " ,ISNULL(p.sort,a.sort) psort ,a.sort ,a.memo" & vbCrLf
                sql &= " ,(CASE a.Levels WHEN '0' THEN (SELECT COUNT(x.FUNID) FROM ID_FUNCTION x WHERE x.parent = a.funid) ELSE 0 END) AS subs" & vbCrLf
                sql &= " FROM ID_FUNCTION a" & vbCrLf
                sql &= " LEFT JOIN ID_FUNCTION p ON p.funid = a.parent" & vbCrLf
                sql &= " WHERE a.VALID = 'Y'" & vbCrLf
                sql &= " AND a.FState NOT IN ('D')" & vbCrLf
                sql &= " AND a.kind = @kind" & vbCrLf
                sql &= " ORDER BY psort,a.kind,a.levels,a.sort" & vbCrLf

                Dim tmpkind As String = If(ddlFun.SelectedIndex <> 0, ddlFun.SelectedValue, arrFun(i))
                With sdaSelect
                    .SelectCommand = New SqlCommand(sql, objconn)
                    .SelectCommand.Parameters.Clear()
                    .SelectCommand.Parameters.Add("kind", SqlDbType.VarChar).Value = tmpkind
                    .Fill(ds, "select_fun" & i)
                End With

                For j As Integer = 0 To ds.Tables("select_fun" & i).Rows.Count - 1
                    Dim dr As DataRow = dt_fun.NewRow
                    dt_fun.Rows.Add(dr)
                    dr("funid") = ds.Tables("select_fun" & i).Rows(j)("funid")
                    dr("name") = ds.Tables("select_fun" & i).Rows(j)("name")
                    dr("spage") = ds.Tables("select_fun" & i).Rows(j)("spage")
                    dr("kind") = ds.Tables("select_fun" & i).Rows(j)("kind")
                    dr("levels") = ds.Tables("select_fun" & i).Rows(j)("levels")
                    dr("parent") = ds.Tables("select_fun" & i).Rows(j)("parent")
                    dr("sort") = ds.Tables("select_fun" & i).Rows(j)("sort")
                    dr("memo") = ds.Tables("select_fun" & i).Rows(j)("memo")
                    dr("subs") = ds.Tables("select_fun" & i).Rows(j)("subs")
                Next

                If ddlFun.SelectedIndex <> 0 Then
                    Exit For
                End If
            Next

            '取得功能種類
            sql = " SELECT DISTINCT kind FROM ID_Function ORDER BY kind "
            With sdaSelect
                .SelectCommand = New SqlCommand(sql, objconn)
                .Fill(ds, "select_kind")
            End With

            If Not ds.Tables("select_kind") Is Nothing Then
                If ds.Tables("select_kind").Rows.Count > 0 Then dt_kind = ds.Tables("select_kind")
            End If

            For i As Integer = 0 To dt_kind.Rows.Count - 1
                Dim dr As DataRow = dt_kind.Rows(i)
                Me.ViewState(dr("kind").ToString) = dt_fun.Select("kind='" & dr("kind").ToString & "'").Length
            Next

            intCntSch = 0

            DataGrid1.DataSource = dt_fun
            DataGrid1.DataKeyField = "FunID"
            DataGrid1.DataBind()

            'TIMS.CloseDbConn(objConn)
            If Not sdaSelect Is Nothing Then sdaSelect.Dispose()
            If Not ds Is Nothing Then ds.Dispose()
        Catch ex As Exception
            'conn.Close()
            Common.MessageBox(Me, ex.ToString)
        End Try
    End Sub

    '清除項目
    Private Sub Clear_Items()
        Me.ViewState("mainmenu") = Nothing

        '鎖定所有勾選項目
        For Each itm As DataGridItem In DataGrid1.Items
            Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
            chkEnable.Enabled = False
        Next

        tr_Fun.Visible = False
        tr_Btn.Visible = False
    End Sub

    '判斷計畫中功能類別是否勾選(true=>勾選,false=>否)
    Function chkData(ByVal dtAPF As DataTable, ByVal intFunID As Integer) As Boolean
        Dim Rst As Boolean = False
        Dim ff As String = ""
        ff = "FUNID=" & intFunID
        If dtAPF.Select(ff).Length > 0 Then Rst = True
        Return Rst
    End Function

#End Region

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        arrFun = TIMS.c_FUNSORT.Split(",") 'arrFun = FunSort.Split(",")

        If Not IsPostBack Then
            tr_Fun.Visible = False
            tr_Btn.Visible = False
            ddlFun = TIMS.Get_ddlFunction(ddlFun, 2)
            TIMS.Get_TPlan(ddlTPlan, , 2)
            btn_Save.Attributes.Add("style", "cursor@hand")
        End If
    End Sub

    Sub SAVEDATA1()
        'Dim sdaDelete As New SqlDataAdapter
        'Dim sdaInsert As New SqlDataAdapter
        'Dim trans As SqlTransaction

        Dim v_ddlTPlan As String = TIMS.GetListValue(ddlTPlan) '.SelectedValue
        TIMS.OpenDbConn(objconn)
        Dim trans As SqlTransaction = objconn.BeginTransaction()

        Dim sql_DEL As String = " DELETE Auth_PlanFunction WHERE TPlanID = @TPlanID"
        Dim sql_DEL2 As String = " DELETE Auth_PlanFunction WHERE TPlanID = @TPlanID AND FunID = @FunID"
        Dim sql_W1_DEL As String = " SELECT 1 FROM Auth_PlanFunction WHERE TPlanID = @TPlanID AND FunID = @FunID"
        Dim sql_W1_ADD As String = " SELECT 1 FROM Auth_PlanFunction WHERE TPlanID = @TPlanID AND FunID = @FunID"
        Dim sql_ADD As String = " INSERT INTO Auth_PlanFunction(TPlanID,FunID,ModifyAcct,ModifyDate) VALUES(@TPlanID,@FunID,@ModifyAcct,GETDATE()) "

        Try
            '刪除舊有功能類別(全刪除 or 部份刪除)
            If ddlFun.SelectedIndex = 0 Then
                Using cmd_DEL As New SqlCommand(sql_DEL, objconn, trans)
                    With cmd_DEL
                        .Parameters.Clear()
                        .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = v_ddlTPlan 'ddlTPlan.SelectedValue
                        .ExecuteNonQuery()
                    End With
                End Using

            Else
                For Each itm As DataGridItem In DataGrid1.Items
                    Dim txtFunID As TextBox = itm.FindControl("txtFunID")
                    Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                    If Not chkEnable.Checked Then
                        Dim fg_can_del As Boolean = False
                        Using cmd_W1 As New SqlCommand(sql_W1_DEL, objconn, trans)
                            Dim dtW As New DataTable
                            With cmd_W1
                                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = v_ddlTPlan 'ddlTPlan.SelectedValue
                                .Parameters.Add("FunID", SqlDbType.Int).Value = Convert.ToInt32(txtFunID.Text)
                                dtW.Load(.ExecuteReader())
                            End With
                            fg_can_del = TIMS.dtHaveDATA(dtW)
                        End Using
                        If (fg_can_del) Then
                            Using cmd_DEL As New SqlCommand(sql_DEL2, objconn, trans)
                                With cmd_DEL
                                    .Parameters.Clear()
                                    .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = v_ddlTPlan 'ddlTPlan.SelectedValue
                                    .Parameters.Add("FunID", SqlDbType.Int).Value = Convert.ToInt32(txtFunID.Text)
                                    .ExecuteNonQuery()
                                End With
                            End Using
                        End If
                    End If
                Next

            End If

            '新增勾選之功能類別
            For Each itm As DataGridItem In DataGrid1.Items
                Dim txtFunID As TextBox = itm.FindControl("txtFunID")
                Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                If chkEnable.Checked Then

                    Dim fg_can_Insert As Boolean = False
                    Using cmd_W1 As New SqlCommand(sql_W1_ADD, objconn, trans)
                        Dim dtW As New DataTable
                        With cmd_W1
                            .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = v_ddlTPlan 'ddlTPlan.SelectedValue
                            .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid1.DataKeys.Item(itm.ItemIndex))
                            dtW.Load(.ExecuteReader())
                        End With
                        fg_can_Insert = TIMS.dtNODATA(dtW)
                    End Using
                    If fg_can_Insert Then
                        Using cmd_ADD As New SqlCommand(sql_ADD, objconn, trans)
                            With cmd_ADD
                                .Parameters.Clear()
                                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = v_ddlTPlan 'ddlTPlan.SelectedValue
                                .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid1.DataKeys.Item(itm.ItemIndex))
                                '.Parameters.Add("FunID", SqlDbType.Int).Value = Convert.ToInt32(TIMS.VAL1(txtFunID.Text))
                                .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID 'Convert.ToString(sm.UserInfo.UserID)
                                .ExecuteNonQuery()
                            End With
                        End Using
                    End If

                End If
            Next

            trans.Commit()
            'TIMS.CloseDbConn(objconn)
            Common.MessageBox(Me, "儲存成功!")

            'If Not sdaDelete Is Nothing Then sdaDelete.Dispose()
            'If Not sdaInsert Is Nothing Then sdaInsert.Dispose()
        Catch ex As Exception
            If trans IsNot Nothing Then trans.Rollback() 'conn.Close()
            Common.MessageBox(Me, ex.Message)
            TIMS.WriteTraceLog(ex.Message, ex)
        End Try
    End Sub

    Private Sub btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click, btn_Save2.Click
        Call SAVEDATA1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                chkEnableAll = e.Item.FindControl("chk_EnableAll") '全選方塊

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim txtFunID As TextBox = e.Item.FindControl("txtFunID") '記錄FunID
                Dim labMainMenu As Label = e.Item.FindControl("lab_MainMenu") '程式類型
                Dim labFunName As Label = e.Item.FindControl("lab_FunName") '選單名稱
                Dim chkEnable As CheckBox = e.Item.FindControl("chk_Enable") '選取方塊

                txtFunID.Text = Convert.ToString(drv("funid"))

                labMainMenu.Text = TIMS.Get_MainMenuName(Convert.ToString(drv("Kind")))

                If Convert.ToString(drv("levels")) = "1" Then
                    labFunName.Text = cst_tab2 & Convert.ToString(drv("Name"))
                Else
                    labFunName.Text = Convert.ToString(drv("Name"))
                End If

                'snoopy特別權限
                If Convert.ToString(sm.UserInfo.UserID) = str_superuser1 Then
                    '連線系統本身功能
                    Dim sPage As String = Convert.ToString(drv("Spage"))
                    If sPage <> "" Then
                        Dim strUrl As String = ""
                        '含有http字眼另開視窗
                        Const strUrl_f1 As String = "<a href=""../../{0}?ID={1}"" target=""_self"">({0})</a>"
                        Const strUrl_f2 As String = "<a href=""{0}"" target=""_blank"">({1})</a>"
                        Dim flag_new_windows As Boolean = TIMS.CHK_FUNCSPAGE2(sm, sPage)
                        If flag_new_windows Then
                            Dim sPage1 As String = sPage
                            If sPage1.Length > 40 Then sPage1 = TIMS.ClearSQM(sPage1.Substring(0, 40)) & "..."
                            strUrl = String.Format(strUrl_f2, sPage, sPage1)
                        Else
                            '回上2層並給ID
                            If sPage.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then sPage = sPage.Substring(0, sPage.Length - 5)
                            strUrl = String.Format(strUrl_f1, sPage, Convert.ToString(drv("funid")))
                        End If
                        labFunName.Text &= strUrl
                    End If
                End If
                chkEnable.Checked = chkData(dtAPF, Convert.ToInt32(drv("funid")))

                If Convert.ToString(Me.ViewState("mmname")) <> labMainMenu.Text Or intCntSch = 0 Then '同類型選單的第一項
                    Dim subs As Integer = Me.ViewState(Convert.ToString(drv("Kind"))) '依類型取得選單數量

                    Me.ViewState("mmname") = labMainMenu.Text
                    e.Item.Cells(0).RowSpan = subs '合併同類型選單
                    e.Item.Cells(0).BackColor = Color.FromArgb(241, 249, 252)
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")))

                    For i As Integer = 0 To e.Item.Cells.Count - 1
                        e.Item.Cells(i).Attributes.Add("id", Convert.ToString(drv("Kind")) & "td" & i)
                    Next

                    trID = 1
                    intCntSch = 1
                Else '非同類型選單第一項的其他項
                    e.Item.Cells(0).Visible = False '隱藏功能類別欄位
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")) & trID)

                    trID += 1
                    intCntSch = 1
                End If

                '設定主選單顏色
                If Convert.ToString(drv("Subs")) <> "0" Then e.Item.BackColor = Color.FromArgb(235, 243, 254)
                If Convert.ToString(drv("SPage")) = "0" Then e.Item.BackColor = Color.FromArgb(235, 243, 254)

                '紀錄第一項的ClientID
                If e.Item.ItemIndex = 0 Then Me.ViewState("itemname") = chkEnable.ClientID
            Case ListItemType.Footer
                '設定全選動作的JavaScript
                chkEnableAll.Attributes.Add("onclick", "Show_SelectAll('" & chkEnableAll.ClientID & "','" & Convert.ToString(Me.ViewState("itemname")) & "'," & DataGrid1.Items.Count & ")")
        End Select
    End Sub

    Private Sub ddlTPlan_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlTPlan.SelectedIndexChanged
        If ddlTPlan.SelectedIndex = 0 Then
            tr_Fun.Visible = False
            tr_Btn.Visible = False
            Clear_Items()
        Else
            tr_Fun.Visible = True
            tr_Btn.Visible = True

            ddlFun.SelectedIndex = 0
            '查詢
            Call search()
        End If
    End Sub

    Private Sub ddlFun_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlFun.SelectedIndexChanged
        '查詢
        Call search()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class