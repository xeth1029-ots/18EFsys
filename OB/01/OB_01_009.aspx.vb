Partial Class OB_01_009
    Inherits AuthBasePage

    Const Cst_TenderSDate As Integer = 5

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), lblTitle1, lblTitle2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            txtJudgeNum.Text = ""
            'btnSave.Visible = False
            ddlyears = TIMS.GetSyear(ddlyears, Year(Now) - 1, Year(Now) + 3, True)
            TPlanID = TIMS.Get_TPlan(TPlanID)

            DataGridTable.Visible = False
            DataGridTable2.Visible = False
            btnSave.Attributes.Add("onclick", "return chkdata();")
            btnSave.Visible = False

            panelSch.Visible = True
            panelSet.Visible = False
        End If

    End Sub


    Sub Query(Optional ByVal sql As String = "")

        Dim parms As New Hashtable
        Dim dt As DataTable
        If sql = "" Then
            sql = "" & vbCrLf
            sql += " select b.PlanName, a.* " & vbCrLf
            sql += " from OB_Tender a " & vbCrLf
            sql += " JOIN OB_Plan b on a.PlanSN=b.PlanSN " & vbCrLf
            sql += " WHERE 1=1 " & vbCrLf
            parms.Clear()
            If ddlyears.SelectedValue <> "" Then
                sql += " AND a.Years=@Years" & vbCrLf
                parms.Add("Years", ddlyears.SelectedValue)
            End If
            If TPlanID.SelectedValue <> "" Then
                sql += " AND a.TPlanID=@TPlanID" & vbCrLf
                parms.Add("TPlanID", TPlanID.SelectedValue)
            End If
            PlanName.Text = TIMS.ClearSQM(PlanName.Text)
            If PlanName.Text <> "" Then
                Dim s_PlanName_lk As String = String.Concat("%", PlanName.Text, "%")
                sql += " AND b.PlanName like @PlanName_lk" & vbCrLf
                parms.Add("PlanName_lk", s_PlanName_lk)
            End If
            TenderName.Text = TIMS.ClearSQM(TenderName.Text)
            If TenderName.Text <> "" Then
                Dim s_TenderName_lk As String = String.Concat("%", TenderName.Text, "%")
                sql += " AND a.TenderName like @TenderName_lk" & vbCrLf
                parms.Add("TenderName_lk", s_TenderName_lk)
            End If
            Sponsor.Text = TIMS.ClearSQM(Sponsor.Text)
            If Sponsor.Text <> "" Then
                Dim s_Sponsor_lk As String = String.Concat("%", Sponsor.Text, "%")
                sql += " AND a.Sponsor like @Sponsor_lk" & vbCrLf
                parms.Add("Sponsor_lk", s_Sponsor_lk)
            End If
        End If


        DataGridTable.Visible = False
        msg.Text = "查無資料!!"

        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Query()

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(Cst_TenderSDate).Text = Common.FormatDate(drv("TenderSDate"))

                Dim btnSet As Button = e.Item.FindControl("btnSet")
                btnSet.CommandArgument = drv("tsn")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        txtJudgeNum.Text = ""
        Select Case e.CommandName
            Case "set"
                panelSch.Visible = False
                panelSet.Visible = True

                Dim sql As String = ""
                Dim dr As DataRow
                Dim dt As DataTable
                Dim i As Integer

                hidTsn.Value = TIMS.ClearSQM(e.CommandArgument)
                If hidTsn.Value = "" Then
                    panelSch.Visible = True
                    panelSet.Visible = False
                    Common.MessageBox(Me, "尚未挑選投標單位！")
                    Return
                End If

                Dim parms As New Hashtable
                parms.Add("TSN", hidTsn.Value)
                sql = " select a.TCsn, a.csn, b.OrgName "
                sql += " from OB_TContractor a "
                sql += " join OB_Contractor b on b.csn=a.csn "
                sql += " WHERE a.tsn=@TSN "
                dt = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count = 0 Then
                    panelSch.Visible = True
                    panelSet.Visible = False
                    Common.MessageBox(Me, "尚未挑選投標單位！")
                    Return
                End If
                'DataGridTable2.Visible = True
                DataGrid2.DataSource = dt
                DataGrid2.DataBind()

                Dim parms2 As New Hashtable
                parms2.Add("TSN", hidTsn.Value)
                sql = "select JudgeNum from OB_Tender WHERE tsn=@TSN "
                dr = DbAccess.GetOneRow(sql, objconn, parms2)

                txtJudgeNum.ReadOnly = False
                btnSend.Visible = True
                DataGridTable2.Visible = False
                btnSave.Visible = False

                If Not dr Is Nothing Then
                    If Convert.ToString(dr("JudgeNum")) <> "" Then
                        If CInt(dr("JudgeNum")) >= 0 Then
                            For i = 1 To CInt(dr("JudgeNum"))
                                DataGrid2.Columns(2 + i).Visible = True
                            Next
                            For i = 2 + CInt(dr("JudgeNum")) + 1 To 12
                                DataGrid2.Columns(i).Visible = False
                            Next

                            txtJudgeNum.ReadOnly = True
                            txtJudgeNum.Text = dr("JudgeNum")

                            btnSend.Visible = False
                            DataGridTable2.Visible = True
                            btnSave.Visible = True
                        End If
                    End If
                End If

        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim sql As String = ""
                'Dim dr As DataRow
                Dim dt As DataTable
                'Dim i As Integer
                Dim txtScore As TextBox = e.Item.FindControl("txtScore")
                'Dim labScore As Label = e.Item.FindControl("labScore")
                Dim lblTCsn As Label = e.Item.FindControl("lblTCsn")
                'Dim txtScore1 As TextBox = e.Item.FindControl("txtScore1")
                Dim txtTemp As TextBox
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                txtScore.Text = 0
                'labScore.Text = 0
                lblTCsn.Text = drv("TCsn")

                For i As Integer = 0 To 10 - 1
                    txtTemp = e.Item.FindControl("txtScore" & (i + 1).ToString)
                    If Not txtTemp Is Nothing Then
                        txtTemp.Attributes.Add("onblur", "sum(" & e.Item.Cells(0).Text & ");")
                    End If
                Next

                Dim parms As New Hashtable
                parms.Add("TCSN", drv("TCsn"))
                sql = " select TScore,TCsn,JudgeNumber from OB_TScore WHERE TCsn=@TCSN order by JudgeNumber"
                dt = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        txtTemp = e.Item.FindControl("txtScore" & (i + 1).ToString)
                        txtTemp.Text = dt.Rows(i)("TScore")
                        'txtTemp.Attributes.Add("onblur", "sum(" & e.Item.Cells(0).Text & ");")
                        txtScore.Text = CInt(txtScore.Text) + CInt(txtTemp.Text)
                        'labScore.Text = CInt(labScore.Text) + CInt(txtTemp.Text)
                    Next
                End If

        End Select
    End Sub

    Private Sub btnSend_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSend.Click
        Dim i As Integer

        If Trim(txtJudgeNum.Text) = "" Or Not IsNumeric(txtJudgeNum.Text) Then
            Common.RespWrite(Me, "<script>alert('請輸入評審委員人數');</script>")
            txtJudgeNum.Text = ""
            btnSave.Visible = False
            Exit Sub
        Else
            If CInt(txtJudgeNum.Text) < 1 Or CInt(txtJudgeNum.Text) > 10 Then
                Common.RespWrite(Me, "<script>alert('評審委員人數介於1~10人');</script>")
                txtJudgeNum.Text = ""
                btnSave.Visible = False
                Exit Sub
            End If
        End If

        For i = 1 To CInt(txtJudgeNum.Text)
            DataGrid2.Columns(2 + i).Visible = True
        Next
        For i = 2 + CInt(txtJudgeNum.Text) + 1 To 12
            DataGrid2.Columns(i).Visible = False
        Next

        btnSave.Visible = True
        DataGridTable2.Visible = True
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim tConn As SqlConnection = DbAccess.GetConnection
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try
            '刪除原評分
            Dim d_sql As String = ""
            d_sql = " delete OB_TScore where TCsn in (select TCsn from OB_TContractor where tsn=@TSN)"
            Dim d_cmd As New SqlCommand(d_sql, tConn, trans)
            With d_cmd
                .Parameters.Clear()
                .Parameters.Add("TSN", SqlDbType.Int).Value = CInt(hidTsn.Value)
                .ExecuteNonQuery()
            End With

            '修改評選者數量
            Dim u_sql As String = " update OB_Tender set JudgeNum=@JudgeNum where tsn=@TSN'"
            Dim u_cmd As New SqlCommand(u_sql, tConn, trans)
            With u_cmd
                .Parameters.Clear()
                .Parameters.Add("JudgeNum", SqlDbType.Int).Value = Val(txtJudgeNum.Text)
                .Parameters.Add("TSN", SqlDbType.Int).Value = CInt(hidTsn.Value)
                .ExecuteNonQuery()
            End With

            '新增評分
            Dim i_sql As String = ""
            i_sql = " insert into OB_TScore (TCsn, JudgeNumber, TScore, CreateAcct, CreateTime) "
            i_sql += " values(@TCsn, @JudgeNumber, @TScore, @CreateAcct, getdate())"
            Dim cmdIns As New SqlCommand(i_sql, tConn, trans)

            For Each item As DataGridItem In DataGrid2.Items

                Dim lblTCsn As Label = item.FindControl("lblTCsn")
                'txtScore = item.FindControl("txtScore")
                For i As Integer = 1 To CInt(txtJudgeNum.Text)
                    Dim txtTemp As TextBox = item.FindControl("txtScore" & i.ToString)
                    If Convert.ToString(txtTemp.Text).Trim = "" Then
                        txtTemp.Text = 0
                    End If
                    With cmdIns
                        .Parameters.Clear()
                        .Parameters.Add("TCsn", SqlDbType.Int).Value = CInt(lblTCsn.Text)
                        .Parameters.Add("JudgeNumber", SqlDbType.Int).Value = i
                        .Parameters.Add("TScore", SqlDbType.Int).Value = CInt(txtTemp.Text)
                        .Parameters.Add("CreateAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .ExecuteNonQuery()
                    End With
                Next

            Next

            trans.Commit()
            Common.RespWrite(Me, "<script>alert('儲存成功');</script>")
            panelSch.Visible = True
            panelSet.Visible = False

        Catch ex As Exception
            trans.Rollback()
            TIMS.CloseDbConn(tConn)
            Common.MessageBox(Me, ex.ToString)
            Return
            'Finally objConn.Close() objConn.Dispose() objConn = Nothing
        End Try

        TIMS.CloseDbConn(tconn)
    End Sub

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExit.Click
        panelSch.Visible = True
        panelSet.Visible = False
    End Sub

End Class
