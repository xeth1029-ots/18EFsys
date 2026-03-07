Partial Class TC_01_004_plan
    Inherits AuthBasePage

    'TC_01_004_InsertPlan.ASPX
    Dim flag_oTestEnv As Boolean = False '測試
    'TIMS/職前/在職/產投

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        If TIMS.sUtl_ChkTest Then flag_oTestEnv = True '測試

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Me.LabTMID.Text = "訓練業別"

        If Not Page.IsPostBack Then
            Me.ViewState("ClassSearchStr") = Session("ClassSearchStr")
            Session("ClassSearchStr") = Nothing
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'save.Enabled = False
        'If au.blnCanSech Then save.Enabled = True
        'check_add.Value = "0"
        'If au.blnCanAdds Then check_add.Value = "1"
        'check_add.Value = "1"
        If Not IsPostBack Then
            '點進功能第1次顯示
            '檢核該計畫班別代碼是否有設定 依大計畫、轄區、年度
            If Not TIMS.CheckCLSID(Me, objconn) Then Common.MessageBox(Me, "班別代碼尚未建檔，請洽本分署承辦人詢問!")
        End If
    End Sub

    '查詢 SQL
    Sub Search1()
        Me.Panel.Visible = True
        Dim aryrow() As String = {"序號", "班別名稱", "訓練計畫", "訓練職類", "通俗職類", "功能"}
        Dim cell As New HtmlTableCell
        Dim row As New HtmlTableRow
        'Dim i, j As Integer
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then aryrow(3) = "訓練業別"

        For i As Integer = 0 To aryrow.Length - 1
            cell = New HtmlTableCell
            cell.InnerText = aryrow(i)
            row.Cells.Add(cell)
            'row.Style("Color") = "#FFFFFF"
        Next
        row.Align = "center"
        row.Attributes("Class") = "head_navy"
        Me.search_tbl.Rows.Add(row)

        Dim parms As Hashtable = New Hashtable()
        'Dim sqlAdapter As SqlDataAdapter = Nothing
        'Dim plan_info As DataTable = Nothing
        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNO" & vbCrLf
        sqlstr &= " ,a.RID ,d.PlanName " & vbCrLf
        sqlstr &= " ,'[' + s.CJOB_No + ']' + s.CJOB_NAME CJOB_NAME " & vbCrLf
        If iPYNum >= 3 Then
            sqlstr &= " ,e.TrainID " & vbCrLf
            sqlstr &= " ,e.TrainName " & vbCrLf
            sqlstr &= " ,e.JobID " & vbCrLf
            sqlstr &= " ,e.JobName " & vbCrLf
        Else
            sqlstr &= " ,CASE WHEN e.JobID IS NULL THEN e.TrainID ELSE e.JobID END TrainID " & vbCrLf
            sqlstr &= " ,CASE WHEN e.JobID IS NULL THEN e.trainName ELSE e.JobName END TrainName " & vbCrLf
            sqlstr &= " ,CASE WHEN e.JobID IS NULL THEN e.TrainID ELSE e.JobID END JobID " & vbCrLf
            sqlstr &= " ,CASE WHEN e.JobID IS NULL THEN e.trainName ELSE e.JobName END JobName " & vbCrLf
        End If
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) ClassName" & vbCrLf
        sqlstr &= " FROM Plan_PlanInfo a " & vbCrLf
        sqlstr &= " JOIN Org_OrgInfo b ON a.ComIDNO = b.ComIDNO " & vbCrLf
        sqlstr &= " JOIN id_plan ip ON ip.PlanID = a.PlanID " & vbCrLf
        sqlstr &= " JOIN Key_Plan d ON a.TPlanID = d.TPlanID " & vbCrLf
        sqlstr &= " JOIN Key_TrainType e ON a.TMID = e.TMID " & vbCrLf
        sqlstr &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY = a.CJOB_UNKEY " & vbCrLf
        sqlstr &= " WHERE 1=1 " & vbCrLf
        'sqlstr &= "  AND a.TransFlag = 'N' " & vbCrLf
        If Not flag_oTestEnv Then sqlstr &= " AND a.TransFlag = 'N' " & vbCrLf '正式
        sqlstr &= " AND a.AppliedResult = 'Y' " & vbCrLf
        sqlstr &= " AND a.IsApprPaper = 'Y' " & vbCrLf
        Select Case sm.UserInfo.RID '登入者
            Case "A" '署(局)
                sqlstr &= " AND ip.TPLANID = @TPLANID " & vbCrLf
                sqlstr &= " AND ip.YEARS = @YEARS " & vbCrLf
                parms.Add("TPLANID", sm.UserInfo.TPlanID)
                parms.Add("YEARS", sm.UserInfo.Years)

            Case Else '分署(中心)或委訓，只能查自已的計畫
                sqlstr &= " AND a.PLANID = @PLANID " & vbCrLf
                parms.Add("PLANID", sm.UserInfo.PlanID)

        End Select

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        If Me.RIDValue.Value <> "" Then
            '有選取,帶選取的 RIDValue
            'sqlstr &= " and a.RID='" & Me.RIDValue.Value & "'" & vbCrLf
            Select Case Convert.ToString(sm.UserInfo.LID)
                Case "0"
                    '沒有選取,帶預設值
                    sqlstr &= "  AND a.RID LIKE @RID " & vbCrLf
                    parms.Add("RID", RIDValue.Value & "%")
                Case "1"
                    '沒有選取,帶預設值
                    sqlstr &= "  AND a.RID LIKE @RID " & vbCrLf
                    parms.Add("RID", RIDValue.Value & "%")
                Case "2"
                    '沒有選取,帶預設值
                    sqlstr &= "  AND a.RID = @RID " & vbCrLf
                    parms.Add("RID", RIDValue.Value)
            End Select
        Else
            '沒有選取,帶預設值
            sqlstr &= "  AND a.RID = @RID " & vbCrLf
            parms.Add("RID", sm.UserInfo.RID)
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If iPYNum >= 3 Then
                If trainValue.Value <> "" Then
                    sqlstr &= "  AND a.TMID = @TMID " & vbCrLf
                    parms.Add("TMID", trainValue.Value)
                End If
            Else
                'Me.LabTMID.Text = "訓練業別"
                If Me.jobValue.Value <> "" Then
                    sqlstr &= "  AND (a.TMID = @TMID " & vbCrLf
                    sqlstr &= "  OR a.TMID IN ( " & vbCrLf
                    sqlstr &= "    SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '職類別
                    sqlstr &= "    SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '業別
                    sqlstr &= "    SELECT TMID FROM Key_TrainType WHERE busid = 'G') " & vbCrLf '產業人才投資方案類
                    sqlstr &= "  AND TMID = @TMID " & vbCrLf
                    sqlstr &= " ))) " & vbCrLf
                    parms.Add("TMID", Me.jobValue.Value)
                End If
            End If
        Else
            If Me.trainValue.Value <> "" Then
                sqlstr &= " and a.TMID=@TMID " & vbCrLf
                parms.Add("TMID", Me.trainValue.Value)
            End If
        End If
        TB_cycltype.Text = TIMS.FmtCyclType(TB_cycltype.Text)
        If TB_cycltype.Text <> "" Then
            sqlstr &= " AND a.CYCLTYPE = @CYCLTYPE "
            parms.Add("CYCLTYPE", TB_cycltype.Text)
        End If
        If txtCJOB_NAME.Text <> "" Then
            sqlstr &= "  AND a.CJOB_UNKEY = @CJOB_UNKEY "
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If
        'If TIMS.sUtl_ChkTest Then sqlstr = cls_test.test2_Sql3() '測試用

        Dim plan_info As DataTable = DbAccess.GetDataTable(sqlstr, objconn, parms)
        If plan_info.Rows.Count = 0 Then
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = "目前沒有計畫可以轉入!!"
            cell.ColSpan = Me.search_tbl.Rows(0).Cells.Count
            row.Cells.Add(cell)
            row.Align = "center"
            Me.search_tbl.Rows.Add(row)
            Exit Sub
        End If

        Dim iNew As Integer = 1
        For Each dr As DataRow In plan_info.Rows
            'Dim classid As String
            row = New HtmlTableRow
            cell = New HtmlTableCell
            cell.InnerText = iNew
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerText = dr("ClassName")
            cell.Align = "left"
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerText = dr("PlanName")
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            If Convert.IsDBNull(dr("TrainName")) Then
                Dim space As String
                space = "&nbsp;"
                cell.InnerHtml = space
            Else
                cell.InnerHtml = String.Format("[{0}]{1}", dr("TrainID"), dr("TrainName"))
            End If
            cell.Align = "left"
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            cell.InnerText = Convert.ToString(dr("CJOB_NAME"))
            cell.Align = "left"
            row.Cells.Add(cell)

            cell = New HtmlTableCell
            'TC_01_004_InsertPlan.ASPX
            If RIDValue.Value = "" Then '沒有選取,帶預設值-轉入
                cell.InnerHtml = "<input type=button class='asp_button_M' value=轉入 onclick=""but_edit(" & dr("PlanID") & ",'" & dr("ComIDNO") & "'," & dr("SeqNO") & "," & Request("ID") & ",'" & sm.UserInfo.RID & "');"">"
            Else '有選取,帶選取的 RIDValue
                cell.InnerHtml = "<input type=button value=轉入 class='asp_button_M' onclick=""but_edit(" & dr("PlanID") & ",'" & dr("ComIDNO") & "'," & dr("SeqNO") & "," & Request("ID") & ",'" & RIDValue.Value & "');"">"
            End If

            'If check_add.Value = "1" Then
            '    'TC_01_004_InsertPlan.ASPX
            '    If RIDValue.Value = "" Then '沒有選取,帶預設值
            '        cell.InnerHtml = "<input type=button class='asp_button_M' value=轉入 onclick=""but_edit(" & dr("PlanID") & ",'" & dr("ComIDNO") & "'," & dr("SeqNO") & "," & Request("ID") & ",'" & sm.UserInfo.RID & "');"">"
            '    Else '有選取,帶選取的 RIDValue
            '        cell.InnerHtml = "<input type=button value=轉入 class='asp_button_M' onclick=""but_edit(" & dr("PlanID") & ",'" & dr("ComIDNO") & "'," & dr("SeqNO") & "," & Request("ID") & ",'" & RIDValue.Value & "');"">"
            '    End If
            'Else
            '    cell.InnerHtml = "<input type=button disabled='true' class='asp_button_M' value=轉入 onclick=""but_edit(" & dr("PlanID") & ",'" & dr("ComIDNO") & "'," & dr("SeqNO") & "," & Request("ID") & ",'" & RIDValue.Value & "');"">"
            'End If

            row.Cells.Add(cell)
            row.Align = "center"
            Me.search_tbl.Rows.Add(row)
            If iNew Mod 2 = 0 Then row.BgColor = "#F5F5F5"
            iNew = iNew + 1
        Next
    End Sub

    '查詢 SQL 設計轉入班級按鈕
    Private Sub save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles save.Click
        Call Search1()
    End Sub

    'Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
    '    Session("ClassSearchStr") = Me.ViewState("ClassSearchStr")
    '   TIMS.Utl_Redirect1(Me, "TC_01_004.aspx?ID=" & Request("ID") & "")
    'End Sub
End Class