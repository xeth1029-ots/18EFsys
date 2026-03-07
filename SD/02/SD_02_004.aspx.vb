Partial Class SD_02_004
    Inherits AuthBasePage

    Const Cst_index As Integer = 0
    Const Cst_OrgName As Integer = 6 '機構名稱
    Const Cst_ClassName As Integer = 8 '班級名稱
    Const Cst_AppliedResult As Integer = 9 '班級審核狀態
    Const Cst_TransFlag As Integer = 11 '是否轉班

    'Dim FunDr As DataRow
    'Dim RelshipTable As DataTable
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        tr_audit1.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.LabTMID.Text = "訓練業別"
        End If

        '20080818 andy
        '--------------------
        IsApprPaper.SelectedValue = "Y"
        audit.SelectedValue = "Y"
        '--------------------
        CyclType.Attributes.Add("onKeyPress", "if(event.keyCode==13){if(isNaN(this.value) ){ alert('「期別」欄位請填寫數字！'); return false;} else { return true; } }")

        Dim v_ddlyears As String = TIMS.GetListValue(ddlyears)
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), v_ddlyears)

        '2018-09-19 經討論此功能改暫不提供快捷功能（因無法從rid正確回推找planid,len(auth_relship.rid)=1 的planid都是記錄0）
        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If

        Dim PlanKind As String
        '依sm.UserInfo.PlanID取得PlanKind
        PlanKind = TIMS.Get_PlanKind(Me, objconn)
        If PlanKind = "1" Then
            dtPlan.Columns(6).Visible = False
        Else
            dtPlan.Columns(6).Visible = True
        End If
        Me.msg.Text = ""
        '分頁設定 Start
        PageControler1.PageDataGrid = dtPlan
        '分頁設定 End

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '    FunDr = FunDrArray(0)
        '    If FunDr("Sech") = 1 Then
        '        btnQuery.Enabled = True
        '    Else
        '        btnQuery.Enabled = False
        '    End If
        'End If

        'btnQuery.Enabled = False
        'If blnCanSech Then btnQuery.Enabled = True

        If Not Me.IsPostBack Then
            DataGridTable.Visible = False
            ddlyears = TIMS.GetSyear(ddlyears)

            ' 年度帶預設值
            Common.SetListItem(ddlyears, sm.UserInfo.Years)

            '取得訓練機構
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '取得訓練計畫
            TPlanid.Value = sm.UserInfo.TPlanID

            'add 取得轄區及計畫代碼
            hidPlanID.Value = sm.UserInfo.PlanID
            hidDistID.Value = sm.UserInfo.DistID
        End If

        ''帶入查詢參數
        'If Not IsPostBack Then
        '    'If Not Session("search") Is Nothing Then
        '    '    Common.SetListItem(ddlyears, TIMS.GetMyValue(Session("search"), "ddlyears"))
        '    '    TB_career_id.Text = TIMS.GetMyValue(Session("search"), "TB_career_id")
        '    '    TrainValue.Value = TIMS.GetMyValue(Session("search"), "trainValue")
        '    '    center.Text = TIMS.GetMyValue(Session("search"), "center")
        '    '    RIDValue.Value = TIMS.GetMyValue(Session("search"), "RIDValue")
        '    '    ClassName.Text = TIMS.GetMyValue(Session("search"), "ClassName")
        '    '    Common.SetListItem(IsApprPaper, TIMS.GetMyValue(Session("search"), "IsApprPaper"))
        '    '    Common.SetListItem(audit, TIMS.GetMyValue(Session("search"), "audit"))
        '    '    PageControler1.PageIndex = TIMS.GetMyValue(Session("search"), "PageIndex")
        '    '    btnQuery_Click(sender, e)
        '    '    Session("search") = Nothing
        '    'End If
        'End If

        '執行 Session("_SearchStr") 保留查詢值
        Call GetKeepSearch() 'sender, e
    End Sub

    Sub LoadData()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        '依sm.UserInfo.PlanID取得PlanKind
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        '取出訓練計畫名稱
        TPlanName.Text = TIMS.GetPlanName(sm.UserInfo.PlanID, objconn)

        '建立訓練的DATATABLE
        Dim StrSql As String = ""
        StrSql &= " select I1.DistID, O1.Orgid, P1.PlanID,O1.OrgName,P1.STDate,P1.FDDate " & vbCrLf
        StrSql &= " ,P1.ComIDNO,P1.PlanYear,P1.TMID,P1.AppliedResult" & vbCrLf
        StrSql &= " ,P1.SeqNO,K1.PlanName" & vbCrLf
        StrSql &= " ,CASE when K2.JobID is null then K2.TrainName else K2.JobName  end TrainName " & vbCrLf
        StrSql &= " ,P1.ClassName,P1.AdmPercent " & vbCrLf
        StrSql &= " ,P1.CyclType,P1.AppliedDate,P2.VerReason,P1.TransFlag,P1.RID,C1.OCID " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql &= " ,P3.FirResult, P3.SecResult, P3.IsApprPaper " & vbCrLf
        Else
            StrSql &= " ,'X' FirResult,'X' SecResult , NULL IsApprPaper" & vbCrLf   '給個假資料
        End If
        StrSql &= " ,f23.OrgName2 " & vbCrLf
        StrSql &= " From Plan_PlanInfo P1 " & vbCrLf
        StrSql &= " JOIN Key_Plan K1 ON P1.TPlanID=K1.TPlanID " & vbCrLf
        StrSql &= " LEFT JOIN Key_TrainType K2 ON P1.TMID=K2.TMID " & vbCrLf
        StrSql &= " JOIN ID_Plan I1 ON P1.PlanID=I1.PlanID " & vbCrLf
        StrSql &= " JOIN Org_OrgInfo O1 ON P1.ComIDNO=O1.ComIDNO " & vbCrLf
        StrSql &= " LEFT JOIN MVIEW_RELSHIP23 f23 ON f23.RID3 =P1.RID" & vbCrLf
        StrSql &= " LEFT JOIN Plan_VerRecord P2 ON P1.PlanID=P2.PlanID and P1.ComIDNO=P2.ComIDNO and P1.SeqNo=P2.SeqNo " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql &= " LEFT JOIN Plan_VerReport P3 ON P1.PlanID=P3.PlanID and P1.ComIDNO=P3.ComIDNO and P1.SeqNo=P3.SeqNo " & vbCrLf
        End If
        StrSql &= " LEFT JOIN Class_ClassInfo C1 ON P1.PlanID=C1.PlanID and P1.ComIDNO=C1.ComIDNO and P1.SeqNo=C1.SeqNo " & vbCrLf
        StrSql &= " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            StrSql &= " AND I1.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            StrSql &= " AND I1.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            StrSql &= " AND I1.TPlanID='" & TPlanid.Value & "' " & vbCrLf
            StrSql &= " AND I1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        StrSql &= " and C1.IsSuccess='Y'" & vbCrLf
        StrSql &= " and P1.IsApprPaper='Y' " & vbCrLf

        '如果不是產學訓計畫 penny 2007/10/17
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Y:已審核,N:審核中
            StrSql &= "and P1.AppliedResult IN ('Y','N') " & vbCrLf
        End If

        If sm.UserInfo.RID = "A" Then
            StrSql &= " and P1.RID IN (SELECT RID FROM Auth_Relship WHERE relship like '" & RelShip & "%')" & vbCrLf
        Else
            If PlanKind = 2 Then
                StrSql &= " and P1.RID IN (SELECT RID FROM Auth_Relship WHERE relship like '" & RelShip & "%')" & vbCrLf
            End If
            If sm.UserInfo.LID = 0 OrElse sm.UserInfo.LID = 1 Then
                StrSql &= " and I1.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            End If
        End If
        If Me.ddlyears.SelectedValue <> "" Then
            StrSql &= " and P1.PlanYear='" & Me.ddlyears.SelectedValue & "'" & vbCrLf
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If Me.jobValue.Value <> "" Then
                StrSql &= " and (" & vbCrLf
                StrSql &= "  P1.TMID = " & Me.jobValue.Value & vbCrLf
                StrSql &= " OR P1.TMID IN ( " & vbCrLf
                StrSql &= " select TMID from Key_TrainType where parent IN ( " & vbCrLf '職類別
                StrSql &= " select TMID from Key_TrainType where parent IN ( " & vbCrLf '業別
                StrSql &= " select TMID from Key_TrainType where busid ='G') " & vbCrLf '產業人才投資方案類
                StrSql &= " AND tmid =" & Me.jobValue.Value & vbCrLf
                StrSql &= " )))" & vbCrLf
            End If
        Else
            If Me.trainValue.Value <> "" Then
                StrSql = StrSql & " and P1.TMID='" & Me.trainValue.Value & "'" & vbCrLf
            End If
        End If

        If txtCJOB_NAME.Text <> "" Then   '通俗職類
            StrSql &= " and P1.CJOB_UNKEY = " & cjobValue.Value & "" & vbCrLf
        End If

        If ClassName.Text <> "" Then
            StrSql &= " and ClassName like'%" & ClassName.Text & "%'" & vbCrLf
        End If
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    StrSql &= " and P1.CyclType='0" & Int(CyclType.Text) & "'" & vbCrLf
                Else
                    StrSql &= " and P1.CyclType='" & CyclType.Text & "'" & vbCrLf
                End If
            End If
        End If

        StrSql &= " and C1.OCID  IS NOT NULL "  '20081030 andy edit
        StrSql &= " and P1.AppliedResult ='Y'  and P1.TransFlag='Y' " & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(StrSql, objconn)

        DataGridTable.Visible = False
        Me.msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            Me.msg.Text = ""

            If ViewState("sort") = "" Then ViewState("sort") = "STDate,ClassName"
            PageControler1.Sort = ViewState("sort")
            'PageControler1.SqlString = StrSql
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dtPlan) '顯示列數不正確

        If CheckInput(ClassName.Text) <> "" Then
            Common.RespWrite(Me, "<script>alert('" & "「班級名稱」 欄位輸入字串中含有不合法字元" & CheckInput(ClassName.Text) & "') ;</script>")
            Exit Sub
        End If

        ClassName.Text = Replace(ClassName.Text, "'", "''")

        dtPlan.CurrentPageIndex = 0
        If CyclType.Text <> "" Then
            If Not IsNumeric(CyclType.Text) Then
                Common.MessageBox(Me, "期別需輸入數字型態!!")
                Exit Sub
            End If
        End If
        LoadData()
        If IsApprPaper.SelectedIndex = 0 Then
            dtPlan.Columns(9).Visible = False
            dtPlan.Columns(10).Visible = False
            dtPlan.Columns(11).Visible = False
            dtPlan.Columns(12).Visible = True
        ElseIf IsApprPaper.SelectedIndex = 1 Then
            dtPlan.Columns(9).Visible = False
            dtPlan.Columns(10).Visible = False
            dtPlan.Columns(11).Visible = False
            dtPlan.Columns(12).Visible = False
        End If
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call Search1()
    End Sub

    Function CheckInput(ByVal parameter As String) As String
        Dim blackList As String() = {"'", "--", ";--", ";", "/*", "*/", "@@",
                                     "@", "char", "nchar", "varchar", "nvarchar", "alter",
                                     "begin", "cast", "create", "cursor", "declare", "delete",
                                     "drop", "end", "exec", "execute", "fetch", "insert",
                                     "kill", "open", "select", "sys", "sysobjects", "syscolumns", "table",
                                     "update"}
        Dim strPos As Integer = 0
        Dim blackListlen As Integer = 0
        Dim InputStr As String = parameter
        Dim errMsg As String = ""
        For i As Integer = 0 To blackList.Length - 1
            blackListlen = blackList(i).Length()
            strPos = InStr(1, InputStr, blackList(i))
            If strPos <> 0 Then
                Select Case blackList(i)
                    Case "'"
                        'InputStr = Replace(InputStr, blackList(i), "''")
                        errMsg += "「 單引號‘」"
                    Case Else
                        errMsg += "「 " & blackList(i) & " 」"
                End Select
            End If
        Next
        'If errMsg <> "" Then Return errMsg
        Return errMsg
    End Function

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but1 As LinkButton = e.Item.FindControl("but1")
                e.Item.Cells(Cst_index).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                'If drv("RID").ToString.Length <> 1 Then
                '    If RelshipTable.Select("RID='" & drv("RID") & "'").Length <> 0 Then
                '        Dim Relship As String
                '        Dim Parent As String
                '        Relship = RelshipTable.Select("RID='" & drv("RID") & "'")(0)("Relship")
                '        Parent = Split(Relship, "/")(Split(Relship, "/").Length - 3)
                '        If RelshipTable.Select("RID='" & Parent & "'").Length <> 0 Then
                '            e.Item.Cells(Cst_OrgName).Text = RelshipTable.Select("RID='" & Parent & "'")(0)("OrgName")
                '        End If
                '    End If
                'End If
                If Convert.ToString(drv("OrgName2")) <> "" Then
                    e.Item.Cells(Cst_OrgName).Text = Convert.ToString(drv("OrgName2"))
                End If
                e.Item.Cells(Cst_ClassName).Text = drv("ClassName").ToString
                If drv("CyclType").ToString <> "" Then
                    If Int(drv("CyclType")) <> 0 Then
                        e.Item.Cells(Cst_ClassName).Text += "第" & Int(drv("CyclType")) & "期"
                    End If
                End If
                Select Case drv("AppliedResult").ToString
                    Case "Y"
                        e.Item.Cells(Cst_AppliedResult).Text += "班級審核通過"
                    Case "N"
                        e.Item.Cells(Cst_AppliedResult).Text += "<font color=red>班級審核不通過</font>"
                    Case "M"
                        e.Item.Cells(Cst_AppliedResult).Text += "請修正資料"
                    Case "O"
                        e.Item.Cells(Cst_AppliedResult).Text += "審核後修正"
                    Case "R"
                        e.Item.Cells(Cst_AppliedResult).Text += "班級退件修正"
                    Case Else
                        e.Item.Cells(Cst_AppliedResult).Text += "班級審核中"
                End Select
                Select Case UCase(drv("IsApprPaper").ToString) 'Plan_VerReport
                    Case "Y"
                    Case "N"
                        e.Item.Cells(Cst_AppliedResult).Text += "<br><font color=red>非正式開班計畫表</font>"
                    Case Else
                        e.Item.Cells(Cst_AppliedResult).Text += "<br><font color=red>未填寫開班計畫表</font>"
                End Select

                e.Item.Cells(13).Text = ""
                If Convert.ToString(drv("DistID")) <> "" Then
                    e.Item.Cells(13).Text = TIMS.Get_DistName1(drv("DistID"))
                End If

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    'Dim FirResult As String = ""
                    'Dim SecResult As String = ""
                    'If IsDBNull(drv("FirResult")) = False Then FirResult = drv("FirResult").ToString
                    'If IsDBNull(drv("SecResult")) = False Then SecResult = drv("SecResult").ToString

                    'Select Case FirResult
                    '    Case "Y"
                    '        e.Item.Cells(Cst_AppliedResult).Text = "初審通過<br>"
                    '    Case "N"
                    '        e.Item.Cells(Cst_AppliedResult).Text = "<font color=red>初審不通過</font><br>"
                    '    Case "R"
                    '        e.Item.Cells(Cst_AppliedResult).Text = "初審退件修正<br>"
                    '    Case Else
                    '        e.Item.Cells(Cst_AppliedResult).Text = "尚未初審<br>"
                    'End Select

                    '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
                    e.Item.Cells(Cst_AppliedResult).Text = ""

                    Select Case drv("AppliedResult").ToString
                        Case "Y"
                            e.Item.Cells(Cst_AppliedResult).Text += "班級審核通過"
                        Case "N"
                            e.Item.Cells(Cst_AppliedResult).Text += "<font color=red>班級審核不通過</font>"
                        Case "M"
                            e.Item.Cells(Cst_AppliedResult).Text += "請修正資料"
                        Case "O"
                            e.Item.Cells(Cst_AppliedResult).Text += "審核後修正"
                        Case "R"
                            e.Item.Cells(Cst_AppliedResult).Text += "班級退件修正"
                        Case Else
                            e.Item.Cells(Cst_AppliedResult).Text += "班級審核中"
                    End Select
                    Select Case UCase(drv("IsApprPaper").ToString) 'Plan_VerReport
                        Case "Y"
                        Case "N"
                            e.Item.Cells(Cst_AppliedResult).Text += "<br><font color=red>非正式開班計畫表</font>"
                        Case Else
                            e.Item.Cells(Cst_AppliedResult).Text += "<br><font color=red>未填寫開班計畫表</font>"
                    End Select
                Else
                    If Convert.IsDBNull(drv("AppliedResult")) Then
                        e.Item.Cells(Cst_AppliedResult).Text = "審核中"
                    Else
                        Select Case drv("AppliedResult")
                            Case "Y"
                                e.Item.Cells(Cst_AppliedResult).Text = "審核通過"
                            Case "N"
                                e.Item.Cells(Cst_AppliedResult).Text = "審核不通過"
                            Case "M"
                                e.Item.Cells(Cst_AppliedResult).Text = "請修正資料"
                            Case "O"
                                e.Item.Cells(Cst_AppliedResult).Text = "審核後修正"
                            Case "R"
                                e.Item.Cells(Cst_AppliedResult).Text = "退件修正"
                        End Select
                    End If
                End If
                Select Case drv("TransFlag").ToString  '是否已轉班
                    Case "Y"
                        e.Item.Cells(Cst_TransFlag).Text = "是"
                    Case "N"
                        e.Item.Cells(Cst_TransFlag).Text = "否"
                End Select

                but1.CommandArgument = "PlanID=" & drv("PlanID") & "&DistID=" & drv("DistID") & "&OCID=" & drv("OCID") & "&OrgID=" & drv("OrgID") & "&RID=" & drv("RID") & "&Year=" & ddlyears.SelectedItem.Text & ""

            Case ListItemType.Header
                If Me.ViewState("sort") <> "" Then
                    Dim mylabel As String
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "ClassName", "ClassName DESC"
                            mylabel = "ComName"
                            i = 8
                            If Me.ViewState("sort") = "ClassName" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "AppliedDate", "AppliedDate DESC"
                            mylabel = "ComName"
                            i = 2
                            If Me.ViewState("sort") = "AppliedDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "STDate", "STDate DESC"
                            mylabel = "ComName"
                            i = 4
                            If Me.ViewState("sort") = "STDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "FDDate", "FDDate DESC"
                            mylabel = "ComName"
                            i = 5
                            If Me.ViewState("sort") = "FDDate" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                        Case "OrgName", "OrgName DESC"
                            mylabel = "ComName"
                            i = 7
                            If Me.ViewState("sort") = "OrgName" Then
                                mysort.ImageUrl = "../../images/SortUp.gif"
                            Else
                                mysort.ImageUrl = "../../images/SortDown.gif"
                            End If
                    End Select
                    If i <> -1 Then
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If
                End If
        End Select

    End Sub

    Public Sub dtPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        KeepSearch()
        If e.CommandName = "update" Then
            TIMS.Utl_Redirect1(Me, "SD_02_004_add.aspx?ID=" & Request("ID") & " &" & e.CommandArgument & "")
        End If
    End Sub

    Private Sub dtPlan_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dtPlan.SortCommand
        If Me.ViewState("sort") <> e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression
        Else
            Me.ViewState("sort") = e.SortExpression & " DESC"
        End If
        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Sub KeepSearch()
        Session("_SearchStr") = "Class=SD_01_004_mail"
        Session("_SearchStr") += "&ddlyears=" & ddlyears.SelectedValue
        Session("_SearchStr") += "&center=" & center.Text
        Session("_SearchStr") += "&RIDValue=" & RIDValue.Value
        Session("_SearchStr") += "&DistID=" & hidDistID.Value
        Session("_SearchStr") += "&PlanID=" & hidPlanID.Value
        Session("_SearchStr") += "&TB_career_id=" & TB_career_id.Text
        Session("_SearchStr") += "&jobValue=" & jobValue.Value
        Session("_SearchStr") += "&txtCJOB_NAME=" & txtCJOB_NAME.Text
        Session("_SearchStr") += "&cjobValue=" & cjobValue.Value
        Session("_SearchStr") += "&ClassName=" & ClassName.Text
        Session("_SearchStr") += "&TrainValue=" & trainValue.Value
        Session("_SearchStr") += "&CyclType=" & CyclType.Text
        Session("_SearchStr") += "&TxtPageSize=" & TxtPageSize.Text
        Session("_SearchStr") += "&PageIndex=" & dtPlan.CurrentPageIndex + 1
        Session("_SearchStr") += "&Button1=" & dtPlan.Visible
        If dtPlan.Visible Then
            Session("_SearchStr") += "&submit=1"
        Else
            Session("_SearchStr") += "&submit=0"
        End If
    End Sub

    '執行 Session("_SearchStr") 保留查詢值
    Sub GetKeepSearch() 'ByVal sender As System.Object, ByVal e As System.EventArgs
        If Not Session("_SearchStr") Is Nothing Then
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
            If TIMS.GetMyValue(Me.ViewState("_SearchStr"), "Class") = "SD_01_004_mail" Then
                Common.SetListItem(ddlyears, TIMS.GetMyValue(Me.ViewState("_SearchStr"), "ddlyears"))
                TB_career_id.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TB_career_id")
                jobValue.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "jobValue")
                txtCJOB_NAME.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "txtCJOB_NAME")
                cjobValue.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "cjobValue")
                trainValue.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TrainValue")
                center.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "center")
                CyclType.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "CyclType")
                RIDValue.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "RIDValue")
                ClassName.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "ClassName")
                Common.SetListItem(IsApprPaper, TIMS.GetMyValue(Me.ViewState("_SearchStr"), "IsApprPaper"))
                Common.SetListItem(audit, TIMS.GetMyValue(Me.ViewState("_SearchStr"), "audit"))
                TxtPageSize.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TxtPageSize")

                Dim MyValue As String = ""
                Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                MyValue = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit")
                If MyValue = "1" Then
                    'btnQuery_Click(sender, e)
                    Call Search1()

                    If IsNumeric(Me.ViewState("PageIndex")) Then
                        '有資料SHOW出 跳頁
                        PageControler1.PageIndex = Me.ViewState("PageIndex")
                        PageControler1.CreateData()
                    End If
                End If
                'If TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit") = "1" Then
                '    Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                '    If IsNumeric(Me.ViewState("PageIndex")) Then PageControler1.PageIndex = Me.ViewState("PageIndex")
                '    btnQuery_Click(sender, e)
                '    'PageControler1.PageIndex = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                '    'PageControler1.CreateData()
                'End If

                hidDistID.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "DistID")
                hidPlanID.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PlanID")

            End If
            Session("_SearchStr") = Nothing
        End If
    End Sub

    Private Sub Btn_OrgSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_OrgSet.Click
        KeepSearch()
        'Dim strDistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        TIMS.CloseDbConn(objconn)

        'TIMS.Utl_Redirect1(Me, "SD_02_004_add.aspx?DistID=" & sm.UserInfo.DistID & "&ID=" & Request("ID") & "&RID=" & RIDValue.Value & "&OrgID=" & Org.Value & "&Year=" & ddlyears.SelectedItem.Text & "")
        If sm.UserInfo.LID = 0 Then
            TIMS.Utl_Redirect1(Me, "SD_02_004_add.aspx?DistID=" & hidDistID.Value & "&ID=" & Request("ID") & "&RID=" & RIDValue.Value & "&OrgID=" & Org.Value & "&Year=" & ddlyears.SelectedItem.Text & "&PlanID=" & hidPlanID.Value)
        Else
            TIMS.Utl_Redirect1(Me, "SD_02_004_add.aspx?DistID=" & sm.UserInfo.DistID & "&ID=" & Request("ID") & "&RID=" & RIDValue.Value & "&OrgID=" & Org.Value & "&Year=" & ddlyears.SelectedItem.Text & "&PlanID=" & sm.UserInfo.PlanID)
        End If

    End Sub
End Class
