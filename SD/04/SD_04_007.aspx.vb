'Imports System.Data.SqlClient
'Imports System.Data
'Imports Turbo

Partial Class SD_04_007
    Inherits AuthBasePage

    Const Cst_index As Integer = 0
    Const Cst_OrgName As Integer = 6 '機構名稱
    Const Cst_ClassName As Integer = 8 '班級名稱
    Const Cst_AppliedResult As Integer = 9 '班級審核狀態
    Const Cst_TransFlag As Integer = 11 '是否轉班

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'Dim sql As String
        'Dim PlanKind As String
        tr_audit1.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.LabTMID.Text = "訓練業別"
            'audit1.Visible = False
            'IsApprPaper.AutoPostBack = False
        End If

        '20080818 andy
        '--------------------
        IsApprPaper.SelectedValue = "Y"
        audit.SelectedValue = "Y"
        '--------------------

        'sql = "SELECT a.RID,a.Relship,b.OrgName FROM "
        'sql += "Auth_Relship a "
        'sql += "JOIN Org_OrgInfo b ON a.OrgID=b.OrgID "
        'RelshipTable = DbAccess.GetDataTable(sql, objConn)

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), v_yearlist)

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'sql = "SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'PlanKind = DbAccess.ExecuteScalar(Sql, objConn)
        Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objConn)
        If sPlanKind = "1" Then
            dtPlan.Columns(6).Visible = False
        Else
            dtPlan.Columns(6).Visible = True
        End If
        Me.msg.Text = ""
        '分頁設定 Start
        PageControler1.PageDataGrid = dtPlan
        '分頁設定 End

        'DistID = sm.UserInfo.DistID
        'LID = sm.UserInfo.LID
        'RID = sm.UserInfo.RID
        'OrgID = sm.UserInfo.OrgID
        'PlanID = sm.UserInfo.PlanID

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    Re_ID.Value = Request("ID")
        '   'Dim FunDr As DataRow = FunDrArray(0)
        '    If FunDr("Sech") = 1 Then
        '        btnQuery.Enabled = True
        '    Else
        '        btnQuery.Enabled = False
        '    End If
        'End If

        'If sm.UserInfo.TPlanID <> "28" Then
        '    audit1.Visible = True
        '    '如果是選擇正式則出現選擇審核中或己審核的 選單 penny 2007/10/17
        '    IsApprPaper.Attributes("onclick") = "checkaudit1();"
        'Else
        '    audit1.Visible = False
        'End If

        'Dim Sqlstr As String = ""
        'Sqlstr = "select TPlanID,PlanKind From ID_Plan where PlanID='" & PlanID & "'"
        'dr = DbAccess.GetOneRow(Sqlstr, objConn)

        If Not Me.IsPostBack Then

            DataGridTable.Visible = False
            yearlist = TIMS.GetSyear(yearlist)

            ' 年度帶預設值
            Common.SetListItem(yearlist, sm.UserInfo.Years)
            '(加強操作便利性)
            RIDValue.Value = sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName 'orgname
            'Dim sqlstring As String = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & RID & "'"
            'Dim orgname As String = DbAccess.ExecuteScalar(sqlstring, objConn)

            '取得訓練計畫
            'Sqlstr = "select TPlanID  from ID_Plan where PlanID=" & sm.UserInfo.PlanID & ""
            'TPlanid.Value = DbAccess.ExecuteScalar(Sqlstr, objConn)
            TPlanid.Value = sm.UserInfo.TPlanID
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            If Session("RestTime") IsNot Nothing Then
                Common.SetListItem(yearlist, TIMS.GetMyValue(Session("RestTime"), "yearlist"))
                TB_career_id.Text = TIMS.GetMyValue(Session("RestTime"), "TB_career_id")
                trainValue.Value = TIMS.GetMyValue(Session("RestTime"), "trainValue")
                txtCJOB_NAME.Text = TIMS.GetMyValue(Session("RestTime"), "txtCJOB_NAME")
                cjobValue.Value = TIMS.GetMyValue(Session("RestTime"), "cjobValue")
                center.Text = TIMS.GetMyValue(Session("RestTime"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("RestTime"), "RIDValue")
                ClassName.Text = TIMS.GetMyValue(Session("RestTime"), "ClassName")
                Common.SetListItem(IsApprPaper, TIMS.GetMyValue(Session("RestTime"), "IsApprPaper"))
                Common.SetListItem(audit, TIMS.GetMyValue(Session("RestTime"), "audit"))
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("RestTime"), "PageIndex")

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("SearchStr"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(Session("RestTime"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                btnQuery_Click(sender, e)
                Session("RestTime") = Nothing
            End If
        End If

    End Sub

    Sub LoadData()
        '取出訓練計畫名稱
        Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objConn)
        TPlanName.Text = TIMS.Get_TPlanName(sm.UserInfo.TPlanID, objConn)

        TPlanid.Value = TIMS.ClearSQM(TPlanid.Value)
        If TPlanid.Value = "" Then TPlanid.Value = sm.UserInfo.TPlanID
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)

        'Dim drTemp As DataRow
        'StrSql = "SELECT PlanName FROM VIEW_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'drTemp = DbAccess.GetOneRow(StrSql, objConn)
        'If Not drTemp Is Nothing Then
        '    TPlanName.Text = drTemp("PlanName").ToString
        'End If

        '建立訓練的DATATABLE
        Dim StrSql As String = ""
        StrSql &= " select I1.DistID, O1.Orgid, P1.PlanID,O1.OrgName,P1.STDate,P1.FDDate " & vbCrLf
        StrSql &= " ,P1.ComIDNO,P1.PlanYear,P1.TMID,P1.AppliedResult" & vbCrLf
        StrSql &= " ,P1.SeqNO,K1.PlanName" & vbCrLf
        StrSql &= " ,CASE when K2.JobID is null then K2.TrainName else K2.JobName end TrainName " & vbCrLf
        StrSql &= " ,P1.ClassName,P1.AdmPercent " & vbCrLf
        StrSql &= " ,P1.CyclType,P1.AppliedDate,P2.VerReason,P1.TransFlag,P1.RID,C1.OCID " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql &= " ,P3.FirResult, P3.SecResult, P3.IsApprPaper " & vbCrLf
        Else
            StrSql &= " ,'X' FirResult, 'X' SecResult, 'X' IsApprPaper " & vbCrLf
        End If
        StrSql &= " ,ar2.OrgName2"
        StrSql &= " FROM Plan_PlanInfo P1" & vbCrLf
        StrSql &= " JOIN Key_Plan K1 ON P1.TPlanID=K1.TPlanID " & vbCrLf
        StrSql &= " JOIN ID_Plan I1 ON P1.PlanID=I1.PlanID " & vbCrLf
        StrSql &= " JOIN Org_OrgInfo O1 ON P1.ComIDNO=O1.ComIDNO " & vbCrLf
        StrSql &= " JOIN Auth_Relship rr ON rr.RID=P1.RID " & vbCrLf
        StrSql &= " JOIN Class_ClassInfo C1 ON P1.PlanID=C1.PlanID and P1.ComIDNO=C1.ComIDNO and P1.SeqNo=C1.SeqNo " & vbCrLf

        StrSql &= " LEFT JOIN KEY_TRAINTYPE K2 ON P1.TMID=K2.TMID " & vbCrLf
        StrSql &= " LEFT JOIN MVIEW_RELSHIP23 ar2 on ar2.RID3=P1.RID " & vbCrLf
        StrSql &= " LEFT JOIN PLAN_VERRECORD P2 ON P1.PlanID=P2.PlanID and P1.ComIDNO=P2.ComIDNO and P1.SeqNo=P2.SeqNo " & vbCrLf

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql += "LEFT JOIN PLAN_VERREPORT P3 ON P1.PlanID=P3.PlanID and P1.ComIDNO=P3.ComIDNO and P1.SeqNo=P3.SeqNo " & vbCrLf
        End If

        StrSql &= " WHERE 1=1" & vbCrLf
        StrSql &= " AND C1.IsSuccess='Y'" & vbCrLf
        StrSql &= " AND P1.IsApprPaper='Y' " & vbCrLf
        StrSql &= " AND P1.AppliedResult ='Y' " & vbCrLf
        StrSql &= " AND P1.TransFlag='Y' " & vbCrLf
        StrSql &= " AND P1.TPlanID='" & TPlanid.Value & "'" & vbCrLf

        If sm.UserInfo.RID = "A" Then
            StrSql &= " AND I1.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            StrSql &= " AND I1.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            StrSql &= " AND I1.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        StrSql &= " AND rr.RELSHIP like '" & sm.UserInfo.RelShip & "%'" & vbCrLf

        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            StrSql &= " AND rr.RID='" & RIDValue.Value & "'" & vbCrLf
        End If
        Select Case sm.UserInfo.LID
            Case 0
                Dim s_DistID As String = TIMS.Get_DistID_RID(If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID), objConn)
                StrSql &= " AND I1.DistID='" & s_DistID & "'" & vbCrLf
            Case 1, 2
                StrSql &= " AND I1.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End Select
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist <> "" Then StrSql &= " and P1.PlanYear='" & v_yearlist & "'" & vbCrLf


        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If jobValue.Value <> "" Then
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
            If trainValue.Value <> "" Then StrSql &= " and P1.TMID='" & trainValue.Value & "'" & vbCrLf
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
        'StrSql &= " and C1.OCID IS NOT NULL "  '20081030 andy edit
        'StrSql &= " and P1.AppliedResult ='Y'  and P1.TransFlag='Y' " & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(StrSql, objConn)

        DataGridTable.Visible = False
        msg.Visible = True
        Me.msg.Text = "查無資料"

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        'If dt.Rows.Count > 0 Then End If

        DataGridTable.Visible = True
        msg.Visible = False
        Me.msg.Text = ""

        If ViewState("sort") = "" Then ViewState("sort") = "STDate,ClassName"
        PageControler1.Sort = ViewState("sort")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dtPlan)

        dtPlan.CurrentPageIndex = 0
        If CyclType.Text <> "" Then
            If Not IsNumeric(CyclType.Text) Then
                Common.MessageBox(Me, "期別需輸入數字型態!!")
                Exit Sub
            End If
        End If

        Call LoadData()

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

    Public Sub dtPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        Select Case e.CommandName
            Case "update" '依班級 作息時間設定
                Call GetRestTime()
                Dim URL_1 As String = String.Concat("SD_04_007_add.aspx?ID=", TIMS.Get_MRqID(Me), "&", e.CommandArgument, "")
                TIMS.Utl_Redirect1(Me, URL_1)
                ' TIMS.Utl_Redirect1(Me, "SD_04_007_add.aspx?ID=" & Request("ID") & " &" & e.CommandArgument & "&RIDValue=" & RIDValue.Value & "&Org=" & Org.Value & "")
        End Select

    End Sub

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        'If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
        '    If CInt(Me.TxtPageSize.Text) >= 1 Then
        '        Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
        '    Else
        '        Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '        Me.TxtPageSize.Text = 10
        '    End If
        'Else
        '    Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
        '    Me.TxtPageSize.Text = 10
        'End If
        'Me.dtPlan.PageSize = Me.TxtPageSize.Text

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but1 As Button = e.Item.FindControl("but1")
                e.Item.Cells(Cst_index).Text = TIMS.Get_DGSeqNo(sender, e) '序號 

                e.Item.Cells(Cst_OrgName).Text = Convert.ToString(drv("OrgName2"))
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
                        Select Case Convert.ToString(drv("AppliedResult"))
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

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "DistID", Convert.ToString(drv("DistID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "OrgID", Convert.ToString(drv("OrgID")))
                TIMS.SetMyValue(sCmdArg, "RID", Convert.ToString(drv("RID")))
                TIMS.SetMyValue(sCmdArg, "Year", yearlist.SelectedItem.Text)
                but1.CommandArgument = sCmdArg

            Case ListItemType.Header
                If Me.ViewState("sort") <> "" Then
                    Dim mylabel As String = ""
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Dim fgSortUpDown As Boolean = False
                    Select Case Convert.ToString(Me.ViewState("sort"))
                        Case "ClassName", "ClassName DESC"
                            mylabel = "ComName"
                            i = 8
                            fgSortUpDown = (ViewState("sort") = "ClassName")
                        Case "AppliedDate", "AppliedDate DESC"
                            mylabel = "ComName"
                            i = 2
                            fgSortUpDown = (ViewState("sort") = "AppliedDate")
                        Case "STDate", "STDate DESC"
                            mylabel = "ComName"
                            i = 4
                            fgSortUpDown = (ViewState("sort") = "STDate")
                        Case "FDDate", "FDDate DESC"
                            mylabel = "ComName"
                            i = 5
                            fgSortUpDown = (ViewState("sort") = "FDDate")
                        Case "OrgName", "OrgName DESC"
                            mylabel = "ComName"
                            i = 7
                            fgSortUpDown = (ViewState("sort") = "OrgName")
                    End Select
                    If i <> -1 AndAlso mylabel <> "" Then
                        mysort.ImageUrl = String.Concat("../../", If(fgSortUpDown, "SortUp.gif", "SortDown.gif"))
                        e.Item.Cells(i).Controls.Add(mysort)
                    End If
                End If
        End Select

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

    'Session("RestTime") 設定
    Sub GetRestTime()
        Dim strRestTime As String = ""
        TIMS.SetMyValue(strRestTime, "center", center.Text)
        TIMS.SetMyValue(strRestTime, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(strRestTime, "ClassName", ClassName.Text)
        TIMS.SetMyValue(strRestTime, "txtCJOB_NAME", txtCJOB_NAME.Text)
        TIMS.SetMyValue(strRestTime, "cjobValue", cjobValue.Value)
        TIMS.SetMyValue(strRestTime, "yearlist", yearlist.SelectedValue)
        TIMS.SetMyValue(strRestTime, "TB_career_id", TB_career_id.Text)
        TIMS.SetMyValue(strRestTime, "trainValue", trainValue.Value)
        TIMS.SetMyValue(strRestTime, "IsApprPaper", IsApprPaper.Text)
        TIMS.SetMyValue(strRestTime, "audit", audit.SelectedValue)
        TIMS.SetMyValue(strRestTime, "Button1", dtPlan.Visible)
        TIMS.SetMyValue(strRestTime, "PageIndex", dtPlan.CurrentPageIndex + 1)
        Session("RestTime") = strRestTime
    End Sub

    '機構作息時間設定
    Private Sub Btn_OrgSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_OrgSet.Click
        'Response.Redirect("SD_04_007_add.aspx?DistID=" & DistID & "&ID=" & Request("ID") & "&RID=" & RIDValue.Value & "&Org=" & Org.Value & "&OCID=''" & "&Year=" & yearlist.SelectedItem.Text & "")
        TIMS.Utl_Redirect1(Me, "SD_04_007_add.aspx?ID=" & Request("ID") & "&DistID=" & sm.UserInfo.DistID & "&RID=" & RIDValue.Value & "&OrgID=" & Org.Value & "&Year=" & yearlist.SelectedItem.Text & "")
    End Sub

End Class
