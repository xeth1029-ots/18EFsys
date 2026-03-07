Partial Class SD_01_004_email
    Inherits AuthBasePage

    Const Cst_index As Integer = 0
    Const Cst_PlanYear As Integer = 1      '計畫年度
    Const Cst_AppliedDate As Integer = 2   '申請日期
    Const Cst_STDate As Integer = 4        '訓練起日
    Const Cst_FDDate As Integer = 5        '訓練迄日
    Const Cst_OrgName As Integer = 6       '機構名稱('管控單位)
    'Const Cst_ClassName As Integer = 8     '班級名稱
    Const Cst_AppliedResult As Integer = 9 '班級審核狀態
    Const Cst_TransFlag As Integer = 11    '是否轉班

    Dim Gv_ddlyears As String = ""
    Dim Gt_ddlyears As String = ""
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    '是否啟用西元年轉民國年機制
    Dim flag_Roc As Boolean = False

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
        tr_audit1.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.labtmid.Text = "訓練業別"
            'audit1.Visible = False
            'IsApprPaper.AutoPostBack = False
        End If

        '分頁設定 Start
        PageControler1.PageDataGrid = dtPlan
        '分頁設定 End

        '20080818 andy
        '--------------------
        IsApprPaper.SelectedValue = "Y"
        audit.SelectedValue = "Y"
        '--------------------

        Gv_ddlyears = TIMS.GetListValue(ddlyears)
        Gt_ddlyears = TIMS.GetListText(ddlyears)

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?selected_year={1}');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"), Gv_ddlyears)

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        Dim PlanKind As String
        PlanKind = TIMS.Get_PlanKind(Me, objconn) '依sm.UserInfo.PlanID取得PlanKind
        If PlanKind = "1" Then
            dtPlan.Columns(6).Visible = False
        Else
            dtPlan.Columns(6).Visible = True
        End If

        'btnQuery.Enabled = False
        'If au.blnCanSech Then btnQuery.Enabled = True

        flag_Roc = TIMS.CHK_REPLACE2ROC_YEARS()

        If Not Me.IsPostBack Then
            Me.msg.Text = ""
            DataGridTable.Visible = False
            ddlyears = TIMS.GetSyear(ddlyears)
            ' 年度帶預設值
            Common.SetListItem(ddlyears, sm.UserInfo.Years)

            ''(加強操作便利性)
            'RIDValue.Value = RID
            'Dim sqlstring As String = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & RID & "'"
            'Dim orgname As String = DbAccess.ExecuteScalar(sqlstring, objConn)
            'center.Text = orgname

            ''取得訓練計畫
            'Sqlstr = "select TPlanID  from ID_Plan where PlanID=" & sm.UserInfo.PlanID & ""
            'TPlanid.Value = DbAccess.ExecuteScalar(Sqlstr, objConn)

            '取得訓練機構
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            '取得訓練計畫
            TPlanid.Value = sm.UserInfo.TPlanID
        End If

        '帶入查詢參數
        GetKeepSearch(sender, e)
    End Sub

    Sub LoadData()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        '取出訓練計畫名稱
        TPlanName.Text = TIMS.GetPlanName(sm.UserInfo.PlanID, objconn)

        Dim s_PlanKind As String = TIMS.Get_PlanKind(Me, objconn) '依sm.UserInfo.PlanID取得PlanKind

        Dim parms As New Hashtable() 'parms.Clear()

        '建立訓練的DATATABLE
        Dim StrSql As String = ""
        StrSql &= " SELECT I1.DistID, O1.Orgid, P1.PlanID, O1.OrgName" & vbCrLf
        StrSql &= " ,P1.STDate, P1.FDDate " & vbCrLf
        StrSql &= " ,P1.ComIDNO, P1.PlanYear, P1.TMID, P1.AppliedResult " & vbCrLf
        StrSql &= " ,P1.SeqNO, K1.PlanName " & vbCrLf
        StrSql &= " ,CASE WHEN K2.JobID IS NULL THEN K2.TrainName ELSE K2.JobName END TrainName " & vbCrLf
        StrSql &= " ,P1.ClassName, P1.AdmPercent " & vbCrLf
        StrSql &= " ,P1.CyclType, P1.AppliedDate" & vbCrLf
        StrSql &= " ,P2.VerReason, P1.TransFlag, P1.RID " & vbCrLf
        StrSql &= " ,C1.OCID " & vbCrLf
        StrSql &= " ,dbo.FN_GET_CLASSCNAME(P1.ClassName,P1.CyclType) CLASSNAME2" & vbCrLf
        StrSql &= " ,rr.RELSHIP " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql &= "    ,P3.FirResult, P3.SecResult, P3.IsApprPaper " & vbCrLf
        Else
            StrSql &= "    ,'X' AS FirResult, 'X' AS SecResult, NULL AS IsApprPaper " & vbCrLf '給個假資料
        End If
        StrSql &= " FROM PLAN_PLANINFO P1 WITH(NOLOCK)" & vbCrLf
        StrSql &= " JOIN Key_Plan K1 WITH(NOLOCK) ON P1.TPlanID = K1.TPlanID" & vbCrLf
        StrSql &= " JOIN Key_TrainType K2 WITH(NOLOCK) ON P1.TMID = K2.TMID" & vbCrLf
        StrSql &= " JOIN ID_Plan I1 WITH(NOLOCK) ON P1.PlanID = I1.PlanID" & vbCrLf
        StrSql &= " JOIN Org_OrgInfo O1 WITH(NOLOCK) ON P1.ComIDNO = O1.ComIDNO" & vbCrLf
        StrSql &= " JOIN AUTH_RELSHIP rr WITH(NOLOCK) ON rr.RID = P1.RID" & vbCrLf
        StrSql &= " LEFT JOIN Plan_VerRecord P2 WITH(NOLOCK) ON P1.PlanID = P2.PlanID and P1.ComIDNO = P2.ComIDNO and P1.SeqNo = P2.SeqNo" & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            StrSql &= " LEFT JOIN Plan_VerReport P3 WITH(NOLOCK) ON P1.PlanID = P3.PlanID AND P1.ComIDNO = P3.ComIDNO AND P1.SeqNo = P3.SeqNo" & vbCrLf
        End If
        StrSql &= " LEFT JOIN CLASS_CLASSINFO C1 WITH(NOLOCK) ON P1.PlanID = C1.PlanID AND P1.ComIDNO = C1.ComIDNO AND P1.SeqNo = C1.SeqNo" & vbCrLf
        StrSql &= " WHERE 1=1" & vbCrLf
        StrSql &= " AND C1.IsSuccess='Y'" & vbCrLf
        StrSql &= " AND P1.IsApprPaper = 'Y' " & vbCrLf
        'StrSql &= " AND C1.IsSuccess = 'Y'" & vbCrLf '/IsSuccess='Y'
        'StrSql += " and C1.OCID IS NOT NULL "  '20081030 andy edit
        StrSql &= " AND P1.AppliedResult = 'Y'" & vbCrLf  '產業人才投資方案類
        StrSql &= " AND P1.TransFlag = 'Y'" & vbCrLf

        StrSql &= " AND P1.TPlanID=@TPlanID2 " & vbCrLf
        If TPlanid.Value = "" Then TPlanid.Value = sm.UserInfo.TPlanID
        parms.Add("TPlanID2", TPlanid.Value)

        If sm.UserInfo.RID = "A" Then
            StrSql &= " and I1.TPlanID=@TPlanID" & vbCrLf
            StrSql &= " and I1.Years=@Years" & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("Years", Convert.ToString(sm.UserInfo.Years))
        Else
            StrSql &= " and I1.TPlanID=@TPlanID" & vbCrLf
            StrSql &= " and I1.PlanID=@PlanID" & vbCrLf
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        '如果不是產學訓計畫 penny 2007/10/17
        'If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '審核狀態有選值
        '    StrSql &= " AND P1.AppliedResult IN ('Y','N') " & vbCrLf
        'End If
        If sm.UserInfo.RID = "A" Then
            StrSql &= " AND rr.RELSHIP LIKE @RELSHIP " & vbCrLf
            parms.Add("RELSHIP", relship & "%")
        Else
            If s_PlanKind = "2" Then
                StrSql &= " AND rr.RELSHIP LIKE @RELSHIP " & vbCrLf
                parms.Add("RELSHIP", relship & "%")
            End If
        End If

        Select Case sm.UserInfo.LID
            Case 0
            Case Else '1/2
                StrSql &= " AND I1.DistID=@DistID " & vbCrLf
                parms.Add("DistID", sm.UserInfo.DistID)
        End Select

        Gv_ddlyears = TIMS.GetListValue(ddlyears)
        If Gv_ddlyears <> "" Then
            StrSql &= " AND P1.PlanYear=@PlanYear " & vbCrLf
            parms.Add("PlanYear", Gv_ddlyears)
        End If

        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If jobValue.Value <> "" Then
                StrSql &= " AND ( " & vbCrLf
                StrSql &= " P1.TMID=@TMID " & vbCrLf
                StrSql &= " OR P1.TMID IN ( " & vbCrLf
                StrSql &= " SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '職類別
                StrSql &= " SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '業別
                StrSql &= " SELECT TMID FROM Key_TrainType WHERE busid = 'G') " & vbCrLf '產業人才投資方案類
                StrSql &= " AND TMID=@TMID " & vbCrLf
                StrSql &= " ))) " & vbCrLf
                parms.Add("TMID", jobValue.Value)
            End If
        Else
            If trainValue.Value <> "" Then
                StrSql = StrSql & " AND P1.TMID=@TMID " & vbCrLf
                parms.Add("TMID", trainValue.Value)
            End If
        End If

        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            StrSql &= " AND P1.CJOB_UNKEY=@CJOB_UNKEY " & vbCrLf  '通俗職類
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If

        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then
            StrSql &= " AND P1.ClassName LIKE '%'+@ClassName+'%' " & vbCrLf
            parms.Add("ClassName", ClassName.Text)
        End If

        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                StrSql &= " AND P1.CyclType = @CyclType " & vbCrLf
                parms.Add("CyclType", If(Int(CyclType.Text) < 10, "0" & Int(CyclType.Text), CyclType.Text))
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(StrSql, objconn, parms)
        DataGridTable.Visible = False
        Me.msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            Me.msg.Text = ""
            If ViewState("sort") = "" Then ViewState("sort") = "STDate,CLASSNAME2"

            PageControler1.PageDataTable = dt '.SqlString = StrSql
            PageControler1.Sort = ViewState("sort")
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.dtPlan)

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

    Private Sub dtPlan_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dtPlan.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but1 As LinkButton = e.Item.FindControl("but1")
                e.Item.Cells(Cst_index).Text = TIMS.Get_DGSeqNo(sender, e)

                'Visible="false" 訓練職類
                '管控單位
                e.Item.Cells(Cst_OrgName).Text = ""
                Dim ParentName As String = TIMS.Get_ParentRID(drv("Relship"), objconn)
                If ParentName <> "" Then e.Item.Cells(Cst_OrgName).Text = ParentName

                'e.Item.Cells(Cst_ClassName).Text = drv("ClassName").ToString
                'If drv("CyclType").ToString <> "" Then
                '    If Int(drv("CyclType")) <> 0 Then e.Item.Cells(Cst_ClassName).Text += "第" & Int(drv("CyclType")) & "期"
                'End If
                Select Case Convert.ToString(drv("AppliedResult"))'.ToString
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

                'e.Item.Cells(13).Text = ""
                'If Convert.ToString(drv("DistID")) <> "" Then e.Item.Cells(13).Text = TIMS.Get_DistName1(drv("DistID"))
                If flag_Roc Then
                    e.Item.Cells(Cst_PlanYear).Text = (CInt(drv("planyear")) - 1911).ToString()
                    e.Item.Cells(Cst_AppliedDate).Text = TIMS.Cdate17(drv("applieddate"))
                    e.Item.Cells(Cst_STDate).Text = TIMS.Cdate17(drv("stdate"))
                    e.Item.Cells(Cst_FDDate).Text = TIMS.Cdate17(drv("fddate"))
                Else
                    e.Item.Cells(Cst_PlanYear).Text = drv("planyear").ToString()  'edit，by:20181018
                    e.Item.Cells(Cst_AppliedDate).Text = drv("applieddate")  'edit，by:20181018
                    e.Item.Cells(Cst_STDate).Text = drv("stdate")  'edit，by:20181018
                    e.Item.Cells(Cst_FDDate).Text = drv("fddate")  'edit，by:20181018
                End If

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
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
                but1.CommandArgument = "PlanID=" & drv("PlanID") & "&DistID=" & drv("DistID") & "&OCID=" & drv("OCID") & "&OrgID=" & drv("OrgID") & "&RID=" & drv("RID") & "&Year=" & Gt_ddlyears & ""

            Case ListItemType.Header
                If Me.ViewState("sort") <> "" Then
                    Dim mysort As New System.Web.UI.WebControls.Image
                    Dim i As Integer = -1
                    Select Case Me.ViewState("sort")
                        Case "CLASSNAME2", "CLASSNAME2 DESC"
                            i = 8
                            mysort.ImageUrl = If(Me.ViewState("sort") = "ClassName", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "AppliedDate", "AppliedDate DESC"
                            i = 2
                            mysort.ImageUrl = If(Me.ViewState("sort") = "AppliedDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "STDate", "STDate DESC"
                            i = 4
                            mysort.ImageUrl = If(Me.ViewState("sort") = "STDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "FDDate", "FDDate DESC"
                            i = 5
                            mysort.ImageUrl = If(Me.ViewState("sort") = "FDDate", "../../images/SortUp.gif", "../../images/SortDown.gif")
                        Case "OrgName", "OrgName DESC"
                            i = 7
                            mysort.ImageUrl = If(Me.ViewState("sort") = "OrgName", "../../images/SortUp.gif", "../../images/SortDown.gif")
                    End Select
                    If i <> -1 Then e.Item.Cells(i).Controls.Add(mysort)
                End If
        End Select
    End Sub

    Public Sub dtPlan_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dtPlan.ItemCommand
        Select Case e.CommandName
            Case "update"
                KeepSearch()
                Call DbAccess.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "SD_01_004_email_add.aspx?ID=" & Request("ID") & " &" & e.CommandArgument & "")
        End Select
    End Sub

    Private Sub dtPlan_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles dtPlan.SortCommand
        Me.ViewState("sort") = If(Me.ViewState("sort") <> e.SortExpression, e.SortExpression, e.SortExpression & " DESC")

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Sub KeepSearch()
        Session("_SearchStr") = "Class=SD_01_004_mail"
        Session("_SearchStr") += "&ddlyears=" & TIMS.ClearSQM(Gv_ddlyears)
        Session("_SearchStr") += "&center=" & TIMS.ClearSQM(center.Text)
        Session("_SearchStr") += "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        Session("_SearchStr") += "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        Session("_SearchStr") += "&jobValue=" & TIMS.ClearSQM(jobValue.Value)
        Session("_SearchStr") += "&txtCJOB_NAME=" & TIMS.ClearSQM(txtCJOB_NAME.Text)
        Session("_SearchStr") += "&cjobValue=" & TIMS.ClearSQM(cjobValue.Value)
        Session("_SearchStr") += "&ClassName=" & TIMS.ClearSQM(ClassName.Text)
        Session("_SearchStr") += "&TrainValue=" & TIMS.ClearSQM(trainValue.Value)
        Session("_SearchStr") += "&CyclType=" & TIMS.ClearSQM(CyclType.Text)
        Session("_SearchStr") += "&TxtPageSize=" & TIMS.ClearSQM(TxtPageSize.Text)
        Session("_SearchStr") += "&PageIndex=" & TIMS.ClearSQM(dtPlan.CurrentPageIndex + 1)
        Session("_SearchStr") += "&Button1=" & dtPlan.Visible

        If dtPlan.Visible Then
            Session("_SearchStr") += "&submit=1"
        Else
            Session("_SearchStr") += "&submit=0"
        End If
    End Sub

    Sub GetKeepSearch(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '執行 Session("_SearchStr") 保留查詢值
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
#Region "(No Use)"

                'If TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit") = "1" Then
                '    Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                '    If IsNumeric(Me.ViewState("PageIndex")) Then PageControler1.PageIndex = Me.ViewState("PageIndex")
                '    btnQuery_Click(sender, e)
                '    'PageControler1.PageIndex = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                '    'PageControler1.CreateData()
                'End If

#End Region
                Dim MyValue As String = ""
                Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
                MyValue = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit")
                If MyValue = "1" Then
                    btnQuery_Click(sender, e)
                    If IsNumeric(Me.ViewState("PageIndex")) Then
                        PageControler1.PageIndex = Me.ViewState("PageIndex")
                        PageControler1.CreateData()
                    End If
                End If
            End If
            Session("_SearchStr") = Nothing
        End If
    End Sub

    Private Sub Btn_OrgSet_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_orgset.Click
        KeepSearch()
        Call DbAccess.CloseDbConn(objconn)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Org.Value = TIMS.ClearSQM(Org.Value)
        Gt_ddlyears = TIMS.GetListText(ddlyears)

        TIMS.Utl_Redirect1(Me, "SD_01_004_email_add.aspx?DistID=" & sm.UserInfo.DistID & "&ID=" & Request("ID") & "&RID=" & RIDValue.Value & "&OrgID=" & Org.Value & "&Year=" & Gt_ddlyears & "")
    End Sub
End Class