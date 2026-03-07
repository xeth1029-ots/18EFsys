Partial Class SD_05_016
    Inherits AuthBasePage

    Const cst_TipMsg1 As String = "已排除 預算別再出發!!"
    Const cst_TipMsg2 As String = "※經費審核通過總額，已排除 預算別再出發"

    Const cst_StatusMsg1_Y As String = "審核通過"
    Const cst_StatusMsg1_N As String = "審核不通過"
    Const cst_StatusMsg1_R As String = "退件修正"
    Const cst_StatusMsg1_S As String = "審核中"
    Const cst_StatusMsg1_NOINFO As String = "無資訊"
    Const cst_AppliedStatusM_NOINFO As String = "NOINFO"

    Const cst_StatusMsg2_1 As String = "已撥款"
    Const cst_StatusMsg2_Y As String = "待撥款" '"撥款中"
    Const cst_StatusMsg2_N As String = "不撥款"
    Const cst_StatusMsg2_R As String = "未撥款"
    Const cst_StatusMsg2_X As String = "不予補助"
    Const cst_StatusMsg2_NOINFO As String = "無資訊"

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
        PageControler1.PageDataGrid = DataGrid1

        Lab_TipMsg2.Text = cst_TipMsg2

        If Not IsPostBack Then
            Call cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()
        Button1.Attributes("onclick") = "return CheckSearch();"
        Button4.Attributes("onclick") = "ClearData();"
        msg.Text = ""
        msg2.Text = ""
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Page1.Visible = True
        Page2.Visible = False
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sSearch1()
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Name.Text = TIMS.ClearSQM(Name.Text)
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        If Name.Text <> "" Then RstMemo &= String.Concat("&Name=", Name.Text)
        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)

        Return RstMemo
    End Function

    Sub sSearch1()
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Name.Text = TIMS.ClearSQM(Name.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)

        Dim flagCanSch1 As Boolean = False
        If Not flagCanSch1 AndAlso OCIDValue1.Value <> "" Then flagCanSch1 = True
        If Not flagCanSch1 AndAlso IDNO.Text <> "" Then flagCanSch1 = True
        If Not flagCanSch1 AndAlso Name.Text <> "" Then flagCanSch1 = True
        If Not flagCanSch1 Then
            Common.MessageBox(Me, "請輸入有效查詢條件!!")
            Exit Sub
        End If

        If IDNO.Text <> "" Then
            TIMS.GetTrainingList2(sm, objconn, IDNO.Text)
        End If
        If OCIDValue1.Value <> "" Then
            TIMS.GetTrainingList2OCID(sm, objconn, OCIDValue1.Value)
        End If

        '(取得單一學員資料)
        Dim sql As String = ""
        sql &= " SELECT DISTINCT SID ,IDNO ,Name ,Birthday "
        sql &= " FROM dbo.VIEW_STUDENTBASICDATA"
        sql &= " WHERE 1=1 " '& SearchStr
        '排除離退訓學員輸入資料 by AMU 20090916
        sql &= " AND STUDSTATUS NOT IN (2,3) " & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND OCID = '" & OCIDValue1.Value & "' "
        If IDNO.Text <> "" Then sql &= " AND IDNO = '" & IDNO.Text & "' " '身分證字號
        If Name.Text <> "" Then sql &= " AND Name LIKE '%" & Name.Text & "%' " '學員姓名

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        'If dt.Rows.Count > 0 AndAlso IDNO.Text <> "" Then
        If IDNO.Text <> "" Then
            Dim dt2 As DataTable = TIMS.GetTrainingList2(sm, objconn, IDNO.Text)
            If dt2.Rows.Count > 0 Then
                For Each dr2 As DataRow In dt2.Rows
                    '資料身分證號未出現
                    If dt.Select("IDNO='" & dr2("IDNO") & "'").Length = 0 Then
                        Dim dr1 As DataRow = dt.NewRow() '新增一筆
                        dr1("SID") = Convert.DBNull
                        dr1("IDNO") = dr2("IDNO")
                        dr1("Name") = CStr(dr2("STDNAME")) & "-" & dr2("PLANNAME")
                        dr1("Birthday") = dr2("BIRTHDAY")
                        dt.Rows.Add(dr1)
                    End If
                Next
                dt.AcceptChanges()
            End If
        End If

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO,BIRTHDAY")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            DataGridTable.Visible = True
            'PageControler1.SqlPrimaryKeyDataCreate(sql, "SID")
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SID"
            PageControler1.ControlerLoad()
        End If
    End Sub

    '顯示學員詳細資料
    Sub show_DG2(ByVal aIDNO As String)
        'https://jira.turbotech.com.tw/browse/TIMSC-57
        '設定訓練單位層級無法觀看訓練機構欄位，使單位觀看查詢結果時無法觀看學員先前參加課程的訓練單位名稱，分署以上層級則維持原呈現方式。
        Dim flag_NoOrgName As Boolean = False '
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0" '署
            Case "1" '分署
            Case Else '其他委訓單位
                flag_NoOrgName = True
        End Select

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= If(flag_NoOrgName, " SELECT ' ' OrgName", " SELECT rr.OrgName") & vbCrLf '機構名稱
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSCNAME " & vbCrLf '班名
        sql &= " ,format(c.STDate,'yyyy/MM/dd') STDate" & vbCrLf '開訓日
        sql &= " ,format(c.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf '結訓日
        sql &= " ,format(e.MODIFYDATE,'yyyy/MM/dd') MODIFYDATE" & vbCrLf '異動日
        sql &= " ,e.SumOfMoney " & vbCrLf
        sql &= " ,e.AppliedStatus " & vbCrLf
        sql &= " ,e.AppliedStatusM " & vbCrLf
        sql &= " ,e.BUDID " & vbCrLf
        sql &= " ,bb.BUDNAME " & vbCrLf '預算別
        sql &= " ,b.SOCID " & vbCrLf 'SOCID
        sql &= " ,b.StudStatus SDSTATUS " & vbCrLf 'StudStatus->SdStatus
        sql &= " ,dbo.DECODE12(b.StudStatus,1,'在訓',2,'離訓',3,'退訓',4,'續訓',5,'結訓','在訓') StudStatus " & vbCrLf 'StudStatus
        sql &= " FROM STUD_STUDENTINFO a " & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.SID = b.SID " & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID = c.OCID " & vbCrLf
        sql &= " JOIN PLAN_PLANINFO d  ON c.PlanID = d.PlanID AND c.ComIDNO = d.ComIDNO AND c.SeqNO = d.SeqNO " & vbCrLf
        sql &= " JOIN VIEW_RIDNAME rr ON rr.RID = c.RID " & vbCrLf
        sql &= " JOIN VIEW_PLAN ip ON ip.planid = c.planid "
        'Cst_TPlanID28_1 '要含充電起飛的包班資料
        'sql &= " AND d.TPlanID IN (" & TIMS.Cst_TPlanID28_1d & ") " & vbCrLf
        sql &= " AND d.TPlanID IN (" & TIMS.Cst_TPlanID28_1a & ") " & vbCrLf
        'sql += " AND d.TPlanID = 28 AND d.AppliedResult = 'Y' AND d.DefStdCost > 0 " & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_QUESTIONFIN qq ON qq.socid = b.socid " & vbCrLf
        '學員經費 已撥款狀態 e.AppliedStatus=1
        sql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST e ON b.SOCID = e.SOCID " & vbCrLf 'AND e.AppliedStatus=1
        sql &= " LEFT JOIN dbo.VIEW_BUDGET bb ON bb.budid = e.budid " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND d.AppliedResult = 'Y' " & vbCrLf
        sql &= " AND c.NotOpen = 'N' " & vbCrLf
        ''sql += "   AND b.StudStatus NOT IN (2,3) " & vbCrLf '排除離退。
        'sql &= " AND (1!=1 "
        'sql &= "    OR ip.TPlanID IN (" & TIMS.Cst_TPlanID28_1a & ") " & vbCrLf
        'sql &= "    OR (ip.TPlanID IN (" & TIMS.Cst_TPlanID28_1b & ") " & vbCrLf
        'sql &= "    AND b.WorkSuppIdent = 'Y' " & vbCrLf
        'sql &= "    AND c.STDate >= CONVERT(DATE, '2010/04/01')) " & vbCrLf
        'sql &= " ) " & vbCrLf
        'sql &= " AND a.IDNO = '" & aIDNO & "' " & vbCrLf
        sql &= " AND ip.TPlanID IN (" & TIMS.Cst_TPlanID28_1a & ")" & vbCrLf
        sql &= " AND a.IDNO = '" & aIDNO & "' " & vbCrLf
        sql &= " ORDER BY d.PlanYear ,c.STDate ASC " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        'aIDNO
        'If IDNO.Text <> "" Then
        'End If
        Dim flag_need_sort_1 As Boolean = False '須要重新排序嗎 false:不用
        Dim dt2 As DataTable = TIMS.GetTrainingList2c(objconn, aIDNO)
        Dim ss3 As String = ""
        Dim ff3 As String = ""

        ff3 = "TPlanID IN (" & TIMS.Cst_TPlanID28_1b1b & ")"
        If dt2.Select(ff3).Length > 0 Then
            For Each dr2 As DataRow In dt2.Rows
                Dim dr1 As DataRow = dt.NewRow()
                dr1("OrgName") = If(flag_NoOrgName, " ", dr2("OrgName")) '機構名稱
                dr1("ClassCName") = String.Concat(dr2("ClassCName"), "-", dr2("PLANNAME"))
                dr1("STDate") = TIMS.Cdate3(dr2("STDate"))
                dr1("FTDate") = TIMS.Cdate3(dr2("FTDate"))
                dr1("SumOfMoney") = dr2("SumOfMoney")
                dr1("AppliedStatus") = dr2("AppliedStatus")
                '強制改為審核通過 
                'dr1("AppliedStatusM") = dr2("AppliedStatusM") 'TIMS.cst_YES 'dr2("AppliedStatusM")
                dr1("AppliedStatusM") = dr2("AppliedStatusM") 'TIMS.cst_YES 'dr2("AppliedStatusM")
                dr1("BUDID") = dr2("BUDID")
                dr1("BUDNAME") = TIMS.GET_BudgetName(Convert.ToString(dr2("BUDID")), objconn)
                dr1("SOCID") = Convert.DBNull
                dr1("SDSTATUS") = dr2("STUDSTATUS")
                dr1("StudStatus") = dr2("TFLAG")
                dt.Rows.Add(dr1)
            Next
            flag_need_sort_1 = True '須要重新排序嗎 true:必須
            dt.AcceptChanges()
        End If

        If flag_need_sort_1 Then
            ff3 = ""
            ss3 = "STDate"
            dt = TIMS.CopyDt(dt, ff3, ss3)
        End If

        '產投 政府補助經費 --產業人才投資方案(三年補助)
        RemainSub.Text = TIMS.Get_3Y_SupplyMoney()

#Region "(No Use)"

        '970508 Andy  學員補助金 依登入年度變更  
        '---------------------------------------
        '2007年前(補助金為2萬)
        '2007年(補助金為為3年3萬)
        '2008年(補助金為為3年5萬)
        '----------------------------------------
        'If sm.UserInfo.Years < "2007" Then
        '    RemainSub.Text = 20000
        'Else
        '    If sm.UserInfo.Years = "2007" Then
        '        RemainSub.Text = 30000
        '    Else
        '        If sm.UserInfo.Years >= "2008" Then RemainSub.Text = 50000
        '    End If
        'End If

#End Region

        Me.LabTotal.Text = RemainSub.Text
        Me.LabTotal.ToolTip = TIMS.gTip_LabTotalSupplyMoney
        'Me.LabTotal.ToolTip = "2007年前，補助金為2萬" & vbCrLf
        'Me.LabTotal.ToolTip += "2007年，補助金為為3年3萬" & vbCrLf
        'Me.LabTotal.ToolTip += "2008年，補助金為為3年5萬" & vbCrLf
        Me.LabSumOfMoney.Text = 0
        'For Each dr In dt.Select("STDate>='" & FormatDateTime(Now.Date.AddYears(-3), 2) & "' and AppliedStatus=1")
        '    Me.LabSumOfMoney.Text += Int(dr("SumOfMoney"))
        '    RemainSub.Text = Int(RemainSub.Text) - dr("SumOfMoney")
        'Next
        Dim STDate As String = TIMS.Cdate3(Now)
        If dt.Rows.Count > 0 Then STDate = CDate(dt.Rows(dt.Rows.Count - 1)("STDate")).ToString("yyyy/MM/dd")

        '每3年所使用的補助金 目前已使用多少政府補助。
        '含職前webservice
        Me.LabSumOfMoney.Text += TIMS.Get_SubsidyCost(aIDNO, STDate, "", "Y", objconn)
        TIMS.Tooltip(Me.LabSumOfMoney, cst_TipMsg1)
        RemainSub.Text = Int(RemainSub.Text) - CInt(Me.LabSumOfMoney.Text)

        Dim sDate As String = String.Empty
        Dim eDate As String = String.Empty
        'Dim aIDNO As String = e.CommandArgument
        Call TIMS.Get_SubSidyCostDay(aIDNO, STDate, sDate, eDate, objconn)
        TIMS.Tooltip(RemainSub, "計算開始日為：" & sDate & "~" & eDate, True)
        LabCostDay.Text = "補助金補助期間：" & sDate & "~" & eDate

        RemainSub.ForeColor = Color.Black
        If Int(RemainSub.Text) < 0 Then RemainSub.ForeColor = Color.Red

        DataGrid2.Visible = False
        msg2.Text = "查無資料"

        If dt.Rows.Count > 0 Then
            DataGrid2.Visible = True
            msg2.Text = ""
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
        End If
    End Sub

    '計算。
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "view1"
                Dim aIDNO As String = e.CommandArgument
                aIDNO = TIMS.ClearSQM(aIDNO)
                If aIDNO = "" Then Exit Sub

                'Dim dt As DataTable = Nothing
                Dim dr As DataRow = Nothing
                Dim sql As String = ""
                sql = "SELECT IDNO,NAME FROM dbo.STUD_STUDENTINFO WHERE IDNO='" & aIDNO & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr Is Nothing Then Exit Sub
                LIDNO.Text = Convert.ToString(dr("IDNO")) '.ToString
                LName.Text = Convert.ToString(dr("Name")) 'dr("Name").ToString
                LName.Text = TIMS.HtmlDecode1(LName.Text)

                Page1.Visible = False
                Page2.Visible = True

                '顯示學員詳細資料
                Call show_DG2(Convert.ToString(dr("IDNO")))
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim Button3 As Button = e.Item.FindControl("Button3")
                Button3.CommandArgument = Convert.ToString(drv("IDNO")) '.ToString
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        'Const Cst_ClassCName As Integer = 1
        'Const Cst_AppliedStatusM As Integer = 6
        'Const Cst_AppliedStatus As Integer = 7
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labClassCName As Label = e.Item.FindControl("labClassCName")
                Dim labSFMTDATE As Label = e.Item.FindControl("labSFMTDATE")
                Dim labAppliedStatusM As Label = e.Item.FindControl("labAppliedStatusM")
                Dim labAppliedStatus As Label = e.Item.FindControl("labAppliedStatus")

                labClassCName.Text = Convert.ToString(drv("ClassCName"))
                labSFMTDATE.Text = String.Concat(drv("STDATE"), "~", drv("FTDATE"), " (", drv("MODIFYDATE"), ")")
                If Convert.ToString(drv("SOCID")) = "" Then labClassCName.ForeColor = Color.Red
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                '申請補助金額 預算別 審核狀態 撥款狀態 訓練狀態 
                '審核狀態
                Dim StatusMMsg1 As String = ""
                Select Case Convert.ToString(drv("AppliedStatusM"))
                    Case cst_AppliedStatusM_NOINFO
                        StatusMMsg1 = cst_StatusMsg1_NOINFO '"審核通過" '"申請成功"
                    Case "Y"
                        StatusMMsg1 = cst_StatusMsg1_Y '"審核通過" '"申請成功"
                    Case "N"
                        StatusMMsg1 = cst_StatusMsg1_N '"審核不通過" '"申請失敗"
                    Case "R"
                        StatusMMsg1 = cst_StatusMsg1_R '"退件修正"
                    Case Else
                        'e.SumOfMoney
                        If Convert.ToString(drv("SumOfMoney")) <> "" AndAlso Val(drv("SumOfMoney")) > 0 Then
                            StatusMMsg1 = cst_StatusMsg1_S '"審核中" '"未審核"
                        End If
                        'If Convert.ToString(drv("SumOfMoney")) <> "" Then StatusMMsg1 = cst_StatusMsg1_S '"審核中" '"未審核"
                End Select
                labAppliedStatusM.Text = StatusMMsg1

                '撥款狀態
                If Convert.ToString(drv("AppliedStatus")) = "1" Then
                    'e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_1 '"已撥款" '"申請成功"
                    labAppliedStatus.Text = cst_StatusMsg2_1 '"已撥款" '"申請成功"
                Else
                    Dim StatusMsg2 As String = ""
                    Select Case Convert.ToString(drv("AppliedStatusM"))
                        Case cst_AppliedStatusM_NOINFO
                            StatusMsg2 = cst_StatusMsg2_NOINFO '"審核通過" '"申請成功"
                        Case "Y" '審核通過
                            StatusMsg2 = cst_StatusMsg2_Y '"撥款中" '"申請中"
                        Case "N" '審核不通過
                            StatusMsg2 = cst_StatusMsg2_N '"不撥款" '"申請中"
                        Case "R" '退件修正
                            StatusMsg2 = cst_StatusMsg2_R '"未撥款" '"申請失敗"
                        Case Else '審核中
                            'StatusMsg2 = cst_StatusMsg2_X 'cst_StatusMsg2_R
                            'e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_R '"未撥款" '"申請失敗"
                            Select Case Convert.ToString(drv("SdStatus"))
                                Case "2", "3"
                                    StatusMsg2 = cst_StatusMsg2_X
                                    'e.Item.Cells(Cst_AppliedStatus).Text = cst_StatusMsg2_X '"不予補助"
                            End Select
                            'Case cst_AppliedStatusM_Y2 '審核通過
                            '    StatusMsg2 = cst_StatusMsg2_Y2 '"撥款中" '"申請中"
                    End Select
                    'e.Item.Cells(Cst_AppliedStatus).Text = StatusMsg2
                    labAppliedStatus.Text = StatusMsg2
                End If
        End Select
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Page1.Visible = True
        Page2.Visible = False
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

#Region "(No Use)"

    'select MIN(g.STDate) adate1
    ',DATEADD(month, 12*3, MIN(g.STDate))-1 adate2
    ',DATEADD(month, 12*3, MIN(g.STDate))  adate3
    'from (
    '  SELECT c.STDate
    '  FROM Stud_StudentInfo a  
    '  JOIN Class_StudentsOfClass b  ON a.SID = b.SID AND a.IDNO = 'B220850788'
    '  JOIN Class_Classinfo c  ON b.OCID = c.OCID and c.NotOpen = 'N'
    '  JOIN Plan_Planinfo d  ON c.PlanID = d.PlanID AND c.ComIDNO =d.ComIDNO AND c.SeqNO =d.SeqNO
    '  AND d.TPlanID in ('28','54','58')
    '  --AND c.STDate >= to_date('2014-04-18','yyyy-MM-dd')
    '  UNION
    '  SELECT c.STDate
    '  FROM Stud_StudentInfo a  
    '  JOIN Class_StudentsOfClass b  ON a.SID = b.SID AND a.IDNO = 'B220850788'
    '  JOIN Class_Classinfo c  ON b.OCID = c.OCID and c.NotOpen = 'N'
    '  JOIN Plan_Planinfo d  ON c.PlanID = d.PlanID AND c.ComIDNO =d.ComIDNO AND c.SeqNO =d.SeqNO
    '  /*  從2011/01/01開始算 且為在職者  and c.STDate>='2011/01/01' and b.WorkSuppIdent='Y'*/
    '  AND c.STDate>= convert(datetime, '2010/04/01', 111) and b.WorkSuppIdent='Y'
    '  AND d.TPlanID in ('46','47') 
    '  --AND c.STDate >= to_date('2014-04-18','yyyy-MM-dd')
    ') g

#End Region

End Class
