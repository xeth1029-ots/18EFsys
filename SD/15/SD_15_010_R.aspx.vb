Partial Class SD_15_010
    Inherits AuthBasePage

    'SD_15_010_R
    'SD_15_010_R_1

#Region "NO USE"
    'Dim objConn As SqlConnection
    'Dim objreader As SqlDataReader
    'Dim objdataset As New DataSet
    'Dim SqlCmd As String
    'Dim dsA As SqlDataAdapter
    'Dim sqlAdapter As SqlDataAdapter
    'Dim sqlTable As DataTable
    'Dim FunDr As DataRow
    'Dim Sqlstr As String
    'Dim DistID As String
    'Dim PlanID As String
    'Dim UserID As String
    'objConn = DbAccess.GetConnection

    'DistID = sm.UserInfo.DistID
    ''PlanID = sm.UserInfo.PlanID
    'UserID = sm.UserInfo.UserID
#End Region

    Dim Auth_Relship As DataTable
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Dim Sqlstr As String = ""
        If sm.UserInfo.RID = "A" Then
            Sqlstr = "SELECT a.RID,b.OrgName FROM Auth_Relship a,Org_OrgInfo b WHERE a.OrgID=b.OrgID and (PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' and Years=(SELECT Years FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "')) or PlanID=0)"
        Else
            Sqlstr = "SELECT a.RID,b.OrgName FROM Auth_Relship a,Org_OrgInfo b WHERE a.OrgID=b.OrgID and (PlanID='" & sm.UserInfo.PlanID & "' or PlanID=0)"
        End If
        Auth_Relship = DbAccess.GetDataTable(Sqlstr, objConn)

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '   'Dim FunDr As DataRow = FunDrArray(0)
        '    If FunDr("Adds") = 1 Then
        '        If sm.UserInfo.Years = 2006 Then Button2.Visible = True Else Button2.Visible = False
        '    End If
        'End If
        If sm.UserInfo.Years = 2006 Then Button2.Visible = True Else Button2.Visible = False

        '不開放給委訓單位
        If sm.UserInfo.LID = "2" Then
            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        End If

        If Not Me.IsPostBack Then
            SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
            Common.SetListItem(SearchPlan, "G")

            DistrictList = TIMS.Get_DistID(DistrictList)
            '取得訓練計畫
            Sqlstr = "select TPlanID  from ID_Plan where PlanID=" & sm.UserInfo.PlanID & ""
            TPlanid.Value = DbAccess.ExecuteScalar(Sqlstr, objConn)
            '(加強操作便利性)2005/4/1-melody
            RIDValue.Value = sm.UserInfo.RID
            Dim sqlstring As String = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & sm.UserInfo.RID & "'"
            Dim orgname As String = DbAccess.ExecuteScalar(sqlstring, objConn)
            center.Text = orgname

            '取得訓練計畫
            'SqlCmd = "select * from Key_Plan order by TPlanID"
            'Me.planlist.Items.Clear()
            'DbAccess.MakeListItem(Me.planlist, SqlCmd)
            'Me.planlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

        End If

        trPlanKind.Style("display") = "none"
        trPackageType.Style("display") = "none"
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trPackageType.Style("display") = TIMS.cst_inline1 '"inline"
        Else
            '28:產業人才投資方案
            '計畫範圍 產投
            If sm.UserInfo.Years >= 2008 Then
                trPlanKind.Style("display") = TIMS.cst_inline1 '"inline"
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        '帶入查詢參數
        If Not IsPostBack Then
            AppStage = TIMS.Get_AppStage(AppStage)
            If Not Session("search") Is Nothing Then
                TB_career_id.Text = TIMS.GetMyValue(Session("search"), "TB_career_id")
                trainValue.Value = TIMS.GetMyValue(Session("search"), "trainValue")
                jobValue.Value = TIMS.GetMyValue(Session("search"), "jobValue")
                center.Text = TIMS.GetMyValue(Session("search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("search"), "RIDValue")
                UNIT_SDATE.Text = TIMS.GetMyValue(Session("search"), "UNIT_SDATE")
                UNIT_EDATE.Text = TIMS.GetMyValue(Session("search"), "UNIT_EDATE")
                Session("search") = Nothing
            End If
        End If

        If sm.UserInfo.DistID = "000" Then
            DistType.Enabled = True
            DistrictList.Enabled = True
        Else
            DistType.Enabled = False
            Common.SetListItem(DistType, "1")
            DistrictList.SelectedValue = sm.UserInfo.DistID
            DistrictList.Enabled = False
        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
        'DistType.Attributes("onclick") = "SetDistType('DistType');"
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '選擇轄區
        'Dim i As Integer
        Dim DistID1 As String
        Dim DistName As String

        '報表要用的轄區參數-begin
        DistID1 = ""
        DistName = ""
        For i As Integer = 0 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")
                If DistName <> "" Then DistName &= ","
                DistName &= Convert.ToString(Me.DistrictList.Items(i).Text)
            End If
        Next

        'If DistID1 <> "" Then
        '    NewDistID = Mid(DistID1, 1, DistID1.Length - 1)
        '    NewDistName = Mid(DistName, 1, DistName.Length - 1)
        'End If
        '報表要用的轄區參數-end

        'Dim relshipstr, relship As String
        ''Dim sql, Years, TMIDStr As String
        'Dim OrgKind2, OrgCName As String
        Dim Years As String = sm.UserInfo.Years - 1911
        If CyclType.Text <> "" Then
            If CyclType.Text.Length < 2 Then
                CyclType.Text = "0" & CyclType.Text
            End If
        End If

        '28:產業人才投資方案 (報表傳入值。)
        Dim Title As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objConn)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim PNAME As String = SearchPlan.SelectedItem.Text
            Select Case SearchPlan.SelectedValue
                Case "G", "W"
                    Title &= "（" & PNAME & "）"
            End Select
        End If

        '28:產業人才投資方案
        Dim OrgKind2 As String = ""
        Select Case SearchPlan.SelectedValue
            Case "A" '不區分
                OrgKind2 = ""
            Case "G" '產業人才投資計畫
                'Title &= "產業人才投資方案（產業人才投資計畫）"
                OrgKind2 = "G"
            Case "W" '提升勞工自主學習計畫
                'Title = "產業人才投資方案（提升勞工自主學習計畫）"
                OrgKind2 = "W"
        End Select
        Dim sPackType As String = ""
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            OrgKind2 = "" '清空
            If PackageType.SelectedValue <> "A" Then
                sPackType = PackageType.SelectedValue
            End If
        End If

        '搜尋型態
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim relship As String = ""
        'Dim relshipstr As String = ""
        Dim OrgCName As String = ""
        If DistType.SelectedValue = "0" Then
            '依轄區，不管訓練機構
            RIDValue.Value = ""
            relship = ""
            OrgCName = ""
        Else
            '依訓練機構，不管轄區
            DistID1 = ""
            DistName = ""
            relship = TIMS.GET_RelshipforRID(RIDValue.Value, objConn)
            OrgCName = center.Text
        End If

        Dim MyValue As String = ""
        MyValue &= "&TMID=" & jobValue.Value
        MyValue &= "&UNIT_SDATE=" & UNIT_SDATE.Text
        MyValue &= "&UNIT_EDATE=" & UNIT_EDATE.Text
        MyValue &= "&start_date=" & start_date.Text
        MyValue &= "&end_date=" & end_date.Text
        MyValue &= "&ClassName=" & Convert.ToString(ClassName.Text)
        MyValue &= "&relship=" & relship
        MyValue &= "&Years=" & Years
        MyValue &= "&Title=" & Title
        MyValue &= "&NewDistID=" & DistID1
        MyValue &= "&NewDistName=" & Convert.ToString(DistName)
        MyValue &= "&OrgCName=" & Convert.ToString(OrgCName)
        MyValue &= "&OrgKind2=" & OrgKind2
        MyValue &= "&PlanYear=" & sm.UserInfo.Years
        MyValue &= "&JobCName=" & Convert.ToString(TB_career_id.Text)
        '依申請階段 
        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If v_AppStage <> "" AndAlso v_AppStage > "0" Then MyValue &= "&AppStage=" & v_AppStage
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&PackageType=" & sPackType

        If sm.UserInfo.LID = "0" Then
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_15_010_R", MyValue)
        Else
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_15_010_R_1", MyValue)
        End If
    End Sub

End Class
