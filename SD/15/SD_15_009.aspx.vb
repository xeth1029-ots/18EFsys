Partial Class SD_15_009
    Inherits AuthBasePage

    'SD_15_009_2010

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            SearchPlan = TIMS.Get_RblOrgPlanKind(SearchPlan, objconn)
            Common.SetListItem(SearchPlan, "G")

            '訓練機構
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

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

            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
            If HistoryRID.Rows.Count <> 0 Then
                center.Attributes("onclick") = "showObj('HistoryList2');"
                center.Style("CURSOR") = "hand"
            End If

            Call Create1()

            Dim sqlstr As String = ""
            sqlstr = "select PlanID,Years from ID_Plan where PlanID= '" & sm.UserInfo.PlanID & "'"
            Dim dr As DataRow = DbAccess.GetOneRow(sqlstr, objconn)
            If Not dr Is Nothing Then
                Common.SetListItem(yearlist, dr("Years"))
            End If

            Common.SetListItem(ClassCate, "")
            Common.SetListItem(JobID, "")
        End If

        '選擇全部轄區
        'DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"
        '選擇全部縣市
        Me.CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');"
        '選擇全部訓練計畫
        'PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"
        '訓練機構
        Button1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"

    End Sub

    Sub Create1()
        AppStage = TIMS.Get_AppStage(AppStage)

        yearlist = TIMS.GetSyear(yearlist)

        DistID = TIMS.Get_DistID(DistID)

        '縣市
        CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing)

        '課程類別
        Dim sqlstr As String = ""
        sqlstr = "SELECT CCID,CCName FROM KEY_CLASSCATELOG ORDER BY SORTKEY"
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        With ClassCate
            .DataSource = dt
            .DataTextField = "CCName"
            .DataValueField = "CCID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", ""))
        End With

        'sqlstr = "SELECT TMID,BusID+'.'+BusName as BusName FROM Key_TrainType WHERE Levels=0"
        '訓練行業別
        sqlstr = "SELECT DISTINCT CONVERT(numeric, JOBID) JOBID,JOBNAME FROM VIEW_TRAINTYPE WHERE BUSID ='G' ORDER BY CONVERT(numeric, JOBID)"
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        With JobID
            .DataSource = dt
            .DataTextField = "JOBNAME"
            .DataValueField = "JOBID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", ""))
        End With

    End Sub

    '列印
    Private Sub Button3_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.ServerClick
        '選擇計畫別
        Dim OrgKind2 As String = ""
        Select Case SearchPlan.SelectedValue
            Case "G" '產業人才投資計畫
                'Title = "產業人才投資方案（產業人才投資計畫）"
                OrgKind2 = "G"
            Case "W" '提升勞工自主學習計畫
                'Title = "產業人才投資方案（提升勞工自主學習計畫）"
                OrgKind2 = "W"
        End Select

        '報表要用的縣市參數
        Dim ICity As String = ""
        For i As Integer = 1 To Me.CityList.Items.Count - 1
            If Me.CityList.Items(i).Selected Then
                If ICity <> "" Then ICity &= ","
                ICity &= Convert.ToString("\'" & Me.CityList.Items(i).Value & "\'")
            End If
        Next

        '課程審核
        '<if test="Result != null and Result != ''">			
        '   and a.AppliedResult ${Result} 	
        '</if>	
        Dim strResult As String = ""
        If rdlResult.SelectedValue <> "A" Then
            If rdlResult.SelectedValue = "T" Then
                strResult = "NOT IN (\'Y\',\'N\')"
            Else
                strResult = "IN (\'" & rdlResult.SelectedValue & "\')"
            End If

        End If

        If RIDValue.Value = "A" Then RIDValue.Value = ""

        Dim sPackType As String = ""
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            OrgKind2 = "" '清空
            If PackageType.SelectedValue <> "A" Then
                sPackType = PackageType.SelectedValue
            End If
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&Years=" & yearlist.SelectedValue
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&Citys=" & ICity
        MyValue &= "&Citys2=" & ICity
        MyValue &= "&ClassCate=" & ClassCate.SelectedValue
        MyValue &= "&JobID=" & JobID.SelectedValue
        MyValue &= "&OrgKind2=" & OrgKind2
        MyValue &= "&Result=" & strResult
        MyValue &= "&UserID=" & sm.UserInfo.UserID
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&PackageType=" & sPackType
        'ReportQuery BussinessTrain SD_15_009_2010
        '依申請階段 
        Dim v_AppStage As String = TIMS.GetListValue(AppStage)
        If v_AppStage <> "" AndAlso v_AppStage > "0" Then MyValue &= "&AppStage=" & v_AppStage

        Dim sCaseYears As String = "2012" '2012:舊年度資訊 2013:新年度資訊
        Dim int_Years As Integer = 2012 'int_Years
        If yearlist.SelectedValue <> "" Then int_Years = CInt(yearlist.SelectedValue)

        sCaseYears = "2012"
        If int_Years > 2012 Then sCaseYears = "2013"

        Select Case sCaseYears
            Case "2012"
                '舊年度資訊 '寫在smartQuery裏面
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_15_009_2010", MyValue)
            Case Else
                '新年度2013
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "BussinessTrain", "SD_15_009_2010", MyValue)
        End Select
    End Sub
End Class

