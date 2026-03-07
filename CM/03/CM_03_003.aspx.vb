Partial Class CM_03_003
    Inherits AuthBasePage

    '<asp:ListItem Value="0" Selected="True">身分別</asp:ListItem>
    '                    <asp:ListItem Value="1" > 年齡</asp:ListItem>
    '                    <asp:ListItem Value="2" > 訓練職類</asp:ListItem>
    '                    <asp:ListItem Value="3" > 教育程度</asp:ListItem>
    '                    <asp:ListItem Value="4" > 性別</asp:ListItem>
    '                    <asp:ListItem Value="5" > 通俗職類</asp:ListItem>
    '                    <asp:ListItem Value="6" > 訓練業別</asp:ListItem>
    '                    <asp:ListItem Value="7" > 縣市別</asp:ListItem>
    '                    <asp:ListItem Value="9" > 上課時數</asp:ListItem>
    Dim str_rblMODE_ALL As String = "身分別,年齡,教育程度,性別,通俗職類,縣市別,上課時數,訓練職類(大類),訓練職類(中類),訓練職類(小類)"
    Const cst_rblMode_身分別 As String = "0"
    Const cst_rblMode_年齡 As String = "1" 'Const cst_rblMode_訓練職類 As String="2"
    Const cst_rblMode_教育程度 As String = "2"
    Const cst_rblMode_性別 As String = "3"
    Const cst_rblMode_通俗職類 As String = "4" 'Const cst_rblMode_訓練業別 As String="6"
    Const cst_rblMode_縣市別 As String = "5" 'Const cst_rblMode_失業週數 As String="8" '2019年06月不使用該項
    Const cst_rblMode_上課時數 As String = "6"
    Const cst_rblMode_訓練職類大類 As String = "7"
    Const cst_rblMode_訓練職類中類 As String = "8"
    Const cst_rblMode_訓練職類小類 As String = "9"

    'Dim t_TrainType1 As DataTable=TIMS.Get_KeyTable("KEY_TRAINTYPE", "busid is not null and busid !='G' and busid !='H'", tConn)  'Cst_訓練職類大類
    '    str_sql=" SELECT TMID JOBTMID, JOBNAME FROM KEY_TRAINTYPE WHERE LEVELS='1' AND PARENT!=197 AND PARENT!=600" & vbCrLf
    '    Dim t_TrainType2 As DataTable=TIMS.Get_KeyTable2(str_sql, tConn)
    '    str_sql=" SELECT TMID, TRAINNAME FROM KEY_TRAINTYPE WHERE LEVELS='2' AND LEN(TRAINID)=4" & vbCrLf
    '    Dim t_TrainType3 As DataTable=TIMS.Get_KeyTable2(str_sql, tConn)



    'iReport '失業週數 '上課時數 'fname
    'CM_03_003*.jrxml

    'Const c_CM_03_003_1 As String="CM_03_003_1" '身分別:0
    'Const c_CM_03_003_2 As String="CM_03_003_2" '年齡:1
    'Const c_CM_03_003_3 As String="CM_03_003_3" '訓練職類:2
    'Const c_CM_03_003_4 As String="CM_03_003_4" '教育程度:3
    'Const c_CM_03_003_5 As String="CM_03_003_5" '性別:4
    'Const c_CM_03_003_6 As String="CM_03_003_6" '身分別'學習券:0
    'Const c_CM_03_003_7 As String="CM_03_003_7" '通俗職類:5(Mode.SelectedIndex)

    Const c_CM_03_003_11 As String = "CM_03_003_11" '身分別:0
    Const c_CM_03_003_12 As String = "CM_03_003_12" '年齡:1
    'Const c_CM_03_003_13 As String="CM_03_003_13" '訓練職類:2
    Const c_CM_03_003_14 As String = "CM_03_003_14" '教育程度:3
    Const c_CM_03_003_15 As String = "CM_03_003_15" '性別:4
    'Const c_CM_03_003_6 As String="CM_03_003_16" '身分別'學習券:0
    Const c_CM_03_003_17 As String = "CM_03_003_17" '通俗職類:5(Mode.SelectedIndex)
    'Const c_CM_03_003_18 As String="CM_03_003_18" '6:訓練業別
    Const c_CM_03_003_19 As String = "CM_03_003_19" '7:縣市別
    'Const c_CM_03_003_20 As String="CM_03_003_20" '8:失業週數
    '上課時數

    '訓練業別:6(Mode.SelectedIndex)
    '縣市別:7(Mode.SelectedIndex)
    'Const c_CM_03_003_8 As String="CM_03_003_8" '失業週數:8(Mode.SelectedIndex)
    'Const c_CM_03_003_9 As String="CM_03_003_9" '上課時數:9(Mode.SelectedIndex)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        'trSYEAR.Visible=False
        'If TIMS.sUtl_ChkTest() Then trSYEAR.Visible=True

        'Button2.Attributes("onclick")="GetOrg();"
        Button3.Style("display") = "none"

        If Not IsPostBack Then
            Call CreateItem()

            FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
            FTDate2.Text = TIMS.Cdate3(Now.Date)
            OCID.Style("display") = "none"
            msg.Text = TIMS.cst_NODATAMsg11

            Button4.Visible = False
            btnExport1.Visible = False
        End If

        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            'DistID.SelectedValue=sm.UserInfo.DistID
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            DistID.Enabled = False
        End If

    End Sub

    Sub CreateItem()
        'DistID.Attributes("onclick")="ClearData();"
        'TPlanID.Attributes("onclick")="ClearData();"
        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '選擇全部縣市
        CityID.Attributes("onclick") = "SelectAll('CityID','CityHidden');"

        'If trSYEAR.Visible Then
        '    ddlSYEAR=TIMS.GetSyear(ddlSYEAR)
        '    Common.SetListItem(ddlSYEAR, sm.UserInfo.Years)
        'End If

        ddlSYEAR = TIMS.GetSyear(ddlSYEAR, 0, 0, True, TIMS.cst_ddl_NotCase)
        'Common.SetListItem(ddlSYEAR, sm.UserInfo.Years)
        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        'DistID.Items(0).Selected=True
        '計畫
        'TPlanID=TIMS.Get_TPlan(TPlanID, , 1, "Y")
        Dim strWHERE As String = "TPLANID IN ('06','07','70')"
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", strWHERE)
        Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
        '縣市
        Dim dtCity As DataTable = TIMS.Get_dtCity(objconn)
        CityID = TIMS.Get_CityName(CityID, dtCity)
        '預算來源
        'BudgetList=TIMS.Get_Budget(BudgetList, 3)
        BudgetList = TIMS.Get_Budget(BudgetList, 29)

        Dim str_rblMODE As String() = str_rblMODE_ALL.Split(",")
        rblMode.Items.Clear()
        For i As Integer = 0 To str_rblMODE.Length - 1
            rblMode.Items.Add(New ListItem(str_rblMODE(i), CStr(i)))
        Next
    End Sub

    'LIST
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim iMaxCells As Integer = DataGrid1.Columns.Count - 1 '7
        Select Case e.Item.ItemType
            Case ListItemType.Footer
                e.Item.Cells(0).ForeColor = Color.Blue
                For i As Integer = 1 To iMaxCells
                    e.Item.Cells(i).Text = 0
                    e.Item.Cells(i).ForeColor = Color.Blue
                Next
                For Each item As DataGridItem In DataGrid1.Items
                    For i As Integer = 1 To iMaxCells
                        e.Item.Cells(i).Text = Int(Val(e.Item.Cells(i).Text)) + Int(Val(item.Cells(i).Text))
                    Next
                Next
        End Select
    End Sub

    '搜尋查詢條件
    Function Get_SchStrX1() As String
        Dim SearchStr As String = ""

        Dim v_ddlSYEAR As String = TIMS.GetListValue(ddlSYEAR)
        If v_ddlSYEAR <> "" Then
            SearchStr &= " and ip.Years='" & v_ddlSYEAR & "'" & vbCrLf
        End If

        Dim sDistID As String = ""
        For Each objitem As ListItem In Me.DistID.Items '選擇轄區
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If sDistID <> "" Then sDistID &= ","
                sDistID &= "'" & objitem.Value & "'"
            End If
        Next
        If sDistID <> "" Then
            SearchStr &= " and ip.DistID IN (" & sDistID & ")" & vbCrLf
        End If

        Dim iteTPlanID As String = ""
        For Each objitem As ListItem In Me.TPlanID.Items   '計畫
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If iteTPlanID <> "" Then iteTPlanID &= ","
                iteTPlanID &= "'" & objitem.Value & "'"
            End If
        Next
        If iteTPlanID <> "" Then
            SearchStr &= " and ip.TPlanID IN (" & iteTPlanID & ")" & vbCrLf
        End If

        '縣市
        Dim itemCtiy As String = ""
        For Each objitem As ListItem In Me.CityID.Items   '計畫
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If itemCtiy <> "" Then itemCtiy += ","
                itemCtiy &= "'" & objitem.Value.ToString & "'"
            End If
        Next
        If itemCtiy <> "" Then
            SearchStr &= " and cc.TADDRESSZIP IN (SELECT ZIPCODE FROM VIEW_ZIPNAME WHERE CTID IN (" & itemCtiy & "))" & vbCrLf
        End If

        'If DistID.SelectedIndex <> 0 Then
        'SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID='" & DistID.SelectedValue & "')"
        'SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID IN ('" & itemstr & "'))"
        'End If
        'If TPlanID.SelectedIndex <> 0 Then
        'SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & TPlanID.SelectedValue & "')"
        'SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID IN ('" & itemplan & "'))"
        'SearchStr += " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID IN ('" & newTPlanID & "'))"
        'End If
        If RIDValue.Value <> "" Then
            SearchStr &= " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
        End If
        If PlanID.Value <> "" Then
            SearchStr &= " and cc.PlanID='" & PlanID.Value & "'" & vbCrLf
        End If

        Dim OCIDStr As String = ""
        OCIDStr = ""
        For Each item As ListItem In OCID.Items
            If item.Selected = True Then
                If item.Value = "%" Then
                    OCIDStr = ""
                    Exit For
                Else
                    If OCIDStr <> "" Then OCIDStr += ","
                    OCIDStr += "'" & item.Value & "'"
                End If
            End If
        Next
        If OCIDStr <> "" Then
            SearchStr &= " and cc.OCID IN (" & OCIDStr & ")" & vbCrLf
        Else
            '開訓區間
            If STDate1.Text <> "" Then
                SearchStr &= " and cc.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
            End If
            If STDate2.Text <> "" Then
                SearchStr &= " and cc.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
            End If

            '結訓區間
            If FTDate1.Text <> "" Then
                SearchStr &= " and cc.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
            End If
            If FTDate2.Text <> "" Then
                SearchStr &= " and cc.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
            End If
        End If

        '選擇預算來源
        'Dim BudgetIDStr As String=""
        Dim itembudget As String = ""
        itembudget = TIMS.CombiSQM2IN(TIMS.GetCblValue(BudgetList))
        'For Each objitem As ListItem In Me.BudgetList.Items
        '    If objitem.Selected=True Then
        '        If itembudget <> "" Then itembudget += ","
        '        itembudget += "'" & objitem.Value & "'"
        '    End If
        'Next
        If itembudget <> "" Then
            SearchStr &= " and cs.BudgetID in (" & itembudget & ") " & vbCrLf
        End If
        Return SearchStr
    End Function

    Sub CheckData1(ByRef sErrmsg As String)
        'Dim sErrmsg As String=""
        Dim v_rblMode As String = TIMS.GetListValue(rblMode)
        If v_rblMode = "" Then sErrmsg &= "請選擇 統計項目 ，必選不可為空" & vbCrLf
        Return
    End Sub

    'SQL 查詢 匯出xls
    Sub Search1()

        Dim v_rblMode As String = TIMS.GetListValue(rblMode)
        Dim SearchStr As String = Get_SchStrX1()
        Dim sql As String = ""

        Select Case v_rblMode'Mode.SelectedIndex
            Case cst_rblMode_身分別 '0 '身分別 '一般身分別
                sql = "" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
                sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

                sql &= " FROM (" & vbCrLf
                sql &= "    select IdentityID sort ,Name Title FROM dbo.KEY_IDENTITY" & vbCrLf
                sql &= "    where 1=1" & vbCrLf
                If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    sql &= "   AND IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")" & vbCrLf
                Else
                    sql &= "   AND IDENTITYID IN (" & TIMS.Cst_Identity28_2019_11 & ")" & vbCrLf
                End If
                sql &= " ) a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.MIdentityID sort" & vbCrLf
                sql &= " 	,COUNT(case when cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(case when cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(case when cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	,COUNT(case when ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
                sql &= " 	FROM dbo.id_plan ip" & vbCrLf
                sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid and cc.NotOpen='N' and cc.IsSuccess='Y'" & vbCrLf
                sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= "    LEFT JOIN dbo.Key_Degree b3 on b3.DegreeID=ss.DegreeID" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
                sql &= " 	WHERE 1=1" & vbCrLf
                sql &= SearchStr
                'sql &= " 	and ip.Years='2018'" & vbCrLf
                'sql &= " 	and ip.DistID IN ('001')" & vbCrLf
                'sql &= " 	and ip.TPlanID IN ('06')" & vbCrLf
                'sql &= " 	and cc.FTDate <= convert(datetime, '2018/10/02')" & vbCrLf
                sql &= " 	GROUP BY cs.MIdentityID" & vbCrLf

                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_年齡 '1 年齡
                'SELECT '19歲以下' Title,1 Sort 
                'UNION SELECT '20歲-24歲' Title,2 Sort 
                'UNION SELECT '25歲-29歲' Title,3 Sort 
                'UNION SELECT '30歲-34歲' Title,4 Sort 
                'UNION SELECT '35歲-39歲' Title,5 Sort 
                'UNION SELECT '40歲-44歲' Title,6 Sort 
                'UNION SELECT '45歲-49歲' Title,7 Sort 
                'UNION SELECT '50歲-54歲' Title,8 Sort 
                'UNION SELECT '55歲-59歲' Title,9 Sort 
                'UNION SELECT '60歲-64歲' Title,10 Sort 
                'UNION SELECT '65歲以上' Title,11 Sort 

                sql = "" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
                sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

                sql &= " FROM (" & vbCrLf
                sql &= " SELECT '19歲以下' Title,1 Sort " & vbCrLf
                sql &= " UNION SELECT '20歲-24歲' Title,2 Sort " & vbCrLf
                sql &= " UNION SELECT '25歲-29歲' Title,3 Sort " & vbCrLf
                sql &= " UNION SELECT '30歲-34歲' Title,4 Sort " & vbCrLf
                sql &= " UNION SELECT '35歲-39歲' Title,5 Sort " & vbCrLf
                sql &= " UNION SELECT '40歲-44歲' Title,6 Sort " & vbCrLf
                sql &= " UNION SELECT '45歲-49歲' Title,7 Sort " & vbCrLf
                sql &= " UNION SELECT '50歲-54歲' Title,8 Sort " & vbCrLf
                sql &= " UNION SELECT '55歲-59歲' Title,9 Sort " & vbCrLf
                sql &= " UNION SELECT '60歲-64歲' Title,10 Sort " & vbCrLf
                sql &= " UNION SELECT '65歲以上' Title,11 Sort " & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                'sql &= " select g.sort " & vbCrLf
                sql &= "    select dbo.FN_YEARSOLDID2D(dbo.FN_YEARSOLD(cc.FTDate, ss.birthday)) YEARSOLD " & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
                sql &= " 	FROM dbo.id_plan ip" & vbCrLf
                sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Key_Degree b3 on b3.DegreeID=ss.DegreeID" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
                sql &= " 	WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf

                sql &= "    GROUP BY dbo.fn_YEARSOLDID2D(dbo.fn_YEARSOLD(cc.ftdate, ss.birthday))" & vbCrLf
                sql &= " ) b on b.YEARSOLD=a.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            'Case cst_rblMode_訓練職類 '2 '訓練職類
            '    sql="" & vbCrLf
            '    sql &= " SELECT a.Title TITLE" & vbCrLf
            '    sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
            '    sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

            '    sql &= " FROM (" & vbCrLf
            '    sql &= " select max(b3.tmkey) sort,b3.BusName+'-'+b3.JobName Title" & vbCrLf
            '    sql &= " FROM VIEW_TRAINTYPE b3" & vbCrLf
            '    sql &= " GROUP BY b3.BusName+'-'+b3.JobName" & vbCrLf
            '    sql &= " ) a" & vbCrLf
            '    sql &= " LEFT JOIN (" & vbCrLf
            '    sql &= "    select b3.BusName+'-'+b3.JobName Title " & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
            '    sql &= " 	FROM dbo.id_plan ip" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
            '    sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
            '    sql &= " 	JOIN dbo.VIEW_TRAINTYPE b3 on b3.TMID=cc.TMID" & vbCrLf
            '    sql &= " 	LEFT JOIN dbo.Key_Degree b4 on b4.DegreeID=ss.DegreeID" & vbCrLf
            '    sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
            '    sql &= " 	WHERE 1=1" & vbCrLf
            '    sql &= SearchStr & vbCrLf

            '    sql &= " GROUP BY  b3.BusName+'-'+b3.JobName" & vbCrLf
            '    sql &= " ) b on a.Title=b.Title" & vbCrLf
            '    sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_教育程度 '3 '教育程度 Const c_CM_03_003_4 As String="CM_03_003_4" '教育程度:3
                sql = "" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
                sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

                sql &= " FROM (" & vbCrLf
                sql &= " select b3.Name Title" & vbCrLf
                sql &= " ,b3.DegreeID+b3.Name sort" & vbCrLf
                sql &= " FROM Key_Degree b3" & vbCrLf
                sql &= " where b3.DEGREETYPE=1" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= "    select b3.DegreeID+b3.Name sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
                sql &= " 	FROM dbo.id_plan ip" & vbCrLf
                sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Key_Degree b3 on b3.DegreeID=ss.DegreeID" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
                sql &= " 	WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " GROUP BY  b3.DegreeID+b3.Name" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_性別 '4 '性別  Const c_CM_03_003_5 As String="CM_03_003_5" '性別:4
                sql = "" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
                sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

                sql &= " FROM (" & vbCrLf
                sql &= " 	SELECT '男' Title,1 Sort ,'M' title2" & vbCrLf
                sql &= " 	UNION SELECT '女' Title,2 Sort ,'F' title2" & vbCrLf
                sql &= " ) a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	SELECT ss.Sex title2" & vbCrLf

                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
                sql &= " 	FROM dbo.id_plan ip" & vbCrLf
                sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Key_Degree b3 on b3.DegreeID=ss.DegreeID" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
                sql &= " 	WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                'sql &= " 	and ip.Years='2018'" & vbCrLf
                'sql &= " 	and ip.DistID IN ('001')" & vbCrLf
                'sql &= " 	and ip.TPlanID IN ('06')" & vbCrLf
                'sql &= " 	and cc.FTDate <= convert(date, '2018/10/02')" & vbCrLf
                sql &= " 	GROUP BY ss.Sex" & vbCrLf
                sql &= " ) b on a.title2=b.title2" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_通俗職類 '5 'CM_03_003_7  '通俗職類  
                '啟用2016年通俗職類
                Dim flag_Cjob2016 As Boolean = TIMS.Get_sCjob2016_USE(Page)
                Dim str_SHARECJOB_YEAR As String = ""
                If flag_Cjob2016 Then str_SHARECJOB_YEAR = TIMS.cst_SHARE_CJOB_2016

                sql = "" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
                sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf
                sql &= " FROM (" & vbCrLf

                Select Case str_SHARECJOB_YEAR
                    Case TIMS.cst_SHARE_CJOB_2016
                        sql &= "  select b3.CJOB_NAME Title " & vbCrLf
                        sql &= "  ,b3.CJOB_UNKEY SORT " & vbCrLf '依此JOIN 
                        sql &= "  ,ISNULL(b3.CJOB_NO,b3.CJOB_TYPE)+ISNULL(b3.JOB_NO ,'') SORT2" & vbCrLf '依此排序
                        sql &= "  FROM SHARE_CJOB b3" & vbCrLf
                        sql &= "  WHERE b3.CYEARS='2019'" & vbCrLf
                    Case Else
                        sql &= "  select b3.CJOB_NAME Title " & vbCrLf
                        sql &= "  ,b3.CJOB_UNKEY SORT " & vbCrLf '依此JOIN 
                        sql &= "  ,CASE WHEN b3.CJOB_NO IS NOT NULL THEN RIGHT('0000'+ISNULL(b3.CJOB_NO ,''),4) ELSE RIGHT('00'+ISNULL(b3.CJOB_TYPE,''),2) END SORT2" & vbCrLf '依此排序
                        sql &= "  FROM SHARE_CJOB b3" & vbCrLf
                        sql &= "  WHERE b3.CYEARS='2014'" & vbCrLf
                End Select

                sql &= " ) a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                'sql &= "  SELECT b3.CJOB_NO+b3.CJOB_Name sort " & vbCrLf
                sql &= "  SELECT b3.CJOB_UNKEY SORT " & vbCrLf

                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
                sql &= " 	FROM dbo.id_plan ip" & vbCrLf
                sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " 	JOIN dbo.SHARE_CJOB b3 on b3.CJOB_UNKEY=cc.CJOB_UNKEY" & vbCrLf
                sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
                sql &= " 	WHERE 1=1" & vbCrLf

                sql &= SearchStr & vbCrLf
                'sql &= " GROUP BY b3.CJOB_NO+b3.CJOB_Name" & vbCrLf
                sql &= "  GROUP BY b3.CJOB_UNKEY" & vbCrLf

                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.SORT2" & vbCrLf

            'Case cst_rblMode_訓練業別 '6 '訓練業別
            '    sql="" & vbCrLf
            '    sql &= " SELECT a.Title TITLE" & vbCrLf
            '    sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
            '    sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

            '    sql &= " FROM (" & vbCrLf
            '    sql &= " select max(b3.tmkey) sort,b3.BusName+'-'+b3.JobName Title" & vbCrLf
            '    sql &= " FROM VIEW_TRAINTYPE b3" & vbCrLf
            '    sql &= " GROUP BY b3.BusName+'-'+b3.JobName" & vbCrLf
            '    sql &= " ) a" & vbCrLf
            '    sql &= " LEFT JOIN (" & vbCrLf
            '    sql &= " select b3.BusName+'-'+b3.JobName Title " & vbCrLf

            '    sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
            '    sql &= " 	FROM dbo.id_plan ip" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
            '    sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
            '    sql &= " 	JOIN dbo.VIEW_TRAINTYPE b3 on b3.TMID=cc.TMID" & vbCrLf
            '    sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
            '    sql &= " 	WHERE 1=1" & vbCrLf
            '    sql &= SearchStr & vbCrLf

            '    sql &= "  GROUP BY  b3.BusName+'-'+b3.JobName" & vbCrLf
            '    sql &= " ) b on a.Title=b.Title" & vbCrLf
            '    sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_縣市別 '7 '縣市別
                sql = "" & vbCrLf
                sql &= " WITH WS1 AS (" & vbCrLf
                sql &= " SELECT cs.socid ,cs.StudStatus ,cc.STDate ,cc.FTDate" & vbCrLf
                sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
                sql &= " ,b3.CTID" & vbCrLf
                sql &= " FROM dbo.ID_PLAN ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sql &= " LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WG1 AS (" & vbCrLf
                sql &= " select b3.CTID sort,b3.CTNAME Title" & vbCrLf
                sql &= " FROM ID_CITY b3 WHERE b3.CTID<>999" & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " FROM WG1 a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.CTID sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cs.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	FROM WS1 cs" & vbCrLf
                sql &= " 	GROUP BY cs.CTID" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            'Case "8" 'cst_rblMode_失業週數 '8 '失業週數
            '    sql="" & vbCrLf
            '    sql &= " SELECT a.Title TITLE" & vbCrLf
            '    sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
            '    sql &= " ,ISNULL(b.SubsidyCount,0) SUBSIDYCOUNT" & vbCrLf
            '    sql &= " ,ISNULL(b.wjobCount,0) WJOBCOUNT" & vbCrLf

            '    sql &= " FROM (" & vbCrLf
            '    sql &= "  select 1 sort, '(失業25週以下)' Title " & vbCrLf
            '    sql &= "  union select 2 sort, '(失業26週以上)' Title " & vbCrLf
            '    sql &= " ) a" & vbCrLf
            '    sql &= " LEFT JOIN (" & vbCrLf
            '    'sql &= " select st.LostJobWeek sort" & vbCrLf
            '    sql &= "  select ISNULL(st.LostJobWeek,1) sort " & vbCrLf

            '    sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN ss3.socid is not null and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() then 1 end) SubsidyCount" & vbCrLf
            '    sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) and cs.WkAheadOfSch='Y' THEN 1 end) wjobCount" & vbCrLf
            '    sql &= " 	FROM dbo.id_plan ip" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
            '    sql &= " 	JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
            '    sql &= " 	JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
            '    sql &= " 	LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
            '    sql &= " 	LEFT JOIN dbo.Sub_SubSidyApply ss3 on ss3.socid=cs.socid and ss3.AppliedStatusFin='Y'" & vbCrLf
            '    sql &= "    LEFT JOIN dbo.VIEW_LOSTJOBWEEK st on st.socid=cs.socid" & vbCrLf
            '    sql &= " 	WHERE 1=1" & vbCrLf
            '    sql &= SearchStr & vbCrLf
            '    sql &= " GROUP BY  st.LostJobWeek" & vbCrLf

            '    sql &= " ) b on a.sort=b.sort" & vbCrLf
            '    sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_上課時數  '6
                sql = "" & vbCrLf
                sql &= " WITH WS1 AS (" & vbCrLf
                sql &= " SELECT cs.socid ,cs.StudStatus ,cc.STDate ,cc.FTDate" & vbCrLf
                sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
                sql &= " ,th.TRAINHOURS1" & vbCrLf
                sql &= " FROM dbo.ID_PLAN ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sql &= " LEFT JOIN V_CLASSTRAINHOURS th on th.OCID=cc.OCID" & vbCrLf
                sql &= " LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WG1 AS (" & vbCrLf
                sql &= " SELECT TRAINHOURS1 sort" & vbCrLf
                sql &= " ,TRAINHOURS Title" & vbCrLf
                sql &= " FROM dbo.V_TRAINHOURS" & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " FROM WG1 a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.TRAINHOURS1 sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cs.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	FROM WS1 cs" & vbCrLf
                sql &= " 	GROUP BY cs.TRAINHOURS1" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_訓練職類大類 '"7"
                sql = "" & vbCrLf
                sql &= " WITH WS1 AS (" & vbCrLf
                sql &= " SELECT cs.socid ,cs.StudStatus ,cc.STDate ,cc.FTDate" & vbCrLf
                sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
                sql &= " FROM dbo.ID_PLAN ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sql &= " LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WG1 AS (" & vbCrLf
                sql &= " SELECT BUSID sort" & vbCrLf
                sql &= " ,BUSNAME Title" & vbCrLf
                sql &= " FROM dbo.KEY_TRAINTYPE WHERE BUSID IS NOT NULL AND BUSID !='G' AND BUSID !='H'" & vbCrLf
                sql &= " )" & vbCrLf

                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " FROM WG1 a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.BUSID sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cs.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	FROM WS1 cs" & vbCrLf
                sql &= " 	GROUP BY cs.BUSID" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_訓練職類中類 '8
                sql = "" & vbCrLf
                sql &= " WITH WS1 AS (" & vbCrLf
                sql &= " SELECT cs.socid ,cs.StudStatus ,cc.STDate ,cc.FTDate" & vbCrLf
                sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
                sql &= " FROM dbo.ID_PLAN ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sql &= " LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WG1 AS (" & vbCrLf
                sql &= " SELECT TMID sort /*JOBTMID*/" & vbCrLf
                sql &= " ,JOBNAME Title FROM KEY_TRAINTYPE WHERE LEVELS='1' AND PARENT!=197 AND PARENT!=600" & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " FROM WG1 a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.JOBTMID sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cs.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	FROM WS1 cs" & vbCrLf
                sql &= " 	GROUP BY cs.JOBTMID" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

            Case cst_rblMode_訓練職類小類 '9
                sql = "" & vbCrLf
                sql &= " WITH WS1 AS (" & vbCrLf
                sql &= " SELECT cs.socid ,cs.StudStatus ,cc.STDate ,cc.FTDate" & vbCrLf
                sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
                sql &= " FROM dbo.ID_PLAN ip" & vbCrLf
                sql &= " JOIN dbo.Class_ClassInfo cc on cc.planid=ip.planid AND cc.NotOpen='N' AND cc.IsSuccess='Y'" & vbCrLf
                sql &= " JOIN dbo.Class_StudentsOfClass cs ON cs.OCID=cc.OCID" & vbCrLf
                sql &= " JOIN dbo.Stud_StudentInfo ss on ss.sid=cs.sid" & vbCrLf
                sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sql &= " LEFT JOIN dbo.VIEW_ZIPNAME b3 on b3.ZipCode=cc.TaddressZip" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= SearchStr & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " ,WG1 AS (" & vbCrLf
                sql &= " SELECT TMID sort" & vbCrLf
                sql &= " ,TRAINNAME Title" & vbCrLf
                sql &= " FROM KEY_TRAINTYPE WHERE LEVELS='2' AND LEN(TRAINID)=4" & vbCrLf
                sql &= " )" & vbCrLf
                sql &= " SELECT a.Title TITLE" & vbCrLf
                sql &= " ,ISNULL(b.JoinStudent,0) JOINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.FinStudent,0) FINSTUDENT" & vbCrLf
                sql &= " ,ISNULL(b.ReStudent,0) RESTUDENT" & vbCrLf
                sql &= " FROM WG1 a" & vbCrLf
                sql &= " LEFT JOIN (" & vbCrLf
                sql &= " 	select cs.TMID sort" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.socid is not null then 1 end) JoinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus Not IN (2,3) and cs.FTDate < GETDATE() then 1 end) FinStudent" & vbCrLf
                sql &= " 	,COUNT(CASE WHEN cs.StudStatus IN (2,3) then 1 end) ReStudent" & vbCrLf
                sql &= " 	FROM WS1 cs" & vbCrLf
                sql &= " 	GROUP BY cs.TMID" & vbCrLf
                sql &= " ) b on a.sort=b.sort" & vbCrLf
                sql &= " order by a.sort" & vbCrLf

        End Select

        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2) ' "查詢時發生錯誤，請重新輸入查詢值!!"
            Exit Sub
        End If

        Dim dt2 As DataTable = Nothing
        Try
            'Dim da As SqlDataAdapter=TIMS.GetOneDA()
            'Dim dt2 As New DataTable
            'TIMS.Fill(sql, da, dt2)
            dt2 = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, TIMS.cst_ErrorMsg9) ' "資料庫效能異常，請重設查詢條件再查詢"
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/*  sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
            Exit Sub
        End Try

        DataGrid1.Visible = False
        Button4.Visible = False
        btnExport1.Visible = False
        If dt2.Rows.Count > 0 Then
            DataGrid1.Visible = True
            Button4.Visible = True
            btnExport1.Visible = True
            FrameTable3.Visible = False
        End If
        DataGrid1.DataSource = dt2
        DataGrid1.DataBind()

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sErrmsg As String = ""
        CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Call Search1()
        Page.RegisterStartupScript("load", "<script>ReStart();</script>")

    End Sub

    '查詢班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        msg.Text = ""
        Dim dt As DataTable = Nothing
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("PlanID", PlanID.Value)
        parms.Add("RID", RIDValue.Value)
        Dim sql As String = ""
        sql = "SELECT PLANID,RID,OCID,CLASSCNAME2 FROM VIEW2 WHERE PlanID=@PlanID and RID=@RID and NotOpen='N' and IsSuccess='Y'"
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then
            msg.Text = "查無此機構底下的班級"
            msg.Visible = True
            OCID.Style("display") = "none"
        Else
            OCID.Items.Clear()
            OCID.Items.Add(New ListItem("全選", "%"))
            For Each dr As DataRow In dt.Rows
                OCID.Items.Add(New ListItem(Convert.ToString(dr("CLASSCNAME2")), Convert.ToString(dr("OCID"))))
            Next
            OCID.Style("display") = "inline"
            msg.Visible = False
        End If
    End Sub

    '回上頁
    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        DataGrid1.Visible = False
        'Button4.Visible=False
        Button3.Style("display") = "none"
        Button4.Visible = False
        btnExport1.Visible = False

        FrameTable3.Visible = True
    End Sub

    'Sub print_iReport1()
    '    'Common.MessageBox(Me, "暫不提供列印!!")
    '    'Exit Sub
    '    'Dim str_rblMODE As String()=str_rblMODE_ALL.Split(",")
    '    Dim v_rblMode As String=TIMS.GetListValue(rblMode)
    '    'Case 0, 1, 2, 3, 4, 5, 6, 7, 8
    '    'Case cst_rblMode_上課時數
    '    Select Case v_rblMode'Mode.SelectedIndex
    '        Case cst_rblMode_身分別
    '        Case cst_rblMode_年齡 'Case cst_rblMode_訓練職類
    '        Case cst_rblMode_教育程度
    '        Case cst_rblMode_性別
    '        Case cst_rblMode_通俗職類 'Case cst_rblMode_訓練業別
    '        Case cst_rblMode_縣市別
    '            'Case cst_rblMode_上課時數' As String="6"
    '            'Case cst_rblMode_訓練職類_大類 'As String="7"
    '            'Case cst_rblMode_訓練職類_中類' As String="8"
    '            'Case cst_rblMode_訓練職類_小類 ' As String="9"
    '        Case Else
    '            Common.MessageBox(Me, "暫不提供列印!!")
    '            Exit Sub
    '    End Select

    '    '報表要用的轄區參數
    '    Dim DistID1 As String=""
    '    'Dim DistName As String=""
    '    For i As Integer=1 To Me.DistID.Items.Count - 1
    '        If Me.DistID.Items(i).Selected AndAlso Me.DistID.Items(i).Value <> "" Then
    '            If DistID1 <> "" Then DistID1 &= ","
    '            DistID1 &= Me.DistID.Items(i).Value

    '            'If DistName <> "" Then DistName &= ","
    '            'DistName &= Me.DistID.Items(i).Text
    '        End If
    '    Next
    '    'If DistID1 <> "" Then
    '    '    newDistID=Mid(DistID1, 1, DistID1.Length - 1)
    '    '    newDistName=Mid(DistName, 1, DistName.Length - 1)
    '    'End If

    '    '報表要用的訓練計畫參數
    '    Dim TPlanID1 As String=""
    '    'Dim TPlanName As String=""
    '    For i As Integer=1 To Me.TPlanID.Items.Count - 1
    '        If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
    '            If TPlanID1 <> "" Then TPlanID1 &= ","
    '            TPlanID1 &= Me.TPlanID.Items(i).Value
    '            'If TPlanName <> "" Then TPlanName &= ","
    '            'TPlanName &= Me.TPlanID.Items(i).Text
    '        End If
    '    Next
    '    'If TPlanID1 <> "" Then
    '    '    newTPlanID=Mid(TPlanID1, 1, TPlanID1.Length - 1)
    '    '    newTPlanIDName=Mid(TPlanName, 1, TPlanName.Length - 1)
    '    'End If

    '    '報表要用的預算別參數
    '    Dim BudID As String=""
    '    BudID=TIMS.GetCblValue(BudgetList)
    '    'Dim BudIDName As String=""
    '    'For i As Integer=1 To Me.BudgetList.Items.Count - 1
    '    '    If Me.BudgetList.Items(i).Selected AndAlso Me.BudgetList.Items(i).Value <> "" Then
    '    '        If BudID <> "" Then BudID &= ","
    '    '        BudID &= Me.BudgetList.Items(i).Value
    '    '    End If
    '    'Next

    '    'If BudID <> "" Then
    '    '    newBudID=Mid(BudID, 1, BudID.Length - 1)
    '    '    newBudIDName=Mid(BudIDName, 1, BudIDName.Length - 1)
    '    'End If

    '    'If Me.DistID.SelectedIndex=0 Then DistStr="" Else DistStr=Me.DistID.SelectedValue
    '    'If Me.TPlanID.SelectedIndex=0 Then TPlanStr="" Else TPlanStr=Me.TPlanID.SelectedValue

    '    '勾選班級後會省略結訓日期的條件
    '    If (RIDValue.Value <> "") Or (PlanID.Value <> "") Then
    '        FTDate1.Text=""
    '        FTDate2.Text=""
    '        STDate1.Text=""
    '        STDate2.Text=""
    '    End If

    '    Dim OCIDStr As String=""
    '    For Each item As ListItem In OCID.Items
    '        If item.Selected=True AndAlso item.Value <> "" Then
    '            If item.Value="%" Then
    '                OCIDStr=""
    '                Exit For
    '            Else
    '                If OCIDStr <> "" Then OCIDStr &= ","
    '                OCIDStr &= item.Value
    '            End If
    '        End If
    '    Next

    '    '報表要用的標題轄區參數
    '    Dim TDistName As String=""
    '    If sm.UserInfo.DistID="000" Then
    '        TDistName=TIMS.Get_DistName1("000")
    '    Else
    '        TDistName=TIMS.Get_DistName1(sm.UserInfo.DistID) 'newDistName
    '    End If

    '    Dim fname As String=""
    '    Select Case v_rblMode'Mode.SelectedIndex
    '        Case cst_rblMode_身分別 '0 '身分別
    '            'Select Case TPlanID.SelectedValue
    '            '    Case "15" '學習券
    '            '        fname=c_CM_03_003_6
    '            '    Case Else
    '            '        fname=c_CM_03_003_1
    '            'End Select
    '            fname=c_CM_03_003_11
    '        Case cst_rblMode_年齡 '1 '年齡
    '            fname=c_CM_03_003_12
    '        'Case cst_rblMode_訓練職類 '2 '訓練職類
    '        '    fname=c_CM_03_003_13
    '        Case cst_rblMode_教育程度 '3 '教育程度
    '            fname=c_CM_03_003_14
    '        Case cst_rblMode_性別 '4 '性別
    '            fname=c_CM_03_003_15
    '        Case cst_rblMode_通俗職類 '5 '通俗
    '            fname=c_CM_03_003_17
    '            'Case 6 '失業週數
    '            '    fname=c_CM_03_003_8 '失業週數
    '            'Case 7 '上課時數
    '            '    fname=c_CM_03_003_9 '上課時數
    '        'Case cst_rblMode_訓練業別 '6
    '        '    fname=c_CM_03_003_18
    '        Case cst_rblMode_縣市別 '7
    '            fname=c_CM_03_003_19
    '    End Select

    '    'strScript="<script language=""javascript"">" + vbCrLf
    '    Dim s_MIdentityID As String=""
    '    s_MIdentityID=TIMS.Cst_Identity28_2019_11
    '    If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then s_MIdentityID=TIMS.Cst_Identity06_2019_11

    '    Dim MyValue As String=""
    '    MyValue="&OCID=" & OCIDStr '多組
    '    MyValue &= "&BudID=" & BudID '多組
    '    MyValue &= "&TPlanID=" & TPlanID1 '多組
    '    MyValue &= "&DistID=" & DistID1 '多組
    '    If fname=c_CM_03_003_11 Then
    '        MyValue &= "&MIdentityID=" & Replace(s_MIdentityID, "'", "") '多組
    '    End If
    '    MyValue &= "&STTDate=" & TIMS.ClearSQM(STDate1.Text)
    '    MyValue &= "&FTTDate=" & TIMS.ClearSQM(STDate2.Text)
    '    MyValue &= "&SFTDate=" & TIMS.ClearSQM(FTDate1.Text)
    '    MyValue &= "&FFTDate=" & TIMS.ClearSQM(FTDate2.Text)
    '    MyValue &= "&RID=" & TIMS.ClearSQM(RIDValue.Value)
    '    MyValue &= "&Planid=" & TIMS.ClearSQM(PlanID.Value)
    '    MyValue &= "&Years=" & TIMS.GetListValue(ddlSYEAR) '.SelectedValue)

    '    'MyValue += "&TDistName=" & TDistName
    '    'MyValue &= "&TPlanName=" & TPlanName
    '    'MyValue &= "&OrgName=" & center.Text
    '    'MyValue &= "&ClassCName=" & OCIDName
    '    'MyValue &= "&DistName=" & DistName
    '    'MyValue &= "&BudIDName=" & BudIDName
    '    'cst_rblMode_失業週數 Case 8  fname=c_CM_03_003_20
    '    'strScript += "</script>"
    '    'Page.RegisterStartupScript("window_onload", strScript)
    '    If fname="" Then
    '        Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    '        Exit Sub
    '    End If

    '    If fname <> "" Then
    '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, fname, MyValue)
    '    End If

    'End Sub

    '列印
    'Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
    '    Dim sErrmsg As String=""
    '    CheckData1(sErrmsg)
    '    If sErrmsg <> "" Then
    '        Common.MessageBox(Me, sErrmsg)
    '        Exit Sub
    '    End If

    '    print_iReport1()
    'End Sub

    '訓練機構選擇
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        Dim msg As String = ""
        Dim DistID1 As String = ""
        Dim N As Integer = 0   '預設 N=0 表示沒有勾選轄區選項
        For i As Integer = 1 To Me.DistID.Items.Count - 1

            If Me.DistID.Items(i).Selected Then '假如有勾選
                N = N + 1  '計算轄區勾選選項的數目
                If N = 1 Then '如果是勾選一個選項
                    DistID1 = Convert.ToString(Me.DistID.Items(i).Value) '取得選項的值
                End If
                If N = 2 Then '如果轄區勾選選項的數目=2
                    'Common.MessageBox(Me, "只能選擇一個轄區")
                    msg += "只能選擇一個轄區!" & vbCrLf
                    DistID1 = ""
                    Exit For
                End If
            End If
        Next

        If N = 0 Then '如果轄區選項沒有選
            'Common.MessageBox(Me, "請選擇轄區")
            msg += "請選擇轄區!" & vbCrLf
        End If

        Dim TPlanID1 As String = ""
        Dim N1 As Integer = 0 '預設 N1=0 表示沒有勾選計畫選項
        For j As Integer = 1 To Me.TPlanID.Items.Count - 1

            If Me.TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(Me.TPlanID.Items(j).Value) '取得選項的值
                End If
                If N1 = 2 Then '如果計畫勾選選項的數目=2
                    'Common.MessageBox(Me, "只能選擇一個計畫")
                    msg += "只能選擇一個計畫!" & vbCrLf
                    TPlanID1 = ""
                    Exit For
                End If

            End If
        Next

        If N1 = 0 Then '如果計畫選項沒有選
            'Common.MessageBox(Me, "請選擇計畫")
            msg += "請選擇計畫!" & vbCrLf
        End If

        If msg <> "" Then
            Common.MessageBox(Me, msg)
        End If

        If DistID1 <> "" And TPlanID1 <> "" Then
            Dim strScript1 As String

            strScript1 = "<script language=""javascript"">" + vbCrLf
            strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "'  + '&BtnName=Button3','查詢機構',400,400,1);"
            strScript1 += "</script>"
            Page.RegisterStartupScript("", strScript1)

        End If
    End Sub

    Sub Utl_Export1()

        DataGrid1.AllowPaging = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        Const sFileName1 As String = "訓練人數綜合查詢"

        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(Cst_功能欄位).Visible=False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        DataGrid1.AllowPaging = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
        'DataGrid1.Columns(Cst_功能欄位).Visible=True
        'Call TIMS.CloseDbConn(objconn)
    End Sub

    '匯出 xls
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        'Const Cst_功能欄位 As Integer=14
        Dim sErrmsg As String = ""
        CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Call Utl_Export1()
    End Sub

End Class


