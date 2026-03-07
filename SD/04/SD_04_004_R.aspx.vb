Partial Class SD_04_004_R
    Inherits AuthBasePage

    '1.不分頁印出，將助教內容拿掉
    'course_list_1_8 (x)'course_list_9_12 (x)'course_list (x)

    '1b.不分頁印出，助教內容保留 '全期 '年
    'course_list_1_8_ya
    'course_list_9_12_ya
    'course_list_ya

    '2.分兩頁印出字體目前是6pt,改為12pt '月
    'course_list_1_8_1
    'course_list_9_12_1
    'course_list_R1

    '其他:'週
    'course_7_1_8
    'course_list_7_9_12 'course_7_9_12
    'course_7

    '日
    'course_list_1_1_8
    'course_list_1_9_12
    'course_list_1

    '1. course_list_1_8_ya.jrxml
    '2. course_list_1_8_1.jrxml
    '3. course_7_1_8.jrxml
    '4. course_list_1_1_8.jrxml
    'V_CLASS_SCHEDULE
    Const cst_printFN1 As String = "course_list_1_8_ya" '68
    Const cst_printFN1B As String = "course_list_1_8_1" '68
    Const cst_printFN2 As String = "course_list_9_12_ya" '68
    Const cst_printFN2B As String = "course_list_9_12_1" '68
    Const cst_printFN3 As String = "course_list_ya"
    Const cst_printFN3B As String = "course_list_R1"

    Const cst_printFN4 As String = "course_7_1_8" '68
    Const cst_printFN5 As String = "course_list_7_9_12"
    Const cst_printFN6 As String = "course_7"
    Const cst_printFN7 As String = "course_list_1_1_8" '68
    Const cst_printFN8 As String = "course_list_1_9_12"
    Const cst_printFN9 As String = "course_list_1"

    Const cst_vsbNotSet As String = "NotSet" '未設定時
    Const cst_vsbclass As String = "class"
    Const cst_vsbrid As String = "rid"

    'select formal,count(1) cnt from Class_Schedule group by formal
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", False, "Button2")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Button1.Attributes("onclick") = "return ReportPrint();"
        RadioButtonList1.Attributes("onclick") = "ChangeMode();"

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            'Dim iYears1 As Integer = Val(sm.UserInfo.Years) - 1
            'Dim iYears2 As Integer = Val(sm.UserInfo.Years) + 1
            Years = TIMS.GetSyear(Years)
            Common.SetListItem(Years, Now.Year)

            For i As Integer = 1 To 12
                Months.Items.Add(New ListItem(i & "月份", i))
            Next

            Months.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Common.SetListItem(Months, Now.Month)
            DecDay.Text = Common.FormatDate(Now)

            If start_date.Text = "" Then start_date.Text = Now.Year & "/" & Now.Month & "/1"
            'SDate_month.Value = start_date.Text '放入hidden欄位來keep value
            If end_date.Text = "" Then end_date.Text = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Now.Year & "/" & Now.Month & "/1")))  '本月底
            ''FDate_month.Value = end_date.Text
            If s_date.Text = "" Then s_date.Text = Now.Year & "/" & Now.Month & "/1"  '本月初
            If e_date.Text = "" Then e_date.Text = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Now.Year & "/" & Now.Month & "/1")))  '本月底

            Button3_Click(sender, e)
        End If

        TR1.Style("display") = "none"
        TR2.Style("display") = "none"
        TR3.Style("display") = "none"
        TR4.Style("display") = "none"
        TR5.Style("display") = "none"
        TR6.Style("display") = "none"

        Select Case RadioButtonList1.SelectedIndex
            Case 0
                TR1.Style("display") = TIMS.cst_inline1 '"inline"
            Case 1
                TR2.Style("display") = TIMS.cst_inline1 '"inline"
            Case 2
                TR2.Style("display") = TIMS.cst_inline1 '"inline"
                TR3.Style("display") = TIMS.cst_inline1 '"inline"
                TR5.Style("display") = TIMS.cst_inline1 '"inline"
            Case 3
                TR6.Style("display") = TIMS.cst_inline1 '"inline"
            Case 4
                TR4.Style("display") = TIMS.cst_inline1 '"inline"
        End Select
    End Sub

    Sub GetOrgID()
        If Me.RIDValue.Value = "" Then Exit Sub
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " SELECT b.OrgID " & vbCrLf
        sqlstr += " FROM dbo.AUTH_RELSHIP a " & vbCrLf
        sqlstr += " JOIN dbo.ORG_ORGINFO b ON a.orgid = b.orgid " & vbCrLf
        sqlstr += " WHERE a.RID = @RID " & vbCrLf
        Dim sCmd As New SqlCommand(sqlstr, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
            OrgidValue.Value = Convert.ToString(.ExecuteScalar())
        End With
    End Sub

    '檢查 作息時間設定  (若未設定)依層級往上找  班級-->機構(分區)-->機構(未分區)
    Function ChkRestTimeSet(ByVal sqlstr As String, ByVal sChkType As String) As DataTable
        Dim sCmd As New SqlCommand(sqlstr, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            '.Parameters.Clear()
            Dim myParam As Hashtable = New Hashtable
            Select Case sChkType
                Case cst_vsbclass 'by 班級
                    '.Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                    '.Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.sm.UserInfo.DistID
                    '.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    '.Parameters.Add("OrgID", SqlDbType.VarChar).Value = Me.OrgidValue.Value
                    myParam.Add("OCID", OCIDValue1.Value)
                    myParam.Add("DistID", Me.sm.UserInfo.DistID)
                    myParam.Add("RID", RIDValue.Value)
                    myParam.Add("OrgID", Me.OrgidValue.Value)
                Case cst_vsbrid 'by 機構
                    '.Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.sm.UserInfo.DistID
                    '.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    '.Parameters.Add("OrgID", SqlDbType.VarChar).Value = Me.OrgidValue.Value
                    myParam.Add("DistID", Me.sm.UserInfo.DistID)
                    myParam.Add("RID", RIDValue.Value)
                    myParam.Add("OrgID", Me.OrgidValue.Value)
            End Select
            'dt.Load(.ExecuteReader())
            dt = DbAccess.GetDataTable(sqlstr, objconn, myParam)
        End With
        Return dt
    End Function

    '檢查 作息時間設定  (若未設定)依層級往上找  班級-->機構(分區)-->機構(未分區)
    Function ChkSeting() As String
        'Dim SetStr As String = ""
        'Dim dt As New DataTable
        'Dim dr As DataRow
        'Dim IsSeted As Boolean = False
        'Dim SqlNotSet, SqlClass, SqlRid, SqlOrg As String

        'by 班級(by ocid)
        Dim sqlClass As String = ""
        sqlClass = ""
        sqlClass &= " SELECT f.ocid, f.orgid, b.DistID, b.OrgLevel, b.Relship, b.rid, c.orgname "
        sqlClass &= " FROM Org_OrgPlanInfo a " & vbCrLf
        sqlClass &= " JOIN Auth_Relship b ON a.RSID = b.RSID " & vbCrLf
        sqlClass &= " JOIN Org_OrgInfo c ON c.OrgID = b.OrgID " & vbCrLf
        sqlClass &= " JOIN Class_ClassInfo d ON b.RID = d.RID " & vbCrLf
        sqlClass &= " JOIN Class_RestTime f ON f.orgid = c.orgid AND f.ocid = d.ocid " & vbCrLf
        sqlClass &= " WHERE 0=0 "
        sqlClass &= " AND f.OrgID = @OrgID "
        sqlClass &= " AND f.ocid = @OCID "
        sqlClass &= " AND b.RID = @RID "
        sqlClass &= " AND b.DistID = @DistID "
        sqlClass &= " GROUP BY f.ocid, f.orgid, b.DistID, b.OrgLevel, b.Relship, b.rid, c.orgname "

        'by 機構(分區 by rid)  檢查是否Class_RestTime是否有該planid的該筆rid資料
        Dim sqlRID As String = ""
        sqlRID = ""
        sqlRID &= " SELECT f.orgid, b.DistID, b.OrgLevel, b.Relship, b.rid, c.orgname "
        sqlRID &= " FROM Org_OrgPlanInfo a "
        sqlRID &= " JOIN Auth_Relship b ON a.RSID = b.RSID "
        sqlRID &= " JOIN Org_OrgInfo c ON c.OrgID = b.OrgID "
        sqlRID &= " JOIN Class_RestTime f ON f.orgid = c.orgid "
        sqlRID &= " WHERE 0=0 "
        sqlRID &= " AND f.TPlanID = '" & Me.sm.UserInfo.TPlanID & "' "
        sqlRID &= " AND f.OrgID = @OrgID "
        sqlRID &= " AND f.RID = @RID "
        sqlRID &= " AND f.ocid IS NULL "
        sqlRID &= " AND b.DistID = @DistID "
        sqlRID &= " GROUP BY f.ocid, f.orgid, b.DistID, b.OrgLevel, b.Relship, b.rid, c.orgname "

        Dim vsSetingBy As String = cst_vsbNotSet  '未設定時
        If Me.OCIDValue1.Value <> "" Then  'by 班級  
            Dim dt As DataTable = ChkRestTimeSet(sqlClass, cst_vsbclass)
            If dt.Rows.Count > 0 Then
                vsSetingBy = cst_vsbclass
            Else
                dt = ChkRestTimeSet(sqlRID, cst_vsbrid)
                If dt.Rows.Count > 0 Then vsSetingBy = cst_vsbrid  'by 機構(分區)
            End If
        End If

        Dim setStr As String = ""
        Select Case vsSetingBy
            Case cst_vsbclass
                setStr = "&RestBy=1"
            Case cst_vsbrid
                setStr = "&RestBy=2"
            Case Else 'cst_vsbNotSet "NotSet"
                setStr = "&RestBy=1"
        End Select

        Return setStr
    End Function

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Errmsg &= "請選擇班級職類!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Errmsg &= "請選擇班級職類!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.ocid, 'x' FROM CLASS_CLASSINFO cc WHERE 1=1 " & vbCrLf
        sql &= "    AND cc.ocid = '" & OCIDValue1.Value & "' " & vbCrLf
        sql &= "    AND EXISTS (" & vbCrLf
        sql &= " 	   SELECT 'x' FROM Plan_CostItem x WHERE 1=1 " & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投--本功能目前應該未開放給產投，但先行設計 by AMU 2012/06
            sql &= " AND x.costid = '09' " & vbCrLf
            sql &= " AND x.CostMode = '5' " & vbCrLf
        Else
            'TIMS
            sql &= " AND x.costid = '21' " & vbCrLf
            sql &= " AND x.CostMode <> '5' " & vbCrLf
        End If

        sql &= " 	AND x.planid = cc.planid AND x.comidno = cc.comidno AND x.seqno = cc.seqno " & vbCrLf
        sql &= " ) " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            'TIMS '有使用 Key_CostItem:21 : 鐘點費(術科 - 助教)
            '產投 '有使用 Key_CostItem2:09 :術科助教費用
            dt = Nothing
            Dim flag_can_pass_07 As Boolean = False '須要pass true:能pass/false:不能pass 
            If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_can_pass_07 = True

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '產投
                sql = "" & vbCrLf
                sql &= " SELECT cc.ocid, 'x' FROM class_classinfo cc " & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= " AND cc.ocid = '" & OCIDValue1.Value & "' " & vbCrLf
                sql &= " AND EXISTS (" & vbCrLf
                sql &= " 	SELECT 'x' FROM Plan_TrainDesc x WHERE x.techid2 IS NOT NULL " & vbCrLf
                sql &= " 	AND x.planid = cc.planid AND x.comidno = cc.comidno AND x.seqno = cc.seqno " & vbCrLf
                sql &= " ) " & vbCrLf
                If Not flag_can_pass_07 Then
                    dt = DbAccess.GetDataTable(sql, objconn)
                End If
            Else
                'TIMS
                sql = "" & vbCrLf
                sql &= " SELECT cc.ocid, 'x' FROM class_classinfo cc " & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                sql &= " AND cc.ocid = '" & OCIDValue1.Value & "' " & vbCrLf
                sql &= " AND EXISTS (" & vbCrLf
                sql &= " 	SELECT 'x' FROM VIEW_CLASS_SCHEDULE x WHERE x.techid2 IS NOT NULL " & vbCrLf
                sql &= " 	AND x.ocid = cc.ocid " & vbCrLf
                sql &= " 	AND x.ocid = '" & OCIDValue1.Value & "' " & vbCrLf
                sql &= " ) " & vbCrLf
                If Not flag_can_pass_07 Then
                    dt = DbAccess.GetDataTable(sql, objconn)
                End If
            End If

            If Not flag_can_pass_07 Then
                If dt IsNot Nothing AndAlso dt.Rows.Count = 0 Then
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '產投  '課表中 未排入助教 
                        Errmsg += "本班有申請「術科助教費用」項目，但課程表內，未排入助教，故無法列印。" & vbCrLf
                    Else
                        'TIMS  '課表中 未排入助教 
                        Errmsg += "本班有申請「鐘點費(術科-助教)」項目，但課程表內，未排入助教，故無法列印。" & vbCrLf
                    End If
                End If
            End If
        End If

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "開始日期有誤。" & vbCrLf
            Else
                start_date.Text = Common.FormatDate(start_date.Text)
            End If
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "結束日期有誤。" & vbCrLf
            Else
                end_date.Text = Common.FormatDate(end_date.Text)
            End If
        End If
        If Errmsg <> "" Then Return False

        sql = " SELECT Formal FROM CLASS_SCHEDULE WHERE OCID = '" & OCIDValue1.Value & "'"
        Dim dtC As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim flag_Formal_Y As Boolean = True
        If dtC.Select("Formal='N'").Length > 0 Then flag_Formal_Y = False
        If Not flag_Formal_Y Then
            '尚未排入正式課程
            Errmsg += "尚未排入正式課程，不可列印。" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        If Not CheckData1(Errmsg) Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim sql As String = ""
        sql = " SELECT TPERIOD FROM CLASS_CLASSINFO WHERE OCID = '" & OCIDValue1.Value & "' "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim dr As DataRow

        Dim str_HourRan As String = ""
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            str_HourRan = Convert.ToString(dr("TPeriod"))
        End If

        'SELECT * FROM Key_HourRan
        Dim rptfileName As String = "" '預設列印報表
        Select Case str_HourRan
            Case "01", "05" '日間,列印1-8
                rptfileName = cst_printFN1 '"course_list_1_8_ya"
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "course_list_1_8", sMyValue)
            Case "02" '晚上,列印9-12
                rptfileName = cst_printFN2 '"course_list_9_12_ya"
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "course_list_9_12", sMyValue)
            Case Else '全日,假日,列印1-12
                rptfileName = cst_printFN3 '"course_list_ya"
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "course_list", sMyValue)
        End Select

        Call GetOrgID()
        'Dim SetStr As String = ""
        '20081105 坐息時間改為 (若未設定)依層級往上找 
        Dim SetStr As String = ChkSeting()

        Dim sMyValue As String = ""
        Select Case RadioButtonList1.SelectedValue
            Case "0" '全期
                'Dim date_end As String = Common.FormatDate(DateAdd(DateInterval.Day, 1, CDate(end_date.Text)))
                'Dim date_end As String = Common.FormatDate(CDate(end_date.Text)) '月底
                Dim date_end As String = end_date.Text '月底
                sMyValue = ""
                sMyValue &= "RID=" & Me.RIDValue.Value
                sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                sMyValue &= "&OCID=" & Me.OCIDValue1.Value
                sMyValue &= "&old_SchoolDate_start=" & start_date.Text 'Common.FormatDate(start_date.Text)
                sMyValue &= "&old_SchoolDate_end=" & end_date.Text 'Common.FormatDate(end_date.Text)
                '檢查當日為星期幾。並帶到星期1
                sMyValue &= "&start_date=" & start_date.Text '& check_date(start_date.Text, objconn)
                sMyValue &= "&end_date=" & date_end
                'sMyValue &= SetStr
            Case "1" '年
                Dim start_year As String = Common.FormatDate(Years.SelectedValue & "/01/01")
                Dim end_year As String = Common.FormatDate(DateAdd(DateInterval.Year, 1, CDate(start_year)))
                sMyValue = ""
                sMyValue &= "RID=" & Me.RIDValue.Value
                sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                sMyValue &= "&OCID=" & Me.OCIDValue1.Value
                sMyValue &= "&old_SchoolDate_start=" & Common.FormatDate(start_year)
                sMyValue &= "&old_SchoolDate_end=" & Common.FormatDate(DateAdd(DateInterval.Day, -1, CDate(end_year))) '年底
                sMyValue &= "&start_date=" & start_year
                sMyValue &= "&end_date=" & end_year
                'sMyValue &= SetStr
            Case "2" '月
                Dim start_month As String = Common.FormatDate(Years.SelectedValue & "/" & Months.SelectedValue & "/1")
                Dim end_month As String = Common.FormatDate(DateAdd(DateInterval.Month, 1, CDate(start_month)))
                Dim old_SchoolDate_end As String = Common.FormatDate(DateAdd(DateInterval.Day, -1, CDate(end_month))) '月底
                sMyValue = ""
                sMyValue &= "RID=" & Me.RIDValue.Value
                sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                sMyValue &= "&OCID=" & Me.OCIDValue1.Value
                sMyValue &= "&old_SchoolDate_start=" & start_month
                sMyValue &= "&old_SchoolDate_end=" & old_SchoolDate_end
                sMyValue &= "&start_date=" & start_month
                sMyValue &= "&end_date=" & old_SchoolDate_end
                'sMyValue &= SetStr
                Select Case str_HourRan
                    Case "01", "05" '日間,列印1-8
                        Select Case Me.Monthlist.SelectedValue
                            Case "1" '不分頁
                                rptfileName = cst_printFN1 '"course_list_1_8_ya" '保留助教
                                'rptfileName = "course_list_1_8" '去除助教
                            Case Else '"0" '分2頁
                                rptfileName = cst_printFN1B '"course_list_1_8_1"
                        End Select
                    Case "02" '晚上,列印9-12
                        Select Case Me.Monthlist.SelectedValue
                            Case "1" '不分頁
                                rptfileName = cst_printFN2 '"course_list_9_12_ya" '保留助教
                                'rptfileName = "course_list_9_12" '去除助教
                            Case Else '"0" '分2頁
                                rptfileName = cst_printFN2B '"course_list_9_12_1"
                        End Select
                    Case Else '全日,假日,列印1-12
                        Select Case Me.Monthlist.SelectedValue
                            Case "1" '不分頁
                                rptfileName = cst_printFN3 ' "course_list_ya" '保留助教
                                'rptfileName = "course_list" '去除助教
                            Case Else '"0" '分2頁
                                rptfileName = cst_printFN3B ' "course_list_R1"
                        End Select
                End Select
            Case "4" '週
                Dim date1_end As String = Common.FormatDate(DateAdd(DateInterval.Day, 1, CDate(e_date.Text)))
                sMyValue = ""
                sMyValue &= "RID=" & Me.RIDValue.Value
                sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                sMyValue &= "&OCID=" & Me.OCIDValue1.Value
                sMyValue &= "&PlanID=" & sm.UserInfo.PlanID
                sMyValue &= "&old_SchoolDate_start=" & Common.FormatDate(s_date.Text)
                sMyValue &= "&old_SchoolDate_end=" & Common.FormatDate(e_date.Text)
                'sMyValue &= "&s_date=" & check_date(s_date.Text, objconn)
                'sMyValue &= "&e_date=" & date1_end
                sMyValue &= "&start_date=" & check_date(s_date.Text, objconn)
                sMyValue &= "&end_date=" & date1_end
                'sMyValue &= SetStr
                Select Case str_HourRan
                    Case "01", "05" '日間,列印1-8
                        rptfileName = cst_printFN4 '"course_7_1_8"
                    Case "02" '晚上,列印9-12
                        rptfileName = cst_printFN5 '"course_list_7_9_12" '"course_7_9_12"
                    Case Else '全日,假日,列印1-12
                        rptfileName = cst_printFN6 '"course_7"
                End Select
            Case "3" '日
                Dim end_date_new As String = Common.FormatDate(DateAdd(DateInterval.Day, 1, CDate(DecDay.Text)))
                sMyValue = ""
                sMyValue &= "RID=" & Me.RIDValue.Value
                sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
                sMyValue &= "&OCID=" & Me.OCIDValue1.Value
                sMyValue &= "&PlanID=" & sm.UserInfo.PlanID
                sMyValue &= "&old_SchoolDate_start=" & Common.FormatDate(DecDay.Text)
                sMyValue &= "&old_SchoolDate_end=" & Common.FormatDate(DecDay.Text)
                sMyValue &= "&start_date=" & Common.FormatDate(DecDay.Text)
                sMyValue &= "&end_date=" & end_date_new
                'sMyValue &= SetStr
                Select Case str_HourRan
                    Case "01", "05" '日間,列印1-8
                        rptfileName = cst_printFN7 '"course_list_1_1_8"
                    Case "02" '晚上,列印9-12
                        rptfileName = cst_printFN8 '"course_list_1_9_12"
                    Case Else '全日,假日,列印1-12
                        rptfileName = cst_printFN9 '"course_list_1"
                End Select
            Case Else
                Common.MessageBox(Me, "未選擇 列印範圍。")
                Exit Sub
        End Select

        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, rptfileName, sMyValue)
    End Sub

    '檢查當日為星期幾。並帶到星期1
    Public Shared Function check_date(ByVal start_date_Text As String, ByRef objconn As SqlConnection) As String
        start_date_Text = TIMS.ClearSQM(start_date_Text)
        start_date_Text = TIMS.Cdate3(start_date_Text)
        Dim date_rst As String = ""
        Dim sql_date As String = " SELECT DATEPART(weekday, CONVERT(DATETIME, '" & start_date_Text & "', 111)) AS week "
        Dim iWeek As Integer = 0
        iWeek = DbAccess.ExecuteScalar(sql_date, objconn)

        '2:星期1
        If iWeek <> 2 Then
            Select Case iWeek
                Case 1 '星期日
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -6, CDate(start_date_Text)))
                Case 3 '星期二
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -1, CDate(start_date_Text)))
                Case 4 '星期三
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -2, CDate(start_date_Text)))
                Case 5 '星期四
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -3, CDate(start_date_Text)))
                Case 6 '星期五
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -4, CDate(start_date_Text)))
                Case 7 '星期六
                    date_rst = Common.FormatDate(DateAdd(DateInterval.Day, -5, CDate(start_date_Text)))
            End Select
        Else
            '星期1是第1天
            date_rst = Common.FormatDate(start_date_Text)
        End If

        Return date_rst
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        sql = " SELECT * FROM Class_ClassInfo WHERE OCID = '" & OCIDValue1.Value & "' "
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            start_date.Text = Common.FormatDate(dr("STDate"))
            end_date.Text = Common.FormatDate(dr("FTDate"))
            s_date.Text = Common.FormatDate(dr("STDate"))
            e_date.Text = Common.FormatDate(dr("FTDate"))
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class