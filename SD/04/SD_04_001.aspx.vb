Partial Class SD_04_001
    Inherits AuthBasePage

    Const cst_ssOCID As String = "_OCID"
    Const vs_TotalHours As String = "vsTotalHours"

    Const cst_errMsg1 As String = "班級資訊有誤，請重新選擇班級!!!"
    Const cst_errMsg2 As String = "班級資訊有誤，請重新操作該功能!!!"

    Const cst_ISintoMsg1 As String = "此班級已使用單月排課" 'ISinto.Text
    Const cst_ISintoMsg2 As String = "此班級已經結訓"
    Const cst_ISintoMsg3 As String = "，已審核確認"
    Const cst_ISintoMsg4 As String = "此班級已審核確認"
    Const cst_ISintoMsg5 As String = "載入班級課程"
    Const cst_ISintoMsg6 As String = "目前尚未排設任何的課程"
    Const cst_ISintoMsg7 As String = "已排入正式課程"
    Const cst_ISintoMsg8 As String = "尚未排入正式課程"
    Const cst_ISintoMsg9 As String = "尚未排入正式課程<BR>(實際時數必須預覽後才可得到正確數字)"
    Const cst_ISintoMsg10 As String = "尚未排入正式課程(權限可刪除)"

    Dim sMSG As String = ""

    Dim dtClsScheG As DataTable  '取得排課群組資料
    Dim dtClsScheG2 As DataTable  '取得排課群組資料2

    Dim sRIDValue As String = "" '紀錄查詢的RID
    Dim dtCourse As DataTable
    Dim dtTeach As DataTable

    'CLASS_TMPSCHEDULE
    'Dim HidClassStartDate.value As String = ""
    'Dim HidClassEndDate.value As String = ""
    'Dim HidClassHours.value As String = ""

    Dim vsformtable As Boolean = False
    Dim vTitle As String = ""

    Dim blnCanDeleteAuth As Boolean = False '特殊刪除權限。
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
        'blnCanDeleteAuth = TIMS.CheckAuthUse(sm.UserInfo.UserID, objconn, 1)
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '檢查帳號的功能權限-----------------------------------Start
        but_add.Enabled = True
        'If Not au.blnCanAdds Then
        '    but_add.Enabled = False
        '    TIMS.Tooltip(but_add, "您無權限使用該功能")
        'End If
        '檢查帳號的功能權限-----------------------------------End

        '增加判斷是否是自辦計劃------------   Start
        'Dim sql As String
        'Dim dr As DataRow
        If Not IsPostBack Then
            'btnDelX2.Visible = False
            RIDValue.Value = sm.UserInfo.RID 'sm.UserInfo.RID
            center.Text = sm.UserInfo.OrgName 'sm.UserInfo.OrgName
            Me.DataGrid1.Style("display") = "none"
        End If

        '載入課程師資資料
        Dim sql As String = ""
        sRIDValue = RIDValue.Value
        sql = " SELECT * FROM Course_CourseInfo WHERE OrgID IN (SELECT OrgID FROM Auth_Relship WHERE RID = '" & RIDValue.Value & "') "
        dtCourse = DbAccess.GetDataTable(sql, objconn)
        sql = " SELECT * FROM Teach_TeacherInfo WHERE RID = '" & RIDValue.Value & "' "
        dtTeach = DbAccess.GetDataTable(sql, objconn)

        '980224 fix只管一班則自動帶入
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');GETvalue();"
        Button6.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        sql = "SELECT PlanKind FROM ID_Plan WHERE PlanID=" & sm.UserInfo.PlanID & " "
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr("PlanKind").ToString = "1" Then
            Button1.Attributes("onclick") = "choose_class(2);"
            Button2.Attributes("onclick") = "choose_class2(1);"
        Else
            Button1.Attributes("onclick") = "choose_class(1);"
            Button2.Attributes("onclick") = "choose_class2(1);"
        End If
        '增加判斷是否是自辦計劃------------   End
        print.Attributes("onclick") = ReportQuery.ReportScript(Me, "list", "Class_Schedule_Total", "OCID='+document.getElementById('OCIDValue1').value+'")

        Me.LoadIntoClass.CausesValidation = False

        '產生Script
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMID1", "TMIDValue1", True, , True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showHistory('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Call TIMS.CreateTeacherScript(Me, RIDValue.Value, objconn) '產生Script

        If Not IsPostBack Then
            Me.OCourseID.Attributes.Add("onDblClick", "javascript:Course('Add','');")
            'Me.OCourseID.Attributes("onchange") = "GetCourseID(this.value,'OCourseID','OCourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','ORoomID');"
            Me.OCourseID.Attributes("onclick") = "GetCourseID(this.value,'OCourseID','OCourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','ORoomID');"
            Me.OLessonTeah1.Attributes.Add("onDblClick", "javascript:LessonTeah1('Add','OLessonTeah1','OLessonTeah1Value');")
            Me.OLessonTeah2.Attributes.Add("onDblClick", "javascript:LessonTeah2('Add','OLessonTeah2','OLessonTeah2Value');")
            Me.OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');"
            Me.OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
            Me.but_add.Attributes.Add("Onclick", "return CheckAddcourse();")
            Me.intoSort.Attributes.Add("OnClick", "return chkintosort();") '排入正式課程

            Me.Cancelinto.Attributes.Add("OnClick", "return chkcancelsort();")
            Me.Button5.Attributes("onclick") = "return confirm('這樣會刪除預排的課程資料，但不會刪除排課列表，確定要刪除?\n\n[提示]:如果您打算改用單月排課作業進行排課，請選擇[確定]。')"
            Me.LoadIntoClass.Attributes("onclick") = "return confirm('這樣將會刪除 [排課班級] 原來的課程資料，載入您所選擇 的 [載入班級] 排課資料，是否確定!!')"

            'objconn = DbAccess.GetConnection()
            'Dim objstr As String
            If OCIDValue1.Value = "" Then
                If Me.Request("OCIDValue1") <> "" Then
                    OCIDValue1.Value = Me.Request("OCIDValue1")
                    Me.OCID1.Text = Me.Request("OCID1")
                End If
            End If

            '從排課列表回上一頁到本頁時，產生基本資料
            If Not Session(cst_ssOCID) Is Nothing Then
                dr = TIMS.GetOCIDDate(Session(cst_ssOCID))
                TMID1.Value = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
                OCID1.Text = Convert.ToString(dr("ClassCName2")) 'dr("ClassCName").ToString
                'If Int(dr("CyclType")) <> 0 Then OCID1.Text += "第" & Int(dr("CyclType")) & "期"
                TMIDValue1.Value = dr("TMID").ToString
                OCIDValue1.Value = dr("OCID").ToString
                RIDValue.Value = dr("RID")
                center.Text = dr("OrgName").ToString

                'Course_CourseInfo = DbAccess.GetDataTable("SELECT * FROM Course_CourseInfo WHERE OrgID IN (SELECT OrgID FROM Auth_Relship WHERE RID='" & RIDValue.Value & "')")
                'Teach_TeacherInfo = DbAccess.GetDataTable("SELECT * FROM Teach_TeacherInfo WHERE RID IN (SELECT RID FROM Auth_Relship WHERE RID='" & RIDValue.Value & "')")
                '載入課程師資資料
                If sRIDValue <> RIDValue.Value Then
                    sRIDValue = RIDValue.Value
                    sql = " SELECT * FROM Course_CourseInfo WHERE OrgID IN (SELECT OrgID FROM Auth_Relship WHERE RID = '" & RIDValue.Value & "') "
                    dtCourse = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT * FROM Teach_TeacherInfo WHERE RID = '" & RIDValue.Value & "' "
                    dtTeach = DbAccess.GetDataTable(sql, objconn)
                End If

                Session(cst_ssOCID) = Nothing
            End If
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If

        Call CreateClass()

        Button4.Attributes("onclick") = "document.form1.OCID2.value='';document.form1.OCIDValue2.value='';"
        LoadIntoClass.Attributes("onclick") = "return CheckImportData();"
    End Sub

    '建立班級基本資料 SQL
    Sub CreateClass()
        If Not (OCIDValue1.Value <> "" AndAlso OCID.Value <> OCIDValue1.Value) Then
            If OCIDValue1.Value = "" Then
                Me.DataGrid1.Style("display") = "none"
                Me.NoFormal.Enabled = False
                Me.intoSort.Enabled = False
                Me.Cancelinto.Enabled = False
                print.Enabled = False
                Button5.Enabled = False
                TIMS.Tooltip(NoFormal, "班級尚未選擇", True)
                TIMS.Tooltip(intoSort, "班級尚未選擇", True)
                TIMS.Tooltip(Cancelinto, "班級尚未選擇", True)
                TIMS.Tooltip(print, "班級尚未選擇", True)
                TIMS.Tooltip(Button5, "班級尚未選擇", True)
            End If
            Exit Sub
        End If

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        OCID.Value = OCIDValue1.Value
        '若單月排已排，則不準排批次
        sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = " & OCIDValue1.Value
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Select("Type=2").Length <> 0 Then
            Me.but_add.Enabled = False
            Me.intoSort.Enabled = False
            Me.NoFormal.Enabled = False
            Me.Cancelinto.Enabled = False
            Me.LoadIntoClass.Enabled = False
            ISinto.Text = cst_ISintoMsg1 '"此班級已使用單月排課"
            print.Enabled = False

            TIMS.Tooltip(but_add, Me.ISinto.Text, True)
            TIMS.Tooltip(intoSort, Me.ISinto.Text, True)
            TIMS.Tooltip(NoFormal, Me.ISinto.Text, True)
            TIMS.Tooltip(Cancelinto, Me.ISinto.Text, True)
            TIMS.Tooltip(LoadIntoClass, Me.ISinto.Text, True)
            TIMS.Tooltip(print, Me.ISinto.Text, True)

            Me.DataGrid1.Style("display") = "none"
            Me.ClassYear.Text = ""
            Me.CyclType.Text = ""
            Me.HourRan.Text = ""
            Me.TranJob.Text = ""
            Me.ClassStart.Text = ""
            Me.ClassEnd.Text = ""
            Me.ClassHours.Text = ""
            Me.Totals.Text = ""
        Else
            sql = "" & vbCrLf
            sql &= " SELECT a.OCID ,a.planid ,a.comidno ,a.seqno " & vbCrLf
            sql &= " ,ip.Years ,a.CyclType " & vbCrLf
            sql &= " ,a.TPeriod ,a.THours " & vbCrLf
            sql &= " ,a.STDate ,a.FTDate " & vbCrLf
            sql &= " ,a.IsClosed " & vbCrLf
            sql &= " ,b.TrainID ,b.TrainName ,c.HourRanName " & vbCrLf
            sql &= " FROM Class_ClassInfo a " & vbCrLf
            sql &= " JOIN id_plan ip ON ip.planid = a.planid " & vbCrLf
            sql &= " JOIN Key_TrainType b ON a.TMID = b.TMID " & vbCrLf
            sql &= " JOIN key_HourRan c ON a.TPeriod = c.HRID " & vbCrLf
            sql &= " WHERE a.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
            If dr Is Nothing Then
                '班級資訊有誤
                Common.MessageBox(Page, cst_errMsg2)
                Exit Sub
            End If

            Session("Class_CourseName") = TIMS.Get_PNamePlanTrainDesc(dr("planid"), dr("comidno"), dr("seqno"), objconn)

            'Me.ClassYear.Text = "20" & dr.Item("Years")
            Me.ClassYear.Text = dr.Item("Years")
            Me.CyclType.Text = "第" & Int(dr.Item("CyclType")) & "期"
            Me.TranJob.Text = "[" & dr("TrainID") & "]" & dr.Item("TrainName")
            Me.HourRan.Text = dr.Item("HourRanName")
            Me.ClassType.Value = dr.Item("TPeriod")
            Me.ClassHours.Text = dr.Item("THours")
            Me.ClassStart.Text = dr.Item("STDate")
            Me.ClassEnd.Text = dr.Item("FTDate")

            HidClassStartDate.Value = Convert.ToString(dr.Item("STDate"))
            HidClassEndDate.Value = Convert.ToString(dr.Item("FTDate"))
            HidClassHours.Value = Convert.ToString(dr.Item("THours"))

            Me.OStartDate.Attributes("ondblclick") = "openCalendar('OStartDate','" & ClassStart.Text & "','" & ClassEnd.Text & "',this.value);"
            Me.OEndDate.Attributes("ondblclick") = "openCalendar('OEndDate','" & ClassStart.Text & "','" & ClassEnd.Text & "',this.value);"

            '訓練時段(日間、晚上、全日制、假日)
            Call GetCourseData(OCIDValue1.Value)

            '假如結訓，則不可以修改資料
            If dr("IsClosed").ToString = "Y" Then
                ISinto.Text = cst_ISintoMsg2 '"此班級已經結訓"

                NoFormal.Enabled = False
                intoSort.Enabled = False
                Cancelinto.Enabled = False
                Button5.Enabled = False
                Me.LoadIntoClass.Enabled = False '不可再載入班級
                DataGrid1.Columns(17).Visible = False
                DataGrid1.Columns(18).Visible = False
                TIMS.Tooltip(NoFormal, Me.ISinto.Text, True)
                TIMS.Tooltip(intoSort, Me.ISinto.Text, True)
                TIMS.Tooltip(Cancelinto, Me.ISinto.Text, True)
                TIMS.Tooltip(Button5, Me.ISinto.Text, True)
                TIMS.Tooltip(LoadIntoClass, Me.ISinto.Text, True)
            End If
        End If

        'Cancelinto '假如審核確認，則不可以修改資料
        If TIMS.Chk_ClassSchVerify(OCIDValue1.Value, objconn) Then
            If ISinto.Text <> "" Then
                ISinto.Text &= cst_ISintoMsg3 '"，已審核確認"
            Else
                ISinto.Text &= cst_ISintoMsg4 '"此班級已審核確認"
            End If
            NoFormal.Enabled = False
            intoSort.Enabled = False
            Cancelinto.Enabled = False
            Button5.Enabled = False

            Me.LoadIntoClass.Enabled = False '不可再載入班級
            DataGrid1.Columns(17).Visible = False
            DataGrid1.Columns(18).Visible = False
            'Cancelinto.ToolTip = "此班級已審核確認"

            TIMS.Tooltip(NoFormal, cst_ISintoMsg4, True)
            TIMS.Tooltip(intoSort, cst_ISintoMsg4, True)
            TIMS.Tooltip(Cancelinto, cst_ISintoMsg4, True)
            TIMS.Tooltip(Button5, cst_ISintoMsg4, True)
            TIMS.Tooltip(LoadIntoClass, cst_ISintoMsg4, True)
        End If

        If Me.OCIDValue2.Value <> "" Then
            Call GetCourseData(OCIDValue2.Value)

            Me.but_add.Enabled = False
            Me.intoSort.Enabled = False
            Me.Cancelinto.Enabled = False
            print.Enabled = False
            ISinto.Text = cst_ISintoMsg5 '"載入班級課程"

            DataGrid1.Columns(17).Visible = False
            DataGrid1.Columns(18).Visible = False

            TIMS.Tooltip(but_add, Me.ISinto.Text, True)
            TIMS.Tooltip(intoSort, Me.ISinto.Text, True)
            TIMS.Tooltip(Cancelinto, Me.ISinto.Text, True)
            TIMS.Tooltip(print, Me.ISinto.Text, True)
        End If

        'If TIMS.sUtl_ChkTest() Then '測試環境開啟
        '    'Me.but_add.Enabled = True
        '    Me.NoFormal.Enabled = True
        '    Me.intoSort.Enabled = True
        '    Me.Cancelinto.Enabled = True
        '    Me.print.Enabled = True
        '    Me.Button5.Enabled = True
        '    TIMS.Tooltip(NoFormal, "測試環境開啟")
        '    TIMS.Tooltip(intoSort, "測試環境開啟")
        '    TIMS.Tooltip(Cancelinto, "測試環境開啟")
        '    TIMS.Tooltip(print, "測試環境開啟")
        '    TIMS.Tooltip(Button5, "測試環境開啟")
        'End If

        If ISinto.Text = cst_ISintoMsg6 Then
            'btnDelX2.Visible = True '提供刪除
            Button5.Enabled = True
            TIMS.Tooltip(Button5, "刪除預排資料")
        End If
    End Sub

    ''' <summary>
    ''' 建立課程DataGrid1 [SQL]
    ''' </summary>
    ''' <param name="OCID"></param>
    Sub GetCourseData(ByVal OCID As String)
        If OCID = "" Then Exit Sub

        Me.ViewState(vs_TotalHours) = 0
        '載入班級
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM CLASS_TMPSCHEDULE"
        sql &= " WHERE OCID = " & CStr(OCID) & vbCrLf
        sql &= " ORDER BY ItemID ,StartDate ,EndDate " & vbCrLf
        sql &= " ,s1 ,s2 ,s3 ,s4 ,s5 ,s6 ,s7 " & vbCrLf
        sql &= " ,e1 ,e2 ,e3 ,e4 ,e5 ,e6 ,e7 " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            ISinto.Text = cst_ISintoMsg6 '"目前尚未排設任何的課程" '系統訊息
            Me.Totals.Text = 0
            print.Enabled = False
            TIMS.Tooltip(print, Me.ISinto.Text, True)
            Me.DataGrid1.Style("display") = "none"
            Exit Sub
        End If

        print.Enabled = True
        TIMS.Tooltip(print, "")
        Me.DataGrid1.Style("display") = ""

        '取得排課群組資料
        dtClsScheG = TIMS.Get_ClassScheduleG(OCID, objconn)
        dtClsScheG2 = TIMS.Get_ClassScheduleG2(OCID, objconn)

        Me.DataGrid1.DataSource = dt
        Me.DataGrid1.DataKeyField = "CTSID"
        Me.DataGrid1.DataBind()

        sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCID & "' ORDER BY SchoolDate "
        Dim dtC As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim flag_Formal_Y As Boolean = False
        Dim flag_Formal_N As Boolean = False
        If dtC.Select("[Formal]='Y'").Length > 0 Then flag_Formal_Y = True
        If dtC.Select("[Formal]='N'").Length > 0 Then flag_Formal_N = True

        If flag_Formal_Y Then
            ISinto.Text = cst_ISintoMsg7 '"已排入正式課程"

            '重建時間配當
            sql = " SELECT * FROM PLAN_SCHEDULE WHERE OCID = '" & OCID & "' ORDER BY PSID "
            Dim dtP As DataTable = DbAccess.GetDataTable(sql, objconn)
            If dtP.Rows.Count = 0 Then Me.Button7.Enabled = True

            Me.but_add.Enabled = False
            Me.intoSort.Enabled = False
            Me.NoFormal.Enabled = False
            TIMS.Tooltip(but_add, Me.ISinto.Text, True)
            TIMS.Tooltip(intoSort, Me.ISinto.Text, True)
            TIMS.Tooltip(NoFormal, Me.ISinto.Text, True)
            'Me.LoadIntoClass.Enabled = False
            Me.Cancelinto.Enabled = False

            'If sm.UserInfo.LID <> 2 Then
            If sm.UserInfo.LID <> 2 Then
                '署(局)、分署(中心)
                Me.Cancelinto.Enabled = True
                'If au.blnCanDel = True Then Me.Cancelinto.Enabled = True
            Else
                '委訓單位
                'mark by nick 060511
                'Me.Cancelinto.Enabled = False
                Me.Cancelinto.Enabled = True
            End If
            vsformtable = True
            DataGrid1.Columns(17).Visible = False
            DataGrid1.Columns(18).Visible = False
            Button5.Enabled = False
        End If

        If flag_Formal_N OrElse Not flag_Formal_Y Then
            ISinto.Text = cst_ISintoMsg8 '"尚未排入正式課程"

            Me.but_add.Enabled = True
            Me.intoSort.Enabled = True
            Me.NoFormal.Enabled = True
            Me.LoadIntoClass.Enabled = True
            Me.Cancelinto.Enabled = False
            TIMS.Tooltip(Cancelinto, cst_ISintoMsg8, True)
            vsformtable = False
            DataGrid1.Columns(17).Visible = True
            DataGrid1.Columns(18).Visible = True

            If dtC.Select("[Formal]='N'").Length <> 0 Then
                Button5.Enabled = True
                ISinto.Text = cst_ISintoMsg8 '"尚未排入正式課程"
                TIMS.Tooltip(Button5, cst_ISintoMsg10, True)
            Else
                Button5.Enabled = False
                ISinto.Text = cst_ISintoMsg9 '"尚未排入正式課程<BR>(實際時數必須預覽後才可得到正確數字)"
                TIMS.Tooltip(Button5, cst_ISintoMsg8, True)
            End If
        End If

    End Sub

    '檢查可能錯誤資料
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        ORecycle.Text = TIMS.ClearSQM(ORecycle.Text)
        'If ORecycle.Text <> "" Then ORecycle.Text = Trim(ORecycle.Text)
        If ORecycle.Text <> "" Then
            'ORecycle.Text = ORecycle.Text.Trim
            Try
                ORecycle.Text = CInt(ORecycle.Text)
                If CInt(ORecycle.Text) <= 0 Then
                    Errmsg += "循環數字有誤 只可輸入大於0的整數數字,請重新確認!" & vbCrLf
                End If
            Catch ex As Exception
                Errmsg += "循環格式有誤 只可輸入數字,請重新確認!" & vbCrLf
            End Try
        End If

        Me.OCourseIDValue.Value = TIMS.ClearSQM(Me.OCourseIDValue.Value)
        'If Me.OCourseIDValue.Value <> "" Then Me.OCourseIDValue.Value = Trim(Me.OCourseIDValue.Value)
        If Me.OCourseIDValue.Value = "" Then Errmsg += "課程資料未輸入,請重新確認!" & vbCrLf

        Me.OLessonTeah1Value.Value = TIMS.ClearSQM(Me.OLessonTeah1Value.Value)
        'If Me.OLessonTeah1Value.Value <> "" Then Me.OLessonTeah1Value.Value = Trim(Me.OLessonTeah1Value.Value)
        If Me.OLessonTeah1Value.Value = "" Then Errmsg += "教師資料為必填,請重新選擇!" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '新增按鈕
    Private Sub but_add_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles but_add.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Dim objstr, chkstr, ClassSDate, ClassEDate, FlowSDate, FlowEDate As String
        'Dim objstr, chkstr, ClassSDate, ClassEDate As String
        'Dim ClassTHours, i, j As Integer
        'Dim FlowTHours As Object
        'Dim weekclass(7, 2) As Integer
        'Dim objtext As TextBox
        'Dim objAdapter As SqlDataAdapter
        'Dim objtable As DataTable
        'Dim dbrow, objrow, chkrow As DataRow
        'Dim sql As String
        Dim weekclass(7, 2) As Integer
        Dim iCalhoursAll As Integer '時數加總
        Dim iCalhoursAll2 As Integer '時數加總2

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇班級!")
            Exit Sub
        End If

        Dim ClassSDate As String = ""
        Dim ClassEDate As String = ""
        Dim ClassTHours As String = ""
        Dim oFlowTHours As Object = Nothing

        '先記算開結訓日期
        Dim sql As String = ""
        sql = " SELECT STDATE ,FTDATE ,THOURS FROM CLASS_CLASSINFO WHERE OCID = " & OCIDValue1.Value
        Dim dbrow As DataRow
        dbrow = DbAccess.GetOneRow(sql, objconn)

        ClassSDate = TIMS.Cdate3(dbrow.Item("STDate"))
        ClassEDate = TIMS.Cdate3(dbrow.Item("FTDate"))
        ClassTHours = dbrow.Item("THours")
        'FlowSDate = Me.OStartDate.Text
        'FlowEDate = Me.OEndDate.Text
        Me.OStartDate.Text = TIMS.ClearSQM(Me.OStartDate.Text)
        Me.OEndDate.Text = TIMS.ClearSQM(Me.OEndDate.Text)
        Me.OCalHours.Text = TIMS.ClearSQM(Me.OCalHours.Text)
        oFlowTHours = DBNull.Value
        If Me.OCalHours.Text <> "" Then oFlowTHours = CInt(Val(Me.OCalHours.Text))

        For i As Integer = 1 To 7
            Dim objtext As TextBox = Page.FindControl("OS" & i)
            If objtext.Text = "" Then
                weekclass(i, 0) = -1
            Else
                weekclass(i, 0) = CInt(objtext.Text)
            End If
            objtext = Page.FindControl("OE" & i)
            If objtext.Text = "" Then
                weekclass(i, 1) = -1
            Else
                weekclass(i, 1) = CInt(objtext.Text)
            End If
        Next

        Dim chkrow As DataRow = Nothing
        sql = " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & " AND ItemID = " & CDbl(Me.OItemID.Text)
        chkrow = DbAccess.GetOneRow(sql, objconn)
        If Not chkrow Is Nothing Then
            Common.MessageBox(Page, "項次重覆，請重新輸入!!!")
            Exit Sub
        End If
        '----------------------------------------------檢查時數加總不能大於訓練時數

        sql = " SELECT ISNULL(SUM(CALHOURS),0) CALHOURSALL FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & ""
        iCalhoursAll = DbAccess.ExecuteScalar(sql, objconn)
        iCalhoursAll2 = iCalhoursAll + OCalHours.Text
        If iCalhoursAll2 > dbrow.Item("THours") Then
            Common.MessageBox(Page, "[排課時數]加總大於[訓練時數]，請重新輸入!!!")
            Exit Sub
        End If
        '---------------------------------------------- end

        Call TIMS.OpenDbConn(objconn)

        Dim objrow As DataRow = Nothing
        Dim objAdapter As SqlDataAdapter = Nothing
        Dim objtable As DataTable = Nothing
        Dim objstr As String = ""
        objstr = " SELECT * FROM CLASS_TMPSCHEDULE WHERE 1<>1 "
        objtable = DbAccess.GetDataTable(objstr, objAdapter, objconn)
        objrow = objtable.NewRow()
        objtable.Rows.Add(objrow)

        objrow("CTSID") = DbAccess.GetNewId(objconn, "CLASS_TMPSCHEDULE_CTSID_SEQ,CLASS_TMPSCHEDULE,CTSID")
        objrow("OCID") = OCIDValue1.Value
        objrow("ItemID") = CDbl(Me.OItemID.Text)
        objrow("CourseID") = Me.OCourseIDValue.Value
        objrow("CalHours") = oFlowTHours '期望排課時數
        objrow("StartDate") = If(OStartDate.Text = "", CDate(ClassSDate), CDate(OStartDate.Text))
        objrow("EndDate") = If(OEndDate.Text = "", CDate(ClassEDate), CDate(OEndDate.Text))
        objrow("RoomID") = Me.ORoomID.Text
        objrow("LessonTeah1") = Me.OLessonTeah1Value.Value
        objrow("LessonTeah2") = Me.OLessonTeah2Value.Value
        For i As Integer = 1 To 7
            If weekclass(i, 0) < 1 Then
                objrow("S" & i) = DBNull.Value
            Else
                objrow("S" & i) = weekclass(i, 0)
            End If
            If weekclass(i, 1) < 1 Then
                objrow("E" & i) = DBNull.Value
            Else
                objrow("E" & i) = weekclass(i, 1)
            End If
        Next
        objrow("RealHours") = CountRealHour(OItemID, OCalHours, OStartDate, OEndDate, OS1, OE1, OS2, OE2, OS3, OE3, OS4, OE4, OS5, OE5, OS6, OE6, OS7, OE7)
        objrow("Recycle") = If(ORecycle.Text <> "", ORecycle.Text, Convert.DBNull)
        objrow("ModifyAcct") = sm.UserInfo.UserID 'sm.UserInfo.UserID
        objrow("ModifyDate") = Now()
        DbAccess.UpdateDataTable(objtable, objAdapter)

        Call ClearItem()
        Call GetCourseData(OCIDValue1.Value) '建立課程DataGrid1
    End Sub

    '計算時計時數
    Function CountRealHour(ByVal OItemID As TextBox, ByVal OCalHours As TextBox, ByVal OStartDate As TextBox, ByVal OEndDate As TextBox, ByVal OS1 As TextBox, ByVal OE1 As TextBox, ByVal OS2 As TextBox, ByVal OE2 As TextBox, ByVal OS3 As TextBox, ByVal OE3 As TextBox, ByVal OS4 As TextBox, ByVal OE4 As TextBox, ByVal OS5 As TextBox, ByVal OE5 As TextBox, ByVal OS6 As TextBox, ByVal OE6 As TextBox, ByVal OS7 As TextBox, ByVal OE7 As TextBox) As Integer
        Dim dr As DataRow
        Dim sql As String
        Dim dt As DataTable
        Dim Holiday As DataTable
        Dim iLeaveHour As Integer
        Dim iUsedHour As Integer

        Dim BDate As String 'BEGIN
        Dim ODate As String 'OVER
        Dim ClassSDate As String 'START
        Dim ClassEDate As String 'END
        'Dim BDate, ODate, ClassSDate, ClassEDate As String
        'Dim CalHours, RealHours, ClassTHours, TotalDays, i As Integer
        Dim iCalHours As Integer = 0
        Dim iClassTHours As Integer = 0

        '取出目前已經排入的課程
        sql = " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & " AND ItemID < '" & OItemID.Text & "' "
        dt = DbAccess.GetDataTable(sql, objconn)

        '計算已經使用過的時數
        sql = " SELECT SUM(RealHours) UsedHour FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & " AND ItemID < '" & OItemID.Text & "' "
        dr = DbAccess.GetOneRow(sql, objconn)
        iUsedHour = If(IsDBNull(dr("UsedHour")), 0, dr("UsedHour"))

        '取出開結訓日、總訓練時數
        sql = " SELECT STDate ,FTDate ,THours FROM CLASS_CLASSINFO WHERE OCID = " & OCIDValue1.Value
        dr = DbAccess.GetOneRow(sql, objconn)

        ClassSDate = TIMS.Cdate3(dr("STDate"))
        ClassEDate = TIMS.Cdate3(dr("FTDate"))
        iClassTHours = dr("THours")
        iLeaveHour = iClassTHours - iUsedHour

        '假如沒有填入課程日期區間，則帶開結訓日
        If OStartDate.Text = "" Then BDate = ClassSDate Else BDate = TIMS.Cdate3(OStartDate.Text)
        If OEndDate.Text = "" Then ODate = ClassEDate Else ODate = TIMS.Cdate3(OEndDate.Text)

        '假如沒有填入課程訓練時數，則帶訓練時數
        OCalHours.Text = TIMS.ClearSQM(OCalHours.Text)
        If OCalHours.Text = "" Then
            iCalHours = iLeaveHour
        Else
            iCalHours = If(Int(OCalHours.Text) > iLeaveHour, iLeaveHour, Int(OCalHours.Text))
        End If

        '計算課程日期的天數
        Dim TotalDays As Integer = 0
        TotalDays = DateDiff(DateInterval.Day, CDate(BDate), CDate(ODate)) + 1

        '實際課程日期的天數
        Dim iRealHours As Integer = 0
        iRealHours = 0

        '取出假日資料
        'sql = " SELECT * FROM Sys_Holiday WHERE RID = '" & sm.UserInfo.RID & "' "
        sql = " SELECT * FROM Sys_Holiday WHERE RID = '" & sm.UserInfo.RID & "' "
        Holiday = DbAccess.GetDataTable(sql, objconn)

        For i As Integer = 1 To TotalDays
            '取出排課日期是否為假日，則是跳過
            If Holiday.Select("HolDate = '" & BDate & "'").Length = 0 Then
                Select Case Weekday(CDate(BDate))
                    Case 2
                        If OE1.Text <> "" Then iRealHours += CInt(OE1.Text) - CInt(OS1.Text) + 1
                    Case 3
                        If OE2.Text <> "" Then iRealHours += CInt(OE2.Text) - CInt(OS2.Text) + 1
                    Case 4
                        If OE3.Text <> "" Then iRealHours += CInt(OE3.Text) - CInt(OS3.Text) + 1
                    Case 5
                        If OE4.Text <> "" Then iRealHours += CInt(OE4.Text) - CInt(OS4.Text) + 1
                    Case 6
                        If OE5.Text <> "" Then iRealHours += CInt(OE5.Text) - CInt(OS5.Text) + 1
                    Case 7
                        If OE6.Text <> "" Then iRealHours += CInt(OE6.Text) - CInt(OS6.Text) + 1
                    Case 1
                        If OE7.Text <> "" Then iRealHours += CInt(OE7.Text) - CInt(OS7.Text) + 1
                End Select
            End If
            BDate = CStr(DateAdd(DateInterval.Day, 1, CDate(BDate)))

            '假如已經到結訓日期，則結束回圈
            If CDate(BDate) > CDate(ClassEDate) Then Exit For
            If iRealHours >= iCalHours Then
                iRealHours = iCalHours
                Exit For
            End If
        Next

        Return iRealHours
    End Function

    ''' <summary>
    ''' 排入正式課程 'SetFormal()
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub intoSort_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles intoSort.Click
#Region "(No Use)"

        'Call SetFormal()
        'Page.RegisterStartupScript("JS1", "<script>alert('系統正在排入課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=Y'</script>")

        ''Call SetFormal()
        'If blnTest Then
        '    Call SetFormal()
        '    Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在排入課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=Y'</script>")
        '    Exit Sub
        'End If

        'Dim myThreadDelegate As New ThreadStart(AddressOf SetFormal)
        'Dim myThread As New Thread(myThreadDelegate)
        'myThread.Start()
        'Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在排入課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=Y'</script>")

#End Region
    End Sub

    ''' <summary>
    ''' 排入正式課程(隱藏觸發) 'SetFormal() -intoSort_Click()
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BTNintoSort_Click(sender As Object, e As System.EventArgs) Handles BTNintoSort.Click
        Call SetFormal()

        intoSort.Enabled = True
        Page.RegisterStartupScript("JS1", "<script>alert('系統正在排入課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=Y'</script>")
    End Sub

    ''' <summary>
    ''' 載入排課班級
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub LoadIntoClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadIntoClass.Click
        'Dim ocid_same As Boolean = False '不可相同
        If OCIDValue1.Value <> "" OrElse Me.OCIDValue2.Value <> "" Then
            If OCIDValue1.Value = Me.OCIDValue2.Value Then
                'ocid_same = True '相同
                Common.MessageBox(Me, "載入班級不可與排課班級相同!!")
                Exit Sub
            End If
        End If

        Dim errmsg1 As String = ""
        ClassStart.Text = TIMS.Cdate3(ClassStart.Text)
        ClassEnd.Text = TIMS.Cdate3(ClassEnd.Text)
        If ClassStart.Text = "" Then errmsg1 &= "排課起日，日期格式有誤!!" & vbCrLf
        If ClassEnd.Text = "" Then errmsg1 &= "排課迄日，日期格式有誤!!" & vbCrLf
        If errmsg1 <> "" Then
            Common.MessageBox(Me, errmsg1)
            Exit Sub
        End If

        If OCIDValue1.Value <> "" AndAlso Me.OCIDValue2.Value <> "" Then
            '刪除資料
            Call TIMS.sUtl_DeleteClassTmpSchedule(OCIDValue1.Value, objconn)
            Call TIMS.SUtl_DeleteClassSchedule(Me, OCIDValue1.Value, objconn)
            Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn)

            '載入班級
            Dim sql As String = ""
            sql = ""
            sql &= " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & Me.OCIDValue2.Value & vbCrLf
            sql &= " ORDER BY ItemID ,StartDate ,EndDate " & vbCrLf
            sql &= " ,s1 ,s2 ,s3 ,s4 ,s5 ,s6 ,s7 " & vbCrLf
            sql &= " ,e1 ,e2 ,e3 ,e4 ,e5 ,e6 ,e7 " & vbCrLf
            Dim table1 As DataTable = DbAccess.GetDataTable(sql, objconn)

            '排課班級
            Call TIMS.OpenDbConn(objconn)
            Dim str2 As String = ""
            Dim table2 As DataTable = Nothing
            Dim adapter2 As SqlDataAdapter = Nothing
            str2 = " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value
            table2 = DbAccess.GetDataTable(str2, adapter2, objconn)
            For Each row1 As DataRow In table1.Rows
                Dim row2 As DataRow = table2.NewRow
                table2.Rows.Add(row2)
                row2("CTSID") = DbAccess.GetNewId(objconn, "CLASS_TMPSCHEDULE_CTSID_SEQ,CLASS_TMPSCHEDULE,CTSID")
                For i As Integer = 1 To table1.Columns.Count - 1
                    Select Case table1.Columns(i).ColumnName
                        Case "OCID"
                            row2("OCID") = OCIDValue1.Value
                            'Case "StartDate"
                        Case "STARTDATE"
                            row2("StartDate") = ClassStart.Text
                            'Case "EndDate"
                        Case "ENDDATE"
                            row2("EndDate") = ClassEnd.Text
                            'Case "RealHours"
                        Case "REALHOURS"
                            row2("RealHours") = 0
                        Case Else
                            row2(table1.Columns(i).ColumnName) = row1(table1.Columns(i).ColumnName)
                    End Select
                Next
            Next
            DbAccess.UpdateDataTable(table2, adapter2)

            Me.OCID2.Text = ""
            Me.OCIDValue2.Value = ""
            Me.but_add.Enabled = True
            Me.intoSort.Enabled = True

            DataGrid1.Columns(17).Visible = True
            DataGrid1.Columns(18).Visible = True

            Call GetCourseData(OCIDValue1.Value)
        End If
    End Sub

    '清除
    Private Sub ClearItem()
        Me.OItemID.Text = ""
        Me.OCourseID.Text = ""
        Me.OCalHours.Text = ""
        Me.OStartDate.Text = ""
        Me.OEndDate.Text = ""
        Me.ORoomID.Text = ""
        Me.OLessonTeah1.Text = ""
        Me.OLessonTeah2.Text = ""
        Me.OE1.Text = ""
        Me.OS1.Text = ""
        Me.OE2.Text = ""
        Me.OS2.Text = ""
        Me.OE3.Text = ""
        Me.OS3.Text = ""
        Me.OE4.Text = ""
        Me.OS4.Text = ""
        Me.OE5.Text = ""
        Me.OS5.Text = ""
        Me.OE6.Text = ""
        Me.OS6.Text = ""
        Me.OE7.Text = ""
        Me.OS7.Text = ""
        Me.ORecycle.Text = ""
    End Sub

    '取消正式課程
    Private Sub Cancelinto_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancelinto.Click
        'Dim delstr As String
        'delstr = " DELETE Class_Schedule WHERE OCID = " & OCIDValue1.Value
        'DbAccess.ExecuteNonQuery(delstr, objconn)

        ''刪除 時間配當預定進度檔
        'delstr = " DELETE Plan_schedule WHERE OCID = " & OCIDValue1.Value
        'DbAccess.ExecuteNonQuery(delstr, objconn)

        '取消正式課程
        Call TIMS.SUtl_DeleteClassSchedule(Me, OCIDValue1.Value, objconn)

        '刪除 時間配當預定進度檔
        Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn)

        Me.intoSort.Enabled = True
        Me.Cancelinto.Enabled = False
        TIMS.Tooltip(Cancelinto, "已取消正式課程", True)
        Me.NoFormal.Enabled = True

        Common.AddClientScript(Page, "alert('正式課表已完全刪除!!!');")
        DataGrid1.Columns(17).Visible = True
        DataGrid1.Columns(18).Visible = True

        ClearItem()
        GetCourseData(OCIDValue1.Value)
    End Sub

    ''' <summary>
    ''' 建立排課列表
    ''' </summary>
    ''' <param name="Formal"></param>
    Public Sub SetClassView(ByVal Formal As String)
        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then Exit Sub
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, cst_errMsg1)
            Exit Sub
        End If

        '刪除正式資料，排課再重算
        Call TIMS.SUtl_DeleteClassSchedule(Me, OCIDValue1.Value, objconn)
        Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn)

        Dim STDate As Date                  '開訓日期
        Dim FTDate As Date                  '結訓日期
        Dim sql As String = ""
        sql = " SELECT STDate ,FTDate FROM CLASS_CLASSINFO WHERE OCID = " & OCIDValue1.Value
        Dim dr As DataRow
        dr = DbAccess.GetOneRow(sql, objconn)
        STDate = dr("STDate")
        FTDate = dr("FTDate")

        Dim CourseDataTable As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        'Dim SUBsql As String = ""
        'SUBsql = " SELECT CourID ,CourseName ,MainCourID FROM Course_CourseInfo "
        'CourseDataTable = DbAccess.GetDataTable(SUBsql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT CourID ,CourseName ,MainCourID FROM COURSE_COURSEINFO p " & vbCrLf
        sql &= " WHERE EXISTS ( " & vbCrLf
        sql &= "    SELECT 'x' FROM AUTH_RELSHIP c " & vbCrLf
        sql &= " 	WHERE EXISTS ( " & vbCrLf
        sql &= " 	    SELECT 'x' FROM AUTH_RELSHIP c2 " & vbCrLf
        sql &= " 		WHERE c2.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= " 		AND c2.OrgID = c.OrgID " & vbCrLf
        sql &= "    ) " & vbCrLf
        sql &= " 	AND c.RID = p.RID " & vbCrLf
        sql &= " ) " & vbCrLf
        CourseDataTable = DbAccess.GetDataTable(sql, objconn)

        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'Plan_Schedule
            '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
            SD_04_002.AddNew_Plan_Schedule(Me, STDate, FTDate, OCIDValue1.Value, dt2, da2, trans, CourseDataTable)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try
    End Sub

#Region "(No Use)"

    'Dim dt As DataTable
    'Dim da As SqlDataAdapter = nothing
    'Dim TempDate As Date
    'Dim i As Integer
    'Dim dr1 As DataRow
    'Dim CourseTable As DataTable                '課程資料
    'sql = "SELECT * FROM Course_CourseInfo WHERE OrgID='" & sm.UserInfo.OrgID & "'"
    'CourseTable = DbAccess.GetDataTable(sql)

    ''建立時間配當表前四項目
    ''1.週數-----------------------------Start
    'dr = dt.NewRow
    'dt.Rows.Add(dr)
    'dr("OCID") = OCIDValue1.Value
    'dr("TitleItem") = 1
    'i = 1
    'TempDate = STDate
    'While TempDate <= FTDate
    '    dr("W" & i) = i
    '    i += 1
    '    TempDate = TempDate.AddDays(7)
    'End While
    'dr("ModifyAcct") = sm.UserInfo.UserID
    'dr("ModifyDate") = Now
    ''1.週數-----------------------------End

    ''2.月份-----------------------------Start
    'dr = dt.NewRow
    'dt.Rows.Add(dr)
    'dr("OCID") = OCIDValue1.Value
    'dr("TitleItem") = 2
    'i = 1
    'TempDate = STDate
    'While TempDate <= FTDate
    '    dr("W" & i) = TempDate.Month
    '    i += 1
    '    TempDate = TempDate.AddDays(7)
    'End While
    'dr("ModifyAcct") = sm.UserInfo.UserID
    'dr("ModifyDate") = Now
    ''2.月份-----------------------------End

    ''3.起日-----------------------------Start
    'dr = dt.NewRow
    'dt.Rows.Add(dr)
    'dr("OCID") = OCIDValue1.Value
    'dr("TitleItem") = 3
    'i = 1
    'TempDate = STDate
    'While TempDate <= FTDate
    '    dr("W" & i) = TempDate.Day
    '    i += 1
    '    TempDate = TempDate.AddDays(7)
    'End While
    'dr("ModifyAcct") = sm.UserInfo.UserID
    'dr("ModifyDate") = Now
    ''3.起日-----------------------------End

    ''4.結束日-----------------------------Start
    'dr = dt.NewRow
    'dt.Rows.Add(dr)
    'dr("OCID") = OCIDValue1.Value
    'dr("TitleItem") = 4
    'i = 1
    'TempDate = STDate.AddDays(6)
    'While TempDate <= FTDate.AddDays(6)
    '    If TempDate > FTDate Then
    '        TempDate = FTDate
    '    End If
    '    dr("W" & i) = TempDate.Day
    '    i += 1
    '    TempDate = TempDate.AddDays(7)
    'End While
    'dr("ModifyAcct") = sm.UserInfo.UserID
    'dr("ModifyDate") = Now
    ''4.結束日-----------------------------End

    'Dim CourID As String
    'i = 1
    'TempDate = STDate
    'While TempDate <= FTDate
    '    For Each dr1 In dt1.Select("SchoolDate>='" & TempDate & "' and SchoolDate<='" & TempDate.AddDays(6) & "'")
    '        For j As Integer = 1 To 12
    '            If Not IsDBNull(dr1("Class" & j)) Then
    '                CourID = GetMainCourse(dr1("Class" & j), CourseTable)

    '                If dt.Select("CourID='" & CourID & "'").Length = 0 Then
    '                    dr = dt.NewRow
    '                    dt.Rows.Add(dr)
    '                    dr("OCID") = OCIDValue1.Value
    '                    dr("CourID") = CourID
    '                    dr("TitleItem") = 5
    '                    dr("ModifyAcct") = sm.UserInfo.UserID
    '                    dr("ModifyDate") = Now
    '                Else
    '                    dr = dt.Select("CourID='" & CourID & "'")(0)
    '                End If

    '                If IsDBNull(dr("W" & i)) Then
    '                    dr("W" & i) = 1
    '                Else
    '                    dr("W" & i) += 1
    '                End If
    '            End If
    '        Next
    '    Next

    '    i += 1
    '    TempDate = TempDate.AddDays(7)
    'End While

    'DbAccess.UpdateDataTable(dt, da)

#End Region

    '建立Plan_Schedule表頭
    Sub SetPlanSchedule()
        Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn) '刪除 時間配當預定進度檔

        Dim dr As DataRow = Nothing
        'Dim dt1 As DataTable
        Dim STDate As Date
        Dim FTDate As Date
        Dim sql As String = ""
        sql = " SELECT STDate ,FTDate FROM Class_ClassInfo WHERE OCID = " & OCIDValue1.Value
        dr = DbAccess.GetOneRow(sql, objconn)
        STDate = dr("STDate")
        FTDate = dr("FTDate")

        Dim CourseDataTable As DataTable = Nothing '課程資料
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        'Dim SUBsql As String = "" '課程資料sql
        'SUBsql = "SELECT CourID,CourseName,MainCourID FROM Course_CourseInfo "
        'CourseDataTable = DbAccess.GetDataTable(SUBsql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT CourID ,CourseName ,MainCourID FROM Course_CourseInfo p " & vbCrLf
        sql &= " WHERE EXISTS ( " & vbCrLf
        sql &= "    SELECT 'x' FROM AUTH_RELSHIP c " & vbCrLf
        sql &= " 	WHERE EXISTS (" & vbCrLf
        sql &= " 	    SELECT 'x' FROM AUTH_RELSHIP c2 " & vbCrLf
        sql &= " 		WHERE c2.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= " 		   AND c2.OrgID = c.OrgID " & vbCrLf
        sql &= " 	   ) " & vbCrLf
        sql &= " 	  AND c.RID = p.RID " & vbCrLf
        sql &= " ) " & vbCrLf
        CourseDataTable = DbAccess.GetDataTable(sql, objconn)

        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'Plan_Schedule
            '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
            SD_04_002.AddNew_Plan_Schedule(Me, STDate, FTDate, OCIDValue1.Value, dt2, da2, trans, CourseDataTable)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try
    End Sub

    ''' <summary>
    ''' 建立正式課程
    ''' </summary>
    Public Sub SetFormal()
        Dim dbrow As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value) '
        If dbrow Is Nothing Then
            '班級資訊有誤
            Common.MessageBox(Page, cst_errMsg1)
            Exit Sub
        End If
        Call SetClassSchedule("Y")
        Call SetPlanSchedule()
    End Sub

    ''' <summary>
    ''' 建立預覽課程
    ''' </summary>
    Public Sub SetView()
        Dim dbrow As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value) '
        If dbrow Is Nothing Then
            '班級資訊有誤
            Common.MessageBox(Page, cst_errMsg1)
            Exit Sub
        End If
        Call SetClassSchedule("N") '先記算開結訓日期
    End Sub

    '建立預覽課程 SQL
    Function SetClassSchedule(ByVal Formal As String) As Boolean
        Dim rst As Boolean = False '異常
        If Convert.ToString(sm.UserInfo.RID) = "" Then Return False
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Return False

        '假別資料
        Dim sql As String = ""
        sql = " SELECT * FROM dbo.SYS_HOLIDAY WHERE RID = '" & sm.UserInfo.RID & "' "
        Dim Holiday As DataTable = DbAccess.GetDataTable(sql, objconn)

        '先刪除已經排入課程的資料
        If Formal = "N" Then
            Call TIMS.sUtl_DeleteClassScheduleFN(Me, OCIDValue1.Value, objconn)
        Else
            'Y
            Call TIMS.SUtl_DeleteClassSchedule(Me, OCIDValue1.Value, objconn)
        End If

        '刪除 時間配當預定進度檔
        Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn)

        sql = " SELECT STDate ,FTDate ,THours FROM CLASS_CLASSINFO WHERE OCID = '" & OCIDValue1.Value & "' "
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        Dim STDate As Date = CDate(dr("STDate"))
        Dim FTDate As Date = CDate(dr("FTDate"))
        Dim iTotalHour As Integer = dr("THours")

        '載入班級
        sql = ""
        sql &= " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & vbCrLf
        sql &= " ORDER BY ItemID ,StartDate ,EndDate " & vbCrLf
        sql &= " ,s1 ,s2 ,s3 ,s4 ,s5 ,s6 ,s7 " & vbCrLf
        sql &= " ,e1 ,e2 ,e3 ,e4 ,e5 ,e6 ,e7 " & vbCrLf
        Dim dt1 As DataTable = Nothing
        Dim da1 As SqlDataAdapter = Nothing
        dt1 = DbAccess.GetDataTable(sql, da1, objconn)

        Call TIMS.OpenDbConn(objconn)
        sql = " SELECT * FROM CLASS_SCHEDULE WHERE 1<>1 "
        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        dt2 = DbAccess.GetDataTable(sql, da2, objconn)
        While STDate <= FTDate
            Dim dr2 As DataRow = dt2.NewRow
            dt2.Rows.Add(dr2)
            'CLASS_SCHEDULE_CSID_SEQ
            dr2("CSID") = DbAccess.GetNewId(objconn, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
            dr2("OCID") = OCIDValue1.Value
            dr2("SchoolDate") = STDate
            dr2("Formal") = Formal
            dr2("Type") = 1
            dr2("ModifyAcct") = sm.UserInfo.UserID 'sm.UserInfo.UserID
            dr2("ModifyDate") = Now
            STDate = STDate.AddDays(1)
        End While

        Dim iRealHours As Integer = 0 '計算已排課時數
        Dim iCalHours As Integer = 0 '期望排課時數
        Dim iWeekIndex As Integer = 0
        For Each dr1 As DataRow In dt1.Rows '暫存排課
            iRealHours = 0
            iCalHours = dr1("CalHours") ''期望排課時數
            iWeekIndex = 1 '每週循環

            '正式課表 (依排課區間取出) 'Class_Schedule
            For Each dr2 As DataRow In dt2.Select("SchoolDate>='" & dr1("StartDate") & "' and SchoolDate<='" & dr1("EndDate") & "'", "SchoolDate")
                '假日--判斷
                If Holiday.Select("HolDate='" & dr2("SchoolDate") & "'").Length = 0 Then
                    'If CDbl(dr1("ItemID")) = CDbl(29) Then
                    '    Dim XXX As String = "X"'TEST
                    '    XXX = "XXX"
                    'End If
                    Dim Recycle As Boolean = True '循環判斷(繼續排課)
                    If dr1("Recycle").ToString <> "" Then
                        If dr1("Recycle") <> 0 Then
                            If iWeekIndex Mod Int(dr1("Recycle")) <> 1 AndAlso Int(dr1("Recycle")) <> 1 Then Recycle = False
                        End If
                    End If
                    If Recycle = True Then
                        For i As Integer = 1 To 7 '星期1~7
                            If Not IsDBNull(dr1("S" & i)) AndAlso (Weekday(dr2("SchoolDate")) = i + 1 OrElse Weekday(dr2("SchoolDate")) = i - 6) Then
                                For j As Integer = Val(dr1("S" & i)) To Val(dr1("E" & i)) '節次
                                    If iTotalHour > 0 AndAlso iCalHours <> 0 Then '有總時數與'期望排課時數(繼續排課)
                                        If j > 12 OrElse j < 1 Then
                                            Common.MessageBox(Me, "班級排課節次超過系統範圍，資訊有誤，請重新確認資料!!!")
                                            Return rst
                                        End If

                                        '尚未排課
                                        If IsDBNull(dr2("Class" & j)) Then
                                            dr2("Class" & j) = dr1("CourseID")
                                            dr2("Teacher" & j) = dr1("LessonTeah1")
                                            dr2("Teacher" & j + 12) = dr1("LessonTeah2")
                                            dr2("Room" & j) = dr1("RoomID")
                                            iRealHours += 1 '實際+1'計算已排課時數
                                            iCalHours -= 1 '期望排課時數-1
                                            iTotalHour -= 1
                                        End If
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If

                If Weekday(dr2("SchoolDate")) = 7 Then iWeekIndex += 1 '下週加1
            Next

            dr1("RealHours") = iRealHours '已排課時數
        Next

        DbAccess.UpdateDataTable(dt1, da1)
        DbAccess.UpdateDataTable(dt2, da2)
        rst = True '正常結束。
        Return rst
    End Function

    '預覽正式課程
    Private Sub NoFormal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NoFormal.Click
        ''Call SetView()
        Call SetView()
        Page.RegisterStartupScript("JS1", "<script>alert('系統正在預排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=N'</script>")
        'Dim blnTest As Boolean = False
        'If blnTest Then
        '    Call SetView()
        '    Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在預排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=N'</script>")
        '    Exit Sub
        'End If
        'Dim myThreadDelegate As New ThreadStart(AddressOf SetView)
        'Dim myThread As New Thread(myThreadDelegate)
        'myThread.Start()
        'Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在預排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & OCIDValue1.Value & "&Formal=N'</script>")
    End Sub

    '列印課程總表
    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        Dim cGuid1 As String = ReportQuery.GetGuid(Page)
        Dim Url As String = ReportQuery.GetUrl(Page)
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.open('" & Url & "GUID=" + cGuid1 + "&SQ_AutoLogout=true&sys=list&filename=Class_Schedule_Total&path=TIMS&OCID=" & OCIDValue1.Value & "');" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    Function GetCourseName(ByVal CourID As String) As String
        CourID = TIMS.ClearSQM(CourID)
        Dim rst As String = CourID
        If CourID <> "" Then
            Dim ff As String = "CourID='" & CourID & "'"
            If Not dtCourse Is Nothing Then
                If dtCourse.Select(ff).Length <> 0 Then rst = dtCourse.Select(ff)(0)("CourseName")
            End If
        End If
        Return rst
    End Function

    Function GetTeacherName(ByVal TechID As String) As String
        TechID = TIMS.ClearSQM(TechID)
        Dim rst As String = TechID
        If TechID <> "" AndAlso TechID <> " " Then
            Dim ff As String = "TechID='" & TechID & "'"
            If Not dtTeach Is Nothing Then
                If dtTeach.Select(ff).Length <> 0 Then rst = dtTeach.Select(ff)(0)("TeachCName")
            End If
        End If
        Return rst
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim IdList As String = ""

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As ImageButton = e.Item.FindControl("ImageButton1") '修改
                Dim btn2 As ImageButton = e.Item.FindControl("ImageButton4") '刪除
                Dim lbtnCopy1 As LinkButton = e.Item.FindControl("lbtnCopy1") 'COPY

                e.Item.Cells(0).Text = drv("ItemID").ToString
                'Dim Course_CourseInfo As DataTable
                'Dim Teach_TeacherInfo As DataTable
                e.Item.Cells(1).Text = GetCourseName(drv("CourseID").ToString)  'class
                e.Item.Cells(2).Text = drv("CalHours").ToString
                e.Item.Cells(3).Text = Common.FormatDate(drv("StartDate"))
                e.Item.Cells(4).Text = Common.FormatDate(drv("EndDate"))
                e.Item.Cells(5).Text = drv("RoomID").ToString
                e.Item.Cells(6).Text = GetTeacherName(drv("LessonTeah1").ToString)
                e.Item.Cells(7).Text = GetTeacherName(drv("LessonTeah2").ToString)
                e.Item.Cells(8).Text = drv("S1").ToString & "-" & drv("E1").ToString
                e.Item.Cells(9).Text = drv("S2").ToString & "-" & drv("E2").ToString
                e.Item.Cells(10).Text = drv("S3").ToString & "-" & drv("E3").ToString
                e.Item.Cells(11).Text = drv("S4").ToString & "-" & drv("E4").ToString
                e.Item.Cells(12).Text = drv("S5").ToString & "-" & drv("E5").ToString
                e.Item.Cells(13).Text = drv("S6").ToString & "-" & drv("E6").ToString
                e.Item.Cells(14).Text = drv("S7").ToString & "-" & drv("E7").ToString
                e.Item.Cells(15).Text = drv("Recycle").ToString

                'returnValue1(ItemID, CourseID, CalHours, StartDate, EndDate, RoomID, LessonTeah1, LessonTeah2, S1, E1, S2, E2, S3, E3, S4, E4, S5, E5, S6, E6, S7, E7, Recycle) {
                Dim sValue1 As String = ""
                sValue1 = ""
                sValue1 &= "'" & e.Item.Cells(0).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(1).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(2).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(3).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(4).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(5).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(6).Text & "'"
                sValue1 &= ",'" & e.Item.Cells(7).Text & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S1")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E1")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S2")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E2")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S3")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E3")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S4")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E4")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S5")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E5")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S6")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("E6")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S7")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("S7")) & "'"
                sValue1 &= ",'" & e.Item.Cells(15).Text & "'"
                sValue1 &= ",'" & Convert.ToString(drv("CourseID")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("LessonTeah1")) & "'"
                sValue1 &= ",'" & Convert.ToString(drv("LessonTeah2")) & "'"
                lbtnCopy1.Attributes("onclick") = "return returnValue1(" & sValue1 & ");"

                btn2.Attributes("onclick") = "return confirm('確定要刪除此課程?');"
                Me.ViewState(vs_TotalHours) += CInt(e.Item.Cells(16).Text)

                '期望排課時數 不等同 已排課時數
                If Val(drv("CalHours")) <> Val(drv("RealHours")) Then
                    e.Item.Cells(16).ForeColor = Color.Red
                    vTitle = ""
                    vTitle += "期望排課時數(" & Convert.ToString(drv("CalHours")) & ")"
                    vTitle += "不等同 已排課時數(" & Convert.ToString(drv("RealHours")) & ")" '& vbCrLf
                    vTitle += ",可能原因:1.排課節數與時數不符。2.轄區行事例放假。3.有衝堂問題(請查詢預覽排課)" '& vbCrLf
                    If Convert.ToString(drv("CourseID")) <> "" Then
                        Dim sCourseID As String = TIMS.ClearSQM(Convert.ToString(drv("CourseID")))
                        '取得排課群組資料
                        Dim sfilter As String = "class='" & sCourseID & "'"
                        If dtClsScheG.Select(sfilter).Length > 0 Then
                            Dim drG As DataRow = dtClsScheG.Select(sfilter)(0)
                            vTitle += vbCrLf & Convert.ToString(drG("coursename"))
                            vTitle += vbCrLf & "從" & Common.FormatDate(drG("min_schooldate"))
                            vTitle += "至" & Common.FormatDate(drG("max_schooldate"))
                            vTitle += ",共排" & Convert.ToString(drG("cnt1")) & "節課."
                        End If
                    End If
                    '為同一天顯示該天詳細資訊
                    If TIMS.Cdate3(drv("StartDate")) = TIMS.Cdate3(drv("EndDate")) Then
                        Dim sfilter As String = "schooldate='" & TIMS.Cdate3(drv("StartDate")) & "'"
                        If dtClsScheG2.Select(sfilter).Length > 0 Then
                            vTitle += vbCrLf & TIMS.Cdate3(drv("StartDate")) & ":"
                            Dim i As Integer = 0
                            For Each dr As DataRow In dtClsScheG2.Select(sfilter)
                                If i Mod 4 = 0 Then vTitle += vbCrLf
                                vTitle &= Convert.ToString(dr("lesson")) & "." & Convert.ToString(dr("coursename"))
                                i += 1
                            Next
                        End If
                    End If
                    TIMS.Tooltip(e.Item.Cells(16), vTitle)
                Else
                    vTitle = ""
                    If Convert.ToString(drv("CourseID")) <> "" Then
                        Dim sCourseID As String = TIMS.ClearSQM(Convert.ToString(drv("CourseID")))
                        '取得排課群組資料
                        Dim sfilter As String = "class='" & sCourseID & "'"
                        If dtClsScheG.Select(sfilter).Length > 0 Then
                            Dim drG As DataRow = dtClsScheG.Select(sfilter)(0)
                            vTitle += Convert.ToString(drG("coursename"))
                            vTitle += vbCrLf & "從" & Common.FormatDate(drG("min_schooldate"))
                            vTitle += "至" & Common.FormatDate(drG("max_schooldate"))
                            vTitle += ",共排" & Convert.ToString(drG("cnt1")) & "節課."
                        End If
                    End If
                    TIMS.Tooltip(e.Item.Cells(16), vTitle)
                End If
                btn1.Enabled = True
                'If au.blnCanMod = True Then
                '    btn1.Enabled = True
                'Else
                '    btn1.Enabled = False
                '    TIMS.Tooltip(btn1, "權限不可修改", True)
                'End If
                btn2.Enabled = True
                'If au.blnCanDel = True Then
                '    btn2.Enabled = True
                'Else
                '    btn2.Enabled = False
                '    TIMS.Tooltip(btn2, "權限不可刪除", True)
                'End If
            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim TItemID As TextBox = e.Item.FindControl("TItemID")
                Dim TCourseID As TextBox = e.Item.FindControl("TCourseID")
                Dim TCalHours As TextBox = e.Item.FindControl("TCalHours")
                Dim TStartDate As TextBox = e.Item.FindControl("TStartDate")
                Dim TEndDate As TextBox = e.Item.FindControl("TEndDate")
                Dim TRoomID As TextBox = e.Item.FindControl("TRoomID")
                Dim HidTLessonTeah1 As HiddenField = e.Item.FindControl("HidTLessonTeah1")
                Dim HidTLessonTeah2 As HiddenField = e.Item.FindControl("HidTLessonTeah2")
                Dim TLessonTeah1 As TextBox = e.Item.FindControl("TLessonTeah1")
                Dim TLessonTeah2 As TextBox = e.Item.FindControl("TLessonTeah2")
                Dim TS1 As TextBox = e.Item.FindControl("TS1")
                Dim TE1 As TextBox = e.Item.FindControl("TE1")
                Dim TS2 As TextBox = e.Item.FindControl("TS2")
                Dim TE2 As TextBox = e.Item.FindControl("TE2")
                Dim TS3 As TextBox = e.Item.FindControl("TS3")
                Dim TE3 As TextBox = e.Item.FindControl("TE3")
                Dim TS4 As TextBox = e.Item.FindControl("TS4")
                Dim TE4 As TextBox = e.Item.FindControl("TE4")
                Dim TS5 As TextBox = e.Item.FindControl("TS5")
                Dim TE5 As TextBox = e.Item.FindControl("TE5")
                Dim TS6 As TextBox = e.Item.FindControl("TS6")
                Dim TE6 As TextBox = e.Item.FindControl("TE6")
                Dim TS7 As TextBox = e.Item.FindControl("TS7")
                Dim TE7 As TextBox = e.Item.FindControl("TE7")
                Dim TRecycle As TextBox = e.Item.FindControl("TRecycle")
                Dim btn1 As ImageButton = e.Item.FindControl("ImageButton2")          '修改存檔
                Dim btn2 As ImageButton = e.Item.FindControl("ImageButton3")          '取消
                Dim btn3 As ImageButton = e.Item.FindControl("ImageButton4")          '刪除

                TItemID.Text = drv("ItemID").ToString
                OCourseIDValue.Value = drv("CourseID").ToString
                TCourseID.Text = GetCourseName(drv("CourseID").ToString)
                TCalHours.Text = drv("CalHours").ToString
                If Convert.ToString(drv("StartDate")) <> "" Then TStartDate.Text = Common.FormatDate(drv("StartDate"))
                If Convert.ToString(drv("EndDate")) <> "" Then TEndDate.Text = Common.FormatDate(drv("EndDate"))
                TRoomID.Text = drv("RoomID").ToString
                TLessonTeah1.Text = GetTeacherName(drv("LessonTeah1").ToString)
                TLessonTeah2.Text = GetTeacherName(drv("LessonTeah2").ToString)
                If TLessonTeah1.Text = "" Then HidTLessonTeah1.Value = ""
                If TLessonTeah2.Text = "" Then HidTLessonTeah2.Value = ""
                HidTLessonTeah1.Value = drv("LessonTeah1").ToString
                HidTLessonTeah2.Value = drv("LessonTeah2").ToString
                'OLessonTeah1Value.Value = drv("LessonTeah1").ToString
                'OLessonTeah2Value.Value = drv("LessonTeah2").ToString

                TS1.Text = drv("S1").ToString
                TE1.Text = drv("E1").ToString
                TS2.Text = drv("S2").ToString
                TE2.Text = drv("E2").ToString
                TS3.Text = drv("S3").ToString
                TE3.Text = drv("E3").ToString
                TS4.Text = drv("S4").ToString
                TE4.Text = drv("E4").ToString
                TS5.Text = drv("S5").ToString
                TE5.Text = drv("E5").ToString
                TS6.Text = drv("S6").ToString
                TE6.Text = drv("E6").ToString
                TS7.Text = drv("S7").ToString
                TE7.Text = drv("E7").ToString
                TRecycle.Text = drv("Recycle").ToString

                IdList = ""
                IdList = "'" & TItemID.ClientID & "'"
                IdList &= ",'" & TCourseID.ClientID & "'"
                IdList &= ",'" & TCalHours.ClientID & "'"
                IdList &= ",'" & TStartDate.ClientID & "'"
                IdList &= ",'" & TEndDate.ClientID & "'"
                IdList &= ",'" & TRoomID.ClientID & "'"
                IdList &= ",'" & TLessonTeah1.ClientID & "'"
                IdList &= ",'" & TLessonTeah2.ClientID & "'"
                IdList &= ",'" & TS1.ClientID & "'"
                IdList &= ",'" & TE1.ClientID & "'"
                IdList &= ",'" & TS2.ClientID & "'"
                IdList &= ",'" & TE2.ClientID & "'"
                IdList &= ",'" & TS3.ClientID & "'"
                IdList &= ",'" & TE3.ClientID & "'"
                IdList &= ",'" & TS4.ClientID & "'"
                IdList &= ",'" & TE4.ClientID & "'"
                IdList &= ",'" & TS5.ClientID & "'"
                IdList &= ",'" & TE5.ClientID & "'"
                IdList &= ",'" & TS6.ClientID & "'"
                IdList &= ",'" & TE6.ClientID & "'"
                IdList &= ",'" & TS7.ClientID & "'"
                IdList &= ",'" & TE7.ClientID & "'"

                'function GetCourseID(CourseID, TextField, ValueField, Tech1Field, TechName1Field, Tech2Field, TechName2Field, RoomField)
                'TCourseID.Attributes("onchange") = "GetCourseID(this.value,'" & TCourseID.ClientID & "','OCourseIDValue','OLessonTeah1Value','" & TLessonTeah1.ClientID & "','OLessonTeah2Value','" & TLessonTeah2.ClientID & "','" & TRoomID.ClientID & "');"
                TCourseID.Attributes("onclick") = "GetCourseID(this.value,'" & TCourseID.ClientID & "','OCourseIDValue','OLessonTeah1Value','" & TLessonTeah1.ClientID & "','OLessonTeah2Value','" & TLessonTeah2.ClientID & "','" & TRoomID.ClientID & "');"
                TCourseID.Attributes.Add("OnDblclick", "javascript:Course('Edit','" & TCourseID.ClientID & "');")
                'TIMS.CreateTeacherScript
                TLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','" & TLessonTeah1.ClientID & "');"
                TLessonTeah1.Attributes("ondblclick") = "javascript:LessonTeah1('Edit','" & TLessonTeah1.ClientID & "','" & OLessonTeah1Value.ClientID & "');"
                TLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','" & TLessonTeah2.ClientID & "');"
                TLessonTeah2.Attributes("ondblclick") = "javascript:LessonTeah2('Edit','" & TLessonTeah2.ClientID & "','" & OLessonTeah2Value.ClientID & "');"
                TStartDate.Attributes("ondblclick") = "openCalendar('" & TStartDate.ClientID & "','" & ClassStart.Text & "','" & ClassEnd.Text & "',this.value);"
                TEndDate.Attributes("ondblclick") = "openCalendar('" & TEndDate.ClientID & "','" & ClassStart.Text & "','" & ClassEnd.Text & "',this.value);"
                btn1.Attributes("onclick") = "return CheckCourse(" & IdList & ");"
                btn3.Attributes("onclick") = "return false;"
        End Select

        If CInt(Val(Me.ViewState(vs_TotalHours))) > CInt(Val(HidClassHours.Value)) Then
            Me.Totals.Text = "<font color=red>" & Val(Me.ViewState(vs_TotalHours)) & "</font>"
            vTitle = ""
            vTitle &= "總時數(" & Convert.ToString(Val(Me.ViewState(vs_TotalHours))) & ")"
            vTitle &= "大於 課程時數(" & Convert.ToString(Val(HidClassHours.Value)) & ")"
            TIMS.Tooltip(Me.Totals, vTitle)
        Else
            Me.Totals.Text = Val(Me.ViewState(vs_TotalHours))
            TIMS.Tooltip(Me.Totals, "")
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Dim da As SqlDataAdapter = nothing
        Select Case e.CommandName
            Case "edit"
                Me.DataGrid1.EditItemIndex = e.Item.ItemIndex
                hidTmpTItemID.Value = Me.DataGrid1.Items(e.Item.ItemIndex).Cells(0).Text

            Case "update" '修改儲存
                Dim chkrow As DataRow
                Dim chkstr As String = ""
                Dim TItemID As TextBox = e.Item.FindControl("TItemID")
                Dim TCalHours As TextBox = e.Item.FindControl("TCalHours") '期望排課時數
                Dim TStartDate As TextBox = e.Item.FindControl("TStartDate")
                Dim TEndDate As TextBox = e.Item.FindControl("TEndDate")
                Dim TRoomID As TextBox = e.Item.FindControl("TRoomID")
                Dim HidTLessonTeah1 As HiddenField = e.Item.FindControl("HidTLessonTeah1")
                Dim HidTLessonTeah2 As HiddenField = e.Item.FindControl("HidTLessonTeah2")
                'Dim TLessonTeah1 As TextBox = e.Item.FindControl("TLessonTeah1")
                'Dim TLessonTeah2 As TextBox = e.Item.FindControl("TLessonTeah2")
                Dim TS1 As TextBox = e.Item.FindControl("TS1")
                Dim TE1 As TextBox = e.Item.FindControl("TE1")
                Dim TS2 As TextBox = e.Item.FindControl("TS2")
                Dim TE2 As TextBox = e.Item.FindControl("TE2")
                Dim TS3 As TextBox = e.Item.FindControl("TS3")
                Dim TE3 As TextBox = e.Item.FindControl("TE3")
                Dim TS4 As TextBox = e.Item.FindControl("TS4")
                Dim TE4 As TextBox = e.Item.FindControl("TE4")
                Dim TS5 As TextBox = e.Item.FindControl("TS5")
                Dim TE5 As TextBox = e.Item.FindControl("TE5")
                Dim TS6 As TextBox = e.Item.FindControl("TS6")
                Dim TE6 As TextBox = e.Item.FindControl("TE6")
                Dim TS7 As TextBox = e.Item.FindControl("TS7")
                Dim TE7 As TextBox = e.Item.FindControl("TE7")
                Dim TRecycle As TextBox = e.Item.FindControl("TRecycle")

                TItemID.Text = TIMS.ClearSQM(TItemID.Text)
                TCalHours.Text = TIMS.ClearSQM(TCalHours.Text)
                TStartDate.Text = TIMS.ClearSQM(TStartDate.Text)
                TEndDate.Text = TIMS.ClearSQM(TEndDate.Text)
                TRoomID.Text = TIMS.ClearSQM(TRoomID.Text)
                HidTLessonTeah1.Value = TIMS.ClearSQM(HidTLessonTeah1.Value)
                HidTLessonTeah2.Value = TIMS.ClearSQM(HidTLessonTeah2.Value)
                OLessonTeah1Value.Value = TIMS.ClearSQM(OLessonTeah1Value.Value)
                TS1.Text = TIMS.ClearSQM(TS1.Text)
                TE1.Text = TIMS.ClearSQM(TE1.Text)
                TS2.Text = TIMS.ClearSQM(TS2.Text)
                TE2.Text = TIMS.ClearSQM(TE2.Text)
                TS3.Text = TIMS.ClearSQM(TS3.Text)
                TE3.Text = TIMS.ClearSQM(TE3.Text)
                TS4.Text = TIMS.ClearSQM(TS4.Text)
                TE4.Text = TIMS.ClearSQM(TE4.Text)
                TS5.Text = TIMS.ClearSQM(TS5.Text)
                TE5.Text = TIMS.ClearSQM(TE5.Text)
                TS6.Text = TIMS.ClearSQM(TS6.Text)
                TE6.Text = TIMS.ClearSQM(TE6.Text)
                TS7.Text = TIMS.ClearSQM(TS7.Text)
                TE7.Text = TIMS.ClearSQM(TE7.Text)
                TRecycle.Text = TIMS.ClearSQM(TRecycle.Text)

                Dim rtnPath As String = ""
                rtnPath = Request.FilePath
                If Convert.ToString(Me.DataGrid1.DataKeys(e.Item.ItemIndex)) = "" Then
                    Common.MessageBox(Me, "請重新查詢有效資料", rtnPath)
                    Exit Sub
                End If
                If TIMS.CheckInput(Convert.ToString(Me.DataGrid1.DataKeys(e.Item.ItemIndex))) Then
                    Common.MessageBox(Me, TIMS.cst_ErrorMsg2, rtnPath)
                    Exit Sub
                End If

                Dim dr As DataRow = Nothing
                Dim da As New SqlDataAdapter
                Dim objtable As New DataTable
                Dim objstr As String = ""
                objstr = " SELECT * FROM CLASS_TMPSCHEDULE WHERE CTSID = @ctsid " '& Me.DataGrid1.DataKeys(e.Item.ItemIndex)
                da.SelectCommand = New SqlCommand(objstr, objconn)
                da.SelectCommand.Parameters.Clear()
                da.SelectCommand.Parameters.Add("ctsid", SqlDbType.Int).Value = Me.DataGrid1.DataKeys(e.Item.ItemIndex)
                da.Fill(objtable)
                If objtable.Rows.Count = 1 Then dr = objtable.Rows(0)
                If dr Is Nothing Then
                    Common.MessageBox(Me, "請重新查詢有效資料", rtnPath)
                    Exit Sub
                End If

                'objstr = "SELECT * FROM CLASS_TMPSCHEDULE where ctsid=" & Me.DataGrid1.DataKeys(e.Item.ItemIndex)
                'objtable = DbAccess.GetDataTable(objstr, da, objconn)
                'dr = objtable.Rows(0)
                If hidTmpTItemID.Value = "" Then
                    Me.DataGrid1.EditItemIndex = -1
                    Common.MessageBox(Page, "項次資料有誤，請重新操作該功能!!!")
                    Exit Sub
                End If
                If CDbl(hidTmpTItemID.Value) <> CDbl(TItemID.Text) Then
                    chkstr = " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & " AND ItemID = " & CDbl(TItemID.Text)
                    chkrow = DbAccess.GetOneRow(chkstr, objconn)
                    If Not chkrow Is Nothing Then
                        Common.MessageBox(Page, "項次重覆，請重新輸入!!!")
                        Exit Sub
                    End If
                End If

                objstr = " SELECT * FROM CLASS_TMPSCHEDULE WHERE ctsid = " & Me.DataGrid1.DataKeys(e.Item.ItemIndex)
                objtable = DbAccess.GetDataTable(objstr, da, objconn)
                dr = objtable.Rows(0)
                'TItemID = e.Item.FindControl("TItemID")
                If hidTmpTItemID.Value = "" Then
                    Me.DataGrid1.EditItemIndex = -1
                    Common.MessageBox(Page, "項次資料有誤，請重新操作該功能!!!")
                    Exit Sub
                End If
                If CDbl(hidTmpTItemID.Value) <> CDbl(TItemID.Text) Then
                    chkstr = " SELECT * FROM CLASS_TMPSCHEDULE WHERE OCID = " & OCIDValue1.Value & " AND ItemID = " & CDbl(TItemID.Text)
                    chkrow = DbAccess.GetOneRow(chkstr, objconn)
                    If Not chkrow Is Nothing Then
                        Common.MessageBox(Page, "項次重覆，請重新輸入!!!")
                        Exit Sub
                    End If
                End If

                Dim Errmsg As String = ""
                'If Me.OCourseIDValue.Value <> "" Then Me.OCourseIDValue.Value = Trim(Me.OCourseIDValue.Value)
                If Me.OCourseIDValue.Value = "" Then '課程名稱 (代碼)
                    Errmsg += "課程名稱資料不可為空,請重新確認!" & vbCrLf
                End If
                'TRecycle.Text
                'If TRecycle.Text <> "" Then TRecycle.Text = Trim(TRecycle.Text)
                If TRecycle.Text <> "" Then
                    Try
                        TRecycle.Text = CInt(TRecycle.Text)
                        If CInt(TRecycle.Text) <= 0 Then Errmsg += "循環數字有誤 只可輸入大於0的整數數字,請重新確認!" & vbCrLf
                    Catch ex As Exception
                        Errmsg += "循環格式有誤 只可輸入數字,請重新確認!" & vbCrLf
                    End Try
                End If

                'If Me.OLessonTeah1Value.Value <> "" Then Me.OLessonTeah1Value.Value = Trim(Me.OLessonTeah1Value.Value)
                If OLessonTeah1Value.Value = "" AndAlso HidTLessonTeah1.Value = "" Then Errmsg += "教師資料不可為空,請重新確認!" & vbCrLf
                If Errmsg <> "" Then
                    Common.MessageBox(Page, Errmsg)
                    Exit Sub
                End If

                dr("ItemID") = CDbl(TItemID.Text)
                dr("CourseID") = Me.OCourseIDValue.Value '課程名稱 (代碼)

                If TCalHours.Text = "" Then
                    dr("CalHours") = DBNull.Value
                Else
                    dr("CalHours") = TCalHours.Text
                End If
                If TStartDate.Text = "" Then
                    dr("StartDate") = HidClassStartDate.Value
                Else
                    dr("StartDate") = TStartDate.Text
                End If
                If TEndDate.Text = "" Then
                    dr("EndDate") = HidClassEndDate.Value
                Else
                    dr("EndDate") = TEndDate.Text
                End If
                dr("RoomID") = TRoomID.Text
                dr("LessonTeah1") = If(OLessonTeah1Value.Value <> "", OLessonTeah1Value.Value, HidTLessonTeah1.Value) 'HidTLessonTeah1.Value
                dr("LessonTeah2") = If(OLessonTeah2Value.Value <> "", OLessonTeah2Value.Value, HidTLessonTeah2.Value) 'HidTLessonTeah2.Value
                dr("S1") = If(TS1.Text = "", Convert.DBNull, TS1.Text)
                dr("E1") = If(TE1.Text = "", Convert.DBNull, TE1.Text)
                dr("S2") = If(TS2.Text = "", Convert.DBNull, TS2.Text)
                dr("E2") = If(TE2.Text = "", Convert.DBNull, TE2.Text)
                dr("S3") = If(TS3.Text = "", Convert.DBNull, TS3.Text)
                dr("E3") = If(TE3.Text = "", Convert.DBNull, TE3.Text)
                dr("S4") = If(TS4.Text = "", Convert.DBNull, TS4.Text)
                dr("E4") = If(TE4.Text = "", Convert.DBNull, TE4.Text)
                dr("S5") = If(TS5.Text = "", Convert.DBNull, TS5.Text)
                dr("E5") = If(TE5.Text = "", Convert.DBNull, TE5.Text)
                dr("S6") = If(TS6.Text = "", Convert.DBNull, TS6.Text)
                dr("E6") = If(TE6.Text = "", Convert.DBNull, TE6.Text)
                dr("S7") = If(TS7.Text = "", Convert.DBNull, TS7.Text)
                dr("E7") = If(TE7.Text = "", Convert.DBNull, TE7.Text)
                dr("Recycle") = If(TRecycle.Text = "", Convert.DBNull, TRecycle.Text)

                '計算實際時數
                Dim CalHours As Integer '期望排課時數
                Dim TotalDays As Integer
                Dim RealHours As Integer
                Dim BDate As String = ""
                Dim ODate As String = ""
                TStartDate = e.Item.FindControl("TStartDate")
                If TStartDate.Text = "" Then BDate = HidClassStartDate.Value Else BDate = TStartDate.Text
                TEndDate = e.Item.FindControl("TEndDate")
                If TEndDate.Text = "" Then ODate = HidClassEndDate.Value Else ODate = TEndDate.Text
                TCalHours = e.Item.FindControl("TCalHours") '期望排課時數
                If TCalHours.Text = "" Then CalHours = HidClassHours.Value Else CalHours = CInt(TCalHours.Text)
                TotalDays = DateDiff(DateInterval.Day, CDate(BDate), CDate(ODate)) + 1
                RealHours = 0 '已排課時數(實際時數)
                dr("RealHours") = CountRealHour(TItemID, TCalHours, TStartDate, TEndDate, TS1, TE1, TS2, TE2, TS3, TE3, TS4, TE4, TS5, TE5, TS6, TE6, TS7, TE7)
                dr("ModifyAcct") = sm.UserInfo.UserID 'sm.UserInfo.UserID
                dr("ModifyDate") = Now()
                DbAccess.UpdateDataTable(objtable, da)
                Me.DataGrid1.EditItemIndex = -1
            Case "cancel"
                Me.DataGrid1.EditItemIndex = -1
            Case "del"
                Try
                    Call TIMS.sUtl_DeleteClassTmpScheduleCTSID(Me.DataGrid1.DataKeys(e.Item.ItemIndex), objconn)
                    Common.MessageBox(Me, "刪除成功")
                Catch ex As Exception
                    Common.MessageBox(Me, "刪除失敗" & ex.ToString)
                End Try
        End Select

        ClearItem()
        GetCourseData(OCIDValue1.Value)
    End Sub

    '刪除預排資料
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            Call TIMS.sUtl_DeleteClassScheduleFN1(Me, OCIDValue1.Value, objconn) '刪除預排資料
            Common.MessageBox(Me, "刪除成功")
            Button5.Enabled = False
            TIMS.Tooltip(Button5, "已刪除預排資料", True)
        Catch ex As Exception
            Common.MessageBox(Me, "刪除失敗" & ex.ToString)
        End Try
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Exit Sub

        'Dim dt2 As DataTable = Nothing
        'Dim da2 As SqlDataAdapter = Nothing
        'Dim sql As String = ""
        'conn = DbAccess.GetConnection()
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim STDate As Date                  '開訓日期
        Dim FTDate As Date                  '結訓日期
        STDate = TIMS.Cdate3(drCC("STDate"))
        FTDate = TIMS.Cdate3(drCC("FTDate"))

        Dim CourseDataTable As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT CourID ,CourseName ,MainCourID FROM Course_CourseInfo p " & vbCrLf
        sql &= " WHERE EXISTS ( " & vbCrLf
        sql &= "    SELECT 'x' FROM AUTH_RELSHIP c " & vbCrLf
        sql &= " 	WHERE EXISTS ( " & vbCrLf
        sql &= " 	    SELECT 'x' FROM AUTH_RELSHIP c2 " & vbCrLf
        sql &= " 		WHERE c2.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= " 		AND c2.OrgID = c.OrgID " & vbCrLf
        sql &= " 	   ) " & vbCrLf
        sql &= " 	   AND c.RID = p.RID " & vbCrLf
        sql &= " ) " & vbCrLf
        CourseDataTable = DbAccess.GetDataTable(sql, objconn)

        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'Plan_Schedule
            '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
            SD_04_002.AddNew_Plan_Schedule(Me, STDate, FTDate, OCIDValue1.Value, dt2, da2, trans, CourseDataTable)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try

        Common.MessageBox(Me, "重建時間配當完成")
    End Sub

#Region "(No Use)"

    ''找出主課程
    'Function GetMainCourse(ByVal CourID As String, ByVal dt As DataTable)
    '    Dim dr As DataRow

    '    If dt.Select("CourID='" & CourID & "'").Length = 0 Then
    '        Return CourID
    '    Else
    '        dr = dt.Select("CourID='" & CourID & "'")(0)
    '        If dr("MainCourID").ToString <> "" Then
    '            Return dr("MainCourID")
    '        Else
    '            Return CourID
    '        End If
    '    End If
    'End Function

    ''提供刪除
    'Protected Sub btnDelX2_Click(sender As Object, e As EventArgs) Handles btnDelX2.Click
    '    'If OCIDValue1.Value = "" Then Exit Sub
    '    If OCIDValue1.Value = "" Then
    '        sMSG = "未選擇班級!!"
    '        Common.MessageBox(Me, sMSG)
    '        Exit Sub
    '    End If
    '    If ISinto.Text <> cst_ISintoMsg6 Then
    '        sMSG = "資料不符合 " & cst_ISintoMsg6
    '        Common.MessageBox(Me, sMSG)
    '        Exit Sub
    '    End If

    '    '刪除資料
    '    Call TIMS.sUtl_DeleteClassTmpSchedule(OCIDValue1.Value, objconn)
    '    Call TIMS.sUtl_DeleteClassSchedule(Me, OCIDValue1.Value, objconn)
    '    Call TIMS.sUtl_DeletePlanSchedule(Me, OCIDValue1.Value, objconn)

    '    btnDelX2.Visible = False
    'End Sub

#End Region
End Class