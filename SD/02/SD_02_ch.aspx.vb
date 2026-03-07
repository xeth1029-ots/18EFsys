Partial Class SD_02_ch
    Inherits AuthBasePage

    'AUTH_ACCRWCLASS,'Dim dt As DataTable,'Dim dtTrainType As DataTable,'Dim PlanKind As Integer,
    Dim objconn As SqlConnection
    Dim cst_dg1_開訓日期 As Int16 = 4
    Dim flag_Roc As Boolean = False '是否啟用西元轉民國年機制

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        TIMS.ChkSession(Me, 9, sm)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

#Region "(No Use)"

        'Request("AcctPlan")=1              表示可以跨計畫選擇
        'Request("special")=1               增加回傳開訓跟結訓日期
        'Request("special")=2               母頁Submit

        'Request("special")=3               SD_05_005.aspx專用
        'Request("special")=4               'EXAM_03_001_add專用
        'Request("special")=5               'SD_05_002_R 專用 //special=5 提供開訓、結訓日期欄位special=5&DateF=start_date,end_date
        'Request("special")=11              '(SD_01_004.aspx)增加回傳 document.form1.hidLockTime1.value
        'Request("special")=12              '(SD_03_002.aspx)增加回傳 document.form1.hidLockTime1.value

        'Request("sort")=1                  表示回傳到TMID1.value
        'Request("sort")=2                  表示回傳到TMID2.value
        'Request("sort")=3                  表示回傳到TMID3.value

        'request("RWClass")=1               被班級計畫賦予限制
        'Request("PlanID")                  程式SD_03_008使用

        'Request("PlanID")                  程式SD_03_008使用
        'Request("PlanID")                  程式SD_03_008使用

#End Region

        flag_Roc = TIMS.CHK_REPLACE2ROC_YEARS()

        'Dim sql As String = ""
        'sql = "SELECT * FROM KEY_TRAINTYPE ORDER BY 1"
        'dtTrainType = DbAccess.GetDataTable(sql, objconn)

        If Not IsPostBack Then
            msg.Text = ""
            search_but.Attributes("onclick") = "javascript:return chkdata();"

            Table2.Visible = False

            Years = TIMS.GetSyear(Years)
            Common.SetListItem(Years, sm.UserInfo.Years)

            YearsTR.Visible = False
            If TIMS.ClearSQM(Request("AcctPlan")) = "1" Then YearsTR.Visible = True '開放年度選擇

            HourRan = TIMS.GET_HOURRAN(HourRan, objconn, sm)

            Dim dt As DataTable = TIMS.GetCookieTable(Me, objconn)

            TB_career_id.Text = TIMS.GetCookieItemValue(dt, "SD_TB_career_id")
            trainValue.Value = TIMS.GetCookieItemValue(dt, "SD_trainValue") 'dt.Select("ItemName='SD_trainValue'")(0)("ItemValue").ToString.Trim
            Common.SetListItem(HourRan, TIMS.GetCookieItemValue(dt, "SD_HourRan"))
            CyclType.Text = TIMS.GetCookieItemValue(dt, "SD_CyclType") 'dt.Select("ItemName='SD_CyclType'")(0)("ItemValue").ToString.Trim
            ClassID.Text = TIMS.GetCookieItemValue(dt, "SD_ClassID") 'dt.Select("ItemName='SD_ClassID'")(0)("ItemValue").ToString.Trim
            Common.SetListItem(ClassRound, TIMS.GetCookieItemValue(dt, "SD_ClassRound"))
            'search_but_Click(sender, e)
            If TB_career_id.Text <> "" AndAlso trainValue.Value <> "" Then Call SUtl_Search1()
        End If
    End Sub

    Sub GetSearch()
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        TIMS.InsertCookieTable(Me, dt, da, "SD_TB_career_id", TB_career_id.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_trainValue", trainValue.Value, False, objconn)
        'TIMS.InsertCookieTable(Me, dt, da, "SD_jobValue", jobValue.Value)
        'TIMS.InsertCookieTable(Me, dt, da, "SD_txtCJOB_NAME", txtCJOB_NAME.Text)
        'TIMS.InsertCookieTable(Me, dt, da, "SD_cjobValue", cjobValue.Value)
        TIMS.InsertCookieTable(Me, dt, da, "SD_HourRan", HourRan.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_CyclType", CyclType.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassID", ClassID.Text.Trim, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassRound", ClassRound.SelectedValue, True, objconn)
    End Sub

#Region "(No Use)"

    'select * from class_classinfo where ocid ='37778'
    'select * from Stud_EnterType where 1=1
    'and ocid1='37778'
    'select * from Stud_EnterTemp p
    'where exists (
    '	select 'x' from Stud_EnterType x where 1=1
    '	and ocid1='37778'
    '	and x.setid=p.setid
    ')
    'select * from Stud_EnterType2 where 1=1
    'and ocid1='37778'
    'select * from Stud_EnterTemp2 p
    'where exists (
    '	select 'x' from Stud_EnterType2 x where 1=1
    '	and ocid1='37778'
    '	and x.esetid=p.esetid
    ')
    'select rid from class_classinfo where planid ='1520' --and cycltype ='01'
    'select rid from PLAN_PLANINFO where planid ='3823' --and cycltype ='01'
    'select rid from class_classinfo CC JOIN ID_CLASS IC ON IC.CLSID=CC.CLSID where CC.planid ='3823' --and cycltype ='01'
    'select * from id_plan where planid ='3823'
    'select * from Sys_OrgVar where rid='G'
    'SELECT * FROM Stud_SelResult WHERE 1=1 and ocid='37778'
    'select * from class_studentsofclass where ocid ='37778'
    'select ss.* 
    'from class_studentsofclass cs 
    'join stud_studentinfo ss on ss.sid =cs.sid
    'where 1=1
    'and cs.ocid ='37778'
    'select ss2.* 
    'from class_studentsofclass cs 
    'join stud_studentinfo ss on ss.sid =cs.sid
    'join stud_subdata ss2 on ss2.sid =cs.sid 
    'where 1=1
    'and cs.ocid ='37778'

    'select top 10 * from Stud_TrainingResults
    'select top 10 * from Key_Sanction
    'select * from Class_Schedule where ocid ='37778'
    'SELECT * FROM Teach_TeacherInfo where rid='F'
    'select * from Plan_Schedule where ocid ='37778'
    'select top 10 * from Stud_Conduct
    'select ss.* 
    'from class_studentsofclass cs 
    'JOIN Stud_TrainingResults ss on ss.socid =cs.socid 
    '--join Course_CourseInfo cc on cc.courid =ss.courid
    'where 1=1
    'and cs.ocid ='37778'
    'select ss.* 
    'from class_studentsofclass cs 
    'JOIN Stud_Conduct ss on ss.socid =cs.socid 
    'where 1=1
    'and cs.ocid ='37778'
    'select ss.* 
    'from class_studentsofclass cs 
    'JOIN Stud_Turnout ss on ss.socid =cs.socid 
    'where 1=1
    'and cs.ocid ='37778'
    'select ss.* 
    'from class_studentsofclass cs 
    'JOIN Stud_Sanction ss on ss.socid =cs.socid 
    'where 1=1
    'and cs.ocid ='37778'

#End Region

    ''' <summary>
    ''' SEARCH 查詢
    ''' </summary>
    Sub SUtl_Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'If ClassCName.Text <> "" Then ClassCName.Text = Trim(ClassCName.Text)
        'If ClassID.Text <> "" Then ClassID.Text = Trim(ClassID.Text)
        ClassCName.Text = TIMS.ClearSQM(ClassCName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        ClassID.Text = TIMS.ClearSQM(ClassID.Text)
        Dim RqRWClass As String = TIMS.ClearSQM(Request("RWClass"))
        Dim RqRID As String = TIMS.ClearSQM(Request("RID"))
        Dim RqAcctPlan As String = TIMS.ClearSQM(Request("AcctPlan"))
        Dim RqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqSelected_year As String = TIMS.ClearSQM(Request("selected_year")) '有輸入年度設定年度

        '年度
        Dim strYears As String = Convert.ToString(sm.UserInfo.Years) 'DEFAULT 登入年度
        If rqSelected_year <> "" Then Common.SetListItem(Years, rqSelected_year) '有輸入年度設定年度
        Dim v_Years As String = TIMS.GetListValue(Years)
        If v_Years <> "" Then strYears = v_Years
        strYears = TIMS.ClearSQM(strYears)

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        jobValue.Value = TIMS.ClearSQM(jobValue.Value)

        Dim sql As String = ""
        sql &= " SELECT cc.PlanID,cc.OCID,cc.ClassCName " & vbCrLf
        sql &= " ,format(cc.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(cc.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " ,cc.TMID " & vbCrLf
        sql &= " ,cc.IsApplic " & vbCrLf
        sql &= " ,CASE WHEN cc.IsApplic='Y' THEN '可挑選志願' ELSE '不可挑選志願' END IsApplic2 " & vbCrLf
        sql &= " ,cc.CLSID,cc.CyclType " & vbCrLf
        sql &= " ,cc.ComIDNO,cc.SeqNO,cc.LevelType " & vbCrLf
        'sql &= " ,cc.Years+'0'+ic.ClassID+cc.CyclType ClassID " & vbCrLf
        '提供期別可以不填寫，但未填補01 維持一致性
        sql &= " ,concat(cc.YEARS,'0',ISNULL(ic.ClassID2,ic.ClassID),ISNULL(cc.CYCLTYPE,'01')) ClassID " & vbCrLf
        sql &= " ,IsNull(t.TrainID ,t.JobID) TrainID " & vbCrLf
        sql &= " ,IsNull(t.TrainName ,t.JobName) TrainName " & vbCrLf
        sql &= " ,'['+ISNULL(t.TrainID,t.JobID)+']'+ISNULL(t.TrainName,t.JobName) TrainName2 " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,pp.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,cc.CJOB_UNKEY,cc.TPeriod,ip.DistID " & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PlanID = cc.PlanID AND pp.COMIDNO = cc.COMIDNO AND pp.SEQNO = cc.SEQNO " & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sql &= " JOIN ID_CLASS ic ON ic.CLSID = cc.CLSID " & vbCrLf
        sql &= " JOIN KEY_TRAINTYPE t ON t.TMID = cc.TMID " & vbCrLf
        sql &= " WHERE cc.IsSuccess='Y'" & vbCrLf
        sql &= " AND cc.NotOpen='N'" & vbCrLf
        sql &= " AND ip.Years='" & strYears & "'" & vbCrLf

        '班別代碼 '提供期別可以不填寫，但未填補01 維持一致性
        If ClassID.Text <> "" Then sql &= " AND concat(cc.YEARS,'0',ISNULL(ic.ClassID2,ic.ClassID),ISNULL(cc.CYCLTYPE,'01')) LIKE '%" & ClassID.Text & "%' " & vbCrLf

        '訓練職類      
        If trainValue.Value <> "" Then sql &= " AND cc.TMID = '" & trainValue.Value & "' " & vbCrLf

        '通俗職類
        If txtCJOB_NAME.Text <> "" AndAlso jobValue.Value <> "" Then sql &= " AND cc.CJOB_UNKEY = " & cjobValue.Value & " " & vbCrLf

        '訓練時段
        Dim v_HourRan As String = TIMS.GetListValue(HourRan)
        If HourRan.SelectedIndex <> 0 AndAlso v_HourRan <> "" Then sql &= " AND cc.TPeriod = '" & v_HourRan & "' " & vbCrLf

        '期別
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        If CyclType.Text <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    sql &= " AND cc.CyclType = '0" & Int(CyclType.Text) & "' " & vbCrLf
                Else
                    sql &= " AND cc.CyclType = '" & CyclType.Text & "' " & vbCrLf
                End If
            End If
        End If

        '班級名稱
        If ClassCName.Text <> "" Then sql &= " AND cc.ClassCName LIKE '%" & ClassCName.Text & "%' " & vbCrLf

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '開訓二週前
                sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),cc.STDate) >=0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                sql &= " AND DATEDIFF(DAY,cc.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,cc.FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                sql &= " AND DATEDIFF(DAY,cc.FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                sql &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                sql &= " AND 1<>1 "
        End Select

        Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objconn)
        If sPlanKind = "1" AndAlso RqRWClass = "1" Then
            If sm.UserInfo.DistID <> "000" Then
                sql &= " AND EXISTS (SELECT 'x' FROM Auth_AccRWClass x WHERE x.Account = '" & sm.UserInfo.UserID & "' AND x.OCID = cc.OCID) " & vbCrLf
            Else
                '署(局)限制。
                If RqRID = "A" Then sql &= " AND EXISTS (SELECT 'x' FROM AUTH_ACCRWCLASS x WHERE x.Account = '" & sm.UserInfo.UserID & "' AND x.OCID = cc.OCID) " & vbCrLf
            End If
        End If

        If sm.UserInfo.RID <> "A" Then
            '(RID!='A')分署或委訓單位 ' 限定計畫/轄區。
            sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND ip.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf

            If RqRID <> "" Then sql &= " AND cc.RID = '" & RqRID & "' " & vbCrLf

            '非使用跨計畫-限定RID
            If RqRID = "" AndAlso RqAcctPlan <> "1" Then
                sql &= " AND cc.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
            End If
        Else
            '署(局)【RID='A'】
            '使用跨計畫 不限定RID，除非有傳入RID '程式SD_03_008使用
            If RqRID <> "" Then sql &= " AND cc.RID = '" & RqRID & "' " & vbCrLf

            '除非有PlanID傳入
            If RqPlanID <> "" Then sql &= " AND cc.PlanID = '" & RqPlanID & "' " & vbCrLf '程式SD_03_008使用

            '不可跨計畫
            If RqAcctPlan <> "1" Then
                '限定年度 限定計畫/轄區
                If Len(RqRID) = 1 Then sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf '(分署)長度為1限定計畫。

                '署(局)不限制。
                If Convert.ToString(RqRID) = "" OrElse RqRID = "A" Then sql &= " AND ip.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
            End If
        End If

        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        'If (flag_test) Then TIMS.writeLog(Me, "SQL:" & sql)

        Dim sCmd As New SqlCommand(sql, objconn)

#Region "(No Use)"

        'select ic.* from ID_Class ic join class_classinfo cc on ic.clsid=cc.clsid where cc.planid ='1672'
        'select * from class_classinfo where planid ='1672'
        'select * from plan_planinfo where planid ='1672'
        'select * from Auth_AccRWClass where Account='e222613795'
        'select OCID from Auth_AccRWClass where Account='e222613795'

        'select c.*
        'FROM Class_ClassInfo i
        'join ID_Class  c on i.CLSID=c.CLSID
        'where 1=1
        'and i.IsSuccess='Y' and i.NotOpen='N' and i.Years='12'
        ' and i.RID='D'
        ' and i.PlanID ='1705'
        'select * from Auth_AccRWClass where Account='H220017636'

        'select ic.* from ID_Class ic join class_classinfo cc on ic.clsid=cc.clsid where cc.planid =1561
        'select cc.* from ID_Class ic join class_classinfo cc on ic.clsid=cc.clsid where cc.planid =1561
        'select * from plan_planinfo where planid =1561
        'select *
        'from Auth_AccRWClass p where 1=1
        'and exists (select 'x' from class_classinfo x where x.ocid =p.ocid and x.planid ='1561')

        'select * from class_classinfo where planid ='1672' and cycltype ='01'
        'select * from Stud_EnterType where 1=1
        'and ocid1='41825'
        'select * from Stud_EnterTemp p
        'where exists (
        '	select 'x' from Stud_EnterType x where 1=1
        '	and ocid1='41825'
        '	and x.setid=p.setid
        ')
        'select rid from class_classinfo where planid ='1672' and cycltype ='01'
        'select * from id_plan where planid ='1672'
        'select * from Sys_OrgVar where  rid='G'
        'SELECT * FROM Stud_SelResult WHERE 1=1 and ocid='41825'

#End Region

        'Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        msg.Text = "查無資料!!"
        Table2.Visible = False
        If dt.Rows.Count = 0 Then Return

        Call GetSearch()
        msg.Text = ""
        Table2.Visible = True

        If Hid_SSSDTRID.Value = "" Then Hid_SSSDTRID.Value = TIMS.GetRnd6Eng()
        PageControler1.SSSDTRID = Hid_SSSDTRID.Value 'TIMS.GetRnd6Eng()
        PageControler1.PageDataTable = dt
        PageControler1.Sort = "ClassID,CyclType"
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>
    ''' SEARCH 查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Search_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search_but.Click
        Call SUtl_Search1()
    End Sub

    ''' <summary>
    ''' SEND 送出。
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Send_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles send.Click
        Dim Rqclass1 As String = TIMS.ClearSQM(Request("class1")) 'OCID 班級代號
        Dim RqSort As String = TIMS.ClearSQM(Request("sort"))
        Dim RqFName As String = TIMS.ClearSQM(Request("FName"))
        Dim RqSpecial As String = TIMS.ClearSQM(Request("special"))
        Dim RqBtnName As String = TIMS.ClearSQM(Request("BtnName"))
        If Rqclass1 = "" Then
            Common.MessageBox(Me, "請先勾選班級!")
            Exit Sub
        End If
        Dim dr As DataRow = TIMS.GetOCIDDate(Val(Rqclass1), objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, "查無資料!")
            Exit Sub
        End If
        Dim className As String = Convert.ToString(dr("CLASSCNAME2"))

        Common.RespWrite(Me, "<script language=javascript>")
        Common.RespWrite(Me, "function returnNum(){")
        If RqSort = "" Then
            Common.RespWrite(Me, "window.opener.document.form1.TMID1.value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue1.value='" & dr("TMID") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCID1.value='" & className & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue1.value='" & dr("OCID") & "';")
        Else
            Common.RespWrite(Me, "window.opener.document.form1.TMID" & RqSort & ".value='[" & dr("TrainID") & "]" & dr("TrainName") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.TMIDValue" & RqSort & ".value='" & dr("TMID") & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCID" & RqSort & ".value='" & className & "';")
            Common.RespWrite(Me, "window.opener.document.form1.OCIDValue" & RqSort & ".value='" & dr("OCID") & "';")
        End If

        If RqFName <> "" Then Common.RespWrite(Me, "window.opener.document.form1.TDate.value ='" & TIMS.Cdate3(dr("FTDate")) & "';") '程式SD_05_017_R 專用

        Select Case Val(RqSpecial)
            Case 11
                '1:鎖定。
                Common.RespWrite(Me, "window.opener.document.form1.hidLockTime1.value='1';")
                If Convert.ToString(dr("ShowOK")) = "Y" Then
                    '2:不鎖定。
                    Common.RespWrite(Me, "window.opener.document.form1.hidLockTime1.value='2';")
                End If
            Case 12
                '1:鎖定。
                Common.RespWrite(Me, "window.opener.document.form1.hidLockTime1.value='1';")
                If Convert.ToString(dr("ShowOK14")) = "Y" Then
                    '2:不鎖定。
                    Common.RespWrite(Me, "window.opener.document.form1.hidLockTime1.value='2';")
                End If
            Case 5
                'SD_05_002_R 專用 //special=5 提供開訓、結訓日期欄位special=5&DateS=start_date&DateF=end_date
                Dim DateS As String = TIMS.ClearSQM(Request("DateS"))
                Dim DateF As String = TIMS.ClearSQM(Request("DateF"))
                If DateS <> "" Then
                    Common.RespWrite(Me, "window.opener.document.form1." & DateS & ".value='" & TIMS.Cdate3(dr("STDate")) & "';")
                End If
                If DateF <> "" Then
                    Common.RespWrite(Me, "window.opener.document.form1." & DateF & ".value='" & TIMS.Cdate3(dr("FTDate")) & "';")
                End If
            Case 4 'EXAM_03_001_add專用
                If IsDBNull(dr("ExamDate")) Then
                    Common.RespWrite(Me, "window.opener.document.form1.txt_examdate1.value='未填寫';")
                Else
                    Common.RespWrite(Me, "window.opener.document.form1.txt_examdate1.value='" & TIMS.Cdate3(dr("ExamDate")) & "';")
                End If
            Case 1 '增加回傳開訓與結訓日期
                Common.RespWrite(Me, "window.opener.document.form1.SDate.value='" & TIMS.Cdate3(dr("STDate")) & "';")
                Common.RespWrite(Me, "window.opener.document.form1.FDate.value='" & TIMS.Cdate3(dr("FTDate")) & "';")
            Case 2
                'SD_04_001   'SD_05_003
                Common.RespWrite(Me, "window.opener.document.form1.submit();")
            Case 3
                'SD_05_005.aspx專用
                Dim iStudCount As Integer = 0
                If Rqclass1 <> "" Then
                    Dim pms1 As New Hashtable From {{"OCID", Val(Rqclass1)}}
                    Dim sql1 As String = "SELECT COUNT(1) STUDCOUNT FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID"
                    iStudCount = DbAccess.ExecuteScalar(sql1, objconn, pms1)
                End If
                Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=false;")
                If dr("IsClosed").ToString = "Y" Then
                    Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=true;")
                    Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級已作過結訓作業';")
                Else
                    If iStudCount > 0 Then
                        Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級尚未作過結訓作業';")
                        Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=false;")
                    Else
                        Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級無學生,無法結訓';")
                        Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=true;")
                        Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=true;")
                    End If
                End If
                Common.RespWrite(Me, "window.opener.document.form1.Usered.value='1';")
                Common.RespWrite(Me, "window.opener.document.form1.submit();")
#Region "(No Use)"

                'Dim FunDr As DataRow
                'Dim FunDt As DataTable = sm.UserInfo.FunDt
                'Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
                'If FunDrArray.Length <> 0 Then
                '    FunDr = FunDrArray(0)
                '    If FunDr("Sech") = 1 Then
                '        'Common.RespWrite(Me, "window.opener.document.form1.Button5.disabled=false;")
                '    End If

                '    If FunDr("Adds") = 1 Then
                '        Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=false;")
                '        If dr("IsClosed").ToString = "Y" Then
                '            Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=true;")
                '            Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級已作過結訓作業';")
                '        Else
                '            If iStudCount > 0 Then
                '                Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級尚未作過結訓作業';")
                '                Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=false;")
                '            Else
                '                Common.RespWrite(Me, "window.opener.document.getElementById('ClassMsg').innerHTML='此班級無學生,無法結訓';")
                '                Common.RespWrite(Me, "window.opener.document.form1.Button6.disabled=true;")
                '                Common.RespWrite(Me, "window.opener.document.form1.Button4.disabled=true;")
                '            End If
                '        End If
                '        Common.RespWrite(Me, "window.opener.document.form1.Usered.value='1';")
                '        Common.RespWrite(Me, "window.opener.document.form1.submit();")
                '    End If
                'End If

#End Region
        End Select

        If RqBtnName <> "" Then Common.RespWrite(Me, "window.opener.document.form1." & RqBtnName & ".click();")

        Common.RespWrite(Me, "window.close();")
        Common.RespWrite(Me, "}")
        Common.RespWrite(Me, "returnNum();")
        Common.RespWrite(Me, "</script>")

        '存入暫存資料
        Dim da As SqlDataAdapter = Nothing
        Dim drTemp As DataRow = Nothing
        Dim dt As DataTable = TIMS.GetCookieTable(Me, da, objconn)
        For i As Integer = 1 To 10
            If dt.Select("ItemName='Temp_OCID" & i & "'").Length = 0 Then
                Dim InsertFlag As Boolean = True
                For j As Integer = 1 To 10
                    If dt.Select("ItemName='Temp_OCID" & j & "' and ItemValue='" & dr("OCID") & "'").Length <> 0 Then InsertFlag = False
                Next
                If InsertFlag Then
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_OCID" & i, dr("OCID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_ClassName" & i, className, False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_TMID" & i, dr("TMID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_TrainName" & i, "[" & dr("TrainID") & "]" & dr("TrainName"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_RID" & i, dr("RID"), False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "Temp_OrgName" & i, dr("OrgName"), True, objconn)
                End If
                Exit For
            Else
                If i = 10 Then
                    For j As Integer = 1 To 9
                        TIMS.SetCookieItemValue(dt, "Temp_OCID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_ClassName", j)
                        TIMS.SetCookieItemValue(dt, "Temp_TMID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_TrainName", j)
                        TIMS.SetCookieItemValue(dt, "Temp_RID", j)
                        TIMS.SetCookieItemValue(dt, "Temp_OrgName", j)
                    Next

                    Dim InsertFlag As Boolean = True
                    For j As Integer = 1 To 10
                        If dt.Select("ItemName='Temp_OCID" & j & "' and ItemValue='" & dr("OCID") & "'").Length <> 0 Then InsertFlag = False
                    Next
                    If InsertFlag = True Then
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_OCID" & i, dr("OCID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_ClassName" & i, className, False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_TMID" & i, dr("TMID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_TrainName" & i, "[" & dr("TrainID") & "]" & dr("TrainName"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_RID" & i, dr("RID"), False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "Temp_OrgName" & i, dr("OrgName"), True, objconn)
                    End If
                End If
            End If
        Next
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If flag_Roc Then e.Item.Cells(cst_dg1_開訓日期).Text = TIMS.Cdate17(drv("STDate"))
        End Select
    End Sub
End Class