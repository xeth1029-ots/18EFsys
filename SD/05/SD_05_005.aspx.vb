Partial Class SD_05_005
    Inherits AuthBasePage

    Public Shared s_log1 As String = ""
    Const cst_iDG2_COL_學號 As Integer = 0
    'Const cst_iDG2_COL_姓名 As Integer = 1
    Const cst_iDG2_COL_出勤扣分 As Integer = 2
    Const cst_iDG2_COL_獎懲扣分 As Integer = 3
    'Const cst_iDG2_COL_實際上課時數 As Integer = 4
    'Const cst_iDG2_COL_上課比率 As Integer = 5
    Const cst_iDG2_COL_導師加減分 As Integer = 6
    Const cst_iDG2_COL_教務課加減分 As Integer = 7
    Const cst_iDG2_COL_操行成績 As Integer = 8
    'Const cst_iDG2_COL_是否核發結訓證書 As Integer = 9

    Dim tmpValue1 As String = "" '組合暫存空間。
    Dim vsTitle As String = ""
    Dim Days1 As Integer = 0
    Dim Days2 As Integer = 0

    Dim dtArc As DataTable
    Dim int_BasicPoint As Integer = 0 '底分
    Dim giCheckMode As Integer = 0 '成績計算方式   ( 1, 各科平均法  2, 訓練時數權重法 )

#Region "NO USE"
    'SELECT * FROM PLAN_SCHEDULE WHERE OCID IN (?)
    'SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (?)
    'SELECT * FROM VIEW_CLASS_SCHEDULE WHERE OCID=?
    'SELECT * FROM STUD_TURNOUT WHERE SOCID=?
    'SELECT * FROM STUD_TRAININGRESULTS WHERE SOCID=?
    'SELECT * FROM STUD_CONDUCT WHERE SOCID=?
#End Region

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '取出設定天數檔 Start
        Call TIMS.Get_SysDays(Days1, Days2)
        '暫時許可權Table
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)

        Const cst_GVID_3 As String = "3"
        Dim sGVID3 As String = TIMS.GetGlobalVar(Me, cst_GVID_3, "1", objconn)
        Select Case sm.UserInfo.LID
            Case 0 '署'選取不同的資訊
                If RIDValue.Value <> sm.UserInfo.RID Then
                    Dim sDistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    sGVID3 = TIMS.GetGlobalVar(sDistID, sm.UserInfo.TPlanID, cst_GVID_3, "1", objconn)
                End If
        End Select
        sGVID3 = TIMS.ClearSQM(sGVID3)
        If sGVID3 = "" Then
            If Not IsPostBack Then Common.MessageBox(Me, "尚未設定系統操行底分,請聯絡分署系統管理者")
            behavior.Enabled = False
            vsTitle = "尚未設定系統操行底分,請聯絡分署系統管理者"
            TIMS.Tooltip(behavior, vsTitle)
        Else
            behavior.Enabled = True
            int_BasicPoint = Val(sGVID3)
        End If
        Const cst_GVID_13 As String = "13"
        Dim sGVID13 As String = TIMS.GetGlobalVar(Me, cst_GVID_13, "1", objconn)
        Select Case sm.UserInfo.LID
            Case 0 '署'選取不同的資訊
                If RIDValue.Value <> sm.UserInfo.RID Then
                    Dim sDistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    sGVID13 = TIMS.GetGlobalVar(sDistID, sm.UserInfo.TPlanID, cst_GVID_13, "1", objconn)
                End If
        End Select

        sGVID13 = TIMS.ClearSQM(sGVID13)
        Select Case sGVID13
            Case "1", "2"
                class_grade.Enabled = True
                giCheckMode = Val(sGVID13)   '成績計算方式(1:各科平均法  2:訓練時數權重法 )

            Case Else
                If Not IsPostBack Then Common.MessageBox(Me, "尚未設定成績計算模式,請聯絡分署系統管理者")
                class_grade.Enabled = False

                vsTitle = "尚未設定成績計算模式,請聯絡分署系統管理者"
                TIMS.Tooltip(class_grade, vsTitle)
        End Select

        If Not IsPostBack Then msg.Text = ""

        '增加Script---   Start
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button10.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button1.Attributes("onclick") = "javascript:return search();"
        Button4.Attributes("onclick") = "javascript:return confirm('您確定要計算這個班級的總成績?');"

        class_grade.Attributes("onclick") = "grade();"
        behavior.Attributes("onclick") = "grade();"
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "Button6")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        '列印空白表
        'Button9.Attributes("onclick") =     ReportQuery.ReportScript(Me, "MultiBlock", "SD_05_005_R1", "OCID='+document.getElementById('OCIDValue1').value+'&ChooseClass='+document.getElementById('CCID').value+'")
        'Me.Button9.Attributes.Add("OnClick", "return PrintWhiteRpt('" & ReportQuery.GetSmartQueryPath & "'," & BasicPoint & ");")  20090930 andy test
        '列印成績表
        'Button5.Attributes("onclick") =     ReportQuery.ReportScript(Me, "MultiBlock", "SD_05_005_R", "OCID='+document.getElementById('OCIDValue1').value+'")
        Button6.Style("display") = "none"

        '搜尋主課程(不啟用)
        Button3.Disabled = True
        If class_grade.Checked Then Button3.Disabled = False '(啟用)

        '檢查帳號的功能許可權---Start,Button2.Enabled = True,If Not au.blnCanAdds Then,Button2.Enabled = False,vsTitle = "您無許可權使用該功能",
        'TIMS.Tooltip(Button2, vsTitle),End If,Button1.Enabled = True,If Not au.blnCanSech Then,Button1.Enabled = False,vsTitle = "您無許可權使用該功能",
        'TIMS.Tooltip(Button1, vsTitle),End If,檢查帳號的功能許可權---End,

        'smpath.Value = Convert.ToString(ReportQuery.GetSmartQueryPath)
        tb_EditResult.Visible = False
        bt_save.Enabled = True

        rbCreditPoints.Style.Item("display") = "none"
        If Not IsPostBack Then
            '計算總成績
            Button4.Enabled = False

            'rbCreditPoints.Visible = False
            rbCreditPoints.Style.Item("display") = "none"
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                rbCreditPoints.Style.Item("display") = TIMS.cst_inline1 '"inline"
                Call SD05015_GetSession()
            End If

            GradeTable.Style.Item("display") = "none"
            PrintTB.Style.Item("display") = "none"
            BehaviorTable.Style.Item("display") = "none"
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '20081208 檢查學術百分比是否設定 andy edit     
            Call chkPercentSet()

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            tb_EditResult.Visible = False
        End If

        If OCIDValue1.Value <> "" Then
            If class_grade.Checked = True Then
                If ChooseClass.Value <> "" Then
                    Call CreateGradeTable()

                    GradeTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    PrintTB.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    TR_SingType.Style.Item("display") = TIMS.cst_inline1 '"inline"

                    If RB_SignType.SelectedValue = 1 Then
                        TR_Signer2.Style.Item("display") = "none"
                        TR_Signer.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    Else
                        TR_Signer2.Style.Item("display") = TIMS.cst_inline1 '"inline"
                        TR_Signer.Style.Item("display") = "none"
                    End If

                    Bn_Print.Style.Item("display") = TIMS.cst_inline1 '"inline"

                Else
                    GradeTable.Style.Item("display") = "none"
                    PrintTB.Style.Item("display") = "none"
                End If
                BehaviorTable.Style.Item("display") = "none"
            ElseIf behavior.Checked = True Then
                GradeTable.Style.Item("display") = "none"
                PrintTB.Style.Item("display") = "none"
                BehaviorTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
            End If
        End If
        '2006/03/28 add conn by matt 'conn = DbAccess.GetConnection
    End Sub

    'SQL 建立成績檔案
    Sub CreateGradeTable()
        Table4.Rows.Clear()
        Button2.Attributes("onclick") = "javascript:return chkdata(1);"
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim PMS1 As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT a.Name + case when b.studstatus in (2,3) then '(*)' else '' end Name" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(b.StudentID) StudentID ,b.SOCID ,b.OCID ,b.StudStatus ,b.CloseDate ,c.FTDate" & vbCrLf
        sql &= " ,case when b.studstatus=2 then a.Name + ' 離訓(' + ISNULL(CONVERT(varchar, b.rejectTDate1, 111),'') + ')'" & vbCrLf
        sql &= " when b.studstatus=3 then a.Name + ' 退訓(' + ISNULL(CONVERT(varchar, b.rejectTDate2, 111),'') + ')' end rejectTDate" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON b.SID=a.SID and b.OCID=@OCID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID" & vbCrLf
        sql &= " JOIN ID_CLASS g on g.CLSID=c.CLSID" & vbCrLf
        sql &= " ORDER BY b.StudentID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "該班無有效學員資料。")
            Exit Sub
        End If

        'Dim SOCID As String = ""
        Dim strTmpSOCID As String = ""
        Dim tmpLable3 As String = ""
        For Each dr As DataRow In dt.Rows
            'Dim tmpValue1 As String = "" '組合暫存空間。
            tmpValue1 = "'" & dr("SOCID") & "'"
            If strTmpSOCID.IndexOf(tmpValue1) = -1 Then '重複字過濾
                If strTmpSOCID <> "" Then strTmpSOCID += ","
                strTmpSOCID += tmpValue1
            End If

            If Convert.ToString(dr("rejectTDate")) <> "" Then
                tmpValue1 = Convert.ToString(dr("rejectTDate"))
                If tmpLable3.IndexOf(tmpValue1) = -1 Then '重複字過濾
                    If tmpLable3 <> "" Then tmpLable3 += "<br>"
                    tmpLable3 += tmpValue1
                End If
            End If
        Next
        Label3.Text = tmpLable3

        ChooseClass.Value = TIMS.CombiSQM2IN(ChooseClass.Value)

        If ChooseClass.Value = "" Then
            GradeTable.Style.Item("display") = "none"
            PrintTB.Style.Item("display") = "none"
            msg.Text = "查無資料"
            Exit Sub
        End If
        'If ChooseClass.Value <> "" Then

        '建立表頭---Start
        'sql = "SELECT *,0 as TotalHours FROM COURSE_COURSEINFO WHERE CourID in (" & ChooseClass.Value & ") ORDER BY CourID"
        sql = ""
        sql &= " SELECT a.CourseName,a.CourID,0 TotalHours"
        sql &= " FROM COURSE_COURSEINFO A"
        sql &= $" WHERE a.CourID IN ({ChooseClass.Value})"
        sql &= " ORDER BY a.CourseName"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt2.Rows.Count > 6 Then 'dt2學科(寬)、dt學員(長)
            scrollDiv.Attributes.Add("class", "DivWidth")
        Else
            scrollDiv.Attributes.Add("class", "DivHeight")
        End If

        Dim MyCell As TableCell = Nothing
        'Dim MyTextBox As TextBox
        Dim MyCheck As HtmlInputCheckBox = Nothing
        Dim MyLabel As Label = Nothing
        Dim MyRow As New TableRow
        MyRow.Attributes.Add("class", "FixedTitleRow")  '20100204 andy edit 固定第一列

        MyCell = New TableCell
        MyRow.Cells.Add(MyCell)
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Attributes.Add("class", "FixedTitleColumn")  '固定位置(前三個欄位不移動)

        MyCheck = New HtmlInputCheckBox
        MyCheck.ID = "totalall"
        MyCheck.Value = "all"

        MyCell.Controls.Add(MyCheck)
        MyCheck.Attributes.Add("OnClick", "return chall();")  '勾選時更新學號 SelSocid.value (表首第一列--全選 )
        MyCell = New TableCell
        MyCell.Attributes.Add("class", "FixedTitleColumn")  '固定位置(前三個欄位不移動)
        MyRow.Cells.Add(MyCell)
        MyCell.Text = "學號"
        MyCell.BorderWidth = Unit.Pixel(1)

        MyCell = New TableCell
        MyCell.Attributes.Add("class", "FixedTitleColumn")  '固定位置(前三個欄位不移動)
        MyRow.Cells.Add(MyCell)
        MyCell.Text = "姓名"
        MyCell.BorderWidth = Unit.Pixel(1)

        If giCheckMode = 2 Then
            '將目前所有的使用課程列出---Start
            Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
            sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID=@OCID"
            Dim dtClass_Schedule As DataTable = DbAccess.GetDataTable(sql, objconn, pmsS)

            For Each dr1 As DataRow In dt2.Rows
                For Each dr As DataRow In dtClass_Schedule.Rows '排課每日課程
                    For i As Integer = 1 To 12
                        If Not IsDBNull(dr("Class" & i)) Then
                            If Convert.ToString(dr("Class" & i)) = Convert.ToString(dr1("CourID")) Then
                                dr1("TotalHours") += 1
                            End If
                        End If
                    Next
                Next
            Next
            '將目前所有的使用課程列出---End
        End If

        For Each dr2 As DataRow In dt2.Select(Nothing, "courseName")
            MyCell = New TableCell
            MyRow.Cells.Add(MyCell)
            MyCell.BorderWidth = Unit.Pixel(1)

            MyCheck = New HtmlInputCheckBox
            MyCheck.ID = $"Course{dr2("CourID")}"
            MyCheck.Value = $"{dr2("CourID")}"

            MyCell.Controls.Add(MyCheck)
            '090408 andy  勾選時更新課程代碼 SelCourID.value (表首第一列)
            MyCheck.Attributes.Add("OnClick", "return ChkSelCourID();")
            MyLabel = New Label
            If giCheckMode = 1 Then
                MyLabel.Text = $"{dr2("CourseName")}"
            Else 'CheckMode = 2
                MyLabel.Text = $"{dr2("CourseName")}({dr2("TotalHours")}Hr)"
            End If

            MyCell.Controls.Add(MyLabel)
        Next
        MyRow.BackColor = Color.FromName("#2aafc0")
        MyRow.ForeColor = Color.White
        Table4.Rows.Add(MyRow)
        '建立表頭---End

        '建立資料---Start
        sql = "" & vbCrLf
        sql &= " SELECT Distinct b.SOCID, a.CourID,a.CourseName,b.Results" & vbCrLf
        sql &= " FROM COURSE_COURSEINFO a" & vbCrLf
        sql &= " JOIN STUD_TRAININGRESULTS b ON a.CourID=b.CourID" & vbCrLf
        sql &= $" WHERE a.CourID IN ({ChooseClass.Value}) AND b.SOCID IN ({strTmpSOCID})" & vbCrLf
        sql &= " ORDER BY b.SOCID, a.CourseName" & vbCrLf
        dt2 = DbAccess.GetDataTable(sql, objconn)

        sql = ""
        sql &= " SELECT * FROM COURSE_COURSEINFO"
        sql &= $" WHERE CourID IN ({ChooseClass.Value})"
        sql &= " ORDER BY courseName"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sql, objconn)

        hidSD05005CHK551.Value = ""
        hidSD05005CHK572.Value = ""
        hidSD05005CHK598.Value = ""
        hidSD05005CHK618.Value = ""
        hidSD05005CHK653.Value = ""
        'Dim i As Integer = 1
        'Dim rowIndex As Int16 = 1
        For Each dr As DataRow In dt.Rows
            MyRow = New TableRow

            MyCell = New TableCell
            MyRow.Cells.Add(MyCell)
            MyCell.BorderWidth = Unit.Pixel(1)

            MyCheck = New HtmlInputCheckBox
            MyCheck.ID = dr("SOCID").ToString
            MyCheck.Value = "'" & dr("SOCID").ToString & "'"
            MyCell.Attributes.Add("class", "FixedDataColumn")  '20100204 andy edit 固定欄位位置
            MyCell.Controls.Add(MyCheck)

            '090408 andy 勾選時更新學號 SelSocid.value 
            MyCheck.Attributes.Add("OnClick", "return ChkSelSocid();")

            MyCell = New TableCell         '學號
            MyCell.Attributes.Add("class", "FixedDataColumn")  '20100204 andy edit 固定欄位位置
            MyRow.Cells.Add(MyCell)
            MyCell.Text = dr("StudentID")   'Right(dr("StudentID"), 2) 
            MyCell.BorderWidth = Unit.Pixel(1)

            MyCell = New TableCell         '姓名
            MyCell.Attributes.Add("class", "FixedDataColumn")  '20100204 andy edit 固定欄位位置
            MyRow.Cells.Add(MyCell)
            MyCell.Text = dr("Name")
            MyCell.BorderWidth = Unit.Pixel(1)

            'For Each dr3 In dt3.Select(Nothing, "CourID")

            Dim i2 As Integer = 1
            Dim tabIndex As Int16 = 0
            For Each dr3 As DataRow In dt3.Select(Nothing, "courseName")
                'If rowIndex <> dt.Rows.Count Then
                '    rowIndex = rowIndex + 1
                'End If
                MyCell = New TableCell
                Dim MyTextBox As TextBox = New TextBox
                MyCell.Controls.Add(MyTextBox)
                MyRow.Cells.Add(MyCell)
                MyCell.BorderWidth = Unit.Pixel(1)
                Dim ff3 As String = $"SOCID='{dr("SOCID")}' and CourID='{dr3("CourID")}'"
                If dt2.Select(ff3, "courseName").Length <> 0 Then
                    MyTextBox.Text = dt2.Select(ff3, "courseName")(0)("Results").ToString
                    'If dt2.Select("SOCID='" & dr("SOCID") & "' and CourID='" & dr3("CourID").ToString & "'", "CourID").Length <> 0 Then
                    'MyTextBox.Text = dt2.Select("SOCID='" & dr("SOCID") & "' and CourID='" & dr3("CourID").ToString & "'", "CourID")(0)("Results").ToString
                Else
                    MyTextBox.Text = ""
                End If

                MyTextBox.Attributes.Add("onKeyDown", "if(event.keyCode==13){event.keyCode=9; }")
                tabIndex = tabIndex + 1
                MyTextBox.TabIndex = tabIndex

                MyTextBox.ID = $"Course{dr3("CourID")}_{dr("SOCID")}"
                MyTextBox.Columns = 5

                Select Case dr("StudStatus")
                    Case 1, 4
                        MyTextBox.Enabled = True
                    Case 2, 3
                        MyTextBox.Enabled = False
                        vsTitle = "學員狀態為離訓、退訓"
                        TIMS.Tooltip(MyTextBox, vsTitle)
                    Case 5
                        If TIMS.Check_Auth_RendClass(dr("OCID"), dtArc) Then
                            MyTextBox.Enabled = True
                        Else
                            Select Case sm.UserInfo.RoleID
                                Case 0, 1
                                    '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時 超過75天後,限制輸入
                                    '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                                    If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 75 Then
                                            hidSD05005CHK551.Value = TIMS.cst_YES
                                            MyTextBox.Enabled = False '每格分數
                                            vsTitle = "補助辦理保母職業訓練(46)/辦理照顧服務員職業訓練(47) 超過75天後,限制輸入"
                                            TIMS.Tooltip(MyTextBox, vsTitle)
                                        Else
                                            hidSD05005CHK551.Value = TIMS.cst_NO
                                            MyTextBox.Enabled = True
                                        End If
                                    Else
                                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days2 Then
                                            hidSD05005CHK572.Value = TIMS.cst_YES
                                            MyTextBox.Enabled = False '每格分數
                                            vsTitle = $"超過{Days2}天後,限制輸入"
                                            TIMS.Tooltip(MyTextBox, vsTitle)
                                        Else
                                            hidSD05005CHK572.Value = TIMS.cst_NO
                                            MyTextBox.Enabled = True
                                        End If
                                    End If

                                Case Else
                                    '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限制天數改成60天
                                    '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                                    If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 60 Then
                                            hidSD05005CHK598.Value = TIMS.cst_YES
                                            MyTextBox.Enabled = False '每格分數
                                            vsTitle = "補助辦理保母職業訓練(46)/辦理照顧服務員職業訓練(47)  超過60天後,限制輸入"
                                            TIMS.Tooltip(MyTextBox, vsTitle)
                                        Else
                                            hidSD05005CHK598.Value = TIMS.cst_NO
                                            MyTextBox.Enabled = True
                                        End If
                                    Else
                                        If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > Days1 Then
                                            hidSD05005CHK618.Value = TIMS.cst_YES
                                            MyTextBox.Enabled = False '每格分數
                                            vsTitle = $"超過{Days1}天後,限制輸入"
                                            TIMS.Tooltip(MyTextBox, vsTitle)
                                        Else
                                            hidSD05005CHK618.Value = TIMS.cst_NO
                                            MyTextBox.Enabled = True
                                        End If
                                    End If

                            End Select
                        End If
                End Select
                '20090227 andy edit  2010/11/29 豪哥說先取消這個定義
                'If DateDiff(DateInterval.Day, dr("FTDate"), Now.Date) > 100 Then '結訓日超過一百天不開放
                '    MyTextBox.Enabled = False
                '    Button2.Enabled = False
                '    Button4.Enabled = False
                'Else
                '    Button4.Enabled = True
                'End If

                '授權設定該班級有設定則開放
                If Not TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_結訓成績登錄, dtArc) Then
                    hidSD05005CHK653.Value = TIMS.cst_YES
                    MyTextBox.Enabled = True '每格分數
                    vsTitle = "授權設定該班級有開放"
                    TIMS.Tooltip(MyTextBox, vsTitle)
                End If

                ''授權設定該班級有設定則開放
                'If ChkIsEndDate(sm.UserInfo.UserID, OCIDValue1.Value.ToString()) = False Then
                '    MyTextBox.Enabled = True
                '    Button2.Enabled = True
                '    Button4.Enabled = True
                '    bt_save.Enabled = True
                '    Button7.Enabled = True
                'End If
            Next
            Table4.Rows.Add(MyRow)

            If i2 Mod 2 = 1 Then
                MyRow.BackColor = Color.FromName("#ecf7ff")
            Else
                MyRow.BackColor = Color.White
            End If
            i2 += 1
        Next

        Select Case hidSD05005CHK551.Value
            Case TIMS.cst_YES
                Button2.Enabled = False '計算儲存
                Button4.Enabled = False '計算總成績
                bt_save.Enabled = False '(科目成績)儲存
                Button7.Enabled = False '(成績)儲存
                vsTitle = "補助辦理保母職業訓練(46)/辦理照顧服務員職業訓練(47) 超過75天後,限制輸入"
                TIMS.Tooltip(Button2, vsTitle)
                TIMS.Tooltip(Button4, vsTitle)
                TIMS.Tooltip(bt_save, vsTitle)
                TIMS.Tooltip(Button7, vsTitle)
            Case TIMS.cst_NO
                Button2.Enabled = True
                Button4.Enabled = True
                bt_save.Enabled = True
                Button7.Enabled = True
        End Select
        Select Case hidSD05005CHK572.Value
            Case TIMS.cst_YES
                Button2.Enabled = False
                Button4.Enabled = False
                bt_save.Enabled = False
                Button7.Enabled = False
                vsTitle = "超過" & Days2 & "天後,限制輸入"
                TIMS.Tooltip(Button2, vsTitle)
                TIMS.Tooltip(Button4, vsTitle)
                TIMS.Tooltip(bt_save, vsTitle)
                TIMS.Tooltip(Button7, vsTitle)
            Case TIMS.cst_NO
                Button2.Enabled = True
                Button4.Enabled = True
                bt_save.Enabled = True
                Button7.Enabled = True
        End Select
        Select Case hidSD05005CHK598.Value
            Case TIMS.cst_YES
                Button2.Enabled = False
                Button4.Enabled = False
                bt_save.Enabled = False
                Button7.Enabled = False
                vsTitle = "補助辦理保母職業訓練(46)/辦理照顧服務員職業訓練(47)  超過60天後,限制輸入"
                TIMS.Tooltip(Button2, vsTitle)
                TIMS.Tooltip(Button4, vsTitle)
                TIMS.Tooltip(bt_save, vsTitle)
                TIMS.Tooltip(Button7, vsTitle)
            Case TIMS.cst_NO
                Button2.Enabled = True
                Button4.Enabled = True
                bt_save.Enabled = True
                Button7.Enabled = True
        End Select
        Select Case hidSD05005CHK618.Value
            Case TIMS.cst_YES
                Button2.Enabled = False
                Button4.Enabled = False
                bt_save.Enabled = False
                Button7.Enabled = False
                vsTitle = "超過" & Days1 & "天後,限制輸入"
                TIMS.Tooltip(Button2, vsTitle)
                TIMS.Tooltip(Button4, vsTitle)
                TIMS.Tooltip(bt_save, vsTitle)
                TIMS.Tooltip(Button7, vsTitle)
            Case TIMS.cst_NO
                Button2.Enabled = True
                Button4.Enabled = True
                bt_save.Enabled = True
                Button7.Enabled = True
        End Select
        Select Case hidSD05005CHK653.Value
            Case TIMS.cst_YES
                Button2.Enabled = True
                Button4.Enabled = True
                bt_save.Enabled = True
                Button7.Enabled = True
                vsTitle = "授權設定該班級有開放"
                TIMS.Tooltip(Button2, vsTitle)
                TIMS.Tooltip(Button4, vsTitle)
                TIMS.Tooltip(bt_save, vsTitle)
                TIMS.Tooltip(Button7, vsTitle)
        End Select

        GradeTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        PrintTB.Style.Item("display") = TIMS.cst_inline1 '"inline"
        'TR_SingType.Style.Item("display") = TIMS.cst_inline1 '"inline"
        'TR_Signer2.Style.Item("display") = "none"
        'TR_Signer.Style.Item("display") = TIMS.cst_inline1 '"inline"
        'Bn_Print.Style.Item("display") = TIMS.cst_inline1 '"inline"
        '建立資料---End
    End Sub

    Sub chkPercentSet()
        'ByRef dtGlobalVar As DataTable
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        PercentSet.Value = "N"
        Dim GVID17_1 As String = TIMS.GetGlobalVar(CStr(sm.UserInfo.DistID), CStr(sm.UserInfo.TPlanID), "17", "1", objconn)
        Dim GVID17_2 As String = TIMS.GetGlobalVar(CStr(sm.UserInfo.DistID), CStr(sm.UserInfo.TPlanID), "17", "2", objconn)
        If GVID17_1 <> "" OrElse GVID17_2 <> "" Then
            PercentSet.Value = "Y"
        End If

    End Sub

    '取出所有使用中的課程ID-COURID
    Public Shared Function Get_AllCourID(ByVal OCIDV1 As String, ByRef oConn As SqlConnection) As String
        Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCIDV1)}}
        Dim sql As String = ""
        sql &= " WITH WC1 AS (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID )" & vbCrLf
        sql &= " ,WC2 AS (SELECT CourID FROM STUD_TRAININGRESULTS WHERE SOCID IN (SELECT SOCID FROM WC1))" & vbCrLf
        sql &= " SELECT DISTINCT b.COURSENAME,''''+CONVERT(varchar(33),a.COURID)+'''' COURID" & vbCrLf
        sql &= " FROM WC2 a" & vbCrLf
        sql &= " JOIN COURSE_COURSEINFO b ON a.CourID=b.CourID" & vbCrLf
        sql &= " ORDER BY b.COURSENAME" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, pmsS)

        Dim DBClsValue As String = ""
        For Each dr As DataRow In dt.Rows
            Dim tmpValue1 As String = Convert.ToString(dr("CourID")) 'CourID必有值
            If DBClsValue.IndexOf(tmpValue1) = -1 Then '重複字過濾
                If DBClsValue <> "" Then DBClsValue &= ","
                DBClsValue &= tmpValue1
            End If
        Next
        Return DBClsValue
    End Function

    'SQL 課程。
    Sub search_ChooseClass()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            msg.Visible = True
            msg.Text = "查無資料"
            Exit Sub
        End If
        'Hid_OCID1.Value = Convert.ToString(drCC("OCID"))
        'Hid_THOURS.Value = Convert.ToString(drCC("THOURS"))
        'If drCC Is Nothing Then Exit Sub
        ChooseClass.Value = Get_AllCourID(OCIDValue1.Value, objconn)
        msg.Visible = False
        msg.Text = ""
        If ChooseClass.Value = "" Then
            msg.Visible = True
            msg.Text = "查無資料"
        End If
    End Sub

    Sub SD05015_KeepSession()
        Dim str_Session_SearchSD05015 As String = ""
        str_Session_SearchSD05015 = "sd=05015"
        str_Session_SearchSD05015 &= "&center=" & center.Text
        str_Session_SearchSD05015 &= "&RIDValue=" & RIDValue.Value
        str_Session_SearchSD05015 &= "&TMID1=" & TMID1.Text
        str_Session_SearchSD05015 &= "&OCID1=" & OCID1.Text
        str_Session_SearchSD05015 &= "&TMIDValue1=" & TMIDValue1.Value
        str_Session_SearchSD05015 &= "&OCIDValue1=" & OCIDValue1.Value
        'str_Session_SearchSD05015 &= "&Button1=1"
        Session("_SearchSD05015") = str_Session_SearchSD05015
    End Sub

    Sub SD05015_GetSession()
        If Session("_SearchSD05015") Is Nothing Then Exit Sub
        Dim str_Session_SearchSD05015 As String = Session("_SearchSD05015")
        Session("_SearchSD05015") = Nothing
        Dim MyValue As String = TIMS.GetMyValue(str_Session_SearchSD05015, "sd")
        If MyValue <> "05015" Then Exit Sub 'str_Session_SearchSD05015 = Nothing

        center.Text = TIMS.GetMyValue(str_Session_SearchSD05015, "center")
        RIDValue.Value = TIMS.GetMyValue(str_Session_SearchSD05015, "RIDValue")
        TMID1.Text = TIMS.GetMyValue(str_Session_SearchSD05015, "TMID1")
        OCID1.Text = TIMS.GetMyValue(str_Session_SearchSD05015, "OCID1")
        TMIDValue1.Value = TIMS.GetMyValue(str_Session_SearchSD05015, "TMIDValue1")
        OCIDValue1.Value = TIMS.GetMyValue(str_Session_SearchSD05015, "OCIDValue1")
        'Session("_SearchSD05015") = Nothing 'ViewState("Button1") = TIMS.GetMyValue(Session("_SearchSD05015"), "Button1")
    End Sub

    Sub Search1()
        '計算總成績
        Button4.Enabled = True
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        GradeTable.Visible = True
        PrintTB.Visible = True
        CCIDVALUE.Value = ""
        If ViewState("IsFirstLoad") Is Nothing Then
            Page.RegisterStartupScript("chkall", "<script> checkAll(); </script>") '20081119 andy
            ViewState("IsFirstLoad") = "N"
        End If

        If OCIDValue1.Value <> "" Then
            Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drCC Is Nothing Then
                msg.Visible = True
                msg.Text = "查無資料"
                Exit Sub
            End If
            Hid_OCID1.Value = Convert.ToString(drCC("OCID"))
            Hid_THOURS.Value = Convert.ToString(drCC("THOURS"))
            UPDATE_CREDITPOINTS(Hid_OCID1.Value)
        End If

        If class_grade.Checked And ChooseClass.Value = "" Then
            search_ChooseClass() 'SQL 課程。
        End If

        Select Case True
            Case class_grade.Checked '課程
                If ChooseClass.Value <> "" Then
                    Call CreateGradeTable()

                    GradeTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    PrintTB.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    TR_SingType.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    TR_Signer2.Style.Item("display") = "none"
                    TR_Signer.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    Bn_Print.Style.Item("display") = TIMS.cst_inline1 '"inline"
                    Print1.Value = 1
                Else
                    GradeTable.Style.Item("display") = "none"
                    PrintTB.Style.Item("display") = "none"
                End If
                BehaviorTable.Style.Item("display") = "none"

            Case behavior.Checked '操行
                Button2.Attributes("onclick") = "javascript:return chkdata(2);" '操行
                'Dim eMP As TIMS.eMinusPoint = TIMS.GetMinusPoint(objconn, OCIDValue1.Value, RIDValue.Value, sm.UserInfo.PlanID)
                'Hid_eMP.Value = eMP
                Dim dt As DataTable = Get_Stud_Turnout(sm.UserInfo.PlanID, RIDValue.Value, OCIDValue1.Value, objconn)

                Dim tmpLable3 As String = ""
                Label4.Text = ""
                For Each dr As DataRow In dt.Rows
                    If Convert.ToString(dr("rejectTDate")) <> "" Then
                        tmpValue1 = Convert.ToString(dr("rejectTDate"))
                        If tmpLable3.IndexOf(tmpValue1) = -1 Then '重複字過濾
                            tmpLable3 &= String.Concat(If(tmpLable3 <> "", "<br>", ""), tmpValue1)
                        End If
                    End If
                Next
                Label4.Text = tmpLable3

                TIMS.Tooltip(Button2, "")
                TIMS.Tooltip(Button4, "")
                msg.Text = "查無資料"
                DataGrid2.Visible = False
                Button2.Visible = False

                If TIMS.dtHaveDATA(dt) Then
                    msg.Text = ""
                    DataGrid2.Visible = True
                    Button2.Visible = True

                    DataGrid2.DataSource = dt
                    DataGrid2.DataKeyField = "SOCID"
                    DataGrid2.DataBind()

                    tb_EditResult.Visible = False
                End If
                'Button2.計算儲存
                Button2.Enabled = True
                If Hid_lockBtn2.Value = "Y" Then Button2.Enabled = False
                If Hid_lockBtn2_MSG.Value <> "" Then TIMS.Tooltip(Button2, Hid_lockBtn2_MSG.Value, True)
                'Button4.計算總成績
                Button4.Enabled = True
                If Hid_lockBtn4.Value = "Y" Then Button4.Enabled = False
                If Hid_lockBtn4_MSG.Value <> "" Then TIMS.Tooltip(Button4, Hid_lockBtn4_MSG.Value, True)

                If dt.Rows.Count = 0 Then
                    '計算總成績
                    Button4.Enabled = False
                    vsTitle = "查無資料"
                    TIMS.Tooltip(Button4, vsTitle)
                End If

            Case rbCreditPoints.Checked '是否取得結訓資格
                Call SD05015_KeepSession()
                'Response.Redirect("SD_05_015.aspx?ID=" & Request("ID"))
                Dim url1 As String = "SD_05_015.aspx?ID=" & Request("ID")
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
        CCID.Value = ChooseClass.Value
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    Private Sub UPDATE_CREDITPOINTS(OCID_VAL As String)
        If OCID_VAL = "" Then Return

        Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCID_VAL)}}
        Dim sSqlS As String = ""
        sSqlS &= " SELECT 1 FROM CLASS_STUDENTSOFCLASS"
        sSqlS &= " WHERE OCID=@OCID AND CREDITPOINTS IS NULL AND STUDSTATUS IN (2,3)" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSqlS, objconn, pmsS)
        If TIMS.dtHaveDATA(dt1) Then
            Dim pmsU As New Hashtable From {{"OCID", TIMS.CINT1(OCID_VAL)}}
            Dim sSqlu As String = ""
            sSqlu &= " UPDATE CLASS_STUDENTSOFCLASS SET CREDITPOINTS=0"
            sSqlu &= " FROM CLASS_STUDENTSOFCLASS"
            sSqlu &= " WHERE OCID=@OCID AND CREDITPOINTS IS NULL AND STUDSTATUS IN (2,3)" & vbCrLf
            DbAccess.ExecuteNonQuery(sSqlu, objconn, pmsU)
        End If

        Dim pmsS2 As New Hashtable From {{"OCID", TIMS.CINT1(OCID_VAL)}}
        Dim sSqlS2 As String = ""
        sSqlS2 &= " SELECT 1 FROM CLASS_STUDENTSOFCLASS"
        sSqlS2 &= " WHERE OCID=@OCID AND CREDITPOINTS=1 AND STUDSTATUS IN (2,3)" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSqlS2, objconn, pmsS2)
        If TIMS.dtHaveDATA(dt2) Then
            Dim pmsU As New Hashtable From {{"OCID", TIMS.CINT1(OCID_VAL)}}
            Dim sSqlu As String = ""
            sSqlu &= " UPDATE CLASS_STUDENTSOFCLASS SET CREDITPOINTS=0"
            sSqlu &= " FROM CLASS_STUDENTSOFCLASS"
            sSqlu &= " WHERE OCID=@OCID AND CREDITPOINTS=1 AND STUDSTATUS IN (2,3)" & vbCrLf
            DbAccess.ExecuteNonQuery(sSqlu, objconn, pmsU)
        End If
    End Sub

    '操行 資料輸出
    Public Shared Function Get_Stud_Turnout(ByVal PlanID As String, ByVal RIDValue As String, ByVal vOCIDValue1 As String,
                                            ByVal tConn As SqlConnection, Optional ByVal int_BasicPoint As Integer = 80) As DataTable

        Dim eMP As TIMS.EMinusPoint = TIMS.GetMinusPoint(tConn, vOCIDValue1, RIDValue, PlanID)
        s_log1 = String.Format("##SD_05_005, TIMS.GetMinusPoint(tConn, OCIDValue1:{0}, RIDValue:{1}, PlanID:{2}), eMP:{3}", vOCIDValue1, RIDValue, PlanID, eMP)
        TIMS.LOG.Debug(s_log1)

        Dim pms1 As New Hashtable From {{"OCID", TIMS.CINT1(vOCIDValue1)}}
        Dim SQL1 As String = ""
        SQL1 &= " WITH WC1 AS (SELECT OCID,SOCID,SID,STUDSTATUS,STUDENTID,BEHAVIORRESULT,CLOSEDATE,REJECTTDATE1,REJECTTDATE2,CREDITPOINTS,ENTERDATE FROM CLASS_STUDENTSOFCLASS WITH(NOLOCK) WHERE OCID=@OCID )" & vbCrLf
        SQL1 &= " SELECT c.NAME+CASE WHEN b.STUDSTATUS in (2,3) then '(*)' ELSE '' END NAME" & vbCrLf
        SQL1 &= " ,'['+d.Years+'0'+g.ClassID+d.CyclType+']'+dbo.FN_GET_CLASSCNAME(d.CLASSCNAME,d.CYCLTYPE) CLASSCNAME" & vbCrLf
        SQL1 &= " ,dbo.FN_CSTUDID2(b.StudentID) StudentID ,b.StudStatus" & vbCrLf
        SQL1 &= " ,dbo.FN_GET_TURNOUT4(b.SOCID) HOURS_ALL" & vbCrLf '(缺席總時數)
        SQL1 &= " ,format(isnull(a.MinusLeave,0), '#######0.##') MinusLeave" & vbCrLf
        SQL1 &= " ,format(isnull(a.MinusSanction,0), '#######0.##') MinusSanction" & vbCrLf
        SQL1 &= " ,a.TechPoint ,a.RemedPoint ,b.BehaviorResult" & vbCrLf
        SQL1 &= " ,b.SOCID ,b.OCID" & vbCrLf
        SQL1 &= " ,b.CloseDate ,d.STDate ,d.FTDate" & vbCrLf
        SQL1 &= " ,format(ISNULL(e.MinusPoint,0), '#######0.##') MinusPoint" & vbCrLf
        SQL1 &= " ,format(ISNULL(f.total,0), '#######0.##') total" & vbCrLf

        SQL1 &= " ,case when b.studstatus=2 then c.Name + ' 離訓(' + isnull(CONVERT(varchar, b.rejectTDate1, 111),'-') + ')'" & vbCrLf
        SQL1 &= "  when b.studstatus=3 then c.Name + ' 退訓(' + isnull(CONVERT(varchar, b.rejectTDate2, 111),'-') + ')' end rejectTDate" & vbCrLf
        SQL1 &= " ,dbo.DECODE6(b.studstatus" & vbCrLf
        SQL1 &= " ,2, dbo.FN_GET_EMPDATE(d.OCID,ISNULL(B.ENTERDATE,d.STDATE) ,b.REJECTTDATE1 -1)" & vbCrLf
        SQL1 &= " ,3, dbo.FN_GET_EMPDATE(d.OCID,ISNULL(B.ENTERDATE,d.STDATE) ,b.REJECTTDATE2 -1)" & vbCrLf
        SQL1 &= " ,dbo.FN_GET_EMPDATE(d.OCID,ISNULL(B.ENTERDATE,d.STDATE) ,CONVERT(DATE,GETDATE()))) CLASSDATE" & vbCrLf

        '是否核發結訓證書
        SQL1 &= " ,b.CREDITPOINTS" & vbCrLf
        '基本分-出缺勤+獎懲+導師加減分+教務課加減分
        Dim s_CONDUCTPOINT As String = String.Format("({0}.0-ISNULL(e.MinusPoint,0.0) + ISNULL(f.total,0.0) + ISNULL(a.TechPoint,0.0) + ISNULL(a.RemedPoint,0.0))", int_BasicPoint)
        SQL1 &= String.Concat(" ,FORMAT(CASE WHEN ", s_CONDUCTPOINT, ">0 THEN ", s_CONDUCTPOINT, " ELSE 0 END,'#######0.##') conductpoint", vbCrLf)
        'sql &= String.Format(" ,FORMAT(({0}.0-ISNULL(e.MinusPoint,0.0) + ISNULL(f.total,0.0) + ISNULL(a.TechPoint,0.0) + ISNULL(a.RemedPoint,0.0)),'#######0.##') conductpoint", CStr(int_BasicPoint)) & vbCrLf

        SQL1 &= " FROM WC1 b" & vbCrLf
        SQL1 &= " JOIN STUD_STUDENTINFO c WITH(NOLOCK) ON b.SID=c.SID" & vbCrLf
        SQL1 &= " JOIN CLASS_CLASSINFO d WITH(NOLOCK) ON b.OCID=d.OCID" & vbCrLf
        SQL1 &= " JOIN ID_CLASS g WITH(NOLOCK) on g.CLSID=d.CLSID" & vbCrLf
        SQL1 &= " LEFT JOIN STUD_CONDUCT a WITH(NOLOCK) ON b.SOCID=a.SOCID" & vbCrLf
        'sm.UserInfo.PlanID, RIDValue.Value, OCIDValue1.Value
        'Select Case chkMinusPoint(PlanID, RIDValue, OCIDValue1, tConn)
        Select Case eMP
            Case TIMS.EMinusPoint.mpKey '"key"
                SQL1 &= " LEFT JOIN ( SELECT e1.SOCID ,sum(e2.MinusPoint*ISNULL(e1.Hours,0)) MinusPoint" & vbCrLf
                SQL1 &= " 	FROM WC1 cs" & vbCrLf
                SQL1 &= " 	JOIN STUD_TURNOUT e1 WITH(NOLOCK) ON e1.SOCID=cs.SOCID" & vbCrLf
                SQL1 &= " 	JOIN KEY_LEAVE e2 WITH(NOLOCK) ON e1.LeaveID=e2.LeaveID" & vbCrLf
                SQL1 &= " 	WHERE ISNULL(e1.TurnoutIgnore,0)=0" & vbCrLf
                SQL1 &= " 	GROUP BY e1.SOCID ) e ON b.SOCID=e.SOCID" & vbCrLf

            Case TIMS.EMinusPoint.mpOrg '"org"
                pms1.Add("PlanID", PlanID)
                'sm.UserInfo.PlanID' RIDValue.Value 
                SQL1 &= " LEFT JOIN ( SELECT e1.SOCID ,sum(e2.MinusPoint*ISNULL(e1.Hours,0)) MinusPoint" & vbCrLf
                SQL1 &= " 	FROM WC1 cs" & vbCrLf
                SQL1 &= " 	JOIN STUD_TURNOUT e1 WITH(NOLOCK) ON e1.SOCID=cs.SOCID" & vbCrLf
                SQL1 &= " 	JOIN ORG_LEAVE e2 WITH(NOLOCK) ON e1.LeaveID=e2.LeaveID" & vbCrLf
                SQL1 &= " 	WHERE e2.PlanID=@PlanID" & vbCrLf
                SQL1 &= "    and e2.OrgID in (SELECT ORGID FROM AUTH_RELSHIP WHERE rid='" & RIDValue & "')" & vbCrLf
                SQL1 &= " 	GROUP BY e1.SOCID ) e ON b.SOCID=e.SOCID" & vbCrLf

            Case TIMS.EMinusPoint.mpClass '"class"
                'OCIDValue1.Value
                SQL1 &= " LEFT JOIN ( SELECT e1.SOCID ,SUM(e2.MinusPoint*ISNULL(e1.Hours,0)) MinusPoint" & vbCrLf
                SQL1 &= " 	FROM WC1 cs" & vbCrLf
                SQL1 &= " 	JOIN STUD_TURNOUT e1 WITH(NOLOCK) ON e1.SOCID=cs.SOCID" & vbCrLf
                SQL1 &= " 	JOIN CLASS_LEAVE e2 WITH(NOLOCK) ON e1.LeaveID=e2.LeaveID" & vbCrLf
                SQL1 &= " 	WHERE e2.OCID=@OCID " & vbCrLf
                SQL1 &= " 	GROUP BY e1.SOCID ) e ON b.SOCID=e.SOCID" & vbCrLf

        End Select
        SQL1 &= " LEFT JOIN ( SELECT f1.SOCID" & vbCrLf
        SQL1 &= " 	,SUM(f1.Times*(CASE WHEN f2.AddMinus='+' THEN ISNULL(f2.Point,0) ELSE 0 - ISNULL(f2.Point,0) END)) total" & vbCrLf
        SQL1 &= " 	FROM WC1 cs" & vbCrLf
        SQL1 &= " 	JOIN STUD_SANCTION f1 WITH(NOLOCK) ON f1.SOCID=cs.SOCID" & vbCrLf
        SQL1 &= " 	JOIN KEY_SANCTION f2 WITH(NOLOCK) on f1.SanID = f2.SanID" & vbCrLf
        SQL1 &= " 	GROUP BY f1.SOCID ) f ON b.SOCID=f.SOCID" & vbCrLf
        SQL1 &= " WHERE d.OCID=@OCID" & vbCrLf
        SQL1 &= " ORDER BY b.StudentID" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(New Page, String.Concat("--", vbCrLf, TIMS.GetMyValue5(pms1), vbCrLf, "--##SD_05_005:", vbCrLf, SQL1))
        End If

        Dim Rst As DataTable = DbAccess.GetDataTable(SQL1, tConn, pms1)
        Return Rst
    End Function

    ''' <summary>是否取得結訓資格 </summary>
    ''' <param name="CreditPoints"></param>
    ''' <param name="oValue"></param>
    Sub SET_CreditPoints_SelectedIndex1(ByRef CreditPoints As DropDownList, ByRef oValue As Object)
        If oValue Is Nothing OrElse IsDBNull(oValue) OrElse Convert.ToString(oValue) = "" Then Return
        CreditPoints.SelectedIndex = If(Convert.ToInt32(oValue) = 1, 1, 2) '1:是／2:否
        'If CreditPoints.SelectedIndex = 1 Then
        '    CreditPoints.Attributes.Add("disabled", "disabled")
        '    CreditPoints.Attributes.Add("readonly", "readonly")
        'Else
        '    CreditPoints.Attributes.Remove("disabled")
        '    CreditPoints.Attributes.Remove("readonly")
        'End If
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim SelectAll As DropDownList = e.Item.FindControl("SelectAll") 'SelectAll.Enabled = If(AppliedResultM = "Y", False, True)
                SelectAll.Attributes("onchange") = "ChangeAll(this.selectedIndex);"

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidDataKeys As HtmlInputHidden = e.Item.FindControl("hidDataKeys")
                Dim labACTUALHOURS As Label = e.Item.FindControl("labACTUALHOURS") '實際上課時數=班級時數-(缺席總時數)
                Dim labTRAINRATIO As Label = e.Item.FindControl("labTRAINRATIO") '上課比率 
                If TIMS.GetValue2(drv("CLASSDATE")) > 0 Then
                    Dim iACTUALHOURS As Integer = (TIMS.GetValue2(drv("CLASSDATE")) - TIMS.GetValue2(drv("HOURS_ALL"))) '實際上課時數=(排課)班級時數-(缺席總時數)
                    labACTUALHOURS.Text = iACTUALHOURS '實際上課時數
                    labTRAINRATIO.Text = String.Concat(TIMS.ROUND(iACTUALHOURS / TIMS.VAL1(Hid_THOURS.Value) * 100, 2), "%") '上課比率=實際上課時數/班級時數
                End If
                'If TIMS.VAL1(Hid_THOURS.Value) > 0 Then
                '    Dim iACTUALHOURS As Integer = (TIMS.GetValue2(Hid_THOURS.Value) - TIMS.GetValue2(drv("HOURS_ALL"))) '實際上課時數=班級時數-(缺席總時數)
                '    labACTUALHOURS.Text = iACTUALHOURS '實際上課時數
                '    labTRAINRATIO.Text = String.Concat(TIMS.ROUND(iACTUALHOURS / TIMS.VAL1(Hid_THOURS.Value) * 100, 2), "%") '上課比率=實際上課時數/班級時數
                'End If

                '是否核發結訓證書
                Dim CreditPoints As DropDownList = e.Item.FindControl("CreditPoints")
                SET_CreditPoints_SelectedIndex1(CreditPoints, drv("CreditPoints"))

                hidDataKeys.Value = Convert.ToString(drv("SOCID"))
                e.Item.Cells(cst_iDG2_COL_學號).Text = Right(drv("StudentID"), 2)
                Select Case Convert.ToString(drv("StudStatus"))
                    Case "2", "3"
                    Case Else
                        Const cst_msgT1 As String = "目前計算結果未儲存完整!!"
                        If msg.Text <> cst_msgT1 Then
                            Try
                                Dim sMinusLeave As String = "0"
                                Dim sMinusPoint As String = "0" 'sMinusLeave = "0" 'sMinusPoint = "0"
                                If Convert.ToString(drv("MinusLeave")) <> "" OrElse Convert.ToString(drv("MinusPoint")) <> "" Then
                                    If Convert.ToString(drv("MinusLeave")) <> "" Then sMinusLeave = Val(drv("MinusLeave"))
                                    If Convert.ToString(drv("MinusPoint")) <> "" Then sMinusPoint = Val(drv("MinusPoint"))
                                    If sMinusLeave <> sMinusPoint Then msg.Text = cst_msgT1
                                End If
                            Catch ex As Exception
                                TIMS.LOG.Error(String.Concat(cst_msgT1, ex.Message), ex)
                            End Try
                        End If
                End Select

                Dim mytext1 As TextBox = e.Item.FindControl("TextBox2") '導師加減分
                mytext1.Text = Convert.ToString(drv("TechPoint")).Trim

                Dim mytext2 As TextBox = e.Item.FindControl("TextBox3") '教務課加減分
                mytext2.Text = Convert.ToString(drv("RemedPoint")).Trim

                mytext1.Attributes.Add("onKeyDown", "if(event.keyCode==13){event.keyCode=9; }")
                mytext1.TabIndex = 1
                mytext2.Attributes.Add("onKeyDown", "if(event.keyCode==13){event.keyCode=9; }")
                mytext2.TabIndex = 2

                Select Case drv("StudStatus")
                    Case 1, 4
                        mytext1.Enabled = True
                        mytext2.Enabled = True
                    Case 2, 3
                        mytext1.Enabled = False
                        mytext2.Enabled = False
                        CreditPoints.Enabled = False
                        vsTitle = "學員狀態為離訓、退訓"
                        TIMS.Tooltip(mytext1, vsTitle)
                        TIMS.Tooltip(mytext2, vsTitle)
                        TIMS.Tooltip(CreditPoints, vsTitle)

                    Case 5
                        If TIMS.Check_Auth_RendClass(drv("OCID"), dtArc) Then
                            mytext1.Enabled = True
                            mytext2.Enabled = True
                        Else
                            Select Case sm.UserInfo.RoleID
                                Case 0, 1
                                    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                        mytext1.Enabled = False
                                        mytext2.Enabled = False
                                        vsTitle = "超過" & Days2 & "天後,限制輸入"
                                        TIMS.Tooltip(mytext1, vsTitle)
                                        TIMS.Tooltip(mytext2, vsTitle)
                                    Else
                                        mytext1.Enabled = True
                                        mytext2.Enabled = True
                                    End If
                                Case Else
                                    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                        mytext1.Enabled = False
                                        mytext2.Enabled = False
                                        vsTitle = "超過" & Days1 & "天後,限制輸入"
                                        TIMS.Tooltip(mytext1, vsTitle)
                                        TIMS.Tooltip(mytext2, vsTitle)
                                    Else
                                        mytext1.Enabled = True
                                        mytext2.Enabled = True
                                    End If
                            End Select
                        End If
                End Select
                '20090227 andy edit
                Hid_lockBtn2.Value = ""
                Hid_lockBtn4.Value = ""
                Hid_lockBtn2_MSG.Value = ""
                Hid_lockBtn4_MSG.Value = ""
                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > 100 Then '結訓日超過一百天不開放
                    mytext1.Enabled = False
                    mytext2.Enabled = False
                    Hid_lockBtn2.Value = "Y"
                    Hid_lockBtn4.Value = "Y"
                    'Button2.Enabled = False
                    'Button4.Enabled = False

                    vsTitle = "超過100天後,限制輸入"
                    TIMS.Tooltip(mytext1, vsTitle)
                    TIMS.Tooltip(mytext2, vsTitle)
                    Hid_lockBtn2_MSG.Value = vsTitle
                    Hid_lockBtn4_MSG.Value = vsTitle
                    'TIMS.Tooltip(Button2, vsTitle)
                    'TIMS.Tooltip(Button4, vsTitle)
                End If

                '授權設定該班級有設定則開放
                If Not TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_結訓成績登錄, dtArc) Then
                    mytext1.Enabled = True
                    mytext2.Enabled = True
                    Hid_lockBtn2.Value = "Y"
                    Hid_lockBtn4.Value = "Y"
                    'Button2.Enabled = True
                    'Button4.Enabled = True
                    vsTitle = "授權設定該班級有開放"
                    TIMS.Tooltip(mytext1, vsTitle)
                    TIMS.Tooltip(mytext2, vsTitle)
                    Hid_lockBtn2_MSG.Value = vsTitle
                    Hid_lockBtn4_MSG.Value = vsTitle
                    'TIMS.Tooltip(Button2, vsTitle)
                    'TIMS.Tooltip(Button4, vsTitle)
                End If

                ''授權設定該班級有設定則開放
                'If ChkIsEndDate(sm.UserInfo.UserID, OCIDValue1.Value.ToString()) = False Then
                '    mytext1.Enabled = True
                '    mytext2.Enabled = True
                '    Button2.Enabled = True
                '    Button4.Enabled = True
                'End If

                Dim sig_mytext1 As Single = TIMS.VAL1(mytext1.Text)
                Dim sig_mytext2 As Single = TIMS.VAL1(mytext2.Text)
                Dim iConduct As Double = CSng(int_BasicPoint) - CSng(drv("MinusPoint")) + CSng(drv("total")) + sig_mytext1 + sig_mytext2
                If iConduct < 0 Then iConduct = 0
                e.Item.Cells(cst_iDG2_COL_操行成績).Text = TIMS.ROUND(iConduct, 2) '4捨5入

        End Select

    End Sub

    '檢核
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Call TIMS.OpenDbConn(objconn)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim sql As String = ""
        sql &= " SELECT 'X' FROM CLASS_STUDENTSOFCLASS" & vbCrLf
        sql &= " WHERE OCID=@OCID AND SOCID=@SOCID" & vbCrLf
        Dim sCmd_CS1 As New SqlCommand(sql, objconn)

        'If Trim(start_date.Text) <> "" Then start_date.Text = Trim(start_date.Text) Else start_date.Text = ""
        'If Trim(end_date.Text) <> "" Then end_date.Text = Trim(end_date.Text) Else end_date.Text = ""
        Dim j As Integer = 0
        For i As Integer = 0 To DataGrid2.Items.Count - 1
            Dim mytext1 As TextBox = DataGrid2.Items(i).Cells(cst_iDG2_COL_導師加減分).FindControl("TextBox2")
            Dim mytext2 As TextBox = DataGrid2.Items(i).Cells(cst_iDG2_COL_教務課加減分).FindControl("TextBox3")

            Dim dtCs1 As New DataTable
            With sCmd_CS1
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = DataGrid2.DataKeys(i)
                dtCs1.Load(.ExecuteReader())
            End With
            If dtCs1.Rows.Count = 0 Then
                Errmsg += "查無有效資料，請重新查詢！" & vbCrLf
                Return False
            End If

            If mytext1.Enabled = True Then
                If Trim(mytext1.Text) <> "" Then mytext1.Text = Trim(mytext1.Text) Else mytext1.Text = ""
                If Trim(mytext2.Text) <> "" Then mytext2.Text = Trim(mytext2.Text) Else mytext2.Text = ""

                If mytext1.Text = "" AndAlso mytext2.Text = "" Then
                Else
                    If Trim(mytext1.Text) <> "" Then
                        mytext1.Text = TIMS.ChangeIDNO(mytext1.Text)
                        If Len(mytext1.Text) > 4 Then
                            Errmsg += "第" & CStr(i + 1) & "筆:導師加減分 長度超過系統範圍限定(4)" & vbCrLf
                        End If
                        If Not IsNumeric(mytext1.Text) Then
                            Errmsg += "第" & CStr(i + 1) & "筆:導師加減分 應為數字格式" & vbCrLf
                        End If
                    End If

                    If Trim(mytext2.Text) <> "" Then
                        mytext2.Text = TIMS.ChangeIDNO(mytext2.Text)
                        If Len(mytext2.Text) > 4 Then
                            Errmsg += "第" & CStr(i + 1) & "筆:輔導課加減分 長度超過系統範圍限定(4)" & vbCrLf
                        End If
                        If Not IsNumeric(mytext2.Text) Then
                            Errmsg += "第" & CStr(i + 1) & "筆:輔導課加減分 應為數字格式" & vbCrLf
                        End If
                    End If
                End If
                j += 1
            End If
        Next

        If j = 0 Then
            Errmsg += "查無有效資料，請重新查詢！" & vbCrLf
        End If
        If Errmsg = "" Then
            Try
                'Dim sql As String = ""
                Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
                sql = "SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID"
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pmsS)
                If TIMS.dtNODATA(dt) Then
                    Errmsg += "查無有效資料，請重新查詢！" & vbCrLf
                    Return False
                End If
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Errmsg += "查詢資料有誤，請重新查詢！" & vbCrLf
                Return False
            End Try
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '計算總成績(做排名)
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            'msg.Visible = True'msg.Text = "查無資料"'Exit Sub
            Common.MessageBox(Me.Page, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        Dim total As Double = 0
        If GradeTable.Visible = False Or GradeTable.Style.Item("display") = "none" Then
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End If
        Try
            If Table4.Rows(0).Cells.Count = 3 Then
                Common.MessageBox(Me, "執行失敗，無學員成績資料!")
                Exit Sub
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End Try
        'Button7_Click(sender, e)
        Call SaveData1()

        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級查詢有誤!!")
            Exit Sub
        End If
        '抓出所有課目(LEFT JOIN 含綜合評量)2007-11-29 -amu

        Dim dt As DataTable = Nothing
        Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT *" & vbCrLf
        sql &= " FROM STUD_TRAININGRESULTS a" & vbCrLf
        sql &= " LEFT JOIN (SELECT * FROM PLAN_SCHEDULE WHERE OCID=@OCID) b ON a.CourID=b.CourID" & vbCrLf
        sql &= " WHERE a.SOCID IN (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID)" & vbCrLf
        sql &= " ORDER BY a.SOCID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn, pmsS)

        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班級查詢有誤!!")
            Exit Sub
        End If

        '抓出CLASS_STUDENTSOFCLASS準備填入
        Dim da As SqlDataAdapter = Nothing
        'Dim sql As String = ""
        sql = ""
        sql &= " SELECT SOCID,TotalResult,Rank,ModifyAcct,ModifyDate "
        sql &= " FROM CLASS_STUDENTSOFCLASS "
        sql &= $" WHERE OCID={TIMS.CINT1(OCIDValue1.Value)}"
        'sql &= " ORDER BY Rank"
        Dim dtResult As DataTable = DbAccess.GetDataTable(sql, da, objconn)
        If dtResult.Rows.Count = 0 Then
            Common.MessageBox(Me, "執行失敗，無學員資料!")
            Exit Sub
        End If
        '先把資料庫中的CLASS_STUDENTSOFCLASS總成績清除
        For Each dr As DataRow In dtResult.Rows
            dr("TotalResult") = 0 '清除每一行
        Next

        Dim Class_Schedule As DataTable = Nothing
        Dim dtCourseInfo As DataTable = Nothing
        Dim AllCourseHours As Integer = 0

        If giCheckMode = 2 Then
            '將目前所有的使用課程列出---Start
            Dim pmsS1 As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
            Dim sql_1 As String = "SELECT DISTINCT CourID,0 as TotalHours FROM STUD_TRAININGRESULTS WHERE SOCID IN (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID)"
            dtCourseInfo = DbAccess.GetDataTable(sql_1, objconn, pmsS1)

            Dim pmsS2 As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
            Dim sql_2 As String = "SELECT * FROM CLASS_SCHEDULE WHERE OCID=@OCID"
            Class_Schedule = DbAccess.GetDataTable(sql_2, objconn, pmsS2)

            For Each dr1 As DataRow In dtCourseInfo.Rows
                For Each dr2 As DataRow In Class_Schedule.Rows
                    For i As Integer = 1 To 12
                        If Not IsDBNull(dr2("Class" & i)) Then
                            If dr2("Class" & i) = dr1("CourID") Then
                                dr1("TotalHours") += 1
                            End If
                        End If
                    Next
                Next
                AllCourseHours += dr1("TotalHours")
            Next
            '將目前所有的使用課程列出---End
        End If

        For Each dr As DataRow In dt.Rows '照著課目跑，有幾堂課的成績就跑幾次
            'SOCIDTemp = dr("SOCID") '暫存這次的SOCID
            If giCheckMode = 1 Then
                total = TIMS.VAL1(dr("Results"))     '取出該SOCID的課目成績
            Else 'CheckMode = 2
                Dim fff4 As String = "CourID='" & dr("CourID") & "'"
                Dim iTotalHours As Integer = If(dtCourseInfo.Select(fff4).Length > 0, dtCourseInfo.Select(fff4)(0)("TotalHours"), 0)
                total = TIMS.VAL1(dr("Results")) * iTotalHours
            End If

            Dim fffR As String = String.Concat("SOCID='", dr("SOCID"), "'")
            If dtResult.Select(fffR).Length = 0 Then
                Common.MessageBox(Me, "執行失敗，無學員資料!")
                Exit Sub
            End If
            Dim drResult As DataRow = dtResult.Select(fffR)(0)   '過濾出CLASS_STUDENTSOFCLASS中個人資料
            If drResult Is Nothing Then
                Common.MessageBox(Me, "執行失敗，無學員資料!")
                Exit Sub
            End If
            Select Case giCheckMode
                Case 1
                    If IsDBNull(drResult("TotalResult")) Then            '假如是一開始尚未輸入的狀態
                        drResult("TotalResult") = total
                    Else
                        drResult("TotalResult") += total
                    End If
                Case Else 'CheckMode = 2
                    If AllCourseHours <> 0 Then
                        If IsDBNull(drResult("TotalResult")) Then            '假如是一開始尚未輸入的狀態
                            drResult("TotalResult") = total / AllCourseHours
                        Else
                            drResult("TotalResult") += total / AllCourseHours
                        End If
                    Else
                        If IsDBNull(drResult("TotalResult")) Then            '假如是一開始尚未輸入的狀態
                            drResult("TotalResult") = total / 1
                        Else
                            drResult("TotalResult") += total / 1
                        End If
                    End If
            End Select
        Next
        For Each dr As DataRow In dtResult.Rows
            dr("TotalResult") = Math.Round(dr("TotalResult"), 2)
        Next

        '準備開始排名
        Dim GradeTemp As Integer = 0
        ''Dim i2 As Integer = 0
        'Dim j As Integer = 1
        If dtResult.Select("1=1", "TotalResult desc").Length = 0 Then
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End If
        '20090622 andy edit  先清除原排名
        For Each dr As DataRow In dtResult.Select("1=1", "TotalResult desc")
            If GradeTemp <> Val(dr("TotalResult")) Then     '假如目前的Temp值跟該學生的總成績不相同時
                GradeTemp = Val(dr("TotalResult"))          '將最大成績存入Temp中
            End If
            'dr("Rank") = i          
            dr("Rank") = Convert.DBNull
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dtResult, da)

        'ViewState("SelAllSocid") = "Y"
        'CountResult()  '20090622 andy edit  排名改以( 1.各科平均法  2, 訓練時數權重法) 計算 '200902 andy 
        Call GetStudRank(Me, OCIDValue1.Value, objconn)
        Common.MessageBox(Me, "成績結算成功!")
    End Sub

    '清除(刪除)目前畫面上未選取科目的學員結訓成績檔
    Public Shared Function chkResults(ByVal CourID As String, ByVal OCIDV1 As String, ByRef oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False 'True
        If CourID = "" Then Return rst
        If OCIDV1 = "" Then Return rst

        Dim pmsS As New Hashtable From {{"OCID", TIMS.CINT1(OCIDV1)}}
        Dim sql As String = ""
        sql &= " DELETE STUD_TRAININGRESULTS" & vbCrLf
        sql &= " WHERE SOCID IN (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID )" & vbCrLf
        sql &= $" AND CourID NOT IN ({CourID})" & vbCrLf
        DbAccess.ExecuteNonQuery(sql, oConn, pmsS)
        Return True
    End Function

    Sub SaveData1()
        Dim CourID As String = ""
        Dim SOCID As String = ""
        If Not GradeTable.Visible OrElse GradeTable.Style.Item("display") = "none" Then
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End If
        Try
            If Table4.Rows(0).Cells.Count = 3 Then
                Common.MessageBox(Me, "執行失敗，無學員成績資料!")
                Exit Sub
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End Try

        'sql = "SELECT * FROM STUD_TRAININGRESULTS WHERE SOCID"
        For i As Integer = 1 To Table4.Rows.Count - 1
            Dim MyRow As TableRow = Table4.Rows(i)
            'Dim tmpValue1 As String = "" '組合暫存空間。
            For j As Integer = 3 To MyRow.Cells.Count - 1
                Dim MyCell As TableCell = MyRow.Cells(j)
                Dim MyTextBox As TextBox = MyCell.Controls(0)
                If i = 1 Then
                    'CourID
                    tmpValue1 = "'" & Split(Replace(MyTextBox.ClientID, "Course", ""), "_")(0) & "'"
                    If CourID.IndexOf(tmpValue1) = -1 Then '重複字過濾
                        If CourID <> "" Then CourID += ","
                        CourID += tmpValue1
                    End If
                End If
                'SOCID 
                tmpValue1 = "'" & Split(Replace(MyTextBox.ClientID, "Course", ""), "_")(1) & "'"
                If SOCID.IndexOf(tmpValue1) = -1 Then '重複字過濾
                    If SOCID <> "" Then SOCID += ","
                    SOCID += tmpValue1
                End If
            Next
        Next
        If CourID <> "" Then
            If chkResults(CourID, OCIDValue1.Value, objconn) = False Then
                Common.MessageBox(Me, "儲存失敗!")
                Exit Sub
            End If
            CourID = $" AND CourID IN ({CourID})"
        End If
        If SOCID <> "" Then
            SOCID = $" AND SOCID IN ({SOCID})"
        End If

        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = $"SELECT * FROM STUD_TRAININGRESULTS WHERE 1=1 {CourID}{SOCID}"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)

        Dim InputOk As Boolean = True
        Dim SaveOk As Boolean = False
        Dim errMsg As String = ""
        'Dim err1, err2
        Dim err1 As String = ""
        Dim err2 As String = ""
        For i As Integer = 1 To Table4.Rows.Count - 1
            Dim MyRow As TableRow = Table4.Rows(i)
            For j As Integer = 3 To MyRow.Cells.Count - 1
                Dim MyCell As TableCell = MyRow.Cells(j)
                Dim MyTextBox As TextBox = MyCell.Controls(0)

                CourID = Split(Replace(MyTextBox.ClientID, "Course", ""), "_")(0)
                SOCID = Split(Replace(MyTextBox.ClientID, "Course", ""), "_")(1)
                MyTextBox.ForeColor = Color.Black
                If MyTextBox.Text <> "" Then
                    If TIMS.IsAbs(MyTextBox.Text) = False Then
                        err1 = "成績欄位填寫有誤!"
                        MyTextBox.ForeColor = Color.Red
                        InputOk = False
                    Else
                        If Convert.ToSingle(MyTextBox.Text) > 100 Then
                            err2 = "成績不得大於100!"
                            MyTextBox.ForeColor = Color.Red
                            InputOk = False
                        End If
                    End If
                End If

                'If MyTextBox.Text = "" Then
                '    '沒輸入成績的時候，必須刪除
                '    If dt.Select("CourID='" & CourID & "' and SOCID='" & SOCID & "'").Length <> 0 Then
                '        dt.Select("CourID='" & CourID & "' and SOCID='" & SOCID & "'")(0).Delete()
                '    End If
                'Else
                'STUD_TRAININGRESULTS
                If InputOk = True Then
                    Dim dr As DataRow = Nothing
                    Dim fff4 As String = $"CourID='{CourID}' and SOCID='{SOCID}'"
                    If dt.Select(fff4).Length = 0 Then '20100513 andy 沒輸入成績預設帶0
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("SOCID") = SOCID
                        dr("CourID") = CourID
                    Else
                        dr = dt.Select(fff4)(0)
                    End If
                    dr("Results") = TIMS.VAL1(MyTextBox.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    SaveOk = True
                End If
                'End If
            Next
        Next

        If InputOk = False Then
            errMsg = err1 & " " & err2
            Page.RegisterStartupScript("err", "<script>alert('" & errMsg & "');</script>")
            Exit Sub
        End If

        Try
            If SaveOk Then
                DbAccess.UpdateDataTable(dt, da)
                Common.MessageBox(Me, "儲存成功!")
            Else
                Common.MessageBox(Me, "無異動資料，請確認儲存條件!")
            End If
            'Button1_Click(sender, e)
            Call Search1()
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Throw ex
        End Try
    End Sub

    '儲存
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Call SaveData1()
    End Sub

    '消除勾選科目
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        'Dim MyCheck As HtmlInputCheckBox

        Try
            If GradeTable.Visible = False Or GradeTable.Style.Item("display") = "none" Then
                Common.MessageBox(Me, "執行失敗，無學員成績資料!")
                Exit Sub
            End If
            If Table4.Rows(0).Cells.Count = 3 Then
                Common.MessageBox(Me, "執行失敗，無學員成績資料!")
                Exit Sub
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, "執行失敗，無學員成績資料!")
            Exit Sub
        End Try

        For i As Integer = 3 To Table4.Rows(0).Cells.Count - 1
            Dim MyCheck As HtmlInputCheckBox
            Try
                MyCheck = Table4.Rows(0).Cells(i).Controls(0)
                If MyCheck.Checked = True Then
                    If ChooseClass.Value.IndexOf("," & MyCheck.Value & ",") <> -1 Then
                        ChooseClass.Value = Replace(ChooseClass.Value, "," & MyCheck.Value & ",", ",")
                    End If
                    If ChooseClass.Value.IndexOf("," & MyCheck.Value) <> -1 Then
                        ChooseClass.Value = Replace(ChooseClass.Value, "," & MyCheck.Value, "")
                    ElseIf ChooseClass.Value.IndexOf(MyCheck.Value & ",") <> -1 Then
                        ChooseClass.Value = Replace(ChooseClass.Value, MyCheck.Value & ",", "")
                    ElseIf ChooseClass.Value.IndexOf(MyCheck.Value) <> -1 Then
                        'ChooseClass.Value = Replace(ChooseClass.Value, MyCheck.Value, "")
                    End If
                End If
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Throw ex
            End Try
        Next
        CCIDVALUE.Value = ChooseClass.Value '20081208 Andy edit
        Call CreateGradeTable()
    End Sub

    '查詢班級資料(隱藏)
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            'msg.Visible = Truemsg.Text = "查無資料"
            Label1.Text = "選擇資料庫中的課程，查無數據!"
            Exit Sub
        End If
        'If drCC Is Nothing Then Exit Sub
        DBClass.Value = Get_AllCourID(OCIDValue1.Value, objconn)
        Label1.Text = "選擇資料庫中的課程，查無數據!"
        If DBClass.Value <> "" Then
            Label1.Text = "已選擇資料庫中的課程資料!"
        End If
        ChooseClass.Value = DBClass.Value

        GradeTable.Style.Item("display") = "none"
        PrintTB.Style.Item("display") = "none"
        BehaviorTable.Style.Item("display") = "none"
        class_grade.Checked = False
        behavior.Checked = False
        Button3.Disabled = True
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        GradeTable.Style.Item("display") = "none"
        PrintTB.Style.Item("display") = "none"
        BehaviorTable.Style.Item("display") = "none"
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        GradeTable.Style.Item("display") = "none"
        PrintTB.Style.Item("display") = "none"
        BehaviorTable.Style.Item("display") = "none"
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '選擇簽核方式。
    Private Sub RB_SignType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles RB_SignType.SelectedIndexChanged
        'Dim i As Integer
        TR_Signer2.Style.Item("display") = "none"
        TR_Signer.Style.Item("display") = "none"
        Bn_Print.Style.Item("display") = "none"

        Select Case RB_SignType.SelectedValue
            Case "1"
                Print1.Value = 1
                'TR_Signer2.Style.Item("display") = "none"
                TR_Signer.Style.Item("display") = TIMS.cst_inline1 '"inline"
                Bn_Print.Style.Item("display") = TIMS.cst_inline1 '"inline"
                cb_signer2.SelectedIndex = -1
                For i As Integer = 0 To cb_signer.Items.Count - 1
                    cb_signer.Items(i).Selected = True
                Next

            Case "2"
                Print1.Value = 2
                TR_Signer2.Style.Item("display") = TIMS.cst_inline1 '"inline"
                'TR_Signer.Style.Item("display") = "none"
                Bn_Print.Style.Item("display") = TIMS.cst_inline1 '"inline"
                cb_signer.SelectedIndex = -1
                For i As Integer = 0 To cb_signer2.Items.Count - 1
                    cb_signer2.Items.Item(i).Selected = True
                Next

        End Select

    End Sub

    '修正學員總分數排名
    Public Shared Sub UpdateRank(ByVal dt As DataTable, ByVal OCIDV1 As String, ByRef oConn As SqlConnection)
        '準備開始排名，先清除原排名
        If dt.Rows.Count > 0 Then
            Dim pmsU As New Hashtable From {{"OCID", OCIDV1}}
            Dim strSql As String = " UPDATE CLASS_STUDENTSOFCLASS SET RANK=null WHERE OCID=@OCID"
            DbAccess.ExecuteNonQuery(strSql, oConn, pmsU)
        End If

        Dim GradeTemp As Double = 0
        Dim m As Integer = 0
        For Each dr As DataRow In dt.Select("", "totalAvg desc")
            '假如目前的Temp值跟該學生的總成績不相同時
            '將最大成績存入Temp中 假如目前的Temp值與該學生的總成績相同時保留名次空位
            If GradeTemp <> dr("totalAvg") Then
                GradeTemp = dr("totalAvg")
                m += 1
            End If
            Dim pmsU2 As New Hashtable From {{"OCID", OCIDV1}, {"SOCID", dr("socid")}, {"Rank", m}}
            Dim strSql2 As String = " UPDATE CLASS_STUDENTSOFCLASS SET Rank =@Rank WHERE OCID=@OCID AND SOCID=@SOCID"
            DbAccess.ExecuteNonQuery(strSql2, oConn, pmsU2)
        Next
    End Sub

    '取出所有學員-並修改排名分數
    Public Shared Sub GetStudRank(ByVal MyPage As Page, ByVal OCIDV1 As String, ByRef oConn As SqlConnection)
        OCIDV1 = TIMS.ClearSQM(OCIDV1)
        If OCIDV1 = "" Then Exit Sub
        If MyPage Is Nothing Then Exit Sub
        '取出所有學員學號
        Dim pmsS As New Hashtable From {{"OCID", OCIDV1}}
        Dim sql As String = ""
        sql &= " SELECT SOCID FROM CLASS_STUDENTSOFCLASS"
        sql &= " WHERE OCID=@OCID AND StudStatus NOT IN (2,3)"
        sql &= " ORDER BY SOCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, pmsS)
        If dt.Rows.Count = 0 Then Exit Sub

        '設定一個空的table
        Dim sql_2 As String = ""
        sql_2 &= " SELECT SOCID ,0.0 totalAvg"
        sql_2 &= " FROM CLASS_STUDENTSOFCLASS"
        sql_2 &= " WHERE 1<>1"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql_2, oConn)

        For Each dr As DataRow In dt.Rows
            Dim dt3 As DataTable = GetGradeTable(MyPage, OCIDV1, Convert.ToString(dr("SOCID")), oConn)
            Dim flagcanCalc As Boolean = False
            If Not dt3 Is Nothing Then
                If dt3.Rows.Count > 0 Then flagcanCalc = True
            End If
            If flagcanCalc Then
                '填入空的table
                Dim dr2 As DataRow = dt2.NewRow
                dr2("SOCID") = dt3.Rows(0)("SOCID")
                dr2("totalAvg") = dt3.Rows(0)("totalAvg")
                dt2.Rows.Add(dr2)
            End If
        Next

        '修正學員總分數排名
        Call UpdateRank(dt2, OCIDV1, oConn)
    End Sub

    '090825
    Public Shared Function getStudResult(ByVal Item As Int16, ByVal OneSOCID As String, ByVal OCIDV1 As String, ByRef oConn As SqlConnection) As DataTable
        Dim sAllCourIDs As String = Get_AllCourID(OCIDV1, oConn)
        '先取出目前所選取的科目  
        Dim sql_1 As String = ""
        sql_1 &= " SELECT CourID,CourseName"
        sql_1 &= " ,case Classification1 when 1 then '學科' when 2 then '術科' end classType"
        sql_1 &= " ,Classification1 ,0 TotalHours" & vbCrLf
        sql_1 &= " FROM COURSE_COURSEINFO"
        If sAllCourIDs <> "" Then
            sql_1 &= " WHERE CourID IN (" & sAllCourIDs & ")" & vbCrLf
        Else
            sql_1 &= " WHERE 1<>1" & vbCrLf
        End If
        sql_1 &= " ORDER BY CourseName "
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql_1, oConn)
        OCIDV1 = TIMS.ClearSQM(OCIDV1)
        If OCIDV1 = "" Then Return dt1

        '取出所有學員成績    
        Dim sql_2 As String = ""
        sql_2 &= " SELECT '" & OCIDV1 & "' OCID ,SOCID ,COURID"
        sql_2 &= " ,'' COURSENAME ,RESULTS"
        sql_2 &= " ,'' CLASSTYPE ,null CLASSIFICATION1 ,0 as HOURS ,0 as RESULTS2"
        sql_2 &= " FROM STUD_TRAININGRESULTS"
        sql_2 &= " WHERE CourID in (" & sAllCourIDs & ")" & vbCrLf
        If OneSOCID <> "" Then
            sql_2 &= " AND SOCID in (" & OneSOCID & ")"
        Else
            sql_2 &= " AND SOCID in (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCIDV1 & "' AND STUDSTATUS NOT IN (2,3))"
        End If
        sql_2 &= " ORDER BY SOCID,CourseName"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql_2, oConn)

        '取出排課檔  
        Dim pmsS3 As New Hashtable From {{"OCID", OCIDV1}}
        Dim sql_3 As String = ""
        sql_3 &= " SELECT * FROM CLASS_SCHEDULE WHERE OCID=@OCID"
        sql_3 &= " ORDER BY SCHOOLDATE" & vbCrLf
        Dim dt3 As DataTable = DbAccess.GetDataTable(sql_3, oConn, pmsS3)

        '產生一個空的table-課程
        Dim sql_A As String = ""
        sql_A &= " SELECT CourID,CourseName"
        sql_A &= " ,case Classification1 when  1  then '學科'  when  2  then  '術科' end classType"
        sql_A &= " ,Classification1 ,0 TotalHours" & vbCrLf
        sql_A &= " FROM COURSE_COURSEINFO WHERE 1<>1 "
        Dim dtReult1 As DataTable = DbAccess.GetDataTable(sql_A, oConn)

        '產生一個空的table-學員
        Dim sql_B As String = ""
        sql_B &= " SELECT '" & OCIDV1 & "' OCID"
        sql_B &= " ,SOCID ,COURID ,'' COURSENAME ,RESULTS"
        sql_B &= " ,'' CLASSTYPE ,null CLASSIFICATION1 ,0 as HOURS ,0 as RESULTS2"
        sql_B &= " FROM STUD_TRAININGRESULTS"
        sql_B &= " WHERE 1<>1"
        Dim dtReult2 As DataTable = DbAccess.GetDataTable(sql_B, oConn)

        Dim CountHours As Int16 = 0
        For Each dr1 As DataRow In dt1.Rows       '列出所有選取的科目
            For Each dr3 As DataRow In dt3.Rows   '依序列出排課
                For i As Integer = 1 To 12
                    If Not IsDBNull(dr3("Class" & i)) Then
                        If dr3("Class" & i) = dr1("CourID") Then
                            dr1("TotalHours") += 1
                        End If
                    End If
                Next
            Next
            Dim dr As DataRow = dtReult1.NewRow      '填入空的table
            dr("CourID") = dr1("CourID")
            dr("CourseName") = dr1("CourseName")
            dr("classType") = dr1("classType")
            dr("Classification1") = dr1("Classification1")
            dr("TotalHours") = dr1("TotalHours")
            dtReult1.Rows.Add(dr)
        Next

        '列出所有選取的學員 
        For Each dr2 As DataRow In dt2.Rows
            For Each dr4 As DataRow In dtReult1.Rows
                If Convert.ToString(dr2("COURID")) = Convert.ToString(dr4("COURID")) Then
                    dr2("COURID") = dr4("COURID")
                    dr2("COURSENAME") = dr4("COURSENAME")
                    dr2("CLASSTYPE") = dr4("CLASSTYPE")
                    dr2("CLASSIFICATION1") = dr4("CLASSIFICATION1")
                    dr2("HOURS") = dr4("TotalHours")
                    dr2("RESULTS2") = dr2("RESULTS") * dr4("TotalHours")

                    Dim dr As DataRow = dtReult2.NewRow
                    dr("OCID") = dr2("OCID")
                    dr("SOCID") = dr2("SOCID")
                    dr("COURID") = dr2("COURID")
                    dr("COURSENAME") = dr2("COURSENAME")
                    dr("RESULTS") = dr2("RESULTS")
                    dr("CLASSTYPE") = dr2("CLASSTYPE")
                    dr("CLASSIFICATION1") = dr2("CLASSIFICATION1")
                    dr("HOURS") = dr2("HOURS")
                    dr("RESULTS2") = dr2("RESULTS2")
                    dtReult2.Rows.Add(dr)
                End If
            Next
        Next
        'Call TIMS.CloseDbConn(conn)

        Select Case Item
            Case 1
                '所選取的科目大於一個且 學員成績檔中只有登錄一個科目成績時 
                '-->以「各科平均法」計算
                'If dt1.Rows.Count >= 1 And dt1_2.Rows.Count = 1 Then
                '    ViewState("ResultTyp") = "1"
                'End If
                Return dt1
            Case 2
                Return dtReult2
        End Select
        Return Nothing
    End Function

    '產生個人成績檔案
    Public Shared Function GetGradeTable(ByRef MyPage As Page, ByVal OCIDV1 As String, ByVal SOCID As String, ByRef oConn As SqlConnection) As DataTable
        'ByVal TPlanID As String, ByVal DistID As String, ByVal OCIDValue1 As String, ByVal SOCID As String
        Dim dtReult As New DataTable
        dtReult.Columns.Add(New DataColumn("socid"))
        'dtReult.Columns.Add(New DataColumn("pClassCount"))
        'dtReult.Columns.Add(New DataColumn("sClassCount"))
        'dtReult.Columns.Add(New DataColumn("pTotal_1"))
        'dtReult.Columns.Add(New DataColumn("sTotal_1"))
        'dtReult.Columns.Add(New DataColumn("pTotal_2"))
        'dtReult.Columns.Add(New DataColumn("sTotal_2"))
        'dtReult.Columns.Add(New DataColumn("pHours"))
        'dtReult.Columns.Add(New DataColumn("sHours"))
        'dtReult.Columns.Add(New DataColumn("totalResults"))
        'dtReult.Columns.Add(New DataColumn("totalResults2"))
        dtReult.Columns.Add(New DataColumn("totalAvg"))
        'dtReult.Columns.Add(New DataColumn("pAvg"))
        'dtReult.Columns.Add(New DataColumn("sAvg"))
        'dtReult.Columns.Add(New DataColumn("ResultTyp"))
        'dtReult.Columns.Add(New DataColumn("percent1"))
        'dtReult.Columns.Add(New DataColumn("percent2"))

        'Dim Errmsg As String = ""
        'Dim i As Integer = 0
        'Dim sql As String = ""
        'Dim dt As New DataTable
        'Dim dt2 As New DataTable
        'Dim da As New SqlDataAdapter

        'Dim Errmsg As String = ""
        'Dim strVar13 As String = TIMS.GetGlobalVar(MyPage, "13", "1", oConn)
        'If strVar13 = "" Then
        '    Errmsg &= "尚未設定成績計算模式,請聯絡中心系統管理者" & vbCrLf
        '    Common.MessageBox(MyPage, Errmsg)
        '    Return Nothing
        'End If
        'CheckMode = Val(strVar13)

        'sql = "SELECT * FROM Sys_GlobalVar WHERE DistID='" & DistID & "' and TPlanID='" & TPlanID & "'"
        'dt = DbAccess.GetDataTable(sql, objconn)
        'If dt.Select("GVID='13'").Length = 0 Then
        '    Errmsg += "尚未設定成績計算模式,請聯絡中心系統管理者" & vbCrLf
        '    'Exit Sub
        'Else
        '    CheckMode = dt.Select("GVID='13'")(0)("ItemVar1")
        'End If

        Dim dt2 As New DataTable
        Dim AllClass As String = Get_AllCourID(OCIDV1, oConn)
        If AllClass = "" Then Return dtReult

        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me.Page, "發生錯誤!! " & ex.Message.ToString())
        '    'Finally 'conn.Close()da.Dispose()dt2.Dispose()
        'End Try

        dt2 = getStudResult(2, SOCID, OCIDV1, oConn)
        '***************************************
        ' 【計算成績方式】
        '        <各科平均法>
        '=======================================
        '  (1)學、術科百分比(無)  (2) 學、術科百分比(有)
        '    ex: 總平均是由(學科全部成績加總/學科總數 4  科) * 學科百分比 20 %      
        '        加上      (術科全部成績加總/術科總數 11 科) * 術科百分比 80 %      
        '        計算產生之結果。
        '---
        '        <訓練時數權重法> 
        '======================================= 
        '  (1)學、術科百分比(無)  
        '   總平均=(科)成績*(科)總時數 /(各科加總)總時數  
        '   總平均是由(各科成績與訓練時數加權/各科時數加總 __ 小時)計算產生之結果。 
        'ex:
        '總平均=
        '(學科成績與時數加權/時數加總 36 小時)* 學科百分比 40 %   
        '     +
        '(術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
        '/2 計算之結果。 

        '---
        '  (2) 學、術科百分比(有)
        '    ex: 學科平均=(科)成績*(科)總時數 /(各科加總)總時數  
        '        總平均是由(學科成績與訓練時數加權/學科時數 __ 小時) * 學科百分比 20 %      
        '        加上      (術科成績與訓練時數加權/術科時數 __ 小時) * 術科百分比 20 %        
        '        計算產生之結果。 
        '***************************************
        'Dim j As Integer = 0
        Dim pClassCount As Integer = 0              '學科科目總數
        Dim sClassCount As Integer = 0              '術科科目總數
        Dim pTotal_1 As Single = 0          '學科總成績---各科平均法
        Dim sTotal_1 As Single = 0        '術科總成績 
        Dim pTotal_2 As Single = 0        '學科總成績---訓練時數權重法
        Dim sTotal_2 As Single = 0         '術科總成績
        Dim pHours As Int16 = 0            '學科總時數
        Dim sHours As Int16 = 0            '術科總時數
        Dim totalResults As Integer = 0    '總成績 
        Dim totalResults2 As Integer = 0   '總成績(時數加權) 
        Dim totalAvg As Single = 0        '總平均 
        Dim pAvg As Single = 0           '學科總平均 
        Dim sAvg As Single = 0           '術科總平均 

        Dim iResultTyp As Integer = 0 'GetPara(13, DistID, TPlanID)      '成績計算方式
        Dim ipercent1 As Single = 0 'GetPara(17, DistID, TPlanID, 1)       '學科百分比
        Dim ipercent2 As Single = 0 'GetPara(17, DistID, TPlanID, 2)      '術科百分比
        Dim GlobalVar13_1 As String = TIMS.GetGlobalVar(MyPage, "13", "1", oConn)
        Dim GlobalVar17_1 As String = TIMS.GetGlobalVar(MyPage, "17", "1", oConn)
        Dim GlobalVar17_2 As String = TIMS.GetGlobalVar(MyPage, "17", "2", oConn)
        If GlobalVar13_1 <> "" Then iResultTyp = Val(GlobalVar13_1)
        If GlobalVar17_1 <> "" Then ipercent1 = Val(GlobalVar17_1)
        If GlobalVar17_2 <> "" Then ipercent2 = Val(GlobalVar17_2)

        'Dim ResultTyp As Integer = GetPara(13, DistID, TPlanID)      '成績計算方式
        'Dim percent1 As Single = GetPara(17, DistID, TPlanID, 1)       '學科百分比
        'Dim percent2 As Single = GetPara(17, DistID, TPlanID, 2)      '術科百分比

        If dt2 Is Nothing Then Return Nothing
        If dt2.Rows.Count = 0 Then Return Nothing

        For j As Integer = 0 To dt2.Rows.Count - 1
            If Convert.ToString(dt2.Rows(j)("Classification1")) = "1" Then        '學科
                pClassCount = pClassCount + 1
                pHours = pHours + Convert.ToInt16(dt2.Rows(j)("hours"))
                pTotal_2 = pTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                pTotal_1 = pTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))   '各科平均法： 1 '默認採用各科平均法(當未設定時)
            ElseIf Convert.ToString(dt2.Rows(j)("Classification1")) = "2" Then    '術科
                sClassCount = sClassCount + 1
                sHours = sHours + Convert.ToInt16(dt2.Rows(j)("hours"))
                sTotal_2 = sTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                sTotal_1 = sTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))
            End If
        Next
        If pClassCount = 0 Then
            pAvg = 0
        Else
            pAvg = pTotal_1 / pClassCount
        End If
        If sClassCount = 0 Then
            sAvg = 0
        Else
            sAvg = sTotal_1 / sClassCount
        End If
        totalResults = sTotal_1 + pTotal_1  '原始成績加總
        totalResults2 = sTotal_2 + pTotal_2 '成績加總(訓練時數權重法)
        If pHours = 0 And sHours = 0 Then
            iResultTyp = 1
        End If
        Select Case iResultTyp
            Case 2        '訓練時數權重法： 2
                '總平均= (學科成績與時數加權/時數加總 36 小時)* 學科百分比 40% + (術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
                '/2 計算之結果。 
                If ipercent1 = 0 And ipercent2 = 0 Then  '百分比未設定時
                    If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
                        totalAvg = (sTotal_2 / sHours)
                    ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
                        totalAvg = (pTotal_2 / pHours)
                    ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                        If sHours = 0 And pHours <> 0 Then
                            totalAvg = (pTotal_2 / pHours)
                        ElseIf sHours <> 0 And pHours = 0 Then
                            totalAvg = (sTotal_2 / sHours)
                        ElseIf sHours <> 0 And pHours <> 0 Then
                            totalAvg = ((sTotal_2 / sHours) + (pTotal_2 / pHours)) / 2
                        End If
                    Else
                        totalAvg = 0
                    End If
                Else
                    If pClassCount = 0 And sClassCount <> 0 And sHours <> 0 Then      '只有術科
                        totalAvg = (sTotal_2 / sHours) * ipercent2
                    ElseIf pClassCount <> 0 And sClassCount = 0 And pHours <> 0 Then  '只有學科
                        totalAvg = (pTotal_2 / pHours) * ipercent1
                    ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                        If sHours = 0 And pHours <> 0 Then
                            totalAvg = (pTotal_2 / pHours) * ipercent1
                        ElseIf sHours <> 0 And pHours = 0 Then
                            totalAvg = (sTotal_2 / sHours) * ipercent2
                        ElseIf sHours <> 0 And pHours <> 0 Then
                            totalAvg = (pTotal_2 / pHours) * ipercent1 + (sTotal_2 / sHours) * ipercent2
                        Else
                            totalAvg = 0
                        End If
                    Else
                        totalAvg = 0
                    End If
                End If
            Case Else     '各科平均法： 1 '默認採用各科平均法(當未設定時)
                If ipercent1 = 0 And ipercent2 = 0 Then  '百分比未設定時
                    totalAvg = totalResults / (pClassCount + sClassCount)
                Else
                    If pClassCount = 0 And sClassCount <> 0 Then      '只有術科
                        totalAvg = (sTotal_1 / sClassCount) * ipercent2
                    ElseIf pClassCount <> 0 And sClassCount = 0 Then  '只有學科
                        totalAvg = (pTotal_1 / pClassCount) * ipercent1
                    ElseIf pClassCount <> 0 And sClassCount <> 0 Then
                        totalAvg = (pTotal_1 / pClassCount) * ipercent1 + (sTotal_1 / sClassCount) * ipercent2
                    Else
                        totalAvg = 0
                    End If
                End If
        End Select

        'Dim dr As DataRow
        Dim dr As DataRow = dtReult.NewRow
        dtReult.Rows.Add(dr)
        dr("socid") = SOCID
        'dr("pClassCount") = pClassCount      '學科科目總數
        'dr("sClassCount") = sClassCount      '術科科目總數
        'dr("pTotal_1") = pTotal_1            '學科總成績---各科平均法
        'dr("sTotal_1") = sTotal_1            '術科總成績 
        'dr("pTotal_2") = pTotal_2            '學科總成績---訓練時數權重法 
        'dr("sTotal_2") = sTotal_2            '術科總成績 
        'dr("pHours") = pHours                '學科總時數
        'dr("sHours") = sHours                '術科總時數
        'dr("totalResults") = totalResults    '總成績 
        'dr("totalResults2") = totalResults2  '總成績(時數加權) 
        dr("totalAvg") = TIMS.ROUND(totalAvg, 1)     '總平均 
        'dr("pAvg") = TIMS.Round(pAvg, 1)             '學科總平均 
        'dr("sAvg") = TIMS.Round(sAvg, 1)             '術科總平均 
        'dr("ResultTyp") = ResultTyp   '成績計算方式 
        'dr("percent1") = percent1     '學科百分比
        'dr("percent2") = percent2     '術科百分比

        '建立資料---End
        Return dtReult
    End Function

    '取出已有登錄之學員成績
    Public Shared Function GetStudResultList(ByVal SOCIDV1 As String, ByVal OCIDV1 As String, ByVal ChooseClassV1 As String, ByRef oConn As SqlConnection) As DataTable
        Dim dt4 As New DataTable
        If OCIDV1 = "" Then Return dt4
        If SOCIDV1 = "" Then Return dt4

        '取出所有學員
        Dim pmsS As New Hashtable From {{"OCID", OCIDV1}}
        Dim sql_1 As String = ""
        sql_1 &= " SELECT a.Name,b.SOCID, b.OCID, b.StudStatus,c.FTDate ,b.StudentID"
        sql_1 &= " FROM STUD_STUDENTINFO a "
        sql_1 &= " JOIN (SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID ) b ON b.SID=a.SID"
        sql_1 &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID "
        sql_1 &= " ORDER BY b.StudentID"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql_1, oConn, pmsS)

        '課程
        Dim flag2 As Boolean = False
        Dim sql_2 As String = ""
        sql_2 &= " SELECT CourseName,CourID FROM COURSE_COURSEINFO WHERE 1=1"
        If ChooseClassV1 <> "" Then
            flag2 = True
            sql_2 &= " and CourID in (" & ChooseClassV1 & ")"
        Else
            Dim tValue1 As String = Get_AllCourID(OCIDV1, oConn)
            If tValue1 <> "" Then
                flag2 = True
                sql_2 &= " and CourID in (" & tValue1 & ")"
            End If
        End If
        If Not flag2 Then sql_2 &= " and 1<>1"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql_2, oConn)

        '取出已有登錄之學員成績
        Dim pmsS3 As New Hashtable From {{"OCID", OCIDV1}}
        Dim sql_3 As String = ""
        sql_3 &= " SELECT a.SOCID, a.COURID, i.CourseName, a.RESULTS "
        sql_3 &= " FROM STUD_TRAININGRESULTS a "
        sql_3 &= " JOIN CLASS_STUDENTSOFCLASS b ON b.socid=a.socid "
        sql_3 &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID "
        sql_3 &= " JOIN ID_CLASS g on g.CLSID=c.CLSID "
        sql_3 &= " left JOIN COURSE_COURSEINFO i  ON a.CourID=i.CourID "
        sql_3 &= " where b.OCID=@OCID"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sql_3, oConn, pmsS3)

        '空表格
        Dim sql_4 As String = ""
        sql_4 &= " SELECT '' Name" & vbCrLf
        sql_4 &= " ,'' SOCID" & vbCrLf
        sql_4 &= " ,'' OCID,null StudStatus" & vbCrLf
        sql_4 &= " ,'' COURID" & vbCrLf
        sql_4 &= " ,'' CourseName" & vbCrLf
        sql_4 &= " ,'' RESULTS" & vbCrLf
        sql_4 &= " ,'' FTDate" & vbCrLf
        dt4 = DbAccess.GetDataTable(sql_4, oConn)

        Dim fff4 As String = "SOCID='" & SOCIDV1 & "'"
        If dt1.Select(fff4).Length = 0 Then Return dt4

        dt4.Clear()
        For Each dr1 As DataRow In dt1.Select(fff4)
            For Each dr2 As DataRow In dt2.Rows
                Dim dr As DataRow = dt4.NewRow
                dr("Name") = dr1("Name")
                dr("SOCID") = dr1("SOCID")
                dr("OCID") = dr1("OCID")
                dr("StudStatus") = dr1("StudStatus")
                dr("COURID") = dr2("COURID")
                dr("CourseName") = dr2("CourseName")
                dr("FTDate") = dr1("FTDate")

                For Each dr3 As DataRow In dt3.Select(fff4)
                    If dr2("COURID") = dr3("COURID") Then
                        dr("RESULTS") = Convert.ToString(dr3("RESULTS"))
                    End If
                Next
                dt4.Rows.Add(dr)
            Next
        Next
        Return dt4
    End Function

    'dl_studName 下拉選單
    Sub getDlstudList()
        dl_studName.Items.Clear()
        tb_EditResult.Visible = False 'Call getStudResultList(objconn)
        Msg2.Text = "查無資料!"

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Exit Sub

        Dim sql As String = ""
        sql &= " SELECT a.Name" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(b.StudentID) StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(b.StudentID)+'  '+a.Name StudentName" & vbCrLf
        'sql &= " ,replace(b.StudentID,c.Years + '0' + g.ClassID + c.CyclType,'') StudentID" & vbCrLf
        'sql &= " ,replace(b.StudentID,c.Years + '0' + g.ClassID + c.CyclType,'') + '  ' + Name StudentName" & vbCrLf
        sql &= " ,b.SOCID ,b.OCID ,b.StudStatus ,b.CloseDate ,c.FTDate" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS b" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO a ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID" & vbCrLf
        sql &= " JOIN ID_CLASS g on g.CLSID=c.CLSID" & vbCrLf
        sql &= " WHERE b.OCID=@OCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then Exit Sub

        'tb_EditResult.Visible = True
        Msg2.Text = ""
        With dl_studName
            .DataSource = dt
            .DataTextField = "StudentName"
            .DataValueField = "SOCID"
            .DataBind()
        End With

        Dim SOCIDV1 As String = dl_studName.SelectedValue
        Dim OCIDV1 As String = OCIDValue1.Value
        Dim ChooseClassV1 As String = ChooseClass.Value
        Dim dt4 As DataTable = GetStudResultList(SOCIDV1, OCIDV1, ChooseClassV1, objconn)
        If dt4.Rows.Count > 0 Then
            tb_EditResult.Visible = True
            dgInput.DataSource = dt4
            dgInput.DataBind()
        End If

        'Msg2.Text = ""
        If dgInput.Items.Count = 0 Then
            Msg2.Text = "查無資料!"
        End If
    End Sub

    Private Sub dl_studName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dl_studName.SelectedIndexChanged
        Dim SOCIDV1 As String = dl_studName.SelectedValue
        Dim OCIDV1 As String = OCIDValue1.Value
        Dim ChooseClassV1 As String = ChooseClass.Value
        Dim dt4 As DataTable = GetStudResultList(SOCIDV1, OCIDV1, ChooseClassV1, objconn)
        If dt4.Rows.Count > 0 Then
            tb_EditResult.Visible = True
            dgInput.DataSource = dt4
            dgInput.DataBind()
        End If
    End Sub

    Private Sub add_Result_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles add_Result.Click
        GradeTable.Visible = False
        PrintTB.Visible = False

        Call getDlstudList()
    End Sub

    Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
        'getDlstudList()
    End Sub

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            'msg.Visible = True'msg.Text = "查無資料"'Exit Sub
            Common.MessageBox(Me.Page, TIMS.cst_NODATAMsg3)
            Exit Sub
        End If
        'If drCC Is Nothing Then Exit Sub

        'Dim da As New SqlDataAdapter
        Dim v_SOCID As String = TIMS.ClearSQM(dl_studName.SelectedValue)
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Try
            TIMS.OpenDbConn(objconn)
            Dim sql As String = ""
            sql = " SELECT * FROM STUD_TRAININGRESULTS WHERE SOCID=@SOCID"
            Dim sCmd As New SqlCommand(sql, objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(v_SOCID)
                dt1.Load(.ExecuteReader())
            End With

            If dt1.Rows.Count > 0 Then
                Dim RecordType As Int16 = 0                   'RecordType 1:新增;2:修改
                For Each item As DataGridItem In dgInput.Items
                    Dim RESULTS As TextBox = item.FindControl("RESULTS")
                    Dim CourIDVal As HtmlInputHidden = item.FindControl("CourIDVal")
                    Dim SocidVal As HtmlInputHidden = item.FindControl("SocidVal")
                    sql = ""
                    If RESULTS.Text <> "" Then
                        If TIMS.IsAbs(RESULTS.Text) = False Then
                            Common.MessageBox(Me.Page, "成績欄位填寫有誤!")
                            Exit Sub
                        End If
                        If Convert.ToSingle(RESULTS.Text) > 100 Then
                            Common.MessageBox(Me.Page, "成績不得大於100!")
                            Exit Sub
                        End If
                    End If
                    If Trim(RESULTS.Text) <> "" Then
                        For Each dr As DataRow In dt1.Rows               '修改
                            If CourIDVal.Value = dr("CourID") And RESULTS.Text <> Convert.ToString(dr("Results")) Then
                                sql = " update STUD_TRAININGRESULTS set Results =@Results, ModifyAcct=@ModifyAcct,ModifyDate = getdate()"
                                sql &= " where SOCID=@SOCID and CourID=@CourID"
                                RecordType = 2
                                Dim uCmd As New SqlCommand(sql, objconn)
                                With uCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("Results", SqlDbType.NVarChar).Value = RESULTS.Text
                                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SocidVal.Value)
                                    .Parameters.Add("CourID", SqlDbType.BigInt).Value = Val(CourIDVal.Value)
                                    .ExecuteNonQuery()
                                End With
                            End If
                        Next
                    End If
                    '刪除
                    If Trim(RESULTS.Text) = "" And dt1.Select("CourID='" & CourIDVal.Value & "'").Length <> 0 Then
                        sql = "DELETE  STUD_TRAININGRESULTS where SOCID=@SOCID and CourID=@CourID"
                        Dim dCmd As New SqlCommand(sql, objconn)
                        With dCmd
                            .Parameters.Clear()
                            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SocidVal.Value)
                            .Parameters.Add("CourID", SqlDbType.BigInt).Value = Val(CourIDVal.Value)
                            .ExecuteNonQuery()
                        End With
                    End If
                    '新增
                    If RESULTS.Text <> "" And RecordType <> 2 And dt1.Select("CourID='" & CourIDVal.Value & "'").Length = 0 Then
                        sql = " insert into STUD_TRAININGRESULTS (SOCID,CourID,Results,ModifyAcct,ModifyDate)"
                        sql &= " VALUES (@SOCID,@CourID,@Results,@ModifyAcct,GETDATE())"
                        Dim iCmd As New SqlCommand(sql, objconn)
                        With iCmd
                            .Parameters.Clear()
                            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SocidVal.Value)
                            .Parameters.Add("CourID", SqlDbType.BigInt).Value = Val(CourIDVal.Value)
                            .Parameters.Add("Results", SqlDbType.NVarChar).Value = RESULTS.Text
                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .ExecuteNonQuery()
                        End With
                    End If

                    'If sql <> "" Then
                    '    With da
                    '        .SelectCommand = New SqlCommand(sql, objconn)
                    '        .SelectCommand.ExecuteNonQuery()
                    '    End With
                    'End If
                Next
            End If
            Common.MessageBox(Me.Page, "儲存成功!!")
            If dl_studName.SelectedIndex < dl_studName.Items.Count - 1 Then
                dl_studName.SelectedIndex = (dl_studName.SelectedIndex) + 1
            ElseIf dl_studName.SelectedIndex = dl_studName.Items.Count - 1 Then
                dl_studName.SelectedIndex = 0
            End If

            Dim SOCIDV1 As String = dl_studName.SelectedValue
            Dim OCIDV1 As String = OCIDValue1.Value
            Dim ChooseClassV1 As String = ChooseClass.Value
            Dim dt4 As DataTable = GetStudResultList(SOCIDV1, OCIDV1, ChooseClassV1, objconn)
            If dt4.Rows.Count > 0 Then
                tb_EditResult.Visible = True
                dgInput.DataSource = dt4
                dgInput.DataBind()
            End If

            'conn.Close()
            'da.Dispose()
            dt1.Dispose()
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me.Page, "發生錯誤!! " & ex.Message.ToString())
        End Try
    End Sub

    Private Sub dgInput_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles dgInput.ItemDataBound
        Dim RESULTS As TextBox = e.Item.FindControl("RESULTS")
        Dim CourIDVal As HtmlInputHidden = e.Item.FindControl("CourIDVal")
        Dim SocidVal As HtmlInputHidden = e.Item.FindControl("SocidVal")
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Select Case drv("StudStatus")
                    Case 1, 4
                        RESULTS.Enabled = True
                    Case 2, 3
                        RESULTS.Enabled = False
                        bt_save.Enabled = False
                        vsTitle = "學員狀態為離訓、退訓"
                        TIMS.Tooltip(RESULTS, vsTitle)
                        TIMS.Tooltip(bt_save, vsTitle)
                    Case 5
                        If TIMS.Check_Auth_RendClass(drv("OCID"), dtArc) Then
                            RESULTS.Enabled = True
                        Else
                            Select Case sm.UserInfo.RoleID
                                Case 0, 1
                                    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days2 Then
                                        RESULTS.Enabled = False
                                        bt_save.Enabled = False
                                        vsTitle = "超過" & Days2 & "天後,限制輸入"
                                        TIMS.Tooltip(RESULTS, vsTitle)
                                        TIMS.Tooltip(bt_save, vsTitle)
                                    Else
                                        RESULTS.Enabled = True
                                    End If
                                Case Else
                                    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) > Days1 Then
                                        RESULTS.Enabled = False
                                        bt_save.Enabled = False
                                        vsTitle = "超過" & Days1 & "天後,限制輸入"
                                        TIMS.Tooltip(RESULTS, vsTitle)
                                        TIMS.Tooltip(bt_save, vsTitle)
                                    Else
                                        RESULTS.Enabled = True
                                    End If
                            End Select
                        End If
                End Select
                RESULTS.Text = Convert.ToString(drv("RESULTS"))
                CourIDVal.Value = Convert.ToString(drv("CourID"))
                SocidVal.Value = Convert.ToString(drv("SOCID"))
                If Convert.ToString(drv("StudStatus")) = "2" Or Convert.ToString(drv("StudStatus")) = "3" Then
                    RESULTS.ReadOnly = True
                    bt_save.Enabled = False
                    vsTitle = "學員狀態為離訓、退訓"
                    TIMS.Tooltip(RESULTS, vsTitle)
                    TIMS.Tooltip(bt_save, vsTitle)
                End If
        End Select
    End Sub

    '列印空白成績表
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            'msg.Visible = True'msg.Text = "查無資料"'Exit Sub
            Common.MessageBox(Me.Page, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        If SelSocid.Value = "" Then
            Common.MessageBox(Me.Page, "請選擇學員")
            Exit Sub
        End If
        '200902 andy 
        Call GetStudRank(Me, OCIDValue1.Value, objconn)
        Dim strScript As String = $"<script>CheckPrint({int_BasicPoint},1);</script>"
        Page.RegisterStartupScript("", strScript)
    End Sub

    '列印成績表
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'Me.Button5.Attributes.Add("OnClick", "return CheckPrint('" & ReportQuery.GetSmartQueryPath & "'," & BasicPoint & ");")
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            'msg.Visible = True'msg.Text = "查無資料"'Exit Sub
            Common.MessageBox(Me.Page, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        If SelSocid.Value = "" Then
            Common.MessageBox(Me.Page, "請選擇學員")
            Exit Sub
        End If
        '200902 andy 
        Call GetStudRank(Me, OCIDValue1.Value, objconn)

        Dim strScript As String = String.Concat("<script>", "CheckPrint(", int_BasicPoint, ",2);", "</script>")
        Page.RegisterStartupScript("", strScript)
    End Sub

    '計算 儲存
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = " SELECT * FROM STUD_CONDUCT WHERE SOCID=@SOCID"
        Dim sCmd As New SqlCommand(sql, objconn)

        'Dim sql_CS1 As String = " SELECT 1 FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID AND OCID=@OCID"
        'Dim sCmd_CS1 As New SqlCommand(sql_CS1, objconn)

        Dim sql_A As String = "" & vbCrLf
        sql_A &= " INSERT INTO STUD_CONDUCT (SOCID ,TECHPOINT,REMEDPOINT,MINUSLEAVE,MINUSSANCTION,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql_A &= " VALUES( @SOCID ,@TECHPOINT,@REMEDPOINT,@MINUSLEAVE,@MINUSSANCTION,@MODIFYACCT,getdate())" & vbCrLf
        Dim iCmd As New SqlCommand(sql_A, objconn)

        Dim sql_U As String = ""
        sql_U &= " UPDATE STUD_CONDUCT" & vbCrLf
        sql_U &= " SET TECHPOINT=@TECHPOINT ,REMEDPOINT=@REMEDPOINT ,MINUSLEAVE=@MINUSLEAVE ,MINUSSANCTION=@MINUSSANCTION" & vbCrLf
        sql_U &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        sql_U &= " WHERE SOCID=@SOCID" & vbCrLf
        Dim uCmd As New SqlCommand(sql_U, objconn)

        '操行成績
        Dim sql_U2 As String = " UPDATE CLASS_STUDENTSOFCLASS SET BehaviorResult=@BehaviorResult WHERE SOCID=@SOCID AND OCID=@OCID"
        Dim uCmd_CS2 As New SqlCommand(sql_U2, objconn)

        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)

        For Each eItem As DataGridItem In DataGrid2.Items
            Dim mytext1 As TextBox = eItem.Cells(cst_iDG2_COL_導師加減分).FindControl("TextBox2")
            Dim mytext2 As TextBox = eItem.Cells(cst_iDG2_COL_教務課加減分).FindControl("TextBox3")
            Dim CreditPoints As DropDownList = eItem.FindControl("CreditPoints") '是否核發結訓證書
            Dim hidDataKeys As HtmlInputHidden = eItem.FindControl("hidDataKeys")
            Dim SOCID_VAL As String = TIMS.ClearSQM(hidDataKeys.Value)
            mytext1.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(mytext1.Text))
            mytext2.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(mytext2.Text))

            Using dtCS1 As New DataTable
                Dim sql_CS1 As String = " SELECT 1 FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID AND OCID=@OCID"
                Using sCmd_CS1 As New SqlCommand(sql_CS1, objconn)
                    With sCmd_CS1
                        .Parameters.Clear()
                        .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL ' DataGrid2.DataKeys(i)
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = Hid_OCID1.Value ' DataGrid2.DataKeys(i)
                        dtCS1.Load(.ExecuteReader())
                    End With
                End Using
                If TIMS.dtNODATA(dtCS1) Then Continue For
            End Using

            'STUD_CONDUCT
            Dim dtCT As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL ' DataGrid2.DataKeys(i)
                dtCT.Load(.ExecuteReader())
            End With

            If dtCT.Rows.Count = 0 Then
                'STUD_CONDUCT
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL
                    .Parameters.Add("TECHPOINT", SqlDbType.VarChar).Value = TIMS.GetValue1(mytext1.Text) '導師加減分
                    .Parameters.Add("REMEDPOINT", SqlDbType.VarChar).Value = TIMS.GetValue1(mytext2.Text) '教務課加減分
                    .Parameters.Add("MINUSLEAVE", SqlDbType.Float).Value = TIMS.GetValue1(eItem.Cells(cst_iDG2_COL_出勤扣分).Text)
                    .Parameters.Add("MINUSSANCTION", SqlDbType.Float).Value = TIMS.GetValue1(eItem.Cells(cst_iDG2_COL_獎懲扣分).Text)
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                End With
            Else
                'STUD_CONDUCT
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("TECHPOINT", SqlDbType.VarChar).Value = TIMS.GetValue1(mytext1.Text) '導師加減分
                    .Parameters.Add("REMEDPOINT", SqlDbType.VarChar).Value = TIMS.GetValue1(mytext2.Text) '教務課加減分
                    .Parameters.Add("MINUSLEAVE", SqlDbType.Float).Value = TIMS.GetValue1(eItem.Cells(cst_iDG2_COL_出勤扣分).Text)
                    .Parameters.Add("MINUSSANCTION", SqlDbType.Float).Value = TIMS.GetValue1(eItem.Cells(cst_iDG2_COL_獎懲扣分).Text)
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL 'DataGrid2.DataKeys(i)
                    .ExecuteNonQuery()
                End With
            End If

            If Convert.ToString(eItem.Cells(cst_iDG2_COL_操行成績).Text) <> "" AndAlso IsNumeric(eItem.Cells(cst_iDG2_COL_操行成績).Text) Then
                'CLASS_STUDENTSOFCLASS
                With uCmd_CS2
                    .Parameters.Clear()
                    .Parameters.Add("BehaviorResult", SqlDbType.Float).Value = TIMS.VAL1(Int(eItem.Cells(cst_iDG2_COL_操行成績).Text))
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL ' DataGrid2.DataKeys(i)
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = Hid_OCID1.Value
                    .ExecuteNonQuery()
                End With
            End If

            If CreditPoints IsNot Nothing Then
                '未選擇／請選擇-CreditPoints-是否核發結訓證書 1:是／0:否
                Dim oCreditPoints As Object = Convert.DBNull
                Select Case CreditPoints.SelectedIndex
                    Case 1 '是 'dr("CreditPoints") = True '1
                        oCreditPoints = 1
                    Case 2 '否 'dr("CreditPoints") = False '0
                        oCreditPoints = 0
                End Select
                Dim sql_U3 As String = " UPDATE CLASS_STUDENTSOFCLASS SET CREDITPOINTS=@CREDITPOINTS WHERE SOCID=@SOCID AND OCID=@OCID"
                Using uCmd_CS3 As New SqlCommand(sql_U3, objconn)
                    With uCmd_CS3
                        .Parameters.Clear()
                        .Parameters.Add("CREDITPOINTS", SqlDbType.Float).Value = oCreditPoints
                        .Parameters.Add("SOCID", SqlDbType.VarChar).Value = SOCID_VAL ' DataGrid2.DataKeys(i)
                        .Parameters.Add("OCID", SqlDbType.VarChar).Value = Hid_OCID1.Value
                        .ExecuteNonQuery()
                    End With
                End Using
            End If
        Next

        Common.MessageBox(Me, "儲存成功!!")
        Call Search1()

    End Sub

End Class
