Partial Class SD_11_001
    Inherits AuthBasePage

    'STUD_QUESTIONARY 'Stud_Questionary.zip 
    'VIEW_STUDSURVEY 'V_STUDQUESTION4 [2016] STUD_RESULTSTUDDATA
    '(VIEW_QUESTIONARY1 VIEW_QUESTIONARY2 VIEW_QUESTIONARY3 VIEW_QUESTIONARY4) [Old]
    'SELECT * FROM STUD_SURVEY WHERE ROWNUM <=10 AND SOCID IN (SELECT DISTINCT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID =86136)
    'Const cst_downloadfileName As String = "Stud_Questionary.zip"

    'Dim sqlAdapter As SqlDataAdapter
    'Dim trans As SqlTransaction
    'Dim dt_Stud As DataTable

    Dim blnPrint2016 As Boolean = False '(職前且為20160501結訓後課程)
    Dim dtDegree As DataTable
    Dim dtMilitary As DataTable
    Dim dtIdentity As DataTable

    Dim FillFormDate As String '讀卡日期 
    Dim SubNo, SOCID, StudentID12, StdName, StdPID, Sex, BirthYear, BirthMonth, BirthDate As String
    Dim StudID As String  '學號 從1開始
    Dim DegreeID As String  '學歷 
    Dim MilitaryID As String '兵役 
    'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
    'BY AMU 2009-07-30
    '非署(局)屬狀況加入 Stud_ResultIdentData.IdentityID  BY AMU 2009-08-25
    Dim sIdentityID As String  '身分 
    Dim Q7 As String   '動機 'Stud_ResultStudData Q7	參訓職訓動機	int
    Dim Q8 As String   '動向 'Stud_ResultStudData Q8	結訓後動向	int
    Dim Q9 As String   '有無工作(有/無) 'Stud_ResultStudData Q9	受訓前有無在工作 char(1)
    Dim Q9Y As String  '有在工作 'Stud_ResultStudData Q9Y	受訓前有無在工作 int
    Dim Q10 As String  '找工作(有/無) 'Stud_ResultStudData  Q10	受訓前有無找工作 char(1)
    'Dim Q11 As String  '找到工作(有/無) 'Stud_ResultStudData Q11 參加訓練後	char(1)
    Dim Q11N As String '工作幫助 'Stud_ResultStudData Q11N	參加訓練後	int
    'Dim Q12 As String  '改進'Stud_ResultTwelveData Q12 資料卡十二項檔	int
    'Dim gQ12A As String  '您參加本次訓練後是否覺得滿意，若不滿意需要改進的為何 VARCHAR Y/N 滿意/不滿意
    'Dim gQ12B As String  ''□1.參訓職類不符就業市場需求  □2.教學課程安排不當  □3.訓練師專業及熱忱不足 □4.訓練設備不符產業需求      □5.訓練時數不足 ' VARCHAR(20)逗號分
    Dim gQ12v1 As String = "" '1.參訓職類符合就業市場需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
    Dim gQ12v2 As String = "" '2.教學課程安排 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
    Dim gQ12v3 As String = "" '3.訓練師專業及熱忱 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
    Dim gQ12v4 As String = "" '4.訓練設備符合產業需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
    Dim gQ12v5 As String = "" '5.訓練時數 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意

    'Stud_Questionary
    'Dim sPage_Url As String = ""
    'Const Cst_Page_Url As String = "SD_11_001_add.aspx"
    Const Cst_Page_Url_A12 As String = "SD_11_001_add12.aspx" '原(Old) /自辦
    Const Cst_Page_Url_A16 As String = "SD_11_001_add16.aspx" '2016年5月

    'Const cst_printFN16 As String = "SD_11_001_16"
    Const Cst_defQA2 As String = "A2" '預設問卷A2 'A'B 'select * from ID_Questionary 'select * from Plan_Questionary where rownum <=10
    Const Cst_defQA16 As String = "A16" '20160501問卷A16 
    'INSERT INTO ID_QUESTIONARY (QID,QNAME,QUESTIONARY,QNOTE ) VALUES (N'4',N'A3',N' ',N'訓練期末學員滿意度調查表2016(職前訓練)')

    'Dim vsQName As String = ""
    'Dim vsQID As String = ""
    'vsQName = dr("QName").ToString
    'vsQID = dr("QID").ToString

    'colArray
    Const cst_FillFormDate As Integer = 0 '讀卡日期
    Const cst_StudID As Integer = 1 '學號
    Const cst_DegreeID As Integer = 2 '學歷
    Const cst_MilitaryID As Integer = 3 '兵役
    Const cst_IdentityID As Integer = 4 '身分
    Const cst_Q7 As Integer = 5 '動機
    Const cst_Q8 As Integer = 6 '動向
    Const cst_Q9 As Integer = 7 '有無工作
    Const cst_Q9Y As Integer = 8 '有在工作
    Const cst_Q10 As Integer = 9 '找工作
    'Const cst_Q11 As Integer = 10 '找到工作(2014停用)
    Const cst_Q11N As Integer = 10 '11 '工作幫助
    Const cst_Q12v1 As Integer = 11 '12 '1.參訓職類符合就業市場需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
    Const cst_Q12v2 As Integer = 12 '13 '2.教學課程安排 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
    Const cst_Q12v3 As Integer = 13 '14 '3.訓練師專業及熱忱 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
    Const cst_Q12v4 As Integer = 14 '15 '4.訓練設備符合產業需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
    Const cst_Q12v5 As Integer = 15 '16 '5.訓練時數 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
    Const cst_Q_Answer_Start_pos As Integer = 16 '17 '其他問卷答案起始位置。
    Const cst_ptPrint As String = "Print" '列印空白表
    Const cst_ptInsert As String = "insert" '新增
    Const cst_ptCheck As String = "check" '查詢
    Const cst_ptEdit As String = "Edit" '修改
    Const cst_ptPrint2 As String = "Print2" '列印單1學員
    Const cst_ptDel As String = "del" '清除重填
    Dim vMsg As String = ""

#Region "(No Use)"

    'Const cst_Q12A As Integer = 12
    'Const cst_Q12B As Integer = 13

    'Try
    'Catch ex As Exception
    '    DbAccess.RollbackTrans(trans)
    '    sr.Close()
    '    srr.Close()
    '    Throw ex
    'End Try

    'If CInt(StudID) >= 100 Then
    '    sql = ""
    '    sql &= " SELECT * FROM stud_resultstuddata "
    '    sql &= " WHERE SOCID IN ( SELECT SOCID FROM Class_StudentsOfClass WHERE (OCID =" & OCIDValue1.Value & " AND RIGHT(studentid,3)='" & RIGHT(StudID, 2) & "' AND LEN(studentid)=12) ) "
    'Else
    '    sql = ""
    '    sql &= " SELECT * FROM stud_resultstuddata "
    '    sql &= " WHERE SOCID IN ( SELECT SOCID FROM Class_StudentsOfClass WHERE (OCID =" & OCIDValue1.Value & " AND RIGHT(studentid,2)='" & RIGHT(StudID, 2) & "' AND LEN(studentid)=11) ) "
    'End If

#End Region

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在-------------------------- Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        'trButton13.Visible = False '自辦／職前計畫顯示。
        trTrnPre1.Visible = False '自辦／職前計畫顯示。
        If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(sm.UserInfo.TPlanID) > -1 Then trTrnPre1.Visible = True '職前計畫顯示。

        If Not IsPostBack Then
            Call cCreate1()

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", , "search")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()
        msg.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        'hidSearchTag.Value = ""
        search.Attributes("onclick") = "javascript:return search1()"

        StudentTable.Style.Item("display") = "none"

        Dim sQuestionarySearchStr As String = "" 'Session("QuestionarySearchStr")
        Dim rProcessType As String = TIMS.ClearSQM(Request("ProcessType"))
        Dim flag_BACK1 As Boolean = False
        If rProcessType = "Back" Then flag_BACK1 = True
        If rProcessType = "Back2" Then flag_BACK1 = True
        If Session("QuestionarySearchStr") IsNot Nothing Then
            sQuestionarySearchStr = Session("QuestionarySearchStr")
            Session("QuestionarySearchStr") = Nothing
        End If
        If sQuestionarySearchStr <> "" AndAlso flag_BACK1 Then
            'Dim sQuestionarySearchStr As String = Session("QuestionarySearchStr")
            Dim MyValue As String = ""
            center.Text = TIMS.GetMyValue(sQuestionarySearchStr, "center")
            RIDValue.Value = TIMS.GetMyValue(sQuestionarySearchStr, "RIDValue")
            TMID1.Text = TIMS.GetMyValue(sQuestionarySearchStr, "TMID1")
            OCID1.Text = TIMS.GetMyValue(sQuestionarySearchStr, "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(sQuestionarySearchStr, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(sQuestionarySearchStr, "OCIDValue1")
            MyValue = TIMS.GetMyValue(sQuestionarySearchStr, "rblprtType1")
            Common.SetListItem(rblprtType1, MyValue)
            MyValue = TIMS.GetMyValue(sQuestionarySearchStr, "Button1")
            'rblprtType1
            If MyValue = "True" Then
                'hidSearchTag.Value = "search"
                'search_Click(sender, e)
                If OCIDValue1.Value <> "" Then
                    MyValue = OCIDValue1.Value
                    blnPrint2016 = False
                    Hid_rblprtType1.Value = rblprtType1.SelectedValue
                    Select Case Hid_rblprtType1.Value
                        Case Cst_defQA16
                            blnPrint2016 = True
                    End Select
                    If blnPrint2016 Then
                        Call ShowStudentsOfClass4(MyValue)
                        Exit Sub
                    End If
                    Call ShowStudentsOfClass(MyValue)
                End If
            End If
        End If
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
    End Sub

    '查詢SQL (20160501)
    Sub sSearch2()
        'VIEW_STUDSURVEY
        'V_STUDQUESTION4
        Me.Panel1.Visible = True
        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False
        'Dim PlanKind As Integer
        'PlanKind = DbAccess.ExecuteScalar("SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'")
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Exit Sub
        Dim cTPlanID As String = TIMS.GetTPlanID(drCC("PlanID"), objconn)
        'blnPrint2016 = False
        'If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(cTPlanID) > -1 Then
        '    If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(drCC("FTDATE"))) >= 0 Then blnPrint2016 = True
        'End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = drCC("RID") ' sm.UserInfo.RID
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,cc.CyclType" & vbCrLf
        sql &= " ,cc.LevelType" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2 " & vbCrLf
        'sql &= " ,cc.STDate " & vbCrLf
        sql &= " ,cc.FTDate " & vbCrLf
        sql &= " ,ISNULL(b.total,0) TOTAL " & vbCrLf
        sql &= " ,ISNULL(b.num1,0) NUM1 " & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN dbo.ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= "    SELECT cs.OCID" & vbCrLf
        sql &= "    ,COUNT(1) total " & vbCrLf
        sql &= "    ,COUNT(CASE WHEN Q4.SOCID IS NOT NULL THEN 1 END ) NUM1 " & vbCrLf
        sql &= "    FROM dbo.CLASS_STUDENTSOFCLASS cs " & vbCrLf
        sql &= "    JOIN dbo.CLASS_CLASSINFO cc ON cc.OCID = cs.OCID " & vbCrLf
        sql &= "    JOIN dbo.ID_Plan ip ON ip.PlanID = cc.PlanID " & vbCrLf
        sql &= "    LEFT JOIN dbo.V_STUDQUESTION4 q4 ON q4.SOCID = cs.SOCID " & vbCrLf
        sql &= "    WHERE 1=1 " & vbCrLf
        sql &= "    AND cs.StudStatus NOT IN (2,3) " & vbCrLf
        sql &= "    AND cc.OCID = '" & OCIDValue1.Value & "' "
        sql &= "    AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = '" & cTPlanID & "' " & vbCrLf
            sql &= " AND cc.Years = '" & drCC("Years") & "' " & vbCrLf
        Else
            sql &= " AND ip.PlanID = '" & drCC("PlanID") & "' " & vbCrLf
        End If
        sql &= "     GROUP BY cs.OCID " & vbCrLf
        sql &= " ) B ON b.ocid = cc.OCID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND cc.OCID = '" & OCIDValue1.Value & "' "
        sql &= " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = '" & cTPlanID & "' " & vbCrLf
            sql &= " AND cc.Years = '" & drCC("Years") & "' " & vbCrLf
        Else
            sql &= " AND ip.PlanID = '" & drCC("PlanID") & "' " & vbCrLf
        End If
        Dim dt_Class As DataTable
        dt_Class = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!"
        DataGrid1.Visible = False
        DataGrid1.Style.Item("display") = "none"
        If dt_Class.Rows.Count > 0 Then
            msg.Text = "" 'False
            DataGrid1.Visible = True
            DataGrid1.Style.Item("display") = ""
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataSource = dt_Class
            DataGrid1.DataBind()
            ''分頁用-   Start
            'DataGridPage1.MyDataTable = stud_table
            'DataGridPage1.FirstTime()
            ''分頁用-   End
        End If
    End Sub

    '查詢SQL (OLD)
    Sub sSearch1()
        Me.Panel1.Visible = True
        StudentTable.Style.Item("display") = "none"
        msg2.Visible = False
        Label1.Visible = False
        'Dim PlanKind As Integer
        'PlanKind = DbAccess.ExecuteScalar("SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'")
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Exit Sub
        Dim cTPlanID As String = TIMS.GetTPlanID(drCC("PlanID"), objconn)
        'blnPrint2016 = False
        'If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(cTPlanID) > -1 Then
        '    If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(drCC("FTDATE"))) >= 0 Then blnPrint2016 = True
        'End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = drCC("RID") ' sm.UserInfo.RID
        Dim sqlstr_class As String = ""
        sqlstr_class = ""
        sqlstr_class &= " SELECT a.OCID"
        sqlstr_class &= " ,a.CyclType "
        sqlstr_class &= " ,a.LevelType "
        sqlstr_class &= " ,a.ClassCName "
        sqlstr_class &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME2 " & vbCrLf
        sqlstr_class &= " ,a.FTDate "
        sqlstr_class &= " ,b.total "
        sqlstr_class &= " ,ISNULL(c.num1,0) num1 " & vbCrLf
        sqlstr_class &= " FROM CLASS_CLASSINFO a " & vbCrLf
        sqlstr_class &= " JOIN ID_Plan ip ON ip.PlanID = a.PlanID " & vbCrLf
        sqlstr_class &= " JOIN (" & vbCrLf
        sqlstr_class &= "   SELECT cs.OCID ,COUNT(1) total " & vbCrLf
        sqlstr_class &= "   FROM CLASS_STUDENTSOFCLASS cs " & vbCrLf
        sqlstr_class &= "   JOIN CLASS_CLASSINFO cc on cc.OCID = cs.OCID " & vbCrLf
        sqlstr_class &= "   JOIN ID_Plan ip on ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr_class &= "   WHERE 1=1 " & vbCrLf
        sqlstr_class &= "   AND cs.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        sqlstr_class &= "   AND cs.StudStatus NOT IN (2,3) " & vbCrLf
        sqlstr_class &= "   AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sqlstr_class &= " AND ip.TPlanID = '" & cTPlanID & "' " & vbCrLf
            sqlstr_class &= " AND cc.Years = '" & drCC("Years") & "' " & vbCrLf
        Else
            sqlstr_class &= " AND ip.PlanID = '" & drCC("PlanID") & "' " & vbCrLf
        End If
        sqlstr_class &= "   GROUP BY cs.OCID " & vbCrLf
        sqlstr_class &= " ) b ON a.ocid = b.ocid " & vbCrLf
        sqlstr_class &= " LEFT JOIN ( " & vbCrLf
        sqlstr_class &= "   SELECT cs.OCID ,COUNT(1) num1 " & vbCrLf
        sqlstr_class &= "   FROM STUD_QUESTIONARY cs " & vbCrLf
        sqlstr_class &= "   JOIN CLASS_CLASSINFO cc ON cc.OCID = cs.OCID " & vbCrLf
        sqlstr_class &= "   JOIN ID_Plan ip on ip.PlanID = cc.PlanID " & vbCrLf
        sqlstr_class &= "   WHERE 1=1 " & vbCrLf
        sqlstr_class &= "   AND cs.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        sqlstr_class &= "   AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sqlstr_class &= " AND ip.TPlanID = '" & cTPlanID & "' " & vbCrLf
            sqlstr_class &= " AND cc.Years = '" & drCC("Years") & "' " & vbCrLf
        Else
            sqlstr_class &= " AND ip.PlanID = '" & drCC("PlanID") & "' " & vbCrLf
            'If PlanKind = 1 Then sqlstr_class &= " AND OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE Account = '" & sm.UserInfo.UserID & "') "
        End If
        sqlstr_class &= "   GROUP BY cs.OCID " & vbCrLf
        sqlstr_class &= " ) c ON c.ocid = a.ocid " & vbCrLf
        sqlstr_class &= " WHERE 1=1 " & vbCrLf
        sqlstr_class &= " AND a.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sqlstr_class &= " AND ip.TPlanID = '" & cTPlanID & "' " & vbCrLf
            sqlstr_class &= " AND a.Years = '" & drCC("Years") & "' " & vbCrLf
        Else
            sqlstr_class &= " AND ip.PlanID = '" & drCC("PlanID") & "' " & vbCrLf
            'If PlanKind = 1 Then sqlstr_class &= " and OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE Account='" & sm.UserInfo.UserID & "') "
        End If
        sqlstr_class &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
        sqlstr_class &= " ORDER BY a.OCID " & vbCrLf
        Dim dt_Class As DataTable
        dt_Class = DbAccess.GetDataTable(sqlstr_class, objconn)

        msg.Text = "查無資料!!"
        DataGrid1.Visible = False
        DataGrid1.Style.Item("display") = "none"
        If dt_Class.Rows.Count > 0 Then
            msg.Text = "" 'False
            DataGrid1.Visible = True
            DataGrid1.Style.Item("display") = ""
            DataGrid1.DataKeyField = "OCID"
            DataGrid1.DataSource = dt_Class
            DataGrid1.DataBind()
            ''分頁用-   Start
            'DataGridPage1.MyDataTable = stud_table
            'DataGridPage1.FirstTime()
            ''分頁用-   End
        End If
    End Sub

    '檢核並設定 blnPrint2016 false:舊true:201605月版本
    Public Shared Function uGet_blnPrint2016(ByVal cTPlanID As String, ByVal cFTDATE As String) As Boolean
        Dim rst As Boolean = False
#Region "(No Use)"

        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then Exit Sub
        'Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        'If drCC Is Nothing Then Exit Sub
        'Dim cTPlanID As String = TIMS.GetTPlanID(drCC("PlanID"), objconn)
        'blnPrint2016 = False

#End Region
        If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(cTPlanID) > -1 Then
            If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(cFTDATE)) >= 0 Then rst = True
        End If
        Return rst
    End Function

    '查詢SQL
    Sub gSearch1()
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇職類班別!!")
            Exit Sub
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Exit Sub
        Dim cTPlanID As String = TIMS.GetTPlanID(drCC("PlanID"), objconn)
        '檢核並設定 blnPrint2016
        blnPrint2016 = uGet_blnPrint2016(cTPlanID, TIMS.Cdate3(drCC("FTDATE")))
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇職類班別!!")
            Exit Sub
        End If
        Common.SetListItem(rblprtType1, Cst_defQA2)
        If blnPrint2016 Then
            Common.SetListItem(rblprtType1, Cst_defQA16)
        End If
        Hid_rblprtType1.Value = rblprtType1.SelectedValue
#Region "(No Use)"

        'blnPrint2016 = False
        'Hid_rblprtType1.Value = rblprtType1.SelectedValue
        'Select Case Hid_rblprtType1.Value
        '    Case Cst_defQA16
        '        blnPrint2016 = True
        'End Select

#End Region
        If blnPrint2016 Then
            Call sSearch2()
            Exit Sub
        End If
        Call sSearch1()
    End Sub

    '查詢鈕
    Private Sub search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles search.Click
        Call gSearch1()
    End Sub

    Sub GetSearchStr()
        Session("QuestionarySearchStr") = Nothing
        Dim schStr1 As String = ""
        TIMS.SetMyValue(schStr1, "center", center.Text)
        TIMS.SetMyValue(schStr1, "RIDValue", RIDValue.Value)
        TIMS.SetMyValue(schStr1, "TMID1", TMID1.Text)
        TIMS.SetMyValue(schStr1, "OCID1", OCID1.Text)
        TIMS.SetMyValue(schStr1, "TMIDValue1", TMIDValue1.Value)
        TIMS.SetMyValue(schStr1, "OCIDValue1", OCIDValue1.Value)
        TIMS.SetMyValue(schStr1, "Button1", CStr(DG_stud.Visible))
        TIMS.SetMyValue(schStr1, "StudentTable", StudentTable.Style.Item("display"))
        TIMS.SetMyValue(schStr1, "rblprtType1", rblprtType1.SelectedValue)
        Session("QuestionarySearchStr") = schStr1
    End Sub

    '學員明細查詢 (OLD)
    Sub ShowStudentsOfClass(ByVal vOCID As String)
        If vOCID = "" Then Exit Sub

        '班級查詢
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT '班別：' +dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME ,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.LevelType " & vbCrLf
        sql &= " ,ISNULL(b.StudentCount,0) StudentCount " & vbCrLf
        sql &= " ,ISNULL(b.TrainCount,0) TrainCount " & vbCrLf
        sql &= " ,ISNULL(b.LeaveCount,0) LeaveCount " & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc " & vbCrLf
        sql &= " JOIN (" & vbCrLf
        sql &= "    SELECT cs.ocid " & vbCrLf
        sql &= "    ,count(1) StudentCount" & vbCrLf
        sql &= "    ,SUM(CASE WHEN cs.StudStatus NOT IN (2,3) THEN 1 END) TrainCount" & vbCrLf
        sql &= "    ,SUM(CASE WHEN cs.StudStatus IN (2,3) THEN 1 END) LeaveCount " & vbCrLf
        sql &= "    FROM CLASS_STUDENTSOFCLASS cs " & vbCrLf
        sql &= "    WHERE 1=1 " & vbCrLf
        sql &= "    AND cs.MakeSOCID IS NULL " & vbCrLf
        sql &= "    AND cs.OCID = @OCID " & vbCrLf
        sql &= "    GROUP BY cs.ocid " & vbCrLf
        sql &= " ) b ON b.OCID = cc.OCID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND cc.OCID = @OCID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = vOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Label1.Text = Convert.ToString(dr("ClassCName"))
            If Convert.ToString(dr("LevelType")) <> "" Then Label1.Text &= "第" & CStr(dr("LevelType")) & "階段"
            Label1.Text &= "("
            Label1.Text &= "開訓人數:" & Convert.ToString(dr("StudentCount"))
            Label1.Text &= "&nbsp;&nbsp;在結訓人數:" & Convert.ToString(dr("TrainCount"))
            Label1.Text &= "&nbsp;&nbsp;離退訓人數:" & Convert.ToString(dr("LeaveCount"))
            Label1.Text &= ")"
        End If

        '班級學員查詢
        Dim sqlstr_stud As String = ""
        sqlstr_stud = ""
        sqlstr_stud += " SELECT a.IsClosed ,b.studentid"
        sqlstr_stud += " ,b.StudStatus ,c.name ,b.OCID ,b.SOCID ,b.RejectTDate1 ,b.RejectTDate2 "
        sqlstr_stud += " FROM class_classinfo a "
        sqlstr_stud += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
        sqlstr_stud += " JOIN stud_studentinfo c ON b.sid = c.sid "
        sqlstr_stud += " WHERE a.ocid = '" & vOCID & "' "
        sqlstr_stud += " AND b.MakeSOCID IS NULL " & vbCrLf
        sqlstr_stud += " AND b.StudStatus NOT IN (2,3) " '排除離退
        sqlstr_stud += " ORDER BY b.studentid "
        dt = DbAccess.GetDataTable(sqlstr_stud, objconn)
        Session("DTable_Stuednt") = Nothing
        If dt.Rows.Count > 0 Then Session("DTable_Stuednt") = dt

        '班級學員查詢
        sqlstr_stud = ""
        sqlstr_stud += " SELECT a.IsClosed ,b.studentid "
        sqlstr_stud += " ,dbo.FN_CSTUDID2(b.STUDENTID) StudentID2"
        'sqlstr_stud += " ,SUBSTRING(b.studentid, LEN(b.studentid)-2, 2) StudentID2 "
        sqlstr_stud += " ,b.StudStatus ,c.name ,b.OCID ,b.SOCID ,b.RejectTDate1 ,b.RejectTDate2 "
        sqlstr_stud += " FROM class_classinfo a "
        sqlstr_stud += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
        sqlstr_stud += " JOIN stud_studentinfo c ON b.sid = c.sid "
        sqlstr_stud += " WHERE a.ocid = '" & vOCID & "' "
        sqlstr_stud += " AND b.MakeSOCID IS NULL " & vbCrLf
        'sqlstr_stud += "   AND b.StudStatus NOT IN (2,3) " '排除離退
        'sqlstr_stud += " ORDER BY SUBSTRING(b.studentid, LEN(b.studentid)-2, 2) "
        sqlstr_stud += " ORDER BY dbo.FN_CSTUDID2(b.STUDENTID) "
        dt = DbAccess.GetDataTable(sqlstr_stud, objconn)

        Session("dtQUE") = Nothing
        sql = " SELECT * FROM STUD_QUESTIONARY WHERE OCID = '" & vOCID & "' "
        Dim dtQUE As DataTable = DbAccess.GetDataTable(sql, objconn)
        Session("dtQUE") = dtQUE
        msg2.Text = "查無此班學生資料!"
        StudentTable.Style.Item("display") = "none"
        Label1.Visible = False

        If dt.Rows.Count > 0 Then
            msg2.Text = ""
            StudentTable.Style.Item("display") = ""
            Label1.Visible = True
            DG_stud.DataSource = dt
            'DG_stud.DataKeyField = "SOCID"
            DG_stud.DataBind()
        End If
    End Sub

    '學員明細查詢 (20160501)
    Sub ShowStudentsOfClass4(ByVal vOCID As String)
        If vOCID = "" Then Exit Sub

        '班級查詢
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT '班別：' +dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME ,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.LevelType " & vbCrLf
        sql &= " ,ISNULL(b.StudentCount,0) StudentCount" & vbCrLf
        sql &= " ,ISNULL(b.TrainCount,0) TrainCount " & vbCrLf
        sql &= " ,ISNULL(b.LeaveCount,0) LeaveCount " & vbCrLf
        sql &= " FROM Class_ClassInfo cc " & vbCrLf
        sql &= " JOIN (SELECT cs.ocid ,COUNT(1) StudentCount " & vbCrLf
        sql &= "    ,SUM(CASE WHEN cs.StudStatus NOT IN (2,3) THEN 1 END) TrainCount " & vbCrLf
        sql &= "    ,SUM(CASE WHEN cs.StudStatus IN (2,3) THEN 1 END) LeaveCount " & vbCrLf
        sql &= "    FROM Class_StudentsOfClass cs " & vbCrLf
        sql &= "    WHERE 1=1 " & vbCrLf
        sql &= "    AND cs.MakeSOCID IS NULL " & vbCrLf
        sql &= "    AND cs.OCID = @OCID " & vbCrLf
        sql &= "    GROUP BY cs.ocid " & vbCrLf
        sql &= " ) b ON b.OCID = cc.OCID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND cc.OCID = @OCID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = vOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Label1.Text = Convert.ToString(dr("ClassCName"))
            If Convert.ToString(dr("LevelType")) <> "" Then Label1.Text &= "第" & CStr(dr("LevelType")) & "階段"
            Label1.Text &= "("
            Label1.Text &= "開訓人數:" & Convert.ToString(dr("StudentCount"))
            Label1.Text &= "&nbsp;&nbsp;在結訓人數:" & Convert.ToString(dr("TrainCount"))
            Label1.Text &= "&nbsp;&nbsp;離退訓人數:" & Convert.ToString(dr("LeaveCount"))
            Label1.Text &= ")"
        End If

        '班級學員查詢
        Dim sqlstr_stud As String = ""
        sqlstr_stud = ""
        sqlstr_stud += " SELECT a.IsClosed ,b.studentid ,b.StudStatus ,c.name ,b.OCID ,b.SOCID ,b.RejectTDate1 ,b.RejectTDate2 "
        sqlstr_stud += " FROM class_classinfo a "
        sqlstr_stud += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
        sqlstr_stud += " JOIN stud_studentinfo c on b.sid = c.sid "
        sqlstr_stud += " WHERE a.ocid = '" & vOCID & "' "
        sqlstr_stud += " AND b.MakeSOCID IS NULL " & vbCrLf
        sqlstr_stud += " AND b.StudStatus NOT IN (2,3) " '排除離退
        sqlstr_stud += " ORDER BY b.studentid "
        dt = DbAccess.GetDataTable(sqlstr_stud, objconn)
        Session("DTable_Stuednt") = Nothing
        If dt.Rows.Count > 0 Then Session("DTable_Stuednt") = dt

        '班級學員查詢
        sqlstr_stud = ""
        sqlstr_stud += " SELECT a.IsClosed ,b.studentid "
        sqlstr_stud += " ,dbo.FN_CSTUDID2(b.STUDENTID) StudentID2"
        'sqlstr_stud += " ,SUBSTRING(b.studentid,LEN(b.studentid)-2, 2) StudentID2"
        sqlstr_stud += " ,b.StudStatus " & vbCrLf
        sqlstr_stud += " ,c.name ,b.OCID ,b.SOCID ,b.RejectTDate1 ,b.RejectTDate2 "
        sqlstr_stud += " FROM class_classinfo a "
        sqlstr_stud += " JOIN class_studentsofclass b ON a.ocid = b.ocid "
        sqlstr_stud += " JOIN stud_studentinfo c on b.sid = c.sid "
        sqlstr_stud += " WHERE a.ocid='" & vOCID & "' "
        sqlstr_stud += " AND b.MakeSOCID IS NULL " & vbCrLf
        'sqlstr_stud += "   AND b.StudStatus NOT IN (2,3) "'排除離退
        sqlstr_stud += " ORDER BY dbo.SUBSTR(b.studentid,-2) "
        dt = DbAccess.GetDataTable(sqlstr_stud, objconn)

        'SELECT * FROM STUD_QUESTIONARY  WHERE ROWNUM <=10
        'SELECT * FROM V_STUDQUESTION4 
        Session("dtQUE") = Nothing
        sql = " SELECT * FROM V_STUDQUESTION4 WHERE OCID = '" & vOCID & "' "
        Dim dtQUE As DataTable = DbAccess.GetDataTable(sql, objconn)
        Session("dtQUE") = dtQUE
        msg2.Text = "查無此班學生資料!"
        StudentTable.Style.Item("display") = "none"
        Label1.Visible = False

        If dt.Rows.Count > 0 Then
            msg2.Text = ""
            StudentTable.Style.Item("display") = ""
            Label1.Visible = True
            DG_stud.DataSource = dt
            'DG_stud.DataKeyField = "SOCID"
            DG_stud.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        Select Case e.CommandName
            Case "view"
                Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
                blnPrint2016 = False
                Select Case Hid_rblprtType1.Value
                    Case Cst_defQA16
                        blnPrint2016 = True
                End Select
                If blnPrint2016 Then
                    DataGrid1.Visible = False
                    Call ShowStudentsOfClass4(vOCID)
                    Exit Sub
                End If
                DataGrid1.Visible = False
                Call ShowStudentsOfClass(vOCID)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 'e.Item.ItemIndex + 1

                Dim drv As DataRowView = e.Item.DataItem
                'Dim OCID_Namestr As String
                'OCID_Namestr = drv("ClassCName").ToString
                'If CInt(e.Item.Cells(6).Text) <> 0 Then OCID_Namestr += "第" & TIMS.GetChtNum(CInt(e.Item.Cells(6).Text)) & "期"
                'e.Item.Cells(1).Text = OCID_Namestr
                Dim mybut1 As Button = e.Item.FindControl("btnView1")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                mybut1.CommandArgument = sCmdArg ' DataGrid1.DataKeys(e.Item.ItemIndex)
                'Case ListItemType.Header, ListItemType.Footer
                'Case Else
        End Select
    End Sub

    Function GetNewSubNo(ByVal Dlid As String, ByVal SOCID As String, ByVal tmpTrans As SqlTransaction) As String
        Dim Rst As String = "1"
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim xFlag As Boolean = False '取得有效值為True
        If Not xFlag Then
            sql = " SELECT SubNo FROM STUD_RESULTSTUDDATA WHERE SOCID = '" & SOCID & "' "
            dr = DbAccess.GetOneRow(sql, tmpTrans)
            If dr IsNot Nothing Then
                xFlag = True
                Rst = Convert.ToString(dr("SubNo"))
            End If
        End If
        If Not xFlag Then
            sql = " SELECT ISNULL(MAX(SubNo),0)+1 NewSubNo FROM STUD_RESULTSTUDDATA WHERE Dlid = '" & Dlid & "' "
            dr = DbAccess.GetOneRow(sql, tmpTrans)
            If dr IsNot Nothing Then
                xFlag = True
                Rst = Convert.ToString(dr("NewSubNo"))
            End If
        End If
        Return Rst
    End Function

    Function ChangeYN(ByVal Str_12 As String) As String
        Dim rst As String = ""
        If Str_12 <> "" Then
            If Left(Str_12, 1) = "1" Then
                rst = "Y"
            Else
                rst = "N"
            End If
        End If
        Return rst
    End Function

    Function ChangeTWDate(ByVal TWDate As String) As String
        Return CStr(CInt(Left(TWDate, 3)) + 1911) & "/" & Mid(TWDate, 4, 2) & "/" & Right(TWDate, 2)
    End Function

    Function getDLIDforOCID(ByVal OCID As String) As String
        Dim Rst As String = ""
        Dim objstr As String
        Dim dt As DataTable
        objstr = " SELECT DLID FROM Stud_DataLid WHERE OCID = '" & OCID & "' "
        dt = DbAccess.GetDataTable(objstr, objconn)
        If dt.Rows.Count > 0 Then Rst = dt.Rows(0)("DLID")
        Return Rst
    End Function

    'Class_StudentsOfClass'檢查學生是否存在
    Function Check_Class_StudentOfClass(ByVal StudID As String, ByVal OCID As String, ByRef StudDLID As String, ByRef SOCID As String, ByRef StudentID12 As String, ByRef StdName As String, ByRef StdPID As String, ByRef Sex As String, ByRef BirthYear As String, ByRef BirthMonth As String, ByRef BirthDate As String) As String
        Dim Reason As String = "" '產生錯誤回應
        Reason = "依據學號(StudID) 及開班編號(OCID) 無法得知班級學員資料(STUD_STUDENTINFO) <BR>"
        '清空下列參數免得重複輸入
        SOCID = ""
        StudentID12 = ""
        StdName = ""
        StdPID = ""
        Sex = ""
        BirthYear = ""
        BirthMonth = ""
        BirthDate = ""

        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        'Dim da As SqlDataAdapter
        'Dim i As Integer
        If IsNumeric(StudID) AndAlso IsNumeric(OCID) Then
            sql = " "
            sql &= " SELECT ISNULL(C.DLID,0) DLID ,ISNULL(C.SUBNO,0) SUBNO ,A.SOCID ,b.name StdName "
            sql &= " ,dbo.FN_CSTUDID2(A.STUDENTID) StudentID "
            'sql &= " ,CASE WHEN LEN(A.StudentID)=12 THEN SUBSTRING(A.StudentID,LEN(A.StudentID)-3,3) ELSE SUBSTRING(A.StudentID,LEN(A.StudentID)-2,2) END AS StudentID "
            sql &= " ,A.StudentID StudentID12 "
            sql &= " ,B.IDNO StdPID "
            sql &= " ,(CASE WHEN B.Sex = 'M' THEN 1 ELSE 2 END) Sex "
            sql &= " ,DATEPART(YEAR, B.Birthday) BirthYear"
            sql &= " ,DATEPART(MONTH, B.Birthday) Birthmonth "
            sql &= " ,DATEPART(DAY, B.Birthday) Birthdate "
            sql &= " FROM CLASS_STUDENTSOFCLASS A "
            sql &= " JOIN STUD_STUDENTINFO B ON A.SID = B.SID "
            sql &= " LEFT JOIN STUD_RESULTSTUDDATA C ON A.SOCID = C.SOCID "
            sql &= " WHERE A.OCID = '" & OCID & "' "
            sql &= " AND dbo.FN_CSTUDID2(A.STUDENTID)='" & StudID & "'"
            'sql &= " AND dbo.FN_CSTUDID2(A.STUDENTID)=CONVERT(NUMERIC, '" & StudID & "') = CASE WHEN LEN(A.StudentID)=12 THEN CONVERT(NUMERIC, SUBSTRING(A.StudentID,LEN(A.StudentID)-3,3)) ELSE CONVERT(NUMERIC, SUBSTRING(A.StudentID,LEN(A.StudentID)-2,2)) END "
            Try
                dt = DbAccess.GetDataTable(sql, objconn)
                'trans = DbAccess.BeginTrans(objconn)
                'DbAccess.CommitTrans(trans)
                If dt.Rows.Count <> 0 Then
                    dr = dt.Rows(0) 'DbAccess.GetOneRow(sql)
                    If Convert.ToString(dr("DLID")) <> "0" AndAlso Convert.ToString(dr("DLID")) <> "" Then
                        StudDLID = Convert.ToString(dr("DLID"))
                        '若是有DLID 表示此班的同學曾經有建立過 Stud_ResultStudData，改變ByRef DLID
                    End If
                    '沒有，則維持原案
                    'If dr("SubNo") <> 0 Then SubNo = dr("SubNo") '若是有SubNo 表示此班的同學曾經有建立過 Stud_ResultStudData，改變ByRef SubNo
                    SOCID = dr("SOCID")
                    StudentID12 = dr("StudentID12")
                    StdName = dr("StdName")
                    StdPID = dr("StdPID")
                    Sex = dr("Sex")
                    BirthYear = dr("BirthYear")
                    BirthMonth = dr("BirthMonth")
                    BirthDate = dr("BirthDate")
                    Reason = "" '正常結束(清空錯誤回應)
                End If
            Catch ex As Exception
                'DbAccess.RollbackTrans(trans)
                'Reason = "依據學號(StudID) 及開班編號(OCID) 無法得知班級學員資料(STUD_STUDENTINFO) <BR>"
                'Throw ex
            End Try
        End If
        Return Reason
    End Function

    Function CheckcolArray1(ByRef StrVal1 As String, ByRef ItemObj As Object) As Boolean
        Dim Rst As Boolean = True
        If StrVal1 <> "" Then StrVal1 = Trim(StrVal1)
        '空白或非數字
        If StrVal1 = "" OrElse Not IsNumeric(Left(StrVal1, 1)) Then
            Rst = False
        Else
            ItemObj = Left(StrVal1, 1)
        End If
        Return Rst
    End Function

    'CHK STUD_RESULTSTUDDATA
    Function CheckImportData(ByVal colArray As Array, ByRef writeflag As Boolean, ByVal int_cst_Len2 As Integer) As String
        'amu 20061221 因為同意可寫入某些錯誤的資料，但還是要show訊息
        writeflag = False '一開始都拒絕寫入 

        Dim Reason As String = ""
        'Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Const cst_Len As Integer = 93 '讀卡機可能產生這樣的長度。
        'Const cst_Len As Integer = 94 '讀卡機可能產生這樣的長度。
        'Const cst_Len2a As Integer = 43
        'Const cst_Len2a2 As Integer = 42 '整理過後的長度
        'Const cst_Len2a2 As Integer = 43 '整理過後的長度
        'Const cst_Len2b As Integer = 34

        'Dim int_cst_Len2 As Integer = 0
        'Select Case vsQName
        '    Case "B" '在職B卷
        '        int_cst_Len2 = cst_Len2b
        '        'Case "A" '職前A卷
        '        '    int_cst_Len2 = cst_Len2a
        '    Case "A2" '職前A2卷
        '        int_cst_Len2 = cst_Len2a2
        'End Select

        '間隔 驗証 OK@True NG@False  預設為@NG '讀卡機可能長度判斷
        Dim stepFlagChk1 As Boolean = False
        If colArray.Length >= cst_Len Then
            stepFlagChk1 = True
        End If

        '間隔 驗証 OK@True NG@False  預設為@NG '整理過後的長度長度判斷
        Dim stepFlagChk2 As Boolean = False
        If colArray.Length >= int_cst_Len2 Then
            stepFlagChk2 = True
        End If

        '若2種長度判斷都不OK
        If Not stepFlagChk1 AndAlso Not stepFlagChk2 Then
            'Reason += "欄位數量不正確(應該為" & cst_Len2 & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            'Reason &= "程式判斷欄位數：" & colArray.Length & "<BR>"
            'Reason &= "系統限定欄位數(" & cst_Len & "或" & int_cst_Len2 & ")<BR>"
            Return Reason '離開。
        End If

        FillFormDate = colArray(cst_FillFormDate).ToString '讀卡日期 
        StudID = Convert.ToString(colArray(cst_StudID)) '學號 
        DegreeID = colArray(cst_DegreeID).ToString '學歷 
        MilitaryID = colArray(cst_MilitaryID).ToString '兵役 
        sIdentityID = colArray(cst_IdentityID).ToString '身分 
        Q7 = colArray(cst_Q7).ToString '動機 'Stud_ResultStudData Q7	參訓職訓動機	int
        Q8 = colArray(cst_Q8).ToString '動向 'Stud_ResultStudData Q8	結訓後動向	int
        Q9 = colArray(cst_Q9).ToString '有無工作(有/無) 'Stud_ResultStudData Q9	受訓前有無在工作 char(1)
        Q9Y = colArray(cst_Q9Y).ToString '有在工作 'Stud_ResultStudData Q9Y	受訓前有無在工作 int
        Q10 = colArray(cst_Q10).ToString '找工作(有/無) 'Stud_ResultStudData  Q10	受訓前有無找工作 char(1)
        'Q11 = colArray(cst_Q11).ToString '找到工作(有/無) 'Stud_ResultStudData Q11 參加訓練後	char(1)
        Q11N = colArray(cst_Q11N).ToString '工作幫助 'Stud_ResultStudData Q11N	參加訓練後	int
        gQ12v1 = colArray(cst_Q12v1).ToString '1.參訓職類符合就業市場需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v2 = colArray(cst_Q12v2).ToString '2.教學課程安排 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v3 = colArray(cst_Q12v3).ToString '3.訓練師專業及熱忱 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v4 = colArray(cst_Q12v4).ToString '4.訓練設備符合產業需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v5 = colArray(cst_Q12v5).ToString '5.訓練時數 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意

        '讀卡日期
        If FillFormDate = "" Or Len(FillFormDate) <> 7 Or IsNumeric(FillFormDate) <> True Then
            Reason += "讀卡日期有誤，必須是民國年格式(yyymmdd)<BR>"
        Else
            'FillFormDate = CStr(CInt(Left(FillFormDate, 3)) + 1911) & "/" & Mid(FillFormDate, 4, 2) & "/" & Right(FillFormDate, 2)
            FillFormDate = ChangeTWDate(FillFormDate)
            If IsDate(FillFormDate) = False Then
                Reason += "讀卡日期有誤，必須是民國年格式(yyymmdd)<BR>"
            Else
                If CDate(FillFormDate) < "1900/1/1" Or CDate(FillFormDate) > "2100/1/1" Then Reason += "讀卡日期範圍有誤<BR>"
            End If
        End If

        '學號
        If StudID = "" Then
            Reason += "學員序號有誤，必須有值<BR>"
        Else
            Try
                StudID = TIMS.ChangeIDNO(StudID)
                StudID = CStr(CInt(StudID))
            Catch ex As Exception
                Reason += "學員序號有誤，必須為數字格式<BR>"
            End Try
        End If
        'If StudID = "" Or Len(StudID) <> 3 Or IsNumeric(StudID) <> True Then
        '    Reason += "學號有誤，必須是格式(XXX)<BR>"
        'Else
        '    If CStr(StudID) < "000" Or CStr(StudID) > "999" Then Reason += "學號範圍有誤<BR>"
        'End If

        DegreeID = Left(DegreeID, 1) '學歷
        If DegreeID = "" Then
            Reason += "必須填寫最高學歷<BR>"
        Else
            Dim MyKey As String = DegreeID
            If DegreeID.Length < 2 Then MyKey = "0" & DegreeID
            If dtDegree.Select("DegreeID='" & MyKey & "'").Length = 0 Then Reason += "學歷值有錯，不符合鍵詞<BR>"
        End If

        MilitaryID = Left(MilitaryID, 1) '兵役
        If MilitaryID = "" Then
            Reason += "必須填寫兵役狀況<BR>"
        Else
            Dim MyKey As String = MilitaryID
            If MilitaryID.Length < 2 Then MyKey = "0" & MilitaryID
            If dtMilitary.Select("MilitaryID='" & MyKey & "'").Length = 0 Then Reason += "兵役狀況有錯，不符合鍵詞<BR>"
        End If

        'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
        'BY AMU 2009-07-30
        '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
        If sIdentityID = "" Then
            'Reason += "必須填寫結訓身分別<BR>"
        Else
            If Len(sIdentityID) Mod 2 = 1 Then sIdentityID = "0" & sIdentityID
            If Len(sIdentityID) Mod 2 <> 0 Then
                Reason += "結訓身分別不符合鍵詞<BR>" '已整理
            Else
                For i As Integer = 1 To (Len(sIdentityID) / 2)
                    Dim MyKey As String = Mid(sIdentityID.ToString, (i * 2 - 1), 2)
                    'If IdentityID.Length < 2 Then MyKey = "0" & IdentityID Else MyKey = IdentityID
                    If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                        Reason += "結訓身分別不符合鍵詞<BR>"
                        Exit For
                    End If
                Next
            End If
        End If

        Q7 = Left(Q7, 1)
        If Q7 = "" Then
            Reason += "必須填寫參訓職訓動機<BR>"
        Else
            Try
                Q7 = CInt(Q7)
                If CInt(Q7) < 1 OrElse CInt(Q7) > 6 Then
                    Reason += "參訓職訓動機 範圍有誤<BR>"
                    writeflag = False
                End If
            Catch ex As Exception
                Reason += "參訓職訓動機 範圍有誤<BR>"
                writeflag = False
            End Try
        End If

        Q8 = Left(Q8, 1)
        If Q8 = "" Then
            Reason += "必須填寫結訓後動向<BR>"
        Else
            Try
                Q8 = CInt(Q8)
                If CInt(Q8) < 1 OrElse CInt(Q8) > 5 Then
                    Reason += "結訓後動向 範圍有誤<BR>"
                    writeflag = False
                End If
            Catch ex As Exception
                Reason += "結訓後動向 範圍有誤<BR>"
                writeflag = False
            End Try
        End If

        '假如上列都有填寫，則可存取，否則不可存取
        If Reason = "" Then writeflag = True Else writeflag = False

        Q9 = Left(Q9, 1)
        Q9Y = Left(Q9Y, 1)
        If Q9 = "" Then
            Reason += "必須填寫訓練前有無工作<BR>"
            'writeflag = True
        Else
            'If CInt(Q9) < 1 Or CInt(Q9) > 2 Then Reason += "訓練前有無工作 範圍有誤<BR>"
            Select Case Q9
                Case "1"
                    If Q9Y = "" Then
                        Reason += "訓練前有工作:必須填寫訓練前工作<BR>"
                        'writeflag = True
                    End If
                Case "2"
                    If Q9Y <> "" Then
                        Reason += "訓練前無工作:不必填寫訓練前工作<BR>"
                        'writeflag = True
                    End If
                Case Else
                    Reason += "訓練前有無工作 範圍有誤<BR>"
                    writeflag = False
            End Select
        End If

        If writeflag Then
            If Q9Y <> "" Then
                Try
                    Q9Y = CInt(Q9Y)
                    If CInt(Q9Y) < 1 OrElse CInt(Q9Y) > 3 Then
                        Reason += "訓練前有無工作 範圍有誤<BR>"
                        writeflag = False
                    End If
                Catch ex As Exception
                    Reason += "訓練前有無工作 範圍有誤<BR>"
                    writeflag = False
                End Try
            End If
        End If

        Q10 = Left(Q10, 1)
        If Q10 = "" Then
            Reason += "必須填寫訓練前有否尋找工作<BR>"
            writeflag = False
        Else
            If writeflag Then
                Try
                    Q10 = CInt(Q10)
                    If CInt(Q10) < 1 Or CInt(Q10) > 2 Then
                        Reason += "訓練前有否尋找工作 範圍有誤<BR>"
                        writeflag = False
                    End If
                Catch ex As Exception
                    Reason += "訓練前有否尋找工作 範圍有誤<BR>"
                    writeflag = False
                End Try
            End If
        End If

#Region "(No Use)"

        'Q11 = Left(Q11, 1)
        'If Q11 = "" Then
        '    Reason += "必須填寫參加本訓後有否找到工作<BR>"
        '    'writeflag = True
        'Else
        '    Select Case Q11
        '        Case "1"
        '            If Q11N <> "" Then
        '                Reason += "參加本訓後找到工作:不必填寫日後工作幫助程度<BR>"
        '                'writeflag = True
        '            End If
        '        Case "2"
        '        Case Else
        '            Reason += "參加本訓後有否找到工作 範圍有誤<BR>"
        '            writeflag = False
        '    End Select
        'End If
        'If writeflag Then
        '    If Q11 = "2" Then
        '        Q11N = Left(Q11N, 1)
        '        If Q11N = "" Then
        '            Reason += "參加本訓後找不到工作:必須填寫日後工作幫助程度<BR>"
        '            writeflag = False
        '        Else
        '            If CInt(Q11N) < 1 Or CInt(Q11N) > 5 Then
        '                Reason += "日後工作幫助程度 範圍有誤<BR>"
        '                writeflag = False
        '            End If
        '        End If
        '    End If
        'End If

#End Region

        If Q11N <> "" Then Q11N = Left(Q11N, 1) '取得最後1碼
        If Q11N <> "" Then Q11N = Val(Q11N) '取得最後1碼轉換數字
        If writeflag Then
            If Q11N = "" Then
                Reason += "參加本訓後找不到工作:必須填寫 您覺得參加本次訓練後，對日後尋找工作幫助的程度<BR>"
                writeflag = False
            Else
                '應該是1~5
                If CInt(Q11N) < 1 Or CInt(Q11N) > 5 Then
                    Reason += "對日後尋找工作幫助的程度 範圍有誤<BR>"
                    writeflag = False
                End If
            End If
        End If
        Select Case gQ12v1
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v1.參訓職類符合就業市場需求 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v2
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v2.教學課程安排 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v3
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v3.訓練師專業及熱忱 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v4
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v4.訓練設備符合產業需求 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v5
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v5.訓練時數 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
#Region "(No Use)"

        'gQ12A = UCase(gQ12A)
        'If gQ12A = "" Then
        '    Reason += "須填寫您參加本次訓練後是否覺得滿意<BR>"
        '    'writeflag = True
        'Else
        '    Select Case gQ12A
        '        Case "Y", "N"
        '        Case Else
        '            Reason += "您參加本次訓練後是否覺得滿意 範圍有誤只能是Y/N<BR>"
        '    End Select
        '    Select Case gQ12A
        '        Case "N"
        '            Dim myvalue As String = ""
        '            If gQ12B <> "" Then
        '                For i As Integer = 1 To 5
        '                    If gQ12B.IndexOf(CStr(i)) > -1 Then
        '                        If myvalue <> "" Then myvalue += ","
        '                        myvalue += CStr(i)
        '                    End If
        '                 Next
        '                gQ12B = myvalue
        '            End If
        '        Case Else
        '            gQ12B = ""
        '    End Select
        'End If

#End Region
        Return Reason
    End Function

    'CHK STUD_RESULTSTUDDATA2
    Function CheckImportData2(ByVal colArray As Array, ByRef writeflag As Boolean) As String
        'amu 20061221 因為同意可寫入某些錯誤的資料，但還是要show訊息
        writeflag = False '一開始都拒絕寫入 
        Dim Reason As String = ""
        Const cst_Len As Integer = 41 '93 '讀卡機可能產生這樣的長度。

        '間隔 驗証 OK@True NG@False  預設為@NG '讀卡機可能長度判斷
        Dim stepFlagChk1 As Boolean = False
        If colArray.Length >= cst_Len Then
            stepFlagChk1 = True
        End If
        ''若長度判斷不OK
        If Not stepFlagChk1 Then
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            Return Reason '離開。
        End If

        FillFormDate = colArray(cst_FillFormDate).ToString '讀卡日期 
        StudID = Convert.ToString(colArray(cst_StudID)) '學號 
        DegreeID = colArray(cst_DegreeID).ToString '學歷 
        MilitaryID = colArray(cst_MilitaryID).ToString '兵役 
        sIdentityID = colArray(cst_IdentityID).ToString '身分 
        Q7 = colArray(cst_Q7).ToString '動機 'Stud_ResultStudData Q7	參訓職訓動機	int
        Q8 = colArray(cst_Q8).ToString '動向 'Stud_ResultStudData Q8	結訓後動向	int
        Q9 = colArray(cst_Q9).ToString '有無工作(有/無) 'Stud_ResultStudData Q9	受訓前有無在工作 char(1)
        Q9Y = colArray(cst_Q9Y).ToString '有在工作 'Stud_ResultStudData Q9Y	受訓前有無在工作 int
        Q10 = colArray(cst_Q10).ToString '找工作(有/無) 'Stud_ResultStudData  Q10	受訓前有無找工作 char(1)
        'Q11 = colArray(cst_Q11).ToString '找到工作(有/無) 'Stud_ResultStudData Q11 參加訓練後	char(1)
        Q11N = colArray(cst_Q11N).ToString '工作幫助 'Stud_ResultStudData Q11N	參加訓練後	int
        gQ12v1 = colArray(cst_Q12v1).ToString '1.參訓職類符合就業市場需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v2 = colArray(cst_Q12v2).ToString '2.教學課程安排 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v3 = colArray(cst_Q12v3).ToString '3.訓練師專業及熱忱 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v4 = colArray(cst_Q12v4).ToString '4.訓練設備符合產業需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v5 = colArray(cst_Q12v5).ToString '5.訓練時數 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意

        '讀卡日期
        If FillFormDate = "" Or Len(FillFormDate) <> 7 Or IsNumeric(FillFormDate) <> True Then
            Reason += "讀卡日期有誤，必須是民國年格式(yyymmdd)<BR>"
        Else
            'FillFormDate = CStr(CInt(Left(FillFormDate, 3)) + 1911) & "/" & Mid(FillFormDate, 4, 2) & "/" & Right(FillFormDate, 2)
            FillFormDate = ChangeTWDate(FillFormDate)
            If IsDate(FillFormDate) = False Then
                Reason += "讀卡日期有誤，必須是民國年格式(yyymmdd)<BR>"
            Else
                If CDate(FillFormDate) < "1900/1/1" Or CDate(FillFormDate) > Now.AddDays(200) Then Reason += "讀卡日期範圍有誤<BR>"
            End If
        End If

        '學號
        If StudID = "" Then
            Reason += "學員序號有誤，必須有值<BR>"
        Else
            Try
                StudID = TIMS.ChangeIDNO(StudID)
                StudID = CStr(CInt(StudID))
            Catch ex As Exception
                Reason += "學員序號有誤，必須為數字格式<BR>"
            End Try
        End If
        'If StudID = "" Or Len(StudID) <> 3 Or IsNumeric(StudID) <> True Then
        '    Reason += "學號有誤，必須是格式(XXX)<BR>"
        'Else
        '    If CStr(StudID) < "000" Or CStr(StudID) > "999" Then Reason += "學號範圍有誤<BR>"
        'End If
        DegreeID = Left(DegreeID, 1) '學歷
        If DegreeID = "" Then
            Reason += "必須填寫最高學歷<BR>"
        Else
            Dim MyKey As String = DegreeID
            If DegreeID.Length < 2 Then MyKey = "0" & DegreeID
            If dtDegree.Select("DegreeID='" & MyKey & "'").Length = 0 Then Reason += "學歷值有錯，不符合鍵詞<BR>"
        End If

        MilitaryID = Left(MilitaryID, 1) '兵役
        If MilitaryID = "" Then
            Reason += "必須填寫兵役狀況<BR>"
        Else
            Dim MyKey As String = MilitaryID
            If MilitaryID.Length < 2 Then MyKey = "0" & MilitaryID
            If dtMilitary.Select("MilitaryID='" & MyKey & "'").Length = 0 Then Reason += "兵役狀況有錯，不符合鍵詞<BR>"
        End If

        'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
        'BY AMU 2009-07-30
        '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
        If sIdentityID = "" Then
            'Reason += "必須填寫結訓身分別<BR>"
        Else
            If Len(sIdentityID) Mod 2 = 1 Then sIdentityID = "0" & sIdentityID
            If Len(sIdentityID) Mod 2 <> 0 Then
                Reason += "結訓身分別不符合鍵詞<BR>" '已整理
            Else
                For i As Integer = 1 To (Len(sIdentityID) / 2)
                    Dim MyKey As String = Mid(sIdentityID.ToString, (i * 2 - 1), 2)
                    'If IdentityID.Length < 2 Then MyKey = "0" & IdentityID Else MyKey = IdentityID
                    If dtIdentity.Select("IdentityID='" & MyKey & "'").Length = 0 Then
                        Reason += "結訓身分別不符合鍵詞<BR>"
                        Exit For
                    End If
                Next
            End If
        End If

        Q7 = Left(Q7, 1)
        If Q7 = "" Then
            Reason += "必須填寫參訓職訓動機<BR>"
        Else
            Try
                Q7 = CInt(Q7)
                If CInt(Q7) < 1 OrElse CInt(Q7) > 6 Then
                    Reason += "參訓職訓動機 範圍有誤<BR>"
                    writeflag = False
                End If
            Catch ex As Exception
                Reason += "參訓職訓動機 範圍有誤<BR>"
                writeflag = False
            End Try
        End If

        Q8 = Left(Q8, 1)
        If Q8 = "" Then
            Reason += "必須填寫結訓後動向<BR>"
        Else
            Try
                Q8 = CInt(Q8)
                If CInt(Q8) < 1 OrElse CInt(Q8) > 5 Then
                    Reason += "結訓後動向 範圍有誤<BR>"
                    writeflag = False
                End If
            Catch ex As Exception
                Reason += "結訓後動向 範圍有誤<BR>"
                writeflag = False
            End Try
        End If

        '假如上列都有填寫，則可存取，否則不可存取
        If Reason = "" Then writeflag = True Else writeflag = False

        Q9 = Left(Q9, 1)
        Q9Y = Left(Q9Y, 1)
        If Q9 = "" Then
            Reason += "必須填寫訓練前有無工作<BR>"
            'writeflag = True
        Else
            'If CInt(Q9) < 1 Or CInt(Q9) > 2 Then Reason += "訓練前有無工作 範圍有誤<BR>"
            Select Case Q9
                Case "1"
                    If Q9Y = "" Then
                        Reason += "訓練前有工作:必須填寫訓練前工作<BR>"
                        'writeflag = True
                    End If
                Case "2"
                    If Q9Y <> "" Then
                        Reason += "訓練前無工作:不必填寫訓練前工作<BR>"
                        'writeflag = True
                    End If
                Case Else
                    Reason += "訓練前有無工作 範圍有誤<BR>"
                    writeflag = False
            End Select
        End If

        If writeflag Then
            If Q9Y <> "" Then
                Try
                    Q9Y = CInt(Q9Y)
                    If CInt(Q9Y) < 1 OrElse CInt(Q9Y) > 3 Then
                        Reason += "訓練前有無工作 範圍有誤<BR>"
                        writeflag = False
                    End If
                Catch ex As Exception
                    Reason += "訓練前有無工作 範圍有誤<BR>"
                    writeflag = False
                End Try
            End If
        End If

        Q10 = Left(Q10, 1)
        If Q10 = "" Then
            Reason += "必須填寫訓練前有否尋找工作<BR>"
            writeflag = False
        Else
            If writeflag Then
                Try
                    Q10 = CInt(Q10)
                    If CInt(Q10) < 1 Or CInt(Q10) > 2 Then
                        Reason += "訓練前有否尋找工作 範圍有誤<BR>"
                        writeflag = False
                    End If
                Catch ex As Exception
                    Reason += "訓練前有否尋找工作 範圍有誤<BR>"
                    writeflag = False
                End Try
            End If
        End If

        If Q11N <> "" Then Q11N = Left(Q11N, 1) '取得最後1碼
        If Q11N <> "" Then Q11N = Val(Q11N) '取得最後1碼轉換數字
        If writeflag Then
            If Q11N = "" Then
                Reason += "參加本訓後找不到工作:必須填寫 您覺得參加本次訓練後，對日後尋找工作幫助的程度<BR>"
                writeflag = False
            Else
                '應該是1~5
                If CInt(Q11N) < 1 Or CInt(Q11N) > 5 Then
                    Reason += "對日後尋找工作幫助的程度 範圍有誤<BR>"
                    writeflag = False
                End If
            End If
        End If

        Select Case gQ12v1
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v1.參訓職類符合就業市場需求 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v2
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v2.教學課程安排 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v3
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v3.訓練師專業及熱忱 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v4
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v4.訓練設備符合產業需求 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select
        Select Case gQ12v5
            Case "1", "2", "3", "4", "5"
            Case Else
                Reason += "12v5.訓練時數 範圍有誤(1~5)<BR>"
                writeflag = False
        End Select

        Return Reason
    End Function

    ''' <summary>'INSERT/UPDATE STUD_RESULTSTUDDATA</summary>
    ''' <param name="colArray"></param>
    ''' <param name="htSS"></param>
    Sub WriteDB1(ByVal colArray As System.Array, ByVal htSS As Hashtable)
        Dim Correct_Dlid As String = TIMS.GetMyValue2(htSS, "Correct_Dlid")
        Dim StudDlid As String = TIMS.GetMyValue2(htSS, "StudDlid")
        Dim vsQID As String = TIMS.GetMyValue2(htSS, "vsQID")
        Dim vsQName As String = TIMS.GetMyValue2(htSS, "vsQName")

        'Dim sql As String = ""
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim da As SqlDataAdapter
        'Dim i As Integer

        FillFormDate = colArray(cst_FillFormDate).ToString '讀卡日期 
        FillFormDate = ChangeTWDate(FillFormDate) '讀卡日期 
        StudID = Convert.ToString(colArray(cst_StudID))  '學號 
        StudID = TIMS.ChangeIDNO(StudID) '學號重整
        StudID = CStr(CInt(StudID))
        StudID = If(StudID >= 100, Right(String.Concat("000", StudID), 3), Right(String.Concat("000", StudID), 2))

        DegreeID = Left(colArray(cst_DegreeID).ToString, 1) '學歷 
        MilitaryID = Left(colArray(cst_MilitaryID).ToString, 1) '兵役 
        'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
        'BY AMU 2009-07-30
        '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
        sIdentityID = colArray(cst_IdentityID).ToString '身分 
        Q7 = Left(colArray(cst_Q7).ToString, 1) '動機 'Stud_ResultStudData Q7	參訓職訓動機	int
        Q8 = Left(colArray(cst_Q8).ToString, 1) '動向 'Stud_ResultStudData Q8	結訓後動向	int
        Q9 = Left(colArray(cst_Q9).ToString, 1) '有無工作(有/無) 'Stud_ResultStudData Q9	受訓前有無在工作 char(1)
        Q9Y = Left(colArray(cst_Q9Y).ToString, 1) '有在工作 'Stud_ResultStudData Q9Y	受訓前有無在工作 int
        Q10 = Left(colArray(cst_Q10).ToString, 1) '找工作(有/無) 'Stud_ResultStudData  Q10	受訓前有無找工作 char(1)
        'Q11 = Left(colArray(cst_Q11).ToString, 1) '找到工作(有/無) 'Stud_ResultStudData Q11 參加訓練後	char(1)
        Q11N = Left(colArray(cst_Q11N).ToString, 1) '工作幫助 'Stud_ResultStudData Q11N	參加訓練後	int
        gQ12v1 = colArray(cst_Q12v1).ToString '1.參訓職類符合就業市場需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v2 = colArray(cst_Q12v2).ToString '2.教學課程安排 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v3 = colArray(cst_Q12v3).ToString '3.訓練師專業及熱忱 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        gQ12v4 = colArray(cst_Q12v4).ToString '4.訓練設備符合產業需求 1.非常同意 2.同意 3.普通 4.不同意 5.非常不同意
        gQ12v5 = colArray(cst_Q12v5).ToString '5.訓練時數 1.非常滿意 2.滿意 3.普通 4.不滿意 5.非常不滿意
        'Q12 = colArray(12).ToString '改進'Stud_ResultTwelveData Q12 資料卡十二項檔	int
        'gQ12A = colArray(cst_Q12A).ToString
        'gQ12B = colArray(cst_Q12B).ToString
        'Q12 = Q12

        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = ""

        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            sql = "SELECT * FROM STUD_RESULTSTUDDATA WHERE SOCID = '" & SOCID & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            Dim dr As DataRow = Nothing
            If dt.Rows.Count = 0 Then
                If Correct_Dlid = "" Then Correct_Dlid = StudDlid
                SubNo = GetNewSubNo(Correct_Dlid, SOCID, trans) 'dr("SubNo")
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("Dlid") = Correct_Dlid
                dr("SubNo") = SubNo 'GetNewSubNo(Correct_Dlid, SOCID, trans)  'CStr(CInt(StudID))
                dr("SOCID") = SOCID
                dr("StdName") = StdName
                'dr("StudentID") = StudID
                dr("StdPID") = StdPID
                dr("Sex") = Sex
                dr("BirthYear") = BirthYear
                dr("BirthMonth") = BirthMonth
                dr("BirthDate") = BirthDate
            Else
                dr = dt.Rows(0)
                'If Correct_Dlid = "" Then Correct_Dlid = dr("DLID")
                Correct_Dlid = dr("Dlid")
                SubNo = dr("SubNo")
                SOCID = dr("SOCID")
            End If
            dr("StudentID") = StudID
            'dr("DegreeID") = DegreeID
            dr("DegreeID") = If(DegreeID.Length < 2, String.Concat("0", DegreeID), DegreeID)
            dr("MilitaryID") = If(MilitaryID.Length < 2, String.Concat("0", MilitaryID), MilitaryID)
            dr("Q7") = If(Q7 <> "", Q7, Convert.DBNull)
            dr("Q8") = If(Q8 <> "", Q8, Convert.DBNull) '
            dr("Q9") = If(Q9 <> "", ChangeYN(Q9), Convert.DBNull) '
            dr("Q9Y") = If(Q9Y <> "", Q9Y, Convert.DBNull) 'Q9Y
            dr("Q10") = If(Q10 <> "", ChangeYN(Q10), Convert.DBNull) 'ChangeYN(Q10)
            'If Q11 <> "" Then dr("Q11") = ChangeYN(Q11)
            dr("Q11N") = If(Q11N <> "", Val(Q11N), Convert.DBNull) ' Val(Q11N)
            dr("Q12v1") = If(gQ12v1 <> "", gQ12v1, Convert.DBNull)
            dr("Q12v2") = If(gQ12v2 <> "", gQ12v2, Convert.DBNull)
            dr("Q12v3") = If(gQ12v3 <> "", gQ12v3, Convert.DBNull)
            dr("Q12v4") = If(gQ12v4 <> "", gQ12v4, Convert.DBNull)
            dr("Q12v5") = If(gQ12v5 <> "", gQ12v5, Convert.DBNull)
            'If gQ12A <> "" Then dr("Q12a") = gQ12A Else dr("Q12a") = Convert.DBNull
            'If gQ12B <> "" Then dr("Q12b") = gQ12B Else dr("Q12b") = Convert.DBNull
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now 'CDate(FillFormDate) 'Now
            DbAccess.UpdateDataTable(dt, da, trans)

            'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
            'BY AMU 2009-07-30
            '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
            sql = "DELETE STUD_RESULTIDENTDATA WHERE DLID='" & Correct_Dlid & "' and SubNo='" & SubNo & "'"
            DbAccess.ExecuteNonQuery(sql, trans)
            'Stud_ResultIdentData不在匯入此 TABLE 改用 Class_StudentsOfClass.IdentityID	參訓身分別代碼
            'BY AMU 2009-07-30
            '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
            'trans = DbAccess.BeginTrans(objconn)
            sql = " SELECT * FROM STUD_RESULTIDENTDATA WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, trans)
            For i As Integer = 0 To (Len(sIdentityID) / 2) - 1
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("DLID") = Correct_Dlid
                dr("SubNo") = SubNo
                dr("IdentityID") = Mid(sIdentityID, i * 2 + 1, 2)
            Next
            DbAccess.UpdateDataTable(dt, da, trans)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            'sr.Close()
            'srr.Close()
            Throw ex
        End Try
    End Sub

    ''' <summary>INSERT/UPDATE STUD_QUESTIONARY</summary>
    ''' <param name="colArray"></param>
    ''' <param name="htSS"></param>
    Sub WriteDB2(ByVal colArray As System.Array, ByVal htSS As Hashtable)
        'Dim Correct_Dlid As String = TIMS.GetMyValue2(htSS, "Correct_Dlid")
        'Dim StudDlid As String = TIMS.GetMyValue2(htSS, "StudDlid")
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then Return
        Dim vsQID As String = TIMS.GetMyValue2(htSS, "vsQID")
        Dim vsQName As String = TIMS.GetMyValue2(htSS, "vsQName")
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = ""
        Dim trans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            'trans = DbAccess.BeginTrans(objconn)
            sql = "DELETE STUD_QUESTIONARY WHERE OCID='" & OCIDValue1.Value & "' AND StudID='" & StudentID12 & "'"
            DbAccess.ExecuteNonQuery(sql, trans)
            'DbAccess.CommitTrans(trans)
            'trans = DbAccess.BeginTrans(objconn)
            'OCID='" & OCIDValue1.Value & "' AND StudID='" & StudentID12 & "' "
            sql = " SELECT * FROM STUD_QUESTIONARY WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, trans)
            Dim dr As DataRow = Nothing
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = OCIDValue1.Value
            dr("StudID") = StudentID12
            dr("FillFormDate") = FillFormDate
            'Dim cxi As Integer = 14 '起始問卷excel欄位位置。
            Dim cxi As Integer = 0  '起始問卷excel欄位位置。
            cxi = cst_Q_Answer_Start_pos '起始問卷excel欄位位置。
            Select Case vsQName
                Case "B" '在職B卷
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_6")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_4")) : cxi += 1
                Case "A" '職前A卷
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_6")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_7")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_6")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_5")) : cxi += 1
                Case "A2" '職前A2卷
                    'Dim cxi As Integer = 14 '起始問卷excel欄位位置。
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q1_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q2_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_6")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q3_7")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_5")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q4_6")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q5_4")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_1")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_2")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_3")) : cxi += 1
                    Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_4")) : cxi += 1
                    'Call CheckcolArray1(colArray(cxi).ToString, dr("Q6_5")) : cxi += 1
            End Select
            'dr("QID") = "1"
            dr("QID") = CInt(vsQID) ' 問卷種類
            'dr_row("QID") = CInt(vsQID)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now ' CDate(FillFormDate) 'Now
            DbAccess.UpdateDataTable(dt, da, trans)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            'sr.Close()
            'srr.Close()
            Throw ex
        End Try
    End Sub

    'INSERT/UPDATE (STUD_SURVEY)
    Sub WriteDB3(ByVal colArray As System.Array, ByVal htSS As Hashtable)
        '16~41
        'colArray(cxi).ToString, "55"
        'Dim SVID As String = TIMS.GetMyValue(SS2, "SVID")
        'Dim vSOCID As String = TIMS.GetMyValue(SS2, "SOCID")
        'Dim SKID2 As String = TIMS.GetMyValue(SS2, "SKID")
        'Dim SQID2 As String = TIMS.GetMyValue(SS2, "SQID")
        'Dim SAID As String = TIMS.GetMyValue(SS2, "SAID")
        Dim sql As String = ""
        sql = " SELECT SQID, SKID FROM ID_SURVEYQUESTION WHERE SVID =4 "
        Dim dtQ As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim vSOCID As String = TIMS.GetMyValue2(htSS, "SOCID")
        Dim SS2 As String = ""
        Dim ff As String = ""
        Dim cxi As Integer = 0  '起始問卷excel欄位位置。
        cxi = cst_Q_Answer_Start_pos '起始問卷excel欄位位置。
        Dim iSQID As Integer = 0
        For iSQID = 55 To 69
            ff = "SQID='" & iSQID & "'"
            If dtQ.Select(ff).Length > 0 Then
                Dim drQ As DataRow = dtQ.Select(ff)(0)
                SS2 = ""
                TIMS.SetMyValue(SS2, "SVID", "4")
                TIMS.SetMyValue(SS2, "SOCID", vSOCID)
                TIMS.SetMyValue(SS2, "SKID", drQ("SKID"))
                TIMS.SetMyValue(SS2, "SQID", drQ("SQID"))
                TIMS.SetMyValue(SS2, "SERIAL2", CStr(colArray(cxi)))
                Call SAVE_STUD_SURVEY(SS2, objconn) '(1~5)
            End If
            cxi += 1
        Next

        iSQID = 71
        ff = "SQID='" & iSQID & "'"
        If dtQ.Select(ff).Length > 0 Then
            Dim drQ As DataRow = dtQ.Select(ff)(0)
            SS2 = ""
            TIMS.SetMyValue(SS2, "SVID", "4")
            TIMS.SetMyValue(SS2, "SOCID", vSOCID)
            TIMS.SetMyValue(SS2, "SKID", drQ("SKID"))
            TIMS.SetMyValue(SS2, "SQID", drQ("SQID"))
            TIMS.SetMyValue(SS2, "SERIAL2", CStr(colArray(cxi)))
            Call SAVE_STUD_SURVEY(SS2, objconn) '(1~5)
        End If
        cxi += 1

        iSQID = 70
        ff = "SQID='" & iSQID & "'"
        If dtQ.Select(ff).Length > 0 Then
            Dim drQ As DataRow = dtQ.Select(ff)(0)
            SS2 = ""
            TIMS.SetMyValue(SS2, "SVID", "4")
            TIMS.SetMyValue(SS2, "SOCID", vSOCID)
            TIMS.SetMyValue(SS2, "SKID", drQ("SKID"))
            TIMS.SetMyValue(SS2, "SQID", drQ("SQID"))
            TIMS.SetMyValue(SS2, "SERIAL2", CStr(colArray(cxi)))
            Call SAVE_STUD_SURVEY(SS2, objconn) '(1~5)
        End If
        cxi += 1

        For iSQID = 72 To 79
            ff = "SQID='" & iSQID & "'"
            If dtQ.Select(ff).Length > 0 Then
                Dim drQ As DataRow = dtQ.Select(ff)(0)
                SS2 = ""
                TIMS.SetMyValue(SS2, "SVID", "4")
                TIMS.SetMyValue(SS2, "SOCID", vSOCID)
                TIMS.SetMyValue(SS2, "SKID", drQ("SKID"))
                TIMS.SetMyValue(SS2, "SQID", drQ("SQID"))
                TIMS.SetMyValue(SS2, "SERIAL2", CStr(colArray(cxi)))
                Call SAVE_STUD_SURVEY(SS2, objconn) '(1~5)
            End If
            cxi += 1
        Next
    End Sub

    Sub SAVE_STUD_SURVEY(ByVal SS2 As String, ByVal oConn As SqlConnection)
        Dim SVID As String = TIMS.GetMyValue(SS2, "SVID")
        Dim vSOCID As String = TIMS.GetMyValue(SS2, "SOCID")
        Dim SKID2 As String = TIMS.GetMyValue(SS2, "SKID")
        Dim SQID2 As String = TIMS.GetMyValue(SS2, "SQID")
        Dim SERIAL2 As String = TIMS.GetMyValue(SS2, "SERIAL2")

        Dim sql As String = ""
        'RadioButtonList (select)
        sql = "" & vbCrLf
        sql &= " SELECT SSID, SQID, SAID " & vbCrLf
        sql &= " FROM STUD_SURVEY " & vbCrLf
        sql &= " WHERE 1=1 AND SOCID = @SOCID AND SQID = @SQID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, oConn)

        ''checkboxlist (delete)
        'sql = "" & vbCrLf
        'sql &= " DELETE STUD_SURVEY WHERE 1=1 AND SOCID = @SOCID AND SQID = @SQID " & vbCrLf
        'Dim dCmd As New SqlCommand(sql, objconn)

        'RadioButtonList /checkboxlist (INSERT)
        sql = "" & vbCrLf
        sql &= " INSERT INTO STUD_SURVEY(SSID, SOCID, DONEDATE, SVID, SKID, SQID, SAID, MODIFYACCT, MODIFYDATE) " & vbCrLf '/*PK*/
        sql &= " VALUES(@SSID, @SOCID, dbo.TRUNC_DATETIME(GETDATE()), @SVID, @SKID, @SQID, @SAID, @MODIFYACCT, GETDATE()) "
        Dim iCmd As New SqlCommand(sql, oConn)
        Dim iSql As String = sql

        'RadioButtonList (update)
        sql = "" & vbCrLf
        sql &= " UPDATE STUD_SURVEY " & vbCrLf
        sql &= " SET SAID = @SAID " & vbCrLf '只修改答案
        sql &= " ,MODIFYACCT = @MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND SOCID = @SOCID" & vbCrLf
        sql &= " AND SQID = @SQID " & vbCrLf
        Dim uCmd As New SqlCommand(sql, oConn)
        Dim uSql As String = sql

        sql = " SELECT SAID, SERIAL FROM ID_SURVEYANSWER WHERE SQID = @SQID "
        Dim sCmd2 As New SqlCommand(sql, oConn)
        Dim dtA As New DataTable
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("SQID", SqlDbType.VarChar).Value = SQID2
            dtA.Load(.ExecuteReader())
        End With
        'Dim dtA As DataTable = DbAccess.GetDataTable(sql, oConn)
        If dtA.Rows.Count = 0 Then Exit Sub
        Dim ff As String = "SERIAL='" & SERIAL2 & "'"
        If dtA.Select(ff).Length = 0 Then Exit Sub
        Dim SAID As Integer = dtA.Select(ff)(0)("SAID")
        Dim dtSS As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = vSOCID
            .Parameters.Add("SQID", SqlDbType.VarChar).Value = SQID2
            dtSS.Load(.ExecuteReader())
        End With
        If dtSS.Rows.Count = 0 Then
            Dim iSSID As Integer = DbAccess.GetNewId(oConn, "STUD_SURVEY_SSID_SEQ,STUD_SURVEY,SSID")
            With iCmd
#Region "(目前不使用)"

                '.Parameters.Clear()
                '.Parameters.Add("SSID", SqlDbType.Int).Value = iSSID
                '.Parameters.Add("SOCID", SqlDbType.Int).Value = vSOCID
                '.Parameters.Add("SVID", SqlDbType.Int).Value = SVID
                '.Parameters.Add("SKID", SqlDbType.Int).Value = Val(SKID2)
                '.Parameters.Add("SQID", SqlDbType.Int).Value = Val(SQID2)
                '.Parameters.Add("SAID", SqlDbType.Int).Value = Val(SAID)
                '.Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.ExecuteNonQuery()

#End Region
                Dim myParam As Hashtable = New Hashtable
                myParam.Add("SSID", iSSID)
                myParam.Add("SOCID", vSOCID)
                myParam.Add("SVID", SVID)
                myParam.Add("SKID", Val(SKID2))
                myParam.Add("SQID", Val(SQID2))
                myParam.Add("SAID", Val(SAID))
                myParam.Add("MODIFYACCT", sm.UserInfo.UserID)
                DbAccess.ExecuteNonQuery(iSql, objconn, myParam)
            End With
        Else
            With uCmd
#Region "(目前不使用)"

                '.Parameters.Clear()
                '.Parameters.Add("SAID", SqlDbType.Int).Value = Val(SAID)
                '.Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("SOCID", SqlDbType.Int).Value = vSOCID
                '.Parameters.Add("SQID", SqlDbType.Int).Value = Val(SQID2)
                '.ExecuteNonQuery()

#End Region
                Dim myParam As Hashtable = New Hashtable
                myParam.Add("SAID", Val(SAID))
                myParam.Add("MODIFYACCT", sm.UserInfo.UserID)
                myParam.Add("SOCID", vSOCID)
                myParam.Add("SQID", Val(SQID2))
                DbAccess.ExecuteNonQuery(uSql, objconn, myParam)
            End With
        End If
    End Sub

    '匯入
    Private Sub Button13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button13.Click
        'Dim MyFile As System.IO.File
        'Dim FileOCIDValue, MyFileName As String
        'Dim MyFileType As String
        'Dim flag As String
        '建立StudentID欄位值
        If OCIDValue1.Value = "" Then
            'Reason += "未選擇 開班編號(OCID) 無法匯入<BR>"
            Common.MessageBox(Me, "未選擇 職類/班別 無法匯入!!")
            Exit Sub
        End If
        If File1.Value = "" Then
            Common.MessageBox(Me, "請確認檔案位置!!")
            Exit Sub
        End If
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drC Is Nothing Then
            Common.MessageBox(Me, "查無班級資料，請重新查詢!!")
            Exit Sub
        End If
        'If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(drC("TPlanID")) > -1 Then
        '    Common.MessageBox(Me, "適用於目前職前訓練計畫，請重新查詢!!")
        '    Exit Sub
        'End If

        Dim vsQName As String = ""
        Dim vsQID As String = ""
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If

        Dim Correct_Dlid As String = "" '取得封面
        Correct_Dlid = getDLIDforOCID(OCIDValue1.Value)
        If Correct_Dlid = "" Then
            Common.MessageBox(Me, "依據開班編號(OCID) 尚未建立結訓資料卡封面檔(STUD_DATALID)!!")
            '離開此功能()
            Exit Sub
        End If

        blnPrint2016 = False
        '班級結訓日大於等於2016/05/01
        '指定2016/5/1(含)之後，所結訓的班級才使用新的問卷內容。
        Dim cTPlanID As String = TIMS.GetTPlanID(drC("PlanID"), objconn)
        If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(cTPlanID) > -1 Then
            If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(drC("FTDATE"))) >= 0 Then
                blnPrint2016 = True
            End If
        End If
        Common.SetListItem(rblprtType1, If(blnPrint2016, Cst_defQA16, Cst_defQA2))

        '檢查檔案格式與大小----------   Start
        Const cst_filetype As String = "csv"
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If

        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue = Split(Split(MyFileName, "-")(1), ".")(0)

        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        Else
            MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
            If LCase(MyFileType) <> cst_filetype Then
                Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                Exit Sub
            End If
        End If

        '檢查檔案格式與大小----------   End
        'If TIMS.sUtl_ChkTest Then
        '    Call sUtl_ImportA16(File1, MyFileName)
        '    Exit Sub
        'End If

        If blnPrint2016 Then
            Common.SetListItem(rblprtType1, Cst_defQA16)
            Call sUtl_ImportA16(File1, MyFileName)
            Exit Sub
        End If
        Call sUtl_ImportA2_B(File1, MyFileName)
    End Sub

    '匯入職前(20160501) 
    Sub sUtl_ImportA16(ByRef File1 As HtmlInputFile, ByVal MyFileName As String)
        Const cst_UploadPath As String = "~/SD/11/Temp/"
        Call TIMS.MyCreateDir(Me, cst_UploadPath)
        Const cst_flag As String = ","

        Dim vsQName As String = ""
        Dim vsQID As String = ""
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If

        '上傳檔案
        File1.PostedFile.SaveAs(Server.MapPath(cst_UploadPath & MyFileName))
        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

        '將檔案讀出放入記憶體
        Dim sr As System.IO.Stream
        Dim srr As System.IO.StreamReader
        sr = IO.File.OpenRead(Server.MapPath(cst_UploadPath & MyFileName))
        srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

        'Dim OneRow As String        'srr.ReadLine 一行一行的資料
        'Dim col As String           '欄位
        'Dim colArray As Array

        '取出資料庫的所有欄位--------   Start
        '建立Next_Dlid值 (新值，不會重複)
        'sql = "Select distinct max(dlid)+1 Next_Dlid from stud_resultstuddata "
        'Dim Next_Dlid As String = DbAccess.ExecuteScalar(sql)
        'Dim Next_Dlid As String
        'Dim writeflag As Boolean

        'Dim BasicSID As String = TIMS.Get_DateNo
        'Dim SIDNum As Integer = 1
        'Dim SID As String
        'Dim Reason As String                '儲存錯誤的原因

        '建立錯誤資料格式Table----------------Start
        'Dim drWrong As DataRow = Nothing
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        dtWrong.Columns.Add(New DataColumn("Index")) '序列
        dtWrong.Columns.Add(New DataColumn("FillFormDate")) '讀卡日期
        dtWrong.Columns.Add(New DataColumn("StudID")) '學號
        dtWrong.Columns.Add(New DataColumn("OCID")) '班級編號
        dtWrong.Columns.Add(New DataColumn("Result")) '情況
        dtWrong.Columns.Add(New DataColumn("Reason")) '說明
        '建立錯誤資料格式Table----------------End

        '取出所有鍵值當判斷-----------------------------------Start
        Dim sql As String = ""
        sql = "SELECT * FROM dbo.KEY_DEGREE WHERE 1=1 AND DEGREETYPE IN ('0','1') ORDER BY 1"
        dtDegree = DbAccess.GetDataTable(sql, objconn) '學歷
        sql = "SELECT * FROM dbo.KEY_MILITARY ORDER BY 1"
        dtMilitary = DbAccess.GetDataTable(sql, objconn) '兵役
        sql = "SELECT * FROM dbo.KEY_IDENTITY ORDER BY 1"
        dtIdentity = DbAccess.GetDataTable(sql, objconn) '身分別
        '取出所有鍵值當判斷-----------------------------------End

        Try
            'Dim int_cst_Len2 As Integer = 0
            'Const cst_Len2a2 As Integer = 43 '整理過後的長度
            'Const cst_Len2b As Integer = 34
            'Select Case vsQName
            '    Case "B" '在職B卷
            '        int_cst_Len2 = cst_Len2b
            '        'Case "A" '職前A卷
            '        '    int_cst_Len2 = cst_Len2a
            '    Case "A2" '職前A2卷
            '        int_cst_Len2 = cst_Len2a2
            'End Select
            Dim Reason As String = "" '儲存錯誤的原因
            Dim iRowIdx As Integer = 0 '讀取行累計數
            Do While srr.Peek >= 0
                Dim OneRow As String = srr.ReadLine
                If Replace(OneRow, cst_flag, "") = "" Then Exit Do '若資料為空白行，則離開回圈

                If iRowIdx <> 0 Then
                    Reason = ""
                    Dim colArray As Array = Split(OneRow, cst_flag)
                    'colArray = Split(OneRow, cst_flag)

                    '==補強EXCEL可能去除零值之可能性==Start
                    'AMU 2006/12/14
                    Try
                        If colArray.Length > 4 Then
                            If Len(colArray(cst_FillFormDate).ToString) = 6 Then
                                colArray(cst_FillFormDate) = String.Concat("0", colArray(cst_FillFormDate)) '.ToString
                            End If
                            Select Case Len(Convert.ToString(colArray(cst_StudID)))
                                Case 1
                                    colArray(cst_StudID) = "00" & Convert.ToString(colArray(cst_StudID))
                                Case 2
                                    colArray(cst_StudID) = "0" & Convert.ToString(colArray(cst_StudID))
                            End Select
                            If Len(Convert.ToString(colArray(cst_IdentityID))) Mod 2 = 1 Then
                                colArray(cst_IdentityID) = "0" & Convert.ToString(colArray(cst_IdentityID))
                            End If
                        Else
                            Reason += "文件格式有誤 無法匯入<BR>"
                            Exit Do
                        End If
                    Catch ex As Exception
                        Reason += "文件格式有誤 無法匯入<BR>"
                        Exit Do
                    End Try
                    '==補強EXCEL可能去除零值之可能性==End

                    '建立StudentID欄位值
                    If OCIDValue1.Value = "" Then
                        Reason += "未選擇 開班編號(OCID) 無法匯入<BR>"
                        Exit Do
                    End If

                    Dim Correct_Dlid As String = "" '取得封面
                    If Reason = "" Then
                        Correct_Dlid = getDLIDforOCID(OCIDValue1.Value)
                        If Correct_Dlid = "" Then
                            Reason += "依據開班編號(OCID) 尚未建立結訓資料卡封面檔(STUD_DATALID) <BR>"
                            Exit Do
                        End If
                    End If

                    Dim StudDlid As String = "" '查詢學員有無填寫。
                    If Reason = "" Then Reason += Check_Class_StudentOfClass(colArray(cst_StudID).ToString, OCIDValue1.Value, StudDlid, SOCID, StudentID12, StdName, StdPID, Sex, BirthYear, BirthMonth, BirthDate) '檢查學生是否存在
                    Dim writeflag As Boolean = False
                    Dim int_cst_Len2 As Integer = 0
                    Const cst_Len2a2 As Integer = 43 '整理過後的長度
                    Select Case vsQName
                        Case "A2" '職前A2卷
                            int_cst_Len2 = cst_Len2a2
                    End Select
                    If Reason = "" Then Reason += CheckImportData2(colArray, writeflag) '檢查資料正確性

                    '通過檢查，開始輸入資料---------------------Start
                    Dim htSS As New Hashtable
                    If SOCID <> "" Then
                        TIMS.SetMyValue2(htSS, "Correct_Dlid", Correct_Dlid)
                        TIMS.SetMyValue2(htSS, "StudDlid", StudDlid)
                        TIMS.SetMyValue2(htSS, "vsQID", vsQID)
                        TIMS.SetMyValue2(htSS, "vsQName", vsQName)
                        TIMS.SetMyValue2(htSS, "SOCID", SOCID)
                        TIMS.SetMyValue2(htSS, "OCID", OCIDValue1.Value)
                    End If

                    If SOCID <> "" AndAlso Reason = "" Then
                        Call WriteDB1(colArray, htSS)
                        Call WriteDB3(colArray, htSS)
                    Else
                        If SOCID <> "" AndAlso writeflag Then
                            Call WriteDB1(colArray, htSS)
                            Call WriteDB3(colArray, htSS)
                        End If

                        '錯誤資料，填入錯誤資料表
                        Dim drWrong As DataRow = Nothing
                        drWrong = dtWrong.NewRow
                        dtWrong.Rows.Add(drWrong)
                        drWrong("Index") = iRowIdx
                        drWrong("FillFormDate") = "資料不足" '讀卡日期
                        drWrong("StudID") = "資料不足" '學號
                        drWrong("OCID") = "資料不足" '開班編號
                        drWrong("Result") = "未轉入"
                        drWrong("Reason") = "資料不足" 'Reason
                        If Reason <> "" Then drWrong("Reason") = Reason 'Reason
                        If colArray.Length > 5 Then
                            drWrong("FillFormDate") = colArray(cst_FillFormDate) '讀卡日期
                            drWrong("StudID") = colArray(cst_StudID) '學號
                            drWrong("OCID") = OCIDValue1.Value '開班編號
                            If writeflag Then
                                drWrong("Result") = "試著執行轉入"
                                If Reason <> "" Then drWrong("Reason") += "-請修正或移至「結訓學員資料卡登入」查看修正"
                            End If
                        End If
                    End If
                End If
                iRowIdx += 1 '讀取行累計數
            Loop
            '開始判別欄位存入------------   End
            If dtWrong.Rows.Count = 0 Then
                If Reason = "" Then
                    Common.MessageBox(Me, "資料匯入成功")
                Else
                    Common.MessageBox(Me, Reason)
                End If
            Else
                Session("MyWrongTable") = dtWrong
                Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_11_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            End If
            sr.Close()
            srr.Close()
            TIMS.MyFileDelete(Server.MapPath(cst_UploadPath & MyFileName))
            'IO.File.Delete(Server.MapPath(cst_UploadPath & MyFileName))
        Catch ex As Exception
            sr.Close()
            srr.Close()
            TIMS.MyFileDelete(Server.MapPath(cst_UploadPath & MyFileName))
            'IO.File.Delete(Server.MapPath(cst_UploadPath & MyFileName))
        End Try
        'Button1_Click(sender, e)
    End Sub

    '匯入職前(OLD)或在職()
    Sub sUtl_ImportA2_B(ByRef File1 As HtmlInputFile, ByVal MyFileName As String)
        Const cst_UploadPath As String = "~/SD/11/Temp/"
        Call TIMS.MyCreateDir(Me, cst_UploadPath)
        Const cst_flag As String = ","
        Dim vsQName As String = ""
        Dim vsQID As String = ""
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If

        '上傳檔案
        File1.PostedFile.SaveAs(Server.MapPath(cst_UploadPath & MyFileName))
        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

        '將檔案讀出放入記憶體
        Dim sr As System.IO.Stream
        Dim srr As System.IO.StreamReader
        sr = IO.File.OpenRead(Server.MapPath(cst_UploadPath & MyFileName))
        srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

        Dim RowIndex As Integer = 0 '讀取行累計數
        Dim OneRow As String        'srr.ReadLine 一行一行的資料
        'Dim col As String           '欄位
        Dim colArray As Array

        '取出資料庫的所有欄位--------   Start
        'Dim sql As String
        'Dim dtStuOfClass As DataTable
        'Dim dr As DataRow
        'Dim da As SqlDataAdapter
        'Dim dt As DataTable
        'Dim STDate As Date
        'Dim i As Integer

        '建立Next_Dlid值 (新值，不會重複)
        'sql = "Select distinct max(dlid)+1 Next_Dlid from stud_resultstuddata "
        'Dim Next_Dlid As String = DbAccess.ExecuteScalar(sql)
        'Dim Next_Dlid As String
        Dim writeflag As Boolean = False

        Dim BasicSID As String = TIMS.Get_DateNo
        Dim SIDNum As Integer = 1
        Dim SID As String = ""
        Dim Reason As String = ""               '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing

        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("FillFormDate"))
        dtWrong.Columns.Add(New DataColumn("StudID"))
        dtWrong.Columns.Add(New DataColumn("OCID"))
        dtWrong.Columns.Add(New DataColumn("Result"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        '取出所有鍵值當判斷-----------------------------------Start
        Dim sql As String = ""
        sql = "SELECT * FROM dbo.KEY_DEGREE WHERE DegreeType IN ('0','1') ORDER BY 1"
        dtDegree = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT * FROM dbo.KEY_MILITARY ORDER BY 1"
        dtMilitary = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT * FROM dbo.KEY_IDENTITY ORDER BY 1"
        dtIdentity = DbAccess.GetDataTable(sql, objconn)
        '取出所有鍵值當判斷-----------------------------------End

        Try
            Dim int_cst_Len2 As Integer = 0
            Const cst_Len2a2 As Integer = 43 '整理過後的長度
            Const cst_Len2b As Integer = 34
            Select Case vsQName
                Case "B" '在職B卷
                    int_cst_Len2 = cst_Len2b
                    'Case "A" '職前A卷 ' int_cst_Len2 = cst_Len2a
                Case "A2" '職前A2卷
                    int_cst_Len2 = cst_Len2a2
            End Select
            Do While srr.Peek >= 0
                OneRow = srr.ReadLine
                If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈
                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, cst_flag)
                    '==補強EXCEL可能去除零值之可能性==Start
                    'AMU 2006/12/14
                    Try
                        If colArray.Length > 4 Then
                            If Len(colArray(cst_FillFormDate).ToString) = 6 Then colArray(cst_FillFormDate) = String.Concat("0", colArray(cst_FillFormDate))
                            Select Case Len(Convert.ToString(colArray(cst_StudID)))
                                Case 1
                                    colArray(cst_StudID) = String.Concat("00", colArray(cst_StudID))
                                Case 2
                                    colArray(cst_StudID) = String.Concat("0", colArray(cst_StudID))
                            End Select
                            If Len(Convert.ToString(colArray(cst_IdentityID))) Mod 2 = 1 Then
                                colArray(cst_IdentityID) = String.Concat("0", colArray(cst_IdentityID))
                            End If
                        Else
                            Reason += "文件格式有誤 無法匯入<BR>"
                            Exit Do
                        End If
                    Catch ex As Exception
                        Reason += "文件格式有誤 無法匯入<BR>"
                        Exit Do
                    End Try
                    '==補強EXCEL可能去除零值之可能性==End

                    '建立StudentID欄位值
                    If OCIDValue1.Value = "" Then
                        Reason += "未選擇 開班編號(OCID) 無法匯入<BR>"
                        Exit Do
                    End If
                    Dim Correct_Dlid As String = ""
                    If Reason = "" Then
                        Correct_Dlid = getDLIDforOCID(OCIDValue1.Value)
                        If Correct_Dlid = "" Then
                            Reason += "依據開班編號(OCID) 尚未建立結訓資料卡封面檔(STUD_DATALID) <BR>"
                            Exit Do
                        End If
                    End If

                    Dim StudDlid As String = "" '查詢學員有無填寫。
                    If Reason = "" Then Reason += Check_Class_StudentOfClass(colArray(cst_StudID).ToString, OCIDValue1.Value, StudDlid, SOCID, StudentID12, StdName, StdPID, Sex, BirthYear, BirthMonth, BirthDate) '檢查學生是否存在
                    writeflag = False
                    If Reason = "" Then Reason += CheckImportData(colArray, writeflag, int_cst_Len2) '檢查資料正確性

                    '通過檢查，開始輸入資料---------------------Start
                    Dim htSS As New Hashtable
                    If SOCID <> "" Then
                        TIMS.SetMyValue2(htSS, "Correct_Dlid", Correct_Dlid)
                        TIMS.SetMyValue2(htSS, "StudDlid", StudDlid) '
                        TIMS.SetMyValue2(htSS, "vsQID", vsQID)
                        TIMS.SetMyValue2(htSS, "vsQName", vsQName)
                    End If
                    If SOCID <> "" AndAlso Reason = "" Then
                        Call WriteDB1(colArray, htSS)
                        Call WriteDB2(colArray, htSS)
                    Else
                        If SOCID <> "" AndAlso writeflag Then
                            '若 writeflag 為true 則可繼續新增到資料庫
                            Call WriteDB1(colArray, htSS)
                            Call WriteDB2(colArray, htSS)
                        End If
                        Try
                            '錯誤資料，填入錯誤資料表
                            drWrong = dtWrong.NewRow
                            dtWrong.Rows.Add(drWrong)
                            drWrong("Index") = RowIndex
                            drWrong("FillFormDate") = "資料不足" '讀卡日期
                            drWrong("StudID") = "資料不足" '學號
                            drWrong("OCID") = "資料不足" '開班編號
                            drWrong("Result") = "未轉入"
                            drWrong("Reason") = "資料不足" 'Reason
                            If Reason <> "" Then drWrong("Reason") = Reason 'Reason
                            If colArray.Length > 5 Then
                                drWrong("FillFormDate") = colArray(cst_FillFormDate) '讀卡日期
                                drWrong("StudID") = colArray(cst_StudID) '學號
                                drWrong("OCID") = OCIDValue1.Value '開班編號
                                If writeflag Then
                                    drWrong("Result") = "試著執行轉入"
                                    If Reason <> "" Then drWrong("Reason") += "-請修正或移至「結訓學員資料卡登入」查看修正"
                                End If
                            End If
                        Catch ex As Exception
                            drWrong("Reason") = "系統錯誤：" & ex.Message.ToString
                        End Try
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            '開始判別欄位存入------------   End
            If dtWrong.Rows.Count = 0 Then
                If Reason = "" Then
                    Common.MessageBox(Me, "資料匯入成功")
                Else
                    Common.MessageBox(Me, Reason)
                End If
            Else
                Session("MyWrongTable") = dtWrong
                Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_11_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            End If
            sr.Close()
            srr.Close()
            TIMS.MyFileDelete(Server.MapPath(cst_UploadPath & MyFileName))
            'IO.File.Delete(Server.MapPath(cst_UploadPath & MyFileName))
        Catch ex As Exception
            sr.Close()
            srr.Close()
            TIMS.MyFileDelete(Server.MapPath(cst_UploadPath & MyFileName))
            'IO.File.Delete(Server.MapPath(cst_UploadPath & MyFileName))
        End Try
        'Button1_Click(sender, e)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        DataGrid1.Style.Item("display") = "none"
    End Sub

    '列印空白表
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Dim vsQName As String = ""
        Dim vsQID As String = ""
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If

        Dim sPrtVal As String = ""
        Dim sPage_Url As String = ""
        sPage_Url = Cst_Page_Url_A12 'Cst_Page_Url
        Select Case vsQID
            Case Cst_defQA2
                sPage_Url = Cst_Page_Url_A12
        End Select
        blnPrint2016 = False
        Select Case rblprtType1.SelectedValue
            Case Cst_defQA16
                blnPrint2016 = True
        End Select
        If blnPrint2016 Then
            sPage_Url = Cst_Page_Url_A16 'Cst_Page_Url
            sPrtVal = ""
            sPrtVal &= "?ID=" & Request("ID")
            sPrtVal &= "&ProcessType=" & cst_ptPrint
            sPrtVal &= "&SVID=4"
            Call TIMS.OpenWin1(Me, sPage_Url & sPrtVal)
            'TIMS.Utl_Redirect1(Me, sPage_Url & sPrtVal)
            Exit Sub
        End If

        'blnPrint2016 = False
        ''班級結訓日大於等於2016/05/01
        ''指定2016/5/1(含)之後，所結訓的班級才使用新的問卷內容。
        'If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(drC("TPlanID")) > -1 Then
        '    If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(drC("FTDATE"))) >= 0 Then blnPrint2016 = True
        'End If
        Call GetSearchStr()
        sPrtVal = ""
        sPrtVal &= "?ID=" & Request("ID")
        sPrtVal &= "&ProcessType=" & cst_ptPrint
        Call TIMS.OpenWin1(Me, sPage_Url & sPrtVal)
        'TIMS.Utl_Redirect1(Me, sPage_Url & "?ProcessType=" & cst_ptPrint & "&ID=" & Request("ID"))
    End Sub

    Private Sub DG_stud_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_stud.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim vStudentid As String = TIMS.GetMyValue(sCmdArg, "studentid")
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim vSOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim vsQName As String = ""
        Dim vsQID As String = ""
        If Not TIMS.GetQType(Me, vsQName, vsQID) Then
            '計畫未設定問卷類型, 請先設定後, 再使用匯入功能
            Common.MessageBox(Me, "計畫未設定問卷類型，請先設定後，再使用此功能")
            '離開此功能()
            Exit Sub
        End If
        Dim sPage_Url As String = Cst_Page_Url_A12 'Cst_Page_Url
        Select Case vsQID
            Case Cst_defQA2
                sPage_Url = Cst_Page_Url_A12
        End Select

        blnPrint2016 = False
        Select Case Hid_rblprtType1.Value
            Case Cst_defQA16
                blnPrint2016 = True
                sPage_Url = Cst_Page_Url_A16 'Cst_Page_Url
                If vSOCID = "" Then Exit Sub
                If vOCID = "" Then Exit Sub
            Case Else
                If vStudentid = "" Then Exit Sub
                If vOCID = "" Then Exit Sub
        End Select

        'Dim drC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        'If drC Is Nothing Then
        '    Common.MessageBox(Me, "查無班級資料，請重新查詢!!")
        '    Exit Sub
        'End If
        'Dim cTPlanID As String = TIMS.GetTPlanID(drC("PlanID"), objconn)
        ''班級結訓日大於等於2016/05/01
        ''指定2016/5/1(含)之後，所結訓的班級才使用新的問卷內容。
        'blnPrint2016 = False
        'If TIMS.Cst_TPlanID_NewSuy201605.IndexOf(cTPlanID) > -1 Then
        '    If DateDiff(DateInterval.Day, CDate(TIMS.cst_NewSuyFD20160501), CDate(drC("FTDATE"))) >= 0 Then
        '        blnPrint2016 = True
        '    End If
        'End If
        'If blnPrint2016 Then
        '    sPage_Url = Cst_Page_Url_A16 'Cst_Page_Url
        'End If

        Call GetSearchStr()
        Dim myValue As String = ""
        myValue = "?ID=" & CStr(Request("ID"))
        myValue &= "&Stuedntid=" & vStudentid
        myValue &= "&socid=" & vSOCID
        myValue &= "&ocid=" & vOCID
        If blnPrint2016 Then myValue &= "&SVID=4"

        Select Case e.CommandName
            Case "insert"
                myValue &= "&ProcessType=" & cst_ptInsert
                TIMS.Utl_Redirect1(Me, sPage_Url & myValue)
            Case "clear"
                'GetSearchStr()
                myValue &= "&ProcessType=" & cst_ptDel
                TIMS.Utl_Redirect1(Me, sPage_Url & myValue)
                'Dim strScript As String
                'strScript = "<script language=""javascript"">" + vbCrLf
                'strScript += "if (window.confirm('此動作會刪除期末學員滿意度調查表資料，是否確定刪除?')){" + vbCrLf
                'strScript += "location.href ='" & sPage_Url & "?ProcessType=del&Stuedntid=" & vStudentid & "&ocid=" & vOCID & "&ID=" & Request("ID") & "';}" + vbCrLf
                'strScript += "</script>"
                'Page.RegisterStartupScript("", strScript)
            Case "check" '查詢
                myValue &= "&ProcessType=" & cst_ptCheck
                TIMS.Utl_Redirect1(Me, sPage_Url & myValue)
            Case "Edit" '修改
                myValue &= "&ProcessType=" & cst_ptEdit
                TIMS.Utl_Redirect1(Me, sPage_Url & myValue)
            Case "print" '列印
                myValue &= "&ProcessType=" & cst_ptPrint2
                Call TIMS.OpenWin1(Me, sPage_Url & myValue)
                'TIMS.Utl_Redirect1(Me, sPage_Url & myValue)
        End Select
    End Sub

    Private Sub DG_stud_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_stud.ItemDataBound
        'Const Cst_學號 As Integer = 0
        Const Cst_姓名 As Integer = 1
        Const Cst_填寫狀態 As Integer = 2
        'Const Cst_功能 As Integer = 3
        'Const Cst_OCID As Integer = 4
        'Const Cst_StudentID As Integer = 5
        Dim ff As String = ""
        Dim dtQUE As DataTable = Nothing
        If Not Session("dtQUE") Is Nothing Then dtQUE = Session("dtQUE")

        Dim drv As DataRowView = e.Item.DataItem '資料取得
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '<asp@Button ID="BtnAdd4" runat="server" Text="新增" CommandName="insert"></asp@Button>
                '<asp@Button ID="BtnEdit" runat="server" Text="修改" CommandName="Edit"></asp@Button>
                '<asp@Button ID="BtnQry5" runat="server" Text="查詢" CommandName="check"></asp@Button>
                '<asp@Button ID="BtnClear6" runat="server" Text="清除重填" CommandName="clear"></asp@Button>
                '<asp@Button ID="BtnPrint" runat="server" Text="列印" CommandName="print"></asp@Button>
                Dim BtnAdd4 As Button = e.Item.FindControl("BtnAdd4") '新增
                Dim BtnEdit As Button = e.Item.FindControl("BtnEdit") '修改
                Dim BtnQry5 As Button = e.Item.FindControl("BtnQry5") '查看
                Dim BtnClear6 As Button = e.Item.FindControl("BtnClear6") '清除重填
                BtnClear6.Attributes.Add("onclick", "return confirm('此動作會刪除期末學員滿意度調查表資料，是否確定刪除?');")
                'BtnClear6.Attributes("onclick") = "return confirm('此動作會刪除期末學員滿意度調查表資料，是否確定刪除?');"
                Dim BtnPrint As Button = e.Item.FindControl("BtnPrint") '列印
                'e.Item.Cells(Cst_學號).Text = Right(drv("StudentID"), 2)
                Dim sName As String = Convert.ToString(drv("NAME"))
                If drv("RejectTDate1").ToString <> "" Then sName &= "(" & FormatDateTime(drv("RejectTDate1"), 2) & ")"
                If drv("RejectTDate2").ToString <> "" Then sName &= "(" & FormatDateTime(drv("RejectTDate2"), 2) & ")"
                e.Item.Cells(Cst_姓名).Text = sName
                e.Item.Cells(Cst_填寫狀態).Text = "否"
                'BtnAdd4.Visible = False '不可新增
                'If check_add.Value = "1" Then BtnAdd4.Visible = True '可新增
                BtnAdd4.Visible = True '可新增
                BtnPrint.Visible = False '不開放列印
                BtnEdit.Visible = False '不可以修改
                BtnQry5.Visible = False '不可查看
                BtnClear6.Visible = False '不可清除重填

                If Not dtQUE Is Nothing Then
                    Select Case Hid_rblprtType1.Value
                        Case Cst_defQA16
                            ff = "SOCID= '" & Convert.ToString(drv("SOCID")) & "'"
                        Case Else
                            ff = "StudID= '" & Convert.ToString(drv("studentid")) & "'"
                    End Select

                    If dtQUE.Select(ff).Length > 0 Then
                        e.Item.Cells(Cst_填寫狀態).Text = "是"
                        BtnAdd4.Visible = False '不可新增
                        BtnEdit.Visible = False '不可以修改
                        If sm.UserInfo.TPlanID = "02" Then '自辦職前可以修改
                            BtnEdit.Visible = True '可以修改
                            BtnPrint.Visible = True '開放列印
                        End If

                        BtnQry5.Visible = True '查看
                        BtnPrint.Visible = True '開放列印

                        BtnClear6.Visible = True '可清除重填(未結訓)
                        If Convert.ToString(drv("StudStatus")) = "5" Then
                            '不可清除重填(該學員已結訓或該班已結訓)
                            BtnClear6.Visible = False
                            vMsg = "該學員已結訓"
                            TIMS.Tooltip(BtnClear6, vMsg)
                            'BtnEdit.Enabled = False '不可以修改
                            'TIMS.Tooltip(BtnEdit, vMsg)
                        End If

                        '兩者功能皆沒有時,不能使用(修、刪功能)
                        'If check_mod.Value = "0" AndAlso check_del.Value = "0" Then
                        '    BtnEdit.Visible = False '不可以修改
                        '    TIMS.Tooltip(BtnEdit, "功能權限未開放")
                        '    BtnClear6.Visible = False '不可以清除重填
                        '    TIMS.Tooltip(BtnClear6, "功能權限未開放")
                        'End If
                    End If
                End If

                'If Convert.ToString(drv("IsClosed")) = "Y" Then
                '    but4.Enabled = False '不可新增
                '    Edit.Enabled = False '不可以修改
                '    TIMS.Tooltip(but4, "該班級已做結訓")
                '    TIMS.Tooltip(Edit, "該班級已做結訓")
                'End If

                If Convert.ToString(drv("RejectTDate1")) <> "" _
                    OrElse Convert.ToString(drv("RejectTDate2")) <> "" Then
                    BtnAdd4.Enabled = False '不開放新增
                    BtnPrint.Enabled = False '不開放列印
                    BtnPrint.Visible = False '不開放列印
                    TIMS.Tooltip(BtnAdd4, "該學員已做離、退訓")
                    TIMS.Tooltip(BtnPrint, "該學員已做離、退訓")
                End If

                Dim sCmdArg As String = ""
                Select Case Hid_rblprtType1.Value
                    Case Cst_defQA16
                        TIMS.SetMyValue(sCmdArg, "studentid", Convert.ToString(drv("studentid")))
                        TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                        TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
                    Case Else
                        TIMS.SetMyValue(sCmdArg, "studentid", Convert.ToString(drv("studentid")))
                        TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                        TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
                End Select

                If BtnAdd4.Enabled Then BtnAdd4.CommandArgument = sCmdArg '新增
                If BtnEdit.Enabled Then BtnEdit.CommandArgument = sCmdArg '修改
                If BtnQry5.Enabled Then BtnQry5.CommandArgument = sCmdArg '查看
                If BtnClear6.Enabled Then BtnClear6.CommandArgument = sCmdArg '清除重填
                If BtnPrint.Enabled Then BtnPrint.CommandArgument = sCmdArg '列印
        End Select
    End Sub

End Class