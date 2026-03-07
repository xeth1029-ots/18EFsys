Partial Class SD_05_019_R_1
    Inherits AuthBasePage

    'Dim CheckMode As String
    Dim PageSize As Int16 = 24   '分頁-- 每頁資料 筆數
    Dim CourseNumSize As Int16 = 4 '每頁科目顯示數目

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
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Dim rqOCID As String = TIMS.sUtl_GetRqValue(Me, "OCID")
        If (Hid_OCID.Value = "") Then Hid_OCID.Value = rqOCID

        If Not IsPostBack Then
            lbmsg.Text = ""
            Dim dt3 As DataTable = GetDatatable(objconn, 3, Hid_OCID.Value, "") '3.取單位名稱
            Me.ViewState("OrgName") = If(dt3.Rows.Count > 0, Convert.ToString(dt3.Rows(0)("OrgName")), "")
            Dim dt7 As DataTable = GetDatatable(objconn, 7, Hid_OCID.Value, "") '7'計畫名稱
            Me.ViewState("PLANNAME") = If(dt7.Rows.Count > 0, Convert.ToString(dt7.Rows(0)("PLANNAME")), "")
            Call PrintResult()    '列出所有分頁資料
        End If
    End Sub

    Private Sub PrintResult(Optional ByVal pageIndex As Int16 = 0)
        Call CreateRpt(10)
    End Sub

    Private Sub PrintEmptPage()  '列印每頁空白資料列
        Dim nl As HtmlGenericControl = New HtmlGenericControl
        print_content.Controls.Add(nl)
        nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
    End Sub


    ''' <summary> 取得參數設定值SYS_GLOBALVAR </summary>
    ''' <param name="GVID"></param>
    ''' <param name="DistID"></param>
    ''' <param name="TPlanID"></param>
    ''' <param name="ItemVar"></param>
    ''' <returns></returns>
    Public Shared Function GetPara(ByRef oConn As SqlConnection, ByVal GVID As String, ByVal DistID As String, ByVal TPlanID As String, ByVal ItemVar As Int16) As Double  '
        Dim iRst As Double = 0

        Dim sTmp As String = ""
        Select Case GVID
            Case "3"
                '操行底分
                sTmp = TIMS.GetGlobalVar(DistID, TPlanID, "3", ItemVar, oConn)
                If sTmp <> "" Then iRst = Val(sTmp)
            Case "13"
                'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
                iRst = 1
                sTmp = TIMS.GetGlobalVar(DistID, TPlanID, "13", ItemVar, oConn)
                'If Me.ViewState("ResultTyp") = "" Then iRst = Val(sTmp) '預設取各科平均法
            Case "17"
                '學術科百分比 (學科ItemVar1 ex:0.4,術科ItemVar2  ex:0.4) 
                'ItemVar = 1 '學科百分比
                'ItemVar = 2 '術科百分比
                iRst = 0
                sTmp = TIMS.GetGlobalVar(DistID, TPlanID, "17", ItemVar, oConn)
                If sTmp <> "" Then iRst = Math.Round(Convert.ToDouble(sTmp), 2)

        End Select
        Return iRst
    End Function

    ''' <summary> 取得課程訓練時數資料表 </summary>
    ''' <param name="OCID"></param>
    ''' <returns></returns>
    Public Shared Function getSubjectHr(ByRef oConn As SqlConnection, ByVal OCID As String) As DataTable
        Call TIMS.OpenDbConn(oConn)

        Dim dt As New DataTable
        OCID = TIMS.ClearSQM(OCID)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, oConn)
        If drCC Is Nothing Then Return dt

        Dim da As New SqlDataAdapter
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WCS AS (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID = @OCID) " & vbCrLf

        sql &= " SELECT DISTINCT b.CourseName ,a.CourID ,b.CLASSIFICATION1 ,0 Hours " & vbCrLf
        sql &= " FROM STUD_TRAININGRESULTS a " & vbCrLf
        sql &= " JOIN WCS ON WCS.SOCID = a.SOCID " & vbCrLf
        sql &= " LEFT JOIN Course_CourseInfo b ON a.CourID = b.CourID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " ORDER BY b.COURSENAME " & vbCrLf
        da.SelectCommand = New SqlCommand(sql, oConn)
        With da.SelectCommand
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
        End With
        da.Fill(dt)

        'Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT Class1, Class2, Class3, Class4, Class5, Class6" & vbCrLf
        sql &= " ,Class7, Class8, Class9, Class10, Class11, Class12 " & vbCrLf
        sql &= " FROM Class_Schedule " & vbCrLf
        sql &= " WHERE OCID = @OCID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, oConn)
        Dim dt2 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            dt2.Load(.ExecuteReader())
        End With

        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                Dim iHours As Integer = 0
                For Each dr2 As DataRow In dt2.Rows
                    For i As Int16 = 1 To 12
                        If Convert.ToString(dr("CourID")) = Convert.ToString(dr2("Class" & i.ToString)) Then iHours += 1
                    Next
                Next
                dr("Hours") = iHours
            Next
        End If
        Return dt
        'sql_3 = "select   thours  from   Class_ClassInfo where   OCID='" & OCID & "'"
    End Function

    Public Shared Function GetDatatable(ByRef oConn As SqlConnection, ByVal Item As Int16, ByVal OCID As String, ByVal SOCID As String) As DataTable
        'Optional ByVal SOCID As String = ""
        ' Item  3'單位名稱變更歷史資料 7'計畫名稱 1'操行底分 4'取出所有科目 5'取出所有學員在 Stud_TrainingResults 的成績  
        '取得
        TIMS.OpenDbConn(oConn)
        Dim dt As New DataTable
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return dt
        'Item: 1~7
        'Dim RID As String = TIMS.ClearSQM(Request("RID"))
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, oConn)
        If drCC Is Nothing Then Return dt
        If OCID = "" Then Return dt
        Dim s_RID As String = Convert.ToString(drCC("RID"))
        Dim s_ORGID As String = Convert.ToString(drCC("ORGID"))
        Dim s_PlanID As String = Convert.ToString(drCC("PlanID"))
        Dim s_TPlanID As String = Convert.ToString(drCC("TPlanID"))
        Dim s_DistID As String = Convert.ToString(drCC("DistID"))
        Dim s_PLANYEARS As String = Convert.ToString(drCC("PLANYEARS"))

        Dim sChooseClass As String = Get_ChooseClass(oConn, OCID)

        Dim sql As String = ""
        Select Case Item
            Case 1
                '操行底分
                Dim sBasicPoint As String = TIMS.GetGlobalVar(s_DistID, s_TPlanID, "3", "1", oConn)
                Dim iBasicPoint As Integer = Val(sBasicPoint)
                '判斷操行分數設定
                Dim eMP As TIMS.EMinusPoint = TIMS.GetMinusPoint(oConn, OCID, s_RID, s_PlanID)
                'Dim ssChkMinusPoint As String = chkMinusPoint(oConn, OCID, s_RID, s_PlanID)
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " WITH WC1 AS ( " & vbCrLf
                sql &= " 	SELECT e1.SOCID ,SUM(e2.MinusPoint*e1.Hours) MinusPoint " & vbCrLf
                sql &= " 	FROM Class_StudentsOfClass cs " & vbCrLf
                sql &= " 	JOIN Stud_Turnout e1 ON e1.socid = cs.socid AND cs.OCID = " & OCID & "" & vbCrLf
                If SOCID <> "" Then sql &= " AND cs.SOCID IN (" & SOCID & ") " & vbCrLf
                Select Case eMP' ssChkMinusPoint
                    Case TIMS.EMinusPoint.mpKey ' "key" '自訂
                        sql &= " JOIN Key_Leave e2 ON e1.LeaveID = e2.LeaveID " & vbCrLf
                    Case TIMS.EMinusPoint.mpOrg '"org"
                        sql &= " JOIN Org_Leave e2 ON e1.LeaveID = e2.LeaveID " & vbCrLf
                        sql &= "  AND e2.PlanID = '" & s_PlanID & "' " & vbCrLf
                        sql &= "  AND e2.OrgID = '" & s_ORGID & "' " & vbCrLf
                    Case TIMS.EMinusPoint.mpClass '"class"
                        sql &= " JOIN Class_Leave e2 ON e1.LeaveID = e2.LeaveID " & vbCrLf
                        sql &= "  And e2.OCID = " & OCID & " " & vbCrLf
                End Select
                sql &= " 	GROUP BY e1.SOCID ) " & vbCrLf

                sql &= " ,WC2 AS ( " & vbCrLf
                sql &= " 	SELECT f1.SOCID ,SUM(f1.Times*(CASE WHEN f2.AddMinus = '+' THEN f2.Point ELSE 0 - f2.Point END)) total " & vbCrLf
                sql &= " 	FROM Class_StudentsOfClass cs " & vbCrLf
                sql &= " 	JOIN Stud_Sanction f1 ON f1.socid = cs.socid AND cs.OCID = " & OCID & " " & vbCrLf
                If SOCID <> "" Then sql &= " AND cs.SOCID IN (" & SOCID & ") " & vbCrLf
                sql &= " 	JOIN Key_Sanction f2 ON f1.SanID = f2.SanID " & vbCrLf
                sql &= " 	GROUP BY f1.SOCID ) " & vbCrLf

                sql &= " SELECT cc.OCID, cs.SOCID, cs.StudentID, ss.Name, cs.StudStatus" & vbCrLf
                sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName" & vbCrLf
                sql &= " ,cc.STDate,cc.FTDate" & vbCrLf
                sql &= " ,(80- ISNULL(e.MinusPoint,0)+ISNULL(f.total,0) + ISNULL(FLOOR(a.TechPoint),0) + ISNULL(FLOOR(a.RemedPoint),0)) conductpoint" & vbCrLf
                sql &= " FROM Class_ClassInfo cc" & vbCrLf
                sql &= " JOIN Class_StudentsOfClass cs ON cs.ocid = cc.ocid" & vbCrLf
                sql &= " JOIN Stud_StudentInfo ss ON ss.SID = cs.SID" & vbCrLf
                sql &= " LEFT JOIN Stud_Conduct a ON a.SOCID = cs.SOCID" & vbCrLf
                sql &= " LEFT JOIN WC1 e ON e.socid = cs.socid" & vbCrLf
                sql &= " LEFT JOIN WC2 f ON f.socid = cs.socid" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= " AND cc.OCID = " & OCID & " " & vbCrLf
                If SOCID <> "" Then sql &= " AND cs.SOCID IN (" & SOCID & ") " & vbCrLf

                'TIMS.writeLog(Me, sql)
            Case 2
                '單位名稱變更歷史資料
                sql = "" & vbCrLf
                sql &= " SELECT o2.ORGNAME " & vbCrLf
                sql &= " FROM CLASS_CLASSINFO c1 " & vbCrLf
                sql &= " JOIN ORG_ORGINFO o1 ON c1.comidno = o1.comidno " & vbCrLf
                sql &= " LEFT JOIN ORG_ORGNAMEHISTORY o2 ON o1.orgid = o2.orgid " & vbCrLf
                sql &= " LEFT JOIN ID_PLAN d ON c1.PlanID = d.PlanID " & vbCrLf
                sql &= " LEFT JOIN KEY_PLAN e ON d.TPlanID = e.TPlanID " & vbCrLf
                sql &= " WHERE c1.ocid = '" & OCID & "' " & vbCrLf
                sql &= " AND o2.years = d.years " & vbCrLf
            Case 3
                '單位名稱變更歷史資料
                'If sm.UserInfo.Years < Now.Year Then
                If Val(s_PLANYEARS) < Now.Year Then
                    sql = "" & vbCrLf
                    sql &= " SELECT o2.ORGNAME " & vbCrLf
                    sql &= " FROM CLASS_CLASSINFO c1 " & vbCrLf
                    sql &= " JOIN ORG_ORGINFO o1 ON c1.comidno = o1.comidno " & vbCrLf
                    sql &= " LEFT JOIN ORG_ORGNAMEHISTORY o2 ON o1.orgid = o2.orgid " & vbCrLf
                    sql &= " LEFT JOIN ID_PLAN d ON c1.PlanID = d.PlanID AND o2.years = d.years " & vbCrLf
                    sql &= " LEFT JOIN KEY_PLAN e ON d.TPlanID = e.TPlanID " & vbCrLf
                    sql &= " WHERE c1.ocid = '" & OCID & "' " & vbCrLf
                    sql &= " ORDER BY o2.editid DESC " & vbCrLf
                Else
                    '登入年度為本年度則直接抓 org_orginfo 內的orgname
                    sql = ""
                    sql &= " SELECT c1.OCID, o1.ORGID, o1.ORGNAME " & vbCrLf
                    sql &= " FROM CLASS_CLASSINFO c1 " & vbCrLf
                    sql &= " JOIN ORG_ORGINFO o1 ON c1.comidno = o1.comidno " & vbCrLf
                    sql &= " WHERE c1.ocid = '" & OCID & "' " & vbCrLf
                End If
            Case 4
                '取出所有科目
                sql = ""
                sql &= " SELECT Distinct b.CourseName, b.CourseID, a.CourID " & vbCrLf
                sql &= " FROM Stud_TrainingResults a " & vbCrLf
                sql &= " LEFT JOIN Course_CourseInfo b ON a.CourID = b.CourID " & vbCrLf
                sql &= " WHERE a.SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID = '" & OCID & "') " & vbCrLf
                sql &= " ORDER BY b.courseName " & vbCrLf
            Case 5
                '取出所有學員在 Stud_TrainingResults 的成績  
                sql = "" & vbCrLf
                sql &= " SELECT SOCID, COURID,RESULTS " & vbCrLf
                sql &= " FROM Stud_TrainingResults " & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                If sChooseClass <> "" Then
                    sql &= " AND CourID IN (" & sChooseClass & ") " & vbCrLf
                Else
                    sql &= " AND 1<>1 " & vbCrLf
                End If
                sql &= " AND socid IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID = '" & OCID & "') " & vbCrLf
            Case 6
                sql = " SELECT SOCID ,OCID, RANK , STUDSTATUS FROM CLASS_STUDENTSOFCLASS WHERE OCID = '" & OCID & "' " & vbCrLf
            Case 7
                '計畫名稱
                sql = "" & vbCrLf
                sql &= " SELECT K1.PLANNAME PLANNAME " & vbCrLf
                sql &= " FROM CLASS_CLASSINFO c1 " & vbCrLf
                sql &= " JOIN ID_PLAN I1 ON C1.PLANID = I1.PLANID " & vbCrLf
                sql &= " JOIN KEY_PLAN K1 ON I1.TPLANID = K1.TPLANID " & vbCrLf
                sql &= " WHERE c1.ocid = '" & OCID & "' " & vbCrLf
        End Select

        Dim sCmd As New SqlCommand(sql, oConn)
        'Call TIMS.OpenDbConn(oConn)
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With

        Return dt
    End Function

#Region "NO USE"
    '判斷操行分數設定
    'Public Shared Function chkMinusPoint(ByRef oConn As SqlConnection, ByVal strOCID As String, ByVal strRID As String, ByVal strPlanID As String) As String
    '    Dim strRtn As String = ""
    '    'Dim sm As SessionModel = SessionModel.Instance()
    '    Call TIMS.OpenDbConn(oConn)
    '    Dim sql As String = ""
    '    sql = ""
    '    sql &= " SELECT leaveid, minuspoint FROM org_leave WHERE planid = @planid AND orgid IN (SELECT orgid FROM auth_relship WHERE rid = @rid) "
    '    Dim sCmd As New SqlCommand(sql, oConn)
    '    Dim dtOO As New DataTable
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("planid", SqlDbType.VarChar).Value = strPlanID 'sm.UserInfo.PlanID
    '        .Parameters.Add("rid", SqlDbType.VarChar).Value = strRID
    '        dtOO.Load(.ExecuteReader())
    '    End With
    '    sql = " SELECT leaveid, minuspoint FROM class_leave WHERE ocid = @ocid "
    '    Dim sCmd2 As New SqlCommand(sql, oConn)
    '    Dim dtCC As New DataTable
    '    With sCmd2
    '        .Parameters.Clear()
    '        .Parameters.Add("ocid", SqlDbType.VarChar).Value = strOCID
    '        dtCC.Load(.ExecuteReader())
    '    End With
    '    If dtCC.Rows.Count > 0 Then
    '        strRtn = "class"
    '    ElseIf dtOO.Rows.Count > 0 Then
    '        strRtn = "org"
    '    Else
    '        strRtn = "key"
    '    End If
    '    Return strRtn
    'End Function
#End Region

    Public Shared Function Get_COURSEINFO(ByRef oConn As SqlConnection, ByRef OCID As String) As DataTable
        TIMS.OpenDbConn(oConn)
        Dim sChooseClass As String = Get_ChooseClass(oConn, OCID)
        Dim sql3 As String = ""
        sql3 &= " SELECT CourID,CourseName" & vbCrLf  '先取出目前所選取的科目  
        sql3 &= "  ,CASE Classification1 WHEN 1 THEN '學科' WHEN 2 THEN '術科' END classType ,Classification1 " & vbCrLf
        sql3 &= " ,0 AS TotalHours " & vbCrLf
        sql3 &= " FROM Course_CourseInfo " & vbCrLf
        sql3 &= " WHERE 1=1 " & vbCrLf
        If sChooseClass <> "" Then
            sql3 &= " AND CourID IN (" & sChooseClass & ") " & vbCrLf
        Else
            sql3 &= " AND 1<>1 " & vbCrLf
        End If
        sql3 &= " ORDER BY CourseName "
        Dim sCmd3 As New SqlCommand(sql3, oConn)
        Dim dt3 As New DataTable
        With sCmd3
            .Parameters.Clear()
            dt3.Load(.ExecuteReader())
        End With
        Return dt3
    End Function

    Public Shared Function Get_ChooseClass(ByRef oConn As SqlConnection, ByRef OCID As String) As String  '取得所有科目
        Dim rst As String = ""
        If OCID = "" Then Return rst
        Dim iOCID As Integer = Val(OCID)

        Call TIMS.OpenDbConn(oConn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT DISTINCT CourID FROM Stud_TrainingResults WHERE SOCID IN (SELECT SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID)" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT b.CourseName, a.CourID" & vbCrLf
        sql &= " from WC1 a" & vbCrLf
        sql &= " LEFT JOIN Course_CourseInfo b ON b.CourID =a.CourID" & vbCrLf

        Dim sCmd As New SqlCommand(sql, oConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = iOCID
            dt.Load(.ExecuteReader())
        End With
        rst = ""
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If rst <> "" Then rst &= ","
                rst &= "'" & dr("CourID") & "'"
            Next
        End If
        Return rst
    End Function

    Public Shared Function getStudResult(ByRef oConn As SqlConnection, ByVal itemTyp As Int16, ByVal str_OCID As String, ByVal OneSocid As String) As DataTable '090825
        'Optional ByVal OneSocid As String = ""
        '成績計算方式 iResultTyp 'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )  
        'itemTyp :權重法-取出參數
        Call TIMS.OpenDbConn(oConn)
        Dim sql_1 As String = ""
        Dim sql_1_2 As String = ""
        Dim sql_2 As String = ""
        Dim sql_3 As String = ""
        Dim sql_A As String = ""
        Dim sql_B As String = ""
        Dim sql_C As String = ""
        Dim ds As New DataSet
        Dim dt1 As New DataTable
        Dim dt1_2 As New DataTable
        Dim dt2 As New DataTable
        Dim dt3 As New DataTable
        Dim dtReult1 As New DataTable
        Dim dtReult2 As New DataTable
        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim dr2 As DataRow
        Dim dr3 As DataRow
        Dim dr4 As DataRow
        Dim da As New SqlDataAdapter
        Dim sChooseClass As String = Get_ChooseClass(oConn, str_OCID)

        '先取出目前所選取的科目  
        sql_1 = ""
        sql_1 &= " SELECT  CourID,CourseName"
        sql_1 &= " ,CASE Classification1 WHEN 1 THEN '學科' WHEN 2 THEN '術科' END AS classType ,Classification1 ,0 AS TotalHours " & vbCrLf
        sql_1 &= " FROM Course_CourseInfo " & vbCrLf
        sql_1 &= " WHERE 1=1 " & vbCrLf
        If sChooseClass <> "" Then
            sql_1 &= " AND CourID IN (" & sChooseClass & ") " & vbCrLf
        Else
            sql_1 &= " AND 1<>1 " & vbCrLf
        End If
        sql_1 &= " ORDER BY CourseName "

        '檢查學員成績的有登錄成績的科目總數    
        sql_1_2 = ""
        sql_1_2 &= " SELECT COURID FROM Stud_TrainingResults " & vbCrLf
        sql_1_2 += " WHERE 1=1 " & vbCrLf
        If sChooseClass <> "" Then
            sql_1_2 += " AND CourID IN (" & sChooseClass & ") " & vbCrLf
        Else
            sql_1_2 += " AND 1<>1 " & vbCrLf
        End If
        sql_1_2 += " AND SOCID in ( " & OneSocid & "  ) "
        sql_1_2 += " GROUP BY courid "
        sql_1_2 += " ORDER BY courid "

        '取出所有學員成績    OCID
        sql_2 = ""
        sql_2 &= " SELECT '" & str_OCID & "' OCID ,SOCID ,COURID ,'' COURSENAME ,RESULTS ,'' CLASSTYPE ,NULL CLASSIFICATION1 ,0 HOURS ,0 RESULTS2 " & vbCrLf
        sql_2 += " FROM Stud_TrainingResults " & vbCrLf
        sql_2 += " WHERE 1=1 " & vbCrLf
        If sChooseClass <> "" Then
            sql_2 += " AND CourID IN (" & sChooseClass & ") " & vbCrLf
        Else
            sql_2 += " AND 1<>1 " & vbCrLf
        End If
        If OneSocid <> "" Then
            sql_2 += " AND SOCID IN (" & OneSocid & ") " & vbCrLf
        Else
            sql_2 += " AND SOCID IN (" & OneSocid & ") " & vbCrLf
        End If
        sql_2 += " ORDER BY SOCID, CourseName " & vbCrLf

        '取出排課檔  
        sql_3 = " SELECT * FROM Class_Schedule WHERE OCID =@OCID" & vbCrLf

        '產生一個空的table"
        sql_A = ""
        sql_A &= " SELECT CourID ,CourseName ,CASE Classification1 WHEN 1 THEN '學科' WHEN 2 THEN '術科' END classType ,Classification1 ,0 TotalHours " & vbCrLf
        sql_A += " FROM Course_CourseInfo " & vbCrLf
        sql_A += " WHERE 1<>1 "

        '產生一個空的table
        str_OCID = TIMS.ClearSQM(str_OCID)
        sql_B = ""
        sql_B = String.Concat(" SELECT '", TIMS.ClearSQM(str_OCID), "' OCID ,SOCID ,COURID ,'' COURSENAME ,RESULTS ,'' CLASSTYPE ,NULL CLASSIFICATION1 ,0 AS HOURS ,0 AS RESULTS2") & vbCrLf
        sql_B &= " FROM STUD_TRAININGRESULTS" & vbCrLf
        sql_B &= " WHERE 1<>1"

        With da
            .SelectCommand = New SqlCommand(sql_1, oConn)
            .Fill(dt1)
            .SelectCommand = New SqlCommand(sql_1_2, oConn)
            .Fill(dt1_2)
            .SelectCommand = New SqlCommand(sql_2, oConn)
            .Fill(dt2)

            .SelectCommand = New SqlCommand(sql_3, oConn)
            .SelectCommand.Parameters.Clear()
            .SelectCommand.Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(str_OCID)
            .Fill(dt3)

            .SelectCommand = New SqlCommand(sql_A, oConn)
            .SelectCommand.Parameters.Clear()
            .Fill(dtReult1)

            .SelectCommand = New SqlCommand(sql_B, oConn)
            .Fill(dtReult2)
        End With

        Dim CountHours As Int16 = 0
        For Each dr1 In dt1.Rows       '列出所有選取的科目
            For Each dr3 In dt3.Rows   '依序列出排課
                For i As Integer = 1 To 12
                    If Not IsDBNull(dr3("Class" & i)) Then
                        If dr3("Class" & i) = dr1("CourID") Then dr1("TotalHours") += 1
                    End If
                Next
            Next
            dr = dtReult1.NewRow      '填入空的table
            dr("CourID") = dr1("CourID")
            dr("CourseName") = dr1("CourseName")
            dr("classType") = dr1("classType")
            dr("Classification1") = dr1("Classification1")
            dr("TotalHours") = dr1("TotalHours")
            dtReult1.Rows.Add(dr)
        Next
        '列出所有選取的學員 
        For Each dr2 In dt2.Rows
            For Each dr4 In dtReult1.Rows
                If Convert.ToString(dr2("COURID")) = Convert.ToString(dr4("COURID")) Then
                    dr2("COURID") = dr4("COURID")
                    dr2("COURSENAME") = dr4("COURSENAME")
                    dr2("CLASSTYPE") = dr4("CLASSTYPE")
                    dr2("CLASSIFICATION1") = dr4("CLASSIFICATION1")
                    dr2("HOURS") = dr4("TotalHours")
                    dr2("RESULTS2") = dr2("RESULTS") * dr4("TotalHours")

                    dr = dtReult2.NewRow
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
        '成績計算方式 iResultTyp 'ResultTyp 成績計算方式(1,各科平均法 / 2,訓練時數權重法 )
        Select Case itemTyp
            Case 1
                '所選取的科目大於一個且 學員成績檔中只有登錄一個科目成績時 -->以「各科平均法」計算
                'If dt1.Rows.Count >= 1 And dt1_2.Rows.Count = 1 Then Me.ViewState("ResultTyp") = "1" End If
                Return dt1
            Case 2
                Return dtReult2
        End Select
        Return dt1

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me.Page, "發生錯誤：" & ex.Message.ToString())
        '    Throw ex '
        'Finally
        '    da.Dispose()
        '    dt1.Dispose()
        '    dt2.Dispose()
        '    dt3.Dispose()
        '    dtReult1.Dispose()
        '    dtReult2.Dispose()
        '    'conn.Close()
        'End Try
    End Function

    ''' <summary> 建立成績檔案 </summary>
    ''' <param name="oConn"></param>
    ''' <param name="sHtb"></param>
    ''' <returns></returns>
    Public Shared Function CreateGradeTable(ByRef oConn As SqlConnection, ByRef sHtb As Hashtable) As DataTable
        ' ByVal TPlanID As String, ByVal DistID As String, ByVal OCIDValue1 As String, ByVal SOCID As String
        Dim TPlanID As String = TIMS.GetMyValue2(sHtb, "TPlanID")
        Dim DistID As String = TIMS.GetMyValue2(sHtb, "DistID")
        Dim OCIDValue1 As String = TIMS.GetMyValue2(sHtb, "OCIDValue1")
        Dim SOCID As String = TIMS.GetMyValue2(sHtb, "SOCID")
        'Dim CheckMode As String = TIMS.GetMyValue2(sHtb, "CheckMode") 'CheckMode = dt.Select("GVID='13'")(0)("ItemVar1")
        '成績計算方式 iResultTyp 'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
        Dim s_ResultTyp As String = TIMS.GetMyValue2(sHtb, "ResultTyp") ' GetPara(13, DistID, TPlanID, 1)   '成績計算方式
        Dim s_percent1 As String = TIMS.GetMyValue2(sHtb, "percent1") ' GetPara(17, DistID, TPlanID, 1)     '學科百分比
        Dim s_percent2 As String = TIMS.GetMyValue2(sHtb, "percent2") ' GetPara(17, DistID, TPlanID, 2)     '術科百分比
        Dim s_pClassCount As String = TIMS.GetMyValue2(sHtb, "pClassCount") '學科科目總數
        Dim s_sClassCount As String = TIMS.GetMyValue2(sHtb, "sClassCount") '術科科目總數
        '成績計算方式 iResultTyp 'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
        Dim iResultTyp As Integer = If(s_ResultTyp <> "", Val(s_ResultTyp), 0) ' GetPara(13, DistID, TPlanID, 1) '成績計算方式
        Dim ipercent1 As Double = If(s_percent1 <> "", Val(s_percent1), 0) 'TIMS.GetMyValue2(sHtb, "GP17_1") ' GetPara(17, DistID, TPlanID, 1)     '學科百分比
        Dim ipercent2 As Double = If(s_percent2 <> "", Val(s_percent2), 0) 'TIMS.GetMyValue2(sHtb, "GP17_2") ' GetPara(17, DistID, TPlanID, 2)     '術科百分比
        Dim i_pClassCount As Integer = If(s_pClassCount <> "", Val(s_pClassCount), 0) '學科科目總數
        Dim i_sClassCount As Integer = If(s_sClassCount <> "", Val(s_sClassCount), 0) '術科科目總數

        Dim dtReult As New DataTable
        dtReult.Columns.Add(New DataColumn("SOCID"))
        dtReult.Columns.Add(New DataColumn("pClassCount"))
        dtReult.Columns.Add(New DataColumn("sClassCount"))
        dtReult.Columns.Add(New DataColumn("pTotal_1"))
        dtReult.Columns.Add(New DataColumn("sTotal_1"))
        dtReult.Columns.Add(New DataColumn("pTotal_2"))
        dtReult.Columns.Add(New DataColumn("sTotal_2"))
        dtReult.Columns.Add(New DataColumn("pHours"))
        dtReult.Columns.Add(New DataColumn("sHours"))
        dtReult.Columns.Add(New DataColumn("totalResults"))
        dtReult.Columns.Add(New DataColumn("totalResults2"))
        dtReult.Columns.Add(New DataColumn("totalAvg"))
        dtReult.Columns.Add(New DataColumn("pAvg"))
        dtReult.Columns.Add(New DataColumn("sAvg"))
        dtReult.Columns.Add(New DataColumn("ResultTyp"))
        dtReult.Columns.Add(New DataColumn("percent1"))
        dtReult.Columns.Add(New DataColumn("percent2"))

        Dim Errmsg As String = ""
        'Dim i As Integer
        'Dim sql As String
        'Dim sql3 As String = ""
        'Dim dt As New DataTable
        'Dim dt2 As New DataTable
        'Dim dt3 As New DataTable
        'Dim da As New SqlDataAdapter
        'Dim pClassCount As Integer = 0  '學科科目總數
        'Dim sClassCount As Integer = 0  '術科科目總數

        'sql = " SELECT * FROM Sys_GlobalVar WHERE DistID = '" & DistID & "' AND TPlanID = '" & TPlanID & "' "
        'Dim sChooseClass As String = Get_ChooseClass(oConn, OCIDValue1)
        'Dim sChooseClass As String = Get_ChooseClass(OCIDValue1)
        'lbmsg.Text = "查無資料"
        'If sChooseClass = "" Then Return dtReult

        'Call TIMS.OpenDbConn(oConn)
        'sql3 = ""
        'sql3 &= " SELECT CourID,CourseName" & vbCrLf  '先取出目前所選取的科目  
        'sql3 &= "  ,CASE Classification1 WHEN 1 THEN '學科' WHEN 2 THEN '術科' END classType ,Classification1 " & vbCrLf
        'sql3 &= " FROM Course_CourseInfo " & vbCrLf
        'sql3 &= " WHERE 1=1 " & vbCrLf
        'If sChooseClass <> "" Then
        '    sql3 &= " AND CourID IN (" & sChooseClass & ") " & vbCrLf
        'Else
        '    sql3 &= " AND 1<>1 " & vbCrLf
        'End If
        'sql3 &= " ORDER BY CourseName "

        'With da
        '    .SelectCommand = New SqlCommand(Sql, oConn)
        '    .Fill(dt)
        '    .SelectCommand = New SqlCommand(sql3, oConn)
        '    .Fill(dt3)
        'End With

        'If dt.Select("GVID='13'").Length = 0 Then
        '    'Errmsg += "尚未設定成績計算模式,請聯絡中心系統管理者" & vbCrLf
        '    Errmsg += "尚未設定成績計算模式,請聯絡分署系統管理者" & vbCrLf
        '    'Exit Sub
        'Else
        '    CheckMode = dt.Select("GVID='13'")(0)("ItemVar1")
        'End If
        'If dt3.Rows.Count > 0 Then
        '    'Dim n As Int16 = 0
        '    For n As Int16 = 0 To dt3.Rows.Count - 1
        '        If Convert.ToString(dt3.Rows(n)("Classification1")) = "1" Then
        '            i_pClassCount = pClassCount + 1
        '        ElseIf Convert.ToString(dt3.Rows(n)("Classification1")) = "2" Then
        '            sClassCount = sClassCount + 1
        '        End If
        '    Next
        'End If

        'conn = DbAccess.GetConnection
        'conn.Open()
        'Dim dt2 As DataTable = getStudResult(oConn, 2, OCIDValue1, SOCID)
        '成績計算方式 iResultTyp 'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
        'Dim itemTyp As Integer = If(s_ResultTyp <> "", Val(s_ResultTyp), 2) 'def:2,訓練時數權重法
        Dim dt2 As DataTable = getStudResult(oConn, 2, OCIDValue1, SOCID)
        '----------------------------------------------------------------  
        '***************************************
        ' 【計算成績方式】
        '        <各科平均法>
        '=======================================
        '  (1)學、術科百分比(無)  (2) 學、術科百分比(有)
        '    ex: 總平均是由(學科全部成績加總/學科總數 4  科) * 學科百分比 20 %      
        '        加上      (術科全部成績加總/術科總數 11 科) * 術科百分比 80 %      
        '        計算產生之結果。
        '---------------------------------------
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

        '---------------------------------------
        '  (2) 學、術科百分比(有)
        '    ex: 學科平均=(科)成績*(科)總時數 /(各科加總)總時數  
        '        總平均是由(學科成績與訓練時數加權/學科時數 __ 小時) * 學科百分比 20 %      
        '        加上      (術科成績與訓練時數加權/術科時數 __ 小時) * 術科百分比 20 %        
        '        計算產生之結果。 
        '***************************************
        'Dim j As Integer = 0
        'Dim pClassCount As Integer = 0              '學科科目總數
        'Dim sClassCount As Integer = 0              '術科科目總數
        Dim pTotal_1 As Decimal = 0        '學科總成績---各科平均法
        Dim sTotal_1 As Decimal = 0        '術科總成績 
        Dim pTotal_2 As Decimal = 0        '學科總成績---訓練時數權重法
        Dim sTotal_2 As Decimal = 0        '術科總成績
        Dim pHours As Decimal = 0          '學科總時數
        Dim sHours As Decimal = 0          '術科總時數
        Dim totalResults As Integer = 0    '總成績 
        Dim totalResults2 As Integer = 0   '總成績(時數加權) 
        Dim totalAvg As Decimal = 0        '總平均 
        Dim pAvg As Decimal = 0            '學科總平均 
        Dim sAvg As Decimal = 0            '術科總平均 
        'Dim ResultTyp As Integer = GetPara(13, DistID, TPlanID, 1)      '成績計算方式
        'Dim percent1 As Double = GetPara(17, DistID, TPlanID, 1)     '學科百分比
        'Dim percent2 As Double = GetPara(17, DistID, TPlanID, 2)     '術科百分比

        If dt2.Rows.Count > 0 Then
            For j As Integer = 0 To dt2.Rows.Count - 1
                If Convert.ToString(dt2.Rows(j)("Classification1")) = "1" Then        '學科
                    'pClassCount = pClassCount + 1
                    pHours = pHours + Convert.ToInt16(dt2.Rows(j)("hours")) '時數
                    pTotal_2 = pTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                    pTotal_1 = pTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))   '各科平均法： 1 '預設採用各科平均法(當未設定時)
                ElseIf Convert.ToString(dt2.Rows(j)("Classification1")) = "2" Then    '術科
                    'sClassCount = sClassCount + 1
                    sHours = sHours + Convert.ToInt16(dt2.Rows(j)("hours")) '時數
                    sTotal_2 = sTotal_2 + Convert.ToDecimal(dt2.Rows(j)("Results2"))  '訓練時數權重法： 2
                    sTotal_1 = sTotal_1 + Convert.ToDecimal(dt2.Rows(j)("Results"))
                End If
            Next

            pAvg = If(i_pClassCount = 0, 0, pTotal_1 / i_pClassCount)

            sAvg = If(i_sClassCount = 0, 0, sTotal_1 / i_sClassCount)

            totalResults = sTotal_1 + pTotal_1  '原始成績加總

            totalResults2 = sTotal_2 + pTotal_2 '成績加總(訓練時數權重法)

            '成績計算方式 iResultTyp 'ResultTyp 成績計算方式( 1, 各科平均法  2, 訓練時數權重法 )
            Select Case iResultTyp
                Case 2
                    '訓練時數權重法： 2
                    '總平均= (學科成績與時數加權/時數加總 36 小時)* 學科百分比 40% + (術科成績與時數加權/時數加總 88 小時)* 術科百分比 60 %        
                    '/2 計算之結果。 
                    If ipercent1 = 0 And ipercent2 = 0 Then  '百分比未設定時
                        If i_pClassCount = 0 And i_sClassCount <> 0 And sHours <> 0 Then      '只有術科
                            totalAvg = (sTotal_2 / sHours)
                        ElseIf i_pClassCount <> 0 And i_sClassCount = 0 And pHours <> 0 Then  '只有學科
                            totalAvg = (pTotal_2 / pHours)
                        ElseIf i_pClassCount <> 0 And i_sClassCount <> 0 Then
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
                        If i_pClassCount = 0 And i_sClassCount <> 0 And sHours <> 0 Then      '只有術科
                            totalAvg = (sTotal_2 / sHours) * ipercent2
                        ElseIf i_pClassCount <> 0 And i_sClassCount = 0 And pHours <> 0 Then  '只有學科
                            totalAvg = (pTotal_2 / pHours) * ipercent1
                        ElseIf i_pClassCount <> 0 And i_sClassCount <> 0 Then
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

                Case Else
                    '各科平均法： 1 '預設採用各科平均法(當未設定時)
                    If ipercent1 = 0 AndAlso ipercent2 = 0 Then  '百分比未設定時
                        totalAvg = totalResults / (i_pClassCount + i_sClassCount)
                    Else
                        If i_pClassCount = 0 And i_sClassCount <> 0 Then      '只有術科
                            totalAvg = (sTotal_1 / i_sClassCount) * ipercent2
                        ElseIf i_pClassCount <> 0 And i_sClassCount = 0 Then  '只有學科
                            totalAvg = (pTotal_1 / i_pClassCount) * ipercent1
                        ElseIf i_pClassCount <> 0 And i_sClassCount <> 0 Then
                            totalAvg = (pTotal_1 / i_pClassCount) * ipercent1 + (sTotal_1 / i_sClassCount) * ipercent2
                        Else
                            totalAvg = 0
                        End If
                    End If
            End Select
            'lbmsg.Text = ""

            'Dim dtReult As New DataTable
            Dim dr As DataRow
            dr = dtReult.NewRow
            dtReult.Rows.Add(dr)
            dr("SOCID") = SOCID
            dr("pClassCount") = i_pClassCount    '學科科目總數
            dr("sClassCount") = i_sClassCount    '術科科目總數
            dr("pTotal_1") = pTotal_1            '學科總成績---各科平均法
            dr("sTotal_1") = sTotal_1            '術科總成績 
            dr("pTotal_2") = pTotal_2            '學科總成績---訓練時數權重法 
            dr("sTotal_2") = sTotal_2            '術科總成績 
            dr("pHours") = pHours                '學科總時數
            dr("sHours") = sHours                '術科總時數
            dr("totalResults") = totalResults    '總成績 
            dr("totalResults2") = totalResults2  '總成績(時數加權) 
            dr("totalAvg") = TIMS.ROUND(totalAvg, 1)     '總平均 
            dr("pAvg") = TIMS.ROUND(pAvg, 1)             '學科總平均 
            dr("sAvg") = TIMS.ROUND(sAvg, 1)             '術科總平均 
            dr("ResultTyp") = iResultTyp   '成績計算方式 iResultTyp   
            dr("percent1") = ipercent1     '學科百分比
            dr("percent2") = ipercent2     '術科百分比
        End If

        Return dtReult

#Region "(No Use)"

        'If sChooseClass <> "" Then
        '    Try

        '    Catch ex As Exception
        '        Common.MessageBox(Me.Page, "發生錯誤!! " & ex.Message.ToString())
        '        Throw ex '
        '    Finally
        '        'conn.Close()
        '        dt2.Dispose()
        '    End Try
        'Else
        '    lbmsg.Text = "查無資料"
        'End If

#End Region
        '建立資料---------------   End
    End Function

    Private Sub CreateRpt(ByVal PageNum As Int16)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim rptTb As HtmlTable
        Dim rptRow As HtmlTableRow
        Dim rptCell As HtmlTableCell

        Dim OCID As String = Hid_OCID.Value 'Request("OCID")
        Dim SelSOCID As String = Request("SOCID")
        Dim BasicPoint As Integer = 0
        BasicPoint = GetPara(objconn, 3, sm.UserInfo.DistID, sm.UserInfo.TPlanID, 1)

        '1'操行底分-SelSOCID
        Dim dt1 As DataTable = GetDatatable(objconn, 1, OCID, SelSOCID)
        'dt2 = GetDatatable(2, OCID) 'dt3 = GetDatatable(3, OCID)
        '4'取出所有科目 
        Dim dt4 As DataTable = GetDatatable(objconn, 4, OCID, "")
        If dt4 Is Nothing Then
            Common.MessageBox(Me.Page, "查無資料，目前學員結訓成績尚未登錄！")
            Return ' Exit Sub
        End If

        '5'取出所有學員在 Stud_TrainingResults 的成績  
        Dim dt5 As DataTable = GetDatatable(objconn, 5, OCID, "")
        '取得課程訓練時數資料表
        Dim dt7 As DataTable = getSubjectHr(objconn, OCID)
        '取得學、術科、總時數
        Dim dt8 As DataTable = getHour(objconn, OCID)

        'Dim SOCID As String = ""
        'Dim dr1 As DataRow
        'Dim da As New SqlDataAdapter

        Dim i As Int16 = 0
        Dim j As Int16 = 0
        Dim recordNote As String = ""

        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
        rptTb.Attributes.Add("align", "left")
        rptTb.Attributes.Add("border", 0)
        rptTb.Attributes.Add("cellpadding", "0")
        rptTb.Attributes.Add("cellspacing", "0")
        rptTb.Attributes.Add("bordercolor", "#FFFFFF")
        print_content.Controls.Add(rptTb)

        '報表列印格式 1:空白 2:有帶資料
        'Dim PrintDataTyp As String = Request("PrintDataTyp")  

        ' 開始
        lbmsg.Text = "查無資料"
        If dt1.Rows.Count > 0 Then lbmsg.Text = ""

        Dim ResultTyp As Integer = GetPara(objconn, 13, sm.UserInfo.DistID, sm.UserInfo.TPlanID, 1)    '成績計算方式 ( 1, 各科平均法  2, 訓練時數權重法 )
        Dim percent1 As Double = GetPara(objconn, 17, sm.UserInfo.DistID, sm.UserInfo.TPlanID, 1)      '學科百分比
        Dim percent2 As Double = GetPara(objconn, 17, sm.UserInfo.DistID, sm.UserInfo.TPlanID, 2)      '術科百分比
        'Dim sGVID13 As String = TIMS.GetGlobalVar(sm.UserInfo.DistID, sm.UserInfo.TPlanID, "13", "1", objconn) '成績計算方式
        'Dim sGVID17_1 As String = TIMS.GetGlobalVar(sm.UserInfo.DistID, sm.UserInfo.TPlanID, "17", "1", objconn) '學科百分比
        'Dim sGVID17_2 As String = TIMS.GetGlobalVar(sm.UserInfo.DistID, sm.UserInfo.TPlanID, "17", "2", objconn) '術科百分比
        Dim pClassCount As Integer = 0  '學科科目總數
        Dim sClassCount As Integer = 0  '術科科目總數
        Dim dt3 As DataTable = Get_COURSEINFO(objconn, OCID)
        If dt3.Rows.Count > 0 Then
            'Dim n As Int16 = 0
            For n As Int16 = 0 To dt3.Rows.Count - 1
                If Convert.ToString(dt3.Rows(n)("Classification1")) = "1" Then
                    pClassCount = pClassCount + 1
                ElseIf Convert.ToString(dt3.Rows(n)("Classification1")) = "2" Then
                    sClassCount = sClassCount + 1
                End If
            Next
        End If

        For n As Int16 = 0 To dt1.Rows.Count - 1
            If dt1.Rows.Count > 2 Then
                WaterPage.Value = Math.Ceiling(dt4.Rows.Count / 40) * dt1.Rows.Count + 2 '浮水印頁數 (每頁最多筆數40筆)
            Else
                WaterPage.Value = dt1.Rows.Count
            End If
            rptTb = Nothing
            Dim SOCID As String = Convert.ToString(dt1.Rows(n)("socid"))

            Dim sHTB6 As New Hashtable
            TIMS.SetMyValue2(sHTB6, "TPlanID", sm.UserInfo.TPlanID)
            TIMS.SetMyValue2(sHTB6, "DistID", sm.UserInfo.DistID)
            TIMS.SetMyValue2(sHTB6, "OCIDValue1", OCID)
            TIMS.SetMyValue2(sHTB6, "SOCID", SOCID)
            'TIMS.SetMyValue2(sHTB6, "CheckMode", sGVID13)
            TIMS.SetMyValue2(sHTB6, "ResultTyp", ResultTyp) '成績計算方式
            TIMS.SetMyValue2(sHTB6, "percent1", percent1) '學科百分比
            TIMS.SetMyValue2(sHTB6, "percent2", percent2) '術科百分比
            TIMS.SetMyValue2(sHTB6, "pClassCount", pClassCount) '學科科目總數
            TIMS.SetMyValue2(sHTB6, "sClassCount", sClassCount) '術科科目總數
            '建立成績檔案
            Dim dt6 As DataTable = CreateGradeTable(objconn, sHTB6)

            '----------------   表首  start --------------------------
            rptTb = New HtmlTable
            rptTb.Attributes.Add("style", "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            rptTb.Attributes.Add("align", "left")
            rptTb.Attributes.Add("border", 0)
            rptTb.Attributes.Add("cellpadding", "0")
            rptTb.Attributes.Add("cellspacing", "0")
            rptTb.Attributes.Add("bordercolor", "#FFFFFF")
            print_content.Controls.Add(rptTb)
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("colspan", 5)
            '-------------------------
            For m1 As Int16 = 1 To 15
                If m1 = 1 Or m1 = 2 Or m1 = 3 Or m1 = 5 Or m1 = 6 Or m1 = 10 Or m1 = 15 Then
                    rptRow = New HtmlTableRow
                    rptTb.Controls.Add(rptRow)
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                Else
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                End If
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                Select Case m1
                    Case 1
                        rptCell.Attributes.Add("colspan", 5)
                        rptCell.Attributes.Add("style", "font-size:16pt;font-family:DFKai-SB")
                        rptCell.InnerHtml = Me.ViewState("OrgName")      '單位  1
                    Case 2
                        rptCell.Attributes.Add("colspan", 5)
                        rptCell.Attributes.Add("style", "font-size:16pt;font-family:DFKai-SB")
                        rptCell.InnerHtml = "受訓學員成績單"            '1
                    Case 3
                        rptCell.Attributes.Add("colspan", 3)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp" & Me.ViewState("PLANNAME")       '計劃名稱 2
                    Case 4
                        rptCell.Attributes.Add("colspan", 2)
                        rptCell.InnerHtml = "&nbsp列印日期： " & Now.ToString("yyyy/MM/dd")  '5
                    Case 5
                        rptCell.Attributes.Add("colspan", 5)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp班　　別： " & Convert.ToString(dt1.Rows(0)("ClassCName"))  '1
                        '------------------------- 姓名：  學號：    訓練期間： 
                    Case 6
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp姓名：" & Convert.ToString(dt1.Rows(n)("Name"))
                    Case 7
                        rptCell.Attributes.Add("colspan", 2)
                        rptCell.InnerHtml = "學號：" & Convert.ToString(dt1.Rows(n)("StudentID"))
                    Case 8
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "訓練期間："
                    Case 9
                        rptCell.Attributes.Add("colspan", 2)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = Convert.ToDateTime(dt1.Rows(0)("STDate")).ToString("yyyy/MM/dd") & "～" & Convert.ToDateTime(dt1.Rows(0)("FTDate")).ToString("yyyy/MM/dd")
                        '------------------------- [CourseID]   課程名稱：    訓練時數：   成績:
                    Case 10
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp"
                    Case 11
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp&nbsp課程名稱"
                    Case 12
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "訓練時數"
                    Case 13
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp"
                    Case 14
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "成績"
                    Case 15
                        rptRow = New HtmlTableRow
                        rptTb.Controls.Add(rptRow)
                        rptCell.Attributes.Add("colspan", 5)
                        rptCell.InnerHtml = "-------------------------------------------------------------------------------------------"
                End Select
            Next
            '------------------------- [CourseID]   課程名稱：    訓練時數：   成績:( 內容  )
            For m As Int16 = 0 To dt4.Rows.Count - 1
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)
                For p As Int16 = 1 To 5
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:11pt;font-family:DFKai-SB")
                    rptCell.Attributes.Add("colspan", 1)
                    Select Case p
                        Case 1
                            rptCell.Attributes.Add("align", "left")
                            rptCell.InnerHtml = "&nbsp&nbsp&nbsp" & Convert.ToString(dt4.Rows(m)("CourseID")).ToUpper '1
                            rptCell.Attributes.Add("width", "120px")
                        Case 2
                            rptCell.InnerHtml = Convert.ToString(dt4.Rows(m)("CourseName"))  '2
                            rptCell.Attributes.Add("style", "layout-flow@horizontal;text-align@left;word-wrap: break-word;word-break:keep-all;font-size:11pt;font-family:DFKai-SB")
                            rptCell.Attributes.Add("width", "250px")
                        Case 3
                            rptCell.InnerHtml = dt7.Select("courid='" & dt4.Rows(m)("courid") & "'")(0)("Hours")  '"訓練時數："  '3
                            rptCell.Attributes.Add("width", "80px")
                        Case 4
                            rptCell.InnerHtml = "&nbsp"     '4
                            rptCell.Attributes.Add("width", "120px")
                        Case 5
                            If dt5.Select("  courid='" & dt4.Rows(m)("courid") & "'  and   socid=" & SOCID).Length <> 0 Then
                                rptCell.InnerHtml = dt5.Select("  courid='" & dt4.Rows(m)("courid") & "'  and   socid=" & SOCID)(0)("results")  '"成績:"  '5
                            Else
                                rptCell.InnerHtml = "&nbsp" '"成績:"  '5
                            End If
                    End Select
                Next
                'If n <> dt1.Rows.Count - 1 Then  '最後一頁不做分頁
                '    If m Mod 12 = 0 Then  '每十二筆做一次分頁
                '        'PrintEmptPage() '分頁
                '    End If
                'End If
            Next

            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("colspan", 5)
            rptCell.InnerHtml = "&nbsp&nbsp"

            For m2 As Int16 = 1 To 26
                'rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                Select Case m2
                    Case 1, 6, 11, 16, 21, 22 '跳行
                        rptRow = New HtmlTableRow
                        rptTb.Controls.Add(rptRow)
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                    Case Else
                        rptCell = New HtmlTableCell
                        rptRow.Controls.Add(rptCell)
                End Select
                Select Case m2
                    '-----------------學科時數：  學科平均成績：
                    Case 1
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp&nbsp&nbsp學科時數："   '1
                    Case 2
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        'rptCell.InnerHtml = Request("H1")  '"學科時數："  '2  091217 Andy 改成抓Plan_PlanInfo
                        rptCell.InnerHtml = Convert.ToString(Int(IIf(dt8.Rows(0)("GenSciHours").ToString = "", 0, dt8.Rows(0)("GenSciHours"))) + Int(IIf(dt8.Rows(0)("ProSciHours").ToString = "", 0, dt8.Rows(0)("ProSciHours").ToString)))
                    Case 3
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp&nbsp "          '3
                    Case 4
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "right")
                        rptCell.InnerHtml = "學科平均成績："  '4
                    Case 5
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = Get_dtValue(dt6, "pAvg") '學科平均成績： 
                        '---------------- -術科時數：  術科平均成績：
                    Case 6
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp&nbsp&nbsp術科時數："
                    Case 7
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        'rptCell.InnerHtml = Request("H2")   '"術科時數："  '2
                        rptCell.InnerHtml = Convert.ToString(dt8.Rows(0)("ProTechHours"))
                    Case 8
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp "   '3
                    Case 9
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "right")
                        rptCell.InnerHtml = "術科平均成績："  '4
                    Case 10
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = Get_dtValue(dt6, "sAvg") '術科平均成績： 
                        '-----------------總時數：  操行成績：
                    Case 11
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = "&nbsp&nbsp&nbsp總時數："   '1'總時數 TotalHours
                    Case 12
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = Get_dtValue(dt8, "TotalHours") '總時數 TotalHours
                    Case 13
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp "   '3
                    Case 14
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "right")
                        rptCell.InnerHtml = "操行成績："  '4
                    Case 15
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = Convert.ToString(dt1.Rows(n)("conductpoint")) '操行成績：
                        rptCell.Attributes.Add("align", "left")
                        '-----------------總平均成績：
                    Case 16
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp"   '1
                    Case 17
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = " "
                    Case 18
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.InnerHtml = "&nbsp "
                        '--------------------- '總平均成績：
                    Case 19
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "right")
                        rptCell.InnerHtml = "總平均成績："  '4
                    Case 20
                        rptCell.Attributes.Add("colspan", 1)
                        rptCell.Attributes.Add("align", "left")
                        rptCell.InnerHtml = Get_dtValue(dt6, "totalAvg") '總平均成績：  pAvg 
                        'rptCell.InnerHtml = Convert.ToString((Convert.ToSingle(dt6.Rows(0)("pAvg")) + Convert.ToSingle(dt6.Rows(0)("pAvg"))) / 2)           '總平均成績： 
                    Case 21
                        rptCell.InnerHtml = "&nbsp"
                        rptCell.Attributes.Add("colspan", 5)
                    Case 22
                        rptCell.InnerHtml = "&nbsp&nbsp&nbsp製表人："
                        rptCell.Attributes.Add("colspan", 1)
                    Case 23
                        rptCell.Attributes.Add("align", "right")
                        rptCell.InnerHtml = "教務課長："
                        rptCell.Attributes.Add("colspan", 1)
                    Case 24
                        rptCell.InnerHtml = "&nbsp"
                        rptCell.Attributes.Add("colspan", 1)
                    Case 25
                        rptCell.InnerHtml = "&nbsp"
                        rptCell.Attributes.Add("colspan", 1)
                    Case 26
                        'rptCell.InnerHtml = "中心主任："
                        rptCell.InnerHtml = "分署長："
                        rptCell.Attributes.Add("colspan", 1)
                End Select
            Next
            If n <> dt1.Rows.Count - 1 Then  '最後一頁不做分頁
                PrintEmptPage() '分頁
            End If
        Next
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("colspan", 5)
        '---------------------------------------- 製表人： 教務課長： 分署長(中心主任)：

#Region "(No Use)"

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me.Page, "發生錯誤!! " & ex.Message.ToString())
        '    Throw ex 'Exit Sub
        'Finally
        '    'conn.Close()
        '    If Not dt1 Is Nothing Then dt1.Dispose()
        '    If Not dt2 Is Nothing Then dt2.Dispose()
        '    If Not dt3 Is Nothing Then dt3.Dispose()
        '    If Not dt4 Is Nothing Then dt4.Dispose()
        'End Try

#End Region
    End Sub


    ''' <summary> 取得學、術科、總時數 </summary>
    ''' <param name="oConn"></param>
    ''' <param name="OCID"></param>
    ''' <returns></returns>
    Public Shared Function getHour(ByRef oConn As SqlConnection, ByVal OCID As String) As DataTable
        Dim dt As New DataTable
        Call TIMS.OpenDbConn(oConn)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.GenSciHours ,a.ProSciHours ,a.ProTechHours ,a.OtherHours ,a.TotalHours " & vbCrLf
        sql &= " FROM Plan_PlanInfo a " & vbCrLf
        sql &= " JOIN Class_ClassInfo b ON b.ComIDNO = a.ComIDNO AND a.SeqNO = b.SeqNO AND b.PlanID = a.PlanID " & vbCrLf
        sql &= " WHERE 1=1 AND b.OCID = @OCID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, oConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function

    ''' <summary>
    ''' 取得第1個欄位值
    ''' </summary>
    ''' <param name="dtX"></param>
    ''' <param name="fieldname"></param>
    ''' <returns></returns>
    Public Shared Function Get_dtValue(ByRef dtX As DataTable, ByVal fieldname As String) As String
        Dim rst As String = "&nbsp"
        If dtX Is Nothing Then Return rst
        If dtX.Rows.Count > 0 Then rst = Convert.ToString(dtX.Rows(0)(fieldname))
        Return rst
    End Function
End Class
