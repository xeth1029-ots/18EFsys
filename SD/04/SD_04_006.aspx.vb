Imports System.Web.Services

Partial Class SD_04_006
    Inherits AuthBasePage

    'SD_04_006_R
    'SD_04_006_R2
    'SD_04_006_R3
    'SD_04_006_R*.jrxml
    Const cst_reportFN1 As String = "SD_04_006_R"
    Const cst_reportFN2 As String = "SD_04_006_R2"
    Const cst_reportFN3 As String = "SD_04_006_R3"
    '2000/13
    'Prepared or callable statement has more than 2000 parameter markers
    Const cst_str_max_parms_count As String = "150"

    Const cst_kindEngage_內 As String = "1"
    Const cst_kindEngage_外 As String = "2"

    Const cst_NOclsdata As String = "(查無該課程資訊)"
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'TIMS.TestDbConn(Me, objConn)
        Call TIMS.OpenDbConn(objConn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            SearchTable.Visible = True
            DataGridTable.Visible = False

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Call CreateItem()

            'Button3_Click(sender, e)
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

            InSelectAll.Attributes("onclick") = "GetAllTeach(1,this.checked);"
            OutSelectAll.Attributes("onclick") = "GetAllTeach(2,this.checked);"
            CourseRound1.Attributes("onclick") = "GetAllCourse(1,this.checked);"
            CourseRound2.Attributes("onclick") = "GetAllCourse(2,this.checked);"
            CourseRound3.Attributes("onclick") = "GetAllCourse(3,this.checked);"

            Button1.Attributes("onclick") = "return check_data();"
            btnPrint.Attributes("onclick") = "return check_data();"

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=btnSchTeach');"
            Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "btnSchTeach", False, "SD_04_006")
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

    ''' <summary>
    ''' 建立物件
    ''' </summary>
    Sub CreateItem()
        'Dim dv As New DataView
        With ClassNum
            .Items.Clear()
            For i As Integer = 1 To 12
                .Items.Add(New ListItem("第" & i & "節", i))
            Next
        End With

        Dim htP As New Hashtable
        '重新 產生內聘老師選項
        htP.Clear()
        htP.Add("kindEngage", cst_kindEngage_內)
        'htP.Add("teachCName", txtSchTeachCName.Text)
        htP.Add("RIDVALUE", RIDValue.Value)
        Utl_ReSetTeach(InTeach, objConn, htP)

        '重新 產生外聘老師選項
        htP.Clear()
        htP.Add("kindEngage", cst_kindEngage_外)
        'htP.Add("teachCName", txtSchTeachCName.Text)
        htP.Add("RIDVALUE", RIDValue.Value)
        Utl_ReSetTeach(OutTeach, objConn, htP)
    End Sub

    ''' <summary>
    ''' 查詢資料SQL
    ''' </summary>
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim TeachID As New ArrayList
        Dim TeachName As New ArrayList
        Dim TeacherCon As String = ""
        For i As Integer = 0 To InTeach.Items.Count - 1
            If InTeach.Items(i).Selected = True AndAlso InTeach.Items(i).Value <> "" Then
                If TeacherCon <> "" Then TeacherCon &= ","
                TeacherCon &= "'" & InTeach.Items(i).Value & "'"

                TeachID.Add(InTeach.Items(i).Value)
                TeachName.Add(InTeach.Items(i).Text)
            End If
        Next
        For i As Integer = 0 To OutTeach.Items.Count - 1
            If OutTeach.Items(i).Selected = True AndAlso OutTeach.Items(i).Value <> "" Then
                If TeacherCon <> "" Then TeacherCon &= ","
                TeacherCon &= "'" & OutTeach.Items(i).Value & "'"

                TeachID.Add(OutTeach.Items(i).Value)
                TeachName.Add(OutTeach.Items(i).Text)
            End If
        Next

        Dim Flag As Boolean = False '選擇節次:沒有
        For i As Integer = 0 To ClassNum.Items.Count - 1
            If ClassNum.Items(i).Selected = True Then
                Flag = True '選擇節次:有
                Exit For
            End If
        Next

        Dim ClassCon As String = ""
        If TeacherCon <> "" Then
            For i As Integer = 0 To ClassNum.Items.Count - 1 '1~12
                If ClassNum.Items(i).Selected Or (Not Flag) Then
                    Select Case rblTeachtype.SelectedValue '1:教師'2:助教1'3:助教2
                        Case "1" '教師
                            Dim i1 As Integer = i + 1
                            If ClassCon <> "" Then ClassCon &= " OR "
                            ClassCon &= "a.Teacher" & i1 & " IN (" & TeacherCon & ")"
                        Case "2" '助教1
                            Dim i1 As Integer = i + 13
                            If ClassCon <> "" Then ClassCon &= " OR "
                            ClassCon &= "a.Teacher" & i1 & " IN (" & TeacherCon & ")"
                        Case "3" '助教2
                            Dim i1 As Integer = i + 25
                            If ClassCon <> "" Then ClassCon &= " OR "
                            ClassCon &= "a.Teacher" & i1 & " IN (" & TeacherCon & ")"
                    End Select
                End If
            Next
            If ClassCon <> "" Then
                ClassCon = " and (" & ClassCon & ")" & vbCrLf
            End If
        End If

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT b.ClassCName" & vbCrLf
        sql &= " ,a.*" & vbCrLf
        sql &= " FROM Class_Schedule a" & vbCrLf
        sql &= " JOIN Class_ClassInfo b ON a.OCID=b.OCID and a.Formal='Y'" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If OCIDValue1.Value <> "" Then
            sql &= " AND b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        If Me.txtCJOB_NAME.Text <> "" AndAlso Me.cjobValue.Value <> "" Then
            sql &= " AND b.CJOB_UNKEY= '" & Me.cjobValue.Value & "'" & vbCrLf
        End If
        If SDate.Text <> "" Then
            sql &= " AND a.SchoolDate>= " & TIMS.To_date(SDate.Text) & vbCrLf
        End If
        If EDate.Text <> "" Then
            sql &= " AND a.SchoolDate<= " & TIMS.To_date(EDate.Text) & vbCrLf
        End If
        If ClassCon <> "" Then
            sql &= ClassCon '1:教師'2:助教1'3:助教2
        End If

        '課表
        Dim dt As New DataTable
        Try
            dt.Load(DbAccess.GetReader(sql, objConn))
        Catch ex As Exception
            TIMS.WriteTraceLog(Me.Page, ex)

            Common.MessageBox(Me, "資料庫效能異常，請縮小範圍重新查詢!!")
            Exit Sub
        End Try
        'If RIDValue.Value <> "" Then
        '    sql = "SELECT * FROM Course_CourseInfo WHERE RID='" & RIDValue.Value & "'"
        'Else
        '    sql = "SELECT * FROM Course_CourseInfo WHERE OrgID='" & sm.UserInfo.OrgID & "'"
        'End If

        '課程名稱
        sql = "SELECT * FROM COURSE_COURSEINFO WHERE OrgID='" & sm.UserInfo.OrgID & "'"
        Dim CourseTable As New DataTable
        CourseTable.Load(DbAccess.GetReader(sql, objConn))
        'CourseTable = DbAccess.GetDataTable(Sql, objConn)

        '組合表
        Dim drTemp As DataRow = Nothing
        Dim TechClassTable As New DataTable
        TechClassTable.Columns.Add(New DataColumn("TeachCName")) '老師
        TechClassTable.Columns.Add(New DataColumn("SchoolDate", System.Type.GetType("System.DateTime"))) '日期
        TechClassTable.Columns.Add(New DataColumn("ClassCName")) '班名
        TechClassTable.Columns.Add(New DataColumn("CourseName")) '課名
        TechClassTable.Columns.Add(New DataColumn("ClassNum")) '節

        Dim ff As String = ""
        'Dim dr As DataRow
        For i As Integer = 0 To ClassNum.Items.Count - 1 '1~12
            'Dim sCourseName As String = cst_NOclsdata
            If ClassNum.Items(i).Selected = True Or Flag = False Then
                For j As Integer = 0 To TeachID.Count - 1
                    Dim s1 As String = CStr(i + 1)
                    'Dim s13 As String = CStr(i + 13)
                    'Dim s25 As String = CStr(i + 25)
                    Select Case rblTeachtype.SelectedValue '1:教師'2:助教1'3:助教2
                        Case "1"
                            Dim t1 As String = CStr(i + 1)
                            ff = "Teacher" & t1 & "='" & TeachID(j) & "'"
                            For Each dr As DataRow In dt.Select(ff)
                                'sCourseName = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp = TechClassTable.NewRow
                                TechClassTable.Rows.Add(drTemp)
                                drTemp("TeachCName") = TeachName(j)
                                drTemp("SchoolDate") = dr("SchoolDate")
                                drTemp("ClassCName") = dr("ClassCName")
                                drTemp("CourseName") = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp("ClassNum") = s1 'i + 1
                            Next

                        Case "2"
                            Dim t1 As String = CStr(i + 13)
                            ff = "Teacher" & t1 & "='" & TeachID(j) & "'"
                            For Each dr As DataRow In dt.Select(ff)
                                'sCourseName = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp = TechClassTable.NewRow
                                TechClassTable.Rows.Add(drTemp)
                                drTemp("TeachCName") = TeachName(j)
                                drTemp("SchoolDate") = dr("SchoolDate")
                                drTemp("ClassCName") = dr("ClassCName")
                                drTemp("CourseName") = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp("ClassNum") = s1 'i + 1
                            Next

                        Case "3"
                            Dim t1 As String = CStr(i + 25)
                            ff = "Teacher" & t1 & "='" & TeachID(j) & "'"
                            For Each dr As DataRow In dt.Select(ff)
                                'sCourseName = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp = TechClassTable.NewRow
                                TechClassTable.Rows.Add(drTemp)
                                drTemp("TeachCName") = TeachName(j)
                                drTemp("SchoolDate") = dr("SchoolDate")
                                drTemp("ClassCName") = dr("ClassCName")
                                drTemp("CourseName") = sUtl_GetCourseName(dr("Class" & s1), CourseTable)
                                drTemp("ClassNum") = s1 'i + 1
                            Next

                    End Select

                Next
            End If
        Next

        msg.Text = "查無資料"
        SearchTable.Visible = True
        DataGridTable.Visible = False

        If TechClassTable.Rows.Count > 0 Then
            msg.Text = ""
            SearchTable.Visible = False
            DataGridTable.Visible = True

            PageControler1.PageDataTable = TechClassTable
            PageControler1.Sort = "TeachCName,SchoolDate"
            PageControler1.ControlerLoad()
        End If

        'sql += "(SELECT TechID FROM Teach_TeacherInfo WHERE RID='" & sm.UserInfo.RID & "'" & TeacherCon & ") a "
        'For i As Integer = 0 To ClassNum.Items.Count - 1
        '    sql += "JOIN (SELECT * FROM Class_Schedule WHERE Formal='Y'" & ClassCon & ") b" & i + 1 & " ON b" & i + 1 & ".Teacher" & i + 1 & "=a.TechID "
        '    sql += "JOIN (SELECT * FROM Class_Schedule WHERE Formal='Y'" & ClassCon & ") b" & i + 13 & " ON b" & i + 13 & ".Teacher" & i + 13 & "=a.TechID "
        'Next
    End Sub

    ''' <summary>
    ''' 查詢資料
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    Function sUtl_GetCourseName(ByVal sCourID As String, ByVal dtCourse As DataTable) As String
        Dim rst As String = cst_NOclsdata '""
        If sCourID = "" Then Return rst

        Dim ff As String = "CourID=" & sCourID
        If dtCourse.Select(ff).Length > 0 Then
            rst = dtCourse.Select(ff)(0)("CourseName")
        End If
        'Try
        'Catch ex As Exception
        '    rst = cst_NOclsdata
        'End Try
        Return rst
    End Function

    ''' <summary>
    ''' 列印
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim i_TeacherCon As Integer = 0
        Dim TeacherCon As String = ""
        For i As Integer = 0 To InTeach.Items.Count - 1
            If InTeach.Items(i).Selected = True AndAlso InTeach.Items(i).Value <> "" Then
                i_TeacherCon += 1
                If TeacherCon <> "" Then TeacherCon &= ","
                TeacherCon &= InTeach.Items(i).Value
            End If
        Next

        For i As Integer = 0 To OutTeach.Items.Count - 1
            If OutTeach.Items(i).Selected = True AndAlso OutTeach.Items(i).Value <> "" Then
                i_TeacherCon += 1
                If TeacherCon <> "" Then TeacherCon &= ","
                TeacherCon &= OutTeach.Items(i).Value
            End If
        Next

        Dim OrgID As String = ""
        If RIDValue.Value = "" Then OrgID = sm.UserInfo.OrgID

        If i_TeacherCon > Val(cst_str_max_parms_count) Then
            Common.MessageBox(Me, "超過" & cst_str_max_parms_count & "個參數標記，資料庫查詢限制!!")
            Exit Sub
        End If

        SDate.Text = TIMS.ClearSQM(SDate.Text)
        EDate.Text = TIMS.ClearSQM(EDate.Text)
        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "SDate", SDate.Text)
        TIMS.SetMyValue(sMyValue, "EDate", EDate.Text)
        TIMS.SetMyValue(sMyValue, "OrgID", OrgID)
        TIMS.SetMyValue(sMyValue, "TechID", TeacherCon)
        TIMS.SetMyValue(sMyValue, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(sMyValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(sMyValue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(sMyValue, "CJOB_UNKEY", cjobValue.Value)

        Dim sFileName As String = cst_reportFN1 ' "SD_04_006_R"
        Select Case rblTeachtype.SelectedValue '1:教師'2:助教1'3:助教2
            Case "1"
                sFileName = cst_reportFN1 '"SD_04_006_R"
            Case "2"
                sFileName = cst_reportFN2 '"SD_04_006_R2"
            Case "3"
                sFileName = cst_reportFN3 '"SD_04_006_R3"
        End Select
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sFileName, sMyValue)
    End Sub

    ''' <summary>
    ''' 清除
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clear.Click
        TMID1.Text = ""
        TMIDValue1.Value = ""
        OCID1.Text = ""
        OCIDValue1.Value = ""
    End Sub

    ''' <summary>
    ''' 回上頁
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        SearchTable.Visible = True
        DataGridTable.Visible = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

#Region "2018 add 查詢條件-師資姓名模糊查詢"
    <WebMethod()>
    Public Shared Function GetInteachList() As String
        Return ""
    End Function

    ''' <summary>
    ''' 老師姓名查詢鈕-click事件
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSchTeach_Click(sender As Object, e As EventArgs) Handles btnSchTeach.Click
        txtSchTeachCName.Text = TIMS.ClearSQM(txtSchTeachCName.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        Dim htP As New Hashtable
        '重新 產生內聘老師選項
        htP.Clear()
        htP.Add("kindEngage", cst_kindEngage_內)
        htP.Add("teachCName", txtSchTeachCName.Text)
        htP.Add("RIDVALUE", RIDValue.Value)
        Utl_ReSetTeach(InTeach, objConn, htP)

        '重新 產生外聘老師選項
        htP.Clear()
        htP.Add("kindEngage", cst_kindEngage_外)
        htP.Add("teachCName", txtSchTeachCName.Text)
        htP.Add("RIDVALUE", RIDValue.Value)
        Utl_ReSetTeach(OutTeach, objConn, htP)
    End Sub


    ''' <summary>
    ''' 老師姓名-模糊查詢
    ''' </summary>
    ''' <param name="chkTeachObj">老師姓名(內外聘)勾選項目物件</param>
    ''' <param name="oConn"></param>
    ''' <param name="htP"></param>
    Public Shared Sub Utl_ReSetTeach(ByRef chkTeachObj As CheckBoxList, ByRef oConn As SqlConnection, ByRef htP As Hashtable)
        Dim sql As String = ""
        Dim dt As New DataTable

        chkTeachObj.Items.Clear()
        Dim s_kindEngage As String = TIMS.GetMyValue2(htP, "kindEngage")
        Dim s_teachCName As String = TIMS.GetMyValue2(htP, "teachCName")
        Dim s_RIDVALUE As String = TIMS.GetMyValue2(htP, "RIDVALUE") '"s_RIDVALUE"

        Call TIMS.OpenDbConn(oConn)
        '查詢所屬單位老師資料
        sql = ""
        sql &= " select techid,teachcname " & vbCrLf
        sql &= " FROM TEACH_TEACHERINFO WITH(NOLOCK)" & vbCrLf
        sql &= " where 1=1 " & vbCrLf
        sql &= " and kindengage=@kindengage " & vbCrLf
        sql &= " and rid=@rid " & vbCrLf
        If s_teachCName <> "" Then sql &= " and teachcname like '%' + @teachcname + '%'"
        sql &= " order by techid "

        Dim parms As Hashtable = New Hashtable()
        parms.Add("kindengage", s_kindEngage) '內外聘別 (1.內, 2.外)
        parms.Add("rid", s_RIDVALUE) '機構RID
        If s_teachCName <> "" Then parms.Add("teachcname", s_teachCName)
        dt = DbAccess.GetDataTable(sql, oConn, parms)

        With chkTeachObj
            .DataSource = dt
            .DataTextField = "teachcname"
            .DataValueField = "techid"
            .DataBind()
        End With
    End Sub
#End Region
End Class


