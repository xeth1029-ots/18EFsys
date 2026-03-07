Public Class SD_04_002_c1
    Inherits AuthBasePage

    'Class_Schedule
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TIMS.ChkSession(Me, 0, sm)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call Create1()
        End If
    End Sub

    Sub Create1()
        Dim flagCanShow As Boolean = False
        Dim rOCID As String = TIMS.ClearSQM(Request("OCID"))
        If rOCID <> "" AndAlso IsNumeric(rOCID) Then flagCanShow = True
        If Not flagCanShow Then
            Common.MessageBox(Me, "資料異常,請重新查詢")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(rOCID, objConn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資料異常,請重新查詢")
            Exit Sub
        End If
        Call sHide_TableShow1(rOCID)
    End Sub

    '將目前所有的使用課程列出 'Public Shared Sub sHide_TableShow1()
    Sub sHide_TableShow1(ByVal sOCID As String)
        Dim drCC As DataRow = TIMS.GetOCIDDate(sOCID, objConn)
        If drCC Is Nothing Then Exit Sub

        'Optional ByVal sType As Integer = 1
        'sType 1@List 班級課表 2@List 某天的班級課表
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        Dim CourseList As String = ""
        Dim CourseArray As Array
        'Const Cst_d1 As String = "convert(varchar(7),Schooldate,111)"
        'Const Cst_ym1 As String = " ym1"
        sOCID = TIMS.ClearSQM(sOCID)
        If sOCID = "" Then Exit Sub
        Call TIMS.OpenDbConn(objConn)
        'sql = "SELECT *,convert(varchar(7),Schooldate,111) ym1 FROM Class_Schedule WHERE OCID='" & OCID.SelectedValue & "'"
        sql = "SELECT a.*, FORMAT(a.Schooldate, 'yyyy_MM') ym3 FROM CLASS_SCHEDULE a WHERE a.OCID='" & sOCID & "'"
        dt = DbAccess.GetDataTable(sql, objConn)
        If dt.Rows.Count = 0 Then Exit Sub

        For Each dr As DataRow In dt.Rows
            For i As Integer = 1 To 12
                Dim vTmpClassI As String = TIMS.ClearSQM(dr("Class" & i))
                If Not IsDBNull(dr("Class" & i)) AndAlso vTmpClassI <> "" Then
                    Dim Flag As Boolean = False
                    CourseArray = Split(CourseList, ",")
                    For j As Integer = 0 To CourseArray.Length - 1
                        If CourseArray(j) = vTmpClassI Then
                            Flag = True
                            Exit For
                        End If
                    Next
                    If Not Flag Then
                        '表示不存在,增加課程代碼
                        If CourseList <> "" Then CourseList &= ","
                        CourseList &= vTmpClassI
                    End If
                End If
            Next
        Next

        'DataGrid3.Visible = False
        DataGrid4.Visible = False '檢視每月已排課時數
        If CourseList <> "" Then '有科目範圍
            sql = ""
            sql &= " SELECT a.CourID, a.CourseName, b.CourseName MCourseName, 0 TotalHours " & vbCrLf
            sql &= " FROM (SELECT * FROM Course_CourseInfo WHERE 1=1 AND CourID IN (" & CourseList & ")) a " & vbCrLf
            sql &= " LEFT JOIN Course_CourseInfo b ON a.MainCourID = b.CourID " & vbCrLf
            sql &= " ORDER BY a.CourID, a.CourseName " & vbCrLf
            Dim dt1 As DataTable '預備塞入課表 (DataGrid3)
            dt1 = DbAccess.GetDataTable(sql, objConn)

            sql = "" & vbCrLf
            'ym1 設計比對 與dt table之中要建立喔
            sql &= " SELECT 0 AS TotalHours, CONVERT(varchar(7), Schooldate, 111) ym1 " & vbCrLf
            sql &= " FROM Class_Schedule " & vbCrLf
            sql &= " WHERE ocid = '" & sOCID & "' " & vbCrLf
            sql &= " GROUP BY CONVERT(varchar(7), Schooldate, 111) " & vbCrLf
            sql &= " ORDER BY CONVERT(varchar(7), Schooldate, 111) " & vbCrLf

            '1.取出上課年月
            sql = "" & vbCrLf
            sql &= " SELECT * FROM (" & vbCrLf
            'ym2 設計比對 與dt table之中要建立喔
            sql &= " SELECT DISTINCT REPLACE(CONVERT(varchar(7), Schooldate, 111),'/','年') + '月' ym2 " & vbCrLf
            sql &= " ,CONVERT(varchar(7), Schooldate, 111) ym1" & vbCrLf
            sql &= " ,FORMAT(Schooldate, 'yyyy_MM') ym3 " & vbCrLf
            sql &= " FROM Class_Schedule " & vbCrLf
            sql &= " WHERE ocid = '" & sOCID & "' " & vbCrLf
            sql &= " ) g " & vbCrLf
            sql &= " ORDER BY ym2 " & vbCrLf
            Dim dt2 As DataTable '預備塞入課表 (DataGrid4)
            dt2 = DbAccess.GetDataTable(sql, objConn)

            '2.取出上課科目，並組合年月
            sql = ""
            sql &= " SELECT a.CourID, a.CourseName, b.CourseName AS MCourseName, 0 AS TotalHours " & vbCrLf
            For i As Integer = 0 To dt2.Rows.Count - 1
                sql &= " ,0 """ & dt2.Rows(i)("ym3") & """" & vbCrLf
                'sql += ",0 as '" & dt2.Rows(i)("ym2") & "'" & vbCrLf
            Next
            sql &= " FROM (SELECT * FROM Course_CourseInfo WHERE 1=1 AND CourID IN (" & CourseList & ")) a " & vbCrLf
            sql &= " LEFT JOIN Course_CourseInfo b ON a.MainCourID = b.CourID " & vbCrLf
            sql &= " ORDER BY a.CourID, a.CourseName " & vbCrLf
            Dim dt3 As DataTable '預備塞入課表 (DataGrid4)
            dt3 = DbAccess.GetDataTable(sql, objConn)

            Const cst_科目 As String = "科目"
            Const cst_累計時數 As String = "累計時數"

            'Empty 中文表格 'dt3b 沒有資料
            sql = ""
            sql &= " SELECT CONVERT(NVARCHAR(50),NULL) """ & cst_科目 & """" & vbCrLf
            For i As Integer = 0 To dt2.Rows.Count - 1
                sql &= " ,0 """ & dt2.Rows(i)("ym2") & """" & vbCrLf '中文年月
            Next
            sql &= " ,0 """ & cst_累計時數 & """ " & vbCrLf
            sql &= " FROM COURSE_COURSEINFO WHERE 1<>1 " & vbCrLf
            Dim dt3b As DataTable '預備塞入課表 (DataGrid4)
            dt3b = DbAccess.GetDataTable(sql, objConn)

            'dt3 沒有資料
            For Each dr1 As DataRow In dt1.Rows '預備塞入課表
                For Each dr As DataRow In dt.Rows '目前排入課程
                    For i As Integer = 1 To 12 '每天12節
                        Dim vTmpClassI As String = TIMS.ClearSQM(dr("Class" & i))
                        If Not IsDBNull(dr("Class" & i)) AndAlso vTmpClassI <> "" Then '有課程
                            If dr("Class" & i) = dr1("CourID") Then
                                dr1("TotalHours") += 1 '時數+1

                                If dt3.Select("CourID='" & vTmpClassI & "'").Length > 0 Then
                                    Dim dr3 As DataRow
                                    dr3 = dt3.Select("CourID='" & vTmpClassI & "'")(0)
                                    dr3(dr("ym3")) += 1 '某年月 時數+1
                                    dr3("TotalHours") += 1 '時數+1
                                End If
                            End If
                        End If
                    Next
                Next
            Next

            'dt3 已有資料
            For i As Integer = 0 To dt3.Rows.Count - 1
                Dim dr3 As DataRow
                dr3 = dt3.Rows(i) '取得

                Dim dr3b As DataRow
                If dr3("TotalHours") > 0 Then '有時數
                    If dt3b.Select(cst_科目 & "='" & dr3("CourseName") & "'").Length > 0 Then
                        '取得資料
                        dr3b = dt3b.Select(cst_科目 & "='" & dr3("CourseName") & "'")(0)
                    Else
                        '建立欄位初始值
                        dr3b = dt3b.NewRow
                        dt3b.Rows.Add(dr3b)
                        dr3b(cst_科目) = dr3("CourseName")
                        For j As Integer = 0 To dt2.Rows.Count - 1
                            dr3b(dt2.Rows(j)("ym2")) = 0
                        Next
                        dr3b(cst_累計時數) = 0
                    End If

                    '補資料
                    For j As Integer = 0 To dt2.Rows.Count - 1
                        dr3b(dt2.Rows(j)("ym2")) += CInt(dr3(dt2.Rows(j)("ym3")))
                    Next
                    dr3b(cst_累計時數) += CInt(dr3("TotalHours"))
                End If
            Next

            If dt2.Rows.Count > 0 Then
                'dt2.AcceptChanges()
                DataGrid4.Visible = True '檢視每月已排課時數
                DataGrid4.DataSource = dt3b 'dt3 'dt2
                DataGrid4.DataBind()
            End If
        End If

    End Sub

    Private Sub btnReNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReNew.Click
        Call Create1()
    End Sub
End Class