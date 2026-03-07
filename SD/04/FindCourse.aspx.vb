Partial Class FindCourse
    Inherits AuthBasePage

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
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
    End Sub

    Sub sUtl_ReturnCourse1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim sRst_msg2 As String = "'','','','','','','','',''"
        RID.Text = TIMS.ClearSQM(RID.Text)
        If RID.Text = "" Then Exit Sub
        CourseID.Text = TIMS.ClearSQM(CourseID.Text)
        If CourseID.Text = "" Then
            Page.RegisterStartupScript("return", sRst_msg2)
            Exit Sub
        End If

        Dim sql As String = ""
        Dim sOrgID As String = ""
        sql = "SELECT ORGID FROM AUTH_RELSHIP WHERE RID='" & RID.Text & "'"
        Dim dtX As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dtX.Rows.Count = 0 Then
            Page.RegisterStartupScript("return", sRst_msg2)
            Exit Sub
        End If
        If dtX.Rows.Count > 0 Then
            sOrgID = dtX.Rows(0)("OrgID")
        End If

        sql = "" & vbCrLf
        sql += " SELECT a.COURID" & vbCrLf
        sql += " ,a.COURSEID" & vbCrLf
        sql += " ,a.COURSENAME" & vbCrLf
        'sql += " ,a.HOURS" & vbCrLf
        'sql += " ,a.CLASSIFICATION1" & vbCrLf
        'sql += " ,a.CLASSIFICATION2" & vbCrLf
        'sql += " ,a.VALID" & vbCrLf
        'sql += " ,a.MAINCOURID" & vbCrLf
        'sql += " ,a.RID" & vbCrLf
        'sql += " ,a.TMID" & vbCrLf
        'sql += " ,a.CLSID" & vbCrLf
        'sql += " ,a.ORGID" & vbCrLf
        sql += " ,a.TECH1" & vbCrLf
        sql += " ,a.TECH2" & vbCrLf
        sql += " ,a.TECH3" & vbCrLf
        sql += " ,a.ROOM" & vbCrLf
        sql += " ,a.ISCOUNTHOURS" & vbCrLf
        sql += " ,b.TeachCName TechName1" & vbCrLf
        sql += " ,c.TeachCName TechName2 " & vbCrLf
        sql += " ,c3.TeachCName TechName3 " & vbCrLf
        sql += " FROM Course_CourseInfo a" & vbCrLf
        sql += " JOIN Auth_Relship r on r.RID=a.RID" & vbCrLf
        sql += " LEFT JOIN Teach_TeacherInfo b ON a.Tech1=b.TechID" & vbCrLf
        sql += " LEFT JOIN Teach_TeacherInfo c ON a.Tech2=c.TechID" & vbCrLf
        sql += " LEFT JOIN Teach_TeacherInfo c3 ON a.Tech3=c3.TechID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        If sOrgID <> "" Then
            sql += " and a.OrgID=@OrgID" & vbCrLf
        Else
            sql += " and a.RID=@RID" & vbCrLf
        End If
        sql += " and a.CourseID=@CourseID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Try
            With sCmd
                .Parameters.Clear()
                If sOrgID <> "" Then
                    .Parameters.Add("OrgID", SqlDbType.VarChar).Value = sOrgID
                Else
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RID.Text
                End If
                .Parameters.Add("CourseID", SqlDbType.VarChar).Value = CourseID.Text
                dt.Load(.ExecuteReader())
            End With
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/* ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += " sOrgID:" & sOrgID & vbCrLf
            strErrmsg += " RID.Text:" & RID.Text & vbCrLf
            strErrmsg += " CourseID.Text:" & CourseID.Text & vbCrLf
            strErrmsg += "/* sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Page.RegisterStartupScript("return", sRst_msg2)
            Exit Sub
        End Try

        Dim dr As DataRow = Nothing
        If dt.Rows.Count > 0 Then dr = dt.Rows(0)

        If Not dr Is Nothing Then
            Dim sRst As String = ""
            sRst = ""
            sRst &= "'" & dr("COURID") & "','" & dr("COURSENAME") & "'"
            sRst &= ",'" & dr("TECH1").ToString & "','" & dr("TechName1").ToString & "'"
            sRst &= ",'" & dr("TECH2").ToString & "','" & dr("TechName2").ToString & "'"
            sRst &= ",'" & dr("TECH3").ToString & "','" & dr("TechName3").ToString & "'"
            sRst &= ",'" & dr("ROOM").ToString & "'"
            Page.RegisterStartupScript("return", "<script>ReturnCourse(" & sRst & ");</script>")
        Else
            Page.RegisterStartupScript("return", "<script>ReturnCourse(" & sRst_msg2 & ");</script>")
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Call sUtl_ReturnCourse1()
    End Sub

End Class
