Partial Class SYS_04_001_add
    Inherits AuthBasePage

    'Dim objtable As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        'back.Attributes("onclick") = "history.go(-1);"
    End Sub

    Sub SaveData1()
        Reason.Text = TIMS.ClearSQM(Reason.Text)
        Dim sDate As String
        Dim eDate As String
        sDate = TIMS.Cdate3(Me.Start_Date.Text)
        eDate = TIMS.Cdate3(Me.End_Date.Text)
        If Reason.Text = "" Then
            Common.MessageBox(Me, "事由不可為空!!")
            Exit Sub
        End If
        If sDate = "" Then
            Common.MessageBox(Me, "起始日期 不可為空!!")
            Exit Sub
        End If
        'If Reason.Text = "" Then Exit Sub
        'If sDate = "" Then Exit Sub
        'If eDate = "" Then Exit Sub
        Dim iTotalDays As Integer = 0
        If eDate <> "" Then
            iTotalDays = DateDiff(DateInterval.Day, CDate(sDate), CDate(eDate))
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select 'x' from SYS_HOLIDAY where RID=@RID and HOLDATE=@HOLDATE" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " insert into SYS_HOLIDAY(RID,HOLDATE,REASON,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql &= " values(@RID,@HOLDATE,@REASON,@MODIFYACCT,getdate())" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " update SYS_HOLIDAY" & vbCrLf
        sql &= " set REASON=@REASON,MODIFYACCT=@MODIFYACCT,MODIFYDATE=getdate()" & vbCrLf
        sql &= " where RID=@RID and HOLDATE=@HOLDATE" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        'Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)

        For i As Integer = 0 To iTotalDays
            Dim dt As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                .Parameters.Add("HOLDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(sDate)
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count = 0 Then
                '新增
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                    .Parameters.Add("HOLDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(sDate)
                    .Parameters.Add("REASON", SqlDbType.VarChar).Value = Reason.Text
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .ExecuteNonQuery()
                    'dt.Load(.ExecuteReader())
                End With
            Else
                '修改
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("REASON", SqlDbType.VarChar).Value = Reason.Text
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                    .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                    .Parameters.Add("HOLDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(sDate)
                    .ExecuteNonQuery()
                    'dt.Load(.ExecuteReader())
                End With
            End If
            sDate = TIMS.Cdate3(DateAdd(DateInterval.Day, 1, CDate(sDate)))
        Next

        'Call CloseDbConn(conn)
        'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))
        Dim url1 As String = TIMS.Get_Url1(Me, "SYS_04_001.aspx")
        TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub

    Private Sub but_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_save.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'Dim objtable As DataTable
        'Dim objadapter As SqlDataAdapter
        'Dim objrow, tmprow As DataRow
        'Dim objstr, tmpstr, sDate, eDate As String
        'Dim TotalDays As Integer

        Call SaveData1()
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Dim url1 As String = TIMS.Get_Url1(Me, "SYS_04_001.aspx")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class
