Partial Class OB_01_001_add
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        If Not IsPostBack Then
            ddlyears = TIMS.GetSyear(ddlyears, Year(Now) - 1, Year(Now) + 3, True)
            TPlanID = TIMS.Get_TPlan(TPlanID)
            ViewState("Action") = TIMS.ClearSQM(UCase(Request("Action")))
            ViewState("TSN") = TIMS.ClearSQM(Request("TSN"))

            Dim sql As String = ""
            Dim dr As DataRow = Nothing

            If ViewState("TSN") <> "" AndAlso (ViewState("Action") = "EDIT" OrElse ViewState("Action") = "VIEW") Then

                sql = ""
                sql += " SELECT b.PlanName, a.* " & vbCrLf
                sql += " FROM OB_Tender a " & vbCrLf
                sql += " JOIN OB_Plan b on a.PlanSN=b.PlanSN " & vbCrLf
                sql += " WHERE TSN='" & ViewState("TSN") & "' " & vbCrLf
                If sm.UserInfo.DistID <> "000" Then
                    sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
                End If
                dr = DbAccess.GetOneRow(sql, objconn)
                txttsn.Text = ViewState("TSN")

                ddlyears = TIMS.GetSyear(ddlyears, dr("years") - 1, dr("years") + 3, True)
                Common.SetListItem(ddlyears, dr("years"))
                If dr("TPlanID").ToString <> "" Then
                    rb1.Checked = True
                    Common.SetListItem(TPlanID, dr("TPlanID"))
                    Me.PlanName.Text = TPlanID.SelectedItem.Text
                Else
                    rb2.Checked = True
                    Me.PlanName.Text = dr("PlanName")
                End If

                Me.TenderCName.Text = dr("TenderCName")
                Me.TenderEName.Text = dr("TenderEName").ToString

                If dr("Sponsor").ToString <> "" Then
                    Me.Sponsor.Text = dr("Sponsor")
                Else
                    Me.Sponsor.Text = ""
                End If
                Me.TenderSDate.Text = Common.FormatDate(dr("TenderSDate"))

                Me.TenderEDate.Text = Common.FormatDate(dr("TenderEDate"))
                '時間格式取前5碼
                Me.TenderEDate2.Text = Left(Common.FormatTime(dr("TenderEDate")), 5)

                Me.ReviewDate.Text = Common.FormatDate(dr("ReviewDate"))

                If dr("ResolutionDate").ToString <> "" Then
                    Me.ResolutionDate.Text = Common.FormatDate(dr("ResolutionDate"))
                Else
                    Me.ResolutionDate.Text = ""
                End If

                If ViewState("Action") = "VIEW" Then
                    Me.btnSave.Visible = False
                End If
            End If

        End If


        PageLoadSetLast1()
    End Sub

    Sub PageLoadSetLast1()
        rb1.Attributes("onclick") = "set_planname1('rb1');"
        rb2.Attributes("onclick") = "set_planname1('rb2');"

        If rb1.Checked Then
            Page.RegisterStartupScript("rb_checked", "<script>set_planname1('rb1');</script>")
        End If
        If rb2.Checked Then
            Page.RegisterStartupScript("rb_checked", "<script>set_planname1('rb2');</script>")
        End If

        'btnSave.Attributes("onclick") = "return CheckData1();"
    End Sub

    Function CheckData(ByRef Errmag As String) As Boolean
        Errmag = ""
        CheckData = False
        If ddlyears.SelectedValue = "" Then
            Errmag += "請選擇年度" & vbCrLf
        End If

        If Not (rb1.Checked Or rb2.Checked) Then
            Errmag += "請選擇採購案類型" & vbCrLf
        End If

        If rb1.Checked Then
            'TPlanID
            If TPlanID.SelectedValue = "" Then
                Errmag += "請選擇訓練計畫名稱(訓練案)" & vbCrLf
            End If
        End If
        If rb2.Checked Then
            If PlanName.Text.Trim = "" Then
                Errmag += "請輸入新訓練計畫名稱(非訓練案)" & vbCrLf
            End If
            PlanName.Text = PlanName.Text.Trim
        End If
        'TenderCName
        If TenderCName.Text.Trim = "" Then
            Errmag += "請輸入標案名稱" & vbCrLf
        End If
        TenderCName.Text = TenderCName.Text.Trim
        'TenderEName
        TenderEName.Text = TIMS.ChangeIDNO(TenderEName.Text.Trim)

        'If Sponsor.Text.Trim = "" Then
        '    Errmag += "請輸入主辦單位" & vbCrLf
        'End If
        Sponsor.Text = Sponsor.Text.Trim


        If TenderSDate.Text.Trim = "" Then
            Errmag += "請選擇投標日期起始日" & vbCrLf
        Else
            If Not IsDate(TenderSDate.Text.Trim) Then
                Errmag += "投標日期起始日應為正確日期格式" & vbCrLf
            End If
        End If

        If TenderEDate.Text.Trim = "" Then
            Errmag += "請選擇投標日期迄日" & vbCrLf
        Else
            If Not IsDate(TenderEDate.Text.Trim) Then
                Errmag += "投標日期迄日應為正確日期格式" & vbCrLf
            End If
        End If

        If TenderEDate2.Text.Trim = "" Then
            Errmag += "請輸入投標迄止時間" & vbCrLf
        Else
            Dim flag As Boolean = True
            If Not TenderEDate2.Text.Trim.IndexOf(":") > -1 Then
                flag = False
            End If
            If Not Len(TenderEDate2.Text.Trim) = 5 Then
                flag = False
            End If
            If Not IsNumeric(Left(TenderEDate2.Text.Trim, 2)) Then
                flag = False
            Else
                If Not (CInt(Left(TenderEDate2.Text.Trim, 2)) >= 0 And CInt(Left(TenderEDate2.Text.Trim, 2)) <= 23) Then
                    flag = False
                End If
            End If
            If Not IsNumeric(Right(TenderEDate2.Text.Trim, 2)) Then
                flag = False
            Else
                If Not (CInt(Right(TenderEDate2.Text.Trim, 2)) >= 0 And CInt(Right(TenderEDate2.Text.Trim, 2)) <= 59) Then
                    flag = False
                End If
            End If
            If Not flag Then
                Errmag += "投標迄止時間應為正確時間格式(23:59)" & vbCrLf
            End If
        End If

        If ReviewDate.Text.Trim = "" Then
            Errmag += "請選擇評選日期" & vbCrLf
        Else
            If Not IsDate(ReviewDate.Text.Trim) Then
                Errmag += "評選日期應為正確日期格式" & vbCrLf
            End If
        End If

        ResolutionDate.Text = ResolutionDate.Text.Trim
        If ResolutionDate.Text.Trim = "" Then
            'Errmag += "請選擇決標日期" & vbCrLf
        Else
            If Not IsDate(ResolutionDate.Text.Trim) Then
                Errmag += "決標日期應為正確日期格式" & vbCrLf
            End If
        End If
        TenderSDate.Text = TenderSDate.Text.Trim
        TenderEDate.Text = TenderEDate.Text.Trim
        If Errmag = "" Then
            If DateDiff(DateInterval.Day, CDate(TenderSDate.Text), CDate(TenderEDate.Text)) < 0 Then
                Errmag += "[投標日期迄日]必需大於[投標日期起始日]" & vbCrLf
            End If
        End If

        If Errmag = "" Then
            CheckData = True
        End If

    End Function

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim Errmsg As String = ""
        If CheckData(Errmsg) Then
            Select Case ViewState("Action")
                Case "ADD"
                    SAVE_Tender()
                Case "EDIT"
                    SAVE_Tender(ViewState("TSN"))
            End Select

        Else
            Common.MessageBox(Me, Errmsg)
        End If
    End Sub

    Sub SAVE_Tender(Optional ByVal TSN As Integer = 0)
        Dim OkErrorflag As Boolean = True
        Dim sqlCmd As SqlCommand
        Dim sqlStr As String = ""
        Select Case ViewState("Action")
            Case "ADD"
                sqlStr = "" & vbCrLf
                sqlStr += " INSERT INTO OB_Tender( years, TPlanID, PlanSN, TenderCName, TenderEName, Sponsor " & vbCrLf
                sqlStr += " , TenderSDate, TenderEDate, ReviewDate, ResolutionDate " & vbCrLf
                sqlStr += " , DistID, CreateAcct, CreateTime, ModifyAcct, ModifyTime) " & vbCrLf
                sqlStr += " VALUES( @Years, @TPlanID , @PlanSN , @TenderCName, @TenderEName, @Sponsor  " & vbCrLf
                sqlStr += " , @TenderSDate, @TenderEDate, @ReviewDate, @ResolutionDate  " & vbCrLf
                sqlStr += " , @DistID, @CreateAcct, getdate(), @ModifyAcct, getdate()) " & vbCrLf

            Case "EDIT"
                sqlStr = "" & vbCrLf
                sqlStr += " UPDATE  OB_Tender" & vbCrLf
                sqlStr += " SET  years=@years " & vbCrLf
                sqlStr += " , TPlanID=@TPlanID " & vbCrLf
                sqlStr += " , PlanSN=@PlanSN " & vbCrLf
                sqlStr += " , TenderCName=@TenderCName " & vbCrLf
                sqlStr += " , TenderEName=@TenderEName " & vbCrLf
                sqlStr += " , Sponsor=@Sponsor " & vbCrLf
                sqlStr += " , TenderSDate=@TenderSDate " & vbCrLf
                sqlStr += " , TenderEDate=@TenderEDate " & vbCrLf
                sqlStr += " , ReviewDate=@ReviewDate " & vbCrLf
                sqlStr += " , ResolutionDate=@ResolutionDate" & vbCrLf
                sqlStr += " , ModifyAcct=@ModifyAcct" & vbCrLf
                sqlStr += " , ModifyTime=getdate()" & vbCrLf
                sqlStr += " WHERE tsn=@TSN " & vbCrLf

        End Select

        Try
            TIMS.OpenDbConn(objconn)
            If rb2.Checked Then
                ViewState("PlanSN") = Get_OBPlanSN(PlanName.Text)
            Else
                ViewState("PlanSN") = Get_OBPlanSN(TPlanID.SelectedItem.Text)
            End If

            sqlCmd = New SqlCommand(sqlStr, objconn)
            With sqlCmd
                .Parameters.Clear()
                .Parameters.Add("Years", SqlDbType.Char, 4).Value = Me.ddlyears.SelectedValue
                If rb2.Checked Then
                    .Parameters.Add("TPlanID", SqlDbType.VarChar, 3).Value = Convert.DBNull
                Else
                    .Parameters.Add("TPlanID", SqlDbType.VarChar, 3).Value = TPlanID.SelectedValue
                End If
                .Parameters.Add("PlanSN", SqlDbType.Decimal).Value = ViewState("PlanSN")
                .Parameters.Add("TenderCName", SqlDbType.NVarChar, 20).Value = TenderCName.Text
                .Parameters.Add("TenderEName", SqlDbType.VarChar, 100).Value = TenderEName.Text

                If Sponsor.Text <> "" Then
                    .Parameters.Add("Sponsor", SqlDbType.NVarChar, 20).Value = Sponsor.Text
                Else
                    '可以不填寫主辦單位
                    .Parameters.Add("Sponsor", SqlDbType.NVarChar, 20).Value = Convert.DBNull
                End If

                .Parameters.Add("TenderSDate", SqlDbType.DateTime).Value = Common.FormatDate(TenderSDate.Text)

                ViewState("TenderEDate") = Common.FormatDate(TenderEDate.Text) & " " & TenderEDate2.Text
                .Parameters.Add("TenderEDate", SqlDbType.DateTime).Value = ViewState("TenderEDate")
                'Common.FormatDate(TenderEDate.Text)

                .Parameters.Add("ReviewDate", SqlDbType.DateTime).Value = Common.FormatDate(ReviewDate.Text)

                If ResolutionDate.Text <> "" Then
                    .Parameters.Add("ResolutionDate", SqlDbType.DateTime).Value = Common.FormatDate(ResolutionDate.Text)
                Else
                    .Parameters.Add("ResolutionDate", SqlDbType.DateTime).Value = Convert.DBNull
                End If

                If ViewState("Action") = "ADD" Then
                    .Parameters.Add("DistID", SqlDbType.VarChar, 3).Value = sm.UserInfo.DistID
                    .Parameters.Add("CreateAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                End If
                .Parameters.Add("ModifyAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                If ViewState("Action") = "EDIT" Then
                    .Parameters.Add("TSN", SqlDbType.Decimal).Value = TSN
                End If
                .ExecuteNonQuery()
            End With


        Catch ex As Exception
            OkErrorflag = False
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Finally
            '    objconn.Close()
            '    sqlCmd.Dispose()
        End Try

        Dim strScript As String = ""

        '若沒有錯誤狀況
        If OkErrorflag Then

            strScript = "<script language=""javascript"">" & vbCrLf

            If ViewState("Action") = "ADD" Then
                strScript += "alert('委外訓練資料查詢-新增成功!!');" & vbCrLf
            End If

            If ViewState("Action") = "EDIT" Then
                strScript += "alert('委外訓練資料查詢-修改成功!!');" & vbCrLf
            End If

            strScript += "location.href='OB_01_001.aspx?ID=" & Request("ID") & "';" & vbCrLf
            strScript += "</script>"

            Page.RegisterStartupScript("", strScript)

        End If
    End Sub

    Private Function Get_OBPlanSN(ByVal PlanName As String) As Integer
        Dim objSqlCmd As SqlCommand
        Dim sqlStr As String = ""
        ViewState("PlanSN") = Nothing
        If PlanName.Trim <> "" Then
            PlanName = PlanName.Trim
            sqlStr = "SELECT PlanSN FROM OB_Plan WHERE PlanName='" & Trim(PlanName) & "'"
            ViewState("PlanSN") = DbAccess.ExecuteScalar(sqlStr, objconn)
            If ViewState("PlanSN") Is Nothing Then
                sqlStr = "" & vbCrLf
                sqlStr += " INSERT INTO  OB_Plan( PlanName, CreateAcct, CreateTime)" & vbCrLf
                sqlStr += " VALUES( @PlanName, @CreateAcct, getdate() )  " & vbCrLf
                Try
                    objSqlCmd = New SqlCommand(sqlStr, objconn)
                    With objSqlCmd
                        .Parameters.Add("PlanName", SqlDbType.NVarChar, 20).Value = Trim(PlanName)
                        .Parameters.Add("CreateAcct", SqlDbType.VarChar, 15).Value = sm.UserInfo.UserID
                        .ExecuteNonQuery()

                        sqlStr = " select OB_Tender_TSN_SEQ.CURRVAL  "
                        objSqlCmd = New SqlCommand(sqlStr, objconn)
                        .Parameters.Clear()

                        ViewState("PlanSN") = .ExecuteScalar()
                    End With
                Catch ex As Exception
                    Common.MessageBox(Me, ex.ToString)
                    'Finally
                    '    objSqlCmd.Dispose()
                End Try
            End If
        End If

        Return ViewState("PlanSN")
    End Function

    Private Sub btnReturn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReturn.Click
        Dim MRqID As String = TIMS.ClearSQM(Request("ID"))
        TIMS.Utl_Redirect1(Me, "OB_01_001.aspx?ID=" & MRqID)
    End Sub
End Class
