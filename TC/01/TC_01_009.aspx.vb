Partial Class TC_01_009
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If TIMS.Get_PlanKind(Me, objconn) <> "1" Then  '1.自辦(不是自辦)
            Common.MessageBox(Me, "此功能只限自辦計劃使用")
            Button1.Enabled = False
        End If
        msg.Text = ""

        If Not IsPostBack Then
            DataGridTable.Visible = False
            TPeriod = TIMS.GET_HOURRAN(TPeriod, objconn, sm)
        End If

        Button1.Attributes("onclick") = "return CheckData();"
        Button2.Attributes("onclick") = "return SaveData();"
    End Sub

    '查詢SQL 
    Sub sSearch1()
        Dim dt As DataTable

        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql += " SELECT a.ClassCName ,a.Years ,a.OCID ,a.CyclType ,a.STDate ,a.FTDate ,a.SEnterDate ,a.FEnterDate ,a.CheckInDate " & vbCrLf
        sql += "  ,a.ExamDate ,a.TMID ,a.TPeriod ,a.TPropertyID ,a.STDate ,a.CJOB_UNKEY ,b.ClassID ,c.OCID CourseFlag " & vbCrLf
        sql += " FROM Class_ClassInfo a " & vbCrLf
        sql += " JOIN ID_Class b ON a.CLSID = b.CLSID " & vbCrLf
        sql += " LEFT JOIN (SELECT DISTINCT OCID FROM dbo.CLASS_SCHEDULE WHERE Formal = 'Y') c ON a.OCID = c.OCID " & vbCrLf
        sql += " WHERE a.IsClosed = 'N' " & vbCrLf
        sql += " AND a.NOTOPEN = @NOTOPEN " & vbCrLf
        sql += " AND a.RID = @RID " & vbCrLf
        sql += " AND a.PLANID = @PLANID " & vbCrLf

        If ClassCName.Text <> "" Then
            sql += " AND a.CLASSCNAME LIKE @CLASSCNAME " & vbCrLf
            parms.Add("CLASSCNAME", "%" & ClassCName.Text & "%")
        End If
        If CyclType.Text <> "" Then
            If CyclType.Text.Length = 1 Then CyclType.Text = "0" & CyclType.Text
            sql += " AND a.CYCLTYPE = @CYCLTYPE " & vbCrLf
            parms.Add("CYCLTYPE", CyclType.Text)
        End If
        If trainValue.Value <> "" Then
            sql += " AND a.TMID = @TMID " & vbCrLf
            parms.Add("TMID", trainValue.Value)
        End If
        If TPeriod.SelectedIndex <> 0 Then
            sql += " AND a.TPERIOD = @TPERIOD " & vbCrLf
            parms.Add("TPERIOD", TPeriod.SelectedValue)
        End If
        If TPropertyID.SelectedIndex <> 0 Then
            sql += " AND a.TPROPERTYID = @TPROPERTYID " & vbCrLf
            parms.Add("TPROPERTYID", TPropertyID.SelectedValue)
        End If
        If start_date.Text <> "" Then
            'sql += " AND a.STDate >= " & TIMS.to_date(start_date.Text) & vbCrLf
            sql += " AND a.STDATE >= @STDATE1 " & vbCrLf
            parms.Add("STDATE1", start_date.Text)
        End If
        If end_date.Text <> "" Then
            'sql += " AND a.STDate <= " & TIMS.to_date(end_date.Text) & vbCrLf '" & end_date.Text & "'" & vbCrLf
            sql += " AND a.STDATE <= @STDATE2 " & vbCrLf
            parms.Add("STDATE2", end_date.Text)
        End If
        If ClassState.SelectedIndex = 1 Then
            'Sql += " AND a.STDate <= GETDATE() " & vbCrLf
            'sql += " AND GETDATE() - a.STDate >= 0 " & vbCrLf
            sql += " AND DATEDIFF(DAY, a.STDate, GETDATE()) >= 0 " & vbCrLf
        ElseIf ClassState.SelectedIndex = 2 Then
            'Sql += " AND a.STDate >= GETDATE() " & vbCrLf
            'sql += " AND GETDATE() - a.STDate <= 0 " & vbCrLf
            sql += " AND DATEDIFF(DAY, a.STDate, GETDATE()) <= 0 " & vbCrLf
        End If
        If txtCJOB_NAME.Text <> "" Then
            sql += " AND a.CJOB_UNKEY = @CJOB_UNKEY " & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If

        sql += " AND a.OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE ACCOUNT = @ACCOUNT) " & vbCrLf
        sql += " ORDER BY b.ClassID, a.CyclType " & vbCrLf
        parms.Add("NOTOPEN", NotOpen.SelectedValue)
        parms.Add("RID", sm.UserInfo.RID)
        parms.Add("PLANID", sm.UserInfo.PlanID)
        parms.Add("ACCOUNT", sm.UserInfo.UserID)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputCheckBox = e.Item.FindControl("OCID")
                Dim hidExamDate As HtmlInputHidden = e.Item.FindControl("hidExamDate")
                Dim STDate As TextBox = e.Item.FindControl("STDate")
                Dim FTDate As TextBox = e.Item.FindControl("FTDate")
                Dim FTDate2 As HtmlInputHidden = e.Item.FindControl("FTDate2")
                Dim SEnterDate As TextBox = e.Item.FindControl("SEnterDate")
                Dim FEnterDate As TextBox = e.Item.FindControl("FEnterDate")
                Dim CheckInDate As TextBox = e.Item.FindControl("CheckInDate")
                Dim IMG1 As HtmlImage = e.Item.FindControl("IMG1")
                Dim IMG2 As HtmlImage = e.Item.FindControl("IMG2")
                Dim IMG3 As HtmlImage = e.Item.FindControl("IMG3")
                Dim IMG4 As HtmlImage = e.Item.FindControl("IMG4")
                Dim IMG5 As HtmlImage = e.Item.FindControl("IMG5")
                e.Item.CssClass = ""
                STDate.Enabled = False
                FTDate.Enabled = False
                SEnterDate.Enabled = False
                FEnterDate.Enabled = False
                CheckInDate.Enabled = False
                IMG1.Style("display") = "none"
                IMG2.Style("display") = "none"
                IMG3.Style("display") = "none"
                IMG4.Style("display") = "none"
                IMG5.Style("display") = "none"
                OCID.Attributes("onclick") = "SelectMyItem(this.checked," & e.Item.ItemIndex + 1 & ")"
                IMG1.Attributes("onclick") = "show_calendar('" & STDate.ClientID & "','','','CY/MM/DD');"
                IMG2.Attributes("onclick") = "show_calendar('" & FTDate.ClientID & "','','','CY/MM/DD');"
                IMG3.Attributes("onclick") = "show_calendar('" & SEnterDate.ClientID & "','','','CY/MM/DD');"
                IMG4.Attributes("onclick") = "show_calendar('" & FEnterDate.ClientID & "','','','CY/MM/DD');"
                IMG5.Attributes("onclick") = "show_calendar('" & CheckInDate.ClientID & "','','','CY/MM/DD');"
                OCID.Value = drv("OCID").ToString
                If IsNumeric(drv("CyclType")) Then
                    If Int(drv("CyclType")) <> 0 Then e.Item.Cells(1).Text += "第" & Int(drv("CyclType")) & "期"
                End If
                e.Item.Cells(2).Text = drv("Years") & "0" & drv("ClassID") & drv("CyclType")
                If drv("STDate").ToString <> "" Then STDate.Text = Common.FormatDate(drv("STDate"))
                If drv("FTDate").ToString <> "" Then
                    FTDate.Text = Common.FormatDate(drv("FTDate"))
                    FTDate2.Value = Common.FormatDate(drv("FTDate"))
                End If
                If drv("SEnterDate").ToString <> "" Then SEnterDate.Text = Common.FormatDate(drv("SEnterDate"))
                If drv("FEnterDate").ToString <> "" Then FEnterDate.Text = Common.FormatDate(drv("FEnterDate"))
                If drv("CheckInDate").ToString <> "" Then CheckInDate.Text = Common.FormatDate(drv("CheckInDate"))
                If drv("ExamDate").ToString <> "" Then hidExamDate.Value = Common.FormatDate(drv("ExamDate"))
                e.Item.Cells(8).Text = "是"
                IMG1.Visible = False
                If drv("CourseFlag").ToString = "" Then
                    e.Item.Cells(8).Text = "否"
                    IMG1.Visible = True
                Else
                    e.Item.Cells(8).ForeColor = Color.Red
                End If
        End Select
    End Sub

    'sUtl_CheckData1 儲存前先檢查輸入資料的正確性 (正式儲存檢核)
    ''' <summary>
    ''' 儲存前先檢查輸入資料的正確性
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function sUtl_CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True '沒有異常為True
        Errmsg = ""

        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    Errmsg &= "於處分日期起的期間，班級申請資料建檔不可儲存。"
            End Select
        End If
        If Errmsg <> "" Then Return False '不可儲存 '有錯誤訊息

        For Each item As DataGridItem In DataGrid1.Items
            Dim OCID As HtmlInputCheckBox = item.FindControl("OCID")
            Dim STDate As TextBox = item.FindControl("STDate")
            Dim FTDate As TextBox = item.FindControl("FTDate")
            Dim SEnterDate As TextBox = item.FindControl("SEnterDate")
            Dim FEnterDate As TextBox = item.FindControl("FEnterDate")
            Dim CheckInDate As TextBox = item.FindControl("CheckInDate")
            Dim hidExamDate As HtmlInputHidden = item.FindControl("hidExamDate")

            If OCID.Checked Then
                If SEnterDate.Text = "" Then
                    Errmsg += "班別資料「報名開始日期」為必填欄位!" & vbCrLf
                    Return False
                End If
                If FEnterDate.Text = "" Then
                    Errmsg += "班別資料「報名結束日期」為必填欄位!" & vbCrLf
                    Return False
                End If
                'If ExamDate.Text = "" Then
                '    Errmsg += "班別資料「甄試日期」為必填欄位!" & vbCrLf
                '    Return False
                'End If
                If (CDate(SEnterDate.Text) >= CDate(FEnterDate.Text)) Then
                    Errmsg += "班別資料[報名結束日期]必須大於[報名開始日期]!" & vbCrLf
                    Return False
                End If
                If (CDate(STDate.Text) <= CDate(FEnterDate.Text)) Then
                    Errmsg += "班別資料[訓練起日]必須大於[報名結束日期]!" & vbCrLf
                    Return False
                End If
                If hidExamDate.Value = "" Then '+2天
                    hidExamDate.Value = TIMS.Cdate3(DateAdd(DateInterval.Day, 2, CDate(FEnterDate.Text)))
                End If
                If DateDiff(DateInterval.Day, CDate(FEnterDate.Text), CDate(hidExamDate.Value)) < 2 Then
                    '「甄試日期」最快得安排於報名截止當日起2日後。
                    'Errmsg += "班別資料「甄試日期」最快得安排於「報名結束日期」當日起2日後!" & vbCrLf
                    '+2天
                    hidExamDate.Value = TIMS.Cdate3(DateAdd(DateInterval.Day, 2, CDate(FEnterDate.Text)))
                End If
                '甄試日期 --自辦在職 為必填
                If hidExamDate.Value <> "" Then
                    Do Until Not TIMS.Chk_HOLDATE(Hid_RID1.Value, hidExamDate.Value, objconn)
                        hidExamDate.Value = CDate(hidExamDate.Value).AddDays(1)
                    Loop
                End If
                If STDate.Text <> "" Then
                    If (CDate(hidExamDate.Value) > CDate(STDate.Text)) Then
                        Errmsg += "班別資料「甄試日期」必須小於或等於[訓練起日]!" & vbCrLf
                        Return False
                        'Common.MessageBox(Me, "[甄試日期]必須小於或等於[開訓日期]!")
                        'Exit Function
                    End If
                End If
            End If
        Next

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存 sql
    Sub SaveData1()
        Hid_RID1.Value = Convert.ToString(sm.UserInfo.RID).Substring(0, 1)
        Call TIMS.OpenDbConn(objconn)

        For Each item As DataGridItem In DataGrid1.Items
            Dim OCID As HtmlInputCheckBox = item.FindControl("OCID")
            Dim STDate As TextBox = item.FindControl("STDate")
            Dim FTDate As TextBox = item.FindControl("FTDate")
            Dim SEnterDate As TextBox = item.FindControl("SEnterDate")
            Dim FEnterDate As TextBox = item.FindControl("FEnterDate")
            Dim CheckInDate As TextBox = item.FindControl("CheckInDate")
            Dim hidExamDate As HtmlInputHidden = item.FindControl("hidExamDate")

            If OCID.Checked Then
                'Dim CourseFlag As Boolean = False
                'Dim ClassMsg As String = ""
                'Dim drSch As DataRow = Nothing
                Dim cPlnaID As String = "" '= Convert.ToString(dr("PlanID"))
                Dim cComIDNO As String = "" '= Convert.ToString(dr("ComIDNO"))
                Dim cSeqNO As String = "" '= Convert.ToString(dr("SeqNO"))
                Dim CourseFlag As Boolean = False

                Dim sql As String = ""
                sql = " SELECT DISTINCT 'x' FROM dbo.CLASS_SCHEDULE WHERE Formal = 'Y' AND OCID = '" & OCID.Value & "' "
                Dim drSch As DataRow = DbAccess.GetOneRow(sql, objconn)
                If Not drSch Is Nothing Then CourseFlag = True

                Dim da As SqlDataAdapter = Nothing
                sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID = '" & OCID.Value & "' "
                '2006/03/28 add conn by matt
                Dim dt As DataTable
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 0 Then
                    Dim dr As DataRow = dt.Rows(0)
                    'FEnterDate2  試著計算
                    Dim sFENTERDATE As String = FEnterDate.Text
                    Dim sEXAMDATE As String = hidExamDate.Value
                    Dim SS1 As String = ""
                    TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                    Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)

                    cPlnaID = Convert.ToString(dr("PlanID"))
                    cComIDNO = Convert.ToString(dr("ComIDNO"))
                    cSeqNO = Convert.ToString(dr("SeqNO"))
                    If STDate.Text <> "" Then dr("STDate") = TIMS.Cdate2(STDate.Text)
                    If FTDate.Text <> "" Then dr("FTDate") = TIMS.Cdate2(FTDate.Text)
                    If SEnterDate.Text <> "" Then dr("SEnterDate") = TIMS.Cdate2(SEnterDate.Text)
                    If FEnterDate.Text <> "" Then dr("FEnterDate") = TIMS.Cdate2(FEnterDate.Text)
                    If CheckInDate.Text <> "" Then dr("CheckInDate") = TIMS.Cdate2(CheckInDate.Text)
                    If hidExamDate.Value <> "" Then dr("ExamDate") = TIMS.Cdate2(hidExamDate.Value)
                    If sFENTERDATE2 <> "" Then dr("FEnterDate2") = CDate(sFENTERDATE2) 'TIMS.cdate2(sFENTERDATE2)
                    dr("LastState") = "M" 'M: 修改(最後異動狀態)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now()
                    DbAccess.UpdateDataTable(dt, da)
                End If

                'Plan_PlanInfo
                sql = " SELECT * FROM Plan_PlanInfo WHERE PlanID = '" & cPlnaID & "' AND ComIDNO = '" & cComIDNO & "' AND SeqNO = '" & cSeqNO & "' "
                '2006/03/28 add conn by matt
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 0 Then
                    Dim dr As DataRow = dt.Rows(0)
                    If STDate.Text <> "" Then dr("STDate") = TIMS.Cdate2(STDate.Text)
                    If FTDate.Text <> "" Then dr("FDDate") = TIMS.Cdate2(FTDate.Text)
                    If SEnterDate.Text <> "" Then dr("SEnterDate") = TIMS.Cdate2(SEnterDate.Text)
                    If FEnterDate.Text <> "" Then dr("FEnterDate") = TIMS.Cdate2(FEnterDate.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now()
                    DbAccess.UpdateDataTable(dt, da)
                End If

                'Class_StudentsOfClass
                sql = " SELECT * FROM Class_StudentsOfClass WHERE OCID = '" & OCID.Value & "' "
                '2006/03/28 add conn by matt
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        Dim dr As DataRow = dt.Rows(i)
                        If STDate.Text <> "" Then dr("OpenDate") = TIMS.Cdate2(STDate.Text)
                        If FTDate.Text <> "" Then dr("CloseDate") = TIMS.Cdate2(FTDate.Text)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, da)
                    Next
                End If
            End If
        Next
    End Sub

    '儲存
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call sUtl_CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
        Common.MessageBox(Me, "儲存成功!")
        Call sSearch1()
        'Button1_Click(sender, e)
    End Sub
End Class