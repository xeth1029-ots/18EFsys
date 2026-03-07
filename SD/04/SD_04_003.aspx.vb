Partial Class SD_04_003
    Inherits AuthBasePage

    Dim strWeek() As String = {"", "星期日", "星期一", "星期二", "星期三", "星期四", "星期五", "星期六"}
    Dim dtCourse As DataTable
    Dim dtTeacher As DataTable

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Dim rqClassID As String = TIMS.ClearSQM(Request("ClassID"))
        Dim rqFormal As String = TIMS.ClearSQM(Request("Formal"))
        Dim rqSingle As String = TIMS.ClearSQM(Request("Single"))

        If Not Page.IsPostBack Then
            Call ChangeCmdVisible(False) '班級審核確認開放控制

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            If rqClassID <> "" Then
                Dim drCC As DataRow = TIMS.GetOCIDDate(rqClassID, objconn)
                If Not drCC Is Nothing Then
                    center.Text = Convert.ToString(drCC("OrgName"))
                    RIDValue.Value = Convert.ToString(drCC("RID"))
                End If
            End If

            Call sSearch1() '列出班級

            Button1.Visible = False
            If rqClassID <> "" Then
                Select Case rqFormal
                    Case "Y"
                        ShowMode.SelectedIndex = 0
                        Common.SetListItem(OCID1, rqClassID)
                    Case "N"
                        ShowMode.SelectedIndex = 1
                        Common.SetListItem(OCID2, rqClassID)
                End Select
                GetCourseData(rqClassID)
                Button1.Visible = True
            End If

            SingleValue.Value = "N"
            Button1.Text = "回全期排課"
            If rqSingle = "Y" Then
                SingleValue.Value = "Y"
                Button1.Text = "回單月排課"
                ShowMode.SelectedIndex = 0
                ShowMode.Enabled = False
            End If
        End If

        Page.RegisterStartupScript("00000", "<script>ChangeMode();</script>")
        ShowMode.Attributes("onclick") = "ChangeMode();"
        OCID1.Attributes("onchange") = "if(this.selectedIndex==0){return false;}"
        OCID2.Attributes("onchange") = "if(this.selectedIndex==0){return false;}"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=Button2');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button2.Style("display") = "none"
        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button2")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Sub GetCourseData(ByVal OCID As String)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Btn_TR1.Visible = True '審核
        Btn_TR2.Visible = True '審核確認
        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = " SELECT * FROM Course_CourseInfo WHERE OrgID IN (SELECT OrgID FROM Auth_Relship WHERE RID = '" & RIDValue.Value & "') "
        dtCourse = DbAccess.GetDataTable(sql, objconn)

        sql = " SELECT * FROM Teach_TeacherInfo WHERE RID IN (SELECT RID FROM Auth_Relship WHERE OrgID IN (SELECT OrgID FROM Auth_Relship WHERE RID = '" & RIDValue.Value & "')) "
        dtTeacher = DbAccess.GetDataTable(sql, objconn)

        Dim TPeriod As String = ""
        '班級審核確認開放控制
        Call ChangeCmdVisible(False)

        If OCID = "" Then Exit Sub

        TIMS.IsAllDateCheck(Me, OCID, "ShowMsg", objconn) '檢查是否為全日制,若是全日制檢查是否符合規則
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        TPeriod = Convert.ToString(drCC("TPeriod")) '上課時段

        Dim IsVisible As Boolean = False
        IsVisible = Not (TPeriod = "01")
        For i As Integer = 10 To 13
            Me.List_Class.Columns(i).Visible = IsVisible
        Next

        IsVisible = Not (TPeriod = "02")
        For i As Integer = 2 To 9
            Me.List_Class.Columns(i).Visible = IsVisible
        Next

        '排課種類(0:正式 1:預排)
        Select Case ShowMode.SelectedIndex
            Case 0  '排課種類(0:正式 1:預排)
                sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = " & OCID & " AND Formal = 'Y' ORDER BY SchoolDate "
                'LID:帳號階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                'OrgLevel:機構階層【0:署(局) 1:分署(中心) 2:委訓(縣市政府) 3：委訓】
                '補助地方政府計畫 LID:【0:署(局) 1:分署(中心) 2:(補助地方政府) 3:委訓】

                'If sm.UserInfo.LID <= 1 OrElse (sm.UserInfo.TPlanID = "17" AndAlso sm.UserInfo.OrgLevel <= 2) Then
                'End If

                '可使用審核確認鈕
                Dim flagCanConfirmat As Boolean = False
                If sm.UserInfo.LID <= 1 Then flagCanConfirmat = True
                If Not flagCanConfirmat AndAlso TIMS.Cst_TPlanID17AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagCanConfirmat = TIMS.ChkRelship232(Me, objconn)
                If flagCanConfirmat Then
                    If Not TIMS.Chk_ClassSchVerify(OCID, objconn) Then
                        ChangeCmdVisible(True, True, "此班級尚未審核確認")
                    Else
                        ChangeCmdVisible(True, False, "此班級已審核確認")
                    End If
                End If
            Case 1  '排課種類(0:正式 1:預排)
                '1:預排
                sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = " & OCID & " AND Formal = 'N' ORDER BY SchoolDate "
            Case Else
                '未選擇排課種類(選擇錯誤!!)
                Exit Sub
        End Select

        Dim objtable As New DataTable
        objtable.Load(DbAccess.GetReader(sql, objconn))

        List_Class.Visible = True
        Me.List_Class.DataSource = objtable
        Me.List_Class.DataBind()
    End Sub

    'List_Class
    Private Sub List_Class_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles List_Class.ItemDataBound
        Dim drv As DataRowView = CType(e.Item.DataItem, DataRowView)
        'Dim Leslable, Tealable, Romlabel, Tealable2 As Label
        '<%--<%#  DbAccess.GetRocDateValue(Container.DataItem("SchoolDate"))%>--%>
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                For i As Integer = 1 To 12
                    Dim i2 As Integer = i + 12
                    Dim i3 As Integer = i + 12 + 12
                    'e.Item.Cells(i).Text = strWeek(Weekday(drv("SchoolDate")))
                    Dim LSchoolDate As Label = e.Item.FindControl("LSchoolDate")
                    Dim LWeekday As Label = e.Item.FindControl("LWeekday")
                    Dim Leslable As Label = e.Item.FindControl("Les" & i)
                    Dim Tealable As Label = e.Item.FindControl("Tea" & i)
                    Dim Tealable2 As Label = e.Item.FindControl("Tea" & i2)
                    Dim Romlabel As Label = e.Item.FindControl("Rom" & i)

                    LSchoolDate.Text = DbAccess.GetRocDateValue(drv("SchoolDate"))
                    LWeekday.Text = strWeek(Weekday(drv("SchoolDate")))
                    Leslable.Text = TIMS.Get_CourseName(drv("Class" & i), dtCourse, objconn)
                    Tealable.Text = TIMS.Get_TeacherName(drv("Teacher" & i), dtTeacher)
                    'add by nick 060525 加入第二位老師
                    Tealable2.Text = ""
                    If Convert.ToString(drv("Teacher" & i2)) <> "" Then Tealable2.Text = TIMS.Get_TeacherName(drv("Teacher" & i2), dtTeacher)
                    If Convert.ToString(drv("Teacher" & i3)) <> "" Then
                        If Tealable2.Text <> "" Then Tealable2.Text &= ","
                        Tealable2.Text &= TIMS.Get_TeacherName(drv("Teacher" & i3), dtTeacher)
                    End If
                    'end
                    Romlabel.Text = Convert.ToString(drv("Room" & i))
                Next
                If Convert.ToString(drv("Vacation")) = "Y" Then
                    For i As Integer = 2 To e.Item.Cells.Count - 2
                        Dim sTmp As String = Trim(e.Item.Cells(i).Text)
                        If sTmp <> "" Then
                            sTmp &= "(假日)"
                            e.Item.Cells(i).Text = sTmp
                        End If
                    Next
                End If
        End Select
    End Sub

    '搜尋預排課程
    Private Sub OCID2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OCID2.SelectedIndexChanged
        ChangeCmdVisible(False) '班級審核確認開放控制
        List_Class.Visible = False
        If OCID2.SelectedIndex <> 0 Then
            OCID1.SelectedIndex = -1
            GetCourseData(OCID2.SelectedValue)
        End If
    End Sub

    '搜尋正式排課
    Private Sub OCID1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OCID1.SelectedIndexChanged
        ChangeCmdVisible(False) '班級審核確認開放控制
        List_Class.Visible = False
        If OCID1.SelectedIndex <> 0 Then
            OCID2.SelectedIndex = -1
            GetCourseData(OCID1.SelectedValue)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim rqClassID As String = TIMS.ClearSQM(Request("ClassID"))
        Dim rqMMID As String = TIMS.ClearSQM(Request("ID"))
        If SingleValue.Value = "Y" Then
            Session("_OCID") = rqClassID
            TIMS.Utl_Redirect1(Me, "SD_04_002.aspx?ID=" & rqMMID)
        Else
            Session("_OCID") = rqClassID
            TIMS.Utl_Redirect1(Me, "SD_04_001.aspx?ID=" & rqMMID)
        End If
    End Sub

    '班級搜尋 SQL
    Function sUtl_GetClassClassInfo(ByRef strErrmsg As String, ByRef sql As String) As DataTable
        Dim Rst As DataTable = Nothing
        Dim PlanKind As Integer = TIMS.Get_PlanKind(Me, objconn)
        'sql1 = "SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'PlanKind = DbAccess.ExecuteScalar(sql1, objconn)

        Dim sDistID As String = sm.UserInfo.DistID
        If sm.UserInfo.DistID = "000" Then sDistID = TIMS.Get_RID2DistID(RIDValue.Value, objconn)

        Dim sql1 As String = ""
        sql1 = ""
        sql1 &= " SELECT OCID, Formal FROM VIEW_SCHEDULE "
        sql1 &= " WHERE 1=1"
        sql1 &= " AND TPlanID = '" & sm.UserInfo.TPlanID & "' "
        sql1 &= " AND Years = '" & sm.UserInfo.Years & "' "
        sql1 &= " AND DistID = '" & sDistID & "' "
        Dim dtVSchedule As New DataTable
        dtVSchedule.Load(DbAccess.GetReader(sql1, objconn))

        Try
            If sm.UserInfo.RID = "A" Then
                sql = "" & vbCrLf
                sql &= " SELECT a.OCID ,a.ClassCName ,a.CyclType ,a.LevelType " & vbCrLf
                sql &= "  ,a.IsSuccess ,a.RID " & vbCrLf
                sql &= "  ,c.ClassID " & vbCrLf
                sql &= "  ,ip.PlanID ,ip.TPlanID ,ip.Years " & vbCrLf
                sql &= "  ,'N' Formal " & vbCrLf
                sql &= " FROM Class_ClassInfo a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN ID_Class c ON a.CLSID = c.CLSID " & vbCrLf
                sql &= " JOIN ID_Plan ip ON ip.PlanID = a.PlanID" & vbCrLf
                sql &= " WHERE a.IsSuccess = 'Y'" & vbCrLf
                sql &= "  AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
                sql &= "  AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "'" & vbCrLf
                sql &= "  AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                sql &= "  AND ip.DistID = '" & sDistID & "' " & vbCrLf
                'sql &= " AND a.RID = '" & RIDValue.Value & "' AND a.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                'sql &= " AND a.OCID IN (SELECT OCID FROM Auth_AccRWClass WHERE Account = '" & sm.UserInfo.UserID & "') " & vbCrLf
                sql &= " ORDER BY c.ClassID ,a.CyclType " & vbCrLf
                Rst = DbAccess.GetDataTable(sql, objconn)
                If Rst.Rows.Count > 0 Then
                    '取得Formal
                    For Each dr As DataRow In Rst.Rows
                        If dtVSchedule.Select("OCID='" & dr("OCID") & "'").Length > 0 Then
                            dr("Formal") = dtVSchedule.Select("OCID='" & dr("OCID") & "'")(0)("Formal")
                        Else
                            Call dr.Delete()
                        End If
                    Next
                    Rst.AcceptChanges()
                End If

            Else
                If PlanKind = 1 Then
                    sql = ""
                    sql &= " SELECT a.OCID ,a.ClassCName ,a.CyclType ,a.LevelType " & vbCrLf
                    sql &= "  ,a.IsSuccess ,a.RID " & vbCrLf
                    sql &= "  ,c.ClassID " & vbCrLf
                    sql &= "  ,ip.PlanID ,ip.TPlanID ,ip.Years " & vbCrLf
                    sql &= "  ,'N' Formal " & vbCrLf
                    sql &= " FROM Class_ClassInfo a " & vbCrLf
                    sql &= " JOIN ID_Class c ON a.CLSID = c.CLSID " & vbCrLf
                    sql &= " JOIN ID_Plan ip ON a.PlanID = ip.PlanID " & vbCrLf
                    sql &= " WHERE a.IsSuccess = 'Y' " & vbCrLf
                    sql &= "  AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
                    sql &= "  AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                    sql &= "  AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                    sql &= " ORDER BY c.ClassID ,a.CyclType " & vbCrLf
                    'Dim oDt As DataTable
                    Rst = DbAccess.GetDataTable(sql, objconn)

                    Dim oDt2 As DataTable
                    Dim oSql As String = ""
                    If Rst.Rows.Count > 0 Then
                        '取得Formal
                        For Each dr As DataRow In Rst.Rows
                            If dtVSchedule.Select("OCID='" & dr("OCID") & "'").Length > 0 Then
                                dr("Formal") = dtVSchedule.Select("OCID='" & dr("OCID") & "'")(0)("Formal")
                            Else
                                Call dr.Delete()
                            End If
                        Next
                        Rst.AcceptChanges()

                        '自辦過濾使用者
                        oSql = " SELECT OCID FROM Auth_AccRWClass WHERE Account = '" & sm.UserInfo.UserID & "' "
                        oDt2 = DbAccess.GetDataTable(oSql, objconn)
                        If oDt2.Rows.Count > 0 Then
                            For Each dr As DataRow In Rst.Rows
                                If oDt2.Select("OCID='" & dr("OCID") & "'").Length = 0 Then Call dr.Delete()
                            Next
                            Rst.AcceptChanges()
                        Else
                            For Each dr As DataRow In Rst.Rows
                                Call dr.Delete()
                            Next
                            Rst.AcceptChanges()
                        End If
                    End If

#Region "(No Use)"

                    ''自辦過濾使用者
                    ''Dim oDt As DataTable
                    'Dim oDt2 As DataTable
                    'Dim oSql As String = ""
                    'oSql = "" & vbCrLf
                    'oSql += " select  a.OCID,a.ClassCName,a.CyclType,a.LevelType " & vbCrLf
                    'oSql += " ,a.Formal " & vbCrLf
                    'If sm.UserInfo.Years <= 2010 Then
                    '    oSql += " from view_ClassinfoSch2 a" & vbCrLf
                    'Else
                    '    oSql += " from view_ClassinfoSch2_2011 a" & vbCrLf
                    'End If
                    ''sql += " join Auth_AccRWClass a2 on a2.ocid =a.ocid" & vbCrLf
                    'oSql += " where 1=1" & vbCrLf
                    'oSql += " AND a.RID='" & RIDValue.Value & "' " & vbCrLf
                    'oSql += " AND a.TPlanID='" & sm.UserInfo.TPlanID & "' " & vbCrLf
                    'oSql += " AND a.Years='" & sm.UserInfo.Years & "'" & vbCrLf
                    ''sql += " AND a2.Account='" & sm.UserInfo.UserID & "'" & vbCrLf
                    'oSql += " Order By a.ClassID,a.CyclType" & vbCrLf
                    'oDt = DbAccess.GetDataTable(oSql)

#End Region
                Else
                    sql = ""
                    sql &= " SELECT a.OCID ,a.ClassCName ,a.CyclType ,a.LevelType " & vbCrLf
                    sql &= " ,a.IsSuccess ,a.RID " & vbCrLf
                    sql &= " ,c.ClassID " & vbCrLf
                    sql &= " ,ip.PlanID ,ip.TPlanID ,ip.Years " & vbCrLf
                    sql &= " ,'N' Formal " & vbCrLf
                    sql &= " FROM Class_ClassInfo a" & vbCrLf
                    sql &= " JOIN ID_Class c ON a.CLSID = c.CLSID " & vbCrLf
                    sql &= " JOIN ID_Plan ip ON a.PlanID = ip.PlanID " & vbCrLf
                    sql &= " WHERE a.IsSuccess = 'Y' " & vbCrLf
                    sql &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
                    sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
                    sql &= " AND ip.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
                    sql &= " ORDER BY c.ClassID ,a.CyclType " & vbCrLf
                    'Dim oDt As DataTable
                    Rst = DbAccess.GetDataTable(sql, objconn)
                    If Rst.Rows.Count > 0 Then
                        '取得Formal
                        For Each dr As DataRow In Rst.Rows
                            If dtVSchedule.Select("OCID='" & dr("OCID") & "'").Length > 0 Then
                                dr("Formal") = dtVSchedule.Select("OCID='" & dr("OCID") & "'")(0)("Formal")
                            Else
                                Call dr.Delete()
                            End If
                        Next
                        Rst.AcceptChanges()
                    End If
                End If
            End If

            'If sql <> "" Then Rst = DbAccess.GetDataTable(sql)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "查詢 已超過連接逾時的設定。" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
            strErrmsg += "ex.ToString:" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            'Common.MessageBox(Me, strErrmsg)
            Call TIMS.WriteTraceLog(Me, ex, ex.Message)
        End Try

        Return Rst
    End Function

    '班級搜尋[SQL]
    Sub sSearch1()
        Dim sql As String = ""
        Dim strErrmsg As String = ""
        Dim dt As DataTable

        dt = sUtl_GetClassClassInfo(strErrmsg, sql)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Try
                strErrmsg += "sql:" & vbCrLf
                strErrmsg += sql & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Exit Sub
        End If

        'Dim dr As DataRow
        OCID1.Items.Clear()
        OCID2.Items.Clear()
        For Each dr As DataRow In dt.Rows
            Dim ClassName As String = CStr(dr("ClassCName")) & "第" & CStr(dr("CyclType")) & "期"
            If TIMS.Chk_ClassSchVerify(dr("OCID"), objconn) Then ClassName += "(已審核)"
            If dr("Formal").ToString = "Y" Then
                OCID1.Items.Add(New ListItem(ClassName, dr("OCID")))
            ElseIf dr("Formal").ToString = "N" Then
                OCID2.Items.Add(New ListItem(ClassName, dr("OCID")))
            End If
        Next

        ''980224 fix 只有一筆資料則將TIMS.cst_ddl_PleaseChoose3拿掉
        'If dt.Rows.Count <> 1 Then
        '    OCID1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        '    OCID2.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        'End If

        Dim ocid1flag As Boolean = False
        Dim ocid2flag As Boolean = False
        '980224 fix 只有一筆資料則將TIMS.cst_ddl_PleaseChoose3拿掉
        If dt.Rows.Count > 0 Then
            If OCID1.Items.Count <> 1 Then
                OCID1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Else
                ocid1flag = True
            End If
            If OCID2.Items.Count <> 1 Then
                OCID2.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Else
                ocid2flag = True
            End If

            If ocid1flag AndAlso ocid2flag Then
                ocid1flag = False
                ocid2flag = False
                OCID1.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                OCID2.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End If
        End If

        If ocid1flag Then
            '班級審核確認開放控制
            ChangeCmdVisible(False)
            '正式1筆
            If OCID1.SelectedValue <> "" Then
                Common.SetListItem(ShowMode, "1")
                OCID2.Style("display") = "none"
                GetCourseData(OCID1.SelectedValue)
            End If
        End If

        If ocid2flag Then
            '班級審核確認開放控制
            ChangeCmdVisible(False)
            '預排1筆
            If OCID2.SelectedValue <> "" Then
                Common.SetListItem(ShowMode, "2")
                OCID1.Style("display") = "none"
                GetCourseData(OCID2.SelectedValue)
            End If
        End If
    End Sub

    '班級搜尋
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call sSearch1()
    End Sub

    '按下審核確認
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) _
        Handles Button3.Click, Button3B.Click, Button4.Click, Button4B.Click
        Const Cst_Y As String = "ResultY" '審核確認
        Const Cst_N As String = "ResultN" '取消審核

        Dim dt As DataTable
        Dim dr As DataRow
        Dim sql As String
        'Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim da As SqlDataAdapter = Nothing
        Dim Scmd As Button = sender

        Select Case Scmd.CommandName
            Case Cst_Y
                If OCID1.SelectedValue <> "" Then
                    sql = " SELECT * FROM CLASS_SCHVERIFY WHERE OCID = '" & OCID1.SelectedValue & "' "
                    dt = DbAccess.GetDataTable(sql, da, objconn)
                    If dt.Rows.Count = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("OCID") = OCID1.SelectedValue
                        dr("AppResult") = "Y"
                    Else
                        dr = dt.Rows(0)
                    End If
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da)
                    Common.MessageBox(Me, "審核確認!!")
                End If
                ChangeCmdVisible(False) '班級審核確認開放控制 (按下後不顯示取消審核)
            Case Cst_N
                If OCID1.SelectedValue <> "" Then
                    sql = " SELECT * FROM CLASS_SCHVERIFY WHERE OCID = '" & OCID1.SelectedValue & "' "
                    dt = DbAccess.GetDataTable(sql, da, objconn)
                    If dt.Rows.Count > 0 Then dt.Rows(0).Delete()
                    DbAccess.UpdateDataTable(dt, da)
                    Common.MessageBox(Me, "取消審核!!")
                End If
                ChangeCmdVisible(True) '班級審核確認開放控制  (按下後顯示審核確認)
        End Select

        If OCID1.SelectedIndex <> 0 Then
            OCID2.SelectedIndex = -1
            GetCourseData(OCID1.SelectedValue)
        End If
    End Sub

    '班級審核確認開放控制
    Sub ChangeCmdVisible(ByVal cmdVisible As Boolean, Optional ByVal cmdEnabled As Boolean = True, Optional ByVal tipMSG As String = "")
        '審核確認(顯示)
        Button3.Visible = cmdVisible
        Button3B.Visible = Button3.Visible
        Button3.Enabled = cmdEnabled
        Button3B.Enabled = cmdEnabled

        If Button3.Visible = True And Button3.Enabled = False Then
            '取消審核(顯示)
            Button4.Visible = True
            Button4B.Visible = True
            Button4.Enabled = True
            Button4B.Enabled = True
        Else
            '取消審核(不顯示)
            Button4.Enabled = False
            Button4B.Enabled = False
            Button4.Visible = False
            Button4B.Visible = False
        End If

        If tipMSG <> "" Then
            Button3.ToolTip = tipMSG
            Button3B.ToolTip = tipMSG
            Button4.ToolTip = tipMSG
            Button4B.ToolTip = tipMSG
        End If
    End Sub

    Private Sub center_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles center.TextChanged
        List_Class.Visible = False
        Btn_TR1.Visible = False '審核
        Btn_TR2.Visible = False '審核確認
    End Sub

    Sub sSearch2()
        List_Class.Visible = False
        '排課種類(0:正式 1:預排)
        Select Case ShowMode.SelectedIndex
            Case 0
                ChangeCmdVisible(False) '班級審核確認開放控制
                If OCID1.SelectedIndex <> 0 Then
                    OCID2.SelectedIndex = -1
                    GetCourseData(OCID1.SelectedValue)
                End If
            Case 1
                ChangeCmdVisible(False) '班級審核確認開放控制
                If OCID2.SelectedIndex <> 0 Then
                    OCID1.SelectedIndex = -1
                    GetCourseData(OCID2.SelectedValue)
                End If
        End Select
    End Sub

    Protected Sub ShowMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ShowMode.SelectedIndexChanged
        Call sSearch2()
    End Sub

    Protected Sub btnSchAct3_Click(sender As Object, e As EventArgs) Handles btnSchAct3.Click
        Call sSearch2()
    End Sub
End Class