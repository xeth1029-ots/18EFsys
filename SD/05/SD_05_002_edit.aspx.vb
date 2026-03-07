Partial Class SD_05_002_edit
    Inherits AuthBasePage

#Region "Function1"
    ''拉出資料庫外使用
    'Sub Get_KeyLeave(ByVal obj As DropDownList)
    '    With obj
    '        .DataSource = Key_Leave
    '        .DataTextField = "Name"
    '        .DataValueField = "LeaveID"
    '        .DataBind()
    '        .Items.Insert(0, New ListItem("請選擇", ""))
    '    End With
    'End Sub

    '依sm.UserInfo.PlanID取得FlexTurnoutKind
    'Function Get_FlexTurnoutKind(ByVal MyPage As Page) As String
    '    Dim Rst As String = ""
    '    ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(MyPage) Then Return Rst

    '    Dim sql As String = ""
    '    sql = "SELECT dbo.NVL(FlexTurnoutKind,0) FlexTurnoutKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
    '    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then Rst = Convert.ToString(dr("FlexTurnoutKind"))
    '    Return Rst
    'End Function
#End Region

    Dim dtLeaveF As DataTable
    Dim dtLeaveM As DataTable

    Dim Stud_Turnout As DataTable

    Dim vMessage As String = ""
    Dim rSOCID As String = ""
    Dim rOCID As String = ""
    Dim rProecess As String = ""
    Dim rLeaveID As String = ""
    Dim lrMsg As String = ""
    Dim ff3 As String = ""
    Const cst_生理假ID As String = "11"
    Const cst_不列入缺曠課 As Integer = 16 'Columns/Cells

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

        '取出假別鍵值----------------------------------------Start
        dtLeaveF = TIMS.GET_LEAVEdt(TIMS.cst_sex_F, objconn)
        dtLeaveM = TIMS.GET_LEAVEdt(TIMS.cst_sex_M, objconn)
        '取出假別鍵值----------------------------------------End

        If Not IsPostBack Then
            btnPrint1.Attributes("onclick") = "return w_printS1();"

            '20080610 Andy  彈性調整出缺勤 start

            '開放彈性調整出缺勤  null,不開放；1,開放
            Dim mydrPl As DataRow = TIMS.Get_FlexTurnoutKind_R(sm.UserInfo.PlanID, objconn)
            Dim flagOpen As Boolean = False
            If Convert.ToString(mydrPl("FlexTurnoutKind")) = "1" Then flagOpen = True
            Dim Plankind As String = mydrPl("Plankind")    '計畫種類 1,自辦； 2,委辦

            'Dim Plankind As String = ""
            'Plankind = TIMS.Get_PlanKind(Me, objconn) '計畫種類 1,自辦； 2,委辦
            '開放彈性調整出缺勤  null,不開放；1,開放
            'Dim flagOpen As Boolean = False
            'If Get_FlexTurnoutKind(Me) = "1" Then flagOpen = True '1,開放

            Select Case flagOpen
                Case True
                    Select Case Plankind
                        Case "1"
                            DataGrid1.Columns(cst_不列入缺曠課).Visible = True
                        Case "2"
                            DataGrid1.Columns(cst_不列入缺曠課).Visible = False
                        Case Else
                            DataGrid1.Columns(cst_不列入缺曠課).Visible = False
                    End Select
                Case False
                    DataGrid1.Columns(cst_不列入缺曠課).Visible = False
                Case Else
                    DataGrid1.Columns(cst_不列入缺曠課).Visible = False
            End Select
            '---------------20080610 Andy  彈性調整出缺勤 end---- 

            Me.ViewState("search") = Session("SearchStr")
            Session("SearchStr") = Nothing
            Call create()
        End If

        '儲存檢查
        Button5.Attributes("onclick") = "return check_data();"
    End Sub

    '取出班級資料 '取出個人資料
    Sub create()
        rSOCID = ""
        rOCID = ""
        rProecess = ""
        rLeaveID = ""
        If Convert.ToString(Request("SOCID")) <> "" Then rSOCID = Convert.ToString(Request("SOCID"))
        If Convert.ToString(Request("OCID")) <> "" Then rOCID = Convert.ToString(Request("OCID"))
        If Convert.ToString(Request("Proecess")) <> "" Then rProecess = Convert.ToString(Request("Proecess"))
        If Convert.ToString(Request("LeaveID")) <> "" Then rLeaveID = Convert.ToString(Request("LeaveID"))
        rSOCID = TIMS.ClearSQM(rSOCID)
        rOCID = TIMS.ClearSQM(rOCID)
        rProecess = TIMS.ClearSQM(rProecess)
        rLeaveID = TIMS.ClearSQM(rLeaveID)
        If rSOCID = "" OrElse rOCID = "" Then
            vMessage = "" & vbCrLf
            vMessage += "查詢資料失敗!!" & vbCrLf
            vMessage += "學員班級資料遺失，請重新查詢學員班級資料!!" & vbCrLf
            Common.MessageBox(Me, vMessage)
            Exit Sub
        End If


        '取出班級資料
        Dim drCC As DataRow = TIMS.GetOCIDDate(rOCID, objconn)
        If drCC Is Nothing Then
            vMessage = "" & vbCrLf
            vMessage += "查詢資料失敗!!" & vbCrLf
            vMessage += "學員班級資料遺失，請重新查詢學員班級資料!!" & vbCrLf
            Common.MessageBox(Me, vMessage)
            Exit Sub
        End If
        OrgName.Text = Convert.ToString(drCC("OrgName"))
        ClassCName.Text = Convert.ToString(drCC("ClassCName2"))
        Call DisabledColumms(Convert.ToString(drCC("TPeriod")))

        If rProecess = "view" Then
            Button5.Enabled = False
            TIMS.Tooltip(Button5, "僅供查看")
        End If

        '取出個人資料
        Dim parms As New Hashtable
        parms.Add("SOCID", rSOCID)
        '可能沒有假別，就使用全部 / '有假別，使用同假別資料
        If rLeaveID <> "" Then parms.Add("LEAVEID", rLeaveID)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WS1 AS (" & vbCrLf
        sql &= " SELECT cs.STUDSTATUS" & vbCrLf
        sql &= " ,cs.STUDENTID" & vbCrLf
        sql &= " ,cs.STUDID2" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf
        sql &= " ,cs.NAME" & vbCrLf
        sql &= " ,cs.SEX" & vbCrLf
        sql &= " ,cs.ISCLOSED" & vbCrLf
        sql &= " FROM V_STUDENTINFO cs" & vbCrLf
        sql &= " WHERE 1=1 AND cs.SOCID=@SOCID " & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT cs.STUDSTATUS,cs.STUDENTID,cs.STUDID2,cs.SOCID" & vbCrLf
        sql &= " ,cs.NAME,cs.SEX,cs.ISCLOSED" & vbCrLf
        sql &= " ,format(st.LEAVEDATE ,'yyyy/MM/dd') LEAVEDATE" & vbCrLf
        sql &= " ,st.SEQNO" & vbCrLf
        sql &= " ,st.LEAVEID,st.HOURS" & vbCrLf
        sql &= " ,st.C1,st.C2,st.C3,st.C4" & vbCrLf
        sql &= " ,st.C5,st.C6,st.C7,st.C8" & vbCrLf
        sql &= " ,st.C9,st.C10,st.C11,st.C12" & vbCrLf
        sql &= " ,st.TURNOUTIGNORE,st.TOTID,st.DASOURCE" & vbCrLf
        sql &= " ,k1.NAME LEAVENAME" & vbCrLf
        sql &= " FROM WS1 cs" & vbCrLf
        sql &= " JOIN STUD_TURNOUT st on st.SOCID=cs.SOCID" & vbCrLf
        sql &= " LEFT JOIN KEY_LEAVE k1 on k1.LEAVEID=st.LEAVEID" & vbCrLf
        If rLeaveID <> "" Then
            sql &= " WHERE 1=1" & vbCrLf
            sql &= " AND k1.LEAVEID=@LEAVEID" & vbCrLf
        End If
        sql &= " ORDER BY st.LEAVEDATE,st.LEAVEID" & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        '順序有差
        Dim parms2 As New Hashtable
        parms2.Add("SOCID", rSOCID)
        Dim sql2 As String = ""
        sql2 = ""
        sql2 &= " SELECT st.SOCID" & vbCrLf
        sql2 &= " ,format(st.LEAVEDATE ,'yyyy/MM/dd') LEAVEDATE" & vbCrLf
        sql2 &= " ,st.SEQNO" & vbCrLf
        sql2 &= " ,st.LEAVEID,st.HOURS" & vbCrLf
        sql2 &= " ,st.C1,st.C2,st.C3,st.C4" & vbCrLf
        sql2 &= " ,st.C5,st.C6,st.C7,st.C8" & vbCrLf
        sql2 &= " ,st.C9,st.C10,st.C11,st.C12" & vbCrLf
        sql2 &= " ,st.TURNOUTIGNORE,st.TOTID,st.DASOURCE" & vbCrLf
        sql2 &= " FROM STUD_TURNOUT st WHERE st.SOCID=@SOCID"
        Stud_Turnout = DbAccess.GetDataTable(sql2, objconn, parms2)

        Call BindDataGrid(dt)
    End Sub

    '建立DataGrid
    Sub BindDataGrid(ByRef dt As DataTable)
        Dim dr As DataRow = Nothing

        If dt.Rows.Count = 0 Then
            Session("SearchStr") = Me.ViewState("search")
            'Common.MessageBox(Me, "查無有效資料!!", "SD_05_002.aspx?ID=" & Request("ID") & "")
            'Button6_ServerClick(Button6, Nothing)
            'lrMsg = "查無有效資料!!"
            'sm.LastResultMessage = lrMsg
            'sm.RedirectUrlAfterBlock = "SD/05/SD_05_002.aspx?ID=" & Request("ID") & ""
            Common.RespWrite(Me, "<script> alert('查無有效資料!!');")
            Common.RespWrite(Me, "location.href='SD_05_002.aspx?ID=" & Request("ID") & "';</script>")
            Return
        End If

        '取得第1筆
        dr = dt.Rows(0)
        Name.Text = Convert.ToString(dr("Name"))
        StudentID.Text = Convert.ToString(dr("STUDID2"))
        StudStatus.Text = TIMS.GET_STUDSTATUS_N(dr("StudStatus"))
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '顯示/隱藏
    Sub DisabledColumms(ByVal num As String)
        For i As Integer = 3 To 14
            DataGrid1.Columns(i).Visible = False
        Next
        Select Case num
            Case "01"
                For i As Integer = 3 To 10
                    DataGrid1.Columns(i).Visible = True
                Next
            Case "02"
                For i As Integer = 11 To 14
                    DataGrid1.Columns(i).Visible = True
                Next
            Case Else
                For i As Integer = 3 To 14
                    DataGrid1.Columns(i).Visible = True
                Next
        End Select
    End Sub

    '儲存資料
    Sub SaveData1()
        Try
            'Try
            'Catch ex As Exception
            '    vMessage = "" & vbCrLf
            '    vMessage += "儲存失敗!!" & vbCrLf
            '    vMessage += "學員資料遺失，請重新查詢學員資料!!" & vbCrLf
            '    Common.MessageBox(Me, vMessage)
            '    Exit Sub
            'End Try
            'Dim dt As DataTable = Nothing
            'Dim conn As SqlConnection = DbAccess.GetConnection()

            Dim dr As DataRow = Nothing
            Dim da As SqlDataAdapter = Nothing
            Dim sql As String = ""
            '取得目前系統所有資料
            sql = "SELECT * FROM STUD_TURNOUT WHERE SOCID='" & rSOCID & "'"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, da, objconn)

            'Dim iSOCID As Integer = Val(rSOCID)
            For Each eItem As DataGridItem In DataGrid1.Items
                '檢核生理假
                vMessage = Chk_LeaveID11(eItem, Val(rSOCID), dt, 1)
                If vMessage <> "" Then
                    Common.MessageBox(Me, vMessage)
                    Exit Sub
                End If
            Next

            For Each eItem As DataGridItem In DataGrid1.Items
                Dim LeaveDate As TextBox = eItem.FindControl("LeaveDate")
                Dim LeaveID As DropDownList = eItem.FindControl("LeaveID")
                Dim C1 As HtmlInputCheckBox = eItem.FindControl("C1")
                Dim C2 As HtmlInputCheckBox = eItem.FindControl("C2")
                Dim C3 As HtmlInputCheckBox = eItem.FindControl("C3")
                Dim C4 As HtmlInputCheckBox = eItem.FindControl("C4")
                Dim C5 As HtmlInputCheckBox = eItem.FindControl("C5")
                Dim C6 As HtmlInputCheckBox = eItem.FindControl("C6")
                Dim C7 As HtmlInputCheckBox = eItem.FindControl("C7")
                Dim C8 As HtmlInputCheckBox = eItem.FindControl("C8")
                Dim C9 As HtmlInputCheckBox = eItem.FindControl("C9")
                Dim C10 As HtmlInputCheckBox = eItem.FindControl("C10")
                Dim C11 As HtmlInputCheckBox = eItem.FindControl("C11")
                Dim C12 As HtmlInputCheckBox = eItem.FindControl("C12")

                Dim TurnoutIgnore As HtmlInputCheckBox = eItem.FindControl("TurnoutIgnore")
                Dim hid_leaveT As HtmlInputHidden = eItem.FindControl("hid_leaveT")
                Dim Hid_SeqNo As HiddenField = eItem.FindControl("Hid_SeqNo")
                'Dim SeqNo As String = Hid_SeqNo.Value 'eItem.Cells(16).Text
                Dim iHours As Integer = 0

                If LeaveID.SelectedValue = "" Then
                    vMessage = "" & vbCrLf
                    vMessage += "儲存失敗!!" & vbCrLf
                    vMessage += "假別不可為空的資料!!" & vbCrLf
                    Common.MessageBox(Me, vMessage)
                    Exit Sub
                End If
                '原日期資料
                Dim ff3 As String = ""
                ff3 = "SOCID='" & rSOCID & "' and LeaveDate='" & hid_leaveT.Value & "' and SeqNo='" & Hid_SeqNo.Value & "'"
                If dt.Select(ff3).Length = 0 Then
                    vMessage = "" & vbCrLf
                    vMessage += "儲存失敗!!" & vbCrLf
                    vMessage += "查無可異動的資料!!" & vbCrLf
                    Common.MessageBox(Me, vMessage)
                    Exit Sub
                End If
                dr = dt.Select(ff3)(0) '應該會有1筆資料(UPDATE 使用)

                '改變日期
                Dim flagChangData1 As Boolean = False
                If LeaveDate.Text <> hid_leaveT.Value Then flagChangData1 = True

                If flagChangData1 Then
                    '判斷原資料是否存在(可能會重複資料日期)
                    LeaveDate.Text = TIMS.Cdate3(LeaveDate.Text)
                    ff3 = "SOCID='" & rSOCID & "' and LeaveDate='" & LeaveDate.Text & "'"
                    If dt.Select(ff3).Length = 0 Then
                        '不存在 可以做update
                        dr("LeaveDate") = TIMS.Cdate2(LeaveDate.Text)
                    Else
                        ''刪除原資料
                        'dr.Delete()
                        ''取得新資料
                        'dr = dt.Select("SOCID='" & rSOCID & "' and LeaveDate='" & LeaveDate.Text & "' and SeqNo='" & SeqNo & "'")(0)
                        'Me.RegisterStartupScript("1111", "<script>alert('儲存失敗\n,異動後的日期資料，與原資料日期有重複，請先使用刪除後再新增資料!!!');</script>")
                        '儲存失敗\n,異動後的日期資料，與原資料日期有重複，請先使用刪除後再新增資料!!!
                        'Dim vMessage As String
                        vMessage = "" & vbCrLf
                        vMessage += "儲存失敗!!" & vbCrLf
                        vMessage += "異動後的日期資料，與原資料日期有重複，請先使用刪除後再新增資料!!" & vbCrLf
                        Common.MessageBox(Me, vMessage)
                        Exit Sub
                    End If
                End If
                dr("LeaveID") = LeaveID.SelectedValue

                dr("C1") = If(C1.Checked, "Y", Convert.DBNull)
                dr("C2") = If(C2.Checked, "Y", Convert.DBNull)
                dr("C3") = If(C3.Checked, "Y", Convert.DBNull)
                dr("C4") = If(C4.Checked, "Y", Convert.DBNull)
                dr("C5") = If(C5.Checked, "Y", Convert.DBNull)
                dr("C6") = If(C6.Checked, "Y", Convert.DBNull)
                dr("C7") = If(C7.Checked, "Y", Convert.DBNull)
                dr("C8") = If(C8.Checked, "Y", Convert.DBNull)
                dr("C9") = If(C9.Checked, "Y", Convert.DBNull)
                dr("C10") = If(C10.Checked, "Y", Convert.DBNull)
                dr("C11") = If(C11.Checked, "Y", Convert.DBNull)
                dr("C12") = If(C12.Checked, "Y", Convert.DBNull)
                iHours += If(C1.Checked, 1, 0)
                iHours += If(C2.Checked, 1, 0)
                iHours += If(C3.Checked, 1, 0)
                iHours += If(C4.Checked, 1, 0)
                iHours += If(C5.Checked, 1, 0)
                iHours += If(C6.Checked, 1, 0)
                iHours += If(C7.Checked, 1, 0)
                iHours += If(C8.Checked, 1, 0)
                iHours += If(C9.Checked, 1, 0)
                iHours += If(C10.Checked, 1, 0)
                iHours += If(C11.Checked, 1, 0)
                iHours += If(C12.Checked, 1, 0)
                '20080616   andy 彈性調整出缺勤 start --------------------------
                dr("TurnoutIgnore") = If(TurnoutIgnore.Checked, 1, Convert.DBNull)
                '20080616   andy 彈性調整出缺勤  end --------------------------
                dr("Hours") = iHours
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now '有異動資料
            Next

            'Dim iSOCID As Integer = Val(rSOCID)
            For Each eItem As DataGridItem In DataGrid1.Items
                '檢核生理假
                vMessage = Chk_LeaveID11(eItem, Val(rSOCID), dt, 2)
                If vMessage <> "" Then
                    Common.MessageBox(Me, vMessage)
                    Exit Sub
                End If
            Next

            'For Each dr In dt.Select("LeaveID='" & rLeaveID & "'")
            '    '沒有異動的資料，便刪除(不做修改)
            '    If dr.RowState = DataRowState.Unchanged Then
            '        dr.Delete()
            '    End If
            'Next

            DbAccess.UpdateDataTable(dt, da)
            Me.RegisterStartupScript("successful", "<script>alert('儲存成功');</script>")

        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)

            vMessage = "" & vbCrLf
            vMessage += "儲存失敗!!" & vbCrLf
            vMessage += "儲存過程中有誤，請重新儲存!!" & vbCrLf
            vMessage += ex.ToString
            Common.MessageBox(Me, vMessage)
            Exit Sub
        End Try
        '20080616   andy

        '取出班級資料
        '取出個人資料
        Call create()
    End Sub

    '儲存資料
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        rSOCID = ""
        rLeaveID = ""
        If Convert.ToString(Request("SOCID")) <> "" Then rSOCID = Convert.ToString(Request("SOCID"))
        If Convert.ToString(Request("LeaveID")) <> "" Then rLeaveID = Convert.ToString(Request("LeaveID"))
        rSOCID = TIMS.ClearSQM(rSOCID)
        rLeaveID = TIMS.ClearSQM(rLeaveID)
        If rSOCID = "" Then
            vMessage = "" & vbCrLf
            vMessage += "儲存失敗!!" & vbCrLf
            vMessage += "學員資料遺失，請重新查詢學員資料!!" & vbCrLf
            Common.MessageBox(Me, vMessage)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        'TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
        'TIMS.SetMyValue(sCmdArg, "LeaveDate", TIMS.cdate3(drv("LeaveDate")))
        'TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
        'btn.CommandArgument = "SOCID='" & drv("SOCID") & "' and LeaveDate = " & TIMS.to_date(drv("LeaveDate")) & " and SeqNo='" & drv("SeqNo") & "'"

        Dim SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim LeaveDate As String = TIMS.GetMyValue(sCmdArg, "LeaveDate")
        Dim SeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        If SOCID = "" Then Exit Sub
        If LeaveDate = "" Then Exit Sub
        If SeqNo = "" Then Exit Sub

        Select Case LCase(e.CommandName)
            Case "del"
                Dim sql As String
                sql = ""
                sql &= " DELETE STUD_TURNOUT"
                sql &= " WHERE 1=1"
                sql &= " AND SOCID=" & SOCID
                sql &= " AND LeaveDate=" & TIMS.To_date(LeaveDate)
                sql &= " AND SeqNo=" & SeqNo
                DbAccess.ExecuteNonQuery(sql, objconn)

                Me.RegisterStartupScript("successful", "<script>alert('刪除成功!');</script>")

                '取出班級資料
                '取出個人資料
                Call create()
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim LeaveDate As TextBox = e.Item.FindControl("LeaveDate")
                Dim dlLeaveID As DropDownList = e.Item.FindControl("LeaveID")
                Dim hid_leaveT As HtmlInputHidden = e.Item.FindControl("hid_leaveT")
                Dim Hid_SeqNo As HiddenField = e.Item.FindControl("Hid_SeqNo")

                Dim Img1 As HtmlImage = e.Item.FindControl("Img1")
                Img1.Attributes("onclick") = "javascript:show_calendar('" & LeaveDate.ClientID & "','','','CY/MM/DD');"

                Dim C1 As HtmlInputCheckBox = e.Item.FindControl("C1")
                Dim C2 As HtmlInputCheckBox = e.Item.FindControl("C2")
                Dim C3 As HtmlInputCheckBox = e.Item.FindControl("C3")
                Dim C4 As HtmlInputCheckBox = e.Item.FindControl("C4")
                Dim C5 As HtmlInputCheckBox = e.Item.FindControl("C5")
                Dim C6 As HtmlInputCheckBox = e.Item.FindControl("C6")
                Dim C7 As HtmlInputCheckBox = e.Item.FindControl("C7")
                Dim C8 As HtmlInputCheckBox = e.Item.FindControl("C8")
                Dim C9 As HtmlInputCheckBox = e.Item.FindControl("C9")
                Dim C10 As HtmlInputCheckBox = e.Item.FindControl("C10")
                Dim C11 As HtmlInputCheckBox = e.Item.FindControl("C11")
                Dim C12 As HtmlInputCheckBox = e.Item.FindControl("C12")
                Dim btn As Button = e.Item.FindControl("Button1") '刪除鈕

                Dim TurnoutIgnore As HtmlInputCheckBox = e.Item.FindControl("TurnoutIgnore")
                Hid_SeqNo.Value = Convert.ToString(drv("SeqNo"))

                If drv("LeaveDate").ToString <> "" Then
                    LeaveDate.Text = Common.FormatDate(drv("LeaveDate"))
                    hid_leaveT.Value = Common.FormatDate(drv("LeaveDate"))
                End If
                'Call TIMS.GET_LEAVE(LeaveID, dtKeyLeave)
                dlLeaveID = TIMS.GET_LEAVE(dlLeaveID, dtLeaveF, dtLeaveM, Convert.ToString(drv("Sex")), 1)

                'Get_KeyLeave(LeaveID )
                Common.SetListItem(dlLeaveID, Convert.ToString(drv("LeaveID")))

                '20080617 ANDY
                '01: 病假'02: 事假'03: 公假'04: 曠課'05: 喪假'06: 遲到'07: 婚假'08: 陪產假
                TurnoutIgnore.Checked = If(IsDBNull(drv("TurnoutIgnore")), False, True)
                '彈性調整只開放婚假
                'Select Case LeaveID.SelectedValue
                '    Case "07"
                '        TurnoutIgnore.Visible = True
                '        TurnoutIgnore.Disabled = False
                '        If IsDBNull(drv("TurnoutIgnore")) Then
                '            TurnoutIgnore.Checked = False
                '        Else
                '            TurnoutIgnore.Checked = True
                '        End If

                '    Case Else
                '        TurnoutIgnore.Visible = False
                '        TurnoutIgnore.Disabled = True
                'End Select

                C1.Checked = If(drv("C1").ToString = "Y", True, False)
                C2.Checked = If(drv("C2").ToString = "Y", True, False)
                C3.Checked = If(drv("C3").ToString = "Y", True, False)
                C4.Checked = If(drv("C4").ToString = "Y", True, False)
                C5.Checked = If(drv("C5").ToString = "Y", True, False)
                C6.Checked = If(drv("C6").ToString = "Y", True, False)
                C7.Checked = If(drv("C7").ToString = "Y", True, False)
                C8.Checked = If(drv("C8").ToString = "Y", True, False)
                C9.Checked = If(drv("C9").ToString = "Y", True, False)
                C10.Checked = If(drv("C10").ToString = "Y", True, False)
                C11.Checked = If(drv("C11").ToString = "Y", True, False)
                C12.Checked = If(drv("C12").ToString = "Y", True, False)

                'sql語法 Stud_Turnout
                Dim sCmdArg As String = "" 'e.CommandArgument
                TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
                TIMS.SetMyValue(sCmdArg, "LeaveDate", TIMS.Cdate3(drv("LeaveDate")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                'btn.CommandArgument = "SOCID='" & drv("SOCID") & "' and LeaveDate = " & TIMS.to_date(drv("LeaveDate")) & " and SeqNo='" & drv("SeqNo") & "'"

                btn.CommandArgument = sCmdArg
                btn.Attributes("onclick") = "return confirm('確定要刪除" & Common.FormatDate(drv("LeaveDate")) & "這一筆資料?');"

                btn.Enabled = Button5.Enabled
                If Not btn.Enabled Then
                    TIMS.Tooltip(btn, "僅供查看")
                    LeaveDate.Enabled = False
                    Img1.Visible = False
                    dlLeaveID.Enabled = False
                End If

                Dim ff3 As String = String.Concat("SeqNo<>'", drv("SeqNo"), "' AND LEAVEDATE='", drv("LEAVEDATE"), "'")
                For Each dr As DataRow In Stud_Turnout.Select(ff3)
                    '20080616  andy 彈性調整出缺勤 start
                    TurnoutIgnore.Checked = If(IsDBNull(dr("TurnoutIgnore")), False, True)
                    '20080616  andy 彈性調整出缺勤 end
                    If Not C1.Disabled Then C1.Disabled = If(dr("C1").ToString = "Y", True, False)
                    If Not C2.Disabled Then C2.Disabled = If(dr("C2").ToString = "Y", True, False)
                    If Not C3.Disabled Then C3.Disabled = If(dr("C3").ToString = "Y", True, False)
                    If Not C4.Disabled Then C4.Disabled = If(dr("C4").ToString = "Y", True, False)
                    If Not C5.Disabled Then C5.Disabled = If(dr("C5").ToString = "Y", True, False)
                    If Not C6.Disabled Then C6.Disabled = If(dr("C6").ToString = "Y", True, False)
                    If Not C7.Disabled Then C7.Disabled = If(dr("C7").ToString = "Y", True, False)
                    If Not C8.Disabled Then C8.Disabled = If(dr("C8").ToString = "Y", True, False)
                    If Not C9.Disabled Then C9.Disabled = If(dr("C9").ToString = "Y", True, False)
                    If Not C10.Disabled Then C10.Disabled = If(dr("C10").ToString = "Y", True, False)
                    If Not C11.Disabled Then C11.Disabled = If(dr("C11").ToString = "Y", True, False)
                    If Not C12.Disabled Then C12.Disabled = If(dr("C12").ToString = "Y", True, False)
                Next
        End Select
    End Sub

    '檢核生理假(變更 前/後 都檢查)
    Function Chk_LeaveID11(ByVal eItem As DataGridItem,
                           ByVal iSOCID As Integer,
                           ByVal dtResult As DataTable, ByVal iType As Integer) As String

        Dim rst As String = ""
        'iType 1@DB面檢核  2:輸入資料面檢核
        Dim LeaveDate As TextBox = eItem.FindControl("LeaveDate")
        Dim LeaveID As DropDownList = eItem.FindControl("LeaveID")
        Dim Hid_SeqNo As HiddenField = eItem.FindControl("Hid_SeqNo")
        Dim STD_CNAME As String = Name.Text

        '1、假別增列「生理假」，每月至多可請生理假1日，但於訓練期間內請假未超過3日者(含第3日)，不併入病假計算(即生理假)；超過3日者，則視為病假計算。
        '2、假別統一排序為：公假、喪假、生理假、病假、事假、曠課，並刪除遲到、婚假、陪產假、未打卡、集會(週會)，範例如圖。
        '3、適用計畫代碼為：02、14、17、20、21、26、34、37、47、58、61、62、64、65、68。
        Select Case LeaveID.SelectedValue
            Case cst_生理假ID '生理假檢核
                Dim ff3 As String = ""
                ff3 = "SOCID='" & iSOCID & "' and LeaveID='" & cst_生理假ID & "'"
                '排除目前(日期)的筆數
                ff3 &= " and LeaveDate<>'" & TIMS.Cdate3(LeaveDate.Text) & "'"
                Select Case iType
                    Case 1 'DB面的檢核
                        If dtResult.Select(ff3).Length >= 3 Then
                            rst = STD_CNAME & "「生理假」超過3日者，請選擇病假!!"
                            Return rst
                        End If
                    Case 2 '輸入資料面檢核
                        If dtResult.Select(ff3).Length >= 3 Then
                            rst = STD_CNAME & "「生理假」超過3日者，請選擇病假!!"
                            Return rst
                        End If
                End Select
                If dtResult.Select(ff3).Length = 0 Then
                    Return rst '沒有請過生理假
                End If

                ff3 = "SOCID='" & iSOCID & "' and LeaveID='" & cst_生理假ID & "'"
                For Each dr As DataRow In dtResult.Select(ff3)
                    Dim chkFlag As Boolean = True
                    If TIMS.Cdate3(dr("LeaveDate")) = TIMS.Cdate3(LeaveDate.Text) _
                        AndAlso Convert.ToString(dr("SeqNo")) = Hid_SeqNo.Value Then

                        chkFlag = False '同日/同號(排除核檢)
                    End If
                    If chkFlag Then
                        Dim sLeaveYM1 As String = TIMS.Cdate5(dr("LeaveDate"))
                        Dim sLeaveYM2 As String = TIMS.Cdate5(LeaveDate.Text)
                        If sLeaveYM1 = sLeaveYM2 Then '同月
                            rst = STD_CNAME & "每月至多可請「生理假」1日，請重新選擇!!"
                            Return rst
                        End If
                    End If
                Next

        End Select

        Return rst
    End Function

    '回上一頁
    Private Sub Button6_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.ServerClick
        Session("SearchStr") = Me.ViewState("search")
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_05_002.aspx?ID=" & Request("ID") & "")
    End Sub

End Class

