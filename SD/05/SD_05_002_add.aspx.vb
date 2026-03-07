Partial Class SD_05_002_add
    Inherits AuthBasePage

    'Dim sql As String = ""
    'dbo.fn_GET_TURNOUT
    Dim dtLeaveF As DataTable 'dtLeaveF = TIMS.GET_LEAVEdt(Me, TIMS.cst_sex_F, objconn)
    Dim dtLeaveM As DataTable 'dtLeaveM = TIMS.GET_LEAVEdt(Me, TIMS.cst_sex_M, objconn)

    Dim Days1 As Integer = 0
    Dim Days2 As Integer = 0

    Dim rqProecess As String = "" 'TIMS.ClearSQM(Request("Proecess"))

    Dim ff3 As String = ""
    Const cst_生理假ID As String = "11"
    Const cst_SearchStr As String = "SearchStr"

    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '取出設定天數檔
        Call TIMS.Get_SysDays(Days1, Days2, objconn)
        '取出假別鍵值----------------------------------------Start
        dtLeaveF = TIMS.GET_LEAVEdt(TIMS.cst_sex_F, objconn)
        dtLeaveM = TIMS.GET_LEAVEdt(TIMS.cst_sex_M, objconn)
        '取出假別鍵值----------------------------------------End

        '檢查系統參數是否有設定------------------------------Start
        HidItemVar1.Value = TIMS.GetGlobalVar(Me, "4", "1", objconn)
        HidItemVar2.Value = TIMS.GetGlobalVar(Me, "4", "2", objconn)
        If HidItemVar1.Value = "" _
            OrElse HidItemVar2.Value = "" Then
            Common.MessageBox(Me, "警告!系統參數尚未設定出缺勤警示")
        End If
        '檢查系統參數是否有設定------------------------------End

        If Session(cst_SearchStr) IsNot Nothing Then
            Session(cst_SearchStr) = Session(cst_SearchStr)
            ViewState(cst_SearchStr) = Session(cst_SearchStr)
        End If

        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        rqProecess = TIMS.ClearSQM(Request("Proecess"))
        If rqProecess <> "add" Then
            Common.RespWrite(Me, "<script language='javascript'>")
            Common.RespWrite(Me, "alert('傳入參數有誤!!:" & rqProecess & "');")
            Common.RespWrite(Me, "location.href='SD_05_002.aspx?ID=" & s_MRqID & "';")
            Common.RespWrite(Me, "</script>")
            Exit Sub
        End If
        'HidOCID1.Value = TIMS.ClearSQM(HidOCID1.Value)

        If Not IsPostBack Then
            msg.Text = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Call sCreate1(0)
        End If

        'If MySessionSet.Value = "1" Then
        '    MySessionSet.Value = 0
        '    Call SetSession()
        'End If
        Button1.Attributes("onclick") = "javascript:return chkdata();"
        Button3.Attributes("onclick") = "javascript:choose_class();"

        'Button1.Enabled = False '儲存
        'If au.blnCanAdds Then Button1.Enabled = True '儲存
    End Sub

    Sub sCreate1(ByVal iType As Integer)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'Dim drCC As DataRow = Nothing
        Select Case iType
            Case 0
                HidOCID1.Value = TIMS.sUtl_GetRqValue(Me, "OCID1", HidOCID1.Value)
                If HidOCID1.Value = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
            Case Else
                If OCIDValue1.Value = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
                HidOCID1.Value = OCIDValue1.Value
        End Select
        Dim drCC As DataRow = TIMS.GetOCIDDate(HidOCID1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        center.Text = Convert.ToString(drCC("OrgName"))
        RIDValue.Value = Convert.ToString(drCC("RID"))
        OCIDValue1.Value = Convert.ToString(drCC("OCID")) 'HidOCID1.Value
        OCID1.Text = Convert.ToString(drCC("ClassCName2"))
        TMIDValue1.Value = Convert.ToString(drCC("TMID"))
        TMID1.Text = Convert.ToString(drCC("TrainName2"))
        LeaveDate.Text = TIMS.Cdate3(Now.Date)
        StDate.Text = TIMS.Cdate3(drCC("STDATE"))
        FtDate.Text = TIMS.Cdate3(drCC("FTDATE"))
        HidThours.Value = Val(drCC("Thours"))

        '取出該班的總訓練時數--------------------------------End
        center.Enabled = True
        TMID1.Enabled = True
        OCID1.Enabled = True
        LeaveDate.Enabled = True
        Button1.Visible = False

        '載入
        'ddl_sSearch2 = TIMS.Get_SOCID_DDL(ddl_sSearch2, objconn, OCIDValue1.Value)
        ddl_sSearch2 = TIMS.Get_SOCID_DDL2(ddl_sSearch2, objconn, OCIDValue1.Value, 10)

        '學生資訊
        Call GetAllClass()

        '---------------20090304 Andy  彈性調整出缺勤 start

        'Dim sql As String = ""
        'sql = " SELECT Plankind,FlexTurnoutKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'Dim mydrPl As DataRow = DbAccess.GetOneRow(sql, objconn)

        Dim mydrPl As DataRow = Nothing
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0"
                '開放彈性調整出缺勤  null,不開放；1,開放
                mydrPl = TIMS.Get_FlexTurnoutKind_R(drCC("PlanID"), objconn)
            Case Else
                '開放彈性調整出缺勤  null,不開放；1,開放
                mydrPl = TIMS.Get_FlexTurnoutKind_R(sm.UserInfo.PlanID, objconn)
        End Select
        Dim blnOpen As Boolean = False
        If Convert.ToString(mydrPl("FlexTurnoutKind")) = "1" Then blnOpen = True
        '計畫種類 1,自辦； 2,委辦
        Dim Plankind As Integer = mydrPl("Plankind")

        '不列入缺曠課
        Const cst_不列入缺曠課 As Integer = 22
        DataGrid1.Columns(cst_不列入缺曠課).Visible = False
        Select Case blnOpen
            Case True
                Select Case Plankind
                    Case 1
                        DataGrid1.Columns(cst_不列入缺曠課).Visible = True
                    Case 2
                        DataGrid1.Columns(cst_不列入缺曠課).Visible = False
                    Case Else
                        DataGrid1.Columns(cst_不列入缺曠課).Visible = False
                End Select
            Case False
                DataGrid1.Columns(cst_不列入缺曠課).Visible = False
        End Select
        '---------------20090304  Andy  彈性調整出缺勤 end 

    End Sub

    'session 
    Sub SetSession()

        Session(cst_SearchStr) = If(Session(cst_SearchStr) IsNot Nothing, Session(cst_SearchStr), ViewState(cst_SearchStr))
        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Dim url1 As String = "SD_05_002.aspx?ID=" & s_MRqID 'Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
        'Page.RegisterStartupScript("Redirect", "<script>location.href='SD_05_002.aspx?ID=" & Request("ID") & "'</script>")
    End Sub

    '學生資訊及其他
    Sub GetAllClass()
        'ddl_sSearch2.Items.Clear()
        If OCIDValue1.Value = "" Then Exit Sub
        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        'Dim sql As String = ""
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        'sql = "SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & OCIDValue1.Value & "'"
        'dr = DbAccess.GetOneRow(sql, objconn)

        If Not flag_test Then
            If Convert.ToString(drC("IsClosed")) = "Y" Then
                Select Case sm.UserInfo.RoleID
                    Case 0, 1
                        '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                        '暫時先改這樣,以後還會再改
                        If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                            If DateDiff(DateInterval.Day, drC("FTDate"), Now.Date) > 75 Then
                                Common.MessageBox(Me, "此班級已經結訓,無法新增!")
                                Button1.Enabled = False
                                StudentTable.Visible = False
                                Exit Sub
                            End If
                        Else
                            If DateDiff(DateInterval.Day, drC("FTDate"), Now.Date) > Days2 Then
                                Common.MessageBox(Me, "此班級已經結訓,無法新增!")
                                Button1.Enabled = False
                                StudentTable.Visible = False
                                Exit Sub
                            End If
                        End If

                    Case Else
                        '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                        '暫時先改這樣,以後還會再改
                        If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                            If DateDiff(DateInterval.Day, drC("FTDate"), Now.Date) > 60 Then
                                Common.MessageBox(Me, "此班級已經結訓,無法新增!")
                                Button1.Enabled = False
                                StudentTable.Visible = False
                                Exit Sub
                            End If
                        Else
                            If DateDiff(DateInterval.Day, drC("FTDate"), Now.Date) > Days1 Then
                                Common.MessageBox(Me, "此班級已經結訓,無法新增!")
                                Button1.Enabled = False
                                StudentTable.Visible = False
                                Exit Sub
                            End If
                        End If

                End Select
            End If

        End If

        For i As Integer = 3 To 18
            DataGrid1.Columns(i).Visible = True
        Next
        Select Case Convert.ToString(drC("TPeriod"))
            Case "01"
                For i As Integer = 15 To 18
                    DataGrid1.Columns(i).Visible = False
                Next
            Case "02"
                For i As Integer = 7 To 14
                    DataGrid1.Columns(i).Visible = False
                Next
            Case Else
        End Select

        Dim flag_can_continue As Boolean = False 'false:異常停止
        'Dim s_SOCID As String = TIMS.ClearSQM(ddl_sSearch2.SelectedValue)
        Dim s_STUDID2 As String = TIMS.ClearSQM(ddl_sSearch2.SelectedValue)
        Dim STUDID2N As String = ""
        Dim STUDID2M As String = ""

        If s_STUDID2 <> "" AndAlso s_STUDID2 <> "ALL" Then
            If s_STUDID2.Split(".").Length > 1 Then
                STUDID2N = s_STUDID2.Split(".")(0) '小值
                STUDID2M = s_STUDID2.Split(".")(1) '大值
                flag_can_continue = True 'ture:正常繼續
            End If
        End If
        If s_STUDID2 = "ALL" Then
            flag_can_continue = True 'ture:正常繼續
        End If
        If Not flag_can_continue Then
            msg.Text = "查無此班的學生資料!"
            Common.MessageBox(Me, "查無此班的學生資料!")
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.StudStatus" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.STUDENTID) STUDID2" & vbCrLf
        sql &= " ,b.Name" & vbCrLf
        sql &= " ,b.Sex" & vbCrLf
        sql &= " ,0 LeaveID" & vbCrLf
        sql &= " ,e.TPeriod" & vbCrLf
        sql &= " ,a.RejectTDate1" & vbCrLf
        sql &= " ,a.RejectTDate2 " & vbCrLf
        sql &= " ,dbo.FN_GET_TURNOUT(a.SOCID) total" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO e" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS a on a.OCID = e.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b on b.SID=a.SID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and a.OCID='" & OCIDValue1.Value & "' " & vbCrLf
        'sql &= " and e.OCID=:'" & OCIDValue1.Value & "'" & vbCrLf
        sql &= " and e.OCID=@OCID" & vbCrLf
        'If s_SOCID <> "" AndAlso s_SOCID <> "ALL" Then
        '    sql &= " and a.SOCID=@SOCID" & vbCrLf
        'End If
        If s_STUDID2 <> "" AndAlso s_STUDID2 <> "ALL" Then
            sql &= " and dbo.FN_CSTUDID2(a.STUDENTID) >= @STUDID2N" & vbCrLf
            sql &= " and dbo.FN_CSTUDID2(a.STUDENTID) <= @STUDID2M" & vbCrLf
        End If
        sql &= " Order By a.StudentID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            'If s_SOCID <> "" AndAlso s_SOCID <> "ALL" Then
            '    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = s_SOCID
            'End If
            If s_STUDID2 <> "" AndAlso s_STUDID2 <> "ALL" Then
                .Parameters.Add("STUDID2N", SqlDbType.VarChar).Value = STUDID2N
                .Parameters.Add("STUDID2M", SqlDbType.VarChar).Value = STUDID2M
            End If
            dt.Load(.ExecuteReader())
        End With
        'da = New SqlDataAdapter(sql, objconn)
        'da.Fill(ds, "student")
        'Dim dt As DataTable = ds.Tables("student")

        StudentTable.Visible = False
        Button1.Visible = False
        msg.Text = "查無此班的學生資料!"
        If dt.Rows.Count > 0 Then
            scrollDiv.Attributes.Add("class", "DivHeight")

            StudentTable.Visible = True
            Button1.Visible = True
            msg.Text = ""

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "SOCID"
            DataGrid1.DataBind()
        End If

    End Sub

    '儲存1
    Function UpdateTable(ByVal drItems As DataGridItem,
                         ByVal dt As DataTable, ByVal SOCID As Integer, ByVal SeqNo As Integer,
                         ByVal dtTemp As DataTable, ByRef sMesbox As String) As DataTable

        '抓取所有假別是否有勾選
        'Dim dr As DataRow = Nothing
        Dim drTemp As DataRow = Nothing
        Dim LeaveID As DropDownList = drItems.FindControl("LeaveID")
        Dim C1 As HtmlInputCheckBox = drItems.FindControl("C1")
        Dim C2 As HtmlInputCheckBox = drItems.FindControl("C2")
        Dim C3 As HtmlInputCheckBox = drItems.FindControl("C3")
        Dim C4 As HtmlInputCheckBox = drItems.FindControl("C4")
        Dim C5 As HtmlInputCheckBox = drItems.FindControl("C5")
        Dim C6 As HtmlInputCheckBox = drItems.FindControl("C6")
        Dim C7 As HtmlInputCheckBox = drItems.FindControl("C7")
        Dim C8 As HtmlInputCheckBox = drItems.FindControl("C8")
        Dim C9 As HtmlInputCheckBox = drItems.FindControl("C9")
        Dim C10 As HtmlInputCheckBox = drItems.FindControl("C10")
        Dim C11 As HtmlInputCheckBox = drItems.FindControl("C11")
        Dim C12 As HtmlInputCheckBox = drItems.FindControl("C12")
        Dim STD_CNAME As String = Convert.ToString(drItems.Cells(1).Text)

        Dim flag As Boolean = False '正常加總為false/ 有狀況為true(停止更動)

        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("SOCID") = SOCID
        'dr("LeaveDate") = Common.FormatDate(LeaveDate.Text)
        dr("LeaveDate") = CDate(LeaveDate.Text)
        dr("SeqNo") = SeqNo
        dr("LeaveID") = LeaveID.SelectedValue

        'Dim sMesbox As String = ""
        Dim iHours As Integer = 0
        flag = False
        If C1.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C1")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第一節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C1") = "Y"
                iHours += 1
            End If
        Else
            dr("C1") = Convert.DBNull
        End If

        flag = False
        If C2.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C2")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第二節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C2") = "Y"
                iHours += 1
            End If
        Else
            dr("C2") = Convert.DBNull
        End If

        flag = False
        If C3.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C3")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第三節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C3") = "Y"
                iHours += 1
            End If
        Else
            dr("C3") = Convert.DBNull
        End If

        flag = False
        If C4.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C4")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第四節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C4") = "Y"
                iHours += 1
            End If
        Else
            dr("C4") = Convert.DBNull
        End If

        flag = False
        If C5.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C5")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第五節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C5") = "Y"
                iHours += 1
            End If
        Else
            dr("C5") = Convert.DBNull
        End If

        flag = False
        If C6.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C6")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第六節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C6") = "Y"
                iHours += 1
            End If
        Else
            dr("C6") = Convert.DBNull
        End If

        flag = False
        If C7.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C7")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第七節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C7") = "Y"
                iHours += 1
            End If
        Else
            dr("C7") = Convert.DBNull
        End If

        flag = False
        If C8.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C8")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第八節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C8") = "Y"
                iHours += 1
            End If
        Else
            dr("C8") = Convert.DBNull
        End If

        flag = False
        If C9.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C9")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第九節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C9") = "Y"
                iHours += 1
            End If
        Else
            dr("C9") = Convert.DBNull
        End If

        flag = False
        If C10.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C10")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第十節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C10") = "Y"
                iHours += 1
            End If
        Else
            dr("C10") = Convert.DBNull
        End If

        flag = False
        If C11.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C11")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第十一節已有請假過"
                End If
            Next
            If flag = False Then
                dr("C11") = "Y"
                iHours += 1
            End If
        Else
            dr("C11") = Convert.DBNull
        End If

        flag = False
        If C12.Checked Then
            For Each drTemp In dtTemp.Select("SOCID='" & SOCID & "'")
                If dr("SeqNo") <> drTemp("SeqNo") And Not IsDBNull(drTemp("C12")) Then
                    If sMesbox <> "" Then sMesbox &= ","
                    sMesbox &= "第十二節已有請假過"
                    flag = True
                End If
            Next
            If flag = False Then
                dr("C12") = "Y"
                iHours += 1
            End If
        Else
            dr("C12") = Convert.DBNull
        End If

        'If iHours = 0 Then
        '    sMesbox &= "至少要勾選1節課程\n"
        '    flag = True
        'End If

        dr("Hours") = iHours
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        '20090305  andy 彈性調整出缺勤 
        '-----------------------------start 
        Dim TurnoutIgnore As HtmlInputCheckBox = drItems.FindControl("TurnoutIgnore")
        If Not TurnoutIgnore.Checked Then
            dr("TurnoutIgnore") = Convert.DBNull
        Else
            dr("TurnoutIgnore") = 1 '有勾選，不列入曠課
        End If
        '-----------------------------end
        If sMesbox <> "" Then
            sMesbox = STD_CNAME & ":" & sMesbox
        End If

        Return dt
    End Function

    '新增(檢核)
    Function ChkAddData1(ByRef errMsg As String,
                         ByRef dtResult As DataTable, ByRef dtTemp As DataTable) As Boolean

        'dtTemp
        Dim rst As Boolean = True
        If rqProecess <> "add" Then
            errMsg = "功能參數有誤，請重新操作!"
            Return False
        End If

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'Dim dtTemp As DataTable
        If Convert.ToString(OCIDValue1.Value).Trim = "" Then
            errMsg = "請重新選擇有效班級!"
            Return False
        End If

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            errMsg = "請重新選擇有效班級!"
            Return False
        End If

        'Dim sql As String = ""
        'sql = "SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & OCIDValue1.Value & "'"
        'dr = DbAccess.GetOneRow(sql, objconn)
        'If dr Is Nothing Then
        '    Common.MessageBox(Me, "請重新選擇有效班級!")
        '    Exit Function
        'End If
        If Convert.ToString(drCC("STDate")) = "" OrElse Convert.ToString(drCC("FTDate")) = "" Then
            errMsg = "該班的開、結訓日期有異常 ,無法新增!"
            Return False
        End If

        LeaveDate.Text = TIMS.ClearSQM(LeaveDate.Text)
        'LeaveDate.Text = Convert.ToString(LeaveDate.Text).Trim
        If Convert.ToString(LeaveDate.Text) = "" Then
            errMsg = "點名日期為必填資料!"
            Return False
            'Common.MessageBox(Me, "點名日期為必填資料!")
            'Exit Function
        End If
        If Not IsDate(Convert.ToString(LeaveDate.Text).Trim) Then
            errMsg = "點名日期必須為日期格式!"
            Return False
            'Common.MessageBox(Me, "點名日期必須為日期格式!")
            'Exit Function
        End If
        Try
            LeaveDate.Text = Common.FormatDate(LeaveDate.Text)
        Catch ex As Exception
            'Common.MessageBox(Me, "點名日期異常!!")
            'Exit Function
            errMsg = "點名日期異常!!"
            Return False
        End Try

        If DateDiff(DateInterval.Day, CDate(LeaveDate.Text), CDate(drCC("STDate"))) > 0 Then
            'Common.MessageBox(Me, "請假日期應大於開班日期" & drCC("STDate") & ",無法新增!")
            'Exit Function
            errMsg = "請假日期應大於開班日期" & TIMS.Cdate3(drCC("STDate")) & ",無法新增!"
            Return False
        End If
        If DateDiff(DateInterval.Day, CDate(drCC("FTDate")), CDate(LeaveDate.Text)) > 0 Then
            'Common.MessageBox(Me, "請假日期應小於結訓日期" & drCC("FTDate") & ",無法新增!")
            'Exit Function
            errMsg = "請假日期應小於結訓日期" & TIMS.Cdate3(drCC("FTDate")) & ",無法新增!"
            Return False
        End If

        '假如處裡狀態是新增狀態時
        '先取出假別資料 (某班)
        Dim sMesbox As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing

        Dim sql As String = ""

        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (select SOCID from CLASS_STUDENTSOFCLASS x where x.OCID ='" & OCIDValue1.Value & "')" & vbCrLf
        sql &= " SELECT m.SOCID" & vbCrLf
        sql &= " ,CONVERT(varchar, m.LeaveDate, 111) LeaveDate" & vbCrLf
        sql &= " ,m.LeaveID" & vbCrLf
        sql &= " ,m.SeqNo" & vbCrLf
        sql &= " FROM STUD_TURNOUT m" & vbCrLf
        sql &= " WHERE EXISTS (select 'x' from WC1 x where x.SOCID =m.SOCID)" & vbCrLf
        dtResult = DbAccess.GetDataTable(sql, objconn)

        '找出歷史資料 (學員某日)
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (select SOCID from CLASS_STUDENTSOFCLASS x where x.OCID ='" & OCIDValue1.Value & "')" & vbCrLf
        sql &= " SELECT * FROM STUD_TURNOUT"
        sql &= " WHERE SOCID IN (SELECT SOCID FROM WC1)"
        sql &= " AND LEAVEDATE= " & TIMS.To_date(LeaveDate.Text) '某天請假
        dtTemp = DbAccess.GetDataTable(sql, objconn)

        '建立空的Table
        sql = "SELECT * FROM STUD_TURNOUT WHERE 1<>1"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each drItems As DataGridItem In DataGrid1.Items  '尋找每一行的DataGrid資料
            Dim myLeaveID As DropDownList = drItems.FindControl("LeaveID")
            If myLeaveID.SelectedIndex <> 0 AndAlso myLeaveID.SelectedValue <> "" Then
                '要是假別不是請選擇的狀態，則更新資料
                '先抓出目前SeqNo的最大值
                Dim iSeqNo As Integer = 0
                Dim iSOCID As Integer = DataGrid1.DataKeys(drItems.ItemIndex)
                ff3 = "SOCID='" & iSOCID & "' AND LeaveDate='" & LeaveDate.Text & "'"
                Dim drResultA() As DataRow = dtResult.Select(ff3, "SeqNo Desc")
                If drResultA.Length = 0 Then
                    iSeqNo = 1
                Else
                    iSeqNo = drResultA(0)("SeqNo") + 1 '最後序號加1
                End If

                '檢核生理假
                sMesbox = Chk_LeaveID11(drItems, iSOCID, TIMS.Cdate3(LeaveDate.Text), iSeqNo, dtResult)
                If sMesbox <> "" Then
                    errMsg = sMesbox
                    Return False
                End If

                '某日的檢核
                dt = UpdateTable(drItems, dt, iSOCID, iSeqNo, dtTemp, sMesbox)
                If sMesbox <> "" Then
                    errMsg = sMesbox
                    Return False
                End If

            End If
        Next

        If errMsg <> "" Then rst = False
        Return rst
    End Function

    '檢核生理假(新增後的檢核) (SD_05_002_add/AC_03_002_add)
    Function Chk_LeaveID11(ByVal drItems As DataGridItem,
                           ByVal iSOCID As Integer, ByVal sLeaveDate As String, ByVal iSeqNo As Integer,
                           ByVal dtResult As DataTable) As String
        Dim rst As String = ""

        Dim LeaveID As DropDownList = drItems.FindControl("LeaveID")
        Dim STD_CNAME As String = Convert.ToString(drItems.Cells(1).Text)

        '1、假別增列「生理假」，每月至多可請生理假1日，但於訓練期間內請假未超過3日者(含第3日)，不併入病假計算(即生理假)；超過3日者，則視為病假計算。
        '2、假別統一排序為：公假、喪假、生理假、病假、事假、曠課，並刪除遲到、婚假、陪產假、未打卡、集會(週會)，範例如圖。
        '3、適用計畫代碼為：02、14、17、20、21、26、34、37、47、58、61、62、64、65、68。
        Select Case LeaveID.SelectedValue
            Case cst_生理假ID '生理假檢核
                Dim ff3 As String = "SOCID='" & iSOCID & "' and LeaveID='" & cst_生理假ID & "'"
                If dtResult.Select(ff3).Length >= 3 Then
                    rst = STD_CNAME & "「生理假」超過3日者，請選擇病假!!"
                    Return rst
                End If
                If dtResult.Select(ff3).Length = 0 Then
                    Return rst '沒有請過生理假
                End If
                For Each dr As DataRow In dtResult.Select(ff3)
                    Dim chkFlag As Boolean = True
                    If TIMS.Cdate3(dr("LeaveDate")) = TIMS.Cdate3(sLeaveDate) _
                        AndAlso Val(dr("SeqNo")) = iSeqNo Then
                        chkFlag = False '同日/同號(排除核檢)
                    End If
                    Dim sLeaveYM1 As String = TIMS.Cdate5(dr("LeaveDate"))
                    Dim sLeaveYM2 As String = TIMS.Cdate5(sLeaveDate)
                    If sLeaveYM1 = sLeaveYM2 Then '同月
                        rst = STD_CNAME & "每月至多可請「生理假」1日，請重新選擇!!"
                        Return rst
                    End If
                Next
        End Select

        Return rst
    End Function

    '儲存(SQL)
    Sub SaveData1(ByRef dtResult As DataTable, ByRef dtTemp As DataTable)
        'Dim sMesbox As String = ""
        'Dim SeqNo As Integer = 0
        'Dim dr As DataRow = Nothing
        'Dim da As SqlDataAdapter = Nothing
        'Dim SOCID As Integer = 0
        '假如處裡狀態是新增狀態時
        '先取出假別資料 (某班)
        'sql = ""
        'sql &= " SELECT m.SOCID, CONVERT(varchar, m.LeaveDate, 111) LeaveDate, m.SeqNo"
        'sql &= " FROM Stud_Turnout m "
        'sql &= " WHERE EXISTS (select 'x' from Class_StudentsOfClass x"
        'sql &= "    where x.SOCID =m.SOCID AND x.OCID ='" & OCIDValue1.Value & "')"
        'dtResult = DbAccess.GetDataTable(sql, objconn)
        '找出歷史資料 (學員某日)
        'sql = ""
        'sql &= " SELECT * FROM Stud_Turnout WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "')"
        'sql &= " and LeaveDate= " & TIMS.to_date(LeaveDate.Text) '某天請假
        'dtTemp = DbAccess.GetDataTable(sql, objconn)

        Dim sMesbox As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim sql As String = ""

        If rqProecess = "add" Then
            '建立空的Table
            sql = "SELECT * FROM STUD_TURNOUT WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, objconn)

            For Each drItems As DataGridItem In DataGrid1.Items '尋找每一行的DataGrid資料
                Dim myLeaveID As DropDownList = drItems.FindControl("LeaveID")
                If myLeaveID.SelectedIndex <> 0 AndAlso myLeaveID.SelectedValue <> "" Then '要是假別不是請選擇的狀態，則更新資料
                    '先抓出目前SeqNo的最大值
                    Dim iSeqNo As Integer = 0
                    Dim iSOCID As Integer = DataGrid1.DataKeys(drItems.ItemIndex)
                    Dim drResultA() As DataRow = dtResult.Select("SOCID='" & iSOCID & "' and LeaveDate='" & LeaveDate.Text & "'", "SeqNo Desc")
                    iSeqNo = If(drResultA.Length = 0, 1, drResultA(0)("SeqNo") + 1) '最後序號加1
                    dt = UpdateTable(drItems, dt, iSOCID, iSeqNo, dtTemp, sMesbox)
                End If
            Next

            LeaveDate.Text = Trim(LeaveDate.Text)
            If LeaveDate.Text <> "" Then LeaveDateHidden.Value = LeaveDate.Text

        End If

        Session(cst_SearchStr) = If(Session(cst_SearchStr) IsNot Nothing, Session(cst_SearchStr), ViewState(cst_SearchStr))

        '任何錯誤訊息，就停止UPDATE
        If sMesbox <> "" Then
            'Button1.Visible = True
            Common.RespWrite(Me, "<script language='javascript'>alert('" & sMesbox & "');</script>")
            Exit Sub
        End If

        Dim s_MRqID As String = TIMS.Get_MRqID(Me) 'Request("ID")
        Try
            DbAccess.UpdateDataTable(dt, da)
            If rqProecess = "add" Then
                LeaveDate.Text = ""
                Call GetAllClass()
                Page.RegisterStartupScript("confirm", "<script language='javascript'>if (!confirm('新增成功，是否要繼續新增?')){document.form1.MySessionSet.value='1';document.form1.submit();}</script>")
            Else
                Common.RespWrite(Me, "<script language='javascript'>alert('修改成功');</script>")
                Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_002.aspx?ID=" & s_MRqID & "';</script>")
            End If
        Catch ex As Exception

            Dim strErrmsg As String = ""
            Dim str_COLUMN_1 As String = TIMS.Get_DataTableCOLUMN2(dt)
            For x As Integer = 0 To dt.Rows.Count - 1
                Dim dr As DataRow = dt.Rows(x)
                Dim str_VALUES_1 As String = TIMS.Get_DataRowValues(str_COLUMN_1, dr)
                strErrmsg &= String.Concat(String.Format("/* x:{0} */ ", x), vbCrLf, " INSERT INTO STUD_TURNOUT (", str_COLUMN_1, ")", vbCrLf)
                strErrmsg &= String.Concat(" VALUES (", str_VALUES_1, ")", vbCrLf)
            Next
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            Call TIMS.WriteTraceLog(strErrmsg)
            Throw ex
        End Try

    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim dtTemp As DataTable = Nothing '找出歷史資料 (學員某日)
        Dim dtResult As DataTable = Nothing '先取出假別資料 (某班)

        Dim errMsg As String = ""
        Call ChkAddData1(errMsg, dtResult, dtTemp) 'out@dtResult, dtTemp
        If errMsg <> "" Then
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        Call SaveData1(dtResult, dtTemp)
    End Sub

    '回上一頁
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        Call SetSession()
    End Sub

    ''取得課程(自動/隱藏)
    'Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
    '    GetAllClass()
    'End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim flagDataItem As Boolean = False
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "head_navy"
                'e.Item.Cells(3).Attributes.Add("class", "")
                'e.Item.Cells(4).Attributes.Add("class", "")
                'e.Item.Cells(5).Attributes.Add("class", "")
                'e.Item.Cells(6).Attributes.Add("class", "")
                e.Item.CssClass = "FixedTitleRow"
                e.Item.Cells(3).Attributes.Add("class", "FixedTitleColumn")
                e.Item.Cells(4).Attributes.Add("class", "FixedTitleColumn")
                e.Item.Cells(5).Attributes.Add("class", "FixedTitleColumn")
                e.Item.Cells(6).Attributes.Add("class", "FixedTitleColumn")

            Case ListItemType.AlternatingItem, ListItemType.Item
                flagDataItem = True
        End Select

        If flagDataItem Then
            '資料列
            Dim drv As DataRowView = e.Item.DataItem
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
            Dim myLeaveID As DropDownList = e.Item.FindControl("LeaveID")
            'by Vicient 2006/08/03
            Dim C14 As HtmlInputCheckBox = e.Item.FindControl("C14")
            Dim C58 As HtmlInputCheckBox = e.Item.FindControl("C58")
            Dim C18 As HtmlInputCheckBox = e.Item.FindControl("C18")
            Dim C912 As HtmlInputCheckBox = e.Item.FindControl("C912")

            Select Case Convert.ToString(drv("TPeriod"))  'add by Kevin 95.11.01 修正aspx 無法判別隱藏物件
                Case "01" '日間(1-8堂)
                    C14.Attributes("onclick") = "GetClassTime1(" & e.Item.ItemIndex + 1 & ");"
                    C14.Disabled = False
                    C58.Attributes("onclick") = "GetClassTime2(" & e.Item.ItemIndex + 1 & ");"
                    C58.Disabled = False
                    C912.Attributes("onclick") = ""
                    C912.Disabled = True
                    C18.Attributes("onclick") = "GetClassTime4(" & e.Item.ItemIndex + 1 & ");"
                    C18.Disabled = False
                Case "02" '晚上
                    C14.Attributes("onclick") = ""
                    C14.Disabled = True
                    C58.Attributes("onclick") = ""
                    C58.Disabled = True
                    C912.Attributes("onclick") = "GetClassTime5(" & e.Item.ItemIndex + 1 & ");"
                    C912.Disabled = False
                    C18.Attributes("onclick") = ""
                    C18.Disabled = True
                Case Else
                    C14.Attributes("onclick") = "GetClassTime1(" & e.Item.ItemIndex + 1 & ");"
                    C14.Disabled = False
                    C58.Attributes("onclick") = "GetClassTime2(" & e.Item.ItemIndex + 1 & ");"
                    C58.Disabled = False
                    C912.Attributes("onclick") = "GetClassTime3(" & e.Item.ItemIndex + 1 & ");"
                    C912.Disabled = False
                    C18.Attributes("onclick") = "GetClassTime4(" & e.Item.ItemIndex + 1 & ");"
                    C18.Disabled = False
                    'Case "03" '全日制(暫停使用) 
                    'Case "04" '假日
                    'Case "05" '日間(1-12堂)
                    'Case "06" '下午及晚上
                    'Case "07" '其他
            End Select

            e.Item.Cells(0).Text = Right(e.Item.Cells(0).Text, 2)
            Select Case Convert.ToString(drv("StudStatus"))
                Case "1"
                    e.Item.Cells(1).Text += "(在訓)"
                Case "2"
                    e.Item.Cells(1).Text += "(離訓)"
                Case "3"
                    e.Item.Cells(1).Text += "(退訓)"
                Case "4"
                    e.Item.Cells(1).Text += "(續訓)"
                Case "5"
                    e.Item.Cells(1).Text += "(結訓)"
            End Select

            With myLeaveID
                Select Case Convert.ToString(drv("Sex"))
                    Case "F"
                        .DataSource = dtLeaveF
                    Case Else
                        .DataSource = dtLeaveM
                End Select
                .DataTextField = "Name"
                .DataValueField = "LeaveID"
                .DataBind()
                .Items.Insert(0, New ListItem("請選擇", 0))
            End With

            Const Cst_備註 As Integer = 20
            e.Item.Cells(Cst_備註).Text = ""

            'Dim dr As DataRow
            Dim iThours As Integer = Val(HidThours.Value)
            Dim iTotal As Integer = 0
            Dim sItemVar1 As String = HidItemVar1.Value 'TIMS.GetGlobalVar(Me, "4", "1")
            Dim sItemVar2 As String = HidItemVar2.Value 'TIMS.GetGlobalVar(Me, "4", "2")
            If Convert.ToString(drv("total")) <> "" Then iTotal = Val(drv("total"))

            If sItemVar1 <> "" AndAlso sItemVar2 <> "" Then
                Try
                    If sItemVar1.ToString <> "0" Then
                        If iTotal <= (iThours * Int(Split(sItemVar1, "/")(0)) / Int(Split(sItemVar1, "/")(1))) Then
                            e.Item.Cells(Cst_備註).Text = ""
                        ElseIf iTotal > (iThours * Int(Split(sItemVar1, "/")(0)) / Int(Split(sItemVar1, "/")(1))) And Int(iThours * Int(Split(sItemVar2, "/")(0)) / Int(Split(sItemVar2, "/")(1))) Then
                            e.Item.Cells(Cst_備註).Text = "警告!!請假總時數：" & e.Item.Cells(Cst_備註).Text & "<br>已超過" & sItemVar1
                            e.Item.Cells(Cst_備註).ForeColor = Color.Red
                        ElseIf iTotal > (iThours * Int(Split(sItemVar2, "/")(0)) / Int(Split(sItemVar2, "/")(1))) Then
                            e.Item.Cells(Cst_備註).Text = "警告!!請假總時數：" & e.Item.Cells(Cst_備註).Text & "<br>已超過" & sItemVar2
                            e.Item.Cells(Cst_備註).ForeColor = Color.Red
                        End If
                    End If
                Catch ex As Exception
                    e.Item.Cells(Cst_備註).Text = " 請假總時數定義有誤，請重新定義!!(" & Convert.ToString(sItemVar1) & ")"
                    e.Item.Cells(Cst_備註).ForeColor = Color.Red
                End Try
            Else
                e.Item.Cells(Cst_備註).Text = "警告!系統參數尚未設定出缺勤警示"
                e.Item.Cells(Cst_備註).ForeColor = Color.Red
            End If

            If rqProecess = "add" Then
                '離訓日期 
                If (drv("RejectTDate1").ToString <> "") And (LeaveDateHidden.Value <> "") Then
                    If CDate(drv("RejectTDate1")) <= CDate(LeaveDateHidden.Value) Then
                        C1.Disabled = True
                        C2.Disabled = True
                        C3.Disabled = True
                        C4.Disabled = True
                        C5.Disabled = True
                        C6.Disabled = True
                        C7.Disabled = True
                        C8.Disabled = True
                        C9.Disabled = True
                        C10.Disabled = True
                        C11.Disabled = True
                        C12.Disabled = True
                        myLeaveID.Enabled = False
                        C14.Disabled = True
                        C58.Disabled = True
                        C18.Disabled = True
                        C912.Disabled = True
                        If e.Item.Cells(Cst_備註).Text <> "" Then e.Item.Cells(Cst_備註).Text += "<br>"
                        e.Item.Cells(Cst_備註).Text += "該學員已於" & Common.FormatDate(drv("RejectTDate1")) & "離訓"
                        e.Item.Cells(Cst_備註).ForeColor = Color.Red
                    End If
                End If

                '退訓日期
                If (drv("RejectTDate2").ToString <> "") AndAlso (LeaveDateHidden.Value <> "") Then
                    If CDate(drv("RejectTDate2")) <= CDate(LeaveDateHidden.Value) Then
                        C1.Disabled = True
                        C2.Disabled = True
                        C3.Disabled = True
                        C4.Disabled = True
                        C5.Disabled = True
                        C6.Disabled = True
                        C7.Disabled = True
                        C8.Disabled = True
                        C9.Disabled = True
                        C10.Disabled = True
                        C11.Disabled = True
                        C12.Disabled = True
                        myLeaveID.Enabled = False
                        C14.Disabled = True
                        C58.Disabled = True
                        C18.Disabled = True
                        C912.Disabled = True
                        If e.Item.Cells(Cst_備註).Text <> "" Then e.Item.Cells(Cst_備註).Text += "<br>"
                        e.Item.Cells(Cst_備註).Text += "該學員已於" & Common.FormatDate(drv("RejectTDate2")) & "退訓"
                        e.Item.Cells(Cst_備註).ForeColor = Color.Red
                    End If
                End If
            End If
        End If
    End Sub

    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call sCreate1(1)
    End Sub

    Protected Sub ddl_sSearch2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddl_sSearch2.SelectedIndexChanged
        'Dim s_SOCID As String = ddl_sSearch2.SelectedValue
        Call GetAllClass() '(s_SOCID)
    End Sub
End Class
