Public Class SYS_05_004
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not Page.IsPostBack Then
            labmsg.Text = ""
            Call sUtl_Cancel1()
            tbSch.Visible = True

            Call initObj()
        End If
    End Sub

    '功能第一次載入初始化
    Sub initObj()
        'Call ListClass.crtDropDownList("org_master", ddlQoyc_status)
        'ddlQTabNum = Get_ddlTabNum(ddlQTabNum)
        'TabNum = Get_ddlTabNum(TabNum)
        TIMS.SUB_SET_HR_MI(HR1, MM1)
        TIMS.SUB_SET_HR_MI(HR2, MM2)
        Common.SetListItem(HR2, 23)
        Common.SetListItem(MM2, 59)
    End Sub

    '取消
    Sub sUtl_Cancel1()
        tbSch.Visible = False
        tbList.Visible = False
        tbEdit.Visible = False
    End Sub

    '清除值(及狀態設定)
    Sub clsValue()
        HidQTYPE.Value = TIMS.GetListValue(RBL_QTYPE)
        HidHN3ID.Value = ""
        HidHN4ID.Value = ""

        labQTYPE_N.Text = TIMS.GetListText(RBL_QTYPE)
        StopSDate.Text = ""
        StopEDate.Text = ""
        txtSubject.Text = ""
        PostDate.Text = ""
    End Sub

    '記錄查詢條件 
    Sub Search1Value()
        '記錄查詢條件
        StopSDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(StopSDate1.Text))
        StopSDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(StopSDate2.Text))
        StopEDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(StopEDate1.Text))
        StopEDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(StopEDate2.Text))

        ViewState("StopSDate1") = StopSDate1.Text
        ViewState("StopSDate2") = StopSDate2.Text
        ViewState("StopEDate1") = StopEDate1.Text
        ViewState("StopEDate2") = StopEDate2.Text
    End Sub

    '查詢 H3
    Sub Search1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.HN3ID" & vbCrLf '/*PK*/ 
        sql &= " ,'H3' QTYPE" & vbCrLf '/*PK*/ 
        sql &= " ,a.SUBJECT" & vbCrLf
        sql &= " ,convert(varchar, a.STOPSDATE, 120) STOPSDATE" & vbCrLf
        sql &= " ,convert(varchar, a.STOPEDATE, 120) STOPEDATE" & vbCrLf
        'sql += " ,a.STOPSDATE" & vbCrLf
        'sql += " ,a.STOPEDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, a.POSTDATE, 111) POSTDATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " ,a.DELETEACCT" & vbCrLf
        sql &= " ,a.DELETEDATE" & vbCrLf
        sql &= " ,ac.NAME MODIFYNAME" & vbCrLf
        sql &= " FROM HOME_NEWS3 a" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.ISDELETE IS NULL" & vbCrLf '顯示未刪除的資料
        If Convert.ToString(ViewState("StopSDate1")) <> "" Then
            sql &= " AND a.STOPSDATE >=@StopSDate1" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopSDate2")) <> "" Then
            sql &= " AND a.STOPSDATE <=@StopSDate2" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopEDate1")) <> "" Then
            sql &= " AND a.STOPEDATE >=@StopEDate1" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopEDate2")) <> "" Then
            sql &= " AND a.STOPEDATE <=@StopEDate2" & vbCrLf
        End If
        sql &= " ORDER BY a.STOPSDATE DESC,a.STOPEDATE DESC" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            If Convert.ToString(ViewState("StopSDate1")) <> "" Then
                'sql += " AND a.STOPSDATE >=@StopSDate1" & vbCrLf
                .Parameters.Add("StopSDate1", SqlDbType.DateTime).Value = CDate(ViewState("StopSDate1"))
            End If
            If Convert.ToString(ViewState("StopSDate2")) <> "" Then
                'sql += " AND a.STOPSDATE <=@StopSDate2" & vbCrLf
                .Parameters.Add("StopSDate2", SqlDbType.DateTime).Value = CDate(ViewState("StopSDate2"))
            End If
            If Convert.ToString(ViewState("StopEDate1")) <> "" Then
                'sql += " AND a.STOPEDATE >=@StopEDate1" & vbCrLf
                .Parameters.Add("StopEDate1", SqlDbType.DateTime).Value = CDate(ViewState("StopEDate1"))
            End If
            If Convert.ToString(ViewState("StopEDate2")) <> "" Then
                'sql += " AND a.STOPEDATE <=@StopEDate2" & vbCrLf
                .Parameters.Add("StopEDate2", SqlDbType.DateTime).Value = CDate(ViewState("StopEDate2"))
            End If
            dt.Load(.ExecuteReader())
        End With

        labmsg.Text = "查無資料"
        tbList.Visible = False
        If dt.Rows.Count = 0 Then Return

        'CPdt = dt.Copy()
        labmsg.Text = ""
        tbList.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢 H4
    Sub Search2()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.HN4ID" & vbCrLf '/*PK*/ 
        sql &= " ,'H4' QTYPE" & vbCrLf '/*PK*/ 
        sql &= " ,a.SUBJECT" & vbCrLf
        sql &= " ,convert(varchar, a.STOPSDATE, 120) STOPSDATE" & vbCrLf
        sql &= " ,convert(varchar, a.STOPEDATE, 120) STOPEDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, a.POSTDATE, 111) POSTDATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " ,a.DELETEACCT" & vbCrLf
        sql &= " ,a.DELETEDATE" & vbCrLf
        sql &= " ,ac.NAME MODIFYNAME" & vbCrLf
        sql &= " FROM HOME_NEWS4 a" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.ISDELETE IS NULL" & vbCrLf '顯示未刪除的資料
        If Convert.ToString(ViewState("StopSDate1")) <> "" Then
            sql &= " AND a.STOPSDATE >=@StopSDate1" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopSDate2")) <> "" Then
            sql &= " AND a.STOPSDATE <=@StopSDate2" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopEDate1")) <> "" Then
            sql &= " AND a.STOPEDATE >=@StopEDate1" & vbCrLf
        End If
        If Convert.ToString(ViewState("StopEDate2")) <> "" Then
            sql &= " AND a.STOPEDATE <=@StopEDate2" & vbCrLf
        End If
        sql &= " ORDER BY a.STOPSDATE DESC,a.STOPEDATE DESC" & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            If Convert.ToString(ViewState("StopSDate1")) <> "" Then
                'sql += " AND a.STOPSDATE >=@StopSDate1" & vbCrLf
                .Parameters.Add("StopSDate1", SqlDbType.DateTime).Value = CDate(ViewState("StopSDate1"))
            End If
            If Convert.ToString(ViewState("StopSDate2")) <> "" Then
                'sql += " AND a.STOPSDATE <=@StopSDate2" & vbCrLf
                .Parameters.Add("StopSDate2", SqlDbType.DateTime).Value = CDate(ViewState("StopSDate2"))
            End If
            If Convert.ToString(ViewState("StopEDate1")) <> "" Then
                'sql += " AND a.STOPEDATE >=@StopEDate1" & vbCrLf
                .Parameters.Add("StopEDate1", SqlDbType.DateTime).Value = CDate(ViewState("StopEDate1"))
            End If
            If Convert.ToString(ViewState("StopEDate2")) <> "" Then
                'sql += " AND a.STOPEDATE <=@StopEDate2" & vbCrLf
                .Parameters.Add("StopEDate2", SqlDbType.DateTime).Value = CDate(ViewState("StopEDate2"))
            End If
            dt.Load(.ExecuteReader())
        End With

        labmsg.Text = "查無資料"
        tbList.Visible = False
        If dt.Rows.Count = 0 Then Return

        'CPdt = dt.Copy()
        labmsg.Text = ""
        tbList.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    'H3
    Sub loadData1()
        If HidHN3ID.Value = "" Then Return

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.HN3ID" & vbCrLf '/*PK*/ 
        sql &= " ,a.SUBJECT" & vbCrLf
        sql &= " ,convert(varchar, a.STOPSDATE, 120) STOPSDATE" & vbCrLf
        sql &= " ,convert(varchar, a.STOPEDATE, 120) STOPEDATE" & vbCrLf
        'sql += " ,a.STOPSDATE" & vbCrLf
        'sql += " ,a.STOPEDATE" & vbCrLf
        sql &= " ,a.POSTDATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " ,a.DELETEACCT" & vbCrLf
        sql &= " ,a.DELETEDATE" & vbCrLf
        sql &= " ,ac.NAME MODIFYNAME" & vbCrLf
        sql &= " FROM HOME_NEWS3 a" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql += " AND a.ISDELETE IS NULL" & vbCrLf '顯示未刪除的資料
        sql &= " AND a.HN3ID=@HN3ID " & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("HN3ID", SqlDbType.Int).Value = Val(HidHN3ID.Value)
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then Return
        Dim dr As DataRow = dt.Rows(0)

        Me.HidHN3ID.Value = Convert.ToString(dr("HN3ID"))

        txtSubject.Text = Convert.ToString(dr("Subject"))
        StopSDate.Text = TIMS.Cdate3(dr("STOPSDATE"))
        If Convert.ToString(dr("STOPSDATE")) <> "" Then
            TIMS.SET_DateHM(CDate(dr("STOPSDATE")), HR1, MM1)
            'Common.SetListItem(HR1, CDate(dr("STOPSDATE")).Hour)
            'Common.SetListItem(MM1, CDate(dr("STOPSDATE")).Minute)
        End If

        StopEDate.Text = TIMS.Cdate3(dr("STOPEDATE"))
        If Convert.ToString(dr("STOPEDATE")) <> "" Then
            TIMS.SET_DateHM(CDate(dr("STOPEDATE")), HR2, MM2)
            'Common.SetListItem(HR2, CDate(dr("STOPEDATE")).Hour)
            'Common.SetListItem(MM2, CDate(dr("STOPEDATE")).Minute)
        End If

        PostDate.Text = TIMS.Cdate3(dr("PostDate"))

    End Sub

    'H4
    Sub loadData2()
        If HidHN4ID.Value = "" Then Return

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.HN4ID" & vbCrLf '/*PK*/ 
        sql &= " ,a.SUBJECT" & vbCrLf
        sql &= " ,convert(varchar, a.STOPSDATE, 120) STOPSDATE" & vbCrLf
        sql &= " ,convert(varchar, a.STOPEDATE, 120) STOPEDATE" & vbCrLf
        'sql += " ,a.STOPSDATE" & vbCrLf
        'sql += " ,a.STOPEDATE" & vbCrLf
        sql &= " ,a.POSTDATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.ISDELETE" & vbCrLf
        sql &= " ,a.DELETEACCT" & vbCrLf
        sql &= " ,a.DELETEDATE" & vbCrLf
        sql &= " ,ac.NAME MODIFYNAME" & vbCrLf
        sql &= " FROM HOME_NEWS4 a" & vbCrLf
        sql &= " JOIN AUTH_ACCOUNT ac on ac.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql += " AND a.ISDELETE IS NULL" & vbCrLf '顯示未刪除的資料
        sql &= " AND a.HN4ID=@HN4ID " & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("HN4ID", SqlDbType.Int).Value = Val(HidHN4ID.Value)
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then Return
        Dim dr As DataRow = dt.Rows(0)

        HidHN4ID.Value = Convert.ToString(dr("HN4ID"))
        txtSubject.Text = Convert.ToString(dr("Subject"))
        StopSDate.Text = TIMS.Cdate3(dr("STOPSDATE"))
        If Convert.ToString(dr("STOPSDATE")) <> "" Then
            TIMS.SET_DateHM(CDate(dr("STOPSDATE")), HR1, MM1)
            'Common.SetListItem(HR1, CDate(dr("STOPSDATE")).Hour)
            'Common.SetListItem(MM1, CDate(dr("STOPSDATE")).Minute)
        End If
        StopEDate.Text = TIMS.Cdate3(dr("STOPEDATE"))
        If Convert.ToString(dr("STOPEDATE")) <> "" Then
            TIMS.SET_DateHM(CDate(dr("STOPEDATE")), HR2, MM2)
            'Common.SetListItem(HR2, CDate(dr("STOPEDATE")).Hour)
            'Common.SetListItem(MM2, CDate(dr("STOPEDATE")).Minute)
        End If
        PostDate.Text = TIMS.Cdate3(dr("PostDate"))

    End Sub


    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        StopSDate.Text = TIMS.ClearSQM(StopSDate.Text)
        StopEDate.Text = TIMS.ClearSQM(StopEDate.Text)
        PostDate.Text = TIMS.ClearSQM(PostDate.Text)

        If StopSDate.Text = "" Then
            Errmsg += "請輸入 停止日期起" & vbCrLf
        End If
        If StopEDate.Text = "" Then
            Errmsg += "請選擇 停止日期迄" & vbCrLf
        End If
        If PostDate.Text = "" Then
            Errmsg += "請輸入 發布日期" & vbCrLf
        End If
        If txtSubject.Text = "" Then
            Errmsg += "請輸入 發布主題" & vbCrLf
        End If

        If StopSDate.Text <> "" Then
            If Not TIMS.IsDate1(StopSDate.Text) Then
                Errmsg += "停止日期起 請輸入正確日期格式" & vbCrLf
            End If
        End If
        If StopEDate.Text <> "" Then
            If Not TIMS.IsDate1(StopEDate.Text) Then
                Errmsg += "停止日期迄 請輸入正確日期格式" & vbCrLf
            End If
        End If
        If PostDate.Text <> "" Then
            If Not TIMS.IsDate1(PostDate.Text) Then
                Errmsg += "發布日期 請輸入正確日期格式" & vbCrLf
            End If
        End If

        If Errmsg = "" Then
            'Dim sHR1 As String = "0" & Me.HR1.SelectedValue
            'Dim sMM1 As String = "0" & Me.MM1.SelectedValue
            'Dim sHR2 As String = "0" & Me.HR2.SelectedValue
            'Dim sMM2 As String = "0" & Me.MM2.SelectedValue
            'If Me.HR1.SelectedValue >= 10 Then sHR1 = Me.HR1.SelectedValue
            'If Me.HR2.SelectedValue >= 10 Then sHR2 = Me.HR2.SelectedValue
            'If Me.MM1.SelectedValue >= 10 Then sMM1 = Me.MM1.SelectedValue
            'If Me.MM2.SelectedValue >= 10 Then sMM2 = Me.MM2.SelectedValue
            StopSDate.Text = TIMS.Cdate3(StopSDate.Text)
            StopEDate.Text = TIMS.Cdate3(StopEDate.Text)
            Dim tStopSDate As DateTime = CDate(TIMS.GET_DateHM(StopSDate, HR1, MM1)) 'CDate(StopSDate.Text & " " & sHR1 & ":" & sMM1)
            Dim tStopEDate As DateTime = CDate(TIMS.GET_DateHM(StopEDate, HR2, MM2)) 'CDate(StopEDate.Text & " " & sHR2 & ":" & sMM2)
            If DateDiff(DateInterval.Minute, tStopSDate, tStopEDate) <= 0 Then
                Errmsg += "停止日期起迄 前後順序有誤!!" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存 H3
    Sub SaveData1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rst As Integer = 0

        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO HOME_NEWS3( " & vbCrLf
        sql &= " HN3ID  " & vbCrLf '/*PK*/ 
        sql &= " ,SUBJECT" & vbCrLf
        sql &= " ,STOPSDATE" & vbCrLf
        sql &= " ,STOPEDATE" & vbCrLf
        sql &= " ,POSTDATE" & vbCrLf
        sql &= " ,CREATEACCT" & vbCrLf
        sql &= " ,CREATEDATE" & vbCrLf
        sql &= " ,MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE" & vbCrLf
        'sql += " ,ISDELETE" & vbCrLf
        'sql += " ,DELETEACCT" & vbCrLf
        'sql += " ,DELETEDATE" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @HN3ID  " & vbCrLf '/*PK*/ 
        sql &= " ,@SUBJECT" & vbCrLf
        sql &= " ,@STOPSDATE" & vbCrLf
        sql &= " ,@STOPEDATE" & vbCrLf
        sql &= " ,@POSTDATE" & vbCrLf
        sql &= " ,@CREATEACCT" & vbCrLf
        sql &= " ,getdate()" & vbCrLf
        sql &= " ,@MODIFYACCT" & vbCrLf
        sql &= " ,getdate()" & vbCrLf
        sql &= " )" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " UPDATE HOME_NEWS3" & vbCrLf
        sql &= " SET SUBJECT=@SUBJECT" & vbCrLf
        sql &= " ,STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " ,STOPEDATE=@STOPEDATE" & vbCrLf
        sql &= " ,POSTDATE=@POSTDATE" & vbCrLf
        'sql += " ,CREATEACCT" & vbCrLf
        'sql += " ,CREATEDATE" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=getdate()" & vbCrLf
        sql &= " WHERE HN3ID =@HN3ID" & vbCrLf '/*PK*/ 
        Dim uCmd As New SqlCommand(sql, objconn)

        '新增重複判斷
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM HOME_NEWS3" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql &= " AND STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " AND STOPEDATE=@STOPEDATE" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        '修改重複判斷
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM HOME_NEWS3" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql &= " AND STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " AND STOPEDATE=@STOPEDATE" & vbCrLf
        sql &= " AND HN3ID!=@HN3ID" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)

        'Dim sHR1 As String = "0" & Me.HR1.SelectedValue
        'Dim sMM1 As String = "0" & Me.MM1.SelectedValue
        'Dim sHR2 As String = "0" & Me.HR2.SelectedValue
        'Dim sMM2 As String = "0" & Me.MM2.SelectedValue
        'If Me.HR1.SelectedValue >= 10 Then sHR1 = Me.HR1.SelectedValue
        'If Me.HR2.SelectedValue >= 10 Then sHR2 = Me.HR2.SelectedValue
        'If Me.MM1.SelectedValue >= 10 Then sMM1 = Me.MM1.SelectedValue
        'If Me.MM2.SelectedValue >= 10 Then sMM2 = Me.MM2.SelectedValue
        StopSDate.Text = TIMS.Cdate3(StopSDate.Text)
        StopEDate.Text = TIMS.Cdate3(StopEDate.Text)
        Dim tStopSDate As DateTime = CDate(TIMS.GET_DateHM(StopSDate, HR1, MM1)) 'CDate(StopSDate.Text & " " & sHR1 & ":" & sMM1)
        Dim tStopEDate As DateTime = CDate(TIMS.GET_DateHM(StopEDate, HR2, MM2)) 'CDate(StopEDate.Text & " " & sHR2 & ":" & sMM2)
        'Dim tStopSDate As DateTime = CDate(StopSDate.Text & " " & sHR1 & ":" & sMM1)
        'Dim tStopEDate As DateTime = CDate(StopEDate.Text & " " & sHR2 & ":" & sMM2)

        If HidHN3ID.Value = "" Then
            '新增
            Dim dt1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該時間區間已新增，請使用修改功能!!")
                Exit Sub
            End If
        Else
            '修改
            Dim dt1 As New DataTable
            With sCmd2
                .Parameters.Clear()
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("HN3ID", SqlDbType.Int).Value = Val(HidHN3ID.Value) '/*PK*/
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該時間區間已存在，請重新輸入!!")
                Exit Sub
            End If
        End If

        txtSubject.Text = TIMS.ClearSQM(txtSubject.Text)
        If HidHN3ID.Value = "" Then
            '新增
            Dim iHN3ID As Integer = DbAccess.GetNewId(objconn, " HOME_NEWS3_HN3ID_SEQ,HOME_NEWS3,HN3ID")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("HN3ID", SqlDbType.Int).Value = iHN3ID  '/*PK*/
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = txtSubject.Text
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = If(PostDate.Text <> "", TIMS.Cdate2(PostDate.Text), Convert.DBNull)
                .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("CREATEDATE", SqlDbType.VarChar).Value = CREATEDATE
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE
                '.Parameters.Add("ISDELETE", SqlDbType.VarChar).Value = ISDELETE
                '.Parameters.Add("DELETEACCT", SqlDbType.VarChar).Value = DELETEACCT
                '.Parameters.Add("DELETEDATE", SqlDbType.VarChar).Value = DELETEDATE
                rst = .ExecuteNonQuery()
            End With
        Else
            '修改
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = txtSubject.Text
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = If(PostDate.Text <> "", TIMS.Cdate2(PostDate.Text), Convert.DBNull)
                '.Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("CREATEDATE", SqlDbType.VarChar).Value = CREATEDATE
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                .Parameters.Add("HN3ID", SqlDbType.Int).Value = Val(HidHN3ID.Value) '/*PK*/
                rst = .ExecuteNonQuery()
            End With
        End If

        If rst = 0 Then
            Common.MessageBox(Page, "執行完畢，無資料更動!")
            Return
        End If

        Call sUtl_Cancel1()

        tbSch.Visible = True
        Call Search1()
    End Sub

    '儲存 H4
    Sub SaveData2()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rst As Integer = 0

        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO HOME_NEWS4(HN4ID ,SUBJECT,STOPSDATE,STOPEDATE,POSTDATE,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql &= "  VALUES (@HN4ID ,@SUBJECT,@STOPSDATE,@STOPEDATE,@POSTDATE,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " UPDATE HOME_NEWS4" & vbCrLf
        sql &= " SET SUBJECT=@SUBJECT" & vbCrLf
        sql &= " ,STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " ,STOPEDATE=@STOPEDATE" & vbCrLf
        sql &= " ,POSTDATE=@POSTDATE" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=getdate()" & vbCrLf
        sql &= " WHERE HN4ID =@HN4ID" & vbCrLf '/*PK*/ 
        Dim uCmd As New SqlCommand(sql, objconn)

        '新增重複判斷
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM HOME_NEWS4" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql &= " AND STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " AND STOPEDATE=@STOPEDATE" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        '修改重複判斷
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM HOME_NEWS4" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND ISDELETE IS NULL" & vbCrLf '使用中。
        sql &= " AND STOPSDATE=@STOPSDATE" & vbCrLf
        sql &= " AND STOPEDATE=@STOPEDATE" & vbCrLf
        sql &= " AND HN4ID!=@HN4ID" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)

        'Dim sHR1 As String = "0" & Me.HR1.SelectedValue
        'Dim sMM1 As String = "0" & Me.MM1.SelectedValue
        'Dim sHR2 As String = "0" & Me.HR2.SelectedValue
        'Dim sMM2 As String = "0" & Me.MM2.SelectedValue
        'If Me.HR1.SelectedValue >= 10 Then sHR1 = Me.HR1.SelectedValue
        'If Me.HR2.SelectedValue >= 10 Then sHR2 = Me.HR2.SelectedValue
        'If Me.MM1.SelectedValue >= 10 Then sMM1 = Me.MM1.SelectedValue
        'If Me.MM2.SelectedValue >= 10 Then sMM2 = Me.MM2.SelectedValue
        StopSDate.Text = TIMS.Cdate3(StopSDate.Text)
        StopEDate.Text = TIMS.Cdate3(StopEDate.Text)
        Dim tStopSDate As DateTime = CDate(TIMS.GET_DateHM(StopSDate, HR1, MM1)) 'CDate(StopSDate.Text & " " & sHR1 & ":" & sMM1)
        Dim tStopEDate As DateTime = CDate(TIMS.GET_DateHM(StopEDate, HR2, MM2)) 'CDate(StopEDate.Text & " " & sHR2 & ":" & sMM2)
        'Dim tStopSDate As DateTime = CDate(StopSDate.Text & " " & sHR1 & ":" & sMM1)
        'Dim tStopEDate As DateTime = CDate(StopEDate.Text & " " & sHR2 & ":" & sMM2)

        If HidHN4ID.Value = "" Then
            '新增
            Dim dt1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該時間區間已新增，請使用修改功能!!")
                Exit Sub
            End If
        Else
            '修改
            Dim dt1 As New DataTable
            With sCmd2
                .Parameters.Clear()
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("HN4ID", SqlDbType.Int).Value = Val(HidHN4ID.Value) '/*PK*/
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該時間區間已存在，請重新輸入!!")
                Exit Sub
            End If
        End If

        txtSubject.Text = TIMS.ClearSQM(txtSubject.Text)

        If HidHN4ID.Value = "" Then
            '新增
            Dim iHN4ID As Integer = DbAccess.GetNewId(objconn, " HOME_NEWS4_HN4ID_SEQ,HOME_NEWS4,HN4ID")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("HN4ID", SqlDbType.Int).Value = iHN4ID  '/*PK*/
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = txtSubject.Text
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = If(PostDate.Text <> "", TIMS.Cdate2(PostDate.Text), Convert.DBNull)
                .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("CREATEDATE", SqlDbType.VarChar).Value = CREATEDATE
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                rst = .ExecuteNonQuery()
            End With
        Else
            '修改
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("SUBJECT", SqlDbType.NVarChar).Value = txtSubject.Text
                .Parameters.Add("STOPSDATE", SqlDbType.DateTime).Value = tStopSDate
                .Parameters.Add("STOPEDATE", SqlDbType.DateTime).Value = tStopEDate
                .Parameters.Add("POSTDATE", SqlDbType.DateTime).Value = If(PostDate.Text <> "", TIMS.Cdate2(PostDate.Text), Convert.DBNull)
                '.Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                '.Parameters.Add("CREATEDATE", SqlDbType.VarChar).Value = CREATEDATE
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                .Parameters.Add("HN4ID", SqlDbType.Int).Value = Val(HidHN4ID.Value) '/*PK*/
                rst = .ExecuteNonQuery()
            End With
        End If

        If rst = 0 Then
            Common.MessageBox(Page, "執行完畢，無資料更動!")
            Return
        End If

        Call sUtl_Cancel1()

        tbSch.Visible = True
        Call Search2()
    End Sub


    '刪除 H3
    Sub Delete1()
        If HidHN3ID.Value = "" Then
            Common.MessageBox(Page, "查無刪除序號，請重新查詢!")
            Exit Sub
        End If

        Dim rst As Integer = 0
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE HOME_NEWS3" & vbCrLf
        sql &= " SET ISDELETE='Y'" & vbCrLf
        sql &= " ,DELETEACCT=@DELETEACCT" & vbCrLf
        sql &= " ,DELETEDATE= getdate()" & vbCrLf
        sql &= " WHERE HN3ID=@HN3ID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)

        Dim dCmd As New SqlCommand(sql, objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("DELETEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            .Parameters.Add("HN3ID", SqlDbType.Int).Value = Val(HidHN3ID.Value)
            rst = .ExecuteNonQuery()
        End With
        If rst = 1 Then
            Common.MessageBox(Page, "刪除成功!")
            Call Search1()
        End If
    End Sub

    '刪除 H4
    Sub Delete2()
        If HidHN4ID.Value = "" Then
            Common.MessageBox(Page, "查無刪除序號，請重新查詢!")
            Exit Sub
        End If

        Dim rst As Integer = 0
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE HOME_NEWS4" & vbCrLf
        sql &= " SET ISDELETE='Y'" & vbCrLf
        sql &= " ,DELETEACCT=@DELETEACCT" & vbCrLf
        sql &= " ,DELETEDATE= getdate()" & vbCrLf
        sql &= " WHERE HN4ID=@HN4ID " & vbCrLf
        Call TIMS.OpenDbConn(objconn)

        Dim dCmd As New SqlCommand(sql, objconn)
        With dCmd
            .Parameters.Clear()
            .Parameters.Add("DELETEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            .Parameters.Add("HN4ID", SqlDbType.Int).Value = Val(HidHN4ID.Value)
            rst = .ExecuteNonQuery()
        End With
        If rst = 1 Then
            Common.MessageBox(Page, "刪除成功!")
            Call Search1()
        End If
    End Sub


    '查詢鈕
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        '記錄查詢條件
        Call Search1Value()
        HidQTYPE.Value = TIMS.GetListValue(RBL_QTYPE)
        Select Case HidQTYPE.Value
            Case "H3"
                Call Search1()
            Case "H4"
                Call Search2()
        End Select
    End Sub

    '新增鈕
    Protected Sub btnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        Call clsValue()

        Call sUtl_Cancel1()
        tbEdit.Visible = True
    End Sub

    '儲存
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        'Call SaveData1()
        'HidQTYPE.Value = TIMS.GetListValue(RBL_QTYPE)
        Select Case HidQTYPE.Value
            Case "H3"
                Call SaveData1()
            Case "H4"
                Call SaveData2()
        End Select
    End Sub

    '回上頁
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Call sUtl_Cancel1()
        tbSch.Visible = True
    End Sub

    Sub UTL_HIDVALUESET(ByRef sCmdArg As String)
        HidHN3ID.Value = TIMS.GetMyValue(sCmdArg, "HN3ID")
        HidHN4ID.Value = TIMS.GetMyValue(sCmdArg, "HN4ID")
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "UPD" '修改
                Call sUtl_Cancel1()
                tbEdit.Visible = True
                Call clsValue()
                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                Call UTL_HIDVALUESET(sCmdArg)
                Select Case HidQTYPE.Value
                    Case "H3"
                        Call loadData1()
                    Case "H4"
                        Call loadData2()
                End Select

            Case "DEL" '刪除
                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                Call UTL_HIDVALUESET(sCmdArg)
                Select Case HidQTYPE.Value
                    Case "H3"
                        Call Delete1()
                    Case "H4"
                        Call Delete2()
                End Select
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SYS_TD2"

                '序號
                e.Item.Cells(0).Text = (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + e.Item.ItemIndex + 1

                Dim lbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate")
                Dim lbtDelete As LinkButton = e.Item.FindControl("lbtDelete")
                lbtDelete.Attributes.Add("onclick", "return confirm('您確定要刪除第" & e.Item.Cells(0).Text & "筆資料嗎?');")

                Dim sCmdArg As String = ""
                Select Case HidQTYPE.Value
                    Case "H3"
                        Call TIMS.SetMyValue(sCmdArg, "HN3ID", drv("HN3ID"))
                    Case "H4"
                        Call TIMS.SetMyValue(sCmdArg, "HN4ID", drv("HN4ID"))
                End Select
                lbtUpdate.CommandArgument = sCmdArg
                lbtDelete.CommandArgument = sCmdArg

        End Select
    End Sub

    Public Shared Function SUtl_UploadFile1(ByRef MyPage As Page, ByRef oFile1 As HtmlInputFile, ByRef s_ErrMsg As String) As Boolean
        Dim rst As Boolean = False '異常-上傳檔案

        s_ErrMsg = "" 'Dim s_ErrMsg As String = ""
        Dim sUpload_Path As String = "~/upload/File/" '上傳位置
        Const Cst_Filetype As String = "xls,pdf" '匯入檔案類型(以逗號分隔)
        Dim sMyFileName As String = ""
        Dim sMyFileType As String = ""
        Const cst_errMsg1 As String = "有檔案嗎？？!"
        Const cst_errMsg2 As String = "檔案位置錯誤!"
        Const cst_errMsg3 As String = "檔案類型錯誤!"
        Dim str_errMsg4 As String = "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!"

        '檢查檔案
        If oFile1.Value = "" Then
            Common.MessageBox(MyPage, cst_errMsg1)
            s_ErrMsg = cst_errMsg1
            Return rst 'Exit Sub
        End If
        '檢查檔案格式與大小
        If oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(MyPage, cst_errMsg2)
            s_ErrMsg = cst_errMsg2
            Return rst 'Exit Sub
        End If
        '取出檔案名稱
        sMyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If sMyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(MyPage, cst_errMsg3)
            s_ErrMsg = cst_errMsg3
            Return rst 'Exit Sub
        End If
        sMyFileType = Split(sMyFileName, ".")((Split(sMyFileName, ".")).Length - 1)
        If LCase(Cst_Filetype).IndexOf(LCase(sMyFileType)) = -1 Then
            Common.MessageBox(MyPage, str_errMsg4)
            s_ErrMsg = str_errMsg4
            Return rst 'Exit Sub
        End If
        Try
            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
            sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            '上傳檔案
            oFile1.PostedFile.SaveAs(MyPage.Server.MapPath(sUpload_Path & sMyFileName))
            '複製檔案(表示上面動作已完成)
            Const cst_uploads1 As String = "uploads1.log"
            IO.File.Copy(MyPage.Server.MapPath("~\Upload\" & cst_uploads1), MyPage.Server.MapPath(sUpload_Path & cst_uploads1), True)

            'lab_UpMsg1.Text = sMyFileName & "上傳檔案完成!(OK)"
            s_ErrMsg = sMyFileName & "上傳檔案完成!(OK)"
            rst = True '上傳檔案OK
        Catch ex As Exception
            'lab_UpMsg1.Text = "上傳檔案有誤!" & ex.ToString
            s_ErrMsg = "上傳檔案有誤!" & ex.ToString
            Common.MessageBox(MyPage, "上傳檔案有誤!" & ex.ToString)
            Return rst 'Exit Sub
        End Try
        Return rst
    End Function

End Class