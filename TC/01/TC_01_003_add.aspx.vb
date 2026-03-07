Partial Class TC_01_003_add
    Inherits AuthBasePage

#Region "NO USE"
    'Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
    '    'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
    '    '請勿使用程式碼編輯器進行修改。
    '    'InitializeComponent()

    '    Dim sql As String
    '    Dim dt As DataTable
    '    Dim dr As DataRow
    '    '★SQL-ORACLE
    '    sql = " SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,DATA_LENGTH "
    '    sql &= " FROM USER_TAB_COLUMNS "
    '    sql &= " WHERE TABLE_NAME IN ('ID_CLASS','KEY_TRAINTYPE') "
    '    sql &= " AND DATA_TYPE IN ('NVARCHAR2','VARCHAR2','CHAR') "
    '    dt = DbAccess.GetDataTable(sql)
    '    For Each dr In dt.Rows
    '        Select Case UCase(dr("COLUMN_NAME"))
    '            Case "CLASSID" '班別代碼
    '                Me.TB_classid.MaxLength = dr("DATA_LENGTH")
    '                Me.TB_classid.ToolTip = "欄位長度" & dr("DATA_LENGTH")
    '            Case "CLASSNAME" '班別名稱
    '                Me.TBclass_name.MaxLength = dr("DATA_LENGTH")
    '                Me.TBclass_name.ToolTip = "欄位長度" & dr("DATA_LENGTH")
    '            Case "CLASSENAME" '英文名稱
    '                Me.ClassEName.MaxLength = dr("DATA_LENGTH")
    '                Me.ClassEName.ToolTip = "欄位長度" & dr("DATA_LENGTH")
    '            Case "TRAINNAME" '訓練職類
    '                Me.TB_career_id.MaxLength = dr("DATA_LENGTH")
    '                Me.TB_career_id.ToolTip = "欄位長度" & dr("DATA_LENGTH")
    '        End Select
    '    Next
    'End Sub
#End Region

    Const cst_sess_sch1_txt As String = "_searchTc0103"

    Dim flgROLEIDx0xLIDx0 As Boolean = False
    'Dim sRe_clsid As String = ""
    Dim ProcessType As String = ""
    'Dim FunDr As DataRow

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        '判斷是否為超級使用者-iType, sm
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(sm, 1)

        If Hid_CLSID.Value = "" Then Hid_CLSID.Value = TIMS.ClearSQM(Request("clsid"))
        'sRe_clsid = TIMS.ClearSQM(Request("clsid"))
        ProcessType = TIMS.ClearSQM(Request("ProcessType"))
        Re_ID.Value = TIMS.ClearSQM(Request("ID"))
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        Select Case ProcessType
            Case "Update"
                '修改學員學號
                tr_cb_USTUDENTID.Visible = If(flgROLEIDx0xLIDx0, True, False)
            Case Else
                '除了修改，其它不顯示
                tr_cb_USTUDENTID.Visible = False
        End Select

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.LabTMID.Text = "訓練業別"
            Me.Re_Class_CEName.Enabled = False
            Me.LabEnameStar.Visible = False
        Else
            Me.Re_Class_CEName.Enabled = True
            Me.LabEnameStar.Visible = True
        End If


        'Select Case ProcessType
        '    Case "Update"
        '        bt_save.Enabled = False
        '        If au.blnCanMod Then bt_save.Enabled = True
        '    Case "Insert"
        '        bt_save.Enabled = False
        '        If au.blnCanAdds Then bt_save.Enabled = True
        'End Select

        If Not Page.IsPostBack Then
            cCreate1()
        End If

        'If sm.UserInfo.RoleID = 5 Then
        '    Common.SetListItem(Plan_List, sm.UserInfo.TPlanID)
        '    Me.Plan_List.Enabled = False
        'End If

        If Not Session(cst_sess_sch1_txt) Is Nothing Then
            Me.ViewState(cst_sess_sch1_txt) = Session(cst_sess_sch1_txt)
            Session(cst_sess_sch1_txt) = Nothing
        End If
    End Sub

    Sub cCreate1()
        'Get_TPlan
        Plan_List = TIMS.Get_TPlan(Plan_List)
        Common.SetListItem(Plan_List, sm.UserInfo.TPlanID) '(新增預設)
        Plan_List.Enabled = False

        ddlYears = TIMS.Get_Years(ddlYears, objconn)
        ddlYears.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(ddlYears, sm.UserInfo.Years) '(新增預設)
        ddlYears.Enabled = False

        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID) '(新增預設)
        ddlDISTID.Enabled = False

        '依據傳入參數，帶出有效資料
        cCreate1_Load1()
    End Sub

    ''' <summary> 依據傳入參數，帶出有效資料 </summary>
    Sub cCreate1_Load1()
        Select Case ProcessType
            Case "Insert"
                Me.lblProecessType.Text = "新增"
                '新增狀態下，指定預設登入的訓練計畫 Start
                Common.SetListItem(Plan_List, sm.UserInfo.TPlanID)
                '20100208 按新增時代查詢之 班別代碼 & 班別名稱
                If Convert.ToString(Me.Request("ClassID")) <> "" Then
                    TB_classid.Text = Convert.ToString(Me.Request("ClassID")).ToUpper
                    hTB_classid.Value = TB_classid.Text
                End If
                TBclass_name.Text = Convert.ToString(Me.Request("ClassName"))
                '新增狀態下，指定預設登入的訓練計畫 End
            Case "Copy"
                Me.lblProecessType.Text = "複製新增"
            Case "Update"
                Me.lblProecessType.Text = "修改"
        End Select

        'Dim objreader As SqlDataReader
        'Dim objconn As SqlConnection
        'TIMS.TestDbConn(Me, objconn)
        'Dim vsTMID As String = "" 'TMID暫存
        Hid_CLSID.Value = TIMS.ClearSQM(Hid_CLSID.Value)

        Dim flag_LoadData As Boolean = False '傳入值正確才Load資料
        If ProcessType = "Update" OrElse ProcessType = "Copy" Then
            If Hid_CLSID.Value <> "" Then flag_LoadData = True '傳入值正確才Load資料
        End If

        '傳入值正確才Load資料
        If Not flag_LoadData Then Return
        '傳入值為空 查屁
        If Hid_CLSID.Value = "" Then Return

        Call TIMS.OpenDbConn(objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.Classid" & vbCrLf
        sql &= " ,a.ClassEName" & vbCrLf
        sql &= " ,a.ClassName" & vbCrLf
        sql &= " ,a.TPlanID" & vbCrLf
        sql &= " ,a.CJOB_UNKEY" & vbCrLf
        'sql += " ,s.CJOB_NO" & vbCrLf
        'sql += " ,s.CJOB_NAME" & vbCrLf
        sql &= " ,a.TMID" & vbCrLf
        sql &= " ,ISNULL(vt.trainname,vt.jobname) TRAINNAME" & vbCrLf
        sql &= " ,a.Content" & vbCrLf
        sql &= " ,vt.GCID3" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,a.DISTID" & vbCrLf
        sql &= " FROM ID_Class a" & vbCrLf
        sql &= " JOIN KEY_PLAN b on a.TPlanID = b.TPlanID" & vbCrLf
        sql &= " LEFT JOIN VIEW_TRAINTYPE vt on a.TMID= vt.TMID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s on a.CJOB_UNKEY = s.CJOB_UNKEY" & vbCrLf
        sql &= " WHERE a.CLSID=@CLSID" & vbCrLf
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("CLSID", SqlDbType.Int).Value = Val(Hid_CLSID.Value)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Return '查無資料搞屁

        '有資料。
        Dim dr As DataRow = dt.Rows(0)
        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)

        Common.SetListItem(ddlYears, Convert.ToString(dr("YEARS")))
        Common.SetListItem(ddlDISTID, Convert.ToString(dr("DISTID")))
        If ProcessType = "Update" Then
            TB_classid.Text = Convert.ToString(dr("Classid")).ToUpper
            hTB_classid.Value = Convert.ToString(dr("Classid")).ToUpper
        End If

        ClassEName.Text = Convert.ToString(dr("ClassEName"))
        'HidClassID1.Value
        HidClassID1.Value = TB_classid.Text
        TBclass_name.Text = Convert.ToString(dr("ClassName"))
        'Get_TPlan'Plan_List.SelectedValue = dr("TPlanID")
        Common.SetListItem(Plan_List, Convert.ToString(dr("TPlanID")))

        If Convert.ToString(dr("CJOB_UNKEY")) <> "" Then
            cjobValue.Value = dr("CJOB_UNKEY")
            txtCJOB_NAME.Text = TIMS.Get_CJOBNAME(dtSCJOB, cjobValue.Value)
            'txtCJOB_NAME.Text = "[" & dr("CJOB_NO").ToString & "]" & dr("CJOB_NAME").ToString
        End If

        'txtCJOB_NAME.Text = dr("CJOB_NAME")
        'vsTMID = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If iPYNum >= 3 Then
                '產投
                If Not Convert.IsDBNull(dr("TMID")) Then
                    'vsTMID = dr("TMID")
                    Me.trainValue.Value = dr("TMID")
                    TB_career_id.Text = Convert.ToString(dr("TRAINNAME"))
                    If Not IsDBNull(dr("GCID3")) Then
                        TB_career_id.Text = TIMS.Get_GCIDName(dr("GCID3").ToString, "2018", objconn)
                    End If
                End If
            Else
                '產投
                If Not Convert.IsDBNull(dr("TMID")) Then
                    'vsTMID = dr("TMID")
                    Me.jobValue.Value = dr("TMID")
                    TB_career_id.Text = Convert.ToString(dr("TRAINNAME"))
                End If
            End If
        Else
            '非產投
            If Not Convert.IsDBNull(dr("TMID")) Then
                'vsTMID = dr("TMID")
                Me.trainValue.Value = dr("TMID")
                TB_career_id.Text = Convert.ToString(dr("TRAINNAME"))
            End If
        End If

        ComSumm.Text = ""
        If Not Convert.IsDBNull(dr("Content")) Then
            ComSumm.Text = dr("Content")
        End If

    End Sub



    '新增，修改判斷是否重複 'TRUE: 重複  FALSE:未重複
    Public Shared Function Check_Double(ByVal ProcessType As String, ByRef hts As Hashtable, ByRef tConn As SqlConnection) As Boolean
        'Optional ByVal CLSID As String = "", Optional ByVal OldClassID As String = "") As Boolean
        Dim s_ClassID As String = TIMS.GetMyValue2(hts, "ClassID")
        Dim s_DistID As String = TIMS.GetMyValue2(hts, "DistID")
        Dim s_TPlanID As String = TIMS.GetMyValue2(hts, "TPlanID")
        Dim s_Years As String = TIMS.GetMyValue2(hts, "Years")
        Dim s_CLSID As String = TIMS.GetMyValue2(hts, "CLSID")
        Dim s_OldClassID As String = TIMS.GetMyValue2(hts, "OldClassID")
        ' ByVal ClassID As String, ByVal DistID As String, ByVal TPlanID As String, ByVal Years As String, ByVal CLSID As String, ByVal OldClassID As String

        Dim Rst As Boolean = False
        'ClassID :新班別代碼
        'Years :年度
        'CLSID:班別代碼PK(流水號)
        'OldClassID :原班別代碼

        Const CstInsert As String = "INSERT"
        Const CstCopy As String = "COPY"
        Const CstUpdate As String = "UPDATE"

        Dim strsql_check As String = ""
        Select Case UCase(ProcessType)
            Case CstInsert, CstCopy
                strsql_check = ""
                '**by Milor 20080508--班級代碼重複應該是用區域別+計畫別去判斷，所以加入了TPlanID的判斷式----start
                'strsql_check = ""
                'strsql_check += " select " & vbCrLf
                'strsql_check += " CLSID,CLASSID,CLASSNAME,CLASSENAME,TPLANID, " & vbCrLf
                'strsql_check += " dbo.SUBSTR(CONTENT, 1, 4000) CONTENT,  " & vbCrLf
                'strsql_check += " TMID, DistID, MODIFYACCT, MODIFYDATE, CJOB_UNKEY, Years " & vbCrLf
                strsql_check = ""
                strsql_check += " SELECT 'X'" & vbCrLf
                strsql_check += " FROM ID_Class " & vbCrLf
                strsql_check += " WHERE 1=1 " & vbCrLf
                strsql_check += " AND ClassID='" & s_ClassID & "' " & vbCrLf
                strsql_check += " AND DistID='" & s_DistID & "' " & vbCrLf
                strsql_check += " AND TPlanID='" & s_TPlanID & "'" & vbCrLf
                strsql_check += " AND Years='" & s_Years & "'" & vbCrLf

            Case CstUpdate
                s_OldClassID = TIMS.ClearSQM(s_OldClassID)
                s_ClassID = TIMS.ClearSQM(s_ClassID)
                If s_OldClassID <> s_ClassID Then '統編有修改過
                    strsql_check = ""
                    '**by Milor 20080508--班級代碼重複應該是用區域別+計畫別去判斷，所以加入了TPlanID的判斷式
                    'strsql_check = ""
                    'strsql_check += " select " & vbCrLf
                    'strsql_check += " CLSID,CLASSID,CLASSNAME,CLASSENAME,TPLANID, " & vbCrLf
                    'strsql_check += " SUBSTRING(CONTENT, 1, 4000) CONTENT, " & vbCrLf
                    'strsql_check += " TMID, DistID, MODIFYACCT, MODIFYDATE, CJOB_UNKEY, Years " & vbCrLf
                    strsql_check = ""
                    strsql_check += " SELECT 'X'" & vbCrLf
                    strsql_check += " FROM ID_Class " & vbCrLf
                    strsql_check += " WHERE 1=1 " & vbCrLf
                    strsql_check += " AND CLSID<>" & Val(s_CLSID) & vbCrLf
                    strsql_check += " AND ClassID='" & s_ClassID & "' " & vbCrLf
                    strsql_check += " AND DistID='" & s_DistID & "' " & vbCrLf
                    strsql_check += " AND TPlanID='" & s_TPlanID & "'" & vbCrLf
                    strsql_check += " AND Years='" & s_Years & "'" & vbCrLf
                End If

        End Select

        If strsql_check <> "" Then
            If DbAccess.GetCount(strsql_check, tConn) > 0 Then
                Rst = True
            End If
        End If

        Return Rst
    End Function

    Function Checkdata1(ByRef sErrmsg1 As String) As Boolean
        Dim rst As Boolean = True '正常為 true 異常 false

        TB_classid.Text = TIMS.ClearSQM(UCase(TB_classid.Text))
        TBclass_name.Text = TIMS.ClearSQM(TBclass_name.Text)
        ClassEName.Text = TIMS.ClearSQM(ClassEName.Text)

        sErrmsg1 = ""
        If TB_classid.Text = "" Then sErrmsg1 &= "班別代碼 不可為空!" & vbCrLf
        If TBclass_name.Text = "" Then sErrmsg1 &= "班別名稱 不可為空!" & vbCrLf
        If ClassEName.Text = "" Then sErrmsg1 &= "英文名稱 不可為空!" & vbCrLf
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If iPYNum >= 3 Then
                If trainValue.Value = "" Then sErrmsg1 &= "訓練職類/業別 不可為空!" & vbCrLf
            Else
                If jobValue.Value = "" Then sErrmsg1 &= "訓練業別 不可為空!" & vbCrLf
            End If
        Else
            If trainValue.Value = "" Then sErrmsg1 &= "訓練職類 不可為空!" & vbCrLf
        End If
        If cjobValue.Value = "" Then sErrmsg1 &= "通俗職類 不可為空!" & vbCrLf

        If sErrmsg1 <> "" Then Return False

        Return rst
    End Function

    ''' <summary> 查詢班級學員是否已經存在了 </summary>
    ''' <returns></returns>
    Function Get_CLASS_STUDENTINFO() As DataTable
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List) 'Get_TPlan
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        If v_ddlDistID = "" Then v_ddlDistID = sm.UserInfo.DistID
        If v_Plan_List = "" Then v_Plan_List = sm.UserInfo.TPlanID
        If v_ddlYears = "" Then v_ddlYears = sm.UserInfo.Years

        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT 1 X" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b ON a.OCID = b.OCID" & vbCrLf
        sql &= " JOIN ID_CLASS c ON b.CLSID = c.CLSID" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PLANID=b.PLANID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND a.StudentID LIKE '%" & HidClassID1.Value & "%'" & vbCrLf
        sql &= " AND c.ClassID=@ClassID " & vbCrLf
        sql &= " AND c.clsid=@clsid" & vbCrLf

        sql &= " AND c.DistID=@DistID " & vbCrLf
        sql &= " AND c.TPlanID=@TPlanID " & vbCrLf
        sql &= " AND c.Years=@Years " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("ClassID", HidClassID1.Value)
            .Parameters.Add("clsid", Val(Hid_CLSID.Value))
            .Parameters.Add("DistID", v_ddlDistID)
            .Parameters.Add("TPlanID", v_Plan_List)
            .Parameters.Add("Years", v_ddlYears)
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function

    ''' <summary> 取出要更新的學員資料 </summary>
    Function Get_CLASS_STUDENTINFO2() As DataTable
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List) 'Get_TPlan
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        If v_ddlDistID = "" Then v_ddlDistID = sm.UserInfo.DistID
        If v_Plan_List = "" Then v_Plan_List = sm.UserInfo.TPlanID
        If v_ddlYears = "" Then v_ddlYears = sm.UserInfo.Years

        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT a.SOCID,a.OCID,c.CLSID" & vbCrLf
        sql &= " ,b.YEARS CLASSYEARS,RIGHT(ip.YEARS,2) YEARS2" & vbCrLf
        sql &= " ,ISNULL(c.CLASSID2,c.CLASSID) CLASSID" & vbCrLf
        sql &= " ,B.CYCLTYPE" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(a.STUDENTID) CSTUDID2" & vbCrLf
        sql &= " ,a.STUDENTID" & vbCrLf
        sql &= " ,concat(b.YEARS,'0',ISNULL(c.CLASSID2,c.CLASSID),ISNULL(b.CYCLTYPE,'01'),dbo.FN_CSTUDID2(a.STUDENTID)) NEWSTUDID" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b ON a.OCID = b.OCID" & vbCrLf
        sql &= " JOIN ID_CLASS c ON b.CLSID = c.CLSID" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PLANID=b.PLANID" & vbCrLf
        sql &= " WHERE concat(b.YEARS,'0',ISNULL(c.CLASSID2,c.CLASSID),ISNULL(b.CYCLTYPE,'01'),dbo.FN_CSTUDID2(a.STUDENTID)) !=a.STUDENTID" & vbCrLf

        sql &= " AND c.CLSID=@CLSID" & vbCrLf
        sql &= " AND c.DistID=@DistID " & vbCrLf
        sql &= " AND c.TPlanID=@TPlanID " & vbCrLf
        sql &= " AND c.Years=@Years " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            '.Parameters.Add("ClassID", HidClassID1.Value)
            .Parameters.Add("CLSID", Val(Hid_CLSID.Value))
            .Parameters.Add("DistID", v_ddlDistID)
            .Parameters.Add("TPlanID", v_Plan_List)
            .Parameters.Add("Years", v_ddlYears)
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function

    ''' <summary> 班級學員檔已有資料,班別代碼不可修改 </summary>
    ''' <returns></returns>
    Function Checkdata2() As Boolean
        Dim rst As Boolean = True '正常為 true 異常 false

        Select Case ProcessType
            Case "Update"
                If TB_classid.Text = hTB_classid.Value Then Return rst '外部帶入 等同 目前修改值 (沒有做代碼的異動)
                HidClassID1.Value = TIMS.ClearSQM(HidClassID1.Value)
            Case Else
                Return rst
        End Select
        If HidClassID1.Value = "" Then Return rst


        '管理者權限足夠／開啟修改學員學號／勾選修改學員學號 可直接修改
        If flgROLEIDx0xLIDx0 AndAlso tr_cb_USTUDENTID.Visible AndAlso cb_USTUDENTID.Checked Then Return True

        Dim dtStd As DataTable = Get_CLASS_STUDENTINFO()
        If dtStd.Rows.Count > 0 Then Return False '正常為 true 異常 false
        Return rst
    End Function

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click

        Dim sErrmsg1 As String = ""
        Dim flag_check As Boolean = Checkdata1(sErrmsg1)
        If sErrmsg1 <> "" Then
            Common.MessageBox(Me, sErrmsg1)
            Exit Sub
        End If
        If Not Page.IsValid Then
            Common.MessageBox(Me, TIMS.cst_SAVENGMsg1)
            Exit Sub '驗證有誤
        End If

        HidClassID1.Value = TIMS.ClearSQM(HidClassID1.Value)
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List)
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)

        '驗証區---Start
        '**by Milor 20080508--班級代碼重複應該是用區域別+計畫別去判斷，所以加入了TPlanID的判斷式----start
        '判斷是否有新增/修改重複的資料
        Dim flag_Double As Boolean = False
        Dim s_parms As New Hashtable
        s_parms.Add("ClassID", TB_classid.Text)
        s_parms.Add("DistID", v_ddlDistID)
        s_parms.Add("TPlanID", v_Plan_List)
        s_parms.Add("Years", v_ddlYears)
        s_parms.Add("CLSID", Val(Hid_CLSID.Value))
        s_parms.Add("OldClassID", HidClassID1.Value)
        flag_Double = Check_Double(ProcessType, s_parms, objconn)

        'Dim strScript1 As String = ""
        If flag_Double Then
            Select Case ProcessType
                Case "Insert"
                    sErrmsg1 = "新增班別代碼重複!!!!"
                    Common.MessageBox(Me, sErrmsg1)
                    Exit Sub
                Case "Copy"
                    sErrmsg1 = "複製班別代碼重複!!!!"
                    Common.MessageBox(Me, sErrmsg1)
                    Exit Sub
                Case "Update"
                    '判斷是否有修改重複的資料
                    sErrmsg1 = "班別代碼設定重複!!!!"
                    Common.MessageBox(Me, sErrmsg1)
                    Exit Sub
            End Select
        End If

        Dim flag_check2 As Boolean = Checkdata2()
        If Not flag_check2 Then
            sErrmsg1 = "班級學員檔已有資料, 班別代碼不可修改!!!!"
            Common.MessageBox(Me, sErrmsg1)
            Exit Sub
        End If
        '驗証區---End

        '儲存
        Call SaveData1()

        '更新的學員資料
        Call SaveData_UpdateStudentID()

        If Session(cst_sess_sch1_txt) Is Nothing Then
            Session(cst_sess_sch1_txt) = Me.ViewState(cst_sess_sch1_txt)
        End If

        Dim strScript1 As String = ""
        Select Case ProcessType
            Case "Insert", "Copy"
                strScript1 = "<script language=""javascript"">" + vbCrLf
                strScript1 += "alert('班別代碼儲存成功!!');" + vbCrLf
                strScript1 += "location.href='TC_01_003.aspx?ID=" & Re_ID.Value & "';" + vbCrLf
                strScript1 += "</script>"
                Page.RegisterStartupScript("", strScript1)

            Case "Update"
                strScript1 = "<script language=""javascript"">" + vbCrLf
                strScript1 += "alert('班別代碼修改成功!!');" + vbCrLf
                strScript1 += "location.href='TC_01_003.aspx?ID=" & Re_ID.Value & "';" + vbCrLf
                strScript1 += "</script>"
                Page.RegisterStartupScript("", strScript1)

        End Select

    End Sub

    ''' <summary> 更新的學員資料 </summary>
    Sub SaveData_UpdateStudentID()
        Dim flag_action As Boolean = False '不可執行此函數
        '儲存成功且修改狀態，且有班別代碼 且有班別代碼的序號 'OldClassID!= TB_classid.Text
        '管理者權限足夠／開啟修改學員學號／勾選修改學員學號 可直接修改 
        If ProcessType = "Update" AndAlso HidClassID1.Value <> "" AndAlso Hid_CLSID.Value <> "" Then
            If flgROLEIDx0xLIDx0 AndAlso tr_cb_USTUDENTID.Visible AndAlso cb_USTUDENTID.Checked Then
                flag_action = True '可執行此函數
                'Call SaveData_UpdateStudentID()
            End If
        End If
        '不可執行此函數
        If Not flag_action Then Return

        Dim dt2 As DataTable = Get_CLASS_STUDENTINFO2()
        If dt2 Is Nothing Then Return
        If dt2.Rows.Count = 0 Then Return

        Dim u_sql As String = ""
        u_sql &= " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
        u_sql &= " SET STUDENTID=@NEWSTUDID" & vbCrLf
        u_sql &= " WHERE OCID=@OCID and SOCID=@SOCID AND STUDENTID=@STUDENTID" & vbCrLf
        For Each dr2 As DataRow In dt2.Rows
            Dim u_par As New Hashtable
            u_par.Add("NEWSTUDID", dr2("NEWSTUDID"))
            u_par.Add("OCID", dr2("OCID"))
            u_par.Add("SOCID", dr2("SOCID"))
            u_par.Add("STUDENTID", dr2("STUDENTID"))
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_par)
        Next
    End Sub

    ''' <summary> 儲存 </summary>
    Sub SaveData1()
        'Dim rst As Boolean = True
        '儲存區---Start
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List)
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        If v_ddlDistID = "" Then v_ddlDistID = sm.UserInfo.DistID
        If v_Plan_List = "" Then v_Plan_List = sm.UserInfo.TPlanID
        If v_ddlYears = "" Then v_ddlYears = sm.UserInfo.Years

        Dim tmpTMID As String = trainValue.Value
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            tmpTMID = If(iPYNum >= 3, trainValue.Value, jobValue.Value)
        End If

        If ComSumm.Text <> "" Then ComSumm.Text = Trim(ComSumm.Text)
        Dim sql As String = ""
        Select Case ProcessType
            Case "Insert", "Copy"
                '新增
                sql = ""
                sql &= " INSERT INTO ID_CLASS ( " & vbCrLf
                sql &= " CLSID,CLASSID,CLASSNAME,CLASSENAME,TPLANID,CONTENT " & vbCrLf
                sql &= " ,TMID,DISTID,MODIFYACCT,MODIFYDATE,CJOB_UNKEY,YEARS " & vbCrLf
                sql &= " ) VALUES ( " & vbCrLf
                sql &= " @CLSID,@CLASSID,@CLASSNAME,@CLASSENAME,@TPLANID,@CONTENT " & vbCrLf
                sql &= " ,@TMID,@DISTID,@MODIFYACCT,GETDATE(),@CJOB_UNKEY,@YEARS " & vbCrLf
                sql &= " ) " & vbCrLf
                Dim iCmd As New SqlCommand(sql, objconn)
                '同一轄區不可有相同的ClassID
                Dim iCLSID As Integer = DbAccess.GetNewId(objconn, "ID_CLASS_CLSID_SEQ,ID_CLASS,CLSID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("CLSID", SqlDbType.Int).Value = iCLSID
                    .Parameters.Add("CLASSID", SqlDbType.VarChar).Value = TB_classid.Text
                    .Parameters.Add("CLASSNAME", SqlDbType.NVarChar).Value = TBclass_name.Text
                    .Parameters.Add("CLASSENAME", SqlDbType.NVarChar).Value = ClassEName.Text
                    .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_Plan_List ' Me.Plan_List.SelectedValue
                    .Parameters.Add("CONTENT", SqlDbType.NVarChar).Value = If(String.IsNullOrEmpty(Me.ComSumm.Text), Convert.DBNull, Me.ComSumm.Text)
                    .Parameters.Add("TMID", SqlDbType.VarChar).Value = tmpTMID
                    .Parameters.Add("DISTID", SqlDbType.VarChar).Value = v_ddlDistID ' sm.UserInfo.DistID
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                    .Parameters.Add("CJOB_UNKEY", SqlDbType.VarChar).Value = cjobValue.Value
                    .Parameters.Add("YEARS", SqlDbType.VarChar).Value = v_ddlYears 'sm.UserInfo.Years
                    DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                End With
                Hid_CLSID.Value = Convert.ToString(iCLSID)
            Case "Update"
                '修改
                sql = ""
                sql &= " UPDATE ID_CLASS " & vbCrLf
                sql &= " SET CLASSID=@CLASSID " & vbCrLf
                sql &= " ,CLASSNAME=@CLASSNAME " & vbCrLf
                sql &= " ,CLASSENAME=@CLASSENAME " & vbCrLf
                sql &= " ,TPLANID=@TPLANID " & vbCrLf
                sql &= " ,CONTENT=@CONTENT " & vbCrLf
                sql &= " ,TMID=@TMID " & vbCrLf
                sql &= " ,DISTID=@DISTID " & vbCrLf
                sql &= " ,MODIFYACCT=@MODIFYACCT " & vbCrLf
                sql &= " ,MODIFYDATE=GETDATE() " & vbCrLf
                sql &= " ,CJOB_UNKEY=@CJOB_UNKEY " & vbCrLf
                sql &= " ,YEARS=@YEARS " & vbCrLf
                sql &= " WHERE CLSID=@CLSID"
                Dim uCmd As New SqlCommand(sql, objconn)

                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("CLSID", SqlDbType.Int).Value = Val(Hid_CLSID.Value) 'sRe_clsid
                    .Parameters.Add("CLASSID", SqlDbType.VarChar).Value = TB_classid.Text
                    .Parameters.Add("CLASSNAME", SqlDbType.NVarChar).Value = TBclass_name.Text
                    .Parameters.Add("CLASSENAME", SqlDbType.NVarChar).Value = ClassEName.Text
                    .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_Plan_List 'Me.Plan_List.SelectedValue
                    .Parameters.Add("CONTENT", SqlDbType.NVarChar).Value = If(String.IsNullOrEmpty(Me.ComSumm.Text), Convert.DBNull, Me.ComSumm.Text)
                    .Parameters.Add("TMID", SqlDbType.VarChar).Value = tmpTMID
                    .Parameters.Add("DISTID", SqlDbType.VarChar).Value = v_ddlDistID ' sm.UserInfo.DistID
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    '.Parameters.Add("MODIFYDATE", SqlDbType.DateTime).Value = Now
                    .Parameters.Add("CJOB_UNKEY", SqlDbType.VarChar).Value = cjobValue.Value
                    .Parameters.Add("YEARS", SqlDbType.VarChar).Value = v_ddlYears 'sm.UserInfo.Years
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                End With
        End Select

        'Return rst
        '儲存區---End
    End Sub



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session(cst_sess_sch1_txt) Is Nothing Then
            Session(cst_sess_sch1_txt) = Me.ViewState(cst_sess_sch1_txt)
        End If
        TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "TC_01_003.aspx?ID=" & Request("ID"))
    End Sub
End Class

