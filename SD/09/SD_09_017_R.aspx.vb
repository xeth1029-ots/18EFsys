Partial Class SD_09_017_R
    Inherits AuthBasePage

    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1 = FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            Call Create1()
        End If

        listYears.Enabled = False
        DistID.Enabled = False
        listTPlanID.Enabled = False
        BtnQuery.Enabled = False
        Select Case sm.UserInfo.LID
            Case "0"
                listYears.Enabled = True
                DistID.Enabled = True
                listTPlanID.Enabled = True
                BtnQuery.Enabled = True
            Case "1"
                BtnQuery.Enabled = True
            Case "2"
                Common.MessageBox(Me, "此功能暫不開放給委訓單位喔!!")
                Exit Sub
        End Select

        '查詢
        'BtnQuery.Attributes("onclick") = "return ChkSearch();"
    End Sub

    Sub Create1()
        msg.Text = ""
        Table4.Visible = False
        PageControler1.Visible = False

        '取出鍵詞-查詢原因-INQUIRY
        Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        listYears = TIMS.Get_Years(Me.listYears)
        DistID = TIMS.Get_DistID(Me.DistID)
        listTPlanID = TIMS.Get_TPlan(Me.listTPlanID)
        listIdentity = TIMS.Get_Identity(Me.listIdentity, 31, objconn)

        Common.SetListItem(listYears, sm.UserInfo.Years)
        Common.SetListItem(DistID, sm.UserInfo.DistID)
        Common.SetListItem(listTPlanID, sm.UserInfo.TPlanID)
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If listYears.SelectedValue = "" Then
            Errmsg += "請選擇 計畫年度!!" & vbCrLf
        End If
        If DistID.SelectedValue = "" Then
            'Errmsg += "請選擇  轄區中心!!" & vbCrLf
            Errmsg += "請選擇  轄區分署!!" & vbCrLf
        End If
        If listTPlanID.SelectedValue = "" Then
            Errmsg += "請選擇 訓練計畫!!" & vbCrLf
        End If
        If listIdentity.SelectedValue = "" Then
            Errmsg += "請選擇 身分別!!" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        '計畫年度,轄區分署,訓練計畫,身分別,
        'listYears,DistID,listTPlanID,listIdentity,
        Dim V_listYears As String = TIMS.GetListValue(listYears)
        Dim V_DistID As String = TIMS.GetListValue(DistID)
        Dim V_listTPlanID As String = TIMS.GetListValue(listTPlanID)
        Dim V_listIdentity As String = TIMS.GetListValue(listIdentity)

        If V_listYears <> "" Then RstMemo &= String.Concat("&計畫年度=", V_listYears)
        If V_DistID <> "" Then RstMemo &= String.Concat("&轄區分署=", V_DistID)
        If V_listTPlanID <> "" Then RstMemo &= String.Concat("&訓練計畫=", V_listTPlanID)
        If V_listIdentity <> "" Then RstMemo &= String.Concat("&身分別=", V_listIdentity)
        Return RstMemo
    End Function

    Sub SSCHEAR1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim sql As String = ""
        sql &= " SELECT cs.socid" & vbCrLf
        sql &= " ,ip.years" & vbCrLf
        sql &= " ,ip.planid" & vbCrLf
        sql &= " ,id1.distid" & vbCrLf
        sql &= " ,id1.name distname" & vbCrLf
        sql &= " ,ip.tplanid" & vbCrLf
        sql &= " ,ip.planname planame" & vbCrLf
        sql &= " ,ip.planKind" & vbCrLf
        sql &= " ,oo.orgid" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,cc.tmid" & vbCrLf
        sql &= " ,tt.trainname" & vbCrLf
        sql &= " ,cc.stdate" & vbCrLf
        sql &= " ,cc.ftdate" & vbCrLf
        sql &= " ,ss.name" & vbCrLf
        sql &= " ,ss.idno" & vbCrLf
        sql &= " ,ss.birthday" & vbCrLf
        sql &= " ,ss2.phoneD" & vbCrLf
        sql &= " ,ss2.phoneN" & vbCrLf
        sql &= " ,ss2.zipcode1" & vbCrLf
        'sql &= " ,ss2.zipcode1_6W" & vbCrLf
        sql &= " ,iz.ZipName" & vbCrLf
        sql &= " ,ss2.address" & vbCrLf
        sql &= " ,CASE WHEN ss.sex='M' THEN '男' WHEN ss.sex='F' THEN '女' ELSE ss.sex END sex" & vbCrLf
        sql &= " ,cs.studStatus" & vbCrLf
        sql &= " ,cs.MIdentityID" & vbCrLf
        'Sql += " ,ISNULL(sg.IsGetJob,'0') IsGetJob" & vbCrLf
        'Sql += " ,case" & vbCrLf
        'Sql += " 	when ISNULL(sg.IsGetJob,'0')='1' then '1.就業' " & vbCrLf
        'Sql += " 	when ISNULL(sg.IsGetJob,'0')='2' then '2.不就業' " & vbCrLf
        'Sql += " 	else '0.未就業' end IsGetJobName" & vbCrLf
        sql &= " FROM view_plan ip" & vbCrLf
        sql &= " join id_district id1 on id1.distid =ip.distid" & vbCrLf
        sql &= " join class_classinfo cc on cc.planid =ip.planid" & vbCrLf
        sql &= " join plan_planinfo pp on pp.planid =cc.planid and pp.comidno =cc.comidno and pp.seqno =cc.seqno" & vbCrLf
        sql &= " join org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        sql &= " left join view_trainType tt on tt.tmid=cc.tmid" & vbCrLf
        sql &= " join class_studentsofclass cs on cs.ocid=cc.ocid" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " join stud_subdata ss2 on ss2.sid =ss.sid" & vbCrLf
        sql &= " join view_zipName iz on iz.zipcode=ss2.zipcode1" & vbCrLf
        'Sql += " left join Stud_GetJobState3 sg on sg.CPoint=1 and sg.socid =cs.socid" & vbCrLf
        sql &= " where cc.IsSuccess ='Y'" & vbCrLf
        sql &= " and cc.NotOpen='N'" & vbCrLf
        sql &= " and cs.studstatus not in (2,3)" & vbCrLf

        sql &= " and ip.years=@years" & vbCrLf
        sql &= " and ip.distid=@distid" & vbCrLf
        sql &= " and ip.tplanid=@tplanid" & vbCrLf
        sql &= " and cs.MIdentityID=@MIdentityID" & vbCrLf

        sql &= " ORDER BY ip.years,ip.tplanid,ip.planid,cc.classcname,cc.ocid,ss.idno" & vbCrLf

        Dim v_listYears As String = TIMS.GetListValue(listYears)
        Dim v_DistID As String = TIMS.GetListValue(DistID)
        Dim v_listTPlanID As String = TIMS.GetListValue(listTPlanID)
        Dim v_listIdentity As String = TIMS.GetListValue(listIdentity)

        Dim parms As New Hashtable
        parms.Add("years", v_listYears)
        parms.Add("distid", v_DistID)
        parms.Add("tplanid", v_listTPlanID)
        parms.Add("MIdentityID", v_listIdentity)
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "YEARS,DISTNAME,PLANAME,ORGNAME,CLASSCNAME,NAME,SEX,MIDENTITYID,IDNO,BIRTHDAY")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, "", sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        msg.Text = "查無資料!!"
        Table4.Visible = False
        PageControler1.Visible = False
        If TIMS.dtNODATA(dt) Then Return

        msg.Text = ""
        Table4.Visible = True
        PageControler1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub
    '查詢
    Private Sub BtnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnQuery.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call SSCHEAR1()
    End Sub
End Class
