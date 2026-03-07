Partial Class SD_05_008_C
    Inherits AuthBasePage

#Region "Web Form"
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()

        'Dim sql As String
        'Dim dr As DataRow
        'Dim dt As DataTable
        ''設定欄位輸入字數的限制
        ''資料表名稱 '欄位名稱 '欄位型態 '欄位長度
        'sql = ""
        'sql &= " SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,CHAR_LENGTH "
        'sql &= " FROM USER_TAB_COLUMNS "
        'sql &= " WHERE TABLE_NAME IN ('STUD_DATALID') "
        'sql &= " AND DATA_TYPE IN ('NVARCHAR2','VARCHAR2','CHAR') "
        'objconn = DbAccess.GetConnection()
        'dt = DbAccess.GetDataTable(sql, objconn)
        'Call TIMS.CloseDbConn(objconn)

        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("STUD_DATALID")
        For Each dr As DataRow In dt.Select("DATA_TYPE IN ('NVARCHAR2','VARCHAR2','CHAR')")
            Select Case UCase(dr("COLUMN_NAME"))
                Case "TRAINCOMD2OTHER"
                    TrainComd2Other.MaxLength = TIMS.CINT1(dr("CHAR_LENGTH"))
                    TrainComd2Other.ToolTip = $"限欄位長度  {dr("CHAR_LENGTH")} 字元"
            End Select
        Next
    End Sub
#End Region

    'update table : STUD_DATALID   
    Dim aSTDNAME As String = ""
    Dim aSTUDENTID As String = ""
    Dim aSTDPID As String = ""
    Dim aSEX As String = ""
    Dim aBIRTHDAY As String = ""
    Dim aDegreeID As String = ""
    Dim aMilitaryID As String = ""
    Dim aIDENTITYIDx1 As String = ""
    Dim aIDENTITYIDx2 As String = ""
    Dim aIDENTITYIDx3 As String = ""
    Dim aIDENTITYIDx4 As String = ""
    Dim aIDENTITYIDx5 As String = ""

    Const cst_Juzhu As String = "1" 'cst_Juzhu 1 署(局)屬 
    Const cst_NonJuzhu As String = "2" 'cst_NonJuzhu 2 非署(局)屬

    Const cst_search As String = "search"
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not Session(cst_search) Is Nothing Then Session(cst_search) = Session(cst_search)

        '加入物件資料-訓練計畫
        If Not IsPostBack Then
            HidDLID.Value = ""
            'If Not Session(cst_search) Is Nothing Then
            '    Me.ViewState("search") = Session(cst_search)
            '    Session(cst_search) = Nothing
            'End If
            '非署(局)屬機構
            UnitCode1 = TIMS.Get_othUnitCode(UnitCode1, 0, objconn)

            Dim sql As String = "SELECT * FROM KEY_PLAN ORDER BY TPLANID"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
            With TPlanID1
                .DataSource = dt
                .DataTextField = "PlanName"
                .DataValueField = "TPlanID"
                .DataBind()
                .Items.Insert(0, New ListItem("--請選擇--", 0))
            End With

            '判斷可以顯示的部份
            Call sUtl_View1()

            '建立資料
            Call Create()

        End If

        Train1.Attributes("onclick") = "change()"
        Train2.Attributes("onclick") = "change()"
    End Sub

    Sub sUtl_View1()
        Dim rqProecess As String = Convert.ToString(Request("Proecess"))
        rqProecess = TIMS.ClearSQM(rqProecess)
        Dim rqDLID As String = Convert.ToString(Request("DLID"))
        rqDLID = TIMS.ClearSQM(rqDLID)

        Select Case rqProecess
            Case "add", "addother", "editother"
                UnitCode.Visible = False
                UnitCode1.Visible = True
                TPlanID.Visible = False
                TPlanID1.Visible = True
                ClassName.Visible = False
                ClassName1.Visible = True
                TrainName.Visible = False
                TB_career_id.Visible = True
                Button2.Visible = True
                ResultCount.Visible = False
                ResultCount1.Visible = True
                ResultDate.Visible = False
                ResultDate1.Visible = True
                TrainingTHour.Visible = False
                TrainingTHour1.Visible = True
                IMG1.Visible = True
            Case "addmy", "editmy"
                UnitCode.Visible = True '署(局)屬機構顯示。
                UnitCode1.Visible = False
                TPlanID.Visible = True
                TPlanID1.Visible = False
                ClassName.Visible = True
                ClassName1.Visible = False
                TrainName.Visible = True
                TB_career_id.Visible = False
                Button2.Visible = False
                ResultCount.Visible = True
                ResultCount1.Visible = False
                ResultDate.Visible = True
                ResultDate1.Visible = False
                TrainingTHour.Visible = True
                TrainingTHour1.Visible = False
                IMG1.Visible = False
        End Select

        btnSaveData1.Visible = False
        btnWriteStud.Visible = False
        Select Case rqProecess
            Case "editmy", "editother"
                '非署(局)屬
                btnSaveData1.Visible = True '填寫學員資料。
            Case Else
                btnWriteStud.Visible = True '儲存
        End Select

        trImport1.Visible = False
        HidDLID.Value = rqDLID
        If HidDLID.Value <> "" Then HidDLID.Value = Trim(HidDLID.Value)
        If HidDLID.Value <> "" Then
            trImport1.Visible = True
        End If

    End Sub

    '查詢
    Sub Create()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqProecess As String = TIMS.ClearSQM(Request("Proecess"))
        Dim rqDLID As String = TIMS.ClearSQM(Request("DLID"))

        Select Case rqProecess
            Case "addmy"
                'If rqOCID = "" Then Exit Sub
                If rqOCID = "" Then
                    'Common.MessageBox(Page, "局屬資訊遺失，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "署屬資訊遺失，請重新操作查詢畫面。")
                    Exit Sub
                End If

                Dim pms1 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}}
                Dim sql As String = ""
                sql &= " SELECT a.ClassCName, c.OrgName, a.CyclType, a.LevelType, e.PlanName,a.OCID "
                sql &= " ,f.TrainName, f.TrainID, f.TMID, a.FTDate, a.THours, g.num, c.OrgID,e.TPlanID,a.RID ,d.DISTID "
                sql &= " FROM Class_ClassInfo a"
                sql &= " JOIN Auth_Relship b ON a.RID=b.RID "
                sql &= " JOIN Org_OrgInfo c ON b.OrgID=c.OrgID "
                sql &= " JOIN ID_Plan d ON d.PlanID=a.PlanID "
                sql &= " JOIN Key_Plan e ON d.TPlanID=e.TPlanID "
                sql &= " JOIN Key_TrainType f ON a.TMID=f.TMID "
                sql &= " LEFT JOIN (SELECT OCID,count(1) num FROM Class_StudentsOfClass WHERE OCID=@OCID and StudStatus='5' GROUP BY OCID) g ON g.OCID=a.OCID"
                sql &= " WHERE a.OCID=@OCID"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)

                If Not dr Is Nothing Then
                    UnitCode.Text = dr("OrgName").ToString
                    'select * from ID_StatistDist where Type = 1 order by statID
                    Dim sDISTID As String = "" 'sm.UserInfo.DistID
                    sDISTID = Convert.ToString(dr("DistID"))

                    Select Case sDISTID '參考統計室-分署代碼表
                        Case "001" '北分署(職訓局北區中心)
                            UnitCodeValue.Value = "003"
                        Case "002" '泰山訓練場(職訓局泰山中心)
                            UnitCodeValue.Value = "001"
                        Case "003" '桃分署(職訓局桃園中心)
                            UnitCodeValue.Value = "005" '?:008
                        Case "004" '中分署(職訓局中區中心)
                            UnitCodeValue.Value = "002"
                        Case "005" '南分署(職訓局台南中心)
                            UnitCodeValue.Value = "006" '?:009
                        Case "006" '高分署(職訓局南區中心)
                            UnitCodeValue.Value = "004"
                        Case "000" '署(職訓局-未編碼)
                            UnitCodeValue.Value = "000"
                        Case Else
                            'Common.RespWrite(Me, "<script>alert('參考統計室-職訓中心代碼表-未建立');</script>")
                            Common.RespWrite(Me, "<script>alert('參考統計室-分署代碼表-未建立');</script>")
                            Common.RespWrite(Me, "<script>top.location.href='../../index';</script>")
                            Response.End()
                    End Select
                    RIDValue.Value = dr("RID").ToString
                    TPlanID.Text = dr("PlanName").ToString
                    TPlanIDValue.Value = dr("TPlanID").ToString
                    OCID.Value = dr("OCID").ToString

                    ClassName.Text = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))

                    If Not IsDBNull(dr("LevelType")) Then
                        If CInt(dr("LevelType")) <> 0 Then
                            ClassName.Text += "第" & TIMS.GetChtNum(CInt(dr("LevelType"))) & "階段"
                        End If
                    End If

                    TrainName.Text = "[" & Convert.ToString(dr("TrainID")) & "]" & Convert.ToString(dr("TrainName"))
                    trainValue.Value = Convert.ToString(dr("TMID"))
                    ResultCount.Text = IIf(dr("num").ToString = "", 0, dr("num"))
                    If Convert.ToString(dr("FTDate")) <> "" Then
                        ResultDate.Text = FormatDateTime(Convert.ToString(dr("FTDate")), DateFormat.ShortDate)
                    End If
                    TrainingTHour.Text = Convert.ToString(dr("THours"))
                End If
                btnWriteStud.Attributes("onclick") = "javascript:return chkdata(1)"
            Case "addother"
                btnWriteStud.Attributes("onclick") = "javascript:return chkdata(2)"
            Case "editmy", "editother"
                HidDLID.Value = rqDLID
                '非署(局)屬
                Dim pms1 As New Hashtable From {{"DLID", TIMS.CINT1(rqDLID)}}
                Dim sql As String = ""
                sql &= " SELECT a.UnitCode,a.RID,a.ClassName,a.OCID,a.Trainice,a.TrainCommend1,a.TrainCommend2" & vbCrLf
                sql &= " ,a.SchoolTime,a.ResultCount,a.OCID,a.ResultDate,a.TrainingTHour,d.TPlanID,d.PlanName" & vbCrLf
                sql &= " ,b.TMID,b.TrainID,b.TrainName" & vbCrLf
                sql &= " FROM STUD_DATALID a " & vbCrLf
                sql &= " LEFT JOIN Key_TrainType b ON a.TMID=b.TMID  " & vbCrLf
                sql &= " LEFT JOIN Key_Plan d ON a.TPlanID=d.TPlanID" & vbCrLf
                sql &= " WHERE a.DLID=@DLID" & vbCrLf
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1)
                If dr IsNot Nothing Then
                    '38416C48D9684B4AA33F4372FFE14DB0A8B1429E24D039BD13264E93539002946ACECDEEDD0EFC2703B176C8862F445FE10A6ABA01E1A452F4792867
                    '98E5798B28B8A7B306C8ED8B7013CC5FAADED74043E7CAF61DF03B0C65E3317B7DA1D498CF06CD18F5A05B8252710A9C9091BD139C13784E3F5C9D02
                    '2904DFAAB1CA093FC63223F75498122941297F80586786EE96BD01D79B442611581464BD6FE222ACD5A5A5BB6209A1EB3CFDC1E2E02E1FF3BA9E670D
                    '4D503E181CDECCC60DA0F4A1BC21720CF4C5EDE9E1B89CE35D8B1280A121BC0B4904644F752CF81F36BB7E4233FAC66B2DB60218CFE026C9B8A304D3
                    '33304BBE3B318CDDE9AB42666ACEDB25B9AB7B0A551972BF27838E21985EDB610284F468D9207D4CD544E55FA52A2AB504740C6BB0F6B440815F6488
                    '21372A9775831AE1DB64D403A8BD595AA9D9F17C6BAA493301DF519823F06E3021C584502D9B265F328F65BCC04498F21B695121DA475E052BC2730E
                    '74D1668FC09C020F082417511752C1E39CE042E0CA168D43A90CDFB8D730254CD94214AA37DDBEFB1FD38BBCB3ACF47FBFD465E7971F1C0966F9D896
                    '439DFF465D9CB4A189059296515E4C39A09929AC4C4E19C6E2D9F8EA908E2AE19332106711F9ED7E884B9F9E955D4995702A36A3D318C17D2D77F888

                    UnitCode.Text = ""
                    If Convert.ToString(dr("UnitCode")) <> "" Then
                        '參考統計室-分署代碼表 (非局屬代號取得局屬)
                        UnitCode.Text = TIMS.Get_UnitCodeName1(Convert.ToString(dr("UnitCode")))
                        Common.SetListItem(UnitCode1, dr("UnitCode").ToString)
                    End If

                    UnitCodeValue.Value = Convert.ToString(dr("UnitCode"))
                    RIDValue.Value = dr("RID").ToString

                    TPlanID.Text = dr("PlanName").ToString
                    Common.SetListItem(TPlanID1, dr("TPlanID").ToString)
                    TPlanIDValue.Value = dr("TPlanID").ToString

                    ClassName.Text = dr("ClassName").ToString
                    ClassName1.Text = dr("ClassName").ToString
                    OCID.Value = dr("OCID").ToString

                    TrainName.Text = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
                    TB_career_id.Text = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
                    trainValue.Value = dr("TMID").ToString

                    Common.SetListItem(Trainice, dr("Trainice").ToString)
                    If dr("TrainCommend1").ToString = "Y" Then
                        Train2.Checked = True
                    ElseIf dr("TrainCommend1").ToString = "N" Then
                        Train1.Checked = True
                    End If
                    Common.SetListItem(TrainCommend2, dr("TrainCommend2").ToString)
                    Common.SetListItem(SchoolTime, dr("SchoolTime").ToString)

                    ResultCount.Text = dr("ResultCount").ToString
                    ResultCount1.Text = dr("ResultCount").ToString
                    Select Case rqProecess
                        Case "editmy"
                            Dim iStudentCount As Integer = 0
                            If $"{dr("OCID")}" <> "" Then
                                Dim pms11 As New Hashtable From {{"OCID", TIMS.CINT1(dr("OCID"))}}
                                Dim sql11 As String = " SELECT COUNT(1) STUDENTCOUNT FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID AND StudStatus=5"
                                iStudentCount = DbAccess.ExecuteScalar(sql11, objconn, pms11)
                            End If
                            If iStudentCount > 0 Then
                                If iStudentCount <> TIMS.CINT1(dr("ResultCount")) Then
                                    ResultCount.Text = iStudentCount
                                    Person.Text = "人(之前封面結訓人數和目前統計不相同，存檔後會更新封面結訓人數)"
                                End If
                            End If

                            '若是使用者新增封面後,更改結訓日期,之後再按修改查看,結訓日期會跟著改
                            Dim pms12 As New Hashtable From {{"OCID", TIMS.CINT1(dr("OCID"))}}
                            Dim sql12 As String = "SELECT FTDATE FROM CLASS_CLASSINFO WHERE OCID=@OCID"
                            Dim FTDate As DateTime = DbAccess.ExecuteScalar(sql12, objconn, pms12)

                            If TIMS.Cdate3(FTDate) <> TIMS.Cdate3($"{dr("ResultDate")}") Then
                                ResultDate.Text = TIMS.Cdate3(FTDate)
                            ElseIf TIMS.Cdate3(FTDate) = TIMS.Cdate3($"{dr("ResultDate")}") Then
                                ResultDate.Text = TIMS.Cdate3($"{dr("ResultDate")}")
                            End If
                    End Select

                    If $"{dr("ResultDate")}" <> "" Then
                        ResultDate1.Text = TIMS.Cdate3(dr("ResultDate"))
                        'ResultDate1.Text = FormatDateTime(dr("ResultDate").ToString, DateFormat.ShortDate)
                    End If

                    TrainingTHour.Text = dr("TrainingTHour").ToString
                    TrainingTHour1.Text = dr("TrainingTHour").ToString
                End If
        End Select
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trainice.SelectedValue = "" Then
            Errmsg += "請選擇 訓練性質" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '新增/修改1筆封面 '署(局)屬/非署(局)屬
    Sub SaveDataNew1(ByRef iDLID As Integer, ByVal rqJuzhu As String)
        '表示沒有，要新增該班封面
        'Dim iDLID As Integer = 1 '從1開始 '先取出最大的DLID
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        If iDLID = 0 Then
            '新增
            iDLID = DbAccess.GetNewId(objconn, "STUD_DATALID_DLID_SEQ,STUD_DATALID,DLID")
            sql = "SELECT * FROM STUD_DATALID WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("DLID") = iDLID
        Else
            sql = "SELECT * FROM STUD_DATALID WHERE DLID='" & iDLID & "'"
            dt = DbAccess.GetDataTable(sql, da, objconn)
            dr = dt.Rows(0)
        End If

        Select Case rqJuzhu
            Case cst_Juzhu
                dr("UnitCode") = UnitCodeValue.Value '參考統計室-分署代碼表
                dr("TPlanID") = TPlanIDValue.Value
                dr("OCID") = OCID.Value
                dr("ClassName") = ClassName.Text
                dr("TMID") = trainValue.Value
                If Trainice.SelectedValue <> "" Then
                    dr("Trainice") = Trainice.SelectedValue
                Else
                    dr("Trainice") = Convert.DBNull
                End If
                If Train1.Checked = True Then
                    dr("TrainCommend1") = "N"
                ElseIf Train2.Checked = True Then
                    dr("TrainCommend1") = "Y"
                End If
                If Not TrainCommend2.SelectedItem Is Nothing Then
                    dr("TrainCommend2") = TrainCommend2.SelectedValue
                End If
                If TrainComd2Other.Text <> "" Then
                    dr("TrainComd2Other") = TrainComd2Other.Text
                End If
                dr("SchoolTime") = SchoolTime.SelectedValue

                dr("ResultCount") = ResultCount.Text
                dr("ResultDate") = ResultDate.Text
                dr("TrainingTHour") = TrainingTHour.Text
                dr("RID") = RIDValue.Value
            Case cst_NonJuzhu
                dr("UnitCode") = UnitCode1.SelectedValue
                If TPlanID1.SelectedValue <> "" Then
                    dr("TPlanID") = TPlanID1.SelectedValue
                Else
                    dr("TPlanID") = Convert.DBNull
                End If
                dr("OCID") = Convert.DBNull
                dr("ClassName") = ClassName1.Text
                If trainValue.Value <> "" Then
                    dr("TMID") = trainValue.Value
                Else
                    dr("TMID") = Convert.DBNull
                End If
                If Trainice.SelectedValue <> "" Then
                    dr("Trainice") = Trainice.SelectedValue
                Else
                    dr("Trainice") = Convert.DBNull
                End If
                If Train1.Checked = True Then
                    dr("TrainCommend1") = "N"
                ElseIf Train2.Checked = True Then
                    dr("TrainCommend1") = "Y"
                End If
                If Not TrainCommend2.SelectedItem Is Nothing Then
                    dr("TrainCommend2") = TrainCommend2.SelectedValue
                End If
                If TrainComd2Other.Text <> "" Then
                    dr("TrainComd2Other") = TrainComd2Other.Text
                End If
                dr("SchoolTime") = SchoolTime.SelectedValue
                If ResultCount1.Text <> "" Then
                    dr("ResultCount") = ResultCount1.Text
                End If
                If ResultDate1.Text <> "" Then
                    dr("ResultDate") = TIMS.Cdate2(ResultDate1.Text)
                End If
                If TrainingTHour1.Text <> "" Then
                    dr("TrainingTHour") = TrainingTHour1.Text
                End If
        End Select
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    '填寫學員資料 
    Private Sub btnWriteStud_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnWriteStud.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rqOCID As String = Convert.ToString(Request("OCID"))
        rqOCID = TIMS.ClearSQM(rqOCID)
        Dim rqProecess As String = Convert.ToString(Request("Proecess"))
        rqProecess = TIMS.ClearSQM(rqProecess)
        Dim rqJuzhu As String = Convert.ToString(Request("Juzhu"))
        rqJuzhu = TIMS.ClearSQM(rqJuzhu)

        Dim sUrl As String = ""
        Dim sJuzhu As String = "" '取得署(局)屬傳遞參數。
        If rqJuzhu <> "" Then sJuzhu = "&Juzhu=" & rqJuzhu
        Select Case rqJuzhu
            Case cst_Juzhu
                If rqProecess <> "addmy" Then
                    'Common.MessageBox(Page, "局屬資訊異常1，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "署屬資訊異常1，請重新操作查詢畫面。")
                    Exit Sub
                End If
                If rqOCID = "" Then
                    'Common.MessageBox(Page, "局屬資訊遺失2，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "署屬資訊遺失2，請重新操作查詢畫面。")
                    Exit Sub
                End If
                '先檢查是否有重複的OCID
                Dim pms13 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}}
                Dim sql13 As String = "SELECT * FROM STUD_DATALID WHERE OCID=@OCID"
                Dim dr As DataRow = DbAccess.GetOneRow(sql13, objconn, pms13)
                If dr IsNot Nothing Then
                    Common.MessageBox(Me, "此班級有封面存在")
                    Exit Sub
                End If

            Case cst_NonJuzhu
                If rqProecess <> "addother" Then
                    'Common.MessageBox(Page, "非局屬資訊異常1，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "非署屬資訊異常1，請重新操作查詢畫面。")
                    Exit Sub
                End If
                If rqOCID <> "" Then
                    'Common.MessageBox(Page, "非局屬資訊異常2，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "非署屬資訊異常2，請重新操作查詢畫面。")
                    Exit Sub
                End If

            Case Else
                'Common.MessageBox(Page, "局屬/非局屬資訊遺失3，請重新操作查詢畫面。")
                Common.MessageBox(Page, "署屬/非署屬資訊遺失3，請重新操作查詢畫面。")
                Exit Sub
        End Select

        Dim iDLID As Integer = 0 '(新增)
        Call SaveDataNew1(iDLID, rqJuzhu) '新增1筆封面
        If Not Session(cst_search) Is Nothing Then Session(cst_search) = Session(cst_search)

        Select Case rqJuzhu
            Case cst_Juzhu
                '署(局)屬
                sUrl = ""
                sUrl &= "SD_05_008_D.aspx?ID=" & Request("ID")
                sUrl &= sJuzhu
                sUrl &= "&Proecess=addmyall"
                sUrl &= "&DLID=" & iDLID
                sUrl &= "&OCID=" & rqOCID '署(局)屬
                TIMS.Utl_Redirect1(Me, sUrl)

            Case cst_NonJuzhu
                '非署(局)屬
                sUrl = ""
                sUrl &= "SD_05_008_D.aspx?ID=" & Request("ID")
                sUrl &= sJuzhu
                sUrl &= "&Proecess=addother"
                sUrl &= "&DLID=" & iDLID
                TIMS.Utl_Redirect1(Me, sUrl)

        End Select

    End Sub

    'SERVER端 檢查
    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        Dim rqProecess As String = Convert.ToString(Request("Proecess"))
        rqProecess = TIMS.ClearSQM(rqProecess)

        Select Case rqProecess
            Case "add", "addmy", "editmy", "addother", "editother"
            Case Else
                Errmsg += "狀態有誤!!" & rqProecess & vbCrLf
        End Select

        If Errmsg = "" Then
            Select Case rqProecess
                Case "addother", "editother"
                    If UnitCode1.SelectedValue = "" Then
                        Errmsg += "請選擇 訓練機構 " & vbCrLf
                    End If
                Case Else '"editmy" 
                    If UnitCodeValue.Value = "" Then
                        Errmsg += "請選擇 訓練機構 " & vbCrLf
                    End If
            End Select
            If Trainice.SelectedValue = "" Then
                Errmsg += "請選擇 訓練性質" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub btnSaveData1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveData1.Click
        Dim Errmsg As String = ""
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim rqProecess As String = Convert.ToString(Request("Proecess"))
        rqProecess = TIMS.ClearSQM(rqProecess)
        Dim rqDLID As String = Convert.ToString(Request("DLID"))
        rqDLID = TIMS.ClearSQM(rqDLID)
        Dim rqOCID As String = Convert.ToString(Request("OCID"))
        rqOCID = TIMS.ClearSQM(rqOCID)
        Dim rqJuzhu As String = Convert.ToString(Request("Juzhu"))
        rqJuzhu = TIMS.ClearSQM(rqJuzhu)

        Select Case rqJuzhu
            Case cst_Juzhu
                If rqProecess <> "editmy" Then
                    'Common.MessageBox(Page, "局屬資訊異常1，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "署屬資訊異常1，請重新操作查詢畫面。")
                    Exit Sub
                End If
                If OCID.Value = "" Then
                    'Common.MessageBox(Page, "局屬資訊遺失2，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "署屬資訊遺失2，請重新操作查詢畫面。")
                    Exit Sub
                End If
                'If rqOCID <> OCID.Value Then
                '    Common.MessageBox(Page, "局屬資訊遺失3，請重新操作查詢畫面。")
                '    Exit Sub
                'End If

            Case cst_NonJuzhu
                If rqProecess <> "editother" Then
                    'Common.MessageBox(Page, "非局屬資訊異常1，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "非署屬資訊異常1，請重新操作查詢畫面。")
                    Exit Sub
                End If
                If rqDLID = "" Then
                    'Common.MessageBox(Page, "非局屬資訊異常2，請重新操作查詢畫面。")
                    Common.MessageBox(Page, "非署屬資訊異常2，請重新操作查詢畫面。")
                    Exit Sub
                End If

            Case Else
                'Common.MessageBox(Page, "局屬/非局屬資訊遺失4，請重新操作查詢畫面。")
                Common.MessageBox(Page, "署屬/非署屬資訊遺失4，請重新操作查詢畫面。")
                Exit Sub
        End Select

        '新增/修改1筆封面 '署(局)屬/非署(局)屬
        Call SaveDataNew1(Val(rqDLID), rqJuzhu)
        Call UpdateStud1(Val(rqDLID), rqJuzhu)

        Page.RegisterStartupScript("winclose", "<script>alert('儲存成功');location.href='SD_05_008.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '修正異常學員資料
    Sub UpdateStud1(ByVal rqDLID As String, ByVal rqJuzhu As String)
        If rqJuzhu <> cst_Juzhu Then Exit Sub '署(局)屬離開

        Dim sql As String = "SELECT * FROM STUD_DATALID WHERE DLID=@DLID and ocid is not null"
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT a.dlid,a.subno ,a.stdpid ,a.stdname" & vbCrLf
        sql &= " ,b.ocid ,ss.socid" & vbCrLf
        sql &= " ,a.socid  asocid" & vbCrLf
        sql &= " FROM STUD_RESULTSTUDDATA a" & vbCrLf
        sql &= " JOIN STUD_DATALID b on b.dlid=a.dlid" & vbCrLf
        sql &= " JOIN V_STUDENTINFO ss on ss.ocid=b.ocid and ss.idno=a.stdpid" & vbCrLf
        sql &= " WHERE a.socid is null AND a.dlid=@DLID" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " UPDATE STUD_RESULTSTUDDATA"
        sql &= " SET SOCID=@SOCID"
        sql &= " WHERE DLID=@DLID AND SUBNO=@SUBNO"
        Dim uCmd2 As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("DLID", SqlDbType.VarChar).Value = rqDLID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Exit Sub '沒資料離開

        Dim dt2 As New DataTable
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("DLID", SqlDbType.VarChar).Value = rqDLID
            dt2.Load(.ExecuteReader())
        End With
        If dt2.Rows.Count = 0 Then Exit Sub '沒資料離開

        '修正
        For Each dr As DataRow In dt2.Rows
            With uCmd2
                .Parameters.Clear()
                .Parameters.Add("socid", SqlDbType.VarChar).Value = dr("socid")
                .Parameters.Add("DLID", SqlDbType.VarChar).Value = dr("DLID")
                .Parameters.Add("SUBNO", SqlDbType.VarChar).Value = dr("SUBNO")
                .ExecuteNonQuery()
            End With
        Next

    End Sub

    Private Sub BtnBack1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnBack1.ServerClick
        If Not Session(cst_search) Is Nothing Then Session(cst_search) = Session(cst_search)
        TIMS.Utl_Redirect1(Me, "SD_05_008.aspx?ID=" & Request("ID"))
    End Sub

    Sub SUB_IMPORT2()

        '姓名：	STDNAME　	　	　	　
        '學號：	(1-99)　STUDENTID	　	　	　
        '一、身分證號碼：	　(要符合身分證邏輯) STDPID	　	　	　
        '二、性別：	　(1m:男,2f: 女)	SEX　	　	 　
        '三、出生日期：(西元年月日)yyyy/MM/dd BirthYear	/BirthMonth	/BirthDate	/aBIRTHDAY
        '四、學歷：1:國中(含)以下,2:高中/職,3:專科,4:大學,5:碩士,6:博士 DegreeID
        '五、兵役：1:已役,2:未役,3:免役,4:在役中 MilitaryID (00為不選擇)
        '六、學員身分(可複選，最多五項) 複選1：
        '六、學員身分(可複選，最多五項) 複選2：
        '六、學員身分(可複選，最多五項) 複選3：
        '六、學員身分(可複選，最多五項) 複選4：
        '六、學員身分(可複選，最多五項) 複選5：

        Dim ff As String = ""
        Dim dtDegree As DataTable
        Dim dtMILITARY As DataTable
        Dim dtIDENTITY As DataTable
        Dim sql As String = ""
        Call TIMS.OpenDbConn(objconn)

        sql = " SELECT DEGREEID FROM KEY_DEGREE WHERE DEGREETYPE='1'" & vbCrLf
        dtDegree = DbAccess.GetDataTable(sql, objconn)

        'sql = " SELECT MILITARYID FROM KEY_MILITARY WHERE MILITARYID !='00'" & vbCrLf
        sql = " SELECT MILITARYID FROM KEY_MILITARY" & vbCrLf '00 為不選擇。
        dtMILITARY = DbAccess.GetDataTable(sql, objconn)

        sql = " SELECT IDENTITYID FROM KEY_IDENTITY" & vbCrLf
        dtIDENTITY = DbAccess.GetDataTable(sql, objconn)

        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM STUD_RESULTSTUDDATA" & vbCrLf
        sql &= " WHERE DLID=@DLID AND StudentID=@StudentID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        ' ,SOCID,@SOCID
        sql = "" & vbCrLf
        sql &= " INSERT INTO STUD_RESULTSTUDDATA (DLID ,SUBNO" & vbCrLf '/*PK*/ 
        sql &= " ,STDNAME ,STUDENTID ,STDPID  ,SEX" & vbCrLf
        sql &= " ,BIRTHYEAR ,BIRTHMONTH ,BIRTHDATE" & vbCrLf
        sql &= " ,DEGREEID,MILITARYID,MODIFYACCT,MODIFYDATE)" & vbCrLf
        sql &= " VALUES (@DLID ,@SUBNO" & vbCrLf '/*PK*/ 
        sql &= " ,@STDNAME,@STUDENTID,@STDPID,@SEX" & vbCrLf
        sql &= " ,@BIRTHYEAR,@BIRTHMONTH,@BIRTHDATE" & vbCrLf
        sql &= " ,@DEGREEID,@MILITARYID,@MODIFYACCT,GETDATE())" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        Const Cst_Filetype As String = "csv" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        'Dim MyPostedFile As HttpPostedFile = Nothing
        'If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return
        Const cst_flag As String = ","
        Dim sMsgBox As String = "" '每筆錯誤收集。
        Dim errMsgBox As String = "" '總錯誤收集。
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        If File1.Value = "" Then
            Common.MessageBox(Me, "請選擇匯入檔案的路徑!")
            Exit Sub
        End If

        '檢查檔案格式與大小----------   Start
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> Cst_Filetype Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小----------   End

        Const Cst_FileSavePath As String = "~/SD/03/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Cst_FileSavePath}{MyFileName}")
        '上傳檔案
        File1.PostedFile.SaveAs(filePath1)

        Try
            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(filePath1)
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0
            Dim OneRow As String
            'Dim col As String                       '欄位
            Dim colArray As Array

            '取出資料庫的所有欄位--------   Start
            Do While srr.Peek >= 0
                sMsgBox = ""

                OneRow = srr.ReadLine
                If Replace(",", OneRow, "") = "" Then
                    Exit Do
                End If

                If RowIndex <> 0 Then
                    colArray = Split(OneRow, cst_flag)

                    aSTDNAME = ""
                    aSTUDENTID = ""
                    aSTDPID = ""
                    aSEX = ""
                    aBIRTHDAY = ""
                    aDegreeID = ""
                    aMilitaryID = ""
                    aIDENTITYIDx1 = ""
                    aIDENTITYIDx2 = ""
                    aIDENTITYIDx3 = ""
                    aIDENTITYIDx4 = ""
                    aIDENTITYIDx5 = ""

                    Try
                        If colArray.Length > 0 Then aSTDNAME = Convert.ToString(colArray(0)) '姓名
                        If colArray.Length > 1 Then aSTUDENTID = Convert.ToString(colArray(1)) '學號
                        If colArray.Length > 2 Then aSTDPID = Convert.ToString(colArray(2)) '身分證號碼
                        If colArray.Length > 3 Then aSEX = Convert.ToString(colArray(3)) '性別
                        If colArray.Length > 4 Then aBIRTHDAY = Convert.ToString(colArray(4)) '出生日期：(西元年月日)yyyy/MM/dd	　	　	
                        If colArray.Length > 5 Then aDegreeID = Convert.ToString(colArray(5)) '學歷：1:國中(含)以下,2:高中/職,3:專科,4:大學,5:碩士,6:博士
                        If colArray.Length > 6 Then aMilitaryID = Convert.ToString(colArray(6)) '兵役：0:不選擇 1:已役,2:未役,3:免役,4:在役中
                        If colArray.Length > 7 Then aIDENTITYIDx1 = Convert.ToString(colArray(7))
                        If colArray.Length > 8 Then aIDENTITYIDx2 = Convert.ToString(colArray(8))
                        If colArray.Length > 9 Then aIDENTITYIDx3 = Convert.ToString(colArray(9))
                        If colArray.Length > 10 Then aIDENTITYIDx4 = Convert.ToString(colArray(10))
                        If colArray.Length > 11 Then aIDENTITYIDx5 = Convert.ToString(colArray(11))

                    Catch ex As Exception
                        sMsgBox += ex.ToString
                    End Try

                    If aSTDNAME <> "" Then aSTDNAME = Trim(aSTDNAME) 'If aSTUDENTID <> "" Then aSTUDENTID = Trim(aSTUDENTID)
                    If aSTDNAME = "" Then
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫姓名!" & vbCrLf
                    End If
                    If aSTUDENTID <> "" Then aSTUDENTID = Trim(aSTUDENTID)
                    If aSTUDENTID <> "" Then
                        If Not TIMS.IsNumeric2(aSTUDENTID) Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學號：" & aSTUDENTID & "檢查有誤!" & vbCrLf
                        End If
                    Else
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫學號!" & vbCrLf
                    End If

                    aSTDPID = TIMS.ClearSQM(aSTDPID)
                    aSTDPID = TIMS.ChangeIDNO(aSTDPID)
                    If aSTDPID <> "" Then
                        '1:國民身分證 
                        Dim flag1 As Boolean = TIMS.CheckIDNO(aSTDPID)
                        '2:居留證 4:居留證2021
                        Dim flag2 As Boolean = TIMS.CheckIDNO2(aSTDPID, 2)
                        Dim flag4 As Boolean = TIMS.CheckIDNO2(aSTDPID, 4)

                        If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,身分證號碼(或居留證號)：" & aSTDPID & "檢查有誤!" & vbCrLf
                        End If
                    Else
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫身分證號碼!" & vbCrLf
                    End If

                    If aSEX <> "" Then
                        Select Case UCase(aSEX)
                            Case "1", "2"
                            Case "M"
                                aSEX = "1"
                            Case "F"
                                aSEX = "2"
                            Case Else
                                sMsgBox += "第 " & RowIndex & " 筆資料,性別：" & aSEX & "檢查有誤!" & vbCrLf
                        End Select
                    Else
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫性別!" & vbCrLf
                    End If

                    If aBIRTHDAY <> "" Then
                        If Not TIMS.IsDate1(aBIRTHDAY) Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,出生日期：" & aBIRTHDAY & "檢查有誤!" & vbCrLf
                        Else
                            aBIRTHDAY = CDate(aBIRTHDAY).ToString("yyyy/MM/dd")
                        End If
                    Else
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫出生日期!" & vbCrLf
                    End If

                    If aDegreeID <> "" Then
                        If aDegreeID.Length = 1 Then
                            aDegreeID = "0" & aDegreeID
                        End If
                        ff = "DegreeID ='" & aDegreeID & "'"
                        If dtDegree.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學歷：" & aDegreeID & "檢查有誤!" & vbCrLf
                        End If
                    Else
                        sMsgBox += "第 " & RowIndex & " 筆資料,未填寫學歷!" & vbCrLf
                    End If

                    aMilitaryID = TIMS.ClearSQM(aMilitaryID)
                    If aMilitaryID <> "" Then
                        If aMilitaryID.Length = 1 Then
                            aMilitaryID = "0" & aMilitaryID
                        End If
                        ff = "MILITARYID ='" & aMilitaryID & "'"
                        If dtMILITARY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,兵役：" & aDegreeID & "檢查有誤!" & vbCrLf
                        End If
                    Else
                        aMilitaryID = "00"
                        'sMsgBox += "第 " & RowIndex & " 筆資料,未填寫兵役!" & vbCrLf
                    End If

                    'If colArray.Length > 7 Then aIDENTITYIDx1 = Convert.ToString(colArray(7))
                    'If colArray.Length > 8 Then aIDENTITYIDx2 = Convert.ToString(colArray(8))
                    'If colArray.Length > 9 Then aIDENTITYIDx3 = Convert.ToString(colArray(9))
                    'If colArray.Length > 10 Then aIDENTITYIDx4 = Convert.ToString(colArray(10))
                    'If colArray.Length > 11 Then aIDENTITYIDx5 = Convert.ToString(colArray(11))

                    If aIDENTITYIDx1 <> "" Then
                        If aIDENTITYIDx1.Length = 1 Then
                            aIDENTITYIDx1 = "0" & aIDENTITYIDx1
                        End If
                        ff = "IDENTITYID='" & aIDENTITYIDx1 & "'"
                        If dtIDENTITY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學員身分別1：" & aIDENTITYIDx1 & "檢查有誤!" & vbCrLf
                        End If
                    End If

                    If aIDENTITYIDx2 <> "" Then
                        If aIDENTITYIDx2.Length = 1 Then
                            aIDENTITYIDx2 = "0" & aIDENTITYIDx2
                        End If
                        ff = "IDENTITYID='" & aIDENTITYIDx2 & "'"
                        If dtIDENTITY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學員身分別2：" & aIDENTITYIDx2 & "檢查有誤!" & vbCrLf
                        End If
                    End If

                    If aIDENTITYIDx3 <> "" Then
                        If aIDENTITYIDx3.Length = 1 Then
                            aIDENTITYIDx3 = "0" & aIDENTITYIDx3
                        End If
                        ff = "IDENTITYID='" & aIDENTITYIDx3 & "'"
                        If dtIDENTITY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學員身分別3：" & aIDENTITYIDx3 & "檢查有誤!" & vbCrLf
                        End If
                    End If

                    If aIDENTITYIDx4 <> "" Then
                        If aIDENTITYIDx4.Length = 1 Then
                            aIDENTITYIDx4 = "0" & aIDENTITYIDx4
                        End If
                        ff = "IDENTITYID='" & aIDENTITYIDx4 & "'"
                        If dtIDENTITY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學員身分別4：" & aIDENTITYIDx4 & "檢查有誤!" & vbCrLf
                        End If
                    End If

                    If aIDENTITYIDx5 <> "" Then
                        If aIDENTITYIDx5.Length = 1 Then
                            aIDENTITYIDx5 = "0" & aIDENTITYIDx5
                        End If
                        ff = "IDENTITYID='" & aIDENTITYIDx1 & "'"
                        If dtIDENTITY.Select(ff).Length = 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學員身分別5：" & aIDENTITYIDx5 & "檢查有誤!" & vbCrLf
                        End If
                    End If

                    If sMsgBox = "" Then
                        '沒有錯誤，試著查詢
                        Dim dt1 As New DataTable
                        With sCmd
                            .Parameters.Clear()
                            .Parameters.Add("DLID", SqlDbType.VarChar).Value = HidDLID.Value
                            .Parameters.Add("STUDENTID", SqlDbType.VarChar).Value = aSTUDENTID
                            dt1.Load(.ExecuteReader())
                        End With
                        If dt1.Rows.Count > 0 Then
                            sMsgBox += "第 " & RowIndex & " 筆資料,學號：" & aSTUDENTID & " 資料已經在存，不可再新增" & vbCrLf
                        End If
                    End If

                    If sMsgBox = "" Then
                        '沒有錯誤，試著新增
                        Dim iSUBNO As Integer = TIMS.Get_nSubNOxResultStudData(HidDLID.Value, objconn)
                        With iCmd 'STUD_RESULTSTUDDATA
                            .Parameters.Clear()
                            .Parameters.Add("DLID", SqlDbType.VarChar).Value = HidDLID.Value
                            .Parameters.Add("SUBNO", SqlDbType.VarChar).Value = iSUBNO

                            .Parameters.Add("STDNAME", SqlDbType.VarChar).Value = aSTDNAME
                            .Parameters.Add("STUDENTID", SqlDbType.VarChar).Value = aSTUDENTID
                            .Parameters.Add("STDPID", SqlDbType.VarChar).Value = aSTDPID
                            .Parameters.Add("SEX", SqlDbType.VarChar).Value = aSEX
                            'vDate1 = CDate(vDate1).ToString("yyyy/MM/dd")
                            'Byear.Text = Year(CDate(vDate1))
                            'Bmonth.Text = Month(CDate(vDate1))
                            'Bday.Text = Day(CDate(vDate1))
                            .Parameters.Add("BIRTHYEAR", SqlDbType.VarChar).Value = Year(CDate(aBIRTHDAY))
                            .Parameters.Add("BIRTHMONTH", SqlDbType.VarChar).Value = Month(CDate(aBIRTHDAY))
                            .Parameters.Add("BIRTHDATE", SqlDbType.VarChar).Value = Day(CDate(aBIRTHDAY))
                            .Parameters.Add("DEGREEID", SqlDbType.VarChar).Value = aDegreeID
                            .Parameters.Add("MILITARYID", SqlDbType.VarChar).Value = IIf(aMilitaryID = "", Convert.DBNull, aMilitaryID)
                            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .ExecuteNonQuery()
                        End With

                        If aIDENTITYIDx1 <> "" Then Call TIMS.INTO_ResultIdentData(HidDLID.Value, iSUBNO, aIDENTITYIDx1, objconn)
                        If aIDENTITYIDx2 <> "" Then Call TIMS.INTO_ResultIdentData(HidDLID.Value, iSUBNO, aIDENTITYIDx2, objconn)
                        If aIDENTITYIDx3 <> "" Then Call TIMS.INTO_ResultIdentData(HidDLID.Value, iSUBNO, aIDENTITYIDx3, objconn)
                        If aIDENTITYIDx4 <> "" Then Call TIMS.INTO_ResultIdentData(HidDLID.Value, iSUBNO, aIDENTITYIDx4, objconn)
                        If aIDENTITYIDx5 <> "" Then Call TIMS.INTO_ResultIdentData(HidDLID.Value, iSUBNO, aIDENTITYIDx5, objconn)

                    End If

                    If sMsgBox <> "" Then
                        errMsgBox &= sMsgBox
                    End If
                End If
                RowIndex = RowIndex + 1
            Loop
            sr.Close()
            srr.Close()
            'TIMS.MyFileDelete(Server.MapPath(Cst_FileSavePath & MyFileName))
        Catch ex As Exception
            errMsgBox &= ex.ToString
        End Try

        TIMS.MyFileDelete(filePath1)

        If errMsgBox <> "" Then
            errMsgBox = "有些資料匯入成功，但有錯誤的資料無法匯入，請檢查下列資料:" & vbCrLf & errMsgBox
            Common.MessageBox(Me, errMsgBox)
        Else
            Common.MessageBox(Me, "資料匯入成功，請按查詢，查看匯入資料")
        End If
    End Sub
    '匯入
    Protected Sub BtnImport1_Click(sender As Object, e As EventArgs) Handles btnImport1.Click
        Dim Errmsg As String = ""
        If TIMS.ClearSQM(HidDLID.Value) = "" Then
            'Errmsg += "非局屬班級查詢 有誤 請重新選擇 " & vbCrLf
            Errmsg += "非署屬班級查詢 有誤 請重新選擇 " & vbCrLf
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SUB_IMPORT2()
    End Sub

End Class
