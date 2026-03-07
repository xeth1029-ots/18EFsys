Public Class SD_05_032
    Inherits AuthBasePage

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    'ADP_RESOLDER
    '../../Doc/ADP_RESOLDERv21.zip
    'alter TABLE [dbo].[ADP_RESOLDER] 	add POSITION [nvarchar] (100)  NULL 
    Const cst_ia姓名 As Integer = 0
    Const cst_ia身分證字號 As Integer = 1 'IDNO
    Const cst_ia出生年月日 As Integer = 2 'BIRTHDAY'[datetime]
    Const cst_ia任職單位全銜 As Integer = 3 'POSITION'[nvarchar] (100)
    Const cst_ia預定退伍日 As Integer = 4 'PREEXDATE'[datetime]
    Const cst_ia送訓至分署 As Integer = 5
    Const cst_iaMaxLength1 As Integer = 6
    'Const cst_excel_exp As String="送訓官兵名冊"

    Dim dtDIST As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '在這裡放置使用者程式碼以初始化網頁
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        HyperLink1.NavigateUrl = "../../Doc/ADP_RESOLDERv21.zip"

        Dim sql As String
        sql = "SELECT DISTID,NAME FROM ID_DISTRICT WHERE DISTID!='000' ORDER BY DISTID"
        dtDIST = DbAccess.GetDataTable(sql, objconn)

        If Not Page.IsPostBack Then
            labmsg.Text = ""
            Call sUtl_Cancel1()
            tbSch.Visible = True
            Call initObj()
        End If
    End Sub

    '功能第一次載入初始化
    Sub initObj()
        sRECOMMDISTID = TIMS.Get_DistID(sRECOMMDISTID, dtDIST)
        sRECOMMDISTID.Enabled = False
        Common.SetListItem(sRECOMMDISTID, sm.UserInfo.DistID)
        ddlRECOMMDISTID = TIMS.Get_DistID(ddlRECOMMDISTID, dtDIST)
        ddlRECOMMDISTID.Enabled = False
        Common.SetListItem(ddlRECOMMDISTID, sm.UserInfo.DistID)

        If sm.UserInfo.DistID = "000" Then
            sRECOMMDISTID.Enabled = True
            ddlRECOMMDISTID.Enabled = True
        End If

        'Call ListClass.crtDropDownList("org_master", ddlQoyc_status)
        'ddlQTabNum=Get_ddlTabNum(ddlQTabNum)
        'TabNum=Get_ddlTabNum(TabNum)
        'For i As Integer=0 To 23
        '    HR1.Items.Add(New ListItem(i, i))
        '    HR2.Items.Add(New ListItem(i, i))
        'Next
        'Common.SetListItem(HR2, 23)

        'For j As Integer=0 To 59
        '    MM1.Items.Add(New ListItem(j, j))
        '    MM2.Items.Add(New ListItem(j, j))
        'Next
        'Common.SetListItem(MM2, 59)
    End Sub

    '取消
    Sub sUtl_Cancel1()
        tbSch.Visible = False
        tbList.Visible = False
        tbEdit.Visible = False
    End Sub

    '清除值(及狀態設定)
    Sub clsValue()
        Hid_ARSID.Value = ""
        tCNAME.Text = ""
        tIDNO.Text = ""
        tBIRTHDAY.Text = ""
        tPOSITION.Text = ""
        tPREEXDATE.Text = ""
        ddlRECOMMDISTID.SelectedIndex = -1
    End Sub

    '將搜尋值加入編輯資料
    Sub CopySch2Value()
        tCNAME.Text = sCNAME.Text
        tIDNO.Text = sIDNO.Text
        'tBIRTHDAY.Text=sBIRTHDAY1.Text
        'tPOSITION.Text=""
        'tPREEXDATE.Text=sPREEXDATE1.Text
        Dim v_sRECOMMDISTID As String = TIMS.GetListValue(sRECOMMDISTID)
        If v_sRECOMMDISTID <> "" Then Common.SetListItem(ddlRECOMMDISTID, v_sRECOMMDISTID)
    End Sub

    '記錄查詢條件 
    Sub Search1Value()
        sCNAME.Text = TIMS.ClearSQM(sCNAME.Text)
        sIDNO.Text = TIMS.ChangeIDNO(sIDNO.Text)
        sBIRTHDAY1.Text = TIMS.ClearSQM(sBIRTHDAY1.Text)
        sBIRTHDAY2.Text = TIMS.ClearSQM(sBIRTHDAY2.Text)
        sPREEXDATE1.Text = TIMS.ClearSQM(sPREEXDATE1.Text)
        sPREEXDATE2.Text = TIMS.ClearSQM(sPREEXDATE2.Text)
        sCREATEDATE1.Text = TIMS.ClearSQM(sCREATEDATE1.Text)
        sCREATEDATE2.Text = TIMS.ClearSQM(sCREATEDATE2.Text)

        sBIRTHDAY1.Text = TIMS.Cdate3(sBIRTHDAY1.Text)
        sBIRTHDAY2.Text = TIMS.Cdate3(sBIRTHDAY2.Text)
        sPREEXDATE1.Text = TIMS.Cdate3(sPREEXDATE1.Text)
        sPREEXDATE2.Text = TIMS.Cdate3(sPREEXDATE2.Text)
        sCREATEDATE1.Text = TIMS.Cdate3(sCREATEDATE1.Text)
        sCREATEDATE2.Text = TIMS.Cdate3(sCREATEDATE2.Text)

        'Dim v_sRECOMMDISTID As String=TIMS.GetListValue(sRECOMMDISTID)
        ViewState("sCNAME") = sCNAME.Text
        ViewState("sIDNO") = sIDNO.Text
        ViewState("sBIRTHDAY1") = sBIRTHDAY1.Text
        ViewState("sBIRTHDAY2") = sBIRTHDAY2.Text
        ViewState("sPREEXDATE1") = sPREEXDATE1.Text
        ViewState("sPREEXDATE2") = sPREEXDATE2.Text
        If sm.UserInfo.DistID = "000" Then ViewState("sRECOMMDISTID") = TIMS.GetListValue(sRECOMMDISTID) 'sRECOMMDISTID.SelectedValue
        If sm.UserInfo.DistID <> "000" Then ViewState("sRECOMMDISTID") = sm.UserInfo.DistID
        ViewState("sCREATEDATE1") = sCREATEDATE1.Text
        ViewState("sCREATEDATE2") = sCREATEDATE2.Text
    End Sub

    '查詢
    Sub Search1()
        'ADP_RESOLDER
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim dt As DataTable = Nothing
        Call sUtl_Search1(dt)
        'dt=DbAccess.GetDataTable(sql, objconn, parms)

        labmsg.Text = "查無資料"
        tbList.Visible = False
        If dt.Rows.Count > 0 Then
            'CPdt=dt.Copy()
            labmsg.Text = ""
            tbList.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '取得單筆資訊(編輯)
    Sub LoadData1()
        If Hid_ARSID.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql &= " SELECT a.ARSID " & vbCrLf   '/*PK*/ 
        sql &= " ,a.CNAME " & vbCrLf
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,a.BIRTHDAY " & vbCrLf
        sql &= " ,a.POSITION " & vbCrLf
        sql &= " ,a.PREEXDATE " & vbCrLf
        sql &= " ,a.RECOMMDISTID " & vbCrLf
        sql &= " ,a.CREATEACCT " & vbCrLf
        sql &= " ,a.CREATEDISTID " & vbCrLf
        sql &= " ,a.CREATEDATE " & vbCrLf
        sql &= " ,a.MODIFYACCT " & vbCrLf
        sql &= " ,a.MODIFYDISTID " & vbCrLf
        sql &= " ,a.MODIFYDATE " & vbCrLf
        'sql += " ,b.NAME RCDISTNAME " & vbCrLf
        sql &= " FROM ADP_RESOLDER a " & vbCrLf
        'sql += " JOIN ID_DISTRICT b on b.DistID=a.RECOMMDISTID " & vbCrLf
        sql &= " WHERE a.ARSID=@ARSID " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then sql &= " AND a.RECOMMDISTID=@RECOMMDISTID " & vbCrLf

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("ARSID", SqlDbType.Int).Value = Val(Hid_ARSID.Value)
            If sm.UserInfo.DistID <> "000" Then
                'sql += " AND a.RECOMMDISTID=@RECOMMDISTID " & vbCrLf
                .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
            End If
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then Return

        Dim dr As DataRow = dt.Rows(0)
        Hid_ARSID.Value = Convert.ToString(dr("ARSID"))
        tCNAME.Text = Convert.ToString(dr("CNAME"))
        tIDNO.Text = Convert.ToString(dr("IDNO"))
        tBIRTHDAY.Text = TIMS.Cdate3(dr("BIRTHDAY"))
        tPOSITION.Text = Convert.ToString(dr("POSITION"))
        tPREEXDATE.Text = TIMS.Cdate3(dr("PREEXDATE"))
        If Convert.ToString(dr("RECOMMDISTID")) <> "" Then Common.SetListItem(ddlRECOMMDISTID, dr("RECOMMDISTID"))

    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        'Dim v_sRECOMMDISTID As String=TIMS.GetListValue(sRECOMMDISTID)
        Dim v_ddlRECOMMDISTID As String = TIMS.GetListValue(ddlRECOMMDISTID)
        tCNAME.Text = TIMS.ClearSQM(tCNAME.Text)
        tIDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(tIDNO.Text))
        tBIRTHDAY.Text = TIMS.ClearSQM(tBIRTHDAY.Text)
        tPOSITION.Text = TIMS.ClearSQM(tPOSITION.Text)
        tPOSITION.Text = TIMS.Get_Substr1(tPOSITION.Text, 100)
        tPREEXDATE.Text = TIMS.ClearSQM(tPREEXDATE.Text)

        If tCNAME.Text = "" Then Errmsg += "請輸入 姓名" & vbCrLf
        If tIDNO.Text = "" Then Errmsg += "請輸入 身分證字號" & vbCrLf
        If tBIRTHDAY.Text = "" Then Errmsg += "請輸入 出生年月日" & vbCrLf
        If tPOSITION.Text = "" Then Errmsg += "請輸入 任職單位(全銜)" & vbCrLf
        If tPREEXDATE.Text = "" Then Errmsg += "請輸入 預定退伍日" & vbCrLf
        If v_ddlRECOMMDISTID = "" Then Errmsg += "請選擇 送訓至分署" & vbCrLf
        If tIDNO.Text <> "" AndAlso Not TIMS.CheckIDNO(tIDNO.Text) Then Errmsg += "身分證字號 檢核有誤" & vbCrLf
        If tBIRTHDAY.Text <> "" AndAlso Not TIMS.IsDate1(tBIRTHDAY.Text) Then Errmsg += "出生年月日 請輸入正確日期格式" & vbCrLf
        If tPREEXDATE.Text <> "" AndAlso Not TIMS.IsDate1(tPREEXDATE.Text) Then Errmsg += "預定退伍日 請輸入正確日期格式" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Sub SaveData1()
        Dim rst As Integer = 0

        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)

        Dim sql As String = ""
        sql &= " INSERT INTO ADP_RESOLDER (ARSID ,CNAME ,IDNO ,BIRTHDAY,POSITION ,PREEXDATE ,RECOMMDISTID" & vbCrLf
        sql &= "  ,CREATEACCT ,CREATEDISTID ,CREATEDATE ,MODIFYACCT ,MODIFYDISTID ,MODIFYDATE) " & vbCrLf
        sql &= " VALUES (@ARSID ,@CNAME ,@IDNO ,@BIRTHDAY,@POSITION ,@PREEXDATE ,@RECOMMDISTID" & vbCrLf
        sql &= "  ,@CREATEACCT ,@CREATEDISTID ,GETDATE() ,@MODIFYACCT ,@MODIFYDISTID ,GETDATE() ) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = ""
        sql &= " UPDATE ADP_RESOLDER " & vbCrLf
        sql &= " SET CNAME=@CNAME ,IDNO=@IDNO " & vbCrLf
        sql &= " ,BIRTHDAY=@BIRTHDAY " & vbCrLf
        sql &= " ,POSITION=@POSITION " & vbCrLf
        sql &= " ,PREEXDATE=@PREEXDATE " & vbCrLf
        sql &= " ,RECOMMDISTID=@RECOMMDISTID " & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT " & vbCrLf
        sql &= " ,MODIFYDISTID=@MODIFYDISTID " & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE() " & vbCrLf
        sql &= " WHERE ARSID=@ARSID " & vbCrLf  '/*PK*/ 
        Dim uCmd As New SqlCommand(sql, objconn)

        '新增重複判斷
        sql = ""
        sql &= " SELECT 'X' FROM ADP_RESOLDER " & vbCrLf
        sql &= " WHERE IDNO=@IDNO AND BIRTHDAY=@BIRTHDAY " & vbCrLf
        sql &= " AND PREEXDATE=@PREEXDATE " & vbCrLf
        sql &= " AND RECOMMDISTID=@RECOMMDISTID " & vbCrLf
        Dim siCmd As New SqlCommand(sql, objconn)

        'sql &= " AND UPPER(CNAME)=@CNAME " & vbCrLf
        '修改重複判斷
        sql = ""
        sql &= " SELECT 'X' FROM ADP_RESOLDER " & vbCrLf
        sql &= " WHERE IDNO=@IDNO  AND BIRTHDAY=@BIRTHDAY " & vbCrLf
        sql &= " AND PREEXDATE=@PREEXDATE " & vbCrLf
        sql &= " AND RECOMMDISTID=@RECOMMDISTID " & vbCrLf
        sql &= " AND ARSID != @ARSID " & vbCrLf
        Dim suCmd As New SqlCommand(sql, objconn)
        Dim v_ddlRECOMMDISTID As String = TIMS.GetListValue(ddlRECOMMDISTID)
        If Hid_ARSID.Value = "" Then
            '新增(檢核)
            Dim dt1 As New DataTable
            With siCmd
                .Parameters.Clear()
                '.Parameters.Add("CNAME", SqlDbType.NVarChar).Value=UCase(tCNAME.Text)
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
                .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = CDate(tBIRTHDAY.Text)
                .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = CDate(tPREEXDATE.Text)
                .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該屆退官兵荐訓資料已新增，請使用修改功能!!")
                Exit Sub
            End If
        Else
            '修改(檢核)
            Dim dt1 As New DataTable
            With suCmd
                .Parameters.Clear()
                '.Parameters.Add("CNAME", SqlDbType.NVarChar).Value=UCase(tCNAME.Text)
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
                .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = CDate(tBIRTHDAY.Text)
                .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = CDate(tPREEXDATE.Text)
                .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
                .Parameters.Add("ARSID", SqlDbType.Int).Value = Val(Hid_ARSID.Value) '/*PK*/
                dt1.Load(.ExecuteReader())
            End With
            If dt1.Rows.Count > 0 Then
                Common.MessageBox(Me, "該屆退官兵荐訓資料已存在，請重新輸入!!")
                Exit Sub
            End If
        End If

        Dim str_saveok_msg As String = ""
        If Hid_ARSID.Value = "" Then
            '新增
            Dim iARSID As Integer = DbAccess.GetNewId(objconn, " ADP_RESOLDER_ARSID_SEQ,ADP_RESOLDER,ARSID")
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("ARSID", SqlDbType.Int).Value = iARSID  '/*PK*/
                .Parameters.Add("CNAME", SqlDbType.NVarChar).Value = tCNAME.Text
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
                .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(tBIRTHDAY.Text)
                .Parameters.Add("POSITION", SqlDbType.NVarChar).Value = tPOSITION.Text
                .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(tPREEXDATE.Text)
                .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
                .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("CREATEDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("MODIFYDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                rst = .ExecuteNonQuery()
            End With
            str_saveok_msg = "新增完成!"
        Else
            '修改
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("CNAME", SqlDbType.NVarChar).Value = tCNAME.Text
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
                .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(tBIRTHDAY.Text)
                .Parameters.Add("POSITION", SqlDbType.NVarChar).Value = tPOSITION.Text
                .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(tPREEXDATE.Text)
                .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
                '.Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value=sm.UserInfo.UserID
                '.Parameters.Add("CREATEDISTID", SqlDbType.VarChar).Value=sm.UserInfo.DistID
                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("MODIFYDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                .Parameters.Add("ARSID", SqlDbType.Int).Value = Val(Hid_ARSID.Value) '/*PK*/
                rst = .ExecuteNonQuery()
            End With
            str_saveok_msg = "修改完成!"
        End If

        If rst = 0 Then
            Common.MessageBox(Page, "執行完畢，無資料更動!")
            Return
        End If

        Call sUtl_Cancel1()
        tbSch.Visible = True

        Call Search1()
        If str_saveok_msg <> "" Then Common.MessageBox(Page, str_saveok_msg)
    End Sub

    '匯入資料塞入編輯欄位(匯入資料)
    Sub EditImportData(ByVal colArray As Array)
        If colArray.Length < cst_iaMaxLength1 Then Exit Sub
        'Dim Errmsg As String=""

        Dim idx As Integer = 0
        idx = cst_ia姓名
        If colArray.Length > idx Then tCNAME.Text = Convert.ToString(colArray(idx))
        idx = cst_ia身分證字號
        If colArray.Length > idx Then tIDNO.Text = Convert.ToString(colArray(idx))
        idx = cst_ia出生年月日
        If colArray.Length > idx Then tBIRTHDAY.Text = Convert.ToString(colArray(idx))
        idx = cst_ia任職單位全銜
        If colArray.Length > idx Then tPOSITION.Text = Convert.ToString(colArray(idx))
        idx = cst_ia預定退伍日
        If colArray.Length > idx Then tPREEXDATE.Text = Convert.ToString(colArray(idx))

        If sm.UserInfo.DistID = "000" Then
            idx = cst_ia送訓至分署
            If Convert.ToString(colArray(idx)) <> "" Then
                Dim ssRECOMMDISTID As String = TIMS.AddZero(colArray(idx), 3)
                Common.SetListItem(ddlRECOMMDISTID, ssRECOMMDISTID)
            End If
        End If
        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(ddlRECOMMDISTID, sm.UserInfo.DistID)
        End If

        tCNAME.Text = TIMS.ClearSQM(tCNAME.Text)
        tIDNO.Text = TIMS.ClearSQM(tIDNO.Text)
        tBIRTHDAY.Text = TIMS.ClearSQM(tBIRTHDAY.Text)
        tPOSITION.Text = TIMS.ClearSQM(tPOSITION.Text) '任職單位全銜
        tPREEXDATE.Text = TIMS.ClearSQM(tPREEXDATE.Text)

        tIDNO.Text = TIMS.ChangeIDNO(tIDNO.Text) '身分證號碼-ChangeIDNO
        tBIRTHDAY.Text = TIMS.Cdate3(tBIRTHDAY.Text)
        tPREEXDATE.Text = TIMS.Cdate3(tPREEXDATE.Text)
    End Sub

    '儲存(匯入資料)
    Function SaveDataImp(ByRef colArray As Array) As String
        Dim Errmsg As String = ""
        If colArray.Length < cst_iaMaxLength1 Then
            Errmsg += "欄位對應有誤" & vbCrLf
            Errmsg += "請檢查Excel格式是否正確" & vbCrLf
            Return Errmsg
        End If
        If Errmsg <> "" Then Return Errmsg

        Call CheckData1(Errmsg)
        If Errmsg <> "" Then Return Errmsg

        Dim rst As Integer = 0
        Call TIMS.OpenDbConn(objconn)
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)

        Dim sql As String = ""
        sql &= " INSERT INTO ADP_RESOLDER (ARSID ,CNAME ,IDNO ,BIRTHDAY,POSITION ,PREEXDATE ,RECOMMDISTID" & vbCrLf
        sql &= " ,CREATEACCT ,CREATEDISTID ,CREATEDATE ,MODIFYACCT ,MODIFYDISTID ,MODIFYDATE) " & vbCrLf
        sql &= " VALUES (@ARSID ,@CNAME ,@IDNO ,@BIRTHDAY,@POSITION ,@PREEXDATE ,@RECOMMDISTID" & vbCrLf
        sql &= " ,@CREATEACCT ,@CREATEDISTID ,GETDATE() ,@MODIFYACCT ,@MODIFYDISTID ,GETDATE() ) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        'sql &= " AND UPPER(CNAME)=@CNAME " & vbCrLf
        '新增重複判斷
        sql = ""
        sql &= " SELECT 'X' FROM ADP_RESOLDER " & vbCrLf
        sql &= " WHERE IDNO=@IDNO AND BIRTHDAY=@BIRTHDAY " & vbCrLf
        sql &= " AND PREEXDATE=@PREEXDATE " & vbCrLf
        sql &= " AND RECOMMDISTID=@RECOMMDISTID " & vbCrLf
        Dim siCmd As New SqlCommand(sql, objconn)

        Dim v_ddlRECOMMDISTID As String = TIMS.GetListValue(ddlRECOMMDISTID)

        '新增(檢核) '新增重複判斷
        Dim dt1 As New DataTable
        With siCmd
            .Parameters.Clear()
            '.Parameters.Add("CNAME", SqlDbType.NVarChar).Value=UCase(tCNAME.Text)
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
            .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = CDate(tBIRTHDAY.Text)
            .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = CDate(tPREEXDATE.Text)
            .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count > 0 Then
            Errmsg = "該屆退官兵荐訓資料已新增，請使用修改功能!!" & vbCrLf
            Return Errmsg
            'Common.MessageBox(Me, "該屆退官兵荐訓資料已新增，請使用修改功能!!")
            'Exit Function
        End If

        '新增
        Dim iARSID As Integer = DbAccess.GetNewId(objconn, "ADP_RESOLDER_ARSID_SEQ,ADP_RESOLDER,ARSID")
        With iCmd
            .Parameters.Clear()
            .Parameters.Add("ARSID", SqlDbType.Int).Value = iARSID  '/*PK*/
            .Parameters.Add("CNAME", SqlDbType.NVarChar).Value = tCNAME.Text
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = tIDNO.Text
            .Parameters.Add("BIRTHDAY", SqlDbType.DateTime).Value = TIMS.Cdate2(tBIRTHDAY.Text)
            .Parameters.Add("POSITION", SqlDbType.VarChar).Value = tPOSITION.Text
            .Parameters.Add("PREEXDATE", SqlDbType.DateTime).Value = TIMS.Cdate2(tPREEXDATE.Text)
            .Parameters.Add("RECOMMDISTID", SqlDbType.VarChar).Value = v_ddlRECOMMDISTID 'ddlRECOMMDISTID.SelectedValue
            .Parameters.Add("CREATEACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            .Parameters.Add("CREATEDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
            .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            .Parameters.Add("MODIFYDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
            rst = .ExecuteNonQuery()
        End With

        Return ""
    End Function

    '查詢鈕
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call Search1Value()  '記錄查詢條件
        Call Search1()
    End Sub

    '新增鈕
    Protected Sub btnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        Call clsValue()
        Call CopySch2Value()
        Call sUtl_Cancel1()
        tbEdit.Visible = True
    End Sub

    '儲存
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return ' Exit Sub
        End If

        Try
            Call SaveData1()
        Catch ex As Exception
            Errmsg = $"儲存有誤，{ex.Message}"
            Call TIMS.WriteTraceLog(ex.Message, ex)
        End Try
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return ' Exit Sub
        End If
    End Sub

    '回上頁
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        Call sUtl_Cancel1()
        tbSch.Visible = True
        If Hid_ARSID.Value <> "" Then Call Search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "" Then Exit Sub
        If e.CommandArgument = "" Then Exit Sub

        Select Case e.CommandName
            Case "UPD" '修改
                Call sUtl_Cancel1()
                tbEdit.Visible = True
                Call clsValue()
                Dim sCmdArg As String = Convert.ToString(e.CommandArgument)
                Hid_ARSID.Value = TIMS.GetMyValue(sCmdArg, "ARSID")
                Call LoadData1()
                'Case "DEL" '刪除
                '    Dim sCmdArg As String=Convert.ToString(e.CommandArgument)
                '    HidHN3ID.Value=TIMS.GetMyValue(sCmdArg, "HN3ID")
                '    Call Delete1()
        End Select
    End Sub

    '表格上的元件配置
    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim objDG1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 
                Dim lbtUpdate As LinkButton = e.Item.FindControl("lbtUpdate")
                'Dim lbtDelete As LinkButton=e.Item.FindControl("lbtDelete")
                'lbtDelete.Attributes.Add("onclick", "return confirm('您確定要刪除第" & e.Item.Cells(0).Text & "筆資料嗎?');")
                Dim sCmdArg As String = ""
                Call TIMS.SetMyValue(sCmdArg, "ARSID", drv("ARSID"))
                lbtUpdate.CommandArgument = sCmdArg
                'lbtDelete.CommandArgument=sCmdArg
        End Select
    End Sub

    '執行匯入動作。(匯入資料)
    Sub SUtl_ImprotX1(ByRef FullFileName1 As String)
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "FullFileName1", FullFileName1)
        TIMS.SetMyValue2(htSS, "FirstCol", "身分證字號") '"身分證字號" '任1欄位名稱(必填)
        Dim Reason As String = ""
        '上傳檔案/取得內容
        Dim dt_xls As DataTable = TIMS.Get_File1data(File1, Reason, htSS, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '建立錯誤資料格式Table----------------Start
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("IDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        Dim iRowIndex As Integer = 0 '讀取行累計數
        For Each dr As DataRow In dt_xls.Rows
            iRowIndex += 1
            Dim colArray As Array = dr.ItemArray
            Call clsValue()
            Call EditImportData(colArray)
            'SaveDataImport
            Dim sReason As String = ""
            sReason = SaveDataImp(colArray)
            If sReason <> "" Then
                '錯誤資料，填入錯誤資料表
                Dim drWrong As DataRow
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = iRowIndex
                drWrong("IDNO") = TIMS.ChangeIDNO(tIDNO.Text)
                drWrong("Reason") = Replace(sReason, vbCrLf, "<br>" & vbCrLf)
            End If
        Next

        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, "資料匯入完成。")
            Exit Sub
        End If

        Session("MyWrongTable") = Nothing
        If dtWrong.Rows.Count > 0 Then
            Session("MyWrongTable") = dtWrong
            Dim strScriptErrMsg1 As String = ""
            strScriptErrMsg1 &= "<script>"
            strScriptErrMsg1 &= "if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?'))"
            strScriptErrMsg1 &= "{window.open('SD_05_032_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}"
            strScriptErrMsg1 &= "</script>"
            Page.RegisterStartupScript(TIMS.xBlockName, strScriptErrMsg1)
        End If
    End Sub

    '匯入鈕。
    Protected Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/SD/05/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        'Call sImport2(FullFileName1)
        Call SUtl_ImprotX1(FullFileName1)
    End Sub

    ''' <summary>
    ''' 查詢dt
    ''' </summary>
    ''' <param name="dt"></param>
    Sub sUtl_Search1(ByRef dt As DataTable)
        'ADP_RESOLDER
        Dim sql As String = ""
        sql &= " SELECT a.ARSID " & vbCrLf   '/*PK*/ 
        sql &= " ,a.CNAME " & vbCrLf
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,CONVERT(varchar, a.BIRTHDAY, 111) BIRTHDAY " & vbCrLf
        sql &= " ,a.POSITION " & vbCrLf '任職單位全銜
        sql &= " ,CONVERT(varchar, a.PREEXDATE, 111) PREEXDATE " & vbCrLf
        sql &= " ,a.RECOMMDISTID " & vbCrLf
        sql &= " ,a.CREATEACCT " & vbCrLf
        sql &= " ,a.CREATEDISTID " & vbCrLf
        sql &= " ,a.CREATEDATE " & vbCrLf
        'CREATEDTE -匯入日期 
        sql &= " ,CONVERT(varchar, a.CREATEDATE, 111) CREATEDTE " & vbCrLf
        sql &= " ,a.MODIFYACCT " & vbCrLf
        sql &= " ,a.MODIFYDISTID " & vbCrLf
        sql &= " ,a.MODIFYDATE " & vbCrLf
        sql &= " ,b.NAME RCDISTNAME " & vbCrLf
        sql &= " FROM ADP_RESOLDER a " & vbCrLf
        sql &= " JOIN ID_DISTRICT b on b.DistID=a.RECOMMDISTID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If Convert.ToString(ViewState("sCNAME")) <> "" Then sql &= " AND a.CNAME like '%' + @sCNAME + '%' " & vbCrLf
        If Convert.ToString(ViewState("sIDNO")) <> "" Then sql &= " AND a.IDNO=@sIDNO " & vbCrLf
        'POSITION
        If Convert.ToString(ViewState("sBIRTHDAY1")) <> "" Then sql &= " AND a.BIRTHDAY >= @sBIRTHDAY1 " & vbCrLf
        If Convert.ToString(ViewState("sBIRTHDAY2")) <> "" Then sql &= " AND a.BIRTHDAY <= @sBIRTHDAY2 " & vbCrLf
        If Convert.ToString(ViewState("sPREEXDATE1")) <> "" Then sql &= " AND a.PREEXDATE >= @sPREEXDATE1 " & vbCrLf
        If Convert.ToString(ViewState("sPREEXDATE2")) <> "" Then sql &= " AND a.PREEXDATE <= @sPREEXDATE2 " & vbCrLf
        If Convert.ToString(ViewState("sRECOMMDISTID")) <> "" Then sql &= " AND a.RECOMMDISTID=@sRECOMMDISTID " & vbCrLf
        If Convert.ToString(ViewState("sCREATEDATE1")) <> "" Then sql &= " AND a.CREATEDATE >= @sCREATEDATE1" & vbCrLf
        If Convert.ToString(ViewState("sCREATEDATE2")) <> "" Then sql &= " AND a.CREATEDATE <= @sCREATEDATE2" & vbCrLf

        sql &= " ORDER BY a.RECOMMDISTID, a.CNAME, a.BIRTHDAY " & vbCrLf

        '(新版查詢功能,by:20180724)
        Call TIMS.OpenDbConn(objconn)
        Dim parms As Hashtable = New Hashtable()
        If Convert.ToString(ViewState("sCNAME")) <> "" Then parms.Add("sCNAME", ViewState("sCNAME"))
        If Convert.ToString(ViewState("sIDNO")) <> "" Then parms.Add("sIDNO", ViewState("sIDNO"))
        'POSITION
        If Convert.ToString(ViewState("sBIRTHDAY1")) <> "" Then parms.Add("sBIRTHDAY1", CDate(ViewState("sBIRTHDAY1")))
        If Convert.ToString(ViewState("sBIRTHDAY2")) <> "" Then parms.Add("sBIRTHDAY2", CDate(ViewState("sBIRTHDAY2")))
        If Convert.ToString(ViewState("sPREEXDATE1")) <> "" Then parms.Add("sPREEXDATE1", CDate(ViewState("sPREEXDATE1")))
        If Convert.ToString(ViewState("sPREEXDATE2")) <> "" Then parms.Add("sPREEXDATE2", CDate(ViewState("sPREEXDATE2")))
        If Convert.ToString(ViewState("sRECOMMDISTID")) <> "" Then parms.Add("sRECOMMDISTID", ViewState("sRECOMMDISTID"))
        If Convert.ToString(ViewState("sCREATEDATE1")) <> "" Then parms.Add("sCREATEDATE1", ViewState("sCREATEDATE1"))
        If Convert.ToString(ViewState("sCREATEDATE2")) <> "" Then parms.Add("sCREATEDATE2", ViewState("sCREATEDATE2"))
        'Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
    End Sub

    ''' <summary>
    ''' 匯出-查詢
    ''' </summary>
    ''' <param name="DGrid1"></param>
    ''' <returns></returns>
    Function Utl_SearchExp1(ByRef DGrid1 As DataGrid) As Boolean
        Dim dt2 As DataTable = Nothing
        Call sUtl_Search1(dt2)

        DGrid1.Visible = False
        labmsg.Text = "查無資料"
        tbList.Visible = False

        If dt2.Rows.Count = 0 Then Return False '沒有資料

        'CPdt=dt.Copy()
        DGrid1.Visible = True
        labmsg.Text = ""
        tbList.Visible = True

        DGrid1.DataSource = dt2
        DGrid1.DataBind()
        Return True '有資料
    End Function

    ''' <summary>
    ''' 匯出
    ''' </summary>
    Sub SUB_Export1(ByRef DGrid1 As DataGrid)
        'Dim cst_功能 As Integer=10
        DGrid1.AllowPaging = False
        'DataGrid1.Columns(cst_功能).Visible=False
        'DataGrid1.Columns(0).Visible=False '班別不顯示
        'If OCIDValue1.Value="" Then
        '    DataGrid1.Columns(0).Visible=True '班別顯示
        'End If
        DGrid1.EnableViewState = False  '把ViewState給關了

        Dim flag_dataExists As Boolean = Utl_SearchExp1(DGrid1)
        If Not flag_dataExists Then Return

        Dim sFileName1 As String = "送訓官兵名冊"
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(cst_功能).Visible=False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div2.RenderControl(objHtmlTextWriter)
        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        DGrid1.AllowPaging = True
        'DataGrid1.Columns(cst_功能).Visible=True
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    ''' <summary>
    ''' 匯出鈕
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Call Search1Value()  '記錄查詢條件
        Call SUB_Export1(DataGrid2)
    End Sub
End Class
