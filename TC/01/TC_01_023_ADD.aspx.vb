Public Class TC_01_023_ADD
    Inherits AuthBasePage

    'ORG_REMOTER
    Dim dtPIC As DataTable
    Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)"
    'Const cst_errMsg_6 As String = "(檔案上傳失敗，請刪除此筆後重新上傳)"
    Const cst_PostedFile_MAX_SIZE_10M As Integer = 10485760 '10*1024*1024=10,485,760  '2*1024*1024=2,097,152
    Const cst_PostedFile_MAX_SIZE_15M As Integer = 15728640 '1024*1024*15=15728640
    'Const cst_errMsg_7 As String = "檔案大小超過2MB!"
    Const cst_errMsg_7_10M As String = "檔案大小超過10MB!"
    Const cst_errMsg_7_15M As String = "檔案大小超過15MB!"
    'Const cst_FileDescMsg_7_10M As String = "PDF(掃瞄畫面需清楚，檔案大小限制10MB以下)!"
    'Const cst_FileDescMsg_7_15M As String = "PDF(掃瞄畫面需清楚，檔案大小限制15MB以下)!"
    Const cst_pic_iWidth As Integer = 960 '480
    Const cst_pic_iHeight As Integer = 480 '240

    Dim rqProcessType As String = ""
    Dim irqProcessType As Integer = 0

    Const cst_rqProcessType As String = "ProcessType"
    Const cst_rqProcessType_Update As String = "Update"
    Const cst_rqProcessType_Insert As String = "Insert"
    Const cst_rqProcessType_View As String = "View"

    Enum eePT_Enum As Int32
        xInsert = 10 'Insert
        xView = 20 'View
        xUpdate = 30 'Update
    End Enum

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        'Call Get_ServerConfigSet()

        rqProcessType = TIMS.ClearSQM(Request(cst_rqProcessType))
        If Hid_RMTID.Value = "" Then Hid_RMTID.Value = TIMS.ClearSQM(Request("RMTID"))
        RIDValue.Value = TIMS.ClearSQM(Request("RID"))
        Dim v_ORGID As String = TIMS.ClearSQM(Request("OrgID"))
        If v_ORGID = "" OrElse v_ORGID = "-1" Then v_ORGID = TIMS.Get_OrgID(RIDValue.Value, objconn)
        If v_ORGID = "" OrElse v_ORGID = "-1" Then v_ORGID = sm.UserInfo.OrgID
        If v_ORGID <> "" Then Hid_ORGID.Value = v_ORGID
        Hid_COMIDNO.Value = TIMS.Get_ComIDNOforOrgID(v_ORGID, objconn)
        labORGNAME.Text = TIMS.GET_ORGNAME(v_ORGID, objconn)

        irqProcessType = eePT_Enum.xUpdate
        Select Case rqProcessType
            Case cst_rqProcessType_Insert '"Insert"
                irqProcessType = eePT_Enum.xInsert
            Case cst_rqProcessType_Update '"Update"
                irqProcessType = eePT_Enum.xUpdate
            Case cst_rqProcessType_View '"View"
                irqProcessType = eePT_Enum.xView
            Case Else
                irqProcessType = eePT_Enum.xView
        End Select

        If Not IsPostBack Then
            '產生新的GUID 避免記憶體相同 而異常
            'Call CREATE_NEW_GUID21()
            Call CCreate1()
            Call SHOW_REMOTER_PIC() '顯示圖檔資料表
        End If
    End Sub

    Sub CCreate1()
        Dim str_rtn_checkFile1 As String = String.Concat("return checkFile1(", cst_PostedFile_MAX_SIZE_10M, ");")
        But1.Attributes.Add("onclick", str_rtn_checkFile1)

        'TIMS.SetMyValue(s_MyValue1, "RMTNO", RMTNO.Text)
        'TIMS.SetMyValue(s_MyValue1, "RMTNAME", RMTNAME.Text)
        'TIMS.SetMyValue(s_MyValue1, "ProcessType", "Insert")
        'TIMS.SetMyValue(s_MyValue1, "RID", RIDValue.Value)
        'TIMS.SetMyValue(s_MyValue1, "OrgID", orgid_value.Value)

        TIMS.PL_placeholder(VIDEODEVICE)
        TIMS.PL_placeholder(SOFTDESC)
        TIMS.PL_placeholder(DEVICEDESC)
        'But1.Attributes("onclick") = "return CheckAddPIC();"

        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)

        cbl_TEACHSOFT = TIMS.GET_CBL_TEACHSOFT(objconn, cbl_TEACHSOFT)
        cbl_TEACHDEVICE = TIMS.GET_CBL_TEACHDEVICE(objconn, cbl_TEACHDEVICE)

        Call Create_DDL_RMTPIC(dtPIC)

        Select Case irqProcessType
            Case eePT_Enum.xInsert
                'ProcessType.Text = "-新增"
                RMTNO.Text = TIMS.ClearSQM(Request("RMTNO"))
                RMTNAME.Text = TIMS.ClearSQM(Request("RMTNAME"))
                CtrlFormEnableInsert()
            Case eePT_Enum.xView '2011-11-13 add檢視
                'ProcessType.Text = "-檢視"
                If Hid_RMTID.Value <> "" Then LoalData1(Hid_RMTID.Value)
                CtrlFormEnable(False)
            Case Else 'Case eePT_Enum.xUpdate
                'ProcessType.Text = "-修改"
                If Hid_RMTID.Value <> "" Then LoalData1(Hid_RMTID.Value)
        End Select

    End Sub

    Sub Create_DDL_RMTPIC(ByRef dtPRI As DataTable)

        With DDL_RMTPIC
            .Items.Clear()
            .Items.Add(New ListItem("==請選擇==", ""))
            .Items.Add(New ListItem("1.教學設備照片", "1"))
            .Items.Add(New ListItem("2.網路環境照片", "2"))
            .Items.Add(New ListItem("3.錄影設備照片", "3"))
            .Items.Add(New ListItem("4.其他照片", "4"))
        End With

        If dtPRI Is Nothing OrElse dtPRI.Rows.Count = 0 Then Return

        For Each dr1 As DataRow In dtPRI.Rows
            Dim depID As String = Convert.ToString(dr1("depID"))
            Dim litem As ListItem = DDL_RMTPIC.Items.FindByValue(depID)
            If depID <> "" AndAlso litem IsNot Nothing Then DDL_RMTPIC.Items.Remove(litem)
        Next

    End Sub

    Private Sub LoalData1(ByVal RMTID As String)
        Hid_ORGID.Value = TIMS.ClearSQM(Hid_ORGID.Value)
        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
        If RMTID = "" OrElse Hid_RMTID.Value = "" OrElse RMTID <> Hid_RMTID.Value Then Return
        If Hid_ORGID.Value = "" Then Return

        Dim hPMS As New Hashtable From {{"ORGID", Val(Hid_ORGID.Value)}, {"RMTID", Val(RMTID)}}
        Dim sSql As String = ""
        sSql &= " SELECT a.RMTID ,a.RMTNO,a.RMTNAME" & vbCrLf
        sSql &= " ,a.ORGID" & vbCrLf
        sSql &= " ,a.KTSID,dbo.FN_TEACHSOFT_N(a.KTSID) TEACHSOFT_N" & vbCrLf
        sSql &= " ,a.KTDID,dbo.FN_TEACHDEVICE_N(a.KTDID) TEACHDEVICE_N" & vbCrLf
        sSql &= " ,a.TEACHSOFT_OTH,a.TEACHDEVICE_OTH" & vbCrLf
        sSql &= " ,a.CABLENETWORK,a.CABLEDLRATE,a.CABLEUPRATE" & vbCrLf
        sSql &= " ,a.WIFINETWORK,a.WIFIDLRATE,a.WIFIUPRATE" & vbCrLf
        sSql &= " ,a.VIDEODEVICE ,a.SOFTDESC,a.DEVICEDESC" & vbCrLf
        'sSql &= " ,a.PIC_DEVICE,a.PIC_NETWORK" & vbCrLf
        sSql &= " ,a.RMTPIC1,a.RMTPIC2,a.RMTPIC3,a.RMTPIC4" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE,a.MODIFYTYPE" & vbCrLf
        sSql &= " FROM ORG_REMOTER a" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.ORGID=a.ORGID" & vbCrLf
        sSql &= " WHERE a.ORGID=@ORGID AND a.RMTID=@RMTID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, hPMS)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt.Rows(0)
        Hid_RMTID.Value = Convert.ToString(dr1("RMTID"))
        RMTNO.Text = Convert.ToString(dr1("RMTNO"))
        RMTNAME.Text = Convert.ToString(dr1("RMTNAME"))
        TIMS.SetCblValue(cbl_TEACHSOFT, Convert.ToString(dr1("KTSID")))
        TIMS.SetCblValue(cbl_TEACHDEVICE, Convert.ToString(dr1("KTDID")))
        TEACHSOFT_OTH.Text = Convert.ToString(dr1("TEACHSOFT_OTH"))
        TEACHDEVICE_OTH.Text = Convert.ToString(dr1("TEACHDEVICE_OTH"))

        CBX_CABLENETWORK.Checked = If(Convert.ToString(dr1("CABLENETWORK")) = "Y", True, False)
        CABLEDLRATE.Text = Convert.ToString(dr1("CABLEDLRATE"))
        CABLEUPRATE.Text = Convert.ToString(dr1("CABLEUPRATE"))

        CBX_WIFINETWORK.Checked = If(Convert.ToString(dr1("WIFINETWORK")) = "Y", True, False)
        WIFIDLRATE.Text = Convert.ToString(dr1("WIFIDLRATE"))
        WIFIUPRATE.Text = Convert.ToString(dr1("WIFIUPRATE"))

        VIDEODEVICE.Text = Convert.ToString(dr1("VIDEODEVICE"))
        SOFTDESC.Text = Convert.ToString(dr1("SOFTDESC"))
        DEVICEDESC.Text = Convert.ToString(dr1("DEVICEDESC"))

    End Sub

    Sub CtrlFormEnableInsert()
        Dim t_msg1 As String = "(待第1次儲存後，序號產生，才可上傳照片)"
        Dim fg_Enable As Boolean = False
        DDL_RMTPIC.Enabled = fg_Enable
        File1.Disabled = (Not fg_Enable)
        But1.Visible = fg_Enable
        TIMS.Tooltip(DDL_RMTPIC, t_msg1)
        TIMS.Tooltip(File1, t_msg1)
        TIMS.Tooltip(But1, t_msg1)
    End Sub

    Private Sub CtrlFormEnable(fg_Enable As Boolean)

        RMTNO.Enabled = fg_Enable
        RMTNAME.Enabled = fg_Enable
        cbl_TEACHSOFT.Enabled = fg_Enable
        TEACHSOFT_OTH.Enabled = fg_Enable
        cbl_TEACHDEVICE.Enabled = fg_Enable
        TEACHDEVICE_OTH.Enabled = fg_Enable
        CBX_CABLENETWORK.Enabled = fg_Enable
        CABLEDLRATE.Enabled = fg_Enable
        CABLEUPRATE.Enabled = fg_Enable
        CBX_WIFINETWORK.Enabled = fg_Enable
        WIFIDLRATE.Enabled = fg_Enable
        WIFIUPRATE.Enabled = fg_Enable
        'Labmsg_NETWORK.Enabled = fg_Enable
        rblMODIFYTYPE.Enabled = fg_Enable
        VIDEODEVICE.Enabled = fg_Enable
        SOFTDESC.Enabled = fg_Enable
        DEVICEDESC.Enabled = fg_Enable
        'labMsg1.Enabled = fg_Enable
        'DataGrid3Table.Enabled = fg_Enable
        'DataGrid3.Enabled = fg_Enable
        'RMTPIC.Enabled = fg_Enable
        '照片種類> .Enabled = fg_Enable
        '圖檔名稱 > .Enabled = fg_Enable
        DDL_RMTPIC.Enabled = fg_Enable
        File1.Disabled = (Not fg_Enable)
        But1.Visible = fg_Enable
        '(儲存鈕)
        BtnSAVEDATA1.Visible = fg_Enable
    End Sub

    ''產生新的GUID 避免記憶體相同 而異常
    'Sub CREATE_NEW_GUID21()
    '    'If IsPostBack Then Exit Sub
    '    Hid_ORG_REMOTER_GUID1.Value = TIMS.GetGUID()
    '    Session(Hid_ORG_REMOTER_GUID1.Value) = Nothing
    'End Sub

    Protected Sub BtnGOBACK_Click(sender As Object, e As EventArgs) Handles BtnGOBACK.Click
        Dim url1 As String = "TC_01_023?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Protected Sub BtnSAVEDATA1_Click(sender As Object, e As EventArgs) Handles BtnSAVEDATA1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, String.Concat(Errmsg, "請確認"))
            Exit Sub
        End If

        Call SAVEDATA1()
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True

        Hid_ORGID.Value = TIMS.ClearSQM(Hid_ORGID.Value)
        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
        Dim v_cbl_TEACHSOFT As String = TIMS.GetChkBoxListValue(cbl_TEACHSOFT)
        Dim v_cbl_TEACHDEVICE As String = TIMS.GetChkBoxListValue(cbl_TEACHDEVICE)
        TEACHSOFT_OTH.Text = TIMS.ClearSQM(TEACHSOFT_OTH.Text)
        TEACHDEVICE_OTH.Text = TIMS.ClearSQM(TEACHDEVICE_OTH.Text)
        CABLEDLRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(CABLEDLRATE.Text)))
        CABLEUPRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(CABLEUPRATE.Text)))
        WIFIDLRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(WIFIDLRATE.Text)))
        WIFIUPRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(WIFIUPRATE.Text)))

        TIMS.Chk_placeholder(VIDEODEVICE)
        TIMS.Chk_placeholder(SOFTDESC)
        TIMS.Chk_placeholder(DEVICEDESC)
        VIDEODEVICE.Text = TIMS.ClearSQM(VIDEODEVICE.Text)
        SOFTDESC.Text = TIMS.ClearSQM(SOFTDESC.Text)
        DEVICEDESC.Text = TIMS.ClearSQM(DEVICEDESC.Text)

        RMTNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(RMTNO.Text))
        'If Hid_RMTID.Value = "" Then
        '    '修改時不檢核
        'End If

        If RMTNO.Text <> "" Then
            If Not TIMS.CheckABC123(RMTNO.Text) Then
                Errmsg &= "環境代號 必須為英數字" & vbCrLf
            ElseIf RMTNO.Text.Length > 10 Then
                Errmsg &= "環境代號 限英數字10字以內" & vbCrLf
            ElseIf RMTNO.Text.Length < 1 Then
                Errmsg &= "環境代號 至少英數字1字以上" & vbCrLf
            ElseIf Not TIMS.CheckABC321(RMTNO.Text) Then
                Errmsg &= "環境代號 必須為英數字10字以內" & vbCrLf
            End If
        End If

        RMTNO.Text = TIMS.ClearSQM(RMTNO.Text)
        If RMTNO.Text = "" Then Errmsg &= "請輸入 環境代號" & vbCrLf
        RMTNAME.Text = TIMS.ClearSQM(RMTNAME.Text)
        If RMTNAME.Text = "" Then Errmsg &= "請輸入 環境名稱" & vbCrLf
        If v_cbl_TEACHSOFT = "" Then Errmsg &= "請選擇 教學軟體" & vbCrLf
        If v_cbl_TEACHDEVICE = "" Then Errmsg &= "請選擇 教學設備" & vbCrLf

        If v_cbl_TEACHSOFT <> "" AndAlso v_cbl_TEACHSOFT.Contains("99") AndAlso TEACHSOFT_OTH.Text = "" Then
            Errmsg &= "教學軟體 有選擇「其他」，請輸入 其他(請說明)" & vbCrLf
        End If
        If v_cbl_TEACHDEVICE <> "" AndAlso v_cbl_TEACHDEVICE.Contains("99") AndAlso TEACHDEVICE_OTH.Text = "" Then
            Errmsg &= "教學設備 有選擇「其他」，請輸入 其他(請說明)" & vbCrLf
        End If

        If CBX_CABLENETWORK.Checked Then
            If CABLEDLRATE.Text = "" Then Errmsg &= "選擇 有線網路 請輸入 下載速率" & vbCrLf
            If CABLEUPRATE.Text = "" Then Errmsg &= "選擇 有線網路 請輸入 上傳速率" & vbCrLf
        End If
        If CBX_WIFINETWORK.Checked Then
            If WIFIDLRATE.Text = "" Then Errmsg &= "選擇 無線網路 請輸入 下載速率" & vbCrLf
            If WIFIUPRATE.Text = "" Then Errmsg &= "選擇 無線網路 請輸入 上傳速率" & vbCrLf
        End If
        If Not CBX_CABLENETWORK.Checked AndAlso Not CBX_WIFINETWORK.Checked Then
            Errmsg &= "請勾選 有線網路 或 無線網路(至少1項)" & vbCrLf
        End If

        If CABLEDLRATE.Text <> "" AndAlso Not TIMS.IsNumberStr(CABLEDLRATE.Text) Then
            Errmsg &= "有線網路-下載速率 請輸入數字" & vbCrLf
        ElseIf CABLEDLRATE.Text <> "" AndAlso Not TIMS.IsNumeric2(CABLEDLRATE.Text) Then
            Errmsg &= "有線網路-下載速率 請輸入(整數)數字" & vbCrLf
        End If
        If CABLEUPRATE.Text <> "" AndAlso Not TIMS.IsNumberStr(CABLEUPRATE.Text) Then
            Errmsg &= "有線網路-上傳速率 請輸入數字" & vbCrLf
        ElseIf CABLEUPRATE.Text <> "" AndAlso Not TIMS.IsNumeric2(CABLEUPRATE.Text) Then
            Errmsg &= "有線網路-上傳速率 請輸入(整數)數字" & vbCrLf
        End If
        If WIFIDLRATE.Text <> "" AndAlso Not TIMS.IsNumberStr(WIFIDLRATE.Text) Then
            Errmsg &= "無線網路-下載速率 請輸入數字" & vbCrLf
        ElseIf WIFIDLRATE.Text <> "" AndAlso Not TIMS.IsNumeric2(WIFIDLRATE.Text) Then
            Errmsg &= "無線網路-下載速率 請輸入(整數)數字" & vbCrLf
        End If
        If WIFIUPRATE.Text <> "" AndAlso Not TIMS.IsNumberStr(WIFIUPRATE.Text) Then
            Errmsg &= "無線網路-上傳速率 請輸入數字" & vbCrLf
        ElseIf WIFIUPRATE.Text <> "" AndAlso Not TIMS.IsNumeric2(WIFIUPRATE.Text) Then
            Errmsg &= "無線網路-上傳速率 請輸入(整數)數字" & vbCrLf
        End If

        If VIDEODEVICE.Text = "" Then Errmsg &= "請輸入 教學錄影設備" & vbCrLf
        If SOFTDESC.Text = "" Then Errmsg &= "請輸入 教學軟體及設備說明 軟體說明文字" & vbCrLf
        If DEVICEDESC.Text = "" Then Errmsg &= "請輸入 教學軟體及設備說明 設備說明文字" & vbCrLf

        'If FactMode.SelectedValue = "" Then Errmsg += "請選擇 場地類型" & vbCrLf
        'ConNum.Text = TIMS.ClearSQM(ConNum.Text)
        'If ConNum.Text = "" Then Errmsg += "請輸入 訓練容納人數" & vbCrLf
        'If ConNum.Text <> "" Then
        '    If Not TIMS.IsNumeric1(ConNum.Text) Then Errmsg += "訓練容納人數 應為數字格式" & vbCrLf
        'End If
        'txtPingNumber.Text = TIMS.ClearSQM(txtPingNumber.Text)
        'If txtPingNumber.Text = "" Then Errmsg &= "請輸入 坪數" & vbCrLf
        'If txtPingNumber.Text <> "" Then
        '    If TIMS.IsNumeric1(txtPingNumber.Text) Then txtPingNumber.Text = Val(TIMS.ROUND(txtPingNumber.Text, 4))
        '    If Not TIMS.IsNumeric1(txtPingNumber.Text) Then Errmsg &= "坪數 應為數字格式(可含小數點4位)" & vbCrLf
        'End If
        'city_code.Value = TIMS.ClearSQM(city_code.Value) '場地郵遞區號
        'ZIPB3.Value = TIMS.ClearSQM(ZIPB3.Value)
        'hidZIP6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZIPB3.Value)
        'Address.Text = TIMS.ClearSQM(Address.Text) '場地地址

        'If Not TIMS.IsZipCode(city_code.Value, objconn) Then Errmsg += "場地地址 郵遞區號前3碼 有誤" & vbCrLf
        'TIMS.CheckZipCODEB3(ZIPB3.Value, "場地地址 郵遞區號後2碼或3碼", True, Errmsg)
        'If Address.Text = "" Then Errmsg += "場地地址 不可為空" & vbCrLf

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    Private Sub SAVEDATA1()

        Dim iRst As Integer = 0
        Hid_ORGID.Value = TIMS.ClearSQM(Hid_ORGID.Value)
        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
        Dim v_cbl_TEACHSOFT As String = TIMS.GetChkBoxListValue(cbl_TEACHSOFT)
        Dim v_cbl_TEACHDEVICE As String = TIMS.GetChkBoxListValue(cbl_TEACHDEVICE)
        TEACHSOFT_OTH.Text = TIMS.ClearSQM(TEACHSOFT_OTH.Text)
        TEACHDEVICE_OTH.Text = TIMS.ClearSQM(TEACHDEVICE_OTH.Text)
        CABLEDLRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(CABLEDLRATE.Text)))
        CABLEUPRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(CABLEUPRATE.Text)))
        WIFIDLRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(WIFIDLRATE.Text)))
        WIFIUPRATE.Text = TIMS.ChangABC123(TIMS.ChangeIDNO(TIMS.ClearSQM(WIFIUPRATE.Text)))

        VIDEODEVICE.Text = TIMS.ClearSQM(VIDEODEVICE.Text)
        SOFTDESC.Text = TIMS.ClearSQM(SOFTDESC.Text)
        DEVICEDESC.Text = TIMS.ClearSQM(DEVICEDESC.Text)
        Dim v_rblMODIFYTYPE As String = TIMS.GetListValue(rblMODIFYTYPE)
        Dim v_MODIFYTYPE As String = If(v_rblMODIFYTYPE = "N", "D", "")
        Dim vORGID As String = TIMS.Get_OrgIDforComIDNO(objconn, Hid_COMIDNO.Value)
        If Hid_COMIDNO.Value = "" OrElse Hid_ORGID.Value = "" OrElse Hid_ORGID.Value <> vORGID Then Return

        If Hid_RMTID.Value = "" Then
            Dim iRMTID As Integer = DbAccess.GetNewId(objconn, "ORG_REMOTER_RMTID_SEQ,ORG_REMOTER,RMTID")
            Dim iParms As New Hashtable
            iParms.Add("RMTID", iRMTID)
            iParms.Add("RMTNO", RMTNO.Text)
            iParms.Add("RMTNAME", RMTNAME.Text)
            iParms.Add("ORGID", Val(Hid_ORGID.Value))
            iParms.Add("KTSID", v_cbl_TEACHSOFT)
            iParms.Add("KTDID", v_cbl_TEACHDEVICE)
            iParms.Add("TEACHSOFT_OTH", TEACHSOFT_OTH.Text)
            iParms.Add("TEACHDEVICE_OTH", TEACHDEVICE_OTH.Text)
            iParms.Add("CABLENETWORK", If(CBX_CABLENETWORK.Checked, "Y", Convert.DBNull))
            iParms.Add("CABLEDLRATE", If(CABLEDLRATE.Text <> "", Val(CABLEDLRATE.Text), Convert.DBNull))
            iParms.Add("CABLEUPRATE", If(CABLEUPRATE.Text <> "", Val(CABLEUPRATE.Text), Convert.DBNull))
            iParms.Add("WIFINETWORK", If(CBX_WIFINETWORK.Checked, "Y", Convert.DBNull))
            iParms.Add("WIFIDLRATE", If(WIFIDLRATE.Text <> "", Val(WIFIDLRATE.Text), Convert.DBNull))
            iParms.Add("WIFIUPRATE", If(WIFIUPRATE.Text <> "", Val(WIFIUPRATE.Text), Convert.DBNull))
            iParms.Add("VIDEODEVICE", VIDEODEVICE.Text)
            iParms.Add("SOFTDESC", SOFTDESC.Text)
            iParms.Add("DEVICEDESC", DEVICEDESC.Text)
            iParms.Add("RMTPIC1", Convert.DBNull)
            iParms.Add("RMTPIC2", Convert.DBNull)
            iParms.Add("RMTPIC3", Convert.DBNull)
            iParms.Add("RMTPIC4", Convert.DBNull)
            iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            iParms.Add("MODIFYTYPE", If(v_MODIFYTYPE <> "", v_MODIFYTYPE, Convert.DBNull))
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_REMOTER(RMTID,RMTNO,RMTNAME" & vbCrLf
            isSql &= " ,ORGID,KTSID,KTDID,TEACHSOFT_OTH,TEACHDEVICE_OTH" & vbCrLf
            isSql &= " ,CABLENETWORK,CABLEDLRATE,CABLEUPRATE ,WIFINETWORK,WIFIDLRATE,WIFIUPRATE" & vbCrLf
            isSql &= " ,VIDEODEVICE,SOFTDESC,DEVICEDESC" & vbCrLf
            isSql &= " ,RMTPIC1,RMTPIC2,RMTPIC3,RMTPIC4" & vbCrLf
            isSql &= " ,MODIFYACCT,MODIFYDATE,MODIFYTYPE)" & vbCrLf
            isSql &= " VALUES (@RMTID,@RMTNO,@RMTNAME" & vbCrLf
            isSql &= " ,@ORGID,@KTSID,@KTDID,@TEACHSOFT_OTH,@TEACHDEVICE_OTH" & vbCrLf
            isSql &= " ,@CABLENETWORK,@CABLEDLRATE,@CABLEUPRATE ,@WIFINETWORK,@WIFIDLRATE,@WIFIUPRATE" & vbCrLf
            isSql &= " ,@VIDEODEVICE,@SOFTDESC,@DEVICEDESC" & vbCrLf
            isSql &= " ,@RMTPIC1,@RMTPIC2,@RMTPIC3,@RMTPIC4" & vbCrLf
            isSql &= " ,@MODIFYACCT,GETDATE(),@MODIFYTYPE)" & vbCrLf
            iRst = DbAccess.ExecuteNonQuery(isSql, objconn, iParms)

        Else
            Dim iRMTID As Integer = If(Hid_RMTID.Value <> "", Val(Hid_RMTID.Value), -1)
            Dim uParms As New Hashtable
            uParms.Add("RMTID", iRMTID)
            uParms.Add("RMTNO", RMTNO.Text)
            uParms.Add("RMTNAME", RMTNAME.Text)
            uParms.Add("ORGID", Val(Hid_ORGID.Value))
            uParms.Add("KTSID", v_cbl_TEACHSOFT)
            uParms.Add("KTDID", v_cbl_TEACHDEVICE)
            uParms.Add("TEACHSOFT_OTH", TEACHSOFT_OTH.Text)
            uParms.Add("TEACHDEVICE_OTH", TEACHDEVICE_OTH.Text)
            uParms.Add("CABLENETWORK", If(CBX_CABLENETWORK.Checked, "Y", Convert.DBNull))
            uParms.Add("CABLEDLRATE", If(CABLEDLRATE.Text <> "", Val(CABLEDLRATE.Text), Convert.DBNull))
            uParms.Add("CABLEUPRATE", If(CABLEUPRATE.Text <> "", Val(CABLEUPRATE.Text), Convert.DBNull))
            uParms.Add("WIFINETWORK", If(CBX_WIFINETWORK.Checked, "Y", Convert.DBNull))
            uParms.Add("WIFIDLRATE", If(WIFIDLRATE.Text <> "", Val(WIFIDLRATE.Text), Convert.DBNull))
            uParms.Add("WIFIUPRATE", If(WIFIUPRATE.Text <> "", Val(WIFIUPRATE.Text), Convert.DBNull))
            uParms.Add("VIDEODEVICE", VIDEODEVICE.Text)
            uParms.Add("SOFTDESC", SOFTDESC.Text)
            uParms.Add("DEVICEDESC", DEVICEDESC.Text)
            'uParms.Add("RMTPIC1", RMTPIC1)
            'uParms.Add("RMTPIC2", RMTPIC2)
            'uParms.Add("RMTPIC3", RMTPIC3)
            'uParms.Add("RMTPIC4", RMTPIC4)
            uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
            'uParms.Add("MODIFYDATE", MODIFYDATE)
            uParms.Add("MODIFYTYPE", If(v_MODIFYTYPE <> "", v_MODIFYTYPE, Convert.DBNull))
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_REMOTER" & vbCrLf
            usSql &= " SET RMTID=@RMTID,RMTNO=@RMTNO,RMTNAME=@RMTNAME" & vbCrLf
            usSql &= " ,ORGID=@ORGID,KTSID=@KTSID,KTDID=@KTDID" & vbCrLf
            usSql &= " ,TEACHSOFT_OTH=@TEACHSOFT_OTH,TEACHDEVICE_OTH=@TEACHDEVICE_OTH" & vbCrLf
            usSql &= " ,CABLENETWORK=@CABLENETWORK,CABLEDLRATE=@CABLEDLRATE,CABLEUPRATE=@CABLEUPRATE" & vbCrLf
            usSql &= " ,WIFINETWORK=@WIFINETWORK,WIFIDLRATE=@WIFIDLRATE,WIFIUPRATE=@WIFIUPRATE" & vbCrLf
            usSql &= " ,VIDEODEVICE=@VIDEODEVICE,SOFTDESC=@SOFTDESC,DEVICEDESC=@DEVICEDESC" & vbCrLf
            'usSql &= " ,RMTPIC1=@RMTPIC1,RMTPIC2=@RMTPIC2,RMTPIC3=@RMTPIC3,RMTPIC4=@RMTPIC4" & vbCrLf
            usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE(),MODIFYTYPE=@MODIFYTYPE" & vbCrLf
            usSql &= " WHERE RMTID=@RMTID"
            iRst = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        End If

        If iRst <= 0 Then Return

        Const Cst_msg5 As String = "資料新增成功!"
        Const Cst_msg6 As String = "資料修改成功!"
        Dim msg_tmp As String = ""
        '訊息選擇
        Select Case rqProcessType
            Case cst_rqProcessType_Insert
                'Cst_新增
                msg_tmp += Cst_msg5
            Case cst_rqProcessType_Update
                'Cst_修改
                msg_tmp += Cst_msg6
        End Select

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim s_Script1 As String = ""
        s_Script1 &= String.Concat("<script language=javascript>", "window.alert('", msg_tmp, "');")
        s_Script1 &= String.Concat("window.location.href='TC_01_023.aspx?ID=", MRqID, "';", "</script>")
        TIMS.Utl_RespWriteEnd(Me, objconn, s_Script1)
    End Sub

    Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click

        Dim iRst As Integer = 0
        Hid_ORGID.Value = TIMS.ClearSQM(Hid_ORGID.Value)
        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
        Dim v_cbl_TEACHSOFT As String = TIMS.GetChkBoxListValue(cbl_TEACHSOFT)
        Dim v_cbl_TEACHDEVICE As String = TIMS.GetChkBoxListValue(cbl_TEACHDEVICE)
        TEACHSOFT_OTH.Text = TIMS.ClearSQM(TEACHSOFT_OTH.Text)
        TEACHDEVICE_OTH.Text = TIMS.ClearSQM(TEACHDEVICE_OTH.Text)
        CABLEDLRATE.Text = TIMS.ClearSQM(CABLEDLRATE.Text)
        CABLEUPRATE.Text = TIMS.ClearSQM(CABLEUPRATE.Text)
        WIFIDLRATE.Text = TIMS.ClearSQM(WIFIDLRATE.Text)
        WIFIUPRATE.Text = TIMS.ClearSQM(WIFIUPRATE.Text)

        VIDEODEVICE.Text = TIMS.ClearSQM(VIDEODEVICE.Text)
        SOFTDESC.Text = TIMS.ClearSQM(SOFTDESC.Text)
        DEVICEDESC.Text = TIMS.ClearSQM(DEVICEDESC.Text)
        Dim v_rblMODIFYTYPE As String = TIMS.GetListValue(rblMODIFYTYPE)
        Dim v_MODIFYTYPE As String = If(v_rblMODIFYTYPE = "N", "D", "")
        Dim vORGID As String = TIMS.Get_OrgIDforComIDNO(objconn, Hid_COMIDNO.Value)
        If Hid_COMIDNO.Value = "" OrElse Hid_ORGID.Value = "" OrElse Hid_ORGID.Value <> vORGID Then Return

        If Hid_RMTID.Value = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(序號為空)，請先儲存資料，再重新操作!")
            Return
        End If
        Dim V_DDL_RMTPIC As String = TIMS.GetListValue(DDL_RMTPIC)
        If V_DDL_RMTPIC = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(上傳類別為空)，請先選擇類別，再重新操作!")
            Return
        End If

        If CHK_UPLOAD_REMOTER_PIC(Hid_RMTID.Value, V_DDL_RMTPIC) Then
            Common.MessageBox(Me, "已有上傳資訊(若要重新上傳請先刪除)，再重新操作!")
            Return
        End If
        Call FILE_UPLOAD_REMOTER_PIC(Hid_RMTID.Value, V_DDL_RMTPIC)
        Call SHOW_REMOTER_PIC()

        '顯示上傳檔案／細項
        'Dim rPMS3 As New Hashtable
        'TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        'TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        'Call SHOW_BIDCASEFL_DG2(rPMS3)
    End Sub

    ''' <summary>檢核資料存在TRUE 不存在FALSE</summary>
    ''' <param name="vRMTID"></param>
    ''' <param name="vDDL_RMTPIC"></param>
    ''' <returns></returns>
    Private Function CHK_UPLOAD_REMOTER_PIC(vRMTID As String, vDDL_RMTPIC As String) As Boolean
        Dim fg_RST As Boolean = False
        Dim hPMS As New Hashtable From {{"RMTID", Val(vRMTID)}, {"ORGID", Val(Hid_ORGID.Value)}}
        Dim ssSql As String = ""
        ssSql &= " SELECT 1 COL1 FROM ORG_REMOTER" & vbCrLf
        ssSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
        Select Case Val(vDDL_RMTPIC)
            Case 1
                ssSql &= " AND LEN(RMTPIC1)>1" & vbCrLf
            Case 2
                ssSql &= " AND LEN(RMTPIC2)>1" & vbCrLf
            Case 3
                ssSql &= " AND LEN(RMTPIC3)>1" & vbCrLf
            Case 4
                ssSql &= " AND LEN(RMTPIC4)>1" & vbCrLf
            Case Else
                Return fg_RST
        End Select
        Dim dtRT As DataTable = DbAccess.GetDataTable(ssSql, objconn, hPMS)
        fg_RST = (dtRT.Rows.Count > 0)
        Return fg_RST
    End Function

    Private Sub SHOW_REMOTER_PIC()
        Hid_RMTID.Value = TIMS.ClearSQM(Hid_RMTID.Value)
        Hid_ORGID.Value = TIMS.ClearSQM(Hid_ORGID.Value)

        If Hid_RMTID.Value = "" OrElse Hid_ORGID.Value = "" Then Return

        Dim rPMS As New Hashtable From {{"RMTID", Val(Hid_RMTID.Value)}, {"ORGID", Val(Hid_ORGID.Value)}}

        dtPIC = TIMS.CreatePICRMTdt(Server, rPMS, objconn)

        Call Create_DDL_RMTPIC(dtPIC)

        DataGrid3Table.Visible = False

        If dtPIC Is Nothing OrElse dtPIC.Rows.Count = 0 Then Return

        DataGrid3Table.Visible = True
        DataGrid3.DataSource = dtPIC
        DataGrid3.DataBind()
    End Sub

    'Private Function CreatePICRMTdt() As DataTable
    '    'Dim dtPRI As DataTable = Nothing
    '    Dim hPMS As New Hashtable From {{"RMTID", Val(Hid_RMTID.Value)}, {"ORGID", Val(Hid_ORGID.Value)}}
    '    Dim sSql As String = ""
    '    sSql &= " WITH WC1 AS (SELECT RMTPIC1,RMTPIC2,RMTPIC3,RMTPIC4 FROM ORG_REMOTER WHERE RMTID=@RMTID AND ORGID=@ORGID)" & vbCrLf
    '    sSql &= " SELECT 1 depID,N'教學設備照片' RMTPIC, RMTPIC1 FileName1,' ' okflag FROM WC1 WHERE RMTPIC1 IS NOT NULL" & vbCrLf
    '    sSql &= " UNION SELECT 2 depID,N'網路環境照片' RMTPIC, RMTPIC2 FileName1,' ' okflag  FROM WC1 WHERE RMTPIC2 IS NOT NULL" & vbCrLf
    '    sSql &= " UNION SELECT 3 depID,N'錄影設備照片' RMTPIC, RMTPIC3 FileName1,' ' okflag  FROM WC1 WHERE RMTPIC3 IS NOT NULL" & vbCrLf
    '    sSql &= " UNION SELECT 4 depID,N'其他照片' RMTPIC, RMTPIC4 FileName1,' ' okflag  FROM WC1 WHERE RMTPIC4 IS NOT NULL" & vbCrLf
    '    Dim dtPRI As DataTable = DbAccess.GetDataTable(sSql, objconn, hPMS)
    '    If dtPRI Is Nothing Then Return dtPRI

    '    Dim Upload_Path As String = TIMS.GET_UPLOADPATH1_RMT()
    '    Dim download_js_Path As String = TIMS.GET_DOWNLOADPATH1_RMT()

    '    For Each dr1 As DataRow In dtPRI.Rows
    '        dr1("okflag") = Convert.ToString(dr1("FileName1"))
    '        Dim filename As String = Convert.ToString(dr1("FileName1"))
    '        If filename <> "" Then
    '            Dim flag_PIC_EXISTS As Boolean = TIMS.CHK_PIC_EXISTS(Server, Upload_Path, filename)
    '            Dim urlA1 As String = "<a class='l' target='_blank' href=""" & download_js_Path & filename & """>" & filename & "</a>"
    '            If Not flag_PIC_EXISTS Then urlA1 = "<font color='red'>" & cst_errMsg_6 & "</font>" '表示 檔案不存在
    '            dr1("okflag") = urlA1
    '        End If
    '    Next
    '    Return dtPRI
    'End Function

    Public Shared Function HttpCHKFile1(ByRef MyPage As Page, ByRef File1 As HtmlInputFile, ByRef MyPostedFile As HttpPostedFile) As Boolean
        'Dim Upload_Path As String = "~/images/Placepic/"
        Dim MyFileColl As HttpFileCollection = Nothing
        Try
            'Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
            MyFileColl = HttpContext.Current.Request.Files
            If MyFileColl Is Nothing OrElse MyFileColl.Count = 0 Then
                Common.MessageBox(MyPage, cst_errMsg_2)
                Return False 'Exit Sub
            End If
        Catch ex As Exception
            TIMS.WriteTraceLog(MyPage, ex)

            Common.MessageBox(MyPage, ex.Message)
            Return False 'Exit Sub
        End Try
        'Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)
        MyPostedFile = MyFileColl.Item(0)
        If MyPostedFile Is Nothing Then
            Common.MessageBox(MyPage, cst_errMsg_2)
            Return False 'Exit Sub
        ElseIf File1.Value = "" Then
            Common.MessageBox(MyPage, "請選擇上傳檔案(不可為空)!")
            Return False 'Exit Sub
        ElseIf File1.PostedFile.ContentLength = 0 OrElse MyPostedFile.ContentLength = 0 Then
            Common.MessageBox(MyPage, "檔案位置錯誤!")
            Return False 'Exit Sub
        End If
        Return True
    End Function

    Function GetThumbNail2(ByRef MyPage As Page, ByRef HtPP As Hashtable, ByRef ImgStream As System.IO.Stream) As String
        Dim strFileName As String = TIMS.GetMyValue2(HtPP, "FileName")
        Dim iWidth As Integer = TIMS.GetMyValue2(HtPP, "iWidth", TIMS.cst_oType_obj)
        Dim iheight As Integer = TIMS.GetMyValue2(HtPP, "iheight", TIMS.cst_oType_obj)
        Dim strContentType As String = TIMS.GetMyValue2(HtPP, "ContentType")
        Dim blnGetFromFile As Boolean = TIMS.GetMyValue2(HtPP, "blnGetFromFile", TIMS.cst_oType_obj)
        Dim Upload_Path As String = TIMS.GetMyValue2(HtPP, "Upload_Path")

        Dim oImg As System.Drawing.Image = If(blnGetFromFile, Drawing.Image.FromFile(strFileName), Drawing.Image.FromStream(ImgStream))
        'oImg = oImg.GetThumbnailImage(iWidth, iheight, Nothing, (New IntPtr).Zero)
        'Dim strGuid As String = (New Guid).NewGuid().ToString().ToUpper()
        oImg = oImg.GetThumbnailImage(iWidth, iheight, Nothing, IntPtr.Zero)
        Dim strGuid As String = Guid.NewGuid().ToString().ToUpper()
        Dim strFileExt As String = strFileName.Substring(strFileName.LastIndexOf("."))

        '保存到本地
        'Dim s_save_file_name As String = String.Concat(MyPage.Server.MapPath(Upload_Path), "\", strGuid, strFileExt)
        Dim s_save_file_name As String = String.Concat(MyPage.Server.MapPath(Upload_Path), "\", strFileName)
        oImg.Save(s_save_file_name, TIMS.GetImageType(strContentType))

        Return strFileName 'String.Concat(strGuid, strFileExt)
    End Function

    Private Sub FILE_UPLOAD_REMOTER_PIC(vRMTID As String, v_DDL_RMTPIC As String)
        'https://developer.mozilla.org/en-US/docs/Web/HTTP/Basics_of_HTTP/MIME_types/Common_types
        'Dim Upload_Path As String = "~/images/Placepic/"

        '.jpg,.gif,.bmp,.png
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not HttpCHKFile1(Me, File1, MyPostedFile) Then Return
        'image/gif image/bmp image/png
        Dim fg_fileType As String = ""
        If MyPostedFile.ContentType.Equals("image/jpeg", StringComparison.OrdinalIgnoreCase) Then
            fg_fileType = "jpg"
        ElseIf MyPostedFile.ContentType.Equals("image/gif", StringComparison.OrdinalIgnoreCase) Then
            fg_fileType = "gif"
        ElseIf MyPostedFile.ContentType.Equals("image/bmp", StringComparison.OrdinalIgnoreCase) Then
            fg_fileType = "bmp"
        ElseIf MyPostedFile.ContentType.Equals("image/png", StringComparison.OrdinalIgnoreCase) Then
            fg_fileType = "png"
        Else
            Common.MessageBox(Me, "無效的檔案格式(限.jpg,.gif,.bmp,.png 檔案)")
            Exit Sub
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, fg_fileType) Then Return

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Dim LowerFileType As String = LCase(MyFileType.ToLower())
        Select Case LowerFileType'LCase(MyFileType)
            Case "jpg", "gif", "bmp", "png"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, "檔案類型錯誤，必須為 (限.jpg,.gif,.bmp,.png) 類型檔案!")
                Exit Sub
        End Select

        '上傳檔案 
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_RMT()
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_RMT(vRMTID, v_DDL_RMTPIC, LowerFileType)
        Dim vSRCFILENAME1 As String = MyFileName
        '上傳檔案/存檔：檔名

        Dim HtPP As New Hashtable
        TIMS.SetMyValue2(HtPP, "FileName", vFILENAME1)
        TIMS.SetMyValue2(HtPP, "iWidth", cst_pic_iWidth)
        TIMS.SetMyValue2(HtPP, "iheight", cst_pic_iHeight)
        TIMS.SetMyValue2(HtPP, "ContentType", MyPostedFile.ContentType)
        TIMS.SetMyValue2(HtPP, "blnGetFromFile", False)
        TIMS.SetMyValue2(HtPP, "Upload_Path", vUploadPath)

        Dim objLock_ThumbNail2 As New Object
        SyncLock objLock_ThumbNail2
            Try
                '上傳檔案
                'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
                vFILENAME1 = GetThumbNail2(Me, HtPP, MyPostedFile.InputStream)
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Common.MessageBox(Me, cst_errMsg_2)

                'Common.MessageBox(Me, ex.ToString)
                Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
                strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
                strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
                strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
                strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
                strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
                'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                TIMS.WriteTraceLog(Me, ex, strErrmsg)
                Exit Sub
            End Try
        End SyncLock

        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "ORGID", Val(Hid_ORGID.Value))
            TIMS.SetMyValue2(rPMS2, "RMTID", vRMTID)
            TIMS.SetMyValue2(rPMS2, "DDL_RMTPIC", v_DDL_RMTPIC)
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Call SAVE_ORG_REMOTER_UPLOAD_RMTPIC(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

    End Sub

    Private Sub SAVE_ORG_REMOTER_UPLOAD_RMTPIC(ByRef rPMS2 As Hashtable)
        Dim iRST As Integer = 0
        Dim vUploadPath As String = TIMS.GetMyValue2(rPMS2, "UploadPath")
        Dim vRMTID As String = TIMS.GetMyValue2(rPMS2, "RMTID")
        Dim vORGID As String = TIMS.GetMyValue2(rPMS2, "ORGID")
        Dim vDDL_RMTPIC As String = TIMS.GetMyValue2(rPMS2, "DDL_RMTPIC")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS2, "FILENAME1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS2, "MODIFYACCT")
        Dim vDELPIC As String = TIMS.GetMyValue2(rPMS2, "DELPIC")

        If vDDL_RMTPIC = "" Then Return

        If vDELPIC = "Y" Then
            Dim hPMS As New Hashtable From {{"MODIFYACCT", vMODIFYACCT}, {"RMTID", Val(vRMTID)}, {"ORGID", Val(vORGID)}}
            hPMS.Add("FILENAME1", vFILENAME1)
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_REMOTER" & vbCrLf
            usSql &= " SET MODIFYACCT=@MODIFYACCT, MODIFYDATE=GETDATE()" & vbCrLf
            Select Case Val(vDDL_RMTPIC)
                Case 1
                    usSql &= " ,RMTPIC1=NULL" & vbCrLf
                Case 2
                    usSql &= " ,RMTPIC2=NULL" & vbCrLf
                Case 3
                    usSql &= " ,RMTPIC3=NULL" & vbCrLf
                Case 4
                    usSql &= " ,RMTPIC4=NULL" & vbCrLf
            End Select
            usSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
            Select Case Val(vDDL_RMTPIC)
                Case 1
                    usSql &= " AND RMTPIC1=@FILENAME1" & vbCrLf
                Case 2
                    usSql &= " AND RMTPIC2=@FILENAME1" & vbCrLf
                Case 3
                    usSql &= " AND RMTPIC3=@FILENAME1" & vbCrLf
                Case 4
                    usSql &= " AND RMTPIC4=@FILENAME1" & vbCrLf
            End Select
            iRST = DbAccess.ExecuteNonQuery(usSql, objconn, hPMS)

        Else
            Dim hPMS As New Hashtable From {{"MODIFYACCT", vMODIFYACCT}, {"RMTID", Val(vRMTID)}, {"ORGID", Val(vORGID)}}
            hPMS.Add("FILENAME1", vFILENAME1)
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_REMOTER" & vbCrLf
            usSql &= " SET MODIFYACCT=@MODIFYACCT, MODIFYDATE=GETDATE()" & vbCrLf
            Select Case Val(vDDL_RMTPIC)
                Case 1
                    usSql &= " ,RMTPIC1=@FILENAME1" & vbCrLf
                Case 2
                    usSql &= " ,RMTPIC2=@FILENAME1" & vbCrLf
                Case 3
                    usSql &= " ,RMTPIC3=@FILENAME1" & vbCrLf
                Case 4
                    usSql &= " ,RMTPIC4=@FILENAME1" & vbCrLf
            End Select
            usSql &= " WHERE RMTID=@RMTID AND ORGID=@ORGID" & vbCrLf
            iRST = DbAccess.ExecuteNonQuery(usSql, objconn, hPMS)

        End If

    End Sub

    Private Sub DataGrid3_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HiddepID As HtmlInputHidden = e.Item.FindControl("HiddepID")
                Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim Butdel4 As Button = e.Item.FindControl("Butdel4")

                HiddepID.Value = Convert.ToString(drv("depID"))
                LabFileName1.Text = Convert.ToString(drv("FileName1"))
                HFileName.Value = Convert.ToString(drv("FileName1"))

                If drv("FileName1") <> drv("okflag") Then LabFileName1.Text = drv("okflag").ToString

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "depID", Convert.ToString(drv("depID")))
                TIMS.SetMyValue(sCmdArg, "FileName1", Convert.ToString(drv("FileName1")))

                Butdel4.CommandArgument = sCmdArg '刪除
                Butdel4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '2011-11-13 訓練單位改只能檢視不能修改
                Butdel4.Visible = If(irqProcessType = eePT_Enum.xView, False, True)
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        If e Is Nothing OrElse e.CommandArgument = "" Then Return
        Dim sCmdArg As String = e.CommandArgument
        Dim depID As String = TIMS.GetMyValue(sCmdArg, "depID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FileName1")
        Dim Upload_Path As String = TIMS.GET_UPLOADPATH1_RMT()
        If depID = "" OrElse vFILENAME1 = "" Then Return
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Select Case e.CommandName
            Case "del4"
                Try
                    Dim rPMS2 As New Hashtable
                    TIMS.SetMyValue2(rPMS2, "UploadPath", Upload_Path)
                    TIMS.SetMyValue2(rPMS2, "ORGID", Val(Hid_ORGID.Value))
                    TIMS.SetMyValue2(rPMS2, "RMTID", Val(Hid_RMTID.Value))
                    TIMS.SetMyValue2(rPMS2, "DDL_RMTPIC", Val(depID))
                    TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
                    TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
                    TIMS.SetMyValue2(rPMS2, "DELPIC", "Y")
                    Call SAVE_ORG_REMOTER_UPLOAD_RMTPIC(rPMS2)

                    Dim flag_PIC_EXISTS As Boolean = TIMS.CHK_PIC_EXISTS(Server, Upload_Path, vFILENAME1)
                    If flag_PIC_EXISTS Then IO.File.Delete(Server.MapPath(String.Concat(Upload_Path, vFILENAME1)))

                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    Common.MessageBox(Me, ex.ToString)

                    Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                    Exit Sub 'Throw ex
                End Try
                DataGrid3.EditItemIndex = -1
        End Select
        Call SHOW_REMOTER_PIC() '顯示教室圖檔資料表
    End Sub
End Class