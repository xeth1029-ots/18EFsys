Partial Class TC_10_001
    Inherits AuthBasePage

    'OA_EXAMINER / OA_EXAMINERJOB 'OA_EXAMINER 審查委員
    'OA_EXAMINERJOB 審查委員職類 (審查職類代碼)
    Const cst_ADD1 As String = "ADD1" '新增

    Const cst_EDIT1 As String = "EDIT1" '修改
    Const cst_STOP1 As String = "STOP1" '停用
    Const cst_DEL1 As String = "DEL1" '刪除
    Const cst_VIEW1 As String = "VIEW1" '查看

    Dim gstr_SQL_1 As String = ""
    Dim gstr_ROWVAL_1 As String = ""

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    '審查委員名單 - OA_EXAMINER
    Dim aRECRUIT As String = "" '遴聘類別代碼 1 A-產業界/B-學術界/C-勞工團體代表
    Dim aMBRNAME As String = "" '姓名 審查委員姓名 50
    'Dim aMBRNMSEQ As String = "" '同名序號,同名可補序號 3
    Dim aUNITNAME As String = "" '現職服務機構 50
    Dim aJOBTITLE As String = "" '職稱 100
    Dim aDEGREE As String = "" '學歷 100
    Dim aCERTIFICAT As String = "" '證照 300
    Dim aSPECIALTY As String = "" '專業背景 300
    Dim aEXAMINERJOB As String = "" '審查職類代碼 60 -- SELECT top 10 * FROM dbo.OA_EXAMINERJOB

    Dim aSERVUNIT1 As String = "" '服務單位1 100
    Dim aSERVTIME1 As String = "" '服務時間1 100
    Dim aJOBTITLE1 As String = "" '職稱1 100
    Dim aSERVUNIT2 As String = "" '服務單位2 100
    Dim aSERVTIME2 As String = "" '服務時間2 100
    Dim aJOBTITLE2 As String = "" '職稱2 100
    Dim aSERVUNIT3 As String = "" '服務單位3 100 
    Dim aSERVTIME3 As String = "" '服務時間3 100
    Dim aJOBTITLE3 As String = "" '職稱3 100

    Dim aPHONE As String = "" '連絡電話 50
    Dim aPHONE2 As String = "" '連絡電話 50
    Dim aCONFAX As String = "" '傳真 20
    Dim aCELLPHONE As String = "" '手機 50
    Dim aCELLPHONE2 As String = "" '手機 50
    Dim aEMAIL As String = "" '電子郵件 70
    Dim aEMAIL2 As String = "" '電子郵件 70

    'Dim aMZIPCODE As String = "" '地址-郵遞區號前3碼 3
    'Dim aMZIPCODE6W As String = "" '地址-郵遞區號6碼 6
    Dim aMADDRESS As String = "" '地址 100
    Dim aMADDRESS2 As String = "" '地址2 100
    Dim aRMNOTE1 As String = "" '備註 300

    Dim aPUSHDISTID As String = "" '推薦分署代碼 100
    Dim aPUSHREASON As String = "" '推薦理由 100
    Dim aRUNTRAIN As String = "" '是否辦訓 10
    Dim aTRAINDISTID As String = "" '辦訓轄區 100
    Dim aADDYEARS_ROC As String = "" '新增年度 3 ROC
    Dim aADDYEARS_AD As String = "" '新增年度 4 AD
    Dim aSTOPUSE As String = "" '停用 3

    Const cst_aRECRUIT As Integer = 0 '遴聘類別代碼
    Const cst_aMBRNAME As Integer = 1 '姓名
    Const cst_aUNITNAME As Integer = 2 '現職服務機構
    Const cst_aJOBTITLE As Integer = 3 '職稱
    Const cst_aDEGREE As Integer = 4 '學歷
    Const cst_aCERTIFICAT As Integer = 5 '證照
    Const cst_aSPECIALTY As Integer = 6 '專業背景
    Const cst_aEXAMINERJOB As Integer = 7 '審查職類代碼

    Const cst_aSERVUNIT1 As Integer = 8 '服務單位1
    Const cst_aSERVTIME1 As Integer = 9 '服務時間1
    Const cst_aJOBTITLE1 As Integer = 10 '職稱1
    Const cst_aSERVUNIT2 As Integer = 11 '服務單位2
    Const cst_aSERVTIME2 As Integer = 12 '服務時間2
    Const cst_aJOBTITLE2 As Integer = 13 '職稱2
    Const cst_aSERVUNIT3 As Integer = 14 '服務單位3
    Const cst_aSERVTIME3 As Integer = 15 '服務時間3
    Const cst_aJOBTITLE3 As Integer = 16 '職稱3

    Const cst_aPHONE As Integer = 17 '連絡電話
    Const cst_aPHONE2 As Integer = 18 '連絡電話
    Const cst_aCONFAX As Integer = 19 '傳真
    Const cst_aCELLPHONE As Integer = 20 '手機
    Const cst_aCELLPHONE2 As Integer = 21 '手機
    Const cst_aEMAIL As Integer = 22 '電子郵件
    Const cst_aEMAIL2 As Integer = 23 '電子郵件2 'Const cst_aMZIPCODE As Integer = 20 '地址-郵遞區號前3碼 'Const cst_aMZIPCODE6W As Integer = 21 '地址-郵遞區號後2碼
    Const cst_aMADDRESS As Integer = 24 '地址
    Const cst_aMADDRESS2 As Integer = 25 '地址
    Const cst_aRMNOTE1 As Integer = 26 '備註

    Const cst_aPUSHDISTID As Integer = 27 '推薦分署代碼
    Const cst_aPUSHREASON As Integer = 28 '推薦理由
    Const cst_aRUNTRAIN As Integer = 29 '是否辦訓
    Const cst_aTRAINDISTID As Integer = 30 '辦訓轄區
    Const cst_aADDYEARS_ROC As Integer = 31 '新增年度
    Const cst_aSTOPUSE As Integer = 32 '啟用
    Const cst_Max_a_Len As Integer = 33 '欄位數量

    Dim dtDist As DataTable = Nothing 'TIMS.Get_DistIDdt(objconn)
    Dim dtGCODE3 As DataTable = Nothing 'TIMS.Get_DistIDdt(objconn)
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
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        dtDist = TIMS.Get_DISTIDdt(objconn) 'Dim dtDist As DataTable = TIMS.Get_DistIDdt(objconn)
        '審查職類代碼
        dtGCODE3 = TIMS.Get_GOVCODE3dt(objconn)
        msg1.Text = ""
        ShowButton1()

        If Not IsPostBack Then cCreate1()

    End Sub

    Sub ShowButton1()
        '署：可使用全部功能【查詢】、【匯出】、【新增】、【修改】、【刪除】、【停用】   (現況OK)
        '分署： 系統管理者 ： 可使用【查詢】、【匯出】，其餘看不到或反灰   (現況OK) / 其他角色： 暫不開放使用(目前功能暫不開放給其他群組用)
        '署：可使用修改、停用、刪除。 / 分署：僅顯示查看。
        btnIMPORT1.Visible = If(sm.UserInfo.LID <> 0, False, True)
        btnSave1.Visible = If(sm.UserInfo.LID <> 0, False, True)
        BtnAddnew.Visible = If(sm.UserInfo.LID <> 0, False, True)

        BtnSearch.Enabled = If(sm.UserInfo.LID = 0, True, False)
        BtnExport.Enabled = If(sm.UserInfo.LID = 0, True, False)
        If (sm.UserInfo.LID = 1) Then
            If (sm.UserInfo.RoleID <= 1) Then
                BtnSearch.Enabled = True
                BtnExport.Enabled = True
            End If
        End If
        Const cst_tipmsg1 As String = "分署-系統管理者：可使用"
        If Not BtnSearch.Enabled Then TIMS.Tooltip(BtnSearch, cst_tipmsg1)
        If Not BtnExport.Enabled Then TIMS.Tooltip(BtnExport, cst_tipmsg1)

        '僅署可新增
        btnSave1.Enabled = If(sm.UserInfo.LID = 0, True, False)
        BtnAddnew.Enabled = If(sm.UserInfo.LID = 0, True, False)
        If Not btnSave1.Enabled Then TIMS.Tooltip(btnSave1, "僅署可新增")
        If Not BtnAddnew.Enabled Then TIMS.Tooltip(BtnAddnew, "僅署可新增")

        '署、分署均可使用匯出功能。

        '僅署均可使用匯入功能。分署不可使用。
        btnIMPORT1.Enabled = If(sm.UserInfo.LID = 0, True, False)
        trImport1.Visible = If(sm.UserInfo.LID = 0, True, False)
        If Not btnIMPORT1.Enabled Then TIMS.Tooltip(btnIMPORT1, "僅署可使用匯入功能,分署不可使用")
    End Sub

    Sub cCreate1()
        PageControler1.Visible = False
        btnSave1.Attributes("onclick") = "return chkSaveData1();"
        'MZIPCODE.Attributes.Add("onblur", "getZipName('City4',this,this.value)")
        ''查詢郵遞區號button控制  

        PageControler1.Visible = False
        'a_apost.HRef = TIMS.cst_PostCodeQry
        'lbWorkZIPB3A = TIMS.Get_WorkZIPB3Link(lbWorkZIPB3A)
        'Dim dtDist As DataTable = TIMS.Get_DistIDdt(objconn)
        SCH_PUSHDISTID = TIMS.Get_DistID(SCH_PUSHDISTID, dtDist, Nothing)
        SCH_PUSHDISTID.Items.Insert(0, New ListItem("全部", ""))
        SCH_PUSHDISTID.Attributes("onclick") = "SelectAll('SCH_PUSHDISTID','SCH_PUSHDISTID_List');"

        '審查課程職類
        SCH_cblGOVCODE3 = TIMS.Get_GOVCODE3(objconn, SCH_cblGOVCODE3)
        SCH_cblGOVCODE3.Attributes("onclick") = "SelectAll('SCH_cblGOVCODE3','SCH_cblGOVCODE3Hidden');"

        '審查課程職類
        cblGOVCODE3 = TIMS.Get_GOVCODE3(objconn, cblGOVCODE3)
        cblGOVCODE3.Attributes("onclick") = "SelectAll('cblGOVCODE3','cblGOVCODE3_Hidden');"

        cbPUSHDISTID = TIMS.Get_DistID(cbPUSHDISTID, dtDist, Nothing)
        cbTRAINDISTID = TIMS.Get_DistID(cbTRAINDISTID, dtDist, Nothing)

        '1.【新增年度】最小年度調整為104
        '2.【新增年度】調整為非必填欄位。
        '下拉式選單。最大值  當年度或登入年度
        Dim iSYears As Integer = 2015
        Dim iEYears As Integer = If(Now.Year > sm.UserInfo.Years, Now.Year, sm.UserInfo.Years)
        ddlADDYEARS = TIMS.GetSyear(ddlADDYEARS, iSYears, iEYears, True) '新增年度
        Common.SetListItem(ddlADDYEARS, sm.UserInfo.Years) '新增年度

        SHOW_PANEL(0)
    End Sub

    Sub ClearData1()
        Hid_EMSEQ.Value = ""
        ddlRECRUIT.SelectedIndex = -1

        MBRNAME.Text = "" 'Convert.ToString(dr1("MBRNAME"))
        MBRNMSEQ.Text = ""
        UNITNAME.Text = "" ' Convert.ToString(dr1("UNITNAME"))
        JOBTITLE.Text = "" ' Convert.ToString(dr1("JOBTITLE"))
        DEGREE.Text = "" 'Convert.ToString(dr1("DEGREE"))
        CERTIFICAT.Text = "" 'Convert.ToString(dr1("CERTIFICAT"))
        SPECIALTY.Text = "" 'Convert.ToString(dr1("SPECIALTY"))

        PHONE.Text = "" 'Convert.ToString(dr1("PHONE"))
        PHONE2.Text = "" 'Convert.ToString(dr1("PHONE"))
        CELLPHONE.Text = "" 'Convert.ToString(dr1("CELLPHONE"))
        CELLPHONE2.Text = "" 'Convert.ToString(dr1("CELLPHONE"))
        CONFAX.Text = "" ' Convert.ToString(dr1("CONFAX"))
        EMAIL.Text = "" ' Convert.ToString(dr1("EMAIL"))
        EMAIL2.Text = "" ' Convert.ToString(dr1("EMAIL"))

        'MZIPCODE.Value = "" 'Convert.ToString(dr1("MZIPCODE"))
        'MZIPCODE6W.Value = "" 'Convert.ToString(dr1("MZIPCODE6W"))
        'ZipCode4_N.Value = "" '(特殊地址)
        MADDRESS.Text = "" 'HttpUtility.HtmlDecode(Convert.ToString(dr1("MADDRESS")))
        MADDRESS2.Text = ""
        'Dim tZipLName As String = ""
        'Dim ZipNameN As String = ""
        'tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr1("MZIPCODE")), objconn)
        'ZipNameN = TIMS.Get_ZipNameN(Convert.ToString(dr1("MZIPCODE")), ZipCode4_N.Value, objconn)
        'City4.Text = "" 'ZipNameN & If(tZipLName <> "", "[" & tZipLName & "]", "")

        RMNOTE1.Text = "" ' Convert.ToString(dr1("RMNOTE1"))
        SERVUNIT1.Text = "" 'Convert.ToString(dr1("SERVUNIT1"))
        SERVTIME1.Text = "" 'Convert.ToString(dr1("SERVTIME1"))
        JOBTITLE1.Text = "" ' Convert.ToString(dr1("JOBTITLE1"))
        SERVUNIT2.Text = "" ' Convert.ToString(dr1("SERVUNIT2"))
        SERVTIME2.Text = "" ' Convert.ToString(dr1("SERVTIME2"))
        JOBTITLE2.Text = "" ' Convert.ToString(dr1("JOBTITLE2"))
        SERVUNIT3.Text = "" ' Convert.ToString(dr1("SERVUNIT3"))
        SERVTIME3.Text = "" ' Convert.ToString(dr1("SERVTIME3"))
        JOBTITLE3.Text = "" ' Convert.ToString(dr1("JOBTITLE3"))
        'cbPUSHDISTID.Text = Convert.ToString(dr1("PUSHDISTID"))
        TIMS.SetCblValue(cbPUSHDISTID, "") 'dr1("PUSHDISTID"))

        PUSHREASON.Text = "" 'Convert.ToString(dr1("PUSHREASON"))
        Common.SetListItem(rblRUNTRAIN, "")

        Common.SetListItem(ddlADDYEARS, sm.UserInfo.Years) '新增年度

        TIMS.SetCblValue(cbTRAINDISTID, "")
        TIMS.SetCblValue(cblGOVCODE3, "") '審查課程職類

        Common.SetListItem(rblSTOPUSE, "")
    End Sub

    Sub LoadData1()
        Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
        If Hid_EMSEQ.Value = "" Then Return
        Dim iEMSEQ As Integer = Val(Hid_EMSEQ.Value)
        Dim parms As New Hashtable
        parms.Add("EMSEQ", iEMSEQ)
        Dim sql As String = "SELECT * FROM OA_EXAMINER WHERE EMSEQ=@EMSEQ"
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr1 Is Nothing Then Return

        'EMSEQ.Text = Convert.ToString(dr1("EMSEQ"))
        Common.SetListItem(ddlRECRUIT, dr1("RECRUIT"))

        MBRNAME.Text = Convert.ToString(dr1("MBRNAME"))
        MBRNMSEQ.Text = $"{dr1("MBRNMSEQ")}"
        UNITNAME.Text = Convert.ToString(dr1("UNITNAME"))
        JOBTITLE.Text = Convert.ToString(dr1("JOBTITLE"))
        DEGREE.Text = Convert.ToString(dr1("DEGREE"))
        CERTIFICAT.Text = Convert.ToString(dr1("CERTIFICAT"))
        SPECIALTY.Text = Convert.ToString(dr1("SPECIALTY"))

        PHONE.Text = Convert.ToString(dr1("PHONE"))
        PHONE2.Text = Convert.ToString(dr1("PHONE2"))
        CELLPHONE.Text = Convert.ToString(dr1("CELLPHONE"))
        CELLPHONE2.Text = Convert.ToString(dr1("CELLPHONE2"))
        CONFAX.Text = Convert.ToString(dr1("CONFAX"))
        EMAIL.Text = Convert.ToString(dr1("EMAIL"))
        EMAIL2.Text = Convert.ToString(dr1("EMAIL2"))

        'MZIPCODE.Value = Convert.ToString(dr1("MZIPCODE"))
        'MZIPCODE6W.Value = Convert.ToString(dr1("MZIPCODE6W"))
        'ZipCode4_N.Value = "" '(特殊地址)
        MADDRESS.Text = HttpUtility.HtmlDecode(Convert.ToString(dr1("MADDRESS")))
        MADDRESS2.Text = HttpUtility.HtmlDecode(Convert.ToString(dr1("MADDRESS2")))
        'Dim tZipLName As String = ""
        'Dim ZipNameN As String = ""
        'tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr1("MZIPCODE")), objconn)
        'ZipNameN = TIMS.Get_ZipNameN(Convert.ToString(dr1("MZIPCODE")), ZipCode4_N.Value, objconn)
        'City4.Text = ZipNameN & If(tZipLName <> "", "[" & tZipLName & "]", "")

        RMNOTE1.Text = Convert.ToString(dr1("RMNOTE1"))
        SERVUNIT1.Text = Convert.ToString(dr1("SERVUNIT1"))
        SERVTIME1.Text = Convert.ToString(dr1("SERVTIME1"))
        JOBTITLE1.Text = Convert.ToString(dr1("JOBTITLE1"))
        SERVUNIT2.Text = Convert.ToString(dr1("SERVUNIT2"))
        SERVTIME2.Text = Convert.ToString(dr1("SERVTIME2"))
        JOBTITLE2.Text = Convert.ToString(dr1("JOBTITLE2"))
        SERVUNIT3.Text = Convert.ToString(dr1("SERVUNIT3"))
        SERVTIME3.Text = Convert.ToString(dr1("SERVTIME3"))
        JOBTITLE3.Text = Convert.ToString(dr1("JOBTITLE3"))
        'cbPUSHDISTID.Text = Convert.ToString(dr1("PUSHDISTID"))
        TIMS.SetCblValue(cbPUSHDISTID, dr1("PUSHDISTID"))

        PUSHREASON.Text = Convert.ToString(dr1("PUSHREASON"))
        Common.SetListItem(rblRUNTRAIN, dr1("RUNTRAIN"))

        TIMS.SetCblValue(cbTRAINDISTID, dr1("TRAINDISTID"))
        '審查課程職類 3
        Dim v_GOVCODE3 As String = GET_EXAMINERJOB_GCODE3(objconn, iEMSEQ, 1)
        TIMS.SetCblValue(cblGOVCODE3, v_GOVCODE3) '審查課程職類

        Common.SetListItem(ddlADDYEARS, dr1("ADDYEARS")) '新增年度
        Common.SetListItem(rblSTOPUSE, dr1("STOPUSE"))

        SHOW_PANEL(1)
    End Sub

    '審查課程職類 3
    Private Function GET_EXAMINERJOB_GCODE3(ByRef oConn As SqlConnection, ByRef iEMSEQ As Integer, ByRef iType As Integer) As String
        Dim parms_3 As New Hashtable From {{"EMSEQ", iEMSEQ}}
        Dim sql_3 As String = "SELECT GCODE FROM dbo.OA_EXAMINERJOB WHERE EMSEQ=@EMSEQ ORDER BY GCODE"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sql_3, oConn, parms_3)
        If dt3 Is Nothing OrElse dt3.Rows.Count = 0 Then Return ""

        Dim v_GCODE3 As String = ""
        For Each dr3 As DataRow In dt3.Rows
            If iType = 1 Then
                v_GCODE3 &= String.Concat(If(v_GCODE3 <> "", ",", ""), dr3("GCODE"))
            ElseIf iType = 2 Then
                Dim fff As String = String.Format("GCODE='{0}'", dr3("GCODE"))
                If dtGCODE3.Select(fff).Length > 0 Then
                    'GCODE,CNAME
                    Dim vTmp1 As String = String.Concat("[", dtGCODE3.Select(fff)(0)("GCODE"), "]", dtGCODE3.Select(fff)(0)("CNAME"))
                    v_GCODE3 &= String.Concat(If(v_GCODE3 <> "", ",", ""), vTmp1)
                End If
            End If
        Next
        Return v_GCODE3
    End Function

    Function GetSchDt() As DataTable
        Dim sParms As New Hashtable
        Dim sql As String = ""
        sql &= " SELECT a.EMSEQ ,a.RECRUIT" & vbCrLf
        sql &= " ,a.MBRNAME,a.MBRNMSEQ" & vbCrLf
        sql &= " ,a.UNITNAME" & vbCrLf
        sql &= " ,a.JOBTITLE" & vbCrLf
        sql &= " ,a.DEGREE" & vbCrLf
        sql &= " ,a.CERTIFICAT" & vbCrLf
        sql &= " ,a.SPECIALTY" & vbCrLf
        sql &= " ,a.PHONE,a.PHONE2" & vbCrLf
        sql &= " ,a.CELLPHONE,a.CELLPHONE2" & vbCrLf
        sql &= " ,a.CONFAX" & vbCrLf
        sql &= " ,a.EMAIL,a.EMAIL2" & vbCrLf
        'sql &= " ,a.MZIPCODE,a.MZIPCODE6W" & vbCrLf
        sql &= " ,a.MADDRESS,a.MADDRESS2" & vbCrLf
        sql &= " ,a.RMNOTE1" & vbCrLf
        sql &= " ,a.SERVUNIT1" & vbCrLf
        sql &= " ,a.SERVTIME1" & vbCrLf
        sql &= " ,a.JOBTITLE1" & vbCrLf
        sql &= " ,a.SERVUNIT2" & vbCrLf
        sql &= " ,a.SERVTIME2" & vbCrLf
        sql &= " ,a.JOBTITLE2" & vbCrLf
        sql &= " ,a.SERVUNIT3" & vbCrLf
        sql &= " ,a.SERVTIME3" & vbCrLf
        sql &= " ,a.JOBTITLE3" & vbCrLf
        sql &= " ,a.PUSHDISTID" & vbCrLf
        sql &= " ,a.PUSHREASON" & vbCrLf
        sql &= " ,a.RUNTRAIN" & vbCrLf
        sql &= " ,a.TRAINDISTID" & vbCrLf
        sql &= " ,a.ADDYEARS" & vbCrLf
        sql &= " ,dbo.FN_GET_ROC_YEAR(a.ADDYEARS) ADDYEARS_ROC" & vbCrLf
        sql &= " ,a.STOPUSE" & vbCrLf

        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.CREATEDATE" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " FROM dbo.OA_EXAMINER a" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        '遴聘類別
        Dim v_SCH_RECRUIT As String = TIMS.GetListValue(SCH_RECRUIT)
        Dim v_SCH_PUSHDISTID As String = TIMS.GetCblValue(SCH_PUSHDISTID)
        SCH_JOBTITLE.Text = TIMS.ClearSQM(SCH_JOBTITLE.Text)
        Dim lkSCH_JOBTITLE As String = If(SCH_JOBTITLE.Text <> "", String.Format("%{0}%", SCH_JOBTITLE.Text), "")
        SCH_MBRNAME.Text = TIMS.ClearSQM(SCH_MBRNAME.Text)
        SCH_UNITNAME.Text = TIMS.ClearSQM(SCH_UNITNAME.Text)
        Dim lkSCH_MBRNAME As String = If(SCH_MBRNAME.Text <> "", String.Format("%{0}%", SCH_MBRNAME.Text), "")
        Dim lkSCH_UNITNAME As String = If(SCH_UNITNAME.Text <> "", String.Format("%{0}%", SCH_UNITNAME.Text), "")

        '審查課程職類
        Dim v_SCH_cblGOVCODE3 As String = TIMS.GetCblValue(SCH_cblGOVCODE3) '審查課程職類
        If v_SCH_cblGOVCODE3 <> "" Then v_SCH_cblGOVCODE3 = TIMS.CombiSQLIN(v_SCH_cblGOVCODE3)
        If v_SCH_cblGOVCODE3 <> "" Then
            sql &= String.Format(" AND EXISTS (SELECT 1 FROM dbo.OA_EXAMINERJOB x WHERE x.EMSEQ=a.EMSEQ AND x.GCODE IN ({0}))", v_SCH_cblGOVCODE3)
        End If

        '遴聘類別
        If v_SCH_RECRUIT <> "" Then sParms.Add("RECRUIT", v_SCH_RECRUIT)
        If v_SCH_RECRUIT <> "" Then sql &= " AND a.RECRUIT=@RECRUIT" & vbCrLf
        '職稱
        If lkSCH_JOBTITLE <> "" Then sParms.Add("JOBTITLE", lkSCH_JOBTITLE)
        If lkSCH_JOBTITLE <> "" Then sql &= " AND a.JOBTITLE like @JOBTITLE" & vbCrLf
        '審查委員姓名
        If lkSCH_MBRNAME <> "" Then sParms.Add("MBRNAME", lkSCH_MBRNAME)
        If lkSCH_MBRNAME <> "" Then sql &= " AND a.MBRNAME like @MBRNAME" & vbCrLf
        '現職服務機構
        If lkSCH_UNITNAME <> "" Then sParms.Add("UNITNAME", lkSCH_UNITNAME)
        If lkSCH_UNITNAME <> "" Then sql &= " AND a.UNITNAME like @UNITNAME" & vbCrLf
        '推薦分署
        If v_SCH_PUSHDISTID <> "" Then
            sql &= " AND (1!=1"
            For Each ss1 As String In v_SCH_PUSHDISTID.Split(",")
                If ss1 <> "" Then sql &= " or a.PUSHDISTID like '" & String.Format("%{0}%", ss1) & "'"
            Next
            sql &= " )" & vbCrLf
        End If
        '含已停用資料
        sql &= If(CHECK_STOPUSE.Checked, "", " AND a.STOPUSE IS NULL") & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sParms)
        Return dt
    End Function

    Sub sSearch1()
        msg1.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False

        Dim dt As DataTable = GetSchDt()
        If dt Is Nothing Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

        msg1.Text = ""
        DataGrid1.Visible = True
        PageControler1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Sub SHOW_PANEL(ByVal iType As Integer)
        panelEdit.Visible = False
        panelSch.Visible = True
        If iType = 0 Then Return

        panelEdit.Visible = True
        panelSch.Visible = False
        If iType = 1 Then Return
    End Sub

#Region "CHANGE-1"
    Function Get_RECRUIT_NN(ByRef Value As String) As String
        Dim rst As String = ""
        If Value = "" Then Return rst
        Select Case Value
            Case "A"
                rst = "A-產業界"
            Case "B"
                rst = "B-學術界"
            Case "C"
                rst = "C-勞工團體代表"
        End Select
        Return rst
    End Function

    'Function Get_PUSHDISTID_NN(ByRef sValue As String) As String
    '    Dim rst As String = ""
    '    'If oValue Is Nothing Then Return rst
    '    'Dim sValue As String = Convert.ToString(oValue)
    '    If sValue = "" Then Return rst

    '    Dim ff3 As String = ""
    '    If sValue.IndexOf(",") = -1 Then
    '        ff3 = String.Format("DISTID='{0}'", sValue)
    '        If dtDist.Select(ff3).Length > 0 Then rst = dtDist.Select(ff3)(0)("NAME")
    '        Return rst
    '    End If
    '    'Dim s_TMP1 As String = ""
    '    For Each V1 As String In sValue.Split(",")
    '        ff3 = String.Format("DISTID='{0}'", V1)
    '        If dtDist.Select(ff3).Length > 0 Then
    '            If rst <> "" Then rst &= ","
    '            rst &= dtDist.Select(ff3)(0)("NAME")
    '        End If
    '    Next
    '    Return rst
    'End Function

    '轉換啟用中文
    Function Get_STOPUSECHANG_NN(ByRef vSTOPUSE As String) As String
        Dim rst As String = "Y"
        If vSTOPUSE Is Nothing Then Return rst
        If vSTOPUSE.Equals("Y") Then Return "停用"
        Return rst
    End Function

#End Region

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Call TIMS.OpenDbConn(objconn)
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        If e Is Nothing Then Return
        If e.CommandName Is Nothing Then Return
        If e.CommandName = "" Then Return
        If e.CommandArgument Is Nothing Then Return
        If e.CommandArgument = "" Then Return
        Dim s_CmdArg As String = e.CommandArgument

        Select Case e.CommandName
            Case cst_EDIT1, cst_VIEW1
                Call ClearData1()
                Hid_EMSEQ.Value = TIMS.GetMyValue(s_CmdArg, "EMSEQ")
                Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
                If Hid_EMSEQ.Value = "" Then Return
                Call LoadData1()

            Case cst_STOP1
                Hid_EMSEQ.Value = TIMS.GetMyValue(s_CmdArg, "EMSEQ")
                Hid_STOPUSE.Value = TIMS.GetMyValue(s_CmdArg, "STOPUSE")
                Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
                If Hid_EMSEQ.Value = "" Then Return
                Dim iEMSEQ As Integer = Val(Hid_EMSEQ.Value)
                Dim s_msg2 As String = "已停用"
                If Hid_STOPUSE.Value.Equals("Y") Then s_msg2 = "已啟用"

                UPDATA_STOPUSE(iEMSEQ, If(Hid_STOPUSE.Value.Equals("Y"), "", "Y"))

                Common.MessageBox(Me, s_msg2)
                sSearch1()
                Return
            Case cst_DEL1
                Hid_EMSEQ.Value = TIMS.GetMyValue(s_CmdArg, "EMSEQ")
                Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
                If Hid_EMSEQ.Value = "" Then Return
                Dim iEMSEQ As Integer = Val(Hid_EMSEQ.Value)

                '	若確定刪除系統判斷此委員在「審查委員出席場次」是否有資料?
                'Common.MessageBox(Me, "測試期間暫不提供刪除!!")
                Dim flag_MEETEXAM As Boolean = CHECK_MEETEXAM(iEMSEQ)
                If flag_MEETEXAM Then
                    Common.MessageBox(Me, "使用中，不可刪除!!")
                    Return
                End If

                Call DELETE_EXAMINER(iEMSEQ)
                Dim s_msg2 As String = "資料已刪除！"
                Common.MessageBox(Me, s_msg2)
                sSearch1()
                Return
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTNEDIT1 As Button = e.Item.FindControl("BTNEDIT1")
                Dim BTNSTOP1 As Button = e.Item.FindControl("BTNSTOP1")
                Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1")
                Dim BTNVIEW1 As Button = e.Item.FindControl("BTNVIEW1")

                Dim labRECRUIT_N As Label = e.Item.FindControl("labRECRUIT_N")
                Dim labPUSHDISTID_N As Label = e.Item.FindControl("labPUSHDISTID_N")

                BTNDEL1.Attributes("onclick") = "javascript:return confirm('此動作會刪除審查委員資料，是否確定刪除?');"

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                labRECRUIT_N.Text = Get_RECRUIT_NN(Convert.ToString(drv("RECRUIT")))
                'labPUSHDISTID_N.Text = Get_PUSHDISTID_NN(Convert.ToString(drv("PUSHDISTID")))
                labPUSHDISTID_N.Text = TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(drv("PUSHDISTID")))

                Dim s_CmdArg As String = ""
                TIMS.SetMyValue(s_CmdArg, "EMSEQ", drv("EMSEQ"))
                TIMS.SetMyValue(s_CmdArg, "STOPUSE", Convert.ToString(drv("STOPUSE")))
                BTNEDIT1.CommandArgument = s_CmdArg
                BTNSTOP1.CommandArgument = s_CmdArg
                BTNDEL1.CommandArgument = s_CmdArg
                BTNVIEW1.CommandArgument = s_CmdArg

                BTNEDIT1.Visible = If(sm.UserInfo.LID <> 0, False, True)
                BTNSTOP1.Visible = If(sm.UserInfo.LID <> 0, False, True)
                BTNDEL1.Visible = If(sm.UserInfo.LID <> 0, False, True)
                BTNVIEW1.Visible = If(sm.UserInfo.LID = 0, False, True)

                '動作：可停用/啟用該筆審查委員資料
                '顯示： 當資料為啟用時， 顯示停用按鈕 / 當資料為停用時， 顯示啟用按鈕
                BTNSTOP1.Text = If(Convert.ToString(drv("STOPUSE")).Equals("Y"), "啟用", "停用")
                If Not Convert.ToString(drv("STOPUSE")).Equals("Y") Then
                    BTNSTOP1.Attributes("onclick") = "javascript:return confirm('此動作會停用審查委員資料，是否確定停用?');"
                End If
                If Convert.ToString(drv("STOPUSE")).Equals("Y") Then TIMS.Tooltip(BTNSTOP1, "已停用", True)
        End Select
    End Sub

    '檢核使用狀況 true:使用中 false:無使用
    Function CHECK_MEETEXAM(ByRef iEMSEQ As Integer) As Boolean
        Dim rst As Boolean = False
        Dim parms As Hashtable = New Hashtable
        parms.Add("EMSEQ", iEMSEQ)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'X' FROM OA_MEETEXAM WHERE EMSEQ=@EMSEQ" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt1.Rows.Count = 0 Then Return rst
        Return True
    End Function

    '刪除
    Sub DELETE_EXAMINER(ByRef iEMSEQ As Integer)
        Dim rst As Integer = 0
        If iEMSEQ = 0 Then Return

        '查詢1筆
        Dim parms As Hashtable = New Hashtable
        parms.Add("EMSEQ", iEMSEQ)
        Dim sql As String = "SELECT 'X' FROM OA_EXAMINER WHERE EMSEQ=@EMSEQ"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt1.Rows.Count <> 1 Then Return

        'Dim parms As Hashtable = New Hashtable
        'Dim sql As String = ""
        '備份存檔
        sql = "" & vbCrLf
        sql &= " UPDATE OA_EXAMINER" & vbCrLf
        sql &= " SET MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE EMSEQ=@EMSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("EMSEQ", iEMSEQ)
        rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)

        sql = "" & vbCrLf
        sql &= " UPDATE OA_EXAMINERJOB" & vbCrLf
        sql &= " SET MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE EMSEQ=@EMSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("EMSEQ", iEMSEQ)
        rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)

        'Dim s_COL As String = "EMSEQ,RECRUIT,MBRNAME,UNITNAME,JOBTITLE,DEGREE,CERTIFICAT,SPECIALTY,PHONE,PHONE2,CELLPHONE,CELLPHONE2,CONFAX,EMAIL,EMAIL2,MADDRESS,MADDRESS2,RMNOTE1,SERVUNIT1,SERVTIME1,JOBTITLE1,SERVUNIT2,SERVTIME2,JOBTITLE2,SERVUNIT3,SERVTIME3,JOBTITLE3,PUSHDISTID,PUSHREASON,RUNTRAIN,TRAINDISTID,ADDYEARS,STOPUSE,RID,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE"
        sql = "SELECT * FROM OA_EXAMINERDEL WHERE 1<>1"
        Dim dtCOL As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim s_COL As String = TIMS.Get_DataTableCOLUMN2(dtCOL)
        '備份存檔
        sql = "" & vbCrLf
        sql &= String.Concat(" INSERT INTO OA_EXAMINERDEL(", s_COL, ")") & vbCrLf
        sql &= String.Concat(" SELECT ", s_COL, " FROM OA_EXAMINER")
        sql &= " WHERE MODIFYACCT=@MODIFYACCT AND EMSEQ=@EMSEQ" & vbCrLf
        parms.Clear()
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("EMSEQ", iEMSEQ)
        rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)

        '刪除
        Dim d_Parms As New Hashtable
        d_Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        d_Parms.Add("EMSEQ", iEMSEQ)
        Dim d_sql As String = " DELETE OA_EXAMINER WHERE MODIFYACCT=@MODIFYACCT AND EMSEQ=@EMSEQ" & vbCrLf
        rst = DbAccess.ExecuteNonQuery(d_sql, objconn, d_Parms)

        '刪除
        d_Parms.Clear()
        d_Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        d_Parms.Add("EMSEQ", iEMSEQ)
        d_sql = "DELETE OA_EXAMINERJOB WHERE MODIFYACCT=@MODIFYACCT AND EMSEQ=@EMSEQ"
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_Parms)
    End Sub

    '停用／啟用設定
    Sub UPDATA_STOPUSE(ByRef iEMSEQ As Integer, ByRef vSTOPUSE As String)
        Dim rst As Integer = 0
        '修改
        Dim parms As Hashtable = New Hashtable
        'Dim iEMSEQ As Integer = Val(Hid_EMSEQ.Value)
        parms.Clear()
        parms.Add("STOPUSE", If(vSTOPUSE <> "", vSTOPUSE, Convert.DBNull))
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("EMSEQ", iEMSEQ)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE OA_EXAMINER" & vbCrLf
        sql &= " SET STOPUSE=@STOPUSE" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE 1=1 AND EMSEQ=@EMSEQ" & vbCrLf

        rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
    End Sub

    '檢查
    Function CheckData1(ByRef s_ERRMSG As String) As Boolean
        Dim rst As Boolean = True
        s_ERRMSG = ""

        Dim v_ddlRECRUIT As String = TIMS.GetListValue(ddlRECRUIT) '遴聘類別

        MBRNAME.Text = TIMS.ClearSQM(MBRNAME.Text) '審查委員姓名
        MBRNMSEQ.Text = TIMS.ClearSQM(MBRNMSEQ.Text) '(同名序號,同名可補序號)
        If MBRNMSEQ.Text <> "" Then MBRNMSEQ.Text = TIMS.CINT1(MBRNMSEQ.Text)
        If MBRNMSEQ.Text <> "" AndAlso TIMS.CINT1(MBRNMSEQ.Text) < 0 Then s_ERRMSG &= "同名序號 請輸入大於0的數字" & vbCrLf
        UNITNAME.Text = TIMS.ClearSQM(UNITNAME.Text) '現職服務機構
        JOBTITLE.Text = TIMS.ClearSQM(JOBTITLE.Text) '職稱
        DEGREE.Text = TIMS.ClearSQM(DEGREE.Text) '學歷

        CERTIFICAT.Text = TIMS.ClearSQM(CERTIFICAT.Text) '證照
        SPECIALTY.Text = TIMS.ClearSQM(SPECIALTY.Text) '專業背景
        PHONE.Text = TIMS.ClearSQM(PHONE.Text) '連絡電話
        PHONE2.Text = TIMS.ClearSQM(PHONE2.Text) '連絡電話2
        CELLPHONE.Text = TIMS.ClearSQM(CELLPHONE.Text) '手機
        CELLPHONE2.Text = TIMS.ClearSQM(CELLPHONE2.Text) '手機2
        CONFAX.Text = TIMS.ClearSQM(CONFAX.Text) '傳真
        EMAIL.Text = TIMS.ChangeEmail(TIMS.ClearSQM(EMAIL.Text)) '電子郵件
        EMAIL2.Text = TIMS.ChangeEmail(TIMS.ClearSQM(EMAIL2.Text)) '電子郵件2

        'MZIPCODE.Value = TIMS.ClearSQM(MZIPCODE.Value) '地址
        'MZIPCODE6W.Value = TIMS.ClearSQM(MZIPCODE6W.Value) '地址
        MADDRESS.Text = TIMS.ClearSQM(MADDRESS.Text) '地址
        MADDRESS2.Text = TIMS.ClearSQM(MADDRESS2.Text) '地址2
        RMNOTE1.Text = TIMS.ClearSQM(RMNOTE1.Text) '備註

        SERVUNIT1.Text = TIMS.ClearSQM(SERVUNIT1.Text) '服務單位1
        SERVTIME1.Text = TIMS.ClearSQM(SERVTIME1.Text) '服務時間1
        JOBTITLE1.Text = TIMS.ClearSQM(JOBTITLE1.Text) '職稱1

        SERVUNIT2.Text = TIMS.ClearSQM(SERVUNIT2.Text) '服務單位2
        SERVTIME2.Text = TIMS.ClearSQM(SERVTIME2.Text) '服務時間2
        JOBTITLE2.Text = TIMS.ClearSQM(JOBTITLE2.Text) '職稱2

        SERVUNIT3.Text = TIMS.ClearSQM(SERVUNIT3.Text) '服務單位3
        SERVTIME3.Text = TIMS.ClearSQM(SERVTIME3.Text) '服務時間3
        JOBTITLE3.Text = TIMS.ClearSQM(JOBTITLE3.Text) '職稱3

        Dim v_cbPUSHDISTID As String = TIMS.GetCblValue(cbPUSHDISTID) '推薦分署
        PUSHREASON.Text = TIMS.ClearSQM(PUSHREASON.Text) '推薦理由
        Dim v_rblRUNTRAIN As String = TIMS.GetListValue(rblRUNTRAIN) '是否辦訓
        Dim v_cbTRAINDISTID As String = TIMS.GetCblValue(cbTRAINDISTID) '辦訓轄區
        Dim v_cblGOVCODE3 As String = TIMS.GetCblValue(cblGOVCODE3) '審查課程職類
        Dim v_ddlADDYEARS As String = TIMS.GetListValue(ddlADDYEARS) '新增年度
        Dim v_rblSTOPUSE As String = TIMS.GetListValue(rblSTOPUSE) '停用

        If v_ddlRECRUIT = "" Then s_ERRMSG &= "請選擇 遴聘類別" & vbCrLf
        If MBRNAME.Text = "" Then s_ERRMSG &= "請輸入 審查委員姓名" & vbCrLf
        If UNITNAME.Text = "" Then s_ERRMSG &= "請輸入 現職服務機構" & vbCrLf
        If JOBTITLE.Text = "" Then s_ERRMSG &= "請輸入 職稱" & vbCrLf
        If SPECIALTY.Text = "" Then s_ERRMSG &= "請輸入 專業背景" & vbCrLf

        If PHONE.Text = "" Then s_ERRMSG &= "請輸入 連絡電話" & vbCrLf
        'If PHONE2.Text = "" Then s_ERRMSG &= "請輸入 連絡電話2" & vbCrLf
        If PHONE.Text <> "" AndAlso PHONE2.Text <> "" AndAlso PHONE.Text.Equals(PHONE2.Text) Then s_ERRMSG &= "連絡電話 1與2不可相同" & vbCrLf
        If CELLPHONE.Text <> "" AndAlso CELLPHONE2.Text <> "" AndAlso CELLPHONE.Text.Equals(CELLPHONE2.Text) Then s_ERRMSG &= "手機 1與2不可相同" & vbCrLf

        If SERVUNIT1.Text = "" Then s_ERRMSG &= "請輸入 服務單位1" & vbCrLf
        If SERVTIME1.Text = "" Then s_ERRMSG &= "請輸入 服務時間1" & vbCrLf
        If JOBTITLE1.Text = "" Then s_ERRMSG &= "請輸入 職稱1" & vbCrLf
        If v_cbPUSHDISTID = "" Then s_ERRMSG &= "請選擇 推薦分署" & vbCrLf
        '是否辦訓/'辦訓轄區
        If v_rblRUNTRAIN = "Y" AndAlso v_cbTRAINDISTID = "" Then s_ERRMSG &= "是否辦訓 若選擇「是」，請選擇 辦訓轄區(至少一筆)" & vbCrLf
        If v_cblGOVCODE3 = "" Then s_ERRMSG &= "請選擇 審查課程職類" & vbCrLf '審查課程職類
        'If v_ddlADDYEARS = "" Then
        '    s_ERRMSG &= "請選擇 新增年度" & vbCrLf
        'Else
        '    If Val(v_ddlADDYEARS) < (106 + 1911) Then
        '        s_ERRMSG &= "新增年度,最小年度為106年" & vbCrLf
        '    ElseIf Val(v_ddlADDYEARS) > (206 + 1911) Then
        '        s_ERRMSG &= "新增年度,最大年度為206年" & vbCrLf
        '    End If
        'End If
        If v_ddlADDYEARS <> "" Then
            If Val(v_ddlADDYEARS) < (104 + 1911) Then
                s_ERRMSG &= "新增年度,最小年度為104年" & vbCrLf
            ElseIf Val(v_ddlADDYEARS) > (206 + 1911) Then
                s_ERRMSG &= "新增年度,最大年度為206年" & vbCrLf
            End If
        End If

        'If PHONE.Text <> "" AndAlso Not TIMS.CheckPhone(PHONE.Text) Then s_ERRMSG &= "請檢查 連絡電話 電話格式有誤" & vbCrLf
        'If CELLPHONE.Text <> "" AndAlso Not TIMS.CheckPhone(CELLPHONE.Text) Then s_ERRMSG &= "請檢查 手機 電話格式有誤" & vbCrLf
        'If CONFAX.Text <> "" AndAlso Not TIMS.CheckPhone(CONFAX.Text) Then s_ERRMSG &= "請檢查 傳真 電話格式有誤" & vbCrLf

        If EMAIL.Text <> "" AndAlso Not TIMS.CheckEmail(EMAIL.Text) Then s_ERRMSG &= "請檢查 電子郵件 EMAIL格式有誤" & vbCrLf
        If EMAIL2.Text <> "" AndAlso Not TIMS.CheckEmail(EMAIL2.Text) Then s_ERRMSG &= "請檢查 電子郵件2 EMAIL格式有誤" & vbCrLf
        If EMAIL.Text <> "" AndAlso EMAIL2.Text <> "" AndAlso EMAIL.Text.Equals(EMAIL2.Text) Then s_ERRMSG &= "電子郵件 1與2不可相同" & vbCrLf
        If MADDRESS.Text <> "" AndAlso MADDRESS2.Text <> "" AndAlso MADDRESS.Text.Equals(MADDRESS2.Text) Then s_ERRMSG &= "地址 1與2不可相同" & vbCrLf

        'If MZIPCODE6W.Value <> "" Then
        '    If Not TIMS.IsNumeric1(MZIPCODE6W.Value) Then s_ERRMSG &= "地址郵遞區號後2碼必須為數字，且不得輸入 00" & vbCrLf
        '    If Val(MZIPCODE6W.Value) < 1 OrElse Val(MZIPCODE6W.Value) > 99 Then s_ERRMSG &= "地址郵遞區號後2碼必須為數字，得輸入 01~99" & vbCrLf
        '    If MZIPCODE6W.Value.Length <> 2 Then s_ERRMSG &= "地址郵遞區號後2碼長度必須為 2 碼(例 01 或 99)" & vbCrLf
        'End If

        ',RECRUIT NVARCHAR(20) -- '遴聘類別
        ',MBRNAME NVARCHAR(50) -- '審查委員姓名
        ',UNITNAME NVARCHAR(50) --'現職服務機構
        ',JOBTITLE NVARCHAR(100) --'職稱
        ',DEGREE NVARCHAR(100) --'學歷
        ',CERTIFICAT NVARCHAR(300)  --'證照
        ',SPECIALTY NVARCHAR(300) --'專業背景
        ',PHONE NVARCHAR(20) --'連絡電話
        ',CELLPHONE NVARCHAR(20) --'手機
        ',CONFAX NVARCHAR(20) --'傳真
        ',EMAIL NVARCHAR(70) --'電子郵件
        ',MZIPCODE VARCHAR(3) --'地址
        ',MZIPCODE6W VARCHAR(2) --'地址
        ',MADDRESS NVARCHAR(200) --'地址
        ',RMNOTE1 NVARCHAR(300) --'備註
        ',SERVUNIT1 NVARCHAR(100) --'服務單位1
        ',SERVTIME1 NVARCHAR(100) --'服務時間1
        ',JOBTITLE1 NVARCHAR(100) --'職稱1
        ',SERVUNIT2 NVARCHAR(100) --'服務單位2
        ',SERVTIME2 NVARCHAR(100) --'服務時間2
        ',JOBTITLE2 NVARCHAR(100) --'職稱2
        ',SERVUNIT3 NVARCHAR(100) --'服務單位3
        ',SERVTIME3 NVARCHAR(100) --'服務時間3
        ',JOBTITLE3 NVARCHAR(100) --'職稱3
        If s_ERRMSG <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>儲存 OA_EXAMINER / OA_EXAMINERJOB</summary>
    Sub SaveData1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return
        End If

        Dim rst As Integer = 0
        Dim flagSaveOK1 As Boolean = False

        'Dim sType As String = cst_ADD1
        'If Hid_EMSEQ.Value <> "" Then sType = cst_UPD1 '修改

        Hid_EMSEQ.Value = TIMS.ClearSQM(Hid_EMSEQ.Value)
        Dim v_ddlRECRUIT As String = TIMS.GetListValue(ddlRECRUIT)
        Dim v_cbPUSHDISTID As String = TIMS.GetCblValue(cbPUSHDISTID)
        Dim v_cbTRAINDISTID As String = TIMS.GetCblValue(cbTRAINDISTID)
        Dim v_rblRUNTRAIN As String = TIMS.GetListValue(rblRUNTRAIN)
        Dim v_ddlADDYEARS As String = TIMS.GetListValue(ddlADDYEARS) '新增年度
        Dim v_rblSTOPUSE As String = TIMS.GetListValue(rblSTOPUSE)
        Dim v_cblGOVCODE3 As String = TIMS.GetCblValue(cblGOVCODE3) '審查課程職類

        Dim iEMSEQ As Integer = 0
        Dim parms As Hashtable = New Hashtable
        If Hid_EMSEQ.Value = "" Then
            '新增
            Dim sql As String = ""
            sql &= " INSERT INTO OA_EXAMINER( EMSEQ ,RECRUIT ,MBRNAME,MBRNMSEQ ,UNITNAME" & vbCrLf
            sql &= " ,JOBTITLE ,DEGREE ,CERTIFICAT ,SPECIALTY ,PHONE ,CELLPHONE ,CONFAX ,EMAIL" & vbCrLf
            sql &= " ,PHONE2 ,CELLPHONE2 ,EMAIL2 ,MADDRESS ,MADDRESS2 ,RMNOTE1" & vbCrLf
            sql &= " ,SERVUNIT1 ,SERVTIME1 ,JOBTITLE1 ,SERVUNIT2 ,SERVTIME2 ,JOBTITLE2 ,SERVUNIT3 ,SERVTIME3 ,JOBTITLE3" & vbCrLf
            sql &= " ,PUSHDISTID ,PUSHREASON ,RUNTRAIN ,TRAINDISTID ,ADDYEARS ,STOPUSE, RID" & vbCrLf
            sql &= " ,CREATEACCT ,CREATEDATE ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
            sql &= " VALUES ( @EMSEQ ,@RECRUIT ,@MBRNAME,@MBRNMSEQ ,@UNITNAME" & vbCrLf
            sql &= " ,@JOBTITLE ,@DEGREE ,@CERTIFICAT ,@SPECIALTY ,@PHONE ,@CELLPHONE ,@CONFAX ,@EMAIL" & vbCrLf
            sql &= " ,@PHONE2 ,@CELLPHONE2 ,@EMAIL2 ,@MADDRESS ,@MADDRESS2 ,@RMNOTE1" & vbCrLf
            sql &= " ,@SERVUNIT1 ,@SERVTIME1 ,@JOBTITLE1 ,@SERVUNIT2 ,@SERVTIME2 ,@JOBTITLE2 ,@SERVUNIT3 ,@SERVTIME3 ,@JOBTITLE3" & vbCrLf
            sql &= " ,@PUSHDISTID ,@PUSHREASON ,@RUNTRAIN ,@TRAINDISTID ,@ADDYEARS ,@STOPUSE, @RID" & vbCrLf
            sql &= " ,@CREATEACCT ,GETDATE() ,@MODIFYACCT ,GETDATE())" & vbCrLf

            iEMSEQ = DbAccess.GetNewId(objconn, "OA_EXAMINER_EMSEQ_SEQ,OA_EXAMINER,EMSEQ")
            parms.Clear()
            parms.Add("EMSEQ", iEMSEQ)
            parms.Add("RECRUIT", v_ddlRECRUIT)
            parms.Add("MBRNAME", If(MBRNAME.Text <> "", MBRNAME.Text, Convert.DBNull))
            parms.Add("MBRNMSEQ", If(MBRNMSEQ.Text <> "", TIMS.CINT1(MBRNMSEQ.Text), Convert.DBNull))
            parms.Add("UNITNAME", If(UNITNAME.Text <> "", UNITNAME.Text, Convert.DBNull))

            parms.Add("JOBTITLE", If(JOBTITLE.Text <> "", JOBTITLE.Text, Convert.DBNull))
            parms.Add("DEGREE", If(DEGREE.Text <> "", DEGREE.Text, Convert.DBNull))
            parms.Add("CERTIFICAT", If(CERTIFICAT.Text <> "", CERTIFICAT.Text, Convert.DBNull))
            parms.Add("SPECIALTY", If(SPECIALTY.Text <> "", SPECIALTY.Text, Convert.DBNull))
            parms.Add("PHONE", If(PHONE.Text <> "", PHONE.Text, Convert.DBNull))
            parms.Add("CELLPHONE", If(CELLPHONE.Text <> "", CELLPHONE.Text, Convert.DBNull))
            parms.Add("CONFAX", If(CONFAX.Text <> "", CONFAX.Text, Convert.DBNull))
            parms.Add("EMAIL", If(EMAIL.Text <> "", EMAIL.Text, Convert.DBNull))

            parms.Add("PHONE2", If(PHONE2.Text <> "", PHONE2.Text, Convert.DBNull))
            parms.Add("CELLPHONE2", If(CELLPHONE2.Text <> "", CELLPHONE2.Text, Convert.DBNull))
            parms.Add("EMAIL2", If(EMAIL2.Text <> "", EMAIL2.Text, Convert.DBNull))
            parms.Add("MADDRESS", If(MADDRESS.Text <> "", MADDRESS.Text, Convert.DBNull))
            parms.Add("MADDRESS2", If(MADDRESS2.Text <> "", MADDRESS2.Text, Convert.DBNull))
            parms.Add("RMNOTE1", If(RMNOTE1.Text <> "", RMNOTE1.Text, Convert.DBNull))

            parms.Add("SERVUNIT1", If(SERVUNIT1.Text <> "", SERVUNIT1.Text, Convert.DBNull))
            parms.Add("SERVTIME1", If(SERVTIME1.Text <> "", SERVTIME1.Text, Convert.DBNull))
            parms.Add("JOBTITLE1", If(JOBTITLE1.Text <> "", JOBTITLE1.Text, Convert.DBNull))
            parms.Add("SERVUNIT2", If(SERVUNIT2.Text <> "", SERVUNIT2.Text, Convert.DBNull))
            parms.Add("SERVTIME2", If(SERVTIME2.Text <> "", SERVTIME2.Text, Convert.DBNull))
            parms.Add("JOBTITLE2", If(JOBTITLE2.Text <> "", JOBTITLE2.Text, Convert.DBNull))
            parms.Add("SERVUNIT3", If(SERVUNIT3.Text <> "", SERVUNIT3.Text, Convert.DBNull))
            parms.Add("SERVTIME3", If(SERVTIME3.Text <> "", SERVTIME3.Text, Convert.DBNull))
            parms.Add("JOBTITLE3", If(JOBTITLE3.Text <> "", JOBTITLE3.Text, Convert.DBNull))

            parms.Add("PUSHDISTID", If(v_cbPUSHDISTID <> "", v_cbPUSHDISTID, Convert.DBNull))
            parms.Add("PUSHREASON", If(PUSHREASON.Text <> "", PUSHREASON.Text, Convert.DBNull))
            parms.Add("RUNTRAIN", If(v_rblRUNTRAIN <> "", v_rblRUNTRAIN, Convert.DBNull))
            parms.Add("TRAINDISTID", If(v_cbTRAINDISTID <> "", v_cbTRAINDISTID, Convert.DBNull))
            parms.Add("ADDYEARS", If(v_ddlADDYEARS <> "", v_ddlADDYEARS, Convert.DBNull))
            parms.Add("STOPUSE", If(v_rblSTOPUSE <> "", v_rblSTOPUSE, Convert.DBNull))
            parms.Add("RID", sm.UserInfo.RID)

            parms.Add("CREATEACCT", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)

            gstr_SQL_1 = sql
            gstr_ROWVAL_1 = TIMS.GetMyValue5(parms)
            rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
            If rst > 0 Then flagSaveOK1 = True
        Else
            '修改
            Dim sql As String = ""
            sql &= " UPDATE OA_EXAMINER" & vbCrLf
            sql &= " SET RECRUIT=@RECRUIT,MBRNAME=@MBRNAME,MBRNMSEQ=@MBRNMSEQ,UNITNAME=@UNITNAME" & vbCrLf
            sql &= " ,JOBTITLE=@JOBTITLE,DEGREE=@DEGREE,CERTIFICAT=@CERTIFICAT,SPECIALTY=@SPECIALTY" & vbCrLf
            sql &= " ,PHONE=@PHONE,CELLPHONE=@CELLPHONE,CONFAX=@CONFAX,EMAIL=@EMAIL ,PHONE2=@PHONE2,CELLPHONE2=@CELLPHONE2,EMAIL2=@EMAIL2" & vbCrLf
            'sql &= " ,MZIPCODE=@MZIPCODE,MZIPCODE6W=@MZIPCODE6W,MADDRESS=@MADDRESS,MADDRESS2=@MADDRESS2,RMNOTE1=@RMNOTE1" & vbCrLf
            sql &= " ,MADDRESS=@MADDRESS,MADDRESS2=@MADDRESS2,RMNOTE1=@RMNOTE1" & vbCrLf
            sql &= " ,SERVUNIT1=@SERVUNIT1,SERVTIME1=@SERVTIME1,JOBTITLE1=@JOBTITLE1" & vbCrLf
            sql &= " ,SERVUNIT2=@SERVUNIT2,SERVTIME2=@SERVTIME2,JOBTITLE2=@JOBTITLE2" & vbCrLf
            sql &= " ,SERVUNIT3=@SERVUNIT3,SERVTIME3=@SERVTIME3,JOBTITLE3=@JOBTITLE3" & vbCrLf
            sql &= " ,PUSHDISTID=@PUSHDISTID,PUSHREASON=@PUSHREASON,RUNTRAIN=@RUNTRAIN" & vbCrLf
            sql &= " ,TRAINDISTID=@TRAINDISTID,ADDYEARS=@ADDYEARS,STOPUSE=@STOPUSE" & vbCrLf
            sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            sql &= " WHERE EMSEQ=@EMSEQ" & vbCrLf

            iEMSEQ = Val(Hid_EMSEQ.Value)
            parms.Clear()
            parms.Add("RECRUIT", v_ddlRECRUIT)
            parms.Add("MBRNAME", If(MBRNAME.Text <> "", MBRNAME.Text, Convert.DBNull))
            parms.Add("MBRNMSEQ", If(MBRNMSEQ.Text <> "", TIMS.CINT1(MBRNMSEQ.Text), Convert.DBNull))
            parms.Add("UNITNAME", If(UNITNAME.Text <> "", UNITNAME.Text, Convert.DBNull))

            parms.Add("JOBTITLE", If(JOBTITLE.Text <> "", JOBTITLE.Text, Convert.DBNull))
            parms.Add("DEGREE", If(DEGREE.Text <> "", DEGREE.Text, Convert.DBNull))
            parms.Add("CERTIFICAT", If(CERTIFICAT.Text <> "", CERTIFICAT.Text, Convert.DBNull))
            parms.Add("SPECIALTY", If(SPECIALTY.Text <> "", SPECIALTY.Text, Convert.DBNull))

            parms.Add("PHONE", If(PHONE.Text <> "", PHONE.Text, Convert.DBNull))
            parms.Add("CELLPHONE", If(CELLPHONE.Text <> "", CELLPHONE.Text, Convert.DBNull))
            parms.Add("CONFAX", If(CONFAX.Text <> "", CONFAX.Text, Convert.DBNull))
            parms.Add("EMAIL", If(EMAIL.Text <> "", EMAIL.Text, Convert.DBNull))
            parms.Add("PHONE2", If(PHONE2.Text <> "", PHONE2.Text, Convert.DBNull))
            parms.Add("CELLPHONE2", If(CELLPHONE2.Text <> "", CELLPHONE2.Text, Convert.DBNull))
            parms.Add("EMAIL2", If(EMAIL2.Text <> "", EMAIL2.Text, Convert.DBNull))

            parms.Add("MADDRESS", If(MADDRESS.Text <> "", MADDRESS.Text, Convert.DBNull))
            parms.Add("MADDRESS2", If(MADDRESS2.Text <> "", MADDRESS2.Text, Convert.DBNull))
            parms.Add("RMNOTE1", If(RMNOTE1.Text <> "", RMNOTE1.Text, Convert.DBNull))

            parms.Add("SERVUNIT1", If(SERVUNIT1.Text <> "", SERVUNIT1.Text, Convert.DBNull))
            parms.Add("SERVTIME1", If(SERVTIME1.Text <> "", SERVTIME1.Text, Convert.DBNull))
            parms.Add("JOBTITLE1", If(JOBTITLE1.Text <> "", JOBTITLE1.Text, Convert.DBNull))
            parms.Add("SERVUNIT2", If(SERVUNIT2.Text <> "", SERVUNIT2.Text, Convert.DBNull))
            parms.Add("SERVTIME2", If(SERVTIME2.Text <> "", SERVTIME2.Text, Convert.DBNull))
            parms.Add("JOBTITLE2", If(JOBTITLE2.Text <> "", JOBTITLE2.Text, Convert.DBNull))
            parms.Add("SERVUNIT3", If(SERVUNIT3.Text <> "", SERVUNIT3.Text, Convert.DBNull))
            parms.Add("SERVTIME3", If(SERVTIME3.Text <> "", SERVTIME3.Text, Convert.DBNull))
            parms.Add("JOBTITLE3", If(JOBTITLE3.Text <> "", JOBTITLE3.Text, Convert.DBNull))

            parms.Add("PUSHDISTID", If(v_cbPUSHDISTID <> "", v_cbPUSHDISTID, Convert.DBNull))
            parms.Add("PUSHREASON", If(PUSHREASON.Text <> "", PUSHREASON.Text, Convert.DBNull))
            parms.Add("RUNTRAIN", If(v_rblRUNTRAIN <> "", v_rblRUNTRAIN, Convert.DBNull))

            parms.Add("TRAINDISTID", If(v_cbTRAINDISTID <> "", v_cbTRAINDISTID, Convert.DBNull))
            parms.Add("ADDYEARS", If(v_ddlADDYEARS <> "", v_ddlADDYEARS, Convert.DBNull))
            parms.Add("STOPUSE", If(v_rblSTOPUSE <> "", v_rblSTOPUSE, Convert.DBNull))

            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("EMSEQ", iEMSEQ)

            gstr_SQL_1 = sql
            gstr_ROWVAL_1 = TIMS.GetMyValue5(parms)
            rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
            If rst > 0 Then flagSaveOK1 = True
        End If

        If iEMSEQ > 0 AndAlso flagSaveOK1 Then
            Call SaveData1_GCODE(iEMSEQ, v_cblGOVCODE3)
        ElseIf Not flagSaveOK1 Then '儲存-失敗
            Common.MessageBox(Me, "儲存失敗!")
            Exit Sub
        End If

        SHOW_PANEL(0)
        '儲存成功 'Hid_EMSEQ.Value = ""
        Call ClearData1()
        Common.MessageBox(Me, "儲存成功!")
        Call sSearch1()
    End Sub

    ''' <summary>SAVE OA_EXAMINERJOB</summary>
    ''' <param name="iEMSEQ"></param>
    ''' <param name="v_cblGOVCODE3"></param>
    Sub SaveData1_GCODE(ByRef iEMSEQ As Integer, ByRef v_cblGOVCODE3 As String)
        Dim d_parms As Hashtable = New Hashtable
        d_parms.Clear()
        d_parms.Add("EMSEQ", iEMSEQ)
        Dim d_sql As String = "DELETE OA_EXAMINERJOB WHERE EMSEQ=@EMSEQ"
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)

        Dim s_parms As Hashtable = New Hashtable
        Dim s_sql As String = "SELECT 1 FROM OA_EXAMINERJOB WHERE EMSEQ=@EMSEQ AND GCODE=@GCODE"

        Dim i_parms As Hashtable = New Hashtable
        Dim i_sql As String = ""
        i_sql &= " INSERT INTO OA_EXAMINERJOB(EMSEQ,GCODE,MODIFYACCT,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@EMSEQ ,@GCODE,@MODIFYACCT,GETDATE())" & vbCrLf

        For Each v_GCODE As String In v_cblGOVCODE3.Split(",")
            'Dim iMTSEQ As Integer xx
            'Dim v_GCODE As String = Convert.ToString(dr1("GCODE"))
            v_GCODE = TIMS.ClearSQM(v_GCODE)
            If v_GCODE <> "" Then
                s_parms.Clear()
                s_parms.Add("EMSEQ", iEMSEQ)
                s_parms.Add("GCODE", v_GCODE)
                Dim s_dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, s_parms)
                If s_dt.Rows.Count = 0 Then
                    'Dim i_parms As New Hashtable
                    i_parms.Clear()
                    i_parms.Add("EMSEQ", iEMSEQ)
                    i_parms.Add("GCODE", v_GCODE)
                    i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
                End If
            End If
        Next
    End Sub

    '查詢
    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        sSearch1()
    End Sub

    '新增
    Protected Sub BtnAddnew_Click(sender As Object, e As EventArgs) Handles BtnAddnew.Click
        ClearData1()
        SHOW_PANEL(1)
    End Sub

    '儲存鈕
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Try
            RMNOTE1.Text = TIMS.Get_Substr1(RMNOTE1.Text, 300)
            PUSHREASON.Text = TIMS.Get_Substr1(PUSHREASON.Text, 1000)

            Call SaveData1()
        Catch ex As Exception
            Dim sErrmsg As String = ""
            sErrmsg = String.Concat("儲存資料有誤，請重新操作!!", ex.Message)
            Common.MessageBox(Me, sErrmsg)

            Dim strErrmsg As String = String.Concat("#btnSave1_Click.儲存資料 : ", vbCrLf, ",ex.Message: ", ex.Message, vbCrLf)
            strErrmsg &= String.Concat("btnSave1_Click", ",_sql: ", gstr_SQL_1, vbCrLf, ",_parms: ", gstr_ROWVAL_1, vbCrLf)
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入 'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
            Return
        End Try
    End Sub

    '回上一頁
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        SHOW_PANEL(0)
        Call ClearData1()
        'Common.MessageBox(Me, "儲存成功!")
        'Call sSearch1()
    End Sub

    ''' <summary>匯出功能</summary>
    Sub ExportData1()
        Dim dt As DataTable = GetSchDt()
        Dim rtnPath As String = Request.FilePath
        If dt Is Nothing Then
            Common.MessageBox(Me, "資料庫查詢失敗，請重新查詢!!", rtnPath)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料，請重新查詢!")
            Exit Sub
        End If
        'msg.Text = ""

        '匯出 Response 
        ExpReport1(dt)
    End Sub

    ''' <summary>匯出 Response</summary>
    ''' <param name="dt"></param>
    Sub ExpReport1(ByRef dt As DataTable)
        '匯出表頭名稱
        Dim sFileName1 As String = "審查委員名單" & TIMS.GetDateNo2()

        Const cst_tit1 As String = "遴聘類別,審查委員姓名,現職服務機構,職稱,學歷,證照,專業背景,審查職類,連絡電話,連絡電話2,手機,手機2,傳真,電子郵件,電子郵件2,地址,地址2,備註,服務單位1,服務時間1,職稱1,服務單位2,服務時間2,職稱2,服務單位3,服務時間3,職稱3,推薦分署,推薦理由,是否辦訓,辦訓轄區,新增年度,啟用"
        Const cst_tit2 As String = "RECRUIT,MBRNAME,UNITNAME,JOBTITLE,DEGREE,CERTIFICAT,SPECIALTY,EXAMINERJOB,PHONE,PHONE2,CELLPHONE,CELLPHONE2,CONFAX,EMAIL,EMAIL2,MADDRESS,MADDRESS2,RMNOTE1,SERVUNIT1,SERVTIME1,JOBTITLE1,SERVUNIT2,SERVTIME2,JOBTITLE2,SERVUNIT3,SERVTIME3,JOBTITLE3,PUSHDISTID,PUSHREASON,RUNTRAIN,TRAINDISTID,ADDYEARS_ROC,STOPUSE"
        Dim sta_tit1 As String() = Split(cst_tit1, ",")
        Dim sta_tit2 As String() = Split(cst_tit2, ",")

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New System.Text.StringBuilder
        'Dim strHTML As String = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""
        '建立抬頭'第1行
        ExportStr = "<tr>" & vbCrLf
        For Each str1 As String In sta_tit1
            ExportStr &= String.Format("<td>{0}</td>", str1) & vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        sbHTML.Append(ExportStr)

        For Each dr As DataRow In dt.Rows
            'For Each dr As DataRow In dt.Rows
            '建立資料面
            ExportStr = "<tr>" & vbCrLf
            For Each s_COLN As String In sta_tit2
                'STOPUSE
                If (s_COLN.Equals("RECRUIT")) Then '遴聘類別
                    ExportStr &= String.Format("<td>{0}</td>", Get_RECRUIT_NN(Convert.ToString(dr(s_COLN)))) & vbTab
                ElseIf (s_COLN.Equals("PUSHDISTID")) Then '推薦分署
                    ExportStr &= String.Format("<td>{0}</td>", TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(dr(s_COLN)))) & vbTab
                ElseIf (s_COLN.Equals("TRAINDISTID")) Then '辦訓轄區
                    ExportStr &= String.Format("<td>{0}</td>", TIMS.Get_PUSHDISTID_NN(dtDist, Convert.ToString(dr(s_COLN)))) & vbTab
                ElseIf (s_COLN.Equals("STOPUSE")) Then '啟用
                    ExportStr &= String.Format("<td>{0}</td>", Get_STOPUSECHANG_NN(Convert.ToString(dr(s_COLN)))) & vbTab
                ElseIf s_COLN.Equals("EXAMINERJOB") Then '審查職類代碼
                    'EXAMINERJOB'審查職類代碼
                    ExportStr &= String.Format("<td>{0}</td>", GET_EXAMINERJOB_GCODE3(objconn, Val(dr("EMSEQ")), 2)) & vbTab
                Else
                    ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(s_COLN))) & vbTab
                End If
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出
    Protected Sub BtnExport_Click(sender As Object, e As EventArgs) Handles BtnExport.Click
        SHOW_PANEL(0)
        '匯出功能
        Call ExportData1()
    End Sub

    Public Shared Sub CHK_TXTVALUE(ByRef Reason As String, ByRef aaTXTNAME As String, ByRef aaTXTVALUE As String, ByRef i_varchar_len_Max As Integer)
        If aaTXTVALUE = "" Then Reason &= String.Format("必須填寫 {0}<BR>", aaTXTNAME)
        If aaTXTVALUE <> "" AndAlso aaTXTVALUE.Length > i_varchar_len_Max Then
            Reason &= String.Format("{0},超過欄位儲存長度({1})<BR>", aaTXTNAME, i_varchar_len_Max)
        End If
    End Sub

    ''' <summary> 將「無」字去除</summary>
    ''' <param name="aaTXTVALUE"></param>
    Public Shared Sub CHG_REPLACE_NODATA(ByRef aaTXTVALUE As String)
        If Not aaTXTVALUE.Equals("無") Then Return
        aaTXTVALUE = ""
    End Sub

    Public Shared Sub CHK_DISTIDVALUE(ByRef Reason As String, ByRef aaTXTNAME As String, ByRef aaTXTVALUE As String, ByRef dtDATA1 As DataTable)
        Dim rst1 As String = ""
        If aaTXTVALUE = "" Then Return
        Dim ff3 As String = ""
        If aaTXTVALUE.IndexOf(",") = -1 Then
            aaTXTVALUE = TIMS.AddZero(aaTXTVALUE, 3)
            ff3 = String.Format("DISTID='{0}'", aaTXTVALUE)
            If aaTXTVALUE <> "" AndAlso dtDATA1.Select(ff3).Length = 0 Then
                Reason &= String.Format("{0},查無資訊代碼，請再確認資料({1})<BR>", aaTXTNAME, aaTXTVALUE)
                Return
            End If
            If rst1 <> "" Then rst1 &= ","
            rst1 &= aaTXTVALUE
        Else
            For Each VV1 As String In aaTXTVALUE.Split(",")
                VV1 = TIMS.AddZero(VV1, 3)
                ff3 = String.Format("DISTID='{0}'", VV1)
                If VV1 <> "" AndAlso dtDATA1.Select(ff3).Length = 0 Then
                    Reason &= String.Format("{0},查無資訊代碼，請再確認資料({1})<BR>", aaTXTNAME, aaTXTVALUE)
                    Return
                End If
                If rst1 <> "" Then rst1 &= ","
                rst1 &= VV1
            Next
        End If
        aaTXTVALUE = rst1
    End Sub

    ''' <summary>審查職類代碼-檢查</summary>
    ''' <param name="Reason"></param>
    ''' <param name="aaTXTNAME"></param>
    ''' <param name="aaTXTVALUE"></param>
    ''' <param name="dtDATA1"></param>
    Public Shared Sub CHK_GOVCODE3VALUE(ByRef Reason As String, ByRef aaTXTNAME As String, ByRef aaTXTVALUE As String, ByRef dtDATA1 As DataTable)
        Dim rst1 As String = ""
        If aaTXTVALUE = "" Then Return
        Dim ff3 As String = ""
        If aaTXTVALUE.IndexOf(",") = -1 Then
            aaTXTVALUE = TIMS.AddZero(aaTXTVALUE, 2)
            ff3 = String.Format("GCODE='{0}'", aaTXTVALUE)
            If aaTXTVALUE <> "" AndAlso dtDATA1.Select(ff3).Length = 0 Then
                Reason &= String.Format("{0},查無資訊代碼，請再確認資料({1})<BR>", aaTXTNAME, aaTXTVALUE)
                Return
            End If
            'If rst1 <> "" Then rst1 &= ","
            rst1 = aaTXTVALUE '(只有1個值)
        Else
            For Each VV1 As String In aaTXTVALUE.Split(",")
                VV1 = TIMS.AddZero(VV1, 2)
                ff3 = String.Format("GCODE='{0}'", VV1)
                If VV1 <> "" AndAlso dtDATA1.Select(ff3).Length = 0 Then
                    Reason &= String.Format("{0},查無資訊代碼，請再確認資料({1})<BR>", aaTXTNAME, aaTXTVALUE)
                    Return
                End If
                rst1 &= String.Concat(If(rst1 <> "", ",", ""), VV1) '(多值)
            Next
        End If
        aaTXTVALUE = rst1
    End Sub

    ''' <summary>檢查匯入資料</summary>
    ''' <param name="colArray"></param>
    ''' <returns></returns>
    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        If colArray Is Nothing Then
            Reason += "匯入資料有誤<BR>"
            Return Reason
        End If
        'Dim dr As DataRow
        If colArray.Length <> cst_Max_a_Len Then
            'Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            If colArray.Length > cst_aMBRNAME Then aMBRNAME = colArray(cst_aMBRNAME).ToString
            Return Reason
        End If

        'aEMSEQ = TIMS.ClearSQM(colArray(cst_aEMSEQ))
        aRECRUIT = TIMS.ClearSQM(colArray(cst_aRECRUIT))
        aMBRNAME = TIMS.ClearSQM(colArray(cst_aMBRNAME))
        aUNITNAME = TIMS.ClearSQM(colArray(cst_aUNITNAME))
        aJOBTITLE = TIMS.ClearSQM(colArray(cst_aJOBTITLE))
        aDEGREE = TIMS.ClearSQM(colArray(cst_aDEGREE))
        aCERTIFICAT = TIMS.ClearSQM(colArray(cst_aCERTIFICAT))
        aSPECIALTY = TIMS.ClearSQM(colArray(cst_aSPECIALTY))
        aEXAMINERJOB = TIMS.ClearSQM(colArray(cst_aEXAMINERJOB))

        aPHONE = TIMS.ClearSQM(colArray(cst_aPHONE))
        aPHONE2 = TIMS.ClearSQM(colArray(cst_aPHONE2))
        aCELLPHONE = TIMS.ClearSQM(colArray(cst_aCELLPHONE))
        aCELLPHONE2 = TIMS.ClearSQM(colArray(cst_aCELLPHONE2))
        aCONFAX = TIMS.ClearSQM(colArray(cst_aCONFAX))
        aEMAIL = TIMS.ClearSQM(colArray(cst_aEMAIL))
        aEMAIL2 = TIMS.ClearSQM(colArray(cst_aEMAIL2)) 'aMZIPCODE = TIMS.ClearSQM(colArray(cst_aMZIPCODE)) 'aMZIPCODE6W = TIMS.ClearSQM(colArray(cst_aMZIPCODE6W))
        aMADDRESS = TIMS.ClearSQM(colArray(cst_aMADDRESS))
        aMADDRESS2 = TIMS.ClearSQM(colArray(cst_aMADDRESS2))
        aRMNOTE1 = TIMS.ClearSQM(colArray(cst_aRMNOTE1))

        aSERVUNIT1 = TIMS.ClearSQM(colArray(cst_aSERVUNIT1))
        aSERVTIME1 = TIMS.ClearSQM(colArray(cst_aSERVTIME1))
        aJOBTITLE1 = TIMS.ClearSQM(colArray(cst_aJOBTITLE1))
        aSERVUNIT2 = TIMS.ClearSQM(colArray(cst_aSERVUNIT2))
        aSERVTIME2 = TIMS.ClearSQM(colArray(cst_aSERVTIME2))
        aJOBTITLE2 = TIMS.ClearSQM(colArray(cst_aJOBTITLE2))
        aSERVUNIT3 = TIMS.ClearSQM(colArray(cst_aSERVUNIT3))
        aSERVTIME3 = TIMS.ClearSQM(colArray(cst_aSERVTIME3))
        aJOBTITLE3 = TIMS.ClearSQM(colArray(cst_aJOBTITLE3))

        aPUSHDISTID = TIMS.ClearSQM(colArray(cst_aPUSHDISTID))
        aPUSHREASON = TIMS.ClearSQM(colArray(cst_aPUSHREASON))
        aRUNTRAIN = TIMS.ClearSQM(colArray(cst_aRUNTRAIN))

        aTRAINDISTID = TIMS.ClearSQM(colArray(cst_aTRAINDISTID))
        aADDYEARS_ROC = TIMS.ClearSQM(colArray(cst_aADDYEARS_ROC))
        aSTOPUSE = TIMS.ClearSQM(colArray(cst_aSTOPUSE))
        'aCREATEACCT = TIMS.ClearSQM(colArray(cst_aCREATEACCT))
        'aCREATEDATE = TIMS.ClearSQM(colArray(cst_aCREATEDATE))
        'aMODIFYACCT = TIMS.ClearSQM(colArray(cst_aMODIFYACCT))
        'aMODIFYDATE = TIMS.ClearSQM(colArray(cst_aMODIFYDATE))

        Dim i_varchar_len_Max As Integer = 0
        If aRECRUIT = "" Then Reason &= "必須填寫 遴聘類別代碼<BR>"
        If aRECRUIT <> "" Then
            Select Case aRECRUIT
                Case "A", "B", "C"
                Case Else
                    Reason &= "遴聘類別代碼,請填寫A/B/C<BR>"
            End Select
        End If

        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "審查委員姓名", aMBRNAME, i_varchar_len_Max)
        'i_varchar_len_Max = 3 : CHK_TXTVALUE(Reason, "同名序號", aMBRNMSEQ, i_varchar_len_Max)
        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "現職服務機構", aUNITNAME, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "職稱", aJOBTITLE, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "學歷", aDEGREE, i_varchar_len_Max)
        i_varchar_len_Max = 300 : CHK_TXTVALUE(Reason, "證照", aCERTIFICAT, i_varchar_len_Max)
        i_varchar_len_Max = 300 : CHK_TXTVALUE(Reason, "專業背景", aSPECIALTY, i_varchar_len_Max)
        i_varchar_len_Max = 60 : CHK_TXTVALUE(Reason, "審查職類代碼", aEXAMINERJOB, i_varchar_len_Max)

        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務單位1", aSERVUNIT1, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務時間1", aSERVTIME1, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "職稱1", aJOBTITLE1, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務單位2", aSERVUNIT2, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務時間2", aSERVTIME2, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "職稱2", aJOBTITLE2, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務單位3", aSERVUNIT3, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "服務時間3", aSERVTIME3, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "職稱3", aJOBTITLE3, i_varchar_len_Max)

        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "連絡電話", aPHONE, i_varchar_len_Max)
        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "連絡電話2", aPHONE2, i_varchar_len_Max)
        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "傳真", aCONFAX, i_varchar_len_Max)
        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "手機", aCELLPHONE, i_varchar_len_Max)
        i_varchar_len_Max = 50 : CHK_TXTVALUE(Reason, "手機2", aCELLPHONE2, i_varchar_len_Max)
        i_varchar_len_Max = 70 : CHK_TXTVALUE(Reason, "電子郵件", aEMAIL, i_varchar_len_Max)
        i_varchar_len_Max = 70 : CHK_TXTVALUE(Reason, "電子郵件2", aEMAIL2, i_varchar_len_Max)

        If aEMAIL <> "" AndAlso aEMAIL <> "無" AndAlso Not TIMS.CheckEmail(aEMAIL) Then Reason &= "電子信箱 EMail格式錯誤<BR>"
        If aEMAIL2 <> "" AndAlso aEMAIL2 <> "無" AndAlso Not TIMS.CheckEmail(aEMAIL2) Then Reason &= "電子信箱2 EMail格式錯誤<BR>"

        'aMZIPCODE = TIMS.ChangeIDNO(aMZIPCODE)
        'aMZIPCODE6W = TIMS.ChangeIDNO(aMZIPCODE6W)
        'i_varchar_len_Max = 3 : CHK_TXTVALUE(Reason, "地址-郵遞區號前3碼", aMZIPCODE, i_varchar_len_Max)
        'i_varchar_len_Max = 2 : CHK_TXTVALUE(Reason, "地址-郵遞區號後2碼", aMZIPCODE6W, i_varchar_len_Max)
        'Dim tmpErrmsg As String = ""
        'If aMZIPCODE6W <> "" AndAlso aMZIPCODE6W <> "無" Then
        '    tmpErrmsg = ""
        '    Call TIMS.CheckZipCODEB3(aMZIPCODE6W, "地址-郵遞區號後2碼", False, tmpErrmsg)
        '    If tmpErrmsg <> "" Then Reason &= Replace(tmpErrmsg, vbCrLf, "<BR>")
        'End If

        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "地址", aMADDRESS, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "地址2", aMADDRESS2, i_varchar_len_Max)
        i_varchar_len_Max = 300 : CHK_TXTVALUE(Reason, "備註", aRMNOTE1, i_varchar_len_Max)

        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "推薦分署代碼", aPUSHDISTID, i_varchar_len_Max)
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "推薦理由", aPUSHREASON, i_varchar_len_Max)
        'i_varchar_len_Max = 10 : CHK_TXTVALUE(Reason, "是否辦訓", aRUNTRAIN, i_varchar_len_Max)
        If aRUNTRAIN <> "" Then
            Select Case aRUNTRAIN
                Case "Y", "N"
                Case Else
                    Reason &= "是否辦訓,請填寫 Y 或 N <BR>"
            End Select
        End If
        i_varchar_len_Max = 100 : CHK_TXTVALUE(Reason, "辦訓轄區代碼", aTRAINDISTID, i_varchar_len_Max)
        i_varchar_len_Max = 3 : CHK_TXTVALUE(Reason, "新增年度", aADDYEARS_ROC, i_varchar_len_Max)
        aADDYEARS_AD = ""
        'If aADDYEARS_ROC = "" Then
        '    Reason &= "新增年度,為必填不可為空白<BR>"
        'ElseIf Not TIMS.IsNumeric2(aADDYEARS_ROC) Then
        '    Reason &= "新增年度,數字格式有誤 <BR>"
        'Else
        '    If Val(aADDYEARS) < 104 Then
        '        Reason &= "新增年度,最小年度為104年 <BR>"
        '    ElseIf Val(aADDYEARS) > 206 Then
        '        Reason &= "新增年度,最大年度為206年 <BR>"
        '    End If
        'End If
        If aADDYEARS_ROC <> "" AndAlso aADDYEARS_ROC <> "無" Then
            If Not TIMS.IsNumeric2(aADDYEARS_ROC) Then
                Reason &= "新增年度,數字格式有誤 <BR>"
            Else
                If Val(aADDYEARS_ROC) < 104 Then
                    Reason &= "新增年度,最小年度為104年 <BR>"
                ElseIf Val(aADDYEARS_ROC) > 206 Then
                    Reason &= "新增年度,最大年度為206年 <BR>"
                End If
            End If
        End If
        If Reason = "" Then aADDYEARS_AD = (Val(aADDYEARS_ROC) + 1911)

        i_varchar_len_Max = 3 : CHK_TXTVALUE(Reason, "停用", aSTOPUSE, i_varchar_len_Max)
        If aSTOPUSE <> "" Then
            Select Case aSTOPUSE
                Case "Y"
                Case "N"
                    aSTOPUSE = ""
                Case Else
                    Reason &= "停用,請填寫 Y 或 N <BR>"
            End Select
        End If
        If Reason <> "" Then
            Reason &= "欄位不可為空白，若無資料，請填寫「無」字！<BR>"
            Return Reason
        End If

        If aEXAMINERJOB <> "" AndAlso aEXAMINERJOB <> "無" Then CHK_GOVCODE3VALUE(Reason, "審查職類代碼", aEXAMINERJOB, dtGCODE3)
        If aPUSHDISTID <> "" AndAlso aPUSHDISTID <> "無" Then CHK_DISTIDVALUE(Reason, "推薦分署代碼", aPUSHDISTID, dtDist)
        If aTRAINDISTID <> "" AndAlso aTRAINDISTID <> "無" Then CHK_DISTIDVALUE(Reason, "辦訓轄區代碼", aTRAINDISTID, dtDist)
        If Reason <> "" Then Return Reason

        'replace 無
        CHG_REPLACE_NODATA(aRECRUIT) ' As String = "" '遴聘類別代碼 1
        CHG_REPLACE_NODATA(aMBRNAME) ' As String = "" '姓名 審查委員姓名 50
        CHG_REPLACE_NODATA(aUNITNAME) ' As String = "" '現職服務機構 50
        CHG_REPLACE_NODATA(aJOBTITLE) ' As String = "" '職稱 100
        CHG_REPLACE_NODATA(aDEGREE) ' As String = "" '學歷 100
        CHG_REPLACE_NODATA(aCERTIFICAT) ' As String = "" '證照 300
        CHG_REPLACE_NODATA(aSPECIALTY) ' As String = "" '專業背景 300
        CHG_REPLACE_NODATA(aEXAMINERJOB) ' As String = "" '審查職類代碼 60

        CHG_REPLACE_NODATA(aSERVUNIT1) ' As String = "" '服務單位1 100
        CHG_REPLACE_NODATA(aSERVTIME1) ' As String = "" '服務時間1 100
        CHG_REPLACE_NODATA(aJOBTITLE1) ' As String = "" '職稱1 100
        CHG_REPLACE_NODATA(aSERVUNIT2) ' As String = "" '服務單位2 100
        CHG_REPLACE_NODATA(aSERVTIME2) ' As String = "" '服務時間2 100
        CHG_REPLACE_NODATA(aJOBTITLE2) ' As String = "" '職稱2 100
        CHG_REPLACE_NODATA(aSERVUNIT3) ' As String = "" '服務單位3 100 
        CHG_REPLACE_NODATA(aSERVTIME3) ' As String = "" '服務時間3 100
        CHG_REPLACE_NODATA(aJOBTITLE3) ' As String = "" '職稱3 100

        CHG_REPLACE_NODATA(aPHONE) ' As String = "" '連絡電話 20
        CHG_REPLACE_NODATA(aPHONE2) ' As String = "" '連絡電話 20
        CHG_REPLACE_NODATA(aCONFAX) 'As String = "" '傳真 20
        CHG_REPLACE_NODATA(aCELLPHONE) ' As String = "" '手機 20
        CHG_REPLACE_NODATA(aCELLPHONE2) ' As String = "" '手機 20
        CHG_REPLACE_NODATA(aEMAIL) ' As String = "" '電子郵件 70
        CHG_REPLACE_NODATA(aEMAIL2) ' As String = "" '電子郵件 70

        'CHG_REPLACE_NODATA(aMZIPCODE) ' As String = "" '地址-郵遞區號前3碼 3
        'CHG_REPLACE_NODATA(aMZIPCODE6W) ' As String = "" '地址-郵遞區號6碼 6
        CHG_REPLACE_NODATA(aMADDRESS) ' As String = "" '地址 100
        CHG_REPLACE_NODATA(aMADDRESS2) ' As String = "" '地址 100
        CHG_REPLACE_NODATA(aRMNOTE1) ' As String = "" '備註 300

        CHG_REPLACE_NODATA(aPUSHDISTID) ' As String = "" '推薦分署代碼 100
        CHG_REPLACE_NODATA(aPUSHREASON) ' As String = "" '推薦理由 100
        CHG_REPLACE_NODATA(aRUNTRAIN) ' As String = "" '是否辦訓 10
        CHG_REPLACE_NODATA(aTRAINDISTID) ' As String = "" '辦訓轄區 100
        CHG_REPLACE_NODATA(aADDYEARS_ROC)
        CHG_REPLACE_NODATA(aSTOPUSE) ' As String = "" '停用 3

        Dim pParms As New Hashtable From {{"MBRNAME", aMBRNAME}}
        Dim S_SQL As String = "SELECT 1 FROM OA_EXAMINER WHERE concat(MBRNAME,MBRNMSEQ)=@MBRNAME"
        Dim dt As DataTable = DbAccess.GetDataTable(S_SQL, objconn, pParms)
        If TIMS.dtHaveDATA(dt) Then
            Reason &= String.Format("(資料重複) 審查委員姓名「{0}」：資料已存在不可再次匯入<BR>", aMBRNAME)
        End If
        Return Reason
    End Function

    ''' <summary>SAVE OA_EXAMINER</summary>
    ''' <param name="iEMSEQ"></param>
    ''' <returns></returns>
    Function Utl_IMPORT07SaveData1(ByRef iEMSEQ As Integer) As Boolean
        Dim rst As Boolean = True '存檔ok

        Dim pParms As New Hashtable From {{"MBRNAME", aMBRNAME}}
        Dim s_sql As String = "SELECT * FROM OA_EXAMINER WHERE concat(MBRNAME,MBRNMSEQ)=@MBRNAME"
        Dim dr1 As DataRow = DbAccess.GetOneRow(s_sql, objconn, pParms)
        If dr1 IsNot Nothing Then Return False '(同名不可再次重建)

        'Dim s_sql As String = "SELECT * FROM OA_EXAMINER WHERE MBRNAME='" & aMBRNAME & "'"
        'Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn)
        'If dt.Rows.Count <> 0 Then Return

        Dim i_rst As Integer = 0
        iEMSEQ = DbAccess.GetNewId(objconn, "OA_EXAMINER_EMSEQ_SEQ,OA_EXAMINER,EMSEQ")
        Dim parms As New Hashtable 'parms.Clear()
        parms.Add("EMSEQ", iEMSEQ)
        parms.Add("RECRUIT", aRECRUIT)
        parms.Add("MBRNAME", If(aMBRNAME <> "", aMBRNAME, Convert.DBNull))
        parms.Add("UNITNAME", If(aUNITNAME <> "", aUNITNAME, Convert.DBNull))

        parms.Add("JOBTITLE", If(aJOBTITLE <> "", aJOBTITLE, Convert.DBNull))
        parms.Add("DEGREE", If(aDEGREE <> "", aDEGREE, Convert.DBNull))
        parms.Add("CERTIFICAT", If(aCERTIFICAT <> "", aCERTIFICAT, Convert.DBNull))
        parms.Add("SPECIALTY", If(aSPECIALTY <> "", aSPECIALTY, Convert.DBNull))

        parms.Add("PHONE", If(aPHONE <> "", aPHONE, Convert.DBNull))
        parms.Add("CELLPHONE", If(aCELLPHONE <> "", aCELLPHONE, Convert.DBNull))
        parms.Add("CONFAX", If(aCONFAX <> "", aCONFAX, Convert.DBNull))
        parms.Add("EMAIL", If(aEMAIL <> "", aEMAIL, Convert.DBNull))

        parms.Add("PHONE2", If(aPHONE2 <> "", aPHONE2, Convert.DBNull))
        parms.Add("CELLPHONE2", If(aCELLPHONE2 <> "", aCELLPHONE2, Convert.DBNull))
        parms.Add("EMAIL2", If(aEMAIL2 <> "", aEMAIL2, Convert.DBNull))

        'parms.Add("MZIPCODE", If(aMZIPCODE <> "", aMZIPCODE, Convert.DBNull))
        'parms.Add("MZIPCODE6W", If(aMZIPCODE6W <> "", aMZIPCODE6W, Convert.DBNull))
        parms.Add("MADDRESS", If(aMADDRESS <> "", aMADDRESS, Convert.DBNull))
        parms.Add("MADDRESS2", If(aMADDRESS2 <> "", aMADDRESS2, Convert.DBNull))
        parms.Add("RMNOTE1", If(aRMNOTE1 <> "", aRMNOTE1, Convert.DBNull))

        parms.Add("SERVUNIT1", If(aSERVUNIT1 <> "", aSERVUNIT1, Convert.DBNull))
        parms.Add("SERVTIME1", If(aSERVTIME1 <> "", aSERVTIME1, Convert.DBNull))
        parms.Add("JOBTITLE1", If(aJOBTITLE1 <> "", aJOBTITLE1, Convert.DBNull))
        parms.Add("SERVUNIT2", If(aSERVUNIT2 <> "", aSERVUNIT2, Convert.DBNull))
        parms.Add("SERVTIME2", If(aSERVTIME2 <> "", aSERVTIME2, Convert.DBNull))
        parms.Add("JOBTITLE2", If(aJOBTITLE2 <> "", aJOBTITLE2, Convert.DBNull))
        parms.Add("SERVUNIT3", If(aSERVUNIT3 <> "", aSERVUNIT3, Convert.DBNull))
        parms.Add("SERVTIME3", If(aSERVTIME3 <> "", aSERVTIME3, Convert.DBNull))
        parms.Add("JOBTITLE3", If(aJOBTITLE3 <> "", aJOBTITLE3, Convert.DBNull))

        parms.Add("PUSHDISTID", If(aPUSHDISTID <> "", aPUSHDISTID, Convert.DBNull))
        parms.Add("PUSHREASON", If(aPUSHREASON <> "", aPUSHREASON, Convert.DBNull))
        parms.Add("RUNTRAIN", If(aRUNTRAIN <> "", aRUNTRAIN, Convert.DBNull))

        parms.Add("TRAINDISTID", If(aTRAINDISTID <> "", aTRAINDISTID, Convert.DBNull))
        parms.Add("ADDYEARS", If(aADDYEARS_AD <> "", aADDYEARS_AD, Convert.DBNull))
        parms.Add("STOPUSE", If(aSTOPUSE <> "", aSTOPUSE, Convert.DBNull))

        parms.Add("RID", sm.UserInfo.RID)
        parms.Add("CREATEACCT", sm.UserInfo.UserID)
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        '新增
        Dim sql As String = ""
        sql &= " INSERT INTO OA_EXAMINER( EMSEQ ,RECRUIT ,MBRNAME ,UNITNAME" & vbCrLf
        sql &= " ,JOBTITLE ,DEGREE ,CERTIFICAT ,SPECIALTY" & vbCrLf
        sql &= " ,PHONE ,CELLPHONE ,CONFAX ,EMAIL" & vbCrLf
        sql &= " ,PHONE2 ,CELLPHONE2 ,EMAIL2" & vbCrLf
        'sql &= " ,MZIPCODE ,MZIPCODE6W ,MADDRESS ,MADDRESS2 ,RMNOTE1" & vbCrLf
        sql &= " ,MADDRESS ,MADDRESS2 ,RMNOTE1" & vbCrLf
        sql &= " ,SERVUNIT1 ,SERVTIME1 ,JOBTITLE1" & vbCrLf
        sql &= " ,SERVUNIT2 ,SERVTIME2 ,JOBTITLE2" & vbCrLf
        sql &= " ,SERVUNIT3 ,SERVTIME3 ,JOBTITLE3" & vbCrLf
        sql &= " ,PUSHDISTID ,PUSHREASON ,RUNTRAIN" & vbCrLf
        sql &= " ,TRAINDISTID ,ADDYEARS ,STOPUSE ,RID" & vbCrLf
        sql &= " ,CREATEACCT ,CREATEDATE" & vbCrLf
        sql &= " ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        sql &= " VALUES ( @EMSEQ ,@RECRUIT ,@MBRNAME ,@UNITNAME" & vbCrLf
        sql &= " ,@JOBTITLE ,@DEGREE ,@CERTIFICAT ,@SPECIALTY" & vbCrLf
        sql &= " ,@PHONE ,@CELLPHONE ,@CONFAX ,@EMAIL" & vbCrLf
        sql &= " ,@PHONE2 ,@CELLPHONE2 ,@EMAIL2" & vbCrLf
        'sql &= " ,@MZIPCODE ,@MZIPCODE6W ,@MADDRESS ,@RMNOTE1" & vbCrLf
        sql &= " ,@MADDRESS ,@MADDRESS2 ,@RMNOTE1" & vbCrLf
        sql &= " ,@SERVUNIT1 ,@SERVTIME1 ,@JOBTITLE1" & vbCrLf
        sql &= " ,@SERVUNIT2 ,@SERVTIME2 ,@JOBTITLE2" & vbCrLf
        sql &= " ,@SERVUNIT3 ,@SERVTIME3 ,@JOBTITLE3" & vbCrLf
        sql &= " ,@PUSHDISTID ,@PUSHREASON ,@RUNTRAIN" & vbCrLf
        sql &= " ,@TRAINDISTID ,@ADDYEARS ,@STOPUSE ,@RID" & vbCrLf
        sql &= " ,@CREATEACCT ,GETDATE()" & vbCrLf
        sql &= " ,@MODIFYACCT ,GETDATE())" & vbCrLf

        i_rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
        If i_rst > 0 Then Return True
        Return rst 'flagSaveOK1 = True
    End Function

    ''' <summary>匯入檔案</summary>
    ''' <param name="FullFileName1"></param>
    Sub Utl_IMPORT07(ByRef FullFileName1 As String)
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "FullFileName1", FullFileName1)
        TIMS.SetMyValue2(htSS, "FirstCol", "姓名")  '任1欄位名稱(必填)
        '儲存錯誤的原因
        Dim Reason As String = ""
        '上傳檔案/取得內容,  hFile1.PostedFile.SaveAs(FullFileName1)
        Dim dt_xls As DataTable = TIMS.Get_File1data(File1, Reason, htSS, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '取出資料庫的所有欄位--------   Start
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("MBRNAME"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End

        Dim iRowIndex As Integer = 1 '讀取行累計數
        Dim sql As String = ""
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            Reason = ""
            Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
            Reason = CheckImportData(colArray) '檢查資料正確性

            If Reason = "" Then
                Dim iEMSEQ As Integer = 0
                Try
                    Dim flag_SAVEOK As Boolean = Utl_IMPORT07SaveData1(iEMSEQ)
                    If flag_SAVEOK AndAlso iEMSEQ > 0 Then SaveData1_GCODE(iEMSEQ, aEXAMINERJOB)
                Catch ex As Exception
                    '取得錯誤資訊寫入
                    Dim strErrmsg1 As String = ""
                    strErrmsg1 &= "TC_10_001.aspx,Utl_IMPORT07()" & vbCrLf
                    strErrmsg1 &= TIMS.GetErrorMsg(Me) & vbCrLf '取得錯誤資訊寫入
                    strErrmsg1 &= "ex.Message:" & ex.Message & vbCrLf
                    'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg1, ex)
                    Reason &= ex.Message
                End Try
            End If

            If Reason <> "" Then
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)

                drWrong("Index") = iRowIndex
                drWrong("MBRNAME") = If(aMBRNAME <> "", aMBRNAME, "***")
                drWrong("Reason") = Reason
            End If
            iRowIndex += 1 '讀取行累計數
        Next
        '開始判別欄位存入------------   End

        '判斷匯出資料是否有誤
        Dim explain As String
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        Dim explain2 As String
        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"
        '開始判別欄位存入------------   End

        If dtWrong.Rows.Count = 0 Then
            '沒有錯誤資料
            If Reason <> "" Then
                Common.MessageBox(Me, explain & Reason)
                Exit Sub
            End If
            If explain <> "" Then
                Common.MessageBox(Me, explain)
                Exit Sub
            End If
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('TC_10_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    End Sub

    '匯入
    Protected Sub btnIMPORT1_Click(sender As Object, e As EventArgs) Handles btnIMPORT1.Click
        Dim sMyFileName As String = ""
        'Dim flag_File1_xls As Boolean = False
        'Dim flag_File1_ods As Boolean = False
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

        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call Utl_IMPORT07(FullFileName1)
    End Sub

End Class