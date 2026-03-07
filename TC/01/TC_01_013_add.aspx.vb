Imports System.IO

Partial Class TC_01_013_add
    Inherits AuthBasePage

    '"(相關變數/參數)"
    'Private components As System.ComponentModel.IContainer
    Dim strUplodFilePath As String = "" 'Server.MapPath(System.Configuration.ConfigurationSettings.AppSettings("uploadPathString2")) '問題單附件上傳路徑
    Dim strSyncUploadEnable As String = "" 'System.Configuration.ConfigurationSettings.AppSettings("syncUploadEnable")               '是否同步
    Dim strSyncUploadServer As String = "" 'System.Configuration.ConfigurationSettings.AppSettings("syncUploadServer")               '目標Server
    Dim strSyncUploadPort As String = "" 'System.Configuration.ConfigurationSettings.AppSettings("syncUploadPort")                   '目標Server
    Dim strSyncUploadPW As String = "" 'System.Configuration.ConfigurationSettings.AppSettings("syncUploadPW")                       '同步目標Server的密碼
    Dim strSyncUploadUser As String = "" 'System.Configuration.ConfigurationSettings.AppSettings("syncUploadUser")                   '同步目標Server的帳號
    Dim strSyncUploadFolder As String = "" '"/upload/Placepic/"                                                                      '附件路徑

    Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    Const cst_errMsg_2 As String = "上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    Const cst_errMsg_3 As String = "檔案位置錯誤!"
    Const cst_errMsg_4 As String = "檔案類型錯誤!"
    Const cst_errMsg_5 As String = "檔案類型錯誤，必須為圖片類型檔案!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗，請刪除此筆後重新上傳)"
    Const cst_PostedFile_max_size As Integer = 10485760
    Const cst_errMsg_7 As String = "檔案大小超過10MB!"
    Const cst_errMsg_8 As String = "請選擇場地圖片!"
    Const cst_errMsg_9 As String = "請選擇場地圖片--隸屬於教室1 或教室2!"
    Const cst_errMsg_10 As String = "無效的圖形格式。"

    Const cst_pic_iWidth As Integer = 960 '480
    Const cst_pic_iHeight As Integer = 480 '240

    'ProcessType:
    Dim rqRID As String = ""
    Dim rqPTID As String = ""
    Dim rqProcessType As String = ""
    Const cst_rqProcessType_Update As String = "Update"
    Const cst_rqProcessType_Insert As String = "Insert"
    Const cst_rqProcessType_View As String = "View"

    Dim irqProcessType As Int32

    Enum eePT_Enum As Int32
        xInsert = 10 'Insert
        xView = 20 'View
        xUpdate = 30 'Update
    End Enum

    'Dim dr As DataRow = Nothing
    'Dim da As SqlDataAdapter = Nothing
    Dim dtPIC As DataTable = Nothing
    Dim dt As DataTable = Nothing

    Const cst_UploadPath As String = "~/images/Placepic/"
    Const cst_downloadPath As String = "../../images/Placepic/"
    Dim Upload_Path As String = ""
    Dim download_Path As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Call Get_ServerConfigSet()

        rqPTID = TIMS.ClearSQM(Request("PTID"))
        rqRID = TIMS.ClearSQM(Request("RID"))
        rqProcessType = TIMS.ClearSQM(Request("ProcessType"))
        Dim v_ORGID As String = TIMS.ClearSQM(Request("OrgID"))
        If v_ORGID = "" Then v_ORGID = TIMS.Get_OrgID(rqRID, objconn)
        If v_ORGID = "" OrElse v_ORGID = "-1" Then v_ORGID = TIMS.Get_OrgID(rqRID, objconn)
        If v_ORGID = "" OrElse v_ORGID = "-1" Then v_ORGID = sm.UserInfo.OrgID
        Hid_comidno.Value = TIMS.Get_ComIDNOforOrgID(v_ORGID, objconn)
        labORGNAME.Text = TIMS.GET_ORGNAME(v_ORGID, objconn)

        irqProcessType = eePT_Enum.xUpdate
        Select Case rqProcessType
            Case cst_rqProcessType_Insert '"Insert"
                irqProcessType = eePT_Enum.xInsert
            Case cst_rqProcessType_View '"View"
                irqProcessType = eePT_Enum.xView
            Case Else 'cst_rqProcessType_Update
                irqProcessType = eePT_Enum.xUpdate
        End Select

        If Not IsPostBack Then
            '產生新的GUID 避免記憶體相同 而異常
            Call CREATE_NEW_GUID21()
            Call cCreate1()
            Call CeratePTIDDesc()
            Call ShowPTIDDesc() '顯示教室圖檔資料表
        End If
        'Const cst_msg1 As String = "建議使用 區碼-電話號碼"
        TIMS.Tooltip(ContactPHone, "建議使用 區碼-電話號碼")

    End Sub

    '產生新的GUID 避免記憶體相同 而異常
    Sub CREATE_NEW_GUID21()
        'If IsPostBack Then Exit Sub
        Hid_TRAINPLACE_GUID1.Value = TIMS.GetGUID()
        Session(Hid_TRAINPLACE_GUID1.Value) = Nothing
    End Sub

    Sub Get_ServerConfigSet()
        '建議圖片大小：480x240
        lab_msg_WH1.Text = "上傳照片檔案需為圖片類型，檔案請小於或等於10M (建議圖片大小：" & CStr(cst_pic_iWidth) & "x" & CStr(cst_pic_iHeight) & ")"

        strUplodFilePath = Server.MapPath(System.Configuration.ConfigurationSettings.AppSettings("uploadPathString2")) '問題單附件上傳路徑
        strSyncUploadEnable = System.Configuration.ConfigurationSettings.AppSettings("syncUploadEnable")               '是否同步
        strSyncUploadServer = System.Configuration.ConfigurationSettings.AppSettings("syncUploadServer")               '目標Server
        strSyncUploadPort = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPort")                   '目標Server
        strSyncUploadPW = System.Configuration.ConfigurationSettings.AppSettings("syncUploadPW")                       '同步目標Server的密碼
        strSyncUploadUser = System.Configuration.ConfigurationSettings.AppSettings("syncUploadUser")                   '同步目標Server的帳號
        strSyncUploadFolder = "/upload/Placepic/"
        Upload_Path = TIMS.Utl_GetConfigSet("Upload_PIC_Path")
        If Upload_Path = "" Then Upload_Path = cst_UploadPath
        download_Path = TIMS.Utl_GetConfigSet("download_PIC_Path")
        If download_Path = "" Then download_Path = cst_downloadPath
    End Sub

    '2011-11-13 add控制編輯視窗
    Private Sub CtrlFormEnable(ByVal blEnable As Boolean)
        'blEnable: True:啟用 /false:停用
        PlaceID.ReadOnly = Not blEnable
        PlaceID.Enabled = blEnable
        PlaceName.ReadOnly = Not blEnable
        IFICation.Enabled = blEnable
        AreaPoss.Enabled = blEnable
        FactMode.Enabled = blEnable
        ModeOther.ReadOnly = Not blEnable
        ContactName.ReadOnly = Not blEnable
        ContactPHone.ReadOnly = Not blEnable
        ContactFax.ReadOnly = Not blEnable
        ContactEMail.ReadOnly = Not blEnable
        MasterName.ReadOnly = Not blEnable
        ConNum.ReadOnly = Not blEnable
        txtPingNumber.ReadOnly = Not blEnable
        rblMODIFYTYPE.Enabled = blEnable
        findzip_but.Attributes.Add("style", "display:" & If(blEnable, "inline", "none"))
        Address.ReadOnly = Not blEnable
        Hwdesc.ReadOnly = Not blEnable
        OtherDesc.ReadOnly = Not blEnable
        depID.Enabled = blEnable
        But1.Visible = blEnable
        btnAdd.Visible = blEnable
        'Common.SetListItem(rblMODIFYTYPE, "Y")
    End Sub

    '檢查是否有重複資料
    Function Check_Repeat(ByVal RID As String, ByVal PTID As Integer) As Boolean
        '"檢查是否有重複資料"
        Dim rst As Boolean = False
        PlaceID.Text = TIMS.ClearSQM(PlaceID.Text)
        RID = TIMS.ClearSQM(RID)

        Dim sParms As New Hashtable
        sParms.Add("PlaceID", PlaceID.Text)
        sParms.Add("RID", RID)
        sParms.Add("PTID", PTID)
        Dim SqlStr As String = ""
        SqlStr &= " SELECT 'x' FROM PLAN_TRAINPLACE pt"
        SqlStr &= " JOIN ORG_ORGINFO oo ON pt.ComIDNO = oo.ComIDNO"
        SqlStr &= " JOIN AUTH_RELSHIP ar ON ar.OrgID = oo.OrgID"
        SqlStr &= " WHERE pt.PlaceID=@PlaceID AND ar.RID=@RID AND PTID <>@PTID"
        Dim dt As DataTable = DbAccess.GetDataTable(SqlStr, objconn, sParms)
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    Sub Check_dtPIC(ByRef dt As DataTable)
        '"Check_dtPIC"
        'Dim MyFile As System.IO.File
        Dim filename As String = ""
        If dt Is Nothing Then Exit Sub
        If dt.Rows.Count = 0 Then Exit Sub
        'foreach (System.Data.DataColumn col in tab.Columns) col.ReadOnly = false; 
        'For Each col As DataColumn In dt.Columns
        '    If col.ReadOnly Then col.ReadOnly = False
        'Next
        For i As Int16 = 0 To dt.Rows.Count - 1
            filename = dt.Rows(i).Item("placepic1").ToString
            If filename <> "" Then
                Dim flag_PIC_EXISTS As Boolean = TIMS.CHK_PIC_EXISTS(Server, Upload_Path, filename)
                Dim urlA1 As String = "<a class='l' target='_blank' href=""" & download_Path & filename & """>" & filename & "</a>"
                If Not flag_PIC_EXISTS Then urlA1 = "<font color='red'>" & cst_errMsg_6 & "</font>" '表示 檔案不存在
                dt.Rows(i)("okflag") = urlA1
            End If
        Next
    End Sub

    Sub cCreate1()
        TIMS.PL_placeholder(Hwdesc)
        But1.Attributes("onclick") = "return CheckAddPIC();"
        btnAdd.Attributes("onclick") = "return chkdata();"
        'Session(Hid_TRAINPLACE_GUID1.Value) = Nothing
        CtrlFormEnable(True)
        hidLID.Value = Convert.ToString(sm.UserInfo.LID) 'add
        Select Case irqProcessType
            Case eePT_Enum.xInsert
                'ProcessType.Text = "-新增"
                '20100208 按新增時代查詢之 場地代碼 & 場地名稱
                PlaceID.Text = TIMS.ClearSQM(Me.Request("PlaceNo"))
                PlaceName.Text = TIMS.ClearSQM(Me.Request("Place"))
            Case eePT_Enum.xView '2011-11-13 add檢視
                'ProcessType.Text = "-檢視"
                If Not IsPostBack Then LoalData1(rqPTID)
                CtrlFormEnable(False)
            Case Else 'Case eePT_Enum.xUpdate
                'ProcessType.Text = "-修改"
                If Not IsPostBack Then LoalData1(rqPTID)
                'PlaceID.ReadOnly = True
                'PlaceID.Enabled = False
        End Select

        With depID
            .Items.Add(New ListItem("==請選擇==", ""))
            .Items.Add(New ListItem("教室1", "1"))
            .Items.Add(New ListItem("教室2", "2"))
        End With

        '郵遞區號查詢
        Litcity_code.Text = TIMS.Get_WorkZIPB3Link2()

        Dim findzip_but_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(city_code, ZIPB3, hidZIP6W, TBCity, Address)
        findzip_but.Attributes.Add("onclick", findzip_but_Attr_VAL)
    End Sub

    ''' <summary> 取圖檔資料，修改則從資料庫取得 </summary>
    Sub CeratePTIDDesc()
        '"CeratePTIDDesc"
        Dim sql As String = ""
        If Session(Hid_TRAINPLACE_GUID1.Value) IsNot Nothing Then
            Session(Hid_TRAINPLACE_GUID1.Value) = Session(Hid_TRAINPLACE_GUID1.Value)
            Exit Sub
        End If

        If rqPTID <> "" Then
            'oooooo_x.jpg取得 x值
            'okflag=placepic1 表示 檔案存在
            sql = "" & vbCrLf
            sql &= " SELECT ptid,LEFT(RIGHT(placepic1,5),1) depID, placepic1 AS placepic1, placepic1 AS okflag FROM Plan_TrainPlace" & vbCrLf
            sql &= " WHERE placepic1 IS NOT NULL AND PTID = '" & rqPTID & "'" & vbCrLf
            sql &= " UNION" & vbCrLf
            sql &= " SELECT ptid,LEFT(RIGHT(placepic2,5),1) depID, placepic2 AS placepic1, placepic2 AS okflag FROM Plan_TrainPlace" & vbCrLf
            sql &= " WHERE placepic2 IS NOT NULL AND PTID = '" & rqPTID & "'" & vbCrLf
            sql &= " ORDER BY placepic1" & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            dt.Columns("PTID").AutoIncrement = True
            dt.Columns("PTID").AutoIncrementSeed = -1
            dt.Columns("PTID").AutoIncrementStep = -1
            Session(Hid_TRAINPLACE_GUID1.Value) = dt
            If dt.Rows.Count > 0 Then Check_dtPIC(Session(Hid_TRAINPLACE_GUID1.Value))
            Exit Sub
        End If

        Call CreateNewPTID() '新增的動作會進入此項

    End Sub

    ''' <summary> 新增的動作會進入此項 </summary>
    Sub CreateNewPTID()
        Dim dtPIC As New DataTable
        dtPIC.TableName = "Table1"
        '建立PIC資料格式Table----------------Start
        dtPIC.Columns.Add(New DataColumn("PTID"))
        dtPIC.Columns.Add(New DataColumn("depID"))
        dtPIC.Columns.Add(New DataColumn("PlacePIC1"))
        dtPIC.Columns.Add(New DataColumn("okflag"))
        'dtPIC.Columns.Add(New DataColumn("GUIDfilename"))
        dtPIC.Columns("PTID").AutoIncrement = True
        dtPIC.Columns("PTID").AutoIncrementSeed = -1
        dtPIC.Columns("PTID").AutoIncrementStep = -1
        '建立PIC資料格式Table----------------End
        Session(Hid_TRAINPLACE_GUID1.Value) = dtPIC
    End Sub

    ''' <summary> 顯示教室圖檔資料表 </summary>
    Sub ShowPTIDDesc()
        DataGrid3Table.Visible = False
        If Session(Hid_TRAINPLACE_GUID1.Value) Is Nothing Then
            CreatePICDT()
            Exit Sub
        End If
        dt = Session(Hid_TRAINPLACE_GUID1.Value)
        If dt.Rows.Count > 0 Then
            DataGrid3Table.Visible = True
            Datagrid1.DataSource = dt
            Datagrid1.DataBind()
        End If
        CreatePICDT()
    End Sub

    '重顯下拉視窗
    Sub CreatePICDT()
        If Session(Hid_TRAINPLACE_GUID1.Value) Is Nothing Then
            With depID
                .Items.Clear()
                .Items.Add(New ListItem("==請選擇==", ""))
                .Items.Add(New ListItem("教室1", "1"))
                .Items.Add(New ListItem("教室2", "2"))
            End With
            Call CreateNewPTID() '新增的動作會進入此項
            Exit Sub
        End If

        Dim dt As DataTable = Session(Hid_TRAINPLACE_GUID1.Value)
        Dim tmpflag1 As Boolean = True
        Dim tmpflag2 As Boolean = True
        Dim tmpstr As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                If tmpstr <> "" Then tmpstr &= ","
                tmpstr &= dt.Rows(i).Item("depID")
            End If
        Next

        Dim tmpary As Array
        tmpary = Split(tmpstr, ",")
        With depID
            .Items.Clear()
            .Items.Add(New ListItem("==請選擇==", ""))
            For i As Integer = LBound(tmpary) To UBound(tmpary)
                If tmpary(i) = "1" Then tmpflag1 = False
                If tmpary(i) = "2" Then tmpflag2 = False
            Next
            If tmpflag1 Then .Items.Add(New ListItem("教室1", "1"))
            If tmpflag2 Then .Items.Add(New ListItem("教室2", "2"))
        End With
        DataGrid3Table.Visible = False
        If dt.Rows.Count > 0 Then
            DataGrid3Table.Visible = True
            Datagrid1.DataSource = dt
            Datagrid1.DataBind()
        End If
    End Sub

    ''' <summary> 依PTID 取得有效資料 </summary>
    ''' <param name="PTID"></param>
    Sub LoalData1(ByVal PTID As String)
        PTID = TIMS.ClearSQM(PTID)
        If PTID = "" Then Return
        Dim flag_PTID_Use As Boolean = False 'false:未被使用
        Select Case sm.UserInfo.LID
            Case 2
                'lrMsg = "該場地目前被使用，不可刪除!!"
                If TIMS.ChkPTIDUse(PTID, objconn) Then flag_PTID_Use = True '被使用-有修改限制
        End Select

        Dim pms_s1 As New Hashtable From {{"PTID", PTID}}
        Dim sql As String = ""
        sql = " SELECT * FROM PLAN_TRAINPLACE WHERE PTID=@PTID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms_s1)
        If dr Is Nothing Then Exit Sub

        'Hid_comidno.Value = dr("COMIDNO").ToString()
        Hid_PlaceID.Value = Convert.ToString(dr("PlaceID"))
        PlaceID.Text = Convert.ToString(dr("PlaceID"))
        PlaceName.Text = dr("PlaceName").ToString()
        If Not IsDBNull(dr("ContactName")) Then ContactName.Text = dr("ContactName").ToString()
        If Not IsDBNull(dr("ContactPHone")) Then ContactPHone.Text = dr("ContactPHone").ToString()
        If Not IsDBNull(dr("ContactFax")) Then ContactFax.Text = dr("ContactFax").ToString()
        If Not IsDBNull(dr("ContactEMail")) Then ContactEMail.Text = dr("ContactEMail").ToString()
        IFICation.SelectedValue = dr("ClassIFICation").ToString()
        If Not IsDBNull(dr("AreaPoss")) Then AreaPoss.SelectedValue = dr("AreaPoss").ToString()
        FactMode.SelectedValue = dr("FactMode").ToString()
        If Not IsDBNull(dr("FactModeOther")) Then ModeOther.Text = dr("FactModeOther").ToString()

        city_code.Value = TIMS.AddZero(Convert.ToString(dr("ZipCode")), 3)
        hidZIP6W.Value = Convert.ToString(dr("ZIP6W"))
        ZIPB3.Value = TIMS.GetZIPCODEB3(hidZIP6W.Value)
        TBCity.Text = TIMS.GET_FullCCTName(objconn, Convert.ToString(dr("ZipCode")), Convert.ToString(dr("ZIP6W")))
        Address.Text = Convert.ToString(dr("Address"))

        ConNum.Text = Convert.ToString(dr("ConNum"))
        txtPingNumber.Text = Convert.ToString(dr("PINGNUMBER"))
        If txtPingNumber.Text <> "" Then txtPingNumber.Text = Val(TIMS.ROUND(txtPingNumber.Text, 4))
        Common.SetListItem(rblMODIFYTYPE, "Y")
        If Convert.ToString(dr("MODIFYTYPE")) = "D" Then Common.SetListItem(rblMODIFYTYPE, "N")
        If Not IsDBNull(dr("MasterName")) Then MasterName.Text = dr("MasterName")
        If Not IsDBNull(dr("Hwdesc")) Then
            Hwdesc.ForeColor = ColorTranslator.FromHtml("#000000")
            Hwdesc.Text = TIMS.ClearSQM2(dr("Hwdesc")) '.ToString()
        End If
        If Not IsDBNull(dr("OtherDesc")) Then OtherDesc.Text = dr("OtherDesc").ToString()

        '被使用 -有修改限制
        If flag_PTID_Use Then
            PlaceID.Enabled = False
            PlaceName.Enabled = False
            city_code.Disabled = True

            ZIPB3.Disabled = True
            Litcity_code.Text = "" '.Style("display") = "none" '不顯示 

            TBCity.Enabled = False
            'findzip_but.Disabled = True
            findzip_but.Style("display") = "none" '不顯示
            Address.Enabled = False
            Const cst_no_use_msg1 As String = "已有班級引用者，該欄位不可修改!"
            TIMS.Tooltip(PlaceID, cst_no_use_msg1)

            TIMS.Tooltip(PlaceName, cst_no_use_msg1)
            TIMS.Tooltip(city_code, cst_no_use_msg1)
            TIMS.Tooltip(ZIPB3, cst_no_use_msg1)
            TIMS.Tooltip(TBCity, cst_no_use_msg1)
            TIMS.Tooltip(Address, cst_no_use_msg1)
        End If
    End Sub

    '回上一頁
    Private Sub But5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But5.Click
        Dim url1 As String = "TC_01_013.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Function GetThumbNail(ByRef MyPage As Page, ByRef HtPP As Hashtable, ByRef ImgStream As System.IO.Stream) As String
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
        'File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & GUIDfilename))
        'oImg.Save(Server.MapPath("images") + "\" + strGuid + strFileExt, GetImageType(strContentType))
        'Dim s_save_file_name As String = String.Concat(MyPage.Server.MapPath(Upload_Path), "\", strGuid, strFileExt)
        Dim s_save_file_name As String = String.Concat(MyPage.Server.MapPath(Upload_Path), "\", strFileName)
        oImg.Save(s_save_file_name, TIMS.GetImageType(strContentType))
        '直接輸出url文件
        'Response.Redirect("images/" + strGuid + strFileExt)
        '以下顯示在屏幕上
        'Response.ContentType = strContentType
        'Dim MemStream As New MemoryStream
        '' 注意：這裡如果直接用 oImg.Save(Response.OutputStream, GetImageType(strContentType))
        '' 對不同的格式可能會出錯，比如Png格式。
        'oImg.Save(MemStream, GetImageType(strContentType))
        'MemStream.WriteTo(Response.OutputStream)
        'Return String.Concat(strGuid, strFileExt)
        Return strFileName
    End Function

    Private Sub But1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim Upload_Path As String = "~/images/Placepic/"
        Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
        Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)
        If LCase(MyPostedFile.ContentType.ToString()).IndexOf("image") < 0 Then
            'Common.RespWrite(Me, "無效的圖形格式。")
            Common.MessageBox(Me, cst_errMsg_10)
            Exit Sub
        End If
        If Me.depID.SelectedIndex = 0 Then
            Common.MessageBox(Me, cst_errMsg_9)
            Exit Sub
        End If
        ' GetThumbNail(MyPostedFile.FileName, 320, 240, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream)
        If File1.Value = "" Then
            Common.MessageBox(Me, cst_errMsg_8)
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, cst_errMsg_3)
            Exit Sub
        End If
        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Dim LowerFileType As String = LCase(MyFileType)
        Select Case LowerFileType 'LCase(MyFileType)
            Case "jpg", "bmp", "gif", "png"
                '檢查檔案格式與大小 End 'flag = "," '=== (edit，by:20181128)
                If File1.PostedFile.ContentLength > cst_PostedFile_max_size Then
                    Common.MessageBox(Me, cst_errMsg_7)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Exit Sub
        End Select
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        'Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, LowerFileType) Then Return

        '上傳檔案
        'FileName1 = "P" & Request("PTID") & "_" & depID.SelectedItem.Value & "." & MyFileType
        'GUIDfilename = Guid.NewGuid.ToString & "." & MyFileType
        'File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & GUIDfilename))
        Dim vFILENAME1 As String = TIMS.GetValidFileName(String.Concat(TIMS.GetGUID(), ".", LowerFileType))
        'Dim s_FileName As String = MyPostedFile.FileName
        TIMS.MyCreateDir(Me, Upload_Path)
        Dim Upload_Path_File As String = String.Concat(Upload_Path, vFILENAME1)

        Dim HtPP As New Hashtable
        TIMS.SetMyValue2(HtPP, "FileName", vFILENAME1)
        TIMS.SetMyValue2(HtPP, "iWidth", cst_pic_iWidth)
        TIMS.SetMyValue2(HtPP, "iheight", cst_pic_iHeight)
        TIMS.SetMyValue2(HtPP, "ContentType", MyPostedFile.ContentType)
        TIMS.SetMyValue2(HtPP, "blnGetFromFile", False)
        TIMS.SetMyValue2(HtPP, "Upload_Path", Upload_Path)

        '上傳檔案/存檔：檔名
        Dim GUIDfilename As String = ""
        Dim objLock_ThumbNail As New Object
        SyncLock objLock_ThumbNail
            Try
                'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
                GUIDfilename = GetThumbNail(Me, HtPP, MyPostedFile.InputStream)
            Catch ex As Exception
                TIMS.LOG.Warn(ex.Message, ex)

                Common.MessageBox(Me, cst_errMsg_2)
                'Common.MessageBox(Me, ex.ToString)
                Dim strErrmsg As String = cst_errMsg_2 & vbCrLf
                'strErrmsg &= String.Concat("GUIDfilename: " & GUIDfilename & vbCrLf
                strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName) & vbCrLf
                strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType) & vbCrLf
                strErrmsg &= String.Concat("Upload_Path: ", Upload_Path) & vbCrLf
                'Server.MapPath(Upload_Path & filename) '   Dim s_save_file_name As String = String.Concat(MyPage.Server.MapPath(Upload_Path), "\", strFileName)
                strErrmsg &= String.Concat("Upload_Path & filename: ", Upload_Path_File) & vbCrLf
                strErrmsg &= String.Concat("Server.MapPath(Upload_Path & filename): ", Server.MapPath(Upload_Path_File)) & vbCrLf
                'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                TIMS.WriteTraceLog(Me, ex, strErrmsg)
                Exit Sub
            End Try
        End SyncLock

        If Session(Hid_TRAINPLACE_GUID1.Value) Is Nothing Then
            sm.LastErrorMessage = cst_errMsg_1
            Exit Sub
        End If

        Try
            '此時 Session(Hid_TRAINPLACE_GUID1.Value) 不管新增或修改都有值了(此動作必為新增)
            dt = Session(Hid_TRAINPLACE_GUID1.Value)
            Dim dr As DataRow = dt.NewRow
            dt.Rows.Add(dr)
            'dr("PTID") = Me.depID.SelectedValue
            dr("depID") = Me.depID.SelectedValue
            dr("PlacePIC1") = GUIDfilename
            dr("okflag") = GUIDfilename
            Session(Hid_TRAINPLACE_GUID1.Value) = dt
            ShowPTIDDesc()
            'CeratePTIDDesc()
            'ShowPTIDDesc()
            'CreatePICDT()
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = "" & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

    End Sub

    Private Sub Datagrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid1.ItemDataBound
        Const Cst_教室 As Integer = 0
        'Const Cst_序號 As Integer = 1
        'Const Cst_圖檔名稱 As Integer = 2
        'Dim i As Integer

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim LabdepID As Label = e.Item.FindControl("LabdepID")
                Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim But4 As Button = e.Item.FindControl("But4")

                e.Item.Cells(Cst_教室).Text = If(e.Item.Cells(Cst_教室).Text = "1", "教室一", "教室二")

                If Not IsDBNull(drv("PlacePIC1")) Then
                    LabFileName1.Text = drv("PlacePIC1").ToString()
                    HFileName.Value = drv("PlacePIC1").ToString()
                End If
                If drv("PlacePIC1") <> drv("okflag") Then LabFileName1.Text = drv("okflag").ToString
                But4.CommandArgument = drv("depID").ToString '刪除
                But4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '2011-11-13 訓練單位改只能檢視不能修改
                But4.Visible = True
                If irqProcessType = eePT_Enum.xView Then But4.Visible = False
        End Select
    End Sub

    Private Sub Datagrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid1.ItemCommand
        'Dim Upload_Path As String = "~/images/Placepic/"
        If Session(Hid_TRAINPLACE_GUID1.Value) Is Nothing Then
            'Turbo2.Common.MessageBox(Me, "資料有誤請重新查詢!!")
            'Common.RespWrite(Me, "<script>alert('資料有誤請重新查詢');</script>")
            'Me.Response.End()
            sm.LastErrorMessage = cst_errMsg_1
            Exit Sub
        End If

        Dim dt1 As DataTable
        Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Select Case e.CommandName
            Case "del"
                Dim sql As String
                Dim dt As DataTable = Session(Hid_TRAINPLACE_GUID1.Value)
                Dim da As SqlDataAdapter = Nothing
                If dt.Select("depID='" & e.CommandArgument & "'").Length <> 0 Then
                    If rqPTID <> "" Then
                        '若是為修改，則可直接動資料庫
                        sql = " SELECT * FROM PLAN_TRAINPLACE WHERE PTID = '" & rqPTID & "'"
                        dt1 = DbAccess.GetDataTable(sql, da, objconn)
                        'UPDATE
                        Dim dr As DataRow = dt1.Rows(0)
                        If Left(Right(HFileName.Value, 5), 1) = "1" And HFileName.Value.IndexOf("_1.") > -1 Then dr("PlacePIC1") = Convert.DBNull
                        If Left(Right(HFileName.Value, 5), 1) = "2" And HFileName.Value.IndexOf("_2.") > -1 Then dr("PlacePIC2") = Convert.DBNull
                        DbAccess.UpdateDataTable(dt1, da)
                    End If
                    dt.Select("depID='" & e.CommandArgument & "'")(0).Delete()
                    If File.Exists(Server.MapPath(Upload_Path & HFileName.Value)) Then File.Delete(Server.MapPath(Upload_Path & HFileName.Value))
                    If strSyncUploadEnable = "1" Then
                        Dim ftp As New EnterpriseDT.Net.Ftp.FTPConnection
                        Try
                            ftp.ServerAddress = strSyncUploadServer
                            ftp.ServerPort = strSyncUploadPort
                            ftp.UserName = strSyncUploadUser
                            ftp.Password = strSyncUploadPW
                            ftp.ServerDirectory = strSyncUploadFolder
                            ftp.Connect()
                            If HFileName.Value <> "" AndAlso ftp.Exists(HFileName.Value) Then ftp.DeleteFile(HFileName.Value)
                            ftp.Close()
                        Catch ex As Exception
                            TIMS.LOG.Warn(ex.Message, ex)
                            'If ftp.IsConnected = True Then ftp.Close()
                        End Try
                    End If
                    Session(Hid_TRAINPLACE_GUID1.Value) = dt
                End If
                Datagrid1.EditItemIndex = -1
            Case "cancel"
                Datagrid1.EditItemIndex = -1
        End Select
        CeratePTIDDesc()
        ShowPTIDDesc() '顯示教室圖檔資料表
        'ShowPTIDDesc()
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        rqPTID = TIMS.ClearSQM(rqPTID)

        PlaceID.Text = TIMS.ClearSQM(PlaceID.Text)
        If PlaceID.Text = "" Then Errmsg += "請輸入 場地代碼" & vbCrLf
        If rqPTID = "" Then
            '修改時不檢核
            If PlaceID.Text <> "" Then
                'PlaceID.Text = PlaceID.Text.Trim
                PlaceID.Text = TIMS.ChangeIDNO(PlaceID.Text)
                If Not TIMS.CheckABC123(PlaceID.Text) Then
                    Errmsg &= "場地代碼必須為英數字" & vbCrLf
                ElseIf Len(PlaceID.Text) > 10 Then
                    Errmsg &= "場地代碼 限英數字10字以內" & vbCrLf
                ElseIf Not TIMS.CheckABC321(PlaceID.Text) Then
                    Errmsg &= "場地代碼必須為英數字10字以內" & vbCrLf
                End If
            End If
        End If
        'If rqPTID <> "" Then
        '    If PlaceID.Text <> "" Then
        '        If Len(PlaceID.Text) > 10 Then Errmsg += "場地代碼 限 10字以內" & vbCrLf
        '    End If
        'End If

        PlaceName.Text = TIMS.ClearSQM(PlaceName.Text)
        If PlaceName.Text = "" Then Errmsg += "請輸入 場地名稱" & vbCrLf
        If IFICation.SelectedValue = "" Then Errmsg += "請選擇 場地類別" & vbCrLf
        If AreaPoss.SelectedValue = "" Then Errmsg += "請選擇 場地屬性" & vbCrLf
        If FactMode.SelectedValue = "" Then Errmsg += "請選擇 場地類型" & vbCrLf
        ConNum.Text = TIMS.ClearSQM(ConNum.Text)
        If ConNum.Text = "" Then Errmsg += "請輸入 訓練容納人數" & vbCrLf
        If ConNum.Text <> "" Then
            If Not TIMS.IsNumeric1(ConNum.Text) Then Errmsg += "訓練容納人數 應為數字格式" & vbCrLf
        End If
        txtPingNumber.Text = TIMS.ClearSQM(txtPingNumber.Text)
        If txtPingNumber.Text = "" Then Errmsg &= "請輸入 坪數" & vbCrLf
        If txtPingNumber.Text <> "" Then
            If TIMS.IsNumeric1(txtPingNumber.Text) Then txtPingNumber.Text = Val(TIMS.ROUND(txtPingNumber.Text, 4))
            If Not TIMS.IsNumeric1(txtPingNumber.Text) Then Errmsg &= "坪數 應為數字格式(可含小數點4位)" & vbCrLf
        End If
        city_code.Value = TIMS.ClearSQM(city_code.Value) '場地郵遞區號
        ZIPB3.Value = TIMS.ClearSQM(ZIPB3.Value)
        hidZIP6W.Value = TIMS.GetZIPCODE6W(city_code.Value, ZIPB3.Value)
        Address.Text = TIMS.ClearSQM(Address.Text) '場地地址

        If Not TIMS.IsZipCode(city_code.Value, objconn) Then Errmsg += "場地地址 郵遞區號前3碼 有誤" & vbCrLf
        TIMS.CheckZipCODEB3(ZIPB3.Value, "場地地址 郵遞區號後2碼或3碼", True, Errmsg)
        If Address.Text = "" Then Errmsg += "場地地址 不可為空" & vbCrLf

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>SAVEDATA</summary>
    Sub sSaveData1()
        'Hid_comidno.Value
        'Dim dr As DataRow = Nothing
        PlaceID.Text = TIMS.ClearSQM(PlaceID.Text)
        Hid_comidno.Value = TIMS.ClearSQM(Hid_comidno.Value)
        rqPTID = TIMS.ClearSQM(rqPTID)
        Hwdesc.Text = TIMS.ClearSQM2(Hwdesc.Text)

        Dim FileName1 As String = ""
        Dim MyFileType As String = ""
        Dim pPMS As New Hashtable
        Dim pSql As String = ""
        If irqProcessType = eePT_Enum.xUpdate Then
            pPMS.Clear()
            pPMS.Add("PTID", rqPTID)
            pPMS.Add("COMIDNO", Hid_comidno.Value)
            pSql = "SELECT PTID,PLACEID FROM PLAN_TRAINPLACE WHERE PTID =@PTID AND COMIDNO =@COMIDNO"
        Else
            pPMS.Clear()
            pPMS.Add("PLACEID", PlaceID.Text)
            pPMS.Add("COMIDNO", Hid_comidno.Value)
            pSql = "SELECT PTID,PLACEID FROM PLAN_TRAINPLACE WHERE PLACEID =@PLACEID AND COMIDNO =@COMIDNO"
        End If
        dt = DbAccess.GetDataTable(pSql, objconn, pPMS)

        Dim iPTID As Integer = 0
        If dt.Rows.Count = 0 Then
            '新增
            Dim sqlA As String = ""
            sqlA &= " INSERT INTO PLAN_TRAINPLACE (PTID,PLACEID,PLACENAME,COMIDNO,CONTACTNAME,CONTACTPHONE,CONTACTFAX,CONTACTEMAIL" & vbCrLf
            sqlA &= " ,CLASSIFICATION ,AREAPOSS,FACTMODE,FACTMODEOTHER" & vbCrLf
            sqlA &= " ,ZIP6W,ZIPCODE,ADDRESS,CONNUM,MASTERNAME,HWDESC,OTHERDESC,PINGNUMBER,MODIFYTYPE,MODIFYACCT,MODIFYDATE)" & vbCrLf
            sqlA &= " VALUES (@PTID,@PLACEID,@PLACENAME,@COMIDNO,@CONTACTNAME,@CONTACTPHONE,@CONTACTFAX,@CONTACTEMAIL" & vbCrLf
            sqlA &= " ,@CLASSIFICATION,@AREAPOSS,@FACTMODE,@FACTMODEOTHER" & vbCrLf
            sqlA &= " ,@ZIP6W,@ZIPCODE,@ADDRESS,@CONNUM,@MASTERNAME,@HWDESC,@OTHERDESC,@PINGNUMBER,@MODIFYTYPE,@MODIFYACCT,GETDATE())" & vbCrLf

            iPTID = DbAccess.GetNewId(objconn, "PLAN_TRAINPLACE_PTID_SEQ,PLAN_TRAINPLACE,PTID")
            Dim ParmsA As New Hashtable
            'ParmsA.Clear()
            ParmsA.Add("PTID", iPTID)
            ParmsA.Add("PLACEID", PlaceID.Text) '場地代碼
            ParmsA.Add("PLACENAME", PlaceName.Text) '場地名稱
            ParmsA.Add("COMIDNO", Hid_comidno.Value) 'drX2("ComIDNO") '統編
            ParmsA.Add("CONTACTNAME", ContactName.Text) '聯絡人
            ParmsA.Add("CONTACTPHONE", If(ContactPHone.Text = "", Convert.DBNull, ContactPHone.Text)) '聯絡人電話
            ParmsA.Add("CONTACTFAX", If(ContactFax.Text = "", Convert.DBNull, ContactFax.Text)) '聯絡人傳真
            ParmsA.Add("CONTACTEMAIL", If(ContactEMail.Text = "", Convert.DBNull, ContactEMail.Text)) '聯絡人email
            ParmsA.Add("CLASSIFICATION", CInt(IFICation.SelectedValue)) '場地類別
            ParmsA.Add("AREAPOSS", AreaPoss.SelectedValue) '場地屬地
            ParmsA.Add("FACTMODE", FactMode.SelectedValue) '場地類型
            ParmsA.Add("FACTMODEOTHER", If(ModeOther.Text = "", Convert.DBNull, ModeOther.Text)) '場地類型其他說明
            ParmsA.Add("ZIPCODE", CInt(city_code.Value)) '場地郵遞區號
            ParmsA.Add("ZIP6W", If(hidZIP6W.Value <> "", hidZIP6W.Value, Convert.DBNull)) '場地郵遞區號後碼 
            ParmsA.Add("ADDRESS", TIMS.ClearSQM(Address.Text)) '場地地址
            ParmsA.Add("CONNUM", If(ConNum.Text = "", Convert.DBNull, ConNum.Text)) 'CInt(ConNum.Text) '容納人數
            ParmsA.Add("MASTERNAME", If(MasterName.Text = "", Convert.DBNull, MasterName.Text)) '負責人
            ParmsA.Add("HWDESC", If(Hwdesc.Text = "", Convert.DBNull, Hwdesc.Text)) '硬體設施說明
            ParmsA.Add("OTHERDESC", If(OtherDesc.Text = "", Convert.DBNull, OtherDesc.Text)) '其他設施說明
            ParmsA.Add("PINGNUMBER", Val(txtPingNumber.Text)) '坪數
            ParmsA.Add("MODIFYTYPE", If(rblMODIFYTYPE.SelectedValue = "N", "D", Convert.DBNull)) '啟用／停用
            ParmsA.Add("MODIFYACCT", sm.UserInfo.UserID)
            'parms.Add("MODIFYDATE", MODIFYDATE)
            DbAccess.ExecuteNonQuery(sqlA, objconn, ParmsA)
        Else
            '修改
            'dr = dt.Rows(0)
            iPTID = dt.Rows(0)("PTID")
            PlaceID.Text = dt.Rows(0)("PlaceID")
            Dim sqlU As String = ""
            sqlU &= " UPDATE PLAN_TRAINPLACE" & vbCrLf 'sql &= " SET PLACEID = @PLACEID" & vbCrLf
            sqlU &= " SET PLACENAME = @PLACENAME" & vbCrLf
            sqlU &= " ,COMIDNO = @COMIDNO" & vbCrLf
            sqlU &= " ,CONTACTNAME = @CONTACTNAME" & vbCrLf
            sqlU &= " ,CONTACTPHONE = @CONTACTPHONE" & vbCrLf
            sqlU &= " ,CONTACTFAX = @CONTACTFAX" & vbCrLf
            sqlU &= " ,CONTACTEMAIL = @CONTACTEMAIL" & vbCrLf
            sqlU &= " ,CLASSIFICATION = @CLASSIFICATION" & vbCrLf
            sqlU &= " ,AREAPOSS = @AREAPOSS" & vbCrLf
            sqlU &= " ,FACTMODE = @FACTMODE" & vbCrLf
            sqlU &= " ,FACTMODEOTHER = @FACTMODEOTHER" & vbCrLf
            sqlU &= " ,ZIPCODE = @ZIPCODE" & vbCrLf
            sqlU &= " ,ZIP6W = @ZIP6W" & vbCrLf
            sqlU &= " ,ADDRESS = @ADDRESS" & vbCrLf
            sqlU &= " ,CONNUM = @CONNUM" & vbCrLf
            sqlU &= " ,MASTERNAME = @MASTERNAME" & vbCrLf
            sqlU &= " ,HWDESC = @HWDESC" & vbCrLf
            sqlU &= " ,OTHERDESC = @OTHERDESC" & vbCrLf
            sqlU &= " ,PINGNUMBER = @PINGNUMBER" & vbCrLf
            sqlU &= " ,MODIFYTYPE = @MODIFYTYPE" & vbCrLf
            sqlU &= " ,MODIFYACCT = @MODIFYACCT" & vbCrLf
            sqlU &= " ,MODIFYDATE = GETDATE()" & vbCrLf
            sqlU &= " WHERE PTID=@PTID" & vbCrLf
            Dim ParmsU As New Hashtable
            'ParmsU.Clear() 'Parms.Add("PLACEID", PlaceID.Text) '場地代碼
            ParmsU.Add("PLACENAME", PlaceName.Text) '場地名稱
            ParmsU.Add("COMIDNO", Hid_comidno.Value) 'drX2("ComIDNO") '統編
            ParmsU.Add("CONTACTNAME", ContactName.Text) '聯絡人
            ParmsU.Add("CONTACTPHONE", If(ContactPHone.Text = "", Convert.DBNull, ContactPHone.Text)) '聯絡人電話
            ParmsU.Add("CONTACTFAX", If(ContactFax.Text = "", Convert.DBNull, ContactFax.Text)) '聯絡人傳真
            ParmsU.Add("CONTACTEMAIL", If(ContactEMail.Text = "", Convert.DBNull, ContactEMail.Text)) '聯絡人email
            ParmsU.Add("CLASSIFICATION", CInt(IFICation.SelectedValue)) '場地類別
            ParmsU.Add("AREAPOSS", AreaPoss.SelectedValue) '場地屬地
            ParmsU.Add("FACTMODE", FactMode.SelectedValue) '場地類型
            ParmsU.Add("FACTMODEOTHER", If(ModeOther.Text = "", Convert.DBNull, ModeOther.Text)) '場地類型其他說明
            ParmsU.Add("ZIPCODE", CInt(city_code.Value)) '場地郵遞區號
            ParmsU.Add("ZIP6W", If(hidZIP6W.Value <> "", hidZIP6W.Value, Convert.DBNull)) '場地郵遞區號後碼 
            ParmsU.Add("ADDRESS", TIMS.ClearSQM(Address.Text)) '場地地址
            ParmsU.Add("CONNUM", If(ConNum.Text = "", Convert.DBNull, ConNum.Text)) 'CInt(ConNum.Text) '容納人數
            ParmsU.Add("MASTERNAME", If(MasterName.Text = "", Convert.DBNull, MasterName.Text)) '負責人
            ParmsU.Add("HWDESC", If(Hwdesc.Text = "", Convert.DBNull, Hwdesc.Text)) '硬體設施說明
            ParmsU.Add("OTHERDESC", If(OtherDesc.Text = "", Convert.DBNull, OtherDesc.Text)) '其他設施說明
            ParmsU.Add("PINGNUMBER", Val(txtPingNumber.Text)) '坪數
            ParmsU.Add("MODIFYTYPE", If(rblMODIFYTYPE.SelectedValue = "N", "D", Convert.DBNull)) '啟用／停用
            ParmsU.Add("MODIFYACCT", sm.UserInfo.UserID)
            'parms.Add("MODIFYDATE", MODIFYDATE)
            ParmsU.Add("PTID", iPTID)
            DbAccess.ExecuteNonQuery(sqlU, objconn, ParmsU)
        End If
        Hid_PlaceID.Value = PlaceID.Text

        'sql = "" & vbCrLf
        'sql &= " SELECT PTID ,PLACEPIC1 ,PLACEPIC2" & vbCrLf
        'sql &= " FROM PLAN_TRAINPLACE" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND PLACEID = '" & PlaceID.Text & "'" & vbCrLf
        'sql &= " AND ComIDNO = '" & Hid_comidno.Value & "'" & vbCrLf
        'sql &= " AND PTID = '" & iPTID & "'" & vbCrLf
        'sql &= " and ModifyDate>= " & TIMS.to_date(Nowstr.ToString("yyyy/MM/dd HH:mm:ss"))
        'dt = DbAccess.GetDataTable(sql, da, objconn)
        'dr = dt.Rows(0)

        If Session(Hid_TRAINPLACE_GUID1.Value) IsNot Nothing Then
            dtPIC = Session(Hid_TRAINPLACE_GUID1.Value)
            'Dim i As Integer
            If dtPIC.Rows.Count > 0 Then
                For i As Integer = 0 To dtPIC.Rows.Count - 1
                    'If i = 0 Then
                    If Not dtPIC.Rows(i).RowState = DataRowState.Deleted Then
                        If dtPIC.Rows(i).Item("okflag").ToString = dtPIC.Rows(i).Item("PlacePIC1").ToString Then
                            Select Case dtPIC.Rows(i).Item("depID").ToString
                                Case "1"
                                    'MyFileType = Split(dtPIC.Rows(i).Item("PlacePIC1").ToString, ".")((Split(dtPIC.Rows(i).Item("PlacePIC1").ToString, ".")).Length - 1)
                                    MyFileType = Get_MyFileType1(dtPIC.Rows(i).Item("PlacePIC1").ToString)
                                    FileName1 = String.Concat("P", iPTID, "_1.", MyFileType)
                                    Dim flag_CanSavePIC As Boolean = If(dtPIC.Rows(i).Item("PlacePIC1").ToString <> FileName1, True, False)
                                    If flag_CanSavePIC Then
                                        'dr("PlacePIC1") = FileName1
                                        Dim tsqlU As String = ""
                                        tsqlU &= " UPDATE PLAN_TRAINPLACE"
                                        tsqlU &= " SET PLACEPIC1=@PLACEPIC1 ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()"
                                        tsqlU &= " WHERE PTID=@PTID AND COMIDNO=@COMIDNO"
                                        Dim Parms As New Hashtable
                                        'Parms.Clear()
                                        Parms.Add("PLACEPIC1", FileName1)
                                        Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                                        Parms.Add("PTID", iPTID)
                                        Parms.Add("COMIDNO", Hid_comidno.Value)
                                        DbAccess.ExecuteNonQuery(tsqlU, objconn, Parms)

                                        Call TIMS.MyFileDelete(Server.MapPath(Upload_Path & FileName1))
                                        'If IO.File.Exists(Server.MapPath(Upload_Path & FileName1)) Then IO.File.Delete(Server.MapPath(Upload_Path & FileName1))
                                        If IO.File.Exists(Server.MapPath(Upload_Path & dtPIC.Rows(i).Item("PlacePIC1").ToString)) Then
                                            IO.File.Move(Server.MapPath(Upload_Path & dtPIC.Rows(i).Item("PlacePIC1").ToString), Server.MapPath(Upload_Path & FileName1))
                                        End If
                                    End If

                                Case "2"
                                    'MyFileType = Split(dtPIC.Rows(i).Item("PlacePIC1").ToString, ".")((Split(dtPIC.Rows(i).Item("PlacePIC1").ToString, ".")).Length - 1)
                                    MyFileType = Get_MyFileType1(dtPIC.Rows(i).Item("PlacePIC1").ToString)
                                    FileName1 = String.Concat("P", iPTID, "_2.", MyFileType)
                                    Dim flag_CanSavePIC As Boolean = If(dtPIC.Rows(i).Item("PlacePIC1").ToString <> FileName1, True, False)
                                    If flag_CanSavePIC Then
                                        'dr("PlacePIC2") = FileName1
                                        Dim tsqlU As String = ""
                                        tsqlU &= " UPDATE PLAN_TRAINPLACE"
                                        tsqlU &= " SET PLACEPIC2 = @PLACEPIC2 ,MODIFYACCT = @MODIFYACCT ,MODIFYDATE = GETDATE()"
                                        tsqlU &= " WHERE PTID = @PTID AND COMIDNO = @COMIDNO"
                                        Dim Parms As New Hashtable
                                        'Parms.Clear()
                                        Parms.Add("PLACEPIC2", FileName1)
                                        Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                                        Parms.Add("PTID", iPTID)
                                        Parms.Add("COMIDNO", Hid_comidno.Value)
                                        DbAccess.ExecuteNonQuery(tsqlU, objconn, Parms)

                                        Call TIMS.MyFileDelete(Server.MapPath(Upload_Path & FileName1))
                                        'If IO.File.Exists(Server.MapPath(Upload_Path & FileName1)) Then IO.File.Delete(Server.MapPath(Upload_Path & FileName1))
                                        If IO.File.Exists(Server.MapPath(Upload_Path & dtPIC.Rows(i).Item("PlacePIC1").ToString)) Then
                                            IO.File.Move(Server.MapPath(Upload_Path & dtPIC.Rows(i).Item("PlacePIC1").ToString), Server.MapPath(Upload_Path & FileName1))
                                        End If
                                    End If

                            End Select
                            '儲存到145上
                            If IO.File.Exists(Server.MapPath(Upload_Path & FileName1)) Then
                                If strSyncUploadEnable = "1" Then
                                    Dim ftp As New EnterpriseDT.Net.Ftp.FTPConnection
                                    Try
                                        ftp.ServerAddress = strSyncUploadServer
                                        ftp.ServerPort = strSyncUploadPort
                                        ftp.UserName = strSyncUploadUser
                                        ftp.Password = strSyncUploadPW
                                        ftp.ServerDirectory = strSyncUploadFolder
                                        ftp.Connect()
                                        ftp.UploadFile(Server.MapPath(Upload_Path & FileName1), FileName1)
                                        ftp.Close()
                                    Catch ex As Exception
                                        TIMS.LOG.Warn(ex.Message, ex)
                                        'If ftp.IsConnected = True Then ftp.Close()
                                    End Try
                                End If
                            End If
                        End If
                    End If
                Next
            End If
        End If
        Session(Hid_TRAINPLACE_GUID1.Value) = Nothing
        'DbAccess.UpdateDataTable(dt, da)

        '代碼修正
        If rqProcessType = cst_rqProcessType_Update Then
            Dim htSS As New Hashtable
            'htSS.Clear()
            htSS.Add("PLACEID", PlaceID.Text)
            htSS.Add("HIDCOMIDNO", Hid_comidno.Value)
            htSS.Add("HIDPLACEID", Hid_PlaceID.Value)
            'SELECT SCIPLACEID,SCIPLACEID2,TECHPLACEID,TECHPLACEID2 FROM PLAN_PLANINFO
            'UPDATE - PLAN_PLANINFO
            TIMS.SetMyValue2(htSS, "TABLE1", "PLAN_PLANINFO")
            'PLAN_PLANINFO 的 學科場地1
            TIMS.SetMyValue2(htSS, "COLUMN1", "SCIPLACEID")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'plan_planinfo 的 學科場地2
            TIMS.SetMyValue2(htSS, "COLUMN1", "SCIPLACEID2")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'plan_planinfo 的 術科場地1
            TIMS.SetMyValue2(htSS, "COLUMN1", "TECHPLACEID")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'plan_planinfo 的 術科場地2
            TIMS.SetMyValue2(htSS, "COLUMN1", "TECHPLACEID2")
            UPDATE_PLACEID_ALL(objconn, htSS)

            'SELECT OLDDATA14_1,OLDDATA14_2,OLDDATA14_3,OLDDATA14_4 ,NEWDATA14_1,NEWDATA14_2,NEWDATA14_3,NEWDATA14_4 FROM PLAN_REVISE
            'UPDATE - PLAN_REVISE
            TIMS.SetMyValue2(htSS, "TABLE1", "PLAN_REVISE")
            'PLAN_REVISE 的 學科場地1變更前
            TIMS.SetMyValue2(htSS, "COLUMN1", "OLDDATA14_1")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 術科場地1變更前
            TIMS.SetMyValue2(htSS, "COLUMN1", "OLDDATA14_2")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 學科場地2變更前
            TIMS.SetMyValue2(htSS, "COLUMN1", "OLDDATA14_3")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 術科場地2變更前
            TIMS.SetMyValue2(htSS, "COLUMN1", "OLDDATA14_4")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'PLAN_REVISE 的 學科場地1變更後
            TIMS.SetMyValue2(htSS, "COLUMN1", "NEWDATA14_1")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 術科場地1變更後
            TIMS.SetMyValue2(htSS, "COLUMN1", "NEWDATA14_2")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 學科場地2變更後
            TIMS.SetMyValue2(htSS, "COLUMN1", "NEWDATA14_3")
            UPDATE_PLACEID_ALL(objconn, htSS)
            'Plan_Revise 的 術科場地2變更後
            TIMS.SetMyValue2(htSS, "COLUMN1", "NEWDATA14_4")
            UPDATE_PLACEID_ALL(objconn, htSS)
        End If
    End Sub

    ''' <summary>儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        '"儲存"
        'Dim rqRID As String = Request("RID") '業務ID
        'Dim Upload_Path As String = "~/images/Placepic/"
        'Dim FileName1 As String = ""
        'Dim MyFileType As String = ""

        PlaceID.Text = TIMS.ClearSQM(PlaceID.Text)
        Hid_comidno.Value = TIMS.ClearSQM(Hid_comidno.Value)
        rqPTID = TIMS.ClearSQM(rqPTID)

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg & "，請確認")
            Exit Sub
        End If

        Dim dr As DataRow = Nothing
        'Dim sql As String = ""
        If Errmsg = "" Then
            Select Case irqProcessType
                Case eePT_Enum.xInsert
                    'sql = "select * from Plan_TrainPlace join Org_OrgInfo on Org_OrgInfo.ComIDNO= Plan_TrainPlace.ComIDNO where PLACEID='" & PlaceID.Text & "' AND OrgID='" & sm.UserInfo.OrgID & "'"
                    Dim rPMS1 As New Hashtable
                    rPMS1.Add("PLACEID", PlaceID.Text)
                    rPMS1.Add("RID", rqRID)
                    Dim sql1 As String = ""
                    sql1 &= " SELECT 'X' "
                    sql1 &= " FROM PLAN_TRAINPLACE a "
                    sql1 &= " JOIN Org_OrgInfo oo ON oo.ComIDNO = a.ComIDNO "
                    sql1 &= " JOIN Auth_Relship ar ON ar.ORGID = oo.ORGID "
                    sql1 &= " WHERE a.PLACEID=@PLACEID AND ar.RID=@RID"
                    dr = DbAccess.GetOneRow(sql1, objconn, rPMS1)
                    If Not dr Is Nothing Then Errmsg += "場地代碼重複" & vbCrLf
                Case eePT_Enum.xUpdate
                    If Check_Repeat(rqRID, rqPTID) = True Then Errmsg += "場地代碼重複" & vbCrLf
            End Select
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg & "，請確認")
            Exit Sub
        End If

        Dim pPMS2 As New Hashtable
        pPMS2.Add("RID", rqRID)
        Dim sqlstr2 As String = "SELECT oo.ComIDNO FROM ORG_ORGINFO oo JOIN Auth_Relship ar ON oo.OrgID = ar.OrgID WHERE ar.RID=@RID"
        Dim drX2 As DataRow = DbAccess.GetOneRow(sqlstr2, objconn, pPMS2)
        If drX2 Is Nothing Then Errmsg += "查無該機構統編，請重新操作此功能。" & vbCrLf '查無該機構統編
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg & "，請確認")
            Exit Sub
        End If

        Dim pPMS3 As New Hashtable
        pPMS3.Add("RID", rqRID)
        Dim sqlstr3 As String = "SELECT oo.ComIDNO FROM ORG_ORGINFO oo JOIN AUTH_RELSHIP ar ON oo.OrgID = ar.OrgID WHERE ar.RID=@RID"
        drX2 = DbAccess.GetOneRow(sqlstr3, objconn, pPMS3)
        If drX2 Is Nothing Then
            'Hid_comidno.Value = ""
            Errmsg = "查無該機構統編，請重新操作此功能。" & vbCrLf '查無該機構統編
            Common.MessageBox(Me, Errmsg & "，請確認")
            Exit Sub
        Else
            If Hid_comidno.Value <> drX2("ComIDNO").ToString Then
                Errmsg = "機構統編與業務代碼不同，請確認業務權限。" & vbCrLf '查無該機構統編
                Common.MessageBox(Me, Errmsg & "，請確認")
                Exit Sub
            End If
        End If

        Call sSaveData1()

        Dim lrMsg As String = "資料修改成功!"
        If irqProcessType = eePT_Enum.xInsert Then lrMsg = "資料新增成功!"

        'If Request("ProcessType") = "Insert" Then
        '    Common.RespWrite(Me, "<script language=javascript>window.alert('資料新增成功!');")
        'Else
        '    Common.RespWrite(Me, "<script language=javascript>window.alert('資料修改成功!');")
        'End If
        'Common.RespWrite(Me, "window.location.href='TC_01_013.aspx';</script>")
        'sm.LastResultMessage = lrMsg  '=== by:before20180824
        'sm.RedirectUrlAfterBlock = ResolveUrl("~/index.aspx")  '=== by:before20180824
        'sm.RedirectUrlAfterBlock = ResolveUrl("~/TC/01/TC_01_013.aspx")  '=== by:20180824

        Common.RespWrite(Me, String.Concat("<script>alert('", lrMsg, "');location.href='TC_01_013.aspx';</script>"))  '=== by:20180827
    End Sub

    Public Shared Sub UPDATE_PLACEID_ALL(ByRef tConn As SqlConnection, ByRef htSS As Hashtable)
        Dim v_PLACEID As String = TIMS.GetMyValue2(htSS, "PLACEID")
        Dim v_HIDCOMIDNO As String = TIMS.GetMyValue2(htSS, "HIDCOMIDNO")
        Dim v_HIDPLACEID As String = TIMS.GetMyValue2(htSS, "HIDPLACEID")
        Dim v_TABLE1 As String = TIMS.GetMyValue2(htSS, "TABLE1")
        Dim v_COLUMN1 As String = TIMS.GetMyValue2(htSS, "COLUMN1")

        Dim dr2 As DataRow = Nothing
        Dim sql2 As String = String.Concat(" SELECT 1 FROM ", v_TABLE1, " WHERE COMIDNO=@HIDCOMIDNO AND ", v_COLUMN1, "=@HIDPLACEID")
        Dim vParms As New Hashtable
        'vParms.Clear()
        vParms.Add("HIDCOMIDNO", v_HIDCOMIDNO)
        vParms.Add("HIDPLACEID", v_HIDPLACEID)
        dr2 = DbAccess.GetOneRow(sql2, tConn, vParms)
        If dr2 IsNot Nothing Then
            Dim sql22 As String = String.Concat(" UPDATE ", v_TABLE1, " SET ", v_COLUMN1, "=@PLACEID WHERE COMIDNO=@HIDCOMIDNO AND ", v_COLUMN1, "=@HIDPLACEID")
            Dim vParms2 As New Hashtable
            'vParms2.Clear()
            vParms2.Add("PLACEID", v_PLACEID)
            vParms2.Add("HIDCOMIDNO", v_HIDCOMIDNO)
            vParms2.Add("HIDPLACEID", v_HIDPLACEID)
            DbAccess.ExecuteNonQuery(sql22, tConn, vParms2)
        End If
    End Sub

    Function Get_MyFileType1(ByVal s_PlacePIC1 As String) As String
        Dim rst As String = ""
        If s_PlacePIC1 = "" OrElse s_PlacePIC1.IndexOf(".") = -1 Then Return rst
        rst = Split(s_PlacePIC1, ".")((Split(s_PlacePIC1, ".")).Length - 1)
        Return rst
    End Function

End Class
