Partial Class RWB_01_011
    Inherits AuthBasePage

    'Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新確認上傳檔案格式!"
    Const cst_errMsg_3 As String = "檔案位置錯誤!"
    Const cst_errMsg_4 As String = "檔案類型錯誤!"
    Const cst_errMsg_5 As String = "檔案類型錯誤，必須為圖片類型檔案(.jpg)!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗，請刪除此筆後重新上傳)"
    Const cst_PostedFile_max_size As Integer = 10485760
    Const cst_errMsg_7 As String = "檔案大小超過10MB!"
    Const cst_errMsg_8 As String = "請選擇一張圖片上傳!"
    Const cst_errMsg_10 As String = "無效的圖形格式。"

    Const Cst_Upload_Path As String = "~/Upload/BANNER/"
    Const cst_NOFILE As String = "(檔案不存在)"
    Const cst_UPLOADFILE_ERR1 As String = "檔案名稱 已經存在 不可重複上傳" '"(上傳檔案失敗,已存在)"

    '首頁資料BANNER
    'Const cst_Sch0 As String = "Sch0"
    Const cst_Sch1 As String = "Sch1"
    Const cst_UPD1 As String = "UPD1"
    Const cst_DEL1 As String = "DEL1"
    Const cst_ADD1 As String = "ADD1"
    Const cst_ses_SearchStr1 As String = "_ses_SearchStr1"

    Dim oSTART_DATE As Object = Nothing
    Dim oEND_DATE As Object = Nothing

    Const cst_TYPEID_B01 As String = "B01"

    Dim objconn As SqlConnection
    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            '處理[分頁設定元件]出現的時機
            'PageControler1.Visible = False
            Call sCreate1() '頁面初始化
        End If
    End Sub

    '頁面初始化
    Sub sCreate1()
        Call TIMS.SUB_SET_HR_MI(ddl_SDATE_HH, ddl_SDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddl_EDATE_HH, ddl_EDATE_MM)
        'Common.SetListItem(ddlC_SDATE_hh, "00")
        'Common.SetListItem(ddlC_SDATE_mm, "00")
        'Common.SetListItem(ddlC_EDATE_hh, "23")
        'Common.SetListItem(ddlC_EDATE_mm, "59")
        showTable1(cst_Sch1)
    End Sub

    '資料查詢
    Sub sSearch1(ByVal sType As String)
        showTable1(cst_Sch1)

        If sType = cst_Sch1 AndAlso Session(cst_ses_SearchStr1) IsNot Nothing Then
            Dim myValue1 As String = ""
            Dim vSearchStr1 As String = Convert.ToString(Session(cst_ses_SearchStr1))
            START_DATE_S1.Text = TIMS.GetMyValue(vSearchStr1, "START_DATE_S1")
            START_DATE_S2.Text = TIMS.GetMyValue(vSearchStr1, "START_DATE_S2")
            END_DATE_S1.Text = TIMS.GetMyValue(vSearchStr1, "END_DATE_S1")
            END_DATE_S2.Text = TIMS.GetMyValue(vSearchStr1, "END_DATE_S2")
            myValue1 = TIMS.GetMyValue(vSearchStr1, "ISUSE_S")
            Common.SetListItem(rblISUSE_S, myValue1)
            Session(cst_ses_SearchStr1) = Nothing
        End If

        msg1.Text = "查無資料"
        tb_Sch.Visible = False

        START_DATE_S1.Text = TIMS.ClearSQM(START_DATE_S1.Text)
        START_DATE_S2.Text = TIMS.ClearSQM(START_DATE_S2.Text)
        END_DATE_S1.Text = TIMS.ClearSQM(END_DATE_S1.Text)
        END_DATE_S2.Text = TIMS.ClearSQM(END_DATE_S2.Text)
        Dim v_rblISUSE_S As String = TIMS.GetListValue(rblISUSE_S)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT TOP 500 a.BANNERID" & vbCrLf '/*PK*/
        sql &= " ,a.TYPEID" & vbCrLf
        sql &= " ,a.B_TITLE" & vbCrLf
        sql &= " ,a.B_CONTENT" & vbCrLf
        sql &= " ,a.B_URL" & vbCrLf
        sql &= " ,a.B_ALT" & vbCrLf
        sql &= " ,a.FILE_NAME" & vbCrLf
        sql &= " ,a.ORG_FILE_NAME" & vbCrLf
        sql &= " ,format(a.START_DATE,'yyyy/MM/dd HH:mm') START_DATE" & vbCrLf
        sql &= " ,format(a.END_DATE,'yyyy/MM/dd HH:mm') END_DATE" & vbCrLf
        'sql &= " ,FORMAT(a.START_DATE,'HH') SDATE_HH " & vbCrLf
        'sql &= " ,FORMAT(a.START_DATE,'mm') SDATE_MM " & vbCrLf
        'sql &= " ,FORMAT(a.END_DATE,'HH') EDATE_HH " & vbCrLf
        'sql &= " ,FORMAT(a.END_DATE,'mm') EDATE_MM " & vbCrLf
        sql &= " ,a.ISUSED" & vbCrLf
        sql &= " ,case when a.ISUSED='Y' THEN '啟用' ELSE '' END ISUSED_N" & vbCrLf
        'sql &= " ,a.CREATEACCT" & vbCrLf
        'sql &= " ,a.CREATEDATE" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.SEQ" & vbCrLf
        sql &= " FROM dbo.TB_BANNER a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.TYPEID='B01'" & vbCrLf
        sql &= If(START_DATE_S1.Text <> "", " AND a.START_DATE >= convert(date,@START_DATE_S1) ", "") & vbCrLf
        sql &= If(START_DATE_S2.Text <> "", " AND a.START_DATE <= convert(date,@START_DATE_S2) ", "") & vbCrLf
        sql &= If(END_DATE_S1.Text <> "", " AND a.END_DATE >= convert(date,@END_DATE_S1) ", "") & vbCrLf
        sql &= If(END_DATE_S2.Text <> "", " AND a.END_DATE <= convert(date,@END_DATE_S2) ", "") & vbCrLf
        sql &= If(v_rblISUSE_S.Equals("X"), "", If(v_rblISUSE_S.Equals("Y"), " AND ISUSED='Y' ", " AND ISUSED IS NULL ")) & vbCrLf
        sql &= " ORDER BY a.BANNERID DESC " & vbCrLf

        Dim parms As New Hashtable()
        parms.Clear()
        If (START_DATE_S1.Text <> "") Then parms.Add("START_DATE_S1", START_DATE_S1.Text)
        If (START_DATE_S2.Text <> "") Then parms.Add("START_DATE_S2", START_DATE_S2.Text)
        If (END_DATE_S1.Text <> "") Then parms.Add("END_DATE_S1", END_DATE_S1.Text)
        If (END_DATE_S2.Text <> "") Then parms.Add("END_DATE_S2", END_DATE_S2.Text)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        msg1.Text = "查無資料"
        tb_Sch.Visible = False
        If dt.Rows.Count = 0 Then Exit Sub

        msg1.Text = ""
        tb_Sch.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Sub ShowData1(ByVal iBANNERID As Integer)
        Call CLS_DATA1()

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT  TOP 500 a.BANNERID" & vbCrLf '/*PK*/
        sql &= " ,a.TYPEID" & vbCrLf
        sql &= " ,a.B_TITLE" & vbCrLf
        sql &= " ,a.B_CONTENT" & vbCrLf
        sql &= " ,a.B_URL" & vbCrLf
        sql &= " ,a.B_ALT" & vbCrLf
        sql &= " ,a.FILE_NAME" & vbCrLf
        sql &= " ,a.ORG_FILE_NAME" & vbCrLf
        sql &= " ,format(a.START_DATE,'yyyy/MM/dd') START_DATE" & vbCrLf
        sql &= " ,format(a.END_DATE,'yyyy/MM/dd') END_DATE" & vbCrLf
        sql &= " ,FORMAT(a.START_DATE,'HH') SDATE_HH " & vbCrLf
        sql &= " ,FORMAT(a.START_DATE,'mm') SDATE_MM " & vbCrLf
        sql &= " ,FORMAT(a.END_DATE,'HH') EDATE_HH " & vbCrLf
        sql &= " ,FORMAT(a.END_DATE,'mm') EDATE_MM " & vbCrLf
        sql &= " ,a.ISUSED" & vbCrLf
        sql &= " ,case when a.ISUSED='Y' THEN '啟用' ELSE '' END ISUSED_N" & vbCrLf
        'sql &= " ,a.CREATEACCT" & vbCrLf
        'sql &= " ,a.CREATEDATE" & vbCrLf
        'sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.SEQ" & vbCrLf
        sql &= " FROM dbo.TB_BANNER a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.TYPEID='B01'" & vbCrLf
        sql &= " AND a.BANNERID=@BANNERID" & vbCrLf
        Dim parms As New Hashtable()
        parms.Clear()
        parms.Add("BANNERID", iBANNERID)

        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr Is Nothing Then Exit Sub
        Hid_BANNERID.Value = Convert.ToString(dr("BANNERID"))
        Hid_TYPEID.Value = Convert.ToString(dr("TYPEID"))
        B_TITLE.Text = Convert.ToString(dr("B_TITLE"))
        Hid_B_CONTENT.Value = Convert.ToString(dr("B_CONTENT"))
        'B_CONTENT.Text = Convert.ToString(dr("B_CONTENT"))
        B_URL.Text = Convert.ToString(dr("B_URL"))
        B_ALT.Text = Convert.ToString(dr("B_ALT"))
        FILE_NAME.Text = Convert.ToString(dr("FILE_NAME"))
        Hid_ORG_FILE_NAME.Value = Convert.ToString(dr("ORG_FILE_NAME"))

        Dim s_Upload_BANNER_Path As String = TIMS.Utl_GetConfigSet("Upload_BANNER_Path")
        Dim s_Upload_Path As String = ""
        s_Upload_Path = If(s_Upload_BANNER_Path <> "", s_Upload_BANNER_Path, Cst_Upload_Path)

        Dim s_EXISTFILE As String = String.Format("(檔案:{0},存在)", FILE_NAME.Text)
        Dim flag_fileExiss As Boolean = TIMS.CHK_PIC_EXISTS(Server, s_Upload_Path, FILE_NAME.Text)
        LabFNAME_MSG.Text = If(flag_fileExiss, s_EXISTFILE, cst_NOFILE)

        START_DATE.Text = Convert.ToString(dr("START_DATE"))
        Common.SetListItem(ddl_SDATE_HH, Convert.ToString(dr("SDATE_HH")))
        Common.SetListItem(ddl_SDATE_MM, Convert.ToString(dr("SDATE_MM")))

        END_DATE.Text = Convert.ToString(dr("END_DATE"))
        Common.SetListItem(ddl_EDATE_HH, Convert.ToString(dr("EDATE_HH")))
        Common.SetListItem(ddl_EDATE_MM, Convert.ToString(dr("EDATE_MM")))

        cb_ISUSED.Checked = If(Convert.ToString(dr("ISUSED")) = "Y", True, False)
        TXT_SEQ.Text = Convert.ToString(dr("SEQ"))

        showTable1(cst_UPD1)
    End Sub

    ''' <summary>
    ''' 畫面狀況調整
    ''' </summary>
    ''' <param name="sType"></param>
    Sub showTable1(ByVal sType As String)
        tb_SchV.Visible = True
        tb_Sch.Visible = False
        tb_Edit1.Visible = False
        'Case cst_Sch1
        Select Case sType
            Case cst_UPD1
                tb_Edit1.Visible = True
                tb_SchV.Visible = False
                tb_Sch.Visible = False
            Case cst_ADD1
                tb_Edit1.Visible = True
                tb_SchV.Visible = False
                tb_Sch.Visible = False
        End Select
    End Sub

    Sub CLS_DATA1()
        Hid_BANNERID.Value = "" ' Convert.ToString(dr("BANNERID"))
        Hid_TYPEID.Value = "" 'Convert.ToString(dr("TYPEID"))
        B_TITLE.Text = "" 'Convert.ToString(dr("B_TITLE"))
        'B_CONTENT.Text = "" 'Convert.ToString(dr("B_CONTENT"))
        Hid_B_CONTENT.Value = "" 'Convert.ToString(dr("B_CONTENT"))

        B_URL.Text = "" 'Convert.ToString(dr("B_URL"))
        B_ALT.Text = "" 'Convert.ToString(dr("B_ALT"))
        FILE_NAME.Text = "" 'Convert.ToString(dr("FILE_NAME"))
        Hid_ORG_FILE_NAME.Value = ""
        LabFNAME_MSG.Text = cst_NOFILE
        START_DATE.Text = "" 'Convert.ToString(dr("START_DATE"))
        END_DATE.Text = "" 'Convert.ToString(dr("END_DATE"))
        Call TIMS.SUB_SET_HR_MI(ddl_SDATE_HH, ddl_SDATE_MM)
        Call TIMS.SUB_SET_HR_MI(ddl_EDATE_HH, ddl_EDATE_MM)
        cb_ISUSED.Checked = True 'If(Convert.ToString(dr("ISUSED")) = "Y", True, False)
        TXT_SEQ.Text = "" 'Convert.ToString(dr("SEQ"))
    End Sub

    Sub SaveData1()
        Dim rst As Integer = 0
        Dim flagSaveOK1 As Boolean = False

        Dim sType As String = cst_ADD1
        If Hid_BANNERID.Value <> "" Then sType = cst_UPD1 '修改
        Dim parms As Hashtable = New Hashtable
        'Dim iBANNERID As Integer = 0
        Dim vORG_FILE_NAME As String = FILE_NAME.Text

        oSTART_DATE = TIMS.Cdate3(START_DATE.Text) & " " & TIMS.GetListValue(ddl_SDATE_HH) & ":" & TIMS.GetListValue(ddl_SDATE_MM)
        oEND_DATE = TIMS.Cdate3(END_DATE.Text) & " " & TIMS.GetListValue(ddl_EDATE_HH) & ":" & TIMS.GetListValue(ddl_EDATE_MM)

        If sType = cst_ADD1 Then
            '新增
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " INSERT INTO TB_BANNER(" & vbCrLf
            sql &= " BANNERID,TYPEID,B_TITLE,B_CONTENT,B_URL,B_ALT" & vbCrLf
            sql &= " ,FILE_NAME,ORG_FILE_NAME,START_DATE,END_DATE" & vbCrLf
            sql &= " ,ISUSED,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE,SEQ" & vbCrLf
            sql &= " ) VALUES (" & vbCrLf
            sql &= " @BANNERID,@TYPEID,@B_TITLE,@B_CONTENT,@B_URL,@B_ALT" & vbCrLf
            'sql &= " ,@FILE_NAME,@ORG_FILE_NAME,CONVERT(DATE,@START_DATE),CONVERT(DATE,@END_DATE)" & vbCrLf
            sql &= " ,@FILE_NAME,@ORG_FILE_NAME,@START_DATE,@END_DATE" & vbCrLf
            sql &= " ,@ISUSED,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE(),@SEQ" & vbCrLf
            sql &= " )" & vbCrLf

            Dim iBANNERID As Integer = DbAccess.GetNewId(objconn, "TB_BANNER_BANNERID_SEQ,TB_BANNER,BANNERID")
            parms.Clear()
            parms.Add("BANNERID", iBANNERID)
            parms.Add("TYPEID", cst_TYPEID_B01)

            parms.Add("B_TITLE", B_TITLE.Text)
            Hid_B_CONTENT.Value = B_TITLE.Text
            parms.Add("B_CONTENT", Hid_B_CONTENT.Value)
            parms.Add("B_URL", B_URL.Text)
            parms.Add("B_ALT", B_ALT.Text)

            parms.Add("FILE_NAME", FILE_NAME.Text)
            parms.Add("ORG_FILE_NAME", vORG_FILE_NAME)
            'parms.Add("START_DATE", START_DATE.Text)
            'parms.Add("END_DATE", END_DATE.Text)
            parms.Add("START_DATE", CDate(oSTART_DATE))
            parms.Add("END_DATE", CDate(oEND_DATE))

            parms.Add("ISUSED", If(cb_ISUSED.Checked, "Y", Convert.DBNull)) '啟用YES 停用 NULL
            parms.Add("CREATEACCT", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("SEQ", Val(TXT_SEQ.Text))
            rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
            flagSaveOK1 = True
        Else
            '修改
            Dim iBANNERID As Integer = Val(Hid_BANNERID.Value)
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " UPDATE TB_BANNER" & vbCrLf 'SET
            'Sql &= " BANNERID=@BANNERID" & vbCrLf 'Sql &= " ,TYPEID=@TYPEID" & vbCrLf
            sql &= " SET B_TITLE=@B_TITLE" & vbCrLf
            sql &= " ,B_CONTENT=@B_CONTENT" & vbCrLf
            sql &= " ,B_URL=@B_URL" & vbCrLf
            sql &= " ,B_ALT=@B_ALT" & vbCrLf
            sql &= " ,FILE_NAME=@FILE_NAME" & vbCrLf
            sql &= " ,ORG_FILE_NAME=@ORG_FILE_NAME" & vbCrLf
            'sql &= " ,START_DATE= CONVERT(DATE,@START_DATE) " & vbCrLf
            'sql &= " ,END_DATE= CONVERT(DATE,@END_DATE)" & vbCrLf
            sql &= " ,START_DATE= @START_DATE " & vbCrLf
            sql &= " ,END_DATE= @END_DATE" & vbCrLf
            sql &= " ,ISUSED=@ISUSED" & vbCrLf
            'Sql &= " ,CREATEACCT=@CREATEACCT" & vbCrLf
            'Sql &= " ,CREATEDATE=@CREATEDATE" & vbCrLf
            sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            sql &= " ,SEQ=@SEQ" & vbCrLf
            sql &= " WHERE 1=1" & vbCrLf
            sql &= " AND BANNERID=@BANNERID" & vbCrLf

            parms.Clear()
            parms.Add("TYPEID", cst_TYPEID_B01)
            parms.Add("B_TITLE", B_TITLE.Text)
            Hid_B_CONTENT.Value = B_TITLE.Text
            parms.Add("B_CONTENT", Hid_B_CONTENT.Value)
            parms.Add("B_URL", B_URL.Text)
            parms.Add("B_ALT", B_ALT.Text)

            parms.Add("FILE_NAME", FILE_NAME.Text)
            parms.Add("ORG_FILE_NAME", vORG_FILE_NAME)
            'parms.Add("START_DATE", START_DATE.Text)
            'parms.Add("END_DATE", END_DATE.Text)
            parms.Add("START_DATE", CDate(oSTART_DATE))
            parms.Add("END_DATE", CDate(oEND_DATE))

            parms.Add("ISUSED", If(cb_ISUSED.Checked, "Y", Convert.DBNull))
            'parms.Add("CREATEACCT", sm.UserInfo.UserID)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("SEQ", Val(TXT_SEQ.Text))

            parms.Add("BANNERID", iBANNERID)
            rst = DbAccess.ExecuteNonQuery(sql, objconn, parms)
            flagSaveOK1 = True
        End If

        If Not flagSaveOK1 Then '儲存-失敗
            Common.MessageBox(Me, "儲存失敗!")
            Exit Sub
        End If

        '儲存成功
        Common.MessageBox(Me, "儲存成功!")
        Call sSearch1(cst_Sch1)
        'TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'DataGrid1功能事件
        If e.CommandArgument = "" Then Exit Sub
        Dim eCmd As String = TIMS.ClearSQM(e.CommandName)
        If eCmd = "" Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        Hid_BANNERID.Value = TIMS.ClearSQM(TIMS.GetMyValue(sCmdArg, "BANNERID"))
        If Hid_BANNERID.Value = "" Then Exit Sub
        Dim iBANNERID As Integer = Val(Hid_BANNERID.Value)

        GetSearchStr()  'edit，by:20190103
        Select Case eCmd
            Case cst_DEL1 '刪除
                Common.MessageBox(Me, "暫不提供刪除功能!!")
                Exit Sub
            Case cst_UPD1 '修改
                Call ShowData1(iBANNERID)
            Case Else
                Common.MessageBox(Me, "查無資料!!:" & eCmd)
                Exit Sub
        End Select
    End Sub

    Sub GetSearchStr()
        'Dim v_rblISUse As String = TIMS.GetListValue(rblISUse)
        Dim vSearchStr1 As String = ""
        vSearchStr1 &= "&START_DATE_S1=" & TIMS.ClearSQM(START_DATE_S1.Text)
        vSearchStr1 &= "&START_DATE_S2=" & TIMS.ClearSQM(START_DATE_S2.Text)
        vSearchStr1 &= "&END_DATE_S1=" & TIMS.ClearSQM(END_DATE_S1.Text)
        vSearchStr1 &= "&END_DATE_S2=" & TIMS.ClearSQM(END_DATE_S2.Text)
        vSearchStr1 &= "&ISUSE_S=" & TIMS.GetListValue(rblISUSE_S)
        Session(cst_ses_SearchStr1) = vSearchStr1
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnUPD1 As Button = e.Item.FindControl("btnUPD1")
                'Dim btnDEL1 As Button = e.Item.FindControl("btnDEL1")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BANNERID", TIMS.CStr1(drv("BANNERID")))
                btnUPD1.CommandArgument = sCmdArg
                'btnDEL1.CommandArgument = sCmdArg
                'btnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg2
        End Select
    End Sub

    '查詢作業
    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Call sSearch1("") '查詢
    End Sub

    '新增
    Protected Sub btnAdd1_Click(sender As Object, e As EventArgs) Handles btnAdd1.Click
        GetSearchStr()  'edit，by:20190103
        showTable1(cst_ADD1)
        CLS_DATA1()
    End Sub

    ''' <summary>
    ''' 檢核 檔案名稱 是否有重複
    ''' </summary>
    ''' <param name="sBANNERID"></param>
    ''' <param name="FILE_NAME"></param>
    ''' <returns></returns>
    Function CHK_DOUBLE_FILENAME(ByRef sBANNERID As String, ByRef FILE_NAME As String) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("FILE_NAME", FILE_NAME)
        sql = ""
        sql &= " SELECT 'x'" & vbCrLf
        sql &= " FROM dbo.TB_BANNER a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND a.TYPEID='B01'" & vbCrLf
        sql &= " AND a.FILE_NAME=@FILE_NAME" & vbCrLf
        If sBANNERID <> "" Then
            parms.Add("BANNERID", sBANNERID)
            sql &= " AND a.BANNERID!=@BANNERID" & vbCrLf
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    '送出前檢核 ---> SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        B_TITLE.Text = TIMS.ClearSQM(B_TITLE.Text)
        'B_CONTENT.Text = TIMS.ClearSQM(B_CONTENT.Text)
        B_URL.Text = TIMS.ClearSQM(B_URL.Text)
        B_ALT.Text = TIMS.ClearSQM(B_ALT.Text)

        FILE_NAME.Text = TIMS.ClearSQM(FILE_NAME.Text)
        'Hid_ORG_FILE_NAME.Value = ""
        START_DATE.Text = TIMS.Cdate3(TIMS.ClearSQM(START_DATE.Text))
        END_DATE.Text = TIMS.Cdate3(TIMS.ClearSQM(END_DATE.Text))
        TXT_SEQ.Text = TIMS.ClearSQM(TXT_SEQ.Text)

        If START_DATE.Text = "" Then Errmsg &= "起始日期 不可為空" & vbCrLf
        If END_DATE.Text = "" Then Errmsg &= "結束日期 不可為空" & vbCrLf

        If B_TITLE.Text = "" Then Errmsg &= "抬頭 不可為空" & vbCrLf
        'If B_CONTENT.Text = "" Then Errmsg &= "內容 不可為空" & vbCrLf
        If B_URL.Text = "" Then Errmsg &= "連線網頁 不可為空" & vbCrLf
        If B_URL.Text <> "" Then
            If Not B_URL.Text.ToLower.StartsWith("https://") AndAlso Not B_URL.Text.ToLower.StartsWith("http://") Then
                TIMS.LOG.Debug("連線網頁 請輸入開頭為https 或 http 的 網址:" & B_URL.Text)
                Errmsg &= "連線網頁 請輸入開頭為https 或 http 的 網址" & vbCrLf
            End If
        End If

        If B_ALT.Text = "" Then Errmsg &= "提示訊息 不可為空" & vbCrLf

        If FILE_NAME.Text = "" Then Errmsg &= "檔名 不可為空" & vbCrLf
        If FILE_NAME.Text <> "" AndAlso Not FILE_NAME.Text.ToLower.EndsWith(".jpg") Then Errmsg &= "檔名有誤 請輸入結尾為.jpg的檔名" & vbCrLf

        If TXT_SEQ.Text = "" Then Errmsg &= "順序 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False

        oSTART_DATE = TIMS.Cdate3(START_DATE.Text) & " " & TIMS.GetListValue(ddl_SDATE_HH) & ":" & TIMS.GetListValue(ddl_SDATE_MM)
        oEND_DATE = TIMS.Cdate3(END_DATE.Text) & " " & TIMS.GetListValue(ddl_EDATE_HH) & ":" & TIMS.GetListValue(ddl_EDATE_MM)
        If Not TIMS.IsDate1(oSTART_DATE) Then Errmsg &= "起始日期 日期格式有誤" & vbCrLf
        If Not TIMS.IsDate1(oEND_DATE) Then Errmsg &= "結束日期 日期格式有誤" & vbCrLf

        Hid_BANNERID.Value = TIMS.ClearSQM(Hid_BANNERID.Value)
        Dim flag_double As Boolean = CHK_DOUBLE_FILENAME(Hid_BANNERID.Value, FILE_NAME.Text)
        If flag_double Then Errmsg &= "檔案名稱 已經存在 不可重複上傳" & vbCrLf

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        SaveData1()
    End Sub

    '取消
    Protected Sub btnCancle1_Click(sender As Object, e As EventArgs) Handles btnCancle1.Click
        If Hid_BANNERID.Value <> "" Then
            '修改中
            Hid_BANNERID.Value = TIMS.ClearSQM(Hid_BANNERID.Value)
            If Hid_BANNERID.Value = "" Then Exit Sub
            Dim iBANNERID As Integer = Val(Hid_BANNERID.Value)
            Call ShowData1(iBANNERID)
        Else
            '新增中
            showTable1(cst_ADD1)
            CLS_DATA1()
        End If
    End Sub

    '回上頁-sch
    Protected Sub btnBack1_Click(sender As Object, e As EventArgs) Handles btnBack1.Click
        CLS_DATA1()
        Call sSearch1(cst_Sch1)
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Protected Sub ButUpload1_Click(sender As Object, e As EventArgs) Handles ButUpload1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim Upload_Path As String = "~/images/Placepic/"
        Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
        Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)

        FILE_NAME.Text = Hid_ORG_FILE_NAME.Value
        LabFNAME_MSG.Text = ""

        'FILE_NAME.Enabled = True
        If LCase(MyPostedFile.ContentType.ToString()).IndexOf("image") < 0 Then
            LabFNAME_MSG.Text = cst_errMsg_10
            'Common.MessageBox(Me, cst_errMsg_10)
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If
        ' GetThumbNail(MyPostedFile.FileName, 320, 240, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream)
        If File1.Value = "" Then
            LabFNAME_MSG.Text = cst_errMsg_8
            'Common.MessageBox(Me, cst_errMsg_8)
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            LabFNAME_MSG.Text = cst_errMsg_3
            'Common.MessageBox(Me, cst_errMsg_3)
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If
        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            LabFNAME_MSG.Text = cst_errMsg_4
            'Common.MessageBox(Me, cst_errMsg_4)
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "jpg"
                '檢查檔案格式與大小 End
                If File1.PostedFile.ContentLength > cst_PostedFile_max_size Then
                    LabFNAME_MSG.Text = cst_errMsg_7
                    'Common.MessageBox(Me, cst_errMsg_7)
                    Common.MessageBox(Me, LabFNAME_MSG.Text)
                    Exit Sub
                End If
            Case Else
                LabFNAME_MSG.Text = cst_errMsg_5
                'Common.MessageBox(Me, cst_errMsg_5)
                Common.MessageBox(Me, LabFNAME_MSG.Text)
                Exit Sub
        End Select

        '上傳檔案
        'FileName1 = "P" & Request("PTID") & "_" & depID.SelectedItem.Value & "." & MyFileType
        'GUIDfilename = Guid.NewGuid.ToString & "." & MyFileType
        'File1.PostedFile.SaveAs(Server.MapPath(Upload_Path & GUIDfilename))

        'Dim filename As String = MyPostedFile.FileName
        '上傳檔案/存檔：檔名
        'Dim GUIDfilename As String = ""

        'Dim Errmsg As String = ""
        Hid_BANNERID.Value = TIMS.ClearSQM(Hid_BANNERID.Value)
        Dim flag_double As Boolean = CHK_DOUBLE_FILENAME(Hid_BANNERID.Value, MyFileName)
        If flag_double Then
            LabFNAME_MSG.Text = cst_UPLOADFILE_ERR1 '""
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If
        'If flag_double Then Errmsg &= "檔案名稱 已經存在 不可重複上傳" & vbCrLf
        'If Errmsg <> "" Then
        '    LabFNAME_MSG.Text = cst_UPLOADFILE_ERR1 '""
        '    Common.MessageBox(Me, Errmsg)
        '    Exit Sub
        'End If

        'Const Cst_Upload_Path As String = "~/Upload/BANNER/"
        Dim s_Upload_BANNER_Path As String = TIMS.Utl_GetConfigSet("Upload_BANNER_Path")
        Dim s_Upload_Path As String = ""
        s_Upload_Path = If(s_Upload_BANNER_Path <> "", s_Upload_BANNER_Path, Cst_Upload_Path)

        Dim s_EXISTFILE As String = String.Format("(檔案:{0},已存在，不可重複上傳)", MyFileName)
        Dim flag_fileExiss As Boolean = TIMS.CHK_PIC_EXISTS(Server, s_Upload_Path, MyFileName)
        If flag_fileExiss Then
            LabFNAME_MSG.Text = s_EXISTFILE '""
            Common.MessageBox(Me, LabFNAME_MSG.Text)
            Exit Sub
        End If

        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, s_Upload_Path, MyFileName)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)

            Common.MessageBox(Me, cst_errMsg_2)
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = "" & vbCrLf
            strErrmsg &= "MyPostedFile.FileName: " & MyPostedFile.FileName & vbCrLf
            strErrmsg &= "MyPostedFile.ContentType: " & MyPostedFile.ContentType.ToString() & vbCrLf
            strErrmsg &= "s_Upload_Path: " & s_Upload_Path & vbCrLf
            'Server.MapPath(Upload_Path & filename)
            strErrmsg &= "MyFileName: " & MyFileName & vbCrLf
            strErrmsg &= "Server.MapPath(Cst_Upload_Path & MyFileName): " & Server.MapPath(Cst_Upload_Path & MyFileName) & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        FILE_NAME.Text = MyFileName
        FILE_NAME.Enabled = False
        LabFNAME_MSG.Text = String.Format("(已上傳:{0})", MyFileName)

        'If Session("dtPIC") Is Nothing Then
        '    sm.LastErrorMessage = cst_errMsg_1
        '    Exit Sub
        'End If

        'Try
        '    '此時 Session("dtPIC") 不管新增或修改都有值了(此動作必為新增)
        '    dt = Session("dtPIC")
        '    Dim dr As DataRow = dt.NewRow
        '    dt.Rows.Add(dr)
        '    'dr("PTID") = Me.depID.SelectedValue
        '    dr("depID") = Me.depID.SelectedValue
        '    dr("PlacePIC1") = GUIDfilename
        '    dr("okflag") = GUIDfilename
        '    Session("dtPIC") = dt
        '    ShowPTIDDesc()
        '    'CeratePTIDDesc()
        '    'ShowPTIDDesc()
        '    'CreatePICDT()
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)

        '    Dim strErrmsg As String = "" & vbCrLf
        '    'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
        '    'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    TIMS.WriteTraceLog(Me, ex, strErrmsg)
        '    Exit Sub
        '    'Throw ex
        'End Try
    End Sub
End Class