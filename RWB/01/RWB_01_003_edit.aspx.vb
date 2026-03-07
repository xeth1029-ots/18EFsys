Partial Class RWB_01_003_edit
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim myFile_Path1 As String = ""
    Dim myFile_Path2 As String = ""
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'ConfigurationManager.AppSettings("UploadPath") + "/DLFILE/"
        myFile_Path1 = TIMS.Utl_GetConfigSet("UploadPath") + "/DLFILE/"
        myFile_Path2 = TIMS.Utl_GetConfigSet("UploadPath") + "/DLFILE/"

        If Not IsPostBack Then Call sCreate1() '頁面初始化
    End Sub

    '頁面初始化
    Sub sCreate1()
        ddlType.SelectedValue = "1"

        ddlPlan.Items.Clear()
        ddlPlan.Items.Add(New ListItem("產業人才投資方案", "1"))
        ddlPlan.Items.Add(New ListItem("自辦在職訓練", "2"))
        ddlPlan.Items.Add(New ListItem("企業委託訓練", "3"))
        ddlPlan.Items.Add(New ListItem("充電起飛", "4"))
        ddlPlan.SelectedIndex = 0
        ddlPlan.Enabled = True

        txtCDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(DateTime.Now.ToString("yyyy/MM/dd")), DateTime.Now.ToString("yyyy/MM/dd"))  'edit，by:20181019

        ddlC_SDATE_hh1.Items.Clear()
        ddlC_EDATE_hh1.Items.Clear()
        For i As Integer = 0 To 23
            ddlC_SDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
            ddlC_EDATE_hh1.Items.Add(New ListItem(i.ToString.PadLeft(2, "0"), i.ToString.PadLeft(2, "0")))
        Next

        ddlC_SDATE_mm1.Items.Clear()
        ddlC_EDATE_mm1.Items.Clear()
        For j As Integer = 0 To 59
            ddlC_SDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
            ddlC_EDATE_mm1.Items.Add(New ListItem(j.ToString.PadLeft(2, "0"), j.ToString.PadLeft(2, "0")))
        Next

        Common.SetListItem(ddlC_SDATE_hh1, "00")
        Common.SetListItem(ddlC_SDATE_mm1, "00")
        Common.SetListItem(ddlC_EDATE_hh1, "23")
        Common.SetListItem(ddlC_EDATE_mm1, "59")

        If TIMS.ClearSQM(Request("A")) = "E" Then
            Dim rSEQNO_E As String = TIMS.DecryptAes(TIMS.ClearSQM(Request("SEQNO_E")))
            Dim rSEQNO As String = TIMS.ClearSQM(Request("DLID"))
            If rSEQNO_E <> "" AndAlso rSEQNO_E = rSEQNO Then hid_V.Value = rSEQNO
            If hid_V.Value <> "" Then LoadData1(Val(hid_V.Value))
        End If
    End Sub

    '進行[(資料下載)類別]異動事件
    Protected Sub ddlType_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlType.SelectedIndexChanged
        resetDDL()
    End Sub

    '實際[(資料下載)類別]異動事件
    Sub resetDDL()
        If ddlType.SelectedValue = "1" Then
            ddlPlan.Items.Clear()
            ddlPlan.Items.Add(New ListItem("產業人才投資方案", "1"))
            ddlPlan.Items.Add(New ListItem("自辦在職訓練", "2"))
            ddlPlan.Items.Add(New ListItem("企業委託訓練", "3"))
            ddlPlan.Items.Add(New ListItem("充電起飛", "4"))
            ddlPlan.SelectedIndex = 0
            ddlPlan.Enabled = True
        Else
            ddlPlan.Items.Clear()
            ddlPlan.Items.Add(New ListItem("[無計畫內容]", ""))
            ddlPlan.SelectedIndex = 0
            ddlPlan.Enabled = False
        End If
    End Sub

    '資料讀取
    Private Sub LoadData1(ByVal iSEQNO As Integer)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.DLID DESC) AS ROWNUM " & vbCrLf
        sql &= "        ,FORMAT(a.START_DATE, 'yyyy-MM-dd') CSDATE " & vbCrLf
        sql &= "        ,FORMAT(a.END_DATE, 'yyyy-MM-dd') CEDATE " & vbCrLf
        sql &= "        ,FORMAT(a.UPLOADDATE, 'yyyy-MM-dd') CCDATE " & vbCrLf
        sql &= "        ,FORMAT(a.MODIFYDATE, 'yyyy-MM-dd') CMDATE " & vbCrLf
        sql &= "        ,FORMAT(a.START_DATE, 'HH') CSDATEHH " & vbCrLf
        sql &= "        ,FORMAT(a.END_DATE, 'HH') CEDATEHH " & vbCrLf
        sql &= "        ,FORMAT(a.START_DATE, 'mm') CSDATEMM " & vbCrLf
        sql &= "        ,FORMAT(a.END_DATE, 'mm') CEDATEMM " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.START_DATE, 111) CSDATED " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.END_DATE, 111) CEDATED " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.UPLOADDATE, 111) CCDATED " & vbCrLf
        sql &= "        ,CONVERT(VARCHAR, a.MODIFYDATE, 111) CMDATED " & vbCrLf
        sql &= "        ,a.DLID " & vbCrLf
        sql &= "        ,a.KINDID " & vbCrLf
        sql &= "        ,a.PLANID " & vbCrLf
        sql &= "        ,a.START_DATE " & vbCrLf
        sql &= "        ,a.END_DATE " & vbCrLf
        sql &= "        ,a.DLTITLE " & vbCrLf
        sql &= "        ,CASE WHEN LEN(a.DLTITLE) > 20 THEN SUBSTRING(a.DLTITLE, 1, 20) + '...' ELSE a.DLTITLE END TITLE1 " & vbCrLf
        sql &= "        ,a.UPLOADDATE " & vbCrLf
        sql &= "        ,a.ISUSED " & vbCrLf
        sql &= "        ,a.MEMO " & vbCrLf
        sql &= "        ,a.MODIFYACCT " & vbCrLf
        sql &= "        ,a.MODIFYDATE " & vbCrLf
        sql &= "        ,a.FILE1_NAME " & vbCrLf
        sql &= "        ,a.FILE1_NAME " & vbCrLf
        sql &= "        ,a.FILE1_SIZE " & vbCrLf
        sql &= "        ,a.FILE1_TYPE " & vbCrLf
        sql &= "        ,a.FILE2_NAME " & vbCrLf
        sql &= "        ,a.FILE2_SIZE " & vbCrLf
        sql &= "        ,a.FILE2_TYPE " & vbCrLf
        sql &= " FROM TB_DLFILE a " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "       AND a.ISUSED = 'Y' " & vbCrLf
        sql &= "       AND a.DLID = @DLID " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If Convert.ToString(iSEQNO) <> "" Then parms.Add("DLID", iSEQNO)
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Dim tType As String = Convert.ToString(dr("KINDID")).Trim
            Dim tPlan As String = ""
            ddlType.SelectedValue = tType
            resetDDL()
            If tType.Equals("1") Then
                tPlan = Convert.ToString(dr("PLANID")).Trim
                ddlPlan.SelectedValue = tPlan
            End If
            txtCDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(Convert.ToString(dr("CCDATED"))), Convert.ToString(dr("CCDATED")))  'edit，by:20181019
            txtSDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(Convert.ToString(dr("CSDATED"))), Convert.ToString(dr("CSDATED")))  'edit，by:20181019
            Common.SetListItem(ddlC_SDATE_hh1, Convert.ToString(dr("CSDATEHH")))
            Common.SetListItem(ddlC_SDATE_mm1, Convert.ToString(dr("CSDATEMM")))
            txtEDATE1.Text = IIf(flag_ROC, TIMS.Cdate17(Convert.ToString(dr("CEDATED"))), Convert.ToString(dr("CEDATED")))  'edit，by:20181019
            Common.SetListItem(ddlC_EDATE_hh1, Convert.ToString(dr("CEDATEHH")))
            Common.SetListItem(ddlC_EDATE_mm1, Convert.ToString(dr("CEDATEMM")))
            Dim myTitle As String = Convert.ToString(dr("DLTITLE"))
            txtTitle.Text = myTitle

            '取得附件相關資訊 ======== Start
#Region "取得附件相關資訊"

            Dim F1_Name As String = Convert.ToString(dr("FILE1_NAME"))
            Dim F1_Type As String = Convert.ToString(dr("FILE1_TYPE"))
            Dim F1_Size As String = Convert.ToString(dr("FILE1_SIZE"))
            If F1_Type <> "" And F1_Size <> "" Then
                lnkF1Name.Text = myTitle.Trim + F1_Type.Trim.ToLower
                lnkF1Name.ToolTip = "[ 附檔名：" + F1_Type.Trim.ToLower.Substring(1) + "  /  檔案大小：" + F1_Size + " ]"
                lblF1Name.Text = F1_Name.Trim
                lblF1Ext.Text = F1_Type.Trim.ToUpper
                divFile1.Visible = True
            Else
                divFile1.Visible = False
            End If

            Dim F2_Name As String = Convert.ToString(dr("FILE2_NAME"))
            Dim F2_Type As String = Convert.ToString(dr("FILE2_TYPE"))
            Dim F2_Size As String = Convert.ToString(dr("FILE2_SIZE"))
            If F2_Type <> "" And F2_Size <> "" Then
                lnkF2Name.Text = myTitle.Trim + F2_Type.Trim.ToLower
                lnkF2Name.ToolTip = "[ 附檔名：" + F2_Type.Trim.ToLower.Substring(1) + "  /  檔案大小：" + F2_Size + " ]"
                lblF2Name.Text = F2_Name.Trim
                lblF2Ext.Text = F2_Type.Trim.ToUpper
                divFile2.Visible = True
            Else
                divFile2.Visible = False
            End If

#End Region
            '====================== End
        Else
            Exit Sub
        End If
    End Sub

    '資料儲存
    Protected Sub bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call SaveData1()
    End Sub

    '送出前檢核 ---> SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        If ddlType.SelectedIndex = -1 Then Errmsg &= "類別 不可為空" & vbCrLf
        If ddlPlan.SelectedIndex = -1 Then Errmsg &= "計畫 不可為空" & vbCrLf

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)  'edit，by:20181019
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)  'edit，by:20181019
        Dim mySDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtTitle.Text = TIMS.ClearSQM(txtTitle.Text)
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019

        If txtSDATE1.Text = "" OrElse TIMS.CStr1(oC_SDATE) = "" Then Errmsg &= "上架日期 不可為空" & vbCrLf
        If txtTitle.Text = "" Then Errmsg &= "標題 不可為空" & vbCrLf
        If txtEDATE1.Text = "" OrElse TIMS.CStr1(oC_EDATE) = "" Then Errmsg &= "停用日期 不可為空" & vbCrLf
        If Errmsg <> "" Then Return False

        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) = 0 Then Errmsg &= "上架日期與停用日期 不可相等!!" & vbCrLf
        If DateDiff(DateInterval.Minute, CDate(oC_SDATE), CDate(oC_EDATE)) < 0 Then Errmsg &= "上架日期與停用日期 順序異常!!" & vbCrLf
        If Errmsg <> "" Then Return False

        '附件檢核 ============================ Start
#Region "[附件檢核]"

        Dim blHasNewF1 As Boolean = fu1.HasFile          '本次上傳的[附件1]是否含檔案
        Dim blHasOldF1 As Boolean = divFile1.Visible     '先前上傳的[附件1]是否含檔案
        Dim blHasNewF2 As Boolean = fu2.HasFile          '本次上傳的[附件2]是否含檔案
        Dim blHasOldF2 As Boolean = divFile2.Visible     '先前上傳的[附件2]是否含檔案
        Dim blPassF1 As Boolean = False                  '驗證[附件1]是否通過OpenOffice條件限制
        Dim blPassF2 As Boolean = True                   '驗證[附件2]是否通過OpenOffice條件限制

        If blHasNewF1 Then
            If fu1.FileBytes.Length > 10485760 Then Errmsg &= "[附件1] 檔案大小超過10MB" & vbCrLf
        End If

        If blHasNewF2 Then
            If fu2.FileBytes.Length > 10485760 Then Errmsg &= "[附件2] 檔案大小超過10MB" & vbCrLf
        End If

        Dim File1Ext As String = ""
        If blHasNewF1 Then
            Dim UpFile1 As HttpPostedFile = fu1.PostedFile
            Dim File1Split() As String = Split(UpFile1.FileName, "\")
            Dim File1Name As String = File1Split(File1Split.Length - 1)
            File1Ext = System.IO.Path.GetExtension(File1Name)
            If CheckFileType(File1Ext) Then blPassF1 = True Else blPassF1 = False
        End If

        If blPassF1 = False Then
            If blHasOldF1 And CheckFileType(lblF1Ext.Text.Trim) Then blPassF1 = True Else blPassF1 = False
            If Not blHasOldF1 Then blPassF1 = False
        End If

        If blPassF1 = False Then blPassF2 = False

        Dim File2Ext As String = ""
        If blHasNewF2 Then
            Dim UpFile2 As HttpPostedFile = fu2.PostedFile
            Dim File2Split() As String = Split(UpFile2.FileName, "\")
            Dim File2Name As String = File2Split(File2Split.Length - 1)
            File2Ext = System.IO.Path.GetExtension(File2Name)
            If CheckFileType(File2Ext) Then blPassF2 = True Else blPassF2 = False
        End If

        If blPassF2 = False Then
            If blHasOldF2 And CheckFileType(lblF2Ext.Text.Trim) Then blPassF2 = True Else blPassF2 = False
            If Not blHasOldF2 Then blPassF2 = False
        End If

        '==========
        If Not blHasNewF1 And Not blHasOldF1 Then Errmsg &= "[附件1]不得為空，請依序上傳檔案，謝謝!" & vbCrLf
        If Not blPassF1 And Not blPassF2 Then Errmsg &= "至少上傳1個含有ODF類型(例如：odp、ods、odt、pdf)的檔案!" & vbCrLf
        If blPassF1 And blPassF2 Then
            If blHasOldF1 And blHasNewF1 And Not CheckFileType(File1Ext) Then Errmsg &= "至少上傳1個含有ODF類型(例如：odp、ods、odt、pdf)的檔案!" & vbCrLf
        End If

#End Region
        '=================================== End

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

#Region "檢查附件是否包含OpenOffile的File類型"

    '檢查附件是否包含OpenOffile的File類型
    Function CheckFileType(ByVal nowFileType As String) As Boolean
        Dim rst As Boolean = False

        '定義目前OpenOffice包含的檔案類型
        Dim acceptTypeList As New List(Of String)
        acceptTypeList.Add(".ODP")
        acceptTypeList.Add(".ODS")
        acceptTypeList.Add(".ODT")
        acceptTypeList.Add(".PDF")

        Dim i As Integer
        For i = 0 To acceptTypeList.Count - 1
            If nowFileType.Trim.ToUpper.Equals(acceptTypeList.Item(i)) Then
                rst = True
                Exit For
            End If
        Next

        Return rst
    End Function

#End Region

    '儲存(part-1)
    Sub SaveData1()
        Dim flagSaveOK1 As Boolean = False

        Try
            flagSaveOK1 = SaveData2()
        Catch ex As Exception
            flagSaveOK1 = False
            Common.MessageBox(Me, ex.Message)
            Exit Sub
        End Try

        If flagSaveOK1 Then
            '儲存成功
            Dim url1 As String = "RWB_01_003.aspx?id1=" & TIMS.Get_MRqID(Me)
            'Common.MessageBox(Me, "儲存成功!", url1)
            Common.MessageBox(Me, "儲存成功!")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If
    End Sub

    '儲存(part-2)
    Function SaveData2() As Boolean
        Dim rst As Boolean = False 'false:異常

        txtSDATE1.Text = TIMS.ClearSQM(txtSDATE1.Text)
        txtEDATE1.Text = TIMS.ClearSQM(txtEDATE1.Text)
        Dim mySDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtSDATE1.Text), txtSDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myEDATE1 As String = IIf(flag_ROC, TIMS.Cdate18(txtEDATE1.Text), txtEDATE1.Text).Replace("/", "-")  'edit，by:20181019
        txtTitle.Text = TIMS.ClearSQM(txtTitle.Text)
        Dim oC_SDATE As Object = mySDATE1 + " " + ddlC_SDATE_hh1.SelectedValue + ":" + ddlC_SDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim oC_EDATE As Object = myEDATE1 + " " + ddlC_EDATE_hh1.SelectedValue + ":" + ddlC_EDATE_mm1.SelectedValue + ":00.000"  'edit，by:20181019
        Dim vITEM1 As String = TIMS.ClearSQM(ddlType.SelectedValue)
        Dim vITEM2 As String = TIMS.ClearSQM(ddlPlan.SelectedValue)

#Region "[附件檔案]相關參數設定"

        Dim myPlan As String = ""
        If vITEM1.Equals("1") Then myPlan = "PDL" Else myPlan = "ODL"

        'Dim myFilePath1 As String = ConfigurationManager.AppSettings("UploadPath") + "/DLFILE/"
        Dim UpFile1 As HttpPostedFile = fu1.PostedFile
        Dim F1oName As String = ""
        Dim F1Ext As String = ""
        Dim F1nName As String = ""
        Dim F1Path As String = ""
        Dim F1oSize As Integer = 0
        Dim F1nSize As String = ""
        Dim F1nTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")

        If fu1.HasFile Then
            F1oSize = fu1.FileBytes.Length
            F1nSize = GetFileSize(F1oSize)
            Dim File1Split() As String = Split(UpFile1.FileName, "\")
            F1oName = File1Split(File1Split.Length - 1)
            F1Ext = System.IO.Path.GetExtension(F1oName).ToUpper
            F1nName = myPlan + "_" + F1nTime + F1Ext
            'F1Path = Server.MapPath(myFilePath1 + F1nName)
            F1Path = Server.MapPath(myFile_Path1 + F1nName)
            fu1.SaveAs(F1Path)
        End If

        'Dim myFilePath2 As String = ConfigurationManager.AppSettings("UploadPath") + "/DLFILE/"
        Dim UpFile2 As HttpPostedFile = fu2.PostedFile
        Dim F2oName As String = ""
        Dim F2Ext As String = ""
        Dim F2nName As String = ""
        Dim F2Path As String = ""
        Dim F2oSize As Integer = 0
        Dim F2nSize As String = ""
        Dim F2nTime As String = DateTime.Now.ToString("yyyyMMddHHmmss")

        If fu2.HasFile Then
            F2oSize = fu2.FileBytes.Length
            F2nSize = GetFileSize(F2oSize)
            Dim File2Split() As String = Split(UpFile2.FileName, "\")
            F2oName = File2Split(File2Split.Length - 1)
            F2Ext = System.IO.Path.GetExtension(F2oName).ToUpper
            F2nName = myPlan + "_" + F2nTime + F2Ext
            'F2Path = Server.MapPath(myFilePath2 + F2nName)
            F2Path = Server.MapPath(myFile_Path2 + F2nName)
            fu2.SaveAs(F2Path)
        End If

#End Region

        '==============================
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO TB_DLFILE " & vbCrLf
        sql &= "   (DLID, KINDID, PLANID, DLTITLE " & vbCrLf
        If fu1.HasFile Then sql &= "    , FILE1_NAME, FILE1_TYPE, FILE1_SIZE " & vbCrLf
        If fu2.HasFile Then sql &= "    , FILE2_NAME, FILE2_TYPE, FILE2_SIZE " & vbCrLf
        sql &= "    , START_DATE, END_DATE, DLCOUNT " & vbCrLf
        sql &= "    , UPLOADDATE, ISUSED, MODIFYACCT, MODIFYDATE) " & vbCrLf
        sql &= " VALUES (@DLID, @KINDID, @PLANID, @DLTITLE " & vbCrLf
        If fu1.HasFile Then sql &= "    , @FILE1_NAME, @FILE1_TYPE, @FILE1_SIZE " & vbCrLf
        If fu2.HasFile Then sql &= "    , @FILE2_NAME, @FILE2_TYPE, @FILE2_SIZE " & vbCrLf
        sql &= "    , @START_DATE, @END_DATE, '0' " & vbCrLf
        sql &= "    , GETDATE(), 'Y', @MODIFYACCT, GETDATE()) " & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)
        Dim iSql As String = sql

        sql = "" & vbCrLf
        sql &= " UPDATE TB_DLFILE " & vbCrLf
        sql &= " SET KINDID = @KINDID " & vbCrLf
        sql &= "     ,PLANID = @PLANID " & vbCrLf
        sql &= "     ,DLTITLE = @DLTITLE " & vbCrLf
        If fu1.HasFile Then
            sql &= "  ,FILE1_NAME = @FILE1_NAME " & vbCrLf
            sql &= "  ,FILE1_TYPE = @FILE1_TYPE " & vbCrLf
            sql &= "  ,FILE1_SIZE = @FILE1_SIZE " & vbCrLf
        End If
        If fu2.HasFile Then
            sql &= "  ,FILE2_NAME = @FILE2_NAME " & vbCrLf
            sql &= "  ,FILE2_TYPE = @FILE2_TYPE " & vbCrLf
            sql &= "  ,FILE2_SIZE = @FILE2_SIZE " & vbCrLf
        End If
        sql &= "     ,START_DATE = @START_DATE " & vbCrLf
        sql &= "     ,END_DATE = @END_DATE " & vbCrLf
        sql &= "     ,MODIFYACCT = @MODIFYACCT " & vbCrLf
        sql &= "     ,MODIFYDATE = GETDATE() " & vbCrLf
        sql &= " WHERE DLID = @DLID " & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        Dim uSql As String = sql

        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Call TIMS.OpenDbConn(objconn)

        Dim iRst As Integer = 0
        If hid_V.Value = "" Then
            '新增
            Dim iSEQNO As Integer = DbAccess.GetNewId(objconn, "TB_DLFILE_DLID_SEQ,TB_DLFILE,DLID")
            With iCmd
                Dim parms As Hashtable = New Hashtable()
                parms.Add("DLID", iSEQNO)
                parms.Add("KINDID", vITEM1)
                parms.Add("PLANID", vITEM2)
                parms.Add("DLTITLE", txtTitle.Text)
                If fu1.HasFile Then
                    parms.Add("FILE1_NAME", F1nName)
                    parms.Add("FILE1_TYPE", F1Ext)
                    parms.Add("FILE1_SIZE", F1nSize)
                End If
                If fu2.HasFile Then
                    parms.Add("FILE2_NAME", F2nName)
                    parms.Add("FILE2_TYPE", F2Ext)
                    parms.Add("FILE2_SIZE", F2nSize)
                End If
                parms.Add("START_DATE", oC_SDATE)
                parms.Add("END_DATE", oC_EDATE)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                iRst += DbAccess.ExecuteNonQuery(iSql, parms)
            End With
            hid_V.Value = iSEQNO
        Else
            '修改
            With uCmd
                Dim parms As Hashtable = New Hashtable()
                parms.Add("DLID", hid_V.Value)
                parms.Add("KINDID", vITEM1)
                parms.Add("PLANID", vITEM2)
                parms.Add("DLTITLE", txtTitle.Text)
                If fu1.HasFile Then
                    parms.Add("FILE1_NAME", F1nName)
                    parms.Add("FILE1_TYPE", F1Ext)
                    parms.Add("FILE1_SIZE", F1nSize)
                End If
                If fu2.HasFile Then
                    parms.Add("FILE2_NAME", F2nName)
                    parms.Add("FILE2_TYPE", F2Ext)
                    parms.Add("FILE2_SIZE", F2nSize)
                End If
                parms.Add("START_DATE", oC_SDATE)
                parms.Add("END_DATE", oC_EDATE)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                iRst += DbAccess.ExecuteNonQuery(uSql, parms)
            End With
        End If

        rst = True
        Return rst
    End Function

    '取消
    Protected Sub bt_cancle_Click(sender As Object, e As EventArgs) Handles bt_cancle.Click
        Dim url1 As String = ""
        url1 = "RWB_01_003.aspx?id1=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

#Region "將附件檔案下載到客戶端"

    '將[附件1]檔案下載到客戶端
    Protected Sub lnkF1Name_Click(sender As Object, e As EventArgs) Handles lnkF1Name.Click
        'Dim myFilePath1 As String = ConfigurationManager.AppSettings("UploadPath") + "/DLFILE/" + lblF1Name.Text.Trim
        Dim myFilePath1 As String = myFile_Path1 + TIMS.ClearSQM(lblF1Name.Text) '.Trim
        Dim myFilePath2 As String = Server.MapPath(myFilePath1)

        If Not System.IO.File.Exists(myFilePath2) Then
            Common.MessageBox(Page, "[附件1]已不存在，請重新上傳，謝謝。")
        Else
            Dim fileByte As Byte() = System.IO.File.ReadAllBytes(myFilePath2)
            Dim myNewFileName As String = txtTitle.Text.Trim + lblF1Ext.Text.Trim.ToLower
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment;filename=" + myNewFileName)
            Response.BinaryWrite(fileByte)
            Response.End()
        End If
    End Sub

    '將[附件2]檔案下載到客戶端
    Protected Sub lnkF2Name_Click(sender As Object, e As EventArgs) Handles lnkF2Name.Click
        'Dim myFilePath1 As String = ConfigurationManager.AppSettings("UploadPath") + "/DLFILE/" + lblF2Name.Text.Trim
        Dim myFilePath1 As String = myFile_Path2 + TIMS.ClearSQM(lblF2Name.Text) '.Trim
        Dim myFilePath2 As String = Server.MapPath(myFilePath1)

        If Not System.IO.File.Exists(myFilePath2) Then
            Common.MessageBox(Page, "[附件2]已不存在，請重新上傳，謝謝。")
        Else
            Dim fileByte As Byte() = System.IO.File.ReadAllBytes(myFilePath2)
            Dim myNewFileName As String = txtTitle.Text.Trim + lblF2Ext.Text.Trim.ToLower
            Response.Clear()
            Response.AddHeader("Content-Disposition", "attachment;filename=" + myNewFileName)
            Response.BinaryWrite(fileByte)
            Response.End()
        End If
    End Sub

#End Region

#Region "取得檔案大小"

    Function GetFileSize(ByVal fileLength As Integer) As String
        Dim rst As String = ""
        Dim tNum As Decimal = fileLength

        If tNum > 1024 Then
            tNum = tNum / 1024

            If tNum > 1024 Then
                tNum = tNum / 1024

                If tNum > 1024 Then
                    tNum = tNum / 1024

                    If tNum > 1024 Then
                        tNum = tNum / 1024
                        rst = Math.Round(tNum, 1).ToString + "TB"  '(暫時設定此系統最大容量為TB)
                    Else
                        rst = Math.Round(tNum, 1).ToString + "GB"
                    End If
                Else
                    rst = Math.Round(tNum, 1).ToString + "MB"
                End If
            Else
                rst = Math.Round(tNum, 1).ToString + "KB"
            End If
        Else
            rst = tNum.ToString + "Byte"
        End If

        rst = rst.Replace(".0", "")   '若檔案容量含有[.0]時,則將小數點拿掉

        '========== 20180718 Note:因Integer型別最大值對應到的大小為2GB，所以[檔案大小]最大值為2GB

        Return rst
    End Function

#End Region
End Class