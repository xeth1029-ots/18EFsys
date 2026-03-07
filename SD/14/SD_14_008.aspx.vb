Partial Class SD_14_008
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_008_2009_b"
    Const cst_printFN2 As String = "SD_14_008_2025_c" '(「職場續航」之課程勾稽投保年資)
    '/**NEW 2014**/ 'SD_14_008_2009_b (SQControl.aspx / Printtype / print_orderyby) (c.IDNO / a.StudentID)
    '/** OLD **/ 'SD_14_008_2009
    'Dim iPYNum14 As Integer = 1 'TIMS.sUtl_GetPYNum14(Me) '若是登入年度為 2014年以後，則傳回2，其餘為1
    'Dim prtFilename As String = "" '列印表件名稱
    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection
    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁 '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End 'iPYNum14 = TIMS.sUtl_GetPYNum14(Me)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            print_orderyby.Value = If(TIMS.GetListValue(print_type) = "2", "a.StudentID", "c.IDNO") 'Printtype

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '使用民國年。
        Years.Value = sm.UserInfo.Years - 1911

        print_type.Attributes("onclick") = "printkind();" '列印時排序方式

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse Convert.ToString(dr("total")) <> "1" Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim v_print_type As String = TIMS.GetListValue(print_type)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "未選擇職類/班別，請選擇班級!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "未選擇職類/班別，請選擇班級!!")
            Exit Sub
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Dim pParms As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sSql As String = ""
        sSql &= " SELECT a.OCID,a.SOCID,a.StudentID,a.SID" & vbCrLf
        sSql &= " ,a.STUDID2,a.CLASSCNAME2,a.STDate,a.FTDate" & vbCrLf
        sSql &= " ,a.NAME,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK2(a.Birthday) BIRTHDAY_MK" & vbCrLf
        sSql &= " ,a.Sex,a.StudStatus,a.AppliedResultM ,a.PlanID ,a.RID" & vbCrLf
        sSql &= " FROM dbo.VIEW_STUDENTBASICDATA a" & vbCrLf
        '排除離退訓學員
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3) AND a.OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDID2,NAME,IDNO")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm列印, TIMS.cst_wmdip2, OCIDValue1.Value, "", objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        '(「職場續航」之課程勾稽投保年資)
        Dim gfg_WYROLE As String = TIMS.CHECK_WYROLE(objconn, OCIDValue1.Value) 'Dim gfg_WYROLE As Boolean = CHECK_WYROLE()

        Dim rptFILENAME As String = If(gfg_WYROLE, cst_printFN2, cst_printFN1)

        Dim V_MSD As String = Convert.ToString(drCC("MSD"))
        Years.Value = TIMS.ClearSQM(Years.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        print_orderyby.Value = If(v_print_type = "2", "a.StudentID", "c.IDNO")

        Dim myvalue As String = ""
        TIMS.SetMyValue(myvalue, "MSD", V_MSD)
        TIMS.SetMyValue(myvalue, "Years", Years.Value) 'sm.UserInfo.Years - 1911
        TIMS.SetMyValue(myvalue, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myvalue, "Printtype", print_orderyby.Value)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, rptFILENAME, myvalue)
    End Sub
End Class