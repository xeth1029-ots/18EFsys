Partial Class SD_15_032
    Inherits AuthBasePage

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

        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub


    Sub cCreate1()
        Labmsg1.Text = "(若未選擇計畫，則依目前登入計畫搜尋)"
        msg1.Text = ""

        ddlYEARS_SCH1 = TIMS.GetSyear(ddlYEARS_SCH1)
        ddlYEARS_SCH2 = TIMS.GetSyear(ddlYEARS_SCH2)
        Common.SetListItem(ddlYEARS_SCH1, sm.UserInfo.Years)
        Common.SetListItem(ddlYEARS_SCH2, sm.UserInfo.Years)

        cblTPLANID = TIMS.Get_TPlan(cblTPLANID, , 1, "Y", "", objconn) 'cblTPLANID.Items(0).Selected = True
        cblTPLANID.Attributes("onclick") = "SelectAll('cblTPLANID','HiddencblTPLANID');"
    End Sub

    Function SET_MY_WHERE_VALUE1(ByRef sPMS As Hashtable) As String
        sPMS = New Hashtable
        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        TEACHCNAME_SCH.Text = TIMS.ClearSQM(TEACHCNAME_SCH.Text)
        IDNO_SCH.Text = TIMS.ClearSQM(IDNO_SCH.Text)

        'Dim v_cblTPLANID As String = TIMS.GetCblValue(cblTPLANID)
        Dim v_cblTPLANIDIn As String = TIMS.GetCblValueIn(cblTPLANID)
        '使用登入計畫
        'Dim fg_USE_SM_TPLANID As Boolean = (v_cblTPLANID = "" OrElse v_cblTPLANIDIn = "")

        'Dim sPMS As New Hashtable
        ''使用登入計畫
        If v_cblTPLANIDIn = "" Then sPMS.Add("TPLANID", sm.UserInfo.TPlanID)
        If (v_ddlYEARS_SCH1 <> "") Then sPMS.Add("YEARS1", v_ddlYEARS_SCH1)
        If (v_ddlYEARS_SCH2 <> "") Then sPMS.Add("YEARS2", v_ddlYEARS_SCH2)
        If (TEACHCNAME_SCH.Text <> "") Then sPMS.Add("TEACHCNAME", TEACHCNAME_SCH.Text)
        If (IDNO_SCH.Text <> "") Then sPMS.Add("IDNO", IDNO_SCH.Text)

        '===sSql
        Dim sSql As String = ""
        If v_cblTPLANIDIn = "" Then
            sSql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
        Else
            sSql &= String.Concat(" AND ip.TPLANID IN (", v_cblTPLANIDIn, ")", vbCrLf)
        End If
        If (v_ddlYEARS_SCH1 <> "") Then sSql &= " AND ip.YEARS>=@YEARS1" & vbCrLf
        If (v_ddlYEARS_SCH2 <> "") Then sSql &= " AND ip.YEARS<=@YEARS2" & vbCrLf
        If (TEACHCNAME_SCH.Text <> "") Then sSql &= " AND tt.TEACHCNAME=@TEACHCNAME" & vbCrLf
        If (IDNO_SCH.Text <> "") Then sSql &= " AND tt.IDNO=@IDNO" & vbCrLf
        Return sSql
    End Function

    Function SET_MY_WHERE_VALUE2(ByRef sPMS As Hashtable) As String
        sPMS = New Hashtable
        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        TEACHCNAME_SCH.Text = TIMS.ClearSQM(TEACHCNAME_SCH.Text)
        IDNO_SCH.Text = TIMS.ClearSQM(IDNO_SCH.Text)
        'Dim v_cblTPLANID As String = TIMS.GetCblValue(cblTPLANID)
        Dim v_cblTPLANIDIn As String = TIMS.GetCblValueIn(cblTPLANID)
        '使用登入計畫
        'Dim fg_USE_SM_TPLANID As Boolean = (v_cblTPLANID = "" OrElse v_cblTPLANIDIn = "")

        'Dim sPMS As New Hashtable
        ''使用登入計畫
        If v_cblTPLANIDIn = "" Then sPMS.Add("TPLANID", sm.UserInfo.TPlanID)
        If (v_ddlYEARS_SCH1 <> "") Then sPMS.Add("YEARS1", v_ddlYEARS_SCH1)
        If (v_ddlYEARS_SCH2 <> "") Then sPMS.Add("YEARS2", v_ddlYEARS_SCH2)
        If (TEACHCNAME_SCH.Text <> "") Then sPMS.Add("TEACHCNAME", TEACHCNAME_SCH.Text)
        If (IDNO_SCH.Text <> "") Then sPMS.Add("IDNO", IDNO_SCH.Text)

        '===sSql
        Dim sSql As String = ""
        If v_cblTPLANIDIn = "" Then
            sSql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
        Else
            sSql &= String.Concat(" AND ip.TPLANID IN (", v_cblTPLANIDIn, ")", vbCrLf)
        End If
        If (v_ddlYEARS_SCH1 <> "") Then sSql &= " AND ip.YEARS>=@YEARS1" & vbCrLf
        If (v_ddlYEARS_SCH2 <> "") Then sSql &= " AND ip.YEARS<=@YEARS2" & vbCrLf
        If (TEACHCNAME_SCH.Text <> "") Then sSql &= " AND tt.TEACHCNAME=@TEACHCNAME" & vbCrLf
        If (IDNO_SCH.Text <> "") Then sSql &= " AND tt.IDNO=@IDNO" & vbCrLf
        Return sSql
    End Function

    Function sSearch1_DATA_dt() As DataTable
        Dim dt1 As DataTable = Nothing
        Dim sPMS1 As New Hashtable
        Dim sSql As String = ""
        'WT1 
        sSql &= " WITH WT1 AS ( SELECT cc.OCID,tt.TECHID,sum(td1.PHOUR) PHOUR" & vbCrLf
        sSql &= " ,min(td1.STRAINDATE) STRAINDATE1,max(td1.STRAINDATE) STRAINDATE2" & vbCrLf
        sSql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sSql &= " JOIN PLAN_TRAINDESC td1 ON td1.PLANID=cc.PLANID AND td1.COMIDNO=cc.COMIDNO AND td1.SEQNO=cc.SEQNO" & vbCrLf
        sSql &= " JOIN TEACH_TEACHERINFO tt ON tt.TECHID=td1.TECHID" & vbCrLf
        sSql &= " JOIN ID_PLAN ip ON ip.PLANID=cc.PLANID" & vbCrLf
        sSql &= " WHERE CC.ISSUCCESS='Y' AND CC.NOTOPEN='N'" & vbCrLf
        sSql &= SET_MY_WHERE_VALUE1(sPMS1)
        sSql &= " GROUP BY cc.OCID,tt.TECHID )" & vbCrLf
        'sql main
        sSql &= " SELECT cc.YEARS" & vbCrLf ' --計畫年度" & vbCrLf
        sSql &= " ,cc.PLANNAME" & vbCrLf ' --計畫" & vbCrLf
        sSql &= " ,cc.ORGNAME" & vbCrLf ' --訓練單位" & vbCrLf
        sSql &= " ,tt.TEACHCNAME" & vbCrLf ' --師資姓名" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(tt.IDNO) IDNO_MK" & vbCrLf 'sSql &= " --,tt.IDNO 身分證號" & vbCrLf
        sSql &= " ,cc.CLASSCNAME" & vbCrLf ' --課程名稱" & vbCrLf
        sSql &= " ,td1.PHOUR" & vbCrLf 'sSql &= " --,cc.THOURS --授課總時數" & vbCrLf
        sSql &= " ,concat(format(td1.STRAINDATE1,'yyyy/MM/dd'),'~',format(td1.STRAINDATE2,'yyyy/MM/dd')) STRAINDATE" & vbCrLf 'sSql &= " --,cc.STDATE --訓練期間,cc.FTDATE --訓練期間" & vbCrLf
        sSql &= " ,cc.OCID" & vbCrLf ' --課程代碼" & vbCrLf
        sSql &= " FROM WT1 td1" & vbCrLf
        sSql &= " JOIN VIEW2 cc on cc.OCID=td1.OCID" & vbCrLf
        sSql &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=td1.TECHID" & vbCrLf
        'sSql &= " --WHERE cc.YEARS='2023' AND cc.TPLANID='28'" & vbCrLf
        dt1 = DbAccess.GetDataTable(sSql, objconn, sPMS1)

        'If dt1 IsNot Nothing AndAlso dt1.Rows.Count = 0 Then
        'End If

        'WT2 
        Dim dt2 As DataTable = Nothing
        Dim sPMS2 As New Hashtable
        Dim sSql2 As String = ""
        sSql2 &= " WITH WT2 AS (select sc.OCID,sc.TECHID,SUM(1.0) PHOUR" & vbCrLf
        sSql2 &= " ,min(sc.SCHOOLDATE) STRAINDATE1,max(sc.SCHOOLDATE) STRAINDATE2" & vbCrLf
        sSql2 &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sSql2 &= " JOIN MV_CLASS_SCHEDULE2 sc on sc.OCID=cc.OCID" & vbCrLf
        sSql2 &= " JOIN TEACH_TEACHERINFO tt ON tt.TECHID=sc.TECHID" & vbCrLf
        sSql2 &= " JOIN ID_PLAN ip ON ip.PLANID=cc.PLANID" & vbCrLf
        sSql2 &= " WHERE CC.ISSUCCESS='Y' AND CC.NOTOPEN='N'" & vbCrLf
        sSql2 &= SET_MY_WHERE_VALUE2(sPMS2)
        sSql2 &= " GROUP BY sc.OCID,sc.TECHID )" & vbCrLf
        'sql main
        sSql2 &= " SELECT cc.YEARS" & vbCrLf ' --計畫年度" & vbCrLf
        sSql2 &= " ,cc.PLANNAME" & vbCrLf ' --計畫" & vbCrLf
        sSql2 &= " ,cc.ORGNAME" & vbCrLf ' --訓練單位" & vbCrLf
        sSql2 &= " ,tt.TEACHCNAME" & vbCrLf ' --師資姓名" & vbCrLf
        sSql2 &= " ,dbo.FN_GET_MASK1(tt.IDNO) IDNO_MK" & vbCrLf 'sSql2 &= " --,tt.IDNO 身分證號" & vbCrLf
        sSql2 &= " ,cc.CLASSCNAME" & vbCrLf ' --課程名稱" & vbCrLf
        sSql2 &= " ,td1.PHOUR" & vbCrLf 'sSql2 &= " --,cc.THOURS --授課總時數" & vbCrLf
        sSql2 &= " ,concat(format(td1.STRAINDATE1,'yyyy/MM/dd'),'~',format(td1.STRAINDATE2,'yyyy/MM/dd')) STRAINDATE" & vbCrLf 'sSql2 &= " --,cc.STDATE --訓練期間,cc.FTDATE --訓練期間" & vbCrLf
        sSql2 &= " ,cc.OCID" & vbCrLf ' --課程代碼" & vbCrLf
        sSql2 &= " FROM WT2 td1" & vbCrLf
        sSql2 &= " JOIN VIEW2 cc on cc.OCID=td1.OCID" & vbCrLf
        sSql2 &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=td1.TECHID" & vbCrLf
        'sSql2 &= " --WHERE cc.YEARS='2023' AND cc.TPLANID='28'" & vbCrLf
        dt2 = DbAccess.GetDataTable(sSql2, objconn, sPMS2)
        '(合併)
        TIMS.CopyDATATABLE(dt1, dt2)

        Return dt1
    End Function

    Sub EXPORT_1()
        Dim sERRMSG1 As String = cCheckData1()
        If sERRMSG1 <> "" Then
            Common.MessageBox(Me, sERRMSG1)
            Return
        End If

        Dim dtXls As DataTable = sSearch1_DATA_dt()
        If dtXls Is Nothing OrElse dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Const cst_TitleS1 As String = "師資資料"
        Dim strFilename1 As String = String.Concat(cst_TitleS1, TIMS.GetDateNo2())
        Dim sTitle1 As String = "師資資料"

        Dim sPattern As String = ""
        sPattern &= "計畫年度,計畫,訓練單位,師資姓名,身分證號,課程名稱,授課總時數,訓練期間,課程代碼"
        Dim sColumn As String = ""
        sColumn &= "YEARS,PLANNAME,ORGNAME,TEACHCNAME,IDNO_MK,CLASSCNAME,PHOUR,STRAINDATE,OCID"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Dim ErrMsg1 As String = cCheckData1()
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If

        Call EXPORT_1()
    End Sub

    Private Function cCheckData1() As String
        Dim ErrMessage1 As String = ""
        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        If v_ddlYEARS_SCH1 = "" Then ErrMessage1 &= "起始計畫年度，必須選擇" & vbCrLf
        If v_ddlYEARS_SCH2 = "" Then ErrMessage1 &= "迄止計畫年度，必須選擇" & vbCrLf

        If v_ddlYEARS_SCH1 <> "" AndAlso v_ddlYEARS_SCH2 <> "" AndAlso TIMS.VAL1(v_ddlYEARS_SCH2) < TIMS.VAL1(v_ddlYEARS_SCH1) Then
            ErrMessage1 &= "迄止計畫年度，必須大於等於 起始計畫年度" & vbCrLf
        End If

        TEACHCNAME_SCH.Text = TIMS.ClearSQM(TEACHCNAME_SCH.Text)
        IDNO_SCH.Text = TIMS.ClearSQM(IDNO_SCH.Text)
        If TEACHCNAME_SCH.Text = "" AndAlso IDNO_SCH.Text = "" Then ErrMessage1 &= "身分證號 或 姓名，至少要擇一輸入" & vbCrLf

        Return ErrMessage1
    End Function
End Class
