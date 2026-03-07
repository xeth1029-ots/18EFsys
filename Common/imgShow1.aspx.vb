Partial Class imgShow1
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        Dim vELNO As String = TIMS.sUtl_GetRqValue(Me, "ELNO")
        Dim CSELNO As String = TIMS.sUtl_GetRqValue(Me, "CSELNO")
        Dim SOCID As String = TIMS.sUtl_GetRqValue(Me, "SOCID")
        Dim OCID As String = TIMS.sUtl_GetRqValue(Me, "OCID")
        Dim SID As String = TIMS.sUtl_GetRqValue(Me, "SID")
        Dim sCmdArg As String = ""
        TIMS.SetMyValue(sCmdArg, "ELNO", vELNO)
        TIMS.SetMyValue(sCmdArg, "CSELNO", CSELNO)
        TIMS.SetMyValue(sCmdArg, "SOCID", SOCID)
        TIMS.SetMyValue(sCmdArg, "OCID", OCID)
        TIMS.SetMyValue(sCmdArg, "SID", SID)
        Call SHOW1_IMG(sCmdArg)
    End Sub


    '顯示簽名
    Private Sub SHOW1_IMG(sCmdArg As String)
        Dim ELNO As String = TIMS.GetMyValue(sCmdArg, "ELNO")
        Dim CSELNO As String = TIMS.GetMyValue(sCmdArg, "CSELNO")
        'Dim ORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        'Dim TPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        'Dim RID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim SOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim SID As String = TIMS.GetMyValue(sCmdArg, "SID")

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drCC Is Nothing Then Return

        Dim pms2 As New Hashtable From {{"SID", SID}}
        Dim sSql2 As String = " SELECT IDNO FROM STUD_STUDENTINFO WHERE SID=@SID" & vbCrLf
        Dim dr2 As DataRow = DbAccess.GetOneRow(sSql2, objconn, pms2)
        If dr2 Is Nothing AndAlso Convert.ToString(dr2("IDNO")) = "" Then Return
        Dim strIDNO As String = Convert.ToString(dr2("IDNO"))

        Dim pms1 As New Hashtable From {{"IDNO", strIDNO}, {"ELNO", Val(ELNO)}, {"SOCID", Val(SOCID)}, {"OCID", Val(OCID)}, {"CSELNO", Val(CSELNO)}}
        Dim sSql1 As String = ""
        sSql1 &= " SELECT CSELNO,ELNO,SOCID,OCID,IDNO,P1_LINK,CREATEACCT,CREATEDATE,SIGNDACCT,SIGNDATE,SENDACCT,SENDDATE,MODIFYACCT,MODIFYDATE,FILEPATH1" & vbCrLf
        sSql1 &= " FROM STUD_ELFORM WHERE IDNO=@IDNO AND ELNO=@ELNO AND SOCID=@SOCID AND OCID=@OCID and CSELNO=@CSELNO" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, pms1)
        If dr1 Is Nothing Then Return

        Dim UploadRootPath As String = TIMS.Utl_GetConfigSet("UploadElSignPath")
        If UploadRootPath = "" Then UploadRootPath = "../../upojt"
        Const cst_ElFormSignPath As String = "ElFormSign"

        '如果有線上簽名資料。//簽名檔位置
        Dim strP1_LINK As String = Convert.ToString(dr1("P1_LINK"))
        Dim sElFormSignPath As String = cst_ElFormSignPath
        Dim webPath As String = String.Concat(UploadRootPath, "/", sElFormSignPath, "/", drCC("PLANID"), "/", drCC("OCID"), "/")
        ' 獲取虛擬路徑相對應的物理路徑
        Dim savePath As String = System.IO.Path.Combine(webPath)
        '組合顯示簽名圖片的路徑
        Dim showImage_Path As String = String.Concat(savePath, strP1_LINK)

        '彈出訊息，顯示簽名圖片
        Image1.ImageUrl = showImage_Path
        'Dim img_buf As String = String.Concat("<img src=", showImage_Path, " width=500 />")
        'Dim js_1 As String = String.Concat("window.open(""", img_buf, """);")
        'TIMS.Utl_RespWriteEnd(Me, objconn, js_1)
    End Sub


End Class