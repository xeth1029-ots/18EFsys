Partial Class SD_03_002_img
    Inherits AuthBasePage

    Const vs_SearchStr As String = "_SearchStr"
    Const vs_HighEduBg As String = "_HighEduBg"
    Const vs_IdentityID As String = "_IdentityID"
    Const vs_STDate As String = "_STDate"

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '載入資料
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            Call SUtl_Create1()
        End If
    End Sub

    '載入資料(首頁載入)(執行1次) If Not IsPostBack Then Call sUtl_Create1()
    Sub SUtl_Create1()
        LabMsg1.Text = TIMS.cst_NODATAMsg1
        'Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        'Dim ETYPE As String = TIMS.GetMyValue(StrCmdArg1, "ETYPE")
        'Dim EMID1 As String = TIMS.GetMyValue(StrCmdArg1, "EMID1")
        'Dim FILENAME As String = TIMS.GetMyValue(StrCmdArg1, "FILENAME")
        Dim rqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqETYPE As String = TIMS.ClearSQM(Request("ETYPE"))
        Dim rqEMID1 As String = TIMS.ClearSQM(Request("EMID1"))
        Dim rqFILENAME As String = TIMS.ClearSQM(Request("FILENAME"))
        Dim rqECMD As String = TIMS.ClearSQM(Request("ECMD"))
        'Dim FuncID As String = TIMS.Get_MRqID(Me)
        'Dim Url_1 As String = String.Concat("SD_03_002_IMG?ID=", FuncID, "&ECMD=", e.CommandName, "&ETYPE=", ETYPE, "&EMID1=", EMID1, "&FILENAME=", FILENAME, "&OCID=", Hid_OCID.Value, "&SOCID=", v_SOCID)
        'Select Case e.CommandName
        '    Case "SF1"
        '        Call TIMS.Utl_Redirect(Me, objconn, Url_1)
        '    Case "SB2"
        '        Call TIMS.Utl_Redirect(Me, objconn, Url_1)
        '    Case "SPB"
        '        Call TIMS.Utl_Redirect(Me, objconn, Url_1)
        'End Select
        Dim drSD As DataRow = TIMS.Get_StudData(rqSOCID, rqOCID, objconn)
        If drSD Is Nothing Then Return
        Dim V_IDNO As String = Convert.ToString(drSD("IDNO"))

        Dim oDt As New DataTable
        Dim SQL_S1 As String = ""
        SQL_S1 &= " SELECT E.ETYPE,E.EMID1,E.IDNO,E.CREATEDATE,E.FILENAME1,E.FILENAME1W,E.SRCFILENAME1,E.FILEPATH1" & vbCrLf
        SQL_S1 &= " ,E.FILENAME2,E.FILENAME2W,E.SRCFILENAME2,E.FILEPATH2" & vbCrLf
        SQL_S1 &= " ,E.ISUSE,E.ISDEL,E.MODIFYACCT,E.MODIFYDATE,E.CATEGORY1,E.ACTION1" & vbCrLf
        SQL_S1 &= " FROM V_EIMG12 E" & vbCrLf
        SQL_S1 &= " WHERE E.MODIFYDATE IS NOT NULL AND LEN(E.FILENAME1)>0 AND E.IDNO=@IDNO" & vbCrLf
        SQL_S1 &= " AND E.ETYPE=@ETYPE AND E.EMID1=@EMID1" & vbCrLf
        Using oCmd As New SqlCommand(SQL_S1, objconn)
            With oCmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = V_IDNO
                .Parameters.Add("ETYPE", SqlDbType.Int).Value = Val(rqETYPE)
                .Parameters.Add("EMID1", SqlDbType.Int).Value = Val(rqEMID1)
                oDt.Load(.ExecuteReader())
            End With
        End Using
        LabMsg1.Text = "(查無圖檔資料)"
        If TIMS.dtNODATA(oDt) Then Return
        Dim dr1 As DataRow = oDt.Rows(0)
        Dim v_imgUrl As String = ""
        Dim v_imgSMP As String = ""
        Select Case rqECMD'e.CommandName
            Case "SF1"
                LabMsg1.Text = ""
                Dim FILEPATH1 As String = Convert.ToString(dr1("FILEPATH1"))
                Dim FILENAME1W As String = Convert.ToString(dr1("FILENAME1"))
                v_imgUrl = String.Concat(FILEPATH1, FILENAME1W)
                v_imgSMP = Server.MapPath(v_imgUrl)
                TIMS.LOG.Debug(String.Concat("v_imgUrl :", v_imgUrl))
                TIMS.LOG.Debug(String.Concat("v_imgSMP :", v_imgSMP))
                If Not IO.File.Exists(v_imgSMP) Then '若檔案不存在製造預設值
                    Hid_ERRMSG1.Value = "圖檔資料有誤!Exists"
                    TIMS.LOG.Debug(Hid_ERRMSG1.Value)
                    'LabMsg1.Text = "圖檔資料有誤!" 'Return
                End If
                Image2.ImageUrl = v_imgUrl
            Case "SB2"
                LabMsg1.Text = ""
                Dim FILEPATH2 As String = Convert.ToString(dr1("FILEPATH2"))
                Dim FILENAME2W As String = Convert.ToString(dr1("FILENAME2"))
                v_imgUrl = String.Concat(FILEPATH2, FILENAME2W)
                v_imgSMP = Server.MapPath(v_imgUrl)
                TIMS.LOG.Debug(String.Concat("v_imgUrl :", v_imgUrl))
                TIMS.LOG.Debug(String.Concat("v_imgSMP :", v_imgSMP))
                If Not IO.File.Exists(v_imgSMP) Then '若檔案不存在製造預設值
                    Hid_ERRMSG1.Value = "圖檔資料有誤!Exists"
                    TIMS.LOG.Debug(Hid_ERRMSG1.Value)
                    'LabMsg1.Text = "圖檔資料有誤!" 'Return
                End If
                Image2.ImageUrl = v_imgUrl
            Case "SPB"
                LabMsg1.Text = ""
                Dim FILEPATH1 As String = Convert.ToString(dr1("FILEPATH1"))
                Dim FILENAME1W As String = Convert.ToString(dr1("FILENAME1"))
                v_imgUrl = String.Concat(FILEPATH1, FILENAME1W)
                v_imgSMP = Server.MapPath(v_imgUrl)
                TIMS.LOG.Debug(String.Concat("v_imgUrl :", v_imgUrl))
                TIMS.LOG.Debug(String.Concat("v_imgSMP :", v_imgSMP))
                If Not IO.File.Exists(v_imgSMP) Then '若檔案不存在製造預設值
                    Hid_ERRMSG1.Value = "圖檔資料有誤!Exists"
                    TIMS.LOG.Debug(Hid_ERRMSG1.Value)
                    'LabMsg1.Text = "圖檔資料有誤!" 'Return
                End If
                Image2.ImageUrl = v_imgUrl
        End Select
    End Sub

    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Session(vs_SearchStr) IsNot Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr) 'Session(vs_SearchStr) = Nothing
        End If

        Dim rqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim s_rqFUNID As String = TIMS.ClearSQM(Request("ID"))
        'Dim s_redirect2 As String = String.Concat("../03/SD_03_002.aspx?ID=", s_rqFUNID, "&todo=2", "&OCID=", rqOCID) 'Call TIMS.Utl_Redirect(Me, objconn, s_redirect2)
        'Dim SOCID_value As String = TIMS.GetMyValue(e.CommandArgument, "SOCID") 'Dim tmpName As String=TIMS.GetMyValue(e.CommandArgument, "tmpName")
        'Session("SearchSOCID") = SOCID_value 'Session("SearchSOCID")=e.Item.Cells(cst_SOCID).Text 'Call GetSearchStr()

        Const cst_SD03002_addaspx As String = "SD_03_002_add.aspx" '28:產業人才投資方案  '有補助比例 (產投)
        Dim str_SDADDASPX As String = cst_SD03002_addaspx '有補助比例 (產投)
        Session("SearchSOCID") = rqSOCID
        Dim Url_1 As String = String.Concat(str_SDADDASPX, "?ID=", s_rqFUNID, "&OCID=", rqOCID, "&SOCID=" & rqSOCID)
        Call TIMS.Utl_Redirect(Me, objconn, Url_1)

    End Sub
End Class
