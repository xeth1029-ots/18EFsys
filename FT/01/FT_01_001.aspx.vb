Public Class FT_01_001
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    'ADP_FDOWNLOAD
    'Batch\dbt_20210217 'dbt_20210217 '108_產投_當年度實務給付

    '一、資料內容：
    '使用系統排程每年3月15日固定撈取一次前一年度撥款日期在1/1-12/31給付參訓學員補助費之資料，並將資料放於前一年度。
    '例如：109/03/15 系統產生一次 撥款日期=108/01/01-108/12/31之資料，當User選擇【年度】為108年時可下載。
    '二、匯出為excel檔，所需欄位包括：
    '計畫年度、計畫名稱、分署、訓練單位、班別名稱、學員姓名、身分證號碼、出生日期、通訊地址、戶籍地址、開訓日期、結訓日期、預算別、撥款日期、補助費用。(欄位如附件)
    '三、資料匯出：
    '年度：依【年度】欄位選擇要匯出之年度
    '計畫：依登入計畫
    '匯出檔名(範例)：108_產投_當年度實務給付_定版數據+日期.xlsx、108_產投_當年度實務給付_定版數據+日期.ods   (依計畫分：充飛、區域據點)
    '資料為定版數據(不因撈取時間不同而變動)。
    '四、權限：此功能僅開放給署，不開放給分署及訓練單位。

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        msg.Text = ""
        If Not Me.IsPostBack Then cCreate1()

    End Sub

    Sub cCreate1()
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, "TPLANID IN ('28','54','70')", objconn)
        TPlanlist1.Items.Add(New ListItem("ECFA", "ECFA"))
        Common.SetListItem(TPlanlist1, sm.UserInfo.TPlanID)

        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        TPlanlist1.Enabled = False
        If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        'Dim sql As String
        'sql = "SELECT DISTID,DISTNAME3 NAME FROM V_DISTRICT ORDER BY DISTID"
        'Dim dtDIST As DataTable = DbAccess.GetDataTable(sql, objconn)
        'ddlDISTID = TIMS.Get_DistID(ddlDISTID, dtDIST)
        'Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        Dim iSYears As Integer = 2020
        Dim iEYearsNowb1 As Integer = Year(Now) - 1
        Dim iEYears As Integer = If(iEYearsNowb1 > 2020, iEYearsNowb1, 2020)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)

        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, "")
    End Sub



    '檢核資料表-取得匯入匯出檔名
    Function GET_FDOWNLOAD(ByRef s_attachmn As String) As String
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        'Dim t_yearlist As String = TIMS.GetListText(yearlist)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = ""

        Const cst_FTNUM_001 As String = "001"
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("FTNUM", cst_FTNUM_001)
        parms.Add("TPLANID", v_TPlanlist1)
        parms.Add("BYEAR", v_yearlist)
        'parms.Add("BMONTH", v_monthlist)
        'parms.Add("BDATAVER", v_BDATAVER)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT FDID " & vbCrLf '/*PK*/
        sql &= " ,FNAME1 ,FNAMEIN ,FNAMEOUT" & vbCrLf
        sql &= " ,FTNUM ,EXPDATE ,TPLANID" & vbCrLf
        sql &= " ,BYEAR ,BMONTH ,BDATAVER" & vbCrLf
        sql &= " ,FSN,FSTATUS,FCOUNT" & vbCrLf
        'sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_FDOWNLOAD" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND FTNUM=@FTNUM " & vbCrLf
        sql &= " AND TPLANID=@TPLANID " & vbCrLf
        sql &= " AND BYEAR=@BYEAR " & vbCrLf
        sql &= " AND FSTATUS IS NULL" & vbCrLf
        sql &= " ORDER BY MODIFYDATE DESC " & vbCrLf
        'sql &= " AND BMONTH=@BMONTH " & vbCrLf
        'sql &= " AND BDATAVER=@BDATAVER " & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then Return s_filename
        Dim dr1 As DataRow = dt.Rows(0)
        s_attachmn = Convert.ToString(dr1("FNAMEOUT"))
        s_filename = Convert.ToString(dr1("FNAMEIN"))
        Return s_filename
    End Function

    '下載匯出檔-取得匯入匯出檔名
    Function GET_FILENAME1(ByRef s_attachmn As String) As String
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        'Dim t_yearlist As String = TIMS.GetListText(yearlist)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = GET_FDOWNLOAD(s_attachmn)
        '若有查詢到值，下面可略過
        If s_filename <> "" AndAlso s_attachmn <> "" Then Return s_filename

        Dim YEARS_ROC As String = CStr(CInt(v_yearlist) - 1911)
        Dim TPLAN As String = "OTH"
        Select Case v_TPlanlist1
            Case "28"
                TPLAN = "產投"
            Case "54"
                TPLAN = "充電起飛"
            Case "70"
                TPLAN = "區域產業據點"
            Case "ECFA"
                TPLAN = "ECFA"
            Case Else
                TPLAN = String.Format("OTH{0}", v_TPlanlist1)
        End Select
        Dim fileN1 As String = String.Format("{0}年_{1}_當年度實務給付", YEARS_ROC, TPLAN)

        'Dim fileN1 As String = ""
        'Dim s_attachmn As String = ""
        'Dim s_filename As String = ""
        If v_yearlist.Equals("2020") AndAlso v_TPlanlist1.Equals("28") Then
            'fileN1 = "109年_產投_當年度實務給付"
            s_attachmn = String.Format("{0}.xlsx", fileN1)
            s_filename = String.Format("{0}_202103151449.xlsx", fileN1)
        End If
        If v_yearlist.Equals("2020") AndAlso v_TPlanlist1.Equals("54") Then
            'fileN1 = "109年_充電起飛_當年度實務給付"
            s_attachmn = String.Format("{0}.xlsx", fileN1)
            s_filename = String.Format("{0}_202103151449.xlsx", fileN1)
        End If
        If v_yearlist.Equals("2020") AndAlso v_TPlanlist1.Equals("70") Then
            'fileN1 = "109年_區域產業據點_當年度實務給付"
            s_attachmn = String.Format("{0}.xlsx", fileN1)
            s_filename = String.Format("{0}_202103151449.xlsx", fileN1)
        End If
        If v_yearlist.Equals("2020") AndAlso v_TPlanlist1.Equals("ECFA") Then
            'fileN1 = "109年_ECFA_當年度實務給付"
            s_attachmn = String.Format("{0}.xlsx", fileN1)
            s_filename = String.Format("{0}_202103171244.xlsx", fileN1)
        End If
        'If s_filename = "" Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Exit Sub
        'End If
        Return s_filename
    End Function

    '檢核查詢參數，有誤為true
    Sub CHK_SchNGVal(ByRef flag_NG_1 As Boolean)
        'Dim flag_NG_1 As Boolean = False
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        Dim t_yearlist As String = TIMS.GetListText(yearlist)
        If (v_TPlanlist1 = "") Then flag_NG_1 = True
        If (v_yearlist = "") Then flag_NG_1 = True
        If (t_TPlanlist1 = "") Then flag_NG_1 = True 'Return 'objdt1
        If (t_yearlist = "") Then flag_NG_1 = True 'Return 'objdt1
    End Sub

    '下載匯出檔
    Sub Utl_DOWNLOAD1()
        Dim flag_NG_1 As Boolean = False
        Call CHK_SchNGVal(flag_NG_1)
        If flag_NG_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        Dim s_attachmn As String = ""
        Dim s_filename As String = ""
        s_filename = GET_FILENAME1(s_attachmn)
        If s_filename = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Dim s_MapFTXLSXPath As String = TIMS.Utl_GetConfigSet("MapFTXLSXPath")
        If s_MapFTXLSXPath = "" Then s_MapFTXLSXPath = "~/XLSX/"
        Dim s_mapfile As String = Server.MapPath(s_MapFTXLSXPath & s_filename)
        If Not IO.File.Exists(s_mapfile) Then
            Dim s_ERR As String = TIMS.cst_NODATAMsg1
            If s_filename <> "" Then s_ERR = (TIMS.cst_NODATAMsg1 & s_filename)
            If s_attachmn <> "" Then s_ERR = (TIMS.cst_NODATAMsg1 & s_attachmn)
            Common.MessageBox(Me, s_ERR)
            Exit Sub
        End If

        Dim Response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        Response.ClearContent()
        Response.Clear()
        Response.ContentType = "application/vnd.ms-excel" '"text/plain"
        Response.AddHeader("Content-Disposition", String.Format("attachment; filename={0};", s_attachmn))
        Response.TransmitFile(s_mapfile)
        Response.Flush()
        Response.End()
    End Sub

    '下載檔案，暫時寫死
    Protected Sub bt_DOWNLOADFILE_Click(sender As Object, e As EventArgs) Handles bt_DOWNLOADFILE.Click
        Utl_DOWNLOAD1()
    End Sub

    'Protected Sub bt_EXPORT_Click(sender As Object, e As EventArgs) Handles bt_EXPORT.Click
    '    Utl_EXPORT1()
    'End Sub

End Class