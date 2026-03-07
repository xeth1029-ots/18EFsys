Public Class FT_01_004
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    'ADP_FDOWNLOAD
    'Batch\dbt_20210217 'dbt_20210217 '當年度參訓學員明細

    '增修需求(提案單)
    '需求編號：OJT-22050301
    '處理等級：普
    '預計完成日期：2022/07/20
    '提案人：發展署 - 金麗芬
    '開發完成後，請先上版至測試環境，待提案人確認OK後，再上版至正式環境
    '以下需求，有疑問再請提出討論，感謝
    '======================================================
    '系統：在職系統
    '計畫：在職進修訓練(分署自辦在職)
    '功能：首頁>>定版數據統計表>>當年度參訓學員明細    (新增功能路徑)
    '需求：
    '參考產投模式，新增自辦的 定版當年度參訓學員明細 功能。
    '資料時間、範圍：
    '每年 5/1 固定撈取前一個年度所有參訓學員資料（含在訓、結訓、離訓、退訓）。       (資料5/1凌晨產出)
    '所需匯出欄位：
    '年度、計畫名稱、轄區、班級名稱、期別、訓練時數、訓練人數、開訓日期、結訓日期、上課縣市別、訓練職類(大類)、訓練職類(中類)、
    '訓練職類(小類)、通俗職類(大類)、通俗職類(小類)、訓練時段、政策性產業、新南向政策、報名開始日期、報名結束日期、甄試日期、
    '報名人數、甄試人數、錄取人數、學員姓名、身分證統一編號、出生日期、性別、身分別、原屬國籍、主要參訓身分別、年齡、年齡級距、
    '最高學歷、學校名稱、科系、畢業狀況、軍種、服役單位名稱、服役起日期、服役迄日期、服役單位電話、通訊地址、戶籍地址、聯絡電話_日、聯絡電話_夜、
    '行動電話、電子信箱、預算別、保險證號、投保單位、投保薪資、實際薪資(勞退)、目前任職公司名稱、公司統一編號、目前任職部門、職務、
    '參訓狀態(在訓/結訓/離訓/退訓)、離訓日期、退訓日期、訓練期末學員滿意度各題項所填資料。
    '(PS：訓練期末學員滿意度各題題目及所填資料，可參考附件(文案)，另提供產投定版之年度參訓學員明細供參考。)
    '檔案為5/1定版數據，以Excel檔案格式供下載，檔名範例為：110_在職進修訓練_0501_年度參訓學員明細 .xlsx

    Const cst_FTNUM_004 As String = "004"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        msg.Text = ""
        If Not Me.IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        Dim where_ff3 As String = String.Format("TPLANID IN ('{0}')", TIMS.Cst_TPlanID06)
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, where_ff3, objconn)
        Common.SetListItem(TPlanlist1, TIMS.Cst_TPlanID06)
        TPlanlist1.Enabled = If(TPlanlist1.SelectedValue <> "", False, True)
        If (Not TPlanlist1.Enabled) Then TIMS.Tooltip(TPlanlist1, "限定計畫")

        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        'TPlanlist1.Enabled = False
        'If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        Dim iSYears As Integer = 2021
        Dim iEYearsNowb1 As Integer = Year(Now) - 1
        Dim iEYears As Integer = If(iEYearsNowb1 > iSYears, iEYearsNowb1, iSYears)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, "")

        Dim vDEF_N As String = "5月1日"
        Dim vDEF As String = "0501" ',"5月1日"
        rbl_BDATAVER.Items.Clear()
        rbl_BDATAVER.Items.Add(New ListItem(vDEF_N, vDEF))
        Common.SetListItem(rbl_BDATAVER, vDEF)
    End Sub

    '檢核資料表-取得匯入匯出檔名
    Function GET_FDOWNLOAD(ByRef s_attachmn As String) As String
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)
        'Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        'Dim t_yearlist As String = TIMS.GetListText(yearlist)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = ""

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("FTNUM", cst_FTNUM_004) '
        parms.Add("TPLANID", v_TPlanlist1)
        parms.Add("BYEAR", v_yearlist)
        'parms.Add("BMONTH", v_monthlist)
        parms.Add("BDATAVER", v_BDATAVER) '版控

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT FDID,FNAME1 ,FNAMEIN ,FNAMEOUT" & vbCrLf
        sql &= " ,FTNUM ,EXPDATE ,TPLANID" & vbCrLf
        sql &= " ,BYEAR ,BMONTH ,BDATAVER" & vbCrLf
        sql &= " ,FSN,FSTATUS,FCOUNT" & vbCrLf
        'sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_FDOWNLOAD" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND FTNUM=@FTNUM " & vbCrLf
        sql &= " AND TPLANID=@TPLANID " & vbCrLf
        sql &= " AND BYEAR=@BYEAR " & vbCrLf
        'sql &= " AND BMONTH=@BMONTH " & vbCrLf
        sql &= " AND BDATAVER=@BDATAVER " & vbCrLf
        sql &= " AND FSTATUS IS NULL" & vbCrLf '排除舊檔案(N/NOUSE)
        sql &= " ORDER BY MODIFYDATE DESC " & vbCrLf
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
        Dim s_TPLAN As String = ""
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)

        Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        Dim t_yearlist As String = TIMS.GetListText(yearlist)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = GET_FDOWNLOAD(s_attachmn)

        '若有查詢到值，下面可略過
        If s_filename <> "" AndAlso s_attachmn <> "" Then Return s_filename

        Dim YEARS_ROC As String = CStr(Val(v_yearlist) - 1911)

        Select Case v_TPlanlist1
            Case "06"
                s_TPLAN = "在職進修訓練"
        End Select

        Dim fileN1 As String = String.Format("{0}_{1}_{2}_年度參訓學員明細", YEARS_ROC, s_TPLAN, v_BDATAVER)
        s_attachmn = String.Format("{0}.xlsx", fileN1)
        Return s_filename
    End Function

    '檢核查詢參數，有誤為true
    Sub CHK_SchNGVal(ByRef flag_NG_1 As Boolean)
        'Dim flag_NG_1 As Boolean = False
        'Dim TPLAN As String = ""
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)
        Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        Dim t_yearlist As String = TIMS.GetListText(yearlist)
        If (v_TPlanlist1 = "") Then flag_NG_1 = True 'Return 'objdt1
        If (v_yearlist = "") Then flag_NG_1 = True
        If (v_BDATAVER = "") Then flag_NG_1 = True
        If (t_TPlanlist1 = "") Then flag_NG_1 = True 'Return 'objdt1
        If (t_yearlist = "") Then flag_NG_1 = True 'Return 'objdt1
        'If flag_NG_1 Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Exit Sub
        'End If
    End Sub

    '下載匯出檔
    Sub Utl_DOWNLOAD1()
        Dim flag_NG_1 As Boolean = False
        Call CHK_SchNGVal(flag_NG_1)
        If flag_NG_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return ' Exit Sub
        End If

        Dim s_attachmn As String = ""
        Dim s_filename As String = ""
        s_filename = GET_FILENAME1(s_attachmn)
        If s_filename = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return ' Exit Sub
        End If

        Dim s_MapFTXLSXPath As String = TIMS.Utl_GetConfigSet("MapFTXLSXPath")
        If s_MapFTXLSXPath = "" Then s_MapFTXLSXPath = "~/XLSX/"
        Dim s_mapfile As String = Server.MapPath(s_MapFTXLSXPath & s_filename)
        If Not IO.File.Exists(s_mapfile) Then
            Dim s_ERR As String = String.Concat("#Not IO.File.Exists :", s_mapfile)
            TIMS.LOG.Debug(s_ERR)

            s_ERR = TIMS.cst_NODATAMsg1
            If s_filename <> "" Then s_ERR = String.Concat(TIMS.cst_NODATAMsg1, vbCrLf, s_filename)
            If s_attachmn <> "" Then s_ERR = String.Concat(TIMS.cst_NODATAMsg1, vbCrLf, s_attachmn)
            Common.MessageBox(Me, s_ERR)
            Return ' Exit Sub
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

    Protected Sub TPlanlist1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TPlanlist1.SelectedIndexChanged

    End Sub

    Protected Sub yearlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles yearlist.SelectedIndexChanged

    End Sub

    'Protected Sub bt_EXPORT_Click(sender As Object, e As EventArgs) Handles bt_EXPORT.Click
    '    Utl_EXPORT1()
    'End Sub

End Class