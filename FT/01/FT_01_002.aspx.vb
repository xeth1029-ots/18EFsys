Public Class FT_01_002
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    'ADP_FDOWNLOAD
    'Batch\dbt_20210217 'dbt_20210217 '當年度參訓學員明細

    '增修需求(提案單)
    '需求編號：OJT-20051205(產投/充飛)
    '處理等級：普
    '預計完成日期：2月底    (總共會有3種定版數據統計表，我陸續提供給你)
    '提案人：發展署 - 金麗芬
    '開發完成後，請先上版至測試環境，待提案人確認OK後，再上版至正式環境
    '以下需求，有疑問再請提出討論，感謝
    '*以資料庫匯出或實體檔案下載方式，待會再電話討論一下!
    '======================================================
    '系統：在職系統
    '計畫：產投、充飛
    '新功能路徑 ：首頁>>定版數據統計表>>當年度參訓學員明細
    '需求：
    '一、資料內容：
    '使用系統排程每年3月1日及4月1日固定撈取一次前一年度開訓日期在1/1-12/31所有參訓學員資料。
    '例如：109/03/1 系統產生一次 開訓日期=108/01/01-108/12/31之所有參訓學員資料，當User選擇【年度】=108、【資料版本】=3月1日版時可下載。
    '二、匯出為excel檔，所需欄位請參考附件。
    '三、資料匯出：
    '年度：依【年度】欄位選擇要匯出之年度
    '計畫：依登入計畫
    '資料版本：依【資料版本】欄位選擇要3/1 or 4/1
    '匯出檔名(範例)：108_產投_當年度參訓學員明細_定版數據+日期.xlsx、108_產投_當年度參訓學員明細_定版數據+日期.ods   (依計畫分：充飛)
    '資料為定版數據(不因撈取時間不同而變動)。
    '四、權限：此功能僅開放給署，不開放給分署及訓練單位。

    Const cst_BDATAVER_3M As String = "0301"
    Const cst_BDATAVER_4M As String = "0401"
    Const cst_BDATAVER_5M As String = "0501" '增加5/1版

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
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, "TPLANID IN ('28','54')", objconn)
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

        Dim vDEF_value As String = cst_BDATAVER_4M '"0401"
        If Now.Month = 3 Then vDEF_value = cst_BDATAVER_3M '"0301"
        If Now.Month = 4 Then vDEF_value = cst_BDATAVER_4M '"0401"
        If Now.Month = 5 Then vDEF_value = cst_BDATAVER_5M '"0501"
        Common.SetListItem(rbl_BDATAVER, vDEF_value)
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

        Const cst_FTNUM_002 As String = "002"
        Dim parms As New Hashtable
        'parms.Clear()
        parms.Add("FTNUM", cst_FTNUM_002)
        parms.Add("TPLANID", v_TPlanlist1)
        parms.Add("BYEAR", v_yearlist)
        'parms.Add("BMONTH", v_monthlist)
        parms.Add("BDATAVER", v_BDATAVER)
        Dim sql As String = ""
        sql &= " SELECT FDID " & vbCrLf '/*PK*/
        sql &= " ,FNAME1 ,FNAMEIN ,FNAMEOUT" & vbCrLf
        sql &= " ,FTNUM ,EXPDATE ,TPLANID" & vbCrLf
        sql &= " ,BYEAR ,BMONTH ,BDATAVER" & vbCrLf
        sql &= " ,FSN,FSTATUS,FCOUNT" & vbCrLf
        'sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_FDOWNLOAD" & vbCrLf
        sql &= " WHERE FTNUM=@FTNUM " & vbCrLf
        sql &= " AND TPLANID=@TPLANID " & vbCrLf
        sql &= " AND BYEAR=@BYEAR " & vbCrLf
        'sql &= " AND BMONTH=@BMONTH " & vbCrLf
        sql &= " AND BDATAVER=@BDATAVER " & vbCrLf
        sql &= " AND FSTATUS IS NULL" & vbCrLf
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
        Dim s_filename As String = ""
        s_filename = GET_FDOWNLOAD(s_attachmn)
        '若有查詢到值，下面可略過
        If s_filename <> "" AndAlso s_attachmn <> "" Then Return s_filename

        Dim YEARS_ROC As String = CStr(Val(v_yearlist) - 1911)
        Select Case v_TPlanlist1
            Case "06"
                s_TPLAN = "在職進修訓練"
            Case "28"
                s_TPLAN = "產業人才投資方案"
            Case "54"
                s_TPLAN = "充電起飛計畫(補助在職勞工及自營作業者參訓)"
        End Select
        Dim fileN1 As String = ""
        fileN1 = String.Format("{0}_{1}_{2}_年度參訓學員明細", YEARS_ROC, s_TPLAN, v_BDATAVER)

        Select Case v_yearlist
            Case "2020"
                Select Case v_TPlanlist1
                    Case "28"
                        Select Case v_BDATAVER
                            Case cst_BDATAVER_3M '"0301"
                                'fileN1 = "109_產業人才投資方案_0301_年度參訓學員明細"
                                s_attachmn = String.Format("{0}.xlsx", fileN1)
                                s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202103161450")
                            Case cst_BDATAVER_4M '"0401"
                                'fileN1 = "109_產業人才投資方案_0401_年度參訓學員明細"
                                s_attachmn = String.Format("{0}.xlsx", fileN1)
                                s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010437")
                        End Select
                    Case "54"
                        Select Case v_BDATAVER
                            Case cst_BDATAVER_3M '"0301"
                                'fileN1 = "109_充電起飛計畫(補助在職勞工及自營作業者參訓)_0301_年度參訓學員明細"
                                s_attachmn = String.Format("{0}.xlsx", fileN1)
                                s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202103161452")
                            Case cst_BDATAVER_4M '"0401"
                                'fileN1 = "109_充電起飛計畫(補助在職勞工及自營作業者參訓)_0401_年度參訓學員明細"
                                s_attachmn = String.Format("{0}.xlsx", fileN1)
                                s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010439")
                        End Select
                End Select
        End Select

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