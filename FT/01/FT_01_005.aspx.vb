Public Class FT_01_005
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    'ADP_FDOWNLOAD
    'Batch\dbt_20210217 'dbt_20210217 '薪資分析明細表

    '確定性需求
    '預計完成日期：2023/02/24 (後面228連假)
    '系統：在職系統    '計畫：產投、自辦在職、充飛、區域據點    '功能：首頁>>定版數據統計表>>薪資分析明細表
    '需求說明：    '為了解參訓學員訓後成效，於學員參訓時及結訓後3、6、9、12個月與勞保局勾稽投保薪資(勞保)、實際薪資(勞退)，
    '並匯出raw data明細表及薪資分析大表，於112年3月1日前完成。
    '詳細說明：     '新增  薪資分析明細表 功能，路徑：首頁>>定版數據統計表>>薪資分析明細表
    '1.介面設計 '【訓練計畫】：產投、充飛、自辦在職、區域據點 (企委不用)
    '【學員參訓年度】：從110開始往後~今年度-2 (都是抓前年度的，所以年度顯示到今年-2) 'ex:113年，顯示110、111

    Const cst_FTNUM_005 As String = "005"

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
        Dim where_ff3 As String = "TPLANID IN ('06','28','54','70')"
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, where_ff3, objconn)
        Common.SetListItem(TPlanlist1, sm.UserInfo.TPlanID)

        TPlanlist1.Enabled = If(TPlanlist1.SelectedValue <> "", False, True)
        If (Not TPlanlist1.Enabled) Then TIMS.Tooltip(TPlanlist1, "限定計畫")

        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        'TPlanlist1.Enabled = False
        'If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        Dim iSYears As Integer = 2021
        Dim iEYearsNowb1 As Integer = Year(Now) - 2
        Dim iEYears As Integer = If(iEYearsNowb1 > iSYears, iEYearsNowb1, iSYears)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, "")

        'Dim vDEF_N As String = "5月1日"
        'Dim vDEF As String = "0501" ',"5月1日"
        'rbl_BDATAVER.Items.Clear()
        'rbl_BDATAVER.Items.Add(New ListItem(vDEF_N, vDEF))
        'Common.SetListItem(rbl_BDATAVER, vDEF)
    End Sub

    ''' <summary>年度+2</summary>
    ''' <returns></returns>
    Public Shared Function GET_BYEAR1(ByRef ddlYEARS As DropDownList) As String
        Dim v_yearlist As String = TIMS.GetListValue(ddlYEARS)
        Return Convert.ToString(If(v_yearlist <> "" AndAlso Val(v_yearlist) > 0, Val(v_yearlist) + 2, Val(v_yearlist)))
    End Function

    ''' <summary>'檢核資料表-取得匯入匯出檔名</summary>
    ''' <param name="s_attachmn"></param>
    ''' <returns></returns>
    Function GET_FDOWNLOAD(ByRef s_attachmn As String) As String
        'Dim s_attachmn As String = ""
        Dim s_filename As String = ""
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'Dim v_BYEAR As String = GET_BYEAR1(yearlist)

        Dim parms As New Hashtable From {
            {"FTNUM", cst_FTNUM_005}, '
            {"TPLANID", v_TPlanlist1},
            {"BYEAR", v_yearlist}
        }
        'parms.Add("BMONTH", v_monthlist)
        'parms.Add("BDATAVER", v_BDATAVER) '版控

        Dim sql As String = ""
        sql &= " SELECT FDID,FNAME1 ,FNAMEIN ,FNAMEOUT ,FTNUM ,EXPDATE ,TPLANID" & vbCrLf
        sql &= " ,BYEAR ,BMONTH ,BDATAVER ,FSN,FSTATUS,FCOUNT" & vbCrLf 'sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_FDOWNLOAD" & vbCrLf
        sql &= " WHERE FTNUM=@FTNUM AND TPLANID=@TPLANID AND BYEAR=@BYEAR" & vbCrLf
        'sql &= " AND BMONTH=@BMONTH " & vbCrLf
        'sql &= " AND BDATAVER=@BDATAVER " & vbCrLf
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
        '若有查詢到值，下面可略過
        Dim s_filename As String = GET_FDOWNLOAD(s_attachmn)
        If s_filename <> "" AndAlso s_attachmn <> "" Then Return s_filename

        Dim s_TPLAN_N As String = ""
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        'Dim t_yearlist As String = TIMS.GetListText(yearlist)
        'Dim s_attachmn As String = ""
        Dim YEARS_ROC As String = CStr(Val(v_yearlist) - 1911)
        '70/54/28/06
        '112年_區域據點_薪資分析明細表
        '112年_產投與充飛計畫(公務ECFA學員)_薪資分析明細表
        '112年_產投(就安就保學員)_薪資分析明細表
        '112年_在職訓練_薪資分析明細表
        Const cst_title005 As String = "薪資分析明細表"
        Select Case v_TPlanlist1
            Case "06"
                s_TPLAN_N = "在職訓練"
            Case "28"
                s_TPLAN_N = "產投(就安就保學員)"
            Case "54"
                s_TPLAN_N = "產投與充飛計畫(公務ECFA學員)"
            Case "70"
                s_TPLAN_N = "區域據點"
        End Select

        Dim fileN1 As String = String.Format("{0}年_{1}_{2}", YEARS_ROC, s_TPLAN_N, cst_title005)
        s_attachmn = String.Concat(fileN1, ".xlsx")
        Return s_filename
    End Function

    '檢核查詢參數，有誤為true
    Sub CHK_SchNGVal(ByRef flag_NG_1 As Boolean)
        'Dim flag_NG_1 As Boolean = False
        'Dim TPLAN As String = ""
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        'Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)
        Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        Dim t_yearlist As String = TIMS.GetListText(yearlist)
        If (v_TPlanlist1 = "") Then flag_NG_1 = True 'Return 'objdt1
        If (v_yearlist = "") Then flag_NG_1 = True
        'If (v_BDATAVER = "") Then flag_NG_1 = True
        If (t_TPlanlist1 = "") Then flag_NG_1 = True 'Return 'objdt1
        If (t_yearlist = "") Then flag_NG_1 = True 'Return 'objdt1
        'If flag_NG_1 Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Exit Sub
        'End If
    End Sub

    ''' <summary>下載匯出檔</summary>
    Sub Utl_DOWNLOAD1()
        Dim flag_NG_1 As Boolean = False
        Call CHK_SchNGVal(flag_NG_1)
        If flag_NG_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return ' Exit Sub
        End If

        Dim s_attachmn As String = "" '取得匯入匯出檔名
        Dim s_filename As String = "" '實際檔案名稱
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

End Class