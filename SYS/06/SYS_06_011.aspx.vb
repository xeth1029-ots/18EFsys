Partial Class SYS_06_011
    Inherits AuthBasePage

    ''SYS_LOGIN_LOG

    'ddlType
    Const cst_dlTYPE1_登入 As String = "LOGIN"
    Const cst_dlTYPE1_登出 As String = "LOGOUT"
    Const cst_dlTYPE1_登入失敗 As String = "LOGINE1"

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        '處理[分頁設定元件]出現的時機
        PageControler1.Visible = False
        If PageControler1.PageDataGrid.Items.Count > 0 Then PageControler1.Visible = True

        If Not IsPostBack Then Call SCreate1() '頁面初始化
    End Sub

    '頁面初始化
    Sub SCreate1()
        Call TIMS.SUB_SET_HR_MI(ddlAccTime1_HH, ddlAccTime1_MM)
        Call TIMS.SUB_SET_HR_MI(ddlAccTime2_HH, ddlAccTime2_MM)
        Common.SetListItem(ddlAccTime1_HH, "00")
        Common.SetListItem(ddlAccTime1_MM, "00")
        Common.SetListItem(ddlAccTime2_HH, "23")
        Common.SetListItem(ddlAccTime2_MM, "59")

        PageControler1.Visible = False
        'ddlType.SelectedValue="1"
        qDATE1.Text = ""
        qDATE2.Text = ""
        'ddlType.SelectedValue=""
        ddlType.SelectedIndex = -1
        Common.SetListItem(ddlType, "")
        qAcc.Text = ""
        qDATE1.Text = TIMS.Cdate17(DateAdd(DateInterval.Month, -1, Date.Today))
        qDATE2.Text = TIMS.Cdate17(Date.Today)

        Dim monthDTV As Date = DateAdd(DateInterval.Month, -1, Now)
        Hid_DTV1.Value = TIMS.Cdate3(monthDTV)
        Dim monthVAL As String = monthDTV.Month
        Dim YearVAL As String = monthDTV.Year

        DDL_YEAR1 = TIMS.GetSyear(DDL_YEAR1)
        Common.SetListItem(DDL_YEAR1, YearVAL)
        DDL_MONTH1 = TIMS.GetSmonth(DDL_MONTH1)
        Common.SetListItem(DDL_MONTH1, monthVAL)
    End Sub

    '準備進行資料查詢作業
    Protected Sub Bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        Call SSearch1() '進行資料查詢作業
    End Sub

    '資料查詢
    Sub SSearch1()
        '「顯示列數」相關設定
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        qDATE1.Text = TIMS.ClearSQM(qDATE1.Text)
        qDATE2.Text = TIMS.ClearSQM(qDATE2.Text)
        qAcc.Text = TIMS.ClearSQM(qAcc.Text)
        qUserName.Text = TIMS.ClearSQM(qUserName.Text)

        '日期
        Dim myqDate1x As String = If(flag_ROC, TIMS.Cdate18(qDATE1.Text), qDATE1.Text).Replace("/", "-")  'edit，by:20181019
        Dim myqDate2x As String = If(flag_ROC, TIMS.Cdate18(qDATE2.Text), qDATE2.Text).Replace("/", "-")  'edit，by:20181019
        'substring(AccessTime,12,5) 23:59
        '時間
        Dim s_qAccTime1 As String = TIMS.GET_DateHM2(ddlAccTime1_HH, ddlAccTime1_MM)
        Dim s_qAccTime2 As String = TIMS.GET_DateHM2(ddlAccTime2_HH, ddlAccTime2_MM)

        Dim s_tmp As String = ""
        '大小順序異常-時間
        If (s_qAccTime1 <> "" AndAlso s_qAccTime2 <> "") AndAlso (s_qAccTime1 >= s_qAccTime2) Then
            s_tmp = s_qAccTime1
            s_qAccTime1 = s_qAccTime2
            s_qAccTime2 = s_tmp
            Common.SetListItem(ddlAccTime1_HH, s_qAccTime1.Substring(0, 2))
            Common.SetListItem(ddlAccTime1_MM, s_qAccTime1.Substring(3, 2))
            Common.SetListItem(ddlAccTime2_HH, s_qAccTime2.Substring(0, 2))
            Common.SetListItem(ddlAccTime2_MM, s_qAccTime2.Substring(3, 2))
        End If
        '大小順序異常-日期
        If (myqDate1x <> "" AndAlso myqDate2x <> "") AndAlso (myqDate1x >= myqDate2x) Then
            s_tmp = myqDate1x
            myqDate1x = myqDate2x
            myqDate2x = s_tmp
            qDATE1.Text = If(flag_ROC, TIMS.Cdate17(myqDate1x), myqDate1x).Trim.Replace("-", "/")
            qDATE2.Text = If(flag_ROC, TIMS.Cdate17(myqDate2x), myqDate2x).Trim.Replace("-", "/")
        End If

        Dim myqDate1 As String = ""
        Dim myqDate2 As String = ""
        If myqDate1x <> "" Then myqDate1 = (myqDate1x & " 00:00:00")
        If myqDate2x <> "" Then myqDate2 = (myqDate2x + " 23:59:59")

        Dim flag_E_MSG As Boolean = False '有訊息者為 True
        Dim myType As String = TIMS.GetListValue(ddlType) '.SelectedValue.Trim
        Select Case myType
            Case cst_dlTYPE1_登入, cst_dlTYPE1_登出
            Case cst_dlTYPE1_登入失敗
                myType = cst_dlTYPE1_登入
                flag_E_MSG = True
        End Select

        Dim myAcc As String = qAcc.Text '.Trim
        Dim myUserName As String = qUserName.Text

        Dim sql As String = ""
        sql &= " SELECT a.UserID,b.NAME UserNAME" & vbCrLf
        sql &= " ,CASE WHEN a.Func='LOGIN' THEN '登入' WHEN a.Func='LOGOUT' THEN '登出' End AS TYPE" & vbCrLf
        sql &= " ,substring(a.AccessTime,1,10) AccessDate" & vbCrLf
        sql &= " ,substring(a.AccessTime,12,8) AccessTime" & vbCrLf
        'sql &= " ,REPLACE(SUBSTRING(a.AccessTime, 1, 19), '.', '') AccessTime" & vbCrLf
        sql &= " ,a.RemoteIP" & vbCrLf
        sql &= " ,a.ResultMessage" & vbCrLf
        sql &= " ,SUBSTRING(a.BrowserInfo, CHARINDEX('Type=', a.BrowserInfo) + len('Type='), CHARINDEX(',', SUBSTRING(a.BrowserInfo, CHARINDEX('Type=', a.BrowserInfo) + len('Type='), 100)) - 1) AS myBrowserInfo" & vbCrLf
        sql &= " FROM dbo.SYS_LOGIN_LOG a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.AUTH_ACCOUNT b WITH(NOLOCK) ON b.ACCOUNT=a.UserID" & vbCrLf ' COLLATE Chinese_Taiwan_Stroke_CI_AS
        sql &= " WHERE 1=1" & vbCrLf

        If s_qAccTime1 <> "" Then sql &= " AND SUBSTRING(a.AccessTime,12,5) >= @qAccTime1" & vbCrLf
        If s_qAccTime2 <> "" Then sql &= " AND SUBSTRING(a.AccessTime,12,5) <= @qAccTime2" & vbCrLf
        If myqDate1 <> "" Then sql &= " AND a.AccessTime>=@qDate1" & vbCrLf
        If myqDate2 <> "" Then sql &= " AND a.AccessTime<=@qDate2" & vbCrLf
        If myType <> "" Then sql &= " AND a.Func=@type1" & vbCrLf
        If flag_E_MSG Then sql &= " AND a.ResultMessage IS NOT NULL" & vbCrLf
        If myAcc <> "" Then sql &= " AND a.UserID LIKE @acct1" & vbCrLf
        If myUserName <> "" Then sql &= " AND b.NAME LIKE @UserName" & vbCrLf

        sql &= " ORDER BY a.AccessTime DESC" & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        If s_qAccTime1 <> "" Then parms.Add("qAccTime1", s_qAccTime1)
        If s_qAccTime2 <> "" Then parms.Add("qAccTime2", s_qAccTime2)
        If myqDate1 <> "" Then parms.Add("qDate1", myqDate1)
        If myqDate2 <> "" Then parms.Add("qDate2", myqDate2)
        If myType <> "" Then parms.Add("type1", myType)
        If myAcc <> "" Then parms.Add("acct1", String.Concat("%", myAcc, "%"))
        If myUserName <> "" Then parms.Add("UserName", String.Concat(myUserName, "%"))

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料"
        tb_Sch.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            tb_Sch.Visible = True
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '調整某欄位所顯示的內容
    Protected Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Lab_AccessTime As Label = e.Item.FindControl("Lab_AccessTime")
                '#Region "依照Web.config的REPLACE2ROC_YEARS(西元年/民國年置換)參數,調整[紀錄時間]顯示格式內容，by:20181019"
                '#End Region
                Dim drv As DataRowView = e.Item.DataItem
                'Dim originLogTime As String=TIMS.cdate3(drv("AccessDate")) '.ToString().Trim
                'Dim tmpLogDate As String=If(Len(originLogTime) > 0, originLogTime.Substring(0, 10), "")
                Dim newLogDate As String = If(flag_ROC, TIMS.Cdate17(drv("AccessDate")), TIMS.Cdate3(drv("AccessDate")))
                Dim tmpLogTime As String = Convert.ToString(drv("AccessTime"))
                'Dim tmpLogTime As String=If(Len(originLogTime) > 0, originLogTime.Substring(11, 8), "")
                Lab_AccessTime.Text = $"{newLogDate} {tmpLogTime}"
                'e.Item.Cells(3).Text = $"{newLogDate} {tmpLogTime}"
        End Select

    End Sub

    ''' <summary>
    ''' 匯出每月登入異常日誌
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_EXP1_Click(sender As Object, e As EventArgs) Handles BTN_EXP1.Click
        EXP_XLXS_V1()
    End Sub

    Sub EXP_XLXS_V1()
        Dim ERRMSG1 As String = ""
        Dim V_DDL_YEAR1 As String = TIMS.GetListValue(DDL_YEAR1)
        If V_DDL_YEAR1 = "" Then ERRMSG1 &= "匯出年份不可為空" & vbCrLf
        Dim V_DDL_MONTH1 As String = TIMS.GetListValue(DDL_MONTH1)
        If V_DDL_MONTH1 = "" Then ERRMSG1 &= "匯出月份不可為空" & vbCrLf
        If ERRMSG1 <> "" Then
            Common.MessageBox(Me, ERRMSG1)
            Return
        End If

        Dim PMS_S1 As New Hashtable From {{"YEARS1", V_DDL_YEAR1}, {"MONTH1", V_DDL_MONTH1}}
        Dim dtH As DataTable = SCH_DATA_1(PMS_S1)
        If TIMS.dtNODATA(dtH) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim S_EXPT1 As String = $"帳號異常登入日誌{V_DDL_YEAR1}年{V_DDL_MONTH1}月"
        Dim dsH As New DataSet
        dsH.Tables.Add(dtH)
        Dim s_fileName1 As String = $"{S_EXPT1}-{TIMS.GetDateNo()}.xlsx"
        DbAccess.CloseDbConn(objconn) : ExpClass1.Utl_Export2_XLSX_Direct(Me, dsH, s_fileName1)
    End Sub

    Private Function SCH_DATA_1(rPMS_1 As Hashtable) As DataTable
        Dim V_YEARS1 As String = TIMS.GetMyValue2(rPMS_1, "YEARS1")
        Dim V_MONTH1 As String = TIMS.GetMyValue2(rPMS_1, "MONTH1")
        Dim V_RID As String = sm.UserInfo.RID

        'DECLARE @RID VARCHAR(4)='A';DECLARE @YEARS1 INT =2025;DECLARE @MONTH1 INT=11;
        Dim PMS_S1 As New Hashtable From {{"RID", V_RID}, {"YEARS1", TIMS.CINT1(V_YEARS1)}, {"MONTH1", TIMS.CINT1(V_MONTH1)}}
        Dim SQL_S1 As String = "
WITH WH1 AS (SELECT RID,HOLDATE,REASON FROM dbo.SYS_HOLIDAY WITH(NOLOCK) WHERE RID=@RID AND YEAR(HOLDATE)=@YEARS1)
,WC1 AS (SELECT a.UserID,a.Func,a.AccessTime,a.RemoteIP ,a.ResultMessage,a.BrowserInfo,b.REASON
FROM dbo.SYS_LOGIN_LOG a WITH(NOLOCK)
LEFT JOIN WH1 b on b.HOLDATE=convert(date,a.AccessTime)
WHERE ((substring(a.AccessTime,12,8)>='00:00:00' AND substring(a.AccessTime,12,8)<='06:00:00') OR b.REASON is not null)
AND YEAR(convert(datetime,AccessTime))=@YEARS1 AND MONTH(convert(datetime,AccessTime))=@MONTH1  )
SELECT a.UserID 帳號,b.NAME 姓名,CASE WHEN a.Func='LOGIN' THEN '登入' WHEN a.Func='LOGOUT' THEN '登出' End 類型,AccessTime 紀錄時間
,a.RemoteIP ADDIP
,SUBSTRING(a.BrowserInfo, CHARINDEX('Type=', a.BrowserInfo) + len('Type='), CHARINDEX(',', SUBSTRING(a.BrowserInfo, CHARINDEX('Type=', a.BrowserInfo) + len('Type='), 100)) - 1) 使用瀏覽器
,dd.DISTNAME 所屬轄區,oo.ORGNAME 單位名稱
FROM WC1 a
JOIN dbo.AUTH_ACCOUNT b WITH(NOLOCK) ON b.ACCOUNT=a.UserID 
JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) on oo.ORGID=b.ORGID
JOIN dbo.ID_PLAN dp WITH(NOLOCK) on dp.PLANID=b.DEFAULT_PLANID
JOIN dbo.V_DISTRICT dd on dd.DISTID=dp.DISTID
"
        Return DbAccess.GetDataTable(SQL_S1, objconn, PMS_S1)
    End Function
End Class
