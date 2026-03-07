Public Class CO_01_003
    Inherits AuthBasePage 'System.Web.UI.Page

    'get_ttqs_1 排程
    'view' V_SENDVER V_RESULT 
    'table' ORG_TTQS2     x不使用了 ORG_TTQS    
    'table' ORG_TTQSVER x不使用了 
    'ORG_TTQSLOCK / ORG_TTQS2

    Dim gflag_test As Boolean = False 'TIMS.sUtl_ChkTest() '測試
    'flag_TTQSLOCK
    Dim flag_TTQSLOCK As Boolean = False '解鎖-非於單位確認之開放時間 'Check_TTQSLOCK()

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        'Dim gflag_test As Boolean = False 'TIMS.sUtl_ChkTest() '測試
        gflag_test = TIMS.sUtl_ChkTest() '測試

        If Not IsPostBack Then
            divSch1.Visible = True
            'divEdt1.Visible = False
            'BtnSaveData1.Visible = False
            PageControler1.Visible = False
            DataGridTable.Visible = False
            BtnSaveData1.Visible = False
            msg1.Text = ""
            '(加強操作便利性)
            sCreate1()
        End If

        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
            ddlDISTID.Enabled = False
        End If

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    'center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
        '    'HistoryRID.Attributes("onclick") = "ShowFrame();"
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If

    End Sub

    Sub sCreate1()
        '年度
        SYEARlist = TIMS.GetSyear(SYEARlist)
        Common.SetListItem(SYEARlist, sm.UserInfo.Years)

        '截止時間
        With rblMONTHS
            .Items.Clear()
            .Items.Add(New ListItem("1月", "01"))
            .Items.Add(New ListItem("7月", "07"))
            .RepeatDirection = RepeatDirection.Horizontal
            .AppendDataBoundItems = True
        End With
        Common.SetListItem(rblMONTHS, "01")

        '單位確認'不拘、正確、資料有誤、未確認
        With rblCONFIRM
            .Items.Clear()
            .Items.Add(New ListItem("不拘", "A")) '(A)ALL
            .Items.Add(New ListItem("正確", "Y")) 'Y
            .Items.Add(New ListItem("資料有誤", "N")) 'N
            .Items.Add(New ListItem("未確認", "S")) '(S)SPACE
            .RepeatDirection = RepeatDirection.Horizontal
            .AppendDataBoundItems = True
        End With
        Common.SetListItem(rblCONFIRM, "A")

        '評核版本
        ddlSENDVER = TIMS.Get_SENDVER_TS(ddlSENDVER, objconn)
        '評核結果
        ddlRESULT = TIMS.Get_RESULT_TS(ddlRESULT, objconn)

        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        If (ddlDISTID.Items.FindByValue("000") IsNot Nothing) Then ddlDISTID.Items.Remove(ddlDISTID.Items.FindByValue("000"))
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        OrgPlanKind = TIMS.Get_RblOrgPlanKind(OrgPlanKind, objconn)
        Common.SetListItem(OrgPlanKind, "G")

        OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)

        '依登入者 LID 判斷是否可自由輸入
        Select Case sm.UserInfo.LID
            Case 0'署
            Case 1 '分署
                Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
                ddlDISTID.Enabled = False
            Case 2 '單位'委訓單位動作
                '依登入者機構判斷計畫種類
                Dim sql As String
                Dim dr As DataRow
                sql = "Select OrgName,ComIDNO,OrgKind2 FROM ORG_ORGINFO where OrgID = '" & sm.UserInfo.OrgID & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr IsNot Nothing Then
                    OrgName.Text = dr("OrgName")
                    COMIDNO.Text = dr("ComIDNO")
                    Select Case Convert.ToString(dr("OrgKind2"))
                        Case "G", "W"
                            Common.SetListItem(OrgPlanKind, Convert.ToString(dr("OrgKind2")))
                    End Select
                Else
                    OrgName.Text = TIMS.cst_ErrorMsg12
                    COMIDNO.Text = TIMS.cst_ErrorMsg12
                End If
                OrgName.Enabled = False
                COMIDNO.Enabled = False
                OrgPlanKind.Enabled = False
        End Select

        '登入年度轉民國年份
        'Years.Value = sm.UserInfo.Years - 1911
        '選擇清除工作
        'SelectValue.Value = ""
        DataGridTable.Visible = False
    End Sub

    ''' <summary>
    ''' false 無開放單位確認
    ''' </summary>
    ''' <returns></returns>
    Function Check_TTQSLOCK() As Boolean
        Dim s_TTQSLOCKMsg1 As String = TIMS.Get_TTQSLOCKMsg1(objconn)
        If s_TTQSLOCKMsg1 = "" Then
            'Dim v_msg_1 As String = "目前無開放單位確認!"
            'Common.MessageBox(Me, v_msg_1)
            Return False
        End If
        Return True
    End Function

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        'Call sClearlist1()
        Call sSearch1()
    End Sub

    ''' <summary>1.查詢／2.匯出 SQL</summary>
    ''' <param name="in_parms"></param>
    ''' <returns></returns>
    Public Function Get_dtORGTTQS2(ByVal in_parms As Hashtable) As DataTable
        Dim vEXP As String = TIMS.GetMyValue2(in_parms, "EXP")
        Dim vExpType As String = TIMS.GetMyValue2(in_parms, "ExpType")

        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID) 'TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vYEARS As String = TIMS.GetListValue(SYEARlist) 'TIMS.ClearSQM(SYEARlist.SelectedValue)
        If vYEARS = "" Then vYEARS = sm.UserInfo.Years
        Dim vMONTHS As String = TIMS.GetListValue(rblMONTHS)
        'Dim vHALFYEAR As String = TIMS.ClearSQM(halfYear.SelectedValue) '1:上年度 /2:下年度

        Dim vORGKIND2 As String = TIMS.GetListValue(OrgPlanKind) 'TIMS.ClearSQM(OrgPlanKind.SelectedValue)
        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND As String = TIMS.GetListValue(OrgKindList) 'TIMS.ClearSQM(OrgKindList.SelectedValue)
        Dim vSENDVER As String = TIMS.GetListValue(ddlSENDVER) 'TIMS.ClearSQM(ddlSENDVER.SelectedValue)
        Dim vRESULT As String = TIMS.GetListValue(ddlRESULT) 'TIMS.ClearSQM(ddlRESULT.SelectedValue)
        Dim vCONFIRM As String = TIMS.GetListValue(rblCONFIRM)

        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        'parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("YEARS", vYEARS)
        'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vDISTID <> "" Then parms.Add("DISTID", vDISTID) 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME)
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf
        If vSENDVER <> "" Then parms.Add("SENDVER", vSENDVER) 'sql &= " AND tr.SENDVER =@SENDVER " & vbCrLf
        If vRESULT <> "" Then parms.Add("RESULT", vRESULT) 'sql &= " AND tr.RESULT =@RESULT " & vbCrLf

        If vMONTHS <> "" Then parms.Add("MONTHS", vMONTHS)
        Select Case vCONFIRM
            Case "Y", "N"
                parms.Add("CONFIRM", vCONFIRM)
        End Select

        'ORG_TTQS2 tr
        Dim sql As String = ""
        sql &= " SELECT tr.OTTID" & vbCrLf
        sql &= " ,tr.TPLANID" & vbCrLf
        sql &= " ,tr.DISTID" & vbCrLf
        sql &= " ,kd.NAME DISTNAME" & vbCrLf
        sql &= " ,tr.ORGID" & vbCrLf
        sql &= " ,tr.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(tr.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,tr.MONTHS" & vbCrLf
        sql &= " ,CASE tr.MONTHS WHEN '01' THEN '1月' WHEN '07' THEN '7月' END MONTHS_N" & vbCrLf
        sql &= " ,tr.APPLIEDRESULT" & vbCrLf
        sql &= " ,tr.COMIDNO" & vbCrLf
        sql &= " ,CONVERT(varchar,tr.SENDDATE,111) SENDDATE" & vbCrLf
        sql &= " ,tr.SENDVER" & vbCrLf
        sql &= " ,tr.RESULT" & vbCrLf
        sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
        sql &= " ,v2.VNAME RESULT_N" & vbCrLf
        sql &= " ,CONVERT(varchar,tr.VALIDDATE,111) VALIDDATE" & vbCrLf
        sql &= " ,tr.EXTLICENS" & vbCrLf
        sql &= " ,tr.GOAL" & vbCrLf '申請目的
        sql &= " ,tr.EXTEND" & vbCrLf
        sql &= " ,CONVERT(varchar,tr.ISSUEDATE,111) ISSUEDATE" & vbCrLf
        sql &= " ,tr.MEMO" & vbCrLf
        sql &= " ,tr.EVALSCOPE" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(tr.EVALSCOPE,'-')='-' then tr.MEMO ELSE concat(tr.MEMO,'(',tr.EVALSCOPE,')') END MEMO2" & vbCrLf

        sql &= " ,CONVERT(varchar,tr.IMPORTDATE,111) IMPORTDATE" & vbCrLf
        sql &= " ,tr.CONFIRM" & vbCrLf
        '訓練單位確認結果--選項：不拘、正確、資料有誤、未確認-分署更新
        sql &= " ,CASE WHEN tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL THEN " & vbCrLf
        sql &= "   case when tr.CONFIRM='N' then '資料有誤'" & vbCrLf
        sql &= "   when tr.CONFIRM='Y' then '正確' ELSE '未確認' END" & vbCrLf
        sql &= "  ELSE '未確認' END" & vbCrLf
        sql &= " +CASE WHEN tr.RENEWACCT IS NOT NULL AND tr.RENEWDATE IS NOT NULL THEN '(分署更新)' ELSE '' END CONFIRM_N" & vbCrLf

        sql &= " ,tr.CONFIRMACCT" & vbCrLf
        sql &= " ,tr.CONFIRMDATE" & vbCrLf
        sql &= " ,tr.REASON1" & vbCrLf
        sql &= " ,tr.APPLIEDRESULT" & vbCrLf
        '分署審核狀況
        sql &= " ,case when tr.APPLIEDRESULT='Y' THEN '已確認' END APPLIEDRESULT_N" & vbCrLf
        sql &= " ,tr.APPLIEDACCT" & vbCrLf
        sql &= " ,tr.APPLIEDDATE" & vbCrLf
        sql &= " ,tr.LOCKACCT" & vbCrLf
        sql &= " ,tr.LOCKDATE" & vbCrLf
        sql &= " ,CASE WHEN tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL THEN 'Y' END IS_LOCK" & vbCrLf
        sql &= " ,CASE WHEN tr.RENEWACCT IS NOT NULL AND tr.RENEWDATE IS NOT NULL THEN 'Y' END IS_RENEW" & vbCrLf
        sql &= " ,o.ORGNAME" & vbCrLf
        sql &= " ,o.ORGKIND" & vbCrLf
        sql &= " ,k1.NAME ORGKIND_N" & vbCrLf
        sql &= " ,o.ORGKIND1" & vbCrLf
        sql &= " ,o.ORGKIND2" & vbCrLf
        'sql &= " ,o.ORGZIPCODE,o.ORGZIPCODE6W" & vbCrLf
        sql &= " ,o.ORGADDRESS" & vbCrLf

        If vEXP = "Y" Then
            Select Case vExpType
                Case "ODS"
                    '匯出重組SQL
                    'Dim sql As String = ""
                    sql = ""
                    sql &= " SELECT CASE WHEN o.ORGKIND2='G' THEN '產投' ELSE '自主' END 計畫" & vbCrLf
                    sql &= " ,concat(kt.TypeID2,'-',kt.TypeID2Name) 單位屬性" & vbCrLf
                    sql &= " ,kd.NAME 分署" & vbCrLf
                    sql &= " ,tr.COMIDNO 統一編號" & vbCrLf
                    sql &= " ,o.ORGNAME 訓練機構" & vbCrLf
                    'sql &= " ,tr.SENDVER 評核版別" & vbCrLf
                    sql &= " ,v1.VNAME 評核版別" & vbCrLf
                    'sql &= " ,tr.GOAL 申請目的" & vbCrLf '申請目的
                    'sql &= " ,tr.RESULT 評核結果" & vbCrLf
                    sql &= " ,v2.VNAME 評核結果" & vbCrLf
                    'sql &= " ,CASE WHEN tr.EXTEND='Y' THEN 'V' ELSE '' END 展延" & vbCrLf
                    sql &= " ,tr.EXTLICENS 展延" & vbCrLf
                    sql &= " ,CONVERT(varchar,tr.SENDDATE,111) 評核日期" & vbCrLf
                    sql &= " ,CONVERT(varchar,tr.ISSUEDATE,111) 發文日期" & vbCrLf
                    sql &= " ,CONVERT(varchar,tr.VALIDDATE,111) 有效期限" & vbCrLf
                    'sql &= " ,tr.MEMO 備註" & vbCrLf
                    'sql &= " ,tr.MEMO ""TTQS訓練機構名稱""" & vbCrLf
                    'sql &= " ,tr.EVALSCOPE" & vbCrLf
                    'MEMO2 ""TTQS訓練機構名稱""" & vbCrLf
                    sql &= " ,CASE WHEN ISNULL(tr.EVALSCOPE,'-')='-' then tr.MEMO ELSE concat(tr.MEMO,'(',tr.EVALSCOPE,')') END ""TTQS訓練機構名稱""" & vbCrLf

                    '訓練單位確認結果--選項：不拘、正確、資料有誤、未確認-分署更新
                    sql &= " ,CASE WHEN tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL THEN " & vbCrLf
                    sql &= "   case when tr.CONFIRM='N' then '資料有誤'" & vbCrLf
                    sql &= "   when tr.CONFIRM='Y' then '正確' ELSE '未確認' END" & vbCrLf
                    sql &= "  ELSE '未確認' END" & vbCrLf
                    sql &= " +CASE WHEN tr.RENEWACCT IS NOT NULL AND tr.RENEWDATE IS NOT NULL THEN '(分署更新)' ELSE '' END 訓練單位確認結果" & vbCrLf
                    '分署審核狀況
                    sql &= " ,case when tr.APPLIEDRESULT='Y' THEN '已確認' END 分署審核狀況" & vbCrLf
                Case Else
                    '匯出重組SQL
                    'Dim sql As String = ""
                    sql = ""
                    sql &= " SELECT CASE WHEN o.ORGKIND2='G' THEN '產投' ELSE '自主' END 計畫" & vbCrLf
                    sql &= " ,concat(kt.TypeID2,'-',kt.TypeID2Name) 單位屬性" & vbCrLf
                    sql &= " ,kd.NAME 分署" & vbCrLf
                    sql &= " ,CHAR(19)+tr.COMIDNO 統一編號" & vbCrLf
                    sql &= " ,o.ORGNAME 訓練機構" & vbCrLf
                    'sql &= " ,tr.SENDVER 評核版別" & vbCrLf
                    sql &= " ,v1.VNAME 評核版別" & vbCrLf
                    'sql &= " ,tr.GOAL 申請目的" & vbCrLf '申請目的
                    'sql &= " ,tr.RESULT 評核結果" & vbCrLf
                    sql &= " ,v2.VNAME 評核結果" & vbCrLf
                    'sql &= " ,CASE WHEN tr.EXTEND='Y' THEN 'V' ELSE '' END 展延" & vbCrLf
                    sql &= " ,tr.EXTLICENS 展延" & vbCrLf
                    sql &= " ,CHAR(19)+CONVERT(varchar,tr.SENDDATE,111) 評核日期" & vbCrLf
                    sql &= " ,CHAR(19)+CONVERT(varchar,tr.ISSUEDATE,111) 發文日期" & vbCrLf
                    sql &= " ,CHAR(19)+CONVERT(varchar,tr.VALIDDATE,111) 有效期限" & vbCrLf
                    'sql &= " ,tr.MEMO 備註" & vbCrLf
                    'sql &= " ,tr.MEMO ""TTQS訓練機構名稱""" & vbCrLf
                    'sql &= " ,tr.EVALSCOPE" & vbCrLf
                    'MEMO2 ""TTQS訓練機構名稱""" & vbCrLf
                    sql &= " ,CASE WHEN ISNULL(tr.EVALSCOPE,'-')='-' then tr.MEMO ELSE concat(tr.MEMO,'(',tr.EVALSCOPE,')') END ""TTQS訓練機構名稱""" & vbCrLf

                    '訓練單位確認結果--選項：不拘、正確、資料有誤、未確認-分署更新
                    sql &= " ,CASE WHEN tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL THEN " & vbCrLf
                    sql &= "   case when tr.CONFIRM='N' then '資料有誤'" & vbCrLf
                    sql &= "   when tr.CONFIRM='Y' then '正確' ELSE '未確認' END" & vbCrLf
                    sql &= "  ELSE '未確認' END" & vbCrLf
                    sql &= " +CASE WHEN tr.RENEWACCT IS NOT NULL AND tr.RENEWDATE IS NOT NULL THEN '(分署更新)' ELSE '' END 訓練單位確認結果" & vbCrLf
                    '分署審核狀況
                    sql &= " ,case when tr.APPLIEDRESULT='Y' THEN '已確認' END 分署審核狀況" & vbCrLf
            End Select
        End If

        sql &= " FROM dbo.ORG_TTQS2 tr" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO o On o.OrgID=tr.OrgID" & vbCrLf
        sql &= " JOIN dbo.ID_DISTRICT kd On kd.DISTID=tr.DISTID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=tr.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=tr.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 On k1.ORGTYPEID=o.ORGKIND" & vbCrLf
        sql &= " LEFT JOIN dbo.KEY_ORGTYPE1 kt on kt.ORGTYPEID1=o.ORGKIND1" & vbCrLf
        sql &= " WHERE tr.TPLANID=@TPLANID" & vbCrLf

        sql &= " And tr.YEARS=@YEARS" & vbCrLf
        'If vHALFYEAR <> "" Then sql &= " And t.HALFYEAR = @HALFYEAR " & vbCrLf '1:上年度 /2:下年度
        If vDISTID <> "" Then sql &= " And tr.DISTID=@DISTID" & vbCrLf
        'If vDISTID <> "" Then sql &= " And EXISTS (Select 'X' FROM dbo.ORG_TTQS x where x.ORGID=t.ORGID AND x.YEARS=t.YEARS AND x.DISTID=@DISTID)" & vbCrLf
        If vORGNAME <> "" Then sql &= String.Concat(" And o.ORGNAME Like '%", vORGNAME, "%'") & vbCrLf
        If vCOMIDNO <> "" Then sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf
        If vSENDVER <> "" Then sql &= " AND tr.SENDVER =@SENDVER" & vbCrLf
        If vRESULT <> "" Then sql &= " AND tr.RESULT =@RESULT" & vbCrLf
        If vMONTHS <> "" Then sql &= " AND tr.MONTHS =@MONTHS" & vbCrLf
        Select Case vCONFIRM
            Case "Y", "N"
                sql &= " AND tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL" & vbCrLf
                sql &= " AND tr.CONFIRM =@CONFIRM " & vbCrLf
            Case "S" 'SPACE/NULL
                sql &= " AND tr.LOCKACCT IS NULL AND tr.LOCKDATE IS NULL" & vbCrLf
                'sql &= " AND tr.CONFIRM IS NULL" & vbCrLf
                'Case Else sql &= " AND tr.LOCKACCT IS NOT NULL AND tr.LOCKDATE IS NOT NULL" & vbCrLf
        End Select
        sql &= " ORDER BY tr.OTTID" & vbCrLf 't.OTSID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    ''' <summary>查詢</summary>
    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        flag_TTQSLOCK = Check_TTQSLOCK() 'Dim flag_TTQSLOCK As Boolean = Check_TTQSLOCK()
        'BtnSaveData1.Visible = False
        PageControler1.Visible = False
        DataGridTable.Visible = False
        BtnSaveData1.Visible = False
        msg1.Text = "查無資料"

        '--SQL--
        Dim dt As DataTable
        Dim in_parms As New Hashtable

        dt = Get_dtORGTTQS2(in_parms)

        'BtnSaveData1.Visible = False
        'PageControler1.Visible = False
        'DataGridTable.Visible = False
        'msg1.Text = "查無資料"
        If dt.Rows.Count = 0 Then Exit Sub

        'BtnSaveData1.Visible = True
        PageControler1.Visible = True
        DataGridTable.Visible = True
        BtnSaveData1.Visible = True
        msg1.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    Sub CheckData1(ByRef s_ERRMSG As String)
        Dim i_cb As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            'Dim drv As DataRowView = e.Item.DataItem
            Dim Hid_OTTID As HiddenField = eItem.FindControl("Hid_OTTID")
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            checkbox1.Value = TIMS.ClearSQM(checkbox1.Value)
            Dim vOTTID As String = TIMS.ClearSQM(Hid_OTTID.Value)
            If (checkbox1.Value <> vOTTID) Then vOTTID = ""

            'checkbox1.Value = Convert.ToString(drv("OTTID"))
            Dim flagCanSave1 As Boolean = False
            If vOTTID <> "" AndAlso checkbox1.Checked Then flagCanSave1 = True

            If flagCanSave1 Then i_cb += 1
        Next
        If i_cb = 0 Then
            s_ERRMSG &= "請勾選有效資料 後再按確認鈕!" & vbCrLf
        End If
    End Sub

    ''' <summary>儲存-審核確認</summary>
    Sub sSaveData1()
        '---updata
        Dim sql As String = ""
        sql &= " UPDATE ORG_TTQS2" & vbCrLf
        sql &= " SET APPLIEDRESULT='Y'" & vbCrLf '審核確認
        sql &= " ,APPLIEDACCT=@APPLIEDACCT ,APPLIEDDATE=GETDATE()" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE OTTID=@OTTID" & vbCrLf
        Dim u_sql As String = sql

        For Each eItem As DataGridItem In DataGrid1.Items
            'Dim drv As DataRowView = e.Item.DataItem
            Dim Hid_OTTID As HiddenField = eItem.FindControl("Hid_OTTID")
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            checkbox1.Value = TIMS.ClearSQM(checkbox1.Value)
            Dim vOTTID As String = TIMS.ClearSQM(Hid_OTTID.Value)
            If (checkbox1.Value <> vOTTID) Then vOTTID = ""

            'checkbox1.Value = Convert.ToString(drv("OTTID"))
            Dim flagCanSave1 As Boolean = False
            If vOTTID <> "" AndAlso checkbox1.Checked Then flagCanSave1 = True

            If flagCanSave1 Then
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("APPLIEDACCT", sm.UserInfo.UserID)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                parms.Add("OTTID", Val(vOTTID))
                DbAccess.ExecuteNonQuery(u_sql, objconn, parms)
            End If
        Next

    End Sub

    Sub unlock_s1(ByRef s_parms As Hashtable)
        Dim s_OTTID As String = TIMS.GetMyValue2(s_parms, "OTTID")
        Dim s_ORGID As String = TIMS.GetMyValue2(s_parms, "ORGID")
        Dim s_COMIDNO As String = TIMS.GetMyValue2(s_parms, "COMIDNO")

        Dim parms As New Hashtable
        parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        parms.Add("OTTID", s_OTTID)
        parms.Add("ORGID", s_ORGID)
        parms.Add("COMIDNO", s_COMIDNO)
        Dim sql As String = ""
        sql &= " UPDATE ORG_TTQS2" & vbCrLf
        sql &= " SET LOCKACCT=null ,LOCKDATE=null" & vbCrLf '解鎖
        sql &= " ,APPLIEDRESULT=null" & vbCrLf '審核-確認-狀況清除
        'sql &= " ,APPLIEDACCT=@APPLIEDACCT" & vbCrLf '保留審核者 
        'sql &= " ,APPLIEDDATE=GETDATE()" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE OTTID=@OTTID" & vbCrLf
        sql &= " AND ORGID=@ORGID" & vbCrLf
        sql &= " AND COMIDNO=@COMIDNO" & vbCrLf
        Dim u_sql As String = sql

        DbAccess.ExecuteNonQuery(u_sql, objconn, parms)
    End Sub


    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        Dim s_OTTID As String = TIMS.GetMyValue(sCmdArg, "OTTID")
        Dim s_ORGID As String = TIMS.GetMyValue(sCmdArg, "ORGID")
        Dim s_COMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
        If s_OTTID = "" OrElse s_ORGID = "" OrElse s_COMIDNO = "" Then Return

        Select Case e.CommandName
            Case "RENEW"
            Case "UNLOCK"
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("OTTID", s_OTTID)
                parms.Add("ORGID", s_ORGID)
                parms.Add("COMIDNO", s_COMIDNO)
                unlock_s1(parms)
                sSearch1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"

            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim Hid_OTTID As HiddenField = e.Item.FindControl("Hid_OTTID")
                Dim checkbox1 As HtmlInputCheckBox = e.Item.FindControl("checkbox1")
                checkbox1.Value = Convert.ToString(drv("OTTID"))

                Dim lbtRENEW As LinkButton = e.Item.FindControl("lbtRENEW")
                Dim lbtUNLOCK As LinkButton = e.Item.FindControl("lbtUNLOCK")
                'checkbox1.Disabled = True
                'If (Convert.ToString(drv("CONFIRM")) = "Y") Then checkbox1.Disabled = False
                Dim flag_CONFIRM_lock As Boolean = False '(鎖)
                '(已確認)(已鎖定)
                If (Convert.ToString(drv("CONFIRM")) = "Y") AndAlso Convert.ToString(drv("IS_LOCK")) = "Y" Then flag_CONFIRM_lock = True '(解)
                '(分署更新)(未鎖定)
                If (Convert.ToString(drv("IS_RENEW")) = "Y") AndAlso Convert.ToString(drv("IS_LOCK")) = "" Then flag_CONFIRM_lock = True '(解)

                checkbox1.Disabled = If(flag_CONFIRM_lock, False, True)
                Dim vv_msg_1 As String = ""
                If checkbox1.Disabled Then vv_msg_1 = String.Format("訓練單位確認結果:{0}", Convert.ToString(drv("CONFIRM_N")))
                If (Convert.ToString(drv("IS_RENEW")) = "Y") Then vv_msg_1 &= "(分署更新)"
                If vv_msg_1 <> "" Then TIMS.Tooltip(checkbox1, vv_msg_1)

                Const cst_cells_訓練單位確認結果 As Integer = 15
                Dim v_REASON1 As String = TIMS.ClearSQM(drv("REASON1"))
                If (Convert.ToString(drv("CONFIRM")) = "N") AndAlso v_REASON1 <> "" Then
                    '【訓練單位確認結果】滑鼠移到上方，可顯示資料有誤的原因!
                    Dim v_msg_1 As String = String.Format("原因:{0}", v_REASON1)
                    Dim js_msg_1 As String = Common.GetJsString(v_msg_1)

                    TIMS.Tooltip(e.Item.Cells(cst_cells_訓練單位確認結果), v_msg_1)
                    e.Item.Cells(cst_cells_訓練單位確認結果).Attributes.Add("Onclick", String.Format("javascript:return alert('{0}');", js_msg_1))
                End If
                'If gflag_test Then
                '    Dim v_msg_1 As String = String.Format("原因(test):{0}", "(test)")
                '    TIMS.Tooltip(e.Item.Cells(cst_cells_訓練單位確認結果), v_msg_1)
                '    'e.Item.Cells(cst_cells_訓練單位確認結果).Attributes.Add("Onclick", String.Format("javascript:return confirm('{0}');", v_msg_1))
                '    e.Item.Cells(cst_cells_訓練單位確認結果).Attributes.Add("Onclick", String.Format("javascript:return alert('{0}');", v_msg_1))
                '    e.Item.Cells(cst_cells_訓練單位確認結果 - 1).Attributes.Add("Onclick", String.Format("javascript:alert('{0}');", v_msg_1))
                '    e.Item.Cells(cst_cells_訓練單位確認結果 - 2).Attributes.Add("Onclick", String.Format("javascript:blockAlert('{0}');", v_msg_1))
                'End If
                If Not flag_TTQSLOCK Then
                    lbtUNLOCK.Enabled = False
                    Dim v_msg_1 As String = "非於單位確認之開放時間"
                    TIMS.Tooltip(lbtUNLOCK, v_msg_1)
                End If
                If flag_TTQSLOCK Then
                    'IS_LOCK
                    lbtUNLOCK.Enabled = False
                    If (Convert.ToString(drv("IS_LOCK")) = "Y") Then lbtUNLOCK.Enabled = True
                    If Not lbtUNLOCK.Enabled Then
                        Dim v_msg_1 As String = "未鎖定無須解鎖" 'String.Format("未鎖定無須解鎖:{0}", Convert.ToString(drv("CONFIRM_N")))
                        TIMS.Tooltip(lbtUNLOCK, v_msg_1)
                    End If
                End If

                Hid_OTTID.Value = Convert.ToString(drv("OTTID"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OTTID", Convert.ToString(drv("OTTID")))
                TIMS.SetMyValue(sCmdArg, "ORGID", Convert.ToString(drv("ORGID")))
                TIMS.SetMyValue(sCmdArg, "COMIDNO", Convert.ToString(drv("COMIDNO")))
                lbtRENEW.CommandArgument = sCmdArg
                lbtUNLOCK.CommandArgument = sCmdArg
                Dim s_open_ss1 As String = "open_CO01003sch1('{0}','{1}','{2}');return false;"
                Dim s_open_ss2 As String = String.Format(s_open_ss1, Convert.ToString(drv("OTTID")), Convert.ToString(drv("ORGID")), Convert.ToString(drv("COMIDNO")))
                lbtRENEW.Attributes("onclick") = s_open_ss2

        End Select

    End Sub

    'Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
    '    sSaveData1()
    '    divSch1.Visible = True
    '    'divEdt1.Visible = False
    '    sm.LastResultMessage = "儲存完畢"
    'End Sub

    ''' <summary> 匯出 </summary>
    Sub sExprot2()
        Dim dtXls As DataTable = Nothing

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS

        Dim in_parms As New Hashtable
        in_parms.Clear()
        in_parms.Add("EXP", "Y") '匯出查詢條件
        in_parms.Add("ExpType", v_ExpType) '匯出查詢條件

        dtXls = Get_dtORGTTQS2(in_parms)

        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If


        '匯出excel /ods
        Dim strFilename1 As String = "ExpFile" & TIMS.GetRnd6Eng

        'sPattern = "序號,訓練單位名稱,班別名稱,提案意願順序,班級課程流水號,訓練時數,訓練人數,每人訓練費用(元),每班總訓練費(元),每班總補助費(元),開訓日期,結訓日期"
        'sPattern &= ",訓練業別編碼,訓練業別,訓練職能編碼,訓練職能,是否為學分班(Y/N),單位屬性,縣市別(辦訓地),聯絡人,聯絡電話,立案縣市,課程分類,統一編號,備註"

        'sColumn = "SEQNUM,ORGNAME,CLASSCNAME,FIRSTSORT,PSNO28,THOURS,TNUM,ONECOST,TOTALCOST,DEFGOVCOST,STDATE,FTDATE"
        'sColumn &= ",GOVCLASS,GOVCLASSNAME,CCID,CCNAME,POINTYN,ORGTYPE,CTNAME,CONTACTNAME,PHONE,CTNAME2,KNAME12,COMIDNO,MEMO1"

        'Dim sPatternA() As String = Split(sPattern, ",")
        'Dim sColumnA() As String = Split(sColumn, ",")
        'Dim iColSpanCount As Integer = sColumnA.Length
        Dim sTitle1 As String = "最近一次TTQS評核結果等級"

        Dim parms As New Hashtable
        parms.Add("ExpType", v_ExpType) 'EXCEL/PDF/ODS
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        'parms.Add("TitleColSpanCnt", iColSpanCount)
        'parms.Add("sPatternA", sPatternA)
        'parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
    End Sub

    Protected Sub btnExp1_Click(sender As Object, e As EventArgs) Handles btnExp1.Click
        Call sExprot2()
    End Sub

    ''' <summary> 審核確認 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Dim s_ERRMSG As String = ""
        Call CheckData1(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        sSaveData1()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
