Partial Class SD_03_012
    Inherits AuthBasePage


    'CLASS_CONFIRM /STUD_CONFIRM
    Const vs_SearchStr1 As String = "vsSearchStr1"
    'Dim au As New cAUTH
    Dim sMemo As String = "" '(查詢原因)
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objConn) '開啟連線
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call Create1()
        End If

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

    End Sub

    '第1次載入
    Sub Create1()
        DataGridTable1.Visible = False
        Msg1.Text = ""

        '取出鍵詞-查詢原因-INQUIRY
        Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objConn, V_INQUIRY)

        PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objConn)
        Common.SetListItem(PlanPoint, "0")

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        If sm.UserInfo.LID <> "2" Then
            '若只有管理一個班級，自動協助帶出班級--by andy 2009-02-25
            Call TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objConn)
        Else
            'Button12_Click(sender, e)
            center.Enabled = False
            Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objConn)
        End If

        If Not Session(vs_SearchStr1) Is Nothing Then
            Dim MyValue As String = ""
            Dim strSearchStr1 As String = Session(vs_SearchStr1)
            'Session(vs_SearchStr1)=Nothing
            MyValue = TIMS.GetMyValue(strSearchStr1, "prg")
            If MyValue = "SD_03_012" Then
                center.Text = TIMS.GetMyValue(strSearchStr1, "center")
                RIDValue.Value = TIMS.GetMyValue(strSearchStr1, "RIDValue")
                TMID1.Text = TIMS.GetMyValue(strSearchStr1, "TMID1")
                OCID1.Text = TIMS.GetMyValue(strSearchStr1, "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "OCIDValue1")
                STDATE1.Text = TIMS.GetMyValue(strSearchStr1, "STDATE1")
                STDATE2.Text = TIMS.GetMyValue(strSearchStr1, "STDATE2")
                FTDATE1.Text = TIMS.GetMyValue(strSearchStr1, "FTDATE1")
                FTDATE2.Text = TIMS.GetMyValue(strSearchStr1, "FTDATE2")
                CONFIRDATE1.Text = TIMS.GetMyValue(strSearchStr1, "CONFIRDATE1") '(建檔日)
                CONFIRDATE2.Text = TIMS.GetMyValue(strSearchStr1, "CONFIRDATE2") '(建檔日)
                MyValue = TIMS.GetMyValue(strSearchStr1, "PlanPoint")
                If MyValue <> "" Then Common.SetListItem(PlanPoint, MyValue)
                'MyValue=TIMS.GetMyValue(strSearchStr1, "PageIndex")
                MyValue = TIMS.GetMyValue(strSearchStr1, "submit")
                If MyValue = "1" Then Call Search1()
            End If
            Session(vs_SearchStr1) = Nothing
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

    End Sub

    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objConn)
        '判斷機構是否只有一個班級 '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = Convert.ToString(dr("trainname"))
        OCID1.Text = Convert.ToString(dr("classname"))
        TMIDValue1.Value = Convert.ToString(dr("trainid"))
        OCIDValue1.Value = Convert.ToString(dr("ocid"))
    End Sub

    Function S_WC1(ByRef parms As Hashtable) As String
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        STDATE1.Text = TIMS.Cdate3(STDATE1.Text)
        STDATE2.Text = TIMS.Cdate3(STDATE2.Text)
        FTDATE1.Text = TIMS.Cdate3(FTDATE1.Text)
        FTDATE2.Text = TIMS.Cdate3(FTDATE2.Text)

        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,cc.ORGNAME" & vbCrLf
        sql &= " ,cc.CLASSCNAME2" & vbCrLf
        sql &= " ,cc.STDATE" & vbCrLf
        sql &= " ,cc.FTDATE" & vbCrLf
        sql &= " FROM dbo.VIEW2 cc" & vbCrLf
        sql &= " WHERE cc.Years=@Years AND cc.TPlanID=@TPlanID" & vbCrLf
        parms.Add("Years", sm.UserInfo.Years)
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sql &= " AND cc.DistID=@DistID" & vbCrLf
                sql &= " AND cc.PlanID=@PlanID" & vbCrLf
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("PlanID", sm.UserInfo.PlanID)
        End Select

        RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Select Case sm.UserInfo.LID
            Case 0
                If RIDValue.Value <> "A" AndAlso RIDValue.Value.Length = 1 Then
                    Dim s_DISTIDSCH As String = TIMS.Get_DistID_RID(RIDValue.Value, objConn)
                    sql &= " AND cc.DistID=@DISTIDSCH" & vbCrLf
                    parms.Add("DISTIDSCH", s_DISTIDSCH)
                Else
                    sql &= " AND cc.RID=@RID" & vbCrLf
                    parms.Add("RID", RIDValue.Value)
                End If
            Case 1
                If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
                    sql &= " AND cc.RID=@RID" & vbCrLf
                    parms.Add("RID", RIDValue.Value)
                End If
            Case 2
                sql &= " AND cc.RID=@RID" & vbCrLf
                parms.Add("RID", RIDValue.Value)
        End Select

        If OCIDValue1.Value <> "" Then
            sql &= " AND cc.OCID=@OCID" & vbCrLf
            parms.Add("OCID", OCIDValue1.Value)
        End If

        If STDATE1.Text <> "" Then
            'sql &= " AND cc.STDate >= " & TIMS.to_date(STDATE1.Text)
            sql &= " AND cc.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", STDATE1.Text)
        End If
        If STDATE2.Text <> "" Then
            'sql &= " AND cc.STDate <= " & TIMS.to_date(STDATE2.Text)
            sql &= " AND cc.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", STDATE2.Text)
        End If
        If FTDATE1.Text <> "" Then
            'sql &= " AND cc.FTDate >= " & TIMS.to_date(FTDATE1.Text)
            sql &= " AND cc.FTDate >= @FTDate1" & vbCrLf
            parms.Add("FTDate1", FTDATE1.Text)
        End If
        If FTDATE2.Text <> "" Then
            'sql &= " AND cc.FTDate <= " & TIMS.to_date(FTDATE2.Text)
            sql &= " AND cc.FTDate <= @FTDate2" & vbCrLf
            parms.Add("FTDate2", FTDATE2.Text)
        End If

        '28:產業人才投資方案
        'KindValue.Value=TIMS.GetTPlanName(sm.UserInfo.TPlanID)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sql &= " AND cc.OrgKind <> '10'" & vbCrLf
                Case "2"
                    '提升勞工自主學習計畫
                    sql &= " AND cc.OrgKind='10'" & vbCrLf
            End Select
        End If
        Return sql
    End Function
    '查詢原因-INQUIRY
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If OCID1.Text <> "" Then RstMemo &= String.Concat("&職類/班別=", OCID1.Text)
        If STDATE1.Text <> "" Then RstMemo &= String.Concat("&STDATE1=", STDATE1.Text)
        If STDATE2.Text <> "" Then RstMemo &= String.Concat("&STDATE2=", STDATE2.Text)
        If FTDATE1.Text <> "" Then RstMemo &= String.Concat("&FTDATE1=", FTDATE1.Text)
        If FTDATE2.Text <> "" Then RstMemo &= String.Concat("&FTDATE2=", FTDATE2.Text)
        If CONFIRDATE1.Text <> "" Then RstMemo &= String.Concat("&CONFIRDATE1=", CONFIRDATE1.Text)
        If CONFIRDATE2.Text <> "" Then RstMemo &= String.Concat("&CONFIRDATE2=", CONFIRDATE2.Text)
        Return RstMemo
    End Function
    '查詢 - SQL
    Sub Search1()
        'Dim flagS1 As Boolean=TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        CONFIRDATE1.Text = TIMS.Cdate3(CONFIRDATE1.Text) '(建檔日)
        CONFIRDATE2.Text = TIMS.Cdate3(CONFIRDATE2.Text) '(建檔日)

        Dim parms As New Hashtable()
        'parms.Clear()

        Dim strWC1 As String = S_WC1(parms)

        Dim sql As String = ""
        sql &= $" WITH WC1 AS ({strWC1})" & vbCrLf
        sql &= " SELECT cc.ocid" & vbCrLf
        sql &= " ,cc.orgname" & vbCrLf
        sql &= " ,cc.classcname2" & vbCrLf
        sql &= " ,CONVERT(VARCHAR ,cc.stdate ,111) stdate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR ,cc.ftdate ,111) ftdate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR ,cf.CONFIRDATE ,111) CONFIRDATE" & vbCrLf '(建檔日)
        sql &= " ,cf.CONFIRACCT" & vbCrLf
        sql &= " ,isnull(aa.name,'帳號有誤') CONFIRNAME" & vbCrLf
        sql &= " ,cf.CFGUID" & vbCrLf
        sql &= " ,cf.CFSEQNO" & vbCrLf
        sql &= " ,ISNULL(aa.LID,1) LID" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN CLASS_CONFIRM cf ON cf.OCID=cc.OCID" & vbCrLf
        sql &= " LEFT JOIN AUTH_ACCOUNT aa ON aa.account=cf.CONFIRACCT" & vbCrLf
        'If Not flagS1 Then sql &= " AND aa.LID<>2" & vbCrLf '排除委訓單位
        sql &= " WHERE aa.LID<>2" & vbCrLf '排除委訓單位
        If CONFIRDATE1.Text <> "" Then
            'sql &= " AND cf.CONFIRDATE >= " & TIMS.to_date(CONFIRDATE1.Text)
            sql &= " AND cf.CONFIRDATE >= @CONFIRDATE1" & vbCrLf
            parms.Add("CONFIRDATE1", CONFIRDATE1.Text)
        End If
        If CONFIRDATE2.Text <> "" Then
            'sql &= " AND cf.CONFIRDATE <= " & TIMS.to_date(CONFIRDATE2.Text)
            sql &= " AND cf.CONFIRDATE <= @CONFIRDATE2" & vbCrLf
            parms.Add("CONFIRDATE2", CONFIRDATE2.Text)
        End If
        Dim sCmd As New SqlCommand(sql, objConn)

        'Call TIMS.OpenDbConn(objConn)
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "ORGNAME,CLASSCNAME2,STDATE,FTDATE")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objConn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        'Call CloseDbConn(conn) 'If dt.Rows.Count > 0 Then Rst=Convert.ToString(dt.Rows(0)("?"))
        DataGridTable1.Visible = False
        Msg1.Text = "查無資料"
        If TIMS.dtNODATA(dt) Then Exit Sub

        DataGridTable1.Visible = True
        Msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call Search1()
    End Sub

    '新增
    Protected Sub BtnInsert1_Click(sender As Object, e As EventArgs) Handles BtnInsert1.Click
        DataGridTable1.Visible = False
        Msg1.Text = ""
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇 職類/班別!!")
            Exit Sub
        End If
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        If drC Is Nothing Then
            Common.MessageBox(Me, "請選擇 職類/班別!!")
            Exit Sub
        End If
        Dim url1 As String = ""
        url1 &= "SD_03_012_add.aspx?ID=" & TIMS.Get_MRqID(Me)
        url1 &= "&ACT=ADD"
        url1 &= "&OCID=" & OCIDValue1.Value
        TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

    Sub KEEPSEARCH()
        Session(vs_SearchStr1) = Nothing
        Dim xSearchStr As String = ""
        xSearchStr = "prg=SD_03_012"
        xSearchStr &= "&center=" & TIMS.ClearSQM(center.Text)
        xSearchStr &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        xSearchStr &= "&TMID1=" & TIMS.ClearSQM(TMID1.Text)
        xSearchStr &= "&OCID1=" & TIMS.ClearSQM(OCID1.Text)
        xSearchStr &= "&OCIDValue1=" & TIMS.ClearSQM(OCIDValue1.Value)
        xSearchStr &= "&TMIDValue1=" & TIMS.ClearSQM(TMIDValue1.Value)
        'xSearchStr += "&IDNO=" & TIMS.ChangeIDNO(IDNO.Text))
        xSearchStr &= "&STDATE1=" & TIMS.ClearSQM(STDATE1.Text)
        xSearchStr &= "&STDATE2=" & TIMS.ClearSQM(STDATE2.Text)
        xSearchStr &= "&FTDATE1=" & TIMS.ClearSQM(FTDATE1.Text)
        xSearchStr &= "&FTDATE2=" & TIMS.ClearSQM(FTDATE2.Text)
        xSearchStr &= "&CONFIRDATE1=" & TIMS.ClearSQM(CONFIRDATE1.Text)
        xSearchStr &= "&CONFIRDATE2=" & TIMS.ClearSQM(CONFIRDATE2.Text)
        xSearchStr &= "&PlanPoint=" & TIMS.ClearSQM(PlanPoint.SelectedValue)
        'xSearchStr += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        xSearchStr &= If(DataGridTable1.Visible, "&submit=1", "&submit=0")

        Session(vs_SearchStr1) = xSearchStr
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument

        Call KEEPSEARCH()
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim CFGUID As String = TIMS.GetMyValue(sCmdArg, "CFGUID")
        Dim CFSEQNO As String = TIMS.GetMyValue(sCmdArg, "CFSEQNO")
        If OCID = "" Then Exit Sub

        Dim url1 As String = ""
        url1 &= $"SD_03_012_add.aspx?ID={TIMS.Get_MRqID(Me)}"
        url1 &= $"&OCID={OCID}"
        url1 &= $"&CFGUID={CFGUID}"
        url1 &= $"&CFSEQNO={CFSEQNO}"
        url1 &= $"&INQUIRY_SCH={TIMS.GetListValue(ddl_INQUIRY_Sch)}"

        Select Case e.CommandName
            Case "BtnEDIT1" '報名名單
                url1 &= "&ACT=EDIT1"
                TIMS.Utl_Redirect(Me, objConn, url1)
            Case "BtnEDIT2" '報到名單
                url1 &= "&ACT=EDIT2"
                TIMS.Utl_Redirect(Me, objConn, url1)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim LabCONFIRDATE As Label = e.Item.FindControl("LabCONFIRDATE")
                Dim LabCONFIRNAME As Label = e.Item.FindControl("LabCONFIRNAME")
                Dim BtnEDIT1 As LinkButton = e.Item.FindControl("BtnEDIT1")
                Dim BtnEDIT2 As LinkButton = e.Item.FindControl("BtnEDIT2")
                LabCONFIRDATE.Text = Convert.ToString(drv("CONFIRDATE"))
                LabCONFIRNAME.Text = Convert.ToString(drv("CONFIRNAME"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "CFGUID", Convert.ToString(drv("CFGUID")))
                TIMS.SetMyValue(sCmdArg, "CFSEQNO", Convert.ToString(drv("CFSEQNO")))
                BtnEDIT1.CommandArgument = sCmdArg
                BtnEDIT2.CommandArgument = sCmdArg
        End Select
    End Sub
End Class
