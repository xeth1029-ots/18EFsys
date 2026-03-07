Partial Class SD_02_018
    Inherits AuthBasePage

    '職前計畫設計
    Const vs_SearchStr1 As String = "vsSearchStr1"
    'Dim au As New cAUTH
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objConn) '開啟連線
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1
        If Not IsPostBack Then Call Create1()
    End Sub

    Sub Create1()
        DataGridTable1.Visible = False
        Msg1.Text = ""

        'PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint)
        'Common.SetListItem(PlanPoint, "0")

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
            'Session(vs_SearchStr1) = Nothing
            MyValue = TIMS.GetMyValue(strSearchStr1, "prg")
            If MyValue = "SD_02_018" Then
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
                ANNMENTDATE1.Text = TIMS.GetMyValue(strSearchStr1, "ANNMENTDATE1")
                ANNMENTDATE2.Text = TIMS.GetMyValue(strSearchStr1, "ANNMENTDATE2")
                MyValue = TIMS.GetMyValue(strSearchStr1, "RBL_ROVED")
                If MyValue <> "" Then Common.SetListItem(RBL_ROVED, MyValue)
                MyValue = TIMS.GetMyValue(strSearchStr1, "RBL_ANNMENT")
                If MyValue <> "" Then Common.SetListItem(RBL_ANNMENT, MyValue)
                MyValue = TIMS.GetMyValue(strSearchStr1, "submit")
                If MyValue = "1" Then Call Search1()
            End If
            Session(vs_SearchStr1) = Nothing
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    ''' <summary>班級資料查詢</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function S_WC1(ByRef parms As Hashtable) As String
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        STDATE1.Text = TIMS.Cdate3(STDATE1.Text)
        STDATE2.Text = TIMS.Cdate3(STDATE2.Text)
        FTDATE1.Text = TIMS.Cdate3(FTDATE1.Text)
        FTDATE2.Text = TIMS.Cdate3(FTDATE2.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cc.ocid " & vbCrLf
        sql &= " ,cc.orgname " & vbCrLf
        sql &= " ,cc.classcname2 " & vbCrLf
        sql &= " ,cc.stdate ,cc.ftdate " & vbCrLf
        sql &= " FROM VIEW2 cc " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf

        sql &= " AND cc.YEARS = @YEARS" & vbCrLf
        sql &= " AND cc.TPlanID = @TPlanID " & vbCrLf
        parms.Add("YEARS", sm.UserInfo.Years.ToString())
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                sql &= " AND cc.DistID = @DistID " & vbCrLf
                sql &= " AND cc.PlanID = @PlanID " & vbCrLf
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("PlanID", sm.UserInfo.PlanID)
        End Select

        RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)

        Select Case sm.UserInfo.LID
            Case 0
                If RIDValue.Value <> "A" AndAlso RIDValue.Value.Length = 1 Then
                    Dim s_DISTIDSCH As String = TIMS.Get_DistID_RID(RIDValue.Value, objConn)
                    sql &= " AND cc.DistID = @DISTIDSCH " & vbCrLf
                    parms.Add("DISTIDSCH", s_DISTIDSCH)
                Else
                    sql &= " AND cc.RID = @RID " & vbCrLf
                    parms.Add("RID", RIDValue.Value)
                End If
            Case 1
                If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
                    sql &= " AND cc.RID = @RID " & vbCrLf
                    parms.Add("RID", RIDValue.Value)
                End If
            Case 2
                sql &= " AND cc.RID = @RID " & vbCrLf
                parms.Add("RID", RIDValue.Value)
        End Select

        If OCIDValue1.Value <> "" Then
            sql &= " AND cc.OCID = @OCID " & vbCrLf
            parms.Add("OCID", OCIDValue1.Value)
        End If
        If STDATE1.Text <> "" Then
            sql &= " AND cc.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", STDATE1.Text)
        End If
        If STDATE2.Text <> "" Then
            sql &= " AND cc.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", STDATE2.Text)
        End If
        If FTDATE1.Text <> "" Then
            sql &= " AND cc.FTDate >= @FTDate1 " & vbCrLf
            parms.Add("FTDate1", FTDATE1.Text)
        End If
        If FTDATE2.Text <> "" Then
            sql &= " AND cc.FTDate <= @FTDate2 " & vbCrLf
            parms.Add("FTDate2", FTDATE2.Text)
        End If

        Return sql
    End Function

    '查詢 - SQL
    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '審核日期、審核者、公告日期、公告者
        ANNMENTDATE1.Text = TIMS.ClearSQM(ANNMENTDATE1.Text) '公告日1
        ANNMENTDATE2.Text = TIMS.ClearSQM(ANNMENTDATE2.Text) '公告日2

        Dim parms As Hashtable = New Hashtable()
        parms.Clear()

        Dim strWC1 As String = S_WC1(parms)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & strWC1 & " )" & vbCrLf
        sql &= " SELECT cc.ocid " & vbCrLf
        sql &= "  ,cc.orgname " & vbCrLf
        sql &= "  ,cc.classcname2 " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, cc.stdate, 111) stdate " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, cc.ftdate, 111) ftdate " & vbCrLf
        sql &= "  ,CONVERT(VARCHAR, cf.ROVEDDATE, 111) ROVEDDATE " & vbCrLf '審核日
        sql &= "  ,CONVERT(VARCHAR, cf.ANNMENTDATE, 111) ANNMENTDATE " & vbCrLf '公告日
        sql &= "  ,cf.ROVEDACCT " & vbCrLf '審核
        sql &= "  ,aa.name ROVEDNAME " & vbCrLf '審核
        sql &= "  ,cf.ANNMENTACCT " & vbCrLf '公告
        sql &= "  ,aa2.name ANNMENTNAME " & vbCrLf '公告
        sql &= "  ,cf.CFGUID " & vbCrLf
        sql &= "  ,cf.CFSEQNO " & vbCrLf
        sql &= " FROM WC1 cc " & vbCrLf
        sql &= " JOIN CLASS_CONFIRM cf ON cf.ocid = cc.ocid " & vbCrLf
        sql &= " LEFT JOIN AUTH_ACCOUNT aa ON aa.account = cf.ROVEDACCT " & vbCrLf '審核
        sql &= " LEFT JOIN AUTH_ACCOUNT aa2 ON aa2.account = cf.ANNMENTACCT " & vbCrLf '公告
        sql &= " WHERE 1=1 " & vbCrLf

        If ANNMENTDATE1.Text <> "" Then
            sql &= " AND cf.ANNMENTDATE >= @ANNMENTDATE1 " & vbCrLf
            parms.Add("ANNMENTDATE1", ANNMENTDATE1.Text)
        End If
        If ANNMENTDATE2.Text <> "" Then
            sql &= " AND cf.ANNMENTDATE <= @ANNMENTDATE2 " & vbCrLf
            parms.Add("ANNMENTDATE2", ANNMENTDATE2.Text)
        End If

        Select Case RBL_ROVED.SelectedValue '審核
            Case "Y"
                sql &= " AND cf.ROVEDDATE IS NOT NULL " & vbCrLf
            Case "N"
                sql &= " AND cf.ROVEDDATE IS NULL " & vbCrLf
        End Select
        Select Case RBL_ANNMENT.SelectedValue '公告
            Case "Y"
                sql &= " AND cf.ANNMENTDATE IS NOT NULL " & vbCrLf
            Case "N"
                sql &= " AND cf.ANNMENTDATE IS NULL " & vbCrLf
        End Select

        'ROVEDACCT  VARCHAR2 (15 CHAR)  ,--審核
        'ROVEDDATE DATE ,
        'ANNMENTACCT  VARCHAR2 (15 CHAR)  ,--公告
        'ANNMENTDATE  DATE  
        Dim sCmd As New SqlCommand(sql, objConn)

        DbAccess.HashParmsChange(sCmd, parms)

        Call TIMS.OpenDbConn(objConn)

        Dim dt As New DataTable

        dt.Load(sCmd.ExecuteReader())

        'Call CloseDbConn(conn)
        'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))
        DataGridTable1.Visible = False
        Msg1.Text = "查無資料"
        If dt.Rows.Count = 0 Then Exit Sub
        Msg1.Text = ""
        DataGridTable1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        Call Search1()
    End Sub

    '保留搜尋值
    Sub KEEPSEARCH()
        Session(vs_SearchStr1) = Nothing
        Dim xSearchStr As String = ""
        xSearchStr = "prg=SD_02_018"
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
        xSearchStr &= "&ANNMENTDATE1=" & TIMS.ClearSQM(ANNMENTDATE1.Text)
        xSearchStr &= "&ANNMENTDATE2=" & TIMS.ClearSQM(ANNMENTDATE2.Text)
        xSearchStr &= "&RBL_ROVED=" & TIMS.ClearSQM(RBL_ROVED.SelectedValue)
        xSearchStr &= "&RBL_ANNMENT=" & TIMS.ClearSQM(RBL_ANNMENT.SelectedValue)
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
        url1 = ""
        url1 &= "SD_02_018_add.aspx?ID=" & TIMS.Get_MRqID(Me)
        url1 &= "&OCID=" & OCID
        url1 &= "&CFGUID=" & CFGUID
        url1 &= "&CFSEQNO=" & CFSEQNO
        Select Case e.CommandName
            Case "BtnEDIT1" '檢視
                url1 &= "&ACT=EDIT1"
                TIMS.Utl_Redirect(Me, objConn, url1)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim LabROVEDDATE As Label = e.Item.FindControl("LabROVEDDATE")
                Dim LabROVEDNAME As Label = e.Item.FindControl("LabROVEDNAME")
                Dim LabANNMENTDATE As Label = e.Item.FindControl("LabANNMENTDATE")
                Dim LabANNMENTNAME As Label = e.Item.FindControl("LabANNMENTNAME")
                Dim BtnEDIT1 As LinkButton = e.Item.FindControl("BtnEDIT1")

                LabROVEDDATE.Text = Convert.ToString(drv("ROVEDDATE"))
                LabROVEDNAME.Text = Convert.ToString(drv("ROVEDNAME"))
                LabANNMENTDATE.Text = Convert.ToString(drv("ANNMENTDATE"))
                LabANNMENTNAME.Text = Convert.ToString(drv("ANNMENTNAME"))

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "CFGUID", Convert.ToString(drv("CFGUID")))
                TIMS.SetMyValue(sCmdArg, "CFSEQNO", Convert.ToString(drv("CFSEQNO")))
                BtnEDIT1.CommandArgument = sCmdArg
        End Select
    End Sub
End Class