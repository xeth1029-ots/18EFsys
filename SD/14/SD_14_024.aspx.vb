Partial Class SD_14_024
    Inherits AuthBasePage

    '//已轉班,'G BussinessTrain: SD_14_021 '產業人才投資計畫 'W BussinessTrain: SD_14_021_B 
    '提升勞工自主學習計畫,'(vp.Years-1911),'助教到底要不要印？ 不要(N),',ISNULL(dbo.FN_GET_CLASS_TEACHER(cc.ocid), '') TEACHER
    'SD_14_024?.jrxml / SD024ON?.jrxml,'SD_14_024G SD_14_024W 招訓簡章,'SD024ONG.jrxml/SD024ONW.jrxml (前台 報名網)
    Const cst_printFN_G1 As String = "SD_14_024G" 'SD024ONG/SD_14_024G
    Const cst_printFN_W1 As String = "SD_14_024W" 'SD024ONW/SD_14_024W
    Dim prtFilename As String = "" '列印表件名稱

    'Dim iPYNum17 As Integer=1 'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線
        PageControler1.PageDataGrid = DataGrid1
        'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
        hidYears.Value = sm.UserInfo.Years - 1911 '設定登入民國年

        If Not IsPostBack Then
            msg.Text = "" '每次 清空
            DataGridTable.Visible = False '預設 隱藏
            hidOCIDValue.Value = ""
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "0")

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2

            '列印 'btnPrint1.Attributes("onclick")="return CheckPrint();"
            Button4.Attributes("onclick") = "ClearData();"
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

    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim sRelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = "
SELECT cc.OCID,concat(pp.PlanID,'_',pp.ComIDNO,'_',pp.SeqNo) PCSValue
,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName
,CONVERT(VARCHAR,cc.STDate,111) STDate
,CONVERT(VARCHAR,cc.FTDate,111) FTDate
,v1.OrgName,v1.OrgKindGW
,concat(dbo.FN_CTIME(PP.MODIFYDATE),dbo.FN_CTIME(CC.MODIFYDATE)) PMD
FROM PLAN_PLANINFO pp
JOIN CLASS_CLASSINFO cc ON cc.PlanID=pp.PlanID AND cc.comidno=pp.comidno AND cc.seqno=pp.seqno
JOIN ID_Plan ip ON ip.PlanID=pp.PlanID
JOIN VIEW_RIDNAME v1 ON v1.RID=pp.RID
WHERE pp.IsApprPaper='Y' AND pp.TransFlag='Y'
"
        '限制為只有正式儲存之班級'已轉班
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND pp.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If sRelShip <> "" Then sql &= " AND v1.RelShip LIKE '" & sRelShip & "%'" & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf

        '28:產業人才投資方案
        'hidorgid.Value=""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case PlanPoint.SelectedValue
                Case "1"
                    '產業人才投資計畫
                    sql &= " AND v1.OrgKind<>'10'" & vbCrLf
                    'hidorgid.Value="G"
                Case "2"
                    '提升勞工自主學習計畫
                    sql &= " AND v1.OrgKind='10'" & vbCrLf
                    'hidorgid.Value="W"
            End Select
        End If

        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGrid1.Visible = True
        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        '1:已轉班
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim flagNG As Boolean = False
        Dim sCmdArg As String = e.CommandArgument
        Dim OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim PCSValue As String = TIMS.GetMyValue(sCmdArg, "PCSValue")
        Dim OrgKindGW As String = TIMS.GetMyValue(sCmdArg, "OrgKindGW")

        If e.CommandArgument = "" Then flagNG = True
        If OCID = "" Then flagNG = True
        If PCSValue = "" Then flagNG = True
        If flagNG Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Select Case OrgKindGW'hidorgid.Value
            Case "G"
                prtFilename = cst_printFN_G1
            Case "W"
                prtFilename = cst_printFN_W1
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        Select Case e.CommandName
            Case "Print1"
                Dim prtstr As String = ""
                prtstr &= "&TPlanID=" & sm.UserInfo.TPlanID
                prtstr &= "&OCID=" & OCID 'hidOCIDValue.Value
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, prtstr)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass="SD_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType=ListItemType.Item Then e.Item.CssClass="SD_TD2"
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HiddenField = e.Item.FindControl("OCID")
                Dim PCSValue As HiddenField = e.Item.FindControl("PCSValue")
                OCID.Value = Convert.ToString(drv("OCID"))
                PCSValue.Value = Convert.ToString(drv("PCSValue"))

                Dim btnPrint1 As Button = e.Item.FindControl("btnPrint1")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "PCSValue", Convert.ToString(drv("PCSValue")))
                TIMS.SetMyValue(sCmdArg, "OrgKindGW", Convert.ToString(drv("OrgKindGW")))
                btnPrint1.CommandArgument = sCmdArg
        End Select
    End Sub

End Class