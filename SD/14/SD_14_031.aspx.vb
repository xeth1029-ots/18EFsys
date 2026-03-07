Partial Class SD_14_031
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_031C" 'SD_14_031C/OJTSD14031C
    Const cst_printFN2 As String = "SD_14_031D" 'SD_14_031D/OJTSD14031D

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Create11()
            'center.Text = sm.UserInfo.OrgName
            'RIDValue.Value = sm.UserInfo.RID
            '判斷機構是否只有一個班級3
            'Call TIMS.GET_OnlyOne_OCID3(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
        End If

        'Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        'Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If

        'TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        'If HistoryTable.Rows.Count <> 0 Then
        '    OCID1.Attributes("onclick") = "showObj('HistoryList');"
        '    OCID1.Style("CURSOR") = "hand"
        'End If
    End Sub

    Private Sub Create11()
        msg.Text = "" '清空
        PanelSCH.Visible = True
        PanelVIEW.Visible = False
        q_APPLIEDDATE1.Text = TIMS.Cdate3(DateAdd(DateInterval.Month, -1, Now))
        q_APPLIEDDATE2.Text = TIMS.Cdate3(Now)
    End Sub

    '判斷機構是否只有一個班級
    'Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
    '    Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn) '判斷機構是否只有一個班級
    'End Sub

    '檢核日期有誤
    Function CHK_DATERANGE(ByRef V_QDATE1 As String, ByRef V_QDATE2 As String) As String
        Dim errMsg As String = ""
        If V_QDATE1 = "" AndAlso V_QDATE2 = "" Then Return ""
        '(_i 正確值比對)
        If V_QDATE1 <> "" AndAlso V_QDATE2 <> "" Then
            '有勾選且有填上架日期的資料，再進一步檢核設定結果不得超過報名日期
            If DateDiff(DateInterval.Minute, CDate(V_QDATE1), CDate(V_QDATE2)) < 0 Then
                errMsg &= "[申請日期區間起日]不能晚於[申請日期區間迄日]!" & vbCrLf
            End If
        ElseIf V_QDATE1 <> "" AndAlso V_QDATE2 = "" Then
            errMsg &= "[申請日期區間起日]有填寫，請填寫[申請日期區間迄日]!" & vbCrLf
        ElseIf V_QDATE1 = "" AndAlso V_QDATE2 <> "" Then
            errMsg &= "[申請日期區間迄日]有填寫，請填寫[申請日期區間起日]!" & vbCrLf
        End If
        Return errMsg
    End Function

    Sub SearchD1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim sErrMsg As String = ""
        q_IDNO.Text = TIMS.ClearSQM(q_IDNO.Text)
        q_CNAME.Text = TIMS.ClearSQM(q_CNAME.Text)
        q_APPLIEDDATE1.Text = TIMS.Cdate3(q_APPLIEDDATE1.Text)
        q_APPLIEDDATE2.Text = TIMS.Cdate3(q_APPLIEDDATE2.Text)
        If q_IDNO.Text = "" AndAlso q_CNAME.Text = "" AndAlso q_APPLIEDDATE1.Text = "" AndAlso q_APPLIEDDATE2.Text = "" Then
            sErrMsg = "請輸入查詢值(任一)"
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If
        If (q_APPLIEDDATE1.Text <> "" OrElse q_APPLIEDDATE2.Text <> "") Then
            Call TIMS.CheckDateErr(q_APPLIEDDATE1.Text, "申請日期區間起日", False, sErrMsg)
            Call TIMS.CheckDateErr(q_APPLIEDDATE2.Text, "申請日期區間迄日", False, sErrMsg)
            If sErrMsg <> "" Then
                Common.MessageBox(Me, sErrMsg)
                Exit Sub
            End If
            sErrMsg = CHK_DATERANGE(q_APPLIEDDATE1.Text, q_APPLIEDDATE2.Text)
            If sErrMsg <> "" Then
                Common.MessageBox(Me, sErrMsg)
                Exit Sub
            End If
        End If

        Dim pms1 As New Hashtable()
        Dim sSql As String = ""
        sSql &= " SELECT a.DCANO,a.DCASENO,a.IDNO,a.CNAME,a.PURID,a.UAGID" & vbCrLf
        sSql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,a.EMAIL,a.EMVCODE" & vbCrLf
        sSql &= " ,a.APPLNACCT,a.APPLNDATE ,format(a.APPLNDATE,'yyyy-MM-dd HH:mm') APPLNDATE_F" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " ,a.SENDACCT,a.SENDDATE" & vbCrLf
        sSql &= " FROM STUD_DIGICERTAPPLY a" & vbCrLf
        sSql &= " WHERE a.SENDDATE IS NOT NULL" & vbCrLf
        If q_APPLIEDDATE1.Text <> "" Then
            pms1.Add("APPLNDATE1", TIMS.Cdate2(q_APPLIEDDATE1.Text))
            sSql &= " AND convert(date,a.APPLNDATE) >= @APPLNDATE1" & vbCrLf
        End If
        If q_APPLIEDDATE2.Text <> "" Then
            pms1.Add("APPLNDATE2", TIMS.Cdate2(q_APPLIEDDATE2.Text))
            sSql &= " AND convert(date,a.APPLNDATE) <= @APPLNDATE2" & vbCrLf
        End If
        If q_IDNO.Text <> "" Then
            pms1.Add("IDNO", q_IDNO.Text)
            sSql &= " AND a.IDNO=@IDNO" & vbCrLf
        End If
        If q_CNAME.Text <> "" Then
            pms1.Add("CNAME", q_CNAME.Text)
            sSql &= " AND a.CNAME=@CNAME" & vbCrLf
        End If

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pms1)

        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        Call SearchD1()
    End Sub
    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If (e Is Nothing) Then Return
        If (e.CommandArgument Is Nothing OrElse e.CommandArgument = "") Then Return

        Dim sCmdArg As String = e.CommandArgument
        Dim sDCANO As String = TIMS.GetMyValue(sCmdArg, "DCANO")
        Dim sDCASENO As String = TIMS.GetMyValue(sCmdArg, "DCASENO")
        Dim sEMVCODE As String = TIMS.GetMyValue(sCmdArg, "EMVCODE")

        Dim flagNG As Boolean = If(sDCANO = "" OrElse sDCASENO = "" OrElse sEMVCODE = "", True, False)
        If flagNG Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Select Case e.CommandName
            Case "DATASHOW1"
                Call SHOW_DETAIL1(sDCANO, sDCASENO, sEMVCODE)
            Case "Print1"  '列印'"Print1"
                Dim prtstr As String = String.Format("&DCANO={0}&DCASENO={1}&EMVCODE={2}", sDCANO, sDCASENO, sEMVCODE)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtstr)
            Case "Print2"  '列印'"Print1"
                Dim prtstr As String = String.Format("&DCANO={0}&DCASENO={1}&EMVCODE={2}", sDCANO, sDCASENO, sEMVCODE)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, prtstr)
        End Select
    End Sub
    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hDCANO As HiddenField = e.Item.FindControl("hDCANO")
                Dim hDCASENO As HiddenField = e.Item.FindControl("hDCASENO")
                Dim hEMVCODE As HiddenField = e.Item.FindControl("hEMVCODE")
                Dim btnPrint1 As Button = e.Item.FindControl("btnPrint1")
                Dim btnPrint2 As Button = e.Item.FindControl("btnPrint2")
                Dim btnDATASHOW1 As Button = e.Item.FindControl("btnDATASHOW1")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                hDCANO.Value = Convert.ToString(drv("DCANO"))
                hDCASENO.Value = Convert.ToString(drv("DCASENO"))
                hEMVCODE.Value = Convert.ToString(drv("EMVCODE"))

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "DCANO", Convert.ToString(drv("DCANO")))
                TIMS.SetMyValue(sCmdArg, "DCASENO", Convert.ToString(drv("DCASENO")))
                TIMS.SetMyValue(sCmdArg, "EMVCODE", Convert.ToString(drv("EMVCODE")))

                btnPrint1.CommandArgument = sCmdArg
                btnPrint2.CommandArgument = sCmdArg
                btnDATASHOW1.CommandArgument = sCmdArg
        End Select
    End Sub

    Private Sub SHOW_DETAIL1(sDCANO As String, sDCASENO As String, sEMVCODE As String)
        Call clearDATA1()

        Dim pms As New Hashtable From {{"DCANO", sDCANO}, {"DCASENO", sDCASENO}, {"EMVCODE", sEMVCODE}}
        Dim sSql As String = ""
        sSql &= " SELECT a.DCANO ,a.DCASENO" & vbCrLf
        sSql &= " ,a.IDNO,a.CNAME" & vbCrLf
        sSql &= " ,a.PURID,(SELECT PURNAME FROM KEY_PURPOSE X WHERE X.PURID=A.PURID) PURNAME" & vbCrLf
        sSql &= " ,a.UAGID,(SELECT UAGNAME FROM KEY_USAGEUNIT X WHERE X.UAGID=A.UAGID) UAGNAME" & vbCrLf
        sSql &= " ,a.EMAIL,a.EMVCODE,a.GUID1" & vbCrLf
        sSql &= " ,a.APPLNACCT,format(a.APPLNDATE,'yyyy-MM-dd HH:mm') APPLNDATE_F,dbo.FN_CDATE(a.APPLNDATE) APPLNDATE_TW" & vbCrLf
        sSql &= " ,a.MODIFYACCT ,a.MODIFYDATE" & vbCrLf
        sSql &= " FROM STUD_DIGICERTAPPLY a" & vbCrLf
        sSql &= " WHERE a.DCANO=@DCANO AND a.DCASENO=@DCASENO AND a.EMVCODE=@EMVCODE " & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, pms)
        If dr1 Is Nothing Then Return

        labCNAME.Text = Convert.ToString(dr1("CNAME"))
        labIDNO.Text = Convert.ToString(dr1("IDNO"))
        labDCASENO.Text = Convert.ToString(dr1("DCASENO"))
        labAPPLNDATE.Text = Convert.ToString(dr1("APPLNDATE_F"))

        labPURPOSE.Text = Convert.ToString(dr1("PURNAME"))
        labUSAGEUNIT.Text = Convert.ToString(dr1("UAGNAME"))
        labAPPLNDATE_TW.Text = Convert.ToString(dr1("APPLNDATE_TW"))

        Hid_DCANO.Value = Convert.ToString(dr1("DCANO"))
        Hid_DCASENO.Value = Convert.ToString(dr1("DCASENO"))
        Hid_EMVCODE.Value = Convert.ToString(dr1("EMVCODE"))

        PanelSCH.Visible = False 'True
        PanelVIEW.Visible = True 'False

        Dim pms3 As New Hashtable From {{"DCANO", sDCANO}}
        Dim sSql3 As String = ""
        sSql3 &= " SELECT COUNT(1) CLSCNT FROM STUD_DIGICERTCLASS x WHERE x.DCANO=@DCANO" & vbCrLf
        Dim dr3 As DataRow = DbAccess.GetOneRow(sSql3, objconn, pms3)
        If dr3 IsNot Nothing Then
            labCLSCNT.Text = Convert.ToString(dr3("CLSCNT"))
        End If

        Dim pms2 As New Hashtable From {{"DCANO", sDCANO}}
        Dim sSql2 As String = ""
        sSql2 &= " SELECT COUNT(1) DLCNT,format(MAX(x.MODIFYDATE),'yyyy-MM-dd HH:mm') LASTDLTIME_F FROM STUD_DIGICERTAPPLYDL x WHERE x.DCANO=@DCANO" & vbCrLf
        Dim dr2 As DataRow = DbAccess.GetOneRow(sSql2, objconn, pms2)
        If dr2 IsNot Nothing Then
            labDLCNT.Text = Convert.ToString(dr2("DLCNT"))
            labLASTDLTIME.Text = Convert.ToString(dr2("LASTDLTIME_F"))
        End If
    End Sub

    Private Sub ClearDATA1()
        labCNAME.Text = ""
        labIDNO.Text = ""
        labDCASENO.Text = ""
        labAPPLNDATE.Text = ""

        labPURPOSE.Text = ""
        labUSAGEUNIT.Text = ""
        labAPPLNDATE_TW.Text = ""

        labCLSCNT.Text = ""
        labDLCNT.Text = "" 'Convert.ToString(dr2("DLCNT"))
        labLASTDLTIME.Text = "" 'Convert.ToString(dr2("LASTDLTIME_F"))

        Hid_DCANO.Value = ""
        Hid_DCASENO.Value = ""
        Hid_EMVCODE.Value = ""
    End Sub

    Protected Sub btnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        ClearDATA1()
        PanelSCH.Visible = True
        PanelVIEW.Visible = False
    End Sub
End Class

