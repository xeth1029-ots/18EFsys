Partial Class SD_14_033
    Inherits AuthBasePage

    Const cst_printFN_1 As String = "SD_14_033"
    Const cst_printFN_2 As String = "SD_14_033_S"
    Const cst_vsSearchStr1 As String = "vsSearchStr1"

    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn) '開啟連線
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

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        Msg1.Text = ""
        'PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objConn)
        'Common.SetListItem(PlanPoint, "1")
        PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objConn)
        Common.SetListItem(PlanPoint, "0")

        SCH_IDNO.Text = ""
        SCH_CNAME.Text = ""
        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objConn, V_INQUIRY)

        If sm.UserInfo.LID <> "2" Then
            '若只有管理一個班級，自動協助帶出班級
            Call TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objConn)
        Else
            'Button12_Click(sender, e)
            center.Enabled = False
            Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objConn)
        End If

        USEKEEPSEARCH()

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(ddl_INQUIRY_Sch, "01")
        End If
    End Sub

    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Protected Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objConn)
        '判斷機構是否只有一個班級 '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = $"{dr("TRAINNAME")}"
        OCID1.Text = $"{dr("CLASSNAME")}"
        TMIDValue1.Value = $"{dr("TRAINID")}"
        OCIDValue1.Value = $"{dr("OCID")}"
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        SCH_IDNO.Text = TIMS.ClearSQM(SCH_IDNO.Text)
        SCH_CNAME.Text = TIMS.ClearSQM(SCH_CNAME.Text)
        'table_sch,center,Button8,Button4,Button3,RIDValue,HistoryList2,HistoryRID,TMID1,OCID1,OCIDValue1,TMIDValue1,HistoryList,HistoryTable
        ',TRPlanPoint28,PlanPoint,SCH_IDNO,SCH_CNAME,tr_ddl_INQUIRY_S,ddl_INQUIRY_Sch,
        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)
        If SCH_IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", SCH_IDNO.Text)
        If SCH_CNAME.Text <> "" Then RstMemo &= String.Concat("&學員姓名=", SCH_CNAME.Text)
        Return RstMemo
    End Function

    '查詢 - SQL
    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        SCH_IDNO.Text = TIMS.ClearSQM(SCH_IDNO.Text)
        SCH_CNAME.Text = TIMS.ClearSQM(SCH_CNAME.Text)

        Dim ERRMSG As String = ""
        If RIDValue.Value = "" Then ERRMSG &= "請選擇訓練機構!" & vbCrLf
        If OCIDValue1.Value = "" Then ERRMSG &= "請選擇職類/班別!" & vbCrLf
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Return
        End If
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Dim parms As New Hashtable() From {{"YEARS", sm.UserInfo.Years}, {"OCID", TIMS.CINT1(OCIDValue1.Value)}, {"RID", RIDValue.Value}}

        Dim sSql As String = ""
        sSql &= " WITH WS1 AS ( SELECT cs.OCID,cs.TPLANID,cs.RID,cs.SOCID,cs.BIRTHDAY,cs.SID,cs.ORGKINDGW,cs.YEARS,cs.CLASSCNAME2 ,cs.STUDID2,cs.NAME,cs.IDNO,cs.SETID" & vbCrLf
        sSql &= " FROM V_STUDENTINFO cs" & vbCrLf
        sSql &= " WHERE cs.YEARS=@YEARS AND cs.OCID=@OCID AND cs.RID=@RID" & vbCrLf

        '28:產業人才投資方案
        HidORGKINDGW.Value = ""
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim V_PlanPoint As String = TIMS.GetListValue(PlanPoint)
            Select Case V_PlanPoint
                Case "1"
                    '產業人才投資計畫
                    HidORGKINDGW.Value = "G"
                    parms.Add("ORGKINDGW", HidORGKINDGW.Value)
                    sSql &= " AND cs.ORGKINDGW=@ORGKINDGW" & vbCrLf
                Case "2"
                    '提升勞工自主學習計畫
                    HidORGKINDGW.Value = "W"
                    parms.Add("ORGKINDGW", HidORGKINDGW.Value)
                    sSql &= " AND cs.ORGKINDGW=@ORGKINDGW" & vbCrLf
            End Select
        End If
        If SCH_IDNO.Text <> "" Then
            parms.Add("IDNO", SCH_IDNO.Text)
            sSql &= " AND cs.IDNO=@IDNO" & vbCrLf
        End If
        If SCH_CNAME.Text <> "" Then
            parms.Add("NAME", SCH_CNAME.Text)
            sSql &= " AND cs.NAME like '%'+@NAME+'%'" & vbCrLf
        End If

        sSql &= " )" & vbCrLf
        sSql &= " ,WS2 AS ( SELECT a.SOCID,b.ETYPE,b.EMID1,b.IDNO,b.CREATEDATE,b.FILENAME1,b.FILENAME1W,b.SRCFILENAME1,b.FILEPATH1" & vbCrLf
        sSql &= " ,b.FILENAME2,b.FILENAME2W,b.SRCFILENAME2,b.FILEPATH2,b.ISUSE,b.ISDEL,b.MODIFYACCT,b.MODIFYDATE,b.CATEGORY1,b.ACTION1" & vbCrLf
        sSql &= " ,b.E2FILENAME1,b.E2FILENAME2,b.E1FILENAME1" & vbCrLf
        sSql &= " FROM WS1 a JOIN V_EIMG12 b on b.IDNO=a.IDNO" & vbCrLf
        sSql &= " WHERE ISUSE='Y' AND ISDEL IS NULL )" & vbCrLf

        sSql &= " SELECT cs.OCID,cs.TPLANID,cs.RID,cs.ORGKINDGW,cs.YEARS,cs.CLASSCNAME2" & vbCrLf
        sSql &= " ,cs.SOCID,cs.STUDID2,cs.NAME CNAME,cs.IDNO,dbo.FN_GET_MASK1(cs.IDNO) IDNO_MK" & vbCrLf
        sSql &= " ,cs.SID,dbo.FN_GET_MASK6(cs.SID) SIDMK6,cs.BIRTHDAY,FORMAT(cs.BIRTHDAY,'yMd') BDAY,cs.SETID" & vbCrLf
        sSql &= " ,s22.E2FILENAME1,s22.E2FILENAME2,s21.E1FILENAME1" & vbCrLf
        sSql &= " FROM WS1 cs" & vbCrLf
        sSql &= " LEFT JOIN WS2 s22 on s22.SOCID=cs.SOCID AND s22.IDNO=cs.IDNO AND s22.ETYPE=2 /*身分證*/" & vbCrLf
        sSql &= " LEFT JOIN WS2 s21 on s21.SOCID=cs.SOCID AND s21.IDNO=cs.IDNO AND s21.ETYPE=1 /*存摺*/" & vbCrLf

        sSql &= " ORDER BY cs.STUDID2" & vbCrLf

        DataGridTable1.Visible = False
        Msg1.Text = "查無資料"

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"{vbCrLf}{TIMS.GetMyValue5(parms)}{vbCrLf}--SD_14_033, sSql:{vbCrLf}{sSql}")
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objConn, parms)

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDID2,IDNO,CNAME")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objConn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then Return 'Exit Sub

        DataGridTable1.Visible = True
        Msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Protected Sub BtnSearch1_Click(sender As Object, e As EventArgs) Handles BtnSearch1.Click
        Call Search1()
    End Sub
    Sub USEKEEPSEARCH()
        If Session(cst_vsSearchStr1) Is Nothing Then Return

        Dim strSearchStr1 As String = Session(cst_vsSearchStr1)
        Session(cst_vsSearchStr1) = Nothing

        Dim MyValue As String = TIMS.GetMyValue(strSearchStr1, "prg")
        If MyValue = "SD_14_033" Then
            center.Text = TIMS.GetMyValue(strSearchStr1, "center")
            RIDValue.Value = TIMS.GetMyValue(strSearchStr1, "RIDValue")
            TMID1.Text = TIMS.GetMyValue(strSearchStr1, "TMID1")
            OCID1.Text = TIMS.GetMyValue(strSearchStr1, "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(strSearchStr1, "OCIDValue1")
            SCH_IDNO.Text = TIMS.GetMyValue(strSearchStr1, "SCH_IDNO")
            SCH_CNAME.Text = TIMS.GetMyValue(strSearchStr1, "SCH_CNAME")
            MyValue = TIMS.GetMyValue(strSearchStr1, "submit")
            If MyValue = "1" Then Call Search1()
        End If
    End Sub
    Sub KEEPSEARCH()
        Session(cst_vsSearchStr1) = Nothing
        Dim xSearchStr As String = "prg=SD_14_033"
        xSearchStr &= $"&center={TIMS.ClearSQM(center.Text)}"
        xSearchStr &= $"&RIDValue={TIMS.ClearSQM(RIDValue.Value)}"
        xSearchStr &= $"&TMID1={TIMS.ClearSQM(TMID1.Text)}"
        xSearchStr &= $"&OCID1={TIMS.ClearSQM(OCID1.Text)}"
        xSearchStr &= $"&OCIDValue1={TIMS.ClearSQM(OCIDValue1.Value)}"
        xSearchStr &= $"&TMIDValue1={TIMS.ClearSQM(TMIDValue1.Value)}"
        xSearchStr &= $"&SCH_IDNO={TIMS.ClearSQM(SCH_IDNO.Text)}"
        xSearchStr &= $"&SCH_CNAME={TIMS.ClearSQM(SCH_CNAME.Text)}"
        xSearchStr &= If(DataGridTable1.Visible, "&submit=1", "&submit=0")
        Session(cst_vsSearchStr1) = xSearchStr
    End Sub


    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument

        Call KEEPSEARCH()

        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vTPlanID As String = TIMS.GetMyValue(sCmdArg, "TPlanID")
        Dim vRID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim vSTUDID2 As String = TIMS.GetMyValue(sCmdArg, "STUDID2")
        Dim vSOCID As String = TIMS.GetMyValue(sCmdArg, "SOCID")
        Dim vSIDMK6 As String = TIMS.GetMyValue(sCmdArg, "SIDMK6")
        Dim vBDAY As String = TIMS.GetMyValue(sCmdArg, "BDAY")
        Dim vSETID As String = TIMS.GetMyValue(sCmdArg, "SETID")
        Dim sE2FILENAME1 As String = TIMS.GetMyValue(sCmdArg, "E2FILENAME1")
        Dim sE2FILENAME2 As String = TIMS.GetMyValue(sCmdArg, "E2FILENAME2")
        Dim sE1FILENAME1 As String = TIMS.GetMyValue(sCmdArg, "E1FILENAME1")
        'Dim BtnDOWNL1 As LinkButton = e.Item.FindControl("BtnDOWNL1") '身分證正面,BtnDOWNL1,
        'Dim BtnDOWNL2 As LinkButton = e.Item.FindControl("BtnDOWNL2") '身分證反面,BtnDOWNL2
        'Dim BtnDOWNL3 As LinkButton = e.Item.FindControl("BtnDOWNL3") '存摺影本,BtnDOWNL3
        Dim tmpFile As String = ""
        Select Case e.CommandName
            Case "Print1"
                Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objConn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
                'Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
                Dim vTSTPRINT As String = If($"{TIMS.Utl_GetConfigSet("printtest")}" = "Y", "2", "1") '正式區1／'測試區2
                Dim vMSD As String = $"{drCC("MSD")}"
                Dim myValue1 As String = ""
                TIMS.SetMyValue(myValue1, "TPlanID", vTPlanID)
                TIMS.SetMyValue(myValue1, "RID", vRID)
                TIMS.SetMyValue(myValue1, "OCID", OCIDValue1.Value)
                TIMS.SetMyValue(myValue1, "MSD", vMSD)
                TIMS.SetMyValue(myValue1, "SOCID", vSOCID)
                TIMS.SetMyValue(myValue1, "STUDID2", vSTUDID2)
                TIMS.SetMyValue(myValue1, "SIDMK6", vSIDMK6)
                TIMS.SetMyValue(myValue1, "BDAY", vBDAY)
                TIMS.SetMyValue(myValue1, "SETID", vSETID)
                TIMS.SetMyValue(myValue1, "TSTPRINT", vTSTPRINT)
                TIMS.CloseDbConn(objConn) : ReportQuery.PrintReport(Me, cst_printFN_2, myValue1)

            Case "BtnDOWNL1"
                If sE2FILENAME1 = "" Then
                    Common.MessageBox(Me, $"檔案不存在 .")
                    Return
                End If
                Dim fgOK As Boolean = TIMS.Utl_DownloadFile(Me, sE2FILENAME1, tmpFile)
                If Not fgOK Then
                    Common.MessageBox(Me, $"檔案下載有誤 . {tmpFile}")
                    Return
                End If
            Case "BtnDOWNL2"
                If sE2FILENAME2 = "" Then
                    Common.MessageBox(Me, $"檔案不存在 .")
                    Return
                End If
                Dim fgOK As Boolean = TIMS.Utl_DownloadFile(Me, sE2FILENAME2, tmpFile)
                If Not fgOK Then
                    Common.MessageBox(Me, $"檔案下載有誤 . {tmpFile}")
                    Return
                End If
            Case "BtnDOWNL3"
                If sE1FILENAME1 = "" Then
                    Common.MessageBox(Me, $"檔案不存在 .")
                    Return
                End If
                Dim fgOK As Boolean = TIMS.Utl_DownloadFile(Me, sE1FILENAME1, tmpFile)
                If Not fgOK Then
                    Common.MessageBox(Me, $"檔案下載有誤 . {tmpFile}")
                    Return
                End If
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                'DataGrid1,序號>,STUDID2,學號>,CNAME,姓名>,IDNO_MK,身分證號碼>,BtnPrint1,列印,Print1></asp:LinkButton>
                'BtnDOWNL1, 身分證正面, BtnDOWNL1, BtnDOWNL2 -, 身分證反面, BtnDOWNL2, BtnDOWNL3, 存摺影本, BtnDOWNL3,
                Dim BtnPrint1 As LinkButton = e.Item.FindControl("BtnPrint1") '列印,Print1
                'Dim BtnDOWNL1 As LinkButton = e.Item.FindControl("BtnDOWNL1") '身分證正面,BtnDOWNL1,
                'Dim BtnDOWNL2 As LinkButton = e.Item.FindControl("BtnDOWNL2") '身分證反面,BtnDOWNL2
                Dim BtnDOWNL3 As LinkButton = e.Item.FindControl("BtnDOWNL3") '存摺影本,BtnDOWNL3
                Dim iDGSeqNo As Integer = TIMS.Get_DGSeqNo(sender, e)
                e.Item.Cells(0).Text = iDGSeqNo
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", $"{drv("ORGKINDGW")}")
                TIMS.SetMyValue(sCmdArg, "TPlanID", $"{drv("TPlanID")}")
                TIMS.SetMyValue(sCmdArg, "RID", $"{drv("RID")}")
                TIMS.SetMyValue(sCmdArg, "OCID", $"{drv("OCID")}")

                TIMS.SetMyValue(sCmdArg, "STUDID2", $"{drv("STUDID2")}")
                TIMS.SetMyValue(sCmdArg, "SOCID", $"{drv("SOCID")}")
                TIMS.SetMyValue(sCmdArg, "SIDMK6", $"{drv("SIDMK6")}")
                TIMS.SetMyValue(sCmdArg, "BDAY", $"{drv("BDAY")}")
                TIMS.SetMyValue(sCmdArg, "SETID", $"{drv("SETID")}")
                'TIMS.SetMyValue(sCmdArg, "YEARS_ROC", $"{drv("YEARS_ROC")}")
                TIMS.SetMyValue(sCmdArg, "E2FILENAME1", $"{drv("E2FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "E2FILENAME2", $"{drv("E2FILENAME2")}")
                TIMS.SetMyValue(sCmdArg, "E1FILENAME1", $"{drv("E1FILENAME1")}")

                BtnPrint1.Visible = $"{drv("E2FILENAME1")}{drv("E2FILENAME2")}{drv("E1FILENAME1")}" <> ""
                BtnPrint1.CommandArgument = sCmdArg
                '暫不提供下載 身分證與存摺
                'BtnDOWNL1.Visible = ($"{drv("E2FILENAME1")}" <> "") '身分證正面,BtnDOWNL1,
                'BtnDOWNL2.Visible = ($"{drv("E2FILENAME2")}" <> "") '身分證反面,BtnDOWNL2
                BtnDOWNL3.Visible = ($"{drv("E1FILENAME1")}" <> "") '存摺影本,BtnDOWNL3
                'If BtnDOWNL1.Visible Then BtnDOWNL1.CommandArgument = sCmdArg
                'If BtnDOWNL2.Visible Then BtnDOWNL2.CommandArgument = sCmdArg
                If BtnDOWNL3.Visible Then BtnDOWNL3.CommandArgument = sCmdArg
        End Select
    End Sub

    ''' <summary>列印空白表單</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnPrint2_Click(sender As Object, e As EventArgs) Handles BtnPrint2.Click

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        SCH_IDNO.Text = TIMS.ClearSQM(SCH_IDNO.Text)
        SCH_CNAME.Text = TIMS.ClearSQM(SCH_CNAME.Text)

        Dim ERRMSG As String = ""
        If RIDValue.Value = "" Then ERRMSG &= "請選擇訓練機構!" & vbCrLf
        If OCIDValue1.Value = "" Then ERRMSG &= "請選擇職類/班別!" & vbCrLf
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Return
        End If

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim vMSD As String = $"{drCC("MSD")}"
        Dim myValue1 As String = ""
        TIMS.SetMyValue(myValue1, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(myValue1, "RID", RIDValue.Value)
        TIMS.SetMyValue(myValue1, "OCID", OCIDValue1.Value)
        TIMS.SetMyValue(myValue1, "MSD", vMSD)
        TIMS.CloseDbConn(objConn) : ReportQuery.PrintReport(Me, cst_printFN_1, myValue1)

    End Sub

End Class
