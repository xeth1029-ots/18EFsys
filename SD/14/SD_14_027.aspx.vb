Public Class SD_14_027
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-21100502：產投 - 新增【學員簽到(退)及教學日誌 】功能
    Const cst_printFN1 As String = "SD_14_027" 'SD_14_027_subreport1 --產投
    Const cst_printFN2 As String = "SD_14_027B" 'SD_14_027B_subreport1 --充電起飛
    Const cst_dg1eCM_Print1 As String = "Print1"

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = "" '清空
            DataGridTable.Visible = False '預設 隱藏
            'STRAINDATE.Text = TIMS.GetSysDate(objconn)

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            'PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn) 'Common.SetListItem(PlanPoint, "0")

            Dim s_javascript_btn2 As String = ""
            Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
            s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');", s_LevOrg)
            Button2.Attributes("onclick") = s_javascript_btn2

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

        Dim sRelShip As String = ""
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        sRelShip = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,concat(pp.PlanID,'_',pp.ComIDNO,'_',pp.SeqNo) PCSValue" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.ClassCName,cc.CyclType) ClassCName" & vbCrLf
        sql &= " ,format(cc.STDate,'yyyy/MM/dd') STDate" & vbCrLf
        sql &= " ,format(cc.FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " ,v1.OrgName" & vbCrLf
        sql &= " ,concat(dbo.FN_CTIME(PP.MODIFYDATE), dbo.FN_CTIME(CC.MODIFYDATE)) PMD" & vbCrLf
        sql &= " ,v1.OrgKindGW" & vbCrLf
        sql &= " ,format(v2.STRAINDATE,'yyyy/MM/dd') STRAINDATE" & vbCrLf

        sql &= " ,(SELECT STUFF((SELECT '；'+concat('(',dc.pname,')',dc.PCONT) FROM V_TRAINDESC dc WHERE dc.PLANID=pp.PLANID and dc.COMIDNO=pp.COMIDNO and dc.SEQNO=pp.SEQNO and dc.STRAINDATE=v2.STRAINDATE AND dc.TPERIOD28_1='Y' ORDER BY dc.pname FOR XML PATH('')),1,1,'')) TPERIOD28_1" & vbCrLf
        sql &= " ,(SELECT STUFF((SELECT '；'+concat('(',dc.pname,')',dc.PCONT) FROM V_TRAINDESC dc WHERE dc.PLANID=pp.PLANID and dc.COMIDNO=pp.COMIDNO and dc.SEQNO=pp.SEQNO and dc.STRAINDATE=v2.STRAINDATE AND dc.TPERIOD28_2='Y' ORDER BY dc.pname FOR XML PATH('')),1,1,'')) TPERIOD28_2" & vbCrLf
        sql &= " ,(SELECT STUFF((SELECT '；'+concat('(',dc.pname,')',dc.PCONT) FROM V_TRAINDESC dc WHERE dc.PLANID=pp.PLANID and dc.COMIDNO=pp.COMIDNO and dc.SEQNO=pp.SEQNO and dc.STRAINDATE=v2.STRAINDATE AND dc.TPERIOD28_3='Y' ORDER BY dc.pname FOR XML PATH('')),1,1,'')) TPERIOD28_3" & vbCrLf
        sql &= " ,(SELECT STUFF((SELECT '；'+concat('(',dc.pname,')',dc.PCONT) FROM V_TRAINDESC dc WHERE dc.PLANID=pp.PLANID and dc.COMIDNO=pp.COMIDNO and dc.SEQNO=pp.SEQNO and dc.STRAINDATE=v2.STRAINDATE ORDER BY dc.pname FOR XML PATH('')),1,1,'')) TPERIOD28_N" & vbCrLf
        sql &= " ,(SELECT STUFF((SELECT '；'+dc.pname FROM V_TRAINDESC dc WHERE dc.PLANID=pp.PLANID and dc.COMIDNO=pp.COMIDNO and dc.SEQNO=pp.SEQNO and dc.STRAINDATE=v2.STRAINDATE ORDER BY dc.pname FOR XML PATH('')),1,1,'')) PONCLASS1" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC3(pp.PLANID,pp.COMIDNO,pp.SEQNO,v2.STRAINDATE,'TP1') TP1" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC3(pp.PLANID,pp.COMIDNO,pp.SEQNO,v2.STRAINDATE,'TP2') TP2" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC3(pp.PLANID,pp.COMIDNO,pp.SEQNO,v2.STRAINDATE,'TP3') TP3" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC3(pp.PLANID,pp.COMIDNO,pp.SEQNO,v2.STRAINDATE,'TPX') TPX" & vbCrLf

        sql &= " FROM dbo.PLAN_PLANINFO pp" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO cc ON cc.PlanID = pp.PlanID AND cc.comidno = pp.comidno AND cc.seqno = pp.seqno" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = pp.PlanID" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME v1 ON v1.RID = pp.RID" & vbCrLf
        sql &= " JOIN dbo.V_TRAINDESC2 v2 ON v2.PLANID=pp.PLANID AND v2.COMIDNO=pp.COMIDNO AND v2.SEQNO=pp.SEQNO" & vbCrLf
        '限制為只有正式儲存之班級
        sql &= " WHERE pp.IsApprPaper = 'Y'" & vbCrLf
        '已轉班
        sql &= " AND pp.TransFlag = 'Y'" & vbCrLf
        sql &= " AND cc.ISSUCCESS = 'Y'" & vbCrLf
        sql &= " AND cc.NOTOPEN = 'N'" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID = '" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years = '" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND pp.PlanID = '" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If

        If sRelShip <> "" Then sql &= " AND v1.RelShip LIKE @RelShip" & vbCrLf
        If Hid_OCID1.Value <> "" Then sql &= " AND cc.OCID = @OCID" & vbCrLf
        If STRAINDATE.Text <> "" Then sql &= " AND v2.STRAINDATE = @STRAINDATE" & vbCrLf

        Dim parms As New Hashtable
        If sRelShip <> "" Then parms.Add("RelShip", String.Concat(sRelShip, "%"))
        If Hid_OCID1.Value <> "" Then parms.Add("OCID", Val(Hid_OCID1.Value))
        If STRAINDATE.Text <> "" Then parms.Add("STRAINDATE", TIMS.Cdate2(STRAINDATE.Text))

        '28:產業人才投資方案
        'Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
        'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    Select Case v_PlanPoint 'PlanPoint.SelectedValue
        '        Case "1"
        '            sql &= " AND v1.OrgKind <> '10'" & vbCrLf '產業人才投資計畫
        '        Case "2"
        '            sql &= " AND v1.OrgKind = '10'" & vbCrLf '提升勞工自主學習計畫
        '    End Select
        'End If
        sql &= " ORDER BY v2.STRAINDATE" & vbCrLf

        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return '查無資料離開

        DataGrid1.Visible = True
        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary> 查詢前檢核 </summary>
    ''' <param name="errMsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False
        Hid_OCID1.Value = ""

        'DataGrid1.Visible = False
        'PageControler1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then errMsg &= "請選擇 職類/班別!" & vbCrLf
        If OCIDValue1.Value <> "" Then
            Dim drOC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drOC Is Nothing Then errMsg &= "職類/班別 選擇有誤!" & vbCrLf
        End If

        STRAINDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(STRAINDATE.Text))
        If STRAINDATE.Text <> "" AndAlso Not TIMS.IsDate1(STRAINDATE.Text) Then errMsg &= "上課日期 格式有誤!" & vbCrLf
        'If STRAINDATE.Text = "" Then errMsg &= "請輸入 上課日期!" & vbCrLf

        rst = If(errMsg <> "", False, True)
        If errMsg <> "" Then Return rst

        '固定 職類/班別
        Hid_OCID1.Value = OCIDValue1.Value

        Return rst
    End Function

    ''' <summary> 列印前檢核 (確認列印參數) </summary>
    ''' <param name="errMsg"></param>
    ''' <returns></returns>
    Function CheckData2(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False

        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        If Hid_OCID1.Value = "" Then errMsg &= "請重新選擇 職類/班別!" & vbCrLf
        If Hid_OCID1.Value <> "" Then
            Dim drOC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
            If drOC Is Nothing Then errMsg &= "職類/班別 有誤!" & vbCrLf
        End If

        hidSTRAINDATE.Value = TIMS.Cdate3(TIMS.ClearSQM(hidSTRAINDATE.Value))
        If hidSTRAINDATE.Value <> "" AndAlso Not TIMS.IsDate1(hidSTRAINDATE.Value) Then errMsg &= "上課日期 格式有誤!" & vbCrLf
        If hidSTRAINDATE.Value = "" Then errMsg &= "請輸入 上課日期" & vbCrLf

        hidPCSValue.Value = TIMS.ClearSQM(hidPCSValue.Value)
        If hidPCSValue.Value = "" Then errMsg &= "查無班級資料!" & vbCrLf
        rst = If(errMsg <> "", False, True)
        If errMsg <> "" Then Return rst

        Dim flag_chk2 As Boolean = TIMS.CHK_TRAINDESC_STRAINDATE(hidPCSValue.Value, hidSTRAINDATE.Value, objconn)
        If Not flag_chk2 Then errMsg &= "查無班級上課日期資料!" & vbCrLf
        rst = If(errMsg <> "", False, True)
        If errMsg <> "" Then Return rst

        '學員名單列印範圍 1:錄訓作業正取名單(正取) 2:完成報到學員名單(學員)
        Dim v_StudlistRange As String = TIMS.GetListValue(rbl_StudlistRange)

        If (v_StudlistRange = "1") Then
            Dim iStd As Integer = CHK_STUD_ADMISSION_Y(Hid_OCID1.Value)
            If iStd = 0 Then errMsg &= "查無班級錄取人數資料!" & vbCrLf
        ElseIf (v_StudlistRange = "2") Then
            Dim iStd As Integer = CHK_STUD_ENTERTYPE(Hid_OCID1.Value)
            If iStd = 0 Then errMsg &= "查無班級完成報到學員資料!" & vbCrLf
        End If
        If errMsg <> "" Then Return rst

        Return rst
    End Function

    ''' <summary>檢核錄取人數</summary>
    ''' <param name="v_OCID"></param>
    ''' <returns></returns>
    Function CHK_STUD_ADMISSION_Y(ByRef v_OCID As String) As Integer
        Dim rst As Integer = 0
        If v_OCID = "" Then Return rst

        Dim parms As New Hashtable From {{"OCID", v_OCID}}
        Dim sql As String = ""
        'sql &= " ,concat(d.NAME,CASE cs.StudStatus WHEN '2' THEN N'(離訓)' WHEN '3' THEN N'(退訓)' ELSE '' END) STDNAME" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY d.SETID) AS ROWSETID" & vbCrLf
        sql &= " ,d.IDNO ,d.SETID ,cs.SOCID" & vbCrLf
        sql &= " FROM STUD_SELRESULT b" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE a ON a.SETID=b.SETID AND a.EnterDate=b.EnterDate AND a.SerNum=b.SerNum AND a.OCID1=b.OCID" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP d ON d.SETID=a.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO e ON e.OCID=b.OCID" & vbCrLf
        sql &= " LEFT JOIN STUD_STUDENTINFO ss on ss.IDNO=d.IDNO" & vbCrLf
        sql &= " LEFT JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid=a.ocid1 AND cs.SID=ss.SID" & vbCrLf
        sql &= " WHERE b.OCID=@OCID" '#{OCID}" & vbCrLf
        sql &= " and b.ADMISSION='Y'" & vbCrLf ' /* 140288 是否錄取:錄取 (除了正取才是錄取) N 不錄取 */" & vbCrLf
        sql &= " and b.SelResultID='01'" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return rst
        rst = dt.Rows.Count
        Return rst
    End Function

    ''' <summary>班級完成報到學員資料</summary>
    ''' <param name="v_OCID"></param>
    ''' <returns></returns>
    Function CHK_STUD_ENTERTYPE(ByRef v_OCID As String) As Integer
        Dim rst As Integer = 0
        If v_OCID = "" Then Return rst
        Dim pms_s1 As New Hashtable From {{"OCID", v_OCID}}

        Dim sSql As String = ""
        'sSql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.NotExam DESC, d2.SIGNNO) AS ROWSETID" & vbCrLf
        'sSql &= " ,concat(d.NAME,CASE cs.StudStatus WHEN '2' THEN N'(離訓)' WHEN '3' THEN N'(退訓)' ELSE '' END) STDNAME" & vbCrLf
        sSql &= " SELECT ROW_NUMBER() OVER(ORDER BY d.SETID) AS ROWSETID" & vbCrLf
        sSql &= " ,d2.SIGNNO ,d.IDNO ,d.SETID ,cs.SOCID" & vbCrLf
        sSql &= " FROM STUD_ENTERTYPE a" & vbCrLf
        sSql &= " JOIN STUD_ENTERTEMP d ON d.SETID=a.SETID" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO e ON e.OCID=a.OCID1" & vbCrLf
        sSql &= " JOIN STUD_STUDENTINFO ss on ss.IDNO=d.IDNO" & vbCrLf
        sSql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.OCID=a.OCID1 AND cs.SID=ss.SID" & vbCrLf
        sSql &= " JOIN STUD_ENTERTYPE2 d2 on d2.SETID=a.SETID AND d2.EnterDate=a.EnterDate AND d2.SerNum=a.SerNum AND d2.OCID1=a.OCID1" & vbCrLf
        sSql &= " WHERE a.OCID1=@OCID" & vbCrLf
        'sSql &= " ORDER BY a.NotExam DESC, d2.SIGNNO" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_s1)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return rst
        rst = dt.Rows.Count
        Return rst
    End Function

    Protected Sub btnSearch1_Click(sender As Object, e As EventArgs) Handles btnSearch1.Click
        Dim sErrMsg As String = ""
        Call CheckData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        '1:已轉班
        Call Search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        Dim s_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim s_PCSValue As String = TIMS.GetMyValue(sCmdArg, "PCSValue")
        Dim s_OrgKindGW As String = TIMS.GetMyValue(sCmdArg, "OrgKindGW")
        Dim s_STRAINDATE As String = TIMS.GetMyValue(sCmdArg, "STRAINDATE")
        Dim flagNG As Boolean = If(e.CommandArgument = "" OrElse s_OCID = "" OrElse s_PCSValue = "", True, False)
        If flagNG Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Select Case s_OrgKindGW
            Case "G", "W"
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
        End Select

        Select Case e.CommandName
            Case cst_dg1eCM_Print1 '"Print1"
                Dim sErrMsg As String = ""
                Call CheckData1(sErrMsg)
                If sErrMsg <> "" Then
                    Common.MessageBox(Me, sErrMsg)
                    Exit Sub
                End If
                hidPCSValue.Value = s_PCSValue
                hidSTRAINDATE.Value = s_STRAINDATE 'STRAINDATE.Text
                Call CheckData2(sErrMsg)
                If sErrMsg <> "" Then
                    Common.MessageBox(Me, sErrMsg)
                    Exit Sub
                End If

                '學員名單列印範圍 1:錄訓作業正取名單(正取) 2:完成報到學員名單(學員)
                Dim v_StudlistRange As String = TIMS.GetListValue(rbl_StudlistRange)
                Dim prtstr As String = ""
                prtstr = String.Format("&TPlanID={0}&OCID={1}&PCSValue={2}&STRAINDATE={3}&STUDLISTRANGE{4}=1", sm.UserInfo.TPlanID, s_OCID, s_PCSValue, s_STRAINDATE, v_StudlistRange)

                If TIMS.Cst_TPlanID54.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, prtstr)
                Else
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, prtstr)
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HiddenField = e.Item.FindControl("OCID")
                Dim PCSValue As HiddenField = e.Item.FindControl("PCSValue")
                OCID.Value = Convert.ToString(drv("OCID"))
                PCSValue.Value = Convert.ToString(drv("PCSValue"))

                Dim btnPrint1 As Button = e.Item.FindControl("btnPrint1")
                btnPrint1.CommandName = cst_dg1eCM_Print1
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "PCSValue", Convert.ToString(drv("PCSValue")))
                TIMS.SetMyValue(sCmdArg, "OrgKindGW", Convert.ToString(drv("OrgKindGW")))
                TIMS.SetMyValue(sCmdArg, "STRAINDATE", Convert.ToString(drv("STRAINDATE")))
                'TIMS.SetMyValue(sCmdArg, "PMD", Convert.ToString(drv("PMD")))
                btnPrint1.CommandArgument = sCmdArg
        End Select
    End Sub

End Class
