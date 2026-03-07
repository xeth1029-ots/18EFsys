Partial Class SD_09_018_R
    Inherits AuthBasePage

    'Const cst_titblue1 As String = "姓名藍色表該學員尚有自辦在職或接受委託訓練課程在訓中"

    Const cst_iPAGENUM As Integer = 10

    Const cst_INSUR_加保 As String = "Y"
    Const cst_INSUR_退保 As String = "N"

    '加保列印 : PRINT1 (正面／背面)
    Const cst_printFN1 As String = "SD_09_018_R"
    Const cst_printFN1B As String = "SD_09_018_RB"

    '退保列印 : PRINT2 (正面／背面)
    Const cst_printFN2 As String = "SD_09_018_R2"
    Const cst_printFN2B As String = "SD_09_018_R2B"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'labblue1.Text = cst_titblue1

        If Not IsPostBack Then
            Call cCreate1()
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

    Sub cCreate1()
        Call Utl_SHOWSCREEN(0)

        msg.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        tb_DataGrid1.Visible = False
        'PageControler1.Visible = False

        Dim s_javascript_btn2 As String = ""
        Dim s_LevOrg As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1")
        s_javascript_btn2 = String.Format("javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();", s_LevOrg)
        Button5.Attributes("onclick") = s_javascript_btn2

        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '查詢
    Private Sub BtnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnQuery.Click
        'Dim Errmsg As String = ""
        'Call CheckData1(Errmsg)
        'If Errmsg <> "" Then
        '    Common.MessageBox(Page, Errmsg)
        '    Exit Sub
        'End If

        Call sUtl_SEARCH1()
    End Sub

    ''' <summary>班級查詢</summary>
    Sub sUtl_SEARCH1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Dim v_rblTYPE1 As String = TIMS.GetListValue(rblTYPE1) '列印狀態:1:正面 /2:背面
        If v_rblTYPE1 = "" Then Return
        Hid_rblTYPE1.Value = v_rblTYPE1

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        '期別CyclType.Text = TIMS.FmtCyclType(CyclType.Text)

        Dim parms As New Hashtable From {{"YEARS", sm.UserInfo.Years.ToString()}, {"TPLANID", sm.UserInfo.TPlanID}}
        If RIDValue.Value <> "" Then parms.Add("RID", RIDValue.Value)
        If OCIDValue1.Value <> "" Then parms.Add("OCID", OCIDValue1.Value)
        '期別If CyclType.Text <> "" Then parms.Add("CyclType", CyclType.Text)

        Dim sql As String = ""
        sql &= " SELECT cc.YEARS,cc.TPLANID" & vbCrLf
        sql &= " ,cc.OCID,cc.RID,cc.CyclType" & vbCrLf
        sql &= " ,cc.DISTNAME,cc.DISTID" & vbCrLf
        sql &= " ,cc.PLANNAME,cc.PLANID" & vbCrLf
        sql &= " ,cc.ORGNAME,cc.COMIDNO" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cc.CLASSCNAME2" & vbCrLf
        sql &= " ,FORMAT(cc.modifydate,'mmssdd') MSD" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " WHERE cc.YEARS=@YEARS" & vbCrLf
        sql &= " AND cc.TPLANID=@TPLANID" & vbCrLf
        If RIDValue.Value <> "" Then sql &= " AND cc.RID=@RID" & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND cc.OCID=@OCID" & vbCrLf
        '期別If CyclType.Text <> "" Then sql &= " AND cc.CyclType=@CyclType" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料!!"
        tb_DataGrid1.Visible = False
        'PageControler1.Visible = False
        If dt.Rows.Count = 0 Then Return

        msg.Text = ""
        tb_DataGrid1.Visible = True
        'PageControler1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lkBtnPrint1 As LinkButton = e.Item.FindControl("lkBtnPrint1") '加保列印 : PRINT1
                Dim lkBtnPrint2 As LinkButton = e.Item.FindControl("lkBtnPrint2") '退保列印 : PRINT2
                Dim lkBtnExp1 As LinkButton = e.Item.FindControl("lkBtnExp1") '加保匯出 : EXPORT1
                Dim lkBtnExp2 As LinkButton = e.Item.FindControl("lkBtnExp2") '退保匯出 : EXPORT2

                'Dim lkBtnSELSTD1 As LinkButton = e.Item.FindControl("lkBtnSELSTD1") '挑選學員列印 : SELSTD1
                Dim lkBtnREMARKS1 As LinkButton = e.Item.FindControl("lkBtnREMARKS1") '【備註】設定 : REMARKS1
                'lkBtnSELSTD1.Visible = (Hid_rblTYPE1.Value = "1")
                Dim s_cmdarg As String = ""
                TIMS.SetMyValue(s_cmdarg, "RID", drv("RID"))
                TIMS.SetMyValue(s_cmdarg, "TPlanID", drv("TPlanID"))
                TIMS.SetMyValue(s_cmdarg, "OCID", drv("OCID"))
                TIMS.SetMyValue(s_cmdarg, "MSD", drv("MSD"))
                lkBtnPrint1.CommandArgument = s_cmdarg
                lkBtnPrint2.CommandArgument = s_cmdarg
                lkBtnExp1.CommandArgument = s_cmdarg
                lkBtnExp2.CommandArgument = s_cmdarg

                '挑選學員列印 : SELSTD1/'【備註】設定 : REMARKS1
                'lkBtnSELSTD1.CommandArgument = s_cmdarg
                lkBtnREMARKS1.CommandArgument = s_cmdarg
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim v_rblTYPE1 As String = TIMS.GetListValue(rblTYPE1) '列印狀態:1:正面 /2:背面
        If v_rblTYPE1 = "" OrElse e.CommandName = "" OrElse e.CommandArgument = "" Then Return
        Dim s_cmdarg As String = e.CommandArgument

        Dim vRID As String = TIMS.GetMyValue(s_cmdarg, "RID")
        Dim vTPlanID As String = TIMS.GetMyValue(s_cmdarg, "TPlanID")
        Dim vOCID As String = TIMS.GetMyValue(s_cmdarg, "OCID")
        Dim vMSD As String = TIMS.GetMyValue(s_cmdarg, "MSD")
        If vRID = "" OrElse vTPlanID = "" OrElse vOCID = "" OrElse vMSD = "" Then Return

        '加保列印 : PRINT1 / '退保列印 : PRINT2
        '挑選學員列印 : SELSTD1/'【備註】設定 : REMARKS1
        Dim s_printFN As String = ""
        Select Case e.CommandName '功能鈕
            Case "PRINT1" '加保列印 正面／反面 '挑選學員列印
                s_printFN = If(v_rblTYPE1 = "1", cst_printFN1, cst_printFN1B)
                If (v_rblTYPE1 = "1") Then
                    Hid_OCID1.Value = vOCID
                    Hid_MSD.Value = vMSD
                    Hid_INSUR.Value = cst_INSUR_加保 '加保列印 正面／反面
                    Call sUtl_SEARCH2(vOCID)
                    Return
                End If

            Case "PRINT2" '退保列印 正面／反面 '挑選學員列印
                s_printFN = If(v_rblTYPE1 = "1", cst_printFN2, cst_printFN2B)
                If (v_rblTYPE1 = "1") Then
                    Hid_OCID1.Value = vOCID
                    Hid_MSD.Value = vMSD
                    Hid_INSUR.Value = cst_INSUR_退保
                    Call sUtl_SEARCH2(vOCID)
                    Return
                End If
            Case "EXPORT1"
                Dim hParams As New Hashtable From {
                    {"RID", vRID},
                    {"TPlanID", vTPlanID},
                    {"OCID", vOCID},
                    {"MSD", vMSD},
                    {"INSUR", cst_INSUR_加保} '加保匯出
                    }
                Call sExprot1_SGR(hParams)
            Case "EXPORT2"
                Dim hParams As New Hashtable From {
                    {"RID", vRID},
                    {"TPlanID", vTPlanID},
                    {"OCID", vOCID},
                    {"MSD", vMSD},
                    {"INSUR", cst_INSUR_退保} '退保匯出
                    }
                Call sExprot1_SGR(hParams)
            Case "SELSTD1" '挑選學員列印
                'Hid_OCID1.Value = vOCID
                'Hid_MSD.Value = vMSD
                'Call sUtl_SEARCH2(vOCID)
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return
            Case "REMARKS1" '【備註】設定
                Call Utl_CLEARDATA3()
                Hid_OCID1.Value = vOCID
                Hid_MSD.Value = vMSD
                Call Utl_SHOWDATA3(vOCID)
                Return
        End Select
        If s_printFN = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If
        Dim iTNUM As Integer = TIMS.Get_STUDENTINFO2(objconn, vOCID).Rows.Count '開訓人數
        Dim iRNUM As Integer = TIMS.GET_PAGEROWNUM(iTNUM, cst_iPAGENUM)
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "RID", vRID)
        TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
        TIMS.SetMyValue(MyValue, "OCID", vOCID)
        TIMS.SetMyValue(MyValue, "MSD", vMSD)
        TIMS.SetMyValue(MyValue, "UserID", sm.UserInfo.UserID)
        TIMS.SetMyValue(MyValue, "RNUM", iRNUM)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, s_printFN, MyValue)
    End Sub

    ''' <summary>職災保加退保申報表-匯出</summary>
    ''' <param name="hParams"></param>
    Private Sub sExprot1_SGR(ByRef hParams As Hashtable)
        Dim vRID As String = TIMS.GetMyValue2(hParams, "RID")
        Dim vTPlanID As String = TIMS.GetMyValue2(hParams, "TPlanID")
        Dim vOCID As String = TIMS.GetMyValue2(hParams, "OCID")
        Dim vMSD As String = TIMS.GetMyValue2(hParams, "MSD")
        Dim vINSUR As String = TIMS.GetMyValue2(hParams, "INSUR")
        'hParams.Add("INSUR", cst_INSUR_退保)
        '1、 【加保匯出】欄位：保險證號、保險證號檢查碼、姓名、身分證號、出生日期、月薪資總額、性別
        '2、 【退保匯出】欄位：保險證號、保險證號檢查碼、姓名、身分證號、出生日期
        Dim sPMS As New Hashtable
        Dim sSql As String = ""
        Dim s_TableName As String = "職災保加保申報匯出"
        Select Case vINSUR
            Case cst_INSUR_加保
                s_TableName = "職災保加保申報匯出"
                sPMS.Clear()
                sPMS.Add("OCID", vOCID)
                sPMS.Add("RID", vRID)
                sPMS.Add("TPLANID", vTPlanID)
                sSql = ""
                sSql &= " SELECT (select substring(dbo.FN_GET_TACTNO(cs.DISTID),1,8)) 保險證號" & vbCrLf
                sSql &= " ,(select substring(dbo.FN_GET_TACTNO(cs.DISTID),9,1)) 保險證號檢查碼" & vbCrLf
                sSql &= " ,cs.NAME 姓名" & vbCrLf
                sSql &= " ,cs.IDNO 身分證號" & vbCrLf
                sSql &= " ,dbo.FN_CDATE4T2(cs.BIRTHDAY) 出生日期" & vbCrLf
                sSql &= " ,dbo.FN_GET_MONTHLYWAGE(cs.STDATE) 月薪資總額" & vbCrLf
                sSql &= " ,cs.SEX 性別" & vbCrLf ',cs.SEX2 性別
                sSql &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
                sSql &= " WHERE cs.OCID=@OCID AND cs.RID=@RID AND cs.TPLANID=@TPLANID" & vbCrLf
                sSql &= " ORDER BY cs.STUDID" & vbCrLf
            Case cst_INSUR_退保
                s_TableName = "職災保退保申報匯出"
                sPMS.Clear()
                sPMS.Add("OCID", vOCID)
                sPMS.Add("RID", vRID)
                sPMS.Add("TPLANID", vTPlanID)
                sSql = ""
                sSql &= " SELECT (select substring(dbo.FN_GET_TACTNO(cs.DISTID),1,8)) 保險證號" & vbCrLf
                sSql &= " ,(select substring(dbo.FN_GET_TACTNO(cs.DISTID),9,1)) 保險證號檢查碼" & vbCrLf
                sSql &= " ,cs.NAME 姓名" & vbCrLf
                sSql &= " ,cs.IDNO 身分證號" & vbCrLf
                sSql &= " ,dbo.FN_CDATE4T2(cs.BIRTHDAY) 出生日期" & vbCrLf
                sSql &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
                sSql &= " WHERE cs.OCID=@OCID AND cs.RID=@RID AND cs.TPLANID=@TPLANID" & vbCrLf
                sSql &= " ORDER BY cs.STUDID" & vbCrLf
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return
        End Select

        Dim dsXlsALL As New DataSet
        Dim dtXls1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        dtXls1.TableName = s_TableName
        dsXlsALL.Tables.Add(dtXls1)
        If dtXls1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        Dim sFileName1 As String = String.Concat(s_TableName, "_", TIMS.GetDateNo())
        'Dim s_titleRange As String = "A1:AG1,A1:AG1,A1:AG1"
        ExpClass1.Utl_Export1_XLSX(Me, dsXlsALL, sFileName1)

    End Sub

    ''' <summary>挑選學員列印</summary>
    ''' <param name="vOCID"></param>
    Private Sub sUtl_SEARCH2(ByVal vOCID As String)
        Hid_OCID1.Value = vOCID
        If vOCID = "" Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Call Utl_SHOWSCREEN(1)
        BtnPrint5.Visible = If(Hid_INSUR.Value = cst_INSUR_加保, True, False)
        BtnPrint6.Visible = If(Hid_INSUR.Value = cst_INSUR_退保, True, False)
        LabMsgSHOW2.Text = Convert.ToString(drCC("CLASSCNAME2"))

        Dim pParms As New Hashtable From {
            {"YEARS", sm.UserInfo.Years.ToString()},
            {"TPLANID", sm.UserInfo.TPlanID},
            {"OCID", vOCID}
        }

        Dim sSql As String = ""
        sSql = " SELECT cs.YEARS,cs.TPLANID" & vbCrLf
        sSql &= " ,cs.OCID,cs.RID,cs.CyclType,cs.CLASSCNAME2" & vbCrLf
        sSql &= " ,cs.DISTNAME,cs.DISTID" & vbCrLf
        sSql &= " ,cs.PLANNAME,cs.PLANID" & vbCrLf
        sSql &= " ,cs.ORGNAME,cs.COMIDNO" & vbCrLf
        sSql &= " ,format(cs.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sSql &= " ,format(cs.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sSql &= " ,cs.IDNO,cs.SOCID,cs.STUDENTID,cs.NAME,cs.BIRTHDAY,cs.STUDSTATUS" & vbCrLf
        sSql &= " ,(SELECT FORMAT(cc.modifydate,'mmssdd') FROM CLASS_CLASSINFO cc WHERE cc.OCID=cs.OCID) MSD" & vbCrLf
        'sSql &= " ,dbo.FN_EXCLUDETRAIN1(cs.IDNO,cs.OCID,cs.PLANID,cs.FTDATE) EXCLUDETRAIN1" & vbCrLf '等於0 表示無其他班級資料在訓 可退保
        'sSql &= " ,dbo.FN_EXCLUDETRAIN2(cs.IDNO,cs.OCID,cs.PLANID,cs.STDATE) EXCLUDETRAIN2" & vbCrLf '等於0 表示無其他班級資料開訓 可加保
        sSql &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
        sSql &= " WHERE cs.YEARS=@YEARS" & vbCrLf
        sSql &= " AND cs.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND cs.OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        labMsg2.Text = "查無資料!!"
        tb_DataGrid2.Visible = False
        'PageControler1.Visible = False
        If dt.Rows.Count = 0 Then Return

        labMsg2.Text = ""
        tb_DataGrid2.Visible = True
        'PageControler1.Visible = True
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()
        'PageControler1.PageDataTable = dt
        'PageControler1.ControlerLoad()
    End Sub

    Protected Sub BtnBACK2_Click(sender As Object, e As EventArgs) Handles BtnBACK2.Click
        Call Utl_SHOWSCREEN(0)
    End Sub

    Protected Sub BTNSAVE3_Click(sender As Object, e As EventArgs) Handles BTNSAVE3.Click
        Call SAVEDATA3()
    End Sub

    Private Sub SAVEDATA3()
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        If Hid_OCID1.Value = "" Then Return

        Dim sParms As New Hashtable From {
            {"OCID", Val(Hid_OCID1.Value)}
        }
        Dim sSql As String = " SELECT 1 FROM CLASS_SUBINFO2 WHERE OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)

        Dim rst As Integer = -1
        NOTEINSUR.Text = TIMS.ClearSQM(NOTEINSUR.Text)
        NOTESURR.Text = TIMS.ClearSQM(NOTESURR.Text)

        NOTEINSUR.Text = TIMS.Get_Substr1(NOTEINSUR.Text, 500)
        NOTESURR.Text = TIMS.Get_Substr1(NOTESURR.Text, 500)
        '新增
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Dim iParms As New Hashtable From {
                {"OCID", Val(Hid_OCID1.Value)},
                {"NOTEINSUR", NOTEINSUR.Text},
                {"NOTESURR", NOTESURR.Text},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim iSql As String = ""
            iSql = " INSERT INTO CLASS_SUBINFO2(OCID,NOTEINSUR,NOTESURR,MODIFYACCT,MODIFYDATE)" & vbCrLf
            iSql &= " VALUES (@OCID,@NOTEINSUR,@NOTESURR,@MODIFYACCT,GETDATE())" & vbCrLf
            rst = DbAccess.ExecuteNonQuery(iSql, objconn, iParms)
            Return
        End If

        '修改
        Dim uParms As New Hashtable From {
            {"OCID", Val(Hid_OCID1.Value)},
            {"NOTEINSUR", NOTEINSUR.Text},
            {"NOTESURR", NOTESURR.Text},
            {"MODIFYACCT", sm.UserInfo.UserID}
        }
        Dim uSql As String = ""
        uSql = " UPDATE CLASS_SUBINFO2" & vbCrLf
        uSql &= " SET NOTEINSUR=@NOTEINSUR ,NOTESURR=@NOTESURR" & vbCrLf
        uSql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " WHERE OCID=@OCID" & vbCrLf
        rst = DbAccess.ExecuteNonQuery(uSql, objconn, uParms)
        Return
    End Sub

    ''' <summary>清理一些欄位</summary>
    Private Sub Utl_CLEARDATA3()
        labMsgORGNAME3.Text = "" 'Convert.ToString(drCC("ORGNAME"))
        labMsgCLASSCNAME3.Text = "" 'Convert.ToString(drCC("CLASSCNAME2"))
        NOTEINSUR.Text = "" ' Convert.ToString(dr("NOTEINSUR"))
        NOTESURR.Text = "" 'Convert.ToString(dr("NOTESURR"))
    End Sub

    '''<summary>【備註】設定</summary>
    Sub Utl_SHOWDATA3(ByVal vOCID As String)
        Hid_OCID1.Value = vOCID 'TIMS.ClearSQM(Hid_OCID1.Value)
        If Hid_OCID1.Value = "" Then Return

        Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Call Utl_SHOWSCREEN(2)
        labMsgORGNAME3.Text = Convert.ToString(drCC("ORGNAME"))
        labMsgCLASSCNAME3.Text = Convert.ToString(drCC("CLASSCNAME2"))
        NOTEINSUR.Text = "" ' Convert.ToString(dr("NOTEINSUR"))
        NOTESURR.Text = "" 'Convert.ToString(dr("NOTESURR"))

        Dim sParms As New Hashtable From {
            {"OCID", Val(Hid_OCID1.Value)}
        }
        Dim sSql As String = " SELECT NOTEINSUR,NOTESURR,OCID FROM CLASS_SUBINFO2 WHERE OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
        Dim dr As DataRow = dt.Rows(0)

        NOTEINSUR.Text = Convert.ToString(dr("NOTEINSUR"))
        NOTESURR.Text = Convert.ToString(dr("NOTESURR"))
    End Sub


    Protected Sub BtnBACK3_Click(sender As Object, e As EventArgs) Handles BtnBACK3.Click
        Call Utl_SHOWSCREEN(0)
    End Sub

    '退保列印 正面
    Protected Sub BtnPrint6_Click(sender As Object, e As EventArgs) Handles BtnPrint6.Click
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        Hid_MSD.Value = TIMS.ClearSQM(Hid_MSD.Value)
        If Hid_OCID1.Value = "" Then Return
        If Hid_MSD.Value = "" Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim s_SOCIDS As String = GET_DG2_SOCIDS()
        If s_SOCIDS = "" Then
            Common.MessageBox(Me, "請勾選要列印的學員!")
            Return
        End If

        Dim vRID As String = Convert.ToString(drCC("RID"))
        Dim vTPlanID As String = Convert.ToString(drCC("TPlanID"))
        Dim vOCID As String = Convert.ToString(drCC("OCID"))
        Dim vMSD As String = Hid_MSD.Value ' drCC("MSD")

        Dim iTNUM As Integer = s_SOCIDS.Split(",").Length '開訓人數(勾選)
        Dim iRNUM As Integer = TIMS.GET_PAGEROWNUM(iTNUM, cst_iPAGENUM)
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "RID", vRID)
        TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
        TIMS.SetMyValue(MyValue, "OCID", vOCID)
        TIMS.SetMyValue(MyValue, "MSD", vMSD)
        TIMS.SetMyValue(MyValue, "UserID", sm.UserInfo.UserID)
        TIMS.SetMyValue(MyValue, "SOCIDS", s_SOCIDS)
        TIMS.SetMyValue(MyValue, "RNUM", iRNUM)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
    End Sub

    '加保列印 正面
    Protected Sub BtnPrint5_Click(sender As Object, e As EventArgs) Handles BtnPrint5.Click
        Hid_OCID1.Value = TIMS.ClearSQM(Hid_OCID1.Value)
        Hid_MSD.Value = TIMS.ClearSQM(Hid_MSD.Value)
        If Hid_OCID1.Value = "" Then Return
        If Hid_MSD.Value = "" Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(Hid_OCID1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim s_SOCIDS As String = GET_DG2_SOCIDS()
        If s_SOCIDS = "" Then
            Common.MessageBox(Me, "請勾選要列印的學員!")
            Return
        End If
        Dim vRID As String = Convert.ToString(drCC("RID"))
        Dim vTPlanID As String = Convert.ToString(drCC("TPlanID"))
        Dim vOCID As String = Convert.ToString(drCC("OCID"))
        Dim vMSD As String = Hid_MSD.Value ' drCC("MSD")

        Dim iTNUM As Integer = s_SOCIDS.Split(",").Length '開訓人數(勾選)
        Dim iRNUM As Integer = TIMS.GET_PAGEROWNUM(iTNUM, cst_iPAGENUM)
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "RID", vRID)
        TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
        TIMS.SetMyValue(MyValue, "OCID", vOCID)
        TIMS.SetMyValue(MyValue, "MSD", vMSD)
        TIMS.SetMyValue(MyValue, "UserID", sm.UserInfo.UserID)
        TIMS.SetMyValue(MyValue, "SOCIDS", s_SOCIDS)
        TIMS.SetMyValue(MyValue, "RNUM", iRNUM)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    ''' <summary>取得勾選資訊</summary>
    ''' <returns></returns>
    Private Function GET_DG2_SOCIDS() As String
        Dim s_SOCIDS_RST As String = ""
        For Each eItem As DataGridItem In DataGrid2.Items
            Dim Hid_SOCID As HiddenField = eItem.FindControl("Hid_SOCID")
            Dim Checkbox1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            If Hid_SOCID IsNot Nothing AndAlso Checkbox1 IsNot Nothing AndAlso Hid_SOCID.Value <> "" AndAlso Checkbox1.Checked Then
                s_SOCIDS_RST &= String.Concat(If(s_SOCIDS_RST <> "", ",", ""), Hid_SOCID.Value)
            End If
        Next
        Return s_SOCIDS_RST
    End Function

    ''' <summary>介面調整,1:班級查詢 /2:班級學員查詢 /3:【備註】設定</summary>
    ''' <param name="iTYPE"></param>
    Sub Utl_SHOWSCREEN(ByVal iTYPE As Integer)
        'iTYPE : 1/2/3 ,1:班級查詢 /2:班級學員查詢 /3:【備註】設定
        tb_CLASSSHOW1.Visible = False
        tb_SELSTD1_1.Visible = False
        tb_EDITDATA3.Visible = False
        If iTYPE = 0 Then tb_CLASSSHOW1.Visible = True '班級查詢
        If iTYPE = 1 Then tb_SELSTD1_1.Visible = True '班級學員查詢
        If iTYPE = 2 Then tb_EDITDATA3.Visible = True '【備註】設定
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lkBtnPrint3 As LinkButton = e.Item.FindControl("lkBtnPrint3") '加保列印 : PRINT3
                Dim lkBtnPrint4 As LinkButton = e.Item.FindControl("lkBtnPrint4") '退保列印 : PRINT4
                Dim Hid_SOCID As HiddenField = e.Item.FindControl("Hid_SOCID")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim labName As Label = e.Item.FindControl("labName") '  (藍色)姓名藍色表該學員尚有自辦在職課程在訓中
                Dim labgreen2 As Label = e.Item.FindControl("labgreen2") ' (綠色)* 該學員於此班已離退訓

                lkBtnPrint3.Visible = If(Hid_INSUR.Value = cst_INSUR_加保, True, False)
                lkBtnPrint4.Visible = If(Hid_INSUR.Value = cst_INSUR_退保, True, False)

                Hid_SOCID.Value = TIMS.ClearSQM(drv("SOCID"))
                labName.Text = TIMS.ClearSQM(drv("NAME"))

                Dim sStdInfo1 As String = GET_STUDCLASSINFO1(drv("IDNO"), drv("OCID"), drv("STDATE"), drv("FTDATE"))
                labName.ForeColor = If(sStdInfo1 <> "", Color.Blue, Color.Black)
                If sStdInfo1 <> "" Then TIMS.Tooltip(labName, sStdInfo1, True)

                Dim fg_STUDSTATUS23 As Boolean = If(Convert.ToString(drv("STUDSTATUS")) = "2", True, If(Convert.ToString(drv("STUDSTATUS")) = "3", True, False))
                labgreen2.Visible = If(fg_STUDSTATUS23, True, False)
                Dim s_STUDSTATUS23 As String = If(Convert.ToString(drv("STUDSTATUS")) = "2", "已離訓", If(Convert.ToString(drv("STUDSTATUS")) = "3", "已退訓", ""))
                If s_STUDSTATUS23 <> "" Then TIMS.Tooltip(labgreen2, s_STUDSTATUS23, True)

                'lkBtnPrint3.Enabled = (Convert.ToString(drv("EXCLUDETRAIN2")) = "0") '等於0 表示 可加保
                'lkBtnPrint4.Enabled = (Convert.ToString(drv("EXCLUDETRAIN1")) = "0") '等於0 表示 可退保
                'If Not lkBtnPrint3.Enabled Then TIMS.Tooltip(lkBtnPrint3, "學員在其他班級已開訓")
                'If Not lkBtnPrint4.Enabled Then TIMS.Tooltip(lkBtnPrint3, "學員在其他班級有開訓")
                Dim s_cmdarg As String = ""
                TIMS.SetMyValue(s_cmdarg, "RID", drv("RID"))
                TIMS.SetMyValue(s_cmdarg, "TPlanID", drv("TPlanID"))
                TIMS.SetMyValue(s_cmdarg, "OCID", drv("OCID"))
                TIMS.SetMyValue(s_cmdarg, "MSD", drv("MSD"))
                TIMS.SetMyValue(s_cmdarg, "SOCID", drv("SOCID"))
                lkBtnPrint3.CommandArgument = s_cmdarg
                lkBtnPrint4.CommandArgument = s_cmdarg
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        If e.CommandName = "" Then Return
        If e.CommandArgument = "" Then Return
        Dim s_cmdarg As String = e.CommandArgument

        Dim vRID As String = TIMS.GetMyValue(s_cmdarg, "RID")
        Dim vTPlanID As String = TIMS.GetMyValue(s_cmdarg, "TPlanID")
        Dim vOCID As String = TIMS.GetMyValue(s_cmdarg, "OCID")
        Dim vMSD As String = TIMS.GetMyValue(s_cmdarg, "MSD")
        Dim vSOCIDS As String = TIMS.GetMyValue(s_cmdarg, "SOCID")

        '加保列印 : PRINT1 / '退保列印 : PRINT2
        '挑選學員列印 : SELSTD1/'【備註】設定 : REMARKS1
        Dim s_printFN As String = ""
        Select Case e.CommandName '功能鈕
            Case "PRINT3" '加保列印 正面
                Dim MyValue As String = ""
                TIMS.SetMyValue(MyValue, "RID", vRID)
                TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
                TIMS.SetMyValue(MyValue, "OCID", vOCID)
                TIMS.SetMyValue(MyValue, "MSD", vMSD)
                TIMS.SetMyValue(MyValue, "UserID", sm.UserInfo.UserID)
                TIMS.SetMyValue(MyValue, "SOCIDS", vSOCIDS)
                TIMS.SetMyValue(MyValue, "RNUM", cst_iPAGENUM)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

            Case "PRINT4" '退保列印 正面
                Dim MyValue As String = ""
                TIMS.SetMyValue(MyValue, "RID", vRID)
                TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
                TIMS.SetMyValue(MyValue, "OCID", vOCID)
                TIMS.SetMyValue(MyValue, "MSD", vMSD)
                TIMS.SetMyValue(MyValue, "UserID", sm.UserInfo.UserID)
                TIMS.SetMyValue(MyValue, "SOCIDS", vSOCIDS)
                TIMS.SetMyValue(MyValue, "RNUM", cst_iPAGENUM)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)

        End Select
    End Sub

    ''' <summary>
    ''' TPLANID:06	在職進修訓練/07	接受企業委託訓練
    ''' '當學員尚有「自辦在職課程」在訓中 (無離退訓)，於該學員姓名呈現藍色字，並於游標移至其姓名時，出現在訓中的「分署、班級名稱、訓練期間」。
    ''' </summary>
    ''' <param name="IDNO"></param>
    ''' <param name="OCID"></param>
    ''' <param name="STDATE"></param>
    ''' <param name="FTDATE"></param>
    ''' <returns></returns>
    Function GET_STUDCLASSINFO1(ByRef IDNO As String, ByRef OCID As String, ByRef STDATE As String, ByRef FTDATE As String) As String
        Dim rst As String = "" '{"TPLANID", sm.UserInfo.TPlanID},
        Dim pParms As New Hashtable From {
            {"IDNO", IDNO},
            {"OCID", Val(OCID)},
            {"YEARS", (sm.UserInfo.Years - 2).ToString()},
            {"STDATE", TIMS.Cdate2(STDATE)},
            {"FTDATE", TIMS.Cdate2(FTDATE)}
        }

        Dim sSql As String = "" 'sSql = "" & vbCrLf
        sSql &= " SELECT cs.IDNO,cs.OCID,cs.DISTNAME,cs.CLASSCNAME2,cs.STUDSTATUS" & vbCrLf
        sSql &= " ,concat(format(cs.STDATE,'yyyy/MM/dd'),'-',format(cs.FTDATE,'yyyy/MM/dd')) SFTDATE" & vbCrLf
        sSql &= " FROM dbo.V_STUDENTINFO cs" & vbCrLf
        sSql &= " WHERE cs.IDNO=@IDNO AND cs.OCID!=@OCID" & vbCrLf
        'sSql &= " AND cs.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND cs.TPLANID IN ('06','07')" & vbCrLf
        sSql &= " AND cs.YEARS>=@YEARS" & vbCrLf
        sSql &= " AND ((cs.STDATE<=@STDATE AND cs.FTDATE>=@STDATE)
OR (cs.STDATE<=@FTDATE AND cs.FTDATE>=@FTDATE) 
OR (@STDATE<=cs.STDATE AND cs.STDATE<=@FTDATE))" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If dt.Rows.Count = 0 Then Return rst
        For Each dr As DataRow In dt.Rows
            rst &= String.Concat(If(rst <> "", "、", ""), String.Format("{0}-{1}-{2}", dr("DISTNAME"), dr("CLASSCNAME2"), dr("SFTDATE")))
        Next
        Return rst
    End Function

End Class
