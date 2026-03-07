Partial Class SD_03_012_add
    Inherits AuthBasePage

    'CLASS_CONFIRM /STUD_CONFIRM
    'SELECT STUDMODE,COUNT(1) CNT ,MIN(MODIFYDATE),MAX(MODIFYDATE) FROM STUD_CONFIRM GROUP BY STUDMODE
    Const cst_printFN1 As String = "SD_CONFIRM_XC"

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
        'PageControler1.PageDataGrid=DataGrid1

        If Not IsPostBack Then Call Create1()
    End Sub

    '第1次載入
    Sub Create1()
        msg.Text = ""
        'DataGridTable1.Visible=False
        'DataGridTable2.Visible=False
        DataGridTable1.Style("display") = "none"
        DataGridTable2.Style("display") = "none"

        BtnSave1.Visible = False
        BtnSave2.Visible = False
        BtnPrint1.Visible = False
        trODNUMBER.Visible = False

        BtnSave1.Attributes.Add("onclick", "return savedataCHK1();")
        BtnSave2.Attributes.Add("onclick", "return savedataCHK2();")

        Dim url1 As String = ""
        Dim ACT As String = TIMS.sUtl_GetRqValue(Me, "ACT")
        Dim OCID As String = TIMS.sUtl_GetRqValue(Me, "OCID")
        If OCID = "" Then
            'url1="SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
            'TIMS.Utl_Redirect(Me, objConn, url1)
            Call TIMS.CloseDbConn(objConn)
            url1 = "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
            Common.MessageBox(Me, TIMS.cst_NODATAMsg92)
            TIMS.Utl_Redirect(Me, objConn, url1)
        End If
        Dim CFGUID As String = TIMS.sUtl_GetRqValue(Me, "CFGUID")
        Dim CFSEQNO As String = TIMS.sUtl_GetRqValue(Me, "CFSEQNO")
        Dim INQUIRY_SCH As String = TIMS.sUtl_GetRqValue(Me, "INQUIRY_SCH")

        Hid_INQUIRY_SCH.Value = INQUIRY_SCH
        Hid_CFGUID.Value = CFGUID
        Hid_OCID.Value = OCID
        Hid_CFSEQNO.Value = CFSEQNO
        MenuTable.Visible = True

        Select Case ACT
            Case "EDIT1" '(檢視)
                'labActText.Text="報名名單"
                'Call Search2() '+參考查詢
                MenuTable.Visible = False
                Call Search1()
                Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")

            Case "EDIT2" '(檢視)
                'labActText.Text="報到名單"
                'BtnSave1.Visible=True
                BtnSave2.Visible = True '解鎖學員參訓作業
                BtnPrint1.Visible = True
                trODNUMBER.Visible = True
                Call Search1() '+參考查詢
                Call Search2()
                Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")

            Case "ADD"
                'labActText.Text="報到名單(新增)"
                BtnSave1.Visible = True
                BtnSave2.Visible = True '解鎖學員參訓作業
                'BtnPrint1.Visible=True
                trODNUMBER.Visible = True
                Call Search1() '+參考查詢
                Call Search2()
                Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")

            Case Else
                'url1="SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
                'TIMS.Utl_Redirect(Me, objConn, url1)
                Call TIMS.CloseDbConn(objConn)
                url1 = "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
                Common.MessageBox(Me, TIMS.cst_NODATAMsg92)
                TIMS.Utl_Redirect(Me, objConn, url1)
        End Select
    End Sub

    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        Hid_CFGUID.Value = TIMS.ClearSQM(Hid_CFGUID.Value)
        Hid_CFSEQNO.Value = TIMS.ClearSQM(Hid_CFSEQNO.Value)
        If Hid_CFGUID.Value <> "" Then RstMemo &= String.Concat("&CFGUID=", Hid_CFGUID.Value)
        If Hid_CFSEQNO.Value <> "" Then RstMemo &= String.Concat("&CFSEQNO=", Hid_CFSEQNO.Value)
        Return RstMemo
    End Function

    '報名名單－檢視
    Sub Search1()
        msg.Text = "查無資料!!"
        'DataGridTable1.Visible=False
        DataGridTable1.Style("display") = "none"
        DataGridTable2.Style("display") = "none"

        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Dim drOCID As DataRow = TIMS.GetOCIDDate(Hid_OCID.Value, objConn)
        If drOCID Is Nothing Then Exit Sub
        LabClassName1.Text = drOCID("ClassCName2") & "，訓練人數" & drOCID("TNum") & "人-報名名單"

        Dim parms As New Hashtable()
        parms.Add("OCID1", Hid_OCID.Value)
        Dim sql As String = ""
        sql &= " SELECT b.esernum" & vbCrLf
        sql &= " ,dbo.LPAD( ROW_NUMBER() OVER (ORDER BY b.esernum), 2, '0') ROWNUM2" & vbCrLf
        sql &= " ,b.esetid" & vbCrLf
        sql &= " ,a.name STUDNAME" & vbCrLf
        sql &= " ,a.idno" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, b.enterdate, 111) enterdate" & vbCrLf
        sql &= " ,b.relenterdate" & vbCrLf
        sql &= " ,dbo.DECODE6(b.EnterPath,'O','外網','o','內網','網路') EnterPath" & vbCrLf
        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " ,b.signUpStatus" & vbCrLf
        sql &= " ,dbo.DECODE6(b.signUpStatus,2,'失敗',0,' ','成功') signUpStatusN" & vbCrLf
        sql &= " ,b.signUpMemo,b.SignNo,b.ExamNo,b2.Uname" & vbCrLf
        sql &= " FROM STUD_ENTERTYPE2 b" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP2 a ON a.eSETID=b.eSETID" & vbCrLf
        sql &= " JOIN STUD_ENTERTRAIN2 b2 ON b2.eSerNum=b.eSerNum" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID=b.OCID1" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid=cc.planid" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME f ON f.RID=cc.RID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo ON oo.comidno=cc.comidno" & vbCrLf
        sql &= " LEFT JOIN KEY_IDENTITY mi ON mi.IdentityID=b2.MIdentityID" & vbCrLf
        sql &= " WHERE b.OCID1=@OCID1" & vbCrLf
        Select Case sm.UserInfo.TPlanID
            Case TIMS.Cst_TPlanID28
                sql &= " AND b.SignNo IS NOT NULL" & vbCrLf
        End Select
        Select Case sm.UserInfo.LID
            Case 0
                sql &= " AND ip.Years=@Years" & vbCrLf
                sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
            Case Else
                sql &= " AND ip.Years=@Years" & vbCrLf
                sql &= " AND ip.DistID=@DistID" & vbCrLf
                sql &= " AND ip.PlanID=@PlanID" & vbCrLf
                sql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("PlanID", sm.UserInfo.PlanID)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
        End Select
        sql &= " ORDER BY b.SignNo ,b.eSerNum" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.ClearSQM(Hid_INQUIRY_SCH.Value)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "ESETID,ESERNUM,STUDNAME,IDNO,RELENTERDATE,ENTERPATH,SIGNUPSTATUSN,SIGNUPMEMO")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, Hid_OCID.Value, sMemo, objConn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then Exit Sub

        msg.Text = ""
        'DataGridTable1.Visible=True
        DataGridTable1.Style("display") = ""
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '報到名單 / 報到名單(新增)
    Sub Search2()
        msg.Text = "查無資料!!"
        'DataGridTable2.Visible=False
        DataGridTable1.Style("display") = "none"
        DataGridTable2.Style("display") = "none"

        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Dim drOCID As DataRow = TIMS.GetOCIDDate(Hid_OCID.Value, objConn)
        If drOCID Is Nothing Then Exit Sub
        LabClassName2.Text = drOCID("ClassCName2") & "，訓練人數" & drOCID("TNum") & "人-報到名單"

        Dim iCFSEQNO As Integer = TIMS.Get_CFSEQNO1(Hid_OCID.Value, objConn) '報到名單(新增)
        If Hid_CFSEQNO.Value <> "" Then iCFSEQNO = Val(Hid_CFSEQNO.Value)

        Dim sql As String = ""
        'sql="" & vbCrLf
        sql &= " SELECT dbo.FN_GET_PLANKIND3(rr.ORGKIND2,vp.TPLANID,vp.PLANNAME) TOPTITLE2" & vbCrLf
        sql &= " ,rr.OrgName" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,ss.Name STUDNAME" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        'sql &= ",dbo.SUBSTR(a.IDNO,1,4) + '*****' + dbo.SUBSTR(a.IDNO,-1,1) IDNO2" & vbCrLf
        sql &= " ,cs.StudentID" & vbCrLf
        sql &= " ,dbo.FN_CSTUDID2(cs.StudentID) STUDID2" & vbCrLf
        sql &= " ,CONVERT(VARCHAR ,cc.STDate ,111) STDate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR ,cc.FTDate ,111) FTDate" & vbCrLf
        sql &= " ,b.RelEnterDate" & vbCrLf
        sql &= " ,c.MDate" & vbCrLf
        sql &= " ,dbo.DECODE10(c.ChangeMode,1,'工作部門或特殊身分異動',2,'退保',3,'調薪',4,'加保','') ChangeMode" & vbCrLf
        sql &= " ,c.ComName" & vbCrLf
        sql &= " ,c.ActNo" & vbCrLf
        sql &= " ,CASE WHEN SUBSTRING(c.ACTNO,1,2)='09' THEN '不符合' WHEN c.CHANGEMODE NOT IN (2) THEN '符合' ELSE '不符合' END" & vbCrLf
        sql &= "  +dbo.DECODE(b.CMASTER1,'Y','(負責人不適用就保)','') CapMode" & vbCrLf
        sql &= " ,cs.STUDSTATUS" & vbCrLf
        sql &= " ,case when bb.BUDID is null then cs.BUDGETID ELSE bb.BUDID end BUDID" & vbCrLf
        sql &= " ,case when bb.BUDID is null then bcs.BUDNAME ELSE bb.BUDNAME end BUDNAME" & vbCrLf
        sql &= " ,b.SETID ,b.ENTERDATE ,b.SERNUM" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf
        sql &= " ,CASE WHEN sf.SFID IS NOT NULL THEN 'Y' END SFIDY" & vbCrLf

        sql &= " FROM STUD_ENTERTYPE b" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a ON a.SETID=b.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID=b.OCID1" & vbCrLf
        sql &= " JOIN VIEW_PLAN vp ON vp.PlanID=cc.PlanID" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME rr ON rr.RID=cc.RID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.IDNO=a.IDNO" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.sid=ss.sid AND cs.ocid=b.ocid1" & vbCrLf
        sql &= " LEFT JOIN STUD_CONFIRM sf ON sf.OCID=cs.OCID AND sf.SOCID=cs.SOCID AND sf.CFSEQNO=@CFSEQNO" & vbCrLf '報到名單(新增)
        sql &= " LEFT JOIN STUD_BLIGATEDATA28 c ON c.socid=cs.socid AND c.idno=a.idno" & vbCrLf
        sql &= " LEFT JOIN VIEW_BUDGET bb ON bb.BUDID=b.BUDID" & vbCrLf
        sql &= " LEFT JOIN VIEW_BUDGET bcs ON bcs.BUDID=cs.BUDGETID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'Select Case sm.UserInfo.TPlanID
        '    Case TIMS.Cst_TPlanID28
        '        sql &= " AND b.SignNo IS NOT NULL" & vbCrLf
        'End Select

        Dim parms As New Hashtable()
        parms.Add("CFSEQNO", iCFSEQNO)
        Select Case sm.UserInfo.LID
            Case "0"
                sql &= " AND vp.Years=@Years" & vbCrLf
                sql &= " AND vp.TPlanID=@TPlanID" & vbCrLf
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
            Case Else
                sql &= " AND vp.Years=@Years" & vbCrLf
                sql &= " AND vp.DistID=@DistID" & vbCrLf
                sql &= " AND vp.PlanID=@PlanID" & vbCrLf
                sql &= " AND vp.TPlanID=@TPlanID" & vbCrLf
                parms.Add("Years", sm.UserInfo.Years)
                parms.Add("DistID", sm.UserInfo.DistID)
                parms.Add("PlanID", sm.UserInfo.PlanID)
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
        End Select

        'sql &= " AND b.OCID1=97711" & vbCrLf
        sql &= " AND b.OCID1=@OCID1" & vbCrLf
        parms.Add("OCID1", Hid_OCID.Value)
        sql &= " ORDER BY cs.StudentID ,a.IDNO" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.ClearSQM(Hid_INQUIRY_SCH.Value)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDSTATUS,STUDNAME,IDNO,ACTNO,BUDNAME,CAPMODE,STUDSTATUS")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, Hid_OCID.Value, sMemo, objConn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            BtnSave2.Visible = False
            BtnPrint1.Visible = False
            Exit Sub
        End If

        msg.Text = ""
        'DataGridTable2.Visible=True
        DataGridTable2.Style("display") = ""
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()

        If Hid_CFGUID.Value = "" Then Exit Sub
        If Hid_CFSEQNO.Value = "" Then Exit Sub
        Dim sql2 As String = ""
        sql2 &= " SELECT cf.* FROM CLASS_CONFIRM cf "
        sql2 &= " WHERE cf.OCID=@OCID AND cf.CFGUID=@CFGUID AND cf.CFSEQNO=@CFSEQNO "
        Dim sCmd2 As New SqlCommand(sql2, objConn)
        'Call TIMS.OpenDbConn(objConn)
        Dim dt2 As New DataTable
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = Hid_OCID.Value
            .Parameters.Add("CFGUID", SqlDbType.VarChar).Value = Hid_CFGUID.Value
            .Parameters.Add("CFSEQNO", SqlDbType.VarChar).Value = Hid_CFSEQNO.Value
            'dt2.Load(.ExecuteReader())
            dt2 = DbAccess.GetDataTable(sCmd2.CommandText, objConn, sCmd2.Parameters)
        End With
        If dt2.Rows.Count = 0 Then Exit Sub

        Dim DR2 As DataRow = dt2.Rows(0)
        ODNUMBER.Text = TIMS.ClearSQM(DR2("ODNUMBER"))
        ODNUMBER.ReadOnly = True
        ODNUMBER.Enabled = False
        If ODNUMBER.Text <> "" Then TIMS.Tooltip(ODNUMBER, "公文文號不可修改,文號若有錯誤，請聯絡系統管理者!", True)
        BtnSave1.Visible = False
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labSeqno As Label = e.Item.FindControl("labSeqno")
                Dim labStdName As Label = e.Item.FindControl("labStdName")
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")

                Dim RelENTERDATE As Label = e.Item.FindControl("RelENTERDATE")
                Dim LabEnterPath As Label = e.Item.FindControl("LabEnterPath")
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                Dim signUpStatusN As Label = e.Item.FindControl("signUpStatusN")
                Dim signUpMemo As Label = e.Item.FindControl("signUpMemo")

                Dim ESETID As HtmlInputHidden = e.Item.FindControl("ESETID")
                Dim ESERNUM As HtmlInputHidden = e.Item.FindControl("ESERNUM")
                ESETID.Value = Convert.ToString(drv("ESETID"))
                ESERNUM.Value = Convert.ToString(drv("ESERNUM"))

                labSeqno.Text = Convert.ToString(drv("SignNo")) 'SignNo 
                If labSeqno.Text = "" Then labSeqno.Text = $"({drv("esernum")})"

                labStdName.Text = Convert.ToString(drv("STUDNAME"))
                labIDNO.Text = Convert.ToString(drv("IDNO"))

                RelENTERDATE.Text = Convert.ToString(drv("RelENTERDATE"))
                LabEnterPath.Text = Convert.ToString(drv("EnterPath"))
                signUpStatusN.Text = Convert.ToString(drv("signUpStatusN"))
                signUpMemo.Text = Convert.ToString(drv("signUpMemo"))
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labSeqno As Label = e.Item.FindControl("labSeqno")
                Dim labStdName As Label = e.Item.FindControl("labStdName")
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")
                Dim labActNo As Label = e.Item.FindControl("labActNo")
                Dim labBudget As Label = e.Item.FindControl("labBudget")
                Dim labCapMode As Label = e.Item.FindControl("labCapMode")
                Dim labStatus23 As Label = e.Item.FindControl("labStatus23")

                Dim SETID As HtmlInputHidden = e.Item.FindControl("SETID")
                Dim EnterDate As HtmlInputHidden = e.Item.FindControl("EnterDate")
                Dim SerNum As HtmlInputHidden = e.Item.FindControl("SerNum")
                Dim SOCID As HtmlInputHidden = e.Item.FindControl("SOCID")
                Dim HStudStatus As HtmlInputHidden = e.Item.FindControl("HStudStatus")

                SETID.Value = Convert.ToString(drv("SETID"))
                EnterDate.Value = TIMS.Cdate3(drv("EnterDate"))
                SerNum.Value = Convert.ToString(drv("SerNum"))
                SOCID.Value = Convert.ToString(drv("SOCID"))
                HStudStatus.Value = Convert.ToString(drv("StudStatus")) '學員狀態1/2/3

                labSeqno.Text = TIMS.Get_DGSeqNo(sender, e) '序號 
                If Convert.ToString(drv("SFIDY")) = "Y" Then labSeqno.Text &= "*"

                labStdName.Text = Convert.ToString(drv("STUDNAME"))
                labIDNO.Text = Convert.ToString(drv("IDNO"))
                labActNo.Text = Convert.ToString(drv("ActNo"))
                labBudget.Text = Convert.ToString(drv("BUDNAME")) '預算別
                labCapMode.Text = Convert.ToString(drv("CapMode"))

                Select Case Convert.ToString(drv("STUDSTATUS"))
                    Case "2", "3"
                        HStudStatus.Value = Convert.ToString(drv("STUDSTATUS"))
                        labStatus23.Text = "是"
                    Case Else
                        HStudStatus.Value = "" '非離退訓為正常清單，清除 HStudStatus 為null
                End Select
        End Select
    End Sub

    '回上一頁
    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Dim url1 As String = "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
        TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

    '儲存1-確認名單 CLASS_CONFIRM/STUD_CONFIRM
    Sub SaveData1()
        Dim sCFGUID As String = TIMS.GetGUID()
        Dim iCFSEQNO As Integer = TIMS.Get_CFSEQNO1(Hid_OCID.Value, objConn)
        ODNUMBER.Text = TIMS.ClearSQM(ODNUMBER.Text)

        Using oConn As SqlConnection = DbAccess.GetConnection()
            Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
            Try
                '遞補學員 使用SCMD's_type 'SF:取得確認狀態 SS:取得學員狀態
                Dim sCmd_SF1 As SqlCommand = TIMS.Get_StudCMDX("SF", oConn, oTrans)
                Dim sCmd_SS1 As SqlCommand = TIMS.Get_StudCMDX("SS", oConn, oTrans)
                'Dim sql As String=""

                Dim pParms As New Hashtable
                pParms.Add("CFGUID", sCFGUID)
                pParms.Add("OCID", Hid_OCID.Value)
                pParms.Add("CFSEQNO", iCFSEQNO) 'get_CFSEQNO(Hid_OCID.Value, objConn)
                pParms.Add("ODNUMBER", ODNUMBER.Text)
                pParms.Add("CREATEACCT", sm.UserInfo.UserID)
                pParms.Add("CONFIRACCT", sm.UserInfo.UserID) '確認者 
                pParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                Dim i1_sql As String = ""
                i1_sql &= " INSERT INTO CLASS_CONFIRM (CFGUID,OCID,CFSEQNO,ODNUMBER,CREATEACCT,CREATEDATE,CONFIRACCT,CONFIRDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
                i1_sql &= " VALUES (@CFGUID,@OCID,@CFSEQNO,@ODNUMBER,@CREATEACCT,GETDATE(),@CONFIRACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
                'Dim iCmd As New SqlCommand(sql, oConn, Trans) pParms.Clear()
                DbAccess.ExecuteNonQuery(i1_sql, oTrans, pParms)

                Dim i2_sql As String = ""
                i2_sql &= " INSERT INTO STUD_CONFIRM (SFID,OCID,CFSEQNO,SETID,ENTERDATE,SERNUM,SOCID,MODIFYACCT,MODIFYDATE,STUDMODE)" & vbCrLf
                i2_sql &= " VALUES (@SFID,@OCID,@CFSEQNO,@SETID,@ENTERDATE,@SERNUM,@SOCID,@MODIFYACCT,GETDATE(),@STUDMODE)" & vbCrLf
                'Dim iCmd2 As New SqlCommand(sql, oConn, Trans)

                For Each eItem As DataGridItem In DataGrid2.Items
                    Dim SETID As HtmlInputHidden = eItem.FindControl("SETID")
                    Dim EnterDate As HtmlInputHidden = eItem.FindControl("EnterDate")
                    Dim SerNum As HtmlInputHidden = eItem.FindControl("SerNum")
                    Dim SOCID As HtmlInputHidden = eItem.FindControl("SOCID")
                    Dim HStudStatus As HtmlInputHidden = eItem.FindControl("HStudStatus")

                    Dim vSTUDMODE_1 As String = "" 'A:遞補學員 2/3:離退訓 NULL:正常清單
                    '非第1次, 可能有-遞補學員(A)
                    If iCFSEQNO <> 1 Then
                        vSTUDMODE_1 = TIMS.Get_STUDMODE1(Hid_OCID.Value, iCFSEQNO - 1, SOCID.Value, sCmd_SF1, sCmd_SS1)
                    End If
                    Select Case HStudStatus.Value
                        Case "2", "3" '離退訓(若為離退，(上次)不管如何都為離退)
                            vSTUDMODE_1 = HStudStatus.Value
                    End Select

                    Dim iSFID As Integer = DbAccess.GetNewId(oTrans, "STUD_CONFIRM_SFID_SEQ,STUD_CONFIRM,SFID")
                    Dim pParms2 As New Hashtable
                    pParms2.Add("SFID", iSFID)
                    pParms2.Add("OCID", Hid_OCID.Value)
                    pParms2.Add("CFSEQNO", iCFSEQNO) 'get_CFSEQNO(Hid_OCID.Value, objConn)
                    pParms2.Add("SETID", SETID.Value)
                    pParms2.Add("EnterDate", TIMS.Cdate2(EnterDate.Value))
                    pParms2.Add("SERNUM", SerNum.Value)
                    pParms2.Add("SOCID", SOCID.Value)
                    pParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
                    pParms2.Add("STUDMODE", TIMS.GetValue1(vSTUDMODE_1))
                    DbAccess.ExecuteNonQuery(i2_sql, oTrans, pParms2)
                Next
                Call DbAccess.CommitTrans(oTrans)

            Catch ex As Exception
                Dim exMessage1 As String = ex.Message

                Dim strErrmsg As String = ""
                strErrmsg &= "ODNUMBER.Text: " & ODNUMBER.Text & vbCrLf
                strErrmsg &= "Hid_OCID.Value: " & Hid_OCID.Value & vbCrLf
                strErrmsg &= "sCFGUID: " & sCFGUID & vbCrLf
                strErrmsg &= "iCFSEQNO: " & iCFSEQNO & vbCrLf

                strErrmsg &= "/* ex.ToString */" & vbCrLf
                strErrmsg &= ex.ToString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                Call DbAccess.RollbackTrans(oTrans)
                Call TIMS.CloseDbConn(oConn)

                strErrmsg = ""
                strErrmsg &= "儲存作業-失敗!!" & vbCrLf
                strErrmsg &= "Message:" & exMessage1 & vbCrLf
                Common.MessageBox(Me, strErrmsg)
                Exit Sub 'Throw ex 'Exit Sub
            End Try
            Call TIMS.CloseDbConn(oConn)
        End Using
        Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")

        'url1=""
        'url1 &= "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
        'TIMS.Utl_Redirect(Me, objConn, url1)
        Call TIMS.CloseDbConn(objConn)
        Dim url1 As String = "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg1)
        TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

    '儲存2-解鎖學員參訓作業
    Sub SaveData2()
        Dim sCFGUID As String = Hid_CFGUID.Value
        Dim iCFSEQNO As Integer = Val(Hid_CFSEQNO.Value)

        Using oConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(oConn)
            Try
                Dim pParms As New Hashtable
                pParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                pParms.Add("CFGUID", sCFGUID)
                pParms.Add("OCID", Hid_OCID.Value)
                pParms.Add("CFSEQNO", iCFSEQNO) 'get_CFSEQNO(Hid_OCID.Value, objConn)
                Dim u_sql As String = ""
                u_sql &= " UPDATE CLASS_CONFIRM" & vbCrLf
                u_sql &= " SET NOLOCK='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                u_sql &= " WHERE CFGUID=@CFGUID AND OCID=@OCID AND CFSEQNO=@CFSEQNO" & vbCrLf
                DbAccess.ExecuteNonQuery(u_sql, Trans, pParms)

                Call DbAccess.CommitTrans(Trans)
            Catch ex As Exception
                Dim exMessage1 As String = ex.Message

                Dim strErrmsg As String = ""
                strErrmsg &= "ODNUMBER.Text: " & ODNUMBER.Text & vbCrLf
                strErrmsg &= "Hid_OCID.Value: " & Hid_OCID.Value & vbCrLf
                strErrmsg &= "sCFGUID: " & sCFGUID & vbCrLf
                strErrmsg &= "iCFSEQNO: " & iCFSEQNO & vbCrLf
                strErrmsg &= "/* ex.ToString */" & vbCrLf
                strErrmsg &= ex.ToString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                Call DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(oConn)

                strErrmsg = ""
                strErrmsg &= "儲存作業-失敗!!" & vbCrLf
                strErrmsg &= "Message:" & exMessage1 & vbCrLf
                Common.MessageBox(Me, strErrmsg)
                Exit Sub 'Throw ex 'Exit Sub
            End Try
            Call TIMS.CloseDbConn(oConn)
        End Using
        Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")
        'url1="SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
        'TIMS.Utl_Redirect(Me, objConn, url1)

        Call TIMS.CloseDbConn(objConn)
        Dim url1 As String = "SD_03_012.aspx?ID=" & TIMS.Get_MRqID(Me)
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg1)
        TIMS.Utl_Redirect(Me, objConn, url1)
    End Sub

    ''' <summary>double:true  1小時內同個文號，不可再次新增</summary>
    ''' <param name="s_ODNUMBER"></param>
    ''' <param name="OCID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function CHK_ODNUMBER_DOUB(ByVal s_ODNUMBER As String, ByVal OCID As String, ByRef oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        '該班已公告過此公文文號! 若為同一文號多次公告，請加註說明如：文號(二次公告)
        Dim sql As String = ""
        sql &= " SELECT 'X'"
        sql &= " FROM CLASS_CONFIRM"
        sql &= " WHERE CONVERT(date,CREATEDATE)=CONVERT(date,getdate())" '當天不行／隔天可再發送
        sql &= " AND CREATEDATE > (getdate()-(1.0/24))" '1小時內同個文號，不可再次新增
        sql &= " AND OCID=@OCID AND ODNUMBER=@ODNUMBER"
        Dim sCmd As New SqlCommand(sql, oConn)

        Dim dtCF As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCID)
            .Parameters.Add("ODNUMBER", SqlDbType.NVarChar).Value = s_ODNUMBER
            dtCF.Load(.ExecuteReader())
        End With
        If TIMS.dtHaveDATA(dtCF) Then rst = True '已存在
        Return rst
    End Function

    ''' <summary>
    ''' 取得上筆資料數
    ''' </summary>
    ''' <param name="iCFSEQNO_OLD"></param>
    ''' <param name="OCID"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function GET_STUDCOUNT1(ByVal iCFSEQNO_OLD As Integer, ByVal OCID As String, ByRef oConn As SqlConnection) As Integer
        Dim sql As String = ""
        sql &= " SELECT SETID ,EnterDate ,SerNum ,SOCID ,STUDMODE" & vbCrLf
        sql &= " FROM STUD_CONFIRM" & vbCrLf
        sql &= " WHERE OCID=@OCID AND CFSEQNO=@CFSEQNO" & vbCrLf
        Dim sCmd_SF As New SqlCommand(sql, oConn)

        Dim dtSF As New DataTable
        With sCmd_SF
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.Int).Value = Val(OCID)
            .Parameters.Add("CFSEQNO", SqlDbType.Int).Value = iCFSEQNO_OLD
            dtSF.Load(.ExecuteReader())
        End With

        Return dtSF.Rows.Count
    End Function

    ''' <summary>
    ''' 檢核1-確認名單
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Errmsg = ""
        ODNUMBER.Text = TIMS.ClearSQM(ODNUMBER.Text)
        If ODNUMBER.Text = "" Then Errmsg &= "公文文號 為必填資料<br>"
        If Errmsg <> "" Then Return False
        TIMS.OpenDbConn(objConn)

        'Dim sCFGUID As String=TIMS.GetGUID()
        Dim iCFSEQNO As Integer = TIMS.Get_CFSEQNO1(Hid_OCID.Value, objConn)
        'ODNUMBER.Text=TIMS.ClearSQM(ODNUMBER.Text)
        If Hid_CFSEQNO.Value = "" AndAlso iCFSEQNO = 1 Then Return True '第1次儲存(直接存)

        'double:true  1小時內同個文號，不可再次新增
        Dim flag_doub As Boolean = CHK_ODNUMBER_DOUB(ODNUMBER.Text, Hid_OCID.Value, objConn)
        If (flag_doub) Then Errmsg &= "該班已公告過此公文文號! 若為同一文號多次公告，請加註說明如：文號(二次公告)"
        If Errmsg <> "" Then Return False

        '非第1次進入檢核
        If iCFSEQNO <> 1 Then
            Dim iCFSEQNO_OLD As Integer = iCFSEQNO - 1

            '取得上筆資料數
            Dim i_SFCNT As Integer = GET_STUDCOUNT1(iCFSEQNO_OLD, Hid_OCID.Value, objConn)

            '筆數不相同
            If DataGrid2.Items.Count <> i_SFCNT Then Return True '可進入儲存

            Dim sCmd_SF1 As SqlCommand = TIMS.Get_StudCMDX("SF", objConn, Nothing)
            Dim sCmd_SS1 As SqlCommand = TIMS.Get_StudCMDX("SS", objConn, Nothing)
            '筆數相同(檢核是否有異動)
            For Each eItem As DataGridItem In DataGrid2.Items
                Dim SETID As HtmlInputHidden = eItem.FindControl("SETID")
                Dim EnterDate As HtmlInputHidden = eItem.FindControl("EnterDate")
                Dim SerNum As HtmlInputHidden = eItem.FindControl("SerNum")
                Dim SOCID As HtmlInputHidden = eItem.FindControl("SOCID")
                Dim HStudStatus As HtmlInputHidden = eItem.FindControl("HStudStatus")
                '非第1次, 可能有-遞補學員
                Dim STUDMODE As String = TIMS.Get_STUDMODE1(Hid_OCID.Value, iCFSEQNO_OLD, SOCID.Value, sCmd_SF1, sCmd_SS1)
                If STUDMODE <> HStudStatus.Value Then Return True '(狀態不同)可進入儲存
            Next
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-251
        '(2)儲存前檢核是否與上次資料不同，有異動才儲存。
        Errmsg &= "本次資料未產生異動，不可儲存!!<br>"
        Return False
    End Function

    ''' <summary>
    ''' 確認名單-儲存1
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSave1_Click(sender As Object, e As EventArgs) Handles BtnSave1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            'Common.MessageBox(Page, Errmsg)
            Page.RegisterStartupScript("Save1Err", "<script>blockAlert('" & Errmsg & "','',function(){ChangeMode(2);});</script>")
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    ''' <summary>
    ''' 解鎖學員參訓作業
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSave2_Click(sender As Object, e As EventArgs) Handles BtnSave2.Click
        '也就是在新增功能下， 不卡控公文文號為必填， 即可解鎖。同樣， 這個解鎖功能可使用角色為分署。
        Select Case sm.UserInfo.LID
            Case 0, 1
            Case Else
                Common.MessageBox(Me, "解鎖功能可使用角色為分署!")
                Exit Sub
        End Select
        SaveData2()
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Hid_CFGUID.Value = TIMS.ClearSQM(Hid_CFGUID.Value)
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Hid_CFSEQNO.Value = TIMS.ClearSQM(Hid_CFSEQNO.Value)
        Dim myValue As String = ""
        myValue &= "&CFGUID=" & Hid_CFGUID.Value
        myValue &= "&OCID=" & Hid_OCID.Value
        myValue &= "&CFSEQNO=" & Hid_CFSEQNO.Value
        TIMS.CloseDbConn(objConn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
        Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")
    End Sub
End Class