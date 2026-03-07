Partial Class TC_01_015
    Inherits AuthBasePage

    Const cst_PunishPeriod_autotxt As String = "(系統自動計算)"

    Const cst_altmsg1 As String = "系統裡查無此統一編號,請輸入系統裡已建立之訓練機構統一編號!!"
    Const cst_altmsg2 As String = "未選擇處分緣由,若無請選擇其他!!"
    Const cst_altmsg2b As String = "未選擇處分緣由!!"

    Const cst_altmsg38 As String = "處分緣由選擇第38項,處分年限 限定0年或1年!!"
    Const cst_altmsg39 As String = "處分緣由選擇第39項,處分年限 限定1年!!"
    Const cst_altmsg40 As String = "處分緣由選擇第40項,處分年限 限定1~3年!!"
    Const cst_altmsg42 As String = "處分緣由選擇第42項,處分年限 限定2年!!"

    Const cst_altmsg6 As String = "未輸入有效統一編號，統一編號為必填資料!!"
    Const cst_altmsg507 As String = "處分緣由選擇第７點,處分年限 限定1~2年!!"
    Const cst_altmsg520 As String = "處分緣由選擇第20點,處分年限 限定1年!!"
    Const cst_altmsg521 As String = "處分緣由選擇第21點,處分年限 限定2年!!"

    Const Cst_事由位置 As Integer = 8
    Const Cst_功能欄位 As Integer = 9

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()
    Dim bl_printMode As Boolean = False   '判斷目前是否為[匯出]模式，by:20181001

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        '(直接在AuthBasePage處理,不用個別檢查Session) TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        '原先『查詢』區塊的『訓練機構』子視窗功能
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        '"20180706 add『新增/修改』區塊的『訓練機構』子視窗功能"
        'Const cst_javascript_openOrg_FMT1 As String="javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        OrgX1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRIDX1, "HistoryList2X1", "RIDValueX1", "centerX1")
        If HistoryRIDX1.Rows.Count <> 0 Then
            centerX1.Attributes("onclick") = "showObj('HistoryList2X1');"
            centerX1.Style("CURSOR") = "hand"
        End If
        '"20180726 add『新增/修改』區塊的『班級名稱』子視窗功能"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValueX1", "centerX1", "TMIDValue1", "TMID1", False, "")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        '"20180706 先暫時拿掉此功能"
        'btnAdds.Enabled=False
        'If blnCanAdds Then btnAdds.Enabled=True
        'btnSearch.Enabled=False
        'If blnCanSech Then btnSearch.Enabled=True

        If Not IsPostBack Then
            Call sCreate1()
            Call sSearch1() '查詢
        End If

        btn_Save.Attributes("onclick") = "return chkdata();"
    End Sub

    Sub sCreate1()
        '(直接在AuthBasePage處理,不用個別檢查Session) If TIMS.ChkSession(Me) Then Exit Sub
        DivOutputDoc.Visible = False

        ddlTPlanIDSch = TIMS.Get_TPlan(ddlTPlanIDSch, , , , , objconn)
        Common.SetListItem(ddlTPlanIDSch, sm.UserInfo.TPlanID)

        ddlTPlanID = TIMS.Get_TPlan(ddlTPlanID, , , , , objconn)
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)

        DistID = TIMS.Get_DistID(DistID)
        ddl_DistID = TIMS.Get_DistID(ddl_DistID)
        Common.SetListItem(ddl_DistID, sm.UserInfo.DistID)

        ddl_DistID.Enabled = False
        Years = TIMS.Get_Years(Years)
        Years.Items.Insert(0, New ListItem("==請選擇==", ""))

        '依計畫顯示不同的緣由條文。
        ddlOBTERMS.Items.Clear()
        ddl_OBYears.Items.Clear()
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            ddlOBTERMS = TIMS.Get_BTERMS(ddlOBTERMS, 1)
            With ddl_OBYears.Items
                .Add(New ListItem("0年", "0"))
                .Add(New ListItem("1年", "1"))
                .Add(New ListItem("2年", "2"))
                .Add(New ListItem("3年", "3"))
            End With
        End If

        'If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    ddlOBTERMS=TIMS.Get_BTERMS(ddlOBTERMS, 4)
        '    '處份年限
        '    With ddl_OBYears.Items
        '        .Add(New ListItem("0年", "0"))
        '        .Add(New ListItem("1年", "1"))
        '        .Add(New ListItem("2年", "2"))
        '    End With
        'End If

        '20180710 (因使用者所屬計畫不屬於上述提及的,導致下拉式選單出不來,故暫時先用28,54那種)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 And TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            ddlOBTERMS = TIMS.Get_BTERMS(ddlOBTERMS, 1)
            With ddl_OBYears.Items
                .Add(New ListItem("0年", "0"))
                .Add(New ListItem("1年", "1"))
                .Add(New ListItem("2年", "2"))
                .Add(New ListItem("3年", "3"))
            End With
        End If
    End Sub

    '查詢
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        Call sSearch1() '查詢
    End Sub

    ''' <summary>
    ''' 顯示狀況設定 Panel
    ''' </summary>
    ''' <param name="iType"></param>
    Sub sUtl_PanelList(ByVal iType As Integer)
        'iType:1 搜尋
        Panel1.Visible = False '搜尋
        Panel2.Visible = False '新增/修改
        Panel3.Visible = False
        Select Case iType
            Case 1
                Panel1.Visible = True '搜尋
            Case 2
                Panel2.Visible = True '新增/修改
            Case 3
                Panel3.Visible = True '檢視
        End Select
        'Panel
    End Sub

    ''' <summary>
    ''' 離開
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btn_lev2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev2.Click
        Call sUtl_PanelList(1) '搜尋
    End Sub

    '離開
    Private Sub btn_lev_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_lev.Click
        Call sUtl_PanelList(1) '搜尋
        Call ClearEdit1()
    End Sub

    Sub SHOW_DATA1(ByRef dr As DataRow)
        If dr Is Nothing Then Return

        ddlTPlanID.Enabled = False '新增/修改鎖定。
        Common.SetListItem(ddlTPlanID, Convert.ToString(dr("TPLANID")))
        ddl_DistID.Enabled = False
        Common.SetListItem(ddl_DistID, Convert.ToString(dr("DISTID")))
        '機構名稱
        centerX1.Text = Convert.ToString(dr("ORGNAME"))    '20180709
        '統一編號
        txt_ComIDNO.Text = Convert.ToString(dr("COMIDNO"))    '20180709
        '處分文號
        txt_No.Text = Convert.ToString(dr("OBNUM"))    '20180709
        TMID1.Text = Convert.ToString(dr("CKNAME"))    '20180709
        '班級名稱
        OCID1.Text = Convert.ToString(dr("CRNAME"))
        '班級名稱代號
        OCIDValue1.Value = Convert.ToString(dr("OCID"))
        TMIDValue1.Value = Convert.ToString(dr("TMID"))
        '申請金額
        txt_ApplyPrice.Text = Convert.ToString(dr("APPLYPRICE"))
        '核定金額
        txt_AuthPrice.Text = Convert.ToString(dr("AUTHPRICE"))
        '處分日期
        txt_OBSdate.Text = If(flag_ROC, TIMS.Cdate17(dr("OBSDATE")), Convert.ToString(dr("OBSDATE"))) 'edit，by:20181001
        Common.SetListItem(ddl_OBYears, Convert.ToString(dr("OBYEARS")))
        '處分期間
        txt_PunishPeriod.Text = If(Convert.ToString(dr("C_PunishPeriod")) <> "", Convert.ToString(dr("C_PunishPeriod")), cst_PunishPeriod_autotxt)
        '處分事由
        txt_OBComment.Text = Convert.ToString(dr("OBCOMMENT"))
        '處分緣由
        Common.SetListItem(ddlOBTERMS, Convert.ToString(dr("OBTERMS")))
        '處分事實
        txt_OBFact.Text = Convert.ToString(dr("OBFACT"))
        '是否會辦政風
        Dim s_ISLAW1 As String = If(Convert.ToString(dr("ISLAW1")).Equals("Y"), "Y", "N")
        Common.SetListItem(rbl_IsLaw1, s_ISLAW1)
        '是否移送檢調
        Dim s_ISLAW2 As String = If(Convert.ToString(dr("ISLAW2")).Equals("Y"), "Y", "N")
        Common.SetListItem(rbl_IsLaw2, s_ISLAW2)
        '移送情形
        txt_Transfer.Text = Convert.ToString(dr("TRANSFER")) '20180709
        '檢調偵查/判決日期
        txt_JudgeDate.Text = If(flag_ROC, TIMS.Cdate17(dr("JUDGEDATE")), Convert.ToString(dr("JUDGEDATE"))) 'edit，by:20181001

        '檢調偵查/判決文號
        txt_JudgeNum.Text = Convert.ToString(dr("JUDGENUM")) '20180709
        '檢調偵查/判決事實
        txt_JudgeFact.Text = Convert.ToString(dr("JUDGEFACT")) '20180709
        '後續待辦事項<br/>(追繳款項、強制執行狀況等
        txt_Tudo.Text = Convert.ToString(dr("TODO"))
        '備註
        txt_Note.Text = Convert.ToString(dr("NOTE"))
        labModifyDate.Text = Convert.ToString(dr("ModifyDate"))
    End Sub

    Sub SHOW_DATA2(ByRef dr As DataRow)
        If dr Is Nothing Then Return

        '計畫別
        lbl_PlanName.Text = Convert.ToString(dr("PLANNAME"))
        lbl_DistID.Text = Convert.ToString(dr("DISTNAME"))
        lab_OrgName.Text = Convert.ToString(dr("ORGNAME"))
        lab_ComIDNO.Text = Convert.ToString(dr("COMIDNO"))

        lbl_CRName.Text = Convert.ToString(dr("CRNAME"))
        lbl_ApplyPrice.Text = Convert.ToString(dr("APPLYPRICE"))
        lbl_AuthPrice.Text = Convert.ToString(dr("AUTHPRICE"))
        lbl_No.Text = Convert.ToString(dr("OBNUM"))
        lblOBTERMS.Text = TIMS.Get_OBTERMSName(TIMS.Get_OBTERM(), Convert.ToString(dr("OBTERMS")))

        'If Not dr("OBSDATE") Is Nothing Then lab_OBSDate.Text=dr("OBSDATE").ToString Else lab_OBSDate.Text=""
        lab_OBSDate.Text = If(flag_ROC, TIMS.Cdate17(dr("OBSDATE")), Convert.ToString(dr("OBSDATE"))) 'edit，by:20181001
        lab_OBYears.Text = Convert.ToString(dr("OBYEARS")) & "年"
        '處分期間
        lab_PunishPeriod.Text = Convert.ToString(dr("C_PunishPeriod"))

        lab_OBComment.Text = Convert.ToString(dr("OBCOMMENT"))
        lbl_OBFact.Text = Convert.ToString(dr("OBFACT"))

        Dim tLaw1 As String = ""
        tLaw1 = If(Convert.ToString(dr("ISLAW1")).ToUpper.Equals("Y"), "是", "否")
        lbl_IsLaw1.Text = tLaw1

        Dim tLaw2 As String = Nothing
        tLaw2 = If(Convert.ToString(dr("ISLAW2")).ToUpper.Equals("Y"), "是", "否")
        lbl_IsLaw2.Text = tLaw2

        lbl_Transfer.Text = Convert.ToString(dr("TRANSFER"))
        lbl_JudgeDate.Text = If(flag_ROC, TIMS.Cdate17(dr("JUDGEDATE")), Convert.ToString(dr("JUDGEDATE"))) 'edit，by:20181001

        lbl_JudgeNum.Text = Convert.ToString(dr("JUDGENUM")) '.ToString Else lbl_JudgeNum.Text=""
        lbl_JudgeFact.Text = Convert.ToString(dr("JUDGEFACT")) '.ToString Else lbl_JudgeFact.Text=""
        lbl_Todo.Text = Convert.ToString(dr("TODO")) '.ToString Else lbl_Todo.Text=""
        lbl_Note.Text = Convert.ToString(dr("NOTE")) '.ToString Else Me.lbl_Note.Text=""
        lab_accName.Text = Convert.ToString(dr("ACCNAME")) '.ToString Else Me.lab_accName.Text=""
        lab_Modifydate.Text = Convert.ToString(dr("ModifyDate"))

    End Sub

    '查詢
    Sub sSearch1()
        '(直接在AuthBasePage處理,不用個別檢查Session) If TIMS.ChkSession(Me) Then Exit Sub
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        RIDValueX1.Value = TIMS.ClearSQM(RIDValueX1.Value)  '20180709
        ComidValue.Text = TIMS.ClearSQM(ComidValue.Text)
        txtOrgName.Text = TIMS.ClearSQM(txtOrgName.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        If RIDValue.Value <> "" Then
            sql &= " WITH WAR1 AS (" & vbCrLf
            sql &= " SELECT r.ORGID FROM AUTH_RELSHIP r WHERE r.RID LIKE @RID+'%'" & vbCrLf
            sql &= " )" & vbCrLf
        End If
        sql &= " SELECT DISTINCT c.NAME DISTNAME ,a.DISTID ,a.OBSN" & vbCrLf
        sql &= " ,ISNULL(b.ORGNAME,'-') ORGNAME" & vbCrLf
        sql &= " ,a.COMIDNO" & vbCrLf
        'sql &= " ,CONVERT(VARCHAR, a.OBSDATE, 111) OBSDATE" & vbCrLf
        'If bl_printMode And flag_ROC Then
        '    sql &= " ,RIGHT('000' + CONVERT(VARCHAR, YEAR(a.OBSDATE)-1911), 3) + FORMAT(a.OBSDATE, '/MM/dd') OBSDATE" & vbCrLf  'edit，by:20181002
        'End If
        sql &= " ,CONVERT(VARCHAR, a.OBSDATE, 111) OBSDATE" & vbCrLf
        sql &= " ,a.OBYEARS" & vbCrLf
        sql &= " ,a.OBCOMMENT" & vbCrLf
        sql &= " ,a.OBTERMS" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        sql &= " ,kp.PLANNAME" & vbCrLf
        '處分期間 (PunishPeriod)
        sql &= " ,CASE WHEN ISNULL(a.OBYEARS,0) >0 then" & vbCrLf
        sql &= " concat(dbo.FN_CDATE1B(a.OBSDATE),'至 ',dbo.FN_CDATE1B(dateadd(YEAR,a.OBYEARS,a.OBSDATE)-1))" & vbCrLf
        sql &= " else '' end C_PunishPeriod" & vbCrLf

        ' (依107年需求,增加其它欄位內容) start
        sql &= " ,YEAR(d.SENTERDATE) C_YEAR" & vbCrLf
        sql &= " ,d.CLASSCNAME C_NAME" & vbCrLf
        sql &= " ,CASE WHEN d.STDATE IS NOT NULL AND d.FTDATE IS NOT NULL" & vbCrLf
        sql &= "   THEN CONVERT(VARCHAR, d.STDATE, 111) + ' ～ ' + CONVERT(VARCHAR, d.FTDATE, 111)" & vbCrLf
        sql &= "   ELSE NULL END  C_PERIOD" & vbCrLf

        sql &= " ,a.APPLYPRICE MY_APPLYPRICE" & vbCrLf
        sql &= " ,a.AUTHPRICE MY_AUTHPRICE" & vbCrLf
        sql &= " ,OBNUM ,OBFACT" & vbCrLf
        sql &= " ,CASE WHEN a.OBYEARS > 0 THEN CONVERT(VARCHAR, a.OBYEARS) + '年' ELSE NULL END MY_OBYEARS" & vbCrLf
        sql &= " ,CASE WHEN a.ISLAW1 IS NOT NULL AND a.ISLAW1='Y' THEN '是'" & vbCrLf
        sql &= "   WHEN a.ISLAW1 IS NOT NULL AND a.ISLAW1='N' THEN '否'" & vbCrLf
        sql &= "   ELSE NULL END MY_LAW1" & vbCrLf
        sql &= " ,CASE WHEN a.ISLAW2 IS NOT NULL AND a.ISLAW2='Y' THEN '是'" & vbCrLf
        sql &= "   WHEN a.ISLAW2 IS NOT NULL AND a.ISLAW2='N' THEN '否'" & vbCrLf
        sql &= "   ELSE NULL END MY_LAW2" & vbCrLf
        sql &= " ,a.TRANSFER MY_TRANSFER" & vbCrLf
        sql &= " ,concat(CONVERT(VARCHAR, a.JUDGEDATE, 111) , CASE WHEN a.JUDGENUM IS NOT NULL THEN ' (' + a.JUDGENUM + ')' ELSE '' END) MY_JUDGE_1" & vbCrLf
        sql &= " ,a.JUDGEFACT MY_JUDGE_2" & vbCrLf
        sql &= " ,a.TODO MY_JUDGE_3" & vbCrLf
        sql &= " ,a.NOTE MY_NOTE" & vbCrLf
        sql &= " ,format(a.MODIFYDATE,'yyyy/MM/dd HH:mm') MY_MODIFYDATE" & vbCrLf
        '===== end
        sql &= " FROM ORG_BLACKLIST a" & vbCrLf
        sql &= " JOIN ORG_ORGINFO b ON b.ComIDNO=a.ComIDNO" & vbCrLf
        sql &= " JOIN KEY_PLAN kp ON kp.TPLANID=a.TPLANID" & vbCrLf

        If RIDValue.Value <> "" Then sql &= " JOIN WAR1 ON WAR1.ORGID=b.ORGID" & vbCrLf
        sql &= " JOIN ID_DISTRICT c ON c.DISTID =a.DISTID" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO d ON d.OCID=a.OCID" & vbCrLf
        sql &= " WHERE a.AVAIL='Y'" & vbCrLf '(刪除註記)

        If ddlTPlanIDSch.SelectedIndex <> 0 AndAlso ddlTPlanIDSch.SelectedValue <> "" Then sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        If ComidValue.Text <> "" Then sql &= " AND b.COMIDNO=@COMIDNO" & vbCrLf
        If txtOrgName.Text <> "" Then sql &= " AND b.ORGNAME LIKE '%'+@ORGNAME+'%'" & vbCrLf
        If DistID.SelectedIndex <> 0 AndAlso DistID.SelectedValue <> "" Then sql &= " AND a.DISTID=@DISTID" & vbCrLf
        If Years.SelectedIndex <> 0 AndAlso Years.SelectedValue <> "" Then sql &= " AND DATEPART(YEAR, a.OBSDATE)=@YEAR" & vbCrLf

        sql &= " ORDER BY a.DISTID, OBSDATE "

        Dim parms As New Hashtable()
        If RIDValue.Value <> "" Then parms.Add("RID", RIDValue.Value)
        If ddlTPlanIDSch.SelectedIndex <> 0 AndAlso ddlTPlanIDSch.SelectedValue <> "" Then parms.Add("TPLANID", ddlTPlanIDSch.SelectedValue)
        If ComidValue.Text <> "" Then parms.Add("COMIDNO", ComidValue.Text)
        If txtOrgName.Text <> "" Then parms.Add("ORGNAME", txtOrgName.Text)
        If DistID.SelectedIndex <> 0 AndAlso DistID.SelectedValue <> "" Then parms.Add("DISTID", DistID.SelectedValue)
        If Years.SelectedIndex <> 0 AndAlso Years.SelectedValue <> "" Then parms.Add("YEAR", Years.SelectedValue)

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Call sUtl_PanelList(1) '搜尋
        msg.Text = "查無資料"
        tb_Sch.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            tb_Sch.Visible = True
        End If

        If bl_printMode Then
            '(edit,by:20180705)
            DataGrid2.DataSource = dt
            DataGrid2.DataBind()
        Else
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If
    End Sub

    '呼叫一筆資料
    Function Loaddata1(ByVal OBSN As String) As DataRow
        Dim rstDr As DataRow = Nothing
        If OBSN = "" Then Return rstDr

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.DISTID ,a.OBSN ,ISNULL(b.ORGNAME,'-') ORGNAME ,a.COMIDNO" & vbCrLf
        sql &= " ,CONVERT(varchar, a.OBSDATE, 111) OBSDATE" & vbCrLf
        sql &= " ,a.OBYEARS ,a.OBCOMMENT ,a.OBNUM" & vbCrLf
        '處分期間 (PunishPeriod)
        sql &= " ,CASE WHEN ISNULL(a.OBYEARS,0) >0 then" & vbCrLf
        sql &= " concat(dbo.FN_CDATE1B(OBSDATE),'至',dbo.FN_CDATE1B(dateadd(YEAR,OBYEARS,OBSDATE)-1))" & vbCrLf
        sql &= " else '' end C_PunishPeriod" & vbCrLf

        sql &= " ,c.NAME DISTNAME" & vbCrLf
        sql &= " ,ISNULL(oo.ORGNAME, '')  + '：' + ISNULL(aa.NAME,'') ACCNAME" & vbCrLf
        sql &= " ,a.OBTERMS ,a.TPLANID" & vbCrLf
        sql &= " ,kp.PLANNAME" & vbCrLf
        '(依107年需求,增加其它欄位內容) start
        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,CASE WHEN e.BUSID IS NOT NULL THEN '[' + e.BUSID + ']' + e.BUSNAME" & vbCrLf
        sql &= "  WHEN e.JOBID IS NOT NULL THEN '[' + e.JOBID + ']' + e.JOBNAME" & vbCrLf
        sql &= "  WHEN e.TRAINID IS NOT NULL THEN '[' + e.TRAINID + ']' + e.TRAINNAME" & vbCrLf
        sql &= "  END CKNAME" & vbCrLf
        sql &= " ,d.TMID TMID" & vbCrLf
        sql &= " ,d.CLASSCNAME CRNAME" & vbCrLf
        sql &= " ,a.APPLYPRICE" & vbCrLf
        sql &= " ,a.AUTHPRICE" & vbCrLf
        sql &= " ,a.OBFACT" & vbCrLf
        sql &= " ,a.ISLAW1" & vbCrLf
        sql &= " ,a.ISLAW2" & vbCrLf
        sql &= " ,a.TRANSFER" & vbCrLf
        sql &= " ,CONVERT(varchar, a.JUDGEDATE, 111) JUDGEDATE" & vbCrLf
        sql &= " ,a.JUDGENUM" & vbCrLf
        sql &= " ,a.JUDGEFACT" & vbCrLf
        sql &= " ,a.TODO" & vbCrLf
        sql &= " ,a.NOTE" & vbCrLf
        sql &= " ,format(a.ModifyDate,'yyyy/MM/dd HH:mm') ModifyDate" & vbCrLf

        sql &= " FROM ORG_BLACKLIST a" & vbCrLf
        sql &= " LEFT JOIN ORG_ORGINFO b ON b.COMIDNO =a.COMIDNO" & vbCrLf
        sql &= " LEFT JOIN KEY_PLAN kp ON kp.TPLANID =a.TPLANID" & vbCrLf
        sql &= " JOIN ID_DISTRICT c ON c.DISTID =a.DISTID" & vbCrLf
        sql &= " LEFT JOIN AUTH_ACCOUNT aa ON aa.ACCOUNT =a.MODIFYACCT" & vbCrLf
        sql &= " LEFT JOIN ORG_ORGINFO oo ON oo.ORGID =aa.ORGID" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSINFO d ON d.OCID =a.OCID" & vbCrLf   '20180709
        sql &= " LEFT JOIN KEY_TRAINTYPE e ON e.TMID =d.TMID" & vbCrLf   '20180709
        sql &= " WHERE a.AVAIL='Y' AND a.OBSN=@OBSN "
        Dim parms As Hashtable = New Hashtable From {{"OBSN", hid_OBSN.Value}}
        rstDr = DbAccess.GetOneRow(sql, objconn, parms)
        Return rstDr
    End Function

    ''' <summary>
    ''' 檢查廠商統一編號x處分起日x分署 資料是否已存在
    ''' </summary>
    ''' <param name="htS"></param>
    ''' <returns></returns>
    Public Function Check_Blacklist(ByRef htS As Hashtable) As String
        '-',ByVal ComIDNO As String, ByVal OBSDate As String
        Dim ComIDNO As String = TIMS.GetMyValue2(htS, "ComIDNO")
        Dim s_DISTID As String = TIMS.GetMyValue2(htS, "DISTID")
        Dim s_DISTNAME As String = TIMS.GetMyValue2(htS, "DISTNAME")
        Dim OBSDate As String = TIMS.GetMyValue2(htS, "OBSDate")

        Dim msg As String = ""
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT OBSN ,ComIDNO ,CONVERT(varchar, OBSDATE, 111) OBSDATE" & vbCrLf
        sql &= " FROM ORG_BLACKLIST" & vbCrLf
        sql &= " WHERE AVAIL='Y'" & vbCrLf
        sql &= " AND COMIDNO=@COMIDNO" & vbCrLf
        sql &= " AND DISTID=@DISTID" & vbCrLf
        'sql += " AND OBSDATE=@OBSDATE" & vbCrLf
        sql &= " AND CONVERT(VARCHAR, OBSDATE, 111)=@OBSDATE" & vbCrLf

        If hid_OBSN.Value <> "" Then sql &= " AND OBSN <> @OBSN "

        Dim dt As DataTable = Nothing
        'parms.Add("OBSDATE", TIMS.to_date(OBSDate))
        Dim parms As New Hashtable From {{"COMIDNO", ComIDNO}, {"DISTID", s_DISTID}, {"OBSDATE", OBSDate}}
        If hid_OBSN.Value <> "" Then parms.Add("OBSN", hid_OBSN.Value)

        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then msg = String.Format("此廠商統一編號 {0} 處分日期 ({1})，資料已存在!", OBSDate, s_DISTNAME) & vbCrLf
        Return msg
    End Function

    ''' <summary>檢核</summary>
    ''' <param name="msg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef msg As String) As Boolean
        Dim rst As Boolean = True
        Dim MyKey As String = ""
        Dim sql2 As String = ""
        Dim dr2 As DataRow = Nothing

        txt_ComIDNO.Text = TIMS.ChangeIDNO(UCase(TIMS.ClearSQM(txt_ComIDNO.Text)))
        If txt_ComIDNO.Text = "" Then '統編為必填資料
            msg += cst_altmsg6 & vbCrLf
            Return False
        End If

        '檢查廠商統一編號&處分起日&分署 資料是否已存在
        msg = ""
        'msg &= Check_Blacklist(Trim(Me.txt_ComIDNO.Text), Common.FormatDate(txt_OBSdate.Text))
        Dim myOBSdate As String = ""  'edit，by:20181001
        myOBSdate = txt_OBSdate.Text  'edit，by:20181001
        If flag_ROC Then myOBSdate = TIMS.Cdate18(myOBSdate)  'edit，by:20181001

        Dim v_ddl_DistID As String = TIMS.GetListValue(ddl_DistID)
        Dim t_ddl_DistID As String = TIMS.GetListText(ddl_DistID)
        Dim in_parms As New Hashtable From {
            {"ComIDNO", txt_ComIDNO.Text},
            {"DISTID", v_ddl_DistID},
            {"DISTNAME", t_ddl_DistID},
            {"OBSDate", Common.FormatDate(myOBSdate)}
        }
        msg &= Check_Blacklist(in_parms)  'edit，by:20181001

        sql2 = " SELECT COMIDNO FROM ORG_ORGINFO WHERE COMIDNO=@COMIDNO "
        Dim parms As Hashtable = New Hashtable()
        parms.Add("COMIDNO", txt_ComIDNO.Text)
        dr2 = DbAccess.GetOneRow(sql2, objconn, parms)
        If dr2 Is Nothing Then msg += cst_altmsg1 & vbCrLf '統編為必填資料，且必須再系統內有該單位

        '檢查處分緣由與處分年限。
        Dim v_ddlOBTERMS As String = TIMS.GetListValue(ddlOBTERMS)
        If v_ddlOBTERMS = "" Then msg += cst_altmsg2 & vbCrLf '處分緣由(條文)不可為空
        If v_ddlOBTERMS <> "" Then
            MyKey = TIMS.Get_OBTERMSName(TIMS.Get_OBTERM(), v_ddlOBTERMS)
            If MyKey = "" Then msg += cst_altmsg2b & vbCrLf

            Dim v_ddl_OBYears As String = TIMS.GetListValue(ddl_OBYears)
            Select Case v_ddlOBTERMS 'ddlOBTERMS.SelectedValue
                Case OBTERMS.cst_c38 'Cst_TPlanID28AppPlan
                    Select Case v_ddl_OBYears'ddl_OBYears.SelectedValue
                        Case "0", "1"
                        Case Else
                            msg += cst_altmsg38 & vbCrLf
                    End Select
                Case OBTERMS.cst_c39 'Cst_TPlanID28AppPlan
                    If v_ddl_OBYears <> "1" Then msg += cst_altmsg39 & vbCrLf

                Case OBTERMS.cst_c40 'Cst_TPlanID28AppPlan
                    Select Case v_ddl_OBYears
                        Case "1", "2", "3"
                        Case Else
                            msg += cst_altmsg40 & vbCrLf
                    End Select
                Case OBTERMS.cst_c42 'Cst_TPlanID28AppPlan
                    If v_ddl_OBYears <> "2" Then msg += cst_altmsg42 & vbCrLf

                Case OBTERMS.cst_c07 'TIMS.Cst_TPlanID68
                    Select Case v_ddl_OBYears
                        Case "1", "2"
                        Case Else
                            msg += cst_altmsg507 & vbCrLf
                    End Select
                Case OBTERMS.cst_c20 'TIMS.Cst_TPlanID68
                    If v_ddl_OBYears <> "1" Then msg += cst_altmsg520 & vbCrLf
                Case OBTERMS.cst_c21 'TIMS.Cst_TPlanID68
                    If v_ddl_OBYears <> "2" Then msg += cst_altmsg521 & vbCrLf
                Case Else
                    '未定義的規則。
            End Select
        End If

        If msg <> "" Then rst = False
        Return rst
    End Function

    '儲存
    Function iSaveData1() As Integer
        Dim iRst As Integer = 0

        '組【Insert】的SQL語法
        Dim iSql As String = ""
        iSql &= " INSERT INTO ORG_BLACKLIST(OBSN, ComIDNO, OBSDATE, OBYears, OBComment, Avail, ModifyAcct, ModifyDate"
        iSql &= " ,OBNum, Distid, OBTERMS, TPlanID"
        iSql &= " ,OCID, APPLYPRICE, AUTHPRICE, OBFACT, ISLAW1, ISLAW2, TRANSFER, JUDGEDATE, JUDGENUM, JUDGEFACT, TODO, NOTE)" & vbCrLf   '20180709
        iSql &= " VALUES(@OBSN, @COMIDNO, @OBSDATE, @OBYEARS, @OBCOMMENT, 'Y', @MODIFYACCT, GETDATE()"
        iSql &= " ,@OBNUM, @DISTID, @OBTERMS, @TPLANID"
        iSql &= " ,@OCID, @APPLYPRICE, @AUTHPRICE, @OBFACT, @ISLAW1, @ISLAW2, @TRANSFER, @JUDGEDATE, @JUDGENUM, @JUDGEFACT, @TODO, @NOTE)" & vbCrLf   '20180709
        'Dim iCmd As New SqlCommand(iSql, objconn)

        '組【Update】的SQL語法
        Dim uSql As String = ""
        uSql &= " UPDATE ORG_BLACKLIST" & vbCrLf
        uSql &= " SET COMIDNO=@COMIDNO" & vbCrLf
        uSql &= " ,OBSDATE=@OBSDATE" & vbCrLf
        uSql &= " ,OBYEARS=@OBYEARS" & vbCrLf
        uSql &= " ,OBCOMMENT=@OBCOMMENT" & vbCrLf
        uSql &= " ,AVAIL='Y'" & vbCrLf
        uSql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        uSql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        uSql &= " ,OBNUM=@OBNUM" & vbCrLf
        uSql &= " ,DISTID=@DISTID" & vbCrLf
        uSql &= " ,OBTERMS=@OBTERMS" & vbCrLf
        uSql &= " ,OCID=@OCID" & vbCrLf   '20180709
        uSql &= " ,APPLYPRICE=@APPLYPRICE" & vbCrLf   '20180709
        uSql &= " ,AUTHPRICE=@AUTHPRICE" & vbCrLf   '20180709
        uSql &= " ,OBFACT=@OBFACT" & vbCrLf   '20180709
        uSql &= " ,ISLAW1=@ISLAW1" & vbCrLf   '20180709
        uSql &= " ,ISLAW2=@ISLAW2" & vbCrLf   '20180709
        uSql &= " ,TRANSFER=@TRANSFER" & vbCrLf   '20180709
        uSql &= " ,JUDGEDATE=@JUDGEDATE" & vbCrLf   '20180709
        uSql &= " ,JUDGENUM=@JUDGENUM" & vbCrLf   '20180709
        uSql &= " ,JUDGEFACT=@JUDGEFACT" & vbCrLf   '20180709
        uSql &= " ,TODO=@TODO" & vbCrLf   '20180709
        uSql &= " ,NOTE=@NOTE" & vbCrLf   '20180709
        uSql &= " WHERE OBSN=@OBSN" & vbCrLf
        'Dim uCmd As New SqlCommand(uSql, objconn)
        Call TIMS.OpenDbConn(objconn)

        Dim v_ddl_OBYears As String = TIMS.GetListValue(ddl_OBYears)
        Dim v_rbl_IsLaw1 As String = TIMS.GetListValue(rbl_IsLaw1).ToUpper
        Dim v_rbl_IsLaw2 As String = TIMS.GetListValue(rbl_IsLaw2).ToUpper
        Dim v_ddlOBTERMS As String = TIMS.GetListValue(ddlOBTERMS)

        If hid_OBSN.Value = "" Then
            '新增
            Dim iOBSN As Integer = DbAccess.GetNewId(objconn, "ORG_BLACKLIST_OBSN_SEQ,ORG_BLACKLIST,OBSN")
            Dim parms As Hashtable = New Hashtable()
            parms.Add("OBSN", iOBSN)
            parms.Add("COMIDNO", Me.txt_ComIDNO.Text)
            'parms.Add("OBSDATE", CDate(txt_OBSdate.Text.Trim))
            parms.Add("OBSDATE", If(flag_ROC, CDate(TIMS.Cdate18(txt_OBSdate.Text)), CDate(txt_OBSdate.Text)))  'edit，by:20181001

            parms.Add("OBYEARS", Val(v_ddl_OBYears))
            parms.Add("OBCOMMENT", txt_OBComment.Text)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("OBNUM", txt_No.Text)
            parms.Add("DISTID", ddl_DistID.SelectedValue)
            parms.Add("OBTERMS", v_ddlOBTERMS)
            parms.Add("TPLANID", ddlTPlanID.SelectedValue)
            parms.Add("OCID", If(OCIDValue1.Value.Trim.Length > 0, Val(OCIDValue1.Value), Convert.DBNull))
            parms.Add("APPLYPRICE", If(txt_ApplyPrice.Text.Trim.Length > 0, Val(txt_ApplyPrice.Text), Convert.DBNull))
            parms.Add("AUTHPRICE", If(txt_AuthPrice.Text.Trim.Length > 0, Val(txt_AuthPrice.Text), Convert.DBNull))
            parms.Add("OBFACT", If(txt_OBFact.Text.Trim.Length > 0, txt_OBFact.Text, Convert.DBNull))
            parms.Add("ISLAW1", If(v_rbl_IsLaw1 <> "", v_rbl_IsLaw1, Convert.DBNull))
            parms.Add("ISLAW2", If(v_rbl_IsLaw2 <> "", v_rbl_IsLaw2, Convert.DBNull))
            parms.Add("TRANSFER", If(txt_Transfer.Text.Trim.Length > 0, txt_Transfer.Text, Convert.DBNull))
            'edit，by:20181001
            parms.Add("JUDGEDATE", If(txt_JudgeDate.Text <> "", If(flag_ROC, CDate(TIMS.Cdate18(txt_JudgeDate.Text)), CDate(txt_JudgeDate.Text)), Convert.DBNull))
            parms.Add("JUDGENUM", If(txt_JudgeNum.Text.Trim.Length > 0, txt_JudgeNum.Text, Convert.DBNull))
            parms.Add("JUDGEFACT", If(txt_JudgeFact.Text.Trim.Length > 0, txt_JudgeFact.Text, Convert.DBNull))
            parms.Add("TODO", If(txt_Tudo.Text.Trim.Length > 0, txt_Tudo.Text, Convert.DBNull))
            parms.Add("NOTE", If(txt_Note.Text.Trim.Length > 0, txt_Note.Text, Convert.DBNull))
            iRst += DbAccess.ExecuteNonQuery(iSql, parms)
        Else
            '修改
            Dim parms As Hashtable = New Hashtable()
            parms.Add("COMIDNO", Me.txt_ComIDNO.Text)
            'parms.Add("OBSDATE", CDate(txt_OBSdate.Text.Trim))
            parms.Add("OBSDATE", If(flag_ROC, CDate(TIMS.Cdate18(txt_OBSdate.Text)), CDate(txt_OBSdate.Text)))  'edit，by:20181001

            parms.Add("OBYEARS", Val(v_ddl_OBYears))
            parms.Add("OBCOMMENT", txt_OBComment.Text)
            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            parms.Add("OBNUM", txt_No.Text)
            parms.Add("DISTID", ddl_DistID.SelectedValue)
            parms.Add("OBTERMS", v_ddlOBTERMS)
            parms.Add("OCID", If(OCIDValue1.Value.Trim.Length > 0, Val(OCIDValue1.Value), Convert.DBNull))  '20180709、20180710
            parms.Add("APPLYPRICE", If(txt_ApplyPrice.Text.Trim.Length > 0, Val(txt_ApplyPrice.Text), Convert.DBNull))  '20180709、20180710
            parms.Add("AUTHPRICE", If(txt_AuthPrice.Text.Trim.Length > 0, Val(txt_AuthPrice.Text), Convert.DBNull))  '20180709、20180710
            parms.Add("OBFACT", If(txt_OBFact.Text.Trim.Length > 0, txt_OBFact.Text, Convert.DBNull))  '20180709、20180710
            parms.Add("ISLAW1", If(v_rbl_IsLaw1 <> "", v_rbl_IsLaw1, Convert.DBNull))
            parms.Add("ISLAW2", If(v_rbl_IsLaw2 <> "", v_rbl_IsLaw2, Convert.DBNull))
            parms.Add("TRANSFER", If(txt_Transfer.Text.Trim.Length > 0, txt_Transfer.Text, Convert.DBNull))  '20180709、20180710
            'edit，by:20181001
            parms.Add("JUDGEDATE", If(txt_JudgeDate.Text <> "", If(flag_ROC, CDate(TIMS.Cdate18(txt_JudgeDate.Text)), CDate(txt_JudgeDate.Text)), Convert.DBNull))
            parms.Add("JUDGENUM", If(txt_JudgeNum.Text.Trim.Length > 0, txt_JudgeNum.Text, Convert.DBNull))
            parms.Add("JUDGEFACT", If(txt_JudgeFact.Text.Trim.Length > 0, txt_JudgeFact.Text, Convert.DBNull))
            parms.Add("TODO", If(txt_Tudo.Text.Trim.Length > 0, txt_Tudo.Text, Convert.DBNull))
            parms.Add("NOTE", If(txt_Note.Text.Trim.Length > 0, txt_Note.Text, Convert.DBNull))
            parms.Add("OBSN", hid_OBSN.Value)
            iRst += DbAccess.ExecuteNonQuery(uSql, parms)
        End If
        Return iRst
    End Function

    '儲存。
    Private Sub btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click
        '(直接在AuthBasePage處理,不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim sErrMsg As String = ""
        Call CheckData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        Dim iRst As Integer = iSaveData1()
        If iRst = 0 Then
            Common.MessageBox(Me, "無資料異動!!")
            Exit Sub
        End If
        If hid_OBSN.Value = "" Then Common.MessageBox(Me, "儲存成功")
        If hid_OBSN.Value <> "" Then Common.MessageBox(Me, "修改成功")
        Call sSearch1()
    End Sub

    '換行
    Function CutRowChar(ByVal str1 As String) As String
        Dim tmpStr As String = ""
        Const Cst_LenNum As Integer = 40 '切斷長度
        Dim num As Int16 = 0
        Dim last As Int16 = 0
        Dim j As Int16 = 0
        num = Len(str1) \ Cst_LenNum '取得總行數
        last = Len(str1) Mod Cst_LenNum '最後剩餘字行數
        j = 1 '計算位置
        For i As Int16 = 0 To num
            If i = num Then
                tmpStr += Mid(str1, j, last)
            Else
                tmpStr += Mid(str1, j, Cst_LenNum) + "<br/>"
                j += Cst_LenNum
            End If
        Next
        Return tmpStr
    End Function

    '修改
    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        '(直接在AuthBasePage處理,不用個別檢查Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sCmdArg As String = e.CommandArgument
        hid_OBSN.Value = TIMS.GetMyValue(sCmdArg, "OBSN")
        If hid_OBSN.Value = "" Then Exit Sub
        hid_OBSN.Value = TIMS.ClearSQM(hid_OBSN.Value)

        Dim dr As DataRow = Loaddata1(hid_OBSN.Value)
        If dr Is Nothing Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        Select Case e.CommandName
            Case "edit" '修改
                Call sUtl_PanelList(2) '修改
                Call SHOW_DATA1(dr)

            Case "view" '檢視
                Call sUtl_PanelList(3) '檢視
                Call SHOW_DATA2(dr)

            Case "del" '刪除(非資料刪除只做註記,使用者無法看到資料)
                Dim u_Parms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"OBSN", hid_OBSN.Value}}
                Dim uSql As String = ""
                uSql &= " UPDATE ORG_BLACKLIST "
                uSql &= " SET AVAIL='N' ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() "
                uSql &= " WHERE OBSN=@OBSN "
                DbAccess.ExecuteNonQuery(uSql, objconn, u_Parms)
                Common.MessageBox(Me, "刪除成功")
                Call sSearch1()

            Case Else
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header, ListItemType.Footer
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn_view As LinkButton = e.Item.FindControl("lbtView")
                Dim btn_edit As LinkButton = e.Item.FindControl("lbtEdit")
                Dim btn_del As LinkButton = e.Item.FindControl("lbtDel")
                Dim labOBTERMS As Label = e.Item.FindControl("labOBTERMS")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                If DataGrid1.AllowPaging Then
                    If Len(e.Item.Cells(Cst_事由位置).Text) > 19 Then e.Item.Cells(Cst_事由位置).Text = Mid(e.Item.Cells(Cst_事由位置).Text, 1, 18) + "..."  '事由內容超過19字後顯示...
                End If
                btn_view.Visible = False
                btn_edit.Visible = True
                btn_del.Visible = True
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OBSN", Convert.ToString(drv("OBSN")))
                btn_view.CommandArgument = ""
                btn_edit.CommandArgument = sCmdArg 'drv("OBSN")
                If sm.UserInfo.DistID <> Convert.ToString(drv("DistID")) Then
                    '如果不是原轄區中心(分署)  只提供檢視
                    btn_view.Visible = True
                    btn_edit.Visible = False '不可修改
                    btn_del.Visible = False '不可刪除
                    btn_view.CommandArgument = sCmdArg 'drv("OBSN")
                    btn_edit.CommandArgument = ""
                End If
                labOBTERMS.Text = If(Convert.ToString(drv("OBTERMS")) <> "", TIMS.Get_OBTERMSName(TIMS.Get_OBTERM(), drv("OBTERMS")), "")
                btn_del.CommandArgument = sCmdArg 'drv("OBSN")
                btn_del.Attributes("onclick") = "return confirm('確定要刪除第" & e.Item.Cells(0).Text & "筆紀錄?');"
                If flag_ROC Then e.Item.Cells(6).Text = TIMS.Cdate17(drv("OBSDATE"))  'edit，by:20181001
        End Select
    End Sub

    '(edit,by:20180705)
    Protected Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labOBTERMS As Label = e.Item.FindControl("labOBTERMS")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                labOBTERMS.Text = If(Convert.ToString(drv("OBTERMS")) <> "", TIMS.Get_OBTERMSName(TIMS.Get_OBTERM(), drv("OBTERMS")), "")
        End Select
    End Sub

    '新增。
    Private Sub btnAdds_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdds.Click
        hid_OBSN.Value = ""
        Call sUtl_PanelList(2) '新增/修改
        Call ClearEdit1()  '新增清除。

        ddlTPlanID.Enabled = False '新增鎖定。
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)
        '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
        ddl_DistID.Enabled = False
        Common.SetListItem(ddl_DistID, sm.UserInfo.DistID)

        '20100208 按新增時代查詢之 統一編號
        ComidValue.Text = TIMS.ClearSQM(ComidValue.Text)
        txt_ComIDNO.Text = ComidValue.Text
        txt_PunishPeriod.Text = cst_PunishPeriod_autotxt
    End Sub

    ''' <summary>
    ''' 新增清除。離開清除
    ''' </summary>
    Sub ClearEdit1()
        txt_ComIDNO.Text = ""
        txt_No.Text = ""
        txt_OBSdate.Text = ""
        Common.SetListItem(ddl_OBYears, "1")

        ddlTPlanID.Enabled = False '新增/修改鎖定。
        Common.SetListItem(ddlTPlanID, sm.UserInfo.TPlanID)
        txt_OBComment.Text = ""
        ddlOBTERMS.SelectedIndex = -1
        hid_OBSN.Value = "" 'Nothing
    End Sub

    ''' <summary>匯出 / DataGrid1 </summary>
    Sub sExport1()
        'Const Cst_xlsFileName As String="訓練單位處分資料匯出.xls"
        Dim oDataGrid1 As DataGrid = DataGrid1 '(DivOutputDoc)
        oDataGrid1.AllowPaging = False
        oDataGrid1.EnableViewState = False  '把ViewState給關了
        bl_printMode = True  'edit，by:20181001
        Call sSearch1()

        Dim sFileName1 As String = "訓練單位處分資料匯出"
        Dim strSTYLE As String = ""
        ''套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        '(配合匯出作業,調整表格欄位) ======
        oDataGrid1.Visible = False
        oDataGrid1.AllowPaging = False
        oDataGrid1.Columns(Cst_功能欄位).Visible = False
        oDataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        'Div1.RenderControl(objHtmlTextWriter)
        DivOutputDoc.Visible = True    '(edit,by:20180705)
        DivOutputDoc.RenderControl(objHtmlTextWriter)   '(edit,by:20180705)
        DivOutputDoc.Visible = False   '(edit,by:20180705)

        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        '回復表格原先狀態 ======
        bl_printMode = False  'edit，by:20181001
        Call sSearch1()       'edit，by:20181001
        oDataGrid1.AllowPaging = True
        oDataGrid1.Columns(Cst_功能欄位).Visible = True

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    ''' <summary>匯出 / DataGrid1</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        '(直接在AuthBasePage處理,不用個別檢查Session) If TIMS.ChkSession(Me) Then Exit Sub
        Call sExport1()
    End Sub

End Class