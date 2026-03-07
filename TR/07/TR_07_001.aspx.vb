Partial Class TR_07_001
    Inherits AuthBasePage

    '年度執行成效
    Const cst_msg_t1 As String = "* 表示該班有以下情形之一：(1)開訓人數比率未達90%、(2)離退訓率超過10%(含)、(3)不開班"
    Const cst_Edit1 As String = "Edit1"
    Const cst_sPeopleFMT1 As String = "{0} 人"
    Const cst_sRateFMT1 As String = "{0} %"
    Dim dtNORID As DataTable = Nothing
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        '1.自辦 '2.委外
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        dtNORID = TIMS.Get_NORIDdt(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        Call TIMS.OpenDbConn(objconn)

        '帶入查詢參數
        If Not IsPostBack Then Call cCreate1()

    End Sub

    Sub cCreate1()
        labmsg_t1.Text = cst_msg_t1
        msg1.Text = ""
        DataGrid1Table.Visible = False
        Utl_display(0)

        Dim dtD As DataTable = TIMS.Get_DISTIDdt(objconn)
        Dim sff As String = "distid='000'"
        If dtD.Select(sff).Length > 0 Then dtD.Select(sff)(0).Delete()
        cblDistid = TIMS.Get_DistID(cblDistid, dtD)
        cblDistid.Items.Insert(0, New ListItem("全部", 0))
        Common.SetListItem(cblDistid, sm.UserInfo.DistID)
        cblDistid.Enabled = If(sm.UserInfo.DistID <> "000", False, True)

        '選擇全部轄區
        cblDistid.Attributes("onclick") = "SelectAll('cblDistid','DistHidden');"

        yearlist = TIMS.GetSyear(yearlist)
        Common.SetListItem(yearlist, sm.UserInfo.Years)

        '取出鍵詞-不開班理由代碼
        Call TIMS.Get_NotOpenReason(NORID, objconn)
    End Sub

    Function Utl_GetData1(ByRef sCmdArg As String) As DataTable

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_cblDistid As String = TIMS.CombiSQM2IN(TIMS.GetCblValue(cblDistid))
        STDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate1.Text))
        STDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate2.Text))
        FTDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(FTDate1.Text))
        FTDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(FTDate2.Text))

        Dim mParms As New Hashtable
        mParms.Add("TPLANID", sm.UserInfo.TPlanID)
        If v_yearlist <> "" Then mParms.Add("YEARS", v_yearlist)
        'If v_cblDistid <> "" Then mParms.Add("DISTID", v_cblDistid)
        If STDate1.Text <> "" Then mParms.Add("STDate1", STDate1.Text)
        If STDate2.Text <> "" Then mParms.Add("STDate2", STDate2.Text)
        If FTDate1.Text <> "" Then mParms.Add("FTDate1", FTDate1.Text)
        If FTDate2.Text <> "" Then mParms.Add("FTDate2", FTDate2.Text)
        If sCmdArg <> "" Then
            'Dim s_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
            Dim s_PLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
            Dim s_COMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
            Dim s_SEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")
            mParms.Clear()
            mParms.Add("TPLANID", sm.UserInfo.TPlanID)
            mParms.Add("PLANID", s_PLANID) ' STDate1.Text)
            mParms.Add("COMIDNO", s_COMIDNO) ' STDate1.Text)
            mParms.Add("SEQNO", s_SEQNO) 'STDate1.Text)
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 As (" & vbCrLf
        sql &= " SELECT vp.YEARS" & vbCrLf
        sql &= " ,vp.DISTID, vp.DISTNAME" & vbCrLf
        sql &= " ,vp.TPLANID, vp.PLANNAME" & vbCrLf
        sql &= " ,CC.OCID, cc.RID" & vbCrLf
        sql &= " ,cc.PLANID, cc.COMIDNO, cc.SEQNO" & vbCrLf
        sql &= " ,cc.CLASSCNAME, cc.CYCLTYPE" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME, cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,cc.TMID, tt.TRAINNAME" & vbCrLf
        sql &= " ,pp.APPLIEDDATE" & vbCrLf
        sql &= " ,cc.STDATE" & vbCrLf
        sql &= " ,cc.FTDATE" & vbCrLf
        sql &= " ,CC.SENTERDATE" & vbCrLf
        sql &= " ,CC.FENTERDATE" & vbCrLf
        sql &= " ,CC.CHECKINDATE" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        sql &= " ,cc.THOURS" & vbCrLf
        sql &= " ,cc.NORID" & vbCrLf
        sql &= " ,cc.OtherReason" & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO PP WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN vp WITH(NOLOCK) On vp.PLANID = PP.PLANID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO OO WITH(NOLOCK) ON OO.COMIDNO = PP.COMIDNO" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO CC WITH(NOLOCK) ON CC.PLANID = PP.PLANID AND CC.COMIDNO = PP.COMIDNO AND CC.SEQNO = PP.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ID_CLASS IC WITH(NOLOCK) ON IC.CLSID = CC.CLSID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND pp.ISAPPRPAPER = 'Y'" & vbCrLf
        'sql &= " and vp.YEARS='2021' and vp.TPLANID='06'" & vbCrLf
        sql &= " and vp.TPLANID=@TPLANID" & vbCrLf
        If sCmdArg <> "" Then
            sql &= " and pp.PLANID=@PLANID" & vbCrLf
            sql &= " and pp.COMIDNO=@COMIDNO" & vbCrLf
            sql &= " and pp.SEQNO=@SEQNO" & vbCrLf
        End If
        If sCmdArg = "" Then
            If v_yearlist <> "" Then sql &= " and vp.YEARS=@YEARS" & vbCrLf 'mParms.Add("YEARS", v_yearlist)
            If v_cblDistid <> "" Then sql &= String.Format(" and vp.DISTID IN ({0})", v_cblDistid) & vbCrLf 'mParms.Add("DISTID", v_cblDistid)
            If STDate1.Text <> "" Then sql &= " and cc.STDate >= @STDate1" & vbCrLf 'mParms.Add("STDate1", STDate1.Text)
            If STDate2.Text <> "" Then sql &= " and cc.STDate <= @STDate2" & vbCrLf 'mParms.Add("STDate2", STDate2.Text)
            If FTDate1.Text <> "" Then sql &= " and cc.FTDate >= @FTDate1" & vbCrLf 'mParms.Add("FTDate1", FTDate1.Text)
            If FTDate2.Text <> "" Then sql &= " and cc.FTDate <= @FTDate2" & vbCrLf 'mParms.Add("FTDate2", FTDate2.Text)
        End If
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " SELECT cc.OCID" & vbCrLf
        'sql &= " /* 報名人數 */" & vbCrLf
        sql &= " ,COUNT(1) STUDETNUM" & vbCrLf '報名人數
        'sql &= " /* 甄試人數 */" & vbCrLf
        sql &= " ,COUNT(CASE WHEN b.TOTALRESULT>=0 THEN 1 END) STUDETNUM2" & vbCrLf '甄試人數
        'sql &= " /* 錄取人數 '2、錄取人數：為【錄訓作業】功能中的「正取」人數。 */" & vbCrLf
        sql &= " ,COUNT(CASE WHEN c.SELRESULTID ='01' THEN 1 END) STUDETNUM3" & vbCrLf '錄取人數
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTYPE b WITH(NOLOCK) on b.OCID1=cc.OCID" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTEMP a WITH(NOLOCK) on a.SETID=b.SETID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SELRESULT c WITH(NOLOCK) on c.setid=b.setid and c.enterdate=b.enterdate and c.sernum=b.sernum and c.ocid=b.ocid1" & vbCrLf
        sql &= " GROUP BY cc.OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " SELECT cc.OCID" & vbCrLf
        'sql &= " /* 開訓人數 */" & vbCrLf
        sql &= " ,COUNT(1) SNum1" & vbCrLf
        'sql &= " /* 結訓人數 */" & vbCrLf
        sql &= " ,COUNT(case when cs.StudStatus=5 then 1 end) ESNum1" & vbCrLf
        'sql &= " /* 離退訓人數 */" & vbCrLf
        sql &= " ,COUNT(case when cs.StudStatus IN (2,3) then 1 end) JSNum1" & vbCrLf
        'sql &= " /* 開訓男性人數 */" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='M' THEN 1 END) CNT1M" & vbCrLf
        'sql &= " /* 開訓女性人數 */" & vbCrLf
        sql &= " ,COUNT(CASE WHEN SS.SEX ='F' THEN 1 END) CNT1F" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) ON cs.OCID=cc.OCID" & vbCrLf
        sql &= " JOIN dbo.STUD_STUDENTINFO ss WITH(NOLOCK) ON ss.sid = cs.sid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cs.MAKESOCID IS NULL" & vbCrLf
        sql &= " GROUP BY cc.OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT cc.OCID" & vbCrLf
        sql &= " ,CC.PLANID,CC.COMIDNO,CC.SEQNO" & vbCrLf
        sql &= " ,cc.YEARS" & vbCrLf
        sql &= " ,cc.PLANNAME" & vbCrLf
        sql &= " ,CC.DISTNAME" & vbCrLf
        sql &= " ,CC.CLASSCNAME2" & vbCrLf
        sql &= " ,CC.TRAINNAME" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        sql &= " ,cc.THOURS" & vbCrLf
        sql &= " ,cc.NORID" & vbCrLf
        sql &= " ,cc.OtherReason" & vbCrLf
        'NORID_N
        sql &= " ,convert(nvarchar(300),NULL) NORID_N" & vbCrLf
        sql &= " ,S1.STUDETNUM" & vbCrLf
        sql &= " ,S1.STUDETNUM2" & vbCrLf
        sql &= " ,S1.STUDETNUM3" & vbCrLf 'CP_04_003_add
        sql &= " ,S2.SNum1" & vbCrLf
        sql &= " ,S2.ESNum1" & vbCrLf
        sql &= " ,S2.JSNum1" & vbCrLf 'CP_04_003_add
        sql &= " ,S2.CNT1M" & vbCrLf
        sql &= " ,S2.CNT1F" & vbCrLf
        sql &= " ,q4.AVERAGE" & vbCrLf 'SD_11_006_R
        'AVERAGE_N
        sql &= " ,convert(nvarchar(30),NULL) AVERAGE_N" & vbCrLf

        '3、錄訓率 錄取率(%)=錄取人數(正取)/甄試人數。 ACCEPtance rate
        sql &= " ,CASE WHEN ISNULL(s1.STUDETNUM2,0)>0 THEN concat(round(cast(isnull(s1.STUDETNUM3,0) as float)/s1.STUDETNUM2*100,2),'%') END ACCEPRATE" & vbCrLf
        '4、開訓人數比率(%)=開訓人數/招生人數。 Number of trainees ratio
        sql &= " ,CASE WHEN cc.TNUM>0 AND S2.SNum1>=0 THEN concat(round(cast(isnull(S2.SNum1,0) as float)/cc.TNUM*100,2),'%') END TRAINRATE" & vbCrLf
        '5、離退訓率(%)=離退訓人數/開訓人數。 Retirement rate
        sql &= " ,CASE WHEN ISNULL(S2.SNum1,0)>0 THEN concat(round(cast(isnull(S2.JSNum1,0) as float)/S2.SNum1*100,2),'%') END RTIRERATE" & vbCrLf

        sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf
        sql &= " ,convert(nvarchar(300),NULL) D20KNAME " & vbCrLf '/*政策性課程類型 D20KNAME*/" & vbCrLf
        'sql &= " ,convert(nvarchar(MAX),NULL) NDATA1" & vbCrLf '（暫時空白資料）
        sql &= " ,pn.Review90" & vbCrLf ' 開訓人數比率未達90%之檢討改善
        sql &= " ,pn.Review10" & vbCrLf '離退訓率超過10%之檢討改善	
        sql &= " ,pn.ReviewNG" & vbCrLf '不開班之檢討措施
        sql &= " ,pn.ReviewOth" & vbCrLf '其他執行說明及檢討改善(非必填)

        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT JOIN WS1 s1 on s1.OCID=cc.OCID" & vbCrLf
        sql &= " LEFT JOIN WS2 s2 on s2.OCID=cc.OCID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_QUESTIONARY4 q4 ON q4.OCID= cc.OCID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=cc.PLANID and dd.COMIDNO=cc.COMIDNO and dd.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_ANNUAL pn on pn.PLANID=cc.PLANID and pn.COMIDNO=cc.COMIDNO and pn.SEQNO=cc.SEQNO" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, mParms)

        Return dt
    End Function

    Sub Utl_display(ByRef i_type As Integer)
        DetailTable.Visible = False
        SearchTable.Visible = True
        If i_type = 0 Then Return

        DetailTable.Visible = True
        SearchTable.Visible = False
        If i_type = 1 Then Return
    End Sub

    Sub Utl_ClearData1()
        Hid_PLANID.Value = "" 'Convert.ToString(dr1("PLANID"))
        Hid_COMIDNO.Value = "" 'Convert.ToString(dr1("COMIDNO"))
        Hid_SEQNO.Value = "" 'Convert.ToString(dr1("SEQNO"))
        Hid_OCID1.Value = "" 'Convert.ToString(dr1("OCID"))

        labCLASSCNAME2.Text = "" ' Convert.ToString(dr1("CLASSCNAME2"))
        labTRAINNAME.Text = "" 'Convert.ToString(dr1("TRAINNAME"))
        labD20KNAME.Text = "" 'Convert.ToString(dr1("D20KNAME"))
        labSTDATE.Text = "" 'Convert.ToString(dr1("STDATE"))
        labFTDATE.Text = "" 'Convert.ToString(dr1("FTDATE"))
        labTNUM.Text = "" ' Convert.ToString(dr1("TNUM"))
        labSTUDETNUM.Text = "" 'Convert.ToString(dr1("STUDETNUM"))
        labSTUDETNUM2.Text = "" 'Convert.ToString(dr1("STUDETNUM2"))
        labSTUDETNUM3.Text = "" 'Convert.ToString(dr1("STUDETNUM3"))
        labSNum1.Text = "" 'Convert.ToString(dr1("SNum1"))

        labACCEPRATE.Text = "" 'Convert.ToString(dr1("ACCEPRATE"))
        labTRAINRATE.Text = "" 'Convert.ToString(dr1("TRAINRATE"))
        TB_Review90.Text = "" 'Convert.ToString(dr1("Review90"))

        labESNum1.Text = "" 'Convert.ToString(dr1("ESNum1"))
        labJSNum1.Text = "" 'Convert.ToString(dr1("JSNum1"))
        labRTIRERATE.Text = "" 'Convert.ToString(dr1("RTIRERATE"))
        TB_Review10.Text = "" 'Convert.ToString(dr1("Review10"))
        labAVERAGE.Text = "" 'Convert.ToString(dr1("AVERAGE"))

        TIMS.SetCblValue(NORID, "")
        NORIDValue.Value = "" 'TIMS.GetCblValue(NORID)
        OtherReason.Text = "" 'dr1("OtherReason").ToString

        TB_ReviewNG.Text = "" 'Convert.ToString(dr1("ReviewNG"))
        TB_ReviewOth.Text = "" 'Convert.ToString(dr1("ReviewOth"))
    End Sub

    Sub Utl_LoadData1(ByRef sCmdArg As String)
        If sCmdArg = "" Then Return

        Dim s_OCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim s_PLANID As String = TIMS.GetMyValue(sCmdArg, "PLANID")
        Dim s_COMIDNO As String = TIMS.GetMyValue(sCmdArg, "COMIDNO")
        Dim s_SEQNO As String = TIMS.GetMyValue(sCmdArg, "SEQNO")

        Dim dt As DataTable = Utl_GetData1(sCmdArg)
        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dt.Rows(0)
        If dr1 Is Nothing Then Return

        UPDATE_DataRow2(dr1)

        Utl_ClearData1()

        Hid_PLANID.Value = Convert.ToString(dr1("PLANID"))
        Hid_COMIDNO.Value = Convert.ToString(dr1("COMIDNO"))
        Hid_SEQNO.Value = Convert.ToString(dr1("SEQNO"))
        Hid_OCID1.Value = Convert.ToString(dr1("OCID"))

        labCLASSCNAME2.Text = Convert.ToString(dr1("CLASSCNAME2"))
        labTRAINNAME.Text = Convert.ToString(dr1("TRAINNAME"))
        labD20KNAME.Text = Convert.ToString(dr1("D20KNAME"))
        labSTDATE.Text = Convert.ToString(dr1("STDATE"))
        labFTDATE.Text = Convert.ToString(dr1("FTDATE"))

        labTNUM.Text = String.Format(cst_sPeopleFMT1, dr1("TNUM"))
        labSTUDETNUM.Text = String.Format(cst_sPeopleFMT1, dr1("STUDETNUM"))
        labSTUDETNUM2.Text = String.Format(cst_sPeopleFMT1, dr1("STUDETNUM2"))
        labSTUDETNUM3.Text = String.Format(cst_sPeopleFMT1, dr1("STUDETNUM3"))
        labSNum1.Text = String.Format(cst_sPeopleFMT1, dr1("SNum1"))

        labACCEPRATE.Text = Convert.ToString(dr1("ACCEPRATE"))
        labTRAINRATE.Text = Convert.ToString(dr1("TRAINRATE"))
        hid_TRAINRATE.Value = If(labTRAINRATE.Text <> "", Replace(labTRAINRATE.Text, "%", ""), "")
        TB_Review90.Text = Convert.ToString(dr1("Review90"))

        labESNum1.Text = String.Format(cst_sPeopleFMT1, dr1("ESNum1"))
        labJSNum1.Text = String.Format(cst_sPeopleFMT1, dr1("JSNum1"))
        labRTIRERATE.Text = Convert.ToString(dr1("RTIRERATE"))
        hid_RTIRERATE.Value = If(labRTIRERATE.Text <> "", Replace(labRTIRERATE.Text, "%", ""), "")
        TB_Review10.Text = Convert.ToString(dr1("Review10"))
        labAVERAGE.Text = String.Format(cst_sRateFMT1, dr1("AVERAGE"))

        TIMS.SetCblValue(NORID, Convert.ToString(dr1("NORID")))
        NORIDValue.Value = TIMS.GetCblValue(NORID)
        OtherReason.Text = dr1("OtherReason").ToString
        NORID.Enabled = False
        OtherReason.Enabled = False
        TB_ReviewNG.Text = Convert.ToString(dr1("ReviewNG"))
        TB_ReviewOth.Text = Convert.ToString(dr1("ReviewOth"))

        '系統請增加檢核：
        '當開訓人數比率(%)<90%(不含)，開訓人數比率未達90%之檢討改善 必須填寫。
        '當離退訓率(%)>=10%(含) ，離退訓率超過10%之檢討改善必須填寫。
        '當此班為不開班， 不開班之檢討改善必須填寫。

        TB_Review90.Enabled = If(hid_TRAINRATE.Value <> "" AndAlso Val(hid_TRAINRATE.Value) < 90, True, False)
        TB_Review10.Enabled = If(hid_RTIRERATE.Value <> "" AndAlso Val(hid_RTIRERATE.Value) >= 10, True, False)
        TB_ReviewNG.Enabled = If(NORIDValue.Value <> "" OrElse OtherReason.Text <> "", True, False)
        TIMS.Tooltip(TB_Review90, "開訓人數比率未達90%之檢討改善", True)
        TIMS.Tooltip(TB_Review10, "離退訓率超過10%之檢討改善", True)
        TIMS.Tooltip(TB_ReviewNG, "不開班之檢討改善", True)

        Dim class_td_TB_Review90 As String = If(TB_Review90.Enabled, "bluecol_need", "bluecol")
        Dim class_td_TB_Review10 As String = If(TB_Review10.Enabled, "bluecol_need", "bluecol")
        Dim class_td_TB_ReviewNG As String = If(TB_ReviewNG.Enabled, "bluecol_need", "bluecol")
        td_TB_Review90.Attributes.Add("class", class_td_TB_Review90)
        td_TB_Review10.Attributes.Add("class", class_td_TB_Review10)
        td_TB_ReviewNG.Attributes.Add("class", class_td_TB_ReviewNG)

    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim dt As DataTable = Utl_GetData1("")

        msg1.Text = "查無資料!"
        DataGrid1Table.Visible = False
        If dt.Rows.Count = 0 Then Return

        msg1.Text = ""
        DataGrid1Table.Visible = True

        'PageControler1.Sort = "RIDValue,ClassID,CyclType"
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Sub sch_check(ByRef Errmsg1 As String)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If v_yearlist = "" Then Errmsg1 &= "查詢年度不可為空" & vbCrLf
    End Sub

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Dim Errmsg1 As String = ""
        sch_check(Errmsg1)
        If Errmsg1 <> "" Then
            Common.MessageBox(Me, Errmsg1)
            Return
        End If

        sSearch1()
    End Sub

    Sub UPDATE_DataRow2(ByRef dr As DataRow)
        Dim s_D20KNAME As String = ""
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME1").ToString())
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME2").ToString())
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME3").ToString())
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME4").ToString())
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME5").ToString())
        TIMS.ADD_PAR2(s_D20KNAME, dr("D20KNAME6").ToString())
        dr("D20KNAME") = s_D20KNAME
        'AVERAGE_N
        Dim s_AVERAGE_N As String = If(Convert.ToString(dr("AVERAGE")) <> "", String.Format("{0}%", dr("AVERAGE")), "")
        dr("AVERAGE_N") = If(s_AVERAGE_N <> "", s_AVERAGE_N, Convert.DBNull)
        Dim sff As String = ""
        'Dim dtNORID As DataTable = Nothing
        If dtNORID Is Nothing Then Return
        If Convert.ToString(dr("NORID")) = "" Then Return
        Dim s_NORID_N As String = ""
        Dim sar_NORID As String() = Convert.ToString(dr("NORID")).Split(",")
        For Each sVal1 As String In sar_NORID
            sff = String.Format("NORID='{0}'", sVal1)
            Dim sNORNAME As String = If(dtNORID.Select(sff).Length > 0, dtNORID.Select(sff)(0)("NORNAME"), "")
            If sNORNAME <> "" Then
                If s_NORID_N <> "" Then s_NORID_N &= ","
                s_NORID_N &= sNORNAME
            End If
        Next
        dr("NORID_N") = If(s_NORID_N <> "", s_NORID_N, "")
        'OtherReason
        If s_NORID_N <> "" AndAlso Convert.ToString(dr("OtherReason")) <> "" Then
            If s_NORID_N <> "" Then s_NORID_N &= ","
            s_NORID_N &= Convert.ToString(dr("OtherReason"))
        End If

        'Dim s_CAPOTHER As String = ""
        'If (dr("CAPOTHER1").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("1.{0}", dr("CAPOTHER1").ToString()))
        'If (dr("CAPOTHER2").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("2.{0}", dr("CAPOTHER2").ToString()))
        'If (dr("CAPOTHER3").ToString() <> "") Then TIMS.ADD_PAR2(s_CAPOTHER, String.Format("3.{0}", dr("CAPOTHER3").ToString()))
        'dr("CAPOTHER") = s_CAPOTHER
    End Sub

    Sub Utl_Export1()
        Dim dtObj As DataTable = Utl_GetData1("")

        If dtObj.Rows.Count = 0 Then Return
        Dim drT1 As DataRow = dtObj.Rows(0)

        'Dim sFileName1 As String = "年度執行成效"
        '序號,
        Dim s_title1 As String = ""
        s_title1 &= "班別名稱,訓練職類,政策性課程類型,開訓日期,結訓日期,預訓人數,報名人數,甄試人數,錄訓人數"
        s_title1 &= ",錄訓率,開訓人數,開訓人數比率,結訓人數,離退訓人數,離退訓率,滿意度"
        s_title1 &= ",開訓人數比率未達90%之檢討改善,離退訓率超過10%之檢討改善,不開班原因,不開班之檢討措施,其他"
        Dim s_data1 As String = ""
        s_data1 &= "CLASSCNAME2,TRAINNAME,D20KNAME,STDATE,FTDATE,TNUM,STUDETNUM,STUDETNUM2,STUDETNUM3"
        s_data1 &= ",ACCEPRATE,SNum1,TRAINRATE,ESNum1,JSNum1,RTIRERATE,AVERAGE_N"
        s_data1 &= ",REVIEW90,REVIEW10,NORID_N,REVIEWNG,REVIEWOTH"

        Dim s_tit2 As String = "" '(純數字)
        s_tit2 = "TNUM,STUDETNUM,STUDETNUM2,STUDETNUM3,SNum1,ESNum1,JSNum1"
        's_tit3 &= ",ACCEPRATE,SNum1,TRAINRATE,ESNum1,JSNum1,RTIRERATE,AVERAGE"
        Dim s_tit3 As String = "" '(文字)/(日期)
        s_tit3 = "STDATE,FTDATE"

        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")
        Dim sary_tit2() As String = s_tit2.Split(",") '(純數字)
        Dim sary_tit3() As String = s_tit3.Split(",") '(文字)/(日期)
        Dim iColSpanCount As Integer = As_title1.Length + 1

        Const cst_ColFmt1 As String = "<td>{0}</td>"
        Const cst_ColFmt2 As String = "<td class=""noDecFormat"">{0}</td>" '(純數字)
        Const cst_ColFmt3 As String = "<td class=""DateFormat"">{0}</td>" '(文字)/(日期)

        Dim vROC_YERS As String = Val(drT1("YEARS")) - 1911
        Dim vORGNAME As String = Convert.ToString(drT1("DISTNAME"))
        Dim vPLANNAME As String = Convert.ToString(drT1("PLANNAME"))
        '勞動部勞動力發展署OO分署 OO年 在職進修訓練 年度執行成效   
        '匯出表頭名稱
        Dim sFileName1 As String = String.Format("export_{0}", TIMS.GetDateNo2())
        Dim s_TitleName As String = String.Format("勞動部{0} {1}年 {2} 年度執行成效", vORGNAME, vROC_YERS, vPLANNAME)

        '套CSS值
        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        'strSTYLE &= ("td{mso-number-format:""\@"";}") '(文字)
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}") '(數字)
        strSTYLE &= (".DateFormat{mso-number-format:""\@"";}") '(文字)/(日期)
        strSTYLE &= ("</style>")

        Dim ExportStr As String '建立輸出文字
        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '表頭及查詢條件列
        ExportStr = String.Format("<tr><td align='center' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_TitleName) & vbCrLf
        sbHTML.Append(ExportStr)

        ExportStr = "<tr>"
        ExportStr &= String.Format(cst_ColFmt1, "序號")
        For Each s_T1 As String In As_title1
            ExportStr &= String.Format(cst_ColFmt1, s_T1) '& vbTab
        Next
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        'Dim iStudCnt As Integer = 0 '申請班數及 預訓人數(x訓練人數)
        Dim iSNum1Cnt As Integer = 0 '開訓人數SNum1
        Dim i_rows As Integer = 0
        For Each oDr1 As DataRow In dtObj.Rows
            UPDATE_DataRow2(oDr1)
            i_rows += 1
            'iStudCnt += Val(If(Convert.ToString(oDr1("TNUM")) <> "", oDr1("TNUM"), 0))
            iSNum1Cnt += Val(If(Convert.ToString(oDr1("SNum1")) <> "", oDr1("SNum1"), 0))
            ExportStr = "<tr>"
            ExportStr &= String.Format(cst_ColFmt1, i_rows) '序號
            For Each s_D1 As String In As_data1
                Dim flag_find2 As Boolean = TIMS.FindValue1(sary_tit2, s_D1) '(純數字)搜尋
                Dim flag_find3 As Boolean = TIMS.FindValue1(sary_tit3, s_D1) '(日期)搜尋
                Dim s_FMT2 As String = If(flag_find2, cst_ColFmt2, If(flag_find3, cst_ColFmt3, cst_ColFmt1)) '(純數字)/'(日期)/一般 切換
                ExportStr &= String.Format(s_FMT2, oDr1(s_D1)) '& vbTab
            Next
            ExportStr &= "</tr>"
            sbHTML.Append(ExportStr)
        Next

        'Dim s_ClassNUM1 As String = String.Format("申請班數 合計 {0} 班", dtObj.Rows.Count)
        Dim s_ClassNUM1 As String = String.Format("合計 {0} 班", dtObj.Rows.Count)
        ExportStr = String.Format("<tr><td align='left' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_ClassNUM1) & vbCrLf
        sbHTML.Append(ExportStr)
        'Dim s_StudTNUM1 As String = String.Format("訓練人數 訓練 {0} 人", iSNum1Cnt) ' iStudCnt)
        Dim s_SNum1NUM As String = String.Format("訓練 {0} 人", iSNum1Cnt) ' iStudCnt)
        ExportStr = String.Format("<tr><td align='left' colspan='{0}'>{1}</td></tr>", iColSpanCount, s_SNum1NUM) & vbCrLf
        sbHTML.Append(ExportStr)

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        Utl_Export1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Const Cst_index As Integer = 0
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtEdit1 As LinkButton = e.Item.FindControl("lbtEdit1") '修改'btnEdit 'ID="lbtEdit1" runat="server" Text="修改" CommandName="Edit1"
                '*表示該班有以下情形之一：(1)開訓人數比率未達90%、(2)離退訓率超過10%(含)、(3)不開班
                Dim LabStar1 As Label = e.Item.FindControl("LabStar1")
                Dim LabSeqno As Label = e.Item.FindControl("LabSeqno") '序號
                LabSeqno.Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "PLANID", drv("PLANID"))
                TIMS.SetMyValue(sCmdArg, "COMIDNO", drv("COMIDNO"))
                TIMS.SetMyValue(sCmdArg, "SEQNO", drv("SEQNO"))
                lbtEdit1.CommandArgument = sCmdArg

                Dim v_TRAINRATE As String = If(Convert.ToString(drv("TRAINRATE")) <> "", Replace(Convert.ToString(drv("TRAINRATE")), "%", ""), "")
                Dim v_RTIRERATE As String = If(Convert.ToString(drv("RTIRERATE")) <> "", Replace(Convert.ToString(drv("RTIRERATE")), "%", ""), "")
                Dim v_NORID As String = Convert.ToString(drv("NORID"))
                Dim v_OtherReason As String = Convert.ToString(drv("OtherReason"))

                Dim s_title1 As String = ""
                Dim f_NG1 As Boolean = If(v_TRAINRATE <> "" AndAlso Val(v_TRAINRATE) < 90, True, False)
                If f_NG1 Then s_title1 &= "開訓人數比率未達90%之檢討改善 必須填寫" & vbCrLf
                Dim f_NG2 As Boolean = If(v_RTIRERATE <> "" AndAlso Val(v_RTIRERATE) >= 10, True, False)
                If f_NG2 Then s_title1 &= "離退訓率超過10%之檢討改善必須填寫" & vbCrLf
                Dim f_NG3 As Boolean = If(v_NORID <> "" OrElse v_OtherReason <> "", True, False)
                If f_NG3 Then s_title1 &= "不開班之檢討改善必須填寫" & vbCrLf
                '*表示該班有以下情形之一：(1)開訓人數比率未達90%、(2)離退訓率超過10%(含)、(3)不開班
                LabStar1.Visible = If((f_NG1 OrElse f_NG2 OrElse f_NG3), True, False)
                TIMS.Tooltip(LabStar1, s_title1, True)
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Select Case e.CommandName
            Case cst_Edit1
                Dim sCmdArg As String = e.CommandArgument
                Utl_LoadData1(sCmdArg)
                Utl_display(1)
        End Select
    End Sub

    Sub CheckData1(ByRef Errmsg1 As String)
        '系統請增加檢核：
        '當開訓人數比率(%)<90%(不含)，開訓人數比率未達90%之檢討改善 必須填寫。
        '當離退訓率(%)>=10%(含) ，離退訓率超過10%之檢討改善必須填寫。
        '當此班為不開班， 不開班之檢討改善必須填寫。
        TB_Review90.Text = TIMS.ClearSQM2(TB_Review90.Text)
        TB_Review10.Text = TIMS.ClearSQM2(TB_Review10.Text)
        TB_ReviewNG.Text = TIMS.ClearSQM2(TB_ReviewNG.Text)

        Dim f_NG1 As Boolean = If(hid_TRAINRATE.Value <> "" AndAlso Val(hid_TRAINRATE.Value) < 90, True, False)
        If f_NG1 AndAlso TB_Review90.Text = "" Then Errmsg1 &= "開訓人數比率未達90%之檢討改善 必須填寫" & vbCrLf

        Dim f_NG2 As Boolean = If(hid_RTIRERATE.Value <> "" AndAlso Val(hid_RTIRERATE.Value) >= 10, True, False)
        If f_NG2 AndAlso TB_Review10.Text = "" Then Errmsg1 &= "離退訓率超過10%之檢討改善必須填寫" & vbCrLf

        Dim f_NG3 As Boolean = If(NORIDValue.Value <> "" OrElse OtherReason.Text <> "", True, False)
        If f_NG3 AndAlso TB_ReviewNG.Text = "" Then Errmsg1 &= "不開班之檢討改善必須填寫" & vbCrLf
    End Sub

    Sub Utl_SaveData1()
        Dim s_PLANID As String = TIMS.ClearSQM(Hid_PLANID.Value)
        Dim s_COMIDNO As String = TIMS.ClearSQM(Hid_COMIDNO.Value)
        Dim s_SEQNO As String = TIMS.ClearSQM(Hid_SEQNO.Value)
        Dim s_OCID1 As String = TIMS.ClearSQM(Hid_OCID1.Value)
        If s_PLANID = "" Then Return
        If s_COMIDNO = "" Then Return
        If s_SEQNO = "" Then Return
        If s_OCID1 = "" Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(s_OCID1, objconn)
        If drCC Is Nothing Then Return
        Dim drPP As DataRow = TIMS.GetPCSDate(s_PLANID, s_COMIDNO, s_SEQNO, objconn)
        If drPP Is Nothing Then Return
        If drCC("OCID").ToString() <> drPP("OCID").ToString() Then Return

        Dim s_parms As New Hashtable
        s_parms.Clear()
        s_parms.Add("PLANID", s_PLANID)
        s_parms.Add("COMIDNO", s_COMIDNO)
        s_parms.Add("SEQNO", s_SEQNO)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 'x' FROM PLAN_ANNUAL" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND PLANID=@PLANID" & vbCrLf
        sql &= " AND COMIDNO=@COMIDNO" & vbCrLf
        sql &= " AND SEQNO=@SEQNO" & vbCrLf

        Dim dt1 As DataTable
        dt1 = DbAccess.GetDataTable(sql, objconn, s_parms)
        Dim flag_is_Add As Boolean = If(dt1.Rows.Count = 0, True, False)

        Dim e_parms As New Hashtable
        e_parms.Clear()
        'where/insert
        e_parms.Add("PLANID", s_PLANID)
        e_parms.Add("COMIDNO", s_COMIDNO)
        e_parms.Add("SEQNO", s_SEQNO)
        e_parms.Add("Review90", If(TB_Review90.Text <> "", TB_Review90.Text, Convert.DBNull))
        e_parms.Add("Review10", If(TB_Review10.Text <> "", TB_Review10.Text, Convert.DBNull))
        e_parms.Add("ReviewNG", If(TB_ReviewNG.Text <> "", TB_ReviewNG.Text, Convert.DBNull))
        e_parms.Add("ReviewOth", If(TB_ReviewOth.Text <> "", TB_ReviewOth.Text, Convert.DBNull))
        e_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        Dim e_sql As String = ""
        If flag_is_Add Then
            e_sql = "" & vbCrLf
            e_sql &= " INSERT INTO PLAN_ANNUAL(" & vbCrLf
            e_sql &= " PLANID,COMIDNO,SEQNO" & vbCrLf
            e_sql &= " ,Review90,Review10" & vbCrLf
            e_sql &= " ,ReviewNG,ReviewOth" & vbCrLf
            e_sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
            e_sql &= " ) VALUES (" & vbCrLf
            e_sql &= " @PLANID,@COMIDNO,@SEQNO" & vbCrLf
            e_sql &= " ,@Review90,@Review10" & vbCrLf
            e_sql &= " ,@ReviewNG,@ReviewOth" & vbCrLf
            e_sql &= " ,@MODIFYACCT,GETDATE()" & vbCrLf
            e_sql &= " )" & vbCrLf
        Else
            e_sql = "" & vbCrLf
            e_sql &= " UPDATE PLAN_ANNUAL" & vbCrLf
            e_sql &= " SET Review90=@Review90" & vbCrLf
            e_sql &= " ,Review10=@Review10" & vbCrLf
            e_sql &= " ,ReviewNG=@ReviewNG" & vbCrLf
            e_sql &= " ,ReviewOth=@ReviewOth" & vbCrLf
            e_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            e_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            e_sql &= " WHERE 1=1" & vbCrLf
            e_sql &= " AND PLANID=@PLANID" & vbCrLf
            e_sql &= " AND COMIDNO=@COMIDNO" & vbCrLf
            e_sql &= " AND SEQNO=@SEQNO" & vbCrLf
        End If
        DbAccess.ExecuteNonQuery(e_sql, objconn, e_parms)

        sm.LastResultMessage = TIMS.cst_SAVEOKMsg3
        Return
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        'Call ReLoad_SB4IDx()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Utl_SaveData1()
        Utl_display(0)
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        Utl_display(0)
    End Sub

End Class

