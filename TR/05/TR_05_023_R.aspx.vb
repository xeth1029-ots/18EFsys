Public Class TR_05_023_R
    Inherits AuthBasePage

    'Dim strTitle1 As String = "身心障礙者參加融合式職業訓練人數統計表"

    'FROM ADP_FDMOI 
    '(key_plan) select '+TPlanID+':'+planname  x
    'FROM key_plan where TPLANID IN ('02','17','21','26','34','37','46','47','49','58','06','28','61')
    'order by 1
    '計畫限定
    'Const cst_TPlanidS1 As String = "'02','17','21','26','34','37','46','47','49','58','06','28','61'"
    Const cst_TPlanidS1 As String = "'06','28'"

    '02:自辦職前訓練                                                         
    '06:在職進修訓練                                                         
    '17:補助地方政府訓練                                                      
    '21:原住民專班訓練                                                       
    '26:外籍與大陸配偶職業訓練                                                 
    '28:產業人才投資方案                                                      
    '34:推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)                      
    '37:委外職前訓練                                                         
    '46:補助辦理保母職業訓練                                                  
    '47:補助辦理照顧服務員職業訓練                                             
    '49:莫拉克風災區專班職前訓練(委外)                                          
    '58:補助辦理托育人員職業訓練  
    '61:推動原住民團體辦理原住民地區失業者職業訓練

    '1.增加「轄區別」該項篩選條件；選項分別為北分署、桃分署(含青年分署)、中分署、雲嘉南分署、高屏澎東分署共5分署，以上扣除泰山職訓中心
    '2.統計表格式變更，除原計算開訓人數外，增加結訓人數、就業人數等統計(格式如附)
    '3.統計表有上方請帶出目前資訊已與衛福部103年○月(月份系統自行帶)之資料核對等文字，文字確定為「資料來源：已核對衛福部提供103年1月份身心障礙者資料」
    '4.如篩選條件為計劃表後，產出表件僅羅列下列計畫名稱    
    '　(1)自辦職前：自辦職前訓練
    '　(2)委託或補助職前：補助地方政府訓練、原住民專班訓練、外籍與大陸配偶職業訓練、推動事業單位辦理職前培訓計畫、委外職前訓練、補助辦理托育人員職業訓練(本計畫101年以前請勾稽保母職業訓練之計畫數據)、補助辦理照顧服務員職業訓練、莫拉克風災專班職前訓練(委外)
    '　(3)在職訓練：自辦在職訓練、產業人才投資方案
    '5.本統計表係以TIMS系統內學員名單比對衛福部資料後回代之數據

    ''' <summary>註解文字回傳</summary>
    ''' <returns></returns>
    Function Get_Memo1() As String
        Dim rst As String = ""
        '資料來源：已核對衛服部提供103年12月份身心障礙者資料
        '資料來源：已核對衛福部提供yyy年MM月份身心障礙者資料，其中新制障礙類別已轉換為舊制
        Dim dt As DataTable
        Dim str_memo1 As String = "資料來源：已核對衛福部提供yyy年MM月份身心障礙者資料，其中新制障礙類別已轉換為舊制"
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT DATEPART(YEAR, MAX(modifydate))-1911 yyy " & vbCrLf
        sql &= " ,DATEPART(MONTH, max(modifydate)) MM " & vbCrLf
        sql &= " FROM dbo.ADP_FDMOI WITH(NOLOCK)" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Dim dr As DataRow = dt.Rows(0)
        str_memo1 = str_memo1.Replace("yyy", Convert.ToString(dr("yyy")))
        str_memo1 = str_memo1.Replace("MM", Convert.ToString(dr("MM")))
        rst = str_memo1
        Return rst
    End Function

    Dim objconn As SqlConnection
    'Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then CreateItem1()
    End Sub

    Sub CreateItem1()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        Syear = TIMS.GetSyear(Syear) '年度
        Common.SetListItem(Syear, sm.UserInfo.Years) ' Now.Year)

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))
        If Not DistID.Items.FindByValue("000") Is Nothing Then DistID.Items.Remove(DistID.Items.FindByValue("000"))
        If Not DistID.Items.FindByValue("002") Is Nothing Then DistID.Items.Remove(DistID.Items.FindByValue("002"))

        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"  '選擇全部轄區
        Export1.Attributes("onclick") = "return search();"
        'Button1.Attributes("onclick") = "return search();"
    End Sub

    Protected Sub Export1_Click(sender As Object, e As EventArgs) Handles Export1.Click
        STDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate1.Text))
        STDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate2.Text))
        FTDate1.Text = TIMS.Cdate3(TIMS.ClearSQM(FTDate1.Text))
        FTDate2.Text = TIMS.Cdate3(TIMS.ClearSQM(FTDate2.Text))

        Dim dt As DataTable
        Dim v_StatsMode As String = TIMS.GetListValue(StatsMode)
        '1:區域別 '2:計畫別 '3:區域計畫別 
        Select Case v_StatsMode 'StatsMode.SelectedValue
            Case "1"
                dt = LoadData1()
                Call ExpReport1(dt, 1)
            Case "2"
                dt = LoadData2()
                Call ExpReport1(dt, 2)
            Case "3"
                dt = LoadData3()
                Call ExpReport1(dt, 3)
            Case Else
                Common.MessageBox(Me, "未選擇統計方式!!")
                Exit Sub
        End Select
    End Sub

    '轄區SQL
    Function LoadData1() As DataTable
        Dim rst As DataTable
        '報表要用的轄區參數,1:為起始位置
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        Dim v_Syear As String = TIMS.GetListValue(Syear)

        Dim strSchcc As String = ""
        strSchcc = ""
        strSchcc &= " AND c.TPlanID IN (" & cst_TPlanidS1 & ") " & vbCrLf '計畫限定。
        If v_Syear <> "" Then strSchcc &= " AND c.years = '" & v_Syear & "' " & vbCrLf
        'If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate1.Text)), TIMS.to_date(STDate1.Text)) & vbCrLf  'edit，by:20181022
        'If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate2.Text)), TIMS.to_date(STDate2.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate1.Text)), TIMS.to_date(FTDate1.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate2.Text)), TIMS.to_date(FTDate2.Text)) & vbCrLf  'edit，by:20181022
        If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf
        If DistID1 <> "" Then strSchcc &= " AND c.DistID IN (" & DistID1.Replace("\'", "'") & ") " & vbCrLf '轉換sql查詢使用

        Dim sWC1 As String = ""
        sWC1 = ""
        sWC1 &= " WITH WC1 AS ( " & vbCrLf
        sWC1 &= " SELECT c.OCID,c.STDATE,c.FTDATE,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " FROM dbo.VIEW2 c " & vbCrLf
        sWC1 &= " WHERE 1=1 " & vbCrLf
        sWC1 &= strSchcc
        sWC1 &= " )" & vbCrLf
        sWC1 &= " ,WC2 AS ( " & vbCrLf
        sWC1 &= " SELECT s.SOCID,s.OCID,s.SEX,s.STUDSTATUS,s.IDNO,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " ,s.WorkSuppIdent " & vbCrLf '在職者 
        sWC1 &= " ,CASE WHEN c.TPLANID='28' THEN dbo.FN_GET_STUDCNT14B(s.STUDSTATUS,s.REJECTTDATE1,s.REJECTTDATE2,c.STDATE) ELSE 1 END STUDCNT" & vbCrLf '開訓人數 
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.V_STUDENTINFO s on c.ocid = s.ocid " & vbCrLf
        sWC1 &= " ) " & vbCrLf
        sWC1 &= " ,WC3 AS ( " & vbCrLf
        sWC1 &= " SELECT s.IDNO,c.OCID,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf '身障報名人數
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.VIEW_ENTERTYPE12 s on s.OCID1=c.OCID" & vbCrLf
        sWC1 &= " JOIN dbo.V_ADPFDMOI_1 a on a.IDNO=s.IDNO" & vbCrLf
        sWC1 &= " ) " & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= sWC1

        sql &= " SELECT d.DISTNAME " & vbCrLf
        sql &= "    ,ISNULL(c.clsNum,0) clsNum " & vbCrLf '班數
        sql &= "    ,ISNULL(c3.entNum,0) entNum" & vbCrLf '身障報名人數
        sql &= "    ,ISNULL(s.A1,0) A1 " & vbCrLf
        sql &= "    ,ISNULL(s.A2,0) A2 " & vbCrLf
        sql &= "    ,ISNULL(s.A,0) A " & vbCrLf '開訓人數

        sql &= "    ,ISNULL(s1.B1,0) B1 " & vbCrLf
        sql &= "    ,ISNULL(s1.B2,0) B2 " & vbCrLf
        sql &= "    ,ISNULL(s1.B,0) B " & vbCrLf '開訓人數(舊制 身障 男女)

        sql &= "    ,ISNULL(s2.D1,0) D1 " & vbCrLf
        sql &= "    ,ISNULL(s2.D2,0) D2 " & vbCrLf
        sql &= "    ,ISNULL(s2.D3,0) D3 " & vbCrLf
        sql &= "    ,ISNULL(s2.D4,0) D4 " & vbCrLf
        sql &= "    ,ISNULL(s2.D,0) D " & vbCrLf '開訓人數(舊制 身障 等級)
        sql &= "    ,ISNULL(s3.E1,0) E1 " & vbCrLf
        sql &= "    ,ISNULL(s3.E2,0) E2 " & vbCrLf
        sql &= "    ,ISNULL(s3.E3,0) E3 " & vbCrLf
        sql &= "    ,ISNULL(s3.E4,0) E4 " & vbCrLf
        sql &= "    ,ISNULL(s3.E5,0) E5 " & vbCrLf
        sql &= "    ,ISNULL(s3.E6,0) E6 " & vbCrLf
        sql &= "    ,ISNULL(s3.E7,0) E7 " & vbCrLf
        sql &= "    ,ISNULL(s3.E8,0) E8 " & vbCrLf
        sql &= "    ,ISNULL(s3.E9,0) E9 " & vbCrLf
        sql &= "    ,ISNULL(s3.E10,0) E10 " & vbCrLf
        sql &= "    ,ISNULL(s3.E11,0) E11 " & vbCrLf
        sql &= "    ,ISNULL(s3.E12,0) E12 " & vbCrLf
        sql &= "    ,ISNULL(s3.E13,0) E13 " & vbCrLf
        sql &= "    ,ISNULL(s3.E14,0) E14 " & vbCrLf
        sql &= "    ,ISNULL(s3.E15,0) E15 " & vbCrLf
        sql &= "    ,ISNULL(s3.E16,0) E16 " & vbCrLf
        sql &= "    ,ISNULL(s3.E17,0) E17 " & vbCrLf
        sql &= "    ,ISNULL(s3.E18,0) E18 " & vbCrLf
        sql &= "    ,ISNULL(s3.E99,0) E99 " & vbCrLf '開訓人數(舊制 身障 類別)
        sql &= "    ,ISNULL(s3.E,0) E " & vbCrLf
        sql &= "    ,ISNULL(s3.F1,0) F1 " & vbCrLf
        sql &= "    ,ISNULL(s3.F2,0) F2 " & vbCrLf
        sql &= "    ,ISNULL(s3.F3,0) F3 " & vbCrLf
        sql &= "    ,ISNULL(s3.F4,0) F4 " & vbCrLf
        sql &= "    ,ISNULL(s3.F5,0) F5 " & vbCrLf
        sql &= "    ,ISNULL(s3.F6,0) F6 " & vbCrLf
        sql &= "    ,ISNULL(s3.F7,0) F7 " & vbCrLf
        sql &= "    ,ISNULL(s3.F8,0) F8 " & vbCrLf
        sql &= "    ,ISNULL(s3.F99,0) F99 " & vbCrLf
        sql &= "    ,ISNULL(s3.F,0) F " & vbCrLf
        sql &= "    ,ISNULL(s.Aa1,0) AA1 " & vbCrLf
        sql &= "    ,ISNULL(s.Aa2,0) AA2 " & vbCrLf
        sql &= "    ,ISNULL(s.Aa,0) AA " & vbCrLf
        sql &= "    ,ISNULL(s1.Ba1,0) BA1 " & vbCrLf
        sql &= "    ,ISNULL(s1.Ba2,0) BA2 " & vbCrLf
        sql &= "    ,ISNULL(s1.Ba,0) BA " & vbCrLf
        sql &= "    ,ISNULL(s2.Da1,0) DA1 " & vbCrLf
        sql &= "    ,ISNULL(s2.Da2,0) DA2 " & vbCrLf
        sql &= "    ,ISNULL(s2.Da3,0) DA3 " & vbCrLf
        sql &= "    ,ISNULL(s2.Da4,0) DA4 " & vbCrLf
        sql &= "    ,ISNULL(s2.Da,0) DA " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea1,0) EA1 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea2,0) EA2 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea3,0) EA3 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea4,0) EA4 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea5,0) EA5 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea6,0) EA6 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea7,0) EA7 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea8,0) EA8 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea9,0) EA9 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea10,0) EA10 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea11,0) EA11 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea12,0) EA12 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea13,0) EA13 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea14,0) EA14 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea15,0) EA15 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea16,0) EA16 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea17,0) EA17 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea18,0) EA18 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea99,0) EA99 " & vbCrLf
        sql &= "    ,ISNULL(s3.Ea,0) EA " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa1,0) FA1 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa2,0) FA2 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa3,0) FA3 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa4,0) FA4 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa5,0) FA5 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa6,0) FA6 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa7,0) FA7 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa8,0) FA8 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa99,0) FA99 " & vbCrLf
        sql &= "    ,ISNULL(s3.Fa,0) FA " & vbCrLf

        'sql &= "    ,ISNULL(s.Ab1,0) Ab1 " & vbCrLf
        'sql &= "    ,ISNULL(s.Ab2,0) Ab2 " & vbCrLf
        'sql &= "    ,ISNULL(s.Ab,0) Ab " & vbCrLf
        'sql &= "    ,ISNULL(s1.Bb1,0) Bb1 " & vbCrLf
        'sql &= "    ,ISNULL(s1.Bb2,0) Bb2 " & vbCrLf
        'sql &= "    ,ISNULL(s1.Bb,0) Bb " & vbCrLf
        'sql &= "    ,ISNULL(s2.Db1,0) Db1 " & vbCrLf
        'sql &= "    ,ISNULL(s2.Db2,0) Db2 " & vbCrLf
        'sql &= "    ,ISNULL(s2.Db3,0) Db3 " & vbCrLf
        'sql &= "    ,ISNULL(s2.Db4,0) Db4 " & vbCrLf
        'sql &= "    ,ISNULL(s2.Db,0) Db " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb1,0) Eb1 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb2,0) Eb2 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb3,0) Eb3 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb4,0) Eb4 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb5,0) Eb5 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb6,0) Eb6 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb7,0) Eb7 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb8,0) Eb8 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb9,0) Eb9 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb10,0) Eb10 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb11,0) Eb11 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb12,0) Eb12 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb13,0) Eb13 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb14,0) Eb14 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb15,0) Eb15 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb16,0) Eb16 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb17,0) Eb17 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb18,0) Eb18 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb99,0) Eb99 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Eb,0) Eb " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb1,0) Fb1 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb2,0) Fb2 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb3,0) Fb3 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb4,0) Fb4 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb5,0) Fb5 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb6,0) Fb6 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb7,0) Fb7 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb8,0) Fb8 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb99,0) Fb99 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Fb,0) Fb " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gm1,0) Gm1 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gf1,0) Gf1 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gt1,0) Gt1 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gm2,0) Gm2 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gf2,0) Gf2 " & vbCrLf
        'sql &= "    ,ISNULL(s3.Gt2,0) Gt2 " & vbCrLf
        'sql &= "    ,ISNULL(s1.B3,0) B3 " & vbCrLf
        'sql &= "    ,ISNULL(s1.B4,0) B4 " & vbCrLf
        'sql &= "    ,ISNULL(s1.B5,0) B5 " & vbCrLf '在職開訓人數(舊制 身障 男女)
        'sql &= "    ,ISNULL(s1.BA3,0) BA3 " & vbCrLf
        'sql &= "    ,ISNULL(s1.BA4,0) BA4 " & vbCrLf
        'sql &= "    ,ISNULL(s1.BA5,0) BA5 " & vbCrLf '在職結訓人數(舊制 身障 男女)

        '排除 署(職訓局)
        '排除 泰山職業訓練中心
        sql &= " FROM (SELECT * FROM dbo.V_DISTRICT v WHERE v.DISTID NOT IN ('000','002') ) d " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.DISTID" & vbCrLf
        sql &= " 	,COUNT(1) clsNum " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.DISTID " & vbCrLf
        sql &= " ) c on c.distid =d.distid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.DISTID" & vbCrLf
        sql &= " 	,COUNT(1) entNum " & vbCrLf
        sql &= " 	FROM WC3 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.DISTID " & vbCrLf
        sql &= " ) c3 on c3.distid =d.distid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.DISTID " & vbCrLf
        sql &= " 	,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) A1 " & vbCrLf '開訓人數男
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) A2 " & vbCrLf '開訓人數女
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) A " & vbCrLf '開訓人數
        '結訓男/'結訓女/'結訓男女
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) AA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT END) AA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) AA " & vbCrLf
        '就業男         
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) AB1 " & vbCrLf
        ''就業女         
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) AB2 " & vbCrLf
        ''就業男女        
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) AB " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql &= " 	GROUP BY s.distid " & vbCrLf
        sql &= " ) s ON s.distid = d.distid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.distid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) B1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) B2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) B " & vbCrLf
        '結訓男/'結訓女/'結訓男女
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) BA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT END) BA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) BA " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) BB1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) BB2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) BB " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B3 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B4 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B5 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA3 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA4 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA5 " & vbCrLf
        sql &= "    FROM WC1 c " & vbCrLf
        sql &= "    JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_1 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= " 	GROUP BY s.distid " & vbCrLf
        sql &= " ) s1 ON s1.distid = d.distid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.distid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='1' THEN s.STUDCNT END) D1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='2' THEN s.STUDCNT END) D2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='3' THEN s.STUDCNT END) D3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='4' THEN s.STUDCNT END) D4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV IN ('1','2','3','4') THEN s.STUDCNT END) D " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN s.STUDCNT END) DA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN s.STUDCNT END) DA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN s.STUDCNT END) DA3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN s.STUDCNT END) DA4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV in ('1','2','3','4') THEN s.STUDCNT END) DA " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN 1 END) DB1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN 1 END) DB2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN 1 END) DB3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN 1 END) DB4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV IN ('1','2','3','4') THEN 1 END) DB " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_2 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= " 	GROUP BY s.distid " & vbCrLf
        sql &= " ) s2 ON s2.distid = d.distid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.distid" & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) E1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) E2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) E3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) E4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) E5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) E6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) E7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) E8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) E9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) E10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) E11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) E12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) E13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) E14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) E15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) E16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) E17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) E18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) E99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) E " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) F1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) F2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) F3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) F4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) F5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) F6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) F7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) F8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) F99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) F " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) Ea1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) Ea2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) Ea3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) Ea4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) Ea5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) Ea6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) Ea7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) Ea8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) Ea9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) Ea10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) Ea11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) Ea12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) Ea13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) Ea14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) Ea15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) Ea16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) Ea17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) Ea18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) Ea99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) Ea " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) Fa1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) Fa2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) Fa3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) Fa4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) Fa5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) Fa6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) Fa7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) Fa8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) Fa99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) Fa " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN 1 END) Eb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN 1 END) Eb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN 1 END) Eb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN 1 END) Eb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN 1 END) Eb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN 1 END) Eb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN 1 END) Eb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN 1 END) Eb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN 1 END) Eb9 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN 1 END) Eb10 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN 1 END) Eb11 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN 1 END) Eb12 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN 1 END) Eb13 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN 1 END) Eb14 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN 1 END) Eb15 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN 1 END) Eb16 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN 1 END) Eb17 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN 1 END) Eb18 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN 1 END) Eb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN 1 END) Eb " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN 1 END) Fb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN 1 END) Fb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN 1 END) Fb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN 1 END) Fb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN 1 END) Fb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN 1 END) Fb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN 1 END) Fb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN 1 END) Fb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN 1 END) Fb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN 1 END) Fb " & vbCrLf
        '身心障礙學員屬公法上救助關係領取津貼就業人數
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt1 " & vbCrLf
        '身心障礙學員屬公法上救助關係領取津貼提前就業人數
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt2 " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN dbo.V_ADPFDMOI_3 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint = 1 " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j9 WITH(NOLOCK) ON j9.SOCID = s.SOCID AND j9.CPoint = 9 AND j.SOCID IS NULL " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.SUB_SUBSIDYAPPLY ss3 WITH(NOLOCK) ON ss3.SOCID = s.SOCID AND ss3.AppliedStatusFin = 'Y' " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql &= " 	GROUP BY s.distid " & vbCrLf
        sql &= " ) s3 ON s3.distid = d.distid " & vbCrLf
        sql &= " ORDER BY d.distid " & vbCrLf
        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    '計畫別SQL
    Function LoadData2() As DataTable
        Dim rst As DataTable

        '報表要用的轄區參數,1:為起始位置
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        Dim strSchcc As String = ""
        If Syear.SelectedValue <> "" Then strSchcc &= " AND c.years = '" & Syear.SelectedValue & "' " & vbCrLf
        'If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate1.Text)), TIMS.to_date(STDate1.Text)) & vbCrLf  'edit，by:20181022
        'If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate2.Text)), TIMS.to_date(STDate2.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate1.Text)), TIMS.to_date(FTDate1.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate2.Text)), TIMS.to_date(FTDate2.Text)) & vbCrLf  'edit，by:20181022
        If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf
        If DistID1 <> "" Then strSchcc &= " AND c.DistID IN (" & DistID1.Replace("\'", "'") & ") " & vbCrLf '轉換sql查詢使用

        Dim sWC1 As String = ""
        sWC1 = ""
        sWC1 &= " WITH WC1 AS ( " & vbCrLf
        sWC1 &= " SELECT c.OCID,c.STDATE,c.FTDATE,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " FROM dbo.VIEW2 c " & vbCrLf
        sWC1 &= " WHERE 1=1 " & vbCrLf
        sWC1 &= strSchcc
        sWC1 &= " )" & vbCrLf
        sWC1 &= " ,WC2 AS ( " & vbCrLf
        sWC1 &= " SELECT s.SOCID,s.OCID,s.SEX,s.STUDSTATUS,s.IDNO,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " ,s.WorkSuppIdent " & vbCrLf '在職者 
        sWC1 &= " ,CASE WHEN c.TPLANID='28' THEN dbo.FN_GET_STUDCNT14B(s.STUDSTATUS,s.REJECTTDATE1,s.REJECTTDATE2,c.STDATE) ELSE 1 END STUDCNT" & vbCrLf '開訓人數 
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.V_STUDENTINFO s on c.ocid = s.ocid " & vbCrLf
        sWC1 &= " ) " & vbCrLf
        sWC1 &= " ,WC3 AS ( " & vbCrLf
        sWC1 &= " SELECT s.IDNO,c.OCID,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf '身障報名人數
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.VIEW_ENTERTYPE12 s on s.OCID1=c.OCID" & vbCrLf
        sWC1 &= " JOIN dbo.V_ADPFDMOI_1 a on a.IDNO=s.IDNO" & vbCrLf
        sWC1 &= " ) " & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= sWC1

        sql &= " SELECT d.planname " & vbCrLf
        sql &= " ,ISNULL(c.clsNum,0) clsNum " & vbCrLf '班數
        sql &= " ,ISNULL(c3.entNum,0) entNum" & vbCrLf '身障報名人數
        sql &= " ,ISNULL(s.A1,0) A1 " & vbCrLf
        sql &= " ,ISNULL(s.A2,0) A2 " & vbCrLf
        sql &= " ,ISNULL(s.A,0) A " & vbCrLf
        sql &= " ,ISNULL(s1.B1,0) B1 " & vbCrLf
        sql &= " ,ISNULL(s1.B2,0) B2 " & vbCrLf
        sql &= " ,ISNULL(s1.B,0) B " & vbCrLf
        sql &= " ,ISNULL(s2.D1,0) D1 " & vbCrLf
        sql &= " ,ISNULL(s2.D2,0) D2 " & vbCrLf
        sql &= " ,ISNULL(s2.D3,0) D3 " & vbCrLf
        sql &= " ,ISNULL(s2.D4,0) D4 " & vbCrLf
        sql &= " ,ISNULL(s2.D,0) D " & vbCrLf
        sql &= " ,ISNULL(s3.E1,0) E1 " & vbCrLf
        sql &= " ,ISNULL(s3.E2,0) E2 " & vbCrLf
        sql &= " ,ISNULL(s3.E3,0) E3 " & vbCrLf
        sql &= " ,ISNULL(s3.E4,0) E4 " & vbCrLf
        sql &= " ,ISNULL(s3.E5,0) E5 " & vbCrLf
        sql &= " ,ISNULL(s3.E6,0) E6 " & vbCrLf
        sql &= " ,ISNULL(s3.E7,0) E7 " & vbCrLf
        sql &= " ,ISNULL(s3.E8,0) E8 " & vbCrLf
        sql &= " ,ISNULL(s3.E9,0) E9 " & vbCrLf
        sql &= " ,ISNULL(s3.E10,0) E10 " & vbCrLf
        sql &= " ,ISNULL(s3.E11,0) E11 " & vbCrLf
        sql &= " ,ISNULL(s3.E12,0) E12 " & vbCrLf
        sql &= " ,ISNULL(s3.E13,0) E13 " & vbCrLf
        sql &= " ,ISNULL(s3.E14,0) E14 " & vbCrLf
        sql &= " ,ISNULL(s3.E15,0) E15 " & vbCrLf
        sql &= " ,ISNULL(s3.E16,0) E16 " & vbCrLf
        sql &= " ,ISNULL(s3.E17,0) E17 " & vbCrLf
        sql &= " ,ISNULL(s3.E18,0) E18 " & vbCrLf
        sql &= " ,ISNULL(s3.E99,0) E99 " & vbCrLf
        sql &= " ,ISNULL(s3.E,0) E " & vbCrLf
        sql &= " ,ISNULL(s3.F1,0) F1 " & vbCrLf
        sql &= " ,ISNULL(s3.F2,0) F2 " & vbCrLf
        sql &= " ,ISNULL(s3.F3,0) F3 " & vbCrLf
        sql &= " ,ISNULL(s3.F4,0) F4 " & vbCrLf
        sql &= " ,ISNULL(s3.F5,0) F5 " & vbCrLf
        sql &= " ,ISNULL(s3.F6,0) F6 " & vbCrLf
        sql &= " ,ISNULL(s3.F7,0) F7 " & vbCrLf
        sql &= " ,ISNULL(s3.F8,0) F8 " & vbCrLf
        sql &= " ,ISNULL(s3.F99,0) F99 " & vbCrLf
        sql &= " ,ISNULL(s3.F,0) F " & vbCrLf
        sql &= " ,ISNULL(s.Aa1,0) AA1 " & vbCrLf
        sql &= " ,ISNULL(s.Aa2,0) AA2 " & vbCrLf
        sql &= " ,ISNULL(s.Aa,0) AA " & vbCrLf
        sql &= " ,ISNULL(s1.Ba1,0) BA1 " & vbCrLf
        sql &= " ,ISNULL(s1.Ba2,0) BA2 " & vbCrLf
        sql &= " ,ISNULL(s1.Ba,0) BA " & vbCrLf
        sql &= " ,ISNULL(s2.Da1,0) DA1 " & vbCrLf
        sql &= " ,ISNULL(s2.Da2,0) DA2 " & vbCrLf
        sql &= " ,ISNULL(s2.Da3,0) DA3 " & vbCrLf
        sql &= " ,ISNULL(s2.Da4,0) DA4 " & vbCrLf
        sql &= " ,ISNULL(s2.Da,0) DA " & vbCrLf
        sql &= " ,ISNULL(s3.Ea1,0) EA1 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea2,0) EA2 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea3,0) EA3 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea4,0) EA4 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea5,0) EA5 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea6,0) EA6 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea7,0) EA7 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea8,0) EA8 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea9,0) EA9 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea10,0) EA10 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea11,0) EA11 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea12,0) EA12 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea13,0) EA13 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea14,0) EA14 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea15,0) EA15 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea16,0) EA16 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea17,0) EA17 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea18,0) EA18 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea99,0) EA99 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea,0) EA " & vbCrLf
        sql &= " ,ISNULL(s3.Fa1,0) FA1 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa2,0) FA2 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa3,0) FA3 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa4,0) FA4 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa5,0) FA5 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa6,0) FA6 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa7,0) FA7 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa8,0) FA8 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa99,0) FA99 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa,0) FA " & vbCrLf

        'sql &= " ,ISNULL(s.Ab1,0) Ab1 " & vbCrLf
        'sql &= " ,ISNULL(s.Ab2,0) Ab2 " & vbCrLf
        'sql &= " ,ISNULL(s.Ab,0) Ab " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb1,0) Bb1 " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb2,0) Bb2 " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb,0) Bb " & vbCrLf
        'sql &= " ,ISNULL(s2.Db1,0) Db1 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db2,0) Db2 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db3,0) Db3 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db4,0) Db4 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db,0) Db " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb1,0) Eb1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb2,0) Eb2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb3,0) Eb3 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb4,0) Eb4 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb5,0) Eb5 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb6,0) Eb6 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb7,0) Eb7 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb8,0) Eb8 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb9,0) Eb9 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb10,0) Eb10 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb11,0) Eb11 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb12,0) Eb12 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb13,0) Eb13 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb14,0) Eb14 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb15,0) Eb15 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb16,0) Eb16 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb17,0) Eb17 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb18,0) Eb18 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb99,0) Eb99 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb,0) Eb " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb1,0) Fb1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb2,0) Fb2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb3,0) Fb3 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb4,0) Fb4 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb5,0) Fb5 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb6,0) Fb6 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb7,0) Fb7 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb8,0) Fb8 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb99,0) Fb99 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb,0) Fb " & vbCrLf
        'sql &= " ,ISNULL(s3.Gm1,0) Gm1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gf1,0) Gf1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gt1,0) Gt1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gm2,0) Gm2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gf2,0) Gf2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gt2,0) Gt2 " & vbCrLf
        'sql &= " ,ISNULL(s1.B3,0) B3 " & vbCrLf
        'sql &= " ,ISNULL(s1.B4,0) B4 " & vbCrLf
        'sql &= " ,ISNULL(s1.B5,0) B5 " & vbCrLf '在職開訓人數(舊制 身障 男女)
        'sql &= " ,ISNULL(s1.BA3,0) BA3 " & vbCrLf
        'sql &= " ,ISNULL(s1.BA4,0) BA4 " & vbCrLf
        'sql &= " ,ISNULL(s1.BA5,0) BA5 " & vbCrLf '在職結訓人數(舊制 身障 男女)
        sql &= " FROM ( " & vbCrLf
        sql &= " SELECT * FROM key_plan WHERE 1=1 " & vbCrLf
        '計畫限定。
        sql &= " AND TPlanID IN (" & cst_TPlanidS1 & ") " & vbCrLf
        sql &= " ) d " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.TPLANID" & vbCrLf
        sql &= " 	,COUNT(1) clsNum " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.TPLANID " & vbCrLf
        sql &= " ) c on c.TPLANID =d.TPLANID " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.TPLANID" & vbCrLf
        sql &= " 	,COUNT(1) entNum " & vbCrLf
        sql &= " 	FROM WC3 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.TPLANID " & vbCrLf
        sql &= " ) c3 on c3.TPLANID =d.TPLANID " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= "  SELECT s.TPLANID " & vbCrLf
        sql &= " 	,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) A1 " & vbCrLf '開訓人數男
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) A2 " & vbCrLf '開訓人數女
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) A " & vbCrLf '開訓人數
        '結訓男/'結訓女/'結訓男女
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) AA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT END) AA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) AA " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) AB1 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) AB2 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) AB " & vbCrLf
        sql &= "  FROM WC1 c " & vbCrLf
        sql &= "  JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        'sql &= "  LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        sql += strSchcc
        sql &= "  GROUP BY s.tplanid " & vbCrLf
        sql &= " ) s ON s.tplanid = d.tplanid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.tplanid" & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) B1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) B2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) B " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) BA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT  END) BA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) BA " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) BB1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) BB2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) BB " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B3 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B4 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B5 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA3 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA4 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA5 " & vbCrLf
        sql &= "    FROM WC1 c " & vbCrLf
        sql &= "    JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_1 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.tplanid " & vbCrLf
        sql &= " ) s1 ON s1.tplanid = d.tplanid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.tplanid" & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='1' THEN s.STUDCNT END) D1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='2' THEN s.STUDCNT END) D2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='3' THEN s.STUDCNT END) D3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='4' THEN s.STUDCNT END) D4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV IN ('1','2','3','4') THEN s.STUDCNT END) D " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN s.STUDCNT END) DA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN s.STUDCNT END) DA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN s.STUDCNT END) DA3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN s.STUDCNT END) DA4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV in ('1','2','3','4') THEN s.STUDCNT END) DA " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN 1 END) DB1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN 1 END) DB2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN 1 END) DB3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN 1 END) DB4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV IN ('1','2','3','4') THEN 1 END) DB " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_2 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.tplanid " & vbCrLf
        sql &= " ) s2 ON s2.tplanid = d.tplanid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.tplanid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) E1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) E2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) E3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) E4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) E5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) E6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) E7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) E8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) E9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) E10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) E11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) E12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) E13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) E14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) E15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) E16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) E17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) E18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) E99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) E " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) F1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) F2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) F3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) F4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) F5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) F6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) F7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) F8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) F99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) F " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) Ea1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) Ea2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) Ea3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) Ea4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) Ea5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) Ea6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) Ea7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) Ea8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) Ea9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) Ea10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) Ea11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) Ea12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) Ea13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) Ea14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) Ea15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) Ea16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) Ea17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) Ea18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) Ea99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) Ea " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) Fa1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) Fa2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) Fa3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) Fa4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) Fa5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) Fa6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) Fa7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) Fa8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) Fa99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) Fa " & vbCrLf

        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN 1 END) Eb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN 1 END) Eb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN 1 END) Eb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN 1 END) Eb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN 1 END) Eb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN 1 END) Eb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN 1 END) Eb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN 1 END) Eb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN 1 END) Eb9 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN 1 END) Eb10 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN 1 END) Eb11 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN 1 END) Eb12 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN 1 END) Eb13 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN 1 END) Eb14 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN 1 END) Eb15 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN 1 END) Eb16 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN 1 END) Eb17 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN 1 END) Eb18 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN 1 END) Eb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN 1 END) Eb " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN 1 END) Fb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN 1 END) Fb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN 1 END) Fb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN 1 END) Fb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN 1 END) Fb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN 1 END) Fb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN 1 END) Fb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN 1 END) Fb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN 1 END) Fb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN 1 END) Fb " & vbCrLf
        ''身心障礙學員屬公法上救助關係領取津貼就業人數
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt1 " & vbCrLf
        ''身心障礙學員屬公法上救助關係領取津貼提前就業人數
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt2 " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN dbo.V_ADPFDMOI_3 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint = 1 " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j9 WITH(NOLOCK) ON j9.SOCID = s.SOCID AND j9.CPoint = 9 AND j.SOCID IS NULL " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.SUB_SUBSIDYAPPLY ss3 WITH(NOLOCK) ON ss3.SOCID = s.SOCID AND ss3.AppliedStatusFin = 'Y' " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.tplanid " & vbCrLf
        sql &= " ) s3 ON s3.tplanid = d.tplanid " & vbCrLf
        sql &= " ORDER BY d.tplanid " & vbCrLf
        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    '轄區計畫別SQL
    Function LoadData3() As DataTable
        Dim rst As DataTable

        '報表要用的轄區參數,1:為起始位置
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        Dim strSchcc As String = ""
        If Syear.SelectedValue <> "" Then strSchcc &= " AND c.years = '" & Syear.SelectedValue & "'" & vbCrLf
        'If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate1.Text)), TIMS.to_date(STDate1.Text)) & vbCrLf  'edit，by:20181022
        'If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(STDate2.Text)), TIMS.to_date(STDate2.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate1.Text)), TIMS.to_date(FTDate1.Text)) & vbCrLf  'edit，by:20181022
        'If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & IIf(flag_ROC, TIMS.to_date(TIMS.cdate18(FTDate2.Text)), TIMS.to_date(FTDate2.Text)) & vbCrLf  'edit，by:20181022
        If STDate1.Text <> "" Then strSchcc &= " AND c.stdate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then strSchcc &= " AND c.stdate <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        If FTDate1.Text <> "" Then strSchcc &= " AND c.ftdate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        If FTDate2.Text <> "" Then strSchcc &= " AND c.ftdate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf
        If DistID1 <> "" Then strSchcc &= " AND c.DistID IN (" & DistID1.Replace("\'", "'") & ") " & vbCrLf '轉換sql查詢使用

        Dim sWC1 As String = ""
        sWC1 = ""
        sWC1 &= " WITH WC1 AS ( " & vbCrLf
        sWC1 &= " SELECT c.OCID,c.STDATE,c.FTDATE,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " FROM dbo.VIEW2 c " & vbCrLf
        sWC1 &= " WHERE 1=1 " & vbCrLf
        sWC1 &= strSchcc
        sWC1 &= " )" & vbCrLf
        sWC1 &= " ,WC2 AS ( " & vbCrLf
        sWC1 &= " SELECT s.SOCID,s.OCID,s.SEX,s.STUDSTATUS,s.IDNO,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf
        sWC1 &= " ,s.WorkSuppIdent " & vbCrLf '在職者 
        sWC1 &= " ,CASE WHEN c.TPLANID='28' THEN dbo.FN_GET_STUDCNT14B(s.STUDSTATUS,s.REJECTTDATE1,s.REJECTTDATE2,c.STDATE) ELSE 1 END STUDCNT" & vbCrLf '開訓人數 
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.V_STUDENTINFO s on c.ocid = s.ocid " & vbCrLf
        sWC1 &= " ) " & vbCrLf
        sWC1 &= " ,WC3 AS ( " & vbCrLf
        sWC1 &= " SELECT s.IDNO,c.OCID,c.DISTID,c.TPLANID,c.YEARS " & vbCrLf '身障報名人數
        sWC1 &= " FROM WC1 c " & vbCrLf
        sWC1 &= " JOIN dbo.VIEW_ENTERTYPE12 s on s.OCID1=c.OCID" & vbCrLf
        sWC1 &= " JOIN dbo.V_ADPFDMOI_1 a on a.IDNO=s.IDNO" & vbCrLf
        sWC1 &= " ) " & vbCrLf

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= sWC1

        sql &= " SELECT d.distname " & vbCrLf
        sql &= " ,d.planname " & vbCrLf
        sql &= " ,ISNULL(c.clsNum,0) clsNum " & vbCrLf '班數
        sql &= " ,ISNULL(c3.entNum,0) entNum" & vbCrLf '身障報名人數

        sql &= " ,ISNULL(s.A1,0) A1 " & vbCrLf
        sql &= " ,ISNULL(s.A2,0) A2 " & vbCrLf
        sql &= " ,ISNULL(s.A,0) A " & vbCrLf
        sql &= " ,ISNULL(s1.B1,0) B1 " & vbCrLf
        sql &= " ,ISNULL(s1.B2,0) B2 " & vbCrLf
        sql &= " ,ISNULL(s1.B,0) B " & vbCrLf
        sql &= " ,ISNULL(s2.D1,0) D1 " & vbCrLf
        sql &= " ,ISNULL(s2.D2,0) D2 " & vbCrLf
        sql &= " ,ISNULL(s2.D3,0) D3 " & vbCrLf
        sql &= " ,ISNULL(s2.D4,0) D4 " & vbCrLf
        sql &= " ,ISNULL(s2.D,0) D " & vbCrLf
        sql &= " ,ISNULL(s3.E1,0) E1 " & vbCrLf
        sql &= " ,ISNULL(s3.E2,0) E2 " & vbCrLf
        sql &= " ,ISNULL(s3.E3,0) E3 " & vbCrLf
        sql &= " ,ISNULL(s3.E4,0) E4 " & vbCrLf
        sql &= " ,ISNULL(s3.E5,0) E5 " & vbCrLf
        sql &= " ,ISNULL(s3.E6,0) E6 " & vbCrLf
        sql &= " ,ISNULL(s3.E7,0) E7 " & vbCrLf
        sql &= " ,ISNULL(s3.E8,0) E8 " & vbCrLf
        sql &= " ,ISNULL(s3.E9,0) E9 " & vbCrLf
        sql &= " ,ISNULL(s3.E10,0) E10 " & vbCrLf
        sql &= " ,ISNULL(s3.E11,0) E11 " & vbCrLf
        sql &= " ,ISNULL(s3.E12,0) E12 " & vbCrLf
        sql &= " ,ISNULL(s3.E13,0) E13 " & vbCrLf
        sql &= " ,ISNULL(s3.E14,0) E14 " & vbCrLf
        sql &= " ,ISNULL(s3.E15,0) E15 " & vbCrLf
        sql &= " ,ISNULL(s3.E16,0) E16 " & vbCrLf
        sql &= " ,ISNULL(s3.E17,0) E17 " & vbCrLf
        sql &= " ,ISNULL(s3.E18,0) E18 " & vbCrLf
        sql &= " ,ISNULL(s3.E99,0) E99 " & vbCrLf
        sql &= " ,ISNULL(s3.E,0) E " & vbCrLf
        sql &= " ,ISNULL(s3.F1,0) F1 " & vbCrLf
        sql &= " ,ISNULL(s3.F2,0) F2 " & vbCrLf
        sql &= " ,ISNULL(s3.F3,0) F3 " & vbCrLf
        sql &= " ,ISNULL(s3.F4,0) F4 " & vbCrLf
        sql &= " ,ISNULL(s3.F5,0) F5 " & vbCrLf
        sql &= " ,ISNULL(s3.F6,0) F6 " & vbCrLf
        sql &= " ,ISNULL(s3.F7,0) F7 " & vbCrLf
        sql &= " ,ISNULL(s3.F8,0) F8 " & vbCrLf
        sql &= " ,ISNULL(s3.F99,0) F99 " & vbCrLf
        sql &= " ,ISNULL(s3.F,0) F " & vbCrLf
        sql &= " ,ISNULL(s.Aa1,0) AA1 " & vbCrLf
        sql &= " ,ISNULL(s.Aa2,0) AA2 " & vbCrLf
        sql &= " ,ISNULL(s.Aa,0) AA " & vbCrLf
        sql &= " ,ISNULL(s1.Ba1,0) BA1 " & vbCrLf
        sql &= " ,ISNULL(s1.Ba2,0) BA2 " & vbCrLf
        sql &= " ,ISNULL(s1.Ba,0) BA " & vbCrLf
        sql &= " ,ISNULL(s2.Da1,0) DA1 " & vbCrLf
        sql &= " ,ISNULL(s2.Da2,0) DA2 " & vbCrLf
        sql &= " ,ISNULL(s2.Da3,0) DA3 " & vbCrLf
        sql &= " ,ISNULL(s2.Da4,0) DA4 " & vbCrLf
        sql &= " ,ISNULL(s2.Da,0) DA " & vbCrLf
        sql &= " ,ISNULL(s3.Ea1,0) EA1 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea2,0) EA2 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea3,0) EA3 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea4,0) EA4 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea5,0) EA5 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea6,0) EA6 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea7,0) EA7 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea8,0) EA8 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea9,0) EA9 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea10,0) EA10 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea11,0) EA11 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea12,0) EA12 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea13,0) EA13 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea14,0) EA14 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea15,0) EA15 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea16,0) EA16 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea17,0) EA17 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea18,0) EA18 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea99,0) EA99 " & vbCrLf
        sql &= " ,ISNULL(s3.Ea,0) EA " & vbCrLf
        sql &= " ,ISNULL(s3.Fa1,0) FA1 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa2,0) FA2 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa3,0) FA3 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa4,0) FA4 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa5,0) FA5 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa6,0) FA6 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa7,0) FA7 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa8,0) FA8 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa99,0) FA99 " & vbCrLf
        sql &= " ,ISNULL(s3.Fa,0) FA " & vbCrLf

        'sql &= " ,ISNULL(s.Ab1,0) Ab1 " & vbCrLf
        'sql &= " ,ISNULL(s.Ab2,0) Ab2 " & vbCrLf
        'sql &= " ,ISNULL(s.Ab,0) Ab " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb1,0) Bb1 " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb2,0) Bb2 " & vbCrLf
        'sql &= " ,ISNULL(s1.Bb,0) Bb " & vbCrLf
        'sql &= " ,ISNULL(s2.Db1,0) Db1 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db2,0) Db2 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db3,0) Db3 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db4,0) Db4 " & vbCrLf
        'sql &= " ,ISNULL(s2.Db,0) Db " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb1,0) Eb1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb2,0) Eb2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb3,0) Eb3 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb4,0) Eb4 " & vbCrLf
        'sql &= "  ,ISNULL(s3.Eb5,0) Eb5 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb6,0) Eb6 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb7,0) Eb7 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb8,0) Eb8 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb9,0) Eb9 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb10,0) Eb10 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb11,0) Eb11 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb12,0) Eb12 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb13,0) Eb13 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb14,0) Eb14 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb15,0) Eb15 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb16,0) Eb16 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb17,0) Eb17 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb18,0) Eb18 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb99,0) Eb99 " & vbCrLf
        'sql &= " ,ISNULL(s3.Eb,0) Eb " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb1,0) Fb1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb2,0) Fb2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb3,0) Fb3 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb4,0) Fb4 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb5,0) Fb5 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb6,0) Fb6 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb7,0) Fb7 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb8,0) Fb8 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb99,0) Fb99 " & vbCrLf
        'sql &= " ,ISNULL(s3.Fb,0) Fb " & vbCrLf
        'sql &= " ,ISNULL(s3.Gm1,0) Gm1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gf1,0) Gf1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gt1,0) Gt1 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gm2,0) Gm2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gf2,0) Gf2 " & vbCrLf
        'sql &= " ,ISNULL(s3.Gt2,0) Gt2 " & vbCrLf
        'sql &= " ,ISNULL(s1.B3,0) B3 " & vbCrLf
        'sql &= " ,ISNULL(s1.B4,0) B4 " & vbCrLf
        'sql &= " ,ISNULL(s1.B5,0) B5 " & vbCrLf '在職開訓人數(舊制 身障 男女)
        'sql &= " ,ISNULL(s1.BA3,0) BA3 " & vbCrLf
        'sql &= " ,ISNULL(s1.BA4,0) BA4 " & vbCrLf
        'sql &= " ,ISNULL(s1.BA5,0) BA5 " & vbCrLf '在職結訓人數(舊制 身障 男女)
        sql &= " FROM ( " & vbCrLf
        sql &= " SELECT DISTINCT a.tplanid, a.planname " & vbCrLf
        sql &= " ,a.distid, a.distname2 distname " & vbCrLf
        sql &= " FROM dbo.VIEW_PLAN a WITH(NOLOCK)" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If Syear.SelectedValue <> "" Then sql &= " AND a.years = '" & Syear.SelectedValue & "' " & vbCrLf
        If DistID1 <> "" Then
            '轉換sql查詢使用
            sql &= " AND a.DistID IN (" & DistID1.Replace("\'", "'") & ") " & vbCrLf
        Else
            sql &= " AND a.DistID NOT IN ('000','002') " & vbCrLf
        End If
        '計畫限定。
        sql &= " AND a.TPlanID IN (" & cst_TPlanidS1 & ") " & vbCrLf
        sql &= " ) d " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.DISTID,c.TPLANID" & vbCrLf
        sql &= " 	,COUNT(1) clsNum " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.DISTID,c.TPLANID " & vbCrLf
        sql &= " ) c on c.distid =d.distid AND c.tplanid = d.tplanid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT c.DISTID,c.TPLANID" & vbCrLf
        sql &= " 	,COUNT(1) entNum " & vbCrLf
        sql &= " 	FROM WC3 c " & vbCrLf
        sql &= " 	where 1=1 " & vbCrLf
        sql &= " 	GROUP BY c.DISTID,c.TPLANID " & vbCrLf
        sql &= " ) c3 on c3.distid =d.distid AND c3.tplanid = d.tplanid" & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= "  SELECT s.distid, s.tplanid " & vbCrLf
        sql &= " 	,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) A1 " & vbCrLf '開訓人數男
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) A2 " & vbCrLf '開訓人數女
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) A " & vbCrLf '開訓人數
        '結訓男/'結訓女/'結訓男女
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) AA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT END) AA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) AA " & vbCrLf

        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) AB1 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) AB2 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) AB " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.distid, s.tplanid " & vbCrLf
        sql &= " ) s ON s.distid = d.distid AND s.tplanid = d.tplanid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= "  SELECT s.distid ,s.tplanid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='M' THEN s.STUDCNT END) B1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex='F' THEN s.STUDCNT END) B2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.sex IN ('M','F') THEN s.STUDCNT END) B " & vbCrLf
        '結訓男/'結訓女/'結訓男女
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN s.STUDCNT END) BA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN s.STUDCNT  END) BA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN s.STUDCNT END) BA " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' THEN 1 END) BB1 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' THEN 1 END) BB2 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') THEN 1 END) BB " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B3 " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B4 " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) B5 " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='M' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA3 " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex='F' AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA4 " & vbCrLf
        'sql &= "  ,COUNT(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND s.sex IN ('M','F') AND ISNULL(s.WorkSuppIdent,' ')='Y' THEN 1 END) BA5 " & vbCrLf
        sql &= "    FROM WC1 c " & vbCrLf
        sql &= "    JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_1 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.distid ,s.tplanid" & vbCrLf
        sql &= " ) s1 ON s1.distid = d.distid AND s1.tplanid = d.tplanid " & vbCrLf

        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.distid ,s.tplanid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='1' THEN s.STUDCNT END) D1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='2' THEN s.STUDCNT END) D2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='3' THEN s.STUDCNT END) D3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV='4' THEN s.STUDCNT END) D4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.HANDLV IN ('1','2','3','4') THEN s.STUDCNT END) D " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN s.STUDCNT END) DA1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN s.STUDCNT END) DA2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN s.STUDCNT END) DA3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN s.STUDCNT END) DA4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV in ('1','2','3','4') THEN s.STUDCNT END) DA " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='1' THEN 1 END) DB1 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='2' THEN 1 END) DB2 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='3' THEN 1 END) DB3 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV='4' THEN 1 END) DB4 " & vbCrLf
        'sql &= "  ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.HANDLV IN ('1','2','3','4') THEN 1 END) DB " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN V_ADPFDMOI_2 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint=1 " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.distid ,s.tplanid " & vbCrLf
        sql &= " ) s2 ON s2.distid = d.distid AND s2.tplanid = d.tplanid " & vbCrLf
        sql &= " LEFT JOIN ( " & vbCrLf
        sql &= " 	SELECT s.distid ,s.tplanid " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) E1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) E2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) E3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) E4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) E5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) E6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) E7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) E8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) E9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) E10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) E11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) E12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) E13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) E14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) E15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) E16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) E17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) E18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) E99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) E " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) F1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) F2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) F3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) F4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) F5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) F6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) F7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) F8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) F99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) F " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN s.STUDCNT END) Ea1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN s.STUDCNT END) Ea2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN s.STUDCNT END) Ea3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN s.STUDCNT END) Ea4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN s.STUDCNT END) Ea5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN s.STUDCNT END) Ea6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN s.STUDCNT END) Ea7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN s.STUDCNT END) Ea8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN s.STUDCNT END) Ea9 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN s.STUDCNT END) Ea10 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN s.STUDCNT END) Ea11 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN s.STUDCNT END) Ea12 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN s.STUDCNT END) Ea13 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN s.STUDCNT END) Ea14 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN s.STUDCNT END) Ea15 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN s.STUDCNT END) Ea16 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN s.STUDCNT END) Ea17 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN s.STUDCNT END) Ea18 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN s.STUDCNT END) Ea99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN s.STUDCNT END) Ea " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN s.STUDCNT END) Fa1 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN s.STUDCNT END) Fa2 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN s.STUDCNT END) Fa3 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN s.STUDCNT END) Fa4 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN s.STUDCNT END) Fa5 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN s.STUDCNT END) Fa6 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN s.STUDCNT END) Fa7 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN s.STUDCNT END) Fa8 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN s.STUDCNT END) Fa99 " & vbCrLf
        sql &= "    ,SUM(CASE WHEN s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN s.STUDCNT END) Fa " & vbCrLf

        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='01' THEN 1 END) Eb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='02' THEN 1 END) Eb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='03' THEN 1 END) Eb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='04' THEN 1 END) Eb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='05' THEN 1 END) Eb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='06' THEN 1 END) Eb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='07' THEN 1 END) Eb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='08' THEN 1 END) Eb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='09' THEN 1 END) Eb9 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='10' THEN 1 END) Eb10 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='11' THEN 1 END) Eb11 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='12' THEN 1 END) Eb12 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='13' THEN 1 END) Eb13 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='14' THEN 1 END) Eb14 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='15' THEN 1 END) Eb15 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='16' THEN 1 END) Eb16 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='17' THEN 1 END) Eb17 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='18' THEN 1 END) Eb18 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP='99' THEN 1 END) Eb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =1 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','99') THEN 1 END) Eb " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='01' THEN 1 END) Fb1 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='02' THEN 1 END) Fb2 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='03' THEN 1 END) Fb3 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='04' THEN 1 END) Fb4 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='05' THEN 1 END) Fb5 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='06' THEN 1 END) Fb6 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='07' THEN 1 END) Fb7 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='08' THEN 1 END) Fb8 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP='99' THEN 1 END) Fb99 " & vbCrLf
        'sql &= "    ,SUM(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE < GETDATE() AND a.IMSOURCE =2 AND a.HANDTP IN ('01','02','03','04','05','06','07','08','99') THEN 1 END) Fb " & vbCrLf
        ''身心障礙學員屬公法上救助關係領取津貼就業人數
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf1 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j.IsGetJob=1 AND s.STUDSTATUS NOT IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt1 " & vbCrLf
        ''身心障礙學員屬公法上救助關係領取津貼提前就業人數
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='M' AND ss3.SOCID IS NOT NULL THEN 1 END) Gm2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex='F' AND ss3.SOCID IS NOT NULL THEN 1 END) Gf2 " & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN j9.IsGetJob=1 AND s.STUDSTATUS IN (2,3) AND c.FTDATE<GETDATE() AND a.IMSOURCE IN (1,2) AND j9.PUBLICRESCUE='Y' AND s.sex IN ('M','F') AND ss3.SOCID IS NOT NULL THEN 1 END) Gt2 " & vbCrLf
        sql &= " 	FROM WC1 c " & vbCrLf
        sql &= " 	JOIN WC2 s ON c.ocid = s.ocid " & vbCrLf
        sql &= " 	JOIN dbo.V_ADPFDMOI_3 a ON a.IDNO = s.IDNO " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j WITH(NOLOCK) ON j.SOCID = s.SOCID AND j.CPoint = 1 " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.STUD_GETJOBSTATE3 j9 WITH(NOLOCK) ON j9.SOCID = s.SOCID AND j9.CPoint = 9 AND j.SOCID IS NULL " & vbCrLf
        'sql &= " 	LEFT JOIN dbo.SUB_SUBSIDYAPPLY ss3 WITH(NOLOCK) ON ss3.SOCID = s.SOCID AND ss3.AppliedStatusFin = 'Y' " & vbCrLf
        sql &= " 	WHERE 1=1 " & vbCrLf
        sql += strSchcc
        sql &= " 	GROUP BY s.distid ,s.tplanid " & vbCrLf
        sql &= " ) s3 ON s3.distid = d.distid AND s3.tplanid = d.tplanid " & vbCrLf
        sql &= " ORDER BY d.tplanid, d.distid " & vbCrLf
        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    ''' <summary>匯出 EXCEL 組合EXCEL( xType:1.分署 2.計畫別 3.分署計畫別)</summary>
    ''' <param name="dt"></param>
    ''' <param name="xType"></param>
    Sub ExpReport1(ByRef dt As DataTable, ByVal xType As Integer)
        'Dim strTitle1 As String = "身心障礙統計表"
        '資料來源：已核對衛服部提供103年12月份身心障礙者資料
        '資料來源：已核對衛福部提供○年○月份身心障礙者資料，其中新制障礙類別已轉換為舊制
        Dim cst_memo1 As String = Get_Memo1()

        'Dim str_Allspan As String = CStr(137)
        Const cst_ALL_COL_COUNT As Integer = 85
        Dim str_Allspan As String = CStr(cst_ALL_COL_COUNT)
        'If xType = 3 Then str_Allspan = CStr(138) ' 3.分署計畫別 '2018-09-06 fix 區域計畫別報表-表頭合併欄寬數
        If xType = 3 Then str_Allspan = CStr(cst_ALL_COL_COUNT + 1) ' 3.分署計畫別 '2018-09-06 fix 區域計畫別報表-表頭合併欄寬數

        Dim dtType1 As DataTable = Nothing
        Dim dtType2 As DataTable = Nothing
        '舊制障礙類別
        'Sql = "SELECT HandTypeID , Name FROM Key_HandicatType WHERE HandTypeID!='00' ORDER BY HandTypeID "
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT HandTypeID ,Name FROM Key_HandicatType WHERE HandTypeID != '00' " & vbCrLf
        'sql += " UNION SELECT '99' HandTypeID ,N'其它' Name " & vbCrLf
        sql &= " ORDER BY HandTypeID " & vbCrLf
        dtType1 = DbAccess.GetDataTable(sql, objconn)

        '新制障礙類別
        sql = " SELECT HandTypeID2 HandTypeID, Name FROM Key_HandicatType2 ORDER BY HandTypeID2 "
        dtType2 = DbAccess.GetDataTable(sql, objconn)

        Dim sFileName1 As String = "身心障礙者參加融合式職業訓練人數統計表"
        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= (".noDecFormat2{mso-number-format:""\@"";}")
        'mso-number-format:"0" 
        strSTYLE &= ("</style>")


        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        'strTitle2 開訓日期
        'strTitle3 結訓日期
        'strTitle4 轄區
        Dim strTitle1 As String = sFileName1
        Dim strTitle2 As String = ""
        Dim strTitle3 As String = ""
        Dim strTitle4 As String = ""
        'strTitle2 = "開訓日期：" & TIMS.CFormatDate(IIf(flag_ROC, TIMS.cdate18(STDate1.Text), STDate1.Text)) & "～" & TIMS.CFormatDate(IIf(flag_ROC, TIMS.cdate18(STDate2.Text), STDate2.Text))  'edit，by:20181022
        'strTitle3 = "結訓日期：" & TIMS.CFormatDate(IIf(flag_ROC, TIMS.cdate18(FTDate1.Text), FTDate1.Text)) & "～" & TIMS.CFormatDate(IIf(flag_ROC, TIMS.cdate18(FTDate2.Text), FTDate2.Text))  'edit，by:20181022

        If Syear.SelectedValue <> "" Then
            If strTitle2 <> "" Then strTitle2 &= "，"
            strTitle2 &= " 年度：" & Syear.SelectedValue
        End If
        If strTitle2 <> "" Then strTitle2 &= "，"
        strTitle2 = "開訓日期：" & TIMS.CFormatDate(STDate1.Text) & "～" & TIMS.CFormatDate(STDate2.Text)  'edit，by:20181022
        strTitle3 = "結訓日期：" & TIMS.CFormatDate(FTDate1.Text) & "～" & TIMS.CFormatDate(FTDate2.Text)  'edit，by:20181022
        strTitle4 = "轄　　區：" & TIMS.GetCheckBoxListRptText(DistID, 1) & "　 分署(以上資料依據篩選條件顯示)"

        Dim ExportStr As String = ""
        '建立抬頭
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""" & str_Allspan & """ align=""center"">" & strTitle1 & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立抬頭2
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""" & str_Allspan & """>" & strTitle2 & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立抬頭3
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""" & str_Allspan & """>" & strTitle3 & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立抬頭4
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""" & str_Allspan & """  align=""right"">" & strTitle4 & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""" & str_Allspan & """ align=""right"">" & cst_memo1 & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf

        '第2行
        ExportStr &= ""
        ExportStr &= "<tr>" & vbCrLf
        'Select Case xType
        '    Case 1 '分署
        '    Case 2 '計畫別
        'End Select
        Select Case xType
            Case 1 '分署
                ExportStr &= "<td rowspan=""3"">分署</td>" & vbTab
            Case 2 '計畫別
                ExportStr &= "<td rowspan=""3"">計畫別</td>" & vbTab
            Case 3 '分署'計畫別
                ExportStr &= "<td rowspan=""3"">分署</td>" & vbTab
                ExportStr &= "<td rowspan=""3"">計畫別</td>" & vbTab
        End Select
        ExportStr &= "<td rowspan=""3"">班數</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">身障報名人數</td>" & vbTab

        ExportStr &= "<td colspan=""41"">開訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""41"">結訓人數</td>" & vbTab
        'ExportStr &= "<td colspan=""47"">就業人數</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf

        '第3行
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td rowspan=""2"">男</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">女</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">小計</td>" & vbTab
        ExportStr &= "<td colspan=""3"">身心障礙身份開訓人數</td>" & vbTab
        'ExportStr &= "<td colspan=""3"">在職者身心障礙開訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""5"">障礙等級</td>" & vbTab
        ExportStr &= "<td colspan=""20"">舊制障礙類別</td>" & vbTab
        ExportStr &= "<td colspan=""10"">新制障礙類別</td>" & vbTab

        ExportStr &= "<td rowspan=""2"">男</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">女</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">小計</td>" & vbTab
        ExportStr &= "<td colspan=""3"">身心障礙身份結訓人數</td>" & vbTab
        'ExportStr &= "<td colspan=""3"">在職者身心障礙結訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""5"">障礙等級</td>" & vbTab
        ExportStr &= "<td colspan=""20"">舊制障礙類別</td>" & vbTab
        ExportStr &= "<td colspan=""10"">新制障礙類別</td>" & vbTab

        'ExportStr &= "<td rowspan=""2"">男</td>" & vbTab
        'ExportStr &= "<td rowspan=""2"">女</td>" & vbTab
        'ExportStr &= "<td rowspan=""2"">小計</td>" & vbTab
        'ExportStr &= "<td colspan=""3"">身心障礙身份就業人數</td>" & vbTab
        'ExportStr &= "<td colspan=""5"">障礙等級</td>" & vbTab
        'ExportStr &= "<td colspan=""20"">舊制障礙類別</td>" & vbTab
        'ExportStr &= "<td colspan=""10"">新制障礙類別</td>" & vbTab
        '身心障礙學員屬公法上救助關係領取津貼就業人數
        'ExportStr &= "<td colspan=""3"">身心障礙學員屬公法上救助關係領取津貼就業人數</td>" & vbTab
        '身心障礙學員屬公法上救助關係領取津貼提前就業人數
        'ExportStr &= "<td colspan=""3"">身心障礙學員屬公法上救助關係領取津貼提前就業人數</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf

        '第4行
        '身心障礙身份開訓人數
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>男</td>" & vbTab
        ExportStr &= "<td>女</td>" & vbTab
        ExportStr &= "<td>小計</td>" & vbTab
        '在職者身心障礙開訓人數
        'ExportStr &= "<td>男</td>" & vbTab
        'ExportStr &= "<td>女</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab
        '障礙等級
        ExportStr &= "<td>極重度</td>" & vbTab
        ExportStr &= "<td>重度</td>" & vbTab
        ExportStr &= "<td>中度</td>" & vbTab
        ExportStr &= "<td>輕度</td>" & vbTab
        ExportStr &= "<td>小計</td>" & vbTab

        '舊制障礙類別
        For i As Integer = 0 To dtType1.Rows.Count - 1
            ExportStr &= "<td class=""noDecFormat2"">" & dtType1.Rows(i)("HandTypeID") & "</td>" & vbTab
        Next
        ExportStr &= "<td class=""noDecFormat2"">99</td>" & vbTab
        ExportStr &= "<td>小計</td>" & vbTab
        '新制障礙類別
        For i As Integer = 0 To dtType2.Rows.Count - 1
            ExportStr &= "<td class=""noDecFormat2"">" & dtType2.Rows(i)("HandTypeID") & "</td>" & vbTab
        Next
        ExportStr &= "<td>小計</td>" & vbTab


        '身心障礙身份結訓人數
        ExportStr &= "<td>男</td>" & vbTab
        ExportStr &= "<td>女</td>" & vbTab
        ExportStr &= "<td>小計</td>" & vbTab
        '在職者身心障礙結訓人數
        'ExportStr &= "<td>男</td>" & vbTab
        'ExportStr &= "<td>女</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab

        ExportStr &= "<td>極重度</td>" & vbTab
        ExportStr &= "<td>重度</td>" & vbTab
        ExportStr &= "<td>中度</td>" & vbTab
        ExportStr &= "<td>輕度</td>" & vbTab

        ExportStr &= "<td>小計</td>" & vbTab
        For i As Integer = 0 To dtType1.Rows.Count - 1
            ExportStr &= "<td class=""noDecFormat2"">" & dtType1.Rows(i)("HandTypeID") & "</td>" & vbTab
        Next
        ExportStr &= "<td class=""noDecFormat2"">99</td>" & vbTab
        ExportStr &= "<td>小計</td>" & vbTab
        For i As Integer = 0 To dtType2.Rows.Count - 1
            ExportStr &= "<td class=""noDecFormat2"">" & dtType2.Rows(i)("HandTypeID") & "</td>" & vbTab
        Next
        ExportStr &= "<td>小計</td>" & vbTab

        '就業人數
        'ExportStr &= "<td>男</td>" & vbTab
        'ExportStr &= "<td>女</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab

        'ExportStr &= "<td>極重度</td>" & vbTab
        'ExportStr &= "<td>重度</td>" & vbTab
        'ExportStr &= "<td>中度</td>" & vbTab
        'ExportStr &= "<td>輕度</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab

        'For i As Integer = 0 To dtType1.Rows.Count - 1
        '    ExportStr &= "<td class=""noDecFormat2"">" & dtType1.Rows(i)("HandTypeID") & "</td>" & vbTab
        'Next
        'ExportStr &= "<td class=""noDecFormat2"">99</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab
        'For i As Integer = 0 To dtType2.Rows.Count - 1
        '    ExportStr &= "<td class=""noDecFormat2"">" & dtType2.Rows(i)("HandTypeID") & "</td>" & vbTab
        'Next
        'ExportStr &= "<td>小計</td>" & vbTab

        ''身心障礙學員屬公法上救助關係領取津貼就業人數
        'ExportStr &= "<td>男</td>" & vbTab
        'ExportStr &= "<td>女</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab
        ''身心障礙學員屬公法上救助關係領取津貼提前就業人數
        'ExportStr &= "<td>男</td>" & vbTab
        'ExportStr &= "<td>女</td>" & vbTab
        'ExportStr &= "<td>小計</td>" & vbTab

        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        'Dim distname As Integer = 0
        Dim clsNum As Integer = 0
        Dim entNum As Integer = 0
        Dim A1 As Integer = 0
        Dim A2 As Integer = 0
        Dim A As Integer = 0
        'Dim B3 As Integer = 0
        'Dim B4 As Integer = 0
        'Dim B5 As Integer = 0
        'Dim BA3 As Integer = 0
        'Dim BA4 As Integer = 0
        'Dim BA5 As Integer = 0
        Dim B1 As Integer = 0
        Dim B2 As Integer = 0
        Dim B As Integer = 0
        Dim D1 As Integer = 0
        Dim D2 As Integer = 0
        Dim D3 As Integer = 0
        Dim D4 As Integer = 0
        Dim D As Integer = 0
        Dim E1 As Integer = 0
        Dim E2 As Integer = 0
        Dim E3 As Integer = 0
        Dim E4 As Integer = 0
        Dim E5 As Integer = 0
        Dim E6 As Integer = 0
        Dim E7 As Integer = 0
        Dim E8 As Integer = 0
        Dim E9 As Integer = 0
        Dim E10 As Integer = 0
        Dim E11 As Integer = 0
        Dim E12 As Integer = 0
        Dim E13 As Integer = 0
        Dim E14 As Integer = 0
        Dim E15 As Integer = 0
        Dim E16 As Integer = 0
        Dim E17 As Integer = 0
        Dim E18 As Integer = 0
        Dim E99 As Integer = 0
        Dim E As Integer = 0
        Dim F1 As Integer = 0
        Dim F2 As Integer = 0
        Dim F3 As Integer = 0
        Dim F4 As Integer = 0
        Dim F5 As Integer = 0
        Dim F6 As Integer = 0
        Dim F7 As Integer = 0
        Dim F8 As Integer = 0
        Dim F99 As Integer = 0
        Dim F As Integer = 0
        Dim Aa1 As Integer = 0
        Dim Aa2 As Integer = 0
        Dim Aa As Integer = 0
        Dim Ba1 As Integer = 0
        Dim Ba2 As Integer = 0
        Dim Ba As Integer = 0
        Dim Da1 As Integer = 0
        Dim Da2 As Integer = 0
        Dim Da3 As Integer = 0
        Dim Da4 As Integer = 0
        Dim Da As Integer = 0
        Dim Ea1 As Integer = 0
        Dim Ea2 As Integer = 0
        Dim Ea3 As Integer = 0
        Dim Ea4 As Integer = 0
        Dim Ea5 As Integer = 0
        Dim Ea6 As Integer = 0
        Dim Ea7 As Integer = 0
        Dim Ea8 As Integer = 0
        Dim Ea9 As Integer = 0
        Dim Ea10 As Integer = 0
        Dim Ea11 As Integer = 0
        Dim Ea12 As Integer = 0
        Dim Ea13 As Integer = 0
        Dim Ea14 As Integer = 0
        Dim Ea15 As Integer = 0
        Dim Ea16 As Integer = 0
        Dim Ea17 As Integer = 0
        Dim Ea18 As Integer = 0
        Dim Ea99 As Integer = 0
        Dim Ea As Integer = 0
        Dim Fa1 As Integer = 0
        Dim Fa2 As Integer = 0
        Dim Fa3 As Integer = 0
        Dim Fa4 As Integer = 0
        Dim Fa5 As Integer = 0
        Dim Fa6 As Integer = 0
        Dim Fa7 As Integer = 0
        Dim Fa8 As Integer = 0
        Dim Fa99 As Integer = 0
        Dim Fa As Integer = 0
        'Dim Ab1 As Integer = 0
        'Dim Ab2 As Integer = 0
        'Dim Ab As Integer = 0
        'Dim Bb1 As Integer = 0
        'Dim Bb2 As Integer = 0
        'Dim Bb As Integer = 0
        'Dim Db1 As Integer = 0
        'Dim Db2 As Integer = 0
        'Dim Db3 As Integer = 0
        'Dim Db4 As Integer = 0
        'Dim Db As Integer = 0

        'Dim Eb1 As Integer = 0
        'Dim Eb2 As Integer = 0
        'Dim Eb3 As Integer = 0
        'Dim Eb4 As Integer = 0
        'Dim Eb5 As Integer = 0
        'Dim Eb6 As Integer = 0
        'Dim Eb7 As Integer = 0
        'Dim Eb8 As Integer = 0
        'Dim Eb9 As Integer = 0
        'Dim Eb10 As Integer = 0
        'Dim Eb11 As Integer = 0
        'Dim Eb12 As Integer = 0
        'Dim Eb13 As Integer = 0
        'Dim Eb14 As Integer = 0
        'Dim Eb15 As Integer = 0
        'Dim Eb16 As Integer = 0
        'Dim Eb17 As Integer = 0
        'Dim Eb18 As Integer = 0
        'Dim Eb99 As Integer = 0
        'Dim Eb As Integer = 0
        'Dim Fb1 As Integer = 0
        'Dim Fb2 As Integer = 0
        'Dim Fb3 As Integer = 0
        'Dim Fb4 As Integer = 0
        'Dim Fb5 As Integer = 0
        'Dim Fb6 As Integer = 0
        'Dim Fb7 As Integer = 0
        'Dim Fb8 As Integer = 0
        'Dim Fb99 As Integer = 0
        'Dim Fb As Integer = 0
        '身心障礙學員屬公法上救助關係領取津貼就業人數
        'Dim Gm1 As Integer = 0
        'Dim Gf1 As Integer = 0
        'Dim Gt1 As Integer = 0
        'Dim Gm2 As Integer = 0
        'Dim Gf2 As Integer = 0
        'Dim Gt2 As Integer = 0

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            clsNum += CInt(Val(dr("clsNum")))
            entNum += CInt(Val(dr("entNum")))
            A1 += CInt(Val(dr("A1")))
            A2 += CInt(Val(dr("A2")))
            A += CInt(Val(dr("A")))

            B1 += CInt(Val(dr("B1")))
            B2 += CInt(Val(dr("B2")))
            B += CInt(Val(dr("B")))
            'B3 += CInt(Val(dr("B3")))
            'B4 += CInt(Val(dr("B4")))
            'B5 += CInt(Val(dr("B5")))
            D1 += CInt(Val(dr("D1")))
            D2 += CInt(Val(dr("D2")))
            D3 += CInt(Val(dr("D3")))
            D4 += CInt(Val(dr("D4")))
            D += CInt(Val(dr("D")))
            E1 += CInt(Val(dr("E1")))
            E2 += CInt(Val(dr("E2")))
            E3 += CInt(Val(dr("E3")))
            E4 += CInt(Val(dr("E4")))
            E5 += CInt(Val(dr("E5")))
            E6 += CInt(Val(dr("E6")))
            E7 += CInt(Val(dr("E7")))
            E8 += CInt(Val(dr("E8")))
            E9 += CInt(Val(dr("E9")))
            E10 += CInt(Val(dr("E10")))
            E11 += CInt(Val(dr("E11")))
            E12 += CInt(Val(dr("E12")))
            E13 += CInt(Val(dr("E13")))
            E14 += CInt(Val(dr("E14")))
            E15 += CInt(Val(dr("E15")))
            E16 += CInt(Val(dr("E16")))
            E17 += CInt(Val(dr("E17")))
            E18 += CInt(Val(dr("E18")))
            E99 += CInt(Val(dr("E99")))
            E += CInt(Val(dr("E")))
            F1 += CInt(Val(dr("F1")))
            F2 += CInt(Val(dr("F2")))
            F3 += CInt(Val(dr("F3")))
            F4 += CInt(Val(dr("F4")))
            F5 += CInt(Val(dr("F5")))
            F6 += CInt(Val(dr("F6")))
            F7 += CInt(Val(dr("F7")))
            F8 += CInt(Val(dr("F8")))
            F99 += CInt(Val(dr("F99")))
            F += CInt(Val(dr("F")))
            Aa1 += CInt(Val(dr("Aa1")))
            Aa2 += CInt(Val(dr("Aa2")))
            Aa += CInt(Val(dr("Aa")))
            Ba1 += CInt(Val(dr("Ba1")))
            Ba2 += CInt(Val(dr("Ba2")))
            Ba += CInt(Val(dr("Ba")))
            'BA3 += CInt(Val(dr("BA3")))
            'BA4 += CInt(Val(dr("BA4")))
            'BA5 += CInt(Val(dr("BA5")))
            Da1 += CInt(Val(dr("Da1")))
            Da2 += CInt(Val(dr("Da2")))
            Da3 += CInt(Val(dr("Da3")))
            Da4 += CInt(Val(dr("Da4")))
            Da += CInt(Val(dr("Da")))
            Ea1 += CInt(Val(dr("Ea1")))
            Ea2 += CInt(Val(dr("Ea2")))
            Ea3 += CInt(Val(dr("Ea3")))
            Ea4 += CInt(Val(dr("Ea4")))
            Ea5 += CInt(Val(dr("Ea5")))
            Ea6 += CInt(Val(dr("Ea6")))
            Ea7 += CInt(Val(dr("Ea7")))
            Ea8 += CInt(Val(dr("Ea8")))
            Ea9 += CInt(Val(dr("Ea9")))
            Ea10 += CInt(Val(dr("Ea10")))
            Ea11 += CInt(Val(dr("Ea11")))
            Ea12 += CInt(Val(dr("Ea12")))
            Ea13 += CInt(Val(dr("Ea13")))
            Ea14 += CInt(Val(dr("Ea14")))
            Ea15 += CInt(Val(dr("Ea15")))
            Ea16 += CInt(Val(dr("Ea16")))
            Ea17 += CInt(Val(dr("Ea17")))
            Ea18 += CInt(Val(dr("Ea18")))
            Ea99 += CInt(Val(dr("Ea99")))
            Ea += CInt(Val(dr("Ea")))
            Fa1 += CInt(Val(dr("Fa1")))
            Fa2 += CInt(Val(dr("Fa2")))
            Fa3 += CInt(Val(dr("Fa3")))
            Fa4 += CInt(Val(dr("Fa4")))
            Fa5 += CInt(Val(dr("Fa5")))
            Fa6 += CInt(Val(dr("Fa6")))
            Fa7 += CInt(Val(dr("Fa7")))
            Fa8 += CInt(Val(dr("Fa8")))
            Fa99 += CInt(Val(dr("Fa99")))
            Fa += CInt(Val(dr("Fa")))

            'Ab1 += CInt(Val(dr("Ab1")))
            'Ab2 += CInt(Val(dr("Ab2")))
            'Ab += CInt(Val(dr("Ab")))
            'Bb1 += CInt(Val(dr("Bb1")))
            'Bb2 += CInt(Val(dr("Bb2")))
            'Bb += CInt(Val(dr("Bb")))
            'Db1 += CInt(Val(dr("Db1")))
            'Db2 += CInt(Val(dr("Db2")))
            'Db3 += CInt(Val(dr("Db3")))
            'Db4 += CInt(Val(dr("Db4")))
            'Db += CInt(Val(dr("Db")))
            'Eb1 += CInt(Val(dr("Eb1")))
            'Eb2 += CInt(Val(dr("Eb2")))
            'Eb3 += CInt(Val(dr("Eb3")))
            'Eb4 += CInt(Val(dr("Eb4")))
            'Eb5 += CInt(Val(dr("Eb5")))
            'Eb6 += CInt(Val(dr("Eb6")))
            'Eb7 += CInt(Val(dr("Eb7")))
            'Eb8 += CInt(Val(dr("Eb8")))
            'Eb9 += CInt(Val(dr("Eb9")))
            'Eb10 += CInt(Val(dr("Eb10")))
            'Eb11 += CInt(Val(dr("Eb11")))
            'Eb12 += CInt(Val(dr("Eb12")))
            'Eb13 += CInt(Val(dr("Eb13")))
            'Eb14 += CInt(Val(dr("Eb14")))
            'Eb15 += CInt(Val(dr("Eb15")))
            'Eb16 += CInt(Val(dr("Eb16")))
            'Eb17 += CInt(Val(dr("Eb17")))
            'Eb18 += CInt(Val(dr("Eb18")))
            'Eb99 += CInt(Val(dr("Eb99")))
            'Eb += CInt(Val(dr("Eb")))
            'Fb1 += CInt(Val(dr("Fb1")))
            'Fb2 += CInt(Val(dr("Fb2")))
            'Fb3 += CInt(Val(dr("Fb3")))
            'Fb4 += CInt(Val(dr("Fb4")))
            'Fb5 += CInt(Val(dr("Fb5")))
            'Fb6 += CInt(Val(dr("Fb6")))
            'Fb7 += CInt(Val(dr("Fb7")))
            'Fb8 += CInt(Val(dr("Fb8")))
            'Fb99 += CInt(Val(dr("Fb99")))
            'Fb += CInt(Val(dr("Fb")))
            'Gm1 += CInt(Val(dr("Gm1")))
            'Gf1 += CInt(Val(dr("Gf1")))
            'Gt1 += CInt(Val(dr("Gt1")))
            'Gm2 += CInt(Val(dr("Gm2")))
            'Gf2 += CInt(Val(dr("Gf2")))
            'Gt2 += CInt(Val(dr("Gt2")))

            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            'For coli As Integer = 0 To dt.Columns.Count - 1
            '    ExportStr &= "<td>" & Convert.ToString(dr(coli)) & "</td>" & vbTab
            'Next
            Select Case xType
                Case 1 '分署
                    ExportStr &= "<td>" & Convert.ToString(dr("distname")) & "</td>" & vbTab
                Case 2 '計畫別
                    ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab
                Case 3 '分署'計畫別
                    ExportStr &= "<td>" & Convert.ToString(dr("distname")) & "</td>" & vbTab
                    ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab
            End Select
            ExportStr &= "<td>" & Convert.ToString(dr("clsNum")) & "</td>" & vbTab '班數
            ExportStr &= "<td>" & Convert.ToString(dr("entNum")) & "</td>" & vbTab '身障報名人數

            ExportStr &= "<td>" & Convert.ToString(dr("A1")) & "</td>" & vbTab '開訓人數-男
            ExportStr &= "<td>" & Convert.ToString(dr("A2")) & "</td>" & vbTab '開訓人數-女
            ExportStr &= "<td>" & Convert.ToString(dr("A")) & "</td>" & vbTab '開訓人數-小計
            '身心障礙身份開訓人數
            ExportStr &= "<td>" & Convert.ToString(dr("B1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B")) & "</td>" & vbTab
            '在職者身心障礙開訓人數		
            'ExportStr &= "<td>" & Convert.ToString(dr("B3")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("B4")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("B5")) & "</td>" & vbTab
            '障礙等級
            ExportStr &= "<td>" & Convert.ToString(dr("D1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("E1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E5")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E6")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E7")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E8")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E9")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E10")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E11")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E12")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E13")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E14")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E15")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E16")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E17")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E18")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E99")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("E")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F5")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F6")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F7")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F8")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F99")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("F")) & "</td>" & vbTab

            '結訓人數																																											
            ExportStr &= "<td>" & Convert.ToString(dr("Aa1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Aa2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Aa")) & "</td>" & vbTab
            '身心障礙身份結訓人數		
            ExportStr &= "<td>" & Convert.ToString(dr("Ba1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ba2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ba")) & "</td>" & vbTab
            '在職者身心障礙結訓人數		
            'ExportStr &= "<td>" & Convert.ToString(dr("BA3")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("BA4")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("BA5")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Da1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Da2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Da3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Da4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Da")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea5")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea6")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea7")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea8")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea9")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea10")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea11")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea12")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea13")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea14")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea15")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea16")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea17")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea18")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea99")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Ea")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa4")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa5")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa6")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa7")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa8")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa99")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("Fa")) & "</td>" & vbTab

            'ExportStr &= "<td>" & Convert.ToString(dr("Ab1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Ab2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Ab")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Bb1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Bb2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Bb")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Db1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Db2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Db3")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Db4")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Db")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb3")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb4")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb5")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb6")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb7")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb8")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb9")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb10")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb11")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb12")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb13")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb14")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb15")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb16")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb17")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb18")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb99")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Eb")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb3")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb4")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb5")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb6")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb7")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb8")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb99")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Fb")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gm1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gf1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gt1")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gm2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gf2")) & "</td>" & vbTab
            'ExportStr &= "<td>" & Convert.ToString(dr("Gt2")) & "</td>" & vbTab
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next

        '建立合計
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        'For coli As Integer = 0 To dt.Columns.Count - 1
        '    ExportStr &= "<td>" & Convert.ToString(dr(coli)) & "</td>" & vbTab
        'Next
        Select Case xType
            Case 1 '分署
                ExportStr &= "<td>合計</td>" & vbTab
            Case 2 '計畫別
                ExportStr &= "<td>合計</td>" & vbTab
            Case 3 '分署'計畫別
                ExportStr &= "<td colspan=2>合計</td>" & vbTab
        End Select
        ExportStr &= "<td>" & Convert.ToString(clsNum) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(entNum) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(A1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(A2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(A) & "</td>" & vbTab

        ExportStr &= "<td>" & Convert.ToString(B1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(B2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(B) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(B3) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(B4) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(B5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(D1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(D2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(D3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(D4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(D) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E6) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E7) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E8) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E9) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E10) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E11) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E12) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E13) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E14) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E15) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E16) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E17) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E18) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E99) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(E) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F6) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F7) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F8) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F99) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(F) & "</td>" & vbTab

        ExportStr &= "<td>" & Convert.ToString(Aa1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Aa2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Aa) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ba1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ba2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ba) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(BA3) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(BA4) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(BA5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Da1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Da2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Da3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Da4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Da) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea6) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea7) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea8) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea9) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea10) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea11) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea12) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea13) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea14) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea15) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea16) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea17) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea18) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea99) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Ea) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa1) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa2) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa3) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa4) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa5) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa6) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa7) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa8) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa99) & "</td>" & vbTab
        ExportStr &= "<td>" & Convert.ToString(Fa) & "</td>" & vbTab

        'ExportStr &= "<td>" & Convert.ToString(Ab1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Ab2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Ab) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Bb1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Bb2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Bb) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Db1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Db2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Db3) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Db4) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Db) & "</td>" & vbTab

        'ExportStr &= "<td>" & Convert.ToString(Eb1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb3) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb4) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb5) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb6) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb7) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb8) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb9) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb10) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb11) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb12) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb13) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb14) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb15) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb16) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb17) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb18) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb99) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Eb) & "</td>" & vbTab

        'ExportStr &= "<td>" & Convert.ToString(Fb1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb3) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb4) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb5) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb6) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb7) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb8) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb99) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Fb) & "</td>" & vbTab

        'ExportStr &= "<td>" & Convert.ToString(Gm1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Gf1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Gt1) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Gm2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Gf2) & "</td>" & vbTab
        'ExportStr &= "<td>" & Convert.ToString(Gt2) & "</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

#Region "(No Use)"

    ''計畫別 組合EXCEL
    'Sub ExpReport2(ByRef dt As DataTable)
    '    Dim strTitle1 As String = "身心障礙統計表"
    '    Dim dtType1 As DataTable
    '    Dim dtType2 As DataTable
    '    Dim Sql As String = ""
    '    Sql = "SELECT HandTypeID , Name FROM Key_HandicatType WHERE HandTypeID!='00' ORDER BY HandTypeID "
    '    dtType1 = DbAccess.GetDataTable(Sql, objconn)
    '    Sql = "SELECT HandTypeID2 , Name FROM Key_HandicatType2 ORDER BY HandTypeID2 "
    '    dtType2 = DbAccess.GetDataTable(Sql, objconn)

    '    Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8) & ".xls")
    '    'Response.ContentType = "Application/octet-stream"
    '    Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
    '    'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
    '    'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
    '    '文件內容指定為Excel
    '    'Response.ContentType = "application/ms-excel;charset=utf-8"
    '    Response.ContentType = "application/ms-excel"
    '    'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
    '    Common.RespWrite(Me, "<html>")
    '    Common.RespWrite(Me, "<head>")
    '    Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
    '    '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>

    '    ''套CSS值
    '    'Common.RespWrite(Me, "<style>")
    '    'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
    '    'Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
    '    ''mso-number-format:"0" 
    '    'Common.RespWrite(Me, "</style>")
    '    Common.RespWrite(Me, "</head>")
    '    Common.RespWrite(Me, "<body>")
    '    Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

    '    Dim ExportStr As String = ""

    '    '建立抬頭
    '    '第1行
    '    ExportStr = ""
    '    ExportStr &= "<tr>" & vbCrLf
    '    ExportStr &= "<td rowspan=""2"">計畫</td>" & vbTab
    '    ExportStr &= "<td rowspan=""2"">班數</td>" & vbTab
    '    ExportStr &= "<td colspan=""3"">開訓人數</td>" & vbTab
    '    ExportStr &= "<td colspan=""3"">身心障礙身份開訓人數</td>" & vbTab
    '    ExportStr &= "<td colspan=""5"">障礙等級</td>" & vbTab
    '    ExportStr &= "<td colspan=""20"">舊制障礙等級</td>" & vbTab
    '    ExportStr &= "<td colspan=""10"">新制障礙等級</td>" & vbTab
    '    ExportStr &= "</tr>" & vbCrLf

    '    '第2行
    '    ExportStr &= "<tr>" & vbCrLf
    '    ExportStr &= "<td>男</td>" & vbTab
    '    ExportStr &= "<td>女</td>" & vbTab
    '    ExportStr &= "<td>小計</td>" & vbTab
    '    ExportStr &= "<td>男</td>" & vbTab
    '    ExportStr &= "<td>女</td>" & vbTab
    '    ExportStr &= "<td>小計</td>" & vbTab

    '    ExportStr &= "<td>輕度</td>" & vbTab
    '    ExportStr &= "<td>中度</td>" & vbTab
    '    ExportStr &= "<td>重度</td>" & vbTab
    '    ExportStr &= "<td>極重度</td>" & vbTab
    '    ExportStr &= "<td>小計</td>" & vbTab

    '    For i As Integer = 0 To dtType1.Rows.Count - 1
    '        ExportStr &= "<td>" & dtType1.Rows(i)("Name") & "</td>" & vbTab
    '    Next
    '    ExportStr &= "<td>其他</td>" & vbTab
    '    ExportStr &= "<td>小計</td>" & vbTab

    '    For i As Integer = 0 To dtType2.Rows.Count - 1
    '        ExportStr &= "<td>" & dtType2.Rows(i)("Name") & "</td>" & vbTab
    '    Next
    '    'ExportStr &= "<td>其他</td>" & vbTab
    '    ExportStr &= "<td>小計</td>" & vbTab

    '    ExportStr &= "</tr>" & vbCrLf
    '    Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

    '    'Dim distname As Integer = 0
    '    Dim clsNum As Integer = 0
    '    Dim A1 As Integer = 0
    '    Dim A2 As Integer = 0
    '    Dim A As Integer = 0
    '    Dim B1 As Integer = 0
    '    Dim B2 As Integer = 0
    '    Dim B As Integer = 0
    '    Dim D1 As Integer = 0
    '    Dim D2 As Integer = 0
    '    Dim D3 As Integer = 0
    '    Dim D4 As Integer = 0
    '    Dim D As Integer = 0
    '    Dim E1 As Integer = 0
    '    Dim E2 As Integer = 0
    '    Dim E3 As Integer = 0
    '    Dim E4 As Integer = 0
    '    Dim E5 As Integer = 0
    '    Dim E6 As Integer = 0
    '    Dim E7 As Integer = 0
    '    Dim E8 As Integer = 0
    '    Dim E9 As Integer = 0
    '    Dim E10 As Integer = 0
    '    Dim E11 As Integer = 0
    '    Dim E12 As Integer = 0
    '    Dim E13 As Integer = 0
    '    Dim E14 As Integer = 0
    '    Dim E15 As Integer = 0
    '    Dim E16 As Integer = 0
    '    Dim E17 As Integer = 0
    '    Dim E18 As Integer = 0
    '    Dim E99 As Integer = 0
    '    Dim E As Integer = 0
    '    Dim F1 As Integer = 0
    '    Dim F2 As Integer = 0
    '    Dim F3 As Integer = 0
    '    Dim F4 As Integer = 0
    '    Dim F5 As Integer = 0
    '    Dim F6 As Integer = 0
    '    Dim F7 As Integer = 0
    '    Dim F8 As Integer = 0
    '    Dim F9 As Integer = 0
    '    Dim F As Integer = 0

    '    Dim iSeqno As Integer = 0
    '    For Each dr As DataRow In dt.Rows
    '        clsNum += CInt(Val(dr("clsNum")))
    '        A1 += CInt(Val(dr("A1")))
    '        A2 += CInt(Val(dr("A2")))
    '        A += CInt(Val(dr("A")))
    '        B1 += CInt(Val(dr("B1")))
    '        B2 += CInt(Val(dr("B2")))
    '        B += CInt(Val(dr("B")))
    '        D1 += CInt(Val(dr("D1")))
    '        D2 += CInt(Val(dr("D2")))
    '        D3 += CInt(Val(dr("D3")))
    '        D4 += CInt(Val(dr("D4")))
    '        D += CInt(Val(dr("D")))
    '        E1 += CInt(Val(dr("E1")))
    '        E2 += CInt(Val(dr("E2")))
    '        E3 += CInt(Val(dr("E3")))
    '        E4 += CInt(Val(dr("E4")))
    '        E5 += CInt(Val(dr("E5")))
    '        E6 += CInt(Val(dr("E6")))
    '        E7 += CInt(Val(dr("E7")))
    '        E8 += CInt(Val(dr("E8")))
    '        E9 += CInt(Val(dr("E9")))
    '        E10 += CInt(Val(dr("E10")))
    '        E11 += CInt(Val(dr("E11")))
    '        E12 += CInt(Val(dr("E12")))
    '        E13 += CInt(Val(dr("E13")))
    '        E14 += CInt(Val(dr("E14")))
    '        E15 += CInt(Val(dr("E15")))
    '        E16 += CInt(Val(dr("E16")))
    '        E17 += CInt(Val(dr("E17")))
    '        E18 += CInt(Val(dr("E18")))
    '        E99 += CInt(Val(dr("E99")))
    '        E += CInt(Val(dr("E")))
    '        F1 += CInt(Val(dr("F1")))
    '        F2 += CInt(Val(dr("F2")))
    '        F3 += CInt(Val(dr("F3")))
    '        F4 += CInt(Val(dr("F4")))
    '        F5 += CInt(Val(dr("F5")))
    '        F6 += CInt(Val(dr("F6")))
    '        F7 += CInt(Val(dr("F7")))
    '        F8 += CInt(Val(dr("F8")))
    '        F9 += CInt(Val(dr("F9")))
    '        F += CInt(Val(dr("F")))
    '        A1 += CInt(Val(dr("A1")))
    '        A2 += CInt(Val(dr("A2")))
    '        A += CInt(Val(dr("A")))
    '        B1 += CInt(Val(dr("B1")))
    '        B2 += CInt(Val(dr("B2")))
    '        B += CInt(Val(dr("B")))
    '        D1 += CInt(Val(dr("D1")))
    '        D2 += CInt(Val(dr("D2")))
    '        D3 += CInt(Val(dr("D3")))
    '        D4 += CInt(Val(dr("D4")))
    '        D += CInt(Val(dr("D")))
    '        E1 += CInt(Val(dr("E1")))
    '        E2 += CInt(Val(dr("E2")))
    '        E3 += CInt(Val(dr("E3")))
    '        E4 += CInt(Val(dr("E4")))
    '        E5 += CInt(Val(dr("E5")))
    '        E6 += CInt(Val(dr("E6")))
    '        E7 += CInt(Val(dr("E7")))
    '        E8 += CInt(Val(dr("E8")))
    '        E9 += CInt(Val(dr("E9")))
    '        E10 += CInt(Val(dr("E10")))
    '        E11 += CInt(Val(dr("E11")))
    '        E12 += CInt(Val(dr("E12")))
    '        E13 += CInt(Val(dr("E13")))
    '        E14 += CInt(Val(dr("E14")))
    '        E15 += CInt(Val(dr("E15")))
    '        E16 += CInt(Val(dr("E16")))
    '        E17 += CInt(Val(dr("E17")))
    '        E18 += CInt(Val(dr("E18")))
    '        E99 += CInt(Val(dr("E99")))
    '        E += CInt(Val(dr("E")))
    '        F1 += CInt(Val(dr("F1")))
    '        F2 += CInt(Val(dr("F2")))
    '        F3 += CInt(Val(dr("F3")))
    '        F4 += CInt(Val(dr("F4")))
    '        F5 += CInt(Val(dr("F5")))
    '        F6 += CInt(Val(dr("F6")))
    '        F7 += CInt(Val(dr("F7")))
    '        F8 += CInt(Val(dr("F8")))
    '        F9 += CInt(Val(dr("F9")))
    '        F += CInt(Val(dr("F")))

    '        '序號+1
    '        iSeqno += 1
    '        '建立資料面
    '        ExportStr = ""
    '        ExportStr &= "<tr>" & vbCrLf
    '        'For coli As Integer = 0 To dt.Columns.Count - 1
    '        '    ExportStr &= "<td>" & Convert.ToString(dr(coli)) & "</td>" & vbTab
    '        'Next
    '        ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("clsNum")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("A1")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("A2")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("A")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("B1")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("B2")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("B")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("D1")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("D2")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("D3")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("D4")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("D")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E1")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E2")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E3")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E4")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E5")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E6")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E7")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E8")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E9")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E10")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E11")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E12")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E13")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E14")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E15")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E16")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E17")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E18")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E99")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("E")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F1")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F2")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F3")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F4")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F5")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F6")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F7")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F8")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F9")) & "</td>" & vbTab
    '        ExportStr &= "<td>" & Convert.ToString(dr("F")) & "</td>" & vbTab
    '        ExportStr &= "</tr>" & vbCrLf
    '        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
    '    Next

    '    '建立合計
    '    ExportStr = ""
    '    ExportStr &= "<tr>" & vbCrLf
    '    'For coli As Integer = 0 To dt.Columns.Count - 1
    '    '    ExportStr &= "<td>" & Convert.ToString(dr(coli)) & "</td>" & vbTab
    '    'Next
    '    ExportStr &= "<td>合計</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(clsNum) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(A1) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(A2) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(A) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(B1) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(B2) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(B) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(D1) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(D2) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(D3) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(D4) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(D) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E1) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E2) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E3) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E4) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E5) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E6) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E7) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E8) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E9) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E10) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E11) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E12) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E13) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E14) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E15) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E16) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E17) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E18) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E99) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(E) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F1) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F2) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F3) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F4) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F5) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F6) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F7) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F8) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F9) & "</td>" & vbTab
    '    ExportStr &= "<td>" & Convert.ToString(F) & "</td>" & vbTab
    '    ExportStr &= "</tr>" & vbCrLf
    '    Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
    '    Common.RespWrite(Me, "</table>")
    '    Common.RespWrite(Me, "</body>")
    '    Response.End()
    'End Sub

#End Region
End Class