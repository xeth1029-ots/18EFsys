Partial Class SD_15_019
    Inherits AuthBasePage

    'Const cst_StrTitle1 As String = "重複參訓統計表"
    Const cst_colspanT2xNum As Integer = 13 '總匯出欄位
    Const cst_StrTitle2 As String = "重複參訓統計表-明細"

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
        Call TIMS.OpenDbConn(objconn)
        'aNow = TIMS.GetSysDateNow(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call create1()
        End If

        'btnPrint1.Attributes("onclick") = "return CheckPrint();"
    End Sub

    Sub create1()
        OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
        Common.SetListItem(OrgKind2, "G")

        ddlYears = TIMS.GetSyear(ddlYears)
        Common.SetListItem(ddlYears, sm.UserInfo.Years)

        'Distid = TIMS.Get_DistID(Distid)
        'Distid.Items.Insert(0, New ListItem("全部", 0))
        Call TIMS.Get_DISTCBL(Distid, objconn)
        Distid.Attributes("onclick") = "SelectAll('Distid','DistHidden');"

        'Dim dt As DataTable
        'Dim sql As String = ""
        'sql = "SELECT DISTID,NAME FROM ID_DISTRICT ORDER BY DISTID"
        'dt = DbAccess.GetDataTable(sql, objconn)
        'ddlDistID = TIMS.Get_DistID(ddlDistID, dt)

        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0"
            Case "1"
                Common.SetListItem(Distid, sm.UserInfo.DistID)
                Distid.Enabled = False
            Case "2"
                Common.SetListItem(ddlYears, sm.UserInfo.Years)
                ddlYears.Enabled = False
                Common.SetListItem(Distid, sm.UserInfo.DistID)
                Distid.Enabled = False
        End Select

        'Distid.Enabled = True
        If sm.UserInfo.DistID <> "000" Then '若登入者非署(局)署，鎖定轄區
            Common.SetListItem(Distid, sm.UserInfo.DistID)
            Distid.Enabled = False
        End If

        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '沒有不區分
        '    trPlanKind.Visible = True
        '    OrgKind2 = TIMS.Get_RblOrgKind2(Me, OrgKind2)
        '    OrgKind2.Items.Insert(0, New ListItem("全部", "A")) 'A/G/W
        '    Common.SetListItem(OrgKind2, "A")
        'End If
    End Sub

    '班級範圍基底 SELECT
    Function get_WC1SQL() As String
        Dim sql As String = ""

        Dim sYear1 As String = TIMS.ClearSQM(ddlYears.SelectedValue)
        If sYear1 = "" Then sYear1 = TIMS.ClearSQM(sm.UserInfo.Years)
        '轄區
        Dim sDistID As String = TIMS.GetCblValueIn(Distid)
        '轄區Dim sDistID As String = TIMS.ClearSQM(Distid.SelectedValue)

        sql = ""
        sql &= " SELECT case when ip.TPLANID='28' then ip.PlanName+'('+v2.ORGPLANNAME+')'" & vbCrLf
        sql &= " else ip.PlanName end plankind" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,ip.years" & vbCrLf
        sql &= " ,ip.distname" & vbCrLf
        sql &= " ,v2.ORGPLANNAME" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,oo.comidno" & vbCrLf
        sql &= " ,ky.name OrgTypeName1" & vbCrLf
        'sql &= " /*單位屬性*/" & vbCrLf
        sql &= " ,nvl2(o1.typeid2,o1.typeid2 + '-' + o1.typeid2name,null) OrgTypeName" & vbCrLf
        sql &= " ,DECODE(dd.APPRESULT,'Y',dd.kname12) KNAME12" & vbCrLf
        'sql &= " /*訓練業別*/" & vbCrLf
        sql &= " ,ig2.GCODE1,ig2.GCODE2,ig2.CNAME" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,COALESCE(iz.CTName,iz1.CTName,iz2.CTName) CTName" & vbCrLf
        sql &= " FROM Class_ClassInfo cc" & vbCrLf
        sql &= " JOIN Plan_PlanInfo pp ON cc.PlanID = pp.PlanID AND cc.ComIDNO = pp.ComIDNO AND cc.SeqNO = pp.SeqNO" & vbCrLf
        sql &= " JOIN ID_CLASS ic on ic.clsid =cc.clsid" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=cc.planid" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.comidno=cc.comidno" & vbCrLf
        sql &= " JOIN KEY_ORGTYPE ky on ky.orgtypeid=oo.OrgKind" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME v2 ON v2.RID=pp.RID" & vbCrLf
        sql &= " JOIN KEY_ORGTYPE1 o1 on o1.OrgTypeID1=oo.OrgKind1" & vbCrLf
        sql &= " JOIN V_GOVCLASSCAST2 ig2 on ig2.GCID2 = pp.GCID2" & vbCrLf
        sql &= " LEFT JOIN view_ZipName iz ON iz.ZipCode = cc.TaddressZip" & vbCrLf
        'sql &= " /* 產投上課地址學科場地代碼 */" & vbCrLf
        sql &= " LEFT JOIN Plan_TrainPlace sp on sp.PTID=pp.AddressSciPTID" & vbCrLf
        sql &= " LEFT JOIN view_zipName iz1 on iz1.zipCode=sp.ZipCode" & vbCrLf
        'sql &= " /* 產投上課地址術科場地代碼 */" & vbCrLf
        sql &= " LEFT JOIN Plan_TrainPlace tp on tp.PTID=pp.AddressTechPTID" & vbCrLf
        sql &= " LEFT JOIN view_zipName iz2 on iz2.zipCode=tp.ZipCode" & vbCrLf
        sql &= " LEFT JOIN V_PLAN_DEPOT dd on dd.planid=pp.planid and dd.comidno=pp.comidno and dd.seqno=pp.seqno" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        sql &= " and ip.TPLANID ='28'" & vbCrLf
        'sql &= " and ip.years ='2016'" & vbCrLf
        'sql &= " and ip.distid ='001'" & vbCrLf
        '是產投計畫而且有打開計畫別功能
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
            AndAlso trPlanKind.Visible Then
            Select Case OrgKind2.SelectedValue 'A/G/W
                Case "A"
                Case "G", "W"
                    sql += " and v2.OrgKind2= '" & OrgKind2.SelectedValue & "'" & vbCrLf
                Case Else
                    sql += " and 1<>1" & vbCrLf
            End Select
        End If
        If sYear1 <> "" Then
            sql &= " AND ip.years='" & sYear1 & "'" & vbCrLf
        End If
        If sDistID <> "" Then
            'sql &= " AND ip.distid='" & sDistID & "'" & vbCrLf
            sql &= " AND ip.DISTID IN (" & sDistID & ")" & vbCrLf
        End If
        If STDate1.Text <> "" Then
            sql &= " AND cc.stdate >=" & TIMS.To_date(STDate1.Text) & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql &= " AND cc.stdate <=" & TIMS.To_date(STDate2.Text) & vbCrLf
        End If

        Return sql
    End Function

    '依班級學員基底 SELECT
    Function get_WC2SQL() As String
        Dim sql As String = ""
        sql &= " SELECT cc.plankind" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.years" & vbCrLf
        sql &= " ,cc.DISTNAME" & vbCrLf
        sql &= " ,cc.ORGPLANNAME" & vbCrLf
        sql &= " ,cc.ORGNAME" & vbCrLf
        sql &= " ,cc.comidno" & vbCrLf
        sql &= " ,cc.OrgTypeName" & vbCrLf 'sql &= " /*單位屬性*/" & vbCrLf
        sql &= " ,cc.KNAME12" & vbCrLf
        sql &= " ,cc.GCODE2" & vbCrLf 'sql &= " /*訓練業別*/" & vbCrLf
        sql &= " ,cc.CNAME" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,cc.CTNAME" & vbCrLf
        sql &= " ,ss.IDNO" & vbCrLf
        sql &= " ,dbo.NVL(sc.SumOfMoney,0) SumOfMoney" & vbCrLf
        sql &= " ,ss.SOCID" & vbCrLf
        sql &= " ,dbo.NVL(sc.AppliedStatus,0) AppliedStatus" & vbCrLf
        sql &= " FROM V_STUDENTINFO ss" & vbCrLf
        sql &= " JOIN STUD_SUBSIDYCOST sc on sc.socid=ss.socid and sc.AppliedStatusM='Y'" & vbCrLf
        sql &= " JOIN WC1 cc on cc.ocid =ss.ocid" & vbCrLf
        Return sql
    End Function

    '相同業別(重複)('總數/領取/未領取) GROUP BY
    Function get_WC3SQL() As String
        Dim sql As String = ""
        sql &= " SELECT cc.IDNO" & vbCrLf
        sql &= " ,cc.GCODE2" & vbCrLf
        sql &= " ,COUNT(1) CNT" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus=1 THEN 1 END) CNT2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus!=1 THEN 1 END) CNT3" & vbCrLf
        sql &= " FROM WC2 cc" & vbCrLf
        sql &= " GROUP BY cc.IDNO" & vbCrLf
        sql &= " ,cc.GCODE2" & vbCrLf
        sql &= " HAVING COUNT(1)>1" & vbCrLf
        Return sql
    End Function

    '業別未重複('總數/領取/未領取) GROUP BY (統計用)
    Function get_WC4NSQL() As String
        Dim sql As String = ""
        sql &= " SELECT cc.IDNO" & vbCrLf
        sql &= " ,COUNT(1) CNT" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus=1 THEN 1 END) CNT1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus!=1 THEN 1 END) CNT2" & vbCrLf
        sql &= " FROM WC2 cc" & vbCrLf
        sql &= " GROUP BY cc.IDNO" & vbCrLf
        Return sql
    End Function

    '業別重複 '總數/領取/未領取 GROUP BY (統計用)
    Function get_WC4YSQL() As String
        Dim sql As String = ""
        sql &= " SELECT cc.IDNO" & vbCrLf
        sql &= " ,cc.GCODE2" & vbCrLf
        sql &= " ,COUNT(1) CNT" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus=1 THEN 1 END) CNT1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cc.AppliedStatus!=1 THEN 1 END) CNT2" & vbCrLf
        sql &= " FROM WC2 cc" & vbCrLf
        sql &= " GROUP BY cc.IDNO" & vbCrLf
        sql &= " ,cc.GCODE2" & vbCrLf
        'sql &= " HAVING COUNT(1)>1" & vbCrLf
        Return sql
    End Function

    '年度次數統計
    Function get_WC40SQL() As String
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 0 NUMSORT" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=1 THEN 1 END) CNT0" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=2 THEN 1 END) CNT1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=3 THEN 1 END) CNT2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=4 THEN 1 END) CNT3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=5 THEN 1 END) CNT4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)=6 THEN 1 END) CNT5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>=1 AND dbo.NVL(CNT,0)<=6 THEN 1 END) CNT05" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>6  AND dbo.NVL(CNT,0)<=11 THEN 1 END) CNT10" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>11 AND dbo.NVL(CNT,0)<=16 THEN 1 END) CNT15" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>16 AND dbo.NVL(CNT,0)<=21 THEN 1 END) CNT20" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>21 AND dbo.NVL(CNT,0)<=26 THEN 1 END) CNT25" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>26 AND dbo.NVL(CNT,0)<=31 THEN 1 END) CNT30" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>31 AND dbo.NVL(CNT,0)<=36 THEN 1 END) CNT35" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>36 AND dbo.NVL(CNT,0)<=41 THEN 1 END) CNT40" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>41 AND dbo.NVL(CNT,0)<=46 THEN 1 END) CNT45" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>46 AND dbo.NVL(CNT,0)<=51 THEN 1 END) CNT50" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>51 THEN 1 END) CNT51" & vbCrLf
        sql &= " ,COUNT(CASE WHEN dbo.NVL(CNT,0)>=1 THEN 1 END) CNTALL" & vbCrLf
        sql &= " FROM WC4 cc" & vbCrLf
        Return sql
    End Function

    '年度次數統計 (領取)
    Function get_WC41SQL() As String
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 1 NUMSORT" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=1 THEN dbo.NVL(CNT1,0) END) CNT0" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=2 THEN dbo.NVL(CNT1,0) END) CNT1" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=3 THEN dbo.NVL(CNT1,0) END) CNT2" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=4 THEN dbo.NVL(CNT1,0) END) CNT3" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=5 THEN dbo.NVL(CNT1,0) END) CNT4" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=6 THEN dbo.NVL(CNT1,0) END) CNT5" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>=1 AND dbo.NVL(CNT,0)<=6 THEN dbo.NVL(CNT1,0) END) CNT05" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>6  AND dbo.NVL(CNT,0)<=11 THEN dbo.NVL(CNT1,0) END) CNT10" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>11 AND dbo.NVL(CNT,0)<=16 THEN dbo.NVL(CNT1,0) END) CNT15" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>16 AND dbo.NVL(CNT,0)<=21 THEN dbo.NVL(CNT1,0) END) CNT20" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>21 AND dbo.NVL(CNT,0)<=26 THEN dbo.NVL(CNT1,0) END) CNT25" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>26 AND dbo.NVL(CNT,0)<=31 THEN dbo.NVL(CNT1,0) END) CNT30" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>31 AND dbo.NVL(CNT,0)<=36 THEN dbo.NVL(CNT1,0) END) CNT35" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>36 AND dbo.NVL(CNT,0)<=41 THEN dbo.NVL(CNT1,0) END) CNT40" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>41 AND dbo.NVL(CNT,0)<=46 THEN dbo.NVL(CNT1,0) END) CNT45" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>46 AND dbo.NVL(CNT,0)<=51 THEN dbo.NVL(CNT1,0) END) CNT50" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>51 THEN dbo.NVL(CNT1,0) END) CNT51" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>=1 THEN dbo.NVL(CNT1,0) END) CNTALL" & vbCrLf
        sql &= " FROM WC4 cc" & vbCrLf
        Return sql
    End Function

    '年度次數統計 (未領取)
    Function get_WC42SQL() As String
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT 2 NUMSORT" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=1 THEN dbo.NVL(CNT2,0) END) CNT0" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=2 THEN dbo.NVL(CNT2,0) END) CNT1" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=3 THEN dbo.NVL(CNT2,0) END) CNT2" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=4 THEN dbo.NVL(CNT2,0) END) CNT3" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=5 THEN dbo.NVL(CNT2,0) END) CNT4" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)=6 THEN dbo.NVL(CNT2,0) END) CNT5" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>=1 AND dbo.NVL(CNT,0)<=6 THEN dbo.NVL(CNT2,0) END) CNT05" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>6  AND dbo.NVL(CNT,0)<=11 THEN dbo.NVL(CNT2,0) END) CNT10" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>11 AND dbo.NVL(CNT,0)<=16 THEN dbo.NVL(CNT2,0) END) CNT15" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>16 AND dbo.NVL(CNT,0)<=21 THEN dbo.NVL(CNT2,0) END) CNT20" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>21 AND dbo.NVL(CNT,0)<=26 THEN dbo.NVL(CNT2,0) END) CNT25" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>26 AND dbo.NVL(CNT,0)<=31 THEN dbo.NVL(CNT2,0) END) CNT30" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>31 AND dbo.NVL(CNT,0)<=36 THEN dbo.NVL(CNT2,0) END) CNT35" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>36 AND dbo.NVL(CNT,0)<=41 THEN dbo.NVL(CNT2,0) END) CNT40" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>41 AND dbo.NVL(CNT,0)<=46 THEN dbo.NVL(CNT2,0) END) CNT45" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>46 AND dbo.NVL(CNT,0)<=51 THEN dbo.NVL(CNT2,0) END) CNT50" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>51 THEN dbo.NVL(CNT2,0) END) CNT51" & vbCrLf
        sql &= " ,SUM(CASE WHEN dbo.NVL(CNT,0)>=1 THEN dbo.NVL(CNT2,0) END) CNTALL" & vbCrLf
        sql &= " FROM WC4 cc" & vbCrLf
        Return sql
    End Function

    '[SQL]
    Function Search2dt(ByVal V_sameJOB1 As String, ByVal V_expType1 As String) As DataTable
        'V_sameJOB1(重複相同/相同訓練業別): N否 Y是
        'V_expType1(產出格式) 1:統計表 /2:明細表
        Dim dt As New DataTable
        Dim sWC1 As String = get_WC1SQL()
        Dim sWC2 As String = get_WC2SQL()
        Dim sWC3 As String = get_WC3SQL()
        Dim sWC4Y As String = get_WC4YSQL()
        Dim sWC4N As String = get_WC4NSQL()
        Dim sWC40 As String = get_WC40SQL()
        Dim sWC41 As String = get_WC41SQL()
        Dim sWC42 As String = get_WC42SQL()

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & sWC1 & ")" & vbCrLf
        sql &= " ,WC2 AS (" & sWC2 & ")" & vbCrLf
        Select Case V_sameJOB1 & V_expType1
            Case "N1" 'V_sameJOB1(相同): N否 Y是'1:統計表 /2:明細表
                sql &= " ,WC4 AS (" & sWC4N & ")" & vbCrLf
                sql &= " ,WC40 AS (" & sWC40 & ")" & vbCrLf
                sql &= " ,WC41 AS (" & sWC41 & ")" & vbCrLf
                sql &= " ,WC42 AS (" & sWC42 & ")" & vbCrLf
                sql &= " SELECT cc.NUMSORT" & vbCrLf
                sql &= " ,cc.CNT0" & vbCrLf
                sql &= " ,cc.CNT1" & vbCrLf
                sql &= " ,cc.CNT2" & vbCrLf
                sql &= " ,cc.CNT3" & vbCrLf
                sql &= " ,cc.CNT4" & vbCrLf
                sql &= " ,cc.CNT5" & vbCrLf
                sql &= " ,cc.CNT05" & vbCrLf
                sql &= " ,cc.CNT10" & vbCrLf
                sql &= " ,cc.CNT15" & vbCrLf
                sql &= " ,cc.CNT20" & vbCrLf
                sql &= " ,cc.CNT25" & vbCrLf
                sql &= " ,cc.CNT30" & vbCrLf
                sql &= " ,cc.CNT35" & vbCrLf
                sql &= " ,cc.CNT40" & vbCrLf
                sql &= " ,cc.CNT45" & vbCrLf
                sql &= " ,cc.CNT50" & vbCrLf
                sql &= " ,cc.CNT51" & vbCrLf
                sql &= " ,cc.CNTALL" & vbCrLf
                sql &= " FROM (" & vbCrLf
                sql &= " SELECT * FROM WC40" & vbCrLf
                sql &= " UNION SELECT * FROM WC41" & vbCrLf
                sql &= " UNION SELECT * FROM WC42" & vbCrLf
                sql &= " ) cc" & vbCrLf
                'sql &= " CROSS JOIN (SELECT COUNT(1) CNTALL FROM WC2) c2" & vbCrLf
                sql &= " ORDER BY cc.NUMSORT" & vbCrLf

            Case "Y1" 'V_sameJOB1(相同): N否 Y是'1:統計表 /2:明細表
                sql &= " ,WC4 AS (" & sWC4Y & ")" & vbCrLf
                sql &= " ,WC40 AS (" & sWC40 & ")" & vbCrLf
                sql &= " ,WC41 AS (" & sWC41 & ")" & vbCrLf
                sql &= " ,WC42 AS (" & sWC42 & ")" & vbCrLf
                sql &= " SELECT cc.NUMSORT" & vbCrLf
                sql &= " ,cc.CNT0" & vbCrLf
                sql &= " ,cc.CNT1" & vbCrLf
                sql &= " ,cc.CNT2" & vbCrLf
                sql &= " ,cc.CNT3" & vbCrLf
                sql &= " ,cc.CNT4" & vbCrLf
                sql &= " ,cc.CNT5" & vbCrLf
                sql &= " ,cc.CNT05" & vbCrLf
                sql &= " ,cc.CNT10" & vbCrLf
                sql &= " ,cc.CNT15" & vbCrLf
                sql &= " ,cc.CNT20" & vbCrLf
                sql &= " ,cc.CNT25" & vbCrLf
                sql &= " ,cc.CNT30" & vbCrLf
                sql &= " ,cc.CNT35" & vbCrLf
                sql &= " ,cc.CNT40" & vbCrLf
                sql &= " ,cc.CNT45" & vbCrLf
                sql &= " ,cc.CNT50" & vbCrLf
                sql &= " ,cc.CNT51" & vbCrLf
                sql &= " ,cc.CNTALL" & vbCrLf
                sql &= " FROM (" & vbCrLf
                sql &= " SELECT * FROM WC40" & vbCrLf
                sql &= " UNION SELECT * FROM WC41" & vbCrLf
                sql &= " UNION SELECT * FROM WC42" & vbCrLf
                sql &= " ) cc" & vbCrLf
                'sql &= " CROSS JOIN (SELECT COUNT(1) CNTALL FROM WC2) c2" & vbCrLf
                sql &= " ORDER BY cc.NUMSORT" & vbCrLf

            Case "N2" 'V_sameJOB1(相同): N否 Y是'1:統計表 /2:明細表
                sql &= " SELECT cc.plankind" & vbCrLf
                sql &= " ,cc.OCID" & vbCrLf
                sql &= " ,cc.years" & vbCrLf
                sql &= " ,cc.distname" & vbCrLf
                sql &= " ,cc.ORGPLANNAME" & vbCrLf
                sql &= " ,cc.orgname" & vbCrLf
                sql &= " ,cc.comidno" & vbCrLf
                sql &= " ,cc.OrgTypeName" & vbCrLf 'sql &= "/*單位屬性*/" & vbCrLf
                sql &= " ,cc.KNAME12" & vbCrLf
                sql &= " ,cc.GCODE2" & vbCrLf 'sql &= "/*訓練業別*/" & vbCrLf
                sql &= " ,cc.CNAME" & vbCrLf
                sql &= " ,cc.classcname" & vbCrLf
                sql &= " ,cc.CTNAME" & vbCrLf
                sql &= " ,cc.SOCID" & vbCrLf
                sql &= " ,cc.idno" & vbCrLf
                sql &= " ,cc.SumOfMoney" & vbCrLf
                sql &= " ,cc.AppliedStatus" & vbCrLf
                sql &= " FROM WC2 cc" & vbCrLf
                sql &= " ORDER BY cc.IDNO,cc.GCODE2,cc.OCID" & vbCrLf

            Case "Y2" 'V_sameJOB1(相同): N否 Y是'1:統計表 /2:明細表
                sql &= " ,WC3 AS (" & sWC3 & ")" & vbCrLf
                sql &= " SELECT cc.plankind" & vbCrLf
                sql &= " ,cc.OCID" & vbCrLf
                sql &= " ,cc.years" & vbCrLf
                sql &= " ,cc.distname" & vbCrLf
                sql &= " ,cc.ORGPLANNAME" & vbCrLf
                sql &= " ,cc.orgname" & vbCrLf
                sql &= " ,cc.comidno" & vbCrLf
                sql &= " ,cc.OrgTypeName" & vbCrLf 'sql &= " /*單位屬性*/" & vbCrLf
                sql &= " ,cc.KNAME12" & vbCrLf
                sql &= " ,cc.GCODE2" & vbCrLf 'sql &= " /*訓練業別*/" & vbCrLf
                sql &= " ,cc.CNAME" & vbCrLf
                sql &= " ,cc.classcname" & vbCrLf
                sql &= " ,cc.CTNAME" & vbCrLf
                sql &= " ,cc.idno" & vbCrLf
                sql &= " ,cc.SumOfMoney" & vbCrLf
                sql &= " ,cc.AppliedStatus" & vbCrLf
                sql &= " FROM WC3 c3" & vbCrLf
                sql &= " JOIN WC2 cc ON cc.idno=c3.idno and cc.GCODE2=c3.GCODE2" & vbCrLf
                sql &= " ORDER BY cc.IDNO,cc.GCODE2,cc.OCID" & vbCrLf

            Case Else
                Return dt
        End Select

        Dim sCmd As New SqlCommand(sql, objconn)

        'TIMS.OpenDbConn(objconn)
        'Dim dt As New DataTable
        With sCmd
            '.Connection = objconn
            '.CommandTimeout = 100
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function


    ''' <summary> '匯出SUB'SQL </summary>
    Sub XLS_ExpRpt()
        'With sCmd2
        '    .Connection = objconn
        '    .CommandTimeout = 100
        '    .Parameters.Clear()
        '    dt.Load(.ExecuteReader())
        'End With
        'dt.DefaultView.Sort = "年度,分署,品名"
        'dt = TIMS.dv2dt(dt.DefaultView)

        Dim dt As DataTable = Nothing
        dt = Search2dt(sameJOB1.SelectedValue, expType1.SelectedValue)
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Select Case expType1.SelectedValue
            Case "1" '統計表
                'Call SHOW_expType1x1()
                Call SHOW_expType1x1(dt)

            Case "2" '明細表
                Call SHOW_expType1x2(dt)

        End Select
    End Sub

    '統計表
    Sub SHOW_expType1x1(ByRef dt As DataTable)
        Const cst_colspanT2xNum As Integer = 19 '總匯出欄位
        Const cst_StrTitle2 As String = "重複參訓統計表"
        Dim strTitleN2 As String = ""
        strTitleN2 &= CStr(Val(ddlYears.SelectedValue) - 1911) & "年度"
        strTitleN2 &= TIMS.GetTPlanName(Convert.ToString(sm.UserInfo.TPlanID), objconn)
        strTitleN2 &= cst_StrTitle2

        Dim sY1 As String = CStr(Val(ddlYears.SelectedValue) - 1911)

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(cst_StrTitle2, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")

        Dim ExportStr As String = ""

        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=" & cst_colspanT2xNum & ">" & strTitleN2 & "</td>"
        ExportStr &= "</tr>"

        ExportStr &= "<tr>"
        ExportStr &= "<td>年度</td>"
        ExportStr &= "<td>0次(未重複)</td>"
        ExportStr &= "<td>1次</td>"
        ExportStr &= "<td>2次</td>"
        ExportStr &= "<td>3次</td>"
        ExportStr &= "<td>4次</td>"
        ExportStr &= "<td>5次</td>"

        ExportStr &= "<td>0-5次</td>"
        ExportStr &= "<td>6-10次</td>"
        ExportStr &= "<td>11-15次</td>"
        ExportStr &= "<td>16-20次</td>"
        ExportStr &= "<td>21-25次</td>"
        ExportStr &= "<td>26-30次</td>"
        ExportStr &= "<td>31-35次</td>"
        ExportStr &= "<td>36-40次</td>"
        ExportStr &= "<td>41-45次</td>"
        ExportStr &= "<td>46-50次</td>"
        ExportStr &= "<td>51次以上</td>"

        ExportStr &= "<td>參訓總人數</td>"
        ExportStr &= "</tr>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Dim dr1 As DataRow = Nothing
        dr1 = dt.Rows(0)
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>" & sY1 & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT0")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT1")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT2")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT3")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT4")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT5")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT05")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT10")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT15")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT20")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT25")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT30")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT35")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT40")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT45")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT50")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT51")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNTALL")) & "</td>"
        ExportStr &= "</tr>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        dr1 = dt.Rows(1)
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>領取補助費</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT0")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT1")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT2")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT3")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT4")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT5")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT05")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT10")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT15")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT20")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT25")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT30")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT35")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT40")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT45")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT50")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT51")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNTALL")) & "</td>"
        ExportStr &= "</tr>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        dr1 = dt.Rows(2)
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>未領取補助費</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT0")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT1")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT2")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT3")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT4")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT5")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT05")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT10")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT15")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT20")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT25")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT30")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT35")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT40")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT45")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT50")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNT51")) & "</td>"
        ExportStr &= "<td>" & Convert.ToString(dr1("CNTALL")) & "</td>"
        ExportStr &= "</tr>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))  'Call TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    Sub SHOW_expType1x2(ByRef dt As DataTable)
        Const cst_colspanT2xNum As Integer = 13 '總匯出欄位
        Const cst_StrTitle2 As String = "重複參訓統計表-明細"
        Dim strTitleN2 As String = ""
        strTitleN2 &= CStr(Val(ddlYears.SelectedValue) - 1911) & "年度"
        strTitleN2 &= TIMS.GetTPlanName(Convert.ToString(sm.UserInfo.TPlanID), objconn)
        strTitleN2 &= cst_StrTitle2

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(cst_StrTitle2, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")

        Dim ExportStr As String = ""
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td colspan=" & cst_colspanT2xNum & ">" & strTitleN2 & "</td>"
        ExportStr &= "</tr>"

        ExportStr &= "<tr>"
        ExportStr &= "<td>年度</td>"
        ExportStr &= "<td>分署</td>"
        ExportStr &= "<td>計畫別</td>"
        ExportStr &= "<td>身分證字號</td>"
        ExportStr &= "<td>單位名稱</td>"
        ExportStr &= "<td>統編</td>"
        ExportStr &= "<td>單位屬性</td>"
        ExportStr &= "<td>課程分類</td>"
        ExportStr &= "<td>訓練業別</td>"
        ExportStr &= "<td>課程名稱</td>"
        ExportStr &= "<td>課程代碼</td>"
        ExportStr &= "<td>補助費用</td>"
        ExportStr &= "<td>上課地點縣市別</td>"
        ExportStr &= "</tr>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows
            ExportStr = "<tr>"
            ExportStr &= "<td>" & Convert.ToString(dr("Years")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("DistName")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ORGPLANNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("IDNO")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("COMIDNO")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("OrgTypeName")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("kname12")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("GCODE2")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("OCID")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("SumOfMoney")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("CTNAME")) & "</td>"
            ExportStr &= "</tr>"
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        'Call TIMS.CloseDbConn(objconn)
        'Response.End()
    End Sub

    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click

        If ddlYears.SelectedValue = "" Then
            Common.MessageBox(Me, "請選擇有效年度!!")
            Exit Sub
        End If
        'If ddlDistID.SelectedValue = "" Then
        '    Common.MessageBox(Me, "請選擇有效轄區!!")
        '    Exit Sub
        'End If

        Dim okFlag As Boolean = False '結束狀況有誤
        Call TIMS.OpenDbConn(objconn)
        Try
            Call XLS_ExpRpt() '匯出SUB'SQL
            okFlag = True '結束狀況無誤
        Catch ex As Exception
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            Exit Sub
        End Try

        '結束狀況無誤
        If okFlag Then TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Call TIMS.CloseDbConn(objconn)'Response.End()
    End Sub
End Class
