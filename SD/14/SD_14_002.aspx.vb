Partial Class SD_14_002
    Inherits AuthBasePage

    'SD_14_002_Q.aspx  '原：TIMS/2017_T28'SD_14_002_R.aspx  '產投：2018:'Const cst_printASPX_Q As String="SD_14_002_Q.aspx?ID=" 'OLD
    Const cst_printASPX_R As String = "SD_14_002_R.aspx?ID=" 'NEW
    'Type: A:已轉班查詢 B:未轉班查詢
    Dim sPrintASPX1 As String = ""

    Const cst_inline1 As String = ""  '"inline"
    Const cst_訓練單位名稱 As Integer = 0
    'SD_14_002_R.aspx(訓練班別計畫表)
    Dim iPYNum As Integer = 1 'iPYNum=TIMS.sUtl_GetPYNum(Me)  '1:2017前 2:2017 3:2018
    'iPYNum=TIMS.sUtl_GetPYNum(Me)
    'Dim au As New cAUTH

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        'Dim Kind, sql, OrgIdKind As String
        PageControler1.PageDataGrid = DataGrid1
        'PageControler2.PageDataGrid=DataGrid2
        sPrintASPX1 = String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))

        If Not IsPostBack Then
            Create1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        If Me.Radio1.SelectedIndex = 1 Then
            TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
            If HistoryTable.Rows.Count <> 0 Then
                OCID1.Attributes("onclick") = "showObj('HistoryList');"
                OCID1.Style("CURSOR") = "hand"
            End If
        End If

        'Years.Value=sm.UserInfo.Years - 1911
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button4.Attributes("onclick") = "ClearData();"

        Call USE_KEEP_SEARCH()

        Dim FG_SHOW1 As Boolean = (sm.UserInfo.LID < 2)
        BTN_EXPORT1.Visible = FG_SHOW1
        tr_RBListExpType.Visible = FG_SHOW1
    End Sub

    Sub USE_KEEP_SEARCH()
        Dim MyValue As String = ""
        If Session("SCH_SD_14_002") Is Nothing Then Return
        Dim strSearch As String = Convert.ToString(Session("SCH_SD_14_002"))
        Session("SCH_SD_14_002") = Nothing

        MyValue = TIMS.GetMyValue(strSearch, "CALLFUNC")
        If MyValue <> "TC11001" Then Return

        center.Text = TIMS.GetMyValue(strSearch, "center")
        RIDValue.Value = TIMS.GetMyValue(strSearch, "RIDValue")
        '班級狀態 0:未轉班
        MyValue = TIMS.GetMyValue(strSearch, "Radio1")
        Common.SetListItem(Radio1, MyValue)
        ClassTR.Visible = If(MyValue = "0", False, True) '未轉班
        TR_5.Visible = If(MyValue = "0", False, True) '未轉班
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'PlanPoint
        STDate1.Text = ""
        STDate2.Text = ""
        ClassName.Text = ""
        If tr_AppStage_TP28.Visible Then
            MyValue = TIMS.GetMyValue(strSearch, "AppStage")
            Common.SetListItem(AppStage, MyValue)
        End If
        TxtPageSize.Text = "30"

        'TR_rblFONTTYPE 'rblFONTTYPE
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call CreateClassPlan()
    End Sub

    Sub Create1()
        BTN_EXPORT1.Visible = False
        tr_RBListExpType.Visible = False
        'search_act1
        '列印字型選擇
        rblFONTTYPE.Attributes("onclick") = "rblFONTTYPE_CHG1();"
        msg.Text = ""
        '計畫 1:產業人才投資計畫 2:提升勞工自主學習計畫 0:不區分
        PlanPoint = TIMS.Get_RblPlanPoint0(Me, PlanPoint, objconn)
        Common.SetListItem(PlanPoint, "0")

        'TR_1.Style.Add("display", "none")
        TR_2.Style.Add("display", "none")
        TR_3.Style.Add("display", "none")
        If sm.UserInfo.OrgLevel <= 1 Then '署(局):0／分署(中心):1
            'TR_1.Style.Add("display", cst_inline1)
            TR_2.Style.Add("display", cst_inline1)
            TR_3.Style.Add("display", cst_inline1)
        End If
        DataGridTable.Visible = False
        ClassTR.Visible = False '未轉班
        TR_5.Visible = False '未轉班
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage=TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
            Call TIMS.SET_MY_APPSTAGE_LIST_VAL(Me, AppStage)
        End If

        '班級狀態 0:未轉班 1:已轉班
        Me.Radio1.SelectedIndex = 0

    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)

        Dim mySTDate1 As String = If(flag_ROC, TIMS.Cdate18(STDate1.Text), STDate1.Text)  'edit，by:20181023
        Dim mySTDate2 As String = If(flag_ROC, TIMS.Cdate18(STDate2.Text), STDate2.Text)  'edit，by:20181023

        If RIDValue.Value = "" Then Errmsg += "機構選擇有誤，請重新選擇!" & vbCrLf

        If mySTDate1 <> "" AndAlso Not TIMS.IsDate1(mySTDate1) Then Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf  'edit，by:20181023
        If mySTDate2 <> "" AndAlso Not TIMS.IsDate1(mySTDate2) Then Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf  'edit，by:20181023
        If Errmsg <> "" Then Return False

        If mySTDate1 <> "" AndAlso mySTDate2 <> "" Then
            If CDate(mySTDate1) > CDate(mySTDate2) Then Errmsg += "【開訓區間】的起日不得大於【開訓區間】的迄日!!" & vbCrLf  'edit，by:20181023
        End If
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Function SSearch1_DATA_dt() As DataTable
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Dim v_rdlResult As String = TIMS.GetListValue(rdlResult)
        Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
        'Dim v_rdo_printOrg As String=TIMS.GetListValue(rdo_printOrg) '顯示訓練單位名稱 'Hid_rdo_printOrg.Value=TIMS.ClearSQM(v_rdo_printOrg)

        '顯示訓練單位名稱 'Dim flag_show_orgname As Boolean = If(v_rdo_printOrg = "N", False, True) ' True
        '機構階層 0.署(局) 1.分署(中心) 2.委訓(補助地方政府) 3.補助地方委訓
        'If flag_show_orgname AndAlso sm.UserInfo.OrgLevel > 1 Then flag_show_orgname = False
        'DataGrid1.Columns(cst_訓練單位名稱).Visible = flag_show_orgname 'False
        Dim PMS1 As New Hashtable

        Dim sql As String = ""
        sql &= " SELECT A.PLANID,A.COMIDNO,A.SEQNO,A.PSNO28,CC.OCID" & vbCrLf
        Select Case v_Radio1'.SelectedValue '0:未轉班,1:已轉班
            Case "0"
                sql &= " ,A.RESULTBUTTON,CONCAT(dbo.FN_GET_CLASSCNAME(A.CLASSNAME,A.CYCLTYPE),CASE WHEN A.RESULTBUTTON IN ('Y','R') THEN '(未送出)' END) CLASSCNAME" & vbCrLf
            Case "1"
                sql &= " ,A.RESULTBUTTON,dbo.FN_GET_CLASSCNAME(A.CLASSNAME,A.CYCLTYPE) CLASSCNAME" & vbCrLf
        End Select
        sql &= " ,CONVERT(VARCHAR, A.STDATE, 111) STDATE ,CONVERT(VARCHAR, A.FDDATE, 111) FTDATE" & vbCrLf
        sql &= " ,B.ORGNAME ,A.RID" & vbCrLf
        sql &= " ,dbo.FN_GET_TRAINDESC(A.PLANID,A.COMIDNO,A.SEQNO,'PCONT') PCONT" '--課程進度/內容" & vbCrLf
        '訓練需求調查(產業人力需求調查,區域人力需求調查,訓練需求概述
        sql &= " ,P2.POWERNEED1,P2.POWERNEED2,P2.POWERNEED3"
        '與政策性產業課程之關聯性概述),訓練目標(單位核心能力介紹
        sql &= " ,P2.POLICYREL,P2.PLANCAUSE"
        '知識,技能,學習成效,其他設施說明
        sql &= " ,P2.PURSCIENCE,P2.PURTECH,P2.PURMORAL,P2.OTHFACDESC23"
        Select Case v_Radio1'.SelectedValue '0:未轉班,1:已轉班
            Case "0"
                sql &= " ,FORMAT(A.MODIFYDATE,'mmssdd') MSD" & vbCrLf 'pp.MSD
            Case "1"
                sql &= " ,FORMAT(CC.MODIFYDATE,'mmssdd') MSD" & vbCrLf 'cc.MSD
        End Select
        sql &= " FROM dbo.PLAN_PLANINFO A" & vbCrLf
        Select Case v_Radio1'.SelectedValue '0:未轉班,1:已轉班
            Case "0" '0:未轉班,1:已轉班
                sql &= " LEFT JOIN dbo.CLASS_CLASSINFO CC ON CC.PLANID=A.PLANID AND CC.COMIDNO=A.COMIDNO AND CC.SEQNO=A.SEQNO" & vbCrLf
            Case "1" '0:未轉班,1:已轉班
                sql &= " JOIN dbo.CLASS_CLASSINFO CC ON CC.PLANID=A.PLANID AND CC.COMIDNO=A.COMIDNO AND CC.SEQNO=A.SEQNO" & vbCrLf
        End Select
        sql &= " JOIN dbo.VIEW_RIDNAME B ON A.RID=B.RID" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN IP ON IP.PLANID=A.PLANID" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT P2 ON P2.PLANID=A.PLANID AND P2.COMIDNO=A.COMIDNO AND P2.SEQNO=A.SEQNO" & vbCrLf
        sql &= $" WHERE ip.TPlanID='{sm.UserInfo.TPlanID}' AND ip.Years='{sm.UserInfo.Years}'" & vbCrLf
        Select Case v_Radio1'.SelectedValue
            Case "0" '0:未轉班,1:已轉班
                sql &= " AND a.TransFlag='N' AND a.IsApprPaper='Y'" & vbCrLf
            Case "1" '0:未轉班,1:已轉班
                sql &= " AND a.TransFlag='Y' AND a.IsApprPaper='Y'" & vbCrLf
                If OCIDValue1.Value <> "" Then sql &= $" AND cc.OCID={TIMS.CINT1(OCIDValue1.Value)}"
        End Select
        If sm.UserInfo.LID <> 0 Then
            sql &= $" AND ip.PlanID={sm.UserInfo.PlanID}" & vbCrLf
        End If
        '沒選擇單位 使用登入者業務權限
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        '機構階層 0.分署(局) 1.分署(中心) 2.委訓(補助地方政府) 3.補助地方委訓
        Select Case $"{sm.UserInfo.OrgLevel}"
            Case "0", "1"
                '有選擇單位
                If RIDValue.Value.Length = 1 Then
                    If RIDValue.Value = "A" Then
                        '長度為1 使用like
                        sql &= $" AND B.RELSHIP LIKE '{RIDValue.Value}%'" & vbCrLf
                    Else
                        '長度為1 使用like
                        sql &= $" AND a.RID LIKE '{RIDValue.Value}%'" & vbCrLf
                    End If
                Else
                    '不為1
                    sql &= $" AND a.RID='{RIDValue.Value}'" & vbCrLf
                End If
            Case Else '"2","3"
                sql &= $" AND a.RID='{RIDValue.Value}'" & vbCrLf
        End Select

        '20100409 andy  edit 加入課程審核狀況
        Select Case v_rdlResult'.SelectedValue
            Case "Y", "N"
                sql &= $" AND a.AppliedResult='{v_rdlResult}'"
        End Select

        '28:產業人才投資方案
        'TRPlanPoint28
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case v_PlanPoint'.SelectedValue
                Case "1"
                    sql &= " AND b.OrgKind<>10" & vbCrLf
                Case "2"
                    sql &= " AND b.OrgKind=10" & vbCrLf
            End Select
        End If

        'START 加入開訓期間條件 2009/07/15 by waiming 
        If STDate1.Text <> "" Then sql &= $" AND a.STDate>={TIMS.To_date(If(flag_ROC, TIMS.Cdate18(STDate1.Text), STDate1.Text))}" 'edit，by:20181023
        If STDate2.Text <> "" Then sql &= $" AND a.STDate<={TIMS.To_date(If(flag_ROC, TIMS.Cdate18(STDate2.Text), STDate2.Text))}" 'edit，by:20181023
        'END 加入開訓期間條件

        '20100409 andy 加入 班級名稱 查詢條件-2 
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        If ClassName.Text <> "" Then sql &= $" AND dbo.FN_GET_CLASSCNAME(A.CLASSNAME,A.CYCLTYPE) LIKE '%{ClassName.Text}%'" & vbCrLf

        '班級名稱,CLASSNAME、課程進度/內容,PCONT、
        '訓練需求調查(產業人力需求調查POWERNEED1、區域人力需求調查POWERNEED2、訓練需求概述POWERNEED3、與政策性產業課程之關聯性概述POLICYREL)、
        '訓練目標(單位核心能力介紹PlanCause、知識PurScience、技能PurTech、學習成效PurMoral)、其他設施說明OTHFACDESC23。
        If ClassKW1.Text <> "" Then
            PMS1.Add("ClassKW1", ClassKW1.Text)
            sql &= " AND ( EXISTS (SELECT 1 FROM PLAN_TRAINDESC PT WITH(NOLOCK) WHERE PT.PLANID=A.PLANID AND PT.COMIDNO=A.COMIDNO AND PT.SEQNO=A.SEQNO AND UPPER(PT.PCONT) LIKE '%'+UPPER(@ClassKW1)+'%')
 OR UPPER(CC.CLASSCNAME) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(A.CLASSNAME) LIKE '%'+UPPER(@ClassKW1)+'%' 
 OR UPPER(P2.POWERNEED1) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.POWERNEED2) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.POWERNEED3) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.POLICYREL) LIKE '%'+UPPER(@ClassKW1)+'%' 
 OR UPPER(P2.PLANCAUSE) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.PURSCIENCE) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.PURTECH) LIKE '%'+UPPER(@ClassKW1)+'%' OR UPPER(P2.PURMORAL) LIKE '%'+UPPER(@ClassKW1)+'%'
 OR UPPER(P2.OTHFACDESC23) LIKE '%'+UPPER(@ClassKW1)+'%' )" & vbCrLf
        End If
        If ClassKW2.Text <> "" Then
            PMS1.Add("ClassKW2", TIMS.CINT1(ClassKW2.Text))
            sql &= " AND cc.OCID=@ClassKW2" & vbCrLf
        End If

        '依申請階段
        Dim v_AppStage As String = If(tr_AppStage_TP28.Visible, TIMS.GetListValue(AppStage), "")
        If (v_AppStage <> "") Then Session(TIMS.SESS_DDL_APPSTAGE_VAL) = v_AppStage
        If v_AppStage <> "" Then sql &= $" AND a.AppStage={TIMS.CINT1(v_AppStage)}" & vbCrLf '依申請階段
        sql &= " ORDER BY a.PlanID ,a.ComIDNO ,a.SeqNo ,cc.OCID" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            '取得目前方法所屬的類別 Type'取得類別名稱（包含命名空間）'取得類別名稱（不包含命名空間）
            'Dim currentType As Type = System.Reflection.MethodBase.GetCurrentMethod().DeclaringType
            'Dim fullName As String = currentType.FullName
            'Dim ClassName As String = currentType.Name
            'TIMS.WriteLog(Me, $"{vbCrLf}--fullName:{fullName}{vbCrLf}--ClassName:{ClassName}")
            TIMS.WriteLog(Me, $"{vbCrLf}--sPMS:{TIMS.GetMyValue5(PMS1)}{vbCrLf}--#{System.Reflection.MethodBase.GetCurrentMethod().DeclaringType.Name}:{vbCrLf}{sql}")
        End If

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, PMS1)
        Return DbAccess.GetDataTable(sql, objconn, PMS1)
    End Function

    '查詢[SQL]
    Sub CreateClassPlan()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        'Dim v_rdlResult As String = TIMS.GetListValue(rdlResult)
        'Dim v_PlanPoint As String = TIMS.GetListValue(PlanPoint)
        'Dim v_rdo_printOrg As String=TIMS.GetListValue(rdo_printOrg) '顯示訓練單位名稱 'Hid_rdo_printOrg.Value=TIMS.ClearSQM(v_rdo_printOrg)
        Select Case v_Radio1'.SelectedValue '0:未轉班,1:已轉班
            Case "0", "1"
                '0:未轉班,1:已轉班
                Hid_Radio1.Value = v_Radio1 'Radio1.SelectedValue
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Exit Sub
        End Select
        DataGrid1.Columns(cst_訓練單位名稱).Visible = True 'OCIDValue.Value=""

        Dim dt As DataTable = SSearch1_DATA_dt()

        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If TIMS.dtNODATA(dt) Then Return

        DataGrid1.Visible = True
        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call CreateClassPlan()

        Dim FG_SHOW1 As Boolean = (sm.UserInfo.LID < 2)
        BTN_EXPORT1.Visible = FG_SHOW1
        tr_RBListExpType.Visible = FG_SHOW1
    End Sub

    '已轉班(依班級查詢) 含列印
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputHidden = e.Item.FindControl("OCID")
                Dim PlanID As HtmlInputHidden = e.Item.FindControl("PlanID")
                Dim ComIDNO As HtmlInputHidden = e.Item.FindControl("ComIDNO")
                Dim SeqNo As HtmlInputHidden = e.Item.FindControl("SeqNo")
                Dim v_rblFONTTYPE As String = TIMS.GetListValue(rblFONTTYPE) '1:細明體/2:標楷體(def)

                e.Item.Cells(2).Text = If(flag_ROC, TIMS.Cdate17(drv("STDate")), drv("STDate"))  'edit，by:20181023
                e.Item.Cells(3).Text = If(flag_ROC, TIMS.Cdate17(drv("FTDate")), drv("FTDate"))  'edit，by:20181023

                '**by Milor 20080429--列印方式改為逐筆列印按鈕，避免多張列印(因為報表設計方式只能查詢單筆)----start
                Dim PrintRpt1 As HtmlInputButton = e.Item.FindControl("PrintRpt1")
                '**by Milor 20080429--列印方式改為逐筆列印按鈕，避免多張列印(因為報表設計方式只能查詢單筆)----start
                'Dim PrintRpt2 As HtmlInputButton=e.Item.FindControl("PrintRpt2")

                '0:未轉班,1:已轉班
                PrintRpt1.Visible = False
                'PrintRpt2.Visible=False
                Select Case Hid_Radio1.Value'.SelectedValue
                    Case "0" '0:未轉班,1:已轉班
                        '未轉班(依計畫查詢) 含列印
                        PrintRpt1.Visible = True
                        PlanID.Value = Convert.ToString(drv("PlanID"))
                        ComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                        SeqNo.Value = Convert.ToString(drv("SeqNo"))
                        'Dim v_MSD As String=Convert.ToString(drv("MSD"))
                        Dim CYearsVal As String = (sm.UserInfo.Years - 1911)
                        'Dim sUrl As String=String.Concat(If(iPYNum >= 3, cst_printASPX_R, cst_printASPX_Q), Request("ID"))

                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "Type", "B") 'Type: A:已轉班查詢 B:未轉班查詢
                        'TIMS.SetMyValue(sCmdArg, "PrintOrg", Hid_rdo_printOrg.Value) '顯示訓練單位名稱
                        TIMS.SetMyValue(sCmdArg, "PrintOrg", "Y") '顯示訓練單位名稱
                        TIMS.SetMyValue(sCmdArg, "Years", CYearsVal)
                        TIMS.SetMyValue(sCmdArg, "PlanID", PlanID.Value)
                        TIMS.SetMyValue(sCmdArg, "ComIDNO", ComIDNO.Value)
                        TIMS.SetMyValue(sCmdArg, "SeqNo", SeqNo.Value)
                        TIMS.SetMyValue(sCmdArg, "FTYPE", v_rblFONTTYPE) '1:細明體/2:標楷體(def)
                        'TIMS.SetMyValue(sCmdArg, "MSD", v_MSD)

                        Dim ock_Value1 As String = String.Concat("window.open('", sPrintASPX1, sCmdArg, "','','resizable=yes,toolbar=no,scrollbars=yes');")
                        '20090107 andy edit 報表改為網頁產生的方式
                        PrintRpt1.Attributes.Add("onclick", ock_Value1)

                    Case "1" '0:未轉班,1:已轉班
                        '已轉班(依班級查詢) 含列印
                        PrintRpt1.Visible = True
                        OCID.Value = Convert.ToString(drv("OCID"))
                        'Dim v_MSD As String=Convert.ToString(drv("MSD"))
                        Dim CYearsVal As String = (sm.UserInfo.Years - 1911)
                        'Dim sUrl As String=String.Concat(If(iPYNum >= 3, cst_printASPX_R, cst_printASPX_Q), Request("ID"))

                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "Type", "A") 'Type: A:已轉班查詢 B:未轉班查詢
                        'TIMS.SetMyValue(sCmdArg, "PrintOrg", rdo_printOrg.SelectedValue) '顯示訓練單位名稱
                        TIMS.SetMyValue(sCmdArg, "PrintOrg", "Y") '顯示訓練單位名稱
                        TIMS.SetMyValue(sCmdArg, "Years", CYearsVal)
                        TIMS.SetMyValue(sCmdArg, "OCID", OCID.Value)
                        TIMS.SetMyValue(sCmdArg, "FTYPE", v_rblFONTTYPE) '1:細明體/2:標楷體(def)
                        'TIMS.SetMyValue(sCmdArg, "MSD", v_MSD)

                        Dim ock_Value1 As String = String.Concat("window.open('", sPrintASPX1, sCmdArg, "','','resizable=yes,toolbar=no,scrollbars=yes');")
                        '20090107 andy edit 報表改為網頁產生的方式
                        PrintRpt1.Attributes.Add("onclick", ock_Value1)

                End Select
        End Select
    End Sub

    '班級狀態
    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        '班級狀態: False:'未轉班／True:已轉班
        ClassTR.Visible = If(Radio1.SelectedIndex = 0, False, True) '未轉班
        TR_5.Visible = If(Radio1.SelectedIndex = 0, False, True) '未轉班
        Call ClsClassTR()
        DataGridTable.Visible = False
    End Sub

    ''' <summary>切換時清理</summary>
    Sub ClsClassTR()
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        ClassName.Text = ""
        ClassKW1.Text = ""
        If Not TR_5.Visible Then ClassKW2.Text = ""
        'ClassKW2.Text = ""
    End Sub

    Sub EXPORT_1()
        Dim sERRMSG As String = ""
        Call CheckData1(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Page, sERRMSG)
            Exit Sub
        End If

        Dim dtXls As DataTable = SSearch1_DATA_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Const cst_TitleS1 As String = "匯出訓練班別"
        Dim strFilename1 As String = String.Concat(cst_TitleS1, TIMS.GetDateNo2())
        'Dim sTitle1 As String = "匯出訓練班別"

        Dim sPattern As String = "課程申請流水號,課程代碼,訓練單位,班級名稱,課程進度/內容,訓練需求調查(產業人力需求調查,區域人力需求調查,訓練需求概述,與政策性產業課程之關聯性概述),訓練目標(單位核心能力介紹,知識,技能,學習成效),其他設施說明"
        Dim sColumn As String = "PSNO28,OCID,ORGNAME,CLASSCNAME,PCONT,POWERNEED1,POWERNEED2,POWERNEED3,POLICYREL,PLANCAUSE,PURSCIENCE,PURTECH,PURMORAL,OTHFACDESC23"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.cst_NO_TITLENAME)
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Call EXPORT_1()
    End Sub

    '顯示訓練單位名稱
    'Private Sub rdo_printOrg_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo_printOrg.SelectedIndexChanged
    '    DataGridTable.Visible=False
    '    'Button1_Click(sender, e)
    'End Sub

End Class
