'Imports Aspose.Words
Imports System.IO
Partial Class SD_14_002_R
    Inherits AuthBasePage

    'SD_14_002_Q.aspx  '原：TIMS/2017_T28'SD_14_002_R.aspx  '產投：2018:後 

    'PLAN_TRAINPLACE
    '確認所選之訓練業別正確性
    'Const cst_TMIDCORRECT_t As String = "確認所選之訓練業別正確性"
    Const cst_TMIDCORRECT_c As String = "貴單位於研提課程時請務必確認所選之訓練業別正確性，如經審查小組審查所選之訓練業別有誤，是否同意協助重新歸類，如不同意，則將依貴單位所選之業別逕行審查。"
    Const cst_msg_memo8a As String = "本課程非屬於「職業安全衛生教育訓練規則」所訂定之訓練課程，無法作為時數認列。"

    'Const cst_images_rptpic_yes_jpg As String = "<img src='../../images/rptpic/yes.jpg' />"
    'Const cst_images_rptpic_no_jpg As String = "<img src='../../images/rptpic/no.jpg' />"
    'Const cst_images_rptpic_yes_jpg As String = "<strong>∨✔☑ √Ⅴ</strong>"
    Const cst_images_rptpic_yes_jpg As String = "<strong>∨</strong>"
    Const cst_images_rptpic_no_jpg As String = "<strong>□</strong>"
    Const cst_images_rptpic_yy_jpg As String = "<strong>■</strong>"

    Dim str_fontfamily_c As String = ""
    'Const cst_fontfamily_c1 As String = "font-family:DFKai-SB"
    'Const cst_fontfamily_c1 As String = "font-family:KaiTi_GB2312,DFKai-SB,KaiU"
    'Const cst_fontfamily_c1 As String = "font-family:DFKai-SB" '1:標楷體(def)/2:細明體
    Const cst_fontfamily_c1 As String = "" '不使用 標楷體 
    Const cst_fontfamily_c2 As String = "font-family:DFKai-SB"

    'Const cst_bgimg_c1 As String = "../../images/rptpic/temple/TIMS_1.jpg"
    'Const cst_style_c1 As String = "BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none"
    Const cst_style_c1 As String = "BORDER-TOP-STYLE: none;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none"
    'Const cst_style_c2 As String = "FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE:2px solid;BORDER-LEFT-STYLE:2px solid;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE:1px solid"
    Const cst_style_c2 As String = "BORDER-RIGHT-STYLE:2px solid;BORDER-LEFT-STYLE:2px solid;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE:1px solid"
    'Const cst_style_c3 As String = "FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: solid;BORDER-LEFT-STYLE: solid;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: solid;border-top-style:solid;"
    Const cst_style_c3 As String = "BORDER-RIGHT-STYLE: solid;BORDER-LEFT-STYLE: solid;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: solid;border-top-style:solid;"
    'Const cst_print_style_x1 As String = "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none"

    Const cst_TeacherDesc_default1 As String = "(依計畫師資及助教資格標準表)"

    Const cst_errorMsg1 As String = "此班資料異常，請連絡系統管理者!!"
    Const cst_ENTERSUPPLY_1 As String = "報名時應先繳全額訓練費用"
    Const cst_ENTERSUPPLY_2 As String = "報名時應先繳50%訓練費用"
    'Const cst_sType_A_已轉班 As String = "A" '已轉班
    'Const cst_sType_B_未轉班 As String = "B" '未轉班
    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me)  '1:2017前 2:2017 3:2018

    Dim flag_OJT22071401 As Boolean = False

    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_t1 As String = "技檢訓練時數,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，該資料不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const TIMS.cst_EHour_Use_TMID As String = "672"

    Public Property o_TMID As Object
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        'OJT-22071401系統-產投-班級申請：新增「訓練業別同意重新歸類」選項 +「與政策性產業課程之關聯性概述」欄位 
        If Hid_OJT22071401.Value = "" Then Hid_OJT22071401.Value = TIMS.Utl_GetConfigVAL(objconn, "OJT22071401")
        flag_OJT22071401 = If(Hid_OJT22071401.Value = "Y", True, False)
        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        imgBt_Pdf.Visible = flag_test

        'actionText.Visible = False
        'img_waiting.Visible = False
        'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
        'iPYNum = TIMS.sUtl_GetPYNum(Me)
        Dim rqType As String = TIMS.ClearSQM(Request("Type"))

        '1:細明體/2:標楷體(def)
        Dim rqFTYPE As String = TIMS.ClearSQM(Request("FTYPE"))
        'Dim v_SD14002RFT As String = TIMS.Utl_GetConfigSet("SD14002RFT") '1:細明體/2:標楷體(def)
        str_fontfamily_c = If(rqFTYPE = "1", cst_fontfamily_c1, cst_fontfamily_c2)

        'Type: A:已轉班查詢 B:未轉班查詢
        Select Case rqType
            Case TIMS.cst_sType_A_已轉班, TIMS.cst_sType_B_未轉班
                Call LoadData(rqType)
        End Select

        Dim PlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim SeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqPCS As String = $"{PlanID}x{ComIDNO}x{SeqNo}"
        Dim tkDecVAL As String = If(Request("tk") IsNot Nothing, TIMS.DecryptAes(Request("tk")), "")

        Dim rqRID As String = TIMS.GetMyValue(tkDecVAL, "RID")
        Dim rqBCID As String = TIMS.GetMyValue(tkDecVAL, "BCID")
        Dim rqBCASENO As String = TIMS.GetMyValue(tkDecVAL, "BCASENO")
        Dim rqORGKINDGW As String = TIMS.GetMyValue(tkDecVAL, "ORGKINDGW")
        Dim rqKBSID As String = TIMS.GetMyValue(tkDecVAL, "KBSID")

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, rqRID, rqBCID, rqBCASENO)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, rqKBSID, rqORGKINDGW)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim rqPDFOUT As String = TIMS.ClearSQM(Request("PDFOUT"))
        'Dim vsPDFOUT As String = If(ViewState("PDFOUT") IsNot Nothing, $"{ViewState("PDFOUT")}", "")
        If rqBCID <> "" AndAlso rqBCASENO <> "" Then
            If rqPDFOUT = "Y" Then
                'ViewState("PDFOUT") = "Y"
                Dim rPMS As New Hashtable From {
                    {"RID", rqRID},
                    {"BCID", rqBCID},
                    {"BCASENO", rqBCASENO},
                    {"ORGKINDGW", rqORGKINDGW},
                    {"KBSID", rqKBSID},
                    {"PCS", rqPCS}
                }
                Call Export_PDF2(rPMS)
            ElseIf rqPDFOUT = "YB" Then
                'actionText.Visible = True
                'img_waiting.Visible = True
                Dim rPMS As New Hashtable From {
                    {"RID", rqRID},
                    {"BCID", rqBCID},
                    {"BCASENO", rqBCASENO},
                    {"ORGKINDGW", rqORGKINDGW},
                    {"KBSID", rqKBSID},
                    {"PCS", rqPCS}
                }
                Call Export_PDF3(rPMS)
            End If
        End If

    End Sub

    ''' <summary>取得列印資料 (含輸出列印)</summary>
    ''' <param name="sType"></param>
    Sub LoadData(ByVal sType As String)
        'Type: A:已轉班查詢 B:未轉班查詢

        'dt1 PLAN_PLANINFO
        Dim flag_RESULTBUTTON_YR As Boolean = False 'TRUE:未送出  'RESULTBUTTON: Y/R cst_ResultButton_尚未送出_未送出
        Dim PlanKind As String = "" '組合計畫名稱
        Dim OrgKind2 As String = "" 'G/W/NULL [提升勞工自主學習計畫@W]
        Dim OrgName As String = ""
        Dim DISTNAME2 As String = ""
        'Dim YearPlan As String = ""
        Dim DISTANCE_NP As String = ""
        Dim jobName As String = " "
        Dim CCName As String = ""
        Dim ClassName As String = ""
        Dim GCName As String = ""
        Dim ClassID As String = ""
        Dim TNum As String = ""
        Dim STDate As String = ""
        Dim FDDate As String = ""
        Dim Weeks As String = ""
        Dim Thours As String = ""
        Dim Week As String = ""
        Dim TrainDemain As String = ""
        Dim Pur As String = ""
        Dim TMIDCORRECT As String = ""
        Dim TGOVEXAM As String = ""
        Dim GOVAGENAME As String = ""
        Dim TGOVEXAMNAME As String = ""
        Dim ENTERSUPPLYSTYLE As String = ""
        Dim TMethod As String = ""
        Dim tPOWERNEED As String = "" '$"{dr1("POWERNEED1")}"
        'Dim vTPERIOD28 As String = ""
        Dim strMyAppStage As String = ""  '申請階段，by:20181023

        'dt3 'PLAN_TRAINPLACE
        Dim SName, Note, CapAll, Address1, Connum1, HWDesc1, Address2, Connum2, HWDesc2, Address3, Connum3, HWDesc3, Address4, Connum4, HWDesc4, OtherDesc3, TMScience, TMTech As String
        'Dim TeacherDesc As String = ""
        'Dim TeacherDesc2 As String = ""
        'TeacherDesc = ""
        'TeacherDesc2 = ""
        SName = ""
        Note = "" '訓練費用編列說明
        CapAll = ""
        Address1 = "" : Connum1 = "" : HWDesc1 = ""
        Address2 = "" : Connum2 = "" : HWDesc2 = ""
        Address3 = "" : Connum3 = "" : HWDesc3 = ""
        Address4 = "" : Connum4 = "" : HWDesc4 = ""
        OtherDesc3 = "" : TMScience = "" : TMTech = ""

        'dt4 PLAN_TRAINDESC

        'dt5 PLAN_VERREPORT
        Dim Rec, RecDesc, Learn, LearnDesc, Act, ActDesc, Rst, ResultDesc, oth, OtherDesc, DefGovCost, Total1, DefStdCost, Total2, Total4, Total3, FirRst, SecRst As String
        Dim MEMO8 As String = $"{cst_images_rptpic_no_jpg}{cst_msg_memo8a}"
        Dim MEMO82 As String = ""
        Dim ISiCAPCOUR As String = "" '是否為iCAP課程 Y/N/NULL
        Dim iCAPCOURDESC As String = "" '課程相關說明
        Dim iCAPNUM As String = "" 'iCAP標章證號
        Dim iCAPMARKDATE As String = "" 'iCAP標章有效期限
        Dim Recruit As String = "" '招訓方式
        Dim Selmethod As String = "" '遴選方式
        Dim Inspire As String = "" '學員激勵辦法
        Dim s_OTHFACDESC23 As String = "" '「裝備與設施」區塊，增加【其他設施說明】欄位
        Dim s_RMTNAME1 As String = "" '遠距課程環境1
        Dim s_RMTNAME2 As String = "" '遠距課程環境2

        Rec = ""
        RecDesc = ""
        Learn = ""
        LearnDesc = ""
        Act = ""
        ActDesc = ""
        Rst = ""
        ResultDesc = ""
        oth = ""
        OtherDesc = ""
        'ISiCAPCOUR = "" '是否為iCAP課程 Y/N/NULL
        'iCAPCOURDESC = "" '課程相關說明
        'iCAPNUM = "" 'iCAP標章證號
        'Recruit = "" '招訓方式
        'Selmethod = "" '遴選方式
        'Inspire = "" '學員激勵辦法
        DefGovCost = ""
        Total1 = ""
        DefStdCost = ""
        Total2 = ""
        Total4 = ""
        Total3 = ""
        FirRst = ""
        SecRst = ""

        Dim Years As String = TIMS.ClearSQM(Request("Years")) 'ROC.YEARS
        Dim fg_ROCYEARS As Boolean = TIMS.CHK_ROCYEARS(objconn, Years)
        If Not fg_ROCYEARS Then Return
        Dim PlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim SeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim OCID As String = TIMS.ClearSQM(Request("OCID"))
        'Dim ReqMSD As String = TIMS.ClearSQM(Request("MSD"))
        Dim PrintOrg As String = TIMS.ClearSQM(Request("PrintOrg")) '顯示訓練單位名稱 Y/*
        Dim errorFlag As Boolean = False
        If OCID = "" AndAlso (PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "") Then errorFlag = True
        If errorFlag Then Return

        Dim dt1 As New DataTable 'CLASS_CLASSINFO / PLAN_PLANINFO / PLAN_VERREPORT
        Dim dt2a As New DataTable '師資  FN_GET_PLAN_TEACHER3 / PLAN_TRAINDESC
        Dim dt2b As New DataTable '助教  FN_GET_PLAN_TEACHER3 / PLAN_TRAINDESC
        Dim dt3 As New DataTable 'PLAN_TRAINPLACE
        Dim dt4 As New DataTable 'PLAN_TRAINDESC
        Dim dt5 As New DataTable 'PLAN_VERREPORT
        Dim dt6 As New DataTable 'PLAN_ABILITY

        Dim iALL_EHOURS As Double = 0 'GET_TRAINDESCD_HOURS(dt4, "EHOURS") '技檢訓練時數
        Dim iALL_AIAHOUR As Double = 0 'GET_TRAINDESCD_HOURS(dt4, "AIAHOUR") 'AI應用時數
        Dim iALL_WNLHOUR As Double = 0 'GET_TRAINDESCD_HOURS(dt4, "WNLHOUR") '職場續航時數

        Dim dtBPA As New DataTable '= Nothing'(企業包班)
        'Dim da As New SqlDataAdapter
        'Dim dr As DataRow = Nothing 'CLASS_CLASSINFO
        Dim rptTb As HtmlTable = Nothing
        Dim rptRow As HtmlTableRow = Nothing
        Dim rptCell As HtmlTableCell = Nothing

        'Dim sql As String
        Call TIMS.OpenDbConn(objconn)
        'Type: A:已轉班查詢 B:未轉班查詢
        Select Case sType
            Case TIMS.cst_sType_A_已轉班
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return ' Exit Sub
                End If
                'If ReqMSD = "" OrElse (ReqMSD <> CStr(drCC("MSD"))) Then
                '    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)'MSD有誤
                '    Return ' Exit Sub
                'End If
                o_TMID = drCC("TMID")
                PlanID = $"{drCC("PlanID")}"
                ComIDNO = $"{drCC("ComIDNO")}"
                SeqNo = $"{drCC("SeqNO")}"
                dtBPA = TIMS.Get_BUSPACKAGEdt(objconn, OCID, PlanID, ComIDNO, SeqNo) '企業包班

            Case TIMS.cst_sType_B_未轉班
                Dim drPP As DataRow = TIMS.GetPCSDate(PlanID, ComIDNO, SeqNo, objconn)
                If drPP Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return ' Exit Sub
                End If
                o_TMID = drPP("TMID")
                'If ReqMSD = "" OrElse (ReqMSD <> CStr(drPP("MSD"))) Then
                '    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)'MSD有誤
                '    Return ' Exit Sub
                'End If
                dtBPA = TIMS.Get_BUSPACKAGEdt(objconn, OCID, PlanID, ComIDNO, SeqNo) '企業包班

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return ' Exit Sub
        End Select

        '進階政策性產業類別 / 政策性產業類別
        Dim s_KNAME1920 As String = ""
        Dim dr2PD As DataRow = TIMS.GET_PLANDEPOT(PlanID, ComIDNO, SeqNo, objconn)
        If dr2PD IsNot Nothing Then
            Dim flag_SHOW_2019_1 As Boolean = TIMS.SHOW_2019_1(sm)
            s_KNAME1920 = If(flag_SHOW_2019_1, TIMS.Get_DEPOTNAME("20", $"{dr2PD("KID20")}", objconn), TIMS.Get_DEPOTNAME("19", $"{dr2PD("KID19")}", objconn))
        End If
        Dim s_KNAME22 As String = If(dr2PD IsNot Nothing, TIMS.Get_DEPOTNAME("22", $"{dr2PD("KID22")}", objconn), "")
        If s_KNAME22 <> "" Then s_KNAME1920 &= $"{If(s_KNAME1920 <> "", "、", "")}{s_KNAME22}"

        Dim s_KNAME25 As String = If(dr2PD IsNot Nothing, TIMS.Get_DEPOTNAME("25", $"{dr2PD("KID25")}", objconn), "")
        If s_KNAME25 <> "" Then s_KNAME1920 &= $"{If(s_KNAME1920 <> "", "、", "")}{s_KNAME25}"

        Dim s_KNAME26 As String = If(dr2PD IsNot Nothing, TIMS.Get_DEPOTNAME("26", $"{dr2PD("KID26")}", objconn), "")
        If s_KNAME26 <> "" Then s_KNAME1920 &= $"{If(s_KNAME1920 <> "", "、", "")}{s_KNAME26}"

        '不管什麼都是「年滿15歲以上」。
        Const cst_AgeOtherDef As Integer = 16 'other Years Start
        Dim yearsOld_N As String = "年滿15歲以上"

        'Type: A:已轉班查詢 B:未轉班查詢
        Select Case sType
            Case TIMS.cst_sType_A_已轉班
                dt1 = TIMS.Get_ClassInfoDt(objconn, OCID)
                If TIMS.dtNODATA(dt1) Then
                    Common.MessageBox(Me, cst_errorMsg1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, cst_errorMsg1)
                    Exit Sub
                End If

                Dim dr1 As DataRow = dt1.Rows(0)
                If $"{dr1("CapAge1")}" <> "" AndAlso TIMS.CINT1(dr1("CapAge1")) >= cst_AgeOtherDef Then
                    yearsOld_N = $"應符合相關法規須年滿{dr1("CapAge1")}以上"
                End If

                'vp.TPlanID
                PlanKind = "&nbsp;" '組合計畫名稱
                Select Case $"{dr1("TPlanID")}"
                    Case "28"
                        'vp.PlanName
                        PlanKind = TIMS.GET_PlanKindName(Me, objconn, $"{dr1("PlanName")}", $"{dr1("OrgKind")}")
                    Case Else
                        PlanKind = $"{dr1("PlanName")}"
                End Select

                OrgName = If(IsDBNull(dr1("OrgName")), "", Trim($"{dr1("OrgName")}"))
                DISTNAME2 = $"{dr1("DISTNAME2")}"
                DISTANCE_NP = $"{dr1("DISTANCE_NP")}"

                OrgKind2 = If(IsDBNull(dr1("OrgKind2")), "", $"{dr1("OrgKind2")}")
                jobName = If(IsDBNull(dr1("jobName")), "", $"{dr1("jobName")}")
                CCName = If(IsDBNull(dr1("CCName")), "", $"{dr1("CCName")}")
                ClassName = If(IsDBNull(dr1("ClassName")), "", $"{dr1("ClassName")}")
                GCName = If(IsDBNull(dr1("GCName")), "", $"{dr1("GCName")}")
                ClassID = If(IsDBNull(dr1("ClassID")), "", If($"{dr1("ClassID")}" = "", "&nbsp;", $"{dr1("ClassID")}"))
                TNum = If(IsDBNull(dr1("TNum")), "", $"{dr1("TNum")}")
                STDate = If(IsDBNull(dr1("STDate")), "", $"{dr1("STDate")}")
                FDDate = If(IsDBNull(dr1("FDDate")), "", $"{dr1("FDDate")}")
                Weeks = If(IsDBNull(dr1("Weeks")), "", $"{dr1("Weeks")}")
                Thours = If(IsDBNull(dr1("Thours")), "", $"{dr1("Thours")}")
                Week = If(IsDBNull(dr1("Week")), "", $"{dr1("Week")}")
                TrainDemain = If(IsDBNull(dr1("TrainDemain")), "", $"{dr1("TrainDemain")}")
                TMIDCORRECT = If(IsDBNull(dr1("TMIDCORRECT")), "", $"{dr1("TMIDCORRECT")}")
                TGOVEXAM = If(IsDBNull(dr1("TGOVEXAM")), "", $"{dr1("TGOVEXAM")}")
                GOVAGENAME = If(IsDBNull(dr1("GOVAGENAME")), "", $"{dr1("GOVAGENAME")}")
                TGOVEXAMNAME = If(IsDBNull(dr1("TGOVEXAMNAME")), "", $"{dr1("TGOVEXAMNAME")}")
                ENTERSUPPLYSTYLE = If(IsDBNull(dr1("ENTERSUPPLYSTYLE")), "", $"{dr1("ENTERSUPPLYSTYLE")}")
                Pur = If(IsDBNull(dr1("PlanCause")), "", "單位核心能力介紹：<br/>" & $"{dr1("PlanCause")}") & "<br/><br/>"  '計畫緣由
                Pur += If(IsDBNull(dr1("PurScience")), "", "知識：<br/>" & $"{dr1("PurScience")}") & "<br/><br/>"  '目標-學科
                Pur += If(IsDBNull(dr1("PurTech")), "", "技能：<br/>" & $"{dr1("PurTech")}") & "<br/><br/>"  '目標-術科
                Pur += If(IsDBNull(dr1("PurMoral")), "", "學習成效：<br/>" & $"{dr1("PurMoral")}") & "<br/><br/>"  '目標-品德
                Select Case $"{dr1("TPlanID")}"
                    Case "28"
                        strMyAppStage = $"{dr1("MyAppStage")}"  '申請階段，by:20181023
                End Select

                Dim str51 As String = ""
                Select Case $"{dr1("FuncLevel")}"
                    Case "01"
                        str51 &= " 級別1(能夠在可預計及有規律的情況中，在密切監督及清楚指示下，執行常規性及重複性的工作。且通常不需要特殊訓練、教育及專業知識與技術)<br/>"
                    Case "02"
                        str51 &= " 級別2(能夠在大部分可預計及有規律的情況中，在經常性監督下，按指導進行需要某些判斷及理解性的工作。需具備基本知識、技術)<br/>"
                    Case "03"
                        str51 &= " 級別3(能夠在部分變動及非常規性的情況中，在一般監督下，獨立完成工作。需要一定程度的專業知識與技術及少許的判斷能力)<br/>"
                    Case "04"
                        str51 &= " 級別4(能夠在經常變動的情況中，在少許監督下，獨立執行涉及規劃設計且需要熟練技巧的工作。需要具備相當的專業知識與技術，及作判斷及決定的能力)<br/>"
                    Case "05"
                        str51 &= " 級別5(能夠在複雜變動的情況中，在最少監督下，自主完成工作。需要具備應用、整合、系統化的專業知識與技術及策略思考與判斷能力)<br/>"
                    Case "06"
                        str51 &= " 級別6(能夠在高度複雜變動的情況中，應用整合的專業知識與技術，獨立完成專業與創新的工作。需要具備策略思考、決策及原創能力)<br/>"
                End Select
                Pur += "職能級別：<br/>" & str51

                TMethod = ""
                '□講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）　　　　　　　　
                '□討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）
                '□演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）
                '□其他教學方法：
                If $"{dr1("TMethodC01")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）")
                End If
                If $"{dr1("TMethodC02")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）")
                End If
                If $"{dr1("TMethodC03")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）")
                End If
                If $"{dr1("TMethodC99")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "其他教學方法：", dr1("TMethodOth"))
                End If

                tPOWERNEED = "" & vbCrLf
                tPOWERNEED &= $" 1.產業人力需求調查：<br/>{dr1("POWERNEED1")}<br/>"
                'ttm1 &= " (應論述調查期間、區域範圍、調查對象、產業發展趨勢及該產業之訓練需求)<br/>" 
                tPOWERNEED &= $" 2.區域人力需求調查：<br/>{dr1("POWERNEED2")}<br/>"
                'ttm1 &= " (依產業人力需求調查結果，進行區域性的人力需求調查，應論述調查期間、區域範圍、調查對象及該產業於該區域之訓練需求)<br/>" 
                tPOWERNEED &= $" 3.訓練需求概述：<br/>{dr1("POWERNEED3")}<br/>"
                If flag_OJT22071401 AndAlso $"{dr1("POLICYREL")}" <> "" Then
                    tPOWERNEED &= $" 4.與政策性產業課程之關聯性概述：<br/>{dr1("POLICYREL")}<br/>"
                End If
                If $"{dr1("POWERNEED4CHK")}" = TIMS.cst_YES Then
                    tPOWERNEED &= $" 課程須符合目的事業主管機關相關規定：<br/>{dr1("POWERNEED4")}<br/>"
                End If

                Dim ss3 As String = $"sType={sType}&OCID={OCID}&TPlanID={dr1("TPlanID")}"
                dt2a = TIMS.Get_TeacherInfoDt2a(objconn, ss3)
                dt2b = TIMS.Get_TeacherInfoDt2b(objconn, ss3)
                'Plan_TrainPlace 裝備與設施
                dt3 = TIMS.Get_TrainPlaceDt3(objconn, ss3)
                If TIMS.dtHaveDATA(dt3) Then
                    Dim dr3 As DataRow = dt3.Rows(0)
                    'TeacherDesc = If(IsDBNull(dr("TeacherDesc")), "", $"{dr("TeacherDesc")}")
                    'TeacherDesc2 = If(IsDBNull(dr("TeacherDesc2")), "", $"{dr("TeacherDesc2")}")
                    'TeacherDesc = If(IsDBNull(dr("TeacherDesc")), "<br/>", Replace($"{dr("TeacherDesc")}", vbCrLf, "<br/>"))
                    'TeacherDesc2 = If(IsDBNull(dr("TeacherDesc2")), "<br/>", Replace($"{dr("TeacherDesc2")}", vbCrLf, "<br/>"))
                    SName = If(IsDBNull(dr3("SName")), "", $"{dr3("SName")}")
                    Note = If(IsDBNull(dr3("Note")), "<br/>", Replace($"{dr3("Note")}", vbCrLf, "<br/>")) '訓練費用編列說明
                    CapAll = If(IsDBNull(dr3("CapAll")), "", $"{dr3("CapAll")}")
                    Address1 = If(IsDBNull(dr3("Address1")), "", $"{dr3("Address1")}")
                    Connum1 = If(IsDBNull(dr3("Connum1")), "", $"{dr3("Connum1")}")
                    HWDesc1 = If(IsDBNull(dr3("HWDesc1")), "", $"{dr3("HWDesc1")}")
                    Address2 = If(IsDBNull(dr3("Address2")), "", $"{dr3("Address2")}")
                    Connum2 = If(IsDBNull(dr3("Connum2")), "", $"{dr3("Connum2")}")
                    HWDesc2 = If(IsDBNull(dr3("HWDesc2")), "", $"{dr3("HWDesc2")}")
                    Address3 = If(IsDBNull(dr3("Address3")), "", $"{dr3("Address3")}")
                    Connum3 = If(IsDBNull(dr3("Connum3")), "", $"{dr3("Connum3")}")
                    HWDesc3 = If(IsDBNull(dr3("HWDesc3")), "", $"{dr3("HWDesc3")}")
                    Address4 = If(IsDBNull(dr3("Address4")), "", $"{dr3("Address4")}")
                    Connum4 = If(IsDBNull(dr3("Connum4")), "", $"{dr3("Connum4")}")
                    HWDesc4 = If(IsDBNull(dr3("HWDesc4")), "", $"{dr3("HWDesc4")}")
                    'If(IsDBNull(dr3("OtherDesc3")), "", $"{dr3("OtherDesc3")}")
                    OtherDesc3 = $"{dr3("OtherDesc3")}"
                    TMScience = If(IsDBNull(dr3("TMScience")), "", $"{dr3("TMScience")}")
                    TMTech = If(Not IsDBNull(dr3("TMTech")), If(Trim($"{dr3("TMTech")}") = "", "<br/>", $"{dr3("TMTech")}"), "<br/>")
                End If

                'dt4:PLAN_TRAINDESC
                dt4 = TIMS.Get_TRAINDESCDt4(objconn, ss3)
                iALL_EHOURS = GET_TRAINDESCD_HOURS(dt4, "EHOURS") '技檢訓練時數
                iALL_AIAHOUR = GET_TRAINDESCD_HOURS(dt4, "AIAHOUR") 'AI應用時數
                iALL_WNLHOUR = GET_TRAINDESCD_HOURS(dt4, "WNLHOUR") '職場續航時數

                'If TIMS.dtHaveDATA( dt4) Then
                '    Dim dr4 As DataRow = dt4.Rows(0)
                '    STrainDate = If(IsDBNull(dr("STrainDate")), "", $"{dr("STrainDate")}")
                '    TechTime = If(IsDBNull(dr("TechTime")), "", $"{dr("TechTime")}")
                '    HOURS = If(IsDBNull(dr("HOURS")), "", $"{dr("HOURS")}")
                '    PCont = If(IsDBNull(dr("PCont")), "", $"{dr("PCont")}")
                '    Classification1 = If(IsDBNull(dr("Classification1")), "&nbsp;", $"{dr("Classification1")}")
                '    PLACENAME = If(IsDBNull(dr("PLACENAME")), "", $"{dr("PLACENAME")}")
                '    TeachCName2 = If(IsDBNull(dr("TeachCName")), "", $"{dr("TeachCName")}")
                'End If

                'PLAN_VERREPORT
                dt5 = TIMS.Get_VERREPORTDt5(objconn, ss3)
                If TIMS.dtHaveDATA(dt5) Then
                    Dim dr5 As DataRow = dt5.Rows(0)
                    Rec = If(dr5("Rec") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    RecDesc = If(IsDBNull(dr5("RecDesc")), "", $"{dr5("RecDesc")}")
                    Learn = If(dr5("Learn") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    LearnDesc = If(IsDBNull(dr5("LearnDesc")), "", $"{dr5("LearnDesc")}")
                    Act = If(dr5("ACT") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    ActDesc = If(IsDBNull(dr5("ActDesc")), "", $"{dr5("ActDesc")}")
                    Rst = If(dr5("Rst") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    ResultDesc = If(IsDBNull(dr5("ResultDesc")), "", $"{dr5("ResultDesc")}")
                    oth = If(dr5("oth") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    OtherDesc = If(IsDBNull(dr5("OtherDesc")), "", $"{dr5("OtherDesc")}")
                    s_OTHFACDESC23 = $"{dr5("OTHFACDESC23")}"
                    s_RMTNAME1 = $"{dr5("RMTNAME1")}"
                    s_RMTNAME2 = $"{dr5("RMTNAME2")}"

                    ISiCAPCOUR = $"{dr5("ISiCAPCOUR")}"
                    iCAPCOURDESC = $"{dr5("iCAPCOURDESC")}"
                    iCAPNUM = $"{dr5("iCAPNUM")}"
                    iCAPMARKDATE = $"{dr5("iCAPMARKDATE")}"
                    Recruit = If(IsDBNull(dr5("Recruit")), "", $"{dr5("Recruit")}")
                    Selmethod = If(IsDBNull(dr5("Selmethod")), "", $"{dr5("Selmethod")}")
                    Inspire = If(IsDBNull(dr5("Inspire")), "", $"{dr5("Inspire")}")

                    DefGovCost = If(IsDBNull(dr5("DefGovCost")), "", $"{dr5("DefGovCost")}")
                    DefStdCost = If(IsDBNull(dr5("DefStdCost")), "", $"{dr5("DefStdCost")}")
                    Total1 = If(IsDBNull(dr5("Total1")), "", $"{dr5("Total1")}")
                    Total2 = If(IsDBNull(dr5("Total2")), "", $"{dr5("Total2")}")
                    If $"{dr5("tplanid")}" = "54" Then
                        DefGovCost = $"{dr5("totalcost")}"
                        DefStdCost = "0"
                        Total1 = If(IsDBNull(dr5("Total1b")), "", $"{dr5("Total1b")}")
                        Total2 = "0"
                    End If
                    Total4 = If(IsDBNull(dr5("Total4")), "", $"{dr5("Total4")}")
                    Total3 = If(IsDBNull(dr5("Total3")), "", $"{dr5("Total3")}")
                    FirRst = If(IsDBNull(dr5("FirRst")), "", $"{dr5("FirRst")}")
                    SecRst = If(IsDBNull(dr5("SecRst")), "", $"{dr5("SecRst")}")
                    MEMO8 = $"{If($"{dr5("MEMO8")}" <> "", cst_images_rptpic_yy_jpg, $"{cst_images_rptpic_no_jpg}{cst_msg_memo8a}")}{dr5("MEMO8")}"
                    MEMO82 = $"{If($"{dr5("MEMO82")}" <> "", $"{cst_images_rptpic_yes_jpg}{dr5("MEMO82")}", "")}"
                End If

                'PLAN_ABILITY
                dt6 = TIMS.GET_PLAN_ABILITYdt(objconn, ss3)

            Case TIMS.cst_sType_B_未轉班
                '未送出()
                Dim PMS1 As New Hashtable From {{"ISAPPRPAPER", "Y"}}
                dt1 = TIMS.Get_PLANINFO_TB(objconn, PlanID, ComIDNO, SeqNo, PMS1)
                If TIMS.dtNODATA(dt1) Then
                    Common.MessageBox(Me, cst_errorMsg1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, cst_errorMsg1)
                    Exit Sub
                End If

                Dim dr1 As DataRow = dt1.Rows(0)
                If $"{dr1("CapAge1")}" <> "" AndAlso TIMS.CINT1(dr1("CapAge1")) >= cst_AgeOtherDef Then
                    yearsOld_N = $"應符合相關法規須年滿{dr1("CapAge1")}以上"
                End If
                'vp.TPlanID
                '組合計畫名稱
                Select Case $"{dr1("TPlanID")}"
                    Case "28"
                        'vp.PlanName
                        PlanKind = TIMS.GET_PlanKindName(Me, objconn, $"{dr1("PlanName")}", $"{dr1("OrgKind")}")
                    Case Else
                        PlanKind = $"{dr1("PlanName")}"
                End Select

                OrgName = If(IsDBNull(dr1("OrgName")), "", Trim($"{dr1("OrgName")}"))
                DISTNAME2 = $"{dr1("DISTNAME2")}"
                DISTANCE_NP = $"{dr1("DISTANCE_NP")}"

                OrgKind2 = If(IsDBNull(dr1("OrgKind2")), "", $"{dr1("OrgKind2")}")
                jobName = If(IsDBNull(dr1("jobName")), "", $"{dr1("jobName")}")
                CCName = If(IsDBNull(dr1("CCName")), "", $"{dr1("CCName")}")
                'RESULTBUTTON: Y/R cst_ResultButton_尚未送出_未送出
                Select Case $"{dr1("RESULTBUTTON")}"
                    Case TIMS.cst_ResultButton_尚未送出_待送審
                        flag_RESULTBUTTON_YR = True 'ClassName &= "(未送出)"
                    Case TIMS.cst_ResultButton_尚未送出_未送出
                        flag_RESULTBUTTON_YR = True 'ClassName &= "(未送出)"
                End Select
                ClassName = $"{dr1("ClassName")}"
                If flag_RESULTBUTTON_YR Then ClassName &= "(未送出)"

                GCName = If(IsDBNull(dr1("GCName")), "", $"{dr1("GCName")}")
                ClassID = If(IsDBNull(dr1("ClassID")), "", $"{dr1("ClassID")}")
                TNum = If(IsDBNull(dr1("TNum")), "", $"{dr1("TNum")}")
                STDate = If(IsDBNull(dr1("STDate")), "", $"{dr1("STDate")}")
                FDDate = If(IsDBNull(dr1("FDDate")), "", $"{dr1("FDDate")}")
                Weeks = If(IsDBNull(dr1("Weeks")), "", $"{dr1("Weeks")}")
                Thours = If(IsDBNull(dr1("Thours")), "", $"{dr1("Thours")}")
                Week = If(IsDBNull(dr1("Week")), "", $"{dr1("Week")}")
                TrainDemain = If(IsDBNull(dr1("TrainDemain")), "", $"{dr1("TrainDemain")}")
                TMIDCORRECT = If(IsDBNull(dr1("TMIDCORRECT")), "", $"{dr1("TMIDCORRECT")}")
                TGOVEXAM = If(IsDBNull(dr1("TGOVEXAM")), "", $"{dr1("TGOVEXAM")}")
                GOVAGENAME = If(IsDBNull(dr1("GOVAGENAME")), "", $"{dr1("GOVAGENAME")}")
                TGOVEXAMNAME = If(IsDBNull(dr1("TGOVEXAMNAME")), "", $"{dr1("TGOVEXAMNAME")}")
                ENTERSUPPLYSTYLE = If(IsDBNull(dr1("ENTERSUPPLYSTYLE")), "", $"{dr1("ENTERSUPPLYSTYLE")}")
                Pur = If(IsDBNull(dr1("PlanCause")), "", "單位核心能力介紹：<br/>" & $"{dr1("PlanCause")}") & "<br/><br/>"  '計畫緣由
                Pur += If(IsDBNull(dr1("PurScience")), "", "知識：<br/>" & $"{dr1("PurScience")}") & "<br/><br/>"  '目標-學科
                Pur += If(IsDBNull(dr1("PurTech")), "", "技能：<br/>" & $"{dr1("PurTech")}") & "<br/><br/>"  '目標-術科
                Pur += If(IsDBNull(dr1("PurMoral")), "", "學習成效：<br/>" & $"{dr1("PurMoral")}") & "<br/><br/>"  '目標-品德
                Select Case $"{dr1("TPlanID")}"
                    Case "28"
                        strMyAppStage = $"{dr1("MyAppStage")}"  '申請階段，by:20181023
                End Select

                Dim str51 As String = ""
                Select Case $"{dr1("FuncLevel")}"
                    Case "01"
                        str51 &= " 級別1(能夠在可預計及有規律的情況中，在密切監督及清楚指示下，執行常規性及重複性的工作。且通常不需要特殊訓練、教育及專業知識與技術)<br/>"
                    Case "02"
                        str51 &= " 級別2(能夠在大部分可預計及有規律的情況中，在經常性監督下，按指導進行需要某些判斷及理解性的工作。需具備基本知識、技術)<br/>"
                    Case "03"
                        str51 &= " 級別3(能夠在部分變動及非常規性的情況中，在一般監督下，獨立完成工作。需要一定程度的專業知識與技術及少許的判斷能力)<br/>"
                    Case "04"
                        str51 &= " 級別4(能夠在經常變動的情況中，在少許監督下，獨立執行涉及規劃設計且需要熟練技巧的工作。需要具備相當的專業知識與技術，及作判斷及決定的能力)<br/>"
                    Case "05"
                        str51 &= " 級別5(能夠在複雜變動的情況中，在最少監督下，自主完成工作。需要具備應用、整合、系統化的專業知識與技術及策略思考與判斷能力)<br/>"
                    Case "06"
                        str51 &= " 級別6(能夠在高度複雜變動的情況中，應用整合的專業知識與技術，獨立完成專業與創新的工作。需要具備策略思考、決策及原創能力)<br/>"
                End Select
                Pur += "職能級別：<br/>" & str51

                TMethod = ""
                '□講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）　　　　　　　　
                '□討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）
                '□演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）
                '□其他教學方法：
                If $"{dr1("TMethodC01")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "講授教學法（運用敘述或講演的方式，傳遞教材知識的一種教學方法，提供相關教材或講義）")
                End If
                If $"{dr1("TMethodC02")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "討論教學法（指團體成員齊聚一起，經由說、聽和觀察的過程，彼此溝通意見，由講師帶領達成教學目標）")
                End If
                If $"{dr1("TMethodC03")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "演練教學法（由講師的帶領下透過設備或教材，進行練習、表現和實作，親自解說示範的技能或程序的一種教學方法）")
                End If
                If $"{dr1("TMethodC99")}" <> "" Then
                    TMethod &= String.Concat(If(TMethod <> "", "<br/>", ""), cst_images_rptpic_yes_jpg, "其他教學方法：", dr1("TMethodOth"))
                End If

                tPOWERNEED = "" & vbCrLf
                tPOWERNEED &= $" 1.產業人力需求調查：<br/>{dr1("POWERNEED1")}<br/>"
                'ttm1 &= " (應論述調查期間、區域範圍、調查對象、產業發展趨勢及該產業之訓練需求)<br/>" 
                tPOWERNEED &= $" 2.區域人力需求調查：<br/>{dr1("POWERNEED2")}<br/>"
                'ttm1 &= " (依產業人力需求調查結果，進行區域性的人力需求調查，應論述調查期間、區域範圍、調查對象及該產業於該區域之訓練需求)<br/>" 
                tPOWERNEED &= $" 3.訓練需求概述：<br/>{dr1("POWERNEED3")}<br/>"
                If flag_OJT22071401 AndAlso $"{dr1("POLICYREL")}" <> "" Then
                    tPOWERNEED &= $" 4.與政策性產業課程之關聯性概述：<br/>{dr1("POLICYREL")}<br/>"
                End If
                If $"{dr1("POWERNEED4CHK")}" = TIMS.cst_YES Then
                    tPOWERNEED &= $" 課程須符合目的事業主管機關相關規定：<br/>{dr1("POWERNEED4")}<br/>"
                End If

                Dim ss3 As String = $"sType={sType}&PlanID={PlanID}&ComIDNO={ComIDNO}&SeqNo={SeqNo}&TPlanID={dr1("TPlanID")}"
                dt2a = TIMS.Get_TeacherInfoDt2a(objconn, ss3)
                dt2b = TIMS.Get_TeacherInfoDt2b(objconn, ss3)
                dt3 = TIMS.Get_TrainPlaceDt3(objconn, ss3)
                If TIMS.dtHaveDATA(dt3) Then
                    Dim dr3 As DataRow = dt3.Rows(0)
                    'TeacherDesc = If(IsDBNull(dr("TeacherDesc")), "", $"{dr("TeacherDesc")}")
                    'TeacherDesc2 = If(IsDBNull(dr("TeacherDesc2")), "", $"{dr("TeacherDesc2")}")
                    'TeacherDesc = If(IsDBNull(dr("TeacherDesc")), "<br/>", Replace($"{dr("TeacherDesc")}", vbCrLf, "<br/>"))
                    'TeacherDesc2 = If(IsDBNull(dr("TeacherDesc2")), "<br/>", Replace($"{dr("TeacherDesc2")}", vbCrLf, "<br/>"))
                    SName = If(IsDBNull(dr3("SName")), "", $"{dr3("SName")}")
                    'Note = If(IsDBNull(dr3("Note")), "", $"{dr3("Note")}")
                    Note = If(IsDBNull(dr3("Note")), "<br/>", Replace($"{dr3("Note")}", vbCrLf, "<br/>")) '訓練費用編列說明
                    CapAll = If(IsDBNull(dr3("CapAll")), "", $"{dr3("CapAll")}")
                    Address1 = If(IsDBNull(dr3("Address1")), "", $"{dr3("Address1")}")
                    Connum1 = If(IsDBNull(dr3("Connum1")), "", $"{dr3("Connum1")}")
                    HWDesc1 = If(IsDBNull(dr3("HWDesc1")), "", $"{dr3("HWDesc1")}")
                    Address2 = If(IsDBNull(dr3("Address2")), "", $"{dr3("Address2")}")
                    Connum2 = If(IsDBNull(dr3("Connum2")), "", $"{dr3("Connum2")}")
                    HWDesc2 = If(IsDBNull(dr3("HWDesc2")), "", $"{dr3("HWDesc2")}")
                    Address3 = If(IsDBNull(dr3("Address3")), "", $"{dr3("Address3")}")
                    Connum3 = If(IsDBNull(dr3("Connum3")), "", $"{dr3("Connum3")}")
                    HWDesc3 = If(IsDBNull(dr3("HWDesc3")), "", $"{dr3("HWDesc3")}")
                    Address4 = If(IsDBNull(dr3("Address4")), "", $"{dr3("Address4")}")
                    Connum4 = If(IsDBNull(dr3("Connum4")), "", $"{dr3("Connum4")}")
                    HWDesc4 = If(IsDBNull(dr3("HWDesc4")), "", $"{dr3("HWDesc4")}")
                    OtherDesc3 = If(IsDBNull(dr3("OtherDesc3")), "", $"{dr3("OtherDesc3")}")
                    TMScience = If(IsDBNull(dr3("TMScience")), "", $"{dr3("TMScience")}")
                    TMTech = If(Not IsDBNull(dr3("TMTech")), If(Trim($"{dr3("TMTech")}") = "", "<br/>", $"{dr3("TMTech")}"), "<br/>")
                End If

                dt4 = TIMS.Get_TRAINDESCDt4(objconn, ss3)
                iALL_EHOURS = GET_TRAINDESCD_HOURS(dt4, "EHOURS") '技檢訓練時數
                iALL_AIAHOUR = GET_TRAINDESCD_HOURS(dt4, "AIAHOUR") 'AI應用時數
                iALL_WNLHOUR = GET_TRAINDESCD_HOURS(dt4, "WNLHOUR") '職場續航時數

                'If TIMS.dtHaveDATA( dt4) Then
                '    dr = dt4.Rows(0)
                '    STrainDate = If(IsDBNull(dr("STrainDate")), "", $"{dr("STrainDate")}")
                '    TechTime = If(IsDBNull(dr("TechTime")), "", $"{dr("TechTime")}")
                '    HOURS = If(IsDBNull(dr("HOURS")), "", $"{dr("HOURS")}")
                '    PCont = If(IsDBNull(dr("PCont")), "", $"{dr("PCont")}")
                '    Classification1 = If(IsDBNull(dr("Classification1")), "&nbsp;", $"{dr("Classification1")}")
                '    PLACENAME = If(IsDBNull(dr("PLACENAME")), "", $"{dr("PLACENAME")}")
                '    TeachCName2 = If(IsDBNull(dr("TeachCName")), "", $"{dr("TeachCName")}")
                'End If

                'PLAN_VERREPORT
                dt5 = TIMS.Get_VERREPORTDt5(objconn, ss3)
                If TIMS.dtHaveDATA(dt5) Then
                    Dim dr5 As DataRow = dt5.Rows(0)
                    Rec = If(dr5("Rec") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    RecDesc = If(IsDBNull(dr5("RecDesc")), "", $"{dr5("RecDesc")}")
                    Learn = If(dr5("Learn") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    LearnDesc = If(IsDBNull(dr5("LearnDesc")), "", $"{dr5("LearnDesc")}")
                    Act = If(dr5("Act") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    ActDesc = If(IsDBNull(dr5("ActDesc")), "", $"{dr5("ActDesc")}")
                    Rst = If(dr5("Rst") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    ResultDesc = If(IsDBNull(dr5("ResultDesc")), "", $"{dr5("ResultDesc")}")
                    oth = If(dr5("oth") = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
                    OtherDesc = If(IsDBNull(dr5("OtherDesc")), "", $"{dr5("OtherDesc")}")
                    s_OTHFACDESC23 = $"{dr5("OTHFACDESC23")}"
                    s_RMTNAME1 = $"{dr5("RMTNAME1")}"
                    s_RMTNAME2 = $"{dr5("RMTNAME2")}"

                    ISiCAPCOUR = $"{dr5("ISiCAPCOUR")}"
                    iCAPCOURDESC = $"{dr5("iCAPCOURDESC")}"
                    iCAPNUM = $"{dr5("iCAPNUM")}"
                    iCAPMARKDATE = $"{dr5("iCAPMARKDATE")}"
                    Recruit = If(IsDBNull(dr5("Recruit")), "", $"{dr5("Recruit")}")
                    Selmethod = If(IsDBNull(dr5("Selmethod")), "", $"{dr5("Selmethod")}")
                    Inspire = If(IsDBNull(dr5("Inspire")), "", $"{dr5("Inspire")}")

                    DefGovCost = If(IsDBNull(dr5("DefGovCost")), "", $"{dr5("DefGovCost")}")
                    DefStdCost = If(IsDBNull(dr5("DefStdCost")), "", $"{dr5("DefStdCost")}")
                    Total1 = If(IsDBNull(dr5("Total1")), "", $"{dr5("Total1")}")
                    Total2 = If(IsDBNull(dr5("Total2")), "", $"{dr5("Total2")}")
                    If $"{dr5("tplanid")}" = "54" Then
                        DefGovCost = $"{dr5("totalcost")}"
                        DefStdCost = "0"
                        Total1 = If(IsDBNull(dr5("Total1b")), "", $"{dr5("Total1b")}")
                        Total2 = "0"
                    End If
                    Total4 = If(IsDBNull(dr5("Total4")), "", $"{dr5("Total4")}")
                    Total3 = If(IsDBNull(dr5("Total3")), "", $"{dr5("Total3")}")
                    FirRst = If(IsDBNull(dr5("FirRst")), "", $"{dr5("FirRst")}")
                    SecRst = If(IsDBNull(dr5("SecRst")), "", $"{dr5("SecRst")}")
                    MEMO8 = $"{If($"{dr5("MEMO8")}" <> "", cst_images_rptpic_yy_jpg, $"{cst_images_rptpic_no_jpg}{cst_msg_memo8a}")}{dr5("MEMO8")}"
                    MEMO82 = $"{If($"{dr5("MEMO82")}" <> "", $"{cst_images_rptpic_yes_jpg}{dr5("MEMO82")}", "")}"
                End If

                'PLAN_ABILITY
                dt6 = TIMS.GET_PLAN_ABILITYdt(objconn, ss3)
        End Select

        'count1 = dt1.Rows.Count
        'count2 = dt2.Rows.Count
        'count2b = dt2b.Rows.Count
        'count3 = dt3.Rows.Count
        'count4 = dt4.Rows.Count
        'count5 = dt5.Rows.Count

        Const cst_iColspanAll_1 As Integer = 18

        dt1 = Nothing
        dt3 = Nothing

        'Dim j As Int16 = 0
        'If dt1.Rows.Count > 0 Then dr = dt1.Rows(0)

        '----- 表首  start
        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", cst_style_c1)
        'rptTb.Attributes.Add("background", cst_bgimg_c1)
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 0)
        print_content.Controls.Add(rptTb)

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:20pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.InnerHtml = $"勞動部勞動力發展署{DISTNAME2}"

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:20pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.InnerHtml = PlanKind

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:20pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        'rptCell.InnerHtml = Years.ToString() & "年度  訓練班別計畫表"

        '=20181023 依照"TIMS 108年 增修項目"，表頭修正為"年度+申請階段+訓練班別計畫表" start
        Dim myYears As String = TIMS.ClearSQM(Years)
        myYears = If(Len(myYears) < 3, Right($"000{myYears}", 3), myYears)
        Dim str_RESULTBTN As String = If(flag_RESULTBUTTON_YR, "(未送出)", "")

        Dim str_Title_Txt As String = If(myYears >= "108" AndAlso strMyAppStage <> "", $"{Years}年度 {strMyAppStage} 訓練班別計畫表{str_RESULTBTN}", $"{Years}年度  訓練班別計畫表{str_RESULTBTN}")
        rptCell.InnerHtml = str_Title_Txt

        '--- 列印時間 090811 andy add
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "right")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1) '18
        rptCell.InnerHtml = String.Format("列印日期： {0}", Now.ToString("yyyy/MM/dd HH:mm"))
        '----- 

        rptTb = New HtmlTable
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 1)
        rptTb.Attributes.Add("style", cst_style_c2)
        'rptTb.Attributes.Add("background", cst_bgimg_c1)
        rptTb.Attributes.Add("bordercolor", "black")
        print_content.Controls.Add(rptTb)

        '----- 
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("width", "18%")
        'rptCell.Attributes.Add("width", "150px")
        rptCell.InnerHtml = "訓練單位名稱"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 4)
        'rptCell.Attributes.Add("width", "150px")
        rptCell.Attributes.Add("width", "20%")
        rptCell.InnerHtml = If(PrintOrg = "Y", OrgName, "&nbsp")

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "18%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "辦理方式" '"訓練計畫"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "44%")
        'rptCell.Attributes.Add("width", "200px")
        rptCell.Attributes.Add("colspan", 8)
        rptCell.InnerHtml = DISTANCE_NP 'YearPlan

        '----- 2
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "訓練職類" '"訓練業別"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 4)
        rptCell.InnerHtml = jobName

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "訓練職能"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 8)
        rptCell.InnerHtml = CCName

        '----- 3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "課程名稱"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 4)
        rptCell.InnerHtml = ClassName

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "業別分類代碼" '"經費分類代碼"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 8)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = GCName 'CName

        '----- 4
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "政策性產業" '"課程班別編號"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 4)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = If(s_KNAME1920 = "", "&nbsp;&nbsp;&nbsp;&nbsp;", s_KNAME1920)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "訓練人數"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = TNum

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "起迄日期"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = String.Format("自{0}<br/>至{1}", TIMS.Cdate3(STDate), TIMS.Cdate3(FDDate))

        '----- 5
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "上課時間"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 4)
        rptCell.InnerHtml = Replace(Weeks, "; ", "; <br/>")

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "訓練時數"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        'rptCell.Attributes.Add("width", "30px")
        rptCell.InnerHtml = Thours

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "訓練週數"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        rptCell.InnerHtml = Week


        '----- 6-0
        '多一行說明
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("valign", "center")
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "訓練計畫內容"

        '----- 6-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        Dim intCnt As Integer = 0 '計算規劃與執行能力 rowspan
        intCnt += 1 '訓練需求調查(是否瞭解區域產業需求)
        intCnt += 1 '訓練目標(是否符合需求並配合訓練單位核心能力)
        'intCnt += 1 '(師資遴選辦法說明)..
        'intCnt += 1 '(助教遴選辦法說明)..
        intCnt += 1 '(學員資格)..'(訓練費用編列說明)..
        'intCnt += 1 '(訓練費用編列說明)..

        '師資(與訓練目標是否切合)
        intCnt += If(TIMS.dtHaveDATA(dt2a), (1 + dt2a.Rows.Count), 2)
        '助教(與訓練目標是否切合)
        intCnt += If(TIMS.dtHaveDATA(dt2b), (1 + dt2b.Rows.Count), 2)

        If TIMS.dtHaveDATA(dtBPA) Then intCnt += 1 + dtBPA.Rows.Count '有資料才顯示(企業包班)

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("valign", "center")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", intCnt)
        'rptCell.InnerHtml = "規劃與執行能力"

        'rptCell.Attributes.Add("colspan", 1)
        'rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("width", "30px")

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "" '規劃與執行能力"
        'rptCell.Attributes.Add("width", "50px")
        'rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("colspan", 1)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.InnerHtml = "訓練需求調查(是否瞭解區域產業需求)"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 15)
        'rptCell.Attributes.Add("width", "20%")
        'rptCell.Attributes.Add("width", "100px")
        rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
        rptCell.InnerHtml = tPOWERNEED 'TrainDemain

        '----- ||
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "規劃與執行能力" '規劃與執行能力"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "18%")
        'rptCell.Attributes.Add("width", "100px")
        rptCell.InnerHtml = "訓練目標(是否符合需求並配合訓練單位核心能力)"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 15)
        'rptCell.Attributes.Add("width", "40%")
        Pur = Replace(Pur, vbCrLf, "<br/>") '單位核心能力介紹：
        rptCell.InnerHtml = Pur '訓練目標(是否符合需求並配合訓練單位核心能力)


        '師資(與訓練目標是否切合)
        '----- 6-2-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", If(dt2a.Rows.Count > 0, 1 + dt2a.Rows.Count, 2))
        rptCell.InnerHtml = "" '規劃與執行能力"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("style", "border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", If(dt2a.Rows.Count > 0, 1 + dt2a.Rows.Count, 2))
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.InnerHtml = "師資(與訓練目標是否切合)"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "7%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "姓名"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "12%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "授課師資條件"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學歷"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 3)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "工作經驗與年資"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "相關證照"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "12%")
        'rptCell.Attributes.Add("width", "250px")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "專業領域"
        '----- 表首 end 

        '----- 6-2-2
        If TIMS.dtHaveDATA(dt2a) Then
            '----- 明細(一) start
            For i As Integer = 0 To dt2a.Rows.Count - 1
                Dim dr2 As DataRow = dt2a.Rows(i)
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
                rptCell.InnerHtml = "" '規劃與執行能力"

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                rptCell.InnerHtml = "" '"師資(與訓練目標是否切合)"

                '姓名
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                rptCell.InnerHtml = If(IsDBNull(dr2("TeachCName")), "&nbsp;", $"{dr2("TeachCName")}")

                '授課師資條件
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                rptCell.InnerHtml = If(IsDBNull(dr2("TeacherDesc")), cst_TeacherDesc_default1, $"{dr2("TeacherDesc")}")

                '學歷
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("colspan", 2)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("DegreeName")), "&nbsp;", $"{dr2("DegreeName")}")

                '工作經驗與年資
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("ExpUnit1")), "&nbsp;", $"{dr2("ExpUnit1")}")

                '相關證照
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("ProLicense")), "&nbsp;", $"{dr2("ProLicense")}")

                '專業領域
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("Specialty1")), "&nbsp;", $"{dr2("Specialty1")}")
            Next
            '----- 明細(一) end
        Else
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "" '規劃與執行能力"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = "" '"師資(與訓練目標是否切合)"

            '姓名
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = "&nbsp;"

            '授課師資條件
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = cst_TeacherDesc_default1 '"(依計畫師資及助教資格標準表)"

            '學歷
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 2)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '工作經驗與年資
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '相關證照
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '專業領域
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"
        End If
        '-----  6-4

        '助教(與訓練目標是否切合)
        '----- 6-2-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", If(dt2b.Rows.Count > 0, 1 + dt2b.Rows.Count, 2))
        rptCell.InnerHtml = "" '規劃與執行能力"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("style", "border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", If(dt2b.Rows.Count > 0, 1 + dt2b.Rows.Count, 2))
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.InnerHtml = "助教(與訓練目標是否切合)"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "7%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "姓名"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "助教條件"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學歷"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 3)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "工作經驗與年資"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 3)
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "相關證照"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("width", "25%")
        'rptCell.Attributes.Add("width", "250px")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "專業領域"
        '-----  表首 end 

        '-----  6-2-2
        If TIMS.dtHaveDATA(dt2b) Then
            '-----  明細(一) start
            For i As Integer = 0 To dt2b.Rows.Count - 1
                Dim dr2 As DataRow = dt2b.Rows(i)
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
                rptCell.InnerHtml = "" '規劃與執行能力"

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                rptCell.InnerHtml = "" '"助教(與訓練目標是否切合)"

                '姓名
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                rptCell.InnerHtml = If(IsDBNull(dr2("TeachCName")), "&nbsp;", $"{dr2("TeachCName")}")

                '授課師資條件
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 2)
                'rptCell.InnerHtml = If(TeacherDesc2 = "", "(依計畫師資及助教資格標準表)", TeacherDesc2)
                rptCell.InnerHtml = If(IsDBNull(dr2("TeacherDesc")), "(依計畫師資及助教資格標準表)", $"{dr2("TeacherDesc")}")

                '學歷
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("colspan", 2)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("DegreeName")), "&nbsp;", $"{dr2("DegreeName")}")

                '工作經驗與年資
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("ExpUnit1")), "&nbsp;", $"{dr2("ExpUnit1")}")

                '相關證照
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("ProLicense")), "&nbsp;", $"{dr2("ProLicense")}")

                '專業領域
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr2("Specialty1")), "&nbsp;", $"{dr2("Specialty1")}")
            Next
            '----- 明細(一) end
        Else
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "" '規劃與執行能力"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = "" '"助教(與訓練目標是否切合)"

            '(助教)姓名
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = "&nbsp;"

            '授課師資條件
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            rptCell.InnerHtml = "(依計畫師資及助教資格標準表)"

            '學歷
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 2)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '工作經驗與年資
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '相關證照
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"

            '專業領域
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"
        End If

        '----- 6-5-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "" '規劃與執行能力"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.InnerHtml = "學員資格(是否明確敘述必備條件與適合對象)"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("valign", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("width", "85%")
        rptCell.Attributes.Add("colspan", 4)
        Dim s_QUALIFY As String = $"學　歷：{SName}<br>年　齡：{yearsOld_N}<br>資格條件：{CapAll}"
        rptCell.InnerHtml = s_QUALIFY

        'x
        'rptRow = New HtmlTableRow
        'rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 3)
        'rptCell.Attributes.Add("rowspan", "1")
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.InnerHtml = "訓練費用編列說明"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", "1")
        rptCell.Attributes.Add("colspan", 8)
        'rptCell.Attributes.Add("width", "25%")
        rptCell.InnerHtml = Note '訓練費用編列說明

        '企業包班 
        If TIMS.dtHaveDATA(dtBPA) Then Call Show_BUSPACKAGE(rptTb, rptRow, rptCell, dtBPA) '有資料才顯示(企業包班)
        '----- 6-5-2

        'rptRow = New HtmlTableRow
        'rptTb.Controls.Add(rptRow)

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("valign", "center")
        'rptCell.Attributes.Add("colspan", 1)
        ''rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = "資格條件:"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("colspan", 3)
        'rptCell.InnerHtml = CapAll

        '----- 7-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("rowspan", 7)
        'rptCell.Attributes.Add("width", "5%")
        rptCell.InnerHtml = "裝備與設施"

        For i_ADDRESS4 As Integer = 1 To 4
            Dim SITE_TXT As String = "" '場地地址
            Dim SITE_VAL As String = "" '場地地址
            Dim Connum_val As String = "" '容納人數
            Dim HWDesc_val As String = "" '硬體設施說明
            Select Case i_ADDRESS4
                Case 1
                    SITE_TXT = "學科場地地址1"
                    SITE_VAL = Address1
                    Connum_val = Connum1
                    HWDesc_val = HWDesc1
                Case 2
                    SITE_TXT = "學科場地地址2"
                    SITE_VAL = Address3
                    Connum_val = Connum3
                    HWDesc_val = HWDesc3
                Case 3
                    SITE_TXT = "術科場地地址1"
                    SITE_VAL = Address2
                    Connum_val = Connum2
                    HWDesc_val = HWDesc2
                Case 4
                    SITE_TXT = "術科場地地址2"
                    SITE_VAL = Address4
                    Connum_val = Connum4
                    HWDesc_val = HWDesc4
            End Select

            If i_ADDRESS4 <> 1 Then
                '第１行有配合抬頭，所以不用加
                '2010/05/31 add 學科場地地址2
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)
            End If

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("colspan", 2)
            'rptCell.Attributes.Add("width", "15%")
            rptCell.InnerHtml = SITE_TXT '"學科場地地址1"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 4)
            'rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = SITE_VAL 'Address1

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 2)
            rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "容納人數"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = Connum_val 'Connum '容納人數

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 2)
            rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "硬體設施說明"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 4)
            rptCell.Attributes.Add("width", "15%")
            rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
            rptCell.InnerHtml = HWDesc_val 'HWDesc''硬體設施說明
        Next

        '----- 7-3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "其他設施說明" '"其他器材設備"
        '其他設施說明
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 15)
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:10pt;" & str_fontfamily_c)
        rptCell.InnerHtml = If(s_OTHFACDESC23 <> "", s_OTHFACDESC23, OtherDesc3)

        '----- 7-3-2
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "遠距課程環境1" '"其他器材設備"
        '遠距課程環境1
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 15)
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:10pt;" & str_fontfamily_c)
        rptCell.InnerHtml = If(s_RMTNAME1 <> "", s_RMTNAME1, "")

        '----- 7-3-3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "遠距課程環境2" '"其他器材設備"
        '遠距課程環境2
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 15)
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:10pt;" & str_fontfamily_c)
        rptCell.InnerHtml = If(s_RMTNAME2 <> "", s_RMTNAME2, "")

        '----- 20090515 andy   拆成兩個table
        rptTb = New HtmlTable
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 1)
        rptTb.Attributes.Add("style", cst_style_c3)
        'rptTb.Attributes.Add("background", cst_bgimg_c1)
        rptTb.Attributes.Add("bordercolor", "black")
        print_content.Controls.Add(rptTb)

        '----- 8-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("valign", "center")
        'rptCell.Attributes.Add("style", "border-left-style@none;border-bottom-style: none;font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("colspan", 1)
        'rptCell.Attributes.Add("rowspan", 1)
        ''rptCell.Attributes.Add("width", "50px")
        'rptCell.Attributes.Add("width", "5%")
        ''rptCell.InnerHtml = "訓練模式特色與創新性"
        'rptCell.InnerHtml = ""

        'https://www.tgos.tw/TGOS/Web/Service/Order/TGOS_ServiceRecord_OrderView.aspx?moid=B71D9B25AA47CA8156E94ED8BA9E9546
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;border-right-style:none ;font-size:10pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("colspan", 1)
        rptCell.InnerHtml = ""

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("width", "90px")
        rptCell.Attributes.Add("width", "10%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;border-bottom-style: none;border-left-style : none;font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("rowspan", 1)
        rptCell.InnerHtml = "教學方法"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("width", "85%")
        'rptCell.Attributes.Add("width", "870px")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 15)
        rptCell.InnerHtml = TMethod 'TMScience

        '----- 8-2
        'rptRow = New HtmlTableRow
        'rptTb.Controls.Add(rptRow)

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        ''rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;border-right-style:none ;font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:10pt;" & str_fontfamily_c)
        ''rptCell.Attributes.Add("width", "50px")
        'rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("colspan", 1)
        'rptCell.InnerHtml = ""

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        ''rptCell.Attributes.Add("width", "90px")
        'rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("style", "word-wrap: break-word;border-bottom-style: none;border-left-style : none;font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("rowspan", 1)
        'rptCell.InnerHtml = "授課時段"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("width", "85%")
        ''rptCell.Attributes.Add("width", "870px")
        'rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("colspan", 9)
        ''If vTPERIOD28 = "" AndAlso TIMS.sUtl_ChkTest Then
        ''    vTPERIOD28 = "早上、下午、晚上" '測試用
        ''End If
        'rptCell.InnerHtml = vTPERIOD28 '"『早上、下午、晚上』"

        '----- 8-3
        'Dim fg_EHour_Use_TMID As Boolean = If(Convert.ToString(o_TMID) <> "" AndAlso Convert.ToString(o_TMID) = TIMS.cst_EHour_Use_TMID, True, False) '符合技能檢定訓練時數
        'Dim fg_OthHour_Use_TMID As Boolean = If(Convert.ToString(o_TMID) <> "" AndAlso Convert.ToString(o_TMID) = TIMS.cst_EHour_Use_TMID, True, False) '符合技能檢定訓練時數
        Dim fg_OthHour_Use_TMID As Boolean = False '符合技能檢定訓練時數'/AI應用時數/職場續航時數
        If Not fg_OthHour_Use_TMID AndAlso (iALL_EHOURS > 0) Then fg_OthHour_Use_TMID = True '技檢訓練時數
        If Not fg_OthHour_Use_TMID AndAlso (iALL_AIAHOUR > 0) Then fg_OthHour_Use_TMID = True 'AI應用時數
        If Not fg_OthHour_Use_TMID AndAlso (iALL_WNLHOUR > 0) Then fg_OthHour_Use_TMID = True '職場續航時數

        Dim i_PCont_colspan As Integer = If(fg_OthHour_Use_TMID, 4, 5) '課程進度/內容 

        Dim i_Other_rowspan As Integer = If(iALL_EHOURS > 0, 2, 1) '其他/rowspan '符合技能檢定訓練時數'/AI應用時數/職場續航時數

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;border-right-style:none ;font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("width", "5%")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.InnerHtml = ""

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "word-wrap: break-word;border-bottom-style:none;font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("width", "90px")
        rptCell.Attributes.Add("width", "5%")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.InnerHtml = ""

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        'rptCell.Attributes.Add("width", "90px")
        rptCell.Attributes.Add("width", "9%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "日期"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        'rptCell.Attributes.Add("width", "90px")
        rptCell.Attributes.Add("width", "5%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "授課<br>時段"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2)
        'rptCell.Attributes.Add("width", "90px")
        rptCell.Attributes.Add("width", "8%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "授課時間"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("width", "4%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "時數"

        '符合技能檢定訓練時數 fg_EHour_Use_TMID
        Dim S_OTH_HOUR As String = ""
        If fg_OthHour_Use_TMID Then
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 1)
            rptCell.Attributes.Add("width", "4%")
            rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
            '技檢訓練時數
            If (iALL_EHOURS > 0) Then S_OTH_HOUR &= String.Concat(If(S_OTH_HOUR <> "", "/", ""), "技檢訓練")
            If (iALL_AIAHOUR > 0) Then S_OTH_HOUR &= String.Concat(If(S_OTH_HOUR <> "", "/", ""), "AI應用")
            If (iALL_WNLHOUR > 0) Then S_OTH_HOUR &= String.Concat(If(S_OTH_HOUR <> "", "/", ""), "職場續航")
            rptCell.InnerHtml = $"{S_OTH_HOUR}(時數)"
        End If

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", i_PCont_colspan)
        'rptCell.Attributes.Add("width", "260px")
        rptCell.Attributes.Add("width", "23%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;border-top-style:none;border-bottom-style:none;border-right-style:none ;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "課程進度/內容"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "7%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學/術科"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("width", "260px")
        'rptCell.Attributes.Add("style", "word-wrap: break-word;border-top-style:none;border-bottom-style:none;border-right-style:none ;font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "授課地點"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "4%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "遠距<br>教學" '遠距教學 FARLEARN 'Ⅴ

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "4%")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "室外<br>教學" '室外教學 OUTLEARN 'Ⅴ

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "7%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "授課教師"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "7%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "助教"

        '----- 8-4
        '(Plan_TrainDesc)
        '技檢訓練時數 '符合技能檢定訓練時數 fg_EHour_Use_TMID / i_TRAINDESCD_EHours /EHOURS
        Dim i_TRAINDESCD_EHours As Decimal = 0

        If TIMS.dtHaveDATA(dt4) Then
            For i As Integer = 0 To dt4.Rows.Count - 1
                '--- 明細(二 ) start
                Dim dr4 As DataRow = dt4.Rows(i)
                If $"{dr4("EHOURS")}" <> "" Then i_TRAINDESCD_EHours += Val(dr4("EHOURS"))

                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)

                '20090820 andy edit
                '訓練模式特色與創新性
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;font-size:13pt;" & str_fontfamily_c)
                'rptCell.Attributes.Add("width", "50px")
                rptCell.Attributes.Add("width", "5%")
                rptCell.Attributes.Add("colspan", 1)
                Dim iType4 As Integer = If((TIMS.dtHaveDATA(dt4) AndAlso dt4.Rows.Count = 1 AndAlso i = 0), 1, 2) '列印方式(1:1筆 2:多筆)

                ' If Convert.ToString(o_TMID) <> "" AndAlso Convert.ToString(o_TMID) <> TIMS.cst_EHour_Use_TMID Then
                Select Case iType4
                    Case 1
                        'rptCell.InnerHtml = "訓練模式特色與創新性"
                        rptCell.InnerHtml = ""
                    Case 2
                        If i = CInt((dt4.Rows.Count) / 2) Then rptCell.InnerHtml = "訓練<br/>模式" '訓練模式特色與創新性
                        'If i = Math.Ceiling((dt4.Rows.Count) / 2) + 1 Then,'rptCell.InnerHtml = "特色<br/>與創<br/>新性",End If,
                End Select

                '課程大綱
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "border-top-style:none;border-bottom-style:none;border-left-style@none;word-wrap: break-word;font-size:13pt;" & str_fontfamily_c)
                'rptCell.Attributes.Add("width", "90px")
                rptCell.Attributes.Add("width", "5%")
                rptCell.Attributes.Add("colspan", 1)
                'Dim iType4 As Integer = 2 '列印方式(1:1筆 2:多筆)
                iType4 = 2
                If TIMS.dtHaveDATA(dt4) AndAlso dt4.Rows.Count = 1 AndAlso i = 0 Then iType4 = 1
                Select Case iType4
                    Case 1
                        rptCell.InnerHtml = "課程大綱"
                    Case 2
                        If i = CInt((dt4.Rows.Count) / 2) Then rptCell.InnerHtml = "課&nbsp程"
                        If i = Math.Ceiling((dt4.Rows.Count) / 2) + 1 Then rptCell.InnerHtml = "大&nbsp綱"
                End Select

                '日期
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap:break-word;word-break:keep-all;font-size:10pt;" & str_fontfamily_c)   'word-break:keep-all; word-wrap: break-word;
                rptCell.Attributes.Add("colspan", 1)
                'Dim S_DR4STRAIN As String = If(Not IsDBNull(dr4("STrainDate")), $"{.ToString("yyyy/MM/dd")}", "")
                rptCell.InnerHtml = TIMS.Cdate3(dr4("STrainDate")) 'S_DR4STRAIN

                '授課時段
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap:break-word;word-break:keep-all;font-size:11pt;" & str_fontfamily_c)   'word-break:keep-all; word-wrap: break-word;
                rptCell.Attributes.Add("colspan", 1)
                rptCell.InnerHtml = TIMS.Chg_TPERIOD28_VAL($"{dr4("TPERIOD28")}")

                '授課時間'星期  20090907 andy add
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "word-wrap:break-word;word-break:keep-all;font-size:11pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                Dim flag_DR4WEEK As Boolean = If($"{dr4("TechTime")}" <> "" AndAlso $"{dr4("STrainDate")}" <> "", True, False)
                Dim s_DR4WEEK As String = If(flag_DR4WEEK, TIMS.GetDayOfWeek(CInt(CDate(dr4("STrainDate")).DayOfWeek)), "")
                rptCell.InnerHtml = s_DR4WEEK

                '授課時間
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "word-wrap:break-word;word-break:keep-all;font-size:10pt;" & str_fontfamily_c)   'word-break:keep-all; word-wrap: break-word;
                rptCell.Attributes.Add("colspan", 1)
                rptCell.InnerHtml = If(IsDBNull(dr4("TechTime")), "", Replace(Replace($"{dr4("TechTime")}", " ", ""), "&nbsp", ""))

                '時數
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                rptCell.InnerHtml = If(IsDBNull(dr4("HOURS")), "", $"{dr4("HOURS")}") 'rptCell.InnerHtml = HOURS

                '技檢訓練時數 '符合技能檢定訓練時數 fg_EHour_Use_TMID
                Dim S_OTH_HOUR_H1 As String = ""
                If fg_OthHour_Use_TMID Then
                    rptCell = New HtmlTableCell
                    rptRow.Controls.Add(rptCell)
                    rptCell.Attributes.Add("align", "center")
                    rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
                    rptCell.Attributes.Add("colspan", 1)
                    If (iALL_EHOURS > 0) Then S_OTH_HOUR_H1 &= String.Concat(If(S_OTH_HOUR_H1 <> "", "/", ""), If(IsDBNull(dr4("EHOURS")), "-", $"{dr4("EHOURS")}"))
                    If (iALL_AIAHOUR > 0) Then S_OTH_HOUR_H1 &= String.Concat(If(S_OTH_HOUR_H1 <> "", "/", ""), If(IsDBNull(dr4("AIAHOUR")), "-", $"{dr4("AIAHOUR")}"))
                    If (iALL_WNLHOUR > 0) Then S_OTH_HOUR_H1 &= String.Concat(If(S_OTH_HOUR_H1 <> "", "/", ""), If(IsDBNull(dr4("WNLHOUR")), "-", $"{dr4("WNLHOUR")}"))
                    rptCell.InnerHtml = S_OTH_HOUR_H1
                End If

                '課程進度/內容 
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "left")
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:10pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", i_PCont_colspan)
                rptCell.InnerHtml = TIMS.Trn_code2($"{dr4("PCont")}")
                'rptCell.InnerHtml = If(IsDBNull(dr4("PCont")), "", $"{dr4("PCont")}")
                'PCont = Replace(Replace($"{dr4("PCont")}", vbCrLf, "<br/>"), " ", "&nbsp;")
                'PCont = Replace($"{dr4("PCont")}", " ", "&nbsp;")  '20090803 andy edit
                'PCont = Replace($"{dr4("PCont")}", " ", "&nbsp;")  '20090820 andy test

                '學/術科
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                rptCell.InnerHtml = If(IsDBNull(dr4("Classification1")), "", $"{dr4("Classification1")}")

                '授課地點
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:11pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                'rptCell.InnerHtml = If(isdbnull(dr4("PLACENAME")), "", $"{dr4("PLACENAME")}")
                Dim PLACENAME As String = ""
                PLACENAME = $"{dr4("PLACENAME")}"
                If PLACENAME.Length > 20 Then
                    For j As Integer = 1 To Convert.ToInt16(PLACENAME.Length / 20) - 1
                        PLACENAME = PLACENAME.Insert(j * 20, "<br/>")
                    Next
                End If
                rptCell.InnerHtml = PLACENAME

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                'rptCell.Attributes.Add("colspan", 1)
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:11pt;" & str_fontfamily_c)
                'rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If($"{dr4("FARLEARN")}".Equals("Y"), "Ⅴ", "")
                '遠距教學 FARLEARN 'Ⅴ

                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:11pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If($"{dr4("OUTLEARN")}".Equals("Y"), "Ⅴ", "")
                '"室外<br>教學" '室外教學 OUTLEARN 'Ⅴ

                '任課老師 授課教師
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                'rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:11pt;" & str_fontfamily_c)
                'rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr4("TeachCName")), "", $"{dr4("TeachCName")}")

                '助教
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                'rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("colspan", 1)
                rptCell.Attributes.Add("style", "word-wrap: break-word;font-size:11pt;" & str_fontfamily_c)
                'rptCell.Attributes.Add("style", "word-wrap: break-word;word-break:keep-all;font-size:10pt;" & str_fontfamily_c)
                rptCell.InnerHtml = If(IsDBNull(dr4("TeachCName2")), "", $"{dr4("TeachCName2")}")
            Next
            '----- 明細(二 ) end
        End If

        rptTb = New HtmlTable
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 1)
        rptTb.Attributes.Add("style", cst_style_c3)
        'rptTb.Attributes.Add("background", cst_bgimg_c1)
        rptTb.Attributes.Add("bordercolor", "black")
        print_content.Controls.Add(rptTb)

        '----- 9-1
        Const cst_9_1_1_colspan As Integer = 3
        Const cst_9_1_2_colspan As Integer = 5
        Const cst_9_1_3_colspan As Integer = 9

        'If dt5.Rows.Count > 0 Then dr = dt5.Rows(0)
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("rowspan", 5)
        rptCell.Attributes.Add("colspan", 1)
        rptCell.Attributes.Add("width", "5%")
        'rptCell.Attributes.Add("width", "30px")
        rptCell.InnerHtml = "訓練績效評估"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "100px")
        rptCell.Attributes.Add("colspan", cst_9_1_1_colspan)
        rptCell.InnerHtml = Rec

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_9_1_2_colspan)
        rptCell.Attributes.Add("width", "30%")
        'rptCell.Attributes.Add("width", "180px")
        rptCell.InnerHtml = "反應評估(滿意度調查機制)："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_3_colspan)
        rptCell.Attributes.Add("width", "55%")
        'rptCell.Attributes.Add("width", "285px")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = RecDesc

        '----- 9-2
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_9_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Learn

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學習評估(考試或報告機制)："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = LearnDesc

        '----- 9-3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_9_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Act

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", cst_9_1_2_colspan)
        rptCell.InnerHtml = "行為評估(課後行動計畫調查機制)："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = ActDesc

        '----- 9-4
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_9_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Rst

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "成果評估(工作績效調查機制)："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = ResultDesc

        '----- 9-5
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_9_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = oth

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "其他機制："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_9_1_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = OtherDesc

        ' --10-1-- colspan::共18
        Const cst_10_1_1_colspan As Integer = 3
        Const cst_10_1_2_colspan As Integer = 14

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("width", "5%")  
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("rowspan", "4")
        rptCell.Attributes.Add("colspan", 1)
        rptCell.InnerHtml = "促進學習機制"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_10_1_1_colspan)
        'rptCell.Attributes.Add("width", "20%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "是否為iCAP課程 ："
        'rptCell.InnerHtml = "招訓及遴選方式  ："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_10_1_2_colspan)
        'rptCell.Attributes.Add("width", "80%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        Dim IMG_ISiCAPCOUR_Y As String = If(ISiCAPCOUR = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg) & "是"
        Dim SS_iCAPCOURDESC As String = "，課程相關說明 ：" & If(iCAPCOURDESC <> "", iCAPCOURDESC, "＿＿＿＿＿＿＿") & "<br>"
        Dim SS_iCAPNUM As String = "　　　iCAP標章證號 ：" & If(iCAPNUM <> "", iCAPNUM, "＿＿＿＿＿＿＿") & "<br>"
        Dim SS_iCAPMARKDATE As String = "　　　iCAP標章有效期限 ：" & If(iCAPMARKDATE <> "", iCAPMARKDATE, "＿＿＿＿＿＿＿") & "<br><br>"
        Dim IMG_ISiCAPCOUR_N As String = If(ISiCAPCOUR = "N", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg) & "否"
        rptCell.InnerHtml = String.Format("{0}{1}{2}{3}", IMG_ISiCAPCOUR_Y, SS_iCAPCOURDESC, SS_iCAPNUM, SS_iCAPMARKDATE, IMG_ISiCAPCOUR_N)

        '------ 10-2
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_10_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "招訓方式 ："
        'rptCell.InnerHtml = "招訓及遴選方式  ："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_10_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Recruit

        '------ 10-3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_10_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "遴選方式 ："
        'rptCell.InnerHtml = "招訓及遴選方式  ："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_10_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Selmethod

        '------ 10-4
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_10_1_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學員激勵辦法  ："

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_10_1_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = Inspire


        '----- 11-1
        Const cst_11_1_colspan As Integer = 2
        Const cst_11_2_colspan As Integer = 5
        Const cst_11_3_colspan As Integer = 2
        Const cst_11_4_colspan As Integer = 5
        Const cst_11_5_colspan As Integer = 2

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        rptCell.Attributes.Add("rowspan", "3")
        rptCell.InnerHtml = "訓練費"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_1_colspan)
        'rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "政府補助"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_2_colspan)
        rptCell.Attributes.Add("width", "25%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = TIMS.Get_CostValue(DefGovCost)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_3_colspan)
        rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每班"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_4_colspan)
        rptCell.Attributes.Add("width", "25%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = Total1
        rptCell.InnerHtml = TIMS.Get_CostValue(Total1)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_5_colspan)
        rptCell.Attributes.Add("width", "15%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每人"
        '----- 11-2
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "學員自付"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = DefStdCost
        rptCell.InnerHtml = TIMS.Get_CostValue(DefStdCost)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每班"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_4_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = Total2
        rptCell.InnerHtml = TIMS.Get_CostValue(Total2)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_5_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每人"

        '-----  11-3
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_1_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "總　計"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_2_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = TIMS.Get_CostValue(Total4)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_3_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每班"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_4_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = TIMS.Get_CostValue(Total3)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_11_5_colspan)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "元/每人"

        '-----  12

        'rptRow = New HtmlTableRow
        'rptTb.Controls.Add(rptRow)

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 2)
        ''rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("style", "font-size:14pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = "初審"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 4)
        ''rptCell.Attributes.Add("width", "40%")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = FirRst

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 2)
        ''rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("style", "font-size:14pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = "複審"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 4)
        ''rptCell.Attributes.Add("width", "40%")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = SecRst
        ''End If

        If OrgKind2 = "W" Then Call Show_ENTERSUPPLY(rptTb, rptRow, rptCell, ENTERSUPPLYSTYLE) '(報名繳費方式)..

        '----- 12
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", String.Concat("font-size:14pt;", str_fontfamily_c))
        rptCell.Attributes.Add("colspan", 2) 'rptCell.Attributes.Add("width", "10%")
        If iALL_EHOURS > 0 AndAlso i_Other_rowspan > 1 Then rptCell.Attributes.Add("rowspan", i_Other_rowspan)
        rptCell.Attributes.Add("rowspan", i_Other_rowspan.ToString())
        rptCell.InnerHtml = "其他"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 7) 'rptCell.Attributes.Add("width", "40%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = "結訓後是否輔導學員參加政府機關辦理相關證照考試或技能檢定" 'rptCell.InnerHtml = "是否輔導學員參加政府機關辦理相關證照考試或技能檢定"
        rptCell.InnerHtml = "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定"

        Dim str_TGOVEXAM_Y1 As String = If(TGOVEXAM = "Y", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
        Dim str_TGOVEXAM_N1 As String = If(TGOVEXAM = "N", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
        Dim str_TGOVEXAM_G1 As String = If(TGOVEXAM = "G", cst_images_rptpic_yes_jpg, cst_images_rptpic_no_jpg)
        Dim v_GOVAGENAME As String = If(TGOVEXAM = "Y" AndAlso GOVAGENAME <> "", GOVAGENAME, "＿＿＿＿＿＿＿")
        Dim v_TGOVEXAMNAME As String = If(TGOVEXAM = "Y" AndAlso TGOVEXAMNAME <> "", TGOVEXAMNAME, "＿＿＿＿＿＿＿")
        Dim str_Y1_TGENM As String = $"{str_TGOVEXAM_Y1}是，{v_GOVAGENAME}，{v_TGOVEXAMNAME}"
        Dim str_N1_TGENM As String = $"{str_TGOVEXAM_N1}否。(包含非政府機關辦理相關證照或檢定)"
        Dim str_G1_TGENM As String = $"{str_TGOVEXAM_G1}本課程結訓後須參加環境部辦理之淨零綠領人才培育課程測驗；測驗成績達及格，即可申請本方案補助。"
        Dim rTGOVEXAM As String = $"{str_Y1_TGENM}<br>{str_N1_TGENM}<br>{str_G1_TGENM}"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell) 'rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", 9) 'rptCell.Attributes.Add("width", "10%")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = rTGOVEXAM

        If iALL_EHOURS > 0 Then
            '符合技能檢定訓練時數
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 7) 'rptCell.Attributes.Add("width", "40%")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "符合技能檢定訓練時數"

            '技檢訓練時數 '符合技能檢定訓練時數 fg_EHour_Use_TMID / i_TRAINDESCD_EHours /EHOURS
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell) 'rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", 9) 'rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = $"{i_TRAINDESCD_EHours} 小時"
        End If

        '-----  12-2
        'PLAN_ABILITY - 專長能力標籤
        If TIMS.dtHaveDATA(dt6) Then
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 4) 'rptCell.Attributes.Add("width", "40%")
            rptCell.Attributes.Add("rowspan", dt6.Rows.Count)
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "專長能力標籤"

            For i6 As Integer = 0 To dt6.Rows.Count - 1
                Dim dr6 As DataRow = dt6.Rows(i6)
                If i6 > 0 Then
                    rptRow = New HtmlTableRow
                    rptTb.Controls.Add(rptRow)
                End If
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("colspan", cst_iColspanAll_1 - 4)
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.InnerHtml = $"{dr6("ABILITYDESC")}"
            Next
        Else
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("colspan", 4) 'rptCell.Attributes.Add("width", "40%")
            'rptCell.Attributes.Add("rowspan", dt6.Rows.Count)
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "專長能力標籤"

            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("colspan", cst_iColspanAll_1 - 4)
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.InnerHtml = "&nbsp;"
        End If

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", 2) 'rptCell.Attributes.Add("width", "40%")
        'rptCell.Attributes.Add("rowspan", dt6.Rows.Count)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "備註"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1 - 2)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = $"{MEMO8}<br>{MEMO82}"

        '-----  13
        'Const cst_13_1_colspan As Integer = cst_iColspanAll_1 '18
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.InnerHtml = "&nbsp;"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 2)
        ''rptCell.Attributes.Add("width", "10%")
        'rptCell.Attributes.Add("style", "font-size:14pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = "審查結果"

        'rptCell = New HtmlTableCell
        'rptRow.Controls.Add(rptCell)
        'rptCell.Attributes.Add("align", "center")
        'rptCell.Attributes.Add("colspan", 10)
        ''rptCell.Attributes.Add("width", "40%")
        'rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        'rptCell.InnerHtml = FirRst

        '-----  表尾  start 
        rptTb = New HtmlTable
        rptTb.Attributes.Add("style", cst_style_c1)
        'rptTb.Attributes.Add("background", cst_bgimg_c1)
        rptTb.Attributes.Add("align", "center")
        rptTb.Attributes.Add("border", 0)
        print_content.Controls.Add(rptTb)

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "<br>"

        '確認所選之訓練業別正確性
        If flag_OJT22071401 Then
            'Const cst_Y1_TMIDCORRECT As String = "<img src='../../images/rptpic/no.jpg' />同意"
            'Const cst_Y2_TMIDCORRECT As String = "<img src='../../images/rptpic/yes.jpg' />同意"
            'Const cst_N1_TMIDCORRECT As String = "<img src='../../images/rptpic/no.jpg' />不同意"
            'Const cst_N2_TMIDCORRECT As String = "<img src='../../images/rptpic/yes.jpg' />不同意"
            Dim rTMIDCORRECT As String = String.Concat(cst_images_rptpic_no_jpg, "同意", "　　", cst_images_rptpic_no_jpg, "不同意")
            Select Case TMIDCORRECT
                Case "Y"
                    rTMIDCORRECT = String.Concat(cst_images_rptpic_yes_jpg, "同意", "　　", cst_images_rptpic_no_jpg, "不同意")
                Case "N"
                    rTMIDCORRECT = String.Concat(cst_images_rptpic_no_jpg, "同意", "　　", cst_images_rptpic_yes_jpg, "不同意")
            End Select

            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "left")
            rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
            'rptCell.Attributes.Add("width", "10%")
            rptCell.Attributes.Add("style", "font-size:14pt;" & str_fontfamily_c)
            rptCell.InnerHtml = $"{cst_TMIDCORRECT_c}<br>{rTMIDCORRECT}"
        End If

        Dim s_bottom_msg1 As String = "※無術科或學科場地者免填該項場地資訊、容納人數與硬體設施說明。<br/>"
        Dim s_bottom_msg2 As String = "※本表如有塗改，應予重新印製或於塗改處加蓋訓練單位主管職銜章。<br/>"
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.InnerHtml = $"{s_bottom_msg1}{s_bottom_msg2}"

        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)
        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("height", "100")
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("colspan", cst_iColspanAll_1)
        rptCell.Attributes.Add("style", "font-size:22pt;" & str_fontfamily_c)
        rptCell.InnerHtml = "<br />"
        '-----  表尾  end  

        'Try
        'Catch ex As Exception
        '    'Common.MessageBox(Me, ex.ToString)
        '    Dim strErrmsg As String = ""
        '    strErrmsg += "/*  ex.ToString: */" & vbCrLf
        '    strErrmsg += ex.ToString & vbCrLf
        '    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.SendMailTest(strErrmsg)
        '    Dim strScript As String = ""
        '    strScript = "<script language=""javascript"">" + vbCrLf
        '    strScript += "alert('發生錯誤!! " & Common.GetJsString(ex.Message.ToString()) & "');" + vbCrLf
        '    strScript += "</script>"
        '    Page.RegisterStartupScript("", strScript)
        '    'Finally 'conn.Close() da.Dispose()
        'End Try

    End Sub

    Private Function GET_TRAINDESCD_HOURS(dt4 As DataTable, COLUMN_N As String) As Double
        Dim rst As Double = 0
        If TIMS.dtNODATA(dt4) Then Return 0
        For Each dr4 As DataRow In dt4.Rows
            If TIMS.IsNumeric1(dr4(COLUMN_N)) AndAlso Val(dr4(COLUMN_N)) > 0 Then rst += Val(dr4(COLUMN_N))
        Next
        Return rst
    End Function

    Sub Export_XLS1()
        Dim fileName As String = "訓練班別計畫表.xls"
        'If Request.Browser.Browser = "IE" Then fileName = Server.UrlPathEncode(fileName)
        'Dim strContentDisposition As String = [String].Format("{0}; filename=""{1}""", "attachment", fileName)
        'Response.AddHeader("Content-Disposition", strContentDisposition)
        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8))
        Response.ContentType = "application/vnd.ms-excel"
        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)

        'htw.Write("<html><head>")
        'htw.Write(" <style> ")
        'htw.Write(" table { width:300px; } ")
        'htw.Write(" </style> ")
        'htw.Write("</head><body><form>")
        'Common.RespWrite(Me, sw.ToString().Replace("<div>", "").Replace("</div>", ""))
        ''htw.Write("</form></body></html>")
        'Response.End()

        print_content.RenderControl(htw)
        Dim strScriptHtml As String = sw.ToString() '.Replace("<div>", "").Replace("</div>", "")
        Call TIMS.Utl_RespWriteEnd(Me, objconn, strScriptHtml)
        'Exit Sub
    End Sub

    ''' <summary>測試PDF輸出 寫LOG</summary>
    Sub Export_PDF1()
        Dim YMDSTR1x As String = DateTime.Now.ToString("ssHHddMMyyyymmss")
        Dim strFileName As String = $"{YMDSTR1x}.pdf"
        Dim s_Charset As String = TIMS.cst_Charset_UTF8 '"UTF-8" 'default

        'Const cst_div_prtConter1 As String = "<div id=""print_content"">"
        'Dim imgPath1 As String = Server.MapPath("~/images/rptpic/temple/TIMS_1.jpg")
        'Dim str_div_prtConter2 As String = String.Concat("<div id=""print_content"" style=""background-image: url(&#39;", imgPath1, "&#39;); background-repeat:repeat; background-position: center center;"">")

        Response.Clear()
        'MyPage.Response.ClearHeaders()
        'MyPage.Response.Charset = "UTF-8"
        'MyPage.Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.Charset = s_Charset
        Response.ContentEncoding = System.Text.Encoding.GetEncoding(s_Charset)
        Response.ContentType = "application/pdf" 'PDF
        Response.AppendHeader("Content-Disposition", $"attachment; filename={strFileName}")

        Dim sw As New System.IO.StringWriter
        Dim htw As New HtmlTextWriter(sw)
        print_content.RenderControl(htw)
        'form1.RenderControl(htw)
        Dim strScriptHtml As String = sw.ToString() '.Replace("<div>", "").Replace("</div>", "")
        'strScriptHtml = strScriptHtml.Replace(cst_div_prtConter1, str_div_prtConter2)
        'TIMS.LOG.Debug(String.Concat("##Export_PDF1,strScriptHtml : ", vbCrLf, strScriptHtml))

        Using stream As New System.IO.MemoryStream
            Dim pdf As HiQPdf.HtmlToPdf = New HiQPdf.HtmlToPdf With {
                .SerialNumber = HiQPdf_SerialNumber'"/7eWrq+b-mbOWnY2e-jYbOz9HP-387fy9/G-yM7fzM7R-zs3RxsbG-xg=="   '「HiQPdf」的SerialNumber
                }
            pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Portrait   '紙張直向
            'pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Landscape   '紙張橫向
            'pdf.Document.DestWidth = 596 '595

            'pdf.ConvertHtmlToStream(strScriptHtml, Nothing, stream)
            pdf.Document.DisplayMaskedImages = True
            Dim pdfBaseUrl As String = ReportQuery.GetBaseUrl(Me) '"https://localhost:44383/"
            TIMS.LOG.Debug($"##SD_14_002_R, ##Export_PDF1,pdfBaseUrl: {pdfBaseUrl}")
            'TIMS.LOG.Debug(String.Concat("##GetHTTP_HOST:", vbCrLf, TIMS.GetHTTP_HOST(Me), vbCrLf))
            pdf.ConvertHtmlToStream(strScriptHtml, pdfBaseUrl, stream)

            '輸出PDF檔案。
            Response.BinaryWrite(stream.ToArray())
        End Using
    End Sub

    ''' <summary>單一輸出 不執行關閉</summary>
    ''' <param name="rPMS"></param>
    Sub Export_PDF2(rPMS As Hashtable)
        If rPMS Is Nothing Then Return
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vPCS As String = TIMS.GetMyValue2(rPMS, "PCS")

        Dim YMDSTR1x As String = DateTime.Now.ToString("ssHHddMMyyyymmss")
        Dim strFileName As String = $"{YMDSTR1x}.pdf"
        Dim s_Charset As String = TIMS.cst_Charset_UTF8 '"UTF-8" 'default

        Response.Clear()
        'MyPage.Response.ClearHeaders()
        'MyPage.Response.Charset = "UTF-8"
        'MyPage.Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.Charset = s_Charset
        Response.ContentEncoding = System.Text.Encoding.GetEncoding(s_Charset)
        Response.ContentType = "application/pdf" 'PDF
        Response.AppendHeader("Content-Disposition", $"attachment; filename={strFileName}")

        Dim sw As New System.IO.StringWriter
        Dim htw As New HtmlTextWriter(sw)
        print_content.RenderControl(htw)
        Dim strScriptHtml As String = sw.ToString() '.Replace("<div>", "").Replace("</div>", "")

        Const cst_08_訓練班別計畫表_WAIVED_PI As String = "PI"
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
        If drOB Is Nothing Then Return
        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"

        Dim PlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim SeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqPCS As String = $"{PlanID}x{ComIDNO}x{SeqNo}"
        Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), PlanID, ComIDNO, SeqNo)
        If iBCPID <= 0 Then Return
        Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, vBCID, vKBSID, cst_08_訓練班別計畫表_WAIVED_PI, drOB)
        If iBCFID <= 0 Then Return

        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PI(vBCID, vKBSID, vPCS, "pdf")
        '檔案儲存
        TIMS.LOG.Debug(String.Concat("##Export_PDF1 , strFileName: ", strFileName))
        Dim vSRCFILENAME1 As String = strFileName ' Convert.ToString(oSRCFILENAME1)

        Using Mstream As New System.IO.MemoryStream
            Dim pdf As HiQPdf.HtmlToPdf = New HiQPdf.HtmlToPdf With {
                .SerialNumber = HiQPdf_SerialNumber'"/7eWrq+b-mbOWnY2e-jYbOz9HP-387fy9/G-yM7fzM7R-zs3RxsbG-xg=="   '「HiQPdf」的SerialNumber
                }
            pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Portrait   '紙張直向
            'pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Landscape   '紙張橫向
            'pdf.ConvertHtmlToStream(strScriptHtml, Nothing, Mstream)
            pdf.Document.DisplayMaskedImages = True
            Dim pdfBaseUrl As String = ReportQuery.GetBaseUrl(Me) '"https://localhost:44383/"
            TIMS.LOG.Debug($"##SD_14_002_R, ##Export_PDF2,pdfBaseUrl: {pdfBaseUrl}")
            'TIMS.LOG.Debug(String.Concat("##GetHTTP_HOST:", vbCrLf, TIMS.GetHTTP_HOST(Me), vbCrLf))
            pdf.ConvertHtmlToStream(strScriptHtml, pdfBaseUrl, Mstream)

            '上傳檔案/存檔：檔名
            Dim save_file_Stream1 As String = Server.MapPath(Path.Combine(vUploadPath, vFILENAME1))
            Try
                TIMS.MyCreateDir(Me, vUploadPath)
                File.WriteAllBytes(save_file_Stream1, Mstream.ToArray())
                'Dim file_Stream1 As New FileStream(save_file_Stream1, FileMode.Append, FileAccess.Write)
                'Mstream..CopyTo(file_Stream1) 'stream.WriteTo(file_Stream1)
                'pdf.ConvertHtmlToStream(strScriptHtml, Nothing, file_Stream1)
                'IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
                '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
                'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
                'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Common.MessageBox(Me, "處理檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)")

                Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
                'strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
                strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
                strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
                'strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
                strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
                'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                TIMS.WriteTraceLog(Me, ex, strErrmsg) 'Exit Sub
                Return
            End Try

            Dim rPMSPI As New Hashtable From {
                {"UploadPath", vUploadPath},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"BCPID", iBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Call TIMS.SAVE_ORG_BIDCASEFL_PI(Me, objconn, rPMSPI)
            '輸出PDF檔案。
            Response.BinaryWrite(Mstream.ToArray())
        End Using
    End Sub

    ''' <summary>批次輸出並且會執行關閉</summary>
    ''' <param name="rPMS"></param>
    Sub Export_PDF3(rPMS As Hashtable)
        If rPMS Is Nothing Then Return
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vPCS As String = TIMS.GetMyValue2(rPMS, "PCS")

        Dim YMDSTR1x As String = DateTime.Now.ToString("ssHHddMMyyyymmss")
        Dim strFileName As String = String.Concat(YMDSTR1x, ".pdf")
        Dim s_Charset As String = TIMS.cst_Charset_UTF8 '"UTF-8" 'default

        Response.Clear()
        'Response.Charset = s_Charset
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding(s_Charset)
        'Response.ContentType = "application/pdf" 'PDF
        'Response.AppendHeader("Content-Disposition", String.Concat("attachment; filename=", strFileName))

        Dim sw As New System.IO.StringWriter
        Dim htw As New HtmlTextWriter(sw)
        print_content.RenderControl(htw)
        'form1.RenderControl(htw)
        'Dim strScriptHtml As String = sw.ToString()
        'print_content.RenderControl(htw)
        'Dim strScriptHtml As String = sw.ToString().Replace("<div>", "").Replace("</div>", "")
        'Dim sHtmlS1 As String = ""
        'sHtmlS1 &= " <style type=""text/css"">" & vbCrLf
        'sHtmlS1 &= " div { background-image: url('../../images/rptpic/temple/TIMS_1.jpg'); background-repeat:repeat; background-position: center center; }" & vbCrLf
        'sHtmlS1 &= " </style>" & vbCrLf
        'Dim strScriptHtml As String = String.Concat(sHtmlS1, sw.ToString())
        Dim strScriptHtml As String = sw.ToString()

        Const cst_08_訓練班別計畫表_WAIVED_PI As String = "PI"

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
        If drOB Is Nothing Then Return
        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"

        Dim PlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim SeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqPCS As String = $"{PlanID}x{ComIDNO}x{SeqNo}"
        Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), PlanID, ComIDNO, SeqNo)
        If iBCPID <= 0 Then Return
        Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, vBCID, vKBSID, cst_08_訓練班別計畫表_WAIVED_PI, drOB)
        If iBCFID <= 0 Then Return

        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PI(vBCID, vKBSID, vPCS, "pdf")
        '檔案儲存
        TIMS.LOG.Debug($"##Export_PDF1 , strFileName: {strFileName}")
        Dim vSRCFILENAME1 As String = strFileName ' Convert.ToString(oSRCFILENAME1)

        'pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Landscape   '紙張橫向
        Using Mstream As New System.IO.MemoryStream
            Dim pdf As HiQPdf.HtmlToPdf = New HiQPdf.HtmlToPdf With {
                .SerialNumber = HiQPdf_SerialNumber'"/7eWrq+b-mbOWnY2e-jYbOz9HP-387fy9/G-yM7fzM7R-zs3RxsbG-xg=="   '「HiQPdf」的SerialNumber
                }
            pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Portrait   '紙張直向
            'pdf.ConvertHtmlToStream(strScriptHtml, Nothing, Mstream)
            pdf.Document.DisplayMaskedImages = True
            Dim pdfBaseUrl As String = ReportQuery.GetBaseUrl(Me) '"https://localhost:44383/"
            TIMS.LOG.Debug($"##SD_14_002_R, ##Export_PDF3,pdfBaseUrl: {pdfBaseUrl}")
            'TIMS.LOG.Debug(String.Concat("##GetHTTP_HOST:", vbCrLf, TIMS.GetHTTP_HOST(Me), vbCrLf))
            pdf.ConvertHtmlToStream(strScriptHtml, pdfBaseUrl, Mstream)

            '上傳檔案/存檔：檔名
            Dim save_file_Stream1 As String = Server.MapPath(Path.Combine(vUploadPath, vFILENAME1))
            Try
                TIMS.MyCreateDir(Me, vUploadPath)
                File.WriteAllBytes(save_file_Stream1, Mstream.ToArray())
                'Dim file_Stream1 As New FileStream(save_file_Stream1, FileMode.Append, FileAccess.Write)
                'Mstream..CopyTo(file_Stream1) 'stream.WriteTo(file_Stream1)
                'pdf.ConvertHtmlToStream(strScriptHtml, Nothing, file_Stream1)
                'IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
                '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
                'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
                'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
            Catch ex As Exception
                TIMS.LOG.Error(ex.Message, ex)
                Common.MessageBox(Me, "處理檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)")

                Dim strErrmsg As String = $"{ex.Message}{vbCrLf}ex.ToString:{ex.ToString}{vbCrLf}"
                strErrmsg &= $"vUploadPath: {vUploadPath}{vbCrLf}"
                strErrmsg &= $"vFILENAME1: {vFILENAME1}{vbCrLf}"
                strErrmsg &= $"vSRCFILENAME1(MyFileName): {vSRCFILENAME1}{vbCrLf}"
                strErrmsg &= $"Server.MapPath(vUploadPath, vFILENAME1): {Server.MapPath($"{vUploadPath}{vFILENAME1}")}{vbCrLf}"
                TIMS.WriteTraceLog(Me, ex, strErrmsg) 'Exit Sub
                Return
            End Try
        End Using

        Dim rPMSPI As New Hashtable From {
                {"UploadPath", vUploadPath},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"BCPID", iBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
        Call TIMS.SAVE_ORG_BIDCASEFL_PI(Me, objconn, rPMSPI)

        '輸出PDF檔案。 'Response.BinaryWrite(Mstream.ToArray())
        Call ReportQuery.CloseWin2(Me)
    End Sub

    ''' <summary>'覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤</summary>
    ''' <param name="Control"></param>
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    ''' <summary>匯出Excel BTN</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub bt_excel_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles bt_excel.Click
        Export_XLS1()
    End Sub

    ''' <summary>顯示(企業包班)</summary>
    ''' <param name="rptTb"></param>
    ''' <param name="rptRow"></param>
    ''' <param name="rptCell"></param>
    ''' <param name="dt2"></param>
    ''' <remarks></remarks>
    Sub Show_BUSPACKAGE(ByRef rptTb As HtmlTable, ByRef rptRow As HtmlTableRow, ByRef rptCell As HtmlTableCell, ByRef dt2 As DataTable)
        Dim intRowSpan As Integer = 2 '預設2行
        If TIMS.dtHaveDATA(dt2) Then intRowSpan = 1 + dt2.Rows.Count '有資料加一行title

        '-----  6-2-1
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("rowspan", intRowSpan)
        rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("colspan", 2)
        rptCell.InnerHtml = "事業單位資料"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "40%")
        'rptCell.Attributes.Add("width", "70px")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "企業名稱"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "30%")
        'rptCell.Attributes.Add("width", "50px")
        rptCell.Attributes.Add("colspan", 3)
        rptCell.InnerHtml = "服務單位統一編號"

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("width", "30%")
        'rptCell.Attributes.Add("width", "250px")
        rptCell.Attributes.Add("colspan", 9)
        rptCell.InnerHtml = "保險證號"
        'End If
        '----- 表首 end

        '----- 6-2-2
        Dim tmpInnerHtml As String = ""
        If TIMS.dtHaveDATA(dt2) Then
            '----- 明細(一) start
            For i As Integer = 0 To dt2.Rows.Count - 1
                Dim dr2 As DataRow = dt2.Rows(i)
                rptRow = New HtmlTableRow
                rptTb.Controls.Add(rptRow)

                '企業名稱
                tmpInnerHtml = If(IsDBNull(dr2("Uname")), "&nbsp;", $"{dr2("Uname")}")
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("width", "40%")
                'rptCell.Attributes.Add("width", "70px")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.InnerHtml = tmpInnerHtml '"企業名稱"

                '服務單位統一編號
                tmpInnerHtml = If(IsDBNull(dr2("Intaxno")), "&nbsp;", $"{dr2("Intaxno")}")
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("width", "30%")
                'rptCell.Attributes.Add("width", "50px")
                rptCell.Attributes.Add("colspan", 3)
                rptCell.InnerHtml = tmpInnerHtml '"服務單位統一編號"

                '保險證號
                tmpInnerHtml = If(IsDBNull(dr2("Ubno")), "&nbsp;", $"{dr2("Ubno")}")
                rptCell = New HtmlTableCell
                rptRow.Controls.Add(rptCell)
                rptCell.Attributes.Add("align", "center")
                rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
                rptCell.Attributes.Add("width", "30%")
                'rptCell.Attributes.Add("width", "250px")
                rptCell.Attributes.Add("colspan", 9)
                rptCell.InnerHtml = tmpInnerHtml '"保險證號"
            Next
            '----- 明細(一) end
        Else
            rptRow = New HtmlTableRow
            rptTb.Controls.Add(rptRow)

            '企業名稱
            tmpInnerHtml = "&nbsp;" ' If(IsDBNull(dr2("Uname")), "&nbsp;", $"{dr2("Uname")}")
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("width", "40%")
            'rptCell.Attributes.Add("width", "70px")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.InnerHtml = tmpInnerHtml '"企業名稱"

            '服務單位統一編號
            tmpInnerHtml = "&nbsp;" 'If(IsDBNull(dr2("Intaxno")), "&nbsp;", $"{dr2("Intaxno")}")
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("width", "30%")
            'rptCell.Attributes.Add("width", "50px")
            rptCell.Attributes.Add("colspan", 3)
            rptCell.InnerHtml = tmpInnerHtml '"服務單位統一編號"

            '保險證號
            tmpInnerHtml = "&nbsp;" 'If(IsDBNull(dr2("Ubno")), "&nbsp;", $"{dr2("Ubno")}")
            rptCell = New HtmlTableCell
            rptRow.Controls.Add(rptCell)
            rptCell.Attributes.Add("align", "center")
            rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
            rptCell.Attributes.Add("width", "30%")
            'rptCell.Attributes.Add("width", "250px")
            rptCell.Attributes.Add("colspan", 9)
            rptCell.InnerHtml = tmpInnerHtml '"保險證號"
        End If
    End Sub

    ''' <summary>顯示 報名繳費方式</summary>
    ''' <param name="rptTb"></param>
    ''' <param name="rptRow"></param>
    ''' <param name="rptCell"></param>
    ''' <param name="ENTERSUPPLYSTYLE"></param>
    ''' <remarks></remarks>
    Sub Show_ENTERSUPPLY(ByRef rptTb As HtmlTable, ByRef rptRow As HtmlTableRow, ByRef rptCell As HtmlTableCell, ByRef ENTERSUPPLYSTYLE As String)
        rptRow = New HtmlTableRow
        rptTb.Controls.Add(rptRow)

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "center")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 2)
        rptCell.InnerHtml = "報名繳費方式"

        'rptCell.Attributes.Add("align", "left")
        'rptCell.Attributes.Add("rowspan", "1")
        'rptCell.Attributes.Add("width", "15%")
        'rptCell.Attributes.Add("width", "70px")

        Dim tENTERSUPPLY As String = String.Concat(cst_images_rptpic_no_jpg, cst_ENTERSUPPLY_1, "<br>", cst_images_rptpic_no_jpg, cst_ENTERSUPPLY_2)
        Select Case ENTERSUPPLYSTYLE
            Case "1"
                tENTERSUPPLY = String.Concat(cst_images_rptpic_yes_jpg, cst_ENTERSUPPLY_1, "<br>", cst_images_rptpic_no_jpg, cst_ENTERSUPPLY_2)
            Case "2"
                tENTERSUPPLY = String.Concat(cst_images_rptpic_no_jpg, cst_ENTERSUPPLY_1, "<br>", cst_images_rptpic_yes_jpg, cst_ENTERSUPPLY_2)
        End Select

        rptCell = New HtmlTableCell
        rptRow.Controls.Add(rptCell)
        rptCell.Attributes.Add("align", "left")
        rptCell.Attributes.Add("style", "font-size:12pt;" & str_fontfamily_c)
        rptCell.Attributes.Add("colspan", 16)
        rptCell.InnerHtml = tENTERSUPPLY  '報名繳費方式

        'rptCell.Attributes.Add("width", "85%")
        'rptCell.Attributes.Add("style", "font-size:10pt;" & str_fontfamily_c)
        'rptCell.Attributes.Add("rowspan", "1")
        'rptCell.Attributes.Add("colspan", 9)
    End Sub

    Protected Sub imgBt_Pdf_Click(sender As Object, e As ImageClickEventArgs) Handles imgBt_Pdf.Click
        Call Export_PDF1()
    End Sub
End Class

