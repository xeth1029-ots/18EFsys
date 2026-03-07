Public Class SD_15_007_R_Prt2
    Inherits AuthBasePage

    Const cst_rpt性別 As String = "SD_15_007_R11"
    Const cst_rpt年齡 As String = "SD_15_007_R12"
    Const cst_rpt教育程度 As String = "SD_15_007_R13"
    Const cst_rpt身分別 As String = "SD_15_007_R14"
    Const cst_rpt工作年資 As String = "SD_15_007_R15"
    Const cst_rpt地理分佈 As String = "SD_15_007_R16"
    Const cst_rpt公司行業別 As String = "SD_15_007_R17"
    Const cst_rpt公司規模 As String = "SD_15_007_R18"
    Const cst_rpt參訓動機 As String = "SD_15_007_R19"
    Const cst_rpt訓後動向 As String = "SD_15_007_R20"
    Const cst_rpt參訓單位類別 As String = "SD_15_007_R21"
    Const cst_rpt參加課程職能別 As String = "SD_15_007_R22"
    Const cst_rpt參加課程型態別 As String = "SD_15_007_R23"

    Dim arrItem(164, 3) As String
    'Dim arrSpanRow(114) As String
    Dim sourceDT As DataTable = Nothing
    'Dim PName As String = ""
    'Dim plankind As String = ""
    Dim filename As String = ""

    'STUD_QUESTIONFAC / STUD_QUESTIONFAC2 /dbo.fn_GET_GOVCNT 'SD_11_004_add17
    'Dim gResDt As DataTable = Nothing '計算行數共用表格
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load 'Me.Load
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        btnCancel.Attributes.Add("onclick", "window.close();")

        Dim TPlanID As String = Request("TPlanID")
        Dim Years As String = Request("Years")
        Dim OCID As String = Request("OCID")
        Dim RID As String = Request("RID")
        Dim SearchPlan As String = Request("SearchPlan") 'G/W
        Dim PackageType As String = Request("PackageType")
        filename = Request("filename")

        TPlanID = TIMS.ClearSQM(TPlanID) 'Request("TPlanID")
        Years = TIMS.ClearSQM(Years) 'Request("Years")
        OCID = TIMS.ClearSQM(OCID) 'Request("OCID")
        RID = TIMS.ClearSQM(RID) 'Request("RID")
        SearchPlan = TIMS.ClearSQM(SearchPlan) 'Request("SearchPlan")
        PackageType = TIMS.ClearSQM(PackageType) ' Request("PackageType")
        filename = TIMS.ClearSQM(filename)
        If RID = "A" Then RID = "" '署(局)屬

        Dim ss As String = ""
        TIMS.SetMyValue(ss, "TPlanID", TPlanID)
        TIMS.SetMyValue(ss, "Years", Years)
        TIMS.SetMyValue(ss, "OCID", OCID)
        TIMS.SetMyValue(ss, "RID", RID)
        TIMS.SetMyValue(ss, "SearchPlan", SearchPlan)
        TIMS.SetMyValue(ss, "PackageType", PackageType)
        'TIMS.SetMyValue(ss, "filename", filename)

        Call SUB_Rpt2A(ss)
    End Sub

    '依filename組這次的群組 ，並且縮小搜尋範圍 STUD_QUESTIONFAC2
    Function Get_WC1_SQL(ByRef ss As String) As String
        'Dim filename As String = TIMS.GetMyValue(ss, "filename")
        Dim RID As String = TIMS.GetMyValue(ss, "RID")
        Dim OCID As String = TIMS.GetMyValue(ss, "OCID")
        Dim Years As String = TIMS.GetMyValue(ss, "Years")
        Dim PackageType As String = TIMS.GetMyValue(ss, "PackageType")
        Dim TPlanID As String = TIMS.GetMyValue(ss, "TPlanID")
        Dim SearchPlan As String = TIMS.GetMyValue(ss, "SearchPlan")

        Dim sql As String = ""
        sql &= " SELECT fc2.SOCID" & vbCrLf '/*PK*/
        'sql &= " ,fc2.S11" & vbCrLf'sql &= " ,fc2.S12" & vbCrLf'sql &= " ,fc2.S13" & vbCrLf'sql &= " ,fc2.S14" & vbCrLf'sql &= " ,fc2.S15" & vbCrLf'sql &= " ,fc2.S16" & vbCrLf'sql &= " ,fc2.S16_NOTE" & vbCrLf'sql &= " ,fc2.S2" & vbCrLf'sql &= " ,fc2.S3" & vbCrLf
        sql &= " ,fc2.A1_1" & vbCrLf
        sql &= " ,fc2.A1_2" & vbCrLf
        sql &= " ,fc2.A1_3" & vbCrLf
        sql &= " ,fc2.A1_4" & vbCrLf
        sql &= " ,fc2.A1_5" & vbCrLf
        sql &= " ,fc2.A1_6" & vbCrLf
        sql &= " ,fc2.A1_7" & vbCrLf
        sql &= " ,fc2.A1_8" & vbCrLf
        sql &= " ,fc2.A1_9" & vbCrLf
        sql &= " ,fc2.A1_10" & vbCrLf
        'sql &= " ,fc2.A1_10_NOTE" & vbCrLf
        sql &= " ,fc2.A2" & vbCrLf
        sql &= " ,fc2.A3" & vbCrLf
        sql &= " ,fc2.A4" & vbCrLf
        sql &= " ,fc2.A5" & vbCrLf
        sql &= " ,fc2.A6" & vbCrLf
        sql &= " ,fc2.A7" & vbCrLf

        sql &= " ,fc2.B11" & vbCrLf
        sql &= " ,fc2.B12" & vbCrLf
        sql &= " ,fc2.B13" & vbCrLf
        sql &= " ,fc2.B14" & vbCrLf
        sql &= " ,fc2.B15" & vbCrLf

        sql &= " ,fc2.B21" & vbCrLf
        sql &= " ,fc2.B22" & vbCrLf
        sql &= " ,fc2.B23" & vbCrLf
        sql &= " ,fc2.B31" & vbCrLf
        sql &= " ,fc2.B32" & vbCrLf
        sql &= " ,fc2.B41" & vbCrLf
        sql &= " ,fc2.B42" & vbCrLf
        sql &= " ,fc2.B43" & vbCrLf
        sql &= " ,fc2.B44" & vbCrLf
        sql &= " ,fc2.B51" & vbCrLf
        sql &= " ,fc2.B61" & vbCrLf
        sql &= " ,fc2.B62" & vbCrLf
        sql &= " ,fc2.B63" & vbCrLf
        sql &= " ,fc2.B71" & vbCrLf
        sql &= " ,fc2.B72" & vbCrLf
        sql &= " ,fc2.B73" & vbCrLf
        sql &= " ,fc2.B74" & vbCrLf
        sql &= " ,fc2.C11" & vbCrLf
        'sql &= " ,fc2.C21_NOTE" & vbCrLf'sql &= " ,fc2.MODIFYACCT" & vbCrLf'sql &= " ,fc2.MODIFYDATE" & vbCrLf'sql &= " ,fc2.DASOURCE" & vbCrLf'sql &= " ,fc2.A2_7_NOTE" & vbCrLf'sql &= " ,fc2.A3_5_NOTE" & vbCrLf
        'sql &= " ,ss.SEX mid11" & vbCrLf
        ''sql &= " ,trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) mid12" & vbCrLf
        'sql &= " ,case when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <15 then '01'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=15 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <20 then '02'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=20 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <25 then '03'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=25 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <30 then '04'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=30 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <35 then '05'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=35 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <40 then '06'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=40 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <45 then '07'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=45 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <50 then '08'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=50 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <55 then '09'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=55 AND trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <60 then '10'" & vbCrLf
        'sql &= " when trunc(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=60 then '11' end mid12" & vbCrLf

        'sql &= " ,ss.DegreeID mid13" & vbCrLf
        'sql &= " ,cs.Midentityid mid14" & vbCrLf
        ''sql &= " --,sb.q61" & vbCrLf
        'sql &= " ,case when sb.q61<3 then '01'" & vbCrLf
        'sql &= " when sb.q61<5 then '02'" & vbCrLf
        'sql &= " when sb.q61<10 then '03'" & vbCrLf
        'sql &= " else '04' end mid15" & vbCrLf
        'sql &= " ,iz.CTID mid16" & vbCrLf
        'sql &= " ,sb.Q4 mid17" & vbCrLf
        'sql &= " ,sb.Q5 mid18" & vbCrLf
        'sql &= " ,sb2.Q2A mid19" & vbCrLf
        'sql &= " ,sb.Q3 mid20" & vbCrLf
        'sql &= " ,oo.OrgKind mid21" & vbCrLf
        'sql &= " ,pp.ClassCate mid22" & vbCrLf

        'sql &= " ,case when pp.PointYN = 'Y' then 1" & vbCrLf
        'sql &= " when pp.PointYN = 'N' then 2" & vbCrLf
        'sql &= " when pp.IsBusiness = 'Y' then 3 end as Mid23" & vbCrLf
        Select Case filename
            Case cst_rpt性別 ' "SD_15_007_R11"
                sql &= " ,ss.SEX mid" & vbCrLf
            Case cst_rpt年齡 '"SD_15_007_R12"
                'trunc --> floor （無條件捨去）
                'dbo.FN_YEARSOLD(cc.ftdate, ss.birthday)
                sql &= " ,CASE WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <15 then '01'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <20 then '02'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <25 then '03'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <30 then '04'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <35 then '05'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <40 then '06'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <45 then '07'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <50 then '08'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <55 then '09'" & vbCrLf
                sql &= "  WHEN dbo.FN_YEARSOLD(cc.STDATE, ss.BIRTHDAY) <60 then '10'" & vbCrLf
                sql &= "  ELSE '11' END mid" & vbCrLf
                'sql &= " ,case when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <15 then '01'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=15 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <20 then '02'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=20 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <25 then '03'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=25 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <30 then '04'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=30 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <35 then '05'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=35 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <40 then '06'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=40 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <45 then '07'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=45 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <50 then '08'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=50 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <55 then '09'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=55 AND floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) <60 then '10'" & vbCrLf
                'sql &= " when floor(dbo.MONTHS_BETWEEN(cc.STDate,ss.birthday)/12) >=60 then '11' end mid" & vbCrLf

            Case cst_rpt教育程度 '"SD_15_007_R13"
                sql &= " ,ss.DegreeID mid" & vbCrLf
            Case cst_rpt身分別 '"SD_15_007_R14"
                sql &= " ,cs.Midentityid mid" & vbCrLf
            Case cst_rpt工作年資 '"SD_15_007_R15"
                sql &= " ,case when sb.q61<3 then '01'" & vbCrLf
                sql &= " when sb.q61<5 then '02'" & vbCrLf
                sql &= " when sb.q61<10 then '03'" & vbCrLf
                sql &= " else '04' end mid" & vbCrLf
            Case cst_rpt地理分佈 '"SD_15_007_R16"
                sql &= " ,iz.CTID mid" & vbCrLf
            Case cst_rpt公司行業別 '"SD_15_007_R17"
                sql &= " ,sb.Q4 mid" & vbCrLf
            Case cst_rpt公司規模 '"SD_15_007_R18"
                sql &= " ,sb.Q5 mid" & vbCrLf
            Case cst_rpt參訓動機 '"SD_15_007_R19"
                sql &= " ,sb2.Q2A mid" & vbCrLf
            Case cst_rpt訓後動向 '"SD_15_007_R20"
                sql &= " ,sb.Q3 mid" & vbCrLf
            Case cst_rpt參訓單位類別 '"SD_15_007_R21"
                sql &= " ,oo.OrgKind mid" & vbCrLf
            Case cst_rpt參加課程職能別 '"SD_15_007_R22"
                sql &= " ,pp.ClassCate mid" & vbCrLf
            Case cst_rpt參加課程型態別 '"SD_15_007_R23"
                sql &= " ,case when pp.PointYN = 'Y' then 1" & vbCrLf
                sql &= " when pp.PointYN = 'N' then 2" & vbCrLf
                sql &= " when pp.IsBusiness = 'Y' then 3 end mid" & vbCrLf
        End Select
        'sql &= " --- select count(1)" & vbCrLf
        sql &= " FROM STUD_QUESTIONFAC2 fc2" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.SOCID=fc2.SOCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss on ss.SID=cs.SID" & vbCrLf
        sql &= " JOIN STUD_SUBDATA ss2 on ss2.SID=cs.SID" & vbCrLf
        sql &= " JOIN STUD_TRAINBG sb on sb.SOCID=cs.SOCID" & vbCrLf
        sql &= " LEFT JOIN V_TRAINBGQ2 sb2 on sb2.SOCID=cs.SOCID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=cs.OCID" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PLANID=cc.PLANID AND pp.COMIDNO=cc.COMIDNO AND pp.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=cc.COMIDNO" & vbCrLf
        sql &= " JOIN ID_Plan ip on ip.PLANID =cc.PLANID" & vbCrLf
        sql &= " JOIN ID_ZIP iz on iz.ZIPCODE=ss2.ZIPCODE1" & vbCrLf
        'sql &= " WHERE 1=1 AND ip.years='2017'  AND ip.tplanid='28'  AND ip.distid='001'" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If RID <> "" Then sql &= " AND cc.RID LIKE '" & RID & "%'" & vbCrLf
        If OCID <> "" Then sql &= " AND cc.OCID='" & OCID & "'" & vbCrLf
        If TPlanID <> "" Then sql &= " AND ip.TPlanID='" & TPlanID & "'" & vbCrLf
        If Years <> "" Then sql &= " AND ip.Years='" & Years & "'" & vbCrLf
        If PackageType <> "" Then sql &= " AND pp.PackageType='" & PackageType & "'" & vbCrLf
        If SearchPlan <> "" Then sql &= " AND oo.OrgKind2='" & SearchPlan & "'" & vbCrLf
        Return sql
    End Function

    '依filename組這次的群組比較
    Function Get_WC2_SQL(ByRef ss As String) As String
        'Dim filename As String = TIMS.GetMyValue(ss, "filename")
        'Dim RID As String = TIMS.GetMyValue(ss, "RID")
        'Dim OCID As String = TIMS.GetMyValue(ss, "OCID")
        'Dim Years As String = TIMS.GetMyValue(ss, "Years")
        'Dim PackageType As String = TIMS.GetMyValue(ss, "PackageType")
        'Dim TPlanID As String = TIMS.GetMyValue(ss, "TPlanID")
        'Dim SearchPlan As String = TIMS.GetMyValue(ss, "SearchPlan")

        Dim sql As String = ""
        Select Case filename
            Case cst_rpt性別 ' "SD_15_007_R11"
                sql = "" & vbCrLf
                sql &= " select 'M' as mid, '男' as Name  " & vbCrLf
                sql &= " union select 'F' as mid, '女' as Name " & vbCrLf

            Case cst_rpt年齡 '"SD_15_007_R12"
                sql = "" & vbCrLf
                sql &= " select '01' as mid, '15以下' as Name  union" & vbCrLf
                sql &= " select '02' as mid, '15~19' as Name  union" & vbCrLf
                sql &= " select '03' as mid, '20~24' as Name  union" & vbCrLf
                sql &= " select '04' as mid, '25~29' as Name  union" & vbCrLf
                sql &= " select '05' as mid, '30~34' as Name  union" & vbCrLf
                sql &= " select '06' as mid, '35~39' as Name  union" & vbCrLf
                sql &= " select '07' as mid, '40~44' as Name  union" & vbCrLf
                sql &= " select '08' as mid, '45~49' as Name  union" & vbCrLf
                sql &= " select '09' as mid, '50~54' as Name  union" & vbCrLf
                sql &= " select '10' as mid, '55~59' as Name  union" & vbCrLf
                sql &= " select '11' as mid, '60以上' as Name " & vbCrLf

            Case cst_rpt教育程度 '"SD_15_007_R13"
                sql = "" & vbCrLf
                sql &= " select DegreeID mid, Name from Key_Degree where DegreeType=1" & vbCrLf

            Case cst_rpt身分別 '"SD_15_007_R14"
                sql = "" & vbCrLf
                sql &= " select IdentityID mid, Name from key_Identity" & vbCrLf

            Case cst_rpt工作年資 '"SD_15_007_R15"
                sql = "" & vbCrLf
                sql &= " Select '01' as mid, '3年以下' Name  union" & vbCrLf
                sql &= " Select '02' as mid, '3~5年' Name  union" & vbCrLf
                sql &= " Select '03' as mid, '5~10年' Name  union" & vbCrLf
                sql &= " Select '04' as mid, '10年以上' Name " & vbCrLf

            Case cst_rpt地理分佈 '"SD_15_007_R16"
                sql = "" & vbCrLf
                sql &= " Select ctid as mid, ctname as Name from id_city" & vbCrLf

            Case cst_rpt公司行業別 '"SD_15_007_R17"
                sql = "" & vbCrLf
                sql &= " SELECT tradeid as mid, tradename as name FROM key_Trade" & vbCrLf

            Case cst_rpt公司規模 '"SD_15_007_R18"
                sql = "" & vbCrLf
                sql &= " Select 1 as mid, '屬於中小企業' Name  union" & vbCrLf
                sql &= " Select 0 as mid, '非中小企業' Name " & vbCrLf

            Case cst_rpt參訓動機 '"SD_15_007_R19"
                sql = "" & vbCrLf
                sql &= " Select 1 as mid, '為補充與原專長相關之技能' Name  union" & vbCrLf
                sql &= " Select 2 as mid, '轉換其他行職業所需技能' Name  union" & vbCrLf
                sql &= " Select 3 as mid, '拓展工作領域及視野' Name  union" & vbCrLf
                sql &= " Select 4 as mid, '其他' Name " & vbCrLf

            Case cst_rpt訓後動向 '"SD_15_007_R20"
                sql = "" & vbCrLf
                sql &= " Select 1 as mid, '轉換工作' Name  union" & vbCrLf
                sql &= " Select 2 as mid, '留任' Name  union" & vbCrLf
                sql &= " Select 3 as mid, '其他' Name " & vbCrLf

            Case cst_rpt參訓單位類別 '"SD_15_007_R21"
                sql = "" & vbCrLf
                sql &= " Select OrgTypeID as mid, Name from Key_OrgType" & vbCrLf

            Case cst_rpt參加課程職能別 '"SD_15_007_R22"
                sql = "" & vbCrLf
                sql &= " SELECT CCID mid, CCName name FROM KEY_CLASSCATELOG" & vbCrLf

            Case cst_rpt參加課程型態別 '"SD_15_007_R23"
                sql = "" & vbCrLf
                sql &= " Select 1 as mid, '學分班' Name  union" & vbCrLf
                sql &= " Select 2 as mid, '非學分班' Name  union" & vbCrLf
                sql &= " Select 3 as mid, '企業包班' Name " & vbCrLf

        End Select

        Return sql
    End Function

    '取得group table 1 (要排序)
    Function GroupDT1(ByRef dt As DataTable, ByVal sWC1 As String, ByVal sWC2 As String) As DataTable
        Dim sql As String = ""
        'sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & sWC1 & " )" & vbCrLf
        sql &= " ,WC2 AS (" & sWC2 & " )" & vbCrLf
        'sql &= " SELECT MID,COUNT(1) CNT" & vbCrLf
        sql &= " SELECT b.MID" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_1='Y' THEN 1 END) A1_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_2='Y' THEN 1 END) A1_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_3='Y' THEN 1 END) A1_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_4='Y' THEN 1 END) A1_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_5='Y' THEN 1 END) A1_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_6='Y' THEN 1 END) A1_6" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_7='Y' THEN 1 END) A1_7" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_8='Y' THEN 1 END) A1_8" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_9='Y' THEN 1 END) A1_9" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A1_10='Y' THEN 1 END) A1_10" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A2=1 THEN 1 END) A2_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=2 THEN 1 END) A2_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=3 THEN 1 END) A2_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=4 THEN 1 END) A2_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=5 THEN 1 END) A2_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=6 THEN 1 END) A2_6" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A2=7 THEN 1 END) A2_7" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A3=1 THEN 1 END) A3_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A3=2 THEN 1 END) A3_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A3=3 THEN 1 END) A3_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A3=4 THEN 1 END) A3_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A3=5 THEN 1 END) A3_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A4=1 THEN 1 END) A4_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=2 THEN 1 END) A4_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=3 THEN 1 END) A4_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=4 THEN 1 END) A4_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=5 THEN 1 END) A4_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=6 THEN 1 END) A4_6" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=7 THEN 1 END) A4_7" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=8 THEN 1 END) A4_8" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A4=9 THEN 1 END) A4_9" & vbCrLf '訓後意見調查統計表-選項重複
        'sql &= " ,COUNT(CASE WHEN a.A4=10 THEN 1 END) A4_10" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A5=1 THEN 1 END) A5_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=2 THEN 1 END) A5_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=3 THEN 1 END) A5_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=4 THEN 1 END) A5_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=5 THEN 1 END) A5_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=6 THEN 1 END) A5_6" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=7 THEN 1 END) A5_7" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=8 THEN 1 END) A5_8" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A5=9 THEN 1 END) A5_9" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN a.A5=10 THEN 1 END) A5_10" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A6=1 THEN 1 END) A6_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A6=2 THEN 1 END) A6_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A6=3 THEN 1 END) A6_3" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.A7=1 THEN 1 END) A7_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A7=2 THEN 1 END) A7_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A7=3 THEN 1 END) A7_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A7=4 THEN 1 END) A7_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.A7=5 THEN 1 END) A7_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B11=1 THEN 1 END) B11_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B11=2 THEN 1 END) B11_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B11=3 THEN 1 END) B11_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B11=4 THEN 1 END) B11_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B11=5 THEN 1 END) B11_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B12=1 THEN 1 END) B12_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B12=2 THEN 1 END) B12_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B12=3 THEN 1 END) B12_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B12=4 THEN 1 END) B12_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B12=5 THEN 1 END) B12_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B13=1 THEN 1 END) B13_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B13=2 THEN 1 END) B13_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B13=3 THEN 1 END) B13_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B13=4 THEN 1 END) B13_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B13=5 THEN 1 END) B13_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B14=1 THEN 1 END) B14_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B14=2 THEN 1 END) B14_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B14=3 THEN 1 END) B14_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B14=4 THEN 1 END) B14_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B14=5 THEN 1 END) B14_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B15=1 THEN 1 END) B15_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B15=2 THEN 1 END) B15_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B15=3 THEN 1 END) B15_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B15=4 THEN 1 END) B15_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B15=5 THEN 1 END) B15_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B21=1 THEN 1 END) B21_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B21=2 THEN 1 END) B21_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B21=3 THEN 1 END) B21_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B21=4 THEN 1 END) B21_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B21=5 THEN 1 END) B21_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B22=1 THEN 1 END) B22_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B22=2 THEN 1 END) B22_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B22=3 THEN 1 END) B22_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B22=4 THEN 1 END) B22_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B22=5 THEN 1 END) B22_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B23=1 THEN 1 END) B23_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B23=2 THEN 1 END) B23_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B23=3 THEN 1 END) B23_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B23=4 THEN 1 END) B23_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B23=5 THEN 1 END) B23_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B31=1 THEN 1 END) B31_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B31=2 THEN 1 END) B31_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B31=3 THEN 1 END) B31_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B31=4 THEN 1 END) B31_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B31=5 THEN 1 END) B31_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B32=1 THEN 1 END) B32_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B32=2 THEN 1 END) B32_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B32=3 THEN 1 END) B32_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B32=4 THEN 1 END) B32_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B32=5 THEN 1 END) B32_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B41=1 THEN 1 END) B41_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B41=2 THEN 1 END) B41_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B41=3 THEN 1 END) B41_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B41=4 THEN 1 END) B41_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B41=5 THEN 1 END) B41_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B42=1 THEN 1 END) B42_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B42=2 THEN 1 END) B42_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B42=3 THEN 1 END) B42_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B42=4 THEN 1 END) B42_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B42=5 THEN 1 END) B42_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B43=1 THEN 1 END) B43_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B43=2 THEN 1 END) B43_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B43=3 THEN 1 END) B43_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B43=4 THEN 1 END) B43_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B43=5 THEN 1 END) B43_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B44=1 THEN 1 END) B44_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B44=2 THEN 1 END) B44_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B44=3 THEN 1 END) B44_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B44=4 THEN 1 END) B44_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B44=5 THEN 1 END) B44_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B51=1 THEN 1 END) B51_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B51=2 THEN 1 END) B51_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B51=3 THEN 1 END) B51_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B51=4 THEN 1 END) B51_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B51=5 THEN 1 END) B51_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B61=1 THEN 1 END) B61_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B61=2 THEN 1 END) B61_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B61=3 THEN 1 END) B61_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B61=4 THEN 1 END) B61_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B61=5 THEN 1 END) B61_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B62=1 THEN 1 END) B62_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B62=2 THEN 1 END) B62_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B62=3 THEN 1 END) B62_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B62=4 THEN 1 END) B62_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B62=5 THEN 1 END) B62_5" & vbCrLf

        sql &= " ,COUNT(CASE WHEN a.B63=1 THEN 1 END) B63_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B63=2 THEN 1 END) B63_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B63=3 THEN 1 END) B63_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B63=4 THEN 1 END) B63_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B63=5 THEN 1 END) B63_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B71=1 THEN 1 END) B71_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B71=2 THEN 1 END) B71_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B71=3 THEN 1 END) B71_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B71=4 THEN 1 END) B71_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B71=5 THEN 1 END) B71_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B72=1 THEN 1 END) B72_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B72=2 THEN 1 END) B72_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B72=3 THEN 1 END) B72_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B72=4 THEN 1 END) B72_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B72=5 THEN 1 END) B72_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B73=1 THEN 1 END) B73_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B73=2 THEN 1 END) B73_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B73=3 THEN 1 END) B73_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B73=4 THEN 1 END) B73_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B73=5 THEN 1 END) B73_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B74=1 THEN 1 END) B74_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B74=2 THEN 1 END) B74_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B74=3 THEN 1 END) B74_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B74=4 THEN 1 END) B74_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.B74=5 THEN 1 END) B74_5" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.C11=1 THEN 1 END) C11_1" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.C11=2 THEN 1 END) C11_2" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.C11=3 THEN 1 END) C11_3" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.C11=4 THEN 1 END) C11_4" & vbCrLf
        sql &= " ,COUNT(CASE WHEN a.C11=5 THEN 1 END) C11_5" & vbCrLf
        sql &= " FROM WC2 b" & vbCrLf
        sql &= " LEFT JOIN WC1 a on a.mid=b.mid" & vbCrLf
        sql &= " GROUP BY b.MID" & vbCrLf
        sql &= " ORDER BY b.MID" & vbCrLf
        'Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        Return dt
    End Function

    '取得group table 2 (要排序)
    Function GroupDT2(ByVal sWC2 As String) As DataTable
        Dim sql As String = ""
        'sql = "" & vbCrLf
        sql &= " WITH WC2 AS (" & sWC2 & " )" & vbCrLf
        sql &= " select mid,name from WC2 ORDER BY mid" & vbCrLf
        sourceDT = DbAccess.GetDataTable(sql, objconn)
        Return sourceDT
    End Function

    Sub SUB_Rpt2A(ByVal ss As String)
        'Dim filename As String = TIMS.GetMyValue(ss, "filename")
        Dim sWC1 As String = Get_WC1_SQL(ss)
        Dim sWC2 As String = Get_WC2_SQL(ss)
        Dim dt As DataTable = Nothing
        dt = GroupDT1(dt, sWC1, sWC2)
        sourceDT = GroupDT2(sWC2)

        Dim tDt As New DataTable
        tDt = RotationDT(dt)
        Select Case filename
            Case cst_rpt性別 ' "SD_15_007_R11"
            Case cst_rpt年齡 '"SD_15_007_R12"
            Case cst_rpt教育程度 '"SD_15_007_R13"
            Case cst_rpt身分別 '"SD_15_007_R14"
            Case cst_rpt工作年資 '"SD_15_007_R15"
            Case cst_rpt地理分佈 '"SD_15_007_R16"
            Case cst_rpt公司行業別 '"SD_15_007_R17"
            Case cst_rpt公司規模 '"SD_15_007_R18"
            Case cst_rpt參訓動機 '"SD_15_007_R19"
            Case cst_rpt訓後動向 '"SD_15_007_R20"
            Case cst_rpt參訓單位類別 '"SD_15_007_R21"
            Case cst_rpt參加課程職能別 '"SD_15_007_R22"
            Case cst_rpt參加課程型態別 '"SD_15_007_R23"
        End Select

        Select Case filename
            Case cst_rpt性別 ' "SD_15_007_R11"
                'tDt = db_R11(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 40, "10", "H", ss)
                'exc_Print("true")
            Case cst_rpt年齡 '"SD_15_007_R12"
                'tDt = db_R12(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "120", "120", 40, "9", "H", ss)
                'exc_Print("true")
            Case cst_rpt教育程度 '"SD_15_007_R13"
                'tDt = db_R13(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "130", "140", 25, "9", "L", ss)
                'exc_Print("false")
            Case cst_rpt身分別 '"SD_15_007_R14"
                'tDt = db_R14(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "80", "60", 40, "6", "L", ss)
                'exc_Print("false")
            Case cst_rpt工作年資 '"SD_15_007_R15"
                'tDt = db_R15(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "180", "160", 50, "10", "H", ss)
                'exc_Print("true")
            Case cst_rpt地理分佈 '"SD_15_007_R16"
                'tDt = db_R16(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "90", "80", 35, "8", "L", ss)
                'exc_Print("false")
            Case cst_rpt公司行業別 '"SD_15_007_R17"
                'tDt = db_R17(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "70", "60", 35, "6", "L", ss)
                'exc_Print("false")
            Case cst_rpt公司規模 '"SD_15_007_R18"
                'tDt = db_R18(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "180", "160", 35, "11", "L", ss)
                'exc_Print("false")
            Case cst_rpt參訓動機 '"SD_15_007_R19"
                'tDt = db_R19(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "120", "120", 40, "10", "H", ss)
                'exc_Print("true")
            Case cst_rpt訓後動向 '"SD_15_007_R20"
                'tDt = db_R20(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 40, "11", "H", ss)
                'exc_Print("true")
            Case cst_rpt參訓單位類別 '"SD_15_007_R21"
                'tDt = db_R21(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "100", "100", 25, "9", "L", ss)
                'exc_Print("false")
            Case cst_rpt參加課程職能別 '"SD_15_007_R22"
                'tDt = db_R22(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "130", "120", 45, "10", "H", ss)
                'exc_Print("true")
            Case cst_rpt參加課程型態別 '"SD_15_007_R23"
                'tDt = db_R23(TPlanID, Years, OCID, RID, SearchPlan, PackageType)
                Call PrintDiv(tDt, filename, "150", "150", 45, "11", "H", ss)
                'exc_Print("true")
        End Select

        'Return tDt
    End Sub

    Sub PrintDiv(ByVal dt As DataTable, ByVal selRpt As String, ByVal Field1_width As String, ByVal Field2_width As String, ByVal RCount As Integer, ByVal font_size As String, ByVal portrait As String, ByVal ss As String)
        'dt:要顯示的資料,selRpt:,Field1_width:標題題目的寬度,Field2_width:標題題目的寬度,RCount:每頁筆數,font_size:內容字型大小,portrait:直式/橫式
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim plankind As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        Dim SearchPlan As String = TIMS.GetMyValue(ss, "SearchPlan")
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select '('+name+')' name from v_orgkind1 where VALUE='" & SearchPlan & "'" & vbCrLf
        Dim dtn As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dtn.Rows.Count > 0 Then
            plankind &= Convert.ToString(dtn.Rows(0)("name"))
        End If

        Dim Years As String = TIMS.GetMyValue(ss, "Years")
        Dim PName As String = String.Concat(Val(Years) - 1911, "年度　訓後意見調查統計表  ")
        Select Case filename
            Case cst_rpt性別 ' "SD_15_007_R11"
                PName &= "性別"
            Case cst_rpt年齡 '"SD_15_007_R12"
                PName &= "年齡"
            Case cst_rpt教育程度 '"SD_15_007_R13"
                PName &= "教育程度"
            Case cst_rpt身分別 '"SD_15_007_R14"
                PName &= "身分別"
            Case cst_rpt工作年資 '"SD_15_007_R15"
                PName &= "工作年資"
            Case cst_rpt地理分佈 '"SD_15_007_R16"
                PName &= "地理分佈"
            Case cst_rpt公司行業別 '"SD_15_007_R17"
                PName &= "公司行業別"
            Case cst_rpt公司規模 '"SD_15_007_R18"
                PName &= "公司規模"
            Case cst_rpt參訓動機 '"SD_15_007_R19"
                PName &= "參訓動機"
            Case cst_rpt訓後動向 '"SD_15_007_R20"
                PName &= "訓後動向"
            Case cst_rpt參訓單位類別 '"SD_15_007_R21"
                PName &= "參訓單位類別"
            Case cst_rpt參加課程職能別 '"SD_15_007_R22"
                PName &= "參加課程職能別"
            Case cst_rpt參加課程型態別 '"SD_15_007_R23"
                PName &= "參加課程型態別 "
        End Select

        Dim tmpDT As New DataTable
        'Dim tmpDR As DataRow
        'Dim tmpObj As Object
        'Dim sql As String = ""
        sql = ""
        Dim PageCount As Int32 = 0  'Pages
        Dim ReportCount As Integer = RCount '每頁筆數
        Dim ColCount As Integer = 0
        Dim intTmp As Integer = 0
        Dim rsCursor As Integer = 0   '報表內容列印的NO
        Dim intPageRecord As Integer = RCount '每頁列印幾筆

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:" + font_size + "pt;font-family:DFKai-SB"
        Dim int_width As Integer
        Dim strWatermarkImg As String
        Dim strWatermarkDiv As String
        Dim intWatermarkTop As Integer
        'Dim divWatermark

        tmpDT = dt
        ColCount = dt.Columns.Count

        intTmp = tmpDT.Rows.Count
        If (intTmp Mod ReportCount) = 0 Then
            PageCount = (intTmp / ReportCount) - 1
        Else
            PageCount = intTmp / ReportCount
        End If

        '表格寬度的設定
        If portrait = "H" Then
            int_width = Int((550 - Field1_width - Field2_width) / sourceDT.Rows.Count)
            strWatermarkImg = "TIMS_1.jpg"
        Else
            int_width = Int((820 - Field1_width - Field2_width) / sourceDT.Rows.Count)
            strWatermarkImg = "TIMS_2.jpg"
        End If

        If dt.Rows.Count > 0 Then

            For i As Integer = 0 To PageCount
                '加背景圖的div
                If portrait = "H" Then
                    intWatermarkTop = i * 800
                Else
                    intWatermarkTop = i * 550
                End If
                strWatermarkDiv = "<div style='position:absolute;z-index:-1; margin:0;padding:0;left:0px;top: " + intWatermarkTop.ToString + "px;'><img src='../../images/rptpic/temple/" + strWatermarkImg + "' /></div>"
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = strWatermarkDiv

                '表頭
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "勞動部勞動力發展署"

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = plankind

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "列印日期：" + Now().ToShortDateString()

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = PName

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "頁數：" + (i + 1).ToString + " / " + PageCount.ToString

                'Column Header
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "2")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", Field1_width)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2")
                nc.InnerHtml = "&nbsp;"

                For j As Integer = 0 To sourceDT.Rows.Count - 1
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    nc.Attributes.Add("align", "center")
                    nc.Attributes.Add("style", strStyle + ";word-break:break-all")
                    nc.Attributes.Add("width", int_width)

                    nc.InnerHtml = sourceDT.Rows(j)("name").ToString
                Next

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("width", int_width)
                nc.InnerHtml = "總計"
                '報表內容
                Dim strTmp As String = ""
                Dim strAlign As String = ""
                For j As Integer = 0 To intPageRecord
                    If rsCursor >= dt.Rows.Count Then
                        GoTo [CONTINUE]
                    End If

                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    For m As Integer = 0 To dt.Columns.Count - 1

                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)

                        If m = 0 Or m = 1 Then
                            strAlign = "left"
                            If m = 0 Then
                                nc.Attributes.Add("width", Field1_width)
                            Else
                                nc.Attributes.Add("width", Field2_width)
                            End If
                        Else
                            strAlign = "right"
                        End If

                        strTmp = dt.Rows(rsCursor)(m).ToString
                        nc.Attributes.Add("align", strAlign)
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = strTmp

                    Next
                    rsCursor += 1
                Next


[CONTINUE]:
                '表尾
                If rsCursor + 1 > tmpDT.Rows.Count Then
                    GoTo out
                End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)

                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"

            Next
out:
        End If

    End Sub

    Sub setItemArray()
        arrItem(0, 0) = "您獲得本次課程的訊息來源（可複選）"
        arrItem(0, 1) = "本署或分署網站"
        arrItem(1, 1) = "就業服務中心"
        arrItem(2, 1) = "訓練單位"
        arrItem(3, 1) = "搜尋網站"
        arrItem(4, 1) = "報紙"
        arrItem(5, 1) = "廣播"
        arrItem(6, 1) = "電視"
        arrItem(7, 1) = "親友介紹"
        arrItem(8, 1) = "社群媒體(ex：臉書、LINE)"
        arrItem(9, 1) = "其他" 'A1_10

        arrItem(10, 0) = "參加本次課程的主要原因"
        arrItem(10, 1) = "課程符合就業市場需求"
        arrItem(11, 1) = "課程符合目前工作需求"
        arrItem(12, 1) = "課程符合個人興趣"
        arrItem(13, 1) = "可取得課程相關證照或證書"
        arrItem(14, 1) = "學習第二專長"
        arrItem(15, 1) = "師資具知名度或專業性"
        arrItem(16, 1) = "其他" 'A2_7

        arrItem(17, 0) = "選擇本訓練單位的主要原因"
        arrItem(17, 1) = "環境、設備良好"
        arrItem(18, 1) = "具課程專業度"
        arrItem(19, 1) = "行政人員服務良好"
        arrItem(20, 1) = "為訓練單位之會員"
        arrItem(21, 1) = "其他" 'A3_5

        arrItem(22, 0) = "沒有參加本方案訓練之前，每年參加訓練支出的費用？"
        arrItem(22, 1) = "0元"
        arrItem(23, 1) = "999元以下"
        arrItem(24, 1) = "1,000元-3,999元"
        arrItem(25, 1) = "4,000元-6,999元"
        arrItem(26, 1) = "7,000元-9,999元"
        arrItem(27, 1) = "10,000元-19,999元"
        arrItem(28, 1) = "20,000元-29,999元"
        arrItem(29, 1) = "30,000元-39,999元"
        arrItem(30, 1) = "40,000元以上" 'A4_9

        arrItem(31, 0) = "如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額？"
        arrItem(31, 1) = "0元"
        arrItem(32, 1) = "999元以下"
        arrItem(33, 1) = "1,000元-3,999元"
        arrItem(34, 1) = "4,000元-6,999元"
        arrItem(35, 1) = "7,000元-9,999元"
        arrItem(36, 1) = "10,000元-19,999元"
        arrItem(37, 1) = "20,000元-29,999元"
        arrItem(38, 1) = "30,000元-39,999元"
        arrItem(39, 1) = "40,000元以上" 'A5_9

        arrItem(40, 0) = "您認為本次課程的訓練費用是否合理？"
        arrItem(40, 1) = "偏高"
        arrItem(41, 1) = "合理"
        arrItem(42, 1) = "偏低" 'A6_3

        arrItem(43, 0) = "結訓後對於工作的規劃？"
        arrItem(43, 1) = "留在原職位"
        arrItem(44, 1) = "轉調較能發揮潛能部門"
        arrItem(45, 1) = "轉換至同業的其他公司"
        arrItem(46, 1) = "轉換至不同行業的公司"
        arrItem(47, 1) = "希望自己創業" 'A7_5

        arrItem(48, 0) = "課程內容符合期望"
        arrItem(48, 1) = "非常同意"
        arrItem(49, 1) = "同意"
        arrItem(50, 1) = "普通"
        arrItem(51, 1) = "不同意"
        arrItem(52, 1) = "非常不同意" 'B11_5

        arrItem(53, 0) = "課程難易安排適當"
        arrItem(53, 1) = "非常同意"
        arrItem(54, 1) = "同意"
        arrItem(55, 1) = "普通"
        arrItem(56, 1) = "不同意"
        arrItem(57, 1) = "非常不同意" 'B12_5

        arrItem(58, 0) = "課程總時數適當"
        arrItem(58, 1) = "非常同意"
        arrItem(59, 1) = "同意"
        arrItem(60, 1) = "普通"
        arrItem(61, 1) = "不同意"
        arrItem(62, 1) = "非常不同意" 'B13_5

        arrItem(63, 0) = "課程符合實務需求"
        arrItem(63, 1) = "非常同意"
        arrItem(64, 1) = "同意"
        arrItem(65, 1) = "普通"
        arrItem(66, 1) = "不同意"
        arrItem(67, 1) = "非常不同意" 'B14_5

        arrItem(68, 0) = "課程符合產業發展趨勢"
        arrItem(68, 1) = "非常同意"
        arrItem(69, 1) = "同意"
        arrItem(70, 1) = "普通"
        arrItem(71, 1) = "不同意"
        arrItem(72, 1) = "非常不同意" 'B15_5

        arrItem(73, 0) = "滿意講師的教學態度"
        arrItem(73, 1) = "非常同意"
        arrItem(74, 1) = "同意"
        arrItem(75, 1) = "普通"
        arrItem(76, 1) = "不同意"
        arrItem(77, 1) = "非常不同意" 'B21_5

        arrItem(78, 0) = "滿意講師的教學方法"
        arrItem(78, 1) = "非常同意"
        arrItem(79, 1) = "同意"
        arrItem(80, 1) = "普通"
        arrItem(81, 1) = "不同意"
        arrItem(82, 1) = "非常不同意" 'B22_5

        arrItem(83, 0) = "滿意講師的課程專業度"
        arrItem(83, 1) = "非常同意"
        arrItem(84, 1) = "同意"
        arrItem(85, 1) = "普通"
        arrItem(86, 1) = "不同意"
        arrItem(87, 1) = "非常不同意" 'B23_5

        arrItem(88, 0) = "對於訓練教材感到滿意"
        arrItem(88, 1) = "非常同意"
        arrItem(89, 1) = "同意"
        arrItem(90, 1) = "普通"
        arrItem(91, 1) = "不同意"
        arrItem(92, 1) = "非常不同意" 'B31_5

        arrItem(93, 0) = "訓練教材能夠輔助課程學習"
        arrItem(93, 1) = "非常同意"
        arrItem(94, 1) = "同意"
        arrItem(95, 1) = "普通"
        arrItem(96, 1) = "不同意"
        arrItem(97, 1) = "非常不同意" 'B32_5

        arrItem(98, 0) = "您對於訓練場地感到滿意"
        arrItem(98, 1) = "非常同意"
        arrItem(99, 1) = "同意"
        arrItem(100, 1) = "普通"
        arrItem(101, 1) = "不同意"
        arrItem(102, 1) = "非常不同意" 'B41_5

        arrItem(103, 0) = "您對於訓練設備感到滿意"
        arrItem(103, 1) = "非常同意"
        arrItem(104, 1) = "同意"
        arrItem(105, 1) = "普通"
        arrItem(106, 1) = "不同意"
        arrItem(107, 1) = "非常不同意" 'B42_5

        arrItem(108, 0) = "您認為實作設備的數量適當"
        arrItem(108, 1) = "非常同意"
        arrItem(109, 1) = "同意"
        arrItem(110, 1) = "普通"
        arrItem(111, 1) = "不同意"
        arrItem(112, 1) = "非常不同意" 'B43_5

        arrItem(113, 0) = "您認為實作設備新穎"
        arrItem(113, 1) = "非常同意"
        arrItem(114, 1) = "同意"
        arrItem(115, 1) = "普通"
        arrItem(116, 1) = "不同意"
        arrItem(117, 1) = "非常不同意" 'B44_5

        arrItem(118, 0) = "能促進學習效果"
        arrItem(118, 1) = "非常同意"
        arrItem(119, 1) = "同意"
        arrItem(120, 1) = "普通"
        arrItem(121, 1) = "不同意"
        arrItem(122, 1) = "非常不同意" 'B51_5

        arrItem(123, 0) = "您認為在訓練課程中，課程內容能讓您專注"
        arrItem(123, 1) = "非常同意"
        arrItem(124, 1) = "同意"
        arrItem(125, 1) = "普通"
        arrItem(126, 1) = "不同意"
        arrItem(127, 1) = "非常不同意" 'B61_5

        arrItem(128, 0) = "您在完成訓練後，已充份學習訓練課程所教授知識或技能"
        arrItem(128, 1) = "非常同意"
        arrItem(129, 1) = "同意"
        arrItem(130, 1) = "普通"
        arrItem(131, 1) = "不同意"
        arrItem(132, 1) = "非常不同意" 'B62_5

        arrItem(133, 0) = "您在完成訓練後，有學習到新的知識或技能"
        arrItem(133, 1) = "非常同意"
        arrItem(134, 1) = "同意"
        arrItem(135, 1) = "普通"
        arrItem(136, 1) = "不同意"
        arrItem(137, 1) = "非常不同意" 'B63_5

        arrItem(138, 0) = "您對於訓練單位的課程安排與授課情形感到滿意"
        arrItem(138, 1) = "非常同意"
        arrItem(139, 1) = "同意"
        arrItem(140, 1) = "普通"
        arrItem(141, 1) = "不同意"
        arrItem(142, 1) = "非常不同意" 'B71_5

        arrItem(143, 0) = "您對於訓練單位的行政服務感到滿意"
        arrItem(143, 1) = "非常同意"
        arrItem(144, 1) = "同意"
        arrItem(145, 1) = "普通"
        arrItem(146, 1) = "不同意"
        arrItem(147, 1) = "非常不同意" 'B72_5

        arrItem(148, 0) = "您對於產業人才投資方案感到滿意"
        arrItem(148, 1) = "非常同意"
        arrItem(149, 1) = "同意"
        arrItem(150, 1) = "普通"
        arrItem(151, 1) = "不同意"
        arrItem(152, 1) = "非常不同意" 'B73_5

        arrItem(153, 0) = "您認為完成本訓練課程對於目前或未來工作有幫助"
        arrItem(153, 1) = "非常同意"
        arrItem(154, 1) = "同意"
        arrItem(155, 1) = "普通"
        arrItem(156, 1) = "不同意"
        arrItem(157, 1) = "非常不同意" 'B74_5

        arrItem(158, 0) = "若本訓練課程沒有補助，是否會全額自費參訓？ "
        arrItem(158, 1) = "一定會"
        arrItem(159, 1) = "應該會"
        arrItem(160, 1) = "普通"
        arrItem(161, 1) = "應該不會"
        arrItem(162, 1) = "一定不會" 'C11_5
    End Sub

    Function RotationDT(ByVal dt As DataTable) As DataTable
        Dim tmpDT As New DataTable
        Dim tmpDR As DataRow
        'Dim tmpObj As Object
        Dim no As Integer = 0
        Dim strItem As String
        Dim strQues As String
        Dim intTotal As Integer = 0

        Call setItemArray() '題目內容

        tmpDT.Columns.Add(New DataColumn("Item"))   '題目
        tmpDT.Columns.Add(New DataColumn("Ques"))   '答案
        For j As Integer = 0 To dt.Rows.Count - 1
            tmpDT.Columns.Add(New DataColumn(dt.Rows(j)(0)))
        Next
        tmpDT.Columns.Add(New DataColumn("Total"))  '答案(總計)

        Const cst_disp As Integer = 1 '位移量()

        For ia1 As Integer = cst_disp To dt.Columns.Count - 1
            tmpDR = tmpDT.NewRow
            tmpDT.Rows.Add(tmpDR)

            strItem = ""
            strQues = ""
            intTotal = 0
            If Not IsNothing(arrItem(ia1 - cst_disp, 0)) Then tmpDR("item") = arrItem(ia1 - cst_disp, 0).ToString
            If Not IsNothing(arrItem(ia1 - cst_disp, 1)) Then tmpDR("Ques") = arrItem(ia1 - cst_disp, 1).ToString

            For j As Integer = 0 To dt.Rows.Count - 1
                no = j + 2 '(前面2格已使用)
                tmpDR(no) = dt.Rows(j)(ia1)
                intTotal += dt.Rows(j)(ia1)
            Next
            tmpDR("Total") = intTotal
        Next

        Return tmpDT
    End Function

    Sub exc_Print(ByVal portrait As String)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return

        'Dim strScript As String = ""
        'strScript = "<script language=""javascript"">" + vbCrLf
        ''strScript = "function print() {"
        'strScript += "if (!factory.object) {"
        ''strScript += "return"
        'strScript += "} else {"
        'strScript += "factory.printing.header = """";"
        'strScript += "factory.printing.footer = """";"
        'strScript += "factory.printing.leftMargin = 5; "
        'strScript += "factory.printing.topMargin = 10; "
        'strScript += "factory.printing.rightMargin = 5; "
        'strScript += "factory.printing.bottomMargin = 10; "
        'strScript += "factory.printing.portrait = " + portrait + ";"
        'strScript += "factory.printing.Print(true);"
        'strScript += "window.close();"
        'strScript += "}"
        ''strScript += "}"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub

    Protected Sub btnPrt_Click(sender As Object, e As EventArgs) Handles btnPrt.Click
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")

        Select Case filename
            Case "SD_15_007_R11"
                Call exc_Print("true")
            Case "SD_15_007_R12"
                Call exc_Print("true")
            Case "SD_15_007_R13"
                Call exc_Print("false")
            Case "SD_15_007_R14"
                Call exc_Print("false")
            Case "SD_15_007_R15"
                Call exc_Print("true")
            Case "SD_15_007_R16"
                Call exc_Print("false")
            Case "SD_15_007_R17"
                Call exc_Print("false")
            Case "SD_15_007_R18"
                Call exc_Print("false")
            Case "SD_15_007_R19"
                Call exc_Print("true")
            Case "SD_15_007_R20"
                Call exc_Print("true")
            Case "SD_15_007_R21"
                Call exc_Print("false")
            Case "SD_15_007_R22"
                Call exc_Print("true")
            Case "SD_15_007_R23"
                Call exc_Print("true")
        End Select

    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出Excel
    Protected Sub btnExcel_Click(sender As Object, e As EventArgs) Handles btnExcel.Click
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")
        trBtn.Visible = False

        'Dim ExcelfileName As String = "訓後動態調查統計表2.xls"
        Dim ExcelfileName As String = "訓後意⾒調查統計表2.xls"
        'If Request.Browser.Browser = "IE" Then
        '    fileName = Server.UrlPathEncode(fileName)
        'End If
        ExcelfileName = HttpUtility.UrlEncode(ExcelfileName, System.Text.Encoding.UTF8)

        Dim strContentDisposition As String = [String].Format("{0}; filename=""{1}""", "attachment", ExcelfileName)
        Response.Clear()
        Response.ClearHeaders()
        Response.Buffer = True
        'Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", strContentDisposition)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        'Response.ContentType = "application/ms-excel;charset=big5"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=Big5>")
        'Response.AddHeader("Content-Disposition", strContentDisposition)
        'Response.ContentType = "Application/vnd.ms-excel"
        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        div_print.RenderControl(htw)
        Common.RespWrite(Me, sw.ToString().Replace("<div ", "<!-- ").Replace("</div>", "<!---->"))
        Call TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

    Protected Sub btnExpOds1_Click(sender As Object, e As EventArgs) Handles btnExpOds1.Click
        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")
        trBtn.Visible = False

        Dim sw As New System.IO.StringWriter
        Dim htw As HtmlTextWriter = New HtmlTextWriter(sw)
        div_print.RenderControl(htw)

        Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()

        Dim strHTML As String = ""
        strHTML &= (sw.ToString().Replace("<div>", "").Replace("</div>", ""))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", "ODS") 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("xlsx_buf", buf)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        'Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub
End Class