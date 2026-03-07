Partial Class SD_15_024
    Inherits AuthBasePage

    '#Region "參數/變數 設定"

    '外部帶入 年度訓練計畫特定對象
    'Const Cst_MIdentityID_507 As String = "'02','03','04','05','06','07','10','13','14','27','28','35','36'"
    'TIMS.Cst_Identity06_2019_2

    Const Cst_性別 As String = "1"
    Const Cst_年齡 As String = "2"
    Const Cst_教育程度 As String = "3"
    Const Cst_身分別 As String = "4"           'Cst_特定對象
    Const Cst_受訓學員地理分布 As String = "5"   'CTID (報名) 通訊地址
    Const Cst_受訓學員地理分布2 As String = "6"  'CTID2 (學員) 戶籍地址
    'Const Cst_參訓單位類別 As String = "7"
    Const Cst_開班縣市 As String = "8"         '開班縣市 -ID_City -TaddCTID='" & dr("CTID") -CTName
    Const Cst_訓練時數 As String = "9"         '訓練時數

    '訓練職類(大類
    Const Cst_訓練職類大類 As String = "21"
    'Const Cst_就職狀況 As String = "22" '就職狀況
    Const Cst_訓練職類中類 As String = "23"
    Const Cst_訓練職類小類 As String = "24"

    Const Cst_報名人數 As String = "11"
    Const Cst_開訓人數 As String = "12"
    Const Cst_結訓人數 As String = "13"
    'Const Cst_就業人數 As String = "14"
    'Const Cst_在職者 As String = "15" '在職者(托育及照服員計畫)
    Const Cst_甄試人數 As String = "31"
    Const Cst_離訓人數 As String = "32"

    '年齡變數範圍 (V_YEARSOLD3)
    'Const cst_ageInStr As String = "1,2,3,4,5,7,8,9,10,11,12"

    Const cst_sql_1 As String = "sql_1" '只要組合sql 
    Const cst_sql_2 As String = "sql_2" '組合sql，要產生查詢
    Const cst_vsSqlString As String = "SqlString"
    Const cst_vs_parms1 As String = "vs_parms1"

    'https://cm.turbotech.com.tw/browse/TIMS-2223
    '首頁>>訓練需求管理>>統計分析>>交叉分析統計表
    '本功能會開放給縣市政府承辦人使用,修改說明如下:
    '1.當縣市政府承辦人登入時,請依登入計畫,鎖定查詢條件"計畫範圍"為勾選於登入計畫,其他的計畫不提供勾選.
    Dim flag_Login1 As Boolean = False '縣市政府承辦人登入時 為 True 其餘為 False
    '2.查詢條件"訓練機構",縣市政府承辦人點選時,直接鎖定登入之年度計畫,並只可選該縣市政府或其底下之訓練單位,若訓練機構選縣市政府,則統計其轄下所有訓練單位之資料.

    Dim objconn As SqlConnection

    '#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '#Region "在這裡放置使用者程式碼以初始化網頁"
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '縣市政府承辦人登入時 為 True 其餘為 False
        flag_Login1 = TIMS.Chk_LoginUserType1(Me)

        If Not IsPostBack Then
            'ClientScript.RegisterHiddenField("isPostBack", "1")
            'Call UPSET_DYNAMIC1(rblDYNAMIC1)
            Call cCreate1()
        End If


    End Sub

#Region "NO USE"
    'Dim strScript As String = "<script>choice_DYNAMIC1();</script>"
    'TIMS.RegisterStartupScript(Me, TIMS.xBlockName(), strScript)
    'Dim strScript As String = "<script>choice_DYNAMIC1();</script>"
    'Common.RespWrite(Page, strScript)
    'Dim xErrmsg As String = Today
    'Common.AddClientScript(Me, "alert('" & xErrmsg & "');")
    'If Not IsPostBack Then Call close_page_load()
    '#End Region
#End Region

#Region "COMMON 1"

    Sub cCreate1()
        Call Get_XRoll_YRoll()
        Call Create301()
        'Call Create501()
        'Call Create502()
        'Call Create507()
        'Call Create508()
        'Call Create510()
        'Call Create511()
        'Call Create303()
        'Call Create304()
        'Call Create307()
        'Call Create308()
    End Sub

    Sub Get_XRoll_YRoll()
        Dim str_AryColum1 As String = "0"
        str_AryColum1 &= ",性別" '1
        str_AryColum1 &= ",年齡" '2
        str_AryColum1 &= ",教育程度" '3
        str_AryColum1 &= ",身分別" '4
        str_AryColum1 &= ",受訓學員(通訊)地理分佈" '5
        'str_AryColum1 &= ",參訓單位類別" '7
        str_AryColum1 &= ",開班縣市" '8
        str_AryColum1 &= ",訓練時數" '9
        str_AryColum1 &= ",訓練職類(大類)" '21
        str_AryColum1 &= ",訓練職類(中類)" '23
        str_AryColum1 &= ",訓練職類(小類)" '24

        Dim str_AryColum2 As String = "0"
        str_AryColum2 &= ",1" '性別"
        str_AryColum2 &= ",2" '年齡"
        str_AryColum2 &= ",3" '教育程度"
        str_AryColum2 &= ",4" '身分別"
        str_AryColum2 &= ",5" '受訓學員(通訊)地理分佈"
        'str_AryColum2 &= ",7" '參訓單位類別"
        str_AryColum2 &= ",8" '開班縣市"
        str_AryColum2 &= ",9" '訓練時數"
        str_AryColum2 &= ",21" '訓練職類(大類)"
        str_AryColum2 &= ",23" '訓練職類(中類)" '23
        str_AryColum2 &= ",24" '訓練職類(小類)" '24

        Dim str_Ary1 As String() = str_AryColum1.Split(",")
        Dim str_Ary2 As String() = str_AryColum2.Split(",")

        XRoll.Items.Clear()
        YRoll.Items.Clear()
        Dim ix As Integer = 0
        For Each strA2 As String In str_Ary2
            If strA2 <> "0" Then
                With XRoll
                    .Items.Add(New ListItem(str_Ary1(ix), strA2))
                End With
                With YRoll
                    .Items.Add(New ListItem(str_Ary1(ix), strA2))
                End With
            End If
            ix += 1
        Next
    End Sub

#End Region


#Region "NOUSE"
    'Sub UPSET_DYNAMIC1(ByRef rblObj As RadioButtonList)
    '    With rblObj
    '        .Items.Clear()
    '        .Items.Add(New ListItem("交叉分析統計表", "CM_03_011")) '(這個base程式) '311:
    '        .Items.Add(New ListItem("年度職業訓練行業別_性別分佈", "TR_05_002_R"))  '502:
    '        .Items.Add(New ListItem("年度訓練人數統計_依行業別", "TR_05_001_R"))    '501:
    '        .Items.Add(New ListItem("年度訓練計畫特定對象人數分佈", "TR_05_007_R")) '507:
    '        .Items.Add(New ListItem("訓練計畫特定對象人數統計表", "TR_05_008_R"))   '508:
    '        .Items.Add(New ListItem("訓練時數統計分析", "TR_05_010_R")) '510
    '        .Items.Add(New ListItem("訓練職類統計分析", "TR_05_011_R")) '511

    '        '.Items.Add(New ListItem("訓練人數綜合查詢", "CM_03_003")) '303 '無法整併
    '        '.Items.Add(New ListItem("結訓人數綜合查詢", "CM_03_004")) '304 '無法整併
    '        '.Items.Add(New ListItem("主要特定對象統計表", "CM_03_007")) '307
    '        '.Items.Add(New ListItem("離退訓人數統計表", "CM_03_008")) '308
    '        '.Items.Add(New ListItem("志願役人數統計表", "CM_03_012"))
    '    End With
    '    'Common.SetListItem(rblObj, "CM_03_011")
    '    'rblObj.Attributes
    '    rblObj.Attributes("onclick") = "return choice_DYNAMIC1();"
    'End Sub
#End Region

#Region "CM_03_011"
    Sub Create301()
        DataGroupTable.Visible = False
        FTDate2.Text = TIMS.Cdate3(Now.Date)
        ddlYear = TIMS.GetSyear(ddlYear, 0, 0, True, TIMS.cst_ddl_NotCase)
        Common.SetListItem(ddlYear, sm.UserInfo.Years)

        DistID = TIMS.Get_DistID(DistID)
        If sm.UserInfo.DistID <> "000" Then
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        Else
            DistID.Enabled = True
        End If
        StudStatus.SelectedIndex = 0 '第1次進來選1
        Dim strWHERE As String = "TPLANID IN ('06','07','70')"
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", strWHERE)
        Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        'btn301_search.Attributes("onclick") = "return CheckSearch();"

        '無法自行選擇。年度及計畫。
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            TPlanID.Enabled = False
            ddlYear.Enabled = False
        End If
    End Sub

    '查詢鈕 [SQL]
    Sub Search301(ByVal sType As String)
        '#Region "查詢"
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""

        Select Case StudStatus.SelectedValue
            Case Cst_開訓人數, Cst_結訓人數, Cst_離訓人數 '班級學員 stud_studentinfo Class_StudentsOfClass Cst_就業人數, Cst_在職者,
                '(學員)
                sql = "" & vbCrLf
                sql &= " SELECT ss.Sex" & vbCrLf  '性別
                sql &= " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=14 then 1" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=15 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=19 then 2" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=20 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=24 then 3" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=25 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=29 then 4" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=30 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=34 then 5" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=35 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=39 then 6" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=40 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=44 then 7" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=45 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=49 then 8" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=50 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=54 then 9" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=55 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=59 then 10" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=60 And DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=64 then 11" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=65 then 12 end age" & vbCrLf
                sql &= " ,ss.DegreeID" & vbCrLf  '學歷
                sql &= " ,g1.CTID,g2.CTID CTID2" & vbCrLf  '縣市轄區('學員通訊郵遞區號)
                sql &= " ,ISNULL(ss.JobState,'0') JobState" & vbCrLf  '就職狀況'0:失業 1:在職

            Case Else '(會員報名) Stud_EnterTemp Stud_EnterType
                sql = "" & vbCrLf
                sql &= " SELECT st.Sex" & vbCrLf  '性別
                sql &= " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=14 then 1" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=15 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=19 then 2" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=20 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=24 then 3" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=25 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=29 then 4" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=30 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=34 then 5" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=35 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=39 then 6" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=40 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=44 then 7" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=45 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=49 then 8" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=50 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=54 then 9" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=55 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=59 then 10" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=60 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=64 then 11" & vbCrLf
                sql &= " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=65 then 12 end age" & vbCrLf
                sql &= " ,st.DegreeID" & vbCrLf  '學歷
                sql &= " ,g.CTID,g2.CTID CTID2" & vbCrLf  '縣市轄區('報名通訊地址郵遞區號)
                sql &= " ,'0' JobState" & vbCrLf  '就職狀況'0:失業 1:在職 

        End Select
        sql &= " ,oo.orgkind" & vbCrLf
        sql &= " ,ISNULL(ky.MergeID,b.MIdentityID) MIdentityID" & vbCrLf
        sql &= " ,case when iz.CTID is null then case when iz1.CTID is null then iz2.CTID else iz1.CTID end else iz.CTID end TaddCTID" & vbCrLf  '開班縣市 (TaddCTID)
        sql &= " ,case When ISNULL(a.THours,0) <=90 then 1" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=91 and ISNULL(a.THours,0) <=149 then 2" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=150 and ISNULL(a.THours,0) <=299 then 3" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=300 and ISNULL(a.THours,0) <=449 then 4" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=450 and ISNULL(a.THours,0) <=599 then 5" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=600 and ISNULL(a.THours,0) <=900 then 6" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=901 and ISNULL(a.THours,0) <=1200 then 7" & vbCrLf
        sql &= "  When ISNULL(a.THours,0) >=1201 then 8 end clshh" & vbCrLf
        'Cst_訓練職類大類/Cst_訓練職類中類/Cst_訓練職類小類
        sql &= " ,tt.BUSID,tt.JOBTMID,tt.TMID" & vbCrLf
        'sql &= " ,tt.BUSID,tt.BUSNAME" & vbCrLf
        'sql &= " ,tt.JOBTMID,tt.JOBNAME" & vbCrLf
        'sql &= " ,a.TMID,tt.TRAINNAME" & vbCrLf
        sql &= " FROM dbo.ID_PLAN d" & vbCrLf
        sql &= " JOIN dbo.PLAN_PLANINFO pp on d.planid=pp.planid" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo on oo.comidno=pp.comidno" & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO a on pp.planid=a.planid and pp.comidno=a.comidno and pp.seqno=a.seqno" & vbCrLf ' and a.rid = pp.rid
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=a.TMID" & vbCrLf '訓練職類 
        'sql &= " LEFT JOIN dbo.MVIEW_RELSHIP23 r3 on r3.RID3 =a.RID" & vbCrLf '補助地方政府單位。

        Select Case StudStatus.SelectedValue
            Case Cst_報名人數, Cst_甄試人數 'Cst_報名人數 (會員報名) Stud_EnterTemp Stud_EnterType
                sql &= " JOIN dbo.STUD_ENTERTYPE sy on sy.ocid1=a.ocid" & vbCrLf '報名班
                sql &= " JOIN dbo.STUD_ENTERTEMP st on st.SETID=sy.SETID" & vbCrLf '報名學員
                sql &= " JOIN dbo.STUD_SELRESULT sel on sel.SETID=sy.SETID AND sel.ENTERDATE=sy.ENTERDATE AND sel.SERNUM=sy.SERNUM AND sel.OCID=sy.OCID1" & vbCrLf '報名學員
                sql &= " LEFT JOIN dbo.ID_ZIP g ON st.ZipCode=g.ZipCode" & vbCrLf '報名通訊地址郵遞區號
                sql &= " LEFT JOIN dbo.STUD_STUDENTINFO ss on ss.idno=st.idno" & vbCrLf '基本學員
                sql &= " LEFT JOIN dbo.STUD_SUBDATA sa on ss.SID=sa.SID" & vbCrLf '基本學員(副)
                sql &= " LEFT JOIN dbo.CLASS_STUDENTSOFCLASS b on b.OCID=sy.ocid1 and b.sid=ss.sid and b.MAKESOCID is null" & vbCrLf

            Case Cst_開訓人數, Cst_結訓人數, Cst_離訓人數 'Cst_就業人數, Cst_在職者,
                '班級學員 stud_studentinfo Class_StudentsOfClass
                sql &= " JOIN dbo.CLASS_STUDENTSOFCLASS b on b.OCID=a.OCID and b.MAKESOCID is null" & vbCrLf '--b.OCID = sy.OCID1 and b.SETID = sy.SETID
                sql &= " JOIN dbo.STUD_STUDENTINFO ss on ss.sid=b.sid" & vbCrLf '基本學員
                sql &= " JOIN dbo.STUD_SUBDATA sa on sa.SID=b.SID" & vbCrLf '基本學員(副)

        End Select
        sql &= " LEFT JOIN dbo.KEY_IDENTITY ky ON ky.IdentityID=b.MIdentityID" & vbCrLf
        'sql &= " LEFT JOIN Stud_GetJobState3 sg on sg.cpoint=1 and b.socid = sg.socid" & vbCrLf
        sql &= " LEFT JOIN dbo.ID_Zip g2 ON sa.ZipCode2=g2.ZipCode" & vbCrLf '學員戶籍郵遞區號
        sql &= " LEFT JOIN dbo.ID_Zip g1 ON sa.ZipCode1=g1.ZipCode" & vbCrLf '學員通訊郵遞區號
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz ON iz.zipCode=a.TaddressZip" & vbCrLf '開班縣市郵遞區號
        sql &= " LEFT JOIN dbo.Plan_TrainPlace sp ON sp.PTID=pp.AddressSciPTID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz1 ON iz1.zipCode=sp.ZipCode" & vbCrLf '(學)開班縣市郵遞區號
        sql &= " left JOIN dbo.Plan_TrainPlace tp ON tp.PTID=pp.AddressTechPTID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz2 ON iz2.zipCode=tp.ZipCode" & vbCrLf '(術)開班縣市郵遞區號
        sql &= " WHERE a.ISSUCCESS='Y'" & vbCrLf
        Dim v_ddlYear As String = TIMS.GetListValue(ddlYear)
        If v_ddlYear <> "" Then
            sql &= " and d.Years=@Years" & vbCrLf
            parms.Add("Years", v_ddlYear)
        End If
        If DistID.SelectedIndex <> 0 Then
            sql &= " and d.DistID=@DistID" & vbCrLf
            parms.Add("DistID", DistID.SelectedValue)
        End If

        '無法自行選擇。年度及計畫。 '縣市政府承辦人登入時 為 True 其餘為 False
        If RIDValue.Value <> "" Then
            sql &= " AND a.RID = @RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        If STDate1.Text <> "" Then
            sql &= " and a.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", STDate1.Text)
        End If
        If STDate2.Text <> "" Then
            sql &= " and a.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", STDate2.Text)
        End If
        If FTDate1.Text <> "" Then
            sql &= " and a.FTDate >= @FTDate1" & vbCrLf
            parms.Add("FTDate1", FTDate1.Text)
        End If
        If FTDate2.Text <> "" Then
            sql &= " and a.FTDate <= @FTDate2" & vbCrLf
            parms.Add("FTDate2", FTDate2.Text)
        End If

        Dim itemplan As String = "" '計畫
        For Each objitem As ListItem In TPlanID.Items  '計畫
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next
        If itemplan <> "" Then sql &= " and d.TPlanID IN (" & itemplan & ")" & vbCrLf

        Select Case StudStatus.SelectedValue
            Case Cst_報名人數
                sql &= " AND a.NOTOPEN ='N'" & vbCrLf
            Case Cst_開訓人數
                sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL" & vbCrLf
            Case Cst_結訓人數
                If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL AND b.STUDSTATUS=5" & vbCrLf
                Else
                    sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL AND b.STUDSTATUS NOT IN (2,3) AND a.FTDate<GETDATE()" & vbCrLf
                End If
            Case Cst_離訓人數
                sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL AND b.STUDSTATUS IN (2,3)" & vbCrLf
            Case Cst_甄試人數
                sql &= " AND a.NOTOPEN ='N' AND sel.AppliedStatus='Y'" & vbCrLf
        End Select
        'Case Cst_就業人數
        '    sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL AND sg.CPoint=1 and sg.IsGetJob=1" & vbCrLf
        '    sql &= " AND b.STUDSTATUS NOT IN (2,3) AND a.FTDate < getdate() "
        'Case Cst_在職者
        '    sql &= " AND a.NOTOPEN ='N' AND b.SOCID IS NOT NULL and b.WORKSUPPIDENT='Y' AND d.TPlanID IN ('58','47')" & vbCrLf
        '    sql &= " AND b.STUDSTATUS NOT IN (2,3) AND a.FTDate < getdate() "

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        ViewState(cst_vsSqlString) = sql
        ViewState(cst_vs_parms1) = parms

        DataGroupTable.Visible = True
        Call CreateData(dt, XRoll.SelectedValue, YRoll.SelectedValue, YRoll.SelectedItem.Text, DataTable1, objconn)

        '#End Region
    End Sub

    '[共用] 與 SD_15_003_R.aspx CM_03_011_R.aspx
    Public Shared Sub CreateData(ByVal dt As DataTable,
        ByVal XRollValue As String, ByVal YRollValue As String, ByVal YRollText As String,
        ByVal DataTable1 As Table, ByVal tConn As SqlConnection)
        Dim MyRow As TableRow = Nothing
        Dim MyCell As TableCell = Nothing
        Dim Key_Degree As DataTable = TIMS.Get_KeyTable("Key_Degree", "DEGREETYPE='1'", tConn)
        'Dim s_where As String = "IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")"
        Dim sm As SessionModel = SessionModel.Instance()
        Dim s_where As String = ""
        s_where = "IDENTITYID IN (" & TIMS.Cst_Identity28_2019_11 & ")"
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then s_where = "IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")"

        Dim str_sql As String = ""
        Dim Key_Identity As DataTable = TIMS.Get_KeyTable("Key_Identity", s_where, tConn)
        Dim ID_City As DataTable = TIMS.Get_KeyTable("ID_City", "", tConn)
        Dim Key_Trade As DataTable = TIMS.Get_KeyTable("Key_Trade", "", tConn)
        Dim Key_ClassCatelog As DataTable = TIMS.Get_KeyTable("Key_ClassCatelog", "", tConn)
        Dim Key_OrgType As DataTable = TIMS.Get_KeyTable("Key_OrgType", "", tConn)
        Dim v_Thours As DataTable = TIMS.Get_KeyTable("V_THOURS2", "", tConn)  'Cst_訓練時數'2
        Dim dtYEARSOLD3 As DataTable = TIMS.Get_KeyTable("V_YEARSOLD3", "", tConn)  'YID,YNAME 年齡
        Dim t_TrainType1 As DataTable = TIMS.Get_KeyTable("KEY_TRAINTYPE", "BUSID NOT IN ('F','G','H')", tConn)  'Cst_訓練職類大類
        'Dim t_TrainType2 As DataTable = TIMS.Get_KeyTable("KEY_TRAINTYPE", "LEVELS='1' AND PARENT!=197 AND PARENT!=600", tConn)  'Cst_訓練職類中類
        'Dim t_TrainType3 As DataTable = TIMS.Get_KeyTable("KEY_TRAINTYPE", "LEVELS='2' AND LEN(TRAINID)=4", tConn)  'Cst_訓練職類小類    
        'str_sql = " SELECT TMID JOBTMID, JOBNAME FROM KEY_TRAINTYPE WHERE LEVELS='1' AND PARENT!=197 AND PARENT!=600" & vbCrLf
        'Cst_訓練職類中類
        str_sql = "" & vbCrLf
        str_sql &= " SELECT c.TMID JOBTMID,concat(p.BUSNAME,'-',c.JOBNAME) JOBNAME, c.TMKEY" & vbCrLf
        str_sql &= " FROM dbo.VIEW_TRAINTYPE3 c" & vbCrLf
        str_sql &= " JOIN dbo.KEY_TRAINTYPE p on p.TMID=c.PARENT" & vbCrLf
        str_sql &= " WHERE c.LEVELS='1' AND c.PARENT NOT IN (6,197,600)" & vbCrLf
        str_sql &= " ORDER BY c.TMKEY" & vbCrLf
        Dim t_TrainType2 As DataTable = TIMS.Get_KeyTable2(str_sql, tConn)
        'Cst_訓練職類小類
        str_sql = "" & vbCrLf
        str_sql &= " WITH WC1 AS ( SELECT c.TMID JOBTMID,concat(p.BUSNAME,'-',c.JOBNAME) JOBNAME" & vbCrLf
        str_sql &= " FROM dbo.KEY_TRAINTYPE c" & vbCrLf
        str_sql &= " JOIN dbo.KEY_TRAINTYPE p on p.TMID=c.PARENT" & vbCrLf
        str_sql &= " WHERE c.LEVELS='1' AND c.PARENT NOT IN (6,197,600) )" & vbCrLf

        str_sql &= " SELECT c.TMID,c.TRAINNAME,c.TMKEY" & vbCrLf
        str_sql &= " FROM dbo.VIEW_TRAINTYPE3 c" & vbCrLf
        str_sql &= " JOIN WC1 p on p.JOBTMID=c.PARENT" & vbCrLf
        str_sql &= " WHERE c.LEVELS='2' AND LEN(c.TRAINID)=4" & vbCrLf
        str_sql &= " ORDER BY c.TMKEY" & vbCrLf
        Dim t_TrainType3 As DataTable = TIMS.Get_KeyTable2(str_sql, tConn)

        Dim dtJOBSTATE As DataTable = TIMS.Get_KeyTable("V_JOBSTATE", "", tConn) '就職狀況'0:失業 1:在職 

        Dim iTotal As Integer = 0
        Select Case XRollValue
            Case Cst_性別
                MyRow = CreateRow(DataTable1)
                MyCell = CreateCell(MyRow, YRollText)
                MyCell.Width = Unit.Pixel(150)
                Select Case YRollValue
                    Case Cst_性別
                        CreateCell(MyRow, "人數")
                        CreateCell(MyRow, "比率")
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, New DataView(dt, "Sex='M'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F')", Nothing, DataViewRowState.CurrentRows).Count, 1)
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, New DataView(dt, "Sex='F'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F')", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Case Cst_年齡
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        '（1）配合65歲以上民眾得參與職業訓練，將X軸及Y軸年齡選項級距修改為：15歲以下；15~19；20~24； 25~29；30~34；35~39；40~44；45~49；50~54；55~64；65歲以上。
                        iTotal = New DataView(dt, "Sex IN ('M','F') and age IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                        For Each dr As DataRow In dtYEARSOLD3.Rows
                            Dim myText As String = Convert.ToString(dr("YNAME"))
                            Dim yid As Integer = Convert.ToString(dr("YID"))
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, myText)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Next

                    Case Cst_教育程度
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        iTotal = New DataView(dt, "Sex IN ('M','F') and DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count
                        For Each dr As DataRow In Key_Degree.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, iTotal)
                        Next
                    Case Cst_身分別
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In Key_Identity.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_受訓學員地理分布
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In ID_City.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CTName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_受訓學員地理分布2
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In ID_City.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CTName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and CTID2 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    'Case Cst_參訓單位類別
                    '    CreateCell(MyRow, "男")
                    '    CreateCell(MyRow, "女")
                    '    CreateCell(MyRow, "小計")
                    '    CreateCell(MyRow, "比率")
                    '    For Each dr As DataRow In Key_OrgType.Rows
                    '        MyRow = CreateRow(DataTable1)
                    '        CreateCell(MyRow, dr("Name").ToString)
                    '        CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                    '        CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                    '        Subtotal(MyRow)
                    '        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                    '    Next
                    Case Cst_開班縣市
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In ID_City.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CTName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and TaddCTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_訓練時數
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In v_Thours.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("CName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and clshh IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_訓練職類大類
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In t_TrainType1.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("busName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and busid IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_訓練職類中類
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In t_TrainType2.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("JOBNAME").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and JOBTMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_訓練職類小類
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In t_TrainType3.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("TRAINNAME").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and TMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

                        'Case Cst_就職狀況
                        '    CreateCell(MyRow, "男")
                        '    CreateCell(MyRow, "女")
                        '    CreateCell(MyRow, "小計")
                        '    CreateCell(MyRow, "比率")
                        '    For Each dr As DataRow In dtJOBSTATE.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("JName").ToString)
                        '        CreateCell(MyRow, New DataView(dt, "Sex='M' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        CreateCell(MyRow, New DataView(dt, "Sex='F' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                End Select
            Case Cst_年齡
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                        Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                        Dim myText As String = Convert.ToString(dr("YNAME"))
                        Dim yid As Integer = Convert.ToString(dr("YID"))
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, myText)
                        CreateCell(MyRow, New DataView(dt, "age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "age IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                        Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                        Dim myText As String = Convert.ToString(dr("YNAME"))
                        Dim yid As Integer = Convert.ToString(dr("YID"))
                        CreateCell(MyRow, myText)   '1~12 Name
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                                Dim yid As Integer = Convert.ToString(dr("YID"))
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and Sex IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                                Dim yid As Integer = Convert.ToString(dr("YID"))
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and Sex IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and CTID2 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                        '            Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                        '            Dim yid As Integer = Convert.ToString(drY("YID"))
                        '            CreateCell(MyRow, New DataView(dt, "age=" & yid & " and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and TaddCTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and clshh IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and busid IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    CreateCell(MyRow, New DataView(dt, "age=" & Convert.ToString(drY("YID")) & " and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and JOBTMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and TMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                            '            Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                            '            Dim yid As Integer = Convert.ToString(drY("YID"))
                            '            CreateCell(MyRow, New DataView(dt, "age=" & yid & " and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_教育程度
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_Degree.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_Degree.Rows
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_Degree.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In Key_Degree.Rows
                        '            CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In Key_Degree.Rows
                            '            CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next

                    End Select
                End If
            Case Cst_身分別
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In Key_Identity.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_Identity.Rows '20090923 andy edit
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_Identity.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_Identity.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In Key_Identity.Rows
                        '            CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and  orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr("CTID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr("Thour") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr("busid") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' AND JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' AND TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In Key_Identity.Rows
                            '            CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL  and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_受訓學員地理分布
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In ID_City.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CTName").ToString)
                        CreateCell(MyRow, New DataView(dt, "CTID='" & dr("CTID").ToString & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In ID_City.Rows
                        CreateCell(MyRow, dr("CTName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr("CTID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In ID_City.Rows
                        '            CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr("CTID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr("Thour") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr("busid") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' AND JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' AND TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In ID_City.Rows
                            '            CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_受訓學員地理分布2
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In ID_City.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CTName").ToString)
                        CreateCell(MyRow, New DataView(dt, "CTID2='" & dr("CTID").ToString & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In ID_City.Rows
                        CreateCell(MyRow, dr("CTName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr("DegreeID") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr1("CTID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr("CTID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In ID_City.Rows
                        '            CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr("CTID") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr("Thour") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr("busid") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr1("CTID") & "' AND JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr1("CTID") & "' AND TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In ID_City.Rows
                            '            CreateCell(MyRow, New DataView(dt, "JobState='" & dr("JobState") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            'Case Cst_參訓單位類別
            '    If YRollValue = XRollValue Then
            '        MyRow = CreateRow(DataTable1)
            '        MyCell = CreateCell(MyRow, YRollText)
            '        MyCell.Width = Unit.Pixel(150)
            '        CreateCell(MyRow, "人數")
            '        CreateCell(MyRow, "比率")
            '        For Each dr As DataRow In Key_OrgType.Rows
            '            MyRow = CreateRow(DataTable1)
            '            CreateCell(MyRow, dr("Name").ToString)
            '            CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
            '        Next
            '    Else
            '        MyRow = CreateRow(DataTable1)
            '        MyCell = CreateCell(MyRow, YRollText)
            '        MyCell.Width = Unit.Pixel(150)
            '        For Each dr As DataRow In Key_OrgType.Rows
            '            CreateCell(MyRow, dr("Name").ToString)
            '        Next
            '        CreateCell(MyRow, "小計")
            '        CreateCell(MyRow, "比率")
            '        Select Case YRollValue
            '            Case Cst_性別
            '                MyRow = CreateRow(DataTable1)
            '                CreateCell(MyRow, "男")
            '                For Each dr As DataRow In Key_OrgType.Rows
            '                    CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '                Subtotal(MyRow)
            '                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                MyRow = CreateRow(DataTable1)
            '                CreateCell(MyRow, "女")
            '                For Each dr As DataRow In Key_OrgType.Rows
            '                    CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '                Subtotal(MyRow)
            '                SubPercent(MyRow, New DataView(dt, " orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '            Case Cst_年齡
            '                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
            '                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
            '                    Dim myText As String = Convert.ToString(drY("YNAME"))
            '                    Dim yid As Integer = Convert.ToString(drY("YID"))
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, myText)
            '                    For Each dr As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "age=" & yid & " and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_教育程度
            '                For Each dr As DataRow In Key_Degree.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("Name").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and ='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_身分別
            '                For Each dr As DataRow In Key_Identity.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("Name").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_受訓學員地理分布
            '                For Each dr As DataRow In ID_City.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("CTName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_受訓學員地理分布2
            '                For Each dr As DataRow In ID_City.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("CTName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_參訓單位類別
            '            Case Cst_開班縣市
            '                For Each dr As DataRow In ID_City.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("CTName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_訓練時數
            '                For Each dr As DataRow In v_Thours.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("CName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_訓練職類大類
            '                For Each dr As DataRow In t_TrainType1.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("busName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '            Case Cst_就職狀況
            '                For Each dr As DataRow In dtJOBSTATE.Rows
            '                    MyRow = CreateRow(DataTable1)
            '                    CreateCell(MyRow, dr("JName").ToString)
            '                    For Each dr1 As DataRow In Key_OrgType.Rows
            '                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
            '                    Next
            '                    Subtotal(MyRow)
            '                    SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
            '                Next
            '        End Select
            '    End If
            Case Cst_開班縣市
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In ID_City.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CTName").ToString)
                        CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In ID_City.Rows
                        CreateCell(MyRow, dr("CTName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In ID_City.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In ID_City.Rows
                        '            CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In ID_City.Rows
                            '            CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_訓練時數
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In v_Thours.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("CName").ToString) 'Title
                        CreateCell(MyRow, New DataView(dt, "clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In v_Thours.Rows
                        CreateCell(MyRow, dr("CName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            Dim myText As String = ""
                            Dim FilterVal As String = ""
                            For i As Integer = 1 To 2
                                Select Case i
                                    Case 1
                                        myText = "男" : FilterVal = "M"
                                    Case 2
                                        myText = "女" : FilterVal = "F"
                                End Select
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In v_Thours.Rows
                        '            CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and JOBTMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In v_Thours.Rows
                            '            CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_訓練職類大類
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In t_TrainType1.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("busName").ToString) 'Title
                        CreateCell(MyRow, New DataView(dt, "busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    'Title
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In t_TrainType1.Rows
                        CreateCell(MyRow, dr("busName").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            Dim myText As String = ""
                            Dim FilterVal As String = ""
                            For i As Integer = 1 To 2
                                Select Case i
                                    Case 1
                                        myText = "男" : FilterVal = "M"
                                    Case 2
                                        myText = "女" : FilterVal = "F"
                                End Select
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and busid='" & dr1("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and busid='" & dr1("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_參訓單位類別
                        '    For Each dr As DataRow In Key_OrgType.Rows
                        '        MyRow = CreateRow(DataTable1)
                        '        CreateCell(MyRow, dr("Name").ToString)
                        '        For Each dr1 As DataRow In t_TrainType1.Rows
                        '            CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        '        Next
                        '        Subtotal(MyRow)
                        '        SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        '    Next
                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and TMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In t_TrainType1.Rows
                            '            CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If
            Case Cst_訓練職類中類
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In t_TrainType2.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("JOBNAME").ToString) 'Title
                        CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    'Title
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In t_TrainType2.Rows
                        CreateCell(MyRow, dr("JOBNAME").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            Dim myText As String = ""
                            Dim FilterVal As String = ""
                            For i As Integer = 1 To 2
                                Select Case i
                                    Case 1
                                        myText = "男" : FilterVal = "M"
                                    Case 2
                                        myText = "女" : FilterVal = "F"
                                End Select
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and JOBTMID='" & dr1("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and JOBTMID='" & dr1("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "JOBTMID='" & dr1("JOBTMID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JOBTMID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("BUSNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr("busid") & "' and JOBTMID='" & dr1("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        'Case Cst_訓練職類中類
                        Case Cst_訓練職類小類
                            For Each dr As DataRow In t_TrainType3.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("TRAINNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType2.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr("TMID") & "' and JOBTMID='" & dr1("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In t_TrainType1.Rows
                            '            CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If

            Case Cst_訓練職類小類
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")
                    For Each dr As DataRow In t_TrainType3.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("TRAINNAME").ToString) 'Title
                        CreateCell(MyRow, New DataView(dt, "TMID='" & dr("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    'Title
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In t_TrainType3.Rows
                        CreateCell(MyRow, dr("TRAINNAME").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")
                    Select Case YRollValue
                        Case Cst_性別
                            Dim myText As String = ""
                            Dim FilterVal As String = ""
                            For i As Integer = 1 To 2
                                Select Case i
                                    Case 1
                                        myText = "男" : FilterVal = "M"
                                    Case 2
                                        myText = "女" : FilterVal = "F"
                                End Select
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and TMID='" & dr1("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and TMID='" & dr1("TMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_身分別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("BUSNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_訓練職類中類
                            For Each dr As DataRow In t_TrainType2.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JOBNAME").ToString)
                                For Each dr1 As DataRow In t_TrainType3.Rows
                                    CreateCell(MyRow, New DataView(dt, "TMID='" & dr1("TMID") & "' and JOBTMID='" & dr("JOBTMID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TMID IS NOT NULL and JOBTMID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'Case Cst_訓練職類小類

                            'Case Cst_就職狀況
                            '    For Each dr As DataRow In dtJOBSTATE.Rows
                            '        MyRow = CreateRow(DataTable1)
                            '        CreateCell(MyRow, dr("JName").ToString)
                            '        For Each dr1 As DataRow In t_TrainType1.Rows
                            '            CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '        Next
                            '        Subtotal(MyRow)
                            '        SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                    End Select
                End If


                'Case Cst_就職狀況
                '    If YRollValue = XRollValue Then
                '        MyRow = CreateRow(DataTable1)
                '        MyCell = CreateCell(MyRow, YRollText)
                '        MyCell.Width = Unit.Pixel(150)
                '        CreateCell(MyRow, "人數")
                '        CreateCell(MyRow, "比率")
                '        For Each dr As DataRow In dtJOBSTATE.Rows
                '            MyRow = CreateRow(DataTable1)
                '            CreateCell(MyRow, dr("JName").ToString) 'Title
                '            CreateCell(MyRow, New DataView(dt, "JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '            SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                '        Next
                '    Else
                '        'Title
                '        MyRow = CreateRow(DataTable1)
                '        MyCell = CreateCell(MyRow, YRollText)
                '        MyCell.Width = Unit.Pixel(150)
                '        For Each dr As DataRow In dtJOBSTATE.Rows
                '            CreateCell(MyRow, dr("JName").ToString)
                '        Next
                '        CreateCell(MyRow, "小計")
                '        CreateCell(MyRow, "比率")
                '        Select Case YRollValue
                '            Case Cst_性別
                '                Dim myText As String = ""
                '                Dim FilterVal As String = ""
                '                For i As Integer = 1 To 2
                '                    Select Case i
                '                        Case 1
                '                            myText = "男" : FilterVal = "M"
                '                        Case 2
                '                            myText = "女" : FilterVal = "F"
                '                    End Select
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, myText)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and JobState='" & dr1("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_年齡
                '                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                '                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                '                    Dim myText As String = Convert.ToString(drY("YNAME"))
                '                    Dim yid As Integer = Convert.ToString(drY("YID"))
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, myText)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "age=" & yid & " and JobState='" & dr1("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_教育程度
                '                For Each dr As DataRow In Key_Degree.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("Name").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_身分別
                '                For Each dr As DataRow In Key_Identity.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("Name").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_受訓學員地理分布
                '                For Each dr As DataRow In ID_City.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("CTName").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_受訓學員地理分布2
                '                For Each dr As DataRow In ID_City.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("CTName").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            'Case Cst_參訓單位類別
                '            '    For Each dr As DataRow In Key_OrgType.Rows
                '            '        MyRow = CreateRow(DataTable1)
                '            '        CreateCell(MyRow, dr("Name").ToString)
                '            '        For Each dr1 As DataRow In dtJOBSTATE.Rows
                '            '            CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '            '        Next
                '            '        Subtotal(MyRow)
                '            '        SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '            '    Next
                '            Case Cst_開班縣市
                '                For Each dr As DataRow In ID_City.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("CTName").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_訓練時數
                '                For Each dr As DataRow In v_Thours.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("CName").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_訓練職類大類
                '                For Each dr As DataRow In t_TrainType1.Rows
                '                    MyRow = CreateRow(DataTable1)
                '                    CreateCell(MyRow, dr("busName").ToString)
                '                    For Each dr1 As DataRow In dtJOBSTATE.Rows
                '                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                '                    Next
                '                    Subtotal(MyRow)
                '                    SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                '                Next
                '            Case Cst_就職狀況
                '        End Select
                '    End If
        End Select

        '小計
        MyRow = CreateRow(DataTable1)
        Call CreateCell(MyRow, "小計")
        For i As Integer = 1 To DataTable1.Rows(0).Cells.Count - 1
            iTotal = 0
            For j As Integer = 1 To DataTable1.Rows.Count - 2
                If IsNumeric(DataTable1.Rows(j).Cells(i).Text) Then
                    iTotal += Val(DataTable1.Rows(j).Cells(i).Text)
                End If
            Next
            If i = DataTable1.Rows(0).Cells.Count - 1 Then
                If DataTable1.Rows(DataTable1.Rows.Count - 1).Cells(DataTable1.Rows(0).Cells.Count - 2).Text = 0 Then
                    Call CreateCell(MyRow, "0%")
                Else
                    Call CreateCell(MyRow, "100%")
                End If
            Else
                Call CreateCell(MyRow, iTotal)
            End If
        Next
    End Sub

    Public Shared Sub Subtotal(ByVal MyRow As TableRow)
        Dim Total As Integer = 0
        For i As Integer = 1 To MyRow.Cells.Count - 1
            Total += Int(MyRow.Cells(i).Text)
        Next
        CreateCell(MyRow, Total)
    End Sub

    Public Shared Sub SubPercent(ByVal MyRow As TableRow, ByVal RecordCount As Integer, Optional ByVal EndNum As Integer = 2)
        Dim Total As Integer = 0
        For i As Integer = 1 To MyRow.Cells.Count - EndNum
            Total += Int(MyRow.Cells(i).Text)
        Next
        If RecordCount = 0 Then
            CreateCell(MyRow, 0)
        Else
            CreateCell(MyRow, Math.Round(Total * 100 / RecordCount, 2) & "%")
        End If
    End Sub

    Public Shared Function CreateRow(ByVal DataTable1 As Table) As TableRow
        Dim MyRow As New TableRow
        DataTable1.Rows.Add(MyRow)

        Return MyRow
    End Function

    Public Shared Function CreateCell(ByRef MyRow As TableRow, ByVal MyText As String) As TableCell
        Dim MyCell As New TableCell
        MyRow.Cells.Add(MyCell)
        MyCell.Text = MyText
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center

        Return MyCell
    End Function

    '選訓練機構
    Private Sub Button3_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.ServerClick
        '#Region "選訓練機構"

        'Dim DistID1 As String = ""
        Dim v_DistID As String = TIMS.GetListValue(DistID)
        If DistID.SelectedIndex = 0 AndAlso v_DistID = "" Then
            Common.MessageBox(Me, "請先選擇轄區!")
            Exit Sub
        End If
        'DistID1 = DistID.SelectedValue

        Dim msg As String = ""
        Dim TPlanID1 As String = ""
        Dim N1 As Integer = 0
        N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        For j As Integer = 1 To TPlanID.Items.Count - 1
            If TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(TPlanID.Items(j).Value) '取得選項的值
                End If
                If N1 = 2 Then '如果計畫勾選選項的數目=2
                    msg += "只能選擇一個計畫!" & vbCrLf
                    TPlanID1 = ""
                    Exit For
                End If
            End If
        Next
        If N1 = 0 Then '如果計畫選項沒有選
            msg += "請選擇計畫!" & vbCrLf
        End If
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        v_DistID = TIMS.ClearSQM(v_DistID)
        TPlanID1 = TIMS.ClearSQM(TPlanID1)
        If v_DistID = "" OrElse TPlanID1 = "" Then
            Common.MessageBox(Me, "請選擇轄區與計畫!!")
            Exit Sub
        End If

        Dim strScript1 As String = ""
        strScript1 = "<script language=""javascript"">" + vbCrLf
        strScript1 &= String.Format("wopen('../../Common/MainOrg.aspx?DistID={0}&TPlanID={1}','查詢機構',400,400,1);", v_DistID, TPlanID1)
        strScript1 &= "</script>"
        Page.RegisterStartupScript("", strScript1)

        '#End Region
    End Sub

    '選擇不同的身分別，有須要改變的資訊。
    Private Sub StudStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudStatus.SelectedIndexChanged
        '#Region "選擇不同的身分別，有須要改變的資訊"

        DataGroupTable.Visible = False

        Select Case StudStatus.SelectedValue
            Case Cst_報名人數
                If Not YRoll.Items.FindByValue(Cst_受訓學員地理分布2) Is Nothing Then
                    YRoll.Items.Remove(YRoll.Items.FindByValue(Cst_受訓學員地理分布2)) '.RemoveAt(5)
                End If
                If Not XRoll.Items.FindByValue(Cst_受訓學員地理分布2) Is Nothing Then
                    XRoll.Items.Remove(XRoll.Items.FindByValue(Cst_受訓學員地理分布2)) '.RemoveAt(5)
                End If
            Case Else
                If YRoll.Items.FindByValue(Cst_受訓學員地理分布2) Is Nothing Then
                    YRoll.Items.Insert(CInt(Cst_受訓學員地理分布2) - 1, New ListItem("受訓學員(戶籍)地理分佈", Cst_受訓學員地理分布2))
                End If
                If XRoll.Items.FindByValue(Cst_受訓學員地理分布2) Is Nothing Then
                    XRoll.Items.Insert(CInt(Cst_受訓學員地理分布2) - 1, New ListItem("受訓學員(戶籍)地理分佈", Cst_受訓學員地理分布2))
                End If
        End Select

        'Select Case StudStatus.SelectedValue
        '    Case Cst_在職者
        '        '計畫限定
        '        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", "TPLANID IN ('58','47')", objconn)
        '    Case Else
        '        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", , objconn)
        '        Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
        'End Select
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", , objconn)
        Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)

        '無法自行選擇。年度及計畫。
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            TPlanID.Enabled = False
            ddlYear.Enabled = False
        End If
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '#End Region
    End Sub

    '匯出 明細資料
    Sub sExport301()
        '#Region "匯出 明細資料"

        'Const Cst_xlsFileName As String = "交叉分析統計表.xls"
        'Dim sFileName As String = ""
        'Dim sFileName1 As String = "交叉分析統計表"
        ''勞保勾稽查詢
        'sFileName = HttpUtility.UrlEncode(Cst_xlsFileName, System.Text.Encoding.UTF8)

        Dim sFileName1 As String = "交叉分析統計表"

        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        'mso-number-format:"0" 
        strSTYLE &= ("</style>")

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Dim strHTML As String = ""
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()

        '#End Region
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub btn301_Export_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn301_Export.Click
        Call Search301(cst_sql_2)
        Call sExport301()
    End Sub

    '列印
    Private Sub Btn301_print1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn301_print1.Click
        Session(cst_vsSqlString) = ViewState(cst_vsSqlString) '= Sql
        Session(cst_vs_parms1) = ViewState(cst_vs_parms1) '= parms

        Call Search301(cst_sql_1)

        'Page.RegisterStartupScript("open", "<script>wopen('../../CM/03/CM_03_011_R.aspx?X=" & XRoll.SelectedValue & "&Y=" & YRoll.SelectedValue & "&YText=" & Server.UrlEncode(YRoll.SelectedItem.Text) & "');</script>")
        Page.RegisterStartupScript("open", "<script>wopen('SD_15_024_R.aspx?X=" & XRoll.SelectedValue & "&Y=" & YRoll.SelectedValue & "&YText=" & Server.UrlEncode(YRoll.SelectedItem.Text) & "');</script>")

    End Sub

    Function CheckData301(ByRef Errmsg As String) As Boolean
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)

        If STDate1.Text <> "" AndAlso Not TIMS.IsDate1(STDate1.Text) Then
            Errmsg &= "開訓起始日期必須為正確日期格式" & vbCrLf
        End If
        If STDate2.Text <> "" AndAlso Not TIMS.IsDate1(STDate2.Text) Then
            Errmsg &= "開訓結束日期必須為正確日期格式" & vbCrLf
        End If
        '結訓期間
        If FTDate1.Text <> "" AndAlso Not TIMS.IsDate1(FTDate1.Text) Then
            Errmsg &= "結訓期間起始日期必須為正確日期格式" & vbCrLf
        End If
        If FTDate2.Text <> "" AndAlso Not TIMS.IsDate1(FTDate2.Text) Then
            Errmsg &= "結訓期間結束日期必須為正確日期格式" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        If StudStatus.SelectedValue = "" Then
            Errmsg &= "請選擇統計範圍" & vbCrLf
            Return False
        End If
        If XRoll.SelectedValue = "" OrElse YRoll.SelectedValue = "" Then
            Errmsg &= "請選擇XY軸分析項目" & vbCrLf
            Return False
        End If
        If Errmsg <> "" Then Return False
        Return True
    End Function

    '查詢鈕  
    Private Sub btn301_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn301_search.Click
        '#Region "查詢鈕"
        Dim Errmsg As String = ""
        Call CheckData301(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Try
            Call Search301(cst_sql_2)
        Catch ex As Exception
            Common.MessageBox(Page, "發生錯誤:" & vbCrLf & ex.Message)
            Dim strErrmsg As String = String.Concat("ex.Message: ", ex.Message, vbCrLf)
            'strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
            'strErrmsg &= ex.ToString & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
        End Try

        '#End Region
    End Sub
#End Region

End Class