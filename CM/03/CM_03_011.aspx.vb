Partial Class CM_03_011
    Inherits AuthBasePage

    Const Cst_性別 As String = "1"
    Const Cst_年齡 As String = "2"
    Const Cst_教育程度 As String = "3"
    Const Cst_身份別 As String = "4" 'Cst_特定對象
    Const Cst_受訓學員地理分布 As String = "5"  '(報名) 通訊地址
    Const Cst_受訓學員地理分布2 As String = "6" '(學員) 戶籍地址
    Const Cst_參訓單位類別 As String = "7"
    Const Cst_開班縣市 As String = "8"  '開班縣市 -ID_City -TaddCTID='" & dr("CTID") -CTName
    Const Cst_訓練時數 As String = "9" '訓練時數
    'Dim v_Thours As DataTable = TIMS.Get_KeyTable2("v_Thours", "", tConn) 
    '1.149時(含)以下
    '2.150~299時
    '3.300~449時
    '4.450~599時
    '5.600時以上
    'v_Thours2
    '1.90小時以下
    '2.91-149小時
    '3.150-299小時
    '4.300-449小時
    '5.450-599小時
    '6.600-900小時
    '7.901-1200小時
    '8.1201小時以上

    '	<asp@ListItem Value="21">訓練職類(大類)</asp@ListItem>
    '訓練職類(大類
    Const Cst_訓練職類大類 As String = "21"
    Const Cst_就職狀況 As String = "22" '就職狀況
    Const Cst_報名人數 As String = "11"
    Const Cst_開訓人數 As String = "12"
    Const Cst_結訓人數 As String = "13"
    Const Cst_就業人數 As String = "14"
    Const Cst_在職者 As String = "15" '在職者(托育及照服員計畫)

    'dtYEARSOLD3 -- select * from V_YEARSOLD3
    Const cst_ageInStr As String = "1,2,3,4,5,7,8,9,10,11,12"

    Const cst_sql_1 As String = "sql_1" '只要組合sql 
    Const cst_sql_2 As String = "sql_2" '組合sql，要產生查詢
    Const cst_vsSqlString As String = "SqlString"
    'Session(cst_vsSqlString) ViewState(cst_vsSqlString) ViewState("SqlString")

    'https://cm.turbotech.com.tw/browse/TIMS-2223
    '首頁>>訓練需求管理>>統計分析>>交叉分析統計表
    '本功能會開放給縣市政府承辦人使用,修改說明如下:
    '1.當縣市政府承辦人登入時,請依登入計畫,鎖定查詢條件"計畫範圍"為勾選於登入計畫,其他的計畫不提供勾選.
    Dim flag_Login1 As Boolean = False '縣市政府承辦人登入時 為 True 其餘為 False

    '2.查詢條件"訓練機構",縣市政府承辦人點選時,直接鎖定登入之年度計畫,並只可選該縣市政府或其底下之訓練單位,若訓練機構選縣市政府,則統計其轄下所有訓練單位之資料.

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        '縣市政府承辦人登入時 為 True 其餘為 False
        flag_Login1 = TIMS.Chk_LoginUserType1(Me)

        If Not IsPostBack Then
            DataGroupTable.Visible = False
            FTDate2.Text = Now.Date

            ddlYear = TIMS.GetSyear(ddlYear)
            Common.SetListItem(ddlYear, sm.UserInfo.Years)
            'ddlYear.SelectedValue = Now.ToString("yyyy")
            DistID = TIMS.Get_DistID(DistID)
            If sm.UserInfo.DistID <> "000" Then
                DistID.SelectedValue = sm.UserInfo.DistID
                DistID.Enabled = False
            Else
                DistID.Enabled = True
            End If
            StudStatus.SelectedIndex = 0 '第1次進來選1
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
        End If

        '無法自行選擇。年度及計畫。
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            TPlanID.Enabled = False
            ddlYear.Enabled = False
        End If

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        Button1.Attributes("onclick") = "return CheckSearch();"
        'Button3.Attributes("onclick") = "OpenOrg();"
    End Sub

    '查詢鈕  
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'If TIMS.sUtl_ChkTest() Then
        '    Call Search1()
        '    Exit Sub
        'End If

        '正式環境使用try Catch
        Try
            Call Search1(cst_sql_2)
        Catch ex As Exception
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)

            Dim strErrmsg As String = ""
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

    End Sub

    '查詢鈕 [SQL]
    Sub Search1(ByVal sType As String)
        'sType cst_sql_1:只要組合sql 
        'sType cst_sql_2:組合sql，要產生查詢
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sql As String = ""
        Select Case StudStatus.SelectedValue
            Case Cst_開訓人數, Cst_結訓人數, Cst_就業人數, Cst_在職者 '班級學員 stud_studentinfo Class_StudentsOfClass
                '(學員)
                '性別
                sql = "" & vbCrLf
                sql += " SELECT ss.Sex" & vbCrLf
                ''年齡
                'sql += " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <15 then 1 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=15 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=24 then 2 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=25 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=34 then 3 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=35 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=44 then 4 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=45 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=54 then 5 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=55 then 6 end  age " & vbCrLf
                '年齡3
                sql += " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=14 then 1 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=15 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=19 then 2 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=20 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=24 then 3 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=25 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=29 then 4 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=30 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=34 then 5 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=35 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=39 then 6 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=40 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=44 then 7 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=45 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=49 then 8 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=50 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=54 then 9 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=55 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=59 then 10 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=60 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) <=64 then 11 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, ss.Birthday) >=65 then 12 end age " & vbCrLf
                '學歷
                sql += " ,ss.DegreeID" & vbCrLf
                '縣市轄區('學員通訊郵遞區號)
                sql += " ,g1.CTID,g2.CTID CTID2" & vbCrLf
                '就職狀況'0:失業 1:在職 
                sql += " ,dbo.NVL(ss.JobState,'0') JobState" & vbCrLf

            Case Else '(會員報名) Stud_EnterTemp Stud_EnterType
                '性別
                sql = "" & vbCrLf
                sql += " SELECT st.Sex" & vbCrLf
                ''年齡
                'sql += " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <15 then 1 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=15 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=24 then 2 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=25 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=34 then 3 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=35 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=44 then 4 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=45 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=54 then 5 " & vbCrLf
                'sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=55 then 6 end age " & vbCrLf
                '年齡3
                sql += " ,case When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=14 then 1 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=15 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=19 then 2 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=20 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=24 then 3 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=25 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=29 then 4 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=30 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=34 then 5 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=35 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=39 then 6 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=40 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=44 then 7 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=45 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=49 then 8 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=50 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=54 then 9 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=55 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=59 then 10 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=60 and DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) <=64 then 11 " & vbCrLf
                sql += " When DATEPART(YEAR, a.STDate) - DATEPART(YEAR, st.Birthday) >=65 then 12 end age " & vbCrLf
                '學歷
                sql += " ,st.DegreeID" & vbCrLf
                '縣市轄區('報名通訊地址郵遞區號)
                sql += " ,g.CTID,g2.CTID CTID2" & vbCrLf
                '就職狀況'0:失業 1:在職 
                sql += " ,'0' JobState" & vbCrLf 'sql += " ,dbo.NVL(ss.JobState,'0') JobState" & vbCrLf

        End Select
        sql += " ,oo.orgkind" & vbCrLf
        sql += " ,dbo.NVL(ky.MergeID,b.MIdentityID) MIdentityID" & vbCrLf
        '開班縣市 (TaddCTID)
        sql += " ,case when iz.CTID is null then case when iz1.CTID is null then iz2.CTID else iz1.CTID end else iz.CTID end TaddCTID" & vbCrLf
        ''訓練時數 v_Thours
        'sql += " ,case When dbo.NVL(a.THours,0) <150 then 1 " & vbCrLf
        'sql += "  When dbo.NVL(a.THours,0) >=150 and dbo.NVL(a.THours,0) <300  then 2 " & vbCrLf
        'sql += "  When dbo.NVL(a.THours,0) >=300 and dbo.NVL(a.THours,0) <450  then 3 " & vbCrLf
        'sql += "  When dbo.NVL(a.THours,0) >=450 and dbo.NVL(a.THours,0) <600  then 4 " & vbCrLf
        'sql += "  When dbo.NVL(a.THours,0) >=600 then 5 end clshh " & vbCrLf
        '訓練時數 v_Thours2
        '	select '1' Thour, '90小時(含)以下' CName
        '	union select '2' Thour, '91-149小時' CName
        '	union select '3' Thour, '150-299小時' CName
        '	union select '4' Thour, '300-449小時' CName
        '	union select '5' Thour, '450-599小時' CName
        '	union select '6' Thour, '600-900小時' CName
        '	union select '7' Thour, '901-1200小時' CName
        '	union select '8' Thour, '1201小時以上' CName
        'v_Thours2
        sql += " ,case When dbo.NVL(a.THours,0) <=90 then 1 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=91 and dbo.NVL(a.THours,0) <=149 then 2 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=150 and dbo.NVL(a.THours,0) <=299 then 3 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=300 and dbo.NVL(a.THours,0) <=449 then 4 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=450 and dbo.NVL(a.THours,0) <=599 then 5 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=600 and dbo.NVL(a.THours,0) <=900 then 6 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=901 and dbo.NVL(a.THours,0) <=1200 then 7 " & vbCrLf
        sql += "  When dbo.NVL(a.THours,0) >=1201 then 8 end clshh " & vbCrLf
        'Cst_訓練職類大類
        sql += " ,tt.busid" & vbCrLf
        sql += " ,tt.busname" & vbCrLf
        sql += " FROM ID_Plan d" & vbCrLf
        sql += " JOIN plan_planinfo pp on d.planid = pp.planid " & vbCrLf
        sql += " JOIN Org_OrgInfo oo on oo.comidno = pp.comidno" & vbCrLf
        sql += " JOIN Class_ClassInfo a on pp.planid = a.planid and pp.comidno = a.comidno and pp.seqno=a.seqno " & vbCrLf ' and a.rid = pp.rid
        sql += " JOIN VIEW_TRAINTYPE tt on tt.TMID =a.TMID" & vbCrLf '訓練職類大類
        sql += " LEFT JOIN MVIEW_RELSHIP23 r3 on r3.RID3 =a.RID" & vbCrLf '補助地方政府單位。

        Select Case StudStatus.SelectedValue
            Case Cst_開訓人數, Cst_結訓人數, Cst_就業人數, Cst_在職者
                '班級學員 stud_studentinfo Class_StudentsOfClass
                sql += " JOIN Class_StudentsOfClass b on b.OCID=a.OCID and b.MAKESOCID is null" & vbCrLf '--b.OCID = sy.OCID1 and b.SETID = sy.SETID
                sql += " JOIN stud_studentinfo ss on ss.sid=b.sid " & vbCrLf '基本學員
                sql += " JOIN Stud_SubData sa on sa.SID = b.SID " & vbCrLf '基本學員(副)

            Case Else '(會員報名) Stud_EnterTemp Stud_EnterType
                sql += " JOIN Stud_EnterType sy on sy.ocid1 =a.ocid " & vbCrLf '報名班
                sql += " JOIN Stud_EnterTemp st on st.SETID =sy.SETID " & vbCrLf '報名學員
                sql += " LEFT JOIN ID_Zip g  ON st.ZipCode=g.ZipCode " & vbCrLf '報名通訊地址郵遞區號

                sql += " LEFT JOIN STUD_STUDENTINFO ss on ss.idno=st.idno " & vbCrLf '基本學員
                sql += " LEFT JOIN Stud_SubData sa on ss.SID = sa.SID " & vbCrLf '基本學員(副)
                sql += " LEFT JOIN Class_StudentsOfClass b on b.OCID=sy.ocid1 and b.sid=ss.sid and b.MAKESOCID is null" & vbCrLf '--b.OCID = sy.OCID1 and b.SETID = sy.SETID
        End Select
        sql += " LEFT JOIN Key_Identity ky  on ky.IdentityID =b.MIdentityID " & vbCrLf
        sql += " LEFT JOIN Stud_GetJobState3 sg on sg.cpoint=1 and b.socid = sg.socid" & vbCrLf

        sql += " LEFT JOIN ID_Zip g2  ON sa.ZipCode2=g2.ZipCode " & vbCrLf '學員戶籍郵遞區號
        sql += " LEFT JOIN ID_Zip g1  ON sa.ZipCode1=g1.ZipCode " & vbCrLf '學員通訊郵遞區號
        sql += " LEFT JOIN VIEW_ZIPNAME iz  on iz.zipCode=a.TaddressZip " & vbCrLf '開班縣市郵遞區號
        'sql += "  /* 產投上課地址學科場地代碼 */" & vbCrLf
        sql += " left join Plan_TrainPlace sp  on sp.PTID=pp.AddressSciPTID " & vbCrLf
        sql += " LEFT JOIN VIEW_ZIPNAME iz1  on iz1.zipCode=sp.ZipCode" & vbCrLf '(學)開班縣市郵遞區號
        'sql += "  /* 產投上課地址術科場地代碼 */" & vbCrLf
        sql += " left JOIN Plan_TrainPlace tp  on tp.PTID=pp.AddressTechPTID " & vbCrLf
        sql += " LEFT JOIN VIEW_ZIPNAME iz2  on iz2.zipCode=tp.ZipCode" & vbCrLf '(術)開班縣市郵遞區號

        sql += " WHERE 1=1" & vbCrLf
        'sql += " and a.NOTOPEN ='N'" & vbCrLf
        sql += " and a.ISSUCCESS='Y'" & vbCrLf
        If ddlYear.SelectedValue <> "" Then
            sql += " and d.Years ='" & ddlYear.SelectedValue & "'" & vbCrLf
        End If
        If DistID.SelectedIndex <> 0 Then
            sql += " and d.DistID='" & DistID.SelectedValue & "'" & vbCrLf
        End If

        '無法自行選擇。年度及計畫。
        '縣市政府承辦人登入時 為 True 其餘為 False
        If RIDValue.Value <> "" Then
            sql += " and (1!=1" & vbCrLf
            sql += " OR a.RID='" & RIDValue.Value & "'" & vbCrLf
            sql += " OR r3.RID2='" & RIDValue.Value & "'" & vbCrLf '(地方政府) 可以依據上層機構。
            sql += " )" & vbCrLf
        End If

        'If PlanID.Value <> "" Then
        '    SearchStr += " and a.PlanID='" & PlanID.Value & "'" & vbCrLf
        'End If
        If STDate1.Text <> "" Then
            sql += " and a.STDate>= " & TIMS.to_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql += " and a.STDate<= " & TIMS.to_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            sql += " and a.FTDate>= " & TIMS.to_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += " and a.FTDate<= " & TIMS.to_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
        End If

        Dim itemplan As String = "" '計畫
        For Each objitem As ListItem In Me.TPlanID.Items  '計畫
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next
        If itemplan <> "" Then
            sql += " and pp.TPlanID IN (" & itemplan & ")" & vbCrLf
        End If

        Select Case StudStatus.SelectedValue
            Case Cst_報名人數
                sql += " and a.NOTOPEN ='N' " & vbCrLf
            Case Cst_開訓人數
                sql += " and a.NOTOPEN ='N' and b.SOCID IS NOT NULL " & vbCrLf
            Case Cst_結訓人數
                sql += " and a.NOTOPEN ='N' and b.SOCID IS NOT NULL and b.STUDSTATUS NOT IN (2,3) and a.FTDate < getdate()" & vbCrLf
            Case Cst_就業人數
                sql += " and a.NOTOPEN ='N' and b.SOCID IS NOT NULL and sg.CPoint=1 and sg.IsGetJob=1 " & vbCrLf
                sql += " and b.STUDSTATUS NOT IN (2,3) and a.FTDate < getdate()"
            Case Cst_在職者
                'SELECT * FROM KEY_PLAN WHERE TPlanID IN ('58','47')
                '--58 補助辦理托育人員職業訓練
                '--47 補助辦理照顧服務員職業訓練
                sql += " and a.NOTOPEN ='N' and b.SOCID IS NOT NULL and b.WORKSUPPIDENT='Y' and d.TPlanID IN ('58','47') " & vbCrLf
                sql += " and b.STUDSTATUS NOT IN (2,3) and a.FTDate < getdate()"
        End Select
        'sql += SearchStr & "" & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn)
        ViewState(cst_vsSqlString) = sql

        DataGroupTable.Visible = True
        Call CreateData(dt, XRoll.SelectedValue, YRoll.SelectedValue, YRoll.SelectedItem.Text, DataTable1, objconn)
    End Sub

#Region "Public1"
    '[共用] 與 SD_15_003_R.aspx CM_03_011_R.aspx
    Public Shared Sub CreateData(ByVal dt As DataTable, _
        ByVal XRollValue As String, ByVal YRollValue As String, ByVal YRollText As String, _
        ByVal DataTable1 As Table, ByVal tConn As SqlConnection)

        Dim MyRow As TableRow = Nothing
        Dim MyCell As TableCell = Nothing
        Dim Key_Degree As DataTable = TIMS.Get_KeyTable("Key_Degree", "", tConn)
        Dim Key_Identity As DataTable = TIMS.Get_KeyTable("Key_Identity", "", tConn)
        Dim ID_City As DataTable = TIMS.Get_KeyTable("ID_City", "", tConn)
        Dim Key_Trade As DataTable = TIMS.Get_KeyTable("Key_Trade", "", tConn)
        Dim Key_ClassCatelog As DataTable = TIMS.Get_KeyTable("Key_ClassCatelog", "", tConn)
        Dim Key_OrgType As DataTable = TIMS.Get_KeyTable("Key_OrgType", "", tConn)
        'Dim v_Thours As DataTable = TIMS.Get_KeyTable2("v_Thours", "", tConn) 'Cst_訓練時數
        Dim v_Thours As DataTable = TIMS.Get_KeyTable("V_THOURS2", "", tConn)  'Cst_訓練時數'2
        Dim dtYEARSOLD3 As DataTable = TIMS.Get_KeyTable("V_YEARSOLD3", "", tConn)  'YID,YNAME 年齡
        Dim t_TrainType1 As DataTable = TIMS.Get_KeyTable("KEY_TRAINTYPE", "busid is not null and busid !='G'", tConn)  'Cst_訓練職類大類
        '就職狀況'0:失業 1:在職 
        Dim dtJOBSTATE As DataTable = TIMS.Get_KeyTable("V_JOBSTATE", "", tConn)

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
                        For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                            Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                            Dim myText As String = Convert.ToString(dr("YNAME"))
                            Dim yid As Integer = Convert.ToString(dr("YID"))
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, myText)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and age IN (1,2,3,4,5,6)", Nothing, DataViewRowState.CurrentRows).Count)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and age IN (" & cst_ageInStr & ")", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

                    Case Cst_教育程度
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")

                        For Each dr As DataRow In Key_Degree.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and DegreeID IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next
                    Case Cst_身份別
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

                    Case Cst_參訓單位類別
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")

                        For Each dr As DataRow In Key_OrgType.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("Name").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

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

                    Case Cst_就職狀況
                        CreateCell(MyRow, "男")
                        CreateCell(MyRow, "女")
                        CreateCell(MyRow, "小計")
                        CreateCell(MyRow, "比率")
                        For Each dr As DataRow In dtJOBSTATE.Rows
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, dr("JName").ToString)
                            CreateCell(MyRow, New DataView(dt, "Sex='M' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            CreateCell(MyRow, New DataView(dt, "Sex='F' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "Sex IN ('M','F') and JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Next

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
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='M' and age=6", Nothing, DataViewRowState.CurrentRows).Count)
                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and Sex IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim dr As DataRow = dtYEARSOLD3.Rows(i)
                                Dim yid As Integer = Convert.ToString(dr("YID"))
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and age=" & yid, Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=1", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=2", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=3", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=4", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=5", Nothing, DataViewRowState.CurrentRows).Count)
                            'CreateCell(MyRow, New DataView(dt, "Sex='F' and age=6", Nothing, DataViewRowState.CurrentRows).Count)
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
                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                    Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                    Dim yid As Integer = Convert.ToString(drY("YID"))
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "age IS NOT NULL and JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
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

                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In Key_Degree.Rows
                                    CreateCell(MyRow, New DataView(dt, "DegreeID='" & dr1("DegreeID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "DegreeID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                    End Select
                End If

            Case Cst_身份別
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
                    'CreateCell(MyRow, Key_Identity.Select(sIdentityID_title2)(0)("Name").ToString) '負擔家計婦女併至獨立負擔家計者
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
                        Case Cst_身份別


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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and MIdentityID='" & dr1("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next

                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL and  orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In Key_Identity.Rows
                                    CreateCell(MyRow, New DataView(dt, "MIdentityID='" & dr1("IdentityID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "MIdentityID IS NOT NULL  and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
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
                        Case Cst_身份別
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
                                'For Each dr As DataRow In Key_OrgType.Rows
                                '    MyRow = CreateRow(DataTable1)
                                '    CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID2='" & dr("CTID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and CTID='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "CTID='" & dr1("CTID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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
                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr("JobState") & "' and CTID2='" & dr1("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "CTID2 IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                    End Select
                End If

            Case Cst_參訓單位類別

                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")

                    For Each dr As DataRow In Key_OrgType.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("Name").ToString)
                        CreateCell(MyRow, New DataView(dt, "orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In Key_OrgType.Rows
                        CreateCell(MyRow, dr("Name").ToString)
                    Next
                    CreateCell(MyRow, "小計")
                    CreateCell(MyRow, "比率")

                    Select Case YRollValue
                        Case Cst_性別
                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "男")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='M' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            MyRow = CreateRow(DataTable1)
                            CreateCell(MyRow, "女")
                            For Each dr As DataRow In Key_OrgType.Rows
                                CreateCell(MyRow, New DataView(dt, "Sex='F' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                            Subtotal(MyRow)
                            SubPercent(MyRow, New DataView(dt, " orgkind IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and ='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_身份別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_參訓單位類別

                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In Key_OrgType.Rows
                                    CreateCell(MyRow, New DataView(dt, "orgkind='" & dr1("OrgTypeID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "orgkind IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                    End Select
                End If

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

                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In ID_City.Rows
                                    CreateCell(MyRow, New DataView(dt, "TaddCTID='" & dr1("CTID") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "TaddCTID IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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
                    'Title
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
                            'MyRow = CreateRow(DataTable1)
                            'CreateCell(MyRow, "男")
                            'For Each dr1 As DataRow In v_Thours.Rows
                            '    CreateCell(MyRow, New DataView(dt, "Sex='M' and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            'MyRow = CreateRow(DataTable1)
                            'CreateCell(MyRow, "女")
                            'For Each dr1 As DataRow In v_Thours.Rows
                            '    CreateCell(MyRow, New DataView(dt, "Sex='F' and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
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

                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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
                            'x
                            'For Each dr As DataRow In ID_City.Rows
                            '    MyRow = CreateRow(DataTable1)
                            '    CreateCell(MyRow, dr("Name").ToString)
                            '    For Each dr1 As DataRow In v_Thours.Rows
                            '        CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            '    Next
                            '    Subtotal(MyRow)
                            '    SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next

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

                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In v_Thours.Rows
                                    CreateCell(MyRow, New DataView(dt, "clshh='" & dr1("Thour") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                    End Select
                End If

            Case Cst_訓練職類大類
                'x
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
                            'MyRow = CreateRow(DataTable1)
                            'CreateCell(MyRow, "男")
                            'For Each dr1 As DataRow In v_Thours.Rows
                            '    CreateCell(MyRow, New DataView(dt, "Sex='M' and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)

                            'MyRow = CreateRow(DataTable1)
                            'CreateCell(MyRow, "女")
                            'For Each dr1 As DataRow In v_Thours.Rows
                            '    CreateCell(MyRow, New DataView(dt, "Sex='F' and clshh='" & dr1("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                            'Next
                            'Subtotal(MyRow)
                            'SubPercent(MyRow, New DataView(dt, "clshh IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
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

                        Case Cst_身份別
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

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

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
                            'x
                        Case Cst_就職狀況
                            For Each dr As DataRow In dtJOBSTATE.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("JName").ToString)
                                For Each dr1 As DataRow In t_TrainType1.Rows
                                    CreateCell(MyRow, New DataView(dt, "busid='" & dr1("busid") & "' and JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "busid IS NOT NULL and JobState IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                    End Select
                End If

            Case Cst_就職狀況
                'x
                If YRollValue = XRollValue Then
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    CreateCell(MyRow, "人數")
                    CreateCell(MyRow, "比率")

                    For Each dr As DataRow In dtJOBSTATE.Rows
                        MyRow = CreateRow(DataTable1)
                        CreateCell(MyRow, dr("JName").ToString) 'Title
                        CreateCell(MyRow, New DataView(dt, "JobState='" & dr("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                        SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL", Nothing, DataViewRowState.CurrentRows).Count, 1)
                    Next
                Else
                    'Title
                    MyRow = CreateRow(DataTable1)
                    MyCell = CreateCell(MyRow, YRollText)
                    MyCell.Width = Unit.Pixel(150)
                    For Each dr As DataRow In dtJOBSTATE.Rows
                        CreateCell(MyRow, dr("JName").ToString)
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
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "Sex='" & FilterVal & "' and JobState='" & dr1("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and Sex IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next
                        Case Cst_年齡
                            For i As Integer = 0 To dtYEARSOLD3.Rows.Count - 1
                                Dim drY As DataRow = dtYEARSOLD3.Rows(i)
                                Dim myText As String = Convert.ToString(drY("YNAME"))
                                Dim yid As Integer = Convert.ToString(drY("YID"))
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, myText)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "age=" & yid & " and JobState='" & dr1("JobState") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and age IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_教育程度
                            For Each dr As DataRow In Key_Degree.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and DegreeID='" & dr("DegreeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and DegreeID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_身份別
                            For Each dr As DataRow In Key_Identity.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and MIdentityID='" & dr("IdentityID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and MIdentityID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_受訓學員地理分布
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and CTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and CTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_受訓學員地理分布2
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and CTID2='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and CTID2 IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_參訓單位類別
                            For Each dr As DataRow In Key_OrgType.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("Name").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and orgkind='" & dr("OrgTypeID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and orgkind IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_開班縣市
                            For Each dr As DataRow In ID_City.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CTName").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and TaddCTID='" & dr("CTID") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and TaddCTID IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_訓練時數
                            For Each dr As DataRow In v_Thours.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("CName").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and clshh='" & dr("Thour") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and clshh IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_訓練職類大類
                            For Each dr As DataRow In t_TrainType1.Rows
                                MyRow = CreateRow(DataTable1)
                                CreateCell(MyRow, dr("busName").ToString)
                                For Each dr1 As DataRow In dtJOBSTATE.Rows
                                    CreateCell(MyRow, New DataView(dt, "JobState='" & dr1("JobState") & "' and busid='" & dr("busid") & "'", Nothing, DataViewRowState.CurrentRows).Count)
                                Next
                                Subtotal(MyRow)
                                SubPercent(MyRow, New DataView(dt, "JobState IS NOT NULL and busid IS Not NULL", Nothing, DataViewRowState.CurrentRows).Count)
                            Next

                        Case Cst_就職狀況
                            'x

                    End Select
                End If

        End Select

        '小計
        MyRow = CreateRow(DataTable1)
        Call CreateCell(MyRow, "小計")
        For i As Integer = 1 To DataTable1.Rows(0).Cells.Count - 1
            iTotal = 0
            For j As Integer = 1 To DataTable1.Rows.Count - 2
                If IsNumeric(DataTable1.Rows(j).Cells(i).Text) Then
                    iTotal += DataTable1.Rows(j).Cells(i).Text
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

#End Region

    '選訓練機構
    Private Sub Button3_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.ServerClick
        Dim DistID1 As String = ""
        If DistID.SelectedIndex = 0 AndAlso DistID.SelectedValue = "" Then
            Common.MessageBox(Me, "請先選擇轄區!")
            Exit Sub
        End If

        DistID1 = DistID.SelectedValue

        Dim msg As String = ""
        Dim TPlanID1 As String = ""
        Dim N1 As Integer = 0
        N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        For j As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(Me.TPlanID.Items(j).Value) '取得選項的值
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

        If DistID1 = "" OrElse TPlanID1 = "" Then
            Common.MessageBox(Me, "請選擇轄區與計畫!!")
            Exit Sub
        End If

        Dim strScript1 As String = ""
        strScript1 = "<script language=""javascript"">" + vbCrLf
        strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "','查詢機構',400,400,1);"
        strScript1 += "</script>"
        Page.RegisterStartupScript("", strScript1)

    End Sub

    '選擇不同的身份別，有須要改變的資訊。
    Private Sub StudStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StudStatus.SelectedIndexChanged
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

        Select Case StudStatus.SelectedValue
            Case Cst_在職者
                '計畫限定
                TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", "TPLANID IN ('58','47')", objconn)
            Case Else
                TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", , objconn)
                Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
        End Select
        '無法自行選擇。年度及計畫。
        '縣市政府承辦人登入時 為 True 其餘為 False
        If flag_Login1 Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            TPlanID.Enabled = False
            ddlYear.Enabled = False
        End If
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

    End Sub

    '匯出 明細資料
    Sub sExport1()
        Const Cst_xlsFileName As String = "交叉分析統計表.xls"
        Dim sFileName As String = ""
        '勞保勾稽查詢
        sFileName = HttpUtility.UrlEncode(Cst_xlsFileName, System.Text.Encoding.UTF8)

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集

        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"

        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        'DataGrid1.AllowPaging = False '關閉分頁
        'DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Response.End()

        'DataGrid1.AllowPaging = True '開啟分頁
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Call Search1(cst_sql_2)
        Call sExport1()
    End Sub

    '列印
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Session("SqlString") = Me.ViewState(cst_vsSqlString)
        Session(cst_vsSqlString) = Me.ViewState(cst_vsSqlString)
        Call Search1(cst_sql_1)

        Page.RegisterStartupScript("open", "<script>wopen('CM_03_011_R.aspx?X=" & XRoll.SelectedValue & "&Y=" & YRoll.SelectedValue & "&YText=" & Server.UrlEncode(YRoll.SelectedItem.Text) & "');</script>")
        'Button1_Click(sender, e)

    End Sub

End Class
