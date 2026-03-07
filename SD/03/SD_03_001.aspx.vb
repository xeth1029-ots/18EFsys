Partial Class SD_03_001
    Inherits AuthBasePage

#Region "NOUSE"
    '#Region "(No Use)"

    'WITH WC1 AS (
    '  select cc.*
    '  FROM VIEW2 cc
    '  where 1=1
    '  and cc.tplanid not in ('28','54')
    '  and cc.years =convert(varchar, 1911+105)
    ')
    ',WC2 AS (
    'select cc.tplanid,cc.planname,cc.distid,cc.distname, 1 CNT 
    'FROM STUD_ENTERTEMP a
    'JOIN STUD_ENTERTYPE b ON a.SETID=b.SETID
    'JOIN STUD_SELRESULT c ON b.SETID=c.SETID and b.EnterDate=c.EnterDate and b.SerNum=c.SerNum AND b.OCID1=c.OCID
    'JOIN WC1 cc on cc.ocid =b.ocid1
    'WHERE 1=1
    'AND c.SelResultID IN ('01','02')
    ')
    'SELECT cc.planname 計畫名稱
    ',cc.distname 轄區
    ',COUNT(1) 報到總人數
    'FROM WC2  CC
    'GROUP BY  cc.planname ,cc.distname


    'SELECT * FROM STUD_SELRESULT 
    'AS OF TIMESTAMP TO_TIMESTAMP('2016-03-10 06:30:00', 'YYYY-MM-DD hh24:mi:ss')
    'where 1=1
    'and setid=1059912	and enterdate=to_date('2016-01-06','yyyy-MM-dd')	and sernum=1
    'begin
    'DELETE STUD_SELRESULT WHERE 1=1  AND SETID='1059912' AND ENTERDATE= to_date('2016-01-06 00:00:00','yyyy-MM-dd hh24:mi:ss') AND SERNUM='1' AND OCID='85983';
    'INSERT INTO STUD_SELRESULT (SETID,ENTERDATE,SERNUM,OCID,SUMOFGRAD,APPLIEDSTATUS,ADMISSION,SELRESULTID,TRNDTYPE,RID,PLANID,MODIFYACCT,MODIFYDATE,SELSORT,NOTES2 )  SELECT N'1059912',to_date('2016-01-06 00:00:00','yyyy-MM-dd hh24:mi:ss'),N'1',N'85983',N'58',N'N',N'N',N'02',NULL,N'G',N'3823',N'stc3422',to_date('2016-01-29 12:00:55','yyyy-MM-dd hh24:mi:ss'),NULL,NULL ;
    'end;

    'If Me.ViewState("TestStr")="AmuTest" Then '測試用
    'Stud_SelResult
    '取消報到
    'UPDATE STUD_SELRESULT SET APPLIEDSTATUS='N' where 1=1 AND SETID='1041773' AND ENTERDATE='2015/7/13 00:00:00' AND SERNUM='1'

    '#Region "(No Use)"

    'Private Sub Button5B_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5B.Click
    '    Dim ErrorMsg1 As String=""
    '    UpdateDbLStudentID(Me.ViewState("OCID"), ErrorMsg1)
    '    Button1_Click(sender, e)
    '    If ErrorMsg1 <> "" Then
    '        Common.MessageBox(Me, ErrorMsg1)
    '    End If
    'End Sub

    'Button2.Attributes("onclick")="return Button2_Send();"
    'Button2B_Click
    'Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    'End Sub

#End Region

    'Const cst_准考證號 As Integer=1
    'Const cst_姓名 As Integer=2
    Const cst_筆試成績 As Integer = 3
    Const cst_口試成績 As Integer = 4
    Const cst_總成績 As Integer = 5
    Const cst_報名日期 As Integer = 6
    Const cst_政府補助 As Integer = 7
    Const cst_名次 As Integer = 8
    Const cst_卷別 As Integer = 9
    Const cst_錄訓結果 As Integer = 10
    'Const cst_放棄報到原因 As Integer=11
    Dim fff As String = ""

    Const vs_SD03001_OCID1 As String = "SD03001_OCID1"
    Const vs_StDate As String = "_StDate" ' ViewState(vs_StDate)
    Const Cst_IDNODoubles As String = "IDNODoubles"
    Const Cst_接近最高額度6 As Integer = 60000 '40000
    Const Cst_接近最高額度9 As Integer = 90000 '40000
    'Const cst_EnterPathW As String="W" '就服站代碼
    Const cst_EnterPathNameW As String = "(就服單位協助報名)" '說明
    Const cst_errTPlanID28Msg1 As String = "請先「勾選」學員再按「完成報到」"

    Const cst_CMasterMsg1 As String = " 民眾 JJJ 具公司/商業負責人身分，非屬失業勞工，不得參加失業者職前訓練。"
    Const cst_CBLIDETMsg1 As String = "學員 @LAB3 不具失、待業身分，不得參加失業者職前訓練。"
    Const cst_CFIRE1Msg1 As String = "學員 @LAB3 為非自願離職者，無法執行報到。"
    Const cst_StdWMsg1 As String = "學員 @LAB3 已有學員資料，無法再次報到!!"
    Const cst_StdWMsg2 As String = "學員 @LAB3 ，姓名長度超過系統範圍，不可儲存!!"
    Const cst_StdWMsg3 As String = "學員 @LAB3 ，依處分日期及年限，仍在處分期間者，學員參訓，被處分者，不可儲存。"
    Const cst_AlertMsg1 As String = "此班級尚未錄取試算!!"

    Const cst_msg219 As String = "※ 姓名前標記「x-」表示民眾已註銷推介"
    Const cst_fgb219 As String = "<font color=blue>x-</font>"
    Const cst_Mgc219 As String = "民眾已註銷推介"

    Const cst_msg_014 As String = "開訓日後14日鎖定功能填寫!"

    'Dim cmd As New SqlCommand
    'Dim cmd2 As New SqlCommand
    'Stud_StudentInfo
    'SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & OCID1 & "'
    'Dim tmpS As String=""

    Dim dtITS As DataTable
    Dim dtStudClass As DataTable
    Dim oflag_Test As Boolean = False '(正式)'(測試用)
    '70: 區域產業據點職業訓練計畫(在職)
    Dim flag_can_ignore_control As Boolean = False
    'If flag_can_repeated_training Then flag_TrainITS=False '(略過)職前卡重參
    'flag_can_repeated_training=TIMS.SD_03_001_Can_Repeated_Training(OCIDValue1.Value)
    Dim flag_can_repeated_training As Boolean = False 'true: (略過)職前卡重參 

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    'Dim au As New cAUTH
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn) '--關閉連線
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objConn) '開啟連線

        '是否為超級使用者
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1)
        '檢查帳號的功能權限 Start
        'Dim flag_can_ignore_control As Boolean=False
        flag_can_ignore_control = False
        '70: 區域產業據點職業訓練計畫(在職)
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso DateDiff(DateInterval.Day, Now, CDate(TIMS.cst_TPlanID70_end_date_1)) >= 0 Then
            flag_can_ignore_control = True '忽視卡關-暫時
        End If

        '(略過)職前卡重參
        flag_can_repeated_training = TIMS.SD_03_001_Can_Repeated_Training(OCIDValue1.Value, objConn)

        hTPlanID2854.Value = "0"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then hTPlanID2854.Value = "1"

        'oflag_Test=TIMS.sUtl_ChkTest()  '測試檢測

        Hidtestflag.Value = If(oflag_Test, "Y", "") '測試環境啟用/'(非)測試環境啟用

        Hid_ignoreflag.Value = If(flag_can_ignore_control, "Y", "") '區域產業據點-啟用

        '#Region "(No Use)"
        'Button1.Enabled=True
        'If Not au.blnCanSech Then
        '    If sm.UserInfo.RoleID <> 0 OrElse sm.UserInfo.LID <> 0 Then
        '        Button1.Enabled=False
        '        TIMS.Tooltip(Button1, "無查詢功能權限", True)
        '    End If
        'End If

        'Button2.Enabled=True
        'If Not au.blnCanAdds Then
        '    If sm.UserInfo.RoleID <> 0 OrElse sm.UserInfo.LID <> 0 Then
        '        Button2.Enabled=False
        '        TIMS.Tooltip(Button2, "使用者權限不可完成報到(無新增權限)", True)
        '    End If
        'End If

        labmsg219.Text = cst_msg219

        If Not IsPostBack Then
            cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        hdatenow.Value = Common.FormatDate(Now) 'yyyy/MM/dd
        'DataGridTable1.Visible=False
        Button1.Attributes("onclick") = "javascript:return search()"  '查詢
        Button2.Attributes("onclick") = "return Button2_Send();"  '完成報到

        '#Region "(No Use)"

        'Button5.Attributes("onclick")="return Button5_Send(); "
        'If Session("Page_Error_MSG") <> "" Then
        '    Common.MessageBox(Me, Session("Page_Error_MSG"))
        '    Session("Page_Error_MSG")=""
        '    Exit Sub
        'End If

        '確認機構是否為黑名單
        Dim vsMsg2 As String = "" '確認機構是否為黑名單
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            Button2.Enabled = False
            TIMS.Tooltip(Button2, vsMsg2)

            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    Sub cCreate1()
        '勾稽檢查學員是否參加職前課程機制
        CheckBoxITS1.Checked = True
        '1. 【檢核】這個選項，改為反灰(改由程式控制) ，承辦人想要強制訓練單位開訓14天內都要勾稽未完成報到的人 '選項仍保留在介面上，讓訓練單位知道有這個功能
        CheckBoxITS1.Enabled = False

        Hid_CAN_IGNORE_RULE1_CNT.Value = ""
        ViewState(vs_SD03001_OCID1) = Nothing

        msg.Text = ""
        DataGridTable1.Visible = False
        'Button2.Style("display")="none"
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        HidLID.Value = sm.UserInfo.LID
        If RIDValue.Value <> sm.UserInfo.RID Then
            If sm.UserInfo.RoleID <> 0 OrElse sm.UserInfo.LID <> 0 Then
                '業務權限不相同,請開班機構完成
                Button2.Enabled = False
                TIMS.Tooltip(Button2, "使用者權限 不可完成報到(請開班機構完成)", True)
            End If
        End If

        '#Region "(No Use)"
        'If sm.UserInfo.TPlanID="36" AndAlso sm.UserInfo.DistID="002" Then '星光幫專用
        '    'Button2.Enabled=True
        '    If Not au.blnCanAdds Then
        '        Button2.Enabled=False
        '        TIMS.Tooltip(Button2, "使用者權限不可完成報到(無新增權限)", True)
        '    End If
        'End If

        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1, objConn)
        'txtEnterDate.Text=Common.FormatDate(Now)
        HidToday.Value = Common.FormatDate(Now)
    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objConn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objConn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    '查詢
    Sub Search1()
        'Button2.Style("display")="none" '完成報到(鈕)
        DataGridTable1.Visible = False
        msg.Text = "查無資料!!"

#Region "查詢時檢核準備項目"
        'Dim ugv_CheckBoxITS1 As String=TIMS.Utl_GetConfigVAL(objConn, "CheckBoxITS1", 1)

        '檢查班級是以成績或者以報名順序來排序
        Dim dt As DataTable
        Dim dr9 As DataRow = Nothing
        'Call TIMS.OpenDbConn(objConn)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim parms As New Hashtable From {{"OCID", TIMS.CINT1(OCIDValue1.Value)}}
        Dim sql As String = ""
        sql &= " SELECT cs.SOCID,CONVERT(VARCHAR, cs.Enterdate, 111) Enterdate" & vbCrLf
        sql &= " ,cs.OCID ,ss.IDNO" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.sid=cs.sid" & vbCrLf
        sql &= " WHERE cs.OCID=@OCID"
        'Sql += " ORDER BY 1" '取單筆資料無須排序
        dtStudClass = DbAccess.GetDataTable(sql, objConn, parms)

        Dim drOCID As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        If drOCID Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        'BtnCheckITS1.Enabled=False
        'If Convert.ToString(drOCID("INPUTOK14")).Equals("Y") Then BtnCheckITS1.Enabled=True
        'If Not BtnCheckITS1.Enabled Then TIMS.Tooltip(BtnCheckITS1, "開訓後超過14日，不再提供查詢!", True)
        'If oflag_Test Then '(測試用)
        '    If Not BtnCheckITS1.Enabled Then
        '        BtnCheckITS1.Enabled=True
        '        TIMS.Tooltip(BtnCheckITS1, "測試用 (開啟「檢核是否同時參加職前課程」)!!")
        '    End If
        'End If

        HidTNum.Value = $"{drOCID("TNum")}" 'dr.Item("TNum")
        ViewState(vs_StDate) = TIMS.Cdate3(drOCID("STDate")) '開訓日期
        hSTDate.Value = TIMS.Cdate3(drOCID("STDate"))
        hSTDate14.Value = TIMS.Cdate3(DateAdd(DateInterval.Day, 13, CDate(drOCID("STDate"))))

        Dim sTPlanID As String = TIMS.GetTPlanID(drOCID("PlanID"), objConn)

        Dim parms2 As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sql2 As String = " SELECT ISNULL(SUM(SumOfGrad),0) TOTAL FROM STUD_SELRESULT WHERE OCID=@OCID GROUP BY OCID "
        Dim drSEL As DataRow = DbAccess.GetOneRow(sql2, objConn, parms2)
        Button2.Style("display") = ""
        If drSEL Is Nothing Then
            Button2.Style("display") = "none"
            Common.MessageBox(Me, cst_AlertMsg1)
            Exit Sub
        End If

        Dim i_TotalGrade As Integer = 0 '沒有甄試成績 '甄試成績
        If Not IsDBNull(drSEL("TOTAL")) Then i_TotalGrade = drSEL("TOTAL") '有甄試成績

        DataGrid1.Columns(cst_卷別).Visible = False
        '沒有甄試成績
        DataGrid1.Columns(cst_筆試成績).Visible = False
        DataGrid1.Columns(cst_口試成績).Visible = False
        DataGrid1.Columns(cst_總成績).Visible = False
        DataGrid1.Columns(cst_報名日期).Visible = True
        If i_TotalGrade <> 0 Then
            '有甄試成績
            DataGrid1.Columns(cst_筆試成績).Visible = True
            DataGrid1.Columns(cst_口試成績).Visible = True
            DataGrid1.Columns(cst_總成績).Visible = True
            DataGrid1.Columns(cst_報名日期).Visible = False
        End If

        DataGrid1.Columns(cst_政府補助).Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then DataGrid1.Columns(cst_政府補助).Visible = True

        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objConn, stdBLACK2TPLANID)
        Dim sqlWSB As String = TIMS.Get_StdBlackWSB(Me, iStdBlackType, stdBLACK2TPLANID, 1)

#End Region

        '錄訓結果
        Dim v_SelResultID As String = TIMS.GetListValue(SelResultID)

        Dim parms3 As New Hashtable From {{"YEARS", sm.UserInfo.Years}, {"OCID", OCIDValue1.Value}}
        Dim sql3 As String = ""
        sql3 &= sqlWSB '(BlackWSB)'WSB
        sql3 &= " SELECT a.SETID ,a.Name" & vbCrLf
        'sql &= " ,(CASE WHEN g.IDNO IS NOT NULL THEN '" & cst_fgb219 & "' ELSE '' END) + a.Name Name" & vbCrLf
        '產業人才投資方案限定，開訓日期後才可完成報到
        'sql &= " ,case when dbo.TRUNC_DATETIME(getdate())-cc.STDATE >=0 then 'Y' else 'N' end CanSave1" & vbCrLf
        sql3 &= " ,CASE WHEN DATEDIFF(day, cc.STDATE, GETDATE())>=0 THEN 'Y' ELSE 'N' END CanSave1" & vbCrLf
        sql3 &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql3 &= " ,b.SerNum, b.RelEnterDate, b.ExamNo, b.WriteResult, b.OralResult, b.TotalResult" & vbCrLf
        sql3 &= " ,CASE WHEN ISNULL(b.TRNDMode,0)=3 THEN 0 ELSE ISNULL(b.TRNDMode,0) END TRNDMode" & vbCrLf
        sql3 &= " ,ISNULL(b.TRNDType,3) TRNDType ,b.TRNDMode TRNDMode2 ,b.NotExam ,c.Admission,c.SelResultID" & vbCrLf
        sql3 &= " ,c.AppliedStatus" & vbCrLf
        '錄訓結果
        sql3 &= " ,d.Name SelResultName ,e.LevelName ,c.OCID ,a.IDNO ,a.Birthday" & vbCrLf
        '放棄報到原因/放棄報到(原因)
        sql3 &= " ,b2.ESETID, b2.ESERNUM,b2.ABANDON,b2.ABANDONReason,b2.ABANDONACCT,b2.ABANDONDATE" & vbCrLf
        sql3 &= " ,CONVERT(varchar, cc.STDate, 111) STDate" & vbCrLf
        sql3 &= " ,CONVERT(varchar, cc.FTDate, 111) FTDate" & vbCrLf
        ',b.EnterPath" ,0 GovCost'計算政府經費另使用查詢',b.CFIRE1'非自願離職者',b.CFIRE1NS '取消提醒',b.CMASTER1'具公司/商業負責人身分',b.CMASTER1NS '已轉知',b.CMASTER1NT '已切結
        'GovCost計算政府經費另使用查詢'CFIRE1非自願離職者'CFIRE1NS取消提醒'CMASTER1具公司/商業負責人身分'CMASTER1NS已轉知'CMASTER1NT已切結
        sql3 &= " ,b.EnterPath,0 GovCost,b.CFIRE1,b.CFIRE1NS,b.CMASTER1,b.CMASTER1NS,b.CMASTER1NT" & vbCrLf
        '有多筆 學員處分資料
        sql3 &= " ,CASE WHEN WSB.IDNO IS NOT NULL THEN 'Y' ELSE 'N' END IsStdBlack" & vbCrLf 'WSB
        'sql &= " ,CASE WHEN g.IDNO IS NOT NULL THEN 'Y' END GOVKILL" & vbCrLf
        sql3 &= " FROM STUD_ENTERTEMP a" & vbCrLf
        sql3 &= " JOIN STUD_ENTERTYPE b ON a.SETID=b.SETID" & vbCrLf
        sql3 &= " JOIN STUD_SELRESULT c ON b.SETID=c.SETID AND b.EnterDate=c.EnterDate AND b.SerNum=c.SerNum AND b.OCID1=c.OCID"
        Select Case v_SelResultID 'SelResultID.SelectedIndex 'SelectedIndex 0:不區分、1:僅顯示正取、2:僅顯示備取
            Case "01"
                sql3 &= " AND c.SelResultID='01'" & vbCrLf
            Case "02"
                sql3 &= " AND c.SelResultID='02'" & vbCrLf
            Case Else
                sql3 &= " AND c.SelResultID IN ('01','02') " '只能是正取或備取
        End Select
        sql3 &= " LEFT JOIN STUD_ENTERTYPE2 b2 ON b2.ESETID=b.ESETID AND b2.ESERNUM=b.ESERNUM" & vbCrLf
        '20090410(Milor)加入只能查登入年度的年度限制。
        sql3 &= " JOIN CLASS_CLASSINFO cc ON b.OCID1=cc.OCID" & vbCrLf
        sql3 &= " JOIN ID_PLAN ip on ip.PlanID=cc.PlanID" & vbCrLf
        sql3 &= " LEFT JOIN KEY_SELRESULT d ON c.SelResultID=d.SelResultID" & vbCrLf
        sql3 &= " LEFT JOIN CLASS_CLASSLEVEL e ON b.CCLID=e.CCLID" & vbCrLf
        '有多筆 學員處分資料 (依系統日期1年內處分。)
        sql3 &= " LEFT JOIN WSB ON WSB.IDNO=a.IDNO" & vbCrLf
        'sql &= " LEFT JOIN WC1G g ON g.IDNO=a.IDNO" & vbCrLf
        sql3 &= " WHERE ip.YEARS =@YEARS AND cc.OCID=@OCID" & vbCrLf
        If i_TotalGrade = 0 Then
            sql3 &= " ORDER BY c.SelResultID, b.NotExam DESC, b.TRNDMode Desc, b.TRNDType, b.RelEnterDate, b.ExamNo" & vbCrLf
        Else
            sql3 &= " ORDER BY c.SelResultID, b.NotExam DESC, b.TRNDMode Desc, b.TRNDType, b.TotalResult DESC, b.ExamNo" & vbCrLf
        End If
        'dr9=DbAccess.GetOneRow(Sql)
        'SELECT * FROM STUD_ENTERTYPE b WHERE b.OCID1='47138'
        'SELECT * FROM STUD_ENTERTYPE2 b WHERE b.OCID1='47138'
        'SELECT a.* FROM Stud_EnterTemp a JOIN STUD_ENTERTYPE b on a.setid =b.setid WHERE b.OCID1='47138'
        'SELECT a.* FROM Stud_EnterTemp2 a JOIN STUD_ENTERTYPE2 b on a.esetid =b.esetid WHERE b.OCID1='47138'
        dt = DbAccess.GetDataTable(sql3, objConn, parms3)
        If dt.Rows.Count > 0 Then
            dr9 = dt.Rows(0)
            '確認內部學員學號是否有重覆
            If CheckDbLStudentID(OCIDValue1.Value) Then
                Dim ErrorMsg1 As String = ""
                Call UpdateDbLStudentID(OCIDValue1.Value, ErrorMsg1)
                If ErrorMsg1 <> "" Then
                    Common.MessageBox(Me, ErrorMsg1)
                    Exit Sub
                Else
                    dt = DbAccess.GetDataTable(sql3, objConn, parms3)
                End If
            End If
        End If

        ViewState(Cst_IDNODoubles) = ""
        Label1.Text = "" 'msg
        'Button2.Visible=False '完成報到(鈕)
        'Button2.Style("display")="none" '完成報到(鈕)
        DataGridTable1.Visible = False
        msg.Text = "查無資料!!"

        If TIMS.dtHaveDATA(dt) Then
            'Call TIMS.OpenDbConn(objConn)
            Dim tSql As String = "SELECT dbo.FN_GET_GOVCOST(@IDNO, @STDate) GovCost" & vbCrLf
            Dim oCmd As New SqlCommand(tSql, objConn)
            Dim v_IDNOs As String = ""
            For Each drSV As DataRow In dt.Rows
                Dim t_IDNO As String = TIMS.ClearSQM(drSV("IDNO"))
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = t_IDNO
                    .Parameters.Add("STDate", SqlDbType.VarChar).Value = Common.FormatDate(drSV("STDate"))
                    drSV("GovCost") = .ExecuteScalar()
                End With
                '(僅針對未完成報到的學員)
                If Not $"{drSV("AppliedStatus")}".Equals("Y") Then
                    v_IDNOs &= String.Concat(If(v_IDNOs <> "", ",", ""), t_IDNO)
                End If
            Next

            'Dim drS As DataRow=Nothing
            'For i As Integer=0 To dt.Rows.Count - 1
            '    drS=dt.Rows(i)
            'Next
            'Dim v_IDNOs As String=""
            'For Each drS As DataRow In dtStudClass.Rows
            '    If v_IDNOs <> "" Then v_IDNOs &= ","
            '    v_IDNOs &= Convert.ToString(drS("IDNO"))
            'Next

            CheckBoxITS1.Checked = True
            '1. 【檢核】這個選項，改為反灰(改由程式控制) ，承辦人想要強制訓練單位開訓14天內都要勾稽未完成報到的人 '選項仍保留在介面上， 讓訓練單位知道有這個功能
            CheckBoxITS1.Enabled = False
            'INPUTOK14 dtITS '開訓後超過14日，該「檢核是否同時參加職前課程」按鈕反灰，不再提供查詢
            dtITS = Nothing

            Dim flag_INPUTOK14 As Boolean = If($"{drOCID("INPUTOK14")}".Equals("Y"), True, False)

            If Not flag_INPUTOK14 Then TIMS.Tooltip(CheckBoxITS1, "開訓後超過14日，不再提供查詢!", True)

            If oflag_Test Then '(測試用)
                v_IDNOs = ""
                For Each drSV As DataRow In dt.Rows
                    Dim t_IDNO As String = TIMS.ClearSQM(drSV("IDNO"))
                    If v_IDNOs <> "" Then v_IDNOs &= ","
                    v_IDNOs &= t_IDNO
                Next
                If CheckBoxITS1.Checked Then
                    TIMS.Tooltip(CheckBoxITS1, "測試用 (開啟「檢核是否同時參加職前課程」)!!")
                    '勾稽檢查學員是否參加職前課程機制
                    'dtITS=TIMS.GetTrainingListS2(v_IDNOs)
                    dtITS = TIMS.GetTrainingListS3(sm, objConn, v_IDNOs)
                End If
            Else
                'If CheckBoxITS1.Checked Then 'End If
                CheckBoxITS1.Checked = flag_INPUTOK14 '開訓後超過14日，不再提供查詢!
                '勾稽檢查學員是否參加職前課程機制
                'If CheckBoxITS1.Checked Then dtITS=TIMS.GetTrainingListS2(v_IDNOs)
                If CheckBoxITS1.Checked Then dtITS = TIMS.GetTrainingListS3(sm, objConn, v_IDNOs)
            End If

            'Label1.Text="" 'msg
            'Button2.Visible=True '完成報到(鈕)
            Button2.Style("display") = "" '"inline" '完成報到(鈕)
            Me.DataGridTable1.Visible = True
            msg.Text = ""

            '#Region "(No Use)"

            ''確認內部學員學號是否有重覆
            'CheckDbLStudentID(OCIDValue1.Value)
            'objConn=DbAccess.GetConnection
            'Dim tSql As String=""
            'tSql="" & vbCrLf
            'tSql += " SELECT dbo.fn_GET_GOVCOST(@IDNO, @STDate)" & vbCrLf
            'cmd=New SqlCommand(tSql, objConn)

            'tSql="" & vbCrLf
            'tSql += " SELECT dbo.fn_GET_GOVCOST(@IDNO, @FTDate)" & vbCrLf
            'cmd2=New SqlCommand(tSql, objConn)

            Label1.Text = OCID1.Text & "，訓練人數" & HidTNum.Value & "人"
            'Button2.Visible=True
            'Me.DataGridTable1.Visible=True
            'DataGrid1.Visible=True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            'If objConn.State=ConnectionState.Open Then objConn.Close() '--關閉連線
            TIMS.Tooltip(Button2, "", True)

            '非系統管理者執行判斷。
            If Not (sm.UserInfo.RoleID = 0 AndAlso sm.UserInfo.LID = 0) Then
                If RIDValue.Value <> sm.UserInfo.RID Then
                    '業務權限不相同,請開班機構完成
                    Button2.Enabled = False '完成報到(鈕)
                    TIMS.Tooltip(Button2, "使用者權限不可完成報到(請開班機構完成)", True)
                End If
            End If
            '#Region "(No Use)"

            'If RIDValue.Value <> sm.UserInfo.RID Then
            '    Button2.Enabled=False '完成報到(鈕)
            '    TIMS.Tooltip(Button2, "使用者權限不可完成報到(請開班機構完成)", True)
            'End If
            'If sTPlanID="36" AndAlso sm.UserInfo.DistID="002" Then '星光幫專用
            '    If Not au.blnCanAdds Then
            '        Button2.Enabled=False '完成報到(鈕)
            '        TIMS.Tooltip(Button2, "使用者權限不可完成報到(無新增權限)")
            '    End If
            'End If
        End If

        If dr9 IsNot Nothing Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                If $"{dr9("CanSave1")}" <> "Y" Then
                    Button2.Enabled = False '完成報到(鈕)
                    TIMS.Tooltip(Button2, "產業人才投資方案限定，開訓日期後才可完成報到", True)
                End If
            End If
        End If

        Dim flag_can_ignore_control As Boolean = False
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso DateDiff(DateInterval.Day, Now, CDate(TIMS.cst_TPlanID70_end_date_1)) >= 0 Then
            '70: 區域產業據點職業訓練計畫(在職)
            flag_can_ignore_control = True '忽視卡關-暫時
        End If

        '未鎖定判斷
        If Button2.Enabled Then
            Dim iTmpDay As Integer = 14
            If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(Today)) > iTmpDay Then
                Button2.Enabled = False '鎖定
                Dim sTmpDay2 As String = "作業日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內完成)!"
                TIMS.Tooltip(Button2, sTmpDay2)
            End If
        End If

        '系統管理者
        If (sm.UserInfo.LID = "0" AndAlso sm.UserInfo.RoleID = "0") Then
            If Not Button2.Enabled Then
                Button2.Enabled = True
                TIMS.Tooltip(Button2, "系統管理者開放權限!!")
            End If
        End If

        If Not flgROLEIDx0xLIDx0 Then
            Dim flagInputOK14NG As Boolean = False '開訓日後14日鎖定功能填寫!
            Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objConn) '暫時權限Table
            'https://jira.turbotech.com.tw/browse/TIMSC-161
            If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_學員參訓, dtArc) Then
                '過了使用期限 True(不可使用)   False(可使用)
                If $"{drOCID("InputOK14")}" <> "Y" Then
                    '開訓日後14日鎖定功能填寫!
                    flagInputOK14NG = True
                End If
            End If
            'If flagInputOK14NG Then
            '    Common.MessageBox(Me, "開訓日後14日鎖定功能填寫!")
            '    Exit Sub
            'End If
            If Not Button2.Enabled AndAlso Not flagInputOK14NG Then
                Button2.Enabled = True
                TIMS.Tooltip(Button2, "取得已結訓班級使用權限-開訓日後14日鎖定功能填寫!!")
            End If
        End If

        If oflag_Test Then '(測試用)
            If Not Button2.Enabled Then
                Button2.Enabled = True
                TIMS.Tooltip(Button2, "測試用 (開啟「完成報到」)!!")
            End If
        End If
        If flag_can_ignore_control Then
            If Not Button2.Enabled Then
                Button2.Enabled = True
                TIMS.Tooltip(Button2, "區域產業暫時開啟 (開啟完成報到)!!")
            End If
        End If

        HidOCID1.Value = OCIDValue1.Value
        txtEnterDate.Text = ViewState(vs_StDate) '帶入開訓日期
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call Search1()
    End Sub

    '若內部學號有重複為true/反之false
    Function CheckDbLStudentID(ByVal OCID1 As String) As Boolean
        'Dim flag As Boolean=False
        If OCID1 = "" Then Return False
        'Dim dt As DataTable=Nothing

        Dim parms As New Hashtable From {{"OCID1", TIMS.CINT1(OCID1)}}
        Dim sql As String = ""
        sql &= " SELECT OCID, StudentID, COUNT(1) cnt" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
        sql &= " WHERE OCID=@OCID1" & vbCrLf
        sql &= " GROUP BY OCID, StudentID" & vbCrLf
        sql &= " HAVING COUNT(1) > 1" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, parms)
        'TIMS.Tooltip(Button5, "內部學號有重複需重整")
        Return (dt.Rows.Count > 0)
    End Function

    ''' <summary>更新有問題的學員號碼</summary>
    ''' <param name="OCID1"></param>
    ''' <param name="ErrorMsg1"></param>
    Sub UpdateDbLStudentID(ByVal OCID1 As String, ByRef ErrorMsg1 As String)
        OCID1 = TIMS.ClearSQM(OCID1)
        If OCID1 = "" Then Return

        Dim VS_STUDENTID As String = ""

        '將報名資料寫入學員資料檔內 Start
        Dim sql As String = ""
        sql &= " SELECT c.OCID, c.TMID" & vbCrLf
        sql &= " ,pp.ClassCate ,pp.IsBusiness" & vbCrLf
        sql &= " ,f.OrgKind ,e.DistID" & vbCrLf
        sql &= " ,c.Years" & vbCrLf
        sql &= " ,c.CyclType" & vbCrLf
        sql &= " ,ISNULL(d.ClassID2,d.ClassID) ClassID" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, c.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, c.FTDate, 111) FTDate" & vbCrLf
        sql &= " ,c.ClassCName ,c.TaddressZip ,c.TAddress" & vbCrLf
        sql &= " ,c.THours" & vbCrLf
        sql &= " ,f.OrgName" & vbCrLf
        sql &= " ,g.ContactName ,g.Phone ,f.ComCIDNO" & vbCrLf
        sql &= " ,g.MasterName ,c.LevelCount" & vbCrLf
        sql &= " ,ip.TPlanID" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO c" & vbCrLf
        sql &= " JOIN ID_CLASS d ON c.CLSID=d.CLSID AND c.OCID='" & OCID1 & "'" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PlanID=c.PlanID" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP e ON e.RID=c.RID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO f ON e.OrgID=f.OrgID" & vbCrLf
        sql &= " JOIN ORG_ORGPLANINFO g ON g.RSID=e.RSID" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PlanID=C.PlanID AND pp.ComIDNO=C.ComIDNO AND pp.SeqNo=C.SeqNo" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objConn) 'nothing
        If dr1 Is Nothing Then Return

        Dim sTPlanID As String = Convert.ToString(dr1("TPlanID"))
        '2020 提供期別可以不填寫，但學號仍加入01 維持一致性
        Dim v_CyclType As String = Convert.ToString(dr1("CyclType"))
        If v_CyclType = "" Then v_CyclType = TIMS.cst_Default_CyclType_forStudentID

        '學號增加的方式應該是要去除前面的固定長度字串，才做流水號處理，並將流水號由原本的2碼變3碼。
        'If sTPlanID="36" AndAlso dr1("ClassID").ToString.Length Then '因為當初轉入時班級ID過長所以重新製作
        '    '取"Y"+RID+流水號(2)  "YC" & Left(drA("RID"), 1) & Right(drA("PlanYear"), 2) & "0000001"
        '    Dim strClassID As String
        '    strClassID="Y" & Right(Left(dr1("ClassID").ToString, 3), 1) & Right(dr1("ClassID").ToString, 2)
        '    dr1("ClassID")=strClassID
        'End If

        '因為有資料庫交易問題所以提前呼叫
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
            Dim StudentID As String = ""
            StudentID = TIMS.Get_TPlanID28_StudentID(
                            dr1("Years").ToString,
                            dr1("DistID").ToString,
                            dr1("OrgKind").ToString,
                            dr1("IsBusiness").ToString,
                            dr1("ClassID").ToString,
                            v_CyclType,
                            dr1("ClassCate").ToString,
                            dr1("TMID").ToString, objConn)
            If StudentID Is Nothing Then
                ErrorMsg1 += "班級：" & dr1("ClassCName") & "參訓學員號有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                StudentID = ""
            Else
                If StudentID = "" Then
                    ErrorMsg1 += "班級：" & dr1("ClassCName") & "參訓學員號有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                    StudentID = ""
                End If
            End If
            VS_STUDENTID = StudentID
        End If

        'Dim dt As DataTable=Nothing
        Dim s_parms As New Hashtable From {{"OCID1", TIMS.CINT1(OCID1)}}
        'Dim sql As String=""
        sql = "" & vbCrLf
        sql &= " SELECT cs.OCID, cs.StudentID" & vbCrLf
        sql &= " ,COUNT(1) cnt" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid=cs.ocid" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid=cc.planid" & vbCrLf
        sql &= " WHERE cs.STUDSTATUS NOT IN (2,3) AND CS.OCID=@OCID1" & vbCrLf
        sql &= " GROUP BY cs.OCID, cs.StudentID" & vbCrLf
        sql &= " HAVING COUNT(1) > 1" & vbCrLf '有重複的學號
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objConn, s_parms)

        Do While dt.Rows.Count > 0
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim dr As DataRow = dt.Rows(i)
                Dim sql2 As String = $" SELECT MAX(SOCID) SOCID FROM CLASS_STUDENTSOFCLASS WHERE OCID={dr("OCID")} AND StudentID='{dr("StudentID")}'"
                Dim iMaxSOCID As Integer = DbAccess.ExecuteScalar(sql2, objConn)

                Dim StudentID As String = ""
                Dim iMaxStudentID As Integer = 0
                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                    '學號增加的方式應該是要去除前面的固定長度字串，才做流水號處理，並將流水號由原本的2碼變3碼。
                    Dim V_ClassID As String = $"{dr1("ClassID")}"
                    If dr1("Years").ToString.Length = 4 Then
                        VS_STUDENTID = String.Concat(Microsoft.VisualBasic.Right(dr1("Years").ToString, 2), "0", V_ClassID, v_CyclType)
                    ElseIf dr1("Years").ToString.Length = 2 Then
                        VS_STUDENTID = String.Concat(dr1("Years").ToString, "0", V_ClassID, v_CyclType)
                    Else
                        VS_STUDENTID = String.Concat(Microsoft.VisualBasic.Right(Now.Year.ToString, 2), "0", V_ClassID, v_CyclType)
                    End If

                    sql2 = $"SELECT ISNULL(MAX(CONVERT(NUMERIC, REPLACE(StudentID,'{VS_STUDENTID}',''))),0)+1 AS MaxNum FROM CLASS_STUDENTSOFCLASS WHERE OCID={dr1("OCID")} AND StudentID LIKE '{VS_STUDENTID}%'"
                    iMaxStudentID = DbAccess.ExecuteScalar(sql2, objConn)
                    StudentID = String.Concat(VS_STUDENTID, Format(iMaxStudentID, "00#"))
                Else
                    sql2 = "" & vbCrLf
                    sql2 += " SELECT ISNULL(MAX(CONVERT(NUMERIC, SUBSTRING(StudentID, LEN(StudentID)-1, 2))),0)+1 MaxNum" & vbCrLf
                    sql2 += " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
                    sql2 += $" WHERE OCID={dr1("OCID")}" & vbCrLf
                    iMaxStudentID = DbAccess.ExecuteScalar(sql2, objConn)
                    StudentID = VS_STUDENTID & Format(iMaxStudentID, "0#")
                End If
                sql2 = $" UPDATE CLASS_STUDENTSOFCLASS SET StudentID='{StudentID}' WHERE SOCID={iMaxSOCID} "
                DbAccess.ExecuteNonQuery(sql2, objConn)
            Next
            dt = DbAccess.GetDataTable(sql, objConn, s_parms)
        Loop

        'ErrorMsg1="test"
    End Sub

    '放棄報到原因/放棄報到(原因)
    Sub UPDATA_STUD_ENTERTYPE2_ABANDON(ByRef myCmdArg As String)
        Dim vESETID As String = TIMS.GetMyValue(myCmdArg, "ESETID")
        Dim vESERNUM As String = TIMS.GetMyValue(myCmdArg, "ESERNUM")
        'Dim vSETID As String=TIMS.GetMyValue(myCmdArg, "SETID")
        'Dim vEnterDate As String=TIMS.GetMyValue(myCmdArg, "EnterDate")
        'Dim vSerNum As String=TIMS.GetMyValue(myCmdArg, "SerNum")
        'Dim vIDNO As String=TIMS.GetMyValue(myCmdArg, "IDNO")
        Dim vOCID As String = TIMS.GetMyValue(myCmdArg, "OCID")
        'Hid_ABANDONReasonSub.Value=TIMS.ClearSQM2(Hid_ABANDONReasonSub.Value)
        Dim u_parms As New Hashtable
        'u_parms.Add("ABANDON", "Y")
        u_parms.Add("ABANDONReason", Hid_ABANDONReasonSub.Value)
        u_parms.Add("ABANDONACCT", sm.UserInfo.UserID)
        'u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("ESERNUM", TIMS.CINT1(vESERNUM))
        u_parms.Add("ESETID", TIMS.CINT1(vESETID))
        u_parms.Add("OCID1", TIMS.CINT1(vOCID))
        Dim u_sql As String = ""
        u_sql &= " UPDATE STUD_ENTERTYPE2" & vbCrLf
        u_sql &= " SET ABANDON='Y'" & vbCrLf
        u_sql &= " ,ABANDONReason=@ABANDONReason" & vbCrLf
        u_sql &= " ,ABANDONACCT=@ABANDONACCT" & vbCrLf
        u_sql &= " ,ABANDONDATE=GETDATE()" & vbCrLf
        'u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        'u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE ESERNUM=@ESERNUM" & vbCrLf
        u_sql &= " AND ESETID=@ESETID" & vbCrLf
        u_sql &= " AND OCID1=@OCID1" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objConn, u_parms)
    End Sub

    '放棄報到原因/放棄報到(原因)
    Sub UPDATA_STUD_SELRESULT_ABANDON(ByRef myCmdArg As String)
        'Dim vESETID As String=TIMS.GetMyValue(myCmdArg, "ESETID")
        'Dim vESERNUM As String=TIMS.GetMyValue(myCmdArg, "ESERNUM")
        Dim vSETID As String = TIMS.GetMyValue(myCmdArg, "SETID")
        Dim vEnterDate As String = TIMS.GetMyValue(myCmdArg, "EnterDate")
        Dim vSerNum As String = TIMS.GetMyValue(myCmdArg, "SerNum")
        'Dim vIDNO As String=TIMS.GetMyValue(myCmdArg, "IDNO")
        Dim vOCID As String = TIMS.GetMyValue(myCmdArg, "OCID")

        'Hid_ABANDONReasonSub.Value=TIMS.ClearSQM2(Hid_ABANDONReasonSub.Value)
        Dim u_parms As New Hashtable
        'u_parms.Add("ABANDON", "Y")
        u_parms.Add("ABANDONReason", Hid_ABANDONReasonSub.Value)
        u_parms.Add("ABANDONACCT", sm.UserInfo.UserID)
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("SETID", TIMS.CINT1(vSETID))
        u_parms.Add("EnterDate", TIMS.Cdate2(vEnterDate))
        u_parms.Add("SerNum", TIMS.CINT1(vSerNum))
        u_parms.Add("OCID", TIMS.CINT1(vOCID))
        Dim u_sql As String = ""
        u_sql &= " UPDATE STUD_SELRESULT" & vbCrLf
        u_sql &= " SET ABANDON='Y'" & vbCrLf
        u_sql &= " ,ABANDONReason=@ABANDONReason" & vbCrLf
        u_sql &= " ,ABANDONACCT=@ABANDONACCT" & vbCrLf
        u_sql &= " ,ABANDONDATE=GETDATE()" & vbCrLf
        'AppliedStatus	是否報到 N
        u_sql &= " ,APPLIEDSTATUS='N'" & vbCrLf        '
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE SETID=@SETID" & vbCrLf
        u_sql &= " AND ENTERDATE=@ENTERDATE" & vbCrLf
        u_sql &= " AND SERNUM=@SERNUM" & vbCrLf
        u_sql &= " AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objConn, u_parms)
    End Sub

    '(Restore) Give up the registration (還原)放棄報到原因/放棄報到(原因)
    Sub UPDATA_STUD_ENTERTYPE2_ABANDON_RESTORE(ByRef myCmdArg As String)
        Dim vESETID As String = TIMS.GetMyValue(myCmdArg, "ESETID")
        Dim vESERNUM As String = TIMS.GetMyValue(myCmdArg, "ESERNUM")
        'Dim vSETID As String=TIMS.GetMyValue(myCmdArg, "SETID")
        'Dim vEnterDate As String=TIMS.GetMyValue(myCmdArg, "EnterDate")
        'Dim vSerNum As String=TIMS.GetMyValue(myCmdArg, "SerNum")
        'Dim vIDNO As String=TIMS.GetMyValue(myCmdArg, "IDNO")
        Dim vOCID As String = TIMS.GetMyValue(myCmdArg, "OCID")

        'Hid_ABANDONReasonSub.Value=TIMS.ClearSQM2(Hid_ABANDONReasonSub.Value)
        Dim u_parms As New Hashtable From {
            {"ESERNUM", TIMS.CINT1(vESERNUM)},
            {"ESETID", TIMS.CINT1(vESETID)},
            {"OCID1", TIMS.CINT1(vOCID)}
        }
        Dim u_sql As String = ""
        u_sql &= " UPDATE STUD_ENTERTYPE2" & vbCrLf
        u_sql &= " SET ABANDON=NULL ,ABANDONReason=NULL ,ABANDONACCT=NULL ,ABANDONDATE=NULL" & vbCrLf
        u_sql &= " WHERE ESERNUM=@ESERNUM" & vbCrLf
        u_sql &= " AND ESETID=@ESETID" & vbCrLf
        u_sql &= " AND OCID1=@OCID1" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objConn, u_parms)
    End Sub

    '(Restore) Give up the registration (還原)放棄報到原因/放棄報到(原因)
    Sub UPDATA_STUD_SELRESULT_ABANDON_RESTORE(ByRef myCmdArg As String)
        'Dim vESETID As String=TIMS.GetMyValue(myCmdArg, "ESETID")
        'Dim vESERNUM As String=TIMS.GetMyValue(myCmdArg, "ESERNUM")
        Dim vSETID As String = TIMS.GetMyValue(myCmdArg, "SETID")
        Dim vEnterDate As String = TIMS.GetMyValue(myCmdArg, "EnterDate")
        Dim vSerNum As String = TIMS.GetMyValue(myCmdArg, "SerNum")
        'Dim vIDNO As String=TIMS.GetMyValue(myCmdArg, "IDNO")
        Dim vOCID As String = TIMS.GetMyValue(myCmdArg, "OCID")
        'Hid_ABANDONReasonSub.Value=TIMS.ClearSQM2(Hid_ABANDONReasonSub.Value)
        Dim u_parms As New Hashtable
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("SETID", TIMS.CINT1(vSETID))
        u_parms.Add("EnterDate", TIMS.Cdate2(vEnterDate))
        u_parms.Add("SerNum", TIMS.CINT1(vSerNum))
        u_parms.Add("OCID", TIMS.CINT1(vOCID))
        Dim u_sql As String = ""
        u_sql &= " UPDATE STUD_SELRESULT" & vbCrLf
        u_sql &= " SET ABANDON=null ,ABANDONReason=null ,ABANDONACCT=null ,ABANDONDATE=null" & vbCrLf
        'AppliedStatus	是否報到 N
        u_sql &= " ,APPLIEDSTATUS='Y'" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE SETID=@SETID" & vbCrLf
        u_sql &= " AND ENTERDATE=@ENTERDATE" & vbCrLf
        u_sql &= " AND SERNUM=@SERNUM" & vbCrLf
        u_sql &= " AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objConn, u_parms)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'If e.CommandName="" Then Return
        If e.CommandArgument = "" Then Return
        Select Case e.CommandName
            Case "ABA" '放棄報到
                Dim myCmdArg As String = e.CommandArgument
                Hid_ABANDONReasonSub.Value = TIMS.ClearSQM2(Hid_ABANDONReasonSub.Value)
                UPDATA_STUD_ENTERTYPE2_ABANDON(myCmdArg)
                UPDATA_STUD_SELRESULT_ABANDON(myCmdArg)
                Dim strMsgBox As String = "報名資料已放棄報到!!"
                Common.MessageBox(Me, strMsgBox)
                '查詢
                Call Search1()
            Case "RESTO" '(還原)放棄報到
                Dim myCmdArg As String = e.CommandArgument
                '(Restore) Give up the registration
                UPDATA_STUD_ENTERTYPE2_ABANDON_RESTORE(myCmdArg)
                UPDATA_STUD_SELRESULT_ABANDON_RESTORE(myCmdArg)
                Dim strMsgBox As String = "報名資料已(還原)放棄報到!!"
                Common.MessageBox(Me, strMsgBox)
                '查詢
                Call Search1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim fg_test As Boolean = TIMS.CHK_IS_TEST_ENVC() 'fg_test OrElse fg_use1 
                Dim fg_use1 As Boolean = TIMS.CanUse3Y10WCost()
                '政府已補助經費   Start
                If drv("GovCost") > 0 Then
                    Dim strTitle As String = ""
                    Dim LabGovCost As Label = e.Item.FindControl("LabGovCost")
                    LabGovCost.Text = TIMS.CINT1(drv("GovCost"))
                    strTitle = "依查詢班級開訓日計算"
                    'If tGovCost2 <> "0" AndAlso tGovCost <> tGovCost2 Then strTitle += "(依查詢班級結訓日計算: " & tGovCost2 & " )"
                    TIMS.Tooltip(LabGovCost, strTitle)
                    If (fg_test OrElse fg_use1) Then
                        If CInt(LabGovCost.Text) >= Cst_接近最高額度9 Then '20080925 andy edit 超過 接近最高額度 的提示，將字變為紅色的
                            LabGovCost.ForeColor = Color.Red '=LabGovCost.ForeColor.Red
                            LabGovCost.Font.Bold = True
                            'TIMS.Tooltip(LabGovCost, "接近最高額度", True)
                        End If
                    Else
                        If CInt(LabGovCost.Text) >= Cst_接近最高額度6 Then '20080925 andy edit 超過 接近最高額度 的提示，將字變為紅色的
                            LabGovCost.ForeColor = Color.Red '=LabGovCost.ForeColor.Red
                            LabGovCost.Font.Bold = True
                            'TIMS.Tooltip(LabGovCost, "接近最高額度", True)
                        End If
                    End If
                End If
                '#Region "(No Use)"

                'If tGovCost <> "" Then
                '    Dim strTitle As String=""
                '    Dim LabGovCost As Label=e.Item.FindControl("LabGovCost")
                '    LabGovCost.Text=tGovCost
                '    strTitle="依查詢班級開訓日計算"
                '    'If tGovCost2 <> "0" AndAlso tGovCost <> tGovCost2 Then strTitle += "(依查詢班級結訓日計算: " & tGovCost2 & " )"
                '    TIMS.Tooltip(LabGovCost, strTitle)
                '    If CInt(LabGovCost.Text) >= Cst_接近最高額度 Then '20080925 andy edit 超過(三)改為-->(五)萬的提示，將字變為紅色的
                '        LabGovCost.ForeColor=LabGovCost.ForeColor.Red
                '        LabGovCost.Font.Bold=True
                '    End If
                'End If
                '政府已補助經費   End

                Dim MyCheck1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1") 'checkbox
                Dim SETID As HtmlInputHidden = e.Item.FindControl("SETID")
                Dim EnterDate As HtmlInputHidden = e.Item.FindControl("EnterDate")
                Dim SerNum As HtmlInputHidden = e.Item.FindControl("SerNum")
                Dim HidCFIRE1 As HtmlInputHidden = e.Item.FindControl("HidCFIRE1")
                Dim HidCFIRE1NS As HtmlInputHidden = e.Item.FindControl("HidCFIRE1NS")
                Dim HidCMASTER1 As HtmlInputHidden = e.Item.FindControl("HidCMASTER1") '認定為公司負責人
                Dim HidCMASTER1NS As HtmlInputHidden = e.Item.FindControl("HidCMASTER1NS") '已轉知
                Dim HidCMASTER1NT As HtmlInputHidden = e.Item.FindControl("HidCMASTER1NT") '已切結
                Dim HidIsStdBlack As HtmlInputHidden = e.Item.FindControl("HidIsStdBlack") '學員處分

                Dim ABANDONReason As TextBox = e.Item.FindControl("ABANDONReason") '放棄報到原因/放棄報到(原因)
                Dim BtnABA As Button = e.Item.FindControl("BtnABA") '放棄報到
                Dim BtnRESTO As Button = e.Item.FindControl("BtnRESTO") '(還原)放棄報到 -- RESTO 
                Dim HidABANDON As HtmlInputHidden = e.Item.FindControl("HidABANDON") '放棄報到

                Dim Label3name As Label = e.Item.FindControl("Label3name")
                Dim Hstar3 As HtmlInputHidden = e.Item.FindControl("Hstar3")
                Dim star3 As Label = e.Item.FindControl("star3")
                Dim Hstar4 As HtmlInputHidden = e.Item.FindControl("Hstar4")
                Dim star4 As Label = e.Item.FindControl("star4")

                ABANDONReason.Text = Convert.ToString(drv("ABANDONReason"))
                BtnABA.Attributes("onclick") = "return chkABANDONReason('" & ABANDONReason.ClientID & "','" & MyCheck1.ClientID & "');"
                HidABANDON.Value = Convert.ToString(drv("ABANDON")) '放棄報到

                HidCFIRE1.Value = Convert.ToString(drv("CFIRE1"))
                HidCFIRE1NS.Value = Convert.ToString(drv("CFIRE1NS"))
                HidCMASTER1.Value = Convert.ToString(drv("CMASTER1"))
                HidCMASTER1NS.Value = Convert.ToString(drv("CMASTER1NS"))
                HidCMASTER1NT.Value = Convert.ToString(drv("CMASTER1NT"))

                star3.Visible = False
                If TIMS.Chk_StudStatus(drv("IDNO").ToString, drv("STDate").ToString, drv("FTDate").ToString, drv("OCID").ToString, objConn) Then star3.Visible = True
                If TIMS.Chk_StudStatus(drv("IDNO").ToString, drv("OCID").ToString, objConn) Then star3.Visible = True
                Hstar3.Value = If(star3.Visible, "1", "")

                Dim ENCIDNO As String = RSA20031.AesEncrypt2($"{drv("IDNO")}")
                'SD_05_010.aspx
                'Label3.Attributes("onclick")="open_History('" & drv("IDNO").ToString & "');" '.CommandArgument=drv("IDNO").ToString
                Dim rqID As String = TIMS.Get_MRqID(Me)
                If Convert.ToString(drv("Name")) = "" Then
                    TIMS.sUtl_404NOTFOUND(Me, objConn)
                    Exit Sub
                End If
                Label3name.Style("CURSOR") = "hand"
                Label3name.Text = Convert.ToString(drv("Name"))
                Label3name.Attributes("onclick") = String.Concat("open_History('", ENCIDNO, "','", rqID, "');") '.CommandArgument=drv("IDNO").ToString

                Dim t_Label3name As String = ""
                t_Label3name = String.Concat("報名資料", vbCrLf, "IDNO:", drv("IDNO"), vbCrLf, "Birthday:", Common.FormatDate(drv("Birthday")))
                TIMS.Tooltip(Label3name, t_Label3name)

                SETID.Value = Convert.ToString(drv("SETID"))
                EnterDate.Value = Convert.ToString(drv("EnterDate"))
                SerNum.Value = Convert.ToString(drv("SerNum"))
                'chkbox
                'dtStudClass.select("idno='" &  Convert.ToString(drv("IDNO")) & "'")(0)("EnterDate")
                fff = "idno='" & Convert.ToString(drv("IDNO")) & "'"
                Dim flag_AppliedStatus_Y2 As Boolean = False
                If dtStudClass.Select(fff).Length > 0 AndAlso Convert.ToString(drv("AppliedStatus")) = "Y" Then
                    flag_AppliedStatus_Y2 = True
                End If
                If Convert.ToString(drv("AppliedStatus")).Equals("Y") Then MyCheck1.Checked = True

                'OJT-21030203： 學員參訓：新增批次勾稽檢查學員是否參加職前課程機制-START
                Dim flag_TrainITS As Boolean = False
                '勾稽檢查學員是否參加職前課程機制
                If Not MyCheck1.Checked Then
                    '(僅針對未完成報到的學員)
                    If CheckBoxITS1.Checked Then flag_TrainITS = TIMS.CheckTrainITS(dtITS, drv("IDNO").ToString, drv("STDate").ToString, drv("FTDate").ToString)
                Else
                    '(已完成報到的學員)
                    If oflag_Test Then '(測試用)
                        If CheckBoxITS1.Checked Then flag_TrainITS = TIMS.CheckTrainITS(dtITS, drv("IDNO").ToString, drv("STDate").ToString, drv("FTDate").ToString)
                    End If
                End If
                'TIMS.Utl_GetConfigSet("SD_03_001")
                If flag_TrainITS Then TIMS.Tooltip(MyCheck1, "學員同時參加職前課程,卡重參!")
                If flag_can_repeated_training AndAlso flag_TrainITS Then TIMS.Tooltip(MyCheck1, "(略過)職前卡重參!")
                If flag_can_repeated_training AndAlso flag_TrainITS Then flag_TrainITS = False '(略過)職前卡重參

                '(測試用)
                'If oflag_Test Then flag_TrainITS=True
                star4.Visible = If(flag_TrainITS, True, False)
                Hstar4.Value = If(star4.Visible, "1", "")
                'OJT-21030203： 學員參訓：新增批次勾稽檢查學員是否參加職前課程機制-END

                If flag_AppliedStatus_Y2 Then
                    Dim vMsg1 As String = ""
                    vMsg1 = "該學員已完成報到，不可取消該紀錄，"
                    fff = "idno='" & Convert.ToString(drv("IDNO")) & "'"
                    If dtStudClass.Select(fff).Length > 0 Then vMsg1 += Convert.ToString(dtStudClass.Select(fff)(0)("EnterDate"))
                    'MyCheck1.Checked=True
                    MyCheck1.Disabled = True
                    TIMS.Tooltip(MyCheck1, vMsg1)
                End If

                '(還原)放棄報到
                BtnRESTO.Visible = False 'If(flagS1, True, False)
                BtnRESTO.Style.Item("display") = "none"
                '放棄報到原因/放棄報到(原因)
                If MyCheck1.Checked Then
                    'MyCheck1.Disabled=True
                    ABANDONReason.Enabled = False
                    BtnABA.Enabled = False
                    ABANDONReason.Visible = False
                    Dim s_tipABA As String = "已報到不可放棄"
                    TIMS.Tooltip(ABANDONReason, s_tipABA)
                    TIMS.Tooltip(BtnABA, s_tipABA)
                ElseIf Convert.ToString(drv("ABANDON")) = "Y" Then
                    '放棄報到原因/放棄報到(原因)
                    MyCheck1.Disabled = True '鎖定 報到
                    ABANDONReason.Enabled = False
                    BtnABA.Enabled = False
                    BtnABA.Visible = False
                    Dim s_tipABA As String = "已放棄報到"
                    TIMS.Tooltip(MyCheck1, s_tipABA)
                    TIMS.Tooltip(ABANDONReason, s_tipABA)
                    TIMS.Tooltip(BtnABA, s_tipABA)

                    '(還原)放棄報到
                    Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                    BtnRESTO.Visible = If(flagS1, True, False)
                    BtnRESTO.Style.Item("display") = "none"
                    BtnRESTO.Attributes("onclick") = "javascript:return confirm('(還原)放棄報到，是否確定?');"
                ElseIf Convert.ToString(drv("ESERNUM")) = "" Then
                    '放棄報到原因/放棄報到(原因)
                    ABANDONReason.Enabled = False
                    BtnABA.Enabled = False
                    ABANDONReason.Visible = False
                    Dim s_tipABA As String = "(缺少線上報名資訊)無法放棄報到!"
                    TIMS.Tooltip(ABANDONReason, s_tipABA)
                    TIMS.Tooltip(BtnABA, s_tipABA)
                End If

                Dim myCmdArg As String = ""
                TIMS.SetMyValue(myCmdArg, "ESETID", Convert.ToString(drv("ESETID")))
                TIMS.SetMyValue(myCmdArg, "ESERNUM", Convert.ToString(drv("ESERNUM")))
                TIMS.SetMyValue(myCmdArg, "SETID", Convert.ToString(drv("SETID")))
                TIMS.SetMyValue(myCmdArg, "EnterDate", Convert.ToString(drv("EnterDate")))
                TIMS.SetMyValue(myCmdArg, "SerNum", Convert.ToString(drv("SerNum")))
                TIMS.SetMyValue(myCmdArg, "IDNO", Convert.ToString(drv("IDNO")))
                TIMS.SetMyValue(myCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(myCmdArg, "SelResultID", Convert.ToString(drv("SelResultID")))
                If BtnABA.Enabled Then BtnABA.CommandArgument = myCmdArg
                '(還原)放棄報到
                BtnRESTO.CommandArgument = myCmdArg

                'chkbox
                If Me.ViewState(Cst_IDNODoubles).ToString.IndexOf(drv("IDNO").ToString) > -1 Then
                    MyCheck1.Disabled = True
                    TIMS.Tooltip(MyCheck1, "身分證號與「" & drv("Name").ToString & "」相同，不可完成報到", True)
                Else
                    If Me.ViewState(Cst_IDNODoubles) <> "" Then Me.ViewState(Cst_IDNODoubles) += ","
                    Me.ViewState(Cst_IDNODoubles) += drv("IDNO").ToString
                End If

                e.Item.Cells(cst_名次).Text = e.Item.ItemIndex + 1

                Dim s_TRNDType As String = "-"
                Select Case drv("TRNDType").ToString
                    Case "1"
                        s_TRNDType = "甲式"
                    Case "2"
                        s_TRNDType = "乙式"
                    Case "3"
                        '星光幫用，但原SQL語法似乎有錯，有空請再確認 AMU
                        If drv("TRNDMode2").ToString = "3" Then s_TRNDType = "推介單"
                    Case Else
                        s_TRNDType = "-"
                End Select
                If Convert.ToString(drv("EnterPath")) = TIMS.cst_EnterPathW Then s_TRNDType = cst_EnterPathNameW '就服單位協助報名
                e.Item.Cells(cst_卷別).Text = s_TRNDType

                Dim s_NotExam As String = Convert.ToString(drv("SelResultName"))
                If drv("NotExam") Then s_NotExam &= "(免試)"
                If Convert.ToString(drv("LevelName")) <> "" AndAlso IsNumeric(drv("LevelName")) Then
                    Select Case Int(drv("LevelName"))
                        Case 1
                            s_NotExam &= "(第一階段插班)"
                        Case 2
                            s_NotExam &= "(第二階段插班)"
                        Case 3
                            s_NotExam &= "(第三階段插班)"
                        Case 4
                            s_NotExam &= "(第四階段插班)"
                        Case 5
                            s_NotExam &= "(第五階段插班)"
                    End Select
                End If
                e.Item.Cells(cst_錄訓結果).Text = s_NotExam

                '學員處分資料
                If Convert.ToString(drv("IsStdBlack")) = "Y" Then
                    If Not MyCheck1.Disabled AndAlso Not MyCheck1.Checked Then
                        '可使用(未失效) '未勾選 
                        MyCheck1.Disabled = True '使其失效，說明原因
                        TIMS.Tooltip(MyCheck1, "學員依處分日期及年限，仍在處分期間者。")
                    End If
                End If
                If Not Button2.Enabled AndAlso Not MyCheck1.Disabled Then
                    MyCheck1.Disabled = False '使其失效，說明原因
                    TIMS.Tooltip(MyCheck1, "完成報到 已鎖定，該資料 同步鎖定。")
                End If
                If oflag_Test Then '(測試用)
                    If MyCheck1.Disabled Then
                        MyCheck1.Disabled = False '開啟checkbox 失效恢復有效功能
                        TIMS.Tooltip(MyCheck1, "測試環境 供測試!!")
                    End If
                End If

                'If Convert.ToString(drv("GOVKILL"))="Y" Then
                '    'e.Item.Enabled=False
                '    MyCheck1.Disabled=True '使其失效，說明原因
                '    TIMS.Tooltip(MyCheck1, cst_Mgc219)
                'End If
        End Select
    End Sub

    '[儲存前]檢核
    Function CheckData1(ByRef Errmsg As String, ByRef dtCS As DataTable) As Boolean
        Dim rst As Boolean = True
        Errmsg = ""

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        HidOCID1.Value = TIMS.ClearSQM(HidOCID1.Value)
        If HidOCID1.Value = "" OrElse OCIDValue1.Value = "" Then
            Errmsg &= "儲存與查詢班級不同，請重新查詢!!" & vbCrLf
            Return False
        End If

        'Dim strTPlanID28AppPlan As String=TIMS.Cst_TPlanID28AppPlan
        '28:產業人才投資方案
        '46:補助辦理保母職業訓練
        '47:補助辦理照顧服務員職業訓練S
        'strTPlanID28AppPlan += ",46,47"
        If HidOCID1.Value <> OCIDValue1.Value Then
            Errmsg &= "儲存與查詢班級不同，請重新查詢!!" & vbCrLf
            Return False
        End If

        txtEnterDate.Text = TIMS.ClearSQM(txtEnterDate.Text)
        If txtEnterDate.Text = "" Then Errmsg &= "請輸入報到日期!" & vbCrLf
        If txtEnterDate.Text <> "" AndAlso Not TIMS.IsDate1(txtEnterDate.Text) Then
            Errmsg &= "報到日期必須是正確的日期格式!" & vbCrLf
        Else
            txtEnterDate.Text = TIMS.Cdate3(txtEnterDate.Text)
        End If
        '報到日期(只能是系統時間今天或今天以前)請確認!\n'; }
        If Errmsg = "" AndAlso DateDiff(DateInterval.Day, CDate(txtEnterDate.Text), CDate(Now)) < 0 Then
            Errmsg &= "報到日期(只能是系統時間今天或今天以前)!" & vbCrLf
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If Errmsg = "" AndAlso DateDiff(DateInterval.Day, CDate(hSTDate.Value), CDate(txtEnterDate.Text)) < 0 Then
                Errmsg &= "「報到日期」僅能選擇開訓日當天~開訓日後14日內,(為開訓日(含)之後14天內)!" & vbCrLf
            End If
            If Errmsg = "" AndAlso DateDiff(DateInterval.Day, CDate(txtEnterDate.Text), CDate(hSTDate14.Value)) < 0 Then
                Errmsg &= "「報到日期」僅能選擇開訓日當天~開訓日後14日內,(為開訓日(含)之後14天內)!" & vbCrLf
            End If
            '報到日期(已超過系統時間30天)請確認!
            If Errmsg = "" AndAlso sm.UserInfo.LID <> 0 AndAlso DateDiff(DateInterval.Day, CDate(txtEnterDate.Text), CDate(Now)) > 30 Then
                Errmsg &= "(非署)「報到日期」(已超過系統時間30天)!" & vbCrLf
            End If
        End If
        If Errmsg <> "" Then Return False

        '未鎖定判斷
        'Dim sql As String=""
        'sql="" & vbCrLf
        'sql &= " SELECT ISNULL(a.ClassID2,a.ClassID) ClassID" & vbCrLf
        'sql &= " ,b.CyclType ,b.Years ,b.TNum" & vbCrLf
        'sql &= " ,CONVERT(VARCHAR, b.STDate, 111) STDate" & vbCrLf
        'sql &= " ,CONVERT(VARCHAR, b.ExamDate, 111) ExamDate" & vbCrLf
        'sql &= " ,ip.TPlanID" & vbCrLf
        'sql &= " FROM ID_CLASS a" & vbCrLf
        'sql &= " JOIN CLASS_CLASSINFO b ON a.CLSID=b.CLSID" & vbCrLf
        'sql &= " JOIN ID_PLAN ip ON ip.planid=b.planid" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        'sql &= " AND b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        'Dim dr As DataRow=DbAccess.GetOneRow(sql, objConn)

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        If drCC Is Nothing Then
            Errmsg &= "查詢班級有誤，請重新查詢!!" & vbCrLf
            Return False
        End If

        HidTNum.Value = Convert.ToString(drCC("TNum")) 'dr.Item("TNum")
        ViewState(vs_StDate) = TIMS.Cdate3(drCC("STDate"))
        Dim sTPlanID As String = Convert.ToString(drCC("TPlanID"))
        'OCIDValue1.Value=TIMS.ClearSQM(OCIDValue1.Value)

        '暫時權限Table
        Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objConn)
        '過了使用期限 True(不可使用)   False(可使用)
        Dim flag_IsOverEndDate As Boolean = TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_學員參訓, dtArc)

        '遞補期限≧報到日期>甄試日期
        Dim dEnterDate As Date = CDate(txtEnterDate.Text) '報到日期
        '非系統管理者 '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "2", "1"
                Dim iTmpDay As Integer = 14 '遞補期限
                If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(Today)) > iTmpDay Then
                    Button2.Enabled = False '鎖定
                    Dim sTmpDay2 As String = "作業日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內完成)!"
                    TIMS.Tooltip(Button2, sTmpDay2)
                    Dim flag_show_error_msg1 As Boolean = False
                    Dim flag_show_error_msg2 As Boolean = False
                    If Not oflag_Test Then flag_show_error_msg1 = True
                    If Not flag_can_ignore_control Then flag_show_error_msg2 = True
                    If flag_show_error_msg1 AndAlso flag_show_error_msg2 Then
                        If flag_IsOverEndDate Then
                            Errmsg &= sTmpDay2 & vbCrLf
                            Return False
                        End If
                    End If
                End If

                '遞補期限≧報到日期>甄試日期
                Dim dStDate14 As Date = CDate(ViewState(vs_StDate)).AddDays(iTmpDay) '遞補期限(開訓日+14)
                If DateDiff(DateInterval.Day, dStDate14, dEnterDate) > 0 Then
                    Button2.Enabled = False '鎖定
                    Dim sTmpDay2 As String = String.Concat("遞補期限與報到日期，已超過", iTmpDay, "天(須於", iTmpDay, "天內完成)!")
                    TIMS.Tooltip(Button2, sTmpDay2)
                    Dim flag_show_error_msg1 As Boolean = False
                    Dim flag_show_error_msg2 As Boolean = False
                    If Not oflag_Test Then flag_show_error_msg1 = True
                    If Not flag_can_ignore_control Then flag_show_error_msg2 = True
                    If flag_show_error_msg1 AndAlso flag_show_error_msg2 Then
                        If flag_IsOverEndDate Then
                            Errmsg &= sTmpDay2 & vbCrLf
                            Return False
                        End If
                    End If
                End If
            Case Else
        End Select

        '遞補期限≧報到日期>甄試日期
        If TIMS.Cdate3(drCC("ExamDate")) <> "" Then
            Dim dExamDate As Date = CDate(drCC("ExamDate")) '甄試日期
            If DateDiff(DateInterval.Day, dEnterDate, dExamDate) > 0 Then
                Button2.Enabled = False '鎖定
                Dim sTmpDay2 As String = "報到日期 不可早於 甄試日期!"
                TIMS.Tooltip(Button2, sTmpDay2)
                If Not oflag_Test Then
                    Errmsg &= sTmpDay2 & vbCrLf
                    Return False
                End If
            End If
        End If

        '若未「勾選」學員就按「完成報到」，跳出提醒視窗『請先「勾選」學員再按「完成報到」。』。
        Dim iChk_Stud_Count As Integer = 0
        Dim iClassCount As Integer = 0 '己經有學員資料的人數
        'S1='Y' (有效學員)
        fff = "S1='Y'"
        If dtCS IsNot Nothing Then iClassCount = dtCS.Select(fff).Length
        'iClassCount=dtCS.Rows.Count

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        Dim v_IDNOs As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim MyCheck1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1") '勾選 'Check1.Disabled=False 有效
            If MyCheck1 Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            Dim SETID As HtmlInputHidden = eItem.FindControl("SETID")
            If SETID Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            Dim drST As DataRow = TIMS.GET_StudEnterTemp(SETID.Value, objConn)
            If drST Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            'Dim sIDNO As String=Convert.ToString(drST("IDNO")) '取得報名學員身分證號。
            'Dim sNAME As String=Convert.ToString(drST("NAME")) '取得報名學員姓名。
            Dim t_IDNO As String = TIMS.ClearSQM(drST("IDNO"))
            If v_IDNOs <> "" Then v_IDNOs &= ","
            v_IDNOs &= t_IDNO
        Next

        Try
            '勾稽檢查學員是否參加職前課程機制
            dtITS = TIMS.GetTrainingListS3(sm, objConn, v_IDNOs)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Errmsg &= String.Concat("勾稽檢查學員是否參加職前課程，查詢時發生錯誤!", vbCrLf, ex.Message) 'TIMS.cst_NODATAMsg2
            Return False
        End Try

        '檢核學員重複參訓。
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim MyCheck1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1") '勾選 'Check1.Disabled=False 有效
            If MyCheck1 Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            Dim SETID As HtmlInputHidden = eItem.FindControl("SETID")
            If SETID Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            Dim drST As DataRow = TIMS.GET_StudEnterTemp(SETID.Value, objConn)
            If drST Is Nothing Then
                Errmsg &= TIMS.cst_NODATAMsg2
                Return False
            End If
            Dim sIDNO As String = TIMS.ClearSQM(drST("IDNO")) '取得報名學員身分證號。
            Dim sNAME As String = Convert.ToString(drST("NAME")) '取得報名學員姓名。
            Dim Label3name As Label = eItem.FindControl("Label3name") 'NAME
            Dim HidCFIRE1 As HtmlInputHidden = eItem.FindControl("HidCFIRE1")
            Dim HidCFIRE1NS As HtmlInputHidden = eItem.FindControl("HidCFIRE1NS")
            Dim HidCMASTER1 As HtmlInputHidden = eItem.FindControl("HidCMASTER1")
            'Dim HidCMASTER1NS As HtmlInputHidden=eItem.FindControl("HidCMASTER1NS") '已轉知
            Dim HidCMASTER1NT As HtmlInputHidden = eItem.FindControl("HidCMASTER1NT") '已切結(不再卡卡)
            Dim HidIsStdBlack As HtmlInputHidden = eItem.FindControl("HidIsStdBlack") '學員處分
            If sNAME.Length > 20 Then
                Dim strMsg1 As String = Replace(cst_StdWMsg2, "@LAB3", sNAME) & vbCrLf
                'strMsg1="學員" & sNAME & "，姓名長度超過系統範圍，不可儲存!!" & vbCrLf
                Errmsg &= strMsg1
                Return False
            End If
            Dim iEnterCnt As Integer = 1
            If dtCS IsNot Nothing AndAlso sIDNO <> "" Then
                '0:該報名學員尚未成為該班學生。
                iEnterCnt = dtCS.Select("IDNO='" & sIDNO & "'").Length
            End If
            'If oflag_Test Then '(測試用)
            '    If iEnterCnt <> 0 Then iEnterCnt=0 '測試環境 供測試!!
            'End If
            If HidIsStdBlack.Value = "Y" Then
                '有多筆 學員處分資料
                If iEnterCnt = 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
                    Dim strMsg1 As String = Replace(cst_StdWMsg3, "@LAB3", sNAME) & vbCrLf
                    Errmsg &= strMsg1
                    Return False
                End If
            End If
            '#Region "(No Use)"

            'If oflag_Test Then
            '    Dim strMsg1 As String=""
            '    strMsg1="學員" & Label3.Text & "，依處分日期及年限，仍在處分期間者，學員參訓，被處分者，不可儲存。" & vbCrLf
            '    Errmsg &= strMsg1
            '    Return False
            'End If

            '未列排除名單要執行 檢核'具公司/商業負責人身分
            'If Not TIMS.Cst_NotTPlanID5.IndexOf(sTPlanID) > -1 Then
            '    If iEnterCnt=0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
            '        Throw New Exception("應確認: 具公司/商業負責人身分 檢核, 在產投/在職 是否不應存在")

            '        '具公司/商業負責人身分
            '        Dim flagCMASTER1 As String=""
            '        If HidCMASTER1.Value="Y" _
            '            OrElse TIMS.Chk_Master(Me, Chkws1, sIDNO, SETID.Value, OCIDValue1.Value)="Y" Then
            '            flagCMASTER1="Y"
            '        End If
            '        '具公司/商業負責人身分 且 未切結
            '        If flagCMASTER1="Y" _
            '            AndAlso HidCMASTER1NT.Value <> "Y" Then
            '            Dim strmessage As String=cst_CMasterMsg1
            '            strmessage=Replace(strmessage, "JJJ", Label3name.Text) & vbCrLf
            '            'strmessage += " 民眾" & Label3.Text & "具公司/商業負責人身分，非屬失業勞工，不得參加失業者職前訓練。" & vbCrLf
            '            'Common.MessageBox(Me, strmessage)
            '            'Exit Function
            '            Errmsg &= strmessage
            '            Return False
            '        End If
            '    End If
            'End If

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                If iEnterCnt = 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
                    iChk_Stud_Count += 1
                    '修改說明:
                    '針對報名民眾於不同階段發生參訓時段重疊(報名)情形，設立無法儲存階段及彈跳提醒視窗
                    '，於「學員參訓」將學員勾選按下「完成報到」，無法儲存時，給予彈跳視窗提醒(如圖3.jpg)。
                    '此外，參訓時段重疊報名情形比對，「錄取作業」項下備取生和未選取者&「e網報名審核」
                    '項下報名審核未點選成功或失敗者，於該課程開訓後第15天取消比對。
                    'by AMU 20160202
                    'Dim drET2 As DataRow=TIMS.Get_ENTERTYPE2(Hid_eSerNum.Value, objconn)
                    OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
                    Dim xStudInfo As String = ""
                    TIMS.SetMyValue(xStudInfo, "IDNO", sIDNO)
                    TIMS.SetMyValue(xStudInfo, "OCID1", OCIDValue1.Value)
                    '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                    Call TIMS.ChkStudDouble(timsSer1, Errmsg, Label3name.Text, xStudInfo)
                End If
            End If

            'Call TIMS.OpenDbConn(objConn)
            ''該民眾不具失、待業身分，不得參加失業者職前訓練。STUD_SELRESULTBLIDET / STUD_SELRESULTBLI
            'Dim dtBLIDET1 As DataTable=TIMS.Get_dtBLIDET1(OCIDValue1.Value, objConn)

            'OJT-21030203： 學員參訓：新增批次勾稽檢查學員是否參加職前課程機制-START
            Dim flag_TrainITS As Boolean = False
            '勾稽檢查學員是否參加職前課程機制
            If MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
                '(僅針對未完成報到的學員)
                flag_TrainITS = TIMS.CheckTrainITS(dtITS, sIDNO, TIMS.Cdate3(drCC("STDate")), TIMS.Cdate3(drCC("FTDate")))
                If flag_TrainITS Then
                    Errmsg &= $"學員「{sNAME}」同時參加職前課程，無法完成報到，請重新確認!。!!!{vbCrLf}"
                    Return False
                End If
            End If

            'TIMS.Utl_GetConfigSet("SD_03_001")
            'If flag_TrainITS Then TIMS.Tooltip(MyCheck1, "學員同時參加職前課程,卡重參!")
            'If flag_can_repeated_training AndAlso flag_TrainITS Then TIMS.Tooltip(MyCheck1, "(略過)職前卡重參!")
            'If flag_can_repeated_training AndAlso flag_TrainITS Then flag_TrainITS=False '(略過)職前卡重參

            If TIMS.Cst_TPlanID28AppPlan4.IndexOf(sTPlanID) > -1 Then
                '產投在職班 '檢查報到人數不能大於訓練人數(限定訓練人數)
                If iEnterCnt = 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
                    iClassCount += 1 '如果訓練人數小於報到人數
                    If TIMS.CINT1(HidTNum.Value) < iClassCount Then '如果訓練人數小於報到人數
                        Errmsg &= $" 訓練人數限制:{HidTNum.Value}{vbCrLf}參訓報到人數不能大於訓練人數{vbCrLf}"
                        Return False
                    End If
                End If
            Else
                '#Region "(No Use)"
                ''其他職前班。
                'If iEnterCnt <> 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled Then
                '    Dim strMsg1 As String=Replace(cst_StdWMsg1, "@LAB3", sNAME) & vbCrLf
                '    Errmsg &= strMsg1 & vbCrLf
                '    Return False
                'End If
                'If iEnterCnt=0 AndAlso MyCheck1.Checked AndAlso MyCheck1.Disabled=False Then
                '    Dim ss As String="IDNO='" & sIDNO & "'"
                '    If dtBLIDET1.Select(ss).Length > 0 Then
                '        Dim strmessage As String=""
                '        '該民眾不具失、待業身分，不得參加失業者職前訓練。
                '        strmessage &= Replace(cst_CBLIDETMsg1, "@LAB3", Label3name.Text) & vbCrLf
                '        Errmsg &= strmessage
                '        Return False
                '    End If
                '    '為非自願離職者 (使用)
                '    If TIMS.Cst_TPlanID_useCFIRE1.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '        If HidCFIRE1.Value="Y" AndAlso HidCFIRE1NS.Value="" Then
                '            '該學員為非自願離職者，無法報到
                '            Dim strmessage As String=""
                '            strmessage &= Replace(cst_CFIRE1Msg1, "@LAB3", Label3name.Text) & vbCrLf
                '            Errmsg &= strmessage
                '            Return False
                '        End If
                '    End If
                'End If
            End If
        Next

        If Errmsg <> "" Then Return False
        If Errmsg = "" Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                If iChk_Stud_Count = 0 Then
                    Errmsg &= cst_errTPlanID28Msg1 & vbCrLf
                    Return False
                End If
            End If
        End If

        HidEnterDate.Value = ""
        Dim strMsgbox As String = ""
        If txtEnterDate.Text = "" Then
            strMsgbox += "請輸入報到日期!" & vbCrLf
        Else
            If Not TIMS.IsDate1(txtEnterDate.Text) Then strMsgbox += "報到日期必須是正確的日期格式!" & vbCrLf
        End If
        If strMsgbox <> "" Then
            '報到日期 發生錯誤
            'Common.MessageBox(Me, " 資料錯誤，錯誤如下:" & vbCrLf & msgbox)
            'Exit Function
            Errmsg &= strMsgbox
            Return False
        End If
        '報到日期未發生錯誤
        txtEnterDate.Text = TIMS.Cdate3(txtEnterDate.Text)
        HidEnterDate.Value = Common.FormatDate(CDate(txtEnterDate.Text))

        If strMsgbox <> "" Then
            '報到日期 發生錯誤
            'Common.MessageBox(Me, " 資料錯誤，錯誤如下:" & vbCrLf & msgbox)
            'Exit Function
            Errmsg &= strMsgbox
            Return False
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>
    ''' 將報名資料寫入學員資料檔內 STUD_ENTERTEMP/STUD_ENTERTYPE/STUD_ENTERTRAIN2
    ''' </summary>
    ''' <param name="SETID"></param>
    ''' <param name="EnterDate"></param>
    ''' <param name="SerNum"></param>
    ''' <returns></returns>
    Function GET_STUD_ENTERTYPE(ByRef SETID As HtmlInputHidden, ByRef EnterDate As HtmlInputHidden, ByRef SerNum As HtmlInputHidden) As DataRow
        'Dim SETID As HtmlInputHidden=eItem.FindControl("SETID")
        'Dim EnterDate As HtmlInputHidden=eItem.FindControl("EnterDate")
        'Dim SerNum As HtmlInputHidden=eItem.FindControl("SerNum")
        Dim dr1 As DataRow = Nothing

        Dim pms1 As New Hashtable From {{"SETID", SETID.Value}, {"EnterDate", TIMS.Cdate2(EnterDate.Value)}, {"SerNum", SerNum.Value}}
        Dim sql As String = ""
        sql &= " select a.IDNO ,a.NAME,a.SEX,a.BIRTHDAY" & vbCrLf
        sql &= " ,a.PASSPORTNO,a.MARITALSTATUS" & vbCrLf
        sql &= " ,a.DEGREEID,a.GRADID,a.SCHOOL,a.DEPARTMENT" & vbCrLf
        sql &= " ,a.MILITARYID ,a.ZIPCODE,a.ADDRESS" & vbCrLf
        sql &= " ,a.PHONE1,a.PHONE2,a.CELLPHONE" & vbCrLf
        sql &= " ,a.EMAIL,a.NOTES,a.ISAGREE" & vbCrLf
        sql &= " ,a.ESETID,a.ZIPCODE6W,a.ZIPCODE_N" & vbCrLf 'STUD_ENTERTEMP a

        sql &= " ,b.ESERNUM,b.SETID ,b.CCLID ,b.SEID ,b.Ticket_NO ,b.TRNDMode ,b.TRNDType ,b.IdentityID" & vbCrLf
        sql &= " ,b.EnterChannel ,b.OCID1 ,b.SupplyID ,b.BudID ,b.HighEduBg ,b.WorkSuppIdent" & vbCrLf
        sql &= " ,b.PriorWorkType1 PriorWorkType3 ,b.PriorWorkOrg1 PriorWorkOrg3" & vbCrLf
        sql &= " ,b.ActNo ActNo2" & vbCrLf
        sql &= " ,b.SOfficeYM1 SOfficeYM3" & vbCrLf
        sql &= " ,b.FOfficeYM1 FOfficeYM3" & vbCrLf 'STUD_ENTERTYPE b 

        sql &= " ,c.TMID ,pp.ClassCate ,pp.IsBusiness ,f.OrgKind ,e.DistID" & vbCrLf
        sql &= " ,ISNULL(d.ClassID2,d.ClassID) ClassID" & vbCrLf
        sql &= " ,c.Years ,c.CyclType ,c.STDate ,c.FTDate ,c.ClassCName" & vbCrLf
        sql &= " ,c.TaddressZip ,c.TAddress ,c.THours ,f.OrgName ,g.ContactName" & vbCrLf
        sql &= " ,g.Phone ,f.ComCIDNO ,g.MasterName" & vbCrLf
        sql &= " ,c.LevelCount ,h.LevelName" & vbCrLf
        sql &= " ,se3.ZIPCODE2,se3.HOUSEHOLDADDRESS,se3.ZIPCODE2_6W" & vbCrLf 'STUD_ENTERTRAIN2
        sql &= " ,se3.MIDENTITYID" & vbCrLf
        sql &= " ,se3.HANDTYPEID,se3.HANDLEVELID" & vbCrLf
        sql &= " ,se3.PRIORWORKORG1,se3.TITLE1" & vbCrLf
        sql &= " ,se3.PRIORWORKORG2,se3.TITLE2" & vbCrLf
        sql &= " ,se3.SOFFICEYM1,se3.FOFFICEYM1" & vbCrLf
        sql &= " ,se3.SOFFICEYM2,se3.FOFFICEYM2" & vbCrLf
        sql &= " ,se3.PRIORWORKPAY,se3.REALJOBLESS,se3.JOBLESSID" & vbCrLf
        sql &= " ,se3.TRAFFIC,se3.SHOWDETAIL" & vbCrLf
        sql &= " ,se3.ACCTMODE" & vbCrLf
        sql &= " ,se3.POSTNO,se3.ACCTHEADNO,se3.BANKNAME" & vbCrLf
        sql &= " ,se3.ACCTEXNO,se3.EXBANKNAME" & vbCrLf
        sql &= " ,se3.ACCTNO,se3.FIRDATE" & vbCrLf
        sql &= " ,se3.UNAME,se3.INTAXNO" & vbCrLf
        sql &= " ,se3.ACTNO,se3.ACTNAME" & vbCrLf
        sql &= " ,se3.SERVDEPT,se3.JOBTITLE" & vbCrLf
        sql &= " ,se3.ZIP,se3.ADDR,se3.TEL,se3.FAX" & vbCrLf
        sql &= " ,se3.SDATE,se3.SJDATE,se3.SPDATE" & vbCrLf
        sql &= " ,se3.Q1,se3.Q2_1,se3.Q2_2,se3.Q2_3,se3.Q2_4" & vbCrLf
        sql &= " ,se3.Q3,se3.Q3_OTHER,se3.Q4,se3.Q5" & vbCrLf
        sql &= " ,se3.Q61,se3.Q62,se3.Q63,se3.Q64" & vbCrLf
        sql &= " ,se3.ISEMAIL" & vbCrLf
        sql &= " ,se3.ACTTYPE,se3.SCALE" & vbCrLf
        sql &= " ,se3.ACTTEL" & vbCrLf
        sql &= " ,se3.ZIPCODE3,se3.ZIPCODE3_6W,se3.ACTADDRESS" & vbCrLf
        sql &= " ,se3.INSURED" & vbCrLf
        sql &= " ,se3.SERVDEPTID,se3.JOBTITLEID" & vbCrLf
        sql &= " ,se3.ZIPCODE3_N,se3.ZIPCODE2_N" & vbCrLf
        sql &= " ,se3.ZIP6W,se3.ZIP_N" & vbCrLf
        sql &= " ,se3.Q2_3 Q2_33" & vbCrLf
        sql &= " ,se3.Q2_4 Q2_44" & vbCrLf
        sql &= " FROM STUD_ENTERTEMP a" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b ON b.SETID=a.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID1=c.OCID" & vbCrLf
        sql &= " JOIN ID_CLASS d ON c.CLSID=d.CLSID" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP e ON e.RID=c.RID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO f ON e.OrgID=f.OrgID" & vbCrLf
        sql &= " JOIN ORG_ORGPLANINFO g ON g.RSID=e.RSID" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.PlanID=C.PlanID AND pp.ComIDNO=C.ComIDNO AND pp.SeqNo=C.SeqNo" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSLEVEL h ON b.CCLID=h.CCLID" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERTRAIN2 se3 ON se3.SEID=b.SEID" & vbCrLf
        'Dim pms As New Hashtable From {{"SETID", SETID.Value}, {"EnterDate", CDate(EnterDate.Value)}, {"SerNum", SerNum.Value}}
        sql &= " WHERE b.SETID=@SETID AND b.EnterDate=@EnterDate AND b.SerNum=@SerNum" & vbCrLf
        'STUD_ENTERTYPE (select b)
        dr1 = DbAccess.GetOneRow(sql, objConn, pms1)
        Return dr1
    End Function

    ''' <summary>儲存-STUD_SELRESULT /新增 CLASS_STUDENTSOFCLASS</summary>
    ''' <param name="strMsgbox"></param>
    ''' <param name="dtCS"></param>
    Sub SaveData1(ByRef strMsgbox As String, ByRef dtCS As DataTable)
        Dim pms1 As New Hashtable From {{"OCID", OCIDValue1.Value}}
        Dim sql As String = ""
        sql &= " SELECT ISNULL(a.ClassID2,a.ClassID) ClassID" & vbCrLf
        sql &= " ,b.CyclType ,b.Years ,b.TNum" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, b.STDate, 111) STDate ,CONVERT(VARCHAR, b.ExamDate, 111) ExamDate" & vbCrLf
        sql &= " ,ip.TPlanID" & vbCrLf
        sql &= " FROM ID_CLASS a" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO b ON a.CLSID=b.CLSID" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.planid=b.planid" & vbCrLf
        sql &= " WHERE b.OCID =@OCID" & vbCrLf
        Dim drOCID As DataRow = DbAccess.GetOneRow(sql, objConn, pms1)
        HidTNum.Value = Convert.ToString(drOCID("TNum")) 'dr.Item("TNum")
        ViewState(vs_StDate) = Convert.ToString(drOCID("STDate"))
        Dim sTPlanID As String = Convert.ToString(drOCID("TPlanID"))

        'Dim drOCID As DataRow=TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        '檢查學員是否在訓中
        Dim vIDNO As String = ""
        Dim vBirthday As String = ""
        Dim vOCID1 As Integer = 0
        'Dim i As Integer=0  '計算Stud_StudentInfo 的筆數
        'Dim j As Integer=0  '計算Stud_StudentInfo 的筆數
        'Dim z As Integer=0  '計算Stud_SubData 的筆數

        'Dim ss As String="" '搜尋字
        'Dim Trans As SqlTransaction=Nothing
        Dim dt As DataTable
        Dim dt9 As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim dr1 As DataRow
        Dim iEnterNum As Integer = 1
        'Dim sql As String=""
        sql = ""
        sql &= " SELECT ss.IDNO" & vbCrLf '取出己經有學員資料的人數
        sql &= " FROM CLASS_STUDENTSOFCLASS cs "
        sql &= " JOIN STUD_STUDENTINFO ss on ss.sid=cs.sid "
        sql &= " WHERE cs.OCID=@OCID AND ss.IDNO=@IDNO" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objConn)

        'Dim strMsgbox As String=""
        For Each eItem As DataGridItem In DataGrid1.Items
            '勾選 'Check1.Disabled=False 有效
            Dim MyCheck1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            If MyCheck1 Is Nothing Then
                strMsgbox &= TIMS.cst_NODATAMsg2
                Exit Sub 'Return False
            End If
            Dim SETID As HtmlInputHidden = eItem.FindControl("SETID")
            Dim EnterDate As HtmlInputHidden = eItem.FindControl("EnterDate")
            Dim SerNum As HtmlInputHidden = eItem.FindControl("SerNum")
            If SETID Is Nothing Then
                strMsgbox &= TIMS.cst_NODATAMsg2
                Exit Sub 'Return False
            End If
            Dim Label3name As Label = eItem.FindControl("Label3name")
            SETID.Value = TIMS.ClearSQM(SETID.Value)
            EnterDate.Value = TIMS.ClearSQM(EnterDate.Value)
            SerNum.Value = TIMS.ClearSQM(SerNum.Value)
            OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
            Dim drST As DataRow = TIMS.GET_StudEnterTemp(SETID.Value, objConn)
            If drST Is Nothing Then
                strMsgbox &= "學員" & Label3name.Text & "參訓 資料有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                Exit For
            End If
            Dim sIDNO As String = Convert.ToString(drST("IDNO")) '取得報名學員身分證號。
            'Dim sNAME As String=Convert.ToString(drST("NAME")) '取得報名學員姓名。
            'Dim sIDNO As String=TIMS.GET_StudEnterTemp IDNO(SETID.Value, objConn) '取得報名學員身分證號。
            Dim iEnterCnt As Integer = 1
            If sIDNO <> "" Then
                '0:該報名學員尚未成為該班學生。
                fff = "IDNO='" & sIDNO & "'"
                iEnterCnt = dtCS.Select(fff).Length
            End If
            If sIDNO = "" Then iEnterCnt = 0
            'ADMISSION
            If iEnterCnt = 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled AndAlso TIMS.CINT1(OCIDValue1.Value) > 0 AndAlso sIDNO <> "" Then
                'Call TIMS.OpenDbConn(objConn)
                Dim dt8 As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("OCID", SqlDbType.Int).Value = TIMS.CINT1(OCIDValue1.Value)
                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = sIDNO
                    dt8.Load(.ExecuteReader())
                End With
                '0:該報名學員尚未成為該班學生。
                iEnterCnt = dt8.Rows.Count
            End If

            'AndAlso MyCheck1.Checked
            If iEnterCnt = 0 AndAlso Not MyCheck1.Disabled AndAlso TIMS.CINT1(OCIDValue1.Value) > 0 AndAlso sIDNO <> "" Then
                '更新 甄試結果試算檔(Stud_SelResult)
                Dim v_AppliedStatus As String = "Y"
                If Not MyCheck1.Checked Then v_AppliedStatus = "N"

                Dim pmsU As New Hashtable From {{"SETID", SETID.Value}, {"EnterDate", TIMS.Cdate2(EnterDate.Value)}, {"SerNum", SerNum.Value}}
                pmsU.Add("AppliedStatus", v_AppliedStatus)
                pmsU.Add("ModifyAcct", sm.UserInfo.UserID)
                Dim sqlU As String = ""
                sqlU &= " UPDATE STUD_SELRESULT"
                sqlU &= " SET AppliedStatus =@AppliedStatus"
                sqlU &= " ,Admission='Y', ModifyAcct =@ModifyAcct ,ModifyDate=GETDATE()"
                sqlU &= " WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum" ' TIMS.to_date(EnterDate.Value)
                DbAccess.ExecuteNonQuery(sqlU, objConn, pmsU)
            End If

            If iEnterCnt = 0 AndAlso MyCheck1.Checked AndAlso Not MyCheck1.Disabled AndAlso TIMS.CINT1(OCIDValue1.Value) > 0 AndAlso sIDNO <> "" Then
                '取得學員身分證號／生日／報名班級
                Dim pms3 As New Hashtable From {{"SETID", SETID.Value}, {"EnterDate", TIMS.Cdate2(EnterDate.Value)}, {"SerNum", SerNum.Value}}
                Dim sql3 As String = ""
                sql3 &= " SELECT a.IDNO ,a.Birthday ,b.OCID1,b.EnterChannel"  '1.網(28:預設);2.現(TIMS:預設);3.通;4.推
                sql3 &= " FROM STUD_ENTERTEMP a "
                sql3 &= " JOIN STUD_ENTERTYPE b ON a.SETID=b.SETID "
                sql3 &= " WHERE b.SETID =@SETID AND b.EnterDate =@EnterDate AND b.SerNum=@SerNum "
                Dim dr As DataRow = DbAccess.GetOneRow(sql3, objConn, pms3)
                If dr Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                    Exit Sub
                End If
                vIDNO = TIMS.ChangeIDNO(dr("IDNO"))
                vBirthday = Common.FormatDate(dr("Birthday"))
                vOCID1 = TIMS.CINT1(dr("OCID1"))

                '檢核是否為網路報名(true:職前班不儲存(清空)就業狀況資料)
                Dim flagIsEtraing As Boolean = False '(false:會儲存)
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                    '1.網(28:預設);2.現(TIMS:預設);3.通;4.推
                    Select Case Convert.ToString(dr("EnterChannel"))
                        Case "1"
                            flagIsEtraing = True
                    End Select
                End If

                '#Region "(No Use)"
                '檢核是否為網路報名(依IDNO,OCID)
                'Dim flagIsEtraing As Boolean=TIMS.Chk_ENTERTYPE2(vIDNO, vOCID1, objConn)

                '2008/7/16 因為學員有可能使用 自費報名 白天晚上班級不同的狀況，故可參訓
                ''看是否有在訓資料而且是非報名此班級的資訊 (排除 06:在職進修訓練 15:學習券 28:產業人才投資方案)
                'sql="SELECT a.StudStatus FROM "
                'sql += "(SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID Not IN (SELECT OCID FROM CLASS_CLASSINFO WHERE PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID IN ('06','15','28'))) and (StudStatus=1 or StudStatus=4) and OCID<>'" & OCID1 & "') a "
                'sql += "JOIN (SELECT * FROM Stud_StudentInfo WHERE IDNO='" & TIMS.ChangeIDNO(IDNO) & "' and Birthday='" & Birthday & "') b ON a.SID=b.SID "
                'dr=DbAccess.GetOneRow(sql, objConn)
                ''dr Is Nothing
                'If dr Is Nothing or sm.UserInfo.TPlanID="06" Or sm.UserInfo.TPlanID="15" Or sm.UserInfo.TPlanID="28" Then
                'If dr1("ClassID").ToString.Length=12 And sm.UserInfo.TPlanID="36" Then '因為當初轉入時班級ID過長所以重新製作
                '    '取"Y"+RID+流水號(2)  "YC" & Left(drA("RID"), 1) & Right(drA("PlanYear"), 2) & "0000001"
                '    Dim strClassID As String
                '    strClassID="Y" & Right(Left(dr1("ClassID").ToString, 3), 1) & Right(dr1("ClassID").ToString, 2)
                '    dr1("ClassID")=strClassID
                'End If

                'vIDNO=TIMS.ChangeIDNO(vIDNO) '前面有做過了
                Dim StudentID As String = ""
                'StudentID="" 'Dim MaxStudentID As Integer 'START 順序改寫 by AMU 2009-05-05  '因為有資料庫交易問題所以提前呼叫 
                '制定SID編號 Start Dim SID As String=""
                Dim pms2 As New Hashtable From {{"IDNO", vIDNO}}
                Dim sql2 As String = " SELECT * FROM STUD_STUDENTINFO WHERE IDNO=@IDNO" '改為IDNO為主 by AMU 2009-05-05 
                dr = DbAccess.GetOneRow(sql2, objConn, pms2)
                '查無SID 產生1個新的/ 有：用原先的
                Dim SID As String = If(dr Is Nothing, TIMS.GET_STUDENT_NEWSID(iEnterNum), Convert.ToString(dr("SID")))

                '將報名資料寫入學員資料檔內 Start
                'dr1=DbAccess.GetOneRow(sql, objConn)
                dr1 = GET_STUD_ENTERTYPE(SETID, EnterDate, SerNum)
                If dr1 Is Nothing Then
                    strMsgbox += "學員" & Label3name.Text & "參訓 資料有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                    StudentID = ""
                    Exit For
                End If

                '2020 提供期別可以不填寫，但學號仍加入01 維持一致性
                Dim v_CyclType As String = Convert.ToString(dr1("CyclType"))
                If v_CyclType = "" Then v_CyclType = TIMS.cst_Default_CyclType_forStudentID

                'select * from STUD_ENTERTRAIN2 where rownum <=10 'select esernum ,count(1) cnt ,max(seid),min(seid) from STUD_ENTERTRAIN2 group by esernum having count(1)>1
                '--刪除異常 --delete STUD_ENTERTRAIN2  where seid in (select max(seid) from STUD_ENTERTRAIN2 group by esernum having count(1)>1)
                '學號增加的方式應該是要去除前面的固定長度字串，才做流水號處理，並將流水號由原本的2碼變3碼。
                'If sm.UserInfo.TPlanID="36" AndAlso dr1("ClassID").ToString.Length=12 Then '因為當初轉入時班級ID過長所以重新製作
                '    '取"Y"+RID+流水號(2)  "YC" & Left(drA("RID"), 1) & Right(drA("PlanYear"), 2) & "0000001"
                '    Dim strClassID As String
                '    strClassID=""
                '    strClassID="Y" & Right(Left(dr1("ClassID").ToString, 3), 1) & Right(dr1("ClassID").ToString, 2)
                '    dr1("ClassID")=strClassID
                'End If

                '因為有資料庫交易問題所以提前呼叫
                'Dim msgbox As String=""
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    StudentID = TIMS.Get_TPlanID28_StudentID(dr1("Years").ToString, dr1("DistID").ToString, dr1("OrgKind").ToString, dr1("IsBusiness").ToString,
                                                             dr1("ClassID").ToString, v_CyclType, dr1("ClassCate").ToString, dr1("TMID").ToString, objConn)
                    If StudentID Is Nothing Then
                        strMsgbox += "學員" & Label3name.Text & "參訓學員號有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                        StudentID = ""
                        Exit For
                    Else
                        If StudentID = "" Then
                            strMsgbox += "學員" & Label3name.Text & "參訓學員號有誤,無法參訓(內部資料異常，請連絡系統管理人員查詢問題)!" & vbCrLf
                            StudentID = ""
                            Exit For
                        End If
                    End If
                End If
                '-- END 順序改寫 by AMU 2009-05-05

                'vIDNO=TIMS.ChangeIDNO(vIDNO) '前面有做過了
                Dim conn As SqlConnection = DbAccess.GetConnection()
                Dim Trans As SqlTransaction = DbAccess.BeginTrans(conn)
                Try
                    '將報名資料寫入學員資料檔內 Start
                    Dim iMaxStudentID As Integer = 0
                    'MaxStudentID=0
                    If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '自辦
                        StudentID = String.Format("{0}0{1}{2}", dr1("Years"), dr1("ClassID"), v_CyclType)
                        sql = "" & vbCrLf
                        sql &= " SELECT ISNULL(MAX(CONVERT(NUMERIC, REPLACE(StudentID,'" & StudentID & "',''))),0)+1 MaxNum" & vbCrLf
                        sql &= " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
                        sql &= " WHERE OCID='" & dr1("OCID1") & "' AND StudentID LIKE '" & StudentID & "%' "
                        iMaxStudentID = DbAccess.ExecuteScalar(sql, Trans)
                        StudentID = String.Concat(StudentID, Format(iMaxStudentID, "00#"))
                    Else
                        '產學訓
                        sql = "" & vbCrLf
                        sql &= " SELECT ISNULL(MAX(CONVERT(NUMERIC, SUBSTRING(StudentID, LEN(StudentID)-1, 2))),0)+1 MaxNum" & vbCrLf
                        sql &= " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
                        sql &= " WHERE OCID='" & dr1("OCID1") & "'" & vbCrLf
                        iMaxStudentID = DbAccess.ExecuteScalar(sql, Trans)
                        If Not String.IsNullOrEmpty(StudentID) AndAlso StudentID <> "" Then
                            StudentID = String.Concat(StudentID, Format(iMaxStudentID, "0#"))
                        Else
                            strMsgbox = String.Concat("學員", Label3name.Text, "參訓學員號有誤,無法參訓!") & vbCrLf
                            StudentID = ""
                            Dim ex As New Exception(strMsgbox)
                            Throw ex
                        End If
                    End If

                    If StudentID <> "" Then
                        '將資料寫入學員主檔---   Start
                        'sql=" SELECT * FROM Stud_StudentInfo WHERE SID='" & SID & "' "
                        sql = " SELECT * FROM STUD_STUDENTINFO WHERE IDNO='" & vIDNO & "'"  '改成以身分證字號判斷 2009/07/16
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        If dt.Rows.Count = 0 Then
                            '如果是新增(找不到idno) '新增1筆
                            'j=dt.Rows.Count
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("SID") = SID
                            Call sUtl_UdpateStudInfo(dr, dr1)
                        Else
                            '可能有多筆。
                            For i As Integer = 0 To dt.Rows.Count - 1
                                dr = dt.Rows(i)
                                Call sUtl_UdpateStudInfo(dr, dr1)
                            Next
                        End If
                        DbAccess.UpdateDataTable(dt, da, Trans)
                        '將資料寫入學員主檔---   End

                        '將資料寫入學員副檔---   Start
                        'sql=" SELECT * FROM Stud_SubData WHERE SID='" & SID & "' "
                        sql = " SELECT SID FROM STUD_STUDENTINFO WHERE IDNO='" & TIMS.ChangeIDNO(vIDNO) & "' " '找出SID'多筆
                        dt9 = DbAccess.GetDataTable(sql, da, Trans)

                        For z As Integer = 0 To dt9.Rows.Count - 1
                            sql = " SELECT * FROM STUD_SUBDATA WHERE SID='" & dt9.Rows(z).Item(0) & "' "
                            dt = DbAccess.GetDataTable(sql, da, Trans)
                            If dt.Rows.Count = 0 Then
                                dr = dt.NewRow()
                                dt.Rows.Add(dr)
                                dr("SID") = dt9.Rows(z)("SID") '.Item(0)' SID
                            Else
                                dr = dt.Rows(0) '已存在SID
                            End If
                            dr("Name") = dr1("Name")
                            dr("School") = dr1("School")
                            dr("Department") = dr1("Department")
                            dr("ZipCode1") = dr1("ZipCode")
                            'VARCHAR(6) numeric
                            dr("ZipCode1_6W") = If(Not IsDBNull(dr1("ZipCODE6W")), dr1("ZipCODE6W"), dr("ZipCode1_6W"))
                            dr("Address") = dr1("Address")
                            dr("Email") = dr1("Email")
                            dr("PhoneD") = dr1("Phone1")
                            dr("PhoneN") = dr1("Phone2")
                            dr("CellPhone") = dr1("CellPhone")
                            'numeric/ 'numeric
                            dr("ZipCode2") = If(Not IsDBNull(dr1("ZipCode2")), Val(dr1("ZipCode2")), dr("ZipCode2"))
                            dr("ZipCode2_6W") = If(Not IsDBNull(dr1("ZipCode2_6W")), Val(dr1("ZipCode2_6W")), dr("ZipCode2_6W"))
                            If dr1("HouseholdAddress").ToString <> "" Or Not IsDBNull(dr1("HouseholdAddress")) Then dr("HouseholdAddress") = dr1("HouseholdAddress")
                            If dr1("HandTypeID").ToString <> "" Or Not IsDBNull(dr1("HandTypeID")) Then dr("HandTypeID") = dr1("HandTypeID")
                            If dr1("HandLevelID").ToString <> "" Or Not IsDBNull(dr1("HandLevelID")) Then dr("HandLevelID") = dr1("HandLevelID")
                            If dr1("PriorWorkOrg1").ToString <> "" Or Not IsDBNull(dr1("PriorWorkOrg1")) Then dr("PriorWorkOrg1") = dr1("PriorWorkOrg1")
                            If dr1("Title1").ToString <> "" Or Not IsDBNull(dr1("Title1")) Then dr("Title1") = dr1("Title1")
                            If dr1("PriorWorkOrg2").ToString <> "" Or Not IsDBNull(dr1("PriorWorkOrg2")) Then dr("PriorWorkOrg2") = dr1("PriorWorkOrg2")
                            If dr1("Title2").ToString <> "" Or Not IsDBNull(dr1("Title2")) Then dr("Title2") = dr1("Title2")
                            If Not IsDBNull(dr1("SOfficeYM1")) Then dr("SOfficeYM1") = dr1("SOfficeYM1")
                            If Not IsDBNull(dr1("FOfficeYM1")) Then dr("FOfficeYM1") = dr1("FOfficeYM1")
                            If Not IsDBNull(dr1("SOfficeYM2")) Then dr("SOfficeYM2") = dr1("SOfficeYM2")
                            If Not IsDBNull(dr1("FOfficeYM2")) Then dr("FOfficeYM2") = dr1("FOfficeYM2")
                            If Not IsDBNull(dr1("PriorWorkPay")) Then dr("PriorWorkPay") = dr1("PriorWorkPay")
                            If Not IsDBNull(dr1("Traffic")) Then dr("Traffic") = dr1("Traffic")
                            Dim v_ShowDetail As String = "N"
                            If Convert.ToString(dr1("ShowDetail")) = "Y" Then v_ShowDetail = Convert.ToString(dr1("ShowDetail"))
                            dr("ShowDetail") = If(v_ShowDetail <> "", v_ShowDetail, Convert.DBNull)

                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                            DbAccess.UpdateDataTable(dt, da, Trans)
                        Next
                        '將資料寫入學員副檔---   End

                        '將資料寫入班級學員檔---   Start
                        sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & vOCID1 & "' AND SID='" & SID & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        Dim iSOCID As Integer = 0
                        If dt.Rows.Count = 0 Then
                            iSOCID = DbAccess.GetNewId(Trans, "CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID")
                            dr = dt.NewRow()
                            dt.Rows.Add(dr)
                            dr("SOCID") = iSOCID
                            dr("OCID") = dr1("OCID1")
                            dr("SID") = SID
                            dr("StudentID") = StudentID
                            '報到日期 開放讓使用者決定 BY AMU 2009-05-05
                            dr("EnterDate") = TIMS.Cdate2(HidEnterDate.Value)
                            'dr("EnterDate")=Common.FormatDate(Now, 2)
                            dr("OpenDate") = dr1("STDate")
                            dr("CloseDate") = dr1("FTDate")
                            dr("StudStatus") = 1
                            dr("TRNDMode") = dr1("TRNDMode")
                            dr("TRNDType") = dr1("TRNDType")
                            '1.網(28:預設);2.現(TIMS:預設);3.通;4.推
                            Dim v_EnterChannel As String = ""
                            If Convert.ToString(dr1("EnterChannel")) <> "" Then
                                v_EnterChannel = dr1("EnterChannel")
                            Else
                                '1.網(28:預設) /'2.現(TIMS:預設)
                                v_EnterChannel = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, "1", "2")
                            End If
                            dr("EnterChannel") = If(v_EnterChannel <> "", v_EnterChannel, Convert.DBNull)

                            Dim v_BudgetID As String = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, Convert.ToString(dr1("BudID")), "")
                            Dim v_SupplyID As String = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, Convert.ToString(dr1("SupplyID")), "")
                            dr("BudgetID") = If(v_BudgetID <> "", v_BudgetID, Convert.DBNull)
                            dr("SupplyID") = If(v_SupplyID <> "", v_SupplyID, Convert.DBNull)
                            'by Vicient
                            If dr1("MIdentityID").ToString <> "" Or Not IsDBNull(dr1("MIdentityID")) Then dr("MIdentityID") = dr1("MIdentityID")
                            dr("IdentityID") = dr1("IdentityID")

                            Dim v_LevelNo As Integer = If(dr1("LevelCount").ToString = "", 0, If(Int(dr1("LevelCount")) = 0, 0, If(IsNumeric(dr1("LevelName")), Int(dr1("LevelName")), 1)))
                            dr("LevelNo") = v_LevelNo
                            dr("SETID") = SETID.Value
                            dr("ETEnterDate") = EnterDate.Value
                            dr("SerNum") = SerNum.Value

                            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                                If Convert.ToString(dr1("ActNo")) <> "" OrElse Not IsDBNull(dr1("ActNo")) Then dr("ActNo") = dr1("ActNo") '產投
                            End If

                            If Not flagIsEtraing Then
                                'CLASS_STUDENTSOFCLASS
                                ' 受訓前任職資料start
                                If Convert.ToString(dr1("ActNo2")) <> "" OrElse Not IsDBNull(dr1("ActNo2")) Then dr("ActNo") = dr1("ActNo2") '職前訓練
                                If dr1("PriorWorkType3").ToString <> "" Then dr("PWType1") = dr1("PriorWorkType3")
                                If dr1("PriorWorkOrg3").ToString <> "" Then dr("PWOrg1") = dr1("PriorWorkOrg3")
                                If dr1("SOfficeYM3").ToString <> "" Then dr("SOfficeYM1") = dr1("SOfficeYM3")
                                If dr1("FOfficeYM3").ToString <> "" Then dr("FOfficeYM1") = dr1("FOfficeYM3")
                                ' 受訓前任職資料end 
                            End If

                            dr("HighEduBg") = dr1("HighEduBg")  '20090326(Milor)加入專上畢業學歷失業者
                            dr("WorkSuppIdent") = dr1("WorkSuppIdent")  '201006 AMU 是否為在職者補助身分
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                            DbAccess.UpdateDataTable(dt, da, Trans)

                            '將報名資料寫入學員服務單位(產學) by Vicient
                            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sTPlanID) > -1 Then
                                sql = " SELECT * FROM STUD_SERVICEPLACE WHERE SOCID='" & iSOCID & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans)
                                If dt.Rows.Count = 0 Then
                                    dr = dt.NewRow()
                                    dt.Rows.Add(dr)
                                    dr("SOCID") = iSOCID
                                Else
                                    dr = dt.Rows(0)
                                End If
                                dr("SERVDEPTID") = dr1("SERVDEPTID")
                                dr("JOBTITLEID") = dr1("JOBTITLEID")
                                If Not IsDBNull(dr1("AcctMode")) Then dr("AcctMode") = dr1("AcctMode")
                                If dr1("PostNo").ToString <> "" Or Not IsDBNull(dr1("PostNo")) Then dr("PostNo") = dr1("PostNo")
                                If dr1("AcctHeadNo").ToString <> "" Or Not IsDBNull(dr1("AcctHeadNo")) Then dr("AcctHeadNo") = dr1("AcctHeadNo")
                                If dr1("BankName").ToString <> "" Or Not IsDBNull(dr1("BankName")) Then dr("BankName") = dr1("BankName")
                                If dr1("AcctNo").ToString <> "" Or Not IsDBNull(dr1("AcctNo")) Then dr("AcctNo") = dr1("AcctNo")
                                If dr1("AcctExNo").ToString <> "" Or Not IsDBNull(dr1("AcctExNo")) Then dr("AcctExNo") = dr1("AcctExNo")
                                If dr1("ExBankName").ToString <> "" Or Not IsDBNull(dr1("ExBankName")) Then dr("ExBankName") = dr1("ExBankName")
                                If Not IsDBNull(dr1("FirDate")) Then dr("FirDate") = dr1("FirDate")
                                If dr1("Uname").ToString <> "" Or Not IsDBNull(dr1("Uname")) Then dr("Uname") = dr1("Uname")
                                If dr1("Intaxno").ToString <> "" Or Not IsDBNull(dr1("Intaxno")) Then dr("Intaxno") = dr1("Intaxno")
                                'If dr1("ActNo").ToString <> "" Or Not IsDBNull(dr1("ActNo")) Then dr("ActNo")=dr1("ActNo")
                                'If dr1("ActName").ToString <> "" Or Not IsDBNull(dr1("ActName")) Then dr("ActName")=dr1("ActName")
                                '產投 + 自辦 勾稽邏輯 --使用者／委訓單位資料失效
                                'dr("ActNo")=dr1("ActNo")
                                'dr("ActName")=dr1("ActName")
                                If dr1("ServDept").ToString <> "" Or Not IsDBNull(dr1("ServDept")) Then dr("ServDept") = dr1("ServDept")
                                If dr1("JobTitle").ToString <> "" Or Not IsDBNull(dr1("JobTitle")) Then dr("JobTitle") = dr1("JobTitle")
                                If Not IsDBNull(dr1("Zip")) Then dr("Zip") = dr1("Zip")
                                If dr1("Addr").ToString <> "" Or Not IsDBNull(dr1("Addr")) Then dr("Addr") = dr1("Addr")
                                If dr1("Tel").ToString <> "" Or Not IsDBNull(dr1("Tel")) Then dr("Tel") = dr1("Tel")
                                If dr1("Fax").ToString <> "" Or Not IsDBNull(dr1("Fax")) Then dr("Fax") = dr1("Fax")
                                If Not IsDBNull(dr1("SDate")) Then dr("SDate") = dr1("SDate")
                                If Not IsDBNull(dr1("SJDate")) Then dr("SJDate") = dr1("SJDate")
                                If Not IsDBNull(dr1("SPDate")) Then dr("SPDate") = dr1("SPDate")
                                '**by Milor 20081017--加入投保單位電話、地址
                                If IsDBNull(dr1("ActTel")) = False Then dr("ActTel") = dr1("ActTel")
                                If IsDBNull(dr1("ZipCode3")) = False Then dr("ActZipCode") = dr1("ZipCode3")
                                If IsDBNull(dr1("ZipCode3_6W")) = False Then dr("ActZipCode_6W") = dr1("ZipCode3_6W")
                                If IsDBNull(dr1("ActAddress")) = False Then dr("ActAddress") = dr1("ActAddress")
                                dr("ModifyAcct") = sm.UserInfo.UserID
                                dr("ModifyDate") = Now
                                '核對是否有線上報名資料.有資料才存檔
                                If Not IsDBNull(dr1("AcctMode")) Then DbAccess.UpdateDataTable(dt, da, Trans)

                                '將報名資料寫入學員參訓背景1(產學)
                                sql = " SELECT * FROM STUD_TRAINBG WHERE SOCID='" & iSOCID & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans)
                                If dt.Rows.Count = 0 Then
                                    dr = dt.NewRow()
                                    dt.Rows.Add(dr)
                                    dr("SOCID") = iSOCID
                                Else
                                    dr = dt.Rows(0)
                                End If
                                If Not IsDBNull(dr1("Q1")) Then dr("Q1") = dr1("Q1")
                                If Not IsDBNull(dr1("Q3")) Then dr("Q3") = dr1("Q3")
                                If dr1("Q3_Other").ToString <> "" Or Not IsDBNull(dr1("Q3_Other")) Then dr("Q3_Other") = dr1("Q3_Other")
                                If Not IsDBNull(dr1("Q4")) Then dr("Q4") = dr1("Q4")
                                If Not IsDBNull(dr1("Q5")) Then dr("Q5") = dr1("Q5")
                                If Not IsDBNull(dr1("Q61")) Then dr("Q61") = dr1("Q61")
                                If Not IsDBNull(dr1("Q62")) Then dr("Q62") = dr1("Q62")
                                If Not IsDBNull(dr1("Q63")) Then dr("Q63") = dr1("Q63")
                                If Not IsDBNull(dr1("Q64")) Then dr("Q64") = dr1("Q64")
                                dr("ModifyAcct") = sm.UserInfo.UserID
                                dr("ModifyDate") = Now
                                '核對是否有線上報名資料.有資料才存檔
                                If Not IsDBNull(dr1("Q1")) Then DbAccess.UpdateDataTable(dt, da, Trans)

                                '將報名資料寫入學員參訓背景2(產學)
                                If Not IsDBNull(dr1("Q2_1")) Then
                                    If dr1("Q2_1") = 1 Then
                                        sql = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' AND Q2=1 "
                                        dt = DbAccess.GetDataTable(sql, da, Trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow()
                                            dt.Rows.Add(dr)
                                            dr("SOCID") = iSOCID
                                            dr("Q2") = 1
                                            DbAccess.UpdateDataTable(dt, da, Trans)
                                        End If
                                    End If
                                End If
                                If Not IsDBNull(dr1("Q2_2")) Then
                                    If dr1("Q2_2") = 1 Then
                                        sql = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' AND Q2=2 "
                                        dt = DbAccess.GetDataTable(sql, da, Trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow()
                                            dt.Rows.Add(dr)
                                            dr("SOCID") = iSOCID
                                            dr("Q2") = 2
                                            DbAccess.UpdateDataTable(dt, da, Trans)
                                        End If
                                    End If
                                End If
                                If Not IsDBNull(dr1("Q2_33")) Then
                                    If dr1("Q2_33") = 1 Then
                                        sql = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' AND Q2=3 "
                                        dt = DbAccess.GetDataTable(sql, da, Trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow()
                                            dt.Rows.Add(dr)
                                            dr("SOCID") = iSOCID
                                            dr("Q2") = 3
                                            DbAccess.UpdateDataTable(dt, da, Trans)
                                        End If
                                    End If
                                End If
                                If Not IsDBNull(dr1("Q2_44")) Then
                                    If dr1("Q2_44") = 1 Then
                                        sql = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' AND Q2=4  "
                                        dt = DbAccess.GetDataTable(sql, da, Trans)
                                        If dt.Rows.Count = 0 Then
                                            dr = dt.NewRow()
                                            dt.Rows.Add(dr)
                                            dr("SOCID") = iSOCID
                                            dr("Q2") = 4
                                            DbAccess.UpdateDataTable(dt, da, Trans)
                                        End If
                                    End If
                                End If
                            End If
                            '將報名資料寫入學員資料檔內---產學訓計畫---End

                        End If
                        '將資料寫入班級學員檔 End
                    End If

                    '將報名資料寫入學員資料檔內 End
                    DbAccess.CommitTrans(Trans)
                    Call TIMS.CloseDbConn(conn)
                Catch ex As Exception
                    Dim strErrmsg As String = ""
                    strErrmsg += "/*  ex.ToString: */" & vbCrLf
                    strErrmsg += ex.ToString & vbCrLf
                    strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg)

                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(conn)
                    'Common.RespWrite(Me, tmpS)
                    'Response.End()
                    strMsgbox += ex.Message & vbCrLf
                    strMsgbox += ex.StackTrace & vbCrLf
                    'iEnterNum -= 1
                    strMsgbox += "學員" & Label3name.Text & "參訓失敗!" & vbCrLf
                    Exit For
                End Try
                Call TIMS.CloseDbConn(conn)
                'Else
                '    msgbox += "學員" & Label3.Text & "在訓中,無法參訓!" & vbCrLf
                'End If
            End If
        Next
    End Sub

    ''' <summary>取出己經有學員資料的人數</summary>
    ''' <param name="vOCIDValue1"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Public Shared Function Get_DatadtCS(ByVal vOCIDValue1 As String, ByRef oConn As SqlConnection) As DataTable
        Dim dtCS As DataTable = Nothing
        If vOCIDValue1 = "" Then Return dtCS

        Dim parms As New Hashtable
        parms.Add("OCID", vOCIDValue1)
        Dim sql As String = ""
        sql &= " SELECT ss.IDNO"
        'S1='Y' (有效學員)
        sql &= " ,CASE WHEN cs.StudStatus NOT IN (2,3) THEN 'Y' END S1" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS cs"
        sql &= " JOIN STUD_STUDENTINFO ss ON ss.sid=cs.sid"
        sql &= " WHERE cs.OCID=@OCID" & vbCrLf 'sql &= "   AND cs.StudStatus NOT IN (2,3) "
        dtCS = DbAccess.GetDataTable(sql, oConn, parms)
        Return dtCS
    End Function

    ''' <summary> 系統使用者，可增加儲存警告次數 </summary>
    Sub Utl_CAN_IGNORE_RULE1_CNT()
        Dim flagS1 As Boolean = flgROLEIDx0xLIDx0 ' TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        If Not flagS1 Then
            ViewState(vs_SD03001_OCID1) = Nothing
            Hid_CAN_IGNORE_RULE1_CNT.Value = ""
            Return
        End If
        If (Hid_CAN_IGNORE_RULE1_CNT.Value = "") Then
            ViewState(vs_SD03001_OCID1) = HidOCID1.Value
            Hid_CAN_IGNORE_RULE1_CNT.Value = "1"
            Return
        End If
        If (Hid_CAN_IGNORE_RULE1_CNT.Value <> "") AndAlso ViewState(vs_SD03001_OCID1) = HidOCID1.Value Then
            Hid_CAN_IGNORE_RULE1_CNT.Value = TIMS.CINT1(Hid_CAN_IGNORE_RULE1_CNT.Value) + 1
        End If
    End Sub

    ''' <summary>
    ''' 取得儲存警告次數
    ''' </summary>
    ''' <returns></returns>
    Function Get_CAN_IGNORE_RULE1_CNT() As Integer
        Dim rst As Integer = 0
        If (Hid_CAN_IGNORE_RULE1_CNT.Value = "") Then Return rst
        If (Convert.ToString(ViewState(vs_SD03001_OCID1)) <> HidOCID1.Value) Then Return rst
        rst = TIMS.CINT1(Hid_CAN_IGNORE_RULE1_CNT.Value)
        Return rst
    End Function

    ''' <summary> 完成報到 'Button2_Send/Button2_Click</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2B_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2B.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg8)
        '    Exit Sub
        'End If

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        ' fix https://jira.turbotech.com.tw/browse/TIMSC-161
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
        If drC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg3)
            Exit Sub
        End If

        'Dim strMsg As String=""
        ''https://jira.turbotech.com.tw/browse/TIMSC-153
        'If TIMS.Chk_DISASTERNG(Me, OCIDValue1.Value, objConn) Then
        '    '針對”重大災害提醒處理”功能中，若有未進行提醒處理者，則無法進行報到作業，並顯示告警訊息。
        '    strMsg="針對【重大災害提醒處理】功能，有未進行提醒處理，無法進行該作業，請先完成 重大災害提醒處理!"
        'End If
        'If strMsg <> "" Then
        '    Common.MessageBox(Me, strMsg)
        '    Exit Sub
        'End If

        If Not flgROLEIDx0xLIDx0 Then
            Dim flagInputOK14NG As Boolean = False
            Dim dtArc As DataTable = TIMS.Get_Auth_REndClass(Me, objConn) '暫時權限Table
            'https://jira.turbotech.com.tw/browse/TIMSC-161
            If TIMS.ChkIsEndDate(OCIDValue1.Value, TIMS.cst_FunID_學員參訓, dtArc) Then
                '過了使用期限 True(不可使用)   False(可使用)
                If $"{drC("InputOK14")}" <> "Y" Then
                    '開訓日後14日鎖定功能填寫!
                    flagInputOK14NG = True
                End If
            End If
            If flag_can_ignore_control Then flagInputOK14NG = False '忽略
            If flagInputOK14NG Then
                Common.MessageBox(Me, cst_msg_014)
                Exit Sub
            End If
        End If

        '取出學員數
        Dim dtCS As DataTable = Get_DatadtCS(OCIDValue1.Value, objConn)
        Dim sErrmsg As String = ""
        '[儲存前]檢核
        Call CheckData1(sErrmsg, dtCS)
        If sErrmsg <> "" Then
            Utl_CAN_IGNORE_RULE1_CNT()
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If
        If HidEnterDate.Value = "" Then
            Common.MessageBox(Me, "報到日期不可為空!")
            Exit Sub
        End If

        Dim strMsgbox As String = ""
        '儲存-STUD_SELRESULT /新增 CLASS_STUDENTSOFCLASS
        Call SaveData1(strMsgbox, dtCS)
        If strMsgbox = "" Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg1)
            'Session("Page_Error_MSG")="儲存成功!!"
        Else
            Common.MessageBox(Me, TIMS.cst_SAVENGMsg1)
            Common.MessageBox(Me, "儲存中斷，有資料錯誤，錯誤如下:" & vbCrLf & strMsgbox)
            'Session("Page_Error_MSG")="儲存成功，但有資料錯誤，錯誤如下:" & vbCrLf & msgbox
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        'Button2.Visible=False
        'Button2.Style("display")="none" '完成報到(鈕)
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objConn)
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        'Button2.Visible=False
        'Button2.Style("display")="none" '完成報到(鈕)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '異動1筆 STUD_STUDENTINFO
    Sub sUtl_UdpateStudInfo(ByRef dr As DataRow, ByRef dr1 As DataRow)
        'Dim rst As Boolean=False
        dr("IDNO") = TIMS.ChangeIDNO(dr1("IDNO"))
        dr("Birthday") = dr1("Birthday")
        dr("Name") = dr1("Name")
        Select Case Convert.ToString(dr1("PassPortNO"))
            Case "1", "2"
                dr("PassPortNO") = Convert.ToString(dr1("PassPortNO"))
            Case Else
                dr("PassPortNO") = "2"
        End Select
        dr("Sex") = dr1("Sex")
        'dr("MaritalStatus")=dr1("MaritalStatus")
        '1.已;2.未(預設)
        If Convert.ToString(dr1("MaritalStatus")) <> "" Then
            dr("MaritalStatus") = dr1("MaritalStatus")
            'dr("MaritalStatus")=Convert.DBNull ' "2"
        End If
        dr("DegreeID") = dr1("DegreeID")
        dr("GraduateStatus") = dr1("GradID")
        dr("MilitaryID") = dr1("MilitaryID")
        'dr("IsAgree")=dr1("IsAgree")
        dr("IsAgree") = If(Convert.ToString(dr1("IsAgree")).Equals("Y"), "Y", "N")
        'by Vicient
        If Not IsDBNull(dr1("RealJobless")) Then
            dr("RealJobless") = dr1("RealJobless")
        End If
        If dr1("JoblessID").ToString <> "" Or Not IsDBNull(dr1("JoblessID")) Then
            dr("JoblessID") = dr1("JoblessID")
        End If
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    'Protected Sub BtnCheckITS1_Click(sender As Object, e As EventArgs) Handles BtnCheckITS1.Click
    '    '未完成報到
    '    'Button2 /  BtnCheckITS1
    '    '1、 點選時，針對該班未完成報到學員，以身分證字號勾稽職前系統學員參訓歷史，比對這些學員於訓練期間是否同時參加職前課程。
    'End Sub

    'Button2B_Click -- Button2_Click
    'Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
    'End Sub

    'Protected Sub center_TextChanged(sender As Object, e As EventArgs) Handles center.TextChanged
    'End Sub
End Class