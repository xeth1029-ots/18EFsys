Partial Class SD_05_004_add2
    Inherits AuthBasePage

    'TIMS使用(非在職)-SD_05_004_add2
    Const cst_inline1 As String = "" '"inline"
    'Dim SOCIDValue As String = HidSOCIDValue.Value '取得學員學號。
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Const cst_str離退 As String = "離退"
    Const cst_str離退訓 As String = "離退訓"

    'Const Cst_2014規則1 As String = "2014" '離訓原因要分離退訓
    'Const Cst_2015規則1 As String = "2015" '離訓原因要分離退訓(改變排序及選項)
    'Const Cst_2016規則1 As String = "2016" '離訓原因要同時顯示'(改變排序及選項)
    'RTReasonID = TIMS.Get_RejectTReason(RTReasonID, "", objconn)
    'Const cst_reject_離 As String = "2"
    'Const cst_reject_離2 As String = "22"
    'Const cst_reject_退 As String = "3"

    ''' <summary>
    ''' 02:提前就業(訓期滿1/2以上)
    ''' </summary>
    ''' <remarks></remarks>
    Const cst_RTRID2_02 As String = "02" '02:提前就業(訓期滿1/2以上)
    Const cst_RTRID2_98 As String = "98" '98:其他(職前訓練須經分署/縣市政府專案認定)
    Const cst_RTRID3_14 As String = "14" '14:訓期未滿1/2找到工作
    Const cst_savedata1 As String = "savedata1"
    Const cst_delete1 As String = "delete1"

    'SELECT ''''+RTReasonID+':'+Reason FROM Key_RejectTReason WHERE SORT2 IS NOT NULL ORDER BY SORT2,RTReasonID
    'SELECT ''''+RTReasonID+':'+Reason FROM Key_RejectTReason WHERE SORT3 IS NOT NULL ORDER BY SORT3,RTReasonID
    'SELECT * FROM Key_RejectTReason WHERE SORT2 IS NOT NULL ORDER BY SORT2,RTReasonID
    '04:患病或遇意外傷害
    '03:遇家庭等災變事故
    '07:奉召服兵役
    '02:提前就業(訓期滿1/2以上)
    '98:其他(職前訓練須經分署/縣市政府專案認定)
    'SELECT * FROM Key_RejectTReason WHERE SORT3 IS NOT NULL ORDER BY SORT3,RTReasonID
    '01:缺課時數超過規定
    '13:參訓期間行為不檢情節重大
    '14:訓期未滿1/2找到工作
    '99:其他

    'https://tims.etraining.gov.tw/SD/05/SD_05_004_add.aspx?ID=117&Proecess=edit&&&&SLTID=87947&TMID=30&OCID=84565
    'SELECT * FROM STUD_LEAVETRAINING WHERE SLTID=87947 AND SOCID='1639461'
    'SELECT RejectTDate1,RejectTDate2 FROM CLASS_STUDENTSOFCLASS WHERE SOCID='1639461'
    'UPDATE CLASS_STUDENTSOFCLASS 
    'SET REJECTTDATE1= convert(datetime, '2015/09/14', 111)
    'WHERE SOCID='1639461'

    '提前就業
    '符合提前就業判斷
    'If TIMS.Chk_WkAheadOfSch(TrainHours.Text, hidTHoours.Value, NeedPay.SelectedValue, RTReasonID.SelectedValue) Then
    '    dr1("WkAheadOfSch") = "Y"
    'End If

    'Me.HidCanOffStudExists.Value =可做離退學員socid集合。
    '遞補規則：
    '1.在職班或產投 且 系統開放 才可用遞補功能，且該學員預算別為不補助 比例為%
    '課程時數：'900小時以下，離退訓遞補期限5天
    '課程時數：'900小時以上，離退訓遞補期限10天
    '1.若沒有 遞補  遞補期限內離退訓 為否
    '2.有 遞補  遞補期限內離退訓 選是
    '遞補期限天數 除上述 5天10天外還要在加上轄區假日即可。
    '201508修正 (show_cbRejectDayIn) 'cbRejectDayIn14
    '450小時以下為3日
    '451-900小時為5日
    '901小時以上為10日 

    Dim rqOCID As String = "" 'TIMS.ClearSQM(Request("OCID"))
    Dim rqSLTID As String = "" 'TIMS.ClearSQM(Request("SLTID"))
    Dim rqProecess As String = "" 'TIMS.ClearSQM(Request("Proecess")) 'add/edit

    'Dim FunDr As DataRow = Nothing
    Dim Days1 As Integer = 0
    Dim Days2 As Integer = 0
    Const vs_StDate As String = "_StDate" 'ViewState
    Const vs_FtDate As String = "_FtDate'"
    Const vs_search As String = "_search" 'ViewState 'Session
    Const vs_OCID As String = "_OCID"

    '修改不可再選學員
    '新增可再選學員

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        '取出設定天數檔 Start
        TIMS.Get_SysDays(Days1, Days2)
        '取出設定天數檔 End
        rqOCID = TIMS.ClearSQM(Request("OCID"))
        rqSLTID = TIMS.ClearSQM(Request("SLTID"))
        rqProecess = TIMS.ClearSQM(Request("Proecess")) 'add/edit

        '非 ROLEID=0 LID=0
        'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
        End If

        HidUseCanOff.Value = "" '可以使用離退判斷功能。1:可以 空白:不作判斷。
        trRejectDayIn14.Visible = True

        '檢查帳號的功能權限-----------------------------------Start
        Button1.Enabled = True
        'If Not au.blnCanAdds Then  ------因測試先拿掉
        '    Button1.Enabled = False
        '    TIMS.Tooltip(Button1, "無權限使用該功能")
        'End If
        '檢查帳號的功能權限-----------------------------------End

        '保留查詢字串
        If Not IsPostBack Then
            Call Create1()
            '儲存。
            Button1.Attributes("onclick") = "javascript:return chkdata() "

            NeedPay.Attributes("onchange") = "NeedPays()"
            RTReasonID2.Attributes("onclick") = "ShowOrg('2');" '(離訓原因)提前就業單位
            RTReasonID22.Attributes("onclick") = "ShowOrg('2');" '(選了「其他(職前訓練須經分署/縣市政府專案認定)」後，會出現)提前就業單位
            RTReasonID3.Attributes("onclick") = "ShowOrg('3');" '(退訓原因)提前就業單位
        End If

        btn_OCID.Style("display") = "none"
    End Sub

    '第1次載入
    Sub Create1()
        If Convert.ToString(rqOCID) = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        rbJobRelateY.Attributes("onclick") = "GetJobRelateYN();"
        rbJobRelateN.Attributes("onclick") = "GetJobRelateYN();"
        cbl_SD12008LB1 = TIMS.GET_SD12008LB1(cbl_SD12008LB1, objconn)

        Call CleanValue() '清除可能值

        '薪資級距檔代碼
        JobSalID = TIMS.Get_Salary(JobSalID, objconn)
        ddlSCJOB = TIMS.Get_SHARECJOB1(Me, ddlSCJOB, objconn) '行業類別 就業職類 2014/2016

        trJOBFIELD.Visible = False
        If TIMS.Cst_TPlanID_USE_JOBFIELD.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trJOBFIELD.Visible = True
            ddlJOBFIELD = TIMS.Get_ddlJOBFIELD1(ddlJOBFIELD) '訓後就業場域
            If TIMS.Cst_TPlanID58_USE_JOBFIELD.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '訓後就業場域58
                ddlJOBFIELD = TIMS.Get_ddlJOBFIELD58(ddlJOBFIELD)
            End If
        End If

        GetJobCode1 = TIMS.Get_GetJob(GetJobCode1, objconn) '就業原因代碼

        If Not Session(vs_search) Is Nothing Then
            ViewState(vs_search) = Session(vs_search)
            Session(vs_search) = Nothing
        End If

        Dim drCC As DataRow = Nothing
        drCC = TIMS.GetOCIDDate(rqOCID, objconn)
        If drCC Is Nothing Then
            'Request("OCID") 找不到班級 '已跳離
            Common.RespWrite(Me, "<script language='javascript'>alert('無此班資料');</script>")
            Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_004.aspx?ID=" & TIMS.Get_MRqID(Me) & "';</script>")
            Exit Sub
        End If

        'Request("OCID") 有找到班級
        RTReasonID2 = TIMS.Get_RejectTReason(Me, RTReasonID2, TIMS.cst_reject_離, objconn, "")
        'RTReasonID22 = TIMS.Get_RejectTReason(Me, RTReasonID22, TIMS.cst_reject_離2, objconn, "")
        RTReasonID3 = TIMS.Get_RejectTReason(Me, RTReasonID3, TIMS.cst_reject_退, objconn, "")

        '查詢該課程開訓日期 & 結訓日期
        Hid_STDate.Value = TIMS.Cdate3(drCC("STDate")) '開訓日期
        Hid_FTDate.Value = TIMS.Cdate3(drCC("FTDate")) '結訓日期
        TMID1.Text = "[" & drCC("TrainID") & "]" & drCC("TrainName")
        TMIDValue1.Value = drCC("TMID")
        OCID1.Text = drCC("ClassCName2")
        'If IsNumeric(drCC("CyclType")) Then
        '    If CInt(drCC("CyclType")) <> 0 Then
        '        OCID1.Text += "第" & TIMS.GetChtNum(CInt(dr("CyclType"))) & "期"
        '    End If
        'End If

        labmsg1.Text = ""
        labmsg1.Text += "　開訓日期：" & Common.FormatDate(drCC("STDATE"))
        'labmsg1.Text += "　訓練時數：" & Convert.ToString(dr("Thours"))

        OCIDValue1.Value = drCC("OCID")
        '****************
        '填入學員資料
        Call Add_Student()

        'rqProecess = TIMS.ClearSQM(rqProecess)
        Select Case rqProecess
            Case "add" '新增
                '已改為由上層選擇班級
                'If Not IsPostBack Then
                '    SOCID.Items.Add(New ListItem("請選擇班別", 0))
                'End If
            Case "edit" '查詢
                TMID1.ReadOnly = True
                OCID1.ReadOnly = True

                SOCID.Enabled = False
                btn_OCID.Visible = False
                OCID1.Enabled = False
                Call EditCreate1()
                Call EditCreate2()
                If HidSOCIDValue.Value <> "" Then
                    Select Case RTReasonID2.SelectedValue
                        Case cst_RTRID2_02 '選擇 提前就業:02
                            Call LoadData2_C9(HidSOCIDValue.Value)
                    End Select
                End If
                Call GetPayMoney()
        End Select

        Kind.Style("display") = "none"
        '自辦(內訓) 顯示 追償狀況 
        If TIMS.Get_PlanKind(Me, objconn) = "1" Then
            Kind.Style("display") = cst_inline1 '"inline"
        End If

        '02:提前就業(訓期滿1/2以上) '提前就業單位。
        trOrgData1.Style("display") = "none" '提前就業
        If RTReasonID2.SelectedValue = cst_RTRID2_02 Then
            trOrgData1.Style("display") = cst_inline1 '"inline"
        End If

        '98:其他(職前訓練須經分署/縣市政府專案認定)
        'spRTReasonID22.Style("display") = "none"
        'If RTReasonID2.SelectedValue = cst_RTRID2_98 Then
        '    spRTReasonID22.Style("display") = cst_inline1 '"inline"
        'End If
        Select Case RTReasonID22.SelectedValue
            Case "20", "21", "22", "23"
                spRTReasonID22.Style("display") = cst_inline1 '"inline"
        End Select
    End Sub

    '清除資料 (學員名單)
    Sub CleanValue()
        JobDate.Text = "" '就業單位到職日
        GetJob1.SelectedIndex = -1 '切結對象(GetJob1)
        JobOrgName.Text = "" '就業單位名稱
        BusGNO.Text = "" '勞保証字號
        JobCity.Text = "" '事業單位地址  JobCity JobZipCode Jobaddress
        JobZipCode.Value = "" '事業單位地址  JobCity JobZipCode Jobaddress
        Jobaddress.Text = "" '事業單位地址  JobCity JobZipCode Jobaddress
        JobTel.Text = "" '事業單位電話 JobTel
        BusFax.Text = "" '事業單位傳真 BusFax
        BusTitle.Text = "" '職稱 BusTitle 
        JobSalID.SelectedIndex = -1 '薪資級距 JobSalID
        hidSBID.Value = ""
        GetJobCode1.SelectedIndex = -1 '就業原因代碼 = TIMS.Get_GetJob(GetJobCode1, objconn)
        'JobCode5.Visible
        SpecTrace.Text = "" '特殊屬性訓練班次結訓學員就業追蹤情形說明
        ddlSCJOB.SelectedIndex = -1 '行業類別 
        PublicRescue.SelectedIndex = -1   '是否為公法救助關係
        'JobRelate.SelectedIndex = -1   '就業關聯性
        rbJobRelateY.Checked = False
        rbJobRelateN.Checked = False
        Call TIMS.SetCblValue(cbl_SD12008LB1, "")

    End Sub

    '鎖定不可修改
    Sub DisableOBJ()
        btnGetZip.Disabled = True
        btnGetZip.Visible = False
        btnClearJobSalID.Enabled = False '清除薪資級距
        btnClearJobSalID.Visible = False '清除薪資級距

        JobDate.Enabled = False '.Text = "" '就業單位到職日
        GetJob1.Enabled = False '.SelectedIndex = -1 '切結對象(GetJob1)
        JobOrgName.Enabled = False '.Text = "" '就業單位名稱
        BusGNO.Enabled = False '.Text = "" '勞保証字號
        JobCity.Enabled = False '.Text = "" '事業單位地址  JobCity JobZipCode Jobaddress
        JobZipCode.Disabled = True '.Value = "" '事業單位地址  JobCity JobZipCode Jobaddress
        Jobaddress.Enabled = False '.Text = "" '事業單位地址  JobCity JobZipCode Jobaddress
        JobTel.Enabled = False '.Text = "" '事業單位電話 JobTel
        BusFax.Enabled = False '.Text = "" '事業單位傳真 BusFax
        BusTitle.Enabled = False '.Text = "" '職稱 BusTitle 
        JobSalID.Enabled = False '.SelectedIndex = -1 '薪資級距 JobSalID
        'hidSBID.Value = ""
        GetJobCode1.Enabled = False '.SelectedIndex = -1 '就業原因代碼 = TIMS.Get_GetJob(GetJobCode1, objconn)
        'JobCode5.Visible
        SpecTrace.Enabled = False '.Text = "" '特殊屬性訓練班次結訓學員就業追蹤情形說明
        ddlSCJOB.Enabled = False '.SelectedIndex = -1 '行業類別 
        PublicRescue.Enabled = False '.SelectedIndex = -1   '是否為公法救助關係
        JobRelate.Enabled = False '.SelectedIndex = -1   '就業關聯性

    End Sub

    '離退訓遞補期限
    Sub show_cbRejectDayIn(ByVal Thours As String)
        Dim tmpStr As String
        Dim vMsg As String = ""
        Dim iThours As Integer = 0

        '201508修正 (show_cbRejectDayIn)
        '450小時以下為3日
        '451-900小時為5日
        '901小時以上為10日 

        '450小時以下為3日
        vMsg = "450小時以下，" & cst_str離退訓 & "遞補期限3天"
        tmpStr = "是(3天內)" & cst_str離退訓
        HidRejectDay.Value = 3

        If IsNumeric(Thours) Then
            iThours = CInt(Thours)
            Select Case iThours
                Case Is <= 450
                    '450小時以下為3日
                    vMsg = "450小時以下，" & cst_str離退訓 & "遞補期限3天"
                    tmpStr = "是(3天內)" & cst_str離退訓
                    HidRejectDay.Value = 3
                Case Is <= 900
                    '451-900小時為5日
                    vMsg = "451-900小時，" & cst_str離退訓 & "遞補期限5天"
                    tmpStr = "是(5天內)" & cst_str離退訓
                    HidRejectDay.Value = 5
                Case Is >= 901
                    '901小時以上為10日 
                    vMsg = "901小時以上，" & cst_str離退訓 & "遞補期限10天"
                    tmpStr = "是(10天內)" & cst_str離退訓
                    HidRejectDay.Value = 10
            End Select
        End If

        cbRejectDayIn14.Text = tmpStr
        If vMsg <> "" Then
            TIMS.Tooltip(cbRejectDayIn14, vMsg)
        End If
    End Sub

    '填入學員資料(依班) 
    Sub Add_Student()
        cbRejectDayIn14.Checked = False '不點選
        cbRejectDayIn14.Enabled = True '不鎖定
        labMakeSOCID.Text = "" '清空

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim sAlterMsg As String = ""
        Dim dr As DataRow = Nothing
        If OCIDValue1.Value <> "" Then
            If ViewState(vs_OCID) Is Nothing Then
                ViewState(vs_OCID) = "0"
            End If
            If ViewState(vs_OCID) <> OCIDValue1.Value Then
                ViewState(vs_OCID) = OCIDValue1.Value
                SOCID.Items.Clear()

                Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " SELECT Thours " & vbCrLf
                sql &= " ,AppliedResultM " & vbCrLf
                sql &= " ,IsClosed " & vbCrLf
                sql &= " ,STDate " & vbCrLf
                sql &= " ,FTDate " & vbCrLf
                sql &= " FROM Class_ClassInfo " & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= " AND OCID='" & OCIDValue1.Value & "'" & vbCrLf
                dr = DbAccess.GetOneRow(sql, objconn)
                If Convert.ToString(dr("Thours")) <> "" Then
                    '離退訓遞補期限 顯示。
                    Call show_cbRejectDayIn(Convert.ToString(dr("Thours")))
                End If
                If dr("AppliedResultM").ToString = "Y" Then
                    sAlterMsg += "學員經費審核結果已經通過，不可新增" & vbCrLf
                End If

                Dim tmpErrMsg1 As String = TIMS.Chk_StopUseDate(Me, Days1, Days2, Convert.ToString(dr("IsClosed")), dr("FTDate"))
                If TIMS.sUtl_ChkTest() Then '測試用
                    tmpErrMsg1 = "" '測試用
                End If
                If tmpErrMsg1 <> "" Then sAlterMsg &= tmpErrMsg1
                If sAlterMsg <> "" Then
                    SOCID.Items.Clear()
                    SOCID.Items.Add(New ListItem("請選擇其他班別", 0))
                    Common.MessageBox(Me, sAlterMsg)
                    Exit Sub
                End If

                'Dim FTDate As Date
                'FTDate = Common.FormatDate(dr("FTDate"), 2)
                'If dr("IsClosed") = "Y" Then
                '    If sm.UserInfo.RoleID <= 1 Then
                '        If DateDiff(DateInterval.Day, FTDate, Now) > Days2 Then
                '            SOCID.Items.Add(New ListItem("請選擇其他班別", 0))
                '            Common.MessageBox(Me, "此班已經結訓!!")
                '            Exit Function
                '        End If
                '    Else
                '        If DateDiff(DateInterval.Day, FTDate, Now) > Days1 Then
                '            SOCID.Items.Add(New ListItem("請選擇其他班別", 0))
                '            Common.MessageBox(Me, "此班已經結訓!!")
                '            Exit Function
                '        End If
                '    End If
                'End If

                If dr("THours").ToString <> "" Then
                    If dr("THours").ToString > "0" Then
                        hidTHoours.Value = dr("THours").ToString
                        LabTHours.Text = "(本班課程總訓練時數為 " & dr("THours").ToString & "小時) "
                    End If
                End If

                Dim parms As New Hashtable
                parms.Add("OCID", OCIDValue1.Value)
                sql = ""
                sql &= " SELECT b.Name,a.SOCID,a.StudStatus "
                sql &= " FROM CLASS_STUDENTSOFCLASS a"
                sql &= " JOIN STUD_STUDENTINFO b ON b.SID=a.SID"
                sql &= " WHERE a.OCID=@OCID"

                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count = 0 Then
                    SOCID.Items.Add(New ListItem("請選擇班別", 0))
                    Common.MessageBox(Me, "查無此班學生資料!!")
                Else
                    'Dim State As String = ""
                    For Each dr In dt.Rows
                        Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(dr("StudStatus"))
                        SOCID.Items.Add(New ListItem(dr("Name") & "(" & STUDSTATUS_N & ")", dr("SOCID") & "&" & dr("StudStatus").ToString))
                    Next
                    SOCID.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, 0))
                End If
                'Call GetPayMoney()
            End If

        End If
    End Sub

    '計畫申請的資料
    Sub GetPayMoney()
        SumOfPay1.Value = ""
        If RejectTDate.Text = "" Then RejectTDate.Text = TIMS.Cdate3(Today)
        If RejectTDate.Text <> "" Then RejectTDate.Text = TIMS.Cdate3(RejectTDate.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.CostID" & vbCrLf
        sql &= " ,ROUND(IsNull(a.OPrice,1)*IsNull(a.ItemCost,1),2) OPICost" & vbCrLf
        sql &= " FROM Plan_CostItem a" & vbCrLf
        sql &= " JOIN Class_ClassInfo b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNO=b.SeqNO" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.CostID IN ('01','02')" & vbCrLf 'sql += " and a.CostID IN ('01','02','03')" & vbCrLf
        sql &= " and b.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Dim iTuition As Double = 0 '學雜費
        Dim iMaterialFee As Double = 0 '材料費
        'Dim iInsurance As Double = 0 '保險費
        For Each dr1 As DataRow In dt.Rows
            Select Case dr1("CostID")
                Case "01"
                    iTuition += Val(dr1("OPICost")) '.ToString
                Case "02"
                    iMaterialFee += Val(dr1("OPICost")) '.ToString
                    'Case "03"
                    '    iInsurance += Val(dr1("OPrice")) '.ToString
            End Select
        Next

        sql = ""
        sql &= " SELECT a.STDate"
        sql &= " ,a.FDDate"
        sql &= " ,a.TNum"
        sql &= " ,a.Thours"
        sql &= " FROM Plan_PlanInfo a "
        sql &= " JOIN Class_ClassInfo b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNO=b.SeqNO"
        sql &= " WHERE 1=1 AND OCID='" & OCIDValue1.Value & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg4)
            Exit Sub
        End If
        'Dim total_days As Integer = 0
        'Dim now_days As Integer = 0
        '一、學雜費：
        '(一)全期訓練時數四百五十小時(含)以下之訓練班次：
        '開訓三日(含)以內離、退訓者，免賠償；
        '開訓逾三日至五日(含)以內離、退訓者，應賠 償全 期費用之四分之一；
        '開訓逾五日離、退訓者，應賠償全期費用之全額。
        '(二)全期訓練時數四百五十一小時至九百小時之訓練班次：開訓五日(含)以內離、退訓者，免賠償；開訓逾五日至十日(含)以內離、退訓者，應 賠償 全期費用之四分之一；開訓逾十日離、退訓者，應賠償全期費用之全額。
        '(三)全期訓練時數九百零一小時(含)以上之訓練班次：開訓十日(含)以內離、退訓者，免賠償；開訓逾十日至二十日(含)以內離、退訓者，應 賠償 全期費用之四分之一；開訓逾二十日離、退訓者，應賠償全期費用之全額。
        '二、材料費：
        '(一)全期訓練時數四百五十小時以下之訓練班次：開訓逾三日離、退訓者，以訓練日數佔全期訓練日數之比例計算賠償金額。
        '(二)全期訓練時數四百五十一小時至九百小時之訓練班次：開訓逾五日離、退訓者，以訓練日數佔全期訓練日數之比例計算賠償金額。
        '(三)全期訓練時數九百零一小時以上之訓練班次：開訓逾十日離、退訓者，以訓練日數佔全期訓練日數之比例計算賠償金額。
        '有不了解的地方，請找我，謝謝!! 
        'Dim OnPersonPay As Integer

        Dim iThours As Integer = Val(dr("Thours"))
        Dim iTotalDays As Integer = DateDiff(DateInterval.Day, CDate(dr("STDate")), CDate(dr("FDDate"))) '總共受訓天數
        Dim iNowDays As Integer = DateDiff(DateInterval.Day, CDate(dr("STDate")), CDate(RejectTDate.Text)) '目前受訓天數

        Select Case iThours
            Case Is <= 450
                If iNowDays < 3 Then
                    iTuition = 0
                    iMaterialFee = 0
                End If
                If iNowDays >= 3 AndAlso iNowDays < 5 Then
                    iTuition = iTuition * 0.25 '計算學雜費
                End If
                If iNowDays >= 3 Then
                    iMaterialFee = iMaterialFee * (iNowDays / iTotalDays) '計算材料費
                End If
            Case Is <= 900
                If iNowDays < 5 Then
                    iTuition = 0
                    iMaterialFee = 0
                End If
                If iNowDays >= 5 AndAlso iNowDays < 10 Then
                    iTuition = iTuition * 0.25 '計算學雜費
                End If
                If iNowDays >= 5 Then
                    iMaterialFee = iMaterialFee * (iNowDays / iTotalDays) '計算材料費
                End If
            Case Else
                If iNowDays < 10 Then
                    iTuition = 0
                    iMaterialFee = 0
                End If
                If iNowDays >= 10 AndAlso iNowDays < 20 Then
                    iTuition = iTuition * 0.25 '計算學雜費
                End If
                If iNowDays >= 10 Then
                    iMaterialFee = iMaterialFee * (iNowDays / iTotalDays) '計算材料費
                End If
        End Select
        '學雜費+材料費
        SumOfPay1.Value = CInt(iTuition + iMaterialFee)

        'If total_days <> 0 Then
        '    If now_days < 0 Then
        '        now_days = 0
        '    End If
        '    '計算學雜費
        '    If now_days < 14 Then
        '        iTuition = iTuition * 0
        '    ElseIf now_days >= 14 And now_days < 28 Then
        '        iTuition = iTuition * 0.25
        '    Else
        '        iTuition = iTuition
        '    End If
        '    '計算材料費
        '    iMaterialFee = iMaterialFee * (now_days / total_days)
        '    '計算保險費
        '    iInsurance = iInsurance * (now_days / total_days)
        '    '學雜費+材料費+保險費
        '    SumOfPay1.Value = CInt(iTuition + iMaterialFee + iInsurance)
        'End If

    End Sub

    '查詢編輯的資料 (學員離退資訊) STUD_LEAVETRAINING / CLASS_STUDENTSOFCLASS
    Sub EditCreate1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.SLTID" & vbCrLf
        sql &= " ,a.NeedPay" & vbCrLf
        sql &= " ,a.SumOfPay" & vbCrLf
        sql &= " ,a.HadPay" & vbCrLf
        sql &= " ,a.PayStatus" & vbCrLf
        sql &= " ,a.NoClose" & vbCrLf
        sql &= " ,a.NoClose_Desc" & vbCrLf
        sql &= " ,a.Other" & vbCrLf
        sql &= " ,a.OtherDesc" & vbCrLf
        sql &= " ,d.Name" & vbCrLf
        sql &= " ,c.RejectTDate1" & vbCrLf
        sql &= " ,c.RejectTDate2" & vbCrLf
        sql &= " ,c.StudStatus" & vbCrLf
        sql &= " ,c.SOCID" & vbCrLf
        sql &= " ,c.RTReasoOther" & vbCrLf
        sql &= " ,c.TrainHours" & vbCrLf
        sql &= " ,c.JobOrgName" & vbCrLf
        sql &= " ,c.JobTel" & vbCrLf
        sql &= " ,c.JobZipCode" & vbCrLf
        sql &= " ,c.Jobaddress" & vbCrLf
        sql &= " ,c.JobDate" & vbCrLf
        sql &= " ,c.JobSalID" & vbCrLf
        sql &= " ,c.RTReasonThat" & vbCrLf '離退訓原因說明
        sql &= " ,c.RejectDayIn14" & vbCrLf
        sql &= " ,c.MakeSOCID" & vbCrLf
        sql &= " ,b.Reason" & vbCrLf
        sql &= " ,b.RTReasonID" & vbCrLf '離退原因ID:
        sql &= " ,a.note" & vbCrLf '備註欄位
        'SELECT 
        sql &= " ,c.WkAheadOfSch " & vbCrLf '提前就業
        '提前就業判斷(依目前系統輸入值)
        sql &= " ,case when ((IsNull(c.TrainHours,0)/cc.THours) >= 0.5 ) AND IsNull(a.NeedPay,'N') ='N' AND c.RTReasonID='02' then 'Y' END WkAheadOfSch2" & vbCrLf
        'JOBFIELD
        sql &= " ,c.JOBFIELD " & vbCrLf '訓後就業場域
        sql &= " FROM STUD_LEAVETRAINING a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS c on a.SOCID=c.SOCID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.OCID=c.OCID" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO d on c.SID=d.SID" & vbCrLf
        sql &= " JOIN KEY_REJECTTREASON b on c.RTReasonID=b.RTReasonID " & vbCrLf
        sql &= " WHERE a.SLTID='" & rqSLTID & "'" & vbCrLf

        Dim dr As DataRow = Nothing
        dr = DbAccess.GetOneRow(sql, objconn)

        cbRejectDayIn14.Checked = False '不點選
        cbRejectDayIn14_N.Checked = False '不點選
        cbRejectDayIn14.Enabled = True '不鎖定
        cbRejectDayIn14_N.Enabled = True '不鎖定
        labMakeSOCID.Text = "" '清空

        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        If trRejectDayIn14.Visible = True Then
            Select Case Convert.ToString(dr("RejectDayIn14"))
                Case "Y"
                    cbRejectDayIn14.Checked = True '點選
                    If TIMS.CheckRejectSOCID(Convert.ToString(dr("SOCID")), objconn) Then '若有遞補學員 則鎖定
                        cbRejectDayIn14.Enabled = False '鎖定
                        cbRejectDayIn14_N.Enabled = False '鎖定
                    End If
                    If Convert.ToString(dr("MakeSOCID")) <> "" Then
                        cbRejectDayIn14.Enabled = False '鎖定
                        cbRejectDayIn14_N.Enabled = False '鎖定
                        labMakeSOCID.Text += " 被遞補學員：" & TIMS.GetSOCIDName(Convert.ToString(dr("MakeSOCID")), objconn)
                        TIMS.Tooltip(cbRejectDayIn14, labMakeSOCID.Text)
                    End If
                Case "N"
                    cbRejectDayIn14_N.Checked = True '點選
                    cbRejectDayIn14.Enabled = False '鎖定
                    cbRejectDayIn14_N.Enabled = False '鎖定
            End Select

            '未鎖定判斷
            If cbRejectDayIn14.Enabled Then
                Dim iTmpDay As Integer = 14
                If DateDiff(DateInterval.Day, CDate(Hid_STDate.Value), CDate(Today)) > iTmpDay Then
                    cbRejectDayIn14.Enabled = False '鎖定
                    cbRejectDayIn14_N.Enabled = False '鎖定
                    Dim sTmpDay2 As String = "作業日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內完成)!"
                    TIMS.Tooltip(cbRejectDayIn14, sTmpDay2)
                    TIMS.Tooltip(cbRejectDayIn14_N, sTmpDay2)
                End If
            End If
        End If

        SLTID.Value = Convert.ToString(dr("SLTID"))
        SOCID.Items.Add(New ListItem(Convert.ToString(dr("Name")), Convert.ToString(dr("SOCID"))))
        Common.SetListItem(SOCID, dr("SOCID").ToString)
        HidSOCIDValue.Value = Convert.ToString(dr("SOCID"))

        Select Case Convert.ToString(dr("StudStatus"))
            Case TIMS.cst_reject_離
                'SELECT * FROM Key_RejectTReason where SORT2 IS NOT NULL 
                'RTReasonID2.RepeatLayout = RepeatLayout.Flow
                'RTReasonID2 = TIMS.Get_RejectTReason(Me, RTReasonID2, TIMS.cst_reject_離, objconn, Convert.ToString(dr("RTReasonID")))
                Select Case Convert.ToString(dr("RTReasonID"))
                    Case "20", "21", "22", "23"
                        Common.SetListItem(RTReasonID22, Convert.ToString(dr("RTReasonID")))
                        Common.SetListItem(RTReasonID2, cst_RTRID2_98)
                    Case Else
                        Common.SetListItem(RTReasonID2, Convert.ToString(dr("RTReasonID")))
                End Select

                Common.SetListItem(StudStatus, TIMS.cst_reject_離)
                RejectTDate.Text = ""
                If Convert.ToString(dr("RejectTDate1")) <> "" Then
                    RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
                End If

            Case TIMS.cst_reject_退
                'SELECT * FROM Key_RejectTReason where SORT3 IS NOT NULL 
                'RTReasonID3 = TIMS.Get_RejectTReason(Me, RTReasonID3, TIMS.cst_reject_退, objconn, Convert.ToString(dr("RTReasonID")))
                Common.SetListItem(RTReasonID3, Convert.ToString(dr("RTReasonID")))
                Common.SetListItem(StudStatus, TIMS.cst_reject_退)
                RejectTDate.Text = ""
                If Convert.ToString(dr("RejectTDate2")) <> "" Then
                    RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
                End If

            Case Else
                If Convert.ToString(dr("RTReasonID")) <> "" Then
                    Select Case Convert.ToString(dr("RTReasonID"))
                        Case "20", "21", "22", "23"
                            ' Common.SetListItem(RTReasonID22, Convert.ToString(dr("RTReasonID")))
                            Common.SetListItem(RTReasonID2, cst_RTRID2_98)
                        Case Else
                            Common.SetListItem(RTReasonID2, Convert.ToString(dr("RTReasonID")))
                    End Select
                    'Common.SetListItem(RTReasonID2, dr("RTReasonID"))
                    'Common.SetListItem(RTReasonID22, dr("RTReasonID"))
                    Common.SetListItem(RTReasonID3, dr("RTReasonID"))
                End If
        End Select
        '原因儲存
        HidRTReasonID.Value = Convert.ToString(dr("RTReasonID"))

        If Convert.ToString(dr("RTReasonID")) = "23" Then
            RTReasoOther2.Text = Convert.ToString(dr("RTReasoOther"))
        End If
        If Convert.ToString(dr("RTReasonID")) = "99" _
            OrElse Convert.ToString(dr("RTReasonID")) = "98" Then
            RTReasoOther2.Text = Convert.ToString(dr("RTReasoOther"))
            RTReasoOther3.Text = Convert.ToString(dr("RTReasoOther"))
        End If

        RTReasonThat.Text = dr("RTReasonThat").ToString

        '20080901 andy  edit  
        SumOfPay.Text = Convert.ToString(dr("SumOfPay"))
        HadPay.Text = Convert.ToString(dr("HadPay"))

        If Convert.ToString(dr("NeedPay")) = "Y" Or Convert.ToString(dr("NeedPay")) = "y" Then
            NeedPay.Items(1).Selected = True '應賠償為是
        Else
            NeedPay.Items(2).Selected = True '應賠償為否
            '20080901 andy edit 
            HadPay.Text = "0"
            SumOfPay.Text = "0"
            SumOfPay.Enabled = False
        End If

        'SumOfPay.Text = Convert.ToString(dr("SumOfPay"))
        'HadPay.Text = Convert.ToString(dr("HadPay"))

        TrainHours.Text = dr("TrainHours").ToString

        If Convert.IsDBNull(dr("PayStatus")) Then
            'PayStatus.SelectedValue = ""
        Else
            PayStatus.SelectedValue = dr("PayStatus")
        End If
        If Convert.IsDBNull(dr("NoClose")) Then
            'NoClose.SelectedValue = ""
        Else
            NoClose.SelectedValue = dr("NoClose")
        End If
        If Convert.IsDBNull(dr("NoClose_Desc")) Then
            NoClose_Desc.Text = ""
        Else
            NoClose_Desc.Text = dr("NoClose_Desc")
        End If
        If Convert.IsDBNull(dr("Other")) Then
            ' Other.SelectedValue = ""
        Else
            Other.SelectedValue = dr("Other")
        End If
        If Convert.IsDBNull(dr("OtherDesc")) Then
            OtherDesc.Text = ""
        Else
            OtherDesc.Text = dr("OtherDesc")
        End If

        If IsDBNull(dr("JobOrgName")) = False Then
            JobOrgName.Text = dr("JobOrgName")
        End If

        If IsDBNull(dr("JobTel")) = False Then
            JobTel.Text = dr("JobTel")
        End If

        If IsDBNull(dr("JobZipCode")) = False Then
            JobZipCode.Value = dr("JobZipCode")
            JobCity.Text = TIMS.Get_ZipName(JobZipCode.Value, objconn)
        End If
        If IsDBNull(dr("Jobaddress")) = False Then
            Jobaddress.Text = dr("Jobaddress")
        End If

        If IsDBNull(dr("JobDate")) = False Then
            JobDate.Text = Common.FormatDate(dr("JobDate"))
        End If

        If IsDBNull(dr("JobSalID")) = False Then
            Common.SetListItem(JobSalID, dr("JobSalID"))
        End If

        'CLASS_STUDENTSOFCLASS.JOBFIELD
        If trJOBFIELD.Visible Then
            '訓後就業場域
            If Convert.ToString(dr("JOBFIELD")) <> "" Then
                Common.SetListItem(ddlJOBFIELD, dr("JOBFIELD"))
            End If
        End If

        '20080815  andy  新增備註欄位
        '--------------------------------
        If Convert.IsDBNull(dr("note")) Then
            tb_note.Text = ""
        Else
            tb_note.Text = dr("note")
        End If
        '--------------------------------
    End Sub

    '查詢班級資訊 CLASS_CLASSINFO
    Sub EditCreate2()
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT b.JobID"
        sql &= " ,b.JobName"
        sql &= " ,b.TrainID"
        sql &= " ,b.TrainName"
        sql &= " ,a.THours"
        sql &= " ,a.TMID"
        sql &= " ,a.ClassCName"
        sql &= " ,a.CyclType "
        sql &= " ,a.OCID "
        sql &= " ,a.AppliedResultM"
        sql &= " ,a.IsClosed"
        sql &= " ,a.STDate"
        sql &= " ,a.FTDate"
        sql &= " FROM CLASS_CLASSINFO a "
        sql &= " JOIN Key_TrainType b ON a.TMID=b.TMID "
        sql &= " WHERE a.OCID='" & rqOCID & "'"
        Dim dr As DataRow = Nothing
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        '學員就業狀況作業_就業關聯性欄位增加
        Dim flag_CanUseJobRelateTypePULB As Boolean = False
        flag_CanUseJobRelateTypePULB = TIMS.Get_JobRelateTypePULB(Me, TIMS.Cdate2(dr("FTDate")))
        trJobRelate1.Visible = True
        trJobRelate2.Visible = False
        If flag_CanUseJobRelateTypePULB Then
            trJobRelate1.Visible = False
            trJobRelate2.Visible = True
        End If

        If dr("THours").ToString <> "" Then
            If dr("THours").ToString > "0" Then
                hidTHoours.Value = dr("THours").ToString
                LabTHours.Text = "(本班課程總訓練時數為 " & dr("THours").ToString & "小時) "
            End If
        End If

        TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
        TMIDValue1.Value = dr("TMID")

        OCID1.Text = TIMS.GET_CLASSNAME(Convert.ToString(dr("ClassCName")), Convert.ToString(dr("CyclType")))
        OCIDValue1.Value = dr("OCID")

        If dr("AppliedResultM").ToString = "Y" Then
            Button1.Enabled = False
            TIMS.Tooltip(Button1, "學員經費審核結果已經通過，不可修改")
        End If
        If dr("IsClosed") = "Y" Then
            If sm.UserInfo.RoleID <= 1 Then
                '系統管理者以上權限
                If DateDiff(DateInterval.Day, dr("FTDate"), Now) > Days2 Then
                    'Button1.Visible = False
                    Button1.Enabled = False
                    TIMS.Tooltip(Button1, "超過結訓日期" & Days2 & "天，停用儲存功能")
                    If TIMS.sUtl_ChkTest() Then '測試用
                        Button1.Enabled = True '測試
                        TIMS.Tooltip(Button1, "超過結訓日期(測試中!!)") '測試
                    End If
                End If
            Else
                '非系統管理者以上權限
                If DateDiff(DateInterval.Day, dr("FTDate"), Now) > Days1 Then
                    'Button1.Visible = False
                    Button1.Enabled = False
                    TIMS.Tooltip(Button1, "超過結訓日期" & Days1 & "天，停用儲存功能")
                    If TIMS.sUtl_ChkTest() Then '測試用
                        Button1.Enabled = True '測試
                        TIMS.Tooltip(Button1, "超過結訓日期(測試中!!)") '測試
                    End If
                End If
            End If
        End If
        ViewState(vs_OCID) = dr("OCID") '資安檢核
    End Sub

    '儲存前檢核1
    Function CheckData1(ByRef sErrmsg As String) As Boolean
        Dim rst As Boolean = True
        sErrmsg = ""

        If HidvStatus.Value <> "" Then
            Common.SetListItem(StudStatus, HidvStatus.Value)
        End If
        Select Case StudStatus.SelectedValue
            Case TIMS.cst_reject_離
            Case TIMS.cst_reject_退
            Case Else
                sErrmsg &= "未選擇!!" & cst_str離退訓 & vbCrLf
                'sErrmsg &= "輸入資料異常，請重新查詢!" & vbCrLf
                Return False
        End Select
        If rqOCID = "" Then
            sErrmsg &= "(無班級資料)輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If
        If RTReasonID2.SelectedValue = "" _
            AndAlso RTReasonID3.SelectedValue = "" Then
            '沒有選到任1筆
            sErrmsg &= "未選擇!!" & cst_str離退訓 & "原因" & vbCrLf
            'sErrmsg &= "輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If
        HidSOCIDValue.Value = TIMS.ClearSQM(HidSOCIDValue.Value)
        Dim SOCIDValue As String = HidSOCIDValue.Value '取得學員學號。
        SOCIDValue = TIMS.ClearSQM(SOCIDValue)
        If SOCIDValue = "" Then
            sErrmsg &= "(無學員資料)輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select cs.SOCID" & vbCrLf
        sql &= " ,cs.StudStatus" & vbCrLf
        sql &= " from CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " join Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql += " and cs.StudStatus in (2,3) " & vbCrLf
        sql &= " and cs.ocid =@OCID " & vbCrLf
        sql &= " and cs.SOCID =@SOCID " & vbCrLf
        Dim sCmd3 As New SqlCommand(sql, objconn)

        Dim dt3 As New DataTable
        Call TIMS.OpenDbConn(objconn)
        With sCmd3
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(rqOCID)
            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SOCIDValue)
            dt3.Load(.ExecuteReader())
        End With
        If dt3.Rows.Count <> 1 Then
            sErrmsg &= "(查無學員資料)輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        '20080923 andy 儲存時狀態為新增判斷該學員離退訓狀態
        If rqProecess = "add" Then
            sql = "" & vbCrLf
            sql &= " select cs.SOCID,cs.StudStatus" & vbCrLf
            sql &= " from CLASS_STUDENTSOFCLASS cs" & vbCrLf
            sql &= " join Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
            sql &= " where 1=1" & vbCrLf
            sql &= " and cs.StudStatus in (2,3) " & vbCrLf
            sql &= " and cs.ocid =@OCID " & vbCrLf
            sql &= " and cs.SOCID =@SOCID " & vbCrLf
            Dim sCmd As New SqlCommand(sql, objconn)

            Dim dt As New DataTable
            Call TIMS.OpenDbConn(objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.BigInt).Value = Val(rqOCID)
                .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SOCIDValue)
                dt.Load(.ExecuteReader())
            End With

            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)

                Dim status As String = If(dr("StudStatus") = 2, "離訓", If(dr("StudStatus") = 3, "退訓", ""))
                'Common.MessageBox(Me, "此學員已是" & status & "狀態無法變更!")
                If status <> "" Then sErrmsg &= String.Concat("此學員已是", status, "狀態無法變更!") & vbCrLf

                Return False 'Exit Function
            End If
        End If

        '判開「離退訓日期」不可超過 開訓日期 & 結訓日期
        If Convert.ToDateTime(Hid_STDate.Value) > Convert.ToDateTime(RejectTDate.Text) Then
            sErrmsg &= cst_str離退訓 & "日期不可小於開訓日期(" & Hid_STDate.Value & ")!" & vbCrLf
            Return False 'Exit Function
        ElseIf Convert.ToDateTime(Hid_FTDate.Value) < Convert.ToDateTime(RejectTDate.Text) Then
            sErrmsg &= cst_str離退訓 & "日期不可大於結訓日期(" & Hid_FTDate.Value & ")!" & vbCrLf
            Return False 'Exit Function
        End If
        TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
        If (IsNumeric(TrainHours.Text) = False) Then
            sErrmsg &= "請於「實際參訓時數」欄位填入數字資料！" & vbCrLf
            Return False 'Exit Function
        End If
        If tb_note.Text.Length > tb_note.MaxLength Then
            sErrmsg &= "「備註] 欄位字數超出最大限制256字" & vbCrLf
            Return False 'Exit Function
        End If

        '遞補規則：,'1.若沒有 遞補  遞補期限內離退訓 為否,'2.有 遞補  遞補期限內離退訓且為14天內 為是,'Dim errMsg As String = "",'errMsg = ""
        Dim chkRejectDayIn14_flag As Boolean = False '是否驗證 遞補期限內離退訓 預設為false
        '有顯示才須判斷
        If trRejectDayIn14.Visible = True Then
            '驗證勾選。
            If Not cbRejectDayIn14.Checked AndAlso Not cbRejectDayIn14_N.Checked Then sErrmsg &= "請勾選，遞補期限內離退訓" & vbCrLf
            If cbRejectDayIn14.Checked AndAlso cbRejectDayIn14_N.Checked Then sErrmsg &= "遞補期限內離退訓 ，是／否請勾選1個" & vbCrLf
            If sErrmsg <> "" Then Return False 'Exit Function 'Common.MessageBox(Page, errMsg) 
            '勾是進入驗證
            If cbRejectDayIn14.Checked Then chkRejectDayIn14_flag = True '是否驗證 遞補期限內離退訓 True 
        End If

        '是否驗證 遞補期限內離退訓 True
        If chkRejectDayIn14_flag Then
            Dim iTmpDay As Integer = 3 '5
            Dim sTmpDay2 As String = ""
            If HidRejectDay.Value <> "" Then iTmpDay = HidRejectDay.Value

            Dim flagOver14 As Boolean = False 'True:已超過14天(或系統規定天數)或未在14天內。
            If Not flagOver14 Then
                '離退日期早開訓日
                If DateDiff(DateInterval.Day, CDate(Hid_STDate.Value), CDate(RejectTDate.Text)) < 0 Then
                    sTmpDay2 = cst_str離退 & "日期早開訓日!!"
                    flagOver14 = True
                End If
            End If

            If Not flgROLEIDx0xLIDx0 Then
                If Not flagOver14 Then
                    'Sys_Holiday
                    sql = "" & vbCrLf
                    sql &= " select 'x'" & vbCrLf
                    sql &= " from Sys_Holiday " & vbCrLf
                    sql &= " where RID =@RID" & vbCrLf
                    sql &= " AND HolDate>=@StDate " & vbCrLf
                    sql &= " AND HolDate<=@RejectTDate " & vbCrLf
                    Dim sCmd As New SqlCommand(sql, objconn)

                    Dim dtH As New DataTable
                    Call TIMS.OpenDbConn(objconn)
                    With sCmd
                        .Parameters.Clear()
                        If Len(sm.UserInfo.RID) = 1 Then
                            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                        Else
                            .Parameters.Add("RID", SqlDbType.VarChar).Value = Left(sm.UserInfo.RID, 1)
                        End If
                        .Parameters.Add("StDate", SqlDbType.DateTime).Value = CDate(Hid_STDate.Value)
                        .Parameters.Add("RejectTDate", SqlDbType.DateTime).Value = CDate(RejectTDate.Text)
                        'dtH = New DataTable
                        dtH.Load(.ExecuteReader())
                    End With


                    '離退日期超過14天 開訓日
                    If DateDiff(DateInterval.Day, CDate(Hid_STDate.Value), CDate(RejectTDate.Text)) > (iTmpDay + dtH.Rows.Count) Then
                        sTmpDay2 = cst_str離退訓 & "日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內)!"
                        flagOver14 = True
                    End If
                End If

                If Not flagOver14 Then
                    iTmpDay = 14
                    If DateDiff(DateInterval.Day, CDate(Hid_STDate.Value), CDate(Today)) > iTmpDay Then
                        sTmpDay2 = "作業日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內完成)!"
                        flagOver14 = True
                    End If
                End If
            End If

            If flagOver14 Then
                'Common.MessageBox(Page, sTmpDay2)
                'Exit Function
                sErrmsg &= sTmpDay2
                Return False
            End If
        End If
        'RTReasonThat.Text
        '發生類型 'System.Web.HttpUnhandledException' 的例外狀況。 ---> System.Data.OracleClient.OracleException (0x80131938): ORA-12899: 資料欄 "DBO_TIMS"."CLASS_STUDENTSOFCLASS"."RTREASONTHAT" 的值太大 (實際: 271, 最大值: 255)
        '於 Turbo2.DbAccess.UpdateDataTable(DataTable objTable, SqlDataAdapter objAdapter, SqlTransaction objTransaction)
        '於 Turbo2.DbAccess.UpdateDataTable(DataTable objTable, SqlDataAdapter objAdapter)
        '於 TIMS.SD_05_004_add.Button1_Click(Object sender, EventArgs e) 於 D

        'errMsg = ""
        TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
        If TrainHours.Text <> "" Then
            Dim v1 As Double = TIMS.VAL1(TrainHours.Text)
            'Dim v2 As Double = CInt(TIMS.VAL1(TrainHours.Text))
            'If Not TIMS.VAL1_Equal(v1, v2) Then sErrmsg &= "請於「實際參訓時數」欄位填入整數數字資料！" & vbCrLf
            Dim iTHours As Double = TIMS.VAL1(hidTHoours.Value)
            If Not TIMS.IsNumeric1(TrainHours.Text) Then
                sErrmsg &= "請於「實際參訓時數」欄位填入數字資料！" & vbCrLf
            ElseIf Not TIMS.IsDivisibleByHalf(v1) Then
                sErrmsg &= "請於「實際參訓時數」欄位填入可整除0.5的數字！" & vbCrLf
            ElseIf v1 < 0 Then
                sErrmsg &= "請於「實際參訓時數」欄位填入大於等於0的數字！" & vbCrLf
            ElseIf v1 > iTHours Then
                sErrmsg &= "請於「實際參訓時數」欄位填入小於等於訓練時數！" & vbCrLf
            End If
        End If

        Dim tmpStr As String = ""
        RTReasoOther2.Text = TIMS.ClearSQM(RTReasoOther2.Text)
        If RTReasonID2.SelectedValue = "98" Then
            Select Case RTReasonID22.SelectedValue
                Case "20", "21", "22"
                Case "23"
                    tmpStr = RTReasoOther2.Text
                    If tmpStr = "" Then
                        sErrmsg &= "若選其他，其他說明為必填。" & vbCrLf
                    End If
                Case Else
                    sErrmsg &= "選了「其他(職前訓練須經分署/縣市政府專案認定)」，請再選擇原因。" & vbCrLf
            End Select
        End If
        RTReasoOther3.Text = TIMS.ClearSQM(RTReasoOther3.Text)
        If RTReasonID3.SelectedValue = "99" Then
            tmpStr = RTReasoOther3.Text
            If tmpStr = "" Then
                sErrmsg &= "若選其他，其他說明為必填。" & vbCrLf
            End If
        End If

        If RTReasonThat.Text <> "" Then
            If Len(RTReasonThat.Text) > 255 Then
                sErrmsg &= "離退訓原因長度範圍 大於指定範圍255" & vbCrLf
            End If
        End If
        'If errMsg <> "" Then
        '    Common.MessageBox(Page, errMsg)
        '    Exit Function
        'End If
        If sErrmsg <> "" Then
            Return False
        End If

        '(SELECT * FROM Key_RejectTReason) '選擇 提前就業:02
        '「提前就業」已明定需於訓期1/2以後就業者，才算是提前就業。
        '故請針對系統中離訓選項為「提前就業」該項目加上程式邏輯，如需選擇該選項
        '該名學員退訓日時需已超過訓練期間1/2以後方才能勾選該選項，於儲存時判斷。
        Select Case RTReasonID2.SelectedValue
            Case cst_RTRID2_02 '選擇 提前就業:02
                If JobDate.Text = "" Then sErrmsg &= "請輸入就業單位到職日" & vbCrLf
                Select Case GetJob1.SelectedValue
                    Case "1", "2" 'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
                    Case ""
                        sErrmsg &= "請選擇切結對象" & vbCrLf
                    Case Else
                        sErrmsg &= "切結對象只能選擇 雇主切結或學員切結!" & vbCrLf
                End Select
                If JobOrgName.Text = "" Then sErrmsg &= "請輸入就業單位名稱" & vbCrLf
                If JobZipCode.Value = "" Then sErrmsg &= "請選擇就業單位郵遞區號" & vbCrLf
                If Jobaddress.Text = "" Then sErrmsg &= "請輸入就業單位地址" & vbCrLf
                If JobTel.Text = "" Then sErrmsg &= "請輸入就業單位電話" & vbCrLf
                If JobSalID.SelectedValue = "" Then sErrmsg &= "請選擇就業薪資級距" & vbCrLf
                '就業單位到職日(JobDate),'切結對象(GetJob1),'就業單位名稱(JobOrgName),'事業單位地址 JobCity    JobZipCode Jobaddress,'事業單位電話(JobTel),'薪資級距(JobSalID)
                If sErrmsg <> "" Then Return False

                TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
                hidTHoours.Value = TIMS.ClearSQM(hidTHoours.Value)
                Dim iTrainHours As Integer = TIMS.VAL1(TrainHours.Text)
                Dim iTHours As Integer = TIMS.VAL1(hidTHoours.Value)
                If Not TIMS.Chk_WkAheadOfSch(iTrainHours, iTHours, NeedPay.SelectedValue, RTReasonID2.SelectedValue) Then
                    sErrmsg &= "該學員 不符合 提前就業認定原則，請重新確認輸入資料。" & vbCrLf
                End If
                Dim iTrainHoursB As Double = TIMS.VAL1(TrainHours.Text)
                Dim iTHoursB As Double = TIMS.VAL1(hidTHoours.Value)
                If (iTrainHoursB / iTHoursB) < 0.5 Then
                    sErrmsg &= "該學員 離退訓原因為提前就業(訓期滿1/2以上)，離退訓日需超過訓期1/2以上!!(訓練時數:" & iTHoursB & ")" & vbCrLf
                End If
        End Select

        '(SELECT * FROM Key_RejectTReason) '選擇 14:訓期未滿1/2找到工作
        Select Case RTReasonID3.SelectedValue
            Case cst_RTRID3_14
                TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
                hidTHoours.Value = TIMS.ClearSQM(hidTHoours.Value)
                Dim iTrainHoursB As Double = TIMS.VAL1(TrainHours.Text)
                Dim iTHoursB As Double = TIMS.VAL1(hidTHoours.Value)
                If (iTrainHoursB / iTHoursB) >= 0.5 Then
                    sErrmsg &= "離訓原因 為訓期未滿1/2找到工作，離退訓日需未超過訓期1/2!!(訓練時數:" & iTHoursB & ")" & vbCrLf
                End If

        End Select

        If sErrmsg <> "" Then
            Return False
        End If
        'If errMsg <> "" Then,Common.MessageBox(Page, errMsg),Exit Function,End If,判斷結束,
        'Dim da1 As SqlDataAdapter = Nothing
        SOCIDValue = TIMS.ClearSQM(SOCIDValue)
        'rqSLTID = TIMS.ClearSQM(rqSLTID)
        sql = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID"
        Dim sCmd2 As New SqlCommand(sql, objconn)
        Dim dt1 As New DataTable
        TIMS.OpenDbConn(objconn)
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.BigInt).Value = Val(SOCIDValue)
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count <> 1 Then
            sErrmsg &= "(查無學員資料)資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        Select Case rqProecess
            Case "add" 'rqProecess
            Case "edit" 'rqProecess
                sql = "SELECT 'x' FROM STUD_LEAVETRAINING Where SLTID=@SLTID"
                Dim sCmd As New SqlCommand(sql, objconn)
                Dim dt As New DataTable
                Call TIMS.OpenDbConn(objconn)
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("SLTID", SqlDbType.BigInt).Value = Val(rqSLTID)
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count <> 1 Then
                    sErrmsg &= "(查無資料)傳入參數有誤,請重新操作，點選功能!!" & vbCrLf
                    Return False
                End If

        End Select

        If sErrmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存  (STUD_LEAVETRAINING)
    Sub Savedata1()
        '儲存開始
        HidSOCIDValue.Value = TIMS.ClearSQM(HidSOCIDValue.Value)
        Dim SOCIDValue As String = HidSOCIDValue.Value '取得學員學號。
        SOCIDValue = TIMS.ClearSQM(SOCIDValue)

        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update 's_TransType  = TIMS.cst_TRANS_LOG_Insert
        Dim s_TargetTable As String = "STUD_LEAVETRAINING"
        Dim s_FuncPath As String = "/SD/05/SD_05_004"
        Const cst_fWHERE As String = "SOCID={0}"
        Dim s_WHERE As String = String.Format(cst_fWHERE, SOCIDValue)

        'ADD / UPDATE STUD_LEAVETRAINING
        Dim iSLTID As Integer = 0
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As New SqlDataAdapter
        Select Case rqProecess
            Case "add"
                sql = " SELECT * FROM STUD_LEAVETRAINING WHERE SOCID='" & SOCIDValue & "'"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count = 0 Then
                    s_TransType = TIMS.cst_TRANS_LOG_Insert
                    iSLTID = DbAccess.GetNewId(objconn, "STUD_LEAVETRAINING_SLTID_SEQ,STUD_LEAVETRAINING,SLTID")
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    'STUD_LEAVETRAINING_SLTID_SEQ
                    dr("SLTID") = iSLTID 'DbAccess.GetNewId(objconn, "STUD_LEAVETRAINING_SLTID_SEQ,STUD_LEAVETRAINING,SLTID")
                    dr("SOCID") = SOCIDValue
                Else
                    dr = dt.Rows(0)
                    iSLTID = dr("SLTID")
                End If

            Case "edit"
                sql = "SELECT * FROM STUD_LEAVETRAINING Where SLTID='" & rqSLTID & "'"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                dr = dt.Rows(0)
                iSLTID = dr("SLTID")
        End Select

        Select Case NeedPay.SelectedValue
            Case "Y", "N"
                dr("NeedPay") = NeedPay.SelectedValue 'Y/N
            Case Else
                dr("NeedPay") = "N"
        End Select
        If SumOfPay.Text = "" Then
            dr("SumOfPay") = Convert.DBNull
        Else
            dr("SumOfPay") = SumOfPay.Text
        End If
        If HadPay.Text = "" Then
            dr("HadPay") = Convert.DBNull
        Else
            dr("HadPay") = HadPay.Text
        End If

        If PayStatus.SelectedValue <> "" Then '追償狀況
            dr("PayStatus") = PayStatus.SelectedValue
        Else
            dr("PayStatus") = Convert.DBNull
        End If

        If NoClose.SelectedValue = "" Then '追償狀況_未結案原因
            dr("NoClose") = Convert.DBNull
        Else
            dr("NoClose") = NoClose.SelectedValue
        End If

        If NoClose_Desc.Text = "" Then
            dr("NoClose_Desc") = Convert.DBNull
        Else
            dr("NoClose_Desc") = NoClose_Desc.Text
        End If

        If Other.SelectedValue = "" Then '追償狀況_其他原因
            dr("Other") = Convert.DBNull
        Else
            dr("Other") = Other.SelectedValue
        End If

        '新增備註欄位
        If tb_note.Text = "" Then ' 處理進度
            dr("note") = Convert.DBNull
        Else
            dr("note") = tb_note.Text
        End If
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Dim htPP As New Hashtable
        htPP.Clear()
        htPP.Add("TransType", s_TransType)
        htPP.Add("TargetTable", s_TargetTable)
        htPP.Add("FuncPath", s_FuncPath)
        htPP.Add("s_WHERE", s_WHERE)
        TIMS.SaveTRANSLOG(sm, objconn, dr, htPP)

        DbAccess.UpdateDataTable(dt, da) 'STUD_LEAVETRAINING

        'update CLASS_STUDENTSOFCLASS
        Dim dt1 As DataTable = Nothing
        Dim dr1 As DataRow = Nothing
        Dim da1 As SqlDataAdapter = Nothing
        sql = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID='" & SOCIDValue & "'"
        dt1 = DbAccess.GetDataTable(sql, da1, objconn)
        If dt1.Rows.Count = 1 Then
            dr1 = dt1.Rows(0)
            '開放時才可修改
            If trRejectDayIn14.Visible = True Then
                '作用中。
                If cbRejectDayIn14.Enabled = True Then
                    dr1("RejectDayIn14") = Convert.DBNull '(兩週內)離退訓
                    If cbRejectDayIn14.Checked Then
                        dr1("RejectDayIn14") = "Y" '(兩週內)離退訓
                    End If
                    If cbRejectDayIn14_N.Checked Then
                        dr1("RejectDayIn14") = "N" '(兩週內)離退訓(非)
                    End If
                End If
            End If

            dr1("StudStatus") = StudStatus.SelectedValue
            Select Case StudStatus.SelectedValue
                Case TIMS.cst_reject_離
                    dr1("WkAheadOfSch") = Convert.DBNull '其他狀況為非提前就業者
                    If TrainHours.Text <> "" Then
                        '符合提前就業人數者  dr1("WkAheadOfSch") = "Y"
                        If Not IsNumeric(TrainHours.Text) Then TrainHours.Text = "0" '檢測數字異常設為0 
                        If Not IsNumeric(hidTHoours.Value) Then hidTHoours.Value = "0" '檢測數字異常設為0                        
                        '符合提前就業判斷
                        If TIMS.Chk_WkAheadOfSch(TrainHours.Text, hidTHoours.Value, NeedPay.SelectedValue, RTReasonID2.SelectedValue) Then dr1("WkAheadOfSch") = "Y"
                    End If
                    dr1("RejectTDate1") = RejectTDate.Text
                    dr1("RejectTDate2") = Convert.DBNull
                    '(離訓原因) 離退訓原因
                    Select Case RTReasonID2.SelectedValue
                        Case "98"
                            Select Case RTReasonID22.SelectedValue
                                Case "20", "21", "22", "23"
                                    dr1("RTReasonID") = RTReasonID22.SelectedValue
                                Case Else
                                    dr1("RTReasonID") = "98"
                            End Select
                        Case Else
                            dr1("RTReasonID") = RTReasonID2.SelectedValue
                    End Select
                    If RTReasoOther2.Text <> "" Then
                        dr1("RTReasoOther") = RTReasoOther2.Text
                    Else
                        dr1("RTReasoOther") = Convert.DBNull
                    End If

                Case TIMS.cst_reject_退
                    dr1("WkAheadOfSch") = Convert.DBNull '其他狀況為非提前就業者

                    dr1("RejectTDate1") = Convert.DBNull
                    dr1("RejectTDate2") = RejectTDate.Text
                    '(退訓原因) 離退訓原因
                    dr1("RTReasonID") = RTReasonID3.SelectedValue
                    If RTReasoOther3.Text <> "" Then
                        dr1("RTReasoOther") = RTReasoOther3.Text
                    Else
                        dr1("RTReasoOther") = Convert.DBNull
                    End If

            End Select

            If RTReasonThat.Text <> "" Then dr1("RTReasonThat") = RTReasonThat.Text Else dr1("RTReasonThat") = Convert.DBNull
            JobOrgName.Text = TIMS.ClearSQM(JobOrgName.Text)
            Select Case RTReasonID2.SelectedValue
                Case cst_RTRID2_02
                    'CLASS_STUDENTSOFCLASS
                    dr1("JobOrgName") = JobOrgName.Text '(必填)
                    dr1("JobTel") = JobTel.Text '(必填)
                    dr1("JobZipCode") = JobZipCode.Value '(必填)
                    dr1("Jobaddress") = Jobaddress.Text '(必填)
                    Dim myValue As String = ""
                    Try
                        myValue = CDate(Common.FormatDate(JobDate.Text)).ToString("yyyy/MM/dd")
                    Catch ex As Exception
                    End Try
                    If myValue <> "" Then
                        dr1("JobDate") = CDate(myValue) '(必填)
                    Else
                        dr1("JobDate") = Convert.DBNull
                    End If
                    dr1("JobSalID") = JobSalID.SelectedValue '(必填)

                    dr1("JOBFIELD") = Convert.DBNull
                    If trJOBFIELD.Visible Then
                        dr1("JOBFIELD") = If(ddlJOBFIELD.SelectedValue = "", Convert.DBNull, ddlJOBFIELD.SelectedValue)
                    End If
                Case Else
                    dr1("JobOrgName") = Convert.DBNull
                    dr1("JobTel") = Convert.DBNull
                    dr1("JobZipCode") = Convert.DBNull
                    dr1("Jobaddress") = Convert.DBNull
                    dr1("JobDate") = Convert.DBNull
                    dr1("JobSalID") = Convert.DBNull
                    dr1("JOBFIELD") = Convert.DBNull
            End Select

            dr1("TrainHours") = If(TrainHours.Text = "", Convert.DBNull, TrainHours.Text)
            dr1("ModifyAcct") = sm.UserInfo.UserID
            '建檔日期(限add用而以)
            If rqProecess = "add" Then dr1("RejectCDate") = Now.ToString("yyyy/MM/dd")
            dr1("ModifyDate") = Now

            'Dim htPP As New Hashtable
            htPP.Clear()
            htPP.Add("TransType", TIMS.cst_TRANS_LOG_Update)
            htPP.Add("TargetTable", "CLASS_STUDENTSOFCLASS")
            htPP.Add("FuncPath", s_FuncPath)
            htPP.Add("s_WHERE", s_WHERE)
            TIMS.SaveTRANSLOG(sm, objconn, dr1, htPP)

            DbAccess.UpdateDataTable(dt1, da1) 'CLASS_STUDENTSOFCLASS
        End If

        'SELECT * FROM ADP_TRNDATA WHERE 1=1 AND MODIFY_DATE >=DATEADD(DAY, getdate(), -300)
        '如果是三合一的學生，要更新狀態
        'UPDATE ADP_TRNDATA
        'sql = "SELECT * FROM Adp_TRNData WHERE SOCID='" & SOCIDValue & "'"
        'dt = DbAccess.GetDataTable(sql, da, objconn)
        'If dt.Rows.Count <> 0 Then
        '    dr = dt.Rows(0)
        '    dr("ARVL_STATE") = 2
        '    dr("ARVL_STP_DATE") = RejectTDate.Text
        '    Select Case RTReasonID.SelectedValue
        '        Case "99", "98"
        '            'If RTReasoOther.Text <> "" Then RTReasoOther.Text = Trim(RTReasoOther.Text)
        '            If Len(RTReasoOther.Text) > 300 Then
        '                RTReasoOther.Text = Mid(RTReasoOther.Text, 1, 300)
        '            End If
        '            dr("ARVL_STP_REASON") = RTReasoOther.Text
        '        Case Else
        '            dr("ARVL_STP_REASON") = RTReasonID.SelectedItem.Text
        '    End Select
        '    dr("SEND_DATE") = Now.Date
        '    dr("TIMSModifyDate") = Now
        '    DbAccess.UpdateDataTable(dt, da)
        'End If

        'SELECT * FROM ADP_DGTRNDATA WHERE 1=1 AND MODIFY_DATE >=DATEADD(DAY, getdate(), -300)
        'UPDATE ADP_DGTRNDATA
        'sql = "SELECT * FROM Adp_DGTRNData WHERE SOCID='" & SOCIDValue & "'"
        'dt = DbAccess.GetDataTable(sql, da, objconn)
        'If dt.Rows.Count <> 0 Then
        '    dr = dt.Rows(0)
        '    dr("ARVL_STATE") = 2
        '    dr("STOP_DATE") = RejectTDate.Text
        '    Select Case RTReasonID.SelectedValue
        '        Case "99", "98"
        '            If Len(RTReasoOther.Text) > 300 Then
        '                RTReasoOther.Text = Mid(RTReasoOther.Text, 1, 300)
        '            End If
        '            dr("STOP_REASON") = RTReasoOther.Text
        '        Case Else
        '            dr("STOP_REASON") = RTReasonID.SelectedItem.Text
        '    End Select
        '    dr("STOP_STATE") = 0
        '    dr("TIMSModifyDate") = Now
        '    DbAccess.UpdateDataTable(dt, da)
        'End If

        'SELECT count(1) cnt1 FROM ADP_GOVTRNDATA WHERE 1=1 AND MODIFY_DATE >=DATEADD(DAY, getdate(), -300)
        'UPDATE ADP_GOVTRNDATA
        sql = "SELECT * FROM ADP_GOVTRNDATA WHERE SOCID='" & SOCIDValue & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            dr("ARVL_STATE") = 2
            dr("ARVL_STP_DATE") = RejectTDate.Text
            Dim tmpREASON As String = ""
            If RTReasonID2.SelectedValue = "98" Then
                tmpREASON = RTReasoOther2.Text
                If Len(tmpREASON) > 300 Then
                    tmpREASON = Mid(tmpREASON, 1, 300)
                End If
            End If
            If RTReasonID3.SelectedValue = "99" Then
                tmpREASON = RTReasoOther3.Text
                If Len(tmpREASON) > 300 Then
                    tmpREASON = Mid(tmpREASON, 1, 300)
                End If
            End If
            dr("ARVL_STP_REASON") = Convert.DBNull
            If tmpREASON <> "" Then
                dr("ARVL_STP_REASON") = tmpREASON
            End If
            dr("TIMSModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da)
        End If

        If Session(vs_search) Is Nothing Then
            Session(vs_search) = ViewState(vs_search)
        End If

        Select Case rqProecess
            Case "add"
                Common.RespWrite(Me, "<script language='javascript'>alert('新增成功');</script>")
            Case "edit"
                Common.RespWrite(Me, "<script language='javascript'>alert('修改成功');</script>")
            Case Else
                Common.RespWrite(Me, "<script language='javascript'>alert('請檢查輸入參數!!');</script>")
                Exit Sub
        End Select

        Dim sMsg1 As String = ""
        sMsg1 = ""
        sMsg1 &= "請記得填寫相關後續作業,若該學員有課程成績,請填寫結訓成績;\n\n"
        sMsg1 &= "若該學員有申請職訓生活津貼,請於職訓生活津貼系統進行" & cst_str離退訓 & "作業。"
        Common.RespWrite(Me, "<script language='javascript'>alert('" & sMsg1 & "');</script>")
        Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_004.aspx?ID=" & TIMS.Get_MRqID(Me) & "';</script>")
    End Sub

    '(儲存) 儲存 sql('選擇 提前就業:02) STUD_GETJOBSTATE3
    Sub SaveData2_C9(ByVal iSOCID As Integer, ByVal sType As String)
        Select Case GetJob1.SelectedValue
            Case "1", "2" 'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
            Case Else
                'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
                Exit Sub
        End Select

        Dim sSysDate As String = TIMS.GetSysDate(objconn)
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        'SD_05_004_add2
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
        Try

            Call SaveData3_C9(tConn, trans, sSysDate, iSOCID, sType)

            '更新三合一資料-----------------------------------------Start
            If Now < DateAdd(DateInterval.Day, 2, CDate(Hid_FTDate.Value)) Then
                sql = "SELECT * FROM ADP_GOVTRNDATA WHERE SOCID='" & iSOCID & "'"
                dt = DbAccess.GetDataTable(sql, da, trans)
                If dt.Rows.Count > 0 Then
                    dr = dt.Rows(0)
                    Select Case sType
                        Case cst_savedata1
                            dr("JOB_STATE") = "2" '[2]：已就業
                            dr("JOB_COMPANY") = JobOrgName.Text 'BusName.Text
                        Case cst_delete1
                            dr("JOB_STATE") = Convert.DBNull
                            dr("JOB_COMPANY") = Convert.DBNull
                    End Select
                    dr("NONJOB_REASON") = Convert.DBNull
                    dr("TIMSModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                End If
            End If
            '更新三合一資料-----------------------------------------End

            Dim IsOnJobV As String = "Y"
            Select Case sType
                Case cst_delete1
                    IsOnJobV = "N"
            End Select

            'UPDATE CLASS_STUDENTSOFCLASS.IsOnJob (STUD_GETJOBSTATE3.IsGetJob)
            'Dim IsGetJobValue As Integer = 1 ' dr("IsGetJob") '1:'就業 0:'未就業 2:'不就業
            sql = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID='" & iSOCID & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            dr = dt.Rows(0)
            dr("IsOnJob") = IsOnJobV '"Y" 'Y/N 就業/未就業
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            DbAccess.CommitTrans(trans)
            'Common.MessageBox(Me, "儲存成功")
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(tConn)
            'Call TIMS.CloseDbConn(objconn)
            Throw ex
            'Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try
        Call TIMS.CloseDbConn(tConn)
    End Sub

    Sub SaveData3_C9(ByRef tConn As SqlConnection, ByRef trans As SqlTransaction, ByVal sSysDate As String, ByVal iSOCID As Integer, ByVal sType As String)
        Const cst_iCPoint9 As Integer = 9 '9:提前就業
        Dim sql As String = ""

        Select Case sType
            Case cst_delete1
                sql = ""
                sql &= " DELETE STUD_GETJOBSTATE3"
                sql &= " WHERE SOCID=@SOCID"
                sql &= " and CPoint=@CPoint"
                Dim dCmd As New SqlCommand(sql, tConn, trans)
                With dCmd
                    .Parameters.Clear()
                    .Parameters.Add("SOCID", SqlDbType.VarChar).Value = iSOCID
                    .Parameters.Add("CPoint", SqlDbType.VarChar).Value = cst_iCPoint9
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(dCmd.CommandText, trans, dCmd.Parameters)  'edit，by:20181031
                End With
                '刪除資料離開
                Exit Sub
        End Select

        Dim jobType As Integer = 0
        If trJobRelate1.Visible Then jobType = 1
        If trJobRelate2.Visible Then jobType = 2
        Dim sJobRelateVal As String = "" '""/Y/N
        Dim sJOBRELATE_Y As String = ""
        Select Case jobType
            Case 1
                sJobRelateVal = TIMS.ClearSQM(JobRelate.SelectedValue)
                'If(JobRelate.SelectedValue = "", Convert.DBNull, JobRelate.SelectedValue)
            Case 2
                If rbJobRelateY.Checked Then sJobRelateVal = "Y"
                If rbJobRelateN.Checked Then sJobRelateVal = "N"
                sJOBRELATE_Y = TIMS.GetCblValue(cbl_SD12008LB1)
        End Select

        'STUD_GETJOBSTATE3-學員就業狀況檔(每3個月,有一天)
        sql = ""
        sql &= " SELECT * FROM STUD_GETJOBSTATE3"
        sql &= " WHERE 1=1"
        sql &= " and SOCID='" & iSOCID & "'"
        sql &= " and CPoint='" & cst_iCPoint9 & "'"
        sql &= " and SBID IS NOT NULL"
        Dim da1 As SqlDataAdapter = Nothing
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, da1, trans)
        If dt1.Rows.Count > 0 Then
            '有(勞保勾稽)資料不用再做什麼了
            Exit Sub
        End If

        'STUD_GETJOBSTATE3-學員就業狀況檔(每3個月,有一天)
        sql = ""
        sql &= " SELECT * FROM STUD_GETJOBSTATE3"
        sql &= " WHERE 1=1"
        sql &= " and SOCID='" & iSOCID & "'"
        sql &= " and CPoint='" & cst_iCPoint9 & "'"
        sql &= " and SBID IS NULL"
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = DbAccess.GetDataTable(sql, da, trans)

        If dt.Rows.Count = 0 Then
            Dim iSGJID As Integer = DbAccess.GetNewId(trans, "STUD_GETJOBSTATE3_SGJID_SEQ,STUD_GETJOBSTATE3,SGJID")
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SGJID") = iSGJID
            dr("SOCID") = iSOCID
            dr("CPoint") = cst_iCPoint9

            '(到職日期) 異動日期最佳化
            'ViewState("MDATE") = ""
            If JobDate.Text <> "" Then
                JobDate.Text = TIMS.Cdate3(JobDate.Text) '.Trim
            Else
                JobDate.Text = TIMS.Cdate3(sSysDate) '.Trim
            End If
            dr("MDate") = CDate(JobDate.Text)
            'dr("CJOB_UNKEY") = Convert.DBNull
            'Dim IsGetJobValue As Integer = 1 '1:'就業 0:'未就業 2:'不就業
            dr("IsGetJob") = 1 '1:'就業 0:'未就業 2:'不就業
            dr("GetJobCode") = GetJobCode1.SelectedValue '依就業(不就業)原因代碼
            dr("NGJobDesc") = Convert.DBNull
            dr("BusName") = JobOrgName.Text 'BusName.Text
            dr("BusGNO") = If(BusGNO.Text = "", Convert.DBNull, BusGNO.Text)
            dr("BusZip") = JobZipCode.Value 'BusZip.Value
            dr("BusAddr") = Jobaddress.Text 'BusAddr.Text
            dr("BusTel") = JobTel.Text 'BusTel.Text
            dr("BusFax") = If(BusFax.Text = "", Convert.DBNull, BusFax.Text)
            dr("BusTitle") = If(BusTitle.Text = "", Convert.DBNull, BusTitle.Text)
            dr("SalID") = If(JobSalID.SelectedValue = "", Convert.DBNull, JobSalID.SelectedValue)
            Select Case GetJob1.SelectedValue
                Case "1", "2", "3" 'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
                    dr("SureItem") = GetJob1.SelectedValue
                Case Else
                    dr("SureItem") = Convert.DBNull
            End Select
            dr("SBID") = If(hidSBID.Value = "", Convert.DBNull, hidSBID.Value)
            If GetJobCode1.SelectedValue = "05" Then
                SpecTrace.Text = TIMS.ClearSQM(SpecTrace.Text)
                dr("SpecTrace") = If(SpecTrace.Text = "", Convert.DBNull, SpecTrace.Text)
            Else
                dr("SpecTrace") = Convert.DBNull
            End If
            '行業類別 
            dr("CJOB_UNKEY") = If(ddlSCJOB.SelectedValue = "", Convert.DBNull, ddlSCJOB.SelectedValue)
            '是否為公法救助關係 (Button8) (SD_12_008_search.aspx?SOCID=) (PublicRescue)
            Select Case PublicRescue.SelectedValue
                Case "Y", "N"
                    dr("PublicRescue") = PublicRescue.SelectedValue
                Case Else
                    dr("PublicRescue") = Convert.DBNull
            End Select
            '就業關聯性 (JobRelate)
            'Select Case JobRelate.SelectedValue
            '    Case "Y", "N"
            '        dr("JobRelate") = JobRelate.SelectedValue
            '    Case Else
            '        dr("JobRelate") = Convert.DBNull
            'End Select
            dr("JobRelate") = If(sJobRelateVal = "", Convert.DBNull, sJobRelateVal)
            dr("JOBRELATE_Y") = If(sJOBRELATE_Y = "", Convert.DBNull, sJOBRELATE_Y)

            If dr("Mode_").ToString = "" Then 'Mode 資料來源 1:系統;2:人工
                dr("Mode_") = 2
            End If
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)

        End If

    End Sub

    '取得提前就業資料 [STUD_GETJOBSTATE3]
    Sub LoadData2_C9(ByVal iSOCID As Integer)
        Const cst_iCPoint9 As Integer = 9 '9:提前就業
        Dim sSysDate As String = TIMS.GetSysDate(objconn)
        'STUD_GETJOBSTATE3-學員就業狀況檔(每3個月,有一天)
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM STUD_GETJOBSTATE3"
        sql &= " WHERE SOCID=@SOCID "
        sql &= " and CPoint='" & cst_iCPoint9 & "'"
        Dim sCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.Int).Value = iSOCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count = 0 Then Exit Sub
        Dim dr As DataRow = dt.Rows(0)
        hidSBID.Value = Convert.ToString(dr("SBID")) '有(勞保勾稽)資料不用再做什麼了
        'JobDate.Text = ""
        '(到職日期) 異動日期最佳化
        If Convert.ToString(dr("MDate")) <> "" Then
            JobDate.Text = TIMS.Cdate3(dr("MDate"))
        End If
        '依就業(不就業)原因代碼
        If Convert.ToString(dr("GetJobCode")) <> "" Then
            Common.SetListItem(GetJobCode1, dr("GetJobCode"))
        End If
        If Convert.ToString(dr("BusName")) <> "" Then
            JobOrgName.Text = Convert.ToString(dr("BusName"))
        End If
        If Convert.ToString(dr("BusGNO")) <> "" Then
            BusGNO.Text = Convert.ToString(dr("BusGNO"))
        End If
        If JobZipCode.Value <> "" Then
            JobZipCode.Value = Convert.ToString(dr("BusZip"))
            Dim tZipCode As String = JobZipCode.Value
            JobCity.Text = TIMS.Get_ZipName(JobZipCode.Value, objconn)
        End If
        If Convert.ToString(dr("BusAddr")) <> "" Then
            Jobaddress.Text = Convert.ToString(dr("BusAddr"))
        End If
        If Convert.ToString(dr("BusTel")) <> "" Then
            JobTel.Text = Convert.ToString(dr("BusTel"))
        End If
        If Convert.ToString(dr("BusFax")) <> "" Then
            BusFax.Text = Convert.ToString(dr("BusFax"))
        End If
        If Convert.ToString(dr("BusTitle")) <> "" Then
            BusTitle.Text = Convert.ToString(dr("BusTitle"))
        End If
        If Convert.ToString(dr("SalID")) <> "" Then
            Common.SetListItem(JobSalID, dr("SalID"))
        End If
        'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
        If Convert.ToString(dr("SureItem")) <> "" Then
            Common.SetListItem(GetJob1, dr("SureItem"))
        End If
        SpecTrace.Text = Convert.ToString(dr("SpecTrace"))
        '行業類別 
        If Convert.ToString(dr("CJOB_UNKEY")) <> "" Then
            Common.SetListItem(ddlSCJOB, dr("CJOB_UNKEY"))
        End If
        '是否為公法救助關係 (Button8) (SD_12_008_search.aspx?SOCID=) (PublicRescue)
        If Convert.ToString(dr("PublicRescue")) <> "" Then
            Common.SetListItem(PublicRescue, dr("PublicRescue"))
        End If
        '就業關聯性 (JobRelate)
        'If Convert.ToString(dr("JobRelate")) <> "" Then
        '    Common.SetListItem(JobRelate, dr("JobRelate"))
        'End If
        '就業關聯性
        rbJobRelateY.Checked = False
        rbJobRelateN.Checked = False
        If Convert.ToString(dr("JOBRELATE")) <> "" Then
            Common.SetListItem(JobRelate, Convert.ToString(dr("JOBRELATE")))
            Select Case Convert.ToString(dr("JOBRELATE"))
                Case TIMS.cst_YES
                    rbJobRelateY.Checked = True
                Case TIMS.cst_NO
                    rbJobRelateN.Checked = True
            End Select
        End If

        Call TIMS.SetCblValue(cbl_SD12008LB1, "")
        If Convert.ToString(dr("JOBRELATE_Y")) <> "" Then
            Call TIMS.SetCblValue(cbl_SD12008LB1, Convert.ToString(dr("JOBRELATE_Y")))
        End If

        Dim flag_diable1 As Boolean = False '不可修改提前就業資訊

        Select Case Convert.ToString(dr("MODE_"))
            Case "1" '1:系統判定;2:人工判定
                Common.SetListItem(GetJob1, "3")
                flag_diable1 = True
            Case "2" '1:系統判定;2:人工判定
                'flag_diable1 = False
                '有序號不可修改，請移至(學員就業狀況作業 SD_12_008 作業)
                If hidSBID.Value <> "" Then flag_diable1 = True

                'GetJob1/SureItem'1:雇主切結 2:學員切結 3:勞保勾稽
                Select Case Convert.ToString(dr("SureItem"))
                    Case "1", "2"
                    Case Else
                        '此情況應該不太可能(但如果有鎖定)
                        Common.SetListItem(GetJob1, "3")
                        flag_diable1 = True
                End Select
        End Select

        If flag_diable1 Then
            Call DisableOBJ()
        End If

    End Sub

    '儲存學生按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '1.取得學員學號。
        Dim SOCIDValue As String = "" '取得學員學號。
        SOCIDValue = Split(SOCID.SelectedValue, "&")(0)
        HidSOCIDValue.Value = SOCIDValue
        '2.檢核
        Dim sErrmsg As String = ""
        Call CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Call ScriptA1()
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        '3.儲存
        Call Savedata1()

        Select Case RTReasonID2.SelectedValue
            Case cst_RTRID2_02 '選擇 提前就業:02
                '新增/修改
                Call SaveData2_C9(SOCIDValue, cst_savedata1)
            Case Else
                '刪除
                Call SaveData2_C9(SOCIDValue, cst_delete1)
                '修改可能性
        End Select

    End Sub

    '依公式計算
    Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
        'If SumOfPay.Enabled = True Then
        '    GetPayMoney()
        '    SumOfPay.Text = SumOfPay1.Value
        '    NeedPay.SelectedIndex = 1
        'End If

        '20080901 andy edit 當「是否賠償」欄位為「否」時則「應賠金額」及「已賠金額」值帶0
        Dim SumOfPayFlag As Boolean = False
        SumOfPayFlag = False
        SumOfPay.Text = TIMS.ClearSQM(SumOfPay.Text)
        If SumOfPay.Text <> "" Then
            If IsNumeric(SumOfPay.Text) Then
                If CInt(SumOfPay.Text) <> 0 Then
                    SumOfPayFlag = True '大於零
                End If
            Else
                Common.MessageBox(Me, "賠償金額必須為數字")
                Exit Sub
            End If
        End If

        Call GetPayMoney()
        If SumOfPay1.Value <> "" Then
            SumOfPay.Text = SumOfPay1.Value
            'NeedPay.SelectedIndex = 1
            Common.SetListItem(NeedPay, "Y") '應賠償為是
        End If
    End Sub

    '清除薪資級距
    Private Sub btnClearJobSalID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearJobSalID.Click
        JobSalID.SelectedIndex = -1 '清空選項
        Call ScriptA1()
    End Sub

    '執行client javascript
    Sub ScriptA1()
        Dim strScript As String = ""
        If RTReasonID2.SelectedValue <> "" Then
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "ShowOrg('2');" + vbCrLf
            strScript += "</script>"
        End If
        If RTReasonID3.SelectedValue <> "" Then
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript += "ShowOrg('3');" + vbCrLf
            strScript += "</script>"
        End If
        If strScript <> "" Then
            TIMS.RegisterStartupScript(Me, TIMS.xBlockName, strScript)
        End If
    End Sub

    '回上一頁
    Protected Sub Btn2back_Click(sender As Object, e As EventArgs) Handles Btn2back.Click
        If Session(vs_search) Is Nothing Then
            Session(vs_search) = ViewState(vs_search)
        End If
        Dim url1 As String = String.Concat("SD_05_004.aspx?ID=", TIMS.Get_MRqID(Me)) ' Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

#Region "NO USE"
    ''提前就業計算原則: 確認 True
    'Function Chk_WkAheadOfSch(ByVal TrainHours As Single, ByVal THours As Single, ByVal NeedPay As String, ByVal RTReasonID As String) As Boolean
    '    Dim WkAheadOfSchFlag As Boolean = False
    '    WkAheadOfSchFlag = False
    '    '提前就業計算原則：1.學員實際參訓時數達總時數 1/2 以上(含)
    '    '提前就業計算原則：2.經分署專案核定免負擔退訓賠償費用者
    '    '3.(選提前就業者)
    '    If (TrainHours / THours) >= 0.5 AndAlso NeedPay = "N" AndAlso RTReasonID = "02" Then
    '        WkAheadOfSchFlag = True
    '    End If
    '    Return WkAheadOfSchFlag
    'End Function

    ''離退訓選項將不同。
    'Protected Sub StudStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles StudStatus.SelectedIndexChanged
    '    HidRTReasonID.Value = ""
    '    If RTReasonID.SelectedValue <> "" Then
    '        HidRTReasonID.Value = RTReasonID.SelectedValue
    '    End If
    '    'SELECT * FROM Key_RejectTReason
    '    'Cst_2016規則1
    '    RTReasonID.RepeatLayout = RepeatLayout.Table
    '    If sm.UserInfo.Years >= Cst_2016規則1 Then

    '    ElseIf sm.UserInfo.Years >= Cst_2015規則1 Then
    '        Select Case StudStatus.SelectedValue
    '            Case TIMS.cst_reject_離
    '                RTReasonID.RepeatLayout = RepeatLayout.Flow
    '                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, StudStatus.SelectedValue, objconn)
    '                Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '            Case TIMS.cst_reject_退
    '                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, StudStatus.SelectedValue, objconn)
    '                Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '            Case Else 'Case TIMS.cst_reject_退
    '                RTReasonID.Items.Clear()
    '                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, "", objconn)
    '                Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '                Common.MessageBox(Me, "請選擇" & cst_str離退訓 & "種類!!")
    '                Exit Sub
    '        End Select
    '    ElseIf sm.UserInfo.Years >= Cst_2014規則1 Then
    '        Select Case StudStatus.SelectedValue
    '            Case TIMS.cst_reject_離, TIMS.cst_reject_退
    '                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, StudStatus.SelectedValue, objconn)
    '                Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '            Case Else 'Case TIMS.cst_reject_退
    '                RTReasonID.Items.Clear()
    '                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, "", objconn)
    '                Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '                Common.MessageBox(Me, "請選擇" & cst_str離退訓 & "種類!!")
    '                Exit Sub
    '        End Select
    '    Else
    '        RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, "", objconn)
    '        Common.SetListItem(RTReasonID, HidRTReasonID.Value)
    '    End If
    'End Sub

    ''離退訓選項將不同。
    'Sub chk_RTReasonID23(ByVal vStudStatus As String)
    '    Select Case vStudStatus
    '        Case TIMS.cst_reject_離
    '            HidRTReasonID.Value = RTReasonID2.SelectedValue
    '            RTReasonID3.SelectedIndex = -1
    '        Case TIMS.cst_reject_退
    '            HidRTReasonID.Value = RTReasonID3.SelectedValue
    '            RTReasonID2.SelectedIndex = -1
    '    End Select
    'End Sub

    'Protected Sub RTReasonID2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RTReasonID2.SelectedIndexChanged
    '    Call chk_RTReasonID23(TIMS.cst_reject_離)
    'End Sub

    'Protected Sub RTReasonID3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RTReasonID3.SelectedIndexChanged
    '    Call chk_RTReasonID23(TIMS.cst_reject_退)
    'End Sub
#End Region

    'Protected Sub StudStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles StudStatus.SelectedIndexChanged

    'End Sub

End Class
