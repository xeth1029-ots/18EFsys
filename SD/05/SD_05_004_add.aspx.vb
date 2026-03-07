Partial Class SD_05_004_add
    Inherits AuthBasePage

    '產投使用-SD_05_004_add
    Dim SOCIDValue As String = "" '取得學員學號。
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Const cst_str離退 As String = "離退"
    Const cst_str離退訓 As String = "離退訓"
    Const Cst_2014規則1 As String = "2014" '離訓原因要分離退訓
    Const Cst_2015規則1 As String = "2015" '離訓原因要分離退訓(改變排序及選項)
    'Const Cst_2016規則1 As String = "2016" '離訓原因要同時顯示'(改變排序及選項)
    'RTReasonID = TIMS.Get_RejectTReason(RTReasonID, "", objconn)

    Dim rqOCID As String = "" 'TIMS.ClearSQM(Request("OCID"))
    Dim rqProecess As String = "" 'TIMS.ClearSQM(Request("Proecess"))
    Dim rqSLTID As String = "" 'TIMS.ClearSQM(Request("SLTID"))

    'https://tims.etraining.gov.tw/SD/05/SD_05_004_add.aspx?ID=117&Proecess=edit&&&&SLTID=87947&TMID=30&OCID=84565
    'SELECT * FROM Stud_LeaveTraining WHERE SLTID=87947 AND SOCID='1639461'
    'SELECT RejectTDate1,RejectTDate2 FROM Class_StudentsOfClass WHERE SOCID='1639461'
    'UPDATE Class_StudentsOfClass 
    'SET REJECTTDATE1= convert(datetime, '2015/09/14', 111)
    'WHERE SOCID='1639461'

    '提前就業
    '符合提前就業判斷
    'If TIMS.Chk_WkAheadOfSch(TrainHours.Text, .hidTHoours.Value, NeedPay.SelectedValue, v_RTReasonID) Then
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
    'Dim FunDr As DataRow = Nothing
    Dim Days1 As Integer = 0
    Dim Days2 As Integer = 0
    Const vs_StDate As String = "_StDate" 'ViewState
    Const vs_FtDate As String = "_FtDate'"
    Const vs_search As String = "_search" 'ViewState 'Session
    Const vs_OCID As String = "_OCID"
    Const vs_msg As String = "_msg"

    '修改不可再選學員
    '新增可再選學員
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '取出設定天數檔 Start
        TIMS.Get_SysDays(Days1, Days2)

        rqOCID = TIMS.ClearSQM(Request("OCID"))
        rqProecess = TIMS.ClearSQM(Request("Proecess"))
        rqSLTID = TIMS.ClearSQM(Request("SLTID"))

        '非 ROLEID=0 LID=0 'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
        End If

        HidUseCanOff.Value = "" '可以使用離退判斷功能。1:可以 空白:不作判斷。
        trRejectDayIn14.Visible = True
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投功能 停用遞補14天功能 用。
            trRejectDayIn14.Visible = False
            '取得可做離退的學員 by AMU 2013/11/07
            HidUseCanOff.Value = "1" '可以使用離退判斷功能。
            'HidCanOffStudExists.Value = TIMS.Get_CanOffStudExists(rqOCID, objconn)
        End If

        '保留查詢字串
        Dim drCC As DataRow = Nothing
        If Not IsPostBack Then
            '薪資級距檔代碼
            JobSalID = TIMS.Get_Salary(JobSalID, objconn)

            If Not Session(vs_search) Is Nothing Then
                ViewState(vs_search) = Session(vs_search)
                Session(vs_search) = Nothing
            End If

            If rqOCID <> "" Then drCC = TIMS.GetOCIDDate(rqOCID, objconn)
            If drCC Is Nothing Then
                'Request("OCID") 找不到班級 '已跳離
                Common.RespWrite(Me, "<script language='javascript'>alert('無此班資料');</script>")
                Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_004.aspx?ID=" & TIMS.Get_MRqID(Me) & "';</script>")
                Exit Sub
            End If

            'Request("OCID") 有找到班級
            RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
            '查詢該課程開訓日期 & 結訓日期
            ViewState(vs_StDate) = drCC("STDate") '開訓日期
            ViewState(vs_FtDate) = drCC("FTDate") '結訓日期
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '如果是產學訓計畫
                If Convert.ToString(drCC("TrainID")) <> "" AndAlso Convert.ToString(drCC("TrainName")) <> "" Then
                    TMID1.Text = "[" & Convert.ToString(drCC("TrainID")) & "]" & Convert.ToString(drCC("TrainName"))
                Else
                    TMID1.Text = "[" & Convert.ToString(drCC("JobID")) & "]" & Convert.ToString(drCC("JobName"))
                End If
            Else
                TMID1.Text = "[" & Convert.ToString(drCC("TrainID")) & "]" & Convert.ToString(drCC("TrainName"))
            End If
            'TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
            TMIDValue1.Value = drCC("TMID")

            OCID1.Text = TIMS.GET_CLASSNAME(Convert.ToString(drCC("ClassCName")), Convert.ToString(drCC("CyclType")))
            OCIDValue1.Value = drCC("OCID")

            labmsg1.Text = ""
            labmsg1.Text += "　開訓日期：" & Common.FormatDate(drCC("STDATE"))
            'labmsg1.Text += "　訓練時數：" & Convert.ToString(dr("Thours"))
            '****************
            '填入學員資料
            Call Add_Student()


            Select Case rqProecess 'Convert.ToString(Request("Proecess"))
                Case "add"
                    '已改為由上層選擇班級
                    'If Not IsPostBack Then
                    '    SOCID.Items.Add(New ListItem("請選擇班別", 0))
                    'End If
                Case "edit"
                    TMID1.ReadOnly = True
                    OCID1.ReadOnly = True

                    SOCID.Enabled = False
                    btn_OCID.Visible = False
                    OCID1.Enabled = False
                    Call EditCreate1()
                    'Call GetPayMoney()
            End Select

            Kind.Style("display") = "none"
            '自辦(內訓) 顯示 追償狀況 
            If TIMS.Get_PlanKind(Me, objconn) = "1" Then
                Kind.Style("display") = "inline"
            End If

        End If

        '20100415 andy 登入計畫為產業人才投資方案,離退訓原因,"缺課時數超過規定"、"提前就業"、"訓練成績不合格" 設定為唯讀 
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Page.RegisterStartupScript("setrdo", "<script>setRdoDisabled('RTReasonID');</script>") '
        End If

        '提前就業單位。
        org_TR.Style("display") = "none" '提前就業
        Dim v_RTReasonID As String = TIMS.GetListValue(RTReasonID)
        If v_RTReasonID = "02" Then
            org_TR.Style("display") = "inline"
        End If

        '儲存。
        Button1.Attributes("onclick") = "javascript:return chkdata() "

        'NeedPay.Attributes("onchange") = "NeedPays()"
        RTReasonID.Attributes("onclick") = "ShowOrg();" '提前就業單位
        btn_OCID.Style("display") = "none"

    End Sub

    '離退訓遞補期限
    Sub show_cbRejectDayIn(ByVal Thours As String)
        Dim tmpStr As String
        Dim vMsg As String = ""
        Dim iThours As Integer = 0

        '201508修正 (show_cbRejectDayIn),'450小時以下為3日,'451-900小時為5日,'901小時以上為10日,

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
        ViewState(vs_msg) = ""
        Dim dr As DataRow
        If OCIDValue1.Value <> "" Then
            If ViewState(vs_OCID) Is Nothing Then
                ViewState(vs_OCID) = "0"
            End If
            If ViewState(vs_OCID) <> OCIDValue1.Value Then
                ViewState(vs_OCID) = OCIDValue1.Value
                SOCID.Items.Clear()

                Dim pms As New Hashtable From {{"OCID", OCIDValue1.Value}}
                Dim sql As String = ""
                sql &= " SELECT Thours,AppliedResultM,IsClosed" & vbCrLf
                sql &= " ,STDate,FTDate" & vbCrLf
                sql &= " FROM Class_ClassInfo" & vbCrLf
                sql &= " WHERE OCID=@OCID" & vbCrLf
                dr = DbAccess.GetOneRow(sql, objconn, pms)
                If Convert.ToString(dr("Thours")) <> "" Then
                    '離退訓遞補期限 顯示。
                    Call show_cbRejectDayIn(Convert.ToString(dr("Thours")))
                End If
                If dr("AppliedResultM").ToString = "Y" Then
                    ViewState(vs_msg) += "學員經費審核結果已經通過，不可新增" & vbCrLf
                End If

                Dim tmpErrMsg1 As String = TIMS.Chk_StopUseDate(Me, Days1, Days2, Convert.ToString(dr("IsClosed")), dr("FTDate"))
                If TIMS.sUtl_ChkTest() Then tmpErrMsg1 = "" '測試用

                If tmpErrMsg1 <> "" Then ViewState(vs_msg) &= tmpErrMsg1

                If ViewState(vs_msg) <> "" Then
                    SOCID.Items.Clear()
                    SOCID.Items.Add(New ListItem("請選擇其他班別", 0))
                    Common.MessageBox(Me, ViewState(vs_msg))
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

                If Convert.ToString(dr("THours")) <> "" Then
                    If dr("THours").ToString > "0" Then
                        hidTHoours.Value = dr("THours").ToString
                        LabTHours.Text = "(本班課程總訓練時數為 " & dr("THours").ToString & "小時) "
                    End If
                End If

                Dim parms As New Hashtable From {{"OCID", OCIDValue1.Value}}
                sql = ""
                sql &= " SELECT b.Name, a.SOCID,a.StudStatus "
                sql &= " FROM Stud_StudentInfo b "
                sql &= " JOIN (SELECT OCID,SOCID,SID,STUDSTATUS FROM CLASS_STUDENTSOFCLASS WHERE OCID=@OCID ) a ON b.SID=a.SID"
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

    '查詢編輯的資料
    Sub EditCreate1()
        Dim sql As String = ""
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
        sql &= " ,c.WkAheadOfSch" & vbCrLf '提前就業
        '提前就業判斷(依目前系統輸入值)
        sql &= " ,case when ((ISNULL(c.TrainHours,0)/cc.THours) >= 0.5 ) AND ISNULL(a.NeedPay,'N') ='N' AND c.RTReasonID='02' then 'Y' END WkAheadOfSch2" & vbCrLf
        sql &= " FROM Stud_LeaveTraining a" & vbCrLf
        sql &= " join Class_StudentsOfClass c  on a.SOCID=c.SOCID" & vbCrLf
        sql &= " join Class_Classinfo cc on cc.OCID=c.OCID" & vbCrLf
        sql &= " join Stud_StudentInfo d on c.SID=d.SID" & vbCrLf
        sql &= " join Key_RejectTReason b on c.RTReasonID=b.RTReasonID" & vbCrLf
        sql &= " WHERE a.SLTID=@SLTID" & vbCrLf

        Dim parms As New Hashtable From {{"SLTID", rqSLTID}}
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

        cbRejectDayIn14.Checked = False '不點選
        cbRejectDayIn14_N.Checked = False '不點選
        cbRejectDayIn14.Enabled = True '不鎖定
        cbRejectDayIn14_N.Enabled = True '不鎖定
        labMakeSOCID.Text = "" '清空

        If Not dr Is Nothing Then
            If trRejectDayIn14.Visible = True Then
                Select Case Convert.ToString(dr("RejectDayIn14"))
                    Case "Y"
                        cbRejectDayIn14.Checked = True '點選
                        If TIMS.CheckRejectSOCID(Convert.ToString(dr("SOCID")), objconn) Then
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
                    If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(Today)) > iTmpDay Then
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

            RTReasonID.RepeatLayout = RepeatLayout.Table
            If sm.UserInfo.Years >= Cst_2015規則1 Then
                Select Case Convert.ToString(dr("StudStatus"))
                    Case TIMS.cst_reject_離
                        'SELECT * FROM Key_RejectTReason where SORT2 IS NOT NULL 
                        RTReasonID.RepeatLayout = RepeatLayout.Flow
                        RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離, objconn, Convert.ToString(dr("RTReasonID")))
                        Common.SetListItem(StudStatus, TIMS.cst_reject_離)
                        If Convert.ToString(dr("RejectTDate1")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
                        End If
                    Case TIMS.cst_reject_退
                        'SELECT * FROM Key_RejectTReason where SORT3 IS NOT NULL 
                        RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_退, objconn, Convert.ToString(dr("RTReasonID")))
                        Common.SetListItem(StudStatus, TIMS.cst_reject_退)
                        If Convert.ToString(dr("RejectTDate2")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
                        End If
                End Select
            ElseIf sm.UserInfo.Years >= Cst_2014規則1 Then
                Select Case Convert.ToString(dr("StudStatus"))
                    Case TIMS.cst_reject_離
                        RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離, objconn, "")
                        Common.SetListItem(StudStatus, TIMS.cst_reject_離)
                        If Convert.ToString(dr("RejectTDate1")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
                        End If
                    Case TIMS.cst_reject_退
                        RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_退, objconn, "")
                        Common.SetListItem(StudStatus, TIMS.cst_reject_退)
                        If Convert.ToString(dr("RejectTDate2")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
                        End If
                End Select
            Else
                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
                Select Case Convert.ToString(dr("StudStatus"))
                    Case TIMS.cst_reject_離
                        Common.SetListItem(StudStatus, TIMS.cst_reject_離)
                        If Convert.ToString(dr("RejectTDate1")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
                        End If
                    Case TIMS.cst_reject_退
                        Common.SetListItem(StudStatus, TIMS.cst_reject_退)
                        If Convert.ToString(dr("RejectTDate2")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
                        End If
                End Select
            End If

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                RTReasonID.RepeatLayout = RepeatLayout.Table
                RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
                Select Case Convert.ToString(dr("StudStatus"))
                    Case TIMS.cst_reject_離
                        Common.SetListItem(StudStatus, TIMS.cst_reject_離)
                        If Convert.ToString(dr("RejectTDate1")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
                        End If
                    Case TIMS.cst_reject_退
                        Common.SetListItem(StudStatus, TIMS.cst_reject_退)
                        If Convert.ToString(dr("RejectTDate2")) <> "" Then
                            RejectTDate.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
                        End If
                End Select
            End If

            HidRTReasonID.Value = Convert.ToString(dr("RTReasonID"))
            Common.SetListItem(RTReasonID, HidRTReasonID.Value)

            If Convert.ToString(dr("RTReasonID")) = "99" _
                OrElse Convert.ToString(dr("RTReasonID")) = "98" Then
                RTReasoOther.Text = Convert.ToString(dr("RTReasoOther"))
            End If
            RTReasonThat.Text = dr("RTReasonThat").ToString

            '20080901 andy  edit  
            'SumOfPay.Text = Convert.ToString(dr("SumOfPay"))
            'HadPay.Text = Convert.ToString(dr("HadPay"))
            'If Convert.ToString(dr("NeedPay")) = "Y" Or Convert.ToString(dr("NeedPay")) = "y" Then
            '    NeedPay.Items(1).Selected = True '應賠償為是
            'Else
            '    NeedPay.Items(2).Selected = True '應賠償為否
            '    '20080901 andy edit 
            '    HadPay.Text = "0"
            '    SumOfPay.Text = "0"
            '    SumOfPay.Enabled = False
            'End If

            'SumOfPay.Text = Convert.ToString(dr("SumOfPay"))
            'HadPay.Text = Convert.ToString(dr("HadPay"))

            TrainHours.Text = dr("TrainHours").ToString
            If Not Convert.IsDBNull(dr("PayStatus")) Then
                Common.SetListItem(PayStatus, dr("PayStatus"))
            End If
            If Not Convert.IsDBNull(dr("NoClose")) Then
                Common.SetListItem(NoClose, dr("NoClose"))
            End If
            NoClose_Desc.Text = Convert.ToString(dr("NoClose_Desc"))
            If Not Convert.IsDBNull(dr("Other")) Then
                Common.SetListItem(Other, dr("Other"))
            End If
            OtherDesc.Text = Convert.ToString(dr("OtherDesc"))
            If (Convert.ToString(dr("JobOrgName")) <> "") Then
                OrgName.Text = Convert.ToString(dr("JobOrgName"))
            End If
            If Not IsDBNull(dr("JobTel")) Then
                JobTel.Text = dr("JobTel")
            End If
            If Not IsDBNull(dr("JobZipCode")) Then
                JobZipCode.Value = dr("JobZipCode")
                JobCity.Text = TIMS.Get_ZipName(JobZipCode.Value)
            End If
            If IsDBNull(dr("Jobaddress")) = False Then
                Jobaddress.Text = dr("Jobaddress")
            End If

            If IsDBNull(dr("JobDate")) = False Then
                'JobDate.Text = Common.FormatDate(dr("JobDate"))
                JobDate.Text = TIMS.Cdate3(dr("JobDate"))
            End If

            If IsDBNull(dr("JobSalID")) = False Then
                Common.SetListItem(JobSalID, dr("JobSalID"))
            End If

            '20080815  andy  新增備註欄位
            tb_note.Text = Convert.ToString(dr("note"))
        End If

        sql = ""
        sql &= " SELECT ISNULL(b.JobID,TrainID) JobID"
        sql &= " ,ISNULL(b.JobName,b.TrainName) JobName"
        sql &= " ,b.TrainID,b.TrainName,a.THours,a.TMID,a.ClassCName,a.CyclType,a.OCID,a.AppliedResultM,a.IsClosed,a.STDate,a.FTDate"
        sql &= " FROM Class_ClassInfo a "
        sql &= " JOIN Key_TrainType b ON a.TMID=b.TMID "
        sql &= " WHERE a.OCID=@OCID "
        parms.Clear()
        parms.Add("OCID", rqOCID)
        dr = DbAccess.GetOneRow(sql, objconn, parms)

        If Convert.ToString(dr("THours")) <> "" Then
            If dr("THours").ToString > "0" Then
                hidTHoours.Value = dr("THours").ToString
                LabTHours.Text = "(本班課程總訓練時數為 " & dr("THours").ToString & "小時) "
            End If
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '如果是產學訓計畫
            TMID1.Text = "[" & dr("JobID") & "]" & dr("JobName")
        Else
            TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
        End If

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
                End If
            Else
                '非系統管理者以上權限
                If DateDiff(DateInterval.Day, dr("FTDate"), Now) > Days1 Then
                    'Button1.Visible = False
                    Button1.Enabled = False
                    TIMS.Tooltip(Button1, "超過結訓日期" & Days1 & "天，停用儲存功能")
                End If
            End If
        End If
        ViewState(vs_OCID) = dr("OCID")
    End Sub

    '儲存前檢核1
    Function Checkdata1(ByRef sErrmsg As String) As Boolean
        Dim rst As Boolean = True
        sErrmsg = ""

        Call TIMS.OpenDbConn(objconn)
        Dim v_StudStatus As String = TIMS.GetListValue(StudStatus)
        Select Case v_StudStatus
            Case TIMS.cst_reject_離
            Case TIMS.cst_reject_退
            Case Else
                sErrmsg &= "輸入資料異常，請重新查詢!" & vbCrLf
                Return False
        End Select

        SOCIDValue = TIMS.ClearSQM(SOCIDValue)
        If SOCIDValue = "" Then
            sErrmsg &= "輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        Dim parms3 As New Hashtable From {{"OCID", rqOCID}, {"SOCID", SOCIDValue}}
        Dim sql3 As String = ""
        sql3 &= " select cs.SOCID,cs.StudStatus" & vbCrLf
        sql3 &= " from Class_StudentsOfClass cs" & vbCrLf
        sql3 &= " join Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
        sql3 &= " where cs.OCID =@OCID and cs.SOCID =@SOCID" & vbCrLf
        'sql += " and cs.StudStatus in (2,3)" & vbCrLf 'Dim sCmd3 As New SqlCommand(sql, objconn)
        Dim dt3 As DataTable = DbAccess.GetDataTable(sql3, objconn, parms3)
        If TIMS.dtNODATA(dt3) OrElse dt3.Rows.Count <> 1 Then
            sErrmsg &= "輸入資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        '20080923 andy 儲存時狀態為新增判斷該學員離退訓狀態
        If rqProecess = "add" Then
            Dim parms1 As New Hashtable From {{"OCID", rqOCID}, {"SOCID", SOCIDValue}}
            Dim sql1 As String = ""
            sql1 &= " select cs.SOCID,cs.StudStatus" & vbCrLf
            sql1 &= " from Class_StudentsOfClass cs" & vbCrLf
            sql1 &= " join Stud_StudentInfo ss on ss.sid =cs.sid" & vbCrLf
            sql1 &= " where cs.StudStatus in (2,3) and cs.OCID =@OCID and cs.SOCID =@SOCID" & vbCrLf
            Dim dt As DataTable = DbAccess.GetDataTable(sql1, objconn, parms1)

            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)
                Dim status As String = If(dr("StudStatus") = 2, "離訓", If(dr("StudStatus") = 3, "退訓", ""))
                'Common.MessageBox(Me, "此學員已是" & status & "狀態無法變更!")
                If status <> "" Then sErrmsg &= String.Concat("此學員已是", status, "狀態無法變更!") & vbCrLf
                Return False 'Exit Function
            End If
        End If

        '判開「離退訓日期」不可超過 開訓日期 & 結訓日期
        If Convert.ToDateTime(ViewState(vs_StDate)) > Convert.ToDateTime(RejectTDate.Text) Then
            'Common.MessageBox(Me, cst_str離退訓 & "日期不可小於開訓日期(" & ViewState(vs_StDate) & ")!")
            sErrmsg &= cst_str離退訓 & "日期不可小於開訓日期(" & ViewState(vs_StDate) & ")!" & vbCrLf
            Return False 'Exit Function
        ElseIf Convert.ToDateTime(ViewState(vs_FtDate)) < Convert.ToDateTime(RejectTDate.Text) Then
            'Common.MessageBox(Me, cst_str離退訓 & "日期不可大於結訓日期(" & ViewState(vs_FtDate) & ")!") 'Exit Function
            sErrmsg &= cst_str離退訓 & "日期不可大於結訓日期(" & ViewState(vs_FtDate) & ")!" & vbCrLf
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

        '遞補規則：'1.若沒有 遞補  遞補期限內離退訓 為否 '2.有 遞補  遞補期限內離退訓且為14天內 為是
        'Dim errMsg As String = "" 'errMsg = ""
        Dim chkRejectDayIn14_flag As Boolean = False '是否驗證 遞補期限內離退訓 預設為false
        '有顯示才須判斷
        If trRejectDayIn14.Visible = True Then
            '驗證勾選。
            If Not cbRejectDayIn14.Checked AndAlso Not cbRejectDayIn14_N.Checked Then
                sErrmsg &= "請勾選，遞補期限內離退訓" & vbCrLf
            End If
            If cbRejectDayIn14.Checked AndAlso cbRejectDayIn14_N.Checked Then
                sErrmsg &= "遞補期限內離退訓 ，是／否請勾選1個" & vbCrLf
            End If
            If sErrmsg <> "" Then
                'Common.MessageBox(Page, errMsg)
                Return False 'Exit Function
            End If
            '勾是進入驗證
            If cbRejectDayIn14.Checked Then
                chkRejectDayIn14_flag = True '是否驗證 遞補期限內離退訓 True
            End If
        End If

        '是否驗證 遞補期限內離退訓 True
        If chkRejectDayIn14_flag Then
            Dim iTmpDay As Integer = 3 '5
            Dim sTmpDay2 As String = ""
            If HidRejectDay.Value <> "" Then iTmpDay = HidRejectDay.Value

            Dim flagOver14 As Boolean = False 'True:已超過14天(或系統規定天數)或未在14天內。
            If Not flagOver14 Then
                '離退日期早開訓日
                If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(RejectTDate.Text)) < 0 Then
                    sTmpDay2 = cst_str離退 & "日期早開訓日!!"
                    flagOver14 = True
                End If
            End If

            If Not flgROLEIDx0xLIDx0 Then
                If Not flagOver14 Then
                    'Sys_Holiday
                    Dim SSQL As String = "" & vbCrLf
                    SSQL &= " select 'x' from Sys_Holiday where RID =@RID" & vbCrLf
                    SSQL &= " AND HolDate>=@StDate AND HolDate<=@RejectTDate" & vbCrLf
                    Dim sCmd As New SqlCommand(SSQL, objconn)
                    Dim dtH As New DataTable 'Call TIMS.OpenDbConn(objconn)
                    With sCmd
                        .Parameters.Clear()
                        Dim V_RID As String = Left(sm.UserInfo.RID, 1)
                        If Len(sm.UserInfo.RID) = 1 Then V_RID = sm.UserInfo.RID
                        .Parameters.Add("RID", SqlDbType.VarChar).Value = Left(sm.UserInfo.RID, 1)
                        .Parameters.Add("StDate", SqlDbType.DateTime).Value = CDate(ViewState(vs_StDate))
                        .Parameters.Add("RejectTDate", SqlDbType.DateTime).Value = CDate(RejectTDate.Text)
                        'dtH = New DataTable 'dtH.Load(.ExecuteReader())
                        dtH = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
                    End With

                    'da = TIMS.GetOneDA()
                    'da.SelectCommand.Parameters.Clear()
                    'If Len(sm.UserInfo.RID) = 1 Then
                    '    da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                    'Else
                    '    da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = Left(sm.UserInfo.RID, 1)
                    'End If
                    'da.SelectCommand.Parameters.Add("StDate", SqlDbType.DateTime).Value = CDate(ViewState(vs_StDate))
                    'da.SelectCommand.Parameters.Add("RejectTDate", SqlDbType.DateTime).Value = CDate(RejectTDate.Text)
                    'Dim dtH As DataTable = New DataTable
                    'TIMS.Fill(Sql, da, dtH)

                    '離退日期超過14天 開訓日
                    If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(RejectTDate.Text)) > (iTmpDay + dtH.Rows.Count) Then
                        sTmpDay2 = cst_str離退訓 & "日期與開訓日期，已超過" & CStr(iTmpDay) & "天(須於" & CStr(iTmpDay) & "天內)!"
                        flagOver14 = True
                    End If
                End If

                If Not flagOver14 Then
                    iTmpDay = 14
                    If DateDiff(DateInterval.Day, CDate(ViewState(vs_StDate)), CDate(Today)) > iTmpDay Then
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

        Dim v_RTReasonID As String = TIMS.GetListValue(RTReasonID)
        If v_RTReasonID = "99" OrElse v_RTReasonID = "98" Then
            Dim tmpStr As String = ""
            If RTReasoOther.Text <> "" Then tmpStr = Trim(RTReasoOther.Text)
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
        If v_RTReasonID = "02" Then
            TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
            hidTHoours.Value = TIMS.ClearSQM(hidTHoours.Value)
            Dim iTrainHours As Double = TIMS.VAL1(TrainHours.Text)
            Dim iTHours As Double = TIMS.VAL1(hidTHoours.Value)
            'If Not TIMS.Chk_WkAheadOfSch(TrainHours.Text, .hidTHoours.Value, NeedPay.SelectedValue, v_RTReasonID) Then
            '    sErrmsg &= "該學員 不符合 提前就業認定原則，請重新確認輸入資料。" & vbCrLf
            'End If
            If (iTrainHours / iTHours) < 0.5 Then
                sErrmsg &= "該學員 離退訓原因為提前就業(訓期滿1/2以上)，離退訓日需超過訓期1/2以上!!" & vbCrLf
            End If
        End If
        '(SELECT * FROM Key_RejectTReason) '選擇 14:訓期未滿1/2找到工作
        If v_RTReasonID = "14" Then
            TrainHours.Text = TIMS.ClearSQM(TrainHours.Text)
            hidTHoours.Value = TIMS.ClearSQM(hidTHoours.Value)
            Dim iTrainHours As Double = TIMS.VAL1(TrainHours.Text)
            Dim iTHours As Double = TIMS.VAL1(hidTHoours.Value)
            If (iTrainHours / iTHours) >= 0.5 Then
                sErrmsg &= "離退訓原因 為訓期未滿1/2找到工作，離退訓日需未超過訓期1/2!!" & vbCrLf
            End If
        End If
        If sErrmsg <> "" Then
            Return False
        End If
        'If errMsg <> "" Then,Common.MessageBox(Page, errMsg),Exit Function,End If, '判斷結束
        'Dim da1 As SqlDataAdapter = Nothing
        'Dim sCmd2 As New SqlCommand(sql, objconn)

        Dim parms2 As New Hashtable From {{"SOCID", SOCIDValue}}
        Dim sql2 As String = "SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID "
        Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, parms2)
        If TIMS.dtNODATA(dt2) OrElse dt2.Rows.Count <> 1 Then
            sErrmsg &= "資料異常，請重新查詢!" & vbCrLf
            Return False
        End If

        Select Case rqProecess
            Case "add" 'rqProecess
            Case "edit" 'rqProecess

                Dim parms5 As New Hashtable From {{"SLTID", rqSLTID}}
                Dim sql5 As String = "SELECT * FROM STUD_LEAVETRAINING WHERE SLTID=@SLTID "
                Dim dt5 As DataTable = DbAccess.GetDataTable(sql5, objconn, parms5)
                If TIMS.dtNODATA(dt5) OrElse dt5.Rows.Count <> 1 Then
                    sErrmsg &= "傳入參數有誤,請重新操作，點選功能!!" & vbCrLf
                    Return False
                End If
        End Select

        If sErrmsg <> "" Then rst = False
        Return rst
    End Function

    '儲存
    Sub Savedata1()
        '儲存開始
        SOCIDValue = TIMS.ClearSQM(SOCIDValue)

        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update
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

        'dr("NeedPay") = Convert.DBNull
        dr("NeedPay") = "N" 'Convert.DBNull
        dr("SumOfPay") = Convert.DBNull
        dr("HadPay") = Convert.DBNull
        '追償狀況
        Dim v_PayStatus As String = TIMS.GetListValue(PayStatus)
        dr("PayStatus") = If(v_PayStatus <> "", v_PayStatus, Convert.DBNull)
        '追償狀況_未結案原因
        Dim v_NoClose As String = TIMS.GetListValue(NoClose)
        dr("NoClose") = If(v_NoClose <> "", v_NoClose, Convert.DBNull)
        dr("NoClose_Desc") = If(NoClose_Desc.Text <> "", NoClose_Desc.Text, Convert.DBNull)
        '追償狀況_其他原因
        Dim v_Other As String = TIMS.GetListValue(Other)
        dr("Other") = If(v_Other <> "", v_Other, Convert.DBNull)
        dr("note") = If(tb_note.Text <> "", tb_note.Text, Convert.DBNull)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Dim htPP As New Hashtable From {
            {"TransType", s_TransType},
            {"TargetTable", s_TargetTable},
            {"FuncPath", s_FuncPath},
            {"s_WHERE", s_WHERE}
        }
        TIMS.SaveTRANSLOG(sm, objconn, dr, htPP)
        DbAccess.UpdateDataTable(dt, da)

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
                    Dim vRejectDayIn14 As String = "" ' dr1("RejectDayIn14") = Convert.DBNull '(兩週內)離退訓
                    If cbRejectDayIn14.Checked Then vRejectDayIn14 = "Y" '(兩週內)離退訓 
                    If cbRejectDayIn14_N.Checked Then vRejectDayIn14 = "N" '(兩週內)離退訓 
                    dr1("RejectDayIn14") = If(vRejectDayIn14 <> "", vRejectDayIn14, Convert.DBNull)
                End If
            End If

            dr1("WkAheadOfSch") = Convert.DBNull '其他狀況為非提前就業者
            If TrainHours.Text <> "" Then
                '符合提前就業人數者  dr1("WkAheadOfSch") = "Y"
                If Not IsNumeric(TrainHours.Text) Then TrainHours.Text = "0" '檢測數字異常設為0 
                If Not IsNumeric(hidTHoours.Value) Then hidTHoours.Value = "0" '檢測數字異常設為0 
                '符合提前就業判斷
                'If TIMS.Chk_WkAheadOfSch(TrainHours.Text, .hidTHoours.Value, NeedPay.SelectedValue, v_RTReasonID) Then dr1("WkAheadOfSch") = "Y" 
            End If

            Dim v_StudStatus As String = TIMS.GetListValue(StudStatus)
            dr1("StudStatus") = v_StudStatus 'v_StudStatus
            Select Case v_StudStatus '2/3
                Case "2"
                    dr1("RejectTDate1") = RejectTDate.Text
                    dr1("RejectTDate2") = Convert.DBNull
                Case "3"
                    dr1("RejectTDate1") = Convert.DBNull
                    dr1("RejectTDate2") = RejectTDate.Text
            End Select

            '離退訓原因
            Dim v_RTReasonID As String = TIMS.GetListValue(RTReasonID)
            dr1("RTReasonID") = v_RTReasonID ' v_RTReasonID
            If v_RTReasonID = "99" OrElse v_RTReasonID = "98" Then
                dr1("RTReasoOther") = If(RTReasoOther.Text <> "", RTReasoOther.Text, Convert.DBNull)
            End If
            dr1("RTReasonThat") = If(RTReasonThat.Text <> "", RTReasonThat.Text, Convert.DBNull)
            OrgName.Text = TIMS.ClearSQM(OrgName.Text)
            If OrgName.Text <> "" AndAlso v_RTReasonID = "02" Then
                dr1("JobOrgName") = OrgName.Text
                dr1("JobTel") = JobTel.Text
                dr1("JobZipCode") = JobZipCode.Value
                dr1("Jobaddress") = Jobaddress.Text
                dr1("JobDate") = TIMS.Cdate2(JobDate.Text) 'If(objValue IsNot Nothing, CDate(objValue), Convert.DBNull)
                dr1("JobSalID") = JobSalID.SelectedValue
            Else
                dr1("JobOrgName") = Convert.DBNull
                dr1("JobTel") = Convert.DBNull
                dr1("JobZipCode") = Convert.DBNull
                dr1("Jobaddress") = Convert.DBNull
                dr1("JobDate") = Convert.DBNull
                dr1("JobSalID") = Convert.DBNull
            End If
            dr1("TrainHours") = If(TrainHours.Text = "", Convert.DBNull, TrainHours.Text)
            dr1("ModifyAcct") = sm.UserInfo.UserID
            '建檔日期(限add用而以)
            If rqProecess = "add" Then dr1("RejectCDate") = Now.ToString("yyyy/MM/dd")
            dr1("ModifyDate") = Now
            'Dim oriSupplyID As String = Convert.ToString(dr1("SupplyID"))
            'Dim oriBudgetID As String = Convert.ToString(dr1("BudgetID"))
            'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0% '請選擇變為空。
            dr1("SupplyID") = "9" 'If(oriSupplyID <> "", "9", Convert.DBNull)
            'BudgetID 01:公務,02:就安,03:就保,04:再出發,97:公務(ECFA),98:特別預算,99:不補助
            dr1("BudgetID") = "99" 'If(oriBudgetID <> "", "99", Convert.DBNull) '"99"

            Dim htPP2 As New Hashtable From {
                {"TransType", TIMS.cst_TRANS_LOG_Update},
                {"TargetTable", "CLASS_STUDENTSOFCLASS"},
                {"FuncPath", s_FuncPath},
                {"s_WHERE", s_WHERE}
            } 'htPP.Clear()
            TIMS.SaveTRANSLOG(sm, objconn, dr1, htPP2)

            DbAccess.UpdateDataTable(dt1, da1)
        End If

        'UPDATE STUD_SUBSIDYCOST
        '2020/12/29
        '產業人才投資方案： 首頁>> 學員動態管理 >> 教務管理 >> 離退訓作業， 訓練單位進行離退訓作業時，
        '系統自動將離退訓學員連動至學員動態管理， 並將補助預算別調整為不補助， 補助比例0%， 不須再先修改學員資料維護即可進行離退訓作業。
        Dim dt2 As DataTable = Nothing
        Dim dr2 As DataRow = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        sql = "SELECT * FROM STUD_SUBSIDYCOST WHERE SOCID='" & SOCIDValue & "'"
        dt2 = DbAccess.GetDataTable(sql, da2, objconn)
        If dt2.Rows.Count = 1 Then
            dr2 = dt2.Rows(0)
            Dim oriSupplyID As String = Convert.ToString(dr2("SupplyID"))
            Dim oriBudID As String = Convert.ToString(dr2("BudID"))
            'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0% '請選擇變為空。
            dr2("SupplyID") = If(oriSupplyID <> "", "9", Convert.DBNull)
            'BudgetID 01:公務,02:就安,03:就保,04:再出發,97:公務(ECFA),98:特別預算,99:不補助
            dr2("BudID") = If(oriBudID <> "", "99", Convert.DBNull) '"99"
            'Dim htPP As New Hashtable
            Dim htPP2 As New Hashtable From {
                {"TransType", TIMS.cst_TRANS_LOG_Update},
                {"TargetTable", "STUD_SUBSIDYCOST"},
                {"FuncPath", s_FuncPath},
                {"s_WHERE", s_WHERE}
            } '.Clear()
            TIMS.SaveTRANSLOG(sm, objconn, dr2, htPP2)
            DbAccess.UpdateDataTable(dt2, da2)
        End If

        If Session(vs_search) Is Nothing Then Session(vs_search) = ViewState(vs_search)

        Select Case rqProecess
            Case "add"
                Common.RespWrite(Me, "<script language='javascript'>alert('新增成功');</script>")
            Case "edit"
                Common.RespWrite(Me, "<script language='javascript'>alert('修改成功');</script>")
            Case Else
                Common.RespWrite(Me, "<script language='javascript'>alert('請檢查輸入參數!!');</script>")
                Exit Sub
        End Select
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '產投不顯示
            Common.RespWrite(Me, "<script language='javascript'>alert('請記得填寫相關後續作業,若該學員有課程成績,請填寫結訓成績;\n\n若該學員有申請職訓生活津貼,請於職訓生活津貼系統進行" & cst_str離退訓 & "作業。');</script>")
        End If
        Common.RespWrite(Me, "<script language='javascript'>location.href='SD_05_004.aspx?ID=" & TIMS.Get_MRqID(Me) & "';</script>")
    End Sub

    '儲存學生按鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '1.取得學員學號。
        SOCIDValue = Split(SOCID.SelectedValue, "&")(0)
        '2.檢核
        Dim sErrmsg As String = ""
        Call Checkdata1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        '3.儲存
        Call Savedata1()
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
#End Region

    '回上一頁
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        If Session(vs_search) Is Nothing Then
            Session(vs_search) = ViewState(vs_search)
        End If
        Dim url1 As String = String.Concat("SD_05_004.aspx?ID=", TIMS.Get_MRqID(Me)) ' Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub

    '清除薪資級距
    Private Sub btnClearJobSalID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearJobSalID.Click
        JobSalID.SelectedIndex = -1 '清空選項
    End Sub

    '離退訓選項將不同。
    Protected Sub StudStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles StudStatus.SelectedIndexChanged
        HidRTReasonID.Value = ""
        Dim v_RTReasonID As String = TIMS.GetListValue(RTReasonID)
        If v_RTReasonID <> "" Then HidRTReasonID.Value = v_RTReasonID

        'SELECT * FROM Key_RejectTReason
        'Cst_2016規則1
        RTReasonID.RepeatLayout = RepeatLayout.Table
        Dim v_StudStatus As String = TIMS.GetListValue(StudStatus)
        If sm.UserInfo.Years >= Cst_2015規則1 Then
            Select Case v_StudStatus
                Case TIMS.cst_reject_離
                    RTReasonID.RepeatLayout = RepeatLayout.Flow
                    RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, v_StudStatus, objconn, "")
                    Common.SetListItem(RTReasonID, HidRTReasonID.Value)
                Case TIMS.cst_reject_退
                    RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, v_StudStatus, objconn, "")
                    Common.SetListItem(RTReasonID, HidRTReasonID.Value)
                Case Else 'Case TIMS.cst_reject_退
                    RTReasonID.Items.Clear()
                    RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
                    Common.SetListItem(RTReasonID, HidRTReasonID.Value)
                    Common.MessageBox(Me, "請選擇" & cst_str離退訓 & "種類!!")
                    Exit Sub
            End Select
        ElseIf sm.UserInfo.Years >= Cst_2014規則1 Then
            Select Case v_StudStatus
                Case TIMS.cst_reject_離, TIMS.cst_reject_退
                    RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, v_StudStatus, objconn, "")
                    Common.SetListItem(RTReasonID, HidRTReasonID.Value)
                Case Else 'Case TIMS.cst_reject_退
                    RTReasonID.Items.Clear()
                    RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
                    Common.SetListItem(RTReasonID, HidRTReasonID.Value)
                    Common.MessageBox(Me, "請選擇" & cst_str離退訓 & "種類!!")
                    Exit Sub
            End Select
        Else
            RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
            Common.SetListItem(RTReasonID, HidRTReasonID.Value)
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            RTReasonID.RepeatLayout = RepeatLayout.Table
            RTReasonID = TIMS.Get_RejectTReason(Me, RTReasonID, TIMS.cst_reject_離退Old, objconn, "")
            Common.SetListItem(RTReasonID, HidRTReasonID.Value)
        End If

    End Sub

End Class
