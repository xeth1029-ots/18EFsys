Partial Class CP_01_006_add8
    Inherits AuthBasePage

#Region "WEBF"
    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("CLASS_UNEXPECTVISITOR", objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "COURSENAME", CourseName)
        'Call TIMS.sUtl_SetMaxLen(dt, "DATA4_NOTE", Data4_Note)
        Call TIMS.sUtl_SetMaxLen(dt, "STUD_NAME", Stud_Name)
        Call TIMS.sUtl_SetMaxLen(dt, "SITEM1_NOTE", SItem1_Note)
        Call TIMS.sUtl_SetMaxLen(dt, "SITEM2_NOTE", SItem2_Note)
        'Call TIMS.sUtl_SetMaxLen(dt, "SITEM3_NOTE", SITEM3_NOTE)
        Call TIMS.sUtl_SetMaxLen(dt, "LITEM2_2_NOTE", LItem2_2_Note)
        Call TIMS.sUtl_SetMaxLen(dt, "CURSENAME", CurseName)
        Call TIMS.sUtl_SetMaxLen(dt, "VISITORNAME", VisitorName)

        Call TIMS.sUtl_SetMaxLen(dt, "STUD_NAME2", Stud_Name2)
        Call TIMS.sUtl_SetMaxLen(dt, "LITEM2_3_NOTE", LItem2_3_Note)
    End Sub
#End Region

    'Const cst_VerDate_20180101 As String = "2018/01/01"
    Const cst_xx26 As String = "%26"
    'Const cst_xx51 As String = "日期^時段^課程單元^師資^助教^場地^人數上限^專職辦訓人員^未依規定收費及退費^出席率為0且歸責於單位^其他"
    'Const cst_xx61 As String = "學員陳情申訴經查屬實，屬訓練單位之缺失^未落實學員簽到退^代簽^未依規定上課^廣宣缺失^招訓違反公平原則^其他"
    'Const cst_xx71 As String = "助教非核定助教^場地住址相同僅內部教室變更^其他"
    'Const cst_xx81 As String = "訓練單位有內部或與外部單位有糾紛導致影響學員上課權益^其他"
    'Dim sdSITEM51 As String() '將變更項目名稱定義到陣列之中
    'Dim sdSITEM61 As String() '將變更項目名稱定義到陣列之中
    'Dim sdSITEM71 As String() '將變更項目名稱定義到陣列之中
    'Dim sdSITEM81 As String() '將變更項目名稱定義到陣列之中

    Const cst_CODE_KIND_SITEM51 As String = "CP_01_006_SITEM51"
    Const cst_CODE_KIND_SITEM61 As String = "CP_01_006_SITEM61"
    Const cst_CODE_KIND_SITEM71 As String = "CP_01_006_SITEM71"
    Const cst_CODE_KIND_SITEM81 As String = "CP_01_006_SITEM81"
    'State
    Const cst_State_Add_新增 As String = "Add"
    Const cst_State_View_檢視 As String = "View" '查詢
    Const cst_State_Edit_修改 As String = "Edit"
    'cst_State_Add_新增/cst_State_View_檢視/cst_State_Edit_修改

    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        Call sUtl_PageInit1()
        '檢查Session是否存在--------------------------End

        'sdSITEM51 = cst_xx51.Split("^")
        'sdSITEM61 = cst_xx61.Split("^")
        'sdSITEM71 = cst_xx71.Split("^")
        'sdSITEM81 = cst_xx81.Split("^")
        'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
        'iPYNum = TIMS.sUtl_GetPYNum(Me)
        'Hid_VerDate.Value = ""
        'If iPYNum = 3 Then Hid_VerDate.Value = cst_VerDate_20180101

        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        'Dim rqDOCID As String = TIMS.ClearSQM(Request("DOCID"))
        Dim rqState As String = TIMS.ClearSQM(Request("State"))
        'cst_State_Add_新增/cst_State_View_檢視/cst_State_Edit_修改
        Hid_State1.Value = rqState
        'Dim rqType As String = TIMS.ClearSQM(Request("Type"))

        If Not IsPostBack Then
            'AuthCount.Attributes("onclick") = "javascript:return Calculate1()"
            'AuthCount.Attributes("onblur") = "javascript:return Calculate1()"
            'RejectCount.Attributes("onclick") = "javascript:return Calculate1()"
            'RejectCount.Attributes("onblur") = "javascript:return Calculate1()"

            AtteRate.Attributes("onclick") = "javascript:return Calculate2()"
            TurthCount.Attributes("onclick") = "javascript:return Calculate2()"
            TurnoutCount.Attributes("onclick") = "javascript:return Calculate2()"
            AtteRate.Attributes("onblur") = "javascript:return Calculate2()"
            TurthCount.Attributes("onblur") = "javascript:return Calculate2()"
            TurnoutCount.Attributes("onblur") = "javascript:return Calculate2()"
            'AuthCount.Attributes("onclick") = "javascript:return Calculate2()"

            OLessonTeah1.Attributes.Add("onDblClick", "javascript:OpenLessonTeah1('Add');")
            teacherbtn.Attributes.Add("onClick", "javascript:OpenLessonTeah1('Add');")
            OLessonTeah2.Attributes.Add("onDblClick", "javascript:OpenLessonTeah2('Add');")
            teacherbtn2.Attributes.Add("onClick", "javascript:OpenLessonTeah2('Add');")
            Button1.Attributes("onclick") = "javascript:return chkdata()"
            'Button1.Enabled = True
            Call sCreate1()

            If rqState = "View" Then Button1.Visible = False
            If rqOCID = "" AndAlso OCIDValue1.Value <> "" Then rqOCID = OCIDValue1.Value
            If rqOCID <> "" Then Call sShowData1(rqOCID, rqSeqNo)
            'If rqDOCID <> "" Then Call sCreate1(rqDOCID, "")
        End If

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If

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
    End Sub

    Sub sCreate1()
        'Select * From SYS_SHAREDCODE Where CODE_KIND ='CP_01_006_SITEM51' ORDER BY SORT_ID
        'Select * From SYS_SHAREDCODE Where CODE_KIND ='CP_01_006_SITEM61' ORDER BY SORT_ID
        'Select * From SYS_SHAREDCODE Where CODE_KIND ='CP_01_006_SITEM71' ORDER BY SORT_ID
        'Select * From SYS_SHAREDCODE Where CODE_KIND ='CP_01_006_SITEM81' ORDER BY SORT_ID
        cblSITEM51 = TIMS.GET_CBLCODE1(cst_CODE_KIND_SITEM51, cblSITEM51, objconn)
        cblSITEM61 = TIMS.GET_CBLCODE1(cst_CODE_KIND_SITEM61, cblSITEM61, objconn)
        cblSITEM71 = TIMS.GET_CBLCODE1(cst_CODE_KIND_SITEM71, cblSITEM71, objconn)
        cblSITEM81 = TIMS.GET_CBLCODE1(cst_CODE_KIND_SITEM81, cblSITEM81, objconn)
        Applytime_HH.Value = "09"
        Applytime_MM.Value = "00"
        'LItem_TR.Attributes.Add("Style", "none")
        'LItem_TR2.Attributes.Add("Style", "none")
        LItem_TR.Style.Item("display") = "none"
        LItem_TR2.Style.Item("display") = "none"
        OrgName.Text = sm.UserInfo.OrgName
        OrgID.Value = sm.UserInfo.OrgID
        If Not Session("SearchStr") Is Nothing Then
            Dim MyValue As String
            MyValue = TIMS.GetMyValue(Session("SearchStr"), "prgid")
            If MyValue = "CP_01_006" Then
                MyValue = TIMS.GetMyValue(Session("SearchStr"), "center")
                center.Text = Replace(MyValue, cst_xx26, "&")
                RIDValue.Value = TIMS.GetMyValue(Session("SearchStr"), "RIDValue")
                MyValue = TIMS.GetMyValue(Session("SearchStr"), "TMID1")
                TMID1.Text = Replace(MyValue, cst_xx26, "&")
                MyValue = TIMS.GetMyValue(Session("SearchStr"), "OCID1")
                OCID1.Text = Replace(MyValue, cst_xx26, "&")
                TMIDValue1.Value = TIMS.GetMyValue(Session("SearchStr"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("SearchStr"), "OCIDValue1")
            End If
            If Not Session("SearchStr") Is Nothing Then Session("SearchStr") = Session("SearchStr")
        End If
    End Sub

    Function GET_STUDCOUNT2(ByVal OCID As String) As Hashtable
        Dim rst As New Hashtable
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT ISNULL(COUNT(1),0) STUDCNT"
        sql &= " ,ISNULL(COUNT(CASE WHEN STUDSTATUS in (2,3) then 1 end),0) REJCNT"
        sql &= " FROM dbo.V_STUDENTINFO WHERE OCID=@OCID"
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("OCID", OCID)
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt.Rows(0)
            rst.Clear()
            rst.Add("STUDCNT", dr1("STUDCNT"))
            rst.Add("REJCNT", dr1("REJCNT"))
        End If
        Return rst
    End Function

    '取得資料庫資料
    Sub sShowData1(ByVal OCID As String, ByVal SeqNo As String)
        '不可為空
        OCID = TIMS.ClearSQM(OCID)
        SeqNo = TIMS.ClearSQM(SeqNo)
        If OCID = "" Then OCID = "0"
        If SeqNo = "" Then SeqNo = "0"
        Dim sql As String = ""
        'Class_Visitor
        'sql = "Select * FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & OCID & "' and SeqNo='" & SeqNo & "'"
        sql = "SELECT * FROM dbo.CLASS_UNEXPECTVISITOR WHERE OCID=@OCID and SeqNo=@SeqNo "
        'Dim sCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        'With sCmd
        '    .Parameters.Clear()
        '    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
        '    .Parameters.Add("SeqNo", SqlDbType.VarChar).Value = SeqNo
        '    dt.Load(.ExecuteReader())
        'End With
        Dim parms As Hashtable = New Hashtable()
        parms.Clear()
        parms.Add("OCID", OCID)
        parms.Add("SeqNo", SeqNo)
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        '在學員動態管理>>教務管理>>不預告實地抽訪紀錄表，由系統自動帶出應到人數及計算出席率
        '公式：
        '應到人數=參訓人數-退訓人數
        'AuthCount=Hid_StudCount-RejectCount
        '出席率=(實到人數+請假人數)/應到人數 "
        'AtteRate=(TurthCount+TurnoutCount) /AuthCount

        Hid_StudCount.Value = ""
        RejectCount.Text = "" 'TIMS.GetMyValue2(getP, "REJCNT") '退訓人數
        AuthCount.Text = "" 'Val(Hid_StudCount.Value) - Val(RejectCount.Text) '應到人數
        If Hid_State1.Value = cst_State_Add_新增 Then
            Dim getP As Hashtable = GET_STUDCOUNT2(OCID)
            Hid_StudCount.Value = TIMS.GetMyValue2(getP, "STUDCNT") '參訓人數
            RejectCount.Text = TIMS.GetMyValue2(getP, "REJCNT") '退訓人數
            AuthCount.Text = Val(Hid_StudCount.Value) - Val(RejectCount.Text) '應到人數
        End If

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)

            If flag_ROC Then
                ApplyDate.Text = TIMS.cdate17(dr("ApplyDate"))  'edit，by:20181018
            Else
                ApplyDate.Text = dr("ApplyDate")  'edit，by:20181018
            End If

            Applytime_HH.Value = Left(dr("ApplyTime"), 2)
            Applytime_MM.Value = Right(dr("ApplyTime"), 2)
            AuthCount.Text = Convert.ToString(dr("AuthCount"))
            TurthCount.Text = Convert.ToString(dr("TurthCount"))
            TruancyCount.Text = Convert.ToString(dr("TruancyCount"))
            TurnoutCount.Text = Convert.ToString(dr("TurnoutCount"))
            RejectCount.Text = Convert.ToString(dr("RejectCount"))
            OtherCount.Text = Convert.ToString(dr("OtherCount"))
            AtteRate.Text = ""
            If Convert.ToString(dr("AtteRate")) <> "" Then AtteRate.Text = dr("AtteRate") * 100
            AtteRate.Text = TIMS.ROUND(AtteRate.Text)
            Me.OLessonTeah1.Text = ""
            Me.OLessonTeah1Value.Value = ""
            If dr("TechID").ToString <> "" Then
                Me.OLessonTeah1.Text = TIMS.Get_TeachCName(dr("TechID").ToString, objconn)
                Me.OLessonTeah1Value.Value = dr("TechID").ToString
            End If
            Me.OLessonTeah2.Text = ""
            Me.OLessonTeah2Value.Value = ""
            If Convert.ToString(dr("TechID2")) <> "" Then
                Me.OLessonTeah2.Text = TIMS.Get_TeachCName(dr("TechID2").ToString, objconn)
                Me.OLessonTeah2Value.Value = dr("TechID2").ToString
            End If
            If dr("CourseName").ToString <> "" Then Me.CourseName.Text = dr("CourseName").ToString

            Common.SetListItem(DATA81, Convert.ToString(dr("Data81")))
            Common.SetListItem(DATA82, Convert.ToString(dr("Data82")))
            Common.SetListItem(DATA83, Convert.ToString(dr("Data83")))
            Common.SetListItem(DATA84, Convert.ToString(dr("Data84")))

            For i As Integer = 0 To Item1.Items.Count - 1
                If Convert.ToString(dr("Item1")) = Item1.Items(i).Value Then Item1.Items(i).Selected = True
            Next
            For i As Integer = 0 To Item2.Items.Count - 1
                If Convert.ToString(dr("Item2")) = Item2.Items(i).Value Then Item2.Items(i).Selected = True
            Next
            For i As Integer = 0 To Item3.Items.Count - 1
                If Convert.ToString(dr("item3")) = Item3.Items(i).Value Then Item3.Items(i).Selected = True
            Next
            For i As Integer = 0 To Item4.Items.Count - 1
                If Convert.ToString(dr("item4")) = Item4.Items(i).Value Then Item4.Items(i).Selected = True
            Next
            Me.Stud_Name.Text = dr("Stud_Name").ToString
            Me.Stud_Name2.Text = dr("Stud_Name2").ToString
            For i As Integer = 0 To SItem1.Items.Count - 1
                If Convert.ToString(dr("SItem1")) = SItem1.Items(i).Value Then SItem1.Items(i).Selected = True
            Next
            For i As Integer = 0 To SItem2.Items.Count - 1
                If Convert.ToString(dr("SItem2")) = SItem2.Items(i).Value Then SItem2.Items(i).Selected = True
            Next

            LItem1.Checked = False
            LItem2.Checked = False
            If dr("LItem1").ToString = "1" Then
                LItem1.Checked = True
                'LItem_TR.Attributes.Add("Style", "none")
                'LItem_TR2.Attributes.Add("Style", "none")
                LItem_TR.Style.Item("display") = "none"
                LItem_TR2.Style.Item("display") = "none"
            End If
            If dr("LItem2").ToString = "1" Then
                LItem2.Checked = True
                'LItem_TR.Attributes.Add("Style", "inline")
                'LItem_TR2.Attributes.Add("Style", "inline")
                LItem_TR.Style.Item("display") = "inline"
                LItem_TR2.Style.Item("display") = "inline"
            End If

            LItem2_1.Checked = False
            If dr("LItem2_1").ToString = "1" Then
                LItem2_1.Checked = True
            End If
            If dr("LItem2_1_Date").ToString <> "" Then
                If flag_ROC Then
                    LItem2_1_Date.Text = TIMS.cdate17(dr("LItem2_1_Date"))  'edit，by:2018108
                Else
                    LItem2_1_Date.Text = Common.FormatDate(dr("LItem2_1_Date").ToString)  'edit，by:2018108
                End If
            End If

            LItem2_2.Checked = False
            If dr("LItem2_2").ToString = "1" Then
                LItem2_2.Checked = True
            End If
            If Convert.ToString(dr("LItem2_2b")) <> "" Then
                Call TIMS.SetCblValue(cblLItem2_2b, dr("LItem2_2b"))
            End If
            If dr("LItem2_2_Note").ToString <> "" Then
                LItem2_2_Note.Text = dr("LItem2_2_Note").ToString
            Else
                LItem2_2_Note.Text = ""
            End If
            If dr("LItem2_3_Note").ToString <> "" Then
                LItem2_3_Note.Text = dr("LItem2_3_Note").ToString
            Else
                LItem2_3_Note.Text = ""
            End If
            If dr("SItem1_Note").ToString <> "" Then
                Me.SItem1_Note.Text = dr("SItem1_Note").ToString
            End If
            If dr("SItem2_Note").ToString <> "" Then
                Me.SItem2_Note.Text = dr("SItem2_Note").ToString
            End If

            If dr("OrgID").ToString <> "" Then
                OrgName.Text = TIMS.GET_OrgName(dr("OrgID"), objconn)
                OrgID.Value = dr("OrgID").ToString 'TIMS.GET_OrgName(dr("OrgID").ToString)
            End If
            CurseName.Text = Convert.ToString(dr("CurseName"))
            VisitorName.Text = Convert.ToString(dr("VisitorName"))

            TIMS.SetCblValue(cblSITEM51, Convert.ToString(dr("SITEM51")))
            TIMS.SetCblValue(cblSITEM61, Convert.ToString(dr("SITEM61")))
            TIMS.SetCblValue(cblSITEM71, Convert.ToString(dr("SITEM71")))
            TIMS.SetCblValue(cblSITEM81, Convert.ToString(dr("SITEM81")))
            SITEM51_NOTE.Text = Convert.ToString(dr("SITEM51_NOTE"))
            SITEM61_NOTE.Text = Convert.ToString(dr("SITEM61_NOTE"))
            SITEM71_NOTE.Text = Convert.ToString(dr("SITEM71_NOTE"))
            SITEM81_NOTE.Text = Convert.ToString(dr("SITEM81_NOTE"))

            chkB_NOINC5.Checked = False
            If Convert.ToString(dr("NOINC5")) = TIMS.cst_YES Then chkB_NOINC5.Checked = True

        End If

        Dim drC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If Not drC Is Nothing Then
            center.Text = drC("OrgName")
            RIDValue.Value = drC("RID").ToString
            TMID1.Text = Convert.ToString(drC("TrainName2"))
            TMIDValue1.Value = drC("TMID")
            OCID1.Text = drC("ClassCName2")
            OCIDValue1.Value = drC("OCID")

            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            Button3.Disabled = True
        End If
    End Sub

    '檢查有效儲存資料
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        ApplyDate.Text = TIMS.ClearSQM(ApplyDate.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If Trim(ApplyDate.Text) <> "" Then ApplyDate.Text = Trim(ApplyDate.Text) Else ApplyDate.Text = ""

        If OCIDValue1.Value = "" Then
            Errmsg += "班級尚未選擇，請重新選擇" & vbCrLf
        Else
            If Not IsNumeric(OCIDValue1.Value) Then Errmsg += "班級選擇有誤，請重新選擇" & vbCrLf
        End If

        If ApplyDate.Text <> "" Then
            If flag_ROC Then
                If Not TIMS.IsDate7(ApplyDate.Text) Then Errmsg += "訪查時間 日期格式有誤" & vbCrLf  'edit，by:20181018
            Else
                If Not TIMS.IsDate1(ApplyDate.Text) Then Errmsg += "訪查時間 日期格式有誤" & vbCrLf  'edit，by:20181018
            End If

            If Errmsg = "" Then
                If flag_ROC Then
                    ApplyDate.Text = TIMS.cdate7(ApplyDate.Text)  'edit，by:20181018
                Else
                    ApplyDate.Text = CDate(ApplyDate.Text).ToString("yyyy/MM/dd")  'edit，by:20181018
                End If
            End If
        Else
            Errmsg += "訪查時間 日期 為必填" & vbCrLf
        End If

        If IsNumeric(Applytime_HH.Value) Then
            If CInt(Applytime_HH.Value) > 23 OrElse CInt(Applytime_HH.Value) < 0 Then Errmsg += "訪查時間 時間(幾點)格式有誤(00~23)，請重新填寫" & vbCrLf
            Applytime_HH.Value = CStr(CInt(Applytime_HH.Value))
        Else
            Errmsg += "訪查時間 時間(幾點)格式有誤(00~23)，請重新填寫" & vbCrLf
        End If

        If IsNumeric(Applytime_MM.Value) Then
            If CInt(Applytime_MM.Value) > 59 OrElse CInt(Applytime_MM.Value) < 0 Then Errmsg += "訪查時間 時間(幾分)格式有誤(00~59)，請重新填寫" & vbCrLf
            Applytime_MM.Value = CStr(CInt(Applytime_MM.Value))
        Else
            Errmsg += "訪查時間 時間(幾分)格式有誤(00~59)，請重新填寫" & vbCrLf
        End If

        If Errmsg = "" Then
            If Len(Applytime_HH.Value) < 2 Then Applytime_HH.Value = "0" & Applytime_HH.Value
            If Len(Applytime_MM.Value) < 2 Then Applytime_MM.Value = "0" & Applytime_MM.Value
        End If

        '當日教師
        If Trim(Me.OLessonTeah1Value.Value) = "" Then
            '當日教師
            OLessonTeah1.Text = ""
            OLessonTeah1Value.Value = ""
            ODegreeID1.Value = ""
            ODegreeIDValue1.Value = ""
            Errmsg += "請輸入 當日教師" & vbCrLf
        End If

        Dim int_Len1 As Integer = 0
        '當日課程
        int_Len1 = 0
        If Trim(CourseName.Text) <> "" Then
            CourseName.Text = Trim(CourseName.Text)
            int_Len1 = CourseName.MaxLength
            If Len(CourseName.Text) > int_Len1 Then Errmsg += "當日課程 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            CourseName.Text = ""
            Errmsg += "請輸入 當日課程" & vbCrLf
        End If

        If Not IsNumeric(AuthCount.Text) Then Errmsg += "應到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        If Not IsNumeric(TurthCount.Text) Then Errmsg += "實到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        If Not IsNumeric(TruancyCount.Text) Then Errmsg += "未到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        If Not IsNumeric(TurnoutCount.Text) Then Errmsg += "請假人數 應為數字格式有誤，請重新填寫" & vbCrLf
        If Not IsNumeric(RejectCount.Text) Then Errmsg += "退訓人數 應為數字格式有誤，請重新填寫" & vbCrLf
        If Not IsNumeric(OtherCount.Text) Then Errmsg += "其他人數 應為數字格式有誤，請重新填寫" & vbCrLf

        AtteRate.Text = TIMS.ClearSQM(AtteRate.Text)
        If Not IsNumeric(AtteRate.Text) Then Errmsg += "出席率 應為數字格式有誤，請重新填寫" & vbCrLf
        If Errmsg = "" AndAlso AtteRate.Text <> "" Then
            'If Not TIMS.IsNumeric2(AtteRate.Text) Then Errmsg &= "出席率 格式有誤，應為正整數數字格式 " & vbCrLf 'int
            If Not TIMS.IsNumeric1(AtteRate.Text) Then
                Errmsg &= "出席率 格式有誤，應為正整數數字格式 " & vbCrLf 'int
            ElseIf CInt(AtteRate.Text) < 0 Then
                Errmsg &= "出席率 格式有誤，應為正整數數字格式 " & vbCrLf 'int
            End If
        End If
        If Errmsg = "" AndAlso AtteRate.Text <> "" Then If Val(AtteRate.Text) > 100 Then Errmsg &= "出席率 數字格式 不可超過100" & vbCrLf 'int

        If Errmsg = "" Then
            AuthCount.Text = CInt(AuthCount.Text)
            TurthCount.Text = CInt(TurthCount.Text)
            TruancyCount.Text = CInt(TruancyCount.Text)
            TurnoutCount.Text = CInt(TurnoutCount.Text)
            RejectCount.Text = CInt(RejectCount.Text)
            OtherCount.Text = CInt(OtherCount.Text)
        End If

        If Trim(Stud_Name.Text) <> "" Then
            Stud_Name.Text = Trim(Stud_Name.Text)
            int_Len1 = Stud_Name.MaxLength
            If Len(Stud_Name.Text) > int_Len1 Then
                Errmsg += "三、現場訪查實況：抽訪學員之姓名1 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            Stud_Name.Text = ""
            Errmsg += "請輸入 三、現場訪查實況：抽訪學員之姓名1" & vbCrLf
        End If

        If Trim(Stud_Name2.Text) <> "" Then
            Stud_Name2.Text = Trim(Stud_Name2.Text)
            int_Len1 = Stud_Name2.MaxLength
            If Len(Stud_Name2.Text) > int_Len1 Then Errmsg += "三、現場訪查實況：抽訪學員之姓名2 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            Stud_Name2.Text = ""
            Errmsg += "請輸入 三、現場訪查實況：抽訪學員之姓名2" & vbCrLf
        End If

        If LItem2_1_Date.Text <> "" Then
            If Not TIMS.IsDate7(LItem2_1_Date.Text) Then Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下： 日期格式有誤" & vbCrLf
            If Errmsg = "" Then
                If flag_ROC Then
                    LItem2_1_Date.Text = TIMS.cdate7(LItem2_1_Date.Text)  'edit，by:20181018
                Else
                    LItem2_1_Date.Text = CDate(LItem2_1_Date.Text).ToString("yyyy/MM/dd")  'edit，by:20181018
                End If
            End If
        End If

        If Trim(LItem2_2_Note.Text) <> "" Then
            LItem2_2_Note.Text = Trim(LItem2_2_Note.Text)
            int_Len1 = LItem2_2_Note.MaxLength
            If Len(LItem2_2_Note.Text) > int_Len1 Then Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下：(2)其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            LItem2_2_Note.Text = ""
        End If

        If Trim(LItem2_3_Note.Text) <> "" Then
            LItem2_3_Note.Text = Trim(LItem2_3_Note.Text)
            int_Len1 = LItem2_3_Note.MaxLength
            If Len(LItem2_3_Note.Text) > int_Len1 Then Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下：(3)其他補充說明 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            LItem2_3_Note.Text = ""
        End If

        If Trim(SItem1_Note.Text) <> "" Then
            SItem1_Note.Text = Trim(SItem1_Note.Text)
            int_Len1 = SItem1_Note.MaxLength
            If Len(SItem1_Note.Text) > int_Len1 Then Errmsg += "三、現場訪查實況1.其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            SItem1_Note.Text = ""
        End If

        If Trim(SItem2_Note.Text) <> "" Then
            SItem2_Note.Text = Trim(SItem2_Note.Text)
            int_Len1 = SItem2_Note.MaxLength
            If Len(SItem2_Note.Text) > int_Len1 Then Errmsg += "三、現場訪查實況2.其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            SItem2_Note.Text = ""
        End If

        'If Trim(SItem3_Note.Text) <> "" Then
        '    SItem3_Note.Text = Trim(SItem3_Note.Text)
        '    int_Len1 = SItem3_Note.MaxLength
        '    If Len(SItem3_Note.Text) > int_Len1 Then
        '        Errmsg += "三、現場訪查實況3.其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        '    End If
        'Else
        '    SItem3_Note.Text = ""
        'End If

        If Trim(CurseName.Text) <> "" Then
            CurseName.Text = Trim(CurseName.Text)
            int_Len1 = CurseName.MaxLength
            If Len(CurseName.Text) > int_Len1 Then Errmsg += "培訓單位人員姓名 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            CurseName.Text = ""
            Errmsg += "培訓單位人員姓名 為必填" & vbCrLf
        End If

        If Trim(VisitorName.Text) <> "" Then
            VisitorName.Text = Trim(VisitorName.Text)
            int_Len1 = VisitorName.MaxLength
            If Len(VisitorName.Text) > 10 Then Errmsg += "訪視人員姓名 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
        Else
            VisitorName.Text = ""
            Errmsg += "訪視人員姓名 為必填" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Sub sSaveData1()
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqDOCID As String = TIMS.ClearSQM(Request("DOCID"))
        Dim rqState As String = TIMS.ClearSQM(Request("State"))
        Dim rqType As String = TIMS.ClearSQM(Request("Type"))
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)

        Try
            Dim sql As String = ""
            Dim da As SqlDataAdapter = Nothing
            Dim dr As DataRow = Nothing
            Dim dt As DataTable = Nothing
            Dim SeqNo As Integer = 0

            If UCase(rqState) = "ADD" Then '表示新增狀態
                '先取出最大SeqNo
                sql = "SELECT MAX(SeqNO) num FROM dbo.CLASS_UNEXPECTVISITOR WHERE OCID='" & OCIDValue1.Value & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If IsDBNull(dr("num")) Then
                        SeqNo = 1
                    Else
                        SeqNo = CInt(dr("num")) + 1
                    End If
                End If

                sql = "SELECT * FROM Class_UnexpectVisitor WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                dr = dt.NewRow
                dt.Rows.Add(dr)
            Else
                sql = "SELECT * FROM Class_UnexpectVisitor WHERE OCID='" & rqOCID & "' and SeqNo='" & rqSeqNo & "'"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                dr = dt.Rows(0)
                SeqNo = rqSeqNo
            End If

            dr("OCID") = OCIDValue1.Value
            dr("SeqNo") = SeqNo

            If flag_ROC Then
                dr("ApplyDate") = TIMS.cdate18(ApplyDate.Text)  'edit，by:20181018
            Else
                dr("ApplyDate") = ApplyDate.Text  'edit，by:20181018
            End If

            If Len(Applytime_HH.Value) < 2 Then
                Applytime_HH.Value = "0" & Applytime_HH.Value
            End If
            If Len(Applytime_MM.Value) < 2 Then
                Applytime_MM.Value = "0" & Applytime_MM.Value
            End If
            dr("ApplyTime") = Applytime_HH.Value & ":" & Applytime_MM.Value

            dr("AuthCount") = 0
            dr("TurthCount") = 0
            dr("TruancyCount") = 0
            dr("TurnoutCount") = 0
            dr("RejectCount") = 0
            dr("OtherCount") = 0
            If AuthCount.Text <> "" Then dr("AuthCount") = AuthCount.Text
            If TurthCount.Text <> "" Then dr("TurthCount") = TurthCount.Text
            If TruancyCount.Text <> "" Then dr("TruancyCount") = TruancyCount.Text
            If TurnoutCount.Text <> "" Then dr("TurnoutCount") = TurnoutCount.Text
            If RejectCount.Text <> "" Then dr("RejectCount") = RejectCount.Text
            If OtherCount.Text <> "" Then dr("OtherCount") = OtherCount.Text

            dr("TechID") = Me.OLessonTeah1Value.Value '不可為空白

            If Me.OLessonTeah2Value.Value <> "" Then '可為空白
                dr("TechID2") = Me.OLessonTeah2Value.Value
            Else
                dr("TechID2") = Convert.DBNull
            End If

            dr("CourseName") = ""
            If CourseName.Text <> "" Then dr("CourseName") = CourseName.Text
            'AtteRate.Text = ""
            'If Convert.ToString(dr("AtteRate")) <> "" Then AtteRate.Text = dr("AtteRate") * 100
            AtteRate.Text = TIMS.ClearSQM(AtteRate.Text)
            If AtteRate.Text = "" Then dr("AtteRate") = Convert.DBNull
            If AtteRate.Text <> "" Then dr("AtteRate") = TIMS.ROUND(Val(AtteRate.Text) / Val(100), 3)

            dr("Data81") = DATA81.SelectedValue
            dr("Data82") = DATA82.SelectedValue
            dr("Data83") = DATA83.SelectedValue
            dr("Data84") = DATA84.SelectedValue

            dr("Item1") = Item1.SelectedValue
            dr("Item2") = Item2.SelectedValue
            dr("Item3") = Item3.SelectedValue
            dr("Item4") = Item4.SelectedValue
            dr("Stud_Name") = Stud_Name.Text
            dr("Stud_Name2") = Stud_Name2.Text

            dr("SItem1") = SItem1.SelectedValue
            dr("SItem2") = SItem2.SelectedValue
            'dr("SItem3") = SItem3.SelectedValue
            '1:是 2:否 3:其他 9:停用
            dr("SItem3") = "9"

            '四、現場處理說明
            dr("LItem1") = IIf(LItem1.Checked, "1", "2")
            dr("LItem2") = IIf(LItem2.Checked, "1", "2")

            dr("LItem2_1") = "2" '未勾選
            If LItem2_1.Checked Then
                dr("LItem2_1") = "1"
            End If

            dr("LItem2_1_Date") = Convert.DBNull
            If Me.LItem2_1_Date.Text <> "" Then
                If flag_ROC Then
                    dr("LItem2_1_Date") = TIMS.cdate18(Me.LItem2_1_Date.Text)  'edit，by:20181018
                Else
                    dr("LItem2_1_Date") = Me.LItem2_1_Date.Text  'edit，by:20181018
                End If
            End If

            dr("LItem2_2") = "2" '未勾選
            If LItem2_2.Checked Then
                dr("LItem2_2") = "1"
            End If

            Dim sLItem2_2b As String = TIMS.GetCblValue(cblLItem2_2b)
            If sLItem2_2b <> "" Then
                dr("LItem2_2b") = sLItem2_2b
            Else
                dr("LItem2_2b") = Convert.DBNull
            End If
            dr("LItem2_2_Note") = ""
            If Me.LItem2_2_Note.Text <> "" Then
                dr("LItem2_2_Note") = Me.LItem2_2_Note.Text
            End If

            dr("LItem2_3_Note") = ""
            If LItem2_3_Note.Text <> "" Then
                dr("LItem2_3_Note") = Me.LItem2_3_Note.Text
            End If

            dr("SItem1_Note") = ""
            If Me.SItem1_Note.Text <> "" Then
                dr("SItem1_Note") = Me.SItem1_Note.Text
            End If
            dr("SItem2_Note") = ""
            If Me.SItem2_Note.Text <> "" Then
                dr("SItem2_Note") = Me.SItem2_Note.Text
            End If
            'dr("SItem3_Note") = ""
            'If Me.SItem3_Note.Text <> "" Then
            '    dr("SItem3_Note") = Me.SItem3_Note.Text
            'End If

            dr("OrgID") = sm.UserInfo.OrgID
            dr("CurseName") = CurseName.Text
            dr("VisitorName") = VisitorName.Text
            dr("RID") = sm.UserInfo.RID

            Dim vSITEM51 As String = TIMS.GetCblValue(cblSITEM51)
            Dim vSITEM61 As String = TIMS.GetCblValue(cblSITEM61)
            Dim vSITEM71 As String = TIMS.GetCblValue(cblSITEM71)
            Dim vSITEM81 As String = TIMS.GetCblValue(cblSITEM81)
            dr("SITEM51") = vSITEM51
            dr("SITEM61") = vSITEM61
            dr("SITEM71") = vSITEM71
            dr("SITEM81") = vSITEM81
            dr("SITEM51_NOTE") = TIMS.ClearSQM(SITEM51_NOTE.Text)
            dr("SITEM61_NOTE") = TIMS.ClearSQM(SITEM61_NOTE.Text)
            dr("SITEM71_NOTE") = TIMS.ClearSQM(SITEM71_NOTE.Text)
            dr("SITEM81_NOTE") = TIMS.ClearSQM(SITEM81_NOTE.Text)

            Dim vNOINC5 As String = TIMS.cst_NO
            If chkB_NOINC5.Checked Then vNOINC5 = TIMS.cst_YES
            dr("NOINC5") = vNOINC5

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now

            DbAccess.UpdateDataTable(dt, da)

            If Not Session("SearchStr") Is Nothing Then Session("SearchStr") = Session("SearchStr")

            'Common.RespWrite(Me, "<script> alert('儲存成功');")
            'If rqType = "CV" Then
            '    Page.RegisterStartupScript("", "<script>alert('儲存成功'); window.location.href='CP_01_008.aspx?ID=" & Request("ID") & "';</script>")
            'Else
            '    Page.RegisterStartupScript("", "<script>alert('儲存成功'); window.location.href='CP_01_006.aspx?ID=" & Request("ID") & "';</script>")
            'End If
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            Common.MessageBox(Me, "儲存失敗!!")
            Exit Sub
            'Common.MessageBox(Me, ex.ToString)
            'Common.RespWrite(Me, ex)
            'Throw ex
        End Try

        '訪視計畫表用
        'Dim rqType As String = UCase(TIMS.ClearSQM(Request("Type")))
        Dim uUrl1 As String = "CP/01/CP_01_006.aspx?ID=" & Request("ID")
        Select Case rqType
            Case "CV"
                uUrl1 = "CP/01/CP_01_008.aspx?ID=" & Request("ID")
        End Select
        Dim sMsg As String = "儲存成功"
        Call TIMS.blockAlert(Me, sMsg, uUrl1)
        'Exit Sub
    End Sub

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call sSaveData1()
    End Sub

    'Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
    '    'Session("_SearchStr") = Me.ViewState("_SearchStr")
    '    If Not Session("SearchStr") Is Nothing Then Session("SearchStr") = Session("SearchStr")
    'End Sub

    Sub sBack2Prev()
        '訪視計畫表用
        Dim rqType As String = UCase(TIMS.ClearSQM(Request("Type")))
        Select Case rqType
            Case "CV"
                Dim url1 As String = "CP_01_008.aspx?ID=" & Request("ID")
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case Else
                Dim url1 As String = "CP_01_006.aspx?ID=" & Request("ID")
                TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub Button5_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.ServerClick
        '訪視計畫表用
        If Not Session("SearchStr") Is Nothing Then Session("SearchStr") = Session("SearchStr")
        Call sBack2Prev()
    End Sub
End Class