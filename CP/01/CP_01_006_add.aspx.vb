Partial Class CP_01_006_add
    Inherits AuthBasePage

    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("CLASS_UNEXPECTVISITOR", objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "COURSENAME", CourseName)
        Call TIMS.sUtl_SetMaxLen(dt, "DATA4_NOTE", Data4_Note)
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

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call sUtl_PageInit1()
        '檢查Session是否存在--------------------------End
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqDOCID As String = TIMS.ClearSQM(Request("DOCID"))
        Dim rqState As String = TIMS.ClearSQM(Request("State"))
        Dim rqType As String = TIMS.ClearSQM(Request("Type"))
        If iPYNum >= 3 Then
            Call TIMS.CloseDbConn(objconn)
            Dim rqMID As String = TIMS.Get_MRqID(Me)
            Dim sUrl1 As String = "CP_01_006_add8.aspx?ID=" & rqMID
            Dim sUrl2 As String = ""
            TIMS.SetMyValue(sUrl2, "OCID", rqOCID)
            TIMS.SetMyValue(sUrl2, "SeqNo", rqSeqNo)
            TIMS.SetMyValue(sUrl2, "DOCID", rqDOCID)
            TIMS.SetMyValue(sUrl2, "State", rqState)
            TIMS.SetMyValue(sUrl2, "Type", rqType)
            Server.Transfer(sUrl1 & sUrl2)
        End If

        If Not IsPostBack Then
            Me.OLessonTeah1.Attributes.Add("onDblClick", "javascript:OpenLessonTeah1('Add');")
            teacherbtn.Attributes.Add("onClick", "javascript:OpenLessonTeah1('Add');")

            Me.OLessonTeah2.Attributes.Add("onDblClick", "javascript:OpenLessonTeah2('Add');")
            teacherbtn2.Attributes.Add("onClick", "javascript:OpenLessonTeah2('Add');")

            Me.Applytime_HH.Value = "09"
            Me.Applytime_MM.Value = "00"

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
                    center.Text = Replace(MyValue, "%26", "&")
                    RIDValue.Value = TIMS.GetMyValue(Session("SearchStr"), "RIDValue")

                    MyValue = TIMS.GetMyValue(Session("SearchStr"), "TMID1")
                    TMID1.Text = Replace(MyValue, "%26", "&")
                    MyValue = TIMS.GetMyValue(Session("SearchStr"), "OCID1")
                    OCID1.Text = Replace(MyValue, "%26", "&")
                    TMIDValue1.Value = TIMS.GetMyValue(Session("SearchStr"), "TMIDValue1")
                    OCIDValue1.Value = TIMS.GetMyValue(Session("SearchStr"), "OCIDValue1")
                    If rqOCID = "" Then rqOCID = OCIDValue1.Value
                End If
                If Not Session("SearchStr") Is Nothing Then
                    Session("SearchStr") = Session("SearchStr")
                End If
            End If

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

        Button1.Attributes("onclick") = "javascript:return chkdata()"

        'LItem1.Attributes("onclick") = "showTR();"
        'LItem2.Attributes("onclick") = "showTR();"

        If Not IsPostBack Then
            If rqOCID <> "" Then
                Call create(rqOCID, rqSeqNo)
            End If
            If rqDOCID <> "" Then
                Call create(rqDOCID, "")
            End If
        End If

        If rqState = "View" Then
            Button1.Visible = False
        End If
        'Button1.Enabled = False
        'If blnCanAdds Then Button1.Enabled = True

        'Dim FunDr As DataRow
        ''檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '    End If
        'End If


        '訪視計畫表用
        If rqType = "CV" AndAlso rqState = "View" Then
            Button4.Visible = False
            Button5.Visible = True
        ElseIf rqType = "CV" Then
            Button4.Visible = False
            Button5.Visible = True
        Else
            Button4.Visible = True
            Button5.Visible = False
        End If
    End Sub

    '取得資料庫資料
    Sub create(ByVal OCID As String, ByVal SeqNo As String)
        'Const cst_str1 As String = "出席率不佳,簽到退未落實,師資不符,課程內容不符,上課地點不符,其他："
        'Dim strA1 As String() = Split(cst_str1, ",")
        'cblLItem2_2b.Items.Clear()
        'With cblLItem2_2b
        '    For i As Integer = 0 To strA1.Length - 1
        '        .Items.Add(New ListItem(strA1(i), CStr(i + 1)))
        '    Next
        'End With
        '不可為空
        If OCID = "" Then OCID = "0"
        If SeqNo = "" Then SeqNo = "0"
        Dim sql As String = ""
        'Class_Visitor
        'sql = "SELECT * FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & OCID & "' and SeqNo='" & SeqNo & "'"
        sql = "SELECT * FROM CLASS_UNEXPECTVISITOR WHERE OCID=@OCID and SeqNo=@SeqNo "
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

        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)

            If flag_ROC Then
                ApplyDate.Text = TIMS.cdate17(dr("ApplyDate")) '西元轉民國，by:20181018
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

            If dr("CourseName").ToString <> "" Then
                Me.CourseName.Text = dr("CourseName").ToString
            End If
            'For i = 0 To Data1.Items.Count - 1
            '    If Convert.ToString(dr("Data1")) = Data1.Items(i).Value Then
            '        Data1.Items(i).Selected = True
            '    End If
            'Next
            For i As Integer = 0 To Data2.Items.Count - 1
                If Convert.ToString(dr("Data2")) = Data2.Items(i).Value Then
                    Data2.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Data3.Items.Count - 1
                If Convert.ToString(dr("Data3")) = Data3.Items(i).Value Then
                    Data3.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Data4.Items.Count - 1
                If Convert.ToString(dr("Data4")) = Data4.Items(i).Value Then
                    Data4.Items(i).Selected = True
                End If
            Next
            Me.Data4_Note.Text = dr("Data4_Note").ToString
            For i As Integer = 0 To Item1.Items.Count - 1
                If Convert.ToString(dr("Item1")) = Item1.Items(i).Value Then
                    Item1.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Item2.Items.Count - 1
                If Convert.ToString(dr("Item2")) = Item2.Items(i).Value Then
                    Item2.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Item3.Items.Count - 1
                If Convert.ToString(dr("item3")) = Item3.Items(i).Value Then
                    Item3.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To Item4.Items.Count - 1
                If Convert.ToString(dr("item4")) = Item4.Items(i).Value Then
                    Item4.Items(i).Selected = True
                End If
            Next
            Me.Stud_Name.Text = dr("Stud_Name").ToString
            Me.Stud_Name2.Text = dr("Stud_Name2").ToString
            For i As Integer = 0 To SItem1.Items.Count - 1
                If Convert.ToString(dr("SItem1")) = SItem1.Items(i).Value Then
                    SItem1.Items(i).Selected = True
                End If
            Next
            For i As Integer = 0 To SItem2.Items.Count - 1
                If Convert.ToString(dr("SItem2")) = SItem2.Items(i).Value Then
                    SItem2.Items(i).Selected = True
                End If
            Next
            'For i = 0 To SItem3.Items.Count - 1
            '    If Convert.ToString(dr("SItem3")) = SItem3.Items(i).Value Then
            '        SItem3.Items(i).Selected = True
            '    End If
            'Next
            If dr("LItem1").ToString = "1" Then
                LItem1.Checked = True
                'LItem_TR.Attributes.Add("Style", "none")
                'LItem_TR2.Attributes.Add("Style", "none")
                LItem_TR.Style.Item("display") = "none"
                LItem_TR2.Style.Item("display") = "none"
            Else
                LItem1.Checked = False
            End If
            If dr("LItem2").ToString = "1" Then
                LItem2.Checked = True
                'LItem_TR.Attributes.Add("Style", "inline")
                'LItem_TR2.Attributes.Add("Style", "inline")
                LItem_TR.Style.Item("display") = "inline"
                LItem_TR2.Style.Item("display") = "inline"
            Else
                LItem2.Checked = False
            End If

            LItem2_1.Checked = False
            If dr("LItem2_1").ToString = "1" Then
                LItem2_1.Checked = True
            End If
            If dr("LItem2_1_Date").ToString <> "" Then
                'LItem2_1_Date.Text = Common.FormatDate(dr("LItem2_1_Date").ToString)
                LItem2_1_Date.Text = TIMS.cdate17(dr("LItem2_1_Date")) '西元轉民國
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
            'If dr("SItem3_Note").ToString <> "" Then
            '    Me.SItem3_Note.Text = dr("SItem3_Note").ToString
            'End If
            If dr("OrgID").ToString <> "" Then
                OrgName.Text = TIMS.GET_OrgName(dr("OrgID"), objconn)
                OrgID.Value = dr("OrgID").ToString 'TIMS.GET_OrgName(dr("OrgID").ToString)
            End If
            CurseName.Text = Convert.ToString(dr("CurseName"))
            VisitorName.Text = Convert.ToString(dr("VisitorName"))
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
            If Not IsNumeric(OCIDValue1.Value) Then
                Errmsg += "班級選擇有誤，請重新選擇" & vbCrLf
            End If
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
            If CInt(Applytime_HH.Value) > 23 OrElse CInt(Applytime_HH.Value) < 0 Then
                Errmsg += "訪查時間 時間(幾點)格式有誤(00~23)，請重新填寫" & vbCrLf
            End If
            Applytime_HH.Value = CStr(CInt(Applytime_HH.Value))
        Else
            Errmsg += "訪查時間 時間(幾點)格式有誤(00~23)，請重新填寫" & vbCrLf
        End If

        If IsNumeric(Applytime_MM.Value) Then
            If CInt(Applytime_MM.Value) > 59 OrElse CInt(Applytime_MM.Value) < 0 Then
                Errmsg += "訪查時間 時間(幾分)格式有誤(00~59)，請重新填寫" & vbCrLf
            End If
            Applytime_MM.Value = CStr(CInt(Applytime_MM.Value))
        Else
            Errmsg += "訪查時間 時間(幾分)格式有誤(00~59)，請重新填寫" & vbCrLf
        End If

        If Errmsg = "" Then
            If Len(Applytime_HH.Value) < 2 Then
                Applytime_HH.Value = "0" & Applytime_HH.Value
            End If
            If Len(Applytime_MM.Value) < 2 Then
                Applytime_MM.Value = "0" & Applytime_MM.Value
            End If
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
            If Len(CourseName.Text) > int_Len1 Then
                Errmsg += "當日課程 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            CourseName.Text = ""
            Errmsg += "請輸入 當日課程" & vbCrLf
        End If

        If Not IsNumeric(AuthCount.Text) Then
            Errmsg += "應到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If
        If Not IsNumeric(TurthCount.Text) Then
            Errmsg += "實到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If
        If Not IsNumeric(TruancyCount.Text) Then
            Errmsg += "未到人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If
        If Not IsNumeric(TurnoutCount.Text) Then
            Errmsg += "請假人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If
        If Not IsNumeric(RejectCount.Text) Then
            Errmsg += "退訓人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If
        If Not IsNumeric(OtherCount.Text) Then
            Errmsg += "其他人數 應為數字格式有誤，請重新填寫" & vbCrLf
        End If

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
        If Errmsg = "" AndAlso AtteRate.Text <> "" Then
            If Val(AtteRate.Text) > 100 Then Errmsg &= "出席率 數字格式 不可超過100" & vbCrLf 'int
        End If

        If Errmsg = "" Then
            AuthCount.Text = CInt(AuthCount.Text)
            TurthCount.Text = CInt(TurthCount.Text)
            TruancyCount.Text = CInt(TruancyCount.Text)
            TurnoutCount.Text = CInt(TurnoutCount.Text)
            RejectCount.Text = CInt(RejectCount.Text)
            OtherCount.Text = CInt(OtherCount.Text)
        End If

        If Trim(Data4_Note.Text) <> "" Then
            Data4_Note.Text = Trim(Data4_Note.Text)
            int_Len1 = Data4_Note.MaxLength
            If Len(Data4_Note.Text) > int_Len1 Then
                Errmsg += "【一、資料文件查核】其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            Data4_Note.Text = ""
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
            If Len(Stud_Name2.Text) > int_Len1 Then
                Errmsg += "三、現場訪查實況：抽訪學員之姓名2 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            Stud_Name2.Text = ""
            Errmsg += "請輸入 三、現場訪查實況：抽訪學員之姓名2" & vbCrLf
        End If

        If LItem2_1_Date.Text <> "" Then
            If Not TIMS.IsDate7(LItem2_1_Date.Text) Then
                Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下： 日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                LItem2_1_Date.Text = TIMS.cdate7(LItem2_1_Date.Text) 'CDate(LItem2_1_Date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            'Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下： 日期 為必填" & vbCrLf
        End If

        If Trim(LItem2_2_Note.Text) <> "" Then
            LItem2_2_Note.Text = Trim(LItem2_2_Note.Text)
            int_Len1 = LItem2_2_Note.MaxLength
            If Len(LItem2_2_Note.Text) > int_Len1 Then
                Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下：(2)其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            LItem2_2_Note.Text = ""
        End If

        If Trim(LItem2_3_Note.Text) <> "" Then
            LItem2_3_Note.Text = Trim(LItem2_3_Note.Text)
            int_Len1 = LItem2_3_Note.MaxLength
            If Len(LItem2_3_Note.Text) > int_Len1 Then
                Errmsg += "四、現場處理說明：2.不預告抽訪結果需修正如下：(3)其他補充說明 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            LItem2_3_Note.Text = ""
        End If

        If Trim(SItem1_Note.Text) <> "" Then
            SItem1_Note.Text = Trim(SItem1_Note.Text)
            int_Len1 = SItem1_Note.MaxLength
            If Len(SItem1_Note.Text) > int_Len1 Then
                Errmsg += "三、現場訪查實況1.其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            SItem1_Note.Text = ""
        End If

        If Trim(SItem2_Note.Text) <> "" Then
            SItem2_Note.Text = Trim(SItem2_Note.Text)
            int_Len1 = SItem2_Note.MaxLength
            If Len(SItem2_Note.Text) > int_Len1 Then
                Errmsg += "三、現場訪查實況2.其他 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
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
            If Len(CurseName.Text) > int_Len1 Then
                Errmsg += "培訓單位人員姓名 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            CurseName.Text = ""
            Errmsg += "培訓單位人員姓名 為必填" & vbCrLf
        End If

        If Trim(VisitorName.Text) <> "" Then
            VisitorName.Text = Trim(VisitorName.Text)
            int_Len1 = VisitorName.MaxLength
            If Len(VisitorName.Text) > 10 Then
                Errmsg += "訪視人員姓名 長度超過系統範圍(" & CStr(int_Len1) & ")" & vbCrLf
            End If
        Else
            VisitorName.Text = ""
            Errmsg += "訪視人員姓名 為必填" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

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
                sql = "Select MAX(SeqNO) NUM FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & OCIDValue1.Value & "'"
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If IsDBNull(dr("num")) Then
                        SeqNo = 1
                    Else
                        SeqNo = CInt(dr("num")) + 1
                    End If
                End If

                sql = "SELECT * FROM CLASS_UNEXPECTVISITOR WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                dr = dt.NewRow
                dt.Rows.Add(dr)
            Else
                sql = "SELECT * FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & rqOCID & "' and SeqNo='" & rqSeqNo & "'"
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

            If AuthCount.Text <> "" Then dr("AuthCount") = Val(AuthCount.Text)
            If TurthCount.Text <> "" Then dr("TurthCount") = Val(TurthCount.Text)
            If TruancyCount.Text <> "" Then dr("TruancyCount") = Val(TruancyCount.Text)
            If TurnoutCount.Text <> "" Then dr("TurnoutCount") = Val(TurnoutCount.Text)
            If RejectCount.Text <> "" Then dr("RejectCount") = Val(RejectCount.Text)
            If OtherCount.Text <> "" Then dr("OtherCount") = Val(OtherCount.Text)

            AtteRate.Text = TIMS.ClearSQM(AtteRate.Text)
            If AtteRate.Text = "" Then dr("AtteRate") = Convert.DBNull
            If AtteRate.Text <> "" Then dr("AtteRate") = TIMS.Round(Val(AtteRate.Text) / 100, 3)

            dr("TechID") = Me.OLessonTeah1Value.Value '不可為空白

            If Me.OLessonTeah2Value.Value <> "" Then '可為空白
                dr("TechID2") = Me.OLessonTeah2Value.Value
            Else
                dr("TechID2") = Convert.DBNull
            End If

            dr("CourseName") = ""
            If CourseName.Text <> "" Then
                dr("CourseName") = CourseName.Text
            End If

            'dr("Data1") = Data1.SelectedValue
            dr("Data1") = "1"
            dr("Data2") = Data2.SelectedValue
            dr("Data3") = Data3.SelectedValue

            dr("Data4") = ""
            dr("Data4_Note") = ""
            If Data4.SelectedIndex <> -1 Then
                dr("Data4") = Data4.SelectedValue
            End If
            If Data4_Note.Text <> "" Then
                dr("Data4_Note") = Data4_Note.Text
            End If

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

            dr("LItem1") = "2"
            If LItem1.Checked Then
                dr("LItem1") = "1"
            End If

            dr("LItem2") = "2" '未勾選
            If LItem2.Checked Then
                dr("LItem2") = "1"
            End If

            dr("LItem2_1") = "2" '未勾選
            If LItem2_1.Checked Then
                dr("LItem2_1") = "1"
            End If

            dr("LItem2_1_Date") = Convert.DBNull
            If Me.LItem2_1_Date.Text <> "" Then
                dr("LItem2_1_Date") = TIMS.cdate18(Me.LItem2_1_Date.Text)
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
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da)

            If Not Session("SearchStr") Is Nothing Then
                Session("SearchStr") = Session("SearchStr")
            End If

#Region "todo for insert sys_trans_log用的寫法"
            ''新增
            'sql = " insert into CLASS_UNEXPECTVISITOR ( "
            'sql += " OCID,SEQNO,APPLYDATE,APPLYTIME,AUTHCOUNT,TURTHCOUNT,TRUANCYCOUNT "
            'sql += " ,TURNOUTCOUNT,REJECTCOUNT,OTHERCOUNT,TECHID,COURSENAME "
            'sql += " ,DATA1,DATA2,DATA3,DATA4,DATA4_NOTE,ITEM1,ITEM2,ITEM3,ITEM4 "
            'sql += " ,STUD_NAME,SITEM1,SITEM2,SITEM3,SITEM1_NOTE,SITEM2_NOTE "
            'sql += " ,LITEM1,LITEM2,LITEM2_1,LITEM2_1_DATE,LITEM2_2,LITEM2_2_NOTE "
            'sql += " ,CURSENAME,VISITORNAME,RID,MODIFYACCT,MODIFYDATE "
            'sql += " ,STUD_NAME2,LITEM2_3_NOTE,ORGID,TECHID2,LITEM2_2B,ATTERATE "
            'sql += " ) values ( "
            'sql += " @OCID,@SEQNO,@APPLYDATE,@APPLYTIME,@AUTHCOUNT,@TURTHCOUNT,@TRUANCYCOUNT "
            'sql += " ,@TURNOUTCOUNT,@REJECTCOUNT,@OTHERCOUNT,@TECHID,@COURSENAME "
            'sql += " ,@DATA1,@DATA2,@DATA3,@DATA4,@DATA4_NOTE,@ITEM1,@ITEM2,@ITEM3,@ITEM4 "
            'sql += " ,@STUD_NAME,@SITEM1,@SITEM2,@SITEM3,@SITEM1_NOTE,@SITEM2_NOTE "
            'sql += " ,@LITEM1,@LITEM2,@LITEM2_1,@LITEM2_1_DATE,@LITEM2_2,@LITEM2_2_NOTE "
            'sql += " ,@CURSENAME,@VISITORNAME,@RID,@MODIFYACCT,@MODIFYDATE "
            'sql += " ,@STUD_NAME2,@LITEM2_3_NOTE,@ORGID,@TECHID2,@LITEM2_2B,@ATTERATE "
            'sql += " ) "
            'Dim iCmd As New SqlCommand(sql, objconn)

            'sql = " UPDATE CLASS_UNEXPECTVISITOR SET "
            'sql += " APPLYDATE=@APPLYDATE "
            'sql += " ,APPLYTIME=@APPLYTIME "
            'sql += " ,AUTHCOUNT=@AUTHCOUNT "
            'sql += " ,TURTHCOUNT=@TURTHCOUNT "
            'sql += " ,TRUANCYCOUNT=@TRUANCYCOUNT "
            'sql += " ,TURNOUTCOUNT=@TURNOUTCOUNT "
            'sql += " ,REJECTCOUNT=@REJECTCOUNT "
            'sql += " ,OTHERCOUNT=@OTHERCOUNT "
            'sql += " ,TECHID=@TECHID "
            'sql += " ,COURSENAME=@COURSENAME "
            'sql += " ,DATA1=@DATA1 "
            'sql += " ,DATA2=@DATA2 "
            'sql += " ,DATA3=@DATA3 "
            'sql += " ,DATA4=@DATA4 "
            'sql += " ,DATA4_NOTE=@DATA4_NOTE "
            'sql += " ,ITEM1=@ITEM1 "
            'sql += " ,ITEM2=@ITEM2 "
            'sql += " ,ITEM3=@ITEM3 "
            'sql += " ,ITEM4=@ITEM4 "
            'sql += " ,STUD_NAME=@STUD_NAME "
            'sql += " ,SITEM1=@SITEM1 "
            'sql += " ,SITEM2=@SITEM2 "
            'sql += " ,SITEM3=@SITEM3 "
            'sql += " ,SITEM1_NOTE=@SITEM1_NOTE "
            'sql += " ,SITEM2_NOTE=@SITEM2_NOTE "
            'sql += " ,LITEM1=@LITEM1 "
            'sql += " ,LITEM2=@LITEM2 "
            'sql += " ,LITEM2_1=@LITEM2_1 "
            'sql += " ,LITEM2_1_DATE=@LITEM2_1_DATE "
            'sql += " ,LITEM2_2=@LITEM2_2 "
            'sql += " ,LITEM2_2_NOTE=@LITEM2_2_NOTE "
            'sql += " ,CURSENAME=@CURSENAME "
            'sql += " ,VISITORNAME=@VISITORNAME "
            'sql += " ,RID=@RID "
            'sql += " ,MODIFYACCT=@MODIFYACCT "
            'sql += " ,MODIFYDATE=@MODIFYDATE "
            'sql += " ,STUD_NAME2=@STUD_NAME2 "
            'sql += " ,LITEM2_3_NOTE=@LITEM2_3_NOTE "
            'sql += " ,ORGID=@ORGID "
            'sql += " ,TECHID2=@TECHID2 "
            'sql += " ,LITEM2_2B=@LITEM2_2B "
            'sql += " ,ATTERATE=@ATTERATE "
            'sql += " WHERE 1=1 "
            'sql += " And OCID=@OCID "
            'sql += " And SEQNO=@SEQNO "
            'Dim uCmd As New SqlCommand(sql, objconn)

            'If Len(Applytime_HH.Value) < 2 Then
            '    Applytime_HH.Value = "0" & Applytime_HH.Value
            'End If
            'If Len(Applytime_MM.Value) < 2 Then
            '    Applytime_MM.Value = "0" & Applytime_MM.Value
            'End If

            'Dim strLItem2_2b As String = TIMS.GetCblValue(cblLItem2_2b)

            ''sql = "SELECT * FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & rqOCID & "' and SeqNo='" & rqSeqNo & "'"
            ''dt = DbAccess.GetDataTable(sql, da, objconn)

            'If UCase(rqState) = "ADD" Then '表示新增狀態
            '    '先取出最大SeqNo
            '    sql = "Select MAX(SeqNO) NUM FROM CLASS_UNEXPECTVISITOR WHERE OCID='" & OCIDValue1.Value & "'"
            '    dr = DbAccess.GetOneRow(sql, objconn)
            '    If Not dr Is Nothing Then
            '        If IsDBNull(dr("num")) Then
            '            SeqNo = 1
            '        Else
            '            SeqNo = CInt(dr("num")) + 1
            '        End If
            '    End If

            '    With iCmd
            '        .Parameters.Clear()
            '        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            '        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = SeqNo
            '        .Parameters.Add("APPLYDATE", SqlDbType.VarChar).Value = ApplyDate.Text
            '        .Parameters.Add("APPLYTIME", SqlDbType.VarChar).Value = Applytime_HH.Value & ":" & Applytime_MM.Value
            '        .Parameters.Add("AUTHCOUNT", SqlDbType.Int).Value = IIf(AuthCount.Text = "", 0, Val(AuthCount.Text))
            '        .Parameters.Add("TURTHCOUNT", SqlDbType.VarChar).Value = IIf(TurthCount.Text = "", 0, Val(TurthCount.Text))
            '        .Parameters.Add("TRUANCYCOUNT", SqlDbType.VarChar).Value = IIf(TruancyCount.Text = "", 0, Val(TruancyCount.Text))
            '        .Parameters.Add("TURNOUTCOUNT", SqlDbType.VarChar).Value = IIf(TurnoutCount.Text = "", 0, Val(TurnoutCount.Text))
            '        .Parameters.Add("REJECTCOUNT", SqlDbType.VarChar).Value = IIf(RejectCount.Text = "", 0, Val(RejectCount.Text))
            '        .Parameters.Add("OTHERCOUNT", SqlDbType.VarChar).Value = IIf(OtherCount.Text = "", 0, Val(OtherCount.Text))
            '        .Parameters.Add("TECHID", SqlDbType.VarChar).Value = Me.OLessonTeah1Value.Value '不可為空白
            '        .Parameters.Add("COURSENAME", SqlDbType.NVarChar).Value = IIf(CourseName.Text = "", "", CourseName.Text)
            '        .Parameters.Add("DATA1", SqlDbType.VarChar).Value = "1"
            '        .Parameters.Add("DATA2", SqlDbType.VarChar).Value = Data2.SelectedValue
            '        .Parameters.Add("DATA3", SqlDbType.VarChar).Value = Data3.SelectedValue
            '        .Parameters.Add("DATA4", SqlDbType.VarChar).Value = IIf(Data4.SelectedValue <> -1, Data4.SelectedValue, "")
            '        .Parameters.Add("DATA4_NOTE", SqlDbType.NVarChar).Value = Data4_Note.Text
            '        .Parameters.Add("ITEM1", SqlDbType.VarChar).Value = Item1.SelectedValue
            '        .Parameters.Add("ITEM2", SqlDbType.VarChar).Value = Item2.SelectedValue
            '        .Parameters.Add("ITEM3", SqlDbType.VarChar).Value = Item3.SelectedValue
            '        .Parameters.Add("ITEM4", SqlDbType.VarChar).Value = Item4.SelectedValue
            '        .Parameters.Add("STUD_NAME", SqlDbType.NVarChar).Value = Stud_Name.Text
            '        .Parameters.Add("SITEM1", SqlDbType.VarChar).Value = SItem1.SelectedValue
            '        .Parameters.Add("SITEM2", SqlDbType.VarChar).Value = SItem2.SelectedValue
            '        .Parameters.Add("SITEM3", SqlDbType.VarChar).Value = "9"
            '        .Parameters.Add("SITEM1_NOTE", SqlDbType.VarChar).Value = IIf(Me.SItem1_Note.Text = "", Convert.DBNull, Me.SItem1_Note.Text)
            '        .Parameters.Add("SITEM2_NOTE", SqlDbType.VarChar).Value = IIf(Me.SItem2_Note.Text = "", Convert.DBNull, Me.SItem2_Note.Text)
            '        .Parameters.Add("LITEM1", SqlDbType.VarChar).Value = IIf(LItem1.Checked, "1", "2")
            '        .Parameters.Add("LITEM2", SqlDbType.VarChar).Value = IIf(LItem2.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_1", SqlDbType.VarChar).Value = IIf(LItem2_1.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_1_DATE", SqlDbType.VarChar).Value = IIf(Me.LItem2_1_Date.Text = "", Convert.DBNull, Me.LItem2_1_Date.Text)
            '        .Parameters.Add("LITEM2_2", SqlDbType.VarChar).Value = IIf(LItem2_2.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_2_NOTE", SqlDbType.NVarChar).Value = IIf(Me.LItem2_2_Note.Text = "", Convert.DBNull, Me.LItem2_2_Note.Text)
            '        .Parameters.Add("CURSENAME", SqlDbType.NVarChar).Value = CurseName.Text
            '        .Parameters.Add("VISITORNAME", SqlDbType.NVarChar).Value = VisitorName.Text
            '        .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            '        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '        .Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = Now
            '        .Parameters.Add("STUD_NAME2", SqlDbType.NVarChar).Value = Stud_Name2.Text
            '        .Parameters.Add("LITEM2_3_NOTE", SqlDbType.NVarChar).Value = IIf(Me.LItem2_3_Note.Text = "", Convert.DBNull, Me.LItem2_3_Note.Text)
            '        .Parameters.Add("ORGID", SqlDbType.VarChar).Value = sm.UserInfo.OrgID
            '        .Parameters.Add("TECHID2", SqlDbType.VarChar).Value = IIf(Me.OLessonTeah2Value.Value = "", Convert.DBNull, Me.OLessonTeah2Value.Value)
            '        .Parameters.Add("LITEM2_2B", SqlDbType.VarChar).Value = IIf(strLItem2_2b = "", Convert.DBNull, strLItem2_2b)

            '        If AtteRate.Text = "" Then
            '            .Parameters.Add("ATTERATE", SqlDbType.VarChar).Value = Convert.DBNull
            '        Else
            '            .Parameters.Add("ATTERATE", SqlDbType.VarChar).Value = TIMS.Round(Val(AtteRate.Text) / 100, 3)
            '        End If

            '        DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
            '    End With
            'Else
            '    '修改
            '    With uCmd
            '        .Parameters.Clear()
            '        .Parameters.Add("APPLYDATE", SqlDbType.VarChar).Value = ApplyDate.Text
            '        .Parameters.Add("APPLYTIME", SqlDbType.VarChar).Value = Applytime_HH.Value & ":" & Applytime_MM.Value
            '        .Parameters.Add("AUTHCOUNT", SqlDbType.Int).Value = IIf(AuthCount.Text = "", 0, Val(AuthCount.Text))
            '        .Parameters.Add("TURTHCOUNT", SqlDbType.VarChar).Value = IIf(TurthCount.Text = "", 0, Val(TurthCount.Text))
            '        .Parameters.Add("TRUANCYCOUNT", SqlDbType.VarChar).Value = IIf(TruancyCount.Text = "", 0, Val(TruancyCount.Text))
            '        .Parameters.Add("TURNOUTCOUNT", SqlDbType.VarChar).Value = IIf(TurnoutCount.Text = "", 0, Val(TurnoutCount.Text))
            '        .Parameters.Add("REJECTCOUNT", SqlDbType.VarChar).Value = IIf(RejectCount.Text = "", 0, Val(RejectCount.Text))
            '        .Parameters.Add("OTHERCOUNT", SqlDbType.VarChar).Value = IIf(OtherCount.Text = "", 0, Val(OtherCount.Text))
            '        .Parameters.Add("TECHID", SqlDbType.VarChar).Value = Me.OLessonTeah1Value.Value '不可為空白
            '        .Parameters.Add("COURSENAME", SqlDbType.NVarChar).Value = IIf(CourseName.Text = "", "", CourseName.Text)
            '        .Parameters.Add("DATA1", SqlDbType.VarChar).Value = "1"
            '        .Parameters.Add("DATA2", SqlDbType.VarChar).Value = Data2.SelectedValue
            '        .Parameters.Add("DATA3", SqlDbType.VarChar).Value = Data3.SelectedValue
            '        .Parameters.Add("DATA4", SqlDbType.VarChar).Value = IIf(Data4.SelectedValue <> -1, Data4.SelectedValue, "")
            '        .Parameters.Add("DATA4_NOTE", SqlDbType.NVarChar).Value = Data4_Note.Text
            '        .Parameters.Add("ITEM1", SqlDbType.VarChar).Value = Item1.SelectedValue
            '        .Parameters.Add("ITEM2", SqlDbType.VarChar).Value = Item2.SelectedValue
            '        .Parameters.Add("ITEM3", SqlDbType.VarChar).Value = Item3.SelectedValue
            '        .Parameters.Add("ITEM4", SqlDbType.VarChar).Value = Item4.SelectedValue
            '        .Parameters.Add("STUD_NAME", SqlDbType.NVarChar).Value = Stud_Name.Text
            '        .Parameters.Add("SITEM1", SqlDbType.VarChar).Value = SItem1.SelectedValue
            '        .Parameters.Add("SITEM2", SqlDbType.VarChar).Value = SItem2.SelectedValue
            '        .Parameters.Add("SITEM3", SqlDbType.VarChar).Value = "9"
            '        .Parameters.Add("SITEM1_NOTE", SqlDbType.VarChar).Value = IIf(Me.SItem1_Note.Text = "", Convert.DBNull, Me.SItem1_Note.Text)
            '        .Parameters.Add("SITEM2_NOTE", SqlDbType.VarChar).Value = IIf(Me.SItem2_Note.Text = "", Convert.DBNull, Me.SItem2_Note.Text)
            '        .Parameters.Add("LITEM1", SqlDbType.VarChar).Value = IIf(LItem1.Checked, "1", "2")
            '        .Parameters.Add("LITEM2", SqlDbType.VarChar).Value = IIf(LItem2.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_1", SqlDbType.VarChar).Value = IIf(LItem2_1.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_1_DATE", SqlDbType.VarChar).Value = IIf(Me.LItem2_1_Date.Text = "", Convert.DBNull, Me.LItem2_1_Date.Text)
            '        .Parameters.Add("LITEM2_2", SqlDbType.VarChar).Value = IIf(LItem2_2.Checked, "1", "2")
            '        .Parameters.Add("LITEM2_2_NOTE", SqlDbType.NVarChar).Value = IIf(Me.LItem2_2_Note.Text = "", Convert.DBNull, Me.LItem2_2_Note.Text)
            '        .Parameters.Add("CURSENAME", SqlDbType.NVarChar).Value = CurseName.Text
            '        .Parameters.Add("VISITORNAME", SqlDbType.NVarChar).Value = VisitorName.Text
            '        .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            '        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
            '        .Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = Now
            '        .Parameters.Add("STUD_NAME2", SqlDbType.NVarChar).Value = Stud_Name2.Text
            '        .Parameters.Add("LITEM2_3_NOTE", SqlDbType.NVarChar).Value = IIf(Me.LItem2_3_Note.Text = "", Convert.DBNull, Me.LItem2_3_Note.Text)
            '        .Parameters.Add("ORGID", SqlDbType.VarChar).Value = sm.UserInfo.OrgID
            '        .Parameters.Add("TECHID2", SqlDbType.VarChar).Value = IIf(Me.OLessonTeah2Value.Value = "", Convert.DBNull, Me.OLessonTeah2Value.Value)
            '        .Parameters.Add("LITEM2_2B", SqlDbType.VarChar).Value = IIf(strLItem2_2b = "", Convert.DBNull, strLItem2_2b)

            '        If AtteRate.Text = "" Then
            '            .Parameters.Add("ATTERATE", SqlDbType.VarChar).Value = Convert.DBNull
            '        Else
            '            .Parameters.Add("ATTERATE", SqlDbType.VarChar).Value = TIMS.Round(Val(AtteRate.Text) / 100, 3)
            '        End If

            '        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            '        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = rqSeqNo

            '        DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
            '    End With

            'End If
#End Region
            ' Common.RespWrite(Me, "<script> alert('儲存成功');")
            'If rqType = "CV" Then
            '    Common.RespWrite(Me, "location.href='CP_01_008.aspx?ID=" & Request("ID") & "';</script>")
            'Else
            '    Common.RespWrite(Me, "location.href='CP_01_006.aspx?ID=" & Request("ID") & "';</script>")
            'End If

            'Common.MessageBox(Me, "儲存成功")
            'If Request("Type") = "CV" Then
            '    TIMS.Utl_Redirect1(Me, "CP_01_008.aspx?ID=" & Request("ID"))
            'Else
            '    TIMS.Utl_Redirect1(Me, "CP_01_006.aspx?ID=" & Request("ID"))
            'End If

            If Request("Type") = "CV" Then
                Page.RegisterStartupScript("", "<script>alert('儲存成功'); window.location.href='CP_01_008.aspx?ID=" & Request("ID") & "';</script>")
            Else
                'Common.RespWrite(Me, "location.href='CP_01_006.aspx?ID=" & Request("ID") & "';</script>")
                Page.RegisterStartupScript("", "<script>alert('儲存成功'); window.location.href='CP_01_006.aspx?ID=" & Request("ID") & "';</script>")
            End If
        Catch ex As Exception
            Common.MessageBox(Me, "儲存失敗!!<br>" + ex.ToString)
            'Common.RespWrite(Me, ex)
            'Throw ex
        End Try
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        'Session("_SearchStr") = Me.ViewState("_SearchStr")
        If Not Session("SearchStr") Is Nothing Then
            Session("SearchStr") = Session("SearchStr")
        End If
        'Response.Redirect("CP_01_006.aspx?ID=" & Request("ID"))
        Dim url1 As String = "CP_01_006.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub Button5_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.ServerClick
        '訪視計畫表用
        If Not Session("SearchStr") Is Nothing Then
            Session("SearchStr") = Session("SearchStr")
        End If
        'Response.Redirect("CP_01_008.aspx?ID=" & Request("ID"))
        Dim url1 As String = "CP_01_008.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Protected Sub TurthCount_TextChanged(sender As Object, e As EventArgs)
    End Sub
End Class