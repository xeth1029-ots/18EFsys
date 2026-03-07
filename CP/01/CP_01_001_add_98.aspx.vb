Partial Class CP_01_001_add_98
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim objConn As SqlConnection

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            If Not Session("SearchStr") Is Nothing Then
                Me.ViewState("SearchStr") = Session("SearchStr")
                Dim MyArray As Array = Split(Session("SearchStr"), "&")
                Dim MyItem As String
                Dim MyValue As String
                For i As Integer = 0 To MyArray.Length - 1
                    MyItem = Split(MyArray(i), "=")(0)
                    MyValue = Split(MyArray(i), "=")(1)
                    Select Case MyItem
                        Case "center"
                            center.Text = MyValue
                        Case "RIDValue"
                            RIDValue.Value = MyValue
                        Case "TMID1"
                            TMID1.Text = MyValue
                        Case "OCID1"
                            OCID1.Text = MyValue
                        Case "TMIDValue1"
                            TMIDValue1.Value = MyValue
                        Case "OCIDValue1"
                            OCIDValue1.Value = MyValue
                        Case "end_date"
                            Me.EndDate.Value = MyValue
                    End Select
                Next
                Session("SearchStr") = Nothing
            End If
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
        End If

        objConn = DbAccess.GetConnection()
        Button1.Attributes("onclick") = "javascript:return chkdata()"

        If Not IsPostBack Then
            If Request("OCID") <> "" Then create(Request("OCID"), Request("SeqNo"))
            If Request("DOCID") <> "" Then create(Request("DOCID"), "")
        End If
        If Request("view") = "1" Then Button1.Visible = False

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Tplanid.Value = sm.UserInfo.TPlanID
        D3c.Attributes("onclick") = "return check('" & D3c.ClientID & "','" & D3c2.ClientID & "')"
        D4c.Attributes("onclick") = "return check('" & D4c.ClientID & "','" & D4c2.ClientID & "')"
        D5c.Attributes("onclick") = "return check('" & D5c.ClientID & "','" & D5c2.ClientID & "')"
        D6c.Attributes("onclick") = "return check('" & D6c.ClientID & "','" & D6c2.ClientID & "')"
        D7c.Attributes("onclick") = "return check('" & D7c.ClientID & "','" & D7c2.ClientID & "')"
        D8c.Attributes("onclick") = "return check('" & D8c.ClientID & "','" & D8c2.ClientID & "')"
        D3c2.Attributes("onclick") = "return check('" & D3c2.ClientID & "','" & D3c.ClientID & "')"
        D4c2.Attributes("onclick") = "return check('" & D4c2.ClientID & "','" & D4c.ClientID & "')"
        D5c2.Attributes("onclick") = "return check('" & D5c2.ClientID & "','" & D5c.ClientID & "')"
        D6c2.Attributes("onclick") = "return check('" & D6c2.ClientID & "','" & D6c.ClientID & "')"
        D7c2.Attributes("onclick") = "return check('" & D7c2.ClientID & "','" & D7c.ClientID & "')"
        D8c2.Attributes("onclick") = "return check('" & D8c2.ClientID & "','" & D8c.ClientID & "')"
    End Sub

    Sub create(ByVal OCID As String, ByVal SeqNo As String)
        Dim sql As String
        sql = " SELECT "
        sql += " SUBSTRING(ITEM10NOTE, 1, 4000) ITEM10NOTE, ITEM11, SUBSTRING(ITEM11PROS, 1, 4000) ITEM11PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM11NOTE, 1, 4000) ITEM11NOTE, ITEM12, SUBSTRING(ITEM12PROS, 1, 4000) ITEM12PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM12NOTE, 1, 4000) ITEM12NOTE, SUBSTRING(ITEM13NOTE, 1, 4000) ITEM13NOTE, " & vbCrLf
        sql += " AHEADJOBCOUNT, STUDYTICKETCOUNT, DATA9, DATACOPY9, DATA9NOTE, DATA10, DATACOPY10, DATA10NOTE, DATA11, DATACOPY11, " & vbCrLf
        sql += " DATA11NOTE, DATA12, ITEM14, ITEM14TECH, ITEM15, ITEM16, ITEM17, ITEM18, ITEM19, ITEM20, ITEM21, ITEM22, ITEM23, ITEM24, " & vbCrLf
        sql += " ITEM25, ITEM26, ITEM27, ITEM28, ITEM28COUNT, ITEM28_2, ITEM29, SUBSTRING(ITEM29PROS, 1, 4000) ITEM29PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM29NOTE, 1, 4000) ITEM29NOTE, ITEM30, SUBSTRING(ITEM31NOTE, 1, 4000) ITEM31NOTE, " & vbCrLf
        sql += " ITEM32, SUBSTRING(ITEM32NOTE, 1, 4000) ITEM32NOTE, D1C, D2C, D3C, D4C, D5C, D6C, D7C, D8C, D3C3, D4C3, D5C3, D6C3, " & vbCrLf
        sql += " D7C3, D8C3, OCID, SEQNO, APPLYDATE, AUTHCOUNT, TURTHCOUNT, TURNOUTCOUNT, TRUANCYCOUNT, REJECTCOUNT, DATA1, DATACOPY1, " & vbCrLf
        sql += " DATA1NOTE, DATA2, DATACOPY2, DATA2NOTE, DATA3, DATACOPY3, DATA3NOTE, DATA4, DATACOPY4, DATA4NOTE, DATA5, DATACOPY5, " & vbCrLf
        sql += " DATA5NOTE, DATA6, DATACOPY6, DATA6NOTE, DATA7, DATACOPY7, DATA7NOTE, ITEM1_1, ITEM1_2, ITEM1_3, " & vbCrLf
        sql += " SUBSTRING(ITEM1PROS, 1, 4000) ITEM1PROS, SUBSTRING(ITEM1NOTE, 1, 4000) ITEM1NOTE, ITEM2_1, " & vbCrLf
        sql += " ITEM2_2, SUBSTRING(ITEM2PROS, 1, 4000) ITEM2PROS, SUBSTRING(ITEM2NOTE, 1, 4000) ITEM2NOTE, " & vbCrLf
        sql += " ITEM3_1, ITEM3_1TECH, ITEM3_1TUTOR, ITEM3_2, SUBSTRING(ITEM3PROS, 1, 4000) ITEM3PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM3NOTE, 1, 4000) ITEM3NOTE, ITEM4_1, SUBSTRING(ITEM4PROS, 1, 4000) ITEM4PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM4NOTE, 1, 4000) ITEM4NOTE, ITEM5_1, SUBSTRING(ITEM5PROS, 1, 4000) ITEM5PROS, " & vbCrLf
        sql += " SUBSTRING(ITEM5NOTE, 1, 4000) ITEM5NOTE, ITEM6_1, ITEM6COUNT1,ITEM6COUNT2,ITEM6COUNT3,ITEM6NAMES, " & vbCrLf
        sql += " SUBSTRING(ITEM6NOTE, 1, 4000) ITEM6NOTE, SUBSTRING(ITEM7NOTE, 1, 4000) ITEM7NOTE, CURSENAME, " & vbCrLf
        sql += " VISITORNAME, MODIFYACCT, MODIFYDATE, RID, ISCLEAR, DATA8, VISITHOUR, TPERIOD, COMCOUNT, BOOKCOUNT, NETCOUNT, PSYCOUNT, " & vbCrLf
        sql += " ITEM8, SUBSTRING(ITEM8PROS, 1, 4000) ITEM8PROS, SUBSTRING(ITEM8NOTE, 1, 4000) ITEM8NOTE, " & vbCrLf
        sql += " ITEM9, ITEM9_TECH, SUBSTRING(ITEM9PROS, 1, 4000) ITEM9PROS, SUBSTRING(ITEM9NOTE, 1, 4000) ITEM9NOTE, " & vbCrLf
        sql += " ITEM10, SUBSTRING(ITEM10PROS, 1, 4000) ITEM10PROS " & vbCrLf
        sql += " FROM Class_Visitor WHERE OCID = '" & OCID & "' AND SeqNo = '" & SeqNo & "' "

        Dim i As Integer = 0
        Dim dr As DataRow = DbAccess.GetOneRow(sql)
        If Not dr Is Nothing Then
            ApplyDate.Text = dr("ApplyDate") '訪查日期
            For i = 0 To Data1.Items.Count - 1 '教學(訓練)日誌
                If Convert.ToString(dr("Data1")) = Data1.Items(i).Value Then Data1.Items(i).Selected = True
            Next
            DataCopy1.Text = Convert.ToString(dr("DataCopy1")) '如附件內容
            Data1Note.Text = Convert.ToString(dr("Data1Note")) '說明
            If dr("D1c").ToString <> "" Then '如附件的check 1表示有勾選
                D1c.Checked = True
            Else
                D1c.Checked = False
            End If

            For i = 0 To Data3.Items.Count - 1 '學員簽到(退)表
                If Convert.ToString(dr("Data3")) = Data3.Items(i).Value Then Data3.Items(i).Selected = True
            Next
            DataCopy3.Text = Convert.ToString(dr("DataCopy3")) '如附件內容
            Data3Note.Text = Convert.ToString(dr("Data3Note")) '說明

            If dr("D2c").ToString <> "" Then '如附件的check 1表示有勾選
                D2c.Checked = True
            Else
                D2c.Checked = False
            End If

            For i = 0 To Data5.Items.Count - 1 '請假單
                If Convert.ToString(dr("Data5")) = Data5.Items(i).Value Then Data5.Items(i).Selected = True
            Next
            DataCopy5.Text = Convert.ToString(dr("DataCopy5")) '如附件內容
            Data5Note.Text = Convert.ToString(dr("Data5Note")) '其他說明

            If dr("D3c").ToString <> "" Then
                If dr("D3c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D3c.Checked = True
                    D3c2.Checked = False
                ElseIf dr("D3c").ToString = 2 Then
                    D3c.Checked = False
                    D3c2.Checked = True
                End If
            End If

            If dr("D3c3").ToString <> "" Then D3c3.SelectedValue = dr("D3c3")  '說明的選項

            For i = 0 To Data6.Items.Count - 1 '退訓/提前就業
                If Convert.ToString(dr("Data6")) = Data6.Items(i).Value Then Data6.Items(i).Selected = True
            Next
            DataCopy6.Text = Convert.ToString(dr("DataCopy6")) '如附件內容
            Data6Note.Text = Convert.ToString(dr("Data6Note")) '其他說明

            If dr("D4c").ToString <> "" Then
                If dr("D4c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D4c.Checked = True
                    D4c2.Checked = False
                ElseIf dr("D4c").ToString = 2 Then
                    D4c.Checked = False
                    D4c2.Checked = True
                End If
            End If

            If dr("D4c3").ToString <> "" Then D4c3.SelectedValue = dr("D4c3")  '說明的選項

            For i = 0 To Data7.Items.Count - 1 '勞保加/退
                If Convert.ToString(dr("Data7")) = Data7.Items(i).Value Then Data7.Items(i).Selected = True
            Next
            DataCopy7.Text = Convert.ToString(dr("DataCopy7")) '如附件內容
            Data7Note.Text = Convert.ToString(dr("Data7Note")) '其他說明

            If dr("D5c").ToString <> "" Then
                If dr("D5c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D5c.Checked = True
                    D5c2.Checked = False
                ElseIf dr("D5c").ToString = 2 Then
                    D5c.Checked = False
                    D5c2.Checked = True
                End If
            End If

            If dr("D5c3").ToString <> "" Then D5c3.SelectedValue = dr("D5c3")  '說明的選項

            For i = 0 To Data9.Items.Count - 1 '學員書籍(講義)
                If Convert.ToString(dr("Data9")) = Data9.Items(i).Value Then Data9.Items(i).Selected = True
            Next
            DataCopy9.Text = Convert.ToString(dr("DataCopy9")) '如附件內容
            Data9Note.Text = Convert.ToString(dr("Data9Note")) '其他說明

            If dr("D6c").ToString <> "" Then
                If dr("D6c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D6c.Checked = True
                    D6c2.Checked = False
                ElseIf dr("D6c").ToString = 2 Then
                    D6c.Checked = False
                    D6c2.Checked = True
                End If
            End If

            If dr("D6c3").ToString <> "" Then D6c3.SelectedValue = dr("D6c3")  '說明的選項

            For i = 0 To Data10.Items.Count - 1 '職訓生活津貼
                If Convert.ToString(dr("Data10")) = Data10.Items(i).Value Then Data10.Items(i).Selected = True
            Next
            DataCopy10.Text = Convert.ToString(dr("DataCopy10")) '如附件內容
            Data10Note.Text = Convert.ToString(dr("Data10Note")) '其他說明

            If dr("D7c").ToString <> "" Then
                If dr("D7c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D7c.Checked = True
                    D7c2.Checked = False
                ElseIf dr("D7c").ToString = 2 Then
                    D7c.Checked = False
                    D7c2.Checked = True
                End If
            End If
            If dr("D7c3").ToString <> "" Then D7c3.SelectedValue = dr("D7c3") '說明的選項

            For i = 0 To Data11.Items.Count - 1 '學員服務手皿
                If Convert.ToString(dr("Data11")) = Data11.Items(i).Value Then Data11.Items(i).Selected = True
            Next
            DataCopy11.Text = Convert.ToString(dr("DataCopy11")) '如附件內容
            Data11Note.Text = Convert.ToString(dr("Data11Note")) '其他說明

            If dr("D8c").ToString <> "" Then
                If dr("D8c").ToString = 1 Then  '如附件的check 1表示勾如附件,2表示勾免提供
                    D8c.Checked = True
                    D8c2.Checked = False
                ElseIf dr("D8c").ToString = 2 Then
                    D8c.Checked = False
                    D8c2.Checked = True
                End If
            End If

            If dr("D8c3").ToString <> "" Then D8c3.SelectedValue = dr("D8c3") '說明的選項

            AuthCount.Text = Convert.ToString(dr("AuthCount")) '核定人數
            TurthCount.Text = Convert.ToString(dr("TurthCount")) '實到人數
            TurnoutCount.Text = Convert.ToString(dr("TurnoutCount")) '請假人數
            TruancyCount.Text = Convert.ToString(dr("TruancyCount")) '缺(曠)課人數
            RejectCount.Text = Convert.ToString(dr("RejectCount")) '退訓人數
            AheadjobCount.Text = Convert.ToString(dr("AheadjobCount")) '提前就業人數

            '課程(師資)實施狀況
            For i = 0 To Item1_1.Items.Count - 1
                If Convert.ToString(dr("Item1_1")) = Item1_1.Items(i).Value Then Item1_1.Items(i).Selected = True
            Next
            For i = 0 To Item1_2.Items.Count - 1
                If Convert.ToString(dr("Item1_2")) = Item1_2.Items(i).Value Then Item1_2.Items(i).Selected = True
            Next
            Item1_3.Text = Convert.ToString(dr("Item1_3"))
            For i = 0 To Item3_1.Items.Count - 1
                If Convert.ToString(dr("Item3_1")) = Item3_1.Items(i).Value Then Item3_1.Items(i).Selected = True
            Next
            Item3_1Tech.Text = Convert.ToString(dr("Item3_1Tech"))
            Item3_1Tutor.Text = Convert.ToString(dr("Item3_1Tutor"))
            Item1Pros.Text = Convert.ToString(dr("Item1Pros"))
            Item1Note.Text = Convert.ToString(dr("Item1Note"))

            '教材設施運用狀況
            For i = 0 To Item19.Items.Count - 1
                If Convert.ToString(dr("Item19")) = Item19.Items(i).Value Then Item19.Items(i).Selected = True
            Next
            For i = 0 To Item20.Items.Count - 1
                If Convert.ToString(dr("Item20")) = Item20.Items(i).Value Then Item20.Items(i).Selected = True
            Next
            For i = 0 To Item21.Items.Count - 1
                If Convert.ToString(dr("Item21")) = Item21.Items(i).Value Then Item21.Items(i).Selected = True
            Next

            Item2Pros.Text = Convert.ToString(dr("Item2Pros"))
            Item2Note.Text = Convert.ToString(dr("Item2Note"))

            '教務管理狀況
            For i = 0 To Item2_1.Items.Count - 1
                If Convert.ToString(dr("Item2_1")) = Item2_1.Items(i).Value Then Item2_1.Items(i).Selected = True
            Next
            For i = 0 To Item2_2.Items.Count - 1
                If Convert.ToString(dr("Item2_2")) = Item2_2.Items(i).Value Then Item2_2.Items(i).Selected = True
            Next
            For i = 0 To Item23.Items.Count - 1
                If Convert.ToString(dr("Item23")) = Item23.Items(i).Value Then Item23.Items(i).Selected = True
            Next
            For i = 0 To Item24.Items.Count - 1
                If Convert.ToString(dr("Item24")) = Item24.Items(i).Value Then Item24.Items(i).Selected = True
            Next
            For i = 0 To Item25.Items.Count - 1
                If Convert.ToString(dr("Item25")) = Item25.Items(i).Value Then Item25.Items(i).Selected = True
            Next
            For i = 0 To Item26.Items.Count - 1
                If Convert.ToString(dr("Item26")) = Item26.Items(i).Value Then Item26.Items(i).Selected = True
            Next

            Item3Pros.Text = Convert.ToString(dr("Item3Pros"))
            Item3Note.Text = Convert.ToString(dr("Item3Note"))

            '費用收核狀況
            If Convert.ToString(dr("Item28")) = "1" Then '出缺勤狀況
                Item28_11.Checked = dr("Item28")
            Else
                Item28_12.Checked = dr("Item28")
            End If
            Item28Count.Text = Convert.ToString(dr("Item28Count"))
            For i = 0 To Item28_2.Items.Count - 1
                If Convert.ToString(dr("Item28_2")) = Item28_2.Items(i).Value Then Item28_2.Items(i).Selected = True
            Next
            For i = 0 To Item6_1.Items.Count - 1
                If Convert.ToString(dr("Item6_1")) = Item6_1.Items(i).Value Then Item6_1.Items(i).Selected = True
            Next
            For i = 0 To Item29.Items.Count - 1
                If Convert.ToString(dr("Item29")) = Item29.Items(i).Value Then Item29.Items(i).Selected = True
            Next
            Item4Pros.Text = Convert.ToString(dr("Item4Pros"))
            Item4Note.Text = Convert.ToString(dr("Item4Note"))

            '職業訓練機構
            For i = 0 To Item30.Items.Count - 1
                If Convert.ToString(dr("Item30")) = Item30.Items(i).Value Then Item30.Items(i).Selected = True
            Next
            Item5Pros.Text = Convert.ToString(dr("Item5Pros"))
            Item5Note.Text = Convert.ToString(dr("Item5Note"))
            Item7Note.Text = Convert.ToString(dr("Item7Note"))
            Item31Note.Text = Convert.ToString(dr("Item31Note"))
            If Convert.ToString(dr("Item32")) <> "" Then
                If Convert.ToString(dr("Item32")) = "1" Then Item32_1.Checked = dr("Item32")
                If Convert.ToString(dr("Item32")) = "2" Then Item32_2.Checked = dr("Item32")
                If Convert.ToString(dr("Item32")) = "3" Then Item32_3.Checked = dr("Item32")
            End If
            Item32Note.Text = Convert.ToString(dr("Item32Note"))
            CurseName.Text = Convert.ToString(dr("CurseName"))
            VisitorName.Text = Convert.ToString(dr("VisitorName"))
        End If

        sql = " SELECT b.TMID, b.TrainID, b.TrainName, a.OCID, a.ClassCName, a.CyclType, a.LevelType, a.RID, c.OrgName "
        sql += " FROM (SELECT * FROM Class_ClassInfo WHERE OCID = '" & OCID & "') a, Key_TrainType b, Auth_Relship d, Org_OrgInfo c "
        sql += " WHERE a.TMID = b.TMID AND d.OrgID = c.OrgID AND d.RID = a.RID "
        dr = DbAccess.GetOneRow(sql)

        If Not dr Is Nothing Then
            center.Text = dr("OrgName")
            RIDValue.Value = dr("RID").ToString
            TMID1.Text = "[" & dr("TrainID") & "]" & dr("TrainName")
            TMIDValue1.Value = dr("TMID")
            OCID1.Text = dr("ClassCName")
            If CInt(dr("CyclType")) <> 0 Then OCID1.Text += "第" & CInt(dr("CyclType")) & "期"
            OCIDValue1.Value = dr("OCID")
            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            Button3.Disabled = True
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim SeqNo As Integer = 0

        If Request("OCID") = "" Then '表示新增狀態
            If OCIDValue1.Value = "" Then
                Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別")
                Exit Sub
            End If
            '先取出最大SeqNo
            sql = "SELECT MAX(SeqNO) as num FROM Class_Visitor WHERE OCID='" & OCIDValue1.Value & "'"
            dr = DbAccess.GetOneRow(sql)
            If Not dr Is Nothing Then
                If IsDBNull(dr("num")) Then
                    SeqNo = 1
                Else
                    SeqNo = CInt(dr("num")) + 1
                End If
            End If

            sql = "SELECT "
            sql += " SUBSTRING(ITEM10NOTE, 1, 4000) ITEM10NOTE, ITEM11, SUBSTRING(ITEM11PROS, 1, 4000) ITEM11PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM11NOTE, 1, 4000) ITEM11NOTE, ITEM12, SUBSTRING(ITEM12PROS, 1, 4000) ITEM12PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM12NOTE, 1, 4000) ITEM12NOTE, SUBSTRING(ITEM13NOTE, 1, 4000) ITEM13NOTE, " & vbCrLf
            sql += " AHEADJOBCOUNT, STUDYTICKETCOUNT, DATA9, DATACOPY9, DATA9NOTE, DATA10, DATACOPY10, DATA10NOTE, DATA11, DATACOPY11, " & vbCrLf
            sql += " DATA11NOTE, DATA12, ITEM14, ITEM14TECH, ITEM15, ITEM16, ITEM17, ITEM18, ITEM19, ITEM20, ITEM21, ITEM22, ITEM23, ITEM24, " & vbCrLf
            sql += " ITEM25, ITEM26, ITEM27, ITEM28, ITEM28COUNT, ITEM28_2, ITEM29, SUBSTRING(ITEM29PROS, 1, 4000) ITEM29PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM29NOTE, 1, 4000) ITEM29NOTE, ITEM30, SUBSTRING(ITEM31NOTE, 1, 4000) ITEM31NOTE, " & vbCrLf
            sql += " ITEM32, SUBSTRING(ITEM32NOTE, 1, 4000) ITEM32NOTE, D1C, D2C, D3C, D4C, D5C, D6C, D7C, D8C, D3C3, D4C3, D5C3, D6C3, " & vbCrLf
            sql += " D7C3, D8C3, OCID, SEQNO, APPLYDATE, AUTHCOUNT, TURTHCOUNT, TURNOUTCOUNT, TRUANCYCOUNT, REJECTCOUNT, DATA1, DATACOPY1, " & vbCrLf
            sql += " DATA1NOTE, DATA2, DATACOPY2, DATA2NOTE, DATA3, DATACOPY3, DATA3NOTE, DATA4, DATACOPY4, DATA4NOTE, DATA5, DATACOPY5, " & vbCrLf
            sql += " DATA5NOTE, DATA6, DATACOPY6, DATA6NOTE, DATA7, DATACOPY7, DATA7NOTE, ITEM1_1, ITEM1_2, ITEM1_3, " & vbCrLf
            sql += " SUBSTRING(ITEM1PROS, 1, 4000) ITEM1PROS, SUBSTRING(ITEM1NOTE, 1, 4000) ITEM1NOTE, ITEM2_1, " & vbCrLf
            sql += " ITEM2_2, SUBSTRING(ITEM2PROS, 1, 4000) ITEM2PROS, SUBSTRING(ITEM2NOTE, 1, 4000) ITEM2NOTE, " & vbCrLf
            sql += " ITEM3_1, ITEM3_1TECH, ITEM3_1TUTOR, ITEM3_2, SUBSTRING(ITEM3PROS, 1, 4000) ITEM3PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM3NOTE, 1, 4000) ITEM3NOTE, ITEM4_1, SUBSTRING(ITEM4PROS, 1, 4000) ITEM4PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM4NOTE, 1, 4000) ITEM4NOTE, ITEM5_1, SUBSTRING(ITEM5PROS, 1, 4000) ITEM5PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM5NOTE, 1, 4000) ITEM5NOTE, ITEM6_1, ITEM6COUNT1, ITEM6COUNT2, ITEM6COUNT3, ITEM6NAMES, " & vbCrLf
            sql += " SUBSTRING(ITEM6NOTE, 1, 4000) ITEM6NOTE, SUBSTRING(ITEM7NOTE, 1, 4000) ITEM7NOTE, CURSENAME, " & vbCrLf
            sql += " VISITORNAME, MODIFYACCT, MODIFYDATE, RID, ISCLEAR, DATA8, VISITHOUR, TPERIOD, COMCOUNT, BOOKCOUNT, NETCOUNT, PSYCOUNT, " & vbCrLf
            sql += " ITEM8, SUBSTRING(ITEM8PROS, 1, 4000) ITEM8PROS, SUBSTRING(ITEM8NOTE, 1, 4000) ITEM8NOTE, " & vbCrLf
            sql += " ITEM9, ITEM9_TECH, SUBSTRING(ITEM9PROS, 1, 4000) ITEM9PROS, SUBSTRING(ITEM9NOTE, 1, 4000) ITEM9NOTE, " & vbCrLf
            sql += " ITEM10, SUBSTRING(ITEM10PROS, 1, 4000) ITEM10PROS " & vbCrLf
            sql += " FROM Class_Visitor WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, da, objConn)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = OCIDValue1.Value
            dr("SeqNo") = SeqNo
        Else
            sql = "SELECT "
            sql += " SUBSTRING(ITEM10NOTE, 1, 4000) ITEM10NOTE, ITEM11,SUBSTRING(ITEM11PROS, 1, 4000) ITEM11PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM11NOTE, 1, 4000) ITEM11NOTE, ITEM12,SUBSTRING(ITEM12PROS, 1, 4000) ITEM12PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM12NOTE, 1, 4000) ITEM12NOTE, SUBSTRING(ITEM13NOTE, 1, 4000) ITEM13NOTE, " & vbCrLf
            sql += " AHEADJOBCOUNT, STUDYTICKETCOUNT, DATA9, DATACOPY9, DATA9NOTE, DATA10, DATACOPY10, DATA10NOTE, DATA11, DATACOPY11, " & vbCrLf
            sql += " DATA11NOTE, DATA12, ITEM14, ITEM14TECH, ITEM15, ITEM16, ITEM17, ITEM18, ITEM19, ITEM20, ITEM21, ITEM22, ITEM23, ITEM24, " & vbCrLf
            sql += " ITEM25, ITEM26, ITEM27, ITEM28, ITEM28COUNT, ITEM28_2, ITEM29, SUBSTRING(ITEM29PROS, 1, 4000) ITEM29PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM29NOTE, 1, 4000) ITEM29NOTE, ITEM30, SUBSTRING(ITEM31NOTE, 1, 4000) ITEM31NOTE, " & vbCrLf
            sql += " ITEM32, SUBSTRING(ITEM32NOTE, 1, 4000) ITEM32NOTE, D1C, D2C, D3C, D4C, D5C, D6C, D7C, D8C, D3C3, D4C3, D5C3, D6C3, " & vbCrLf
            sql += " D7C3, D8C3, OCID, SEQNO, APPLYDATE, AUTHCOUNT, TURTHCOUNT, TURNOUTCOUNT, TRUANCYCOUNT, REJECTCOUNT, DATA1, DATACOPY1, " & vbCrLf
            sql += " DATA1NOTE, DATA2, DATACOPY2, DATA2NOTE, DATA3, DATACOPY3, DATA3NOTE, DATA4, DATACOPY4, DATA4NOTE, DATA5, DATACOPY5, " & vbCrLf
            sql += " DATA5NOTE, DATA6, DATACOPY6, DATA6NOTE, DATA7, DATACOPY7, DATA7NOTE, ITEM1_1, ITEM1_2, ITEM1_3, " & vbCrLf
            sql += " SUBSTRING(ITEM1PROS, 1, 4000) ITEM1PROS, SUBSTRING(ITEM1NOTE, 1, 4000) ITEM1NOTE, ITEM2_1, " & vbCrLf
            sql += " ITEM2_2, SUBSTRING(ITEM2PROS, 1, 4000) ITEM2PROS, SUBSTRING(ITEM2NOTE, 1, 4000) ITEM2NOTE, " & vbCrLf
            sql += " ITEM3_1, ITEM3_1TECH, ITEM3_1TUTOR, ITEM3_2, SUBSTRING(ITEM3PROS, 1, 4000) ITEM3PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM3NOTE, 1, 4000) ITEM3NOTE, ITEM4_1, SUBSTRING(ITEM4PROS, 1, 4000) ITEM4PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM4NOTE, 1, 4000) ITEM4NOTE, ITEM5_1, SUBSTRING(ITEM5PROS, 1, 4000) ITEM5PROS, " & vbCrLf
            sql += " SUBSTRING(ITEM5NOTE, 1, 4000) ITEM5NOTE, ITEM6_1, ITEM6COUNT1, ITEM6COUNT2, ITEM6COUNT3, ITEM6NAMES, " & vbCrLf
            sql += " SUBSTRING(ITEM6NOTE, 1, 4000) ITEM6NOTE, SUBSTRING(ITEM7NOTE, 1, 4000) ITEM7NOTE,CURSENAME, " & vbCrLf
            sql += " VISITORNAME, MODIFYACCT, MODIFYDATE, RID, ISCLEAR, DATA8, VISITHOUR, TPERIOD, COMCOUNT, BOOKCOUNT, NETCOUNT, PSYCOUNT, " & vbCrLf
            sql += " ITEM8,SUBSTRING(ITEM8PROS, 1, 4000) ITEM8PROS,SUBSTRING(ITEM8NOTE, 1, 4000) ITEM8NOTE, " & vbCrLf
            sql += " ITEM9, ITEM9_TECH, SUBSTRING(ITEM9PROS, 1, 4000) ITEM9PROS, SUBSTRING(ITEM9NOTE, 1, 4000) ITEM9NOTE, " & vbCrLf
            sql += " ITEM10, SUBSTRING(ITEM10PROS, 1, 4000) ITEM10PROS " & vbCrLf
            sql += " FROM Class_Visitor WHERE OCID = '" & Request("OCID") & "' AND SeqNo = '" & Request("SeqNo") & "' "
            dt = DbAccess.GetDataTable(sql, da, objConn)
            dr = dt.Rows(0)
            OCIDValue1.Value = Request("OCID")
            SeqNo = Request("SeqNo")
        End If

        '**by Milor 20080623--訪查日期未填時，給予告警不存檔----start
        If ApplyDate.Text <> "" Then
            dr("ApplyDate") = Convert.ToDateTime(ApplyDate.Text) '訪查日期
        Else
            Common.MessageBox(Me, "請填選查訪日期！")
            Exit Sub
        End If
        '**by Milor 20080623----end

        dr("Data1") = Data1.SelectedValue '教學(訓練)日誌
        If DataCopy1.Text <> "" Then dr("DataCopy1") = DataCopy1.Text 'Data1 如附件內容
        If Data1Note.Text <> "" Then dr("Data1Note") = Data1Note.Text 'Data1 說明
        If D1c.Checked = True Then dr("D1c") = 1 '如附件 check
        dr("Data3") = Data3.SelectedValue '學員簽到(退)表
        If DataCopy3.Text <> "" Then dr("DataCopy3") = DataCopy3.Text '如附件內容
        If Data3Note.Text <> "" Then dr("Data3Note") = Data3Note.Text '說明
        If D2c.Checked = True Then dr("D2c") = 1 '如附件 check
        dr("Data5") = Data5.SelectedValue '請假單
        If DataCopy5.Text <> "" Then dr("DataCopy5") = DataCopy5.Text '如附件內容
        If Data5Note.Text <> "" Then dr("Data5Note") = Data5Note.Text '說明
        If D3c.Checked = True Then '如附件 check
            dr("D3c") = 1
        ElseIf D3c2.Checked = True Then
            dr("D3c") = 2
        End If
        If D3c3.SelectedIndex <> -1 Then dr("D3c3") = D3c3.SelectedValue '說明的check
        dr("Data6") = Data6.SelectedValue '退訓/提前就業申請表
        If DataCopy6.Text <> "" Then dr("DataCopy6") = DataCopy6.Text
        If Data6Note.Text <> "" Then dr("Data6Note") = Data6Note.Text
        If D4c.Checked = True Then '如附件 check
            dr("D4c") = 1
        ElseIf D4c2.Checked = True Then
            dr("D4c") = 2
        End If
        If D4c3.SelectedIndex <> -1 Then dr("D4c3") = D4c3.SelectedValue '說明的check
        dr("Data7") = Data7.SelectedValue '勞保加/退明細表
        If DataCopy7.Text <> "" Then dr("DataCopy7") = DataCopy7.Text
        If Data7Note.Text <> "" Then dr("Data7Note") = Data7Note.Text
        If D5c.Checked = True Then '如附件 check
            dr("D5c") = 1
        ElseIf D5c2.Checked = True Then
            dr("D5c") = 2
        End If
        If D5c3.SelectedIndex <> -1 Then dr("D5c3") = D5c3.SelectedValue  '說明的check
        dr("Data9") = Data9.SelectedValue '學員書籍(講義)、材料領用表
        If DataCopy9.Text <> "" Then dr("DataCopy9") = DataCopy9.Text
        If Data9Note.Text <> "" Then dr("Data9Note") = Data9Note.Text
        If D6c.Checked = True Then '如附件 check
            dr("D6c") = 1
        ElseIf D6c2.Checked = True Then
            dr("D6c") = 2
        End If
        If D6c3.SelectedIndex <> -1 Then dr("D6c3") = D6c3.SelectedValue '說明的check
        dr("Data10") = Data10.SelectedValue '職訓生活津補助印領清冊
        If DataCopy10.Text <> "" Then dr("DataCopy10") = DataCopy10.Text
        If Data10Note.Text <> "" Then dr("Data10Note") = Data10Note.Text
        If D7c.Checked = True Then '如附件 check
            dr("D7c") = 1
        ElseIf D7c2.Checked = True Then
            dr("D7c") = 2
        End If
        If D7c3.SelectedIndex <> -1 Then dr("D7c3") = D7c3.SelectedValue '說明的check
        dr("Data11") = Data11.SelectedValue '學員服務手冊或權利義務公告相關文件
        If DataCopy11.Text <> "" Then dr("DataCopy11") = DataCopy11.Text
        If Data11Note.Text <> "" Then dr("Data11Note") = Data11Note.Text
        If D8c.Checked = True Then '如附件 check
            dr("D8c") = 1
        ElseIf D8c2.Checked = True Then
            dr("D8c") = 2
        End If
        If D8c3.SelectedIndex <> -1 Then dr("D8c3") = D8c3.SelectedValue '說明的check
        If AuthCount.Text <> "" Then dr("AuthCount") = AuthCount.Text '核定人數
        If TurthCount.Text <> "" Then dr("TurthCount") = TurthCount.Text '實到人數
        If TurnoutCount.Text <> "" Then dr("TurnoutCount") = TurnoutCount.Text '請假人數
        If TruancyCount.Text <> "" Then dr("TruancyCount") = TruancyCount.Text '缺(曠)課人數
        If RejectCount.Text <> "" Then dr("RejectCount") = RejectCount.Text '退訓人數
        If AheadjobCount.Text <> "" Then dr("AheadjobCount") = AheadjobCount.Text '提前就業人數
        dr("Item1_1") = Item1_1.SelectedValue '有無週(月)課程表
        dr("Item1_2") = Item1_2.SelectedValue '是否依課程表授課
        dr("Item1_3") = Item1_3.Text '課目
        dr("Item3_1") = Item3_1.SelectedValue '教師與助教是否與計晝相符
        dr("Item3_1Tech") = Item3_1Tech.Text '教師
        If Item3_1Tutor.Text <> "" Then dr("Item3_1Tutor") = Item3_1Tutor.Text '助教
        If Item1Pros.Text <> "" Then dr("Item1Pros") = Item1Pros.Text
        If Item1Note.Text <> "" Then dr("Item1Note") = Item1Note.Text
        dr("Item19") = Item19.SelectedValue '有無書籍(講義)領用表
        dr("Item20") = Item20.SelectedValue '有無材料領用表
        dr("Item21") = Item21.SelectedValue '訓練設施設備是否依契約提供學員使用
        If Item2Pros.Text <> "" Then dr("Item2Pros") = Item2Pros.Text
        If Item2Note.Text <> "" Then dr("Item2Note") = Item2Note.Text
        dr("Item2_1") = Item2_1.SelectedValue '教學(訓練)日誌是否確實填寫
        dr("Item2_2") = Item2_2.SelectedValue '有否按時呈主管核閱
        dr("Item23") = Item23.SelectedValue '學員生活、就業輔導與管理機制是否依契約規範辦理
        dr("Item24") = Item24.SelectedValue '是否依契約規範提供學員問題反應申訴管道
        dr("Item25") = Item25.SelectedValue '是否為參訓學員辦理勞工保險加退保
        dr("Item26") = Item26.SelectedValue '是否依契約規範公告學員權益義務或編製參訓學員服務手冊
        If Item3Pros.Text <> "" Then dr("Item3Pros") = Item3Pros.Text
        If Item3Note.Text <> "" Then dr("Item3Note") = Item3Note.Text
        If Item28_11.Checked = True Then
            dr("Item28") = Item28_11.Value '有無自費參訓學員
        Else
            dr("Item28") = Item28_12.Value '有無自費參訓學員
        End If
        dr("Item28_2") = Item28_2.SelectedValue '訓練費用是否繳交主辦單位
        If Item28Count.Text <> "" Then dr("Item28Count") = Item28Count.Text '有自費參訓學員人數
        dr("Item6_1") = Item6_1.SelectedValue '職業訓練生活津貼是否依規定申請並核發
        dr("Item29") = Item29.SelectedValue '培訓單位無巧立名目強制收取費用
        If Item4Pros.Text <> "" Then dr("Item4Pros") = Item4Pros.Text
        If Item4Note.Text <> "" Then dr("Item4Note") = Item4Note.Text
        dr("Item30") = Item30.SelectedValue '職業訓練機構是否依規定懸掛設立許可證書
        If Item5Pros.Text <> "" Then dr("Item5Pros") = Item5Pros.Text
        If Item5Note.Text <> "" Then dr("Item5Note") = Item5Note.Text
        If Item7Note.Text <> "" Then dr("Item7Note") = Item7Note.Text '參訓學員反映意見及問題
        If Item31Note.Text <> "" Then dr("Item31Note") = Item31Note.Text '綜合建議
        If Item32_1.Checked = True Then
            dr("Item32") = Item32_1.Value '缺失處理
        ElseIf Item32_2.Checked = True Then
            dr("Item32") = Item32_2.Value '缺失處理
        ElseIf Item32_3.Checked = True Then
            dr("Item32") = Item32_3.Value '缺失處理
        End If
        If Item32Note.Text <> "" Then dr("Item32Note") = Item32Note.Text
        dr("CurseName") = CurseName.Text '培訓單位人員姓名
        dr("VisitorName") = VisitorName.Text '訪視人員姓名
        dr("RID") = sm.UserInfo.RID
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Try
            DbAccess.UpdateDataTable(dt, da)
            Session("SearchStr") = Me.ViewState("SearchStr")
            Session("_SearchStr") = Me.ViewState("_SearchStr")
            Common.RespWrite(Me, "<script> alert('儲存成功');")
            If Request("DOCID") <> "" Then
                Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID") & "';</script>")
            Else
                Common.RespWrite(Me, "location.href='CP_01_001.aspx?ID=" & Request("ID") & "';</script>")
            End If
        Catch ex As Exception
            'Common.RespWrite(Me, ex)
            Throw ex
        End Try
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        Session("SearchStr") = Me.ViewState("SearchStr")
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        If Request("DOCID") <> "" Then
            TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID") & "&DOCID=" & Request("DOCID"))
        Else
            TIMS.Utl_Redirect1(Me, "CP_01_001.aspx?ID=" & Request("ID"))
        End If
    End Sub
End Class