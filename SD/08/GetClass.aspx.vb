Partial Class GetClass
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            TPeriod = TIMS.GET_HOURRAN(TPeriod, objconn, sm)
            DataGridTable.Visible = False
        End If

        msg.Text = ""
        PageControler1.PageDataGrid = DataGrid1
        Button2.Attributes("onclick") = "if(document.form1.OCID.value==''){alert('請選擇班級');return false;}"

        If Not IsPostBack Then
            Dim dt As DataTable = TIMS.GetCookieTable(Me, objconn)

            If dt.Select("ItemName='SD_TB_career_id'").Length <> 0 Then
                TB_career_id.Text = dt.Select("ItemName='SD_TB_career_id'")(0)("ItemValue")
                trainValue.Value = dt.Select("ItemName='SD_trainValue'")(0)("ItemValue")
                Common.SetListItem(TPeriod, dt.Select("ItemName='SD_HourRan'")(0)("ItemValue"))
                CyclType.Text = dt.Select("ItemName='SD_CyclType'")(0)("ItemValue")
                ClassID.Text = dt.Select("ItemName='SD_ClassID'")(0)("ItemValue")
                Common.SetListItem(ClassRound, dt.Select("ItemName='SD_ClassRound'")(0)("ItemValue"))

                Button1_Click(sender, e)
            End If
        End If
    End Sub

    Sub GetSearch()
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        TIMS.InsertCookieTable(Me, dt, da, "SD_TB_career_id", TB_career_id.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_trainValue", trainValue.Value, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_HourRan", TPeriod.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_CyclType", CyclType.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassID", ClassID.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD_ClassRound", ClassRound.SelectedValue, True, objconn)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        'Dim V_HourRan As String = TIMS.ClearSQM(HourRan.SelectedValue)
        Dim rq_RID As String = TIMS.ClearSQM(Request("RID"))
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)
        ClassID.Text = TIMS.ClearSQM(ClassID.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        Dim V_TPeriod As String = TIMS.ClearSQM(TPeriod.SelectedValue)

        OCID.Value = ""
        Dim SqlStr As String = ""
        SqlStr = ""
        SqlStr += " SELECT cc.OCID,c.TrainName,b.ClassID,cc.CyclType"
        SqlStr += " ,cc.LevelType,cc.ClassCName"
        SqlStr += " ,format(cc.STDate,'yyyy/MM/dd') STDate"
        SqlStr += " ,format(cc.FTDate,'yyyy/MM/dd') FTDate"
        SqlStr += " ,cc.IsApplic " & vbCrLf
        SqlStr &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2 " & vbCrLf
        SqlStr += " FROM Class_ClassInfo cc " & vbCrLf
        SqlStr += " LEFT JOIN ID_Class b ON cc.CLSID=b.CLSID " & vbCrLf
        SqlStr += " LEFT JOIN Key_TrainType c ON cc.TMID=c.TMID " & vbCrLf
        SqlStr += " WHERE 1=1" & vbCrLf
        SqlStr += " and cc.NotOpen='N' " & vbCrLf
        SqlStr += " and cc.IsSuccess='Y' " & vbCrLf
        If sm.UserInfo.RID = "A" Then
            SqlStr += " and cc.PlanID in (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            SqlStr += "    and Years='" & sm.UserInfo.Years & "' )" & vbCrLf

            SqlStr += " and cc.RID='" & rq_RID & "'" & vbCrLf
        Else
            SqlStr += " and cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            SqlStr += " and cc.RID='" & rq_RID & "'" & vbCrLf
        End If
        If Trim(TB_career_id.Text) <> "" Then
            SqlStr += " and cc.TMID='" & trainValue.Value & "'" & vbCrLf
        End If
        If TPeriod.SelectedIndex <> 0 AndAlso V_TPeriod <> "" Then
            SqlStr += " and cc.TPeriod='" & V_TPeriod & "'" & vbCrLf
        End If

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '開訓二週前
                SqlStr &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),cc.STDate) >0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                SqlStr &= " AND DATEDIFF(DAY,cc.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,cc.FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                SqlStr &= " AND DATEDIFF(DAY,cc.FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                SqlStr &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                SqlStr &= " AND 1<>1 "
        End Select

        If Trim(CyclType.Text) <> "" Then
            If IsNumeric(CyclType.Text) Then
                If Int(CyclType.Text) < 10 Then
                    SqlStr += " and cc.CyclType='0" & Int(CyclType.Text) & "'" & vbCrLf
                Else
                    SqlStr += " and cc.CyclType='" & Int(CyclType.Text) & "'" & vbCrLf
                End If
            End If
        End If
        If ClassID.Text <> "" Then
            SqlStr += " and cc.CLSID IN (SELECT CLSID FROM ID_Class WHERE ClassID Like'%" & ClassID.Text & "%')" & vbCrLf
        End If
        If ClassName.Text <> "" Then
            SqlStr += " and cc.ClassCName like '%" & ClassName.Text & "%'"
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(SqlStr, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            Call GetSearch() 'InsertCookieTable

            DataGridTable.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim sql As String
        'Dim dt As DataTable
        'Dim da As SqlDataAdapter
        'Dim dr As DataRow
        Dim ScriptStr As String = ""
        Dim ClassName As String = ""

        OCID.Value = TIMS.ClearSQM(OCID.Value)
        Dim rq_RID As String = TIMS.ClearSQM(Request("RID"))
        Dim rqOCIDField As String = TIMS.ClearSQM(Request("OCIDField"))
        Dim rqSubmitBtn As String = TIMS.ClearSQM(Request("SubmitBtn"))

        If rqOCIDField <> "" Then ScriptStr += "opener.document.getElementById('" & rqOCIDField & "').value='" & OCID.Value & "';"
        If rqSubmitBtn <> "" Then ScriptStr += "opener.document.getElementById('" & rqSubmitBtn & "').click();"

        '若 OCID.Value 有多選，把他變成單選
        OCID.Value = TIMS.GetOneValue(OCID.Value)
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        dt = TIMS.GetCookieTable(Me, da, objconn)
        For i As Integer = 1 To 5
            If dt.Select("ItemName='SubsidyRID" & i & "'").Length = 0 Then
                Dim InsertFlag As Boolean = True
                For j As Integer = 1 To 5
                    If dt.Select("ItemName='SubsidyRID" & j & "' and ItemValue='" & rq_RID & "'").Length <> 0 Then
                        InsertFlag = False

                        Dim Olddr As DataRow
                        Olddr = dt.Select("ItemName='SubsidyRID" & j & "' and ItemValue='" & rq_RID & "'")(0)
                        TIMS.InsertCookieTable(Me, dt, da, "SubsidyClass" & j, OCID.Value, True, objconn)
                        Exit For
                    End If
                Next

                If InsertFlag Then
                    TIMS.InsertCookieTable(Me, dt, da, "SubsidyRID" & i, rq_RID, False, objconn)
                    TIMS.InsertCookieTable(Me, dt, da, "SubsidyClass" & i, OCID.Value, True, objconn)
                End If
                Exit For
            Else
                If i = 5 Then
                    For j As Integer = 1 To 4
                        dt.Select("ItemName='SubsidyRID" & j & "'")(0)("ItemValue") = dt.Select("ItemName='SubsidyRID" & j + 1 & "'")(0)("ItemValue")
                        dt.Select("ItemName='SubsidyClass" & j & "'")(0)("ItemValue") = dt.Select("ItemName='SubsidyClass" & j + 1 & "'")(0)("ItemValue")
                    Next

                    Dim InsertFlag As Boolean = True
                    For j As Integer = 1 To 5
                        If dt.Select("ItemName='SubsidyRID" & j & "' and ItemValue='" & rq_RID & "'").Length <> 0 Then
                            InsertFlag = False

                            Dim Olddr As DataRow
                            Olddr = dt.Select("ItemName='SubsidyRID" & j & "' and ItemValue='" & rq_RID & "'")(0)
                            Olddr("ItemValue") = OCID.Value
                        End If
                    Next

                    If InsertFlag Then
                        TIMS.InsertCookieTable(Me, dt, da, "SubsidyRID" & i, rq_RID, False, objconn)
                        TIMS.InsertCookieTable(Me, dt, da, "SubsidyClass" & i, OCID.Value, True, objconn)
                    End If
                End If
            End If
        Next

        Page.RegisterStartupScript("close", "<script>" & ScriptStr & "window.close();</script>")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                Checkbox2.Attributes("onclick") = "select_all(this.checked)"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")

                Checkbox1.Value = drv("OCID")
                Checkbox1.Attributes("onclick") = "GetOCID(this.checked,this.value)"
                Dim OCIDArray As Array = Split(OCID.Value, ",")
                For i As Integer = 0 To OCIDArray.Length - 1
                    If drv("OCID").ToString = OCIDArray(i) Then Checkbox1.Checked = True
                Next

                'If IsNumeric(drv("CyclType")) Then
                '    If Int(drv("CyclType")) <> 0 Then
                '        e.Item.Cells(3).Text += "第" & Int(drv("CyclType")) & "期"
                '    End If
                'End If

                e.Item.Cells(5).Text = If(Convert.ToString(drv("IsApplic")) = "Y", "可挑選志願", "不可挑選志願")
        End Select
    End Sub
End Class
