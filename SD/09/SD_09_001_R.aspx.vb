Partial Class SD_09_001_R
    Inherits AuthBasePage

    'Const cst_printFN2 As String = "Student_Report2"
    Const cst_printFN3 As String = "Student_Report3"

    'Student_Report2
    Dim PlanKind As String = ""
    Dim vsOkPrintTPlanID As String = "" '=ViewState("OkPrintTPlanID")
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        ''分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        ''分頁設定 End

        vsOkPrintTPlanID = Get_OkPrintTPlanID(objconn)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '職類/班別 顯示
            'Me.ClassPanel.Visible = True
            msg.Text = ""
            Table4.Visible = False
            PageControler1.Visible = False
            Button1.Visible = False
            'msg.Visible = False
            'Request("AcctPlan")=1              表示可以跨計畫選擇
            'request("RWClass")=1               被班級計畫賦予限制

            '依sm.UserInfo.PlanID取得PlanKind
            PlanKind = TIMS.Get_PlanKind(Me, objconn)
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button12_Click(sender, e)
            End If

            'Button1.Attributes("onclick") = "CheckPrint('" & ReportQuery.GetSmartQueryPath & "');return false;"
            'Button1.Attributes("onclick") = "CheckPrint('" & ReportQuery.GetSmartQueryPath & "');"
            Button1.Attributes("onclick") = "CheckPrint();"
        End If

        'Button1.Attributes("onclick") = "if(ReportPrint()){"
        'Button1.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "MultiBlock", "Student_Report", "RID='+document.getElementById('RIDValue').value+'&OCID='+document.getElementById('OCIDValue1').value+'&TMID='+document.getElementById('TMIDValue1').value+'")
        'Button1.Attributes("onclick") += "}return false;"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        ' TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, , True)
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim dr_Class As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"

            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim labelStar As Label = e.Item.FindControl("LabelStar")
                Dim labOCID As Label = e.Item.FindControl("labOCID")

                labOCID.Style("display") = "none"
                labOCID.Visible = False
                labOCID.Text = drv("OCID")

                Checkbox1.Value = drv("OCID")
                Checkbox1.Attributes("onclick") = "InsertValue(this.checked,this.value)"
                'If PrintValue.Value.IndexOf(Checkbox1.Value) <> -1 Then
                If PrintValue.Value.IndexOf(drv("OCID")) <> -1 Then
                    Checkbox1.Checked = True '選擇
                End If
                'by Milor 20080804--此限制要排除學習卷----start
                'by Milor 20080715----start
                '加入判斷學員資料中是否有必要欄位未填，有則顯示星號並鎖定無法核取。
                ''15:學習卷
                ''26:外籍與大陸配偶職業訓練(暫時提供開放列印)
                ''03:彌補轄區不足委外訓練
                ''05:移地訓練
                ''12:職訓券
                ''35:中長期失業者職前訓練

                labelStar.Visible = False '資料齊全
                Checkbox1.Disabled = False '開放
                Dim rMsg1 As String = ""
                If TIMS.Chk_ClassStdApprAll(Me, drv("OCID"), objconn, rMsg1) Then
                    labelStar.Visible = True '資料不齊全
                    If rMsg1 <> "" Then
                        TIMS.Tooltip(labelStar, rMsg1, True) '學員資料中有必要欄位未填
                    Else
                        TIMS.Tooltip(labelStar, "學員資料中有必要欄位未填!!", True) '學員資料中有必要欄位未填
                    End If
                    TIMS.Tooltip(Checkbox1, "學員資料中有必要欄位未填!!", True) '學員資料中有必要欄位未填
                End If

                If vsOkPrintTPlanID <> "" AndAlso vsOkPrintTPlanID.ToString.IndexOf(sm.UserInfo.TPlanID.ToString) > -1 Then
                    Checkbox1.Disabled = False '開放
                    If labelStar.Visible Then  '資料不齊全
                        TIMS.Tooltip(labelStar, "學員資料中有必要欄位未填!!(暫時提供列印)", True) '學員資料中有必要欄位未填
                        TIMS.Tooltip(Checkbox1, "學員資料中有必要欄位未填!!(暫時提供列印)", True) '學員資料中有必要欄位未填
                    End If
                Else
                    If labelStar.Visible = True Then  '資料不齊全
                        Checkbox1.Disabled = True '不開放
                    End If
                End If
                'by Milor 20080804

                'e.Item.Cells(1).Text = "[" & drv("TrainID").ToString & "]" & drv("TrainName").ToString
                'If Int(drv("CyclType")) <> 0 Then
                '    e.Item.Cells(3).Text += "第" & TIMS.GetChtNum(Int(drv("CyclType"))) & "期"
                'End If
                'If e.Item.Cells(5).Text = "Y" Or e.Item.Cells(5).Text = "y" Then
                '    e.Item.Cells(5).Text = "可挑選志願"
                'Else
                '    e.Item.Cells(5).Text = "不可挑選志願"
                'End If

                Dim v_IsApplicTxt As String = "不可挑選志願"
                If Convert.ToString(drv("IsApplic")) = "Y" Then v_IsApplicTxt = "可挑選志願"
                e.Item.Cells(5).Text = v_IsApplicTxt

        End Select

    End Sub

    Function Get_OkPrintTPlanID(ByVal tConn As SqlConnection) As String
        Dim rst As String = ""
        Dim dt As DataTable
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT TPLANID FROM SYS_VAR"
        sql &= " WHERE 1=1"
        sql &= " AND UPPER(SPAGE)='SD_09_001_R'"
        sql &= " AND LOWER(ItemName)='print'"
        sql &= " and ItemValue='Y'"
        dt = DbAccess.GetDataTable(sql, tConn)
        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                If rst <> "" Then rst += ","
                rst += Convert.ToString(dt.Rows(i)("TPlanID"))
            Next
        End If
        Return rst
    End Function

    '查詢
    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim rqAcctPlan As String = "" 'Request("AcctPlan")
        Dim rqRWClass As String = "" 'Request("RWClass")
        rqAcctPlan = Request("AcctPlan")
        rqRWClass = Request("RWClass")
        rqAcctPlan = TIMS.ClearSQM(rqAcctPlan)
        rqRWClass = TIMS.ClearSQM(rqRWClass)

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        CyclType.Value = TIMS.ClearSQM(CyclType.Value)

        Dim sqlStr As String = ""
        sqlStr = "" & vbCrLf
        sqlStr &= " select cc.Years" & vbCrLf
        sqlStr &= " ,cc.PlanID,cc.OCID,cc.ClassCName" & vbCrLf
        sqlStr &= " ,cc.STDate,cc.FTDate" & vbCrLf
        sqlStr &= " ,cc.TMID" & vbCrLf
        sqlStr &= " ,cc.IsApplic" & vbCrLf
        sqlStr &= " ,cc.CLSID" & vbCrLf
        sqlStr &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sqlStr &= " ,cc.CyclType" & vbCrLf
        sqlStr &= " ,cc.ComIDNO,cc.SeqNO" & vbCrLf
        sqlStr &= " ,cc.LevelType,cc.CJOB_UNKEY " & vbCrLf
        sqlStr &= " ,ip.TPlanID" & vbCrLf
        sqlStr &= " ,cc.Years+'0'+ic.ClassID+cc.CyclType ClassID" & vbCrLf
        sqlStr &= " ,t.TrainID,t.TrainName" & vbCrLf
        sqlStr &= " ,'['+t.TrainID+']'+t.TrainName TrainName2" & vbCrLf
        sqlStr &= " FROM Class_ClassInfo cc" & vbCrLf
        sqlStr &= " join ID_Class ic on ic.CLSID =cc.CLSID" & vbCrLf
        sqlStr &= " join ID_Plan ip on ip.planid =cc.planid " & vbCrLf
        sqlStr &= " join Auth_Relship rr on rr.RID =cc.RID" & vbCrLf
        sqlStr &= " join Key_TrainType t on t.TMID=cc.TMID" & vbCrLf
        sqlStr &= " where 1=1 " & vbCrLf
        sqlStr &= " and cc.IsSuccess='Y' " & vbCrLf
        sqlStr &= " and cc.NotOpen='N'" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            '署(局)
            sqlStr &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'  " & vbCrLf
            sqlStr &= " and ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            '其他轄區
            sqlStr &= " and ip.TPlanID='" & sm.UserInfo.TPlanID & "'  " & vbCrLf
            sqlStr &= " and ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
            'Request("AcctPlan")=1              表示可以跨計畫選擇
            'request("RWClass")=1               被班級計畫賦予限制
            If rqAcctPlan = "1" Then '可以跨計畫選擇
                sqlStr &= " and rr.OrgID in (SELECT OrgID FROM AUTH_RELSHIP WHERE RID ='" & RIDValue.Value & "')" & vbCrLf
            Else
                If RIDValue.Value <> "" Then
                    sqlStr &= " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
                Else
                    sqlStr &= " and cc.RID='" & sm.UserInfo.RID & "'" & vbCrLf
                End If
                sqlStr &= " and ip.PlanID='" & sm.UserInfo.PlanID & "'  " & vbCrLf
            End If
            If PlanKind = "1" AndAlso rqRWClass = 1 Then '班級計畫賦予限制
                sqlStr &= " JOIN Auth_AccRWClass c on cc.OCID=c.OCID and c.Account='" & sm.UserInfo.UserID & "'" & vbCrLf
            End If
        End If

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '開訓二週前
                sqlStr &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),cc.STDate) >0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                sqlStr &= " AND DATEDIFF(DAY,cc.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,cc.FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                sqlStr &= " AND DATEDIFF(DAY,cc.FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                sqlStr &= " AND DATEDIFF(DAY,GETDATE(),cc.STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                sqlStr &= " AND 1<>1 "
        End Select

        If OCIDValue1.Value.ToString <> "" Then
            sqlStr &= " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        If Me.txtCJOB_NAME.Text <> "" Then
            sqlStr &= " and cc.CJOB_UNKEY='" & Me.cjobValue.Value & "'" & vbCrLf '通俗職類
        End If
        If CyclType.Value.ToString <> "" Then
            If Len(CyclType.Value.ToString) = 1 Then
                sqlStr &= " and cc.CyclType= '0" & CyclType.Value & "'" & vbCrLf
            Else
                sqlStr &= " and cc.CyclType='" & CyclType.Value & "'" & vbCrLf
            End If
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlStr, objconn)

        msg.Text = "查無資料!!"
        Table4.Visible = False
        PageControler1.Visible = False
        Button1.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True
            PageControler1.Visible = True
            Button1.Visible = True

            'PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
            'DataGrid1.DataSource = dt
            'DataGrid1.DataBind()
        End If

    End Sub

    '列印鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'If Me.PrintValue.Value = "" Then
        'End If
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim OCIDs As String = ""
        For Each Item As DataGridItem In DataGrid1.Items
            Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
            Dim labelStar As Label = Item.FindControl("LabelStar")
            Dim labOCID As Label = Item.FindControl("labOCID")
            If Checkbox1.Checked Then
                If OCIDs <> "" Then OCIDs += ","
                OCIDs += Convert.ToString("\'" & labOCID.Text & "\'")
            End If
        Next

        If OCIDs = "" Then
            Common.MessageBox(Me, "至少要勾選一班級")
            Exit Sub
        End If

        Dim MyValue As String = "ijk=ijk"
        MyValue &= "&OCID=" & OCIDs 'Me.PrintValue.Value
        MyValue &= "&RID=" & Me.RIDValue.Value
        MyValue &= "&UserID=" & sm.UserInfo.UserID
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN3, MyValue)
    End Sub

    '**by Milor 20080715----end
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        PageControler1.Visible = False
        Button1.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

End Class
