Partial Class SD_09_015_R
    Inherits AuthBasePage

    Dim sPlanKind As String = ""
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
        '檢查Session是否存在 End
        PageControler1 = Me.FindControl("PageControler1")
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        sPlanKind = TIMS.Get_PlanKind(Me, objconn)

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            'Me.ClassPanel.Visible = True
            PageControler1.Visible = False
            Button1.Visible = False
            Button3.Visible = False
            msg.Visible = False

            'Request("AcctPlan")=1              表示可以跨計畫選擇
            'request("RWClass")=1               被班級計畫賦予限制
            'sPlanKind = DbAccess.ExecuteScalar("SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'", objconn)

            Button1.Attributes("onclick") = "CheckPrint('Y');return false;"
            Button3.Attributes("onclick") = "CheckPrint('N');return false;"

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button12_Click(sender, e)
            End If
        End If

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

    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click
        Dim dt As DataTable
        Dim RWClass As String = Convert.ToString(Request("RWClass"))

        Dim SqlStr As String = ""
        SqlStr = "" & vbCrLf
        SqlStr += " Select i.Years,i.PlanID,i.OCID,i.ClassCName,i.STDate,i.FTDate,i.TMID" & vbCrLf
        SqlStr += " ,i.IsApplic,i.CLSID,i.CyclType,i.ComIDNO,i.SeqNO,i.LevelType " & vbCrLf
        SqlStr += " ,i.Years+'0'+c.ClassID+i.CyclType ClassID" & vbCrLf
        SqlStr += " ,c.TPlanID" & vbCrLf
        SqlStr += " ,t.TrainID,t.TrainName " & vbCrLf
        SqlStr += " FROM Class_ClassInfo i" & vbCrLf
        SqlStr += " JOIN ID_Class c on  i.CLSID=c.CLSID and c.DistID = '" & sm.UserInfo.DistID & "'" & vbCrLf
        SqlStr += " LEFT JOIN Key_TrainType t on t.TMID=i.TMID " & vbCrLf
        If sPlanKind = "1" AndAlso RWClass = "1" Then
            SqlStr += "JOIN Auth_AccRWClass AAC on a.OCID=AAC.OCID AND AAC.Account='" & sm.UserInfo.UserID & "'" & vbCrLf
        End If
        SqlStr += " where i.IsSuccess='Y' and i.NotOpen='N' " & vbCrLf
        If sm.UserInfo.RID = "A" Then '署(局)
            SqlStr += " and c.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            SqlStr += " and i.Years ='" & sm.UserInfo.Years & "'" & vbCrLf

        Else '其他轄區
            If Request("AcctPlan") <> "1" Then '表示不可以跨計畫選擇
                SqlStr += " and i.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf
            Else
                '表示可跨計畫選擇，依使用權限決定
                If RIDValue.Value <> "" Then
                    SqlStr += " and i.RID='" & RIDValue.Value & "'" & vbCrLf
                Else
                    SqlStr += " and i.RID='" & sm.UserInfo.RID & "'" & vbCrLf
                End If
            End If
        End If

        If OCIDValue1.Value.ToString <> "" Then
            SqlStr += " and i.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        If CyclType.Value.ToString <> "" Then
            If Len(CyclType.Value.ToString) = 1 Then
                SqlStr += " and i.CyclType= '0" & CyclType.Value & "'" & vbCrLf
            Else
                SqlStr += " and i.CyclType='" & CyclType.Value & "'" & vbCrLf
            End If
        End If

        '班級範圍
        Select Case ClassRound.SelectedIndex
            Case 0 '開訓二週前
                SqlStr &= " AND DATEDIFF(DAY,GETDATE(),i.STDate) <= (2*7) AND DATEDIFF(DAY,GETDATE(),i.STDate) >0" & vbCrLf
            Case 1 '已開訓('己開訓改成除排己經結訓的班級)
                SqlStr &= " AND DATEDIFF(DAY,i.STDate,GETDATE()) >=0 AND DATEDIFF(DAY,i.FTDate,GETDATE()) < 0" & vbCrLf
            Case 2 '已結訓
                SqlStr &= " AND DATEDIFF(DAY,i.FTDate,GETDATE()) >=0" & vbCrLf
            Case 3 '未開訓
                SqlStr &= " AND DATEDIFF(DAY,GETDATE(),i.STDate) > 0 " & vbCrLf
            Case 4 '全部
            Case Else '異常
                SqlStr &= " AND 1<>1 "
        End Select

        dt = DbAccess.GetDataTable(SqlStr, objconn)

        msg.Visible = True
        msg.Text = "查無資料!!"

        Table4.Visible = False
        PageControler1.Visible = False
        Button1.Visible = False
        Button3.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Visible = False
            msg.Text = ""

            Table4.Visible = True
            PageControler1.Visible = True
            Button1.Visible = True
            Button3.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "ClassID,CyclType"
            PageControler1.ControlerLoad()

        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dr_Class As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem

                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim labelStar As Label = e.Item.FindControl("LabelStar")
                Checkbox1.Value = Val(drv("OCID"))
                Checkbox1.Attributes("onclick") = "InsertValue(this.checked,this.value)"
                'If PrintValue.Value.IndexOf(Checkbox1.Value) <> -1 Then
                If PrintValue.Value.IndexOf(drv("OCID")) <> -1 Then
                    Checkbox1.Checked = True
                End If
                'by Milor 20080804--此限制要排除學習卷----start
                If sm.UserInfo.TPlanID <> "15" Then
                    'by Milor 20080715----start
                    '加入判斷學員資料中是否有必要欄位未填，有則顯示星號並鎖定無法核取。
                    If Chk_ApprAll(Convert.ToInt32(drv("OCID"))) = True Then
                        labelStar.Visible = True
                        Checkbox1.Disabled = True
                        TIMS.Tooltip(labelStar, "學員資料中有必要欄位未填!!", True)
                        TIMS.Tooltip(Checkbox1, "學員資料中有必要欄位未填!!", True)
                    Else
                        labelStar.Visible = False
                        Checkbox1.Disabled = False
                    End If
                    'by Milor 20080715----end
                Else
                    labelStar.Visible = False
                    Checkbox1.Disabled = False
                End If
                'by Milor 20080804

                e.Item.Cells(1).Text = "[" & drv("TrainID").ToString & "]" & drv("TrainName").ToString

                e.Item.Cells(3).Text = TIMS.GET_CLASSNAME(Convert.ToString(drv("ClassCName")), Convert.ToString(drv("CyclType")))

                If e.Item.Cells(5).Text = "Y" Or e.Item.Cells(5).Text = "y" Then
                    e.Item.Cells(5).Text = "可挑選志願"
                Else
                    e.Item.Cells(5).Text = "不可挑選志願"
                End If

        End Select
    End Sub

    '**by Milor 20080715----start
    Private Function Chk_ApprAll(ByVal tmpOCID As Integer) As Boolean
        '此函數用來判斷學員資料維護的必要欄位是否有填寫完整，為了避免資料庫中被存入NULL，
        '所以判斷時都一致使用nvl轉換。
        'PS.此檢核對應學員資料維護的*號標示。
        Dim rst As Boolean = False

        Dim dt As DataTable = Nothing
        Dim sqlStr As String = ""
        sqlStr = "" & vbCrLf
        sqlStr += " select a.SOCID " & vbCrLf
        sqlStr += " from Class_StudentsOfClass a " & vbCrLf
        sqlStr += " join Stud_StudentInfo b on a.SID=b.SID " & vbCrLf
        sqlStr += " join Stud_SubData c on c.SID=a.SID " & vbCrLf
        sqlStr += " where a.OCID=" & tmpOCID & vbCrLf
        sqlStr += " and (1!=1" & vbCrLf
        sqlStr += " or dbo.NVL(a.MidentityID,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(a.IdentityID,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(b.MilitaryID,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(c.Email,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(c.school,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(c.Department,' ')=' ' " & vbCrLf
        sqlStr += " or dbo.NVL(c.PhoneD,' ')=' '" & vbCrLf
        sqlStr += " or dbo.NVL(c.address,' ')=' ' " & vbCrLf
        '產學訓與非產學訓的例外檢核
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sqlStr += " or dbo.NVL(a.SupplyID,' ')=' ' " & vbCrLf
        Else
            sqlStr += " or dbo.NVL(a.BudgetID,' ')=' ' " & vbCrLf
            sqlStr += " or dbo.NVL(a.SubsidyID,' ')=' ' " & vbCrLf
            sqlStr += " or dbo.NVL(b.EngName,' ')=' ' " & vbCrLf
            sqlStr += " or dbo.NVL(b.IsAgree,' ')=' ' " & vbCrLf
            sqlStr += " or dbo.NVL(c.EmergencyContact,' ')=' ' " & vbCrLf
            sqlStr += " or dbo.NVL(c.EmergencyRelation,' ')=' ' " & vbCrLf
            sqlStr &= " or c.ZipCode3 IS NULL" & vbCrLf
            sqlStr += " or dbo.NVL(c.ShowDetail,' ')=' ' " & vbCrLf
        End If
        sqlStr += " )" & vbCrLf
        dt = DbAccess.GetDataTable(sqlStr, objconn)
        If dt.Rows.Count > 0 Then
            rst = True
        End If
        Return rst
    End Function

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
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        Table4.Visible = False
        PageControler1.Visible = False
        Button1.Visible = False
    End Sub


End Class
