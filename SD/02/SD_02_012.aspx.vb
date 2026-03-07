Partial Class SD_02_012
    Inherits AuthBasePage

    Const cst_msgAlt1 As String = "該計畫，未提供非自願離職者提醒處理!!"
    'Dim aNow As Date
    'Dim FunDr As DataRow
    'Dim Days1 As Integer
    'Dim Days2 As Integer
    'Dim gFlagEnv As Boolean = True '正式環境。(測試用) / TestStr
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
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End
        ''取出設定天數檔 Start
        'TIMS.Get_SysDays(Days1, Days2)
        ''取出設定天數檔 End
        'Call TIMS.OpenDbConn(objconn)
        'aNow = TIMS.GetSysDateNow(objconn)
        'Dim dr As DataRow

        If Not IsPostBack Then
            msg.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If TIMS.Cst_TPlanID_useCFIRE1.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                Common.MessageBox(Me, cst_msgAlt1)
                Exit Sub
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button5.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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

        '查詢檢查。
        BtnSearch1.Attributes("onclick") = "javascript:return search()"

        If Not IsPostBack Then
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub

    '查詢Sql
    Sub sUtl_Search1()
        'If TestStr = "AmuTest" Then gFlagEnv = False '測試用。
        Dim gFlagEnv As Boolean = True '正式環境。(測試用) / TestStr
        If TIMS.sUtl_ChkTest() Then
            Common.MessageBox(Me, "測試環境使用!!")
            gFlagEnv = False '測試用。
        End If
        Dim sql As String = ""
        If Not gFlagEnv Then
            sql = "" & vbCrLf
            sql += " SELECT a.SETID" & vbCrLf
            sql += " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
            sql += " ,b.SERNUM" & vbCrLf
            sql += " ,a.IDNO" & vbCrLf
            sql += " ,b.OCID1" & vbCrLf
            sql += " ,b.EnterPath" & vbCrLf
            sql += " ,'1' type1" & vbCrLf
            sql += " ,0 eSETID" & vbCrLf
            sql += " ,b.CFIRE1NS" & vbCrLf '取消提醒
            'sql += " ,b.CFIRE1Reason" & vbCrLf
            sql += " ,b.CFIRE1R2" & vbCrLf
            sql += " ,b.CFIRE1MACCT" & vbCrLf
            sql += " ,a.NAME" & vbCrLf
            sql += " ,cc.Years" & vbCrLf
            sql += " ,cc.PlanName" & vbCrLf
            sql += " ,cc.DistName" & vbCrLf
            sql += " ,cc.OrgName" & vbCrLf
            sql += " ,cc.ClassCName" & vbCrLf
            sql += " FROM STUD_ENTERTEMP a" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE b ON b.SETID =a.SETID" & vbCrLf
            sql += " JOIN VIEW2 cc on cc.ocid =b.OCID1" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            sql += " AND cc.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND b.ENTERPATH NOT IN ('W','R')" & vbCrLf
            sql += " AND b.CFIRE1='Y'" & vbCrLf '非自願離職者
            'sql += " AND b.CFIRE1NS IS NULL" & vbCrLf

            sql += " AND NOT EXISTS (" & vbCrLf
            sql += " SELECT 'X'" & vbCrLf
            sql += " FROM STUD_ENTERTEMP2 xa" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE2 xb on xb.eSETID =xa.eSETID" & vbCrLf
            sql += " JOIN CLASS_CLASSINFO ccx on ccx.OCID=xb.OCID1" & vbCrLf
            sql += " JOIN ID_PLAN ipx on ipx.PlanID=ccx.PlanID" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            sql += " AND ipx.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND xb.CFIRE1='Y'" & vbCrLf '非自願離職者
            'sql += " AND xb.CFIRE1NS IS NULL" & vbCrLf '取消提醒
            sql += " AND xb.OCID1=b.OCID1 AND xa.IDNO=a.IDNO" & vbCrLf
            sql += " )" & vbCrLf

            sql += " UNION" & vbCrLf
            sql += " SELECT a.eSETID" & vbCrLf
            sql += " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
            sql += " ,b.eSERNUM" & vbCrLf
            sql += " ,a.IDNO" & vbCrLf
            sql += " ,b.OCID1" & vbCrLf
            sql += " ,b.EnterPath" & vbCrLf
            sql += " ,'2' type1" & vbCrLf
            sql += " ,b.eSETID" & vbCrLf
            sql += " ,b.CFIRE1NS" & vbCrLf '取消提醒
            'sql += " ,b.CFIRE1Reason" & vbCrLf
            sql += " ,b.CFIRE1R2" & vbCrLf
            sql += " ,b.CFIRE1MACCT" & vbCrLf
            sql += " ,a.NAME" & vbCrLf
            sql += " ,cc.Years" & vbCrLf
            sql += " ,cc.PlanName" & vbCrLf
            sql += " ,cc.DistName" & vbCrLf
            sql += " ,cc.OrgName" & vbCrLf
            sql += " ,cc.ClassCName" & vbCrLf
            sql += " FROM STUD_ENTERTEMP2 a" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE2 b on b.ESETID =a.ESETID" & vbCrLf
            sql += " JOIN VIEW2 cc on cc.ocid =b.OCID1" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            'sm.UserInfo.TPlanID
            sql += " AND cc.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND b.CFIRE1='Y'" & vbCrLf '非自願離職者
        End If

        If gFlagEnv Then
            '正式
            If RIDValue.Value = "" Then
                Common.MessageBox(Me, "未選擇有效機構!")
                Exit Sub
            End If
            If OCIDValue1.Value = "" Then
                Common.MessageBox(Me, "未選擇有效班級!")
                Exit Sub
            End If

            sql = "" & vbCrLf
            sql += " SELECT a.SETID" & vbCrLf
            sql += " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
            sql += " ,b.SERNUM" & vbCrLf
            sql += " ,a.IDNO" & vbCrLf
            sql += " ,b.OCID1" & vbCrLf
            sql += " ,b.EnterPath" & vbCrLf
            sql += " ,'1' type1" & vbCrLf
            sql += " ,0 eSETID" & vbCrLf
            sql += " ,b.CFIRE1NS" & vbCrLf '取消提醒
            'sql += " ,b.CFIRE1Reason" & vbCrLf
            sql += " ,b.CFIRE1R2" & vbCrLf
            sql += " ,b.CFIRE1MACCT" & vbCrLf
            sql += " ,a.NAME" & vbCrLf
            sql += " ,cc.Years" & vbCrLf
            sql += " ,cc.PlanName" & vbCrLf
            sql += " ,cc.DistName" & vbCrLf
            sql += " ,cc.OrgName" & vbCrLf
            sql += " ,cc.ClassCName" & vbCrLf
            sql += " FROM STUD_ENTERTEMP a" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE b on b.SETID =a.SETID" & vbCrLf
            sql += " JOIN VIEW2 cc on cc.ocid =b.OCID1" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            'sm.UserInfo.TPlanID
            sql += " AND cc.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND b.ENTERPATH NOT IN ('W','R')" & vbCrLf
            sql += " AND b.CFIRE1='Y'" & vbCrLf '非自願離職者
            'sql += " AND b.CFIRE1NS IS NULL" & vbCrLf
            sql += " and cc.RID=@RID " & vbCrLf
            If OCIDValue1.Value <> "" Then
                sql += " and cc.OCID=@OCID" & vbCrLf
            End If

            sql += " AND NOT EXISTS (" & vbCrLf
            sql += " SELECT 'X'" & vbCrLf
            sql += " FROM STUD_ENTERTEMP2 xa" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE2 xb on xb.eSETID =xa.eSETID" & vbCrLf
            sql += " JOIN CLASS_CLASSINFO ccx on ccx.OCID=xb.OCID1" & vbCrLf
            sql += " JOIN ID_PLAN ipx on ipx.PlanID=ccx.PlanID" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            'sm.UserInfo.TPlanID
            sql += " AND ipx.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND xb.CFIRE1='Y'" & vbCrLf '非自願離職者
            'sql += " AND xb.CFIRE1NS IS NULL" & vbCrLf '取消提醒
            sql += " AND xb.OCID1=b.OCID1 AND xa.IDNO=a.IDNO" & vbCrLf
            sql += " and ccx.RID=@RID " & vbCrLf
            If OCIDValue1.Value <> "" Then
                sql += " and xb.OCID1=@OCID" & vbCrLf
            End If
            sql += " )" & vbCrLf

            sql += " UNION" & vbCrLf

            sql += " SELECT a.eSETID" & vbCrLf
            sql += " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
            sql += " ,b.eSERNUM" & vbCrLf
            sql += " ,a.IDNO" & vbCrLf
            sql += " ,b.OCID1" & vbCrLf
            sql += " ,b.EnterPath" & vbCrLf
            sql += " ,'2' type1" & vbCrLf
            sql += " ,b.eSETID" & vbCrLf
            sql += " ,b.CFIRE1NS" & vbCrLf '取消提醒
            'sql += " ,b.CFIRE1Reason" & vbCrLf
            sql += " ,b.CFIRE1R2" & vbCrLf
            sql += " ,b.CFIRE1MACCT" & vbCrLf
            sql += " ,a.NAME" & vbCrLf
            sql += " ,cc.Years" & vbCrLf
            sql += " ,cc.PlanName" & vbCrLf
            sql += " ,cc.DistName" & vbCrLf
            sql += " ,cc.OrgName" & vbCrLf
            sql += " ,cc.ClassCName" & vbCrLf
            sql += " FROM STUD_ENTERTEMP2 a" & vbCrLf
            sql += " JOIN STUD_ENTERTYPE2 b on b.ESETID =a.ESETID" & vbCrLf
            sql += " JOIN VIEW2 cc on cc.ocid =b.OCID1" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            'sm.UserInfo.TPlanID
            sql += " AND cc.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql += " AND b.CFIRE1='Y'" & vbCrLf '非自願離職者
            'sql += " AND b.CFIRE1NS IS NULL" & vbCrLf '取消提醒
            sql += " and cc.RID=@RID " & vbCrLf
            If OCIDValue1.Value <> "" Then
                sql += " and cc.OCID=@OCID" & vbCrLf
            End If
        End If

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            If gFlagEnv Then
                '.Parameters.Add("xxx", SqlDbType.VarChar).Value = ""
                If RIDValue.Value <> "" Then
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    'sql += " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
                Else
                    .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
                End If
                If OCIDValue1.Value <> "" Then
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                    'sql += " and cc.OCID='" & OCIDValue1.Value & "'" & vbCrLf
                End If
            End If
            dt.Load(.ExecuteReader())
        End With

        msg.Text = "查無資料!"
        Table4.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch1.Click
        If TIMS.Cst_TPlanID_useCFIRE1.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, cst_msgAlt1)
            Exit Sub
        End If

        '查詢Sql
        Call sUtl_Search1()
    End Sub

    '設定 RadioButtonList
    Sub SetRdoCFIRE1R2(ByRef oCFIRE1R2 As RadioButtonList)
        With oCFIRE1R2
            .Items.Clear()
            .Items.Add(New ListItem("持推介單", "1"))
            .Items.Add(New ListItem("簽立權益說明暨同意書", "2"))
            .Items.Add(New ListItem("未甄試", "3"))
            .Items.Add(New ListItem("未錄取", "4"))
            .Items.Add(New ListItem("未報到", "5"))
        End With
        'Common.SetListItem(oCFIRE1R2, "1")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header, ListItemType.Footer
            Case Else
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                'Dim CFIRE1Reason As TextBox = e.Item.FindControl("CFIRE1Reason") '取消提醒
                Dim CFIRE1R2 As RadioButtonList = e.Item.FindControl("CFIRE1R2") '取消提醒(處理說明選擇)
                Dim hid_OCID1 As HiddenField = e.Item.FindControl("hid_OCID1")
                Dim hid_IDNO As HiddenField = e.Item.FindControl("hid_IDNO")
                Dim hid_TYPE As HiddenField = e.Item.FindControl("hid_TYPE")
                Dim hid_SETID As HiddenField = e.Item.FindControl("hid_SETID")
                Dim hid_eSETID As HiddenField = e.Item.FindControl("hid_eSETID")
                'Dim CheckBox1 As CheckBox = e.Item.FindControl("CheckBox1")
                'CFIRE1Reason.Text = Convert.ToString(drv("CFIRE1Reason"))
                'CFIRE1Reason.MaxLength = 100
                Call SetRdoCFIRE1R2(CFIRE1R2) '設定 RadioButtonList

                If Convert.ToString(drv("CFIRE1R2")) <> "" Then
                    Common.SetListItem(CFIRE1R2, drv("CFIRE1R2"))
                End If
                hid_OCID1.Value = drv("OCID1")
                hid_IDNO.Value = drv("IDNO")
                hid_TYPE.Value = Convert.ToString(drv("TYPE1"))
                hid_SETID.Value = Convert.ToString(drv("SETID"))
                hid_eSETID.Value = Convert.ToString(drv("eSETID"))

                'CheckBox1.Checked = False
                'If Convert.ToString(drv("cFire1NS")) = "Y" Then '取消提醒
                '    CheckBox1.Checked = True
                'End If
        End Select
    End Sub

    '單一班級搜尋
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        Table4.Visible = False
    End Sub

    '單一班級搜尋
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    '儲存。
    Protected Sub btnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Call SaveData1()
    End Sub

    '儲存。
    Sub SaveData1()
        '2006/03/ add conn by matt
        'aNow = TIMS.GetSysDateNow(objconn)

        'CHECK
        Dim ERRMSG As String = ""
        ERRMSG = ""

        'Const cst_姓名 As Integer = 1          ' Cells(cst_姓名)
        'Const cst_報名機構 As Integer = 2
        'Const cst_報名班級 As Integer = 3
        'Const cst_報名日期 As Integer = 4      'Cells(5) = Cells(cst_報名日期)

        Dim i As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            i += 1
            'Dim CFIRE1Reason As TextBox = eItem.FindControl("CFIRE1Reason") '取消提醒說明
            Dim CFIRE1R2 As RadioButtonList = eItem.FindControl("CFIRE1R2") '取消提醒說明
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_TYPE As HiddenField = eItem.FindControl("hid_TYPE")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")
            Dim hid_eSETID As HiddenField = eItem.FindControl("hid_eSETID")
            'Dim CheckBox1 As CheckBox = eItem.FindControl("CheckBox1")
            'CFIRE1Reason.Text = TIMS.ClearSQM(CFIRE1Reason.Text)
            'If CheckBox1.Checked AndAlso CFIRE1Reason.Text = "" Then
            '    ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ") 取消提醒,請輸入處理說明。"
            'End If
            'If CFIRE1Reason.Text.Length > 100 Then
            '    CFIRE1Reason.Text = Microsoft.VisualBasic.Left(CFIRE1Reason.Text, 100)
            'End If
            'If CheckBox1.Checked AndAlso CFIRE1R2.SelectedValue = "" Then
            '    ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ") 取消提醒,請選擇處理說明。"
            'End If
            'If CheckBox1.Enabled AndAlso CheckBox1.Checked Then
            '    If CFIRE1R2.SelectedValue = "" Then
            '        ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ") 取消提醒,請選擇處理說明。" & vbCrLf
            '    End If
            'End If
        Next
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " UPDATE STUD_ENTERTYPE" & vbCrLf
        sql += " SET cFire1NS=@cFire1NS " & vbCrLf '取消提醒
        'sql += " ,CFIRE1Reason=@CFIRE1Reason " & vbCrLf
        sql += " ,CFIRE1R2=@CFIRE1R2" & vbCrLf
        sql += " ,CFIRE1MACCT=@CFIRE1MACCT" & vbCrLf
        sql += " ,CFIRE1MDATE= getdate()" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND CFIRE1='Y'" & vbCrLf '非自願離職者
        sql += " AND SETID =@SETID " & vbCrLf
        'sql += " AND ENTERDATE =@ENTERDATE" & vbCrLf
        'sql += " AND SERNUM =@SERNUM" & vbCrLf
        sql += " AND OCID1 =@OCID1" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " SELECT b.SETID,b.OCID1 "
        sql += " FROM STUD_ENTERTEMP a" & vbCrLf
        sql += " JOIN STUD_ENTERTYPE b ON b.SETID=a.SETID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.IDNO =@IDNO" & vbCrLf
        sql += " AND b.OCID1 =@OCID1" & vbCrLf
        Dim sCmd1 As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " UPDATE STUD_ENTERTYPE2" & vbCrLf
        sql += " SET cFire1NS=@cFire1NS " & vbCrLf '取消提醒
        'sql += " ,CFIRE1Reason=@CFIRE1Reason " & vbCrLf
        sql += " ,CFIRE1R2=@CFIRE1R2" & vbCrLf
        sql += " ,CFIRE1MACCT=@CFIRE1MACCT" & vbCrLf
        sql += " ,CFIRE1MDATE= getdate()" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND CFIRE1='Y'" & vbCrLf '非自願離職者
        sql += " AND eSETID =@eSETID " & vbCrLf
        'sql += " AND eSERNUM =@eSERNUM" & vbCrLf
        sql += " AND OCID1 =@OCID1" & vbCrLf
        Dim uCmd2 As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " SELECT b.eSETID,b.OCID1 "
        sql += " FROM STUD_ENTERTEMP2 a" & vbCrLf
        sql += " JOIN STUD_ENTERTYPE2 b ON b.eSETID=a.eSETID" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND a.IDNO =@IDNO" & vbCrLf
        sql += " AND b.OCID1 =@OCID1" & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)

        'SAVE
        Call TIMS.OpenDbConn(objconn)
        For Each eItem As DataGridItem In DataGrid1.Items
            i += 1
            'Dim CFIRE1Reason As TextBox = eItem.FindControl("CFIRE1Reason") '取消提醒
            Dim CFIRE1R2 As RadioButtonList = eItem.FindControl("CFIRE1R2") '取消提醒(處理說明選擇)
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_TYPE As HiddenField = eItem.FindControl("hid_TYPE")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")
            Dim hid_eSETID As HiddenField = eItem.FindControl("hid_eSETID")
            Dim CheckBox1 As CheckBox = eItem.FindControl("CheckBox1")

            Select Case hid_TYPE.Value
                Case "1"
                    If CFIRE1R2.SelectedValue <> "" Then
                        Dim dt2 As New DataTable
                        With sCmd2
                            .Parameters.Clear()
                            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = hid_IDNO.Value
                            .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(hid_OCID1.Value)
                            dt2.Load(.ExecuteReader())
                        End With
                        With uCmd
                            .Parameters.Clear()
                            .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                            'If CheckBox1.Checked Then
                            '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                            'Else
                            '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = Convert.DBNull '取消提醒
                            'End If
                            '.Parameters.Add("CFIRE1Reason", SqlDbType.NVarChar).Value = CFIRE1Reason.Text
                            .Parameters.Add("CFIRE1R2", SqlDbType.VarChar).Value = CFIRE1R2.SelectedValue
                            .Parameters.Add("CFIRE1MACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("SETID", SqlDbType.Int).Value = Val(hid_SETID.Value)
                            .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(hid_OCID1.Value)
                            .ExecuteNonQuery()
                        End With
                        For Each dr As DataRow In dt2.Rows '(STUD_ENTERTYPE2)
                            '同OCID1資料修正。
                            With uCmd2
                                .Parameters.Clear()
                                .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                                'If CheckBox1.Checked Then
                                '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                                'Else
                                '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = Convert.DBNull '取消提醒
                                'End If
                                '.Parameters.Add("CFIRE1Reason", SqlDbType.NVarChar).Value = CFIRE1Reason.Text
                                .Parameters.Add("CFIRE1R2", SqlDbType.VarChar).Value = CFIRE1R2.SelectedValue
                                .Parameters.Add("CFIRE1MACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                .Parameters.Add("eSETID", SqlDbType.Int).Value = Val(dr("eSETID"))
                                .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(dr("OCID1"))
                                .ExecuteNonQuery()
                            End With
                        Next
                    End If


                Case "2"
                    If CFIRE1R2.SelectedValue <> "" Then
                        Dim dt1 As New DataTable
                        With sCmd1
                            .Parameters.Clear()
                            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = hid_IDNO.Value
                            .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(hid_OCID1.Value)
                            dt1.Load(.ExecuteReader())
                        End With
                        With uCmd2
                            .Parameters.Clear()
                            .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                            'If CheckBox1.Checked Then
                            '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                            'Else
                            '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = Convert.DBNull '取消提醒
                            'End If
                            '.Parameters.Add("CFIRE1Reason", SqlDbType.NVarChar).Value = CFIRE1Reason.Text
                            .Parameters.Add("CFIRE1R2", SqlDbType.VarChar).Value = CFIRE1R2.SelectedValue
                            .Parameters.Add("CFIRE1MACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("eSETID", SqlDbType.Int).Value = Val(hid_eSETID.Value)
                            .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(hid_OCID1.Value)
                            .ExecuteNonQuery()
                        End With
                        For Each dr As DataRow In dt1.Rows '(STUD_ENTERTYPE)
                            '同OCID1資料修正。
                            With uCmd
                                .Parameters.Clear()
                                .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                                'If CheckBox1.Checked Then
                                '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = "Y" '取消提醒
                                'Else
                                '    .Parameters.Add("cFire1NS", SqlDbType.VarChar).Value = Convert.DBNull '取消提醒
                                'End If
                                '.Parameters.Add("CFIRE1Reason", SqlDbType.NVarChar).Value = CFIRE1Reason.Text
                                .Parameters.Add("CFIRE1R2", SqlDbType.VarChar).Value = CFIRE1R2.SelectedValue
                                .Parameters.Add("CFIRE1MACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                .Parameters.Add("SETID", SqlDbType.Int).Value = Val(dr("SETID"))
                                .Parameters.Add("OCID1", SqlDbType.Int).Value = Val(dr("OCID1"))
                                .ExecuteNonQuery()
                            End With
                        Next

                    End If

            End Select
        Next

        Common.MessageBox(Me, "儲存成功!")

    End Sub

End Class