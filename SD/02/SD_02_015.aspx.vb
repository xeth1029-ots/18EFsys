Partial Class SD_02_015
    Inherits AuthBasePage

    'STUD_SELRESULTBLI 
    'STUD_SELRESULTBLIDET 
    'STATUSNC1: 1:已提供勞保明細表 /2:未甄試/3:未錄取/4:未報到
    'STATUSNC2: 1:未有工作事實 /2:已做離退訓
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim vMsg As String = ""
    Dim strSS As String = ""
    'Const cst_saveMsgST2_2 As String = "請將該名學員進行離退訓作業。倘使用者未完成渠等學員離退訓作業，則限制使用者無法使用學員出缺勤作業功能"
    Const cst_saveMsgST2_2 As String = "請將該名學員進行離退訓作業。"

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        flgROLEIDx0xLIDx0 = False
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'BtnSearch1.Enabled = True
        'If Not au.blnCanSech Then
        '    BtnSearch1.Enabled = False
        '    TIMS.Tooltip(BtnSearch1, "無查詢權限")
        'End If
        '檢查帳號的功能權限-----------------------------------End

        If Not IsPostBack Then
            btnSave1.Attributes("onclick") = "return confirm('儲存後無法再做修正\n請確認儲存資料正確性\n是否確定儲存?');"
            '查詢檢查。
            BtnSearch1.Attributes("onclick") = "javascript:return search1();"

            msg.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
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

    End Sub

    '查詢 SQL
    Sub sUtl_Search1()
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "未選擇有效機構")
            Exit Sub
        End If
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "未選擇有效班級")
            Exit Sub
        End If
        Dim drC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drC Is Nothing Then
            Common.MessageBox(Me, "未選擇有效班級!!")
            Exit Sub
        End If
        HidPlanID.Value = CStr(drC("PlanID"))

        Dim dtStud As DataTable = Nothing
        Select Case rblType1.SelectedValue
            Case "1" '甄試日前
            Case "2" '開訓日前
            Case "3", "4" '3:開訓日後 '4:訓期已滿1/2 
                '查詢目前計畫所有學員
                strSS = "" 'Dim strSS As String = ""
                TIMS.SetMyValue(strSS, "PlanID", HidPlanID.Value)
                TIMS.SetMyValue(strSS, "RID", RIDValue.Value)
                TIMS.SetMyValue(strSS, "OCID", OCIDValue1.Value)
                dtStud = TIMS.Get_STUDENTINFO(strSS, objconn)
            Case Else
                Common.MessageBox(Me, "未選擇查詢種類!!")
                Exit Sub
        End Select

        '查詢種類
        Me.ViewState("rblType1") = rblType1.SelectedValue
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.SETID" & vbCrLf
        sql &= " ,CONVERT(varchar, b.EnterDate, 111) EnterDate" & vbCrLf
        sql &= " ,b.SERNUM" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,dbo.fn_GET_MASK1(a.IDNO) MKIDNO" & vbCrLf
        sql &= " ,b.OCID1" & vbCrLf
        sql &= " ,b.EnterPath" & vbCrLf
        'sql &= " ,'1' type1" & vbCrLf 'sql &= " ,0 eSETID" & vbCrLf
        sql &= " ,b.modifyacct" & vbCrLf
        sql &= " ,bd.SBDID" & vbCrLf
        sql &= " ,bd.SB3ID" & vbCrLf
        sql &= " ,bd.STATUSPT" & vbCrLf
        sql &= " ,bd.STATUSNC1" & vbCrLf
        sql &= " ,bd.STATUSNC2" & vbCrLf
        sql &= " ,bd.STATUSNCREASON" & vbCrLf
        'sql &= " --,b.CMASTER1NS,b.CMASTER1NT,b.CMASTER1Reason,b.CMASTER1MACCT" & vbCrLf
        sql &= " ,a.NAME" & vbCrLf
        sql &= " ,cc.Years" & vbCrLf
        sql &= " ,cc.PlanName" & vbCrLf
        sql &= " ,cc.DistName" & vbCrLf
        sql &= " ,cc.OrgName" & vbCrLf
        sql &= " ,cc.ClassCName" & vbCrLf
        sql &= " ,bi.ACTNO" & vbCrLf
        sql &= " ,bi.COMNAME" & vbCrLf

        sql &= " FROM STUD_SELRESULTBLI bi " & vbCrLf
        sql &= " JOIN STUD_SELRESULTBLIDET bd on bi.SB3ID=bd.SB3ID" & vbCrLf
        'sql &= " JOIN STUD_SELRESULTBLIDET bd on bi.SB3ID=bd.SB3ID AND bd.STATUSPT IN ('ES3','ES2')" & vbCrLf
        'sql &= " LEFT JOIN STUD_SELRESULTBLIDET bd2 on bi.SB3ID=bd2.SB3ID AND bd.STATUSPT IN ('AS2')" & vbCrLf
        'sql &= " LEFT JOIN STUD_SELRESULTBLIDET bd2 on bi.SB3ID=bd2.SB3ID AND bd.STATUSPT IN ('ST1')" & vbCrLf
        'sql &= " --and bd.STATUSPT in ('AS2') and bi.ACTNO is not null" & vbCrLf
        sql &= " JOIN VIEW2 cc on cc.ocid =bd.ocid" & vbCrLf
        sql &= " JOIN VIEW_PLAN ip on ip.planid =cc.planid" & vbCrLf
        'sql &= " join Stud_SelResult t1 on t1.setid=bi.setid and t1.enterdate=bi.enterdate and t1.sernum=bi.sernum and dbo.NVL(t1.Admission,'Y')='Y'" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP a on a.SETID =bi.SETID" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b on b.setid=bi.setid and b.enterdate=bi.enterdate and b.sernum=bi.sernum" & vbCrLf
        'sql &= " JOIN STUD_ENTERTYPE b on b.setid=bi.setid and b.OCID1=bd.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " /* AND cc.STDATE > dbo.TRUNC_DATETIME(getdate()) */" & vbCrLf
        sql &= " and ip.tplanid in (" & TIMS.Cst_TPlanID_BliDet201605 & ")" & vbCrLf
        sql &= " and ip.tplanid not in (" & TIMS.Cst_TPlanID_BliDet201605_NG & ")" & vbCrLf
        'sql &= " and ip.tplanid in ('02','14','62','17','20','21','61','26','34','37','47','50','51','53','55','58','59','64','65')" & vbCrLf
        'sql &= " and ip.tplanid not in ('58','47')" & vbCrLf
        sql &= " AND cc.IsSuccess = 'Y'" & vbCrLf
        sql &= " AND cc.NotOpen = 'N'" & vbCrLf
        If Not flgROLEIDx0xLIDx0 Then
            sql &= " and cc.PlanID=@PlanID" & vbCrLf
        End If
        sql &= " AND cc.RID =@RID" & vbCrLf
        sql &= " AND cc.OCID =@OCID" & vbCrLf
        '<asp@ListItem Selected="True" Value="1">甄試日前</asp@ListItem>
        '<asp@ListItem Value="2">開訓日前</asp@ListItem>
        '<asp@ListItem Value="3">開訓日後</asp@ListItem>
        Select Case rblType1.SelectedValue
            Case "1" '甄試日前
                sql &= " AND bd.STATUSPT IN ('ES3','ES2')" & vbCrLf
            Case "2" '開訓日前
                sql &= " AND bd.STATUSPT IN ('AS2')" & vbCrLf
            Case "3" '開訓日後
                sql &= " AND bd.STATUSPT IN ('ST1')" & vbCrLf
            Case "4" '4:訓期已滿1/2 
                sql &= " AND bd.STATUSPT IN ('ST2')" & vbCrLf
            Case Else
                sql &= " AND 1<>1" & vbCrLf
        End Select
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            If Not flgROLEIDx0xLIDx0 Then
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
            End If
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            dt.Load(.ExecuteReader())
        End With

        msg.Text = "查無資料!"
        Table4.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Table4.Visible = True

            Select Case rblType1.SelectedValue
                Case "1" '甄試日前
                    PageControler1.PageDataTable = dt
                    PageControler1.ControlerLoad()

                Case "2" '開訓日前
                    PageControler1.PageDataTable = dt
                    PageControler1.ControlerLoad()

                Case "3", "4"
                    '3:開訓日後 '4:訓期已滿1/2 
                    Dim dt2 As DataTable = dt.Copy.Clone
                    For Each dr As DataRow In dt.Rows
                        '查詢目前計畫所有學員
                        Dim ff As String = ""
                        ff = "IDNO='" & dr("IDNO") & "' AND OCID='" & dr("OCID1") & "'"
                        If dtStud.Select(ff).Length > 0 Then
                            Dim dr2 As DataRow = dt2.NewRow
                            dr2.ItemArray = dr.ItemArray
                            dt2.Rows.Add(dr2)
                        End If
                    Next
                    dt2.AcceptChanges()

                    msg.Text = "查無資料!"
                    Table4.Visible = False
                    If dt2.Rows.Count > 0 Then
                        msg.Text = ""
                        Table4.Visible = True

                        PageControler1.PageDataTable = dt2
                        PageControler1.ControlerLoad()
                    End If

                Case Else
                    Common.MessageBox(Me, "未選擇查詢種類!!")
                    Exit Sub
            End Select

        End If

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch1.Click
        '查詢Sql
        Call sUtl_Search1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Call SHOW_ITEM1(sender, e)
                'Case ListItemType.Header, ListItemType.Footer
        End Select
    End Sub

    Sub SHOW_ITEM1(ByRef sender As Object, ByRef e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim drv As DataRowView = e.Item.DataItem
        e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

        Dim LabIDNO2 As Label = e.Item.FindControl("LabIDNO2")
        Dim STATUSNC1 As RadioButtonList = e.Item.FindControl("STATUSNC1")
        Dim STATUSNC2 As RadioButtonList = e.Item.FindControl("STATUSNC2")
        Dim STATUSNCREASON As TextBox = e.Item.FindControl("STATUSNCREASON")
        Dim Hid_STATUSPT As HiddenField = e.Item.FindControl("Hid_STATUSPT")
        Dim Hid_SBDID As HiddenField = e.Item.FindControl("Hid_SBDID")
        Dim Hid_SB3ID As HiddenField = e.Item.FindControl("Hid_SB3ID")
        Dim hid_OCID1 As HiddenField = e.Item.FindControl("hid_OCID1")
        Dim hid_IDNO As HiddenField = e.Item.FindControl("hid_IDNO")
        Dim hid_SETID As HiddenField = e.Item.FindControl("hid_SETID")
        Dim BtnVIEW1 As Button = e.Item.FindControl("BtnVIEW1")

        STATUSNCREASON.Text = Convert.ToString(drv("STATUSNCREASON"))
        STATUSNCREASON.MaxLength = 100

        Hid_STATUSPT.Value = drv("STATUSPT")
        Hid_SBDID.Value = drv("SBDID")
        Hid_SB3ID.Value = drv("SB3ID")
        hid_OCID1.Value = drv("OCID1")
        hid_IDNO.Value = drv("IDNO")
        hid_SETID.Value = Convert.ToString(drv("SETID"))

        STATUSNCREASON.Enabled = True
        STATUSNC1.Enabled = True
        STATUSNC2.Enabled = True
        If Not Me.ViewState("rblType1") Is Nothing Then
            Select Case Me.ViewState("rblType1")
                Case "1" '甄試日前
                    STATUSNC2.Enabled = False
                Case "2" '開訓日前
                    STATUSNC2.Enabled = False
                Case "3", "4"
                    '3:開訓日後 '4:訓期已滿1/2 
                    STATUSNC1.Enabled = False
                Case Else
                    STATUSNCREASON.Enabled = False
                    STATUSNC1.Enabled = False
                    STATUSNC2.Enabled = False
            End Select
        End If

        If Convert.ToString(drv("STATUSNC1")) <> "" Then '已轉知1
            Common.SetListItem(STATUSNC1, drv("STATUSNC1"))
        End If
        If Convert.ToString(drv("STATUSNC2")) <> "" Then '已轉知2
            Common.SetListItem(STATUSNC2, drv("STATUSNC2"))
        End If

        Dim flagNoUpdate As Boolean = False 'false:可修改'(未)待已轉知
        If Convert.ToString(drv("STATUSNC1")) <> "" Then '已轉知1
            flagNoUpdate = True '不可修改'待已轉知
        End If
        If Convert.ToString(drv("STATUSNC2")) <> "" Then '已轉知2
            flagNoUpdate = True '不可修改'待已轉知
        End If
        LabIDNO2.Text = Convert.ToString(drv("idno"))
        If flagNoUpdate Then
            '待已轉知作業完成，則請遮罩身分證字號(變成模糊顯示)
            LabIDNO2.Text = Convert.ToString(drv("MKIDNO"))
        End If

        If Not flgROLEIDx0xLIDx0 AndAlso flagNoUpdate Then
            vMsg = "已轉知資訊不可取消!"
            STATUSNCREASON.Enabled = False
            TIMS.Tooltip(STATUSNCREASON, vMsg)
            STATUSNC1.Enabled = False
            TIMS.Tooltip(STATUSNC1, vMsg)
            STATUSNC2.Enabled = False
            TIMS.Tooltip(STATUSNC2, vMsg)
        End If

        If flgROLEIDx0xLIDx0 Then
            vMsg = "該使用者可以變更資料!!"
            STATUSNCREASON.Enabled = True
            STATUSNC1.Enabled = True
            STATUSNC2.Enabled = True
            TIMS.Tooltip(STATUSNCREASON, vMsg)
            TIMS.Tooltip(STATUSNC1, vMsg)
            TIMS.Tooltip(STATUSNC2, vMsg)
        End If

        '檢視內顯示最新一筆勞保加保證號及投保單位名稱
        Dim strVIEW1 As String = ""
        If Convert.ToString(drv("ACTNO")) <> "" Then
            strVIEW1 &= "勞保證號：" & CStr(drv("ACTNO")) & "\n"
        End If
        If Convert.ToString(drv("COMNAME")) <> "" Then
            strVIEW1 &= "投保單位：" & CStr(drv("COMNAME")) & "\n"
        End If
        If strVIEW1 = "" Then strVIEW1 &= "查無資料!!"
        BtnVIEW1.Attributes("onclick") = "alert('" & strVIEW1 & "');return false;"

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
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
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

    '檢核與儲存。 UPDATE STUD_SELRESULTBLIDET
    Sub SaveData1()
        '2006/03/ add conn by matt 'aNow = TIMS.GetSysDateNow(objconn)
        'CHECK
        Dim ERRMSG As String = ""
        Const cst_姓名 As Integer = 1          ' Cells(cst_姓名)
        'Const cst_報名機構 As Integer = 2
        'Const cst_報名班級 As Integer = 3
        'Const cst_報名日期 As Integer = 4      'Cells(5) = Cells(cst_報名日期)
        Dim i As Integer = 0
        Dim iSTATUSNC2_2 As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            i += 1
            Dim STATUSNC1 As RadioButtonList = eItem.FindControl("STATUSNC1")
            Dim STATUSNC2 As RadioButtonList = eItem.FindControl("STATUSNC2")
            Dim STATUSNCREASON As TextBox = eItem.FindControl("STATUSNCREASON") '已轉知 處理說明
            Dim Hid_STATUSPT As HiddenField = eItem.FindControl("Hid_STATUSPT")
            Dim Hid_SBDID As HiddenField = eItem.FindControl("Hid_SBDID")
            Dim Hid_SB3ID As HiddenField = eItem.FindControl("Hid_SB3ID")
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")

            STATUSNCREASON.Text = TIMS.ClearSQM(STATUSNCREASON.Text)
            Select Case STATUSNC1.SelectedValue
                Case "1", "2", "3", "4"
                    If STATUSNC1.Enabled Then
                        If STATUSNCREASON.Text = "" Then
                            ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ") 已轉知,請輸入處理說明。"
                        End If
                    End If
            End Select
            If ERRMSG = "" Then

                Select Case STATUSNC2.SelectedValue
                    Case "1", "2"
                        If STATUSNC2.Enabled Then
                            If STATUSNCREASON.Text = "" Then
                                ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ") 已轉知,請輸入處理說明。"
                            End If
                        End If
                End Select
                Select Case STATUSNC2.SelectedValue
                    Case "2" '2:已做離退訓
                        If ERRMSG = "" AndAlso STATUSNC2.Enabled Then
                            iSTATUSNC2_2 += 1
                        End If
                End Select

            End If

            If STATUSNCREASON.Text.Length > 100 Then
                STATUSNCREASON.Text = Microsoft.VisualBasic.Left(STATUSNCREASON.Text, 100)
            End If
        Next
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If

        Dim sql As String = ""
        'STUD_SELRESULTBLIDET
        sql = "" & vbCrLf
        sql += " UPDATE STUD_SELRESULTBLIDET" & vbCrLf
        sql += " SET STATUSNC1=@STATUSNC1 " & vbCrLf '已轉知1
        sql += " ,STATUSNC2=@STATUSNC2" & vbCrLf  '已轉知2
        sql += " ,STATUSNCREASON=@STATUSNCREASON " & vbCrLf '處理說明
        sql += " ,STATUSNCMACCT=@STATUSNCMACCT" & vbCrLf
        sql += " ,STATUSNCMDATE= getdate()" & vbCrLf
        sql += " WHERE SBDID =@SBDID " & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)
        'sql += " AND SB3ID =@SB3ID " & vbCrLf
        'sql += " AND IDNO =@IDNO " & vbCrLf
        'sql += " AND OCID =@OCID " & vbCrLf
        'sql += " AND STATUSPT =@STATUSPT " & vbCrLf

        Dim iUPD As Integer = 0
        'SAVE
        Call TIMS.OpenDbConn(objconn)
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim STATUSNC1 As RadioButtonList = eItem.FindControl("STATUSNC1")
            Dim STATUSNC2 As RadioButtonList = eItem.FindControl("STATUSNC2")
            Dim STATUSNCREASON As TextBox = eItem.FindControl("STATUSNCREASON") '已轉知 處理說明
            Dim Hid_STATUSPT As HiddenField = eItem.FindControl("Hid_STATUSPT")
            Dim Hid_SBDID As HiddenField = eItem.FindControl("Hid_SBDID")
            Dim Hid_SB3ID As HiddenField = eItem.FindControl("Hid_SB3ID")
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")

            Select Case STATUSNC1.SelectedValue
                Case "1", "2", "3", "4"
                    If STATUSNC1.Enabled Then
                        With uCmd
                            .Parameters.Clear()
                            .Parameters.Add("STATUSNC1", SqlDbType.VarChar).Value = STATUSNC1.SelectedValue '取消提醒1
                            .Parameters.Add("STATUSNC2", SqlDbType.VarChar).Value = TIMS.GetValue1(STATUSNC2.SelectedValue) '取消提醒2
                            .Parameters.Add("STATUSNCREASON", SqlDbType.NVarChar).Value = STATUSNCREASON.Text
                            .Parameters.Add("STATUSNCMACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            .Parameters.Add("SBDID", SqlDbType.Int).Value = Val(Hid_SBDID.Value)
                            .ExecuteNonQuery()
                        End With
                        iUPD += 1
                    End If
                Case Else
                    Select Case STATUSNC2.SelectedValue
                        Case "1", "2"
                            If STATUSNC2.Enabled Then
                                With uCmd
                                    .Parameters.Clear()
                                    .Parameters.Add("STATUSNC1", SqlDbType.VarChar).Value = Convert.DBNull '取消提醒1
                                    .Parameters.Add("STATUSNC2", SqlDbType.VarChar).Value = STATUSNC2.SelectedValue '取消提醒2
                                    .Parameters.Add("STATUSNCREASON", SqlDbType.NVarChar).Value = STATUSNCREASON.Text
                                    .Parameters.Add("STATUSNCMACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                    .Parameters.Add("SBDID", SqlDbType.Int).Value = Val(Hid_SBDID.Value)
                                    .ExecuteNonQuery()
                                End With
                                iUPD += 1
                            End If
                    End Select
            End Select
        Next


        If iSTATUSNC2_2 > 0 Then
            Common.MessageBox(Me, cst_saveMsgST2_2)
        End If
        If iUPD = 0 Then
            Common.MessageBox(Me, "無異動資料!")
            Exit Sub
        End If
        If iUPD > 0 Then
            Common.MessageBox(Me, "儲存成功!")
            Exit Sub
        End If
    End Sub

End Class