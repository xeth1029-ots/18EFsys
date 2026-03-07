Partial Class SD_02_017
    Inherits AuthBasePage

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            msg.Text = ""
            Table4.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            '查詢檢查。
            BtnSearch1.Attributes("onclick") = "javascript:return search()"
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
    Sub SUtl_Search1()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "未選擇有效機構")
            Exit Sub
        End If
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "未選擇有效班級")
            Exit Sub
        End If

        Dim pms1 As New Hashtable From {
            {"TPlanID", sm.UserInfo.TPlanID},
            {"Years", sm.UserInfo.Years},
            {"OCID", CInt(OCIDValue1.Value)}
        }
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.OCID,cc.Years,cc.PlanName,cc.DistName,cc.OrgName,cc.ClassCName" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " WHERE cc.TPlanID=@TPlanID and cc.Years=@Years" & vbCrLf
        If sm.UserInfo.LID > 0 Then
            pms1.Add("PlanID", sm.UserInfo.PlanID)
            pms1.Add("RID", RIDValue.Value)
            sql &= " and cc.PlanID=@PlanID and cc.RID=@RID" & vbCrLf
        End If
        sql &= " and cc.OCID=@OCID )" & vbCrLf

        sql &= " SELECT sd.SDID,sd.ADID,sd.OCID,sd.IDNO,sd.ALRINFORM,sd.ALRMACCT,sd.ALRMDATE" & vbCrLf
        sql &= " ,dbo.NVL(t1.Name,t2.Name) Name,t1.SETID,t2.eSETID,cc.Years,cc.PlanName,cc.DistName,cc.OrgName,cc.ClassCName" & vbCrLf
        sql &= " FROM STUD_DISASTER sd" & vbCrLf
        sql &= " JOIN WC1 cc on cc.OCID =sd.OCID" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERTEMP t1 on t1.IDNO=sd.IDNO AND t1.SETID=sd.SETID" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERTEMP2 t2 on t2.IDNO=sd.IDNO AND t2.eSETID=sd.eSETID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)

        msg.Text = "查無資料!"
        Table4.Visible = False
        If TIMS.dtNODATA(dt) Then Return

        msg.Text = ""
        Table4.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header, ListItemType.Footer
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號

                Dim hid_OCID1 As HiddenField = e.Item.FindControl("hid_OCID1")
                Dim hid_IDNO As HiddenField = e.Item.FindControl("hid_IDNO")
                Dim hid_SETID As HiddenField = e.Item.FindControl("hid_SETID")
                Dim hid_eSETID As HiddenField = e.Item.FindControl("hid_eSETID")
                Dim CheckBox1 As CheckBox = e.Item.FindControl("CheckBox1")

                hid_OCID1.Value = Convert.ToString(drv("OCID"))
                hid_IDNO.Value = Convert.ToString(drv("IDNO"))
                hid_SETID.Value = Convert.ToString(drv("SETID"))
                hid_eSETID.Value = Convert.ToString(drv("eSETID"))

                CheckBox1.Checked = False
                If Convert.ToString(drv("ALRINFORM")) = "Y" Then '已轉知
                    CheckBox1.Checked = True
                    CheckBox1.Enabled = False
                End If

        End Select
    End Sub

    '儲存。
    Sub SaveData1()
        '2006/03/ add conn by matt 'aNow = TIMS.GetSysDateNow(objconn)
        Call TIMS.OpenDbConn(objconn)
        'CHECK
        Dim sERRMSG As String = ""
        'sERRMSG = ""
        'Const cst_姓名 As Integer = 1          ' Cells(cst_姓名)
        'Const cst_報名機構 As Integer = 2
        'Const cst_報名班級 As Integer = 3
        'Const cst_報名日期 As Integer = 4      'Cells(5) = Cells(cst_報名日期)

        Dim i As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")
            Dim hid_eSETID As HiddenField = eItem.FindControl("hid_eSETID")
            Dim CheckBox1 As CheckBox = eItem.FindControl("CheckBox1") '已轉知  

            If CheckBox1.Enabled AndAlso CheckBox1.Checked Then
                i += 1
                Exit For
            End If
        Next
        If i = 0 Then sERRMSG &= "未勾選，任一項目!"
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Exit Sub
        End If

        Dim sql As String = ""
        sql &= " UPDATE STUD_DISASTER" & vbCrLf
        sql &= " SET ALRINFORM='Y',ALRMACCT=@ALRMACCT ,ALRMDATE=GETDATE()" & vbCrLf
        sql &= " where OCID=@OCID and IDNO=@IDNO" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        'SAVE
        'Call TIMS.OpenDbConn(objconn)
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim hid_OCID1 As HiddenField = eItem.FindControl("hid_OCID1")
            Dim hid_IDNO As HiddenField = eItem.FindControl("hid_IDNO")
            Dim hid_SETID As HiddenField = eItem.FindControl("hid_SETID")
            Dim hid_eSETID As HiddenField = eItem.FindControl("hid_eSETID")
            Dim CheckBox1 As CheckBox = eItem.FindControl("CheckBox1") '已轉知  

            If CheckBox1.Enabled AndAlso CheckBox1.Checked Then
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("ALRMACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("OCID", SqlDbType.Int).Value = Val(hid_OCID1.Value)
                    .Parameters.Add("IDNO", SqlDbType.VarChar).Value = hid_IDNO.Value
                    .ExecuteNonQuery()
                End With
            End If
        Next

        Common.MessageBox(Me, "儲存成功!")

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

    ''查詢Sql
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSearch1.Click
        Call SUtl_Search1()
    End Sub

    '儲存。
    Protected Sub BtnSave1_Click(sender As Object, e As EventArgs) Handles btnSave1.Click
        Call SaveData1()
    End Sub

End Class