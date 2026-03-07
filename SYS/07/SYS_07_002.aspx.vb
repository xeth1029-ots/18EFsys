Partial Class SYS_07_002
    Inherits System.Web.UI.Page
 
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        bt_update.Attributes.Add("onclick", "return confirm('是否執行!?執行後原資料將不存在!');")
        bt_reserve.Attributes.Add("onclick", "return confirm('是否執行!?');")
        bt_restoreID.Attributes.Add("onclick", "return confirm('確定還原!');")
        bt_fixID.Attributes.Add("onclick", "return confirm('是否執行!?');")

        'If sm.UserInfo.UserID = "amuting" Then
        '    tr01.Visible = True
        'Else
        '    tr01.Visible = False
        'End If
    End Sub

    Private Sub bt_reserve_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reserve.Click
        run(1)
    End Sub
    Private Sub bt_update_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_update.Click
        run(2)
    End Sub
    Private Sub bt_restoreID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_restoreID.Click
        run(3)
    End Sub
    Private Sub bt_fixID_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_fixID.Click
        run(4)
    End Sub

    Sub run(ByVal mode As Int16)
        Dim myobj As New FixJoblessID
        Dim ocid As String = ""
        Try
            myobj.chkJoblessID(mode, ocid)
            Common.MessageBox(Me.Page, "資料更新完成!")
        Catch ex As Exception
            Common.MessageBox(Me.Page, "更新失敗,發生錯誤：" & ex.Message.ToString)
        End Try
    End Sub

End Class
