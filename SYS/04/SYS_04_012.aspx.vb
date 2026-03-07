Partial Class SYS_04_012
    Inherits System.Web.UI.Page


    'SELECT 'UPDATE Key_Plan'
    '+' set PropertyID='+ isnull(convert(varchar,PropertyID),'Null')+ ''
    '+' WHERE TPlanID='''+ TPlanID+ '''' xStr FROM Key_Plan

    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            Call Create()
        End If
    End Sub

    Function Create()
        Dim sql As String
        Dim dt As DataTable

        sql = "" & vbCrLf
        sql += " SELECT TPlanID" & vbCrLf
        sql += " ,'('+TPlanID+')'+PlanName PlanName " & vbCrLf
        sql += " ,PropertyID" & vbCrLf
        sql += " FROM Key_Plan  " & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "TPlanID"
            DataGrid1.DataBind()
        Else
            Common.MessageBox(Me, "資料異常，請連絡系統管理者!!")
            Exit Function
        End If
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "SYS_TD1"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim PropertyID0 As HtmlInputRadioButton = e.Item.FindControl("PropertyID0")
                Dim PropertyID1 As HtmlInputRadioButton = e.Item.FindControl("PropertyID1")
                Dim PropertyIDX As HtmlInputRadioButton = e.Item.FindControl("PropertyIDX")

                Dim drv As DataRowView = e.Item.DataItem

                e.Item.CssClass = "SYS_TD2"

                'null:其他(停用)
                '0:職前
                '1:在職
                Select Case Convert.ToString(drv("PropertyID"))
                    Case "0"
                        PropertyID0.Checked = True
                    Case "1"
                        PropertyID1.Checked = True
                    Case Else
                        PropertyIDX.Checked = True
                End Select
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim da As OracleDataAdapter = Nothing
        Dim dr As DataRow
        'Dim conn As OracleConnection
        '2006/03/28 add conn by matt
        'conn = DbAccess.GetConnection

        sql = "SELECT * FROM Key_Plan"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim PropertyID0 As HtmlInputRadioButton = eItem.FindControl("PropertyID0")
            Dim PropertyID1 As HtmlInputRadioButton = eItem.FindControl("PropertyID1")
            Dim PropertyIDX As HtmlInputRadioButton = eItem.FindControl("PropertyIDX")
            Dim PropertyID As Integer = -1
            'null:其他(停用)
            '0:職前
            '1:在職
            PropertyID = -1 'null:其他(停用)
            Select Case True
                Case PropertyID0.Checked
                    PropertyID = PropertyID0.Value
                Case PropertyID1.Checked
                    PropertyID = PropertyID1.Value
            End Select

            dr = dt.Select("TPlanID='" & DataGrid1.DataKeys(eItem.ItemIndex) & "'")(0)
            Select Case PropertyID
                Case -1
                    dr("PropertyID") = Convert.DBNull
                Case Else
                    dr("PropertyID") = PropertyID
            End Select
            'dr("ModifyAcct") = sm.UserInfo.UserID
            'dr("ModifyDate") = Now
        Next
        DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功")
    End Sub

End Class
