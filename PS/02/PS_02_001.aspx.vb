Partial Class PS_02_001
    Inherits AuthBasePage

    ''VISUALCHART PS_VISUALCHART
    Dim AllCount As Int32 = 0
    Dim ChkCount As Int32 = 0

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'objConn.Open()
        If Not IsPostBack Then Call LoadNeedData()
    End Sub

    Sub LoadNeedData()
        '依照計畫別代碼,取得該計畫對應的所有統計圖表
        Dim myDT1 As New DataTable
        Dim sql1 As String = ""
        sql1 &= " SELECT CRTID, CNAME, PIC, TYPEID, DEFAULTSHOW " & vbCrLf
        sql1 &= " FROM VISUALCHART " & vbCrLf
        sql1 &= " WHERE ISUSED = 'Y' AND TPLANID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
        sql1 &= "  ORDER BY SORT ASC " & vbCrLf
        myDT1 = DbAccess.GetDataTable(sql1, objconn)
        If myDT1.Rows.Count > 0 Then
            Me.ListView1.DataSource = myDT1
            AllCount = myDT1.Rows.Count
            Me.ListView1.DataBind()
            For Each item As ListViewItem In ListView1.Items
                Dim objTooltip As Label = item.FindControl("myTooltip")
                If objTooltip.ToolTip.Trim.Equals("N") Then
                    objTooltip.Text = ""
                Else
                    objTooltip.Text = "(預設)"
                End If
            Next

            '讀取目前使用者有勾選的項目內容
            Dim myDT2 As New DataTable
            Dim sql2 As String = ""
            sql2 &= " SELECT DISTINCT TOP(3) a.CRTID, b.CNAME, b.PIC, b.TYPEID, b.SORT " & vbCrLf
            sql2 &= " FROM PS_VISUALCHART a " & vbCrLf
            sql2 &= " JOIN VISUALCHART b ON a.CRTID = b.CRTID " & vbCrLf
            sql2 &= " WHERE b.ISUSED = 'Y' AND a.SHOW = 'Y' " & vbCrLf
            sql2 &= "    AND b.TPLANID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf  '06 在職計畫; 28 產投
            sql2 &= "    AND a.ACCOUNT = '" & sm.UserInfo.UserID & "' " & vbCrLf
            myDT2 = DbAccess.GetDataTable(sql2, objconn)
            If myDT2.Rows.Count > 0 Then
                Dim myList As New List(Of String)
                Dim i As Integer = 0
                For i = 0 To myDT2.Rows.Count - 1
                    myList.Add(myDT2.Rows(i).Item("CRTID"))
                Next

                For Each item As ListViewItem In ListView1.Items
                    Dim objLbl As Label = item.FindControl("CRTID")
                    Dim objChk As CheckBox = item.FindControl("CHECK")
                    If myList.IndexOf(objLbl.Text.Trim) > -1 Then objChk.Checked = True
                Next
            Else
                For Each item As ListViewItem In ListView1.Items
                    Dim objTooltip As Label = item.FindControl("myTooltip")
                    Dim objChk As CheckBox = item.FindControl("CHECK")
                    If objTooltip.ToolTip.Trim.Equals("Y") Then objChk.Checked = True
                Next
            End If
        Else
            Common.MessageBox(Me, "找不到計畫別的「範例圖表」!!")
        End If
    End Sub

    Private Sub btn_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_save.Click
        Dim mySql1 As String = ""
        mySql1 &= " DELETE a FROM PS_VISUALCHART a " & vbCrLf
        mySql1 &= " JOIN VISUALCHART b ON a.CRTID = b.CRTID " & vbCrLf
        mySql1 &= " WHERE b.TPLANID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
        mySql1 &= " AND a.ACCOUNT = '" & sm.UserInfo.UserID & "' " & vbCrLf
        DbAccess.ExecuteNonQuery(mySql1, objconn)

        Dim mySql2 As String = ""
        mySql2 &= " INSERT INTO PS_VISUALCHART (PSVID, ACCOUNT, CRTID, SHOW, SORT, MODIFYACCT, MODIFYDATE) " & vbCrLf
        mySql2 &= " VALUES (@PSVID, @ACCOUNT, @CRTID, @SHOW, @SORT, @ModifyAcct, GETDATE()) " & vbCrLf
        Dim newSort As Integer = 0
        For Each item As ListViewItem In ListView1.Items
            Dim objLbl As Label = item.FindControl("CRTID")
            Dim objChk As CheckBox = item.FindControl("CHECK")
            If objChk.Checked = True Then '被勾選
                Dim myParam As Hashtable = New Hashtable
                Dim PSVID As Integer = DbAccess.GetNewId(objconn, "PS_VISUALCHART_PSVID_SEQ,PS_VISUALCHART,PSVID")
                myParam.Add("PSVID", PSVID)
                myParam.Add("ACCOUNT", sm.UserInfo.UserID)
                myParam.Add("CRTID", CInt(objLbl.Text.Trim))
                myParam.Add("SHOW", "Y")
                newSort += 1
                myParam.Add("SORT", newSort)
                myParam.Add("ModifyAcct", Convert.ToString(sm.UserInfo.UserID))
                DbAccess.ExecuteNonQuery(mySql2, objconn, myParam)
            End If
        Next

        Common.MessageBox(Me, "儲存成功!!")
    End Sub
End Class