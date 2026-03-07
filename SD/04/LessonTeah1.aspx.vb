Partial Class LessonTeah1
    Inherits AuthBasePage

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'tc_03_006 產投 Addx產投任課教師 Addy產投助教
        Me.modifytype.Value = Me.Request("type") 'Add'Edit / 'Add2 / 'Addx(產投任課教師) 'Addy(產投助教)
        Me.fieldname.Value = Me.Request("fieldname")
        Me.hiddenname.Value = Me.Request("hiddenname")
        Me.ExistTech.Value = Me.Request("ExistTech") '排除1名
        'Request("RID") 

        If Not Page.IsPostBack Then
            Call Search_Query("")
        End If

        txtSearch1.Attributes("onkeypress") = "Search_click();return false;"
        Me.btnSearch.Attributes("onclick") = "Search_click2();return false;"
        Me.btnSearch.Attributes("onkeypress") = "Search_click2();return false;"
    End Sub

#Region "Function"
    Sub Search_Query(ByVal TeachCNameVal As String)
        Dim rqRID As String = TIMS.ClearSQM(Request("RID"))
        Dim sql As String = ""
        sql &= " select tt.TechID ,tt.TeachCName ,tt.KindEngage" & vbCrLf
        sql &= " ,te.TechID TechIDEmp " & vbCrLf
        sql &= " from Teach_TeacherInfo tt   " & vbCrLf
        sql &= " left join (" & vbCrLf
        sql &= " 	select DISTINCT TechID from Teacher_Employs where TEDate2 >= getdate()" & vbCrLf '未到結束簽約使用日期。.內(沒簽約)
        sql &= " ) te on tt.TechID=te.TechID " & vbCrLf
        sql &= " WHERE tt.WorkStatus='1'" & vbCrLf
        '產投任課教師 或 '產投助教
        'If sqlTechType <> "" Then 
        '    sql += sqlTechType & vbCrLf
        'End If
        'Dim sqlTechType As String=""
        '啟動年度2013
        If sm.UserInfo.Years >= 2013 Then
            Select Case Me.modifytype.Value
                Case "Addx" '產投任課教師
                    sql &= " AND tt.TechType1='Y'" '& vbCrLf
                Case "Addy" '產投助教
                    sql &= " AND tt.TechType2='Y'" '& vbCrLf
            End Select
        End If

        If rqRID = "" Then 'RID (業務權限)
            sql &= " AND tt.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        Else
            sql &= " AND tt.RID='" & rqRID & "'" & vbCrLf
        End If
        '排除1名
        If Me.ExistTech.Value <> "" Then
            sql &= " and tt.TechID <> " & Val(Me.ExistTech.Value) & vbCrLf
        End If
        '搜尋名稱
        TeachCNameVal = TIMS.ClearSQM(TeachCNameVal)
        If TeachCNameVal <> "" Then sql &= String.Format(" and tt.TeachCName like N'%{0}%'", TeachCNameVal) & vbCrLf
        'sql &= " ORDER BY tt.TeachCName" & vbCrLf

        Dim objtable As DataTable = DbAccess.GetDataTable(sql, objconn)
        'objtable2=DbAccess.GetDataTable(objstr2, objconn)

        Dim mydv1 As New DataView(objtable) 'KindEngage:(1.內
        Dim mydv2 As New DataView(objtable) 'KindEngage:(2.外
        Dim mydv3 As New DataView(objtable)
        'KindEngage:(1.內;2.外)
        mydv1.RowFilter = "KindEngage='1'" '1.內
        mydv2.RowFilter = "KindEngage='2' and TechIDEmp is NOT NULL" '1.內2.外 (有簽約)'有在外聘師資管理功能建資料且聘約迄日尚未到期的老師
        mydv3.RowFilter = "KindEngage='2' and TechIDEmp is NULL" '1.內2.外(沒簽約)

        mydv1.Sort = "TeachCName,TechID"
        mydv2.Sort = "TeachCName,TechID"
        mydv3.Sort = "TeachCName,TechID"

        Me.LessonTeahList1.DataSource = mydv1
        Me.LessonTeahList1.DataTextField = "TeachCName"
        Me.LessonTeahList1.DataValueField = "TechID"
        Me.LessonTeahList1.DataBind()

        Me.LessonTeahList2.DataSource = mydv2
        Me.LessonTeahList2.DataTextField = "TeachCName"
        Me.LessonTeahList2.DataValueField = "TechID"
        Me.LessonTeahList2.DataBind()

        Me.LessonTeahList3.DataSource = mydv3
        Me.LessonTeahList3.DataTextField = "TeachCName"
        Me.LessonTeahList3.DataValueField = "TechID"
        Me.LessonTeahList3.DataBind()

    End Sub

    Sub sUtl_ReVal(ByVal MyPage As Page, ByRef objTechList As Object, ByVal MType As String)
        Dim oTechList As RadioButtonList = CType(objTechList, RadioButtonList)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">" + vbCrLf
        Select Case MType 'Me.modifytype.Value
            Case "Add", "Edit", "Addx", "Addy"
                strScript += "returnValue('" & oTechList.SelectedValue & "','" & oTechList.Items(oTechList.SelectedIndex).Text & "');" + vbCrLf
            Case "Add2"
                strScript += "returnValue2('" & oTechList.SelectedValue & "','" & oTechList.Items(oTechList.SelectedIndex).Text & "');" + vbCrLf
            Case Else
                strScript += "returnValue('" & oTechList.SelectedValue & "','" & oTechList.Items(oTechList.SelectedIndex).Text & "');" + vbCrLf
        End Select
        strScript += "</script>"
        MyPage.RegisterStartupScript("", strScript)
    End Sub

    Sub sUtl_ReClsVal(ByVal MyPage As Page, ByVal MType As String)
        Dim strScript As String
        strScript = "<script language=""javascript"">" + vbCrLf
        Select Case MType 'Me.modifytype.Value
            Case "Add", "Edit", "Addx", "Addy"
                strScript += "returnValue('','');" + vbCrLf
            Case "Add2"
                strScript += "returnValue2('','');" + vbCrLf
            Case Else
                strScript += "returnValue('','');" + vbCrLf
        End Select
        strScript += "</script>"

        MyPage.RegisterStartupScript("", strScript)
    End Sub

#End Region

    Private Sub LessonTeahList1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList1.SelectedIndexChanged
        'LessonTeahList1
        Call sUtl_ReVal(Me, sender, Me.modifytype.Value)
    End Sub

    Private Sub LessonTeahList2_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LessonTeahList2.SelectedIndexChanged
        'LessonTeahList2
        Call sUtl_ReVal(Me, sender, Me.modifytype.Value)
    End Sub

    Private Sub LessonTeahList3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LessonTeahList3.SelectedIndexChanged
        'LessonTeahList3
        Call sUtl_ReVal(Me, sender, Me.modifytype.Value)
    End Sub

    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        'Clear
        Call sUtl_ReClsVal(Me, Me.modifytype.Value)
    End Sub

    Private Sub hbtnSearch_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles hbtnSearch.ServerClick
        Dim InputVal As String = ""
        'Me.txtSearch1.Text=Trim(Me.txtSearch1.Text)
        Me.txtSearch1.Text = TIMS.ClearSQM(Me.txtSearch1.Text)
        InputVal = Me.txtSearch1.Text
        InputVal = TIMS.ClearSQM(InputVal)
        If InputVal <> "" Then
            If TIMS.CheckInput(InputVal) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        End If

        If InputVal <> "" Then
            Call Search_Query(InputVal)
        Else
            Call Search_Query("")
        End If

    End Sub

End Class
