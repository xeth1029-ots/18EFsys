Partial Class CP_04_002
    Inherits AuthBasePage

    'CP_04_002_add.aspx
    'CP_04_002_add_01.aspx
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

        '檢查日期格式
        Me.SSTDate.Attributes("onchange") = "check_date();"
        Me.ESTDate.Attributes("onchange") = "check_date();"

        If Not Page.IsPostBack Then
            Dim dt As DataTable
            'Dim dr As DataRow
            Dim sqlstr As String

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            yearlist_SelectedIndexChanged(sender, e)

            sqlstr = "SELECT Name,DistID FROM ID_District ORDER BY DistID"
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            Me.DistrictList.DataSource = dt
            Me.DistrictList.DataTextField = "Name"
            Me.DistrictList.DataValueField = "DistID"
            Me.DistrictList.DataBind()
            Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

            '計畫
            PlanList = TIMS.Get_TPlan(PlanList, , 1, "Y")
            '縣市
            CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing)

        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '選擇全部縣市
        Me.CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');"

        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"

        ' ShowConUnit.Attributes("onclick") = "ShowUnit();return false;"

    End Sub

    '明細查詢
    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        '選擇轄區
        Dim objitem As ListItem
        Dim itemstr As String = ""
        For Each objitem In Me.DistrictList.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value.ToString & "'"
            End If
        Next

        '選擇縣市
        Dim itemcity As String = ""
        For Each objitem In Me.CityList.Items
            If objitem.Selected = True Then
                If itemcity <> "" Then itemcity += ","
                itemcity += "'" & objitem.Value.ToString & "'"
            End If
        Next

        '選擇訓練計畫
        Dim itemplan As String = ""
        For Each objitem In Me.PlanList.Items
            If objitem.Selected = True Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next

        Dim ConRID As String = ""
        ConRID = ""
        For i As Integer = 1 To 7
            Dim chkList1 As CheckBoxList
            chkList1 = Me.FindControl("IsConUnit" & CStr(i))
            If Not chkList1 Is Nothing Then
                For Each objitem In chkList1.Items
                    If objitem.Selected = True Then
                        If ConRID <> "" Then ConRID += ","
                        ConRID += objitem.Value
                    End If
                Next
            End If
        Next

        'For Each objitem In IsConUnit2.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next
        'For Each objitem In IsConUnit3.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next
        'For Each objitem In IsConUnit4.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next
        'For Each objitem In IsConUnit5.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next
        'For Each objitem In IsConUnit6.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next
        'For Each objitem In IsConUnit7.Items
        '    If objitem.Selected = True Then
        '        If ConRID = "" Then
        '            ConRID = objitem.Value
        '        Else
        '            ConRID += "," & objitem.Value
        '        End If
        '    End If
        'Next

        Session("itemstr") = itemstr
        Session("itemplan") = itemplan
        Session("itemcity") = itemcity
        Session("ConRID") = ConRID
        Session("SSTDate") = Me.SSTDate.Text
        Session("ESTDate") = Me.ESTDate.Text
        'Response.Redirect("CP_04_002_add.aspx?yearlist=" & Me.yearlist.SelectedValue)
        Dim url1 As String = "CP_04_002_add.aspx?ID=" & Request("ID") & "&yearlist=" & Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub bt_reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_reset.Click
        'Reset
        Me.yearlist.SelectedIndex = 1

        Dim i As Short
        For i = 0 To Me.DistrictList.Items.Count - 1
            Me.DistrictList.Items(i).Selected = False
        Next

        For i = 0 To Me.CityList.Items.Count - 1
            Me.CityList.Items(i).Selected = False
        Next

        For i = 0 To Me.PlanList.Items.Count - 1
            Me.PlanList.Items(i).Selected = False
        Next

        Me.SSTDate.Text = ""
        Me.ESTDate.Text = ""

    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        Dim dt As DataTable
        Dim dv As New DataView

        Dist0.Style("display") = "none"
        Dist1.Style("display") = "none"
        Dist2.Style("display") = "none"
        Dist3.Style("display") = "none"
        Dist4.Style("display") = "none"
        Dist5.Style("display") = "none"
        Dist6.Style("display") = "none"

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.OrgName+'('+d.PlanName+')' OrgName " & vbCrLf
        sql += " ,b.RID " & vbCrLf
        sql += " ,c.DistID " & vbCrLf
        sql += " FROM Org_OrgInfo a " & vbCrLf
        sql += " JOIN Auth_Relship b ON a.OrgID=b.OrgID " & vbCrLf
        sql += " JOIN ID_Plan c ON b.PlanID=c.PlanID  " & vbCrLf
        sql += " JOIN view_LoginPlan d ON c.PlanID=d.PlanID " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += " and a.IsConUnit=1" & vbCrLf
        sql += " and c.Years='" & yearlist.SelectedValue & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        dt.TableName = "UnitName"
        dv.Table = dt

        If dt.Rows.Count = 0 Then
            msg.Text = "查無" & yearlist.SelectedValue & "年的管控單位"
        Else
            msg.Text = ""
            dv.RowFilter = "DistID='000'"
            With IsConUnit1
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='001'"
            With IsConUnit2
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='002'"
            With IsConUnit3
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='003'"
            With IsConUnit4
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='004'"
            With IsConUnit5
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='005'"
            With IsConUnit6
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With

            dv.RowFilter = "DistID='006'"
            With IsConUnit7
                .DataSource = dv
                .DataTextField = "OrgName"
                .DataValueField = "RID"
                .DataBind()
                .Visible = True
            End With
        End If
    End Sub

    Private Sub bt_search1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search1.Click
        '選擇轄區
        'Dim objitem As ListItem
        'Dim itemstr As String
        'Dim DistID, DistName, ICity, TPlanID, ConRID1 As String
        'Dim newDistID, newICity, newTPlanID, newConRID As String
        'Dim newDistName, ICityName, newICityName, TPlanName, newTPlanIDName As String
        'Dim ConRID1_Name, newConRIDName As String
        'Dim i As Integer



        Dim itemstr As String = "" '選擇轄區
        itemstr = "" '選擇轄區
        For Each objitem As ListItem In Me.DistrictList.Items
            If objitem.Selected = True Then
                If itemstr <> "" Then itemstr += ","
                itemstr += "'" & objitem.Value.ToString & "'"
            End If
        Next
        Dim itemcity As String '選擇縣市
        itemcity = "" '選擇縣市
        For Each objitem As ListItem In Me.CityList.Items
            If objitem.Selected = True Then
                If itemcity <> "" Then itemcity += ","
                itemcity += "'" & objitem.Value.ToString & "'"

            End If
        Next
        Dim itemplan As String '選擇訓練計畫
        itemplan = "" '選擇訓練計畫
        For Each objitem As ListItem In Me.PlanList.Items
            If objitem.Selected = True Then
                If itemplan <> "" Then itemplan += ","
                itemplan += "'" & objitem.Value.ToString & "'"
            End If
        Next

        '報表要用的轄區參數
        Dim DistID As String = ""
        Dim DistName As String = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                If DistID <> "" Then DistID &= ","
                DistID &= Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")
                If DistName <> "" Then DistName &= ","
                DistName &= Me.DistrictList.Items(i).Text
            End If
        Next
        If DistID <> "" Then
            '選擇全部
            If Me.DistrictList.Items(0).Selected Then
                DistName = "全部"
            End If
        End If

        '報表要用的縣市參數
        Dim ICity As String = ""
        Dim ICityName As String = ""
        ICity = ""
        ICityName = ""
        For i As Integer = 1 To Me.CityList.Items.Count - 1
            If Me.CityList.Items(i).Selected Then
                If ICity <> "" Then ICity &= ","
                ICity &= "\'" & Me.CityList.Items(i).Value & "\'"

                If ICityName <> "" Then ICityName &= ","
                ICityName &= Me.CityList.Items(i).Text
            End If
        Next
        If ICity <> "" Then
            If Me.CityList.Items(0).Selected Then
                ICityName = "全部"
            End If
        End If
        'If ICity <> "" Then
        '    newICity = Mid(ICity, 1, ICity.Length - 1)
        'End If

        'If ICity <> "" Then
        '    If Me.CityList.Items(0).Selected Then
        '        newICity = Mid(ICity, 1, ICity.Length - 1)
        '        newICityName = "全部"
        '    Else
        '        newICity = Mid(ICity, 1, ICity.Length - 1)
        '        newICityName = Mid(ICityName, 1, ICityName.Length - 1)
        '    End If
        'End If


        '報表要用的訓練計畫參數
        Dim TPlanID As String = ""
        Dim TPlanName As String = ""
        TPlanID = ""
        TPlanName = ""
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected Then
                If TPlanID <> "" Then TPlanID &= ","
                TPlanID &= "\'" & Me.PlanList.Items(i).Value & "\'"

                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Me.PlanList.Items(i).Text
            End If
        Next
        If TPlanID <> "" Then
            If Me.PlanList.Items(0).Selected Then
                TPlanName = "全部"
            End If
        End If


        'If TPlanID <> "" Then
        '    newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        'End If

        'If TPlanID <> "" Then
        '    If Me.PlanList.Items(0).Selected Then
        '        newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        '        newTPlanIDName = "全部"
        '    Else
        '        newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        '        newTPlanIDName = Mid(TPlanName, 1, TPlanName.Length - 1)
        '    End If
        'End If

        '選擇控管單位
        Dim ConRID As String = ""
        For Each objitem As ListItem In IsConUnit1.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next

        For Each objitem As ListItem In IsConUnit2.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next
        For Each objitem As ListItem In IsConUnit3.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next
        For Each objitem As ListItem In IsConUnit4.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next
        For Each objitem As ListItem In IsConUnit5.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next
        For Each objitem As ListItem In IsConUnit6.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next
        For Each objitem As ListItem In IsConUnit7.Items
            If objitem.Selected = True Then
                If ConRID <> "" Then ConRID += ","
                ConRID += objitem.Value
            End If
        Next

        '報表要用的控管單位參數
        Dim ConRID1 As String = ""
        Dim ConRID1_Name As String = ""
        ConRID1 = ""
        ConRID1_Name = ""
        For i As Integer = 0 To Me.IsConUnit1.Items.Count - 1
            If Me.IsConUnit1.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit2.Items.Count - 1
            If Me.IsConUnit2.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit3.Items.Count - 1
            If Me.IsConUnit3.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit4.Items.Count - 1
            If Me.IsConUnit4.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit5.Items.Count - 1
            If Me.IsConUnit5.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit6.Items.Count - 1
            If Me.IsConUnit6.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next
        For i As Integer = 0 To Me.IsConUnit7.Items.Count - 1
            If Me.IsConUnit7.Items(i).Selected Then
                If ConRID1 <> "" Then ConRID1 += ","
                ConRID1 += Convert.ToString("\'" & Me.IsConUnit1.Items(i).Value & "\'")

                If ConRID1_Name <> "" Then ConRID1_Name += ","
                ConRID1_Name += Me.IsConUnit1.Items(i).Text
            End If
        Next

        'If ConRID1 <> "" Then
        '    newConRID = Mid(ConRID1, 1, ConRID1.Length - 1)
        '    newConRIDName = Mid(ConRID1_Name, 1, ConRID1_Name.Length - 1)
        'End If

        Session("itemstr") = itemstr
        Session("itemplan") = itemplan
        Session("itemcity") = itemcity
        Session("ConRID") = ConRID
        Session("SSTDate") = Me.SSTDate.Text
        Session("ESTDate") = Me.ESTDate.Text

        Session("newDistID") = DistID 'newDistID
        Session("newICity") = ICity
        Session("newTPlanID") = TPlanID
        Session("newConRID") = ConRID1 ' newConRID

        Session("newDistName") = DistName 'newDistName
        Session("newICityName") = ICityName
        Session("newTPlanIDName") = TPlanName
        Session("newConRIDName") = ConRID1_Name 'newConRIDName

        'Response.Redirect("CP_04_002_add_01.aspx?yearlist=" & Me.yearlist.SelectedValue)
        Dim url1 As String = "CP_04_002_add_01.aspx?ID=" & Request("ID") & "&yearlist=" & Me.yearlist.SelectedValue
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class

