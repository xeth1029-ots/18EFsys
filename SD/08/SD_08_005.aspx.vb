Partial Class SD_08_005
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_08_005"

    'SD_08_005
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not Page.IsPostBack Then
            '代入年度
            ddlYear = TIMS.GetSyear(ddlYear)
            Common.SetListItem(ddlYear, sm.UserInfo.Years)

            '代入轄區
            cklDistID = TIMS.Get_DistID(cklDistID)
            cklDistID.Items.Insert(0, New ListItem("全部", ""))
            cklDistID.SelectedValue = sm.UserInfo.DistID

            '代入訓練計畫
            cklTPlanID = TIMS.Get_TPlan(cklTPlanID)
            cklTPlanID.Items.Insert(0, New ListItem("全部", ""))
            cklTPlanID.SelectedValue = sm.UserInfo.TPlanID

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PlanID.Value = sm.UserInfo.PlanID

            lsbOCID.Style("display") = "none"
            msg.Text = TIMS.cst_NODATAMsg11

            btnSchClass_Click(sender, e)

            '改成若是委訓單位登入下列欄位就不顯示
            If sm.UserInfo.LID = "2" Then
                Year_TR.Style("display") = "none"
                DistID_TR.Style("display") = "none"
                PlanID_TR.Style("display") = "none"
                Check_TR.Style("display") = "none"
                Button2.Style("display") = "none"
            Else
                If sm.UserInfo.DistID <> "000" Then
                    DistID_TR.Style("display") = "none"
                    PlanID_TR.Style("display") = "none"
                End If
            End If

            btnSchClass.Attributes.Add("style", "display:none")
            cklDistID.Attributes("onclick") = "SelectAll('cklDistID','hidDistID');"
            cklTPlanID.Attributes("onclick") = "SelectAll('cklTPlanID','hidTPlanID');"
            btnPrt.Attributes("onclick") = "return chkPrt();return false;"
            chkData.Attributes("onclick") = "Enabled_OCID('" & sm.UserInfo.OrgName & "','" & sm.UserInfo.RID & "','" & sm.UserInfo.PlanID & "');"
        End If

        Class_TR.Style("display") = "none"
        Org_TR.Style("display") = "none"
        If Not chkData.Checked Then
            Class_TR.Style("display") = TIMS.cst_inline1 '"inline"
            Org_TR.Style("display") = TIMS.cst_inline1 '"inline"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick
        Dim DistID1 As String = ""
        Dim TPlanID1 As String = ""
        Dim N As Integer = 0
        Dim N1 As Integer = 0
        Dim strMsg As String = ""
        Dim strScript1 As String = ""

        '判斷轄區(不可附選)
        For i As Integer = 1 To cklDistID.Items.Count - 1
            If cklDistID.Items(i).Selected Then
                N = N + 1

                If N = 1 Then
                    DistID1 = Convert.ToString(cklDistID.Items(i).Value) '取得選項的值
                Else
                    Exit For
                End If
            End If
        Next

        Select Case N
            Case 0
                strMsg += "請選擇轄區!" & vbCrLf
            Case 1
                strMsg += ""
            Case Else
                strMsg += "只能選擇一個轄區!" & vbCrLf
                DistID1 = ""
        End Select

        '判斷計畫(不可附選)
        For j As Integer = 1 To cklTPlanID.Items.Count - 1
            If cklTPlanID.Items(j).Selected Then
                N1 = N1 + 1

                If N1 = 1 Then
                    TPlanID1 = Convert.ToString(cklTPlanID.Items(j).Value) '取得選項的值
                Else
                    Exit For
                End If
            End If
        Next

        Select Case N1
            Case 0
                strMsg += "請選擇計畫!" & vbCrLf
            Case 1
                strMsg += ""
            Case Else
                strMsg += "只能選擇一個計畫!" & vbCrLf
                TPlanID1 = ""
        End Select

        If strMsg <> "" Then
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If


        center.Text = ""
        lsbOCID.Items.Clear()

        strScript1 = "<script language=""javascript"">" + vbCrLf
        strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "'  + '&BtnName=btnSchClass','查詢機構',400,400,1);"
        strScript1 += "</script>"
        Page.RegisterStartupScript("", strScript1)
    End Sub

    Private Sub btnSchClass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSchClass.Click
        'Dim conn As SqlConnection = DbAccess.GetConnection()
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        'Dim dr As DataRow = Nothing
        'Dim strSelected As String = ""
        'Dim relship As String = ""
        'Dim ClassName As String = ""
        msg.Text = ""

        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql &= " Select cc.ocid ,cc.classcname2 ClassName"
        sql &= " FROM VIEW2 cc"
        sql &= " where cc.planid='" & PlanID.Value & "'"
        sql &= " and cc.rid='" & RIDValue.Value & "' "
        sql &= " and cc.notopen='N' and cc.issuccess='Y'"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count = 0 Then
            msg.Text = "查無此機構底下的班級"
            lsbOCID.Style("display") = "none"
        End If

        msg.Text = ""
        TIMS.GET_OnlyOne_OCID(Me, txtTMID, hidTMID, txtOCID, hidOCID)

        lsbOCID.Items.Clear()
        lsbOCID.Items.Add(New ListItem("全選", "%"))

        Dim strSelected As String = ""
        For Each dr As DataRow In dt.Rows
            'Select Case sm.UserInfo.TPlanID
            '    Case "17" '補助地方政府訓練
            '        'ClassName = dr("OrgName").ToString & "_" & dr("ClassCName").ToString
            '    Case Else
            '        ClassName = dr("ClassCName").ToString
            'End Select
            'If IsNumeric(dr("CyclType")) Then
            '    If Int(dr("CyclType")) <> 0 Then
            '        ClassName += "第" & Int(dr("CyclType")) & "期"
            '    End If
            'End If

            lsbOCID.Items.Add(New ListItem(CStr(dr("ClassName")), dr("OCID")))
            If Convert.ToString(dr("OCID")) = hidOCID.Value Then
                strSelected = Convert.ToString(dr("OCID"))
            End If
        Next

        lsbOCID.Style("display") = TIMS.cst_inline1 '"inline"

        If strSelected.ToString <> "" Then
            lsbOCID.SelectedValue = strSelected
        End If
    End Sub

    Private Sub btnPrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrt.Click
        Dim strMsg As String = ""
        Dim strDistID As String = ""
        Dim strDistName As String = ""

        Dim strTPlanID As String = ""
        Dim strTPlanName As String = ""

        Dim strOCID As String = ""
        Dim strOCName As String = ""

        For i As Integer = 1 To cklDistID.Items.Count - 1
            If cklDistID.Items(i).Selected Then
                If strDistID <> "" Then strDistID += ","
                strDistID += Convert.ToString(cklDistID.Items(i).Value)
                If strDistName <> "" Then strDistName += ","
                strDistName += Convert.ToString(cklDistID.Items(i).Text)

                'If strDistID = "" Then
                '    strDistID = Convert.ToString(cklDistID.Items(i).Value)
                '    strDistName = Convert.ToString(cklDistID.Items(i).Text)
                'Else
                '    strDistID += "," & Convert.ToString(cklDistID.Items(i).Value)
                '    strDistName += ", " & Convert.ToString(cklDistID.Items(i).Text)
                'End If
            End If
        Next

        If strDistID = "" Then strMsg += "請選擇轄區!" & vbCrLf

        For i As Integer = 1 To cklTPlanID.Items.Count - 1
            If cklTPlanID.Items(i).Selected Then
                If strTPlanID <> "" Then strTPlanID += ","
                strTPlanID += Convert.ToString(cklTPlanID.Items(i).Value)
                If strTPlanName <> "" Then strTPlanName += ","
                strTPlanName += Convert.ToString(cklTPlanID.Items(i).Text)

                'If strTPlanID = "" Then
                '    strTPlanID = Convert.ToString(cklTPlanID.Items(i).Value)
                '    strTPlanName = Convert.ToString(cklTPlanID.Items(i).Text)
                'Else
                '    strTPlanID += "," & Convert.ToString(cklTPlanID.Items(i).Value)
                '    strTPlanName += ", " & Convert.ToString(cklTPlanID.Items(i).Text)
                'End If
            End If
        Next

        If strTPlanID = "" Then strMsg += "請選擇計畫!" & vbCrLf

        If chkData.Checked = False Then
            For i As Integer = 0 To lsbOCID.Items.Count - 1
                If lsbOCID.Items(i).Selected Then
                    If strOCID <> "" Then strOCID += ","
                    strOCID += Convert.ToString(lsbOCID.Items(i).Value)
                    If strOCName <> "" Then strOCName += ","
                    strOCName += Convert.ToString(lsbOCID.Items(i).Text)

                    'If strOCID = "" Then
                    '    strOCID = Convert.ToString(lsbOCID.Items(i).Value)
                    '    strOCName = Convert.ToString(lsbOCID.Items(i).Text)
                    'Else
                    '    strOCID += "," & Convert.ToString(lsbOCID.Items(i).Value)
                    '    strOCName += ", " & Convert.ToString(lsbOCID.Items(i).Text)
                    'End If
                End If
            Next

            If strOCID = "" Then strMsg += "請選擇班級!" & vbCrLf
        Else
            strOCID = ""
            strOCName = ""
        End If


        If strMsg <> "" Then
            Common.MessageBox(Me, strMsg)
            Exit Sub
        End If

        '年度、轄區ID、轄區Name、訓練計畫ID、訓練計畫Name、查詢範圍、訓練單位RID、訓練單位Name、班級ID、班級Name、開訓起日、開訓迄日、結訓起日、結訓迄日
        'Years,DistID,DistName,TPlanID,TPlanName,schFlag,RID,RName,OCID,OCName,SCSDate,ECSDate,STEDate,ETEDate
        Dim strMyValue As String = ""

        strMyValue += "&Years=" & ddlYear.SelectedValue
        strMyValue += "&DistID=" & strDistID & "&DistName=" & strDistName
        strMyValue += "&TPlanID=" & strTPlanID & "&TPlanName=" & strTPlanName

        If chkData.Checked = False Then
            strMyValue += "&schFlag=0"

            If lsbOCID.Items(0).Selected = True Then
                strMyValue += "&RID=" & RIDValue.Value & "&RName=" & center.Text
            Else
                strMyValue += "&RID=" & RIDValue.Value & "&RName=" & center.Text
                strMyValue += "&OCID=" & strOCID & "&OCName=" & strOCName
            End If
        Else
            strMyValue += "&schFlag=1"
        End If

        If txtSCSDate.Text <> "" Then
            strMyValue += "&SCSDate=" & txtSCSDate.Text
        End If

        If txtECSDate.Text <> "" Then
            strMyValue += "&ECSDate=" & txtECSDate.Text
        End If

        If txtSTEDate.Text <> "" Then
            strMyValue += "&STEDate=" & txtSTEDate.Text
        End If

        If txtETEDate.Text <> "" Then
            strMyValue += "&ETEDate=" & txtETEDate.Text
        End If

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, strMyValue)
    End Sub
End Class
