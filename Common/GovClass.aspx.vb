Partial Class GovClass
    Inherits AuthBasePage

    'SELECT * FROM V_GOVCLASSCAST2 ORDER BY GCODE2
    'SELECT * FROM ID_GOVCLASSCAST2 WHERE PARENTS IS NULL ORDER BY GCODE
    'SELECT * FROM ID_GOVCLASSCAST2 WHERE PARENTS IS NOT NULL ORDER BY PARENTS,GCODE
    Dim strYears As String = "" '2014 / 2015'(經費分類代碼。)
    Const cst_y2014 As String = "2014"
    Const cst_y2015 As String = "2015"
    Const cst_y2018 As String = "2018"
    Const cst_y2019 As String = "2019"

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start, Me.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        '(經費分類代碼。)
        strYears = cst_y2014 '"2014" '2014年  顯示層級。
        If sm.UserInfo.Years >= cst_y2015 Then strYears = cst_y2015 '2015年 不顯示層級。
        If sm.UserInfo.Years >= cst_y2018 Then strYears = cst_y2018 '2018年。
        If sm.UserInfo.Years >= cst_y2019 Then strYears = cst_y2019 '2019年。

        trRadio1.Visible = True '2014年 顯示層級。
        Select Case strYears
            Case cst_y2015 '"2015"
                trRadio1.Visible = False '2015年 不顯示層級。
            Case Is >= cst_y2018 '"2018"
                trRadio1.Visible = False '2018年 不顯示層級。
        End Select

        If Not IsPostBack Then
            errortable.Style("display") = "none" '嚴重錯誤警示
            Call cCreate1()
        End If
    End Sub

#Region "V_GOVCLASSCAST3"

    Function Get_DDLGCode3P(ByVal obj As DropDownList, ByRef conn As SqlConnection) As DropDownList
        Dim dt As New DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT GCODE31 ,'['+GCODE31+']'+CNAME CNAME " & vbCrLf
        sql &= " FROM V_GOVCLASSCAST3 " & vbCrLf
        sql &= " WHERE 1=1 AND PGCID3 IS NULL " & vbCrLf
        sql &= " ORDER BY GCODE31 " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        With obj
            .DataSource = dt
            .DataTextField = "CNAME"
            .DataValueField = "GCODE31"
            .DataBind()
            '.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

    Function Get_DDLGCode3C(ByVal obj As DropDownList, ByVal GCODE31 As String, ByRef conn As SqlConnection) As DropDownList
        Dim dt As New DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT GCID3,GCODE31 ,GCODE32 ,'['+GCODE2+']'+CNAME CNAME " & vbCrLf
        sql &= " FROM V_GOVCLASSCAST3 " & vbCrLf
        sql &= " WHERE 1=1 AND PGCID3 IS NOT NULL AND GCODE31 = @GCODE31 " & vbCrLf
        sql &= " ORDER BY GCODE32 " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("GCODE31", SqlDbType.VarChar).Value = GCODE31
            dt.Load(.ExecuteReader())
        End With
        With obj
            .DataSource = dt
            .DataTextField = "CNAME"
            .DataValueField = "GCID3"
            .DataBind()
            '.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        Return obj
    End Function

#End Region

    '2014年前經費分類代碼顯示。
    Sub cCreate1()
        fieldname.Value = TIMS.ClearSQM(Request("fieldname")) '回傳值。
        Dim rqtrainValue As String = TIMS.ClearSQM(Request("trainValue"))
        Dim rqJobValue As String = TIMS.ClearSQM(Request("jobValue"))
        Dim rqPointYN As String = TIMS.ClearSQM(Request("PointYN"))

        Select Case strYears
            Case Is >= cst_y2018 '"2018"
                If rqtrainValue = "" Then Exit Sub
                Dim sql As String = ""
                sql = " SELECT GCID3 ,GCODE31 FROM V_GOVCLASSCAST3 WHERE TMID = " & rqtrainValue
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    GCode1 = Get_DDLGCode3P(GCode1, objconn)
                    Common.SetListItem(GCode1, dr("GCODE31"))
                    GCode2 = Get_DDLGCode3C(GCode2, dr("GCODE31"), objconn)
                    Common.SetListItem(GCode2, dr("GCID3"))
                End If
            Case cst_y2015 '"2015"
                GCode1 = TIMS.GetGCode2B(GCode1, "", objconn)
                '若有帶 jobValue 值，自動選擇功能如下：
                If Convert.ToString(rqJobValue) <> "" Then
                    Dim sql As String = ""
                    sql = "" & vbCrLf
                    sql += " SELECT ke.GCID2 " & vbCrLf '代碼
                    sql += " ,ig2.GCODE " & vbCrLf '順序編號(父層)
                    sql += " FROM KEY_TRAINTYPE ke " & vbCrLf
                    sql += " JOIN ID_GOVCLASSCAST2 ig2 ON ke.GCID2 = ig2.GCID2 " & vbCrLf
                    sql += " WHERE ke.TMID = " & Convert.ToString(rqJobValue) & vbCrLf
                    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

                    If Not dr Is Nothing Then
                        If Convert.ToString(dr("GCODE")) <> "" Then
                            Common.SetListItem(GCode1, dr("GCODE"))
                            GCode2 = TIMS.GetGCode2B(GCode2, dr("GCODE"), objconn)
                        End If
                    Else
                        '訓練業別超過預期資料，嚴重錯誤警示
                        normaltable.Style("display") = "none"
                        'errortable.Style("display") = "inline"
                        errortable.Style("display") = ""
                        'Common.RespWrite(Me, "<script>ClearValue();alert('資料異常，請重新選擇訓練業別!!!');window.close();</script>")
                        'Exit Sub
                    End If
                End If
                '非 學分班不可選擇 ;學分班 可選擇：
                'Radio1.Enabled = True
                GCode1.Enabled = True
                If Convert.ToString(rqPointYN) <> "" Then
                    If Convert.ToString(rqPointYN) = "N" Then
                        'Radio1.Enabled = False
                        GCode1.Enabled = False
                    End If
                End If
            Case cst_y2014 '"2014"
                '2014年前經費分類代碼顯示。
                '新增語法如下：
                'INSERT INTO ID_GovClassCast SELECT 101,101,NULL,'3C共通核心職能課程',NULL,NULL
                '補充資料庫 Radio1 新值如下：
                'Dim sql As String = ""
                'sql = ""
                'sql += " SELECT GCode1 as Value,CNAME "
                'sql += " FROM ID_GovClassCast "
                'sql += " WHERE GCode2 is null and GovClass >=99 "
                'sql += " ORDER BY GCode1"
                'Dim dt As DataTable = DbAccess.GetDataTable(sql)
                'If dt.Rows.Count > 0 Then
                '    For i As Integer = 0 To dt.Rows.Count - 1
                '        Radio1.Items.Add(New ListItem(dt.Rows(i).Item("CNAME").ToString, dt.Rows(i).Item("Value").ToString))
                '        'RADIO1.Items.
                '    Next
                'End If
                '若有帶 jobValue 值，自動選擇功能如下：
                If Convert.ToString(rqJobValue) <> "" Then
                    Dim sql5 As String = ""
                    sql5 = ""
                    sql5 += " SELECT ig.GovClass ,ig.GCode1 "
                    sql5 += " FROM KEY_TRAINTYPE ke "
                    sql5 += " JOIN ID_GOVCLASSCAST ig ON ke.GCID = ig.GCID "
                    sql5 += " WHERE ke.TMID = " & Convert.ToString(rqJobValue)
                    Dim dr As DataRow = DbAccess.GetOneRow(sql5, objconn)
                    If Not dr Is Nothing Then
                        If Convert.ToString(dr("GovClass")) <> "" Then Common.SetListItem(Radio1, dr("GovClass"))
                        'Radio1.SelectedValue = dr("GovClass")
                        GCode1 = TIMS.Get_GCode1(GCode1, Radio1.SelectedValue, objconn)
                        If Convert.ToString(dr("GCode1")) <> "" Then Common.SetListItem(GCode1, dr("GCode1"))
                        'GCode1.SelectedValue = dr("GCode1")
                        GCode2 = TIMS.Get_GCode2(GCode2, Radio1.SelectedValue, GCode1.SelectedValue, objconn)
                    Else
                        '訓練業別超過預期資料，嚴重錯誤警示
                        normaltable.Style("display") = "none"
                        'errortable.Style("display") = "inline"
                        errortable.Style("display") = ""
                        'Common.RespWrite(Me, "<script>ClearValue();alert('資料異常，請重新選擇訓練業別!!!');window.close();</script>")
                        'Exit Sub
                    End If
                Else
                    GCode1 = TIMS.Get_GCode1(GCode1, Radio1.SelectedValue, objconn) '未帶有效值。
                End If

                '非 學分班不可選擇 ;學分班 可選擇：
                Radio1.Enabled = True
                GCode1.Enabled = True
                If Convert.ToString(rqPointYN) <> "" Then
                    If Convert.ToString(rqPointYN) = "N" Then
                        Radio1.Enabled = False
                        GCode1.Enabled = False
                    End If
                End If
        End Select
    End Sub

    '層級選擇後。
    Private Sub Radio1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        If Radio1.SelectedValue <> "" Then
            Select Case strYears
                Case cst_y2015'"2015"
                    '無此選項。 'Radio1
                Case cst_y2014 '"2014" 'else
                    GCode1 = TIMS.Get_GCode1(GCode1, Radio1.SelectedValue, objconn)
            End Select
        Else
            Common.MessageBox(Me, "請選擇有效資料")
        End If
        GCode2.Items.Clear()
    End Sub

    '類別選擇後。
    Private Sub GCode1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GCode1.SelectedIndexChanged
        If Radio1.SelectedValue <> "" And GCode1.SelectedValue <> "" Then
            Select Case strYears
                Case Is >= cst_y2018 '"2018"
                    GCode2 = Get_DDLGCode3C(GCode2, GCode1.SelectedValue, objconn)
                Case cst_y2015 '"2015"
                    'GCode1.SelectedValue 為父層順序。
                    GCode2 = TIMS.GetGCode2B(GCode2, GCode1.SelectedValue, objconn)
                Case cst_y2014 '"2014" 'else
                    GCode2 = TIMS.Get_GCode2(GCode2, Radio1.SelectedValue, GCode1.SelectedValue, objconn)
            End Select
        Else
            GCode2.Items.Clear()
            Common.MessageBox(Me, "請選擇有效資料")
        End If
    End Sub

    '選擇後
    Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick, Button2.ServerClick
        Common.RespWrite(Me, "<script>window.close();</script>")
    End Sub
End Class