Public Class TeachDesc1
    Inherits AuthBasePage

    'Const cst_01 As String = "01"
    'Const cst_02 As String = "02"
    'Const cst_03 As String = "03"
    'Const cst_04 As String = "04"
    'Const cst_05 As String = "05"
    Const cst_99 As String = "99"
    Const cst_TCTYPE_A1 As String = "A1"
    Const cst_TCTYPE_A As String = "A"
    Const cst_TCTYPE_AX As String = "AX"
    Const cst_TCTYPE_B As String = "B"
    Dim ff As String = ""
    Dim SSS3 As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Hid_VALUE1.Value = ""
            'Hid_Text1.Value = ""
            Call sCreate1()
        End If
    End Sub

    Sub sCreate1()
        Dim rqType1 As String = TIMS.ClearSQM(Request("TCTYPE")) '"A"
        Dim rqRID As String = TIMS.ClearSQM(Request("RID")) '"A"
        Hid_OPTextBox1.Value = TIMS.ClearSQM(Request("TB1")) '"A"
        btnSend1.Attributes("onclick") = "return ReturnValue();"
        'rqType1 = "2"
        Label1.Visible = False
        Label2.Visible = False
        'btnSend1.Value = "選擇"
        btnSend1.Text = "選擇"
        LabSB3.Visible = False
        txtSubject2.Visible = False
        ddlTeachingCond1.Visible = False
        CblTeachingCond2.Items.Clear()

        Select Case UCase(rqType1)
            Case cst_TCTYPE_A
                Dim dtX1 As DataTable = Get_DtX(cst_TCTYPE_A1)
                '師資遴選辦法說明 2層-1
                Label1.Visible = True
                ddlTeachingCond1.Visible = True
                With ddlTeachingCond1
                    .Items.Add(New ListItem("==請選擇==", ""))
                    For Each drv As DataRow In dtX1.Rows
                        .Items.Add(New ListItem(Convert.ToString(drv("TCVTEXT1")), CStr(drv("TC1"))))
                    Next
                    '.Items.Add(New ListItem("(一)符合相關教師資格", cst_01))
                    '.Items.Add(New ListItem("(二)符合相關技術專業技能", cst_02))
                    '.Items.Add(New ListItem("(三)產業界專業人士", cst_03))
                    '.Items.Add(New ListItem("(四)其他", cst_99))
                End With
            Case cst_TCTYPE_B
                Dim dtX1 As DataTable = Get_DtX(cst_TCTYPE_B)
                '助教遴選辦法說明 1層(B01~B99)
                Label2.Visible = True
                ff = "1=1"
                SSS3 = "TCSORT"
                For Each drv As DataRow In dtX1.Select(ff, SSS3)
                    Utl_Addlist1(CblTeachingCond2, Convert.ToString(drv("TCVTEXT1")), CStr(drv("TCSEQ")))
                Next
#Region "(No Use)"

                'With CblTeachingCond2
                '    Utl_Addlist1(CblTeachingCond2, "(一)大專學歷以上畢業，持有與該班次職類群相關之證照、或擔任相關技術工作累計達 2 年以上者。", "B01")
                '    Utl_Addlist1(CblTeachingCond2, "(二)大專學歷以上相關科系畢業，並任職與課程相關之專業領域行業 1 年以上者。", "B02")
                '    Utl_Addlist1(CblTeachingCond2, "(三)高中職畢業，持有與該班次職類群相關之證照、或擔任相關技術工作累計達 3 年以上者。", "B03")
                '    Utl_Addlist1(CblTeachingCond2, "(四)具特殊專業技藝者(師傅)，並從事該行業累積達 3 年以上。", "B04")
                '    Utl_Addlist1(CblTeachingCond2, "(五)未符合上述規定者，請檢具其他足以擔任助教資格之相關證明文件，經審查核可者。", "B05")
                '    Utl_Addlist1(CblTeachingCond2, "(六)其他(如為TTQS相關性課程，請填寫此欄位)", "B99")
                'End With

#End Region
                LabSB3.Visible = True
                txtSubject2.Visible = True
        End Select

    End Sub

    Sub Utl_Addlist1(ByVal oChk As CheckBoxList, ByVal sText1 As String, ByVal sValue1 As String)
        Dim oListItem1 As New ListItem(sText1, sValue1)
        If sValue1 <> "" Then oListItem1.Attributes("onclick") = "setTDValue(this.checked,'" & sValue1 & "');"

        oChk.Items.Add(oListItem1)

        If sValue1 <> "" AndAlso Hid_VALUE1.Value.IndexOf(sValue1) > -1 Then oChk.Items.FindByValue(sValue1).Selected = True
    End Sub

    Function Get_DtX(ByVal TCTYPE As String) As DataTable
        Dim sql As String = ""
        sql &= " SELECT TCSEQ ,TCTYPE ,TC1 ,TC2 ,TCVTEXT1 ,TCSORT "
        sql &= " FROM ID_TEACHDESC "
        sql &= " WHERE 1=1 "
        Select Case TCTYPE
            Case cst_TCTYPE_A1
                sql &= " AND TCTYPE = 'A' "
                sql &= " AND TC2 IS NULL "
            Case cst_TCTYPE_A
                sql &= " AND TCTYPE = 'A' "
                sql &= " AND TC2 IS NOT NULL "
            Case cst_TCTYPE_B
                sql &= " AND TCTYPE = 'B' "
                sql &= " AND TC2 IS NULL "
            Case cst_TCTYPE_AX
                Hid_VALUE1.Value = TIMS.CombiSQM2IN(Hid_VALUE1.Value)
                If Hid_VALUE1.Value <> "" Then
                    sql &= " AND TCSEQ IN (" & Hid_VALUE1.Value & ") "
                Else
                    sql &= " AND 1<>1 "
                End If
        End Select
        Dim dtX1 As DataTable = DbAccess.GetDataTable(sql, objconn)
        Return dtX1
    End Function

    Protected Sub ddlTeachingCond1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTeachingCond1.SelectedIndexChanged
        btnSend1.Text = "選擇"
        btnSend1.Visible = False
        txtSubject2.Visible = False
        CblTeachingCond2.Items.Clear()
        If ddlTeachingCond1.SelectedValue = "" Then Exit Sub

        Select Case ddlTeachingCond1.SelectedValue
            Case cst_99
                'btnSend1.Value = "送出"
                btnSend1.Text = "送出"
                btnSend1.Visible = True
                txtSubject2.Visible = True
            Case Else
                Dim dtX1 As DataTable = Get_DtX(cst_TCTYPE_A)
                ff = "TC1=" & ddlTeachingCond1.SelectedValue
                SSS3 = "TCSORT"
                If dtX1.Select(ff, SSS3).Length > 0 Then btnSend1.Visible = True
                '師資遴選辦法說明 2層-2
                For Each drv As DataRow In dtX1.Select(ff, SSS3)
                    Utl_Addlist1(CblTeachingCond2, Convert.ToString(drv("TCVTEXT1")), CStr(drv("TCSEQ")))
                Next
        End Select

#Region "(No Use)"

        'Dim objList1 As New ListItem
        '師資遴選辦法說明 2層-2
        'Select Case ddlTeachingCond1.SelectedValue
        '    Case cst_01
        '        Utl_Addlist1(CblTeachingCond2, "1.具教育部審定與授課課程相關之合格教(講)師證書影本或各分署所核發職業訓練師、助理研究員聘書者。", "A0101")
        '        Utl_Addlist1(CblTeachingCond2, "2.具有與授課課程相關之博士學歷畢業者。", "A0102")
        '        Utl_Addlist1(CblTeachingCond2, "3.具碩士學歷畢業，曾任相關課程專任教師累計達 1 年或兼任教師 2 年以上者。", "A0103")
        '        Utl_Addlist1(CblTeachingCond2, "4.大學或獨立學院相關學系畢業，曾任相關課程專任教師累計達 2 年或兼任教師 4 年以上者。", "A0104")
        '        Utl_Addlist1(CblTeachingCond2, "5.專科以上學校相關科系畢業，曾任相關課程專任教師累計達 3 年或兼任教師 5 年以上者。", "A0105")
        '    Case cst_02
        '        Utl_Addlist1(CblTeachingCond2, "1.具與授課課程相關之甲級技術士證照或專門職業及技術人員高等考試及格者。", "A0201")
        '        Utl_Addlist1(CblTeachingCond2, "2.大學或獨立學院相關科系畢業，並取得與授課課程相關之乙級技術士證者。", "A0202")
        '        Utl_Addlist1(CblTeachingCond2, "3.專科學歷以上畢業，曾任與授課課程相關之專業或技術工作滿 2 年，並取得相關之乙級技術士證者。", "A0203")
        '        Utl_Addlist1(CblTeachingCond2, "4.高中(職)學校畢業，曾任相關之專業或技術工作滿 4 年，並取得與授課課目相關之乙級技術士證者。", "A0204")
        '    Case cst_03
        '        Utl_Addlist1(CblTeachingCond2, "1.在技術上有特殊造詣（師傅），具與應聘職類相關之教學工作或專業或技術工作累計達 5 年以上者。", "A0301")
        '        Utl_Addlist1(CblTeachingCond2, "2.大專學歷以上畢業，任職事業單位 2 年以上，其專業技能足以擔任授課講師者。", "A0302")
        '        Utl_Addlist1(CblTeachingCond2, "3.未符合上述規定者，請檢具其他足以擔任授課師資之相關證明文件，經審查核可者。", "A0303")
        '    Case cst_99
        '        btnSend1.Value = "送出"
        '        txtSubject2.Visible = True
        'End Select

#End Region
    End Sub

    ''' <summary>取得選取值</summary>
    ''' <param name="TCSEQ"></param>
    ''' <returns></returns>
    Function Get_X2Value(ByVal TCSEQ As String) As String
        Dim rst As String = ""
        Dim dtX1 As DataTable = Get_DtX(cst_TCTYPE_AX)
        For Each drv As DataRow In dtX1.Rows
            If Convert.ToString(drv("TCVTEXT1")) <> "" Then
                rst &= String.Concat(If(rst <> "", vbCrLf, ""), drv("TCVTEXT1"))
            End If
        Next
        txtSubject2.Text = TIMS.ClearSQM(txtSubject2.Text)
        If txtSubject2.Text <> "" Then
            rst &= String.Concat(If(rst <> "", vbCrLf, ""), "其他：", txtSubject2.Text)
        End If
        Return rst
    End Function

    ''' <summary>送出按鈕</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnSend1_Click(sender As Object, e As EventArgs) Handles btnSend1.Click
        If Hid_OPTextBox1.Value = "" Then Exit Sub
        Hid_OPTextBox1.Value = TIMS.ClearSQM(Hid_OPTextBox1.Value)
        'Dim rqType1 As String = TIMS.ClearSQM(Request("TCTYPE")) '"A"
        Dim JSvalue1 As String = Get_X2Value(Hid_VALUE1.Value)
        'Dim rqRID As String = TIMS.ClearSQM(Request("RID")) '"A"
        Dim strScript1 As String = ""
        strScript1 &= "<script language=javascript>"
        strScript1 &= "function ValueTD1(){"
        strScript1 &= "    if (opener == undefined) { window.close(); return; }"
        strScript1 &= "    var OP1=window.opener.document.getElementById('" & Hid_OPTextBox1.Value & "');"
        strScript1 &= "    if(OP1!=null) OP1.value='" & Common.GetJsString(JSvalue1) & "';"
        strScript1 &= "    window.close();"
        strScript1 &= "}"
        strScript1 &= "ValueTD1();"
        strScript1 &= "window.close();"
        strScript1 &= "</script>"
        Common.RespWrite(Me, strScript1)
    End Sub
End Class