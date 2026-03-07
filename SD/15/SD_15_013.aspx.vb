Public Class SD_15_013
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents TitleLab1 As System.Web.UI.WebControls.Label
    Protected WithEvents TitleLab2 As System.Web.UI.WebControls.Label
    Protected WithEvents Distid As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents Tcitycode As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents Ocitycode As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents GovClassName As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents CCID As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents KID_6 As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents KID_10 As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents KID_4 As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents KID_7 As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents PointYN As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Apppass As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Endclass As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Appmoney As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents Stopclass As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents SDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents SDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents EDate1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents EDate2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents ChbExit As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents BtnExp As System.Web.UI.WebControls.Button
    Protected WithEvents DistHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents TcityHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents OcityHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents GovClassHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents CCIDHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents KID_6_TR As System.Web.UI.HtmlControls.HtmlTableRow
    Protected WithEvents KID_6_hid As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents KID_10_TR As System.Web.UI.HtmlControls.HtmlTableRow
    Protected WithEvents KID_10_hid As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents KID_4_TR As System.Web.UI.HtmlControls.HtmlTableRow
    Protected WithEvents KID_4_hid As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents TR_7dep As System.Web.UI.HtmlControls.HtmlTableRow
    Protected WithEvents KID_7_hid As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents ChbExitHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents PackageType As System.Web.UI.WebControls.CheckBoxList
    Protected WithEvents PackageHidden As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents center As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button3 As System.Web.UI.WebControls.Button
    Protected WithEvents HistoryRID As System.Web.UI.WebControls.Table
    Protected WithEvents TMID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents OCID1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents HistoryTable As System.Web.UI.WebControls.Table
    Protected WithEvents Plankind As System.Web.UI.WebControls.RadioButtonList
    Protected WithEvents RIDValue As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Button2 As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents OCIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents TMIDValue1 As System.Web.UI.HtmlControls.HtmlInputHidden

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Const Cst_全部 As Integer = 0
    Const Cst_申請人次 As Integer = 1
    Const Cst_申請補助費 As Integer = 2
    Const Cst_核定人次 As Integer = 3
    Const Cst_核定補助費 As Integer = 4
    Const Cst_實際開訓人次 As Integer = 5
    Const Cst_實際開訓人次加總 As Integer = 6
    Const Cst_預估補助費 As Integer = 7
    Const Cst_預估補助費加總 As Integer = 8
    Const Cst_結訓人次 As Integer = 9
    Const Cst_撥款人次 As Integer = 10
    Const Cst_撥款補助費 As Integer = 11
    Const Cst_不預告訪視次數_實地抽訪 As Integer = 12
    Const Cst_不預告訪視次數_電話抽訪 As Integer = 13
    Const Cst_累積訪視異常次數 As Integer = 14
    Const Cst_會計查帳次數 As Integer = 15
    Const Cst_離訓人次 As Integer = 16
    Const Cst_退訓人次 As Integer = 17
    Const Cst_訓練時數 As Integer = 18
    Const Cst_人時成本 As Integer = 19
    Const Cst_上課時間 As Integer = 20

    'Const Cst_撥款日期 As Integer = 21
    'Const Cst_統一編號 As Integer = 22
    'Const Cst_立案縣市 As Integer = 23
    'Const Cst_包班事業單位 As Integer = 24


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        If Not IsPostBack Then

            Distid = TIMS.Get_DistID(Distid)
            Distid.Items.Insert(0, New ListItem("全部", 0))
            PackageType = TIMS.GetPackageType(PackageType)
            Tcitycode = TIMS.Get_CityName(Tcitycode)
            Ocitycode = TIMS.Get_CityName(Ocitycode)
            GovClassName = Get_GovClass(GovClassName) '訓練業別
            TIMS.Get_ClassCatelog(CCID)    '課程職能
            CCID.Items.Insert(0, New ListItem("全部", 0))
            PointYN = AddList(PointYN)
            Apppass = AddList(Apppass)
            Endclass = AddList(Endclass)
            'Appmoney = AddList(Appmoney)
            Stopclass = AddList(Stopclass)
            GET_ExitCell(ChbExit) '匯出欄位 

            Distid.Attributes("onclick") = "SelectAll('Distid','DistHidden');"
            PackageType.Attributes("onclick") = "SelectAll('PackageType','PackageHidden');"
            Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"
            Ocitycode.Attributes("onclick") = "SelectAll('Ocitycode','OcityHidden');"
            GovClassName.Attributes("onclick") = "SelectAll('GovClassName','GovClassHidden');"
            CCID.Attributes("onclick") = "SelectAll('CCID','CCIDHidden');"
            ChbExit.Attributes("onclick") = "SelectAll('ChbExit','ChbExitHidden');"

            'Plankind.SelectedIndex = 0
            PointYN.SelectedIndex = 0
            Apppass.SelectedIndex = 0
            Endclass.SelectedIndex = 0
            Appmoney.SelectedIndex = 0
            Stopclass.SelectedIndex = 0

            Distid.Enabled = True
            If sm.UserInfo.DistID <> "000" Then
                Distid.SelectedValue = sm.UserInfo.DistID
                Distid.Enabled = False
            End If

            Get_KeyBusiness(KID_6, "01")
            Get_KeyBusiness(KID_10, "02")
            Get_KeyBusiness(KID_4, "03")
            ' get_Key_BusID(KID_7, "04") '舊的不管

            '產業別鍵詞
            KID_6.Attributes("onclick") = "SelectAll('KID_6','KID_6_hid');"
            KID_10.Attributes("onclick") = "SelectAll('KID_10','KID_10_hid');"
            KID_4.Attributes("onclick") = "SelectAll('KID_4','KID_4_hid');"
            'KID_7.Attributes("onclick") = "SelectAll('KID_7','KID_7_hid');"

            'If sm.UserInfo.Years >= 2011 Then
            '    KID_6_TR.Visible = True
            '    KID_10_TR.Visible = True
            '    KID_4_TR.Visible = True
            'Else
            '    KID_6_TR.Visible = False
            '    KID_10_TR.Visible = False
            '    KID_4_TR.Visible = False
            'End If
        End If

        'If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
        '    Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg.aspx');"
        'Else
        '    Button2.Attributes("onclick") = "javascript@openOrg('../../Common/LevOrg1.aspx');"
        'End If

    End Sub

    '產業別鍵詞
    Function Get_KeyBusiness(ByVal obj As ListControl, ByVal DepID As String) As ListControl
        Dim dt As DataTable = Nothing
        Dim sql As String

        sql = " select  a.KNAME ,a.SeqNo from  Key_Business a"
        sql += " join   Key_Depot b on  b.Depid=a.DepID"
        sql += " where 1=1"
        If DepID <> "" Then
            sql += " and a.DepID='" & DepID & "'"
        End If
        sql += " and a.Status is  null"
        dt = DbAccess.GetDataTable(sql)
        With obj
            .Items.Clear()
            .DataSource = dt
            .DataTextField = "KNAME"
            .DataValueField = "SeqNo"
            .DataBind()
            If TypeOf obj Is DropDownList Then
                .Items.Insert(0, New ListItem("===請選擇===", ""))
            End If
            If TypeOf obj Is CheckBoxList Then
                .Items.Insert(0, New ListItem("全部", ""))
            End If
        End With

        Return obj
    End Function

    '是否鍵詞
    Function AddList(ByVal obj As ListControl) As ListControl
        obj.Items.Clear()
        obj.Items.Insert(0, New ListItem("不區分", "A"))
        obj.Items.Insert(1, New ListItem("是", "Y"))
        obj.Items.Insert(2, New ListItem("否", "N"))
        Return obj
    End Function

    '職類課程人時成本分類檔
    Function Get_GovClass(ByVal Obj As ListControl) As ListControl
        Dim sql As String
        Dim dt As DataTable
        sql = "" & vbCrLf
        sql += " Select " & vbCrLf
        sql += " convert(varchar, GovClass) +',' +convert(varchar,GCode1) as Gcid" & vbCrLf
        sql += " ,case GovClass " & vbCrLf
        sql += "  when 1 Then '院' + '-' + right('0' +GCode1,2) " & vbCrLf
        sql += "  when 2 Then '局' + '-' + right('0' +GCode1,2) end as GovClass" & vbCrLf
        sql += " from ID_GovClassCast" & vbCrLf
        sql += " where 1=1 " & vbCrLf
        sql += " and GCode2 is NULL " & vbCrLf
        sql += " and GovClass in (1,2)" & vbCrLf
        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            With Obj
                .Items.Clear()
                .DataSource = dt
                .DataTextField = "GovClass"
                .DataValueField = "Gcid"
                .DataBind()
                If TypeOf Obj Is DropDownList Then
                    .Items.Insert(0, New ListItem("===請選擇===", ""))
                End If
                If TypeOf Obj Is CheckBoxList Then
                    .Items.Insert(0, New ListItem("全部", 0))
                End If
            End With
        End If
        Return Obj
    End Function

    'GET_ExitCell(ChbExit)'匯出欄位 
    Function GET_ExitCell(ByVal obj As ListControl) As ListControl
        obj.Items.Clear()
        obj.Items.Insert(Cst_全部, New ListItem("全部", Cst_全部))
        obj.Items.Insert(Cst_申請人次, New ListItem("申請人次", Cst_申請人次))
        obj.Items.Insert(Cst_申請補助費, New ListItem("申請補助費", Cst_申請補助費))
        obj.Items.Insert(Cst_核定人次, New ListItem("核定人次", Cst_核定人次))
        obj.Items.Insert(Cst_核定補助費, New ListItem("核定補助費", Cst_核定補助費))
        obj.Items.Insert(Cst_實際開訓人次, New ListItem("實際開訓人次", Cst_實際開訓人次))
        obj.Items.Insert(Cst_實際開訓人次加總, New ListItem("實際開訓人次加總", Cst_實際開訓人次加總))
        obj.Items.Insert(Cst_預估補助費, New ListItem("預估補助費", Cst_預估補助費))
        obj.Items.Insert(Cst_預估補助費加總, New ListItem("預估補助費加總", Cst_預估補助費加總))
        obj.Items.Insert(Cst_結訓人次, New ListItem("結訓人次", Cst_結訓人次))
        obj.Items.Insert(Cst_撥款人次, New ListItem("撥款人次", Cst_撥款人次))
        obj.Items.Insert(Cst_撥款補助費, New ListItem("撥款補助費", Cst_撥款補助費))
        obj.Items.Insert(Cst_不預告訪視次數_實地抽訪, New ListItem("不預告訪視次數-實地抽訪", Cst_不預告訪視次數_實地抽訪))
        obj.Items.Insert(Cst_不預告訪視次數_電話抽訪, New ListItem("不預告訪視次數-電話抽訪", Cst_不預告訪視次數_電話抽訪))
        obj.Items.Insert(Cst_累積訪視異常次數, New ListItem("累積訪視異常次數", Cst_累積訪視異常次數))
        obj.Items.Insert(Cst_會計查帳次數, New ListItem("會計查帳次數", Cst_會計查帳次數))
        obj.Items.Insert(Cst_離訓人次, New ListItem("離訓人次", Cst_離訓人次))
        obj.Items.Insert(Cst_退訓人次, New ListItem("退訓人次", Cst_退訓人次))
        'obj.Items.Insert(18, New ListItem("訓練職能", "18"))
        obj.Items.Insert(Cst_訓練時數, New ListItem("訓練時數", Cst_訓練時數))
        obj.Items.Insert(Cst_人時成本, New ListItem("人時成本", Cst_人時成本))
        obj.Items.Insert(Cst_上課時間, New ListItem("上課時間", Cst_上課時間))
    End Function

    '匯出SUB
    Private Sub ExpRpt(ByVal da As SqlDataAdapter)

        Dim sql As String = ""
        Dim dt As New DataTable

        Dim strSearch As String = ""
        Dim DistID2 As String = ""
        Dim TCityCode2 As String = ""
        Dim OCityCode2 As String = ""
        Dim GovClassName2 As String = ""
        Dim PackageType2 As String = ""
        Dim ExportStr As String = ""
        Dim CCID2 As String = ""
        'Dim GovClass2 As Array
        'Dim GCode1 As String = ""
        'Dim i As Integer = 0
        'Dim a As Integer = 0
        'Dim j As Integer = 0
        Dim dr As DataRow
        Dim SeqNostr1 As String = ""
        Dim SeqNostr2 As String = ""
        Dim SeqNostr3 As String = ""
        Dim TMID As String = ""

        ' dt = New DataTable

        '轄區
        If sm.UserInfo.DistID = "000" Then
            For i As Integer = 0 To Distid.Items.Count - 1
                If Distid.Items.Item(i).Selected = True Then
                    If Distid.Items.Item(i).Text <> "全部" Then
                        DistID2 += "'" & Distid.Items.Item(i).Value & "'" & ","
                    End If
                End If
            Next
            If DistID2 <> "" Then
                DistID2 = Left(DistID2, Len(DistID2) - 1)
                strSearch += " and ip.Distid in (" & DistID2 & ")"
            End If
        Else
            strSearch += " and ip.Distid = '" & sm.UserInfo.DistID & "'"
        End If

        'If RIDValue.Value <> "" Then
        '    strSearch += " and ar.RID = '" & RIDValue.Value & "'"
        'End If
        'If OCIDValue1.Value <> "" Then
        '    strSearch += " and cc.ocid = '" & OCIDValue1.Value & "'"
        'End If
        'If Plankind.SelectedIndex <> 0 Then
        '    strSearch += " and oo.OrgKind2 = '" & Plankind.SelectedValue & "'"
        'End If

        '辦訓地縣市
        TCityCode2 = ""
        For i As Integer = 0 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True Then
                If Tcitycode.Items.Item(i).Text <> "全部" Then
                    TCityCode2 += Tcitycode.Items.Item(i).Value & ","
                End If
            End If
        Next

        If sm.UserInfo.Years <= 2010 Then
            If TCityCode2 <> "" Then
                TCityCode2 = Left(TCityCode2, Len(TCityCode2) - 1)
                strSearch += " and ic.CTID in (" & TCityCode2 & ")"
            End If
        ElseIf sm.UserInfo.Years >= 2011 Then
            If TCityCode2 <> "" Then
                TCityCode2 = Left(TCityCode2, Len(TCityCode2) - 1)
                strSearch += " and (ic3.CTID in (" & TCityCode2 & ") or ic4.CTID in (" & TCityCode2 & "))"
            End If
        End If

        '包班總類
        For i As Integer = 0 To PackageType.Items.Count - 1
            If PackageType.Items.Item(i).Selected = True Then
                If PackageType.Items.Item(i).Text <> "全部" Then
                    PackageType2 += PackageType.Items.Item(i).Value & ","
                End If
            End If
        Next
        If PackageType2 <> "" Then
            PackageType2 = Left(PackageType2, Len(PackageType2) - 1)
            strSearch += " and pp.PackageType in (" & PackageType2 & ")"
        End If

        '立案地縣市
        For i As Integer = 0 To Ocitycode.Items.Count - 1
            If Ocitycode.Items.Item(i).Selected = True Then
                If Ocitycode.Items.Item(i).Text <> "全部" Then
                    OCityCode2 += Ocitycode.Items.Item(i).Value & ","
                End If
            End If
        Next
        If OCityCode2 <> "" Then
            OCityCode2 = Left(OCityCode2, Len(OCityCode2) - 1)
            strSearch += " and ic2.CTID in (" & OCityCode2 & ")"
        End If

        '訓練業別
        For i As Integer = 0 To GovClassName.Items.Count - 1
            If GovClassName.Items.Item(i).Selected = True Then
                If GovClassName.Items.Item(i).Text <> "全部" Then
                    GovClassName2 += "'" & GovClassName.Items.Item(i).Value & "'" & ","
                End If
            End If
        Next
        If GovClassName2 <> "" Then
            GovClassName2 = Left(GovClassName2, Len(GovClassName2) - 1)
            strSearch += " and convert(varchar,ig.GovClass)+','+convert(varchar,ig.GCode1) in (" & GovClassName2 & ")"
        End If

        '訓練職能
        For i As Integer = 0 To CCID.Items.Count - 1
            If CCID.Items.Item(i).Selected = True Then
                If CCID.Items.Item(i).Text <> "全部" Then
                    CCID2 += CCID.Items.Item(i).Value & ","
                End If
            End If
        Next
        If CCID2 <> "" Then
            CCID2 = Left(CCID2, Len(CCID2) - 1)
            strSearch += " and pp.ClassCate in (" & CCID2 & ")"
        End If

        '六大新興產業
        For i As Integer = 0 To KID_6.Items.Count - 1
            If KID_6.Items.Item(i).Selected = True Then
                If KID_6.Items.Item(i).Text <> "全部" Then
                    SeqNostr1 += KID_6.Items.Item(i).Value & ","
                End If
            End If
        Next
        If SeqNostr1 <> "" Then
            SeqNostr1 = Left(SeqNostr1, Len(SeqNostr1) - 1)
            strSearch += " and (vd1.seqno in (" & SeqNostr1 & ")"
        End If

        '十大重點服務業
        For i As Integer = 0 To KID_10.Items.Count - 1
            If KID_10.Items.Item(i).Selected = True Then
                If KID_10.Items.Item(i).Text <> "全部" Then
                    SeqNostr2 += KID_10.Items.Item(i).Value & ","
                End If
            End If
        Next
        If SeqNostr2 <> "" Then
            SeqNostr2 = Left(SeqNostr2, Len(SeqNostr2) - 1)
            If SeqNostr1 <> "" Then
                strSearch += " or vd2.seqno in (" & SeqNostr2 & ")"
            Else
                strSearch += " and (vd2.seqno in (" & SeqNostr2 & ")"
            End If
        End If

        '四大新興智慧型產業
        For i As Integer = 0 To KID_4.Items.Count - 1
            If KID_4.Items.Item(i).Selected = True Then
                If KID_4.Items.Item(i).Text <> "全部" Then
                    SeqNostr3 += KID_4.Items.Item(i).Value & ","
                End If
            End If
        Next
        If SeqNostr3 <> "" Then
            SeqNostr3 = Left(SeqNostr3, Len(SeqNostr3) - 1)
            If SeqNostr2 <> "" Then
                strSearch += " or vd3.SeqNo in (" & SeqNostr3 & ")"
            Else
                strSearch += " and (vd3.SeqNo in (" & SeqNostr3 & ")"
            End If
        End If

        If (SeqNostr1 <> "" OrElse SeqNostr2 <> "" OrElse SeqNostr3 <> "") Then
            strSearch += ")"
        End If

        If PointYN.SelectedIndex <> 0 Then
            strSearch += " and pp.PointYN = '" & PointYN.SelectedValue & "'"
        End If

        If Apppass.SelectedIndex <> 0 Then
            strSearch += " and pp.AppliedResult = '" & Apppass.SelectedValue & "'"
        End If

        If Endclass.SelectedIndex <> 0 Then   '是否結訓
            If Endclass.SelectedValue = "Y" Then
                strSearch += " and cc.FTDate < getdate() "
            Else
                strSearch += " and cc.FTDAte >= getdate()"
            End If
        End If

        If Appmoney.SelectedIndex <> 0 Then
            strSearch += " and ss.AppliedStatus = '" & Appmoney.SelectedValue & "'"
        End If

        'If Apppass.SelectedIndex <> 0 Then
        '    str += " and pp.AppliedResult = '" & Apppass.SelectedValue & "'"
        'End If

        If Stopclass.SelectedIndex <> 0 Then
            If Stopclass.SelectedValue = "Y" Then
                '不開班
                strSearch += " and cc.NotOpen = 'Y'"
            Else
                '開班
                strSearch += " and cc.NotOpen = 'N'"
            End If
        End If

        If SDate1.Text <> "" Then
            'str += " and cc.STDate >='" & SDate1.Text & "'"
            strSearch += " and pp.STDate >='" & SDate1.Text & "'"
        End If

        If SDate2.Text <> "" Then
            'str += " and cc.STDate <='" & SDate2.Text & "'"
            strSearch += " and pp.STDate <='" & SDate2.Text & "'"
        End If
        If EDate1.Text <> "" Then
            'str += " and cc.FTDate >='" & EDate1.Text & "'"
            strSearch += " and pp.FDDate >='" & EDate1.Text & "'"
        End If
        If EDate2.Text <> "" Then
            'str += " and cc.FTDate <='" & EDate2.Text & "'"
            strSearch += " and pp.FDDate <='" & EDate2.Text & "'"
        End If


        sql += "select" & vbCrLf
        sql += " a.tmid," & vbCrLf
        sql += "a.OrgTypeName," & vbCrLf '單位屬性 
        sql += "a.DistName," & vbCrLf
        sql += "a.ocid," & vbCrLf
        sql += "a.orgname," & vbCrLf
        sql += "a.ClassName," & vbCrLf
        sql += "a.CyclType, " & vbCrLf
        sql += "a.ClassID," & vbCrLf
        sql += "a.STDate," & vbCrLf
        sql += "a.FDDate," & vbCrLf
        sql += "a.PackageType," & vbCrLf
        sql += "a.ADefGovCost," & vbCrLf
        sql += "a.ATNum," & vbCrLf
        sql += "a.DefGovCost," & vbCrLf
        sql += "a.TNum," & vbCrLf
        sql += "a.THours," & vbCrLf
        '/*人時成本*/
        sql += "(a.DefGovCost/(case when a.TNum = 0 then 1 else a.TNum end))/(case when a.THours = 0 then 1 else a.THours end) as PhCost," & vbCrLf
        sql += "a.WEEKSTIME," & vbCrLf
        sql += "a.kname1," & vbCrLf
        sql += "a.kname2," & vbCrLf
        sql += "a.kname3," & vbCrLf
        'sql += "a.GovClass," & vbCrLf
        'sql += "a.GCode1," & vbCrLf
        'sql += "a.GCode2," & vbCrLf
        sql += "a.GCodeName," & vbCrLf
        sql += "a.CCName," & vbCrLf
        sql += "a.AddressSciPTID," & vbCrLf
        sql += "a.AddressTechPTID," & vbCrLf
        sql += "a.openstudcount1," & vbCrLf
        sql += "a.openstudcount2," & vbCrLf
        sql += "a.openstudcount3," & vbCrLf
        sql += "a.openstudcount97," & vbCrLf
        sql += "a.openstudcountall," & vbCrLf
        sql += "isnull(a.cost1,0) as cost1," & vbCrLf
        sql += "isnull(a.cost2,0) as cost2," & vbCrLf
        sql += "isnull(a.cost3,0) as cost3," & vbCrLf
        sql += "isnull(a.cost97,0) as cost97," & vbCrLf
        sql += "isnull(a.costAll,0) as costAll," & vbCrLf
        'sql += "a.closestudcout," & vbCrLf
        sql += "a.closestudcout01," & vbCrLf
        sql += "a.closestudcout02," & vbCrLf
        sql += "a.closestudcout03," & vbCrLf
        sql += "a.closestudcout97," & vbCrLf
        sql += "a.closestudcoutall," & vbCrLf
        sql += "a.std_cnt2," & vbCrLf
        sql += "a.std_cnt3," & vbCrLf
        sql += "a.budcountall," & vbCrLf
        sql += "a.budcountall2," & vbCrLf
        sql += "a.budcountall3," & vbCrLf
        sql += "a.budcountall97," & vbCrLf
        sql += "a.bud03count, " & vbCrLf
        sql += "a.bud03count07," & vbCrLf
        sql += "a.bud03count05," & vbCrLf
        sql += "a.bud03count06," & vbCrLf
        sql += "a.bud03count04," & vbCrLf
        sql += "a.bud03count28," & vbCrLf
        sql += "a.bud03count10," & vbCrLf
        sql += "a.bud03count26," & vbCrLf
        sql += "a.bud02count," & vbCrLf
        sql += "a.bud02count07," & vbCrLf
        sql += "a.bud02count05," & vbCrLf
        sql += "a.bud02count06," & vbCrLf
        sql += "a.bud02count04," & vbCrLf
        sql += "a.bud02count28," & vbCrLf
        sql += "a.bud02count10," & vbCrLf
        sql += "a.bud02count26," & vbCrLf
        sql += "a.bud01count," & vbCrLf
        sql += "a.bud01count07," & vbCrLf
        sql += "a.bud01count05," & vbCrLf
        sql += "a.bud01count06," & vbCrLf
        sql += "a.bud01count04," & vbCrLf
        sql += "a.bud01count28," & vbCrLf
        sql += "a.bud01count10," & vbCrLf
        sql += "a.bud01count26," & vbCrLf
        sql += "a.bud97count, " & vbCrLf
        sql += "a.bud97count07," & vbCrLf
        sql += "a.bud97count05," & vbCrLf
        sql += "a.bud97count06," & vbCrLf
        sql += "a.bud97count04," & vbCrLf
        sql += "a.bud97count28," & vbCrLf
        sql += "a.bud97count10," & vbCrLf
        sql += "a.bud97count26," & vbCrLf
        sql += "a.budmoneyall," & vbCrLf
        sql += "a.budmoneyall2," & vbCrLf
        sql += "a.budmoneyall3," & vbCrLf
        sql += "a.budmoneyall97," & vbCrLf
        sql += "a.bud03money," & vbCrLf
        sql += "a.bud03money07," & vbCrLf
        sql += "a.bud03money05," & vbCrLf
        sql += "a.bud03money06," & vbCrLf
        sql += "a.bud03money04," & vbCrLf
        sql += "a.bud03money28," & vbCrLf
        sql += "a.bud03money10," & vbCrLf
        sql += "a.bud03money26," & vbCrLf
        sql += "a.bud02money," & vbCrLf
        sql += "a.bud02money07," & vbCrLf
        sql += "a.bud02money05," & vbCrLf
        sql += "a.bud02money06," & vbCrLf
        sql += "a.bud02money04," & vbCrLf
        sql += "a.bud02money28," & vbCrLf
        sql += "a.bud02money10," & vbCrLf
        sql += "a.bud02money26," & vbCrLf
        sql += "a.bud01money," & vbCrLf
        sql += "a.bud01money07," & vbCrLf
        sql += "a.bud01money05," & vbCrLf
        sql += "a.bud01money06," & vbCrLf
        sql += "a.bud01money04," & vbCrLf
        sql += "a.bud01money28," & vbCrLf
        sql += "a.bud01money10," & vbCrLf
        sql += "a.bud01money26," & vbCrLf
        sql += "a.bud97money," & vbCrLf
        sql += "a.bud97money07," & vbCrLf
        sql += "a.bud97money05," & vbCrLf
        sql += "a.bud97money06," & vbCrLf
        sql += "a.bud97money04," & vbCrLf
        sql += "a.bud97money28," & vbCrLf
        sql += "a.bud97money10," & vbCrLf
        sql += "a.bud97money26," & vbCrLf
        '/*總特殊學員就保人次*/
        sql += "a.budcountall - a.bud03count as Sbudcount01," & vbCrLf
        '/*總特殊學員就安人次*/
        sql += "a.budcountall2 - a.bud02count as Sbudcount02," & vbCrLf
        '/*總特殊學員公務人次*/
        sql += "a.budcountall3 - a.bud01count as Sbudcount03," & vbCrLf
        '/*總特殊學員協助人次*/
        sql += "a.budcountall97 - a.bud97count as Sbudcount97," & vbCrLf
        '/*總特殊學員就保金額*/
        sql += "a.budmoneyall - a.bud03money as Sbudmoney01," & vbCrLf
        '/*總特殊學員就安金額*/
        sql += "a.budmoneyall2 - a.bud02money as Sbudmoney02," & vbCrLf
        '/*總特殊學員公務金額*/
        sql += "a.budmoneyall3 - a.bud01money as Sbudmoney03," & vbCrLf
        '/*總特殊學員協助金額*/
        sql += "a.budmoneyall97 - a.bud97money as Sbudmoney97," & vbCrLf
        sql += "isnull(e.vitN,0) as vitN," & vbCrLf
        sql += "isnull(e.cuall,0) as cuall," & vbCrLf
        sql += "isnull(f.VitTelN,0) as VitTelN," & vbCrLf
        sql += "isnull(f.ctall,0) as ctall," & vbCrLf
        sql += "isnull(e.vitN,0) + isnull(f.VitTelN,0) as vtn" & vbCrLf
        sql += "from" & vbCrLf
        sql += "(" & vbCrLf '轄區中心
        sql += " select pp.tmid, o1.typeid2+'-'+o1.typeid2name   as  OrgTypeName,  " '單位屬性
        sql += " case when ip.Distid = '001' then '北區職訓中心'" & vbCrLf
        sql += " when ip.Distid = '002' then '泰山職訓中心'" & vbCrLf
        sql += " when ip.Distid = '003' then '桃園職訓中心'" & vbCrLf
        sql += " when ip.Distid = '004' then '中區職訓中心'" & vbCrLf
        sql += " when ip.Distid = '005' then '台南職訓中心'" & vbCrLf
        sql += " when ip.Distid = '006' then '南區職訓中心'" & vbCrLf
        sql += " end as DistName," & vbCrLf
        sql += " cc.ocid," & vbCrLf
        '/*單位名稱*/
        'sql += "oo.orgname," & vbCrLf
        sql += "Replace(oo.orgname, CHAR(9), CHAR(32)) as orgname, " & vbCrLf  '將tab鍵產生的空白取代為空白,避免匯出的欄位多出空格沒有對齊
        '/*課程名稱*/
        'sql += "pp.ClassName, " & vbCrLf
        sql += "Replace(pp.ClassName, CHAR(9), CHAR(32))  as ClassName, " & vbCrLf '將tab鍵產生的空白取代為空白,避免匯出的欄位多出空格沒有對齊
        '/*期別*/
        sql += "pp.CyclType, " & vbCrLf
        '/*課程代碼*/
        'sql += "cc.Years+ '0' + f.ClassID + cc.CyclType as ClassID, " & vbCrLf
        sql += " cc.ocid as ClassID," & vbCrLf
        '/*開訓練日期*/
        sql += "pp.STDate," & vbCrLf
        '"/*結訓日期*/"
        sql += "pp.FDDate, " & vbCrLf
        '"/*包班總類*/"
        sql += "case when pp.PackageType = '1' then '非包班' when  pp.PackageType = '2' then '企業包班'  when  pp.PackageType = '3' then '聯合企業包班' end as PackageType, " & vbCrLf
        '/六大新興產業/
        sql += "vd1.kname AS kname1," & vbCrLf
        '/十大重點服務/
        sql += "vd2.kname AS kname2," & vbCrLf
        '/四大新興智慧型產業/
        sql += "vd3.kname AS kname3," & vbCrLf
        '/訓練業別編碼/
        'sql += "case when ig.GovClass in(1,2) then ig.GovClass else '' end as　GovClass," & vbCrLf
        'sql += "case when ig.GCode1 in (99,100,101)," & vbCrLf
        'sql += "ig.GCode2," & vbCrLf
        sql += "case when ig.GovClass = 1 then '院'+'-'+ right('0'+convert(varchar,ig.GCode1),2) + '-' +right('0'+convert(varchar,ig.GCode2),2)" & vbCrLf
        sql += "     when ig.GovClass = 2 then '局'+'-'+ right('0'+convert(varchar,ig.GCode1),2) + '-' +right('0'+convert(varchar,ig.GCode2),2) " & vbCrLf
        sql += " else '' end as GCodeName," & vbCrLf
        '/訓練職能/
        sql += "kc.CCName," & vbCrLf
        '/學科場地地址/
        sql += "case when convert(int,pp.PlanYear) <= 2010 then isnull(ic.CTName,'') else isnull( ic3.CTName,'') end as AddressSciPTID," & vbCrLf
        '/術科場地地址/
        sql += "isnull(ic4.CTName,'') as AddressTechPTID," & vbCrLf
        '/*申請總補助費*/
        sql += "isnull(pp.DefGovCost,0)  as ADefGovCost," & vbCrLf
        '/*申請人次*/
        sql += "isnull(pp.TNum,0)  as ATNum," & vbCrLf
        '/訓練時數/
        sql += "isnull(pp.THours,0) as THours," & vbCrLf
        '/*核定總補助費*/
        sql += "case when pp.AppliedResult = 'Y' then isnull(pp.DefGovCost,0) else 0 end as DefGovCost," & vbCrLf
        '/*核定人次*/
        sql += "case when pp.AppliedResult = 'Y' then isnull(pp.TNum,0) else 0 end as TNum," & vbCrLf
        '/*上課時間*/
        sql += "replace(dbo.dbo.fn_GET_plan_onclass(pp.PlanID,pp.ComIDNO,pp.SeqNo,'WEEKTIME'),' ','' ) as WEEKSTIME," & vbCrLf
        '/*開訓人次-就保*/
        sql += "sum(case when cc.Notopen = 'N' and cs.IsApprPaper = 'Y' and cs.BudgetID = '03'  then 1 else 0 end ) as openstudcount1," & vbCrLf
        '/*開訓人次-就安*/"
        sql += "sum(case when cc.Notopen = 'N' and cs.IsApprPaper = 'Y' and cs.BudgetID = '02'  then 1 else 0 end ) as openstudcount2," & vbCrLf
        '/*開訓人次-公務*/
        sql += "sum(case when cc.Notopen = 'N' and cs.IsApprPaper = 'Y' and cs.BudgetID = '01'  then 1 else 0 end ) as openstudcount3," & vbCrLf
        '/*開訓人次-協助*/
        sql += "sum(case when cc.Notopen = 'N' and cs.IsApprPaper = 'Y' and cs.BudgetID = '97'  then 1 else 0 end ) as openstudcount97," & vbCrLf
        '/*開訓人次-合計*/
        sql += "sum(case when cc.Notopen = 'N' and cs.IsApprPaper = 'Y' and cs.BudgetID in ('01','02','03','97')  then 1 else 0 end ) as openstudcountall," & vbCrLf

        '/*預估補助費-就保*/
        sql += "sum( case when isnull(cs.SupplyID,0)=0 or pp.TNum = 0 then 0" & vbCrLf
        sql += "when cs.SupplyID='1' and pp.TNum <> 0 and cs.BudgetID = '03' and cs.IsApprPaper = 'Y' then ( isnull(pp.TotalCost,0)/isnull(pp.TNum,1))*0.8 " & vbCrLf
        sql += "when cs.SupplyID='9' and pp.TNum <> 0 and cs.BudgetID = '03' and cs.IsApprPaper = 'Y' then 0"
        sql += "when cs.SupplyID='2' and pp.TNum <> 0 and cs.BudgetID = '03' and cs.IsApprPaper = 'Y' then  isnull(pp.TotalCost,0)/isnull(pp.TNum,1) end ) as cost1," & vbCrLf
        '/*預估補助費-就安*/
        sql += "sum( case when isnull(cs.SupplyID,0)=0 or pp.TNum = 0 then 0" & vbCrLf
        sql += "when cs.SupplyID='1' and pp.TNum <> 0 and cs.BudgetID = '02' and cs.IsApprPaper = 'Y' then ( isnull(pp.TotalCost,0)/isnull(pp.TNum,1))*0.8 " & vbCrLf
        sql += "when cs.SupplyID='9' and pp.TNum <> 0 and cs.BudgetID = '02' and cs.IsApprPaper = 'Y' then 0"
        sql += "when cs.SupplyID='2' and pp.TNum <> 0 and cs.BudgetID = '02' and cs.IsApprPaper = 'Y' then  isnull(pp.TotalCost,0)/isnull(pp.TNum,1) end ) as cost2," & vbCrLf
        '/*預估補助費-公務*/
        sql += "sum( case when isnull(cs.SupplyID,0)=0 or pp.TNum = 0 then 0" & vbCrLf
        sql += "when cs.SupplyID='1' and pp.TNum <> 0 and cs.BudgetID = '01' and cs.IsApprPaper = 'Y' then ( isnull(pp.TotalCost,0)/isnull(pp.TNum,1))*0.8" & vbCrLf
        sql += "when cs.SupplyID='9' and pp.TNum <> 0 and cs.BudgetID = '01' and cs.IsApprPaper = 'Y' then 0"
        sql += "when cs.SupplyID='2' and pp.TNum <> 0 and cs.BudgetID = '01' and cs.IsApprPaper = 'Y' then  isnull(pp.TotalCost,0)/isnull(pp.TNum,1) end ) as cost3," & vbCrLf
        '/*預估補助費-協助*/
        sql += "sum( case when isnull(cs.SupplyID,0)=0 or pp.TNum = 0 then 0" & vbCrLf
        sql += "when cs.SupplyID='1' and pp.TNum <> 0 and cs.BudgetID = '97' and cs.IsApprPaper = 'Y' then ( isnull(pp.TotalCost,0)/isnull(pp.TNum,1))*0.8" & vbCrLf
        sql += "when cs.SupplyID='9' and pp.TNum <> 0 and cs.BudgetID = '97' and cs.IsApprPaper = 'Y' then 0"
        sql += "when cs.SupplyID='2' and pp.TNum <> 0 and cs.BudgetID = '97' and cs.IsApprPaper = 'Y' then  isnull(pp.TotalCost,0)/isnull(pp.TNum,1) end ) as cost97," & vbCrLf
        '/*預估補助費-合計*/
        sql += "sum( case when isnull(cs.SupplyID,0)=0 or pp.TNum = 0 then 0" & vbCrLf
        sql += "when cs.SupplyID='1' and pp.TNum <> 0 and cs.BudgetID IN('01','02','03','97') and cs.IsApprPaper = 'Y' then ( isnull(pp.TotalCost,0)/isnull(pp.TNum,1))*0.8 " & vbCrLf
        sql += "when cs.SupplyID='9' and pp.TNum <> 0 and cs.BudgetID IN('01','02','03','97') and cs.IsApprPaper = 'Y' then 0"
        sql += "when cs.SupplyID='2' and pp.TNum <> 0 and cs.BudgetID IN('01','02','03','97') and cs.IsApprPaper = 'Y' then  isnull(pp.TotalCost,0)/isnull(pp.TNum,1) end ) as costAll," & vbCrLf

        '/*結訓-就保人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate() and cs.BudgetID = '03' " & vbCrLf
        sql += "then 1 else 0 end ) as closestudcout03," & vbCrLf

        '/*結訓-就安人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate() and cs.BudgetID = '02' " & vbCrLf
        sql += "then 1 else 0 end ) as closestudcout02," & vbCrLf

        '/*結訓-公務人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate() and cs.BudgetID = '01' " & vbCrLf
        sql += "then 1 else 0 end ) as closestudcout01," & vbCrLf

        '/*結訓-協助人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate() and cs.BudgetID = '97' " & vbCrLf
        sql += "then 1 else 0 end ) as closestudcout97," & vbCrLf

        '/*結訓-合計人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate() " & vbCrLf
        sql += "then 1 else 0 end ) as closestudcoutall," & vbCrLf

        '/*離訓人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.StudStatus = 2 " & vbCrLf
        sql += "then 1 else 0 end ) as std_cnt2," & vbCrLf
        '/*退訓人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.StudStatus = 3 " & vbCrLf
        sql += "then 1 else 0 end ) as std_cnt3," & vbCrLf
        '/*就保合計撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "then 1 else 0 end ) as budcountall," & vbCrLf

        '/*就安合計撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "then 1 else 0 end ) as budcountall2," & vbCrLf

        '/*公務合計撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "then 1 else 0 end ) as budcountall3," & vbCrLf

        '/*協助合計撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "then 1 else 0 end ) as budcountall97," & vbCrLf

        '/*就保一般身分學員撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count, " & vbCrLf

        '/*就保特殊身分學員撥款人次-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count07, " & vbCrLf

        '/*就保特殊身分學員撥款人次-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count05, " & vbCrLf

        '/*就保特殊身分學員撥款人次-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count06, " & vbCrLf

        '/*就保特殊身分學員撥款人次-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count04, " & vbCrLf
        '/*就保特殊身分學員撥款人次-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count28, " & vbCrLf

        '/*就保特殊身分學員撥款人次-更生保護者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count10, " & vbCrLf

        '/*就保特殊身分學員撥款人次-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then 1 else 0 end ) as bud03count26, " & vbCrLf

        '/*就保特殊身分學員撥款人次-加總*/
        'sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        'sql += "and cs.CreditPoints is not NULL " & vbCrLf
        'sql += "and ss.AppliedStatus = '1'" & vbCrLf
        'sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        'sql += "and cc.FTDate < getdate()" & vbCrLf
        'sql += "and cs.BudgetID = '03' " & vbCrLf
        'sql += "and cs.MIdentityID ='26'" & vbCrLf
        'sql += "then 1 else 0 end ) as bud03countSall, " & vbCrLf

        '/*就安一般身分學員撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count, " & vbCrLf

        '/*就安特殊身分學員撥款人次-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count07, " & vbCrLf

        '/*就安特殊身分學員撥款人次-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count05, " & vbCrLf

        '/*就安特殊身分學員撥款人次-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count06, " & vbCrLf

        '/*就安特殊身分學員撥款人次-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count04, " & vbCrLf

        '/*就安特殊身分學員撥款人次-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count28, " & vbCrLf

        '/*就安特殊身分學員撥款人次-更生保護者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count10, " & vbCrLf

        '/*就安特殊身分學員撥款人次-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then 1 else 0 end ) as bud02count26, " & vbCrLf

        '/*公務一般身分學員撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count, " & vbCrLf

        '/*公務特殊身分學員撥款人次-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count07, " & vbCrLf

        '/*公務特殊身分學員撥款人次-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count05, " & vbCrLf

        '/*公務特殊身分學員撥款人次-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count06, " & vbCrLf

        '/*公務特殊身分學員撥款人次-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count04, " & vbCrLf

        '/*公務特殊身分學員撥款人次-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count28, " & vbCrLf

        '/*公務特殊身分學員撥款人次-更生保護者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count10, " & vbCrLf

        '/*公務特殊身分學員撥款人次-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then 1 else 0 end ) as bud01count26, " & vbCrLf

        '/*協助一般身分學員撥款人次*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count, " & vbCrLf

        '/*協助特殊身分學員撥款人次-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count07, " & vbCrLf

        '/*協助特殊身分學員撥款人次-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count05, " & vbCrLf

        '/*協助特殊身分學員撥款人次-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count06, " & vbCrLf

        '/*協助特殊身分學員撥款人次-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count04, " & vbCrLf
        '/*協助特殊身分學員撥款人次-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count28, " & vbCrLf

        '/*協助特殊身分學員撥款人次-更生保護者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count10, " & vbCrLf

        '/*協助特殊身分學員撥款人次-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then 1 else 0 end ) as bud97count26, " & vbCrLf

        '/*就保合計撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as budmoneyall, " & vbCrLf

        '/*就安合計撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as budmoneyall2, " & vbCrLf

        '/*公務合計撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as budmoneyall3, " & vbCrLf

        '/*協助合計撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as budmoneyall97, " & vbCrLf

        '/*就保一般身分學員撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money, " & vbCrLf

        '/*就保特殊身分學員撥款金額-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money07,  " & vbCrLf

        '/*就保特殊身分學員撥款金額-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money05,  " & vbCrLf

        '/*就保特殊身分學員撥款金額-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money06,  " & vbCrLf

        '/*就保特殊身分學員撥款金額-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money04,  " & vbCrLf

        '/*就保特殊身分學員撥款金額-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money28,  " & vbCrLf

        '/*就保特殊身分學員撥款金額-更生人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money10, " & vbCrLf

        '/*就保特殊身分學員撥款金額-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '03' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud03money26, " & vbCrLf

        '/*就安一般身分學員撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money07,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money05,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money06,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money04,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money28,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-更生人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money10,  " & vbCrLf

        '/*就安特殊身分學員撥款金額-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '02' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud02money26,  " & vbCrLf

        '/*公務一般身分學員撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money, " & vbCrLf

        '/*公務特殊身分學員撥款金額-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money07, " & vbCrLf

        '/*公務特殊身分學員撥款金額-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money05," & vbCrLf

        '/*公務特殊身分學員撥款金額-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money06," & vbCrLf

        '/*公務特殊身分學員撥款金額-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money04," & vbCrLf

        '/*公務特殊身分學員撥款金額-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money28," & vbCrLf

        '/*公務特殊身分學員撥款金額-更生人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money10," & vbCrLf

        '/*公務特殊身分學員撥款金額-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '01' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud01money26," & vbCrLf

        '/*協助一般身分學員撥款金額*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='01'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money, " & vbCrLf

        '/*協助特殊身分學員撥款金額-生活扶助戶*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='07'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money07,  " & vbCrLf

        '/*協助特殊身分學員撥款金額-原住民*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y'" & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='05'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money05,  " & vbCrLf

        '/*協助特殊身分學員撥款金額-身心障礙者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='06'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money06,  " & vbCrLf

        '/*協助特殊身分學員撥款金額-中高齡*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='04'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money04,  " & vbCrLf

        '/*協助特殊身分學員撥款金額-獨力負擔家計者*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='28'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money28,  " & vbCrLf

        '/*協助特殊身分學員撥款金額-更生人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='10'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money10, " & vbCrLf

        '/*協助特殊身分學員撥款金額-犯罪被害人*/
        sql += "sum (case when cc.NotOpen = 'N' and cs.IsApprPaper = 'Y' " & vbCrLf
        sql += "and cs.CreditPoints is not NULL " & vbCrLf
        sql += "and ss.AppliedStatus = '1'" & vbCrLf
        sql += "and cs.StudStatus Not IN (2,3) " & vbCrLf
        sql += "and cc.FTDate < getdate()" & vbCrLf
        sql += "and cs.BudgetID = '97' " & vbCrLf
        sql += "and cs.MIdentityID ='26'" & vbCrLf
        sql += "then ss.SumOfMoney else 0 end ) as bud97money26 " & vbCrLf

        sql += " from" & vbCrLf
        sql += " Key_OrgType ky" & vbCrLf
        sql += " Join" & vbCrLf
        sql += " dbo.Org_OrgInfo oo" & vbCrLf
        sql += " on oo.orgkind = ky.orgTypeid" & vbCrLf
        '單位屬性  
        sql += " left join  Key_OrgType1 o1  on  oo.OrgKind1=o1.OrgTypeID1 " & vbCrLf
        sql += " join Plan_PlanInfo pp" & vbCrLf
        sql += " on pp.comidno = oo.comidno" & vbCrLf
        sql += " join ID_Plan ip on ip.planid = pp.planid" & vbCrLf
        sql += " join Auth_Relship ar on pp.rid = ar.rid" & vbCrLf
        sql += " join Org_OrgPlanInfo op on op.RSID = ar.RSID " & vbCrLf
        'sql += " join key_traintype kt on kt.TMID = pp.TMID " & vbCrLf
        sql += " left join ID_GovClassCast ig on pp.GCID = ig.GCID" & vbCrLf
        sql += " left join view_Depot01 vd1 on vd1.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot02 vd2 on vd2.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot03 vd3 on vd3.GCID = pp.GCID " & vbCrLf
        sql += " join Key_ClassCatelog kc" & vbCrLf
        sql += " on pp.ClassCate = kc.CCID" & vbCrLf
        sql += " left join class_classinfo cc" & vbCrLf
        sql += " on pp.planid = cc.planid and pp.rid = cc.rid and pp.seqno = cc.seqno" & vbCrLf
        sql += " left join ID_Class f ON cc.CLSID=f.CLSID " & vbCrLf
        sql += " left join Class_StudentsOfClass cs on cs.ocid = cc.ocid" & vbCrLf
        sql += " left join iD_Zip iz on pp.TaddressZip = iz.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic on ic.CTID = iz.CTID" & vbCrLf
        sql += " left join iD_Zip iz2 on op.ZipCode = iz2.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic2 on ic2.CTID = iz2.CTID" & vbCrLf
        sql += " left join Stud_SubsidyCost ss on cs.socid = ss.socid" & vbCrLf
        sql += " left join Plan_TrainPlace pp1 on pp.AddressSciPTID =pp1.PTID " & vbCrLf
        sql += " left join Plan_TrainPlace pp2 on pp.AddressTechPTID =pp2.PTID " & vbCrLf
        sql += " left join iD_Zip iz3 on pp1.ZipCode = iz3.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic3 on ic3.CTID = iz3.CTID" & vbCrLf
        sql += " left join iD_Zip iz4 on pp2.ZipCode = iz4.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic4 on ic4.CTID = iz4.CTID" & vbCrLf
        sql += " where" & vbCrLf
        sql += " 1=1" & vbCrLf
        sql += " and pp.Tplanid = '28'" & vbCrLf
        sql += " and pp.IsApprPaper ='Y'" & vbCrLf
        sql += strSearch & vbCrLf
        sql += " group by " & vbCrLf
        sql += " o1.typeid2,o1.typeid2name," & vbCrLf
        sql += " ip.DistID," & vbCrLf
        sql += " cc.ocid," & vbCrLf
        sql += " cc.Years," & vbCrLf
        sql += " cc.CyclType," & vbCrLf
        sql += " f.ClassID," & vbCrLf
        sql += " oo.orgname, " & vbCrLf
        sql += " pp.planid," & vbCrLf
        sql += " pp.comidno," & vbCrLf
        sql += " pp.seqno," & vbCrLf
        sql += " pp.ClassName," & vbCrLf
        sql += " pp.THours," & vbCrLf
        sql += " pp.AppliedResult, " & vbCrLf
        sql += " pp.STDate, " & vbCrLf
        sql += " pp.FDDate, " & vbCrLf
        sql += " pp.DefGovCost," & vbCrLf
        sql += " pp.TNum,pp.PlanYear," & vbCrLf
        sql += " pp.tmid,pp.CyclType,pp.PackageType," & vbCrLf
        sql += " vd1.kname,vd2.kname,vd3.kname,ig.GovClass," & vbCrLf
        sql += " kc.CCName,ic.CTName,ic3.CTName,ic4.CTName," & vbCrLf
        sql += " ig.GCode1,ig.GCode2" & vbCrLf
        sql += ") a"

        sql += " Left Join" & vbCrLf

        sql += " (" & vbCrLf
        sql += " select" & vbCrLf
        sql += " cc.ocid," & vbCrLf
        '/*累計訪視異常次數*/
        sql += " sum (case when cu.LItem2 ='Y' then 1 else 0 end) as vitN," & vbCrLf
        '/*累計不預告訪視次數*/
        sql += " (select count(cu1.ocid) from Class_UnexpectVisitor cu1 where cu1.ocid = cc.ocid) as cuall " & vbCrLf
        sql += " from" & vbCrLf
        sql += " Plan_PlanInfo pp" & vbCrLf
        sql += " join dbo.Org_OrgInfo oo" & vbCrLf
        sql += " on pp.comidno = oo.comidno" & vbCrLf
        sql += " join ID_Plan ip on ip.planid = pp.planid" & vbCrLf
        sql += " join Auth_Relship ar on pp.rid = ar.rid" & vbCrLf
        sql += " join Org_OrgPlanInfo op on ar.RSID = op.RSID " & vbCrLf
        'sql += " join key_traintype kt on kt.TMID = pp.TMID " & vbCrLf
        sql += " join ID_GovClassCast ig on pp.GCID = ig.GCID" & vbCrLf
        sql += " left join view_Depot01 vd1 on vd1.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot02 vd2 on vd2.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot03 vd3 on vd3.GCID = pp.GCID " & vbCrLf
        sql += " join class_classinfo cc on pp.planid = cc.planid and pp.rid = cc.rid and pp.seqno = cc.seqno" & vbCrLf
        sql += " left join Class_StudentsOfClass cs on cs.ocid = cc.ocid" & vbCrLf
        sql += " join Class_UnexpectVisitor cu on cc.ocid = cu.ocid" & vbCrLf
        sql += " left join iD_Zip iz on pp.TaddressZip = iz.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic on ic.CTID = iz.CTID" & vbCrLf
        sql += " left join iD_Zip iz2 on op.ZipCode = iz2.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic2 on ic2.CTID = iz2.CTID" & vbCrLf
        sql += " left join Stud_SubsidyCost ss on cs.socid = ss.socid" & vbCrLf
        sql += " left join Plan_TrainPlace pp1 on pp.AddressSciPTID =pp1.PTID " & vbCrLf
        sql += " left join Plan_TrainPlace pp2 on pp.AddressTechPTID =pp2.PTID " & vbCrLf
        sql += " left join iD_Zip iz3 on pp1.ZipCode = iz3.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic3 on ic3.CTID = iz3.CTID" & vbCrLf
        sql += " left join iD_Zip iz4 on pp2.ZipCode = iz4.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic4 on ic4.CTID = iz4.CTID" & vbCrLf
        sql += " where 1 = 1" & vbCrLf
        sql += " and pp.Tplanid = '28'" & vbCrLf
        sql += " and pp.IsApprPaper ='Y'" & vbCrLf
        sql += strSearch & vbCrLf

        sql += " group by " & vbCrLf
        sql += " cc.ocid " & vbCrLf
        sql += " ) e on a.ocid =e.ocid" & vbCrLf

        sql += " Left Join" & vbCrLf

        sql += " (select" & vbCrLf
        sql += " cc.ocid," & vbCrLf
        '/*累計訪視異常次數電話*/
        sql += " sum(case when ct.Item10 ='2' then 1 else 0 end) as VitTelN, " & vbCrLf
        '/*累計不預告訪視次數電話*/ 
        sql += " (select count(ct1.ocid) from Class_UnexpectTel ct1 where ct1.ocid =cc.ocid) as ctall " & vbCrLf
        sql += " from" & vbCrLf
        sql += " Plan_PlanInfo pp" & vbCrLf
        sql += " join dbo.Org_OrgInfo oo" & vbCrLf
        sql += " on pp.comidno = oo.comidno" & vbCrLf
        sql += " join ID_Plan ip on ip.planid = pp.planid" & vbCrLf
        sql += " join Auth_Relship ar on pp.rid = ar.rid" & vbCrLf
        sql += " join Org_OrgPlanInfo op on ar.RSID = op.RSID " & vbCrLf
        'sql += " join key_traintype kt on kt.TMID = pp.TMID " & vbCrLf
        sql += " join ID_GovClassCast ig on pp.GCID = ig.GCID" & vbCrLf
        sql += " left join view_Depot01 vd1 on vd1.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot02 vd2 on vd2.GCID = pp.GCID " & vbCrLf
        sql += " left join view_Depot03 vd3 on vd3.GCID = pp.GCID " & vbCrLf
        sql += " join class_classinfo cc on pp.planid = cc.planid and pp.rid = cc.rid and pp.seqno = cc.seqno" & vbCrLf
        sql += " left join Class_StudentsOfClass cs on cs.ocid = cc.ocid" & vbCrLf
        sql += " join Class_UnexpectTel ct on cc.ocid = ct.ocid" & vbCrLf
        sql += " left join iD_Zip iz on pp.TaddressZip = iz.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic on ic.CTID = iz.CTID" & vbCrLf
        sql += " left join iD_Zip iz2 on op.ZipCode = iz2.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic2 on ic2.CTID = iz2.CTID" & vbCrLf
        sql += " left join Stud_SubsidyCost ss on cs.socid = ss.socid" & vbCrLf
        sql += " left join Plan_TrainPlace pp1 on pp.AddressSciPTID =pp1.PTID " & vbCrLf
        sql += " left join Plan_TrainPlace pp2 on pp.AddressTechPTID =pp2.PTID " & vbCrLf
        sql += " left join iD_Zip iz3 on pp1.ZipCode = iz3.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic3 on ic3.CTID = iz3.CTID" & vbCrLf
        sql += " left join iD_Zip iz4 on pp2.ZipCode = iz4.ZipCode" & vbCrLf
        sql += " left join iD_CITY ic4 on ic4.CTID = iz4.CTID" & vbCrLf
        sql += " where 1 = 1" & vbCrLf

        sql += " and pp.Tplanid = '28'" & vbCrLf
        sql += " and pp.IsApprPaper ='Y'" & vbCrLf
        sql += strSearch & vbCrLf

        sql += " group by " & vbCrLf
        sql += " cc.ocid ) f" & vbCrLf

        sql += " on a.ocid = f.ocid" & vbCrLf

        da.SelectCommand.CommandText = sql
        da.SelectCommand.Parameters.Clear()
        da.Fill(dt)

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("綜合查詢統計表", System.Text.Encoding.UTF8) & ".xls")
        Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        If ChbExit.SelectedIndex = -1 Then
            ExportStr = "職訓中心" & vbTab & "單位屬性" & vbTab & "訓練機構" & vbTab & "班別名稱" & vbTab & "期別" & vbTab & "課程代碼" & vbTab & "開訓日期" & vbTab & "結訓日期" & vbTab
            ExportStr = ExportStr & "新興產業" & vbTab & "重點服務業" & vbTab & "新興智慧型產業" & vbTab & "訓練業別編碼" & vbTab & "訓練職能" & vbTab    '20100114 add andy
            ExportStr = ExportStr & "學科辦訓地縣市" & vbTab & "術科辦訓地縣市" & vbTab & "包班總類" & vbTab
        Else
            ExportStr = "職訓中心" & vbTab & "單位屬性" & vbTab & "訓練機構" & vbTab & "班別名稱" & vbTab & "期別" & vbTab & "課程代碼" & vbTab & "開訓日期" & vbTab & "結訓日期" & vbTab
            ExportStr = ExportStr & "新興產業" & vbTab & "重點服務業" & vbTab & "新興智慧型產業" & vbTab & "訓練業別編碼" & vbTab & "訓練職能" & vbTab     '20100114 add andy
            ExportStr = ExportStr & "學科辦訓地縣市" & vbTab & "術科辦訓地縣市" & vbTab & "包班總類" & vbTab
            For a As Integer = 0 To Me.ChbExit.Items.Count - 1
                If ChbExit.Items.Item(a).Selected Then
                    Select Case ChbExit.Items.Item(a).Value
                        Case "1"
                            ExportStr = ExportStr & "申請人次" & vbTab
                        Case "2"
                            ExportStr = ExportStr & "申請補助費" & vbTab
                        Case "3"
                            ExportStr = ExportStr & "核定人次" & vbTab
                        Case "4"
                            ExportStr = ExportStr & "核定補助費" & vbTab
                        Case "5"
                            ExportStr = ExportStr & "實際就保開訓人次" & vbTab & "實際就安開訓人次" & vbTab & "實際公務開訓人次" & vbTab & "實際協助開訓人次" & vbTab
                        Case "6"
                            ExportStr = ExportStr & "實際合計開訓人次" & vbTab
                        Case "7"
                            ExportStr = ExportStr & "就保預估補助費金額" & vbTab & "就安預估補助費金額" & vbTab & "公務預估補助費金額" & vbTab & "協助預估補助費金額" & vbTab
                        Case "8"
                            ExportStr = ExportStr & "合計預估補助費金額" & vbTab
                        Case "9" '合計結訓人次
                            ExportStr = ExportStr & "就保結訓人次" & vbTab & "就安結訓人次" & vbTab & "公務結訓人次" & vbTab & "協助結訓人次" & vbTab & "合計結訓人次" & vbTab
                        Case "10"
                            ExportStr = ExportStr & "就保撥款人次" & vbTab & "就安撥款人次" & vbTab & "公務撥款人次" & vbTab & "協助撥款人次" _
                                        & vbTab & "就保一般身分撥款人次" & vbTab & "就保特殊身分(生活扶助戶)撥款人次" & vbTab & "就保特殊身分(原住民)撥款人次" & vbTab & "就保特殊身分(身心障礙者)撥款人次" & vbTab & "就保特殊身分(中高齡)撥款人次" & vbTab & "就保特殊身分(獨力負擔家計者)撥款人次" & vbTab & "就保特殊身分(更生受保護者)撥款人次" & vbTab & "就保特殊身分(犯罪被害人及其親屬)撥款人次" & vbTab & "就保特殊身分總撥款人次" _
                                        & vbTab & "就安一般身分撥款人次" & vbTab & "就安特殊身分(生活扶助戶)撥款人次" & vbTab & "就安特殊身分(原住民)撥款人次" & vbTab & "就安特殊身分(身心障礙者)撥款人次" & vbTab & "就安特殊身分(中高齡)撥款人次" & vbTab & "就安特殊身分(獨力負擔家計者)撥款人次" & vbTab & "就安特殊身分(更生受保護者)撥款人次" & vbTab & "就安特殊身分(犯罪被害人及其親屬)撥款人次" & vbTab & "就安特殊身分總撥款人次" _
                                        & vbTab & "公務一般身分撥款人次" & vbTab & "公務特殊身分(生活扶助戶)撥款人次" & vbTab & "公務特殊身分(原住民)撥款人次" & vbTab & "公務特殊身分(身心障礙者)撥款人次" & vbTab & "公務特殊身分(中高齡)撥款人次" & vbTab & "公務特殊身分(獨力負擔家計者)撥款人次" & vbTab & "公務特殊身分(更生受保護者)撥款人次" & vbTab & "公務特殊身分(犯罪被害人及其親屬)撥款人次" & vbTab & "公務特殊身分總撥款人次" _
                                        & vbTab & "協助一般身分撥款人次" & vbTab & "協助特殊身分(生活扶助戶)撥款人次" & vbTab & "協助特殊身分(原住民)撥款人次" & vbTab & "協助特殊身分(身心障礙者)撥款人次" & vbTab & "協助特殊身分(中高齡)撥款人次" & vbTab & "協助特殊身分(獨力負擔家計者)撥款人次" & vbTab & "協助特殊身分(更生受保護者)撥款人次" & vbTab & "協助特殊身分(犯罪被害人及其親屬)撥款人次" & vbTab & "協助特殊身分總撥款人次" & vbTab
                        Case "11"
                            ExportStr = ExportStr & "就保撥款補助費" & vbTab & "就安撥款補助費" & vbTab & "公務撥款補助費" & vbTab & "協助撥款補助費" _
                                        & vbTab & "就保一般身分撥款補助費" & vbTab & "就保特殊身分(生活扶助戶)撥款補助費" & vbTab & "就保特殊身分(原住民)撥款補助費" & vbTab & "就保特殊身分(身心障礙者)撥款補助費" & vbTab & "就保特殊身分(中高齡)撥款補助費" & vbTab & "就保特殊身分(獨力負擔家計者)撥款補助費" & vbTab & "就保特殊身分(更生受保護者)撥款補助費" & vbTab & "就保特殊身分(犯罪被害人及其親屬)撥款補助費" & vbTab & "就保特殊身分總撥款補助費" _
                                        & vbTab & "就安一般身分撥款補助費" & vbTab & "就安特殊身分(生活扶助戶)撥款補助費" & vbTab & "就安特殊身分(原住民)撥款補助費" & vbTab & "就安特殊身分(身心障礙者)撥款補助費" & vbTab & "就安特殊身分(中高齡)撥款補助費" & vbTab & "就安特殊身分(獨力負擔家計者)撥款補助費" & vbTab & "就安特殊身分(更生受保護者)撥款補助費" & vbTab & "就安特殊身分(犯罪被害人及其親屬)撥款補助費" & vbTab & "就安特殊身分總撥款補助費" _
                                        & vbTab & "公務一般身分撥款補助費" & vbTab & "公務特殊身分(生活扶助戶)撥款補助費" & vbTab & "公務特殊身分(原住民)撥款補助費" & vbTab & "公務特殊身分(身心障礙者)撥款補助費" & vbTab & "公務特殊身分(中高齡)撥款補助費" & vbTab & "公務特殊身分(獨力負擔家計者)撥款補助費" & vbTab & "公務特殊身分(更生受保護者)撥款補助費" & vbTab & "公務特殊身分(犯罪被害人及其親屬)撥款補助費" & vbTab & "公務特殊身分總撥款補助費" _
                                        & vbTab & "協助一般身分撥款補助費" & vbTab & "協助特殊身分(生活扶助戶)撥款補助費" & vbTab & "協助特殊身分(原住民)撥款補助費" & vbTab & "協助特殊身分(身心障礙者)撥款補助費" & vbTab & "協助特殊身分(中高齡)撥款補助費" & vbTab & "協助特殊身分(獨力負擔家計者)撥款補助費" & vbTab & "協助特殊身分(更生受保護者)撥款補助費" & vbTab & "協助特殊身分(犯罪被害人及其親屬)撥款補助費" & vbTab & "協助特殊身分總撥款補助費" & vbTab
                        Case "12"
                            ExportStr = ExportStr & "累計不預告實地抽訪次數" & vbTab
                        Case "13"
                            ExportStr = ExportStr & "累計不預告電話抽訪次數" & vbTab
                        Case "14"
                            ExportStr = ExportStr & "累積訪視異常次數" & vbTab
                        Case "15"
                            ExportStr = ExportStr & "會計查帳次數" & vbTab
                        Case "16"
                            ExportStr = ExportStr & "離訓人次" & vbTab
                        Case "17"
                            ExportStr = ExportStr & "退訓人次" & vbTab
                            'Case "18"
                            '    ExportStr = ExportStr & "訓練職能" & vbTab
                        Case "18"
                            ExportStr = ExportStr & "訓練時數" & vbTab
                        Case "19"
                            ExportStr = ExportStr & "人時成本" & vbTab
                        Case "20"
                            ExportStr = ExportStr & "上課時間" & vbTab
                    End Select
                End If
            Next

        End If

        ExportStr += vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        For Each dr In dt.Rows
            ExportStr = ""
            If ChbExit.SelectedIndex = -1 Then
                ExportStr = ExportStr & dr("DistName") & vbTab  '職訓中心
                ExportStr = ExportStr & Convert.ToString(dr("OrgTypeName")) & vbTab    '單位屬性
                ExportStr = ExportStr & dr("orgname") & vbTab  '訓練機構
                ExportStr = ExportStr & dr("ClassName") & vbTab  '班別名稱
                ExportStr = ExportStr & dr("CyclType") & vbTab  '期別
                ExportStr = ExportStr & dr("ClassID") & vbTab  '課程代碼
                ExportStr = ExportStr & dr("STDate") & vbTab  '開訓日期
                ExportStr = ExportStr & dr("FDDate") & vbTab  '結訓日期

                ExportStr = ExportStr & dr("kname1") & vbTab      '六大   '20100114 add andy
                ExportStr = ExportStr & dr("kname2") & vbTab      '十大
                ExportStr = ExportStr & dr("kname3") & vbTab      '四大
                ExportStr = ExportStr & dr("GCodeName") & vbTab      '訓練業別編碼
                ExportStr = ExportStr & Convert.ToString(dr("CCName")) & vbTab   '訓練職能
                'If Convert.ToString(dr("GovClass")) = "1" Then   '訓練業別編碼
                '    ExportStr = ExportStr & "院"
                'ElseIf Convert.ToString(dr("GovClass")) = "2" Then
                '    ExportStr = ExportStr & "局"
                'End If
                ''ExportStr = ExportStr & "-" & Format(dr("GCode1"), "#0") & vbTab
                'ExportStr = ExportStr & "-" & dr("GCode1") & "-" & dr("GCode2") & vbTab
                ExportStr = ExportStr & dr("AddressSciPTID") & vbTab  '學科辦訓地縣市
                ExportStr = ExportStr & dr("AddressTechPTID") & vbTab '術科辦訓地縣市
                ExportStr = ExportStr & dr("PackageType") & vbTab  '包班種類
            Else
                ExportStr = ExportStr & dr("DistName") & vbTab  '職訓中心
                ExportStr = ExportStr & Convert.ToString(dr("OrgTypeName")) & vbTab    '單位屬性
                ExportStr = ExportStr & dr("orgname") & vbTab  '訓練機構
                ExportStr = ExportStr & dr("ClassName") & vbTab  '班別名稱
                ExportStr = ExportStr & dr("CyclType") & vbTab  '期別
                ExportStr = ExportStr & dr("ClassID") & vbTab  '課程代碼
                ExportStr = ExportStr & dr("STDate") & vbTab  '開訓日期
                ExportStr = ExportStr & dr("FDDate") & vbTab  '結訓日期

                ExportStr = ExportStr & dr("kname1") & vbTab      '六大   '20100114 add andy
                ExportStr = ExportStr & dr("kname2") & vbTab      '十大
                ExportStr = ExportStr & dr("kname3") & vbTab      '四大
                ExportStr = ExportStr & dr("GCodeName") & vbTab      '訓練業別編碼
                ExportStr = ExportStr & Convert.ToString(dr("CCName")) & vbTab   '訓練職能
                'If Convert.ToString(dr("GovClass")) = "1" Then      '
                '    ExportStr = ExportStr & "院"
                'ElseIf Convert.ToString(dr("GovClass")) = "2" Then
                '    ExportStr = ExportStr & "局"
                'End If
                ''ExportStr = ExportStr & "-" & Format(dr("GCode1"), "#0") & vbTab
                'ExportStr = ExportStr & "-" & dr("GCode1") & "-" & dr("GCode2") & vbTab
                ExportStr = ExportStr & dr("AddressSciPTID") & vbTab  '學科辦訓地縣市
                ExportStr = ExportStr & dr("AddressTechPTID") & vbTab '術科辦訓地縣市
                ExportStr = ExportStr & dr("PackageType") & vbTab  '包班種類

                For j As Integer = 0 To Me.ChbExit.Items.Count - 1
                    If Me.ChbExit.Items(j).Selected Then
                        Select Case ChbExit.Items.Item(j).Value
                            Case "1"
                                ExportStr = ExportStr & dr("ATNum") & vbTab  '申請人數
                            Case "2"
                                ExportStr = ExportStr & dr("ADefGovCost") & vbTab  '申請補助費
                            Case "3"
                                ExportStr = ExportStr & dr("TNum") & vbTab  '核定人數
                            Case "4"
                                ExportStr = ExportStr & dr("DefGovCost") & vbTab '核定補助費
                            Case "5"
                                ExportStr = ExportStr & dr("openstudcount1") & vbTab & dr("openstudcount2") & vbTab & dr("openstudcount3") & vbTab & dr("openstudcount97") & vbTab  '開訓人次
                            Case "6"
                                ExportStr = ExportStr & dr("openstudcountall") & vbTab  '開訓人次加總
                            Case "7"
                                ExportStr = ExportStr & dr("cost1") & vbTab & dr("cost2") & vbTab & dr("cost3") & vbTab & dr("cost97") & vbTab '預估補助費
                            Case "8"
                                ExportStr = ExportStr & dr("costAll") & vbTab '預估補助費加總
                            Case "9" '合計結訓人次
                                ExportStr = ExportStr & dr("closestudcout03") & vbTab & dr("closestudcout02") & vbTab & dr("closestudcout01") & vbTab & dr("closestudcout97") & vbTab & dr("closestudcoutall") & vbTab '結訓人次
                            Case "10"
                                ExportStr = ExportStr & dr("budcountall") & vbTab & dr("budcountall2") & vbTab & dr("budcountall3") & vbTab & dr("budcountall97") & vbTab _
                                            & dr("bud03count") & vbTab & dr("bud03count07") & vbTab & dr("bud03count05") & vbTab & dr("bud03count06") & vbTab & dr("bud03count04") & vbTab & dr("bud03count28") & vbTab & dr("bud03count10") & vbTab & dr("bud03count26") & vbTab & dr("Sbudcount01") & vbTab _
                                            & dr("bud02count") & vbTab & dr("bud02count07") & vbTab & dr("bud02count05") & vbTab & dr("bud02count06") & vbTab & dr("bud02count04") & vbTab & dr("bud02count28") & vbTab & dr("bud02count10") & vbTab & dr("bud02count26") & vbTab & dr("Sbudcount02") & vbTab _
                                            & dr("bud01count") & vbTab & dr("bud01count07") & vbTab & dr("bud01count05") & vbTab & dr("bud01count06") & vbTab & dr("bud01count04") & vbTab & dr("bud01count28") & vbTab & dr("bud01count10") & vbTab & dr("bud01count26") & vbTab & dr("Sbudcount03") & vbTab _
                                            & dr("bud97count") & vbTab & dr("bud97count07") & vbTab & dr("bud97count05") & vbTab & dr("bud97count06") & vbTab & dr("bud97count04") & vbTab & dr("bud97count28") & vbTab & dr("bud97count10") & vbTab & dr("bud97count26") & vbTab & dr("Sbudcount97") & vbTab   '撥款人次
                            Case "11"
                                ExportStr = ExportStr & dr("budmoneyall") & vbTab & dr("budmoneyall2") & vbTab & dr("budmoneyall3") & vbTab & dr("budmoneyall97") & vbTab _
                                            & dr("bud03money") & vbTab & dr("bud03money07") & vbTab & dr("bud03money05") & vbTab & dr("bud03money06") & vbTab & dr("bud03money04") & vbTab & dr("bud03money28") & vbTab & dr("bud03money10") & vbTab & dr("bud03money26") & vbTab & dr("Sbudmoney01") & vbTab _
                                            & dr("bud02money") & vbTab & dr("bud02money07") & vbTab & dr("bud02money05") & vbTab & dr("bud02money06") & vbTab & dr("bud02money04") & vbTab & dr("bud02money28") & vbTab & dr("bud02money10") & vbTab & dr("bud02money26") & vbTab & dr("Sbudmoney02") & vbTab _
                                            & dr("bud01money") & vbTab & dr("bud01money07") & vbTab & dr("bud01money05") & vbTab & dr("bud01money06") & vbTab & dr("bud01money04") & vbTab & dr("bud01money28") & vbTab & dr("bud01money10") & vbTab & dr("bud01money26") & vbTab & dr("Sbudmoney03") & vbTab _
                                            & dr("bud97money") & vbTab & dr("bud97money07") & vbTab & dr("bud97money05") & vbTab & dr("bud97money06") & vbTab & dr("bud97money04") & vbTab & dr("bud97money28") & vbTab & dr("bud97money10") & vbTab & dr("bud97money26") & vbTab & dr("Sbudmoney97") & vbTab    '撥款補助費
                            Case "12"
                                ExportStr = ExportStr & dr("cuall") & vbTab '不預告訪視次數-實地抽訪
                            Case "13"
                                ExportStr = ExportStr & dr("ctall") & vbTab '不預告訪視次數-電話抽訪
                            Case "14"
                                ExportStr = ExportStr & dr("vtn") & vbTab    '累計訪視異常次數
                            Case "15"
                                ExportStr = ExportStr & "" & vbTab '會計查帳次數
                            Case "16"
                                ExportStr = ExportStr & Convert.ToString(dr("std_cnt2")) & vbTab   '離訓人次
                            Case "17"
                                ExportStr = ExportStr & Convert.ToString(dr("std_cnt3")) & vbTab   '退訓人次
                                'Case "18"
                                '    ExportStr = ExportStr & Convert.ToString(dr("CCName")) & vbTab   '訓練職能
                            Case "18"
                                ExportStr = ExportStr & Convert.ToString(dr("THours")) & vbTab   '訓練時數
                            Case "19"
                                ExportStr = ExportStr & Convert.ToString(dr("PhCost")) & vbTab   '人時成本
                            Case "20"
                                ExportStr = ExportStr & Convert.ToString(dr("WEEKSTIME")) & vbTab   '上課時間
                        End Select
                    End If
                Next
            End If

            ExportStr += vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Next
    End Sub

    '訓練機構選擇
    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim DistID1 As String
        Dim N As Integer
        Dim i As Integer
        Dim msg As String = ""

        If sm.UserInfo.DistID = "000" Then
            DistID1 = ""
            N = 0   '預設 N =0 表示沒有勾選轄區選項
            For i = 1 To Me.Distid.Items.Count - 1

                If Me.Distid.Items(i).Selected Then '假如有勾選
                    N = N + 1  '計算轄區勾選選項的數目
                    If N = 1 Then '如果是勾選一個選項
                        DistID1 = Convert.ToString(Me.Distid.Items(i).Value) '取得選項的值
                    End If
                    'If N = 2 Then '如果轄區勾選選項的數目=2
                    '    msg += "只能選擇一個轄區!" & vbCrLf
                    '    DistID1 = ""
                    '    Exit For
                    'End If
                End If
            Next

            If N = 0 Then '如果轄區選項沒有選
                msg += "請選擇轄區!" & vbCrLf
            End If

            If msg <> "" Then
                Turbo.Common.MessageBox(Me, msg)
            End If

        End If
    End Sub

    '匯出
    Private Sub BtnExp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnExp.Click
        Dim okFlag As Boolean = False
        Dim conn As New SqlConnection
        Try
            conn = DbAccess.GetConnection()
            If conn.State = ConnectionState.Closed Then conn.Open()

            Dim da As New SqlDataAdapter
            da.SelectCommand = New SqlCommand
            da.SelectCommand.Connection = conn
            da.SelectCommand.CommandTimeout = 100

            ExpRpt(da) '匯出SUB
            okFlag = True
            'da.Dispose()
            If conn.State = ConnectionState.Open Then conn.Close()
        Catch ex As Exception
            If conn.State = ConnectionState.Open Then conn.Close()
            Turbo.Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.ToString)
            'Me.Page.RegisterStartupScript("Errmsg", "<script>alert('發生錯誤:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        End Try

        If okFlag Then
            Response.End()
        End If
    End Sub
End Class


