Public Class menu
    Inherits AuthBasePage

    Const cst_errmsg1 As String = "請重新選擇計畫!!"
    Const cst_errmsg2 As String = "該使用者/計畫無此功能，請重新選擇計畫!!"

    Dim intKind As Integer = -1
    Dim tnCnt As Integer = 0
    Const cst_text_color As String = "blue"
    ''style
    ''class
    'select * from auth_groupacct where account ='k14631'

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Response.Cache.SetExpires(DateTime.Now())

        objconn = DbAccess.GetConnection
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(objconn) Then Exit Sub

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    labmsg.Text = cst_errmsg1 '"請重新選擇計畫!!"
        '    Exit Sub
        'End If

        If Not IsPostBack Then
            Try
                Call search()
            Catch ex As Exception
                labmsg.Text = cst_errmsg2 & ".pl" '"該計畫無此功能，請重新選擇計畫!!"
            End Try

            '取得或設定值，指出是否顯示展開節點指示器。
            TreeView1.ShowExpandCollapse = False

            If sm.FunctionPage <> "" Then
                '以程式設計方式展開節點
                Call TreeView1.ExpandAll()
            Else
                '關閉樹狀結構中的每個節點。收合節點
                Call TreeView1.CollapseAll()
            End If

        End If
    End Sub

    '建置功能列
    Private Sub search()

        '群組查詢
        Dim dt_group As DataTable = TIMS.sGet_CanUseSchDt() ' FROM 
        Dim blnGroupF As Boolean = True '查詢有資料。
        If dt_group Is Nothing Then
            blnGroupF = False '查詢無資料。
        Else
            dt_group.DefaultView.RowFilter = "Valid='Y'"
            dt_group = TIMS.dv2dt(dt_group.DefaultView)
            If dt_group.Rows.Count = 0 Then blnGroupF = False '查詢無資料。
        End If
        If Not blnGroupF Then
            labmsg.Text = cst_errmsg2 & ".ms" '""
            Exit Sub
        End If
        If blnGroupF Then
            '有資料執行顯示菜單功能。
            Call AddTreeView(dt_group)
        End If
    End Sub

    '檢查是否有功能子集。
    Function ChkKindDataTable(ByRef dt As DataTable, ByVal Kind As String) As Boolean
        Dim rst As Boolean = True
        labmsg.Text = ""
        If dt.Select("Kind='" & Kind & "' and Levels=0", "newSort").Length = 0 Then
            rst = False
            labmsg.Text = cst_errmsg2 & ".mc" '"該計畫無此功能，請重新選擇計畫!!"
        End If
        Return rst
    End Function

    '主menu顯示順序
    Private Sub AddTreeView(ByVal dt As DataTable)
        Select Case sm.FunctionPage
            Case "TC", "SD", "CP", "TR", "CM", "FM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"
                '檢查是否有功能子集。
                If Not ChkKindDataTable(dt, sm.FunctionPage) Then Exit Sub
                AddTitleNode(dt, sm.FunctionPage)
            Case Else
                AddTitleNode(dt, "TC")
                AddTitleNode(dt, "SD")
                AddTitleNode(dt, "CP")
                AddTitleNode(dt, "TR")
                AddTitleNode(dt, "CM")
                AddTitleNode(dt, "FM")
                AddTitleNode(dt, "OB")
                AddTitleNode(dt, "SE")
                AddTitleNode(dt, "EXAM")
                AddTitleNode(dt, "SV")
                AddTitleNode(dt, "SYS")
                AddTitleNode(dt, "FAQ")
                AddTitleNode(dt, "OO")
        End Select

    End Sub

    Private Sub AddTitleNode(ByVal dt As DataTable, ByVal Kind As String)
        If dt.Select("Kind='" & Kind & "' and Levels=0").Length <> 0 Or Kind = "FAQ" Then
            Dim MyNode As TreeNode = New TreeNode
            Dim vsStudExamUrl As String = ""

            '功能類別對照 取得中文名稱
            Dim rst As String = TIMS.Get_MainMenuName(UCase(Kind))
            If rst <> "" Then
                intKind += 1

                Dim strjs As String = ""
                strjs = "" & vbCrLf
                strjs += " <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""background-color@White"">" & vbCrLf
                strjs += " <tr style=""height:0px""><td width=""12px""></td><td width=""30px""></td><td width=""14px""></td><td width=""14px""></td><td width=""142px""></td></tr>" & vbCrLf
                'strjs += " <tr style=""height:36px; background-image@url(./images/i2/icon/left_bd2.gif); cursor@pointer"" onclick=""chkMenu('" & intKind & "');"">" & vbCrLf
                strjs += " <tr style=""height:36px; background-image@url(./images/i2/icon/left_bd2.gif); cursor@pointer"" onclick=""TreeView_ToggleNode(TreeView1_Data," & tnCnt & ", TreeView1n" & tnCnt & ", '', TreeView1n" & tnCnt & "Nodes); chkMenu(" & intKind & ");"">" & vbCrLf
                strjs += " <td></td>" & vbCrLf
                strjs += " <td>" & vbCrLf
                strjs += " <br style=""line-height:10px"" />" & vbCrLf
                strjs += " <img id=""img" & intKind & "_off"" src=""./images/i2/icon/left_off.gif"" alt="""" style=""display@inline"" />" & vbCrLf
                strjs += " <img id=""img" & intKind & "_on""  src=""./images/i2/icon/left_on.gif"" alt="""" style=""display@none"" />" & vbCrLf
                strjs += " </td>" & vbCrLf
                strjs += " <td colspan=""3"">" & vbCrLf
                strjs += " <br style=""line-height:8px"" />" & vbCrLf
                strjs += " <label id=""labFunName" & intKind & """ style=""font-size:18px; font-weight@bold; color: #5288b2"">" & rst & "</label>" & vbCrLf
                strjs += " </td>" & vbCrLf
                strjs += " </tr>" & vbCrLf
                strjs += " </table>" & vbCrLf
                MyNode.Text = strjs
                'TreeView1.ShowLines = False

                'MyNode.Text = rst
                ''images/i2/aimg01_cm.jpg
                ''images/menuback_01.gif
                ''MyNode.ImageUrl = "images/menuback_01.gif"
                ''MyNode.ImageUrl = "images\i2\aimg01_cm.jpg"
                ''MyNode.ImageUrl = "images/i2/icon/left_bd.gif"
                ' '' ''MyNode.ImageUrl = "images/i2/icon/left_bd.gif"
                ' '' ''MyNode.SelectAction = TreeNodeSelectAction.SelectExpand
            End If
            Select Case Kind
                Case "SE"
                    'vsStudExamUrl = ConfigurationManager.AppSettings("StudExamUrl")
                    vsStudExamUrl = ConfigurationSettings.AppSettings("StudExamUrl")
                    'MyNode.ImageUrl = "images/menuback_21.gif"
                    MyNode.NavigateUrl = vsStudExamUrl '& "?ID=" & dr("FunID").ToString
                    MyNode.Target = "_new"
            End Select


            'MyNode.ImageToolTip = rst
            'MyNode.SelectAction = TreeNodeSelectAction.Expand
            'Select Case Kind
            '    Case "TC"
            '        'MyNode.Parent.ImageUrl = "images/icon/left_bd.gif"
            '        'menuback_01
            '        MyNode.ImageUrl = "images/menuback_01.gif"
            '    Case "SD"
            '        MyNode.ImageUrl = "images/menuback_03.gif"
            '    Case "CP"
            '        MyNode.ImageUrl = "images/menuback_05.gif"
            '    Case "SYS"
            '        MyNode.ImageUrl = "images/menuback_07.gif"
            '    Case "TR"
            '        MyNode.ImageUrl = "images/menuback_12.gif"
            '    Case "CM"
            '        MyNode.ImageUrl = "images/menuback_18.gif"
            '    Case "FM"
            '        MyNode.ImageUrl = "images/menuback_24.gif"
            '    Case "SV"
            '        MyNode.ImageUrl = "images/menuback_22.gif"
            '    Case "SE"
            '        MyNode.ImageUrl = "images/menuback_21.gif"
            '    Case "EXAM"
            '        MyNode.ImageUrl = "images/menuback_20.gif"
            '    Case "OB"
            '        MyNode.ImageUrl = "images/menuback_19.gif"
            '    Case "FAQ"
            '        MyNode.ImageUrl = "images/menuback_11.gif"
            '    Case "OO"
            '        MyNode.ImageUrl = "images/menuback_23.gif"
            'End Select
            '用戶點擊文字可以展開節點
            MyNode.SelectAction = TreeNodeSelectAction.Expand
            TreeView1.Nodes.Add(MyNode)
            tnCnt += 1

            Select Case Kind
                Case "SE"
                Case Else
                    Call AddNode(dt, Kind, , MyNode)
            End Select
        End If
    End Sub

    '子節點樣式
    Private Sub AddNode(ByVal dt As DataTable, ByVal Kind As String, Optional ByVal drChild As DataRow = Nothing, Optional ByVal ParentNode As TreeNode = Nothing)
        Dim vsUserID As String = "" & Convert.ToString(sm.UserInfo.UserID)
        Dim vsDistID As String = "" & Convert.ToString(sm.UserInfo.DistID)
        Dim vsLID As String = "" & Convert.ToString(sm.UserInfo.LID)
        If vsLID <> "" Then vsLID = CInt(vsLID)

        Dim dr As DataRow
        Dim MyNode As TreeNode

        'Const cst_FunLength_set As Integer = 7 '10
        If drChild Is Nothing Then
            For Each dr In dt.Select("Kind='" & Kind & "' and Levels=0", "newSort")
                Const cst_FunLength_set As Integer = 9 '預計最大字串長度。
                Dim FunName As String = "" '上層功能名稱  Levels=0 
                Dim FunLength As Integer = 0 '功能字串長度。
                FunName = dr("Name").ToString
                If FunName <> "" Then FunName = Trim(FunName)
                FunLength = FunName.Length

                MyNode = New TreeNode
                If FunLength > cst_FunLength_set Then
                    Const cst_addSpace1 As String = "<BR>&nbsp;&nbsp;&nbsp;&nbsp;"
                    Dim i As Integer = 0
                    Dim mFunName As String = ""
                    While FunLength > 0
                        If mFunName = "" Then
                            mFunName = Mid(FunName, 1, cst_FunLength_set)
                        Else
                            mFunName &= cst_addSpace1 & Mid(FunName, i * cst_FunLength_set + 1, cst_FunLength_set)
                        End If
                        FunLength -= cst_FunLength_set '減少字串
                        i += 1
                    End While
                    FunName = mFunName '最後存進 FunName
                End If

                'If FunLength > cst_FunLength_set Then
                '    Dim i As Integer = 0
                '    While FunLength > 0
                '        If IsDBNull(dr("SPage")) Then
                '            If FunName = "" Then
                '                FunName = Mid(dr("Name"), 1, cst_FunLength_set)
                '            Else
                '                FunName += "<BR>" & Mid(dr("Name"), i * cst_FunLength_set + 1, cst_FunLength_set)
                '            End If
                '            FunLength -= cst_FunLength_set
                '            i += 1
                '        Else
                '            If FunName = "" Then
                '                FunName = Mid(dr("Name"), 1, cst_FunLength_set)
                '            Else
                '                FunName += "<BR>" & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Mid(dr("Name"), i * cst_FunLength_set + 1, cst_FunLength_set)
                '            End If
                '            FunLength -= cst_FunLength_set
                '            i += 1
                '        End If
                '    End While
                'Else
                '    FunName = dr("Name")
                'End If

                If IsDBNull(dr("SPage")) Then
                    '沒有連結
                    'MyNode.ImageUrl = "images/i2/minus.gif"
                    'MyNode.ImageUrl = "images/i2/plus.gif"
                    'MyNode.SelectAction = TreeNodeSelectAction.Expand
                    Dim strjs As String = ""
                    strjs = "" & vbCrLf
                    strjs += " <table width=""100%"" border=""0"" cellpadding=""0"" cellspacing=""0"" style=""background-color@White"">" & vbCrLf
                    strjs += " <tr style=""height:18px;"" onclick=""TreeView_ToggleNode(TreeView1_Data," & tnCnt & ", TreeView1n" & tnCnt & ", '', TreeView1n" & tnCnt & "Nodes);"">" & vbCrLf
                    strjs += " <td>" & vbCrLf
                    'strjs += " <br style=""line-height:8px"" />" & vbCrLf
                    strjs += " <label  style=""font-size:16px; font-weight@bold;"" ><img src='./images/point.gif' border='0' align='absmiddle'><font color='" & cst_text_color & "'>" & FunName & "</font></label>" & vbCrLf
                    strjs += " </td>" & vbCrLf
                    strjs += " </tr>" & vbCrLf
                    strjs += " </table>" & vbCrLf
                    MyNode.Text = strjs
                    'MyNode.Text = "<font color='" & cst_text_color & "'>" & FunName & "</font>"
                Else
                    '有連結
                    'MyNode.Text = "<img src='images/i2/point.gif' border='0' align='absmiddle'><font color='" & cst_text_color & "'>" & "&nbsp;&nbsp;" & FunName & "</font>"
                    MyNode.Text = "<img src='./images/point.gif' border='0' align='absmiddle'><font color='" & cst_text_color & "'>" & FunName & "</font>"
                End If

                Dim vsIDNO As String = ""
                Dim vsuser_ID As String = ""
                Dim vsWOassPrd As String = ""

                If IsDBNull(dr("SPage")) Then
                    '父層無子頁
                    MyNode.SelectAction = TreeNodeSelectAction.Expand
                    'onclick=""TreeView_ToggleNode(TreeView1_Data," & tnCnt & ", TreeView1n" & tnCnt & ", '', TreeView1n" & tnCnt & "Nodes);
                Else
                    Select Case Kind
                        Case "OO"
                            Dim flag_SPage As Boolean = True 'True:正常網頁 'False:特殊網頁
                            flag_SPage = True 'True:正常網頁

                            'lapm
                            If dr("SPage").ToString.ToLower.IndexOf("lapm") > -1 Then
                                flag_SPage = False '特殊網頁
                                vsIDNO = TIMS.Get_Account_IDNO(vsUserID, objconn)
                                vsuser_ID = TIMS.Get_SubSysUser(vsIDNO, "user_ID", objconn)
                                vsWOassPrd = TIMS.Get_SubSysUser(vsIDNO, "password", objconn)

                                If vsuser_ID <> "" AndAlso vsWOassPrd <> "" Then
                                    MyNode.NavigateUrl = dr("SPage").ToString & "index.aspx?User=" & vsuser_ID & "&Pwd=" & vsWOassPrd
                                    MyNode.Target = "_new" '其他系統 另外開視窗
                                Else
                                    MyNode.NavigateUrl = "javascript:NoAccount();"
                                    MyNode.Target = "_self" '自身錯誤訊息'職業訓練生活津貼管理系統,無此帳號!"
                                End If
                            End If

                            'cognos8
                            If dr("SPage").ToString.ToLower.IndexOf("cognos8") > -1 Then
                                flag_SPage = False '特殊網頁
                                vsuser_ID = ""
                                vsWOassPrd = ""
                                '限定局或中心使用者可用
                                Select Case vsLID
                                    Case 2
                                    Case Else ' 0, 1 '局:0 '中心:1
                                        Select Case vsDistID
                                            Case "000" '"職訓局"
                                                vsuser_ID = "user10"
                                                vsWOassPrd = "evtaaccuser10"
                                            Case "001" '"北區"
                                                vsuser_ID = "user01"
                                                vsWOassPrd = "evtaaccuser01"
                                            Case "002" '"泰山"
                                                vsuser_ID = "user02"
                                                vsWOassPrd = "evtaaccuser02"
                                            Case "003" '"桃園"
                                                vsuser_ID = "user03"
                                                vsWOassPrd = "evtaaccuser03"
                                            Case "004" '"中區"
                                                vsuser_ID = "user04"
                                                vsWOassPrd = "evtaaccuser04"
                                            Case "005" '"台南""臺南"
                                                vsuser_ID = "user05"
                                                vsWOassPrd = "evtaaccuser05"
                                            Case "006" '"南區"
                                                vsuser_ID = "user06"
                                                vsWOassPrd = "evtaaccuser06"
                                        End Select
                                End Select


                                If vsuser_ID <> "" AndAlso vsWOassPrd <> "" Then
                                    MyNode.NavigateUrl = ""
                                    MyNode.NavigateUrl += dr("SPage").ToString.ToLower
                                    MyNode.NavigateUrl += "/cgi-bin/cognos.cgi?b_action=xts.run&m=portal/cc.xts&CAMNamespace=Local NT&CAMUsername="
                                    MyNode.NavigateUrl += vsuser_ID
                                    MyNode.NavigateUrl += "&CAMPassword=" & vsWOassPrd
                                    MyNode.Target = "_new" '其他系統 另外開視窗
                                Else
                                    MyNode.NavigateUrl = "javascript:NoAccount();"
                                    MyNode.Target = "_self" '自身錯誤訊息'職業訓練生活津貼管理系統,無此帳號!"
                                End If
                            End If

                            If flag_SPage Then 'True:正常網頁
                                MyNode.NavigateUrl = dr("SPage").ToString
                                MyNode.Target = "_new" '其他系統 另外開視窗
                            End If

                        Case Else
                            MyNode.NavigateUrl = Convert.ToString(dr("SPage")) & "?ID=" & dr("FunID").ToString
                            MyNode.Target = "mainFrame"
                    End Select
                End If

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    AddNode(dt, Kind, dr, MyNode)
                End If

                ParentNode.ChildNodes.Add(MyNode)
                tnCnt += 1
            Next
        Else
            For Each dr In dt.Select("Kind='" & Kind & "' and [Parent]='" & drChild("FunID") & "'", "newSort")
                Const cst_FunLength_set As Integer = 9 '預計最大字串長度。
                Dim FunName As String = "" '下層功能名稱  Levels=0 
                Dim FunLength As Integer = 0 '功能字串長度。
                FunName = dr("Name").ToString
                If FunName <> "" Then FunName = Trim(FunName)
                FunLength = FunName.Length

                MyNode = New TreeNode

                If FunLength > cst_FunLength_set Then
                    Const cst_addSpace1 As String = "<BR>&nbsp;&nbsp;&nbsp;&nbsp;"
                    Dim i As Integer = 0
                    Dim mFunName As String = ""
                    While FunLength > 0
                        If mFunName = "" Then
                            mFunName = Mid(FunName, 1, cst_FunLength_set)
                        Else
                            mFunName &= cst_addSpace1 & Mid(FunName, i * cst_FunLength_set + 1, cst_FunLength_set)
                        End If
                        FunLength -= cst_FunLength_set '減少字串
                        i += 1
                    End While
                    MyNode.Text = "<font color='" & cst_text_color & "'>。" & mFunName & "</font>"
                Else
                    MyNode.Text = "<font color='" & cst_text_color & "'>。" & FunName & "</font>"
                End If

                If IsDBNull(dr("SPage")) Then
                    MyNode.SelectAction = TreeNodeSelectAction.Expand
                Else
                    MyNode.NavigateUrl = Convert.ToString(dr("SPage")) & "?ID=" & dr("FunID").ToString
                    MyNode.Target = "mainFrame"
                End If
                'If Not IsDBNull(dr("SPage")) Then
                '    MyNode.NavigateUrl = dr("SPage").ToString & "?ID=" & dr("FunID").ToString
                '    MyNode.Target = "mainFrame"
                'End If

                'MyNode.SelectAction = TreeNodeSelectAction.Expand
                ParentNode.ChildNodes.Add(MyNode)
                tnCnt += 1

                If dt.Select("Kind='" & Kind & "' and [Parent]='" & dr("FunID") & "'").Length <> 0 Then
                    AddNode(dt, Kind, dr, MyNode)
                End If
            Next
        End If
    End Sub

    'Private Sub ExpandNode()

    'End Sub
End Class