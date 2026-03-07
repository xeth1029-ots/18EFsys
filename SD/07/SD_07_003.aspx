<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_07_003.aspx.vb" Inherits="WDAIIP.SD_07_003" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>技能檢定查詢</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../../css/style.css" type="text/css" rel="stylesheet">
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
    </script>
    <script type="text/javascript" language="javascript">
        function ShowType1() {
            //alert('');
            //debugger;
            //var cst_inline1='inline';
            var cst_inline1 = '';
            var rdoType1 = document.getElementById('rdoType1');
            var rdoType2 = document.getElementById('rdoType2');
            var TR_STDate = document.getElementById('TR_STDate');
            var TR_FTDate = document.getElementById('TR_FTDate');
            var TR_ExamDate = document.getElementById('TR_ExamDate');
            var txtSTDateS = document.getElementById('txtSTDateS');
            var txtSTDateE = document.getElementById('txtSTDateE');
            var txtFTDateS = document.getElementById('txtFTDateS');
            var txtFTDateE = document.getElementById('txtFTDateE');
            var txtExamDateS = document.getElementById('txtExamDateS');
            var txtExamDateE = document.getElementById('txtExamDateE');

            if (!rdoType1) return false;
            if (!rdoType2) return false;

            TR_STDate.style.display = 'none';
            TR_FTDate.style.display = 'none';
            TR_ExamDate.style.display = 'none';
            if (rdoType1.checked) {
                txtExamDateS.value = '';
                txtExamDateE.value = '';
                TR_STDate.style.display = cst_inline1;
                TR_FTDate.style.display = cst_inline1;
            }
            if (rdoType2.checked) {
                txtSTDateS.value = '';
                txtSTDateE.value = '';
                txtFTDateS.value = '';
                txtFTDateE.value = '';
                TR_ExamDate.style.display = cst_inline1;
            }
        }

        function IsDate(MyDate) { //判斷日期格式
            if (MyDate != '') {
                if (!checkDate(MyDate))
                    return false;
            }
            return true;
        }

        function IsNumeric(sText) {	//判斷是否為數值
            var ValidChars = "0123456789.";
            var IsNumber = true;
            var Char;

            for (i = 0; i < sText.length && IsNumber == true; i++) {
                Char = sText.charAt(i);
                if (ValidChars.indexOf(Char) == -1) { IsNumber = false; }
            }
            return IsNumber;
        }

        function check() {
            var Item = '';
            var msg = '';
            var TPlanID = document.getElementById('TPlanID');

            var rdoType1 = document.getElementById('rdoType1');
            var rdoType2 = document.getElementById('rdoType2');
            var TR_STDate = document.getElementById('TR_STDate');
            var TR_FTDate = document.getElementById('TR_FTDate');
            var TR_ExamDate = document.getElementById('TR_ExamDate');
            var txtSTDateS = document.getElementById('txtSTDateS');
            var txtSTDateE = document.getElementById('txtSTDateE');
            var txtFTDateS = document.getElementById('txtFTDateS');
            var txtFTDateE = document.getElementById('txtFTDateE');
            var txtExamDateS = document.getElementById('txtExamDateS');
            var txtExamDateE = document.getElementById('txtExamDateE');

            var txtYearsOldS = document.getElementById('txtYearsOldS');
            var txtYearsOldE = document.getElementById('txtYearsOldE');

            //debugger;
            if (isEmpty(TPlanID)) { msg += '請選擇 訓練計畫\n'; if (Item == '') Item = 'TPlanID'; }
            if (rdoType1.checked) {
                if (txtSTDateS.value != '' && !IsDate(txtSTDateS.value)) {
                    msg += '開訓起始日期不是正確的格式\n'; if (Item == '') Item = 'txtSTDateS';
                }
                if (txtSTDateE.value != '' && !IsDate(txtSTDateE.value)) {
                    msg += '開訓迄止日期不是正確的格式\n'; if (Item == '') Item = 'txtSTDateE';
                }
                if (txtFTDateS.value != '' && !IsDate(txtFTDateS.value)) {
                    msg += '結訓起始日期不是正確的格式\n'; if (Item == '') Item = 'txtFTDateS';
                }
                if (txtFTDateE.value != '' && !IsDate(txtFTDateE.value)) {
                    msg += '結訓迄止日期不是正確的格式\n'; if (Item == '') Item = 'txtFTDateE';
                }
            }
            else {
                //2009/06/03 拿掉發證日期
                //if (document.form1.txtSendoutCertDateS.value=='' && document.form1.txtSendoutCertDateE.value=='') { 
                //	msg+='請選擇發證日期區間\n';if(Item=='') Item='txtSendoutCertDateS';}
                //if (document.form1.txtSendoutCertDateS.value!='' && !IsDate(document.form1.txtSendoutCertDateS.value)) {
                //	msg+='發證起始日期不是正確的格式\n';if(Item=='') Item='txtSendoutCertDateS';}
                //if (document.form1.txtSendoutCertDateE.value!='' && !IsDate(document.form1.txtSendoutCertDateE.value)) {
                //	msg+='發證迄止日期不是正確的格式\n';if(Item=='') Item='txtSendoutCertDateE';}
                if (txtExamDateS.value != '' && !IsDate(txtExamDateS.value)) {
                    msg += '檢定日起始日期不是正確的格式\n'; if (Item == '') Item = 'txtExamDateS';
                }
                if (txtExamDateE.value != '' && !IsDate(txtExamDateE.value)) {
                    msg += '檢定日迄止日期不是正確的格式\n'; if (Item == '') Item = 'txtExamDateE';
                }
            }
            if (txtYearsOldS.value == '' && txtYearsOldE.value == '') {
                msg += '請填寫年齡層區間\n'; if (Item == '') Item = 'txtYearsOldS';
            }
            if (txtYearsOldS.value != '' && !IsNumeric(txtYearsOldS.value)) {
                msg += '起始年齡層數字格式有誤\n'; if (Item == '') Item = 'txtYearsOldS';
            }
            if (txtYearsOldE.value != '' && !IsNumeric(txtYearsOldE.value)) {
                msg += '迄止年齡層數字格式有誤\n'; if (Item == '') Item = 'txtYearsOldE';
            }

            if (msg != '') {
                if (document.getElementById(Item))
                    document.getElementById(Item).focus();
                alert(msg);
                return false;
            }
        }

    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <table id="FrameTable" cellspacing="1" cellpadding="1" width="100%" border="0">
            <tr>
                <td>
                    <table class="font" id="Tab_Title" cellspacing="1" cellpadding="1" width="100%" border="0">
                        <tr>
                            <td>
                                <asp:Label ID="TitleLab1" runat="server"></asp:Label><asp:Label ID="TitleLab2" runat="server">
							    首頁&gt;&gt;學員動態管理&gt;&gt;技能檢定管理&gt;&gt;技能檢定查詢
                                </asp:Label>
                            </td>
                        </tr>
                    </table>
                    <div id="Div1" runat="server">
                        <table class="table_nw" id="Table2" cellspacing="1" cellpadding="1" width="100%">
                            <tr>
                                <td class="bluecol_need" style="width: 20%">訓練計畫 </td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="TPlanID" runat="server">
                                    </asp:DropDownList>
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">依循條件 </td>
                                <td class="whitecol">
                                    <asp:RadioButton ID="rdoType1" runat="server" GroupName="rdoType" Checked="True" />計畫年度
								<asp:RadioButton ID="rdoType2" runat="server" GroupName="rdoType" />考試年度
								<asp:DropDownList ID="ddlYears" runat="server"></asp:DropDownList>
                                </td>
                            </tr>
                            <tr id="TR_STDate" runat="server">
                                <td class="bluecol">開訓日期區間 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtSTDateS" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtSTDateS.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
								<asp:TextBox ID="txtSTDateE" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtSTDateE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </td>
                            </tr>
                            <tr id="TR_FTDate" runat="server">
                                <td class="bluecol">結訓日期區間 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtFTDateS" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtFTDateS.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
								<asp:TextBox ID="txtFTDateE" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtFTDateE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </td>
                            </tr>
                            <tr id="TR_ExamDate" runat="server">
                                <td class="bluecol">檢定日期區間 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtExamDateS" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtExamDateS.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />～
								<asp:TextBox ID="txtExamDateE" runat="server" Width="15%"></asp:TextBox>
                                    <img style="cursor: pointer" onclick="javascript:show_calendar('<%= txtExamDateE.ClientId %>','','','CY/MM/DD');" alt="" src="../../images/show-calendar.gif" align="top" width="30" height="30" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">年齡層區間 </td>
                                <td class="whitecol">
                                    <asp:TextBox ID="txtYearsOldS" runat="server" MaxLength="3" Columns="3" Width="10%"></asp:TextBox>歲～
								<asp:TextBox ID="txtYearsOldE" runat="server" MaxLength="3" Columns="3" Width="10%"></asp:TextBox>歲 </td>
                            </tr>
                            <tr>
                                <td class="bluecol_need">統計項目 </td>
                                <td class="whitecol">
                                    <asp:RadioButtonList ID="rbCounttype" runat="server" CssClass="font" RepeatLayout="Flow" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="1" Selected="True">依班別</asp:ListItem>
                                        <asp:ListItem Value="2">依檢定類別</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <table width="100%">
                            <tr>
                                <td class="whitecol" align="center">
                                    <asp:Button ID="btnSend" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div id="Div2" runat="server">
                        <table class="font" id="DataGridTable1" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>轄區：
								<asp:Label ID="Label1" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" ShowFooter="True" CellPadding="8">
                                        <FooterStyle HorizontalAlign="Center" BackColor="#B0E2FF"></FooterStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="訓練計畫" FooterText="合計" HeaderStyle-Width="18%">
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="LinkButton1" runat="server" ForeColor="Blue"></asp:LinkButton>
                                                </ItemTemplate>
                                                <FooterStyle HorizontalAlign="Left"></FooterStyle>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="StudCount" HeaderText="報檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ExamCount" HeaderText="到檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="okPassCount" HeaderText="及格人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="noPassCount" HeaderText="缺考人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="及格率(%)" HeaderStyle-Width="10%">
                                                <ItemTemplate>
                                                    <asp:Label ID="passrate" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                                    </asp:DataGrid>
                                    <%--
										<asp:datagrid id="DataGrid1" runat="server" 
											OnItemCommand="DataGrid1_ItemCommand" AutoGenerateColumns="False"
											 PageSize="2" AllowPaging="True">
										
                                    --%>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Button ID="Button1" runat="server" Text="回上層" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div id="Div3" runat="server">
                        <table class="font" id="DataGridTable2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>轄區：
								<asp:Label ID="Label2A" runat="server"></asp:Label>&nbsp;&nbsp;訓練計畫：
								<asp:Label ID="Label2B" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid2" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" ShowFooter="True" CellPadding="8">
                                        <FooterStyle HorizontalAlign="Center" BackColor="#B0E2FF"></FooterStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="申請機構" FooterText="合計" HeaderStyle-Width="18%">
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:LinkButton ID="Linkbutton2" runat="server" ForeColor="Blue"></asp:LinkButton>
                                                </ItemTemplate>
                                                <FooterStyle HorizontalAlign="Left"></FooterStyle>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="StudCount" HeaderText="報檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ExamCount" HeaderText="到檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="okPassCount" HeaderText="及格人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="noPassCount" HeaderText="缺考人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="及格率(%)" HeaderStyle-Width="10%">
                                                <ItemTemplate>
                                                    <asp:Label ID="passrate2" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                                    </asp:DataGrid>
                                    <%--
										<asp:datagrid id="DataGrid1" runat="server" 
											OnItemCommand="DataGrid1_ItemCommand" AutoGenerateColumns="False"
											 PageSize="2" AllowPaging="True">
										
                                    --%>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Button ID="Button2" runat="server" Text="回上層" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                    <div id="Div4" runat="server">
                        <table class="font" id="DataGridTable3" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>轄區：
								<asp:Label ID="Label3A" runat="server"></asp:Label>&nbsp;&nbsp;訓練計畫：
								<asp:Label ID="Label3B" runat="server"></asp:Label>&nbsp;&nbsp;申請機構：
								<asp:Label ID="Label3C" runat="server"></asp:Label>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="DataGrid3" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" ShowFooter="True" CellPadding="8">
                                        <FooterStyle HorizontalAlign="Center" BackColor="#B0E2FF"></FooterStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="班別" FooterText="合計" HeaderStyle-Width="18%">
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="labClassName" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <FooterStyle HorizontalAlign="Left"></FooterStyle>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="StudCount" HeaderText="報檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ExamCount" HeaderText="到檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="okPassCount" HeaderText="及格人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="noPassCount" HeaderText="缺考人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="及格率(%)" HeaderStyle-Width="10%">
                                                <ItemTemplate>
                                                    <asp:Label ID="passrate3" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                                    </asp:DataGrid>
                                    <%--
										<asp:datagrid id="DataGrid1" runat="server" 
											OnItemCommand="DataGrid1_ItemCommand" AutoGenerateColumns="False"
											 PageSize="2" AllowPaging="True">
										
                                    --%>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <asp:DataGrid ID="Datagrid4" runat="server" CssClass="font" Width="100%" AutoGenerateColumns="False" ShowFooter="True" CellPadding="8">
                                        <FooterStyle HorizontalAlign="Center" BackColor="#B0E2FF"></FooterStyle>
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn HeaderText="檢定類別" FooterText="合計" HeaderStyle-Width="18%">
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                <ItemTemplate>
                                                    <asp:Label ID="Examkind" runat="server"></asp:Label>
                                                </ItemTemplate>
                                                <FooterStyle HorizontalAlign="Left"></FooterStyle>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="StudCount" HeaderText="報檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="ExamCount" HeaderText="到檢人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="okPassCount" HeaderText="及格人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:BoundColumn DataField="noPassCount" HeaderText="缺考人數" HeaderStyle-Width="18%"></asp:BoundColumn>
                                            <asp:TemplateColumn HeaderText="及格率(%)" HeaderStyle-Width="10%">
                                                <ItemTemplate>
                                                    <asp:Label ID="passrate4" runat="server"></asp:Label>
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                        </Columns>
                                        <PagerStyle HorizontalAlign="Center" Position="Top"></PagerStyle>
                                    </asp:DataGrid>
                                    <%--
										<asp:datagrid id="DataGrid1" runat="server" 
											OnItemCommand="DataGrid1_ItemCommand" AutoGenerateColumns="False"
											 PageSize="2" AllowPaging="True">
										
                                    --%>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" colspan="2" class="whitecol">
                                    <asp:Button ID="Button3" runat="server" Text="回上層" CssClass="asp_button_M"></asp:Button>
                                </td>
                            </tr>
                        </table>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
