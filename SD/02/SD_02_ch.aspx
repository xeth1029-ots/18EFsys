<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="SD_02_ch.aspx.vb" Inherits="WDAIIP.SD_02_ch" %>

<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>職類班級選擇</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR" />
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="../../css/style.css" type="text/css" rel="stylesheet" />
    <script type="text/javascript" language="javascript" src="../../js/date-picker.js"></script>
    <script type="text/javascript" language="javascript" src="../../js/openwin/openwin.js"></script>
    <script type="text/javascript" src="../../js/common.js"></script>
    <script type="text/javascript" src="../../Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" language="javascript">
        var _UA = window.navigator.userAgent;
        var _isIE = (_UA.indexOf("MSIE") != -1 || _UA.indexOf("Trident") != -1) ? true : false;
        if (document.body) { window.scroll(0, document.body.scrollHeight); }
        //if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }
        //if (!_isIE) { if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); } }
    </script>
    <script type="text/javascript" language="javascript">
        //查詢時的檢查。
        function chkdata() {
            var msg = ''
            if (document.getElementById('CyclType').value != '' && !isUnsignedInt(document.getElementById('CyclType').value)) msg += '期別請輸入數字\n'
            if (msg != '') {
                //opener.alert(msg);
                $('#myMsg').text(msg);
                return false;
            }
            else {
                $('#myMsg').text(msg);
                return true;
            }
        }

        //清除職業類別
        function ClearTMID() {
            document.getElementById('TB_career_id').value = '';
            document.getElementById('trainValue').value = '';
        }

        //清除通俗職類
        function ClearCjob() {
            document.getElementById('txtCJOB_NAME').value = '';
            document.getElementById('cjobValue').value = '';
        }
    </script>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div style="overflow-y: auto; height: 630px;">
            <%--style="height: 650px; overflow-y: auto; overflow: scroll;"--%>
            <table id="Table3" cellspacing="1" cellpadding="1" width="100%" border="0">
                <tr>
                    <td>
                        <table class="table_nw" id="Table1" cellspacing="1" cellpadding="1" width="100%">
                            <tr id="YearsTR" runat="server">
                                <td class="bluecol" style="width: 20%">年度</td>
                                <td class="whitecol" colspan="3">
                                    <asp:DropDownList ID="Years" runat="server"></asp:DropDownList></td>
                            </tr>
                            <tr>
                                <td class="bluecol" style="width: 20%">班別代碼</td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="ClassID" runat="server" Columns="15" Width="50%" MaxLength="30"></asp:TextBox></td>
                                <td class="bluecol" style="width: 20%">訓練職類</td>
                                <td class="whitecol" style="width: 30%">
                                    <asp:TextBox ID="TB_career_id" runat="server" Columns="15" onfocus="this.blur()" Width="60%"></asp:TextBox>
                                    <input onclick="openTrain(document.getElementById('trainValue').value);" type="button" value="..." class="button_b_Mini" />
                                    <input id="Button1" onclick="ClearTMID();" type="button" value="清除" name="Button1" class="button_b_S" />
                                    <input id="trainValue" type="hidden" name="trainValue" runat="server" />
                                    <input id="jobValue" type="hidden" name="jobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">通俗職類</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="txtCJOB_NAME" runat="server" Columns="30" onfocus="this.blur()" Width="40%"></asp:TextBox>
                                    <input id="btu_sel2" onclick="openCjob(document.getElementById('cjobValue').value);" type="button" value="..." name="btu_sel2" runat="server" class="button_b_Mini" />
                                    <input id="Button2" onclick="ClearCjob();" type="button" value="清除" name="Button2" class="button_b_S" />
                                    <input id="cjobValue" type="hidden" name="cjobValue" runat="server" />
                                </td>
                            </tr>
                            <tr>
                                <td class="bluecol">訓練時段</td>
                                <td class="whitecol">
                                    <asp:DropDownList ID="HourRan" runat="server"></asp:DropDownList></td>
                                <td class="bluecol">期別</td>
                                <td class="whitecol">
                                    <asp:TextBox ID="CyclType" runat="server" Columns="5" Width="30%" MaxLength="5"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級名稱</td>
                                <td class="whitecol" colspan="3">
                                    <asp:TextBox ID="ClassCName" runat="server" Columns="50" Width="60%" MaxLength="100"></asp:TextBox></td>
                            </tr>
                            <tr>
                                <td class="bluecol">班級範圍</td>
                                <td class="whitecol" colspan="3">
                                    <asp:RadioButtonList ID="ClassRound" runat="server" CssClass="font" RepeatDirection="Horizontal">
                                        <asp:ListItem Value="開訓二週前" Selected="True">開訓二週前</asp:ListItem>
                                        <asp:ListItem Value="已開訓">已開訓</asp:ListItem>
                                        <asp:ListItem Value="已結訓">已結訓</asp:ListItem>
                                        <asp:ListItem Value="未開訓">未開訓</asp:ListItem>
                                        <asp:ListItem Value="全部">全部</asp:ListItem>
                                    </asp:RadioButtonList>
                                </td>
                            </tr>
                        </table>
                        <div align="center" class="whitecol">
                            <asp:Button ID="search_but" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button></div>
                        <div align="center">
                            <asp:Label ID="myMsg" runat="server" ForeColor="Red"></asp:Label><br />
                            <asp:Label ID="msg" runat="server" ForeColor="Red"></asp:Label>
                        </div>
                        <table id="Table2" cellspacing="1" cellpadding="1" width="100%" border="0" runat="server">
                            <tr>
                                <td>
                                    <%--<div style="overflow-y: auto; height: 310px;"></div>--%>
                                    <%----%>
                                    <asp:DataGrid ID="DataGrid1" runat="server" CssClass="font" AllowPaging="True" AutoGenerateColumns="False" Width="100%" CellPadding="6">
                                        <AlternatingItemStyle BackColor="#F5F5F5" />
                                        <HeaderStyle CssClass="head_navy" />
                                        <Columns>
                                            <asp:TemplateColumn>
                                                <HeaderStyle Width="8%" />
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                                <ItemTemplate>
                                                    <input id="radio1" value='<%# DataBinder.Eval(Container.DataItem, "OCID")%>' type="radio" name="class1">
                                                </ItemTemplate>
                                            </asp:TemplateColumn>
                                            <asp:BoundColumn DataField="TrainName2" HeaderText="訓練職類">
                                                <HeaderStyle HorizontalAlign="Center" Width="28%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="ClassID" HeaderText="班級代碼">
                                                <HeaderStyle HorizontalAlign="Center" Width="20%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="CLASSCNAME2" HeaderText="班級名稱">
                                                <HeaderStyle HorizontalAlign="Center" Width="28%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="STDate" HeaderText="開訓日期">
                                                <HeaderStyle HorizontalAlign="Center" Width="16%"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                            <asp:BoundColumn DataField="IsApplic2" HeaderText="志願班別" Visible="false">
                                                <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                                <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                            </asp:BoundColumn>
                                        </Columns>
                                        <PagerStyle Visible="False"></PagerStyle>
                                    </asp:DataGrid>

                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <div align="center">
                                        <uc1:PageControler ID="PageControler1" runat="server"></uc1:PageControler>
                                    </div>
                                </td>
                            </tr>
                            <tr>
                                <td align="center" class="whitecol">
                                    <asp:Button ID="send" runat="server" Text="送出" CssClass="asp_button_M"></asp:Button></td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </div>
        <asp:HiddenField ID="Hid_SSSDTRID" runat="server" />
    </form>
</body>
</html>
