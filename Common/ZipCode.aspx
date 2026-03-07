<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ZipCode.aspx.vb" Inherits="WDAIIP.ZipCode1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
    <title>郵遞區號</title>
    <meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
    <meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
    <meta content="JavaScript" name="vs_defaultClientScript">
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
    <link href="../css/style.css" type="text/css" rel="stylesheet">
    <style type="text/css">
        A:link { color: #000000; text-decoration: none; }
        A:visited { color: #000000; text-decoration: none; }
        A:active { color: #666666; text-decoration: none; }
        A:hover { color: #333333; text-decoration: underline; }
    </style>
    <script type="text/javascript" src="../js/common.js"></script>
    <script language="javascript" type="text/javascript">
        function getVal(ctid, zipcode, zipcode2, zipcode6, ctname, zipname, zipcode_n, road) {
            var CtID = document.getElementById('hidCtID').value;
            var ZipCode = document.getElementById('hidZipCode').value;
            var ZipCode2 = document.getElementById('hidZipCode2').value;
            //flag_work2022xZIP6
            var ZipCode6 = document.getElementById('hidZipCode6').value;
            var CtName = document.getElementById('hidCtName').value;
            var CityName = document.getElementById('hidCityName').value;
            var ZipName = document.getElementById('hidZipName').value;
            var ZIPCODE_N = document.getElementById('hidZIPCODE_N').value;
            var Road = document.getElementById('hidRoad').value;
            //var hidno = document.getElementById('hidno').value;
            var oZipCode6 = window.opener.document.getElementById(ZipCode6);
            var oCtName = window.opener.document.getElementById(CtName);

            if (CtID != '') window.opener.document.getElementById(CtID).value = ctid;
            if (ZipCode != '') window.opener.document.getElementById(ZipCode).value = zipcode;
            if (ZipCode2 != '') window.opener.document.getElementById(ZipCode2).value = zipcode2;
            //ZipCode6,flag_work2022xZIP6
            if (ZipCode6 != '' && oZipCode6) oZipCode6.value = zipcode6;
            if (CtName != '' && oCtName) {
                oCtName.value = '';
                if (zipcode != '') oCtName.value = '(' + zipcode + ')' + ctname + zipname;
            }
            if (CityName != '') window.opener.document.getElementById(CityName).value = ctname;
            if (ZipName != '') window.opener.document.getElementById(ZipName).value = zipname;
            if (ZIPCODE_N != '') window.opener.document.getElementById(ZIPCODE_N).value = zipcode_n;
            if (Road != '') window.opener.document.getElementById(Road).value = road;
            //alert(ctname);alert(zipname);alert(CityName);alert(ZipName);
            window.close();
        }
    </script>
</head>
<body bgcolor="white">
    <form id="form1" method="post" runat="server">
        <input id="hidSN" type="hidden" runat="server" />
        <input id="hidCtID" type="hidden" runat="server" />
        <input id="hidZipCode" type="hidden" runat="server" />
        <input id="hidZipCode2" type="hidden" runat="server" />
        <input id="hidZipCode6" type="hidden" runat="server" />
        <input id="hidCtName" type="hidden" runat="server" />
        <input id="hidCityName" type="hidden" runat="server" />
        <input id="hidZipName" type="hidden" runat="server" />
        <input id="hidZIPCODE_N" type="hidden" runat="server" />
        <input id="hidRoad" type="hidden" runat="server" />
        <div style="font-weight: bold" align="center">選擇郵遞區號</div>
        <br />
        <table cellspacing="0" cellpadding="0" width="100%" align="left" border="0">
            <tr>
                <td>
                    <table class="table_nw" cellspacing="1" cellpadding="1" width="100%" align="left" border="0">
                        <tr>
                            <td width="20%" class="bluecol">郵遞區號種類 </td>
                            <td width="80%" class="whitecol">
                                <asp:RadioButtonList ID="rblPOSTTYPE1" runat="server" RepeatDirection="Horizontal" RepeatLayout="Flow">
                                    <asp:ListItem Value="3" Selected="True">3+3郵遞區號</asp:ListItem>
                                    <asp:ListItem Value="2">3+2郵遞區號</asp:ListItem>
                                </asp:RadioButtonList>
                        </tr>
                        <tr>
                            <td class="bluecol">郵遞區號 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtZip" runat="server" MaxLength="7" Width="18%"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td class="bluecol">縣市/鄉鎮區 </td>
                            <td class="whitecol">
                                <asp:DropDownList ID="ddlCity" runat="server" AutoPostBack="true"></asp:DropDownList>／
                                <asp:DropDownList ID="ddlZip" runat="server" Enabled="false" AutoPostBack="True">
                                    <asp:ListItem Value="請選擇">請選擇</asp:ListItem>
                                </asp:DropDownList>
                            </td>
                        </tr>
                        <tr>
                            <td class="bluecol">街道名稱 </td>
                            <td class="whitecol">
                                <asp:TextBox ID="txtRoad" runat="server" CssClass="ipt" MaxLength="30" Width="50%"></asp:TextBox></td>
                        </tr>
                    </table>
                </td>
            </tr>
            <tr>
                <td align="center" class="whitecol">
                    <asp:Button ID="btnSch" runat="server" Text="查詢" CssClass="asp_button_M"></asp:Button>
                    &nbsp;<asp:Button ID="btnCancel" runat="server" Text="取消" CssClass="asp_button_M"></asp:Button>
                    &nbsp;<asp:Button ID="btnClear1" runat="server" Text="清空" Visible="False" CssClass="asp_button_M"></asp:Button>
                </td>
            </tr>
            <tr>
                <td align="left">
                    <br>
                    <div align="center">
                        <asp:Label ID="labMsg" Style="color: red;" runat="server"></asp:Label>
                    </div>
                    <div style="padding-bottom: 0px; margin-top: 0px; padding-left: 0px; width: 100%; padding-right: 0px; margin-bottom: 0px; margin-left: 0px; overflow: auto; padding-top: 0px">
                        <div style="overflow-y: auto; height: 400px;">
                            <asp:DataGrid ID="DataGrid1" runat="server" Width="100%" AutoGenerateColumns="False" AllowPaging="True" BorderColor="AliceBlue" CellPadding="8">
                                <AlternatingItemStyle BackColor="#F5F5F5"></AlternatingItemStyle>
                                <HeaderStyle HorizontalAlign="Center" CssClass="head_navy"></HeaderStyle>
                                <Columns>
                                    <asp:TemplateColumn ItemStyle-Width="4%" ItemStyle-HorizontalAlign="Center">
                                        <ItemTemplate>
                                            <asp:RadioButton ID="rdoSelect" runat="server" GroupName="grdoSelect"></asp:RadioButton>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="郵遞區號" ItemStyle-Width="14%" ItemStyle-HorizontalAlign="Center">
                                        <ItemTemplate>
                                            <asp:Label ID="labZipCode" runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="地區" ItemStyle-Width="38%" ItemStyle-HorizontalAlign="Center">
                                        <ItemTemplate>
                                            <asp:Label ID="labZipArea" runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="街道名稱" ItemStyle-Width="23%" ItemStyle-HorizontalAlign="Center">
                                        <ItemTemplate>
                                            <asp:Label ID="labRoad" runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                    <asp:TemplateColumn HeaderText="說明" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="21%">
                                        <ItemTemplate>
                                            <asp:Label ID="labNote" runat="server"></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateColumn>
                                </Columns>
                                <PagerStyle HorizontalAlign="center" Mode="NumericPages"></PagerStyle>
                            </asp:DataGrid>
                        </div>
                    </div>
                </td>
            </tr>
        </table>
    </form>
</body>
</html>