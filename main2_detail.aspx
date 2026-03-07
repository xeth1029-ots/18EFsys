<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="main2_detail.aspx.vb" Inherits="WDAIIP.main2_detail" %>

<!DOCTYPE html>
<html>
<head runat="server">
    <title>首頁</title>
    <meta charset="utf-8" />
    <meta content="JavaScript" name="vs_defaultClientScript" />
    <meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema" />
    <link href="css/style.css" type="text/css" rel="stylesheet" />
    <link href="css/homebase.css" type="text/css" rel="stylesheet" />
    <link href="css/jquery-confirm.min.css" rel="stylesheet" />
    <link href="css/bootstrap3-3-6.min.css" rel="stylesheet" />
    <link href="css/bootstrap-treeview.css" rel="stylesheet" />
    <link href="css/font-awesome-4.7.0.min.css" rel="stylesheet" />
    <link href="css/font-awesome.css" rel="stylesheet" />
    <link href="css/font-awesome.min.css" rel="stylesheet" />
    <link rel="stylesheet" href="https://use.fontawesome.com/releases/v5.2.0/css/all.css" integrity="sha384-hWVjflwFxL6sNzntih27bfxkr27PmbbK/iSvJ+a4+0owXq79v+lsFkW54bOGbiDQ" crossorigin="anonymous" />

    <script type="text/javascript" src="js/date-picker.js"></script>
    <script type="text/javascript" src="js/openwin/openwin.js"></script>
    <script type="text/javascript" src="Scripts/respond.min.js"></script>

    <script type="text/javascript" src="Scripts/jquery-3.7.1.min.js"></script>
    <script type="text/javascript" src="Scripts/jquery-migrate-3.4.1.min.js"></script>
    <script type="text/javascript" src="Scripts/bootstrap.min.js"></script>
    <%--<script type="text/javascript" src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.0/jquery.min.js"></script>
    <script type="text/javascript" src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/js/bootstrap.min.js"></script>--%>
</head>
<body>
    <form id="form1" method="post" runat="server">
        <div class="wrap">
            <div class="index-info">
                <div class="col-xs-12" runat="server" id="divArea1" visible="false">
                    <div class="theme-news">
                        <h3 class="theme-news-titleC"><i class="far fa-bell"></i>作業提醒</h3>
                        <div class="table-responsive">
                            <asp:GridView ID="gv1" runat="server" Width="100%" ShowHeader="False" AutoGenerateColumns="False" BorderWidth="0px" CellPadding="0" CssClass="table-news" EmptyDataText="本日無系統作業提醒。">
                                <Columns>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Top" ItemStyle-HorizontalAlign="left" ItemStyle-BorderStyle="None">
                                        <ItemTemplate>
                                            <asp:Label ID="labSubject" runat="server" Text='<%# Bind("Subject") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:Button ID="btnMore1" runat="server" Text="back" CssClass="btn btn-news-more" />
                    </div>
                </div>
                <div class="col-xs-12" runat="server" id="divArea2" visible="false">
                    <div class="theme-news">
                        <h3 class="theme-news-titleD"><i class="fas fa-newspaper"></i>最新消息</h3>
                        <div class="table-responsive">
                            <asp:GridView ID="gv2" runat="server" AllowPaging="false" Width="100%" ShowHeader="false" AutoGenerateColumns="false" BorderWidth="0" CellSpacing="0" CellPadding="0" CssClass="table-news" EmptyDataText="無最新消息。">
                                <Columns>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="center" ItemStyle-Width="15%" ItemStyle-BorderStyle="None" ItemStyle-ForeColor="#127abc">
                                        <ItemTemplate>
                                            <asp:Label ID="labPostDate" runat="server" Text='<%# Bind("PostDate") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Top" ItemStyle-HorizontalAlign="left" ItemStyle-Width="85%" ItemStyle-BorderStyle="None">
                                        <ItemTemplate>
                                            <asp:Label ID="labSubject" runat="server" Text='<%# Bind("Subject") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                            <asp:Button ID="btnMore2" runat="server" Text="back" CssClass="btn btn-news-more" />
                        </div>
                    </div>
                </div>
                <div class="col-xs-12" runat="server" id="divArea3" visible="false">
                    <div class="theme-news">
                        <h3 class="theme-news-titleE"><i class="fas fa-wrench"></i>功能增修說明</h3>
                        <div class="table-responsive">
                            <asp:GridView ID="gv3" runat="server" AllowPaging="false" Width="100%" ShowHeader="false" AutoGenerateColumns="false" BorderWidth="0" CellSpacing="0" CellPadding="0" CssClass="table-news" EmptyDataText="功能增修說明。">
                                <Columns>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="center" ItemStyle-Width="15%" ItemStyle-BorderStyle="None" ItemStyle-ForeColor="#127abc">
                                        <ItemTemplate>
                                            <asp:Label ID="labPostDate" runat="server" Text='<%# Bind("PostDate") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Top" ItemStyle-HorizontalAlign="left" ItemStyle-Width="85%" ItemStyle-BorderStyle="None">
                                        <ItemTemplate>
                                            <asp:Label ID="labSubject" runat="server" Text='<%# Bind("Subject") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                            <asp:Button ID="btnMore3" runat="server" Text="back" CssClass="btn btn-news-more" />
                        </div>
                    </div>
                </div>
                <div class="col-xs-12" runat="server" id="divArea4" visible="false">
                    <div class="theme-download">
                        <h3 class="theme-download-title">下載專區</h3>
                        <div class="table-responsive">
                            <asp:GridView ID="gv4" runat="server" AllowPaging="false" Width="100%" ShowHeader="false" AutoGenerateColumns="false" BorderWidth="0" CellSpacing="0" CellPadding="0" CssClass="table-news" EmptyDataText="功能增修說明。">
                                <Columns>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="center" ItemStyle-Width="15%" ItemStyle-BorderStyle="None" ItemStyle-ForeColor="#127abc">
                                        <ItemTemplate>
                                            <asp:Label ID="labPostDate" runat="server" Text='<%# Bind("PostDate") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Top" ItemStyle-HorizontalAlign="left" ItemStyle-Width="85%" ItemStyle-BorderStyle="None">
                                        <ItemTemplate>
                                            <asp:Label ID="labSubject" runat="server" Text='<%# Bind("Subject") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:Button ID="btnMore4" runat="server" Text="back" CssClass="btn btn-news-more"/>
                    </div>
                </div>
                <div class="col-xs-12" runat="server" id="divArea5" visible="false">
                    <div class="theme-svsteaching">
                        <h3 class="theme-svsteaching-title">影音教學專區</h3>
                        <div class="table-responsive">
                            <asp:GridView ID="gv5" runat="server" AllowPaging="false" Width="100%" ShowHeader="false" AutoGenerateColumns="false" BorderWidth="0" CellSpacing="0" CellPadding="0" CssClass="table-news" EmptyDataText="功能增修說明。">
                                <Columns>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Middle" ItemStyle-HorizontalAlign="center" ItemStyle-Width="15%" ItemStyle-BorderStyle="None" ItemStyle-ForeColor="#127abc">
                                        <ItemTemplate>
                                            <asp:Label ID="labPostDate" runat="server" Text='<%# Bind("PostDate") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:TemplateField ItemStyle-VerticalAlign="Top" ItemStyle-HorizontalAlign="left" ItemStyle-Width="85%" ItemStyle-BorderStyle="None">
                                        <ItemTemplate>
                                            <asp:Label ID="labSubject" runat="server" Text='<%# Bind("Subject") %>'></asp:Label>
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:Button ID="btnMore5" runat="server" Text="back" CssClass="btn btn-news-more"/>
                    </div>
                </div>
                <div class="col-xs-12" runat="server">
                    <br/>
                </div>
            </div>
        </div>
        <script type="text/javascript" src="../../Scripts/jquery-3.7.1.min.js"></script>
    </form>
</body>
</html>
