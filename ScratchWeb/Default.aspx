<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="ScratchWeb._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <div class="jumbotron">
        <h1>ASP.NET</h1>
        <p class="lead">ASP.NET is a free web framework for building great Web sites and Web applications using HTML, CSS and JavaScript.</p>
        <p><a href="http://www.asp.net" class="btn btn-primary btn-lg">Learn more &raquo;</a></p>
    </div>

    <div class="row">
        <div class="col-md-4">
            <h2>Getting Started</h2>
            <p>
                ASP.NET Web Forms lets you build dynamic websites using a familiar drag-and-drop, event-driven model.
            A design surface and hundreds of controls and components let you rapidly build sophisticated, powerful UI-driven sites with data access.
            </p>
            <p>
                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301948">Learn more &raquo;</a>
            </p>
        </div>
        <div class="col-md-4">
            <h2>Get more libraries</h2>
            <p>
                NuGet is a free Visual Studio extension that makes it easy to add, remove, and update libraries and tools in Visual Studio projects.
            </p>
            <p>
                <a class="btn btn-default" href="https://go.microsoft.com/fwlink/?LinkId=301949">Learn more &raquo;</a>
            </p>
        </div>
        <!--<div class="row">-->
        <div class="col-md-4">
            <i class="material-icons prefix"><i class="fa fa-user" aria-hidden="true"></i></i>
            <asp:DropDownList ID="GenderListBox" runat="server" ValidationGroup="default"></asp:DropDownList>
            <label>Gender List</label>
        </div>
        <!--</div>-->

        <!--<div class="row">-->
        <div class="col-md-4">
            <i class="material-icons prefix"><i class="fa fa-user" aria-hidden="true"></i></i>
            <asp:DropDownList ID="GenderDropDown" runat="server" ValidationGroup="default"></asp:DropDownList>
            <label>Gender Drop</label>
        </div>
        <!--</div>-->

        <asp:Button ID="SubmitButton" runat="server" Text="Submit" CssClass="btn btn-large" />
        <br />

        <asp:Button Text="text"  runat="server" OnClick="Unnamed_Click" />
    </div>

</asp:Content>
