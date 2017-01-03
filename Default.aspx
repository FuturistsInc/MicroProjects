<%@ Page Title="Welcome to MicroProjects" Language="C#" MasterPageFile="~/Master/MasterPage.master" AutoEventWireup="true" CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContentPlaceHolder" Runat="Server">
    <h3>Welcome to MicroProjects</h3>    
    <p id="p_Status" runat="server"></p>
    <asp:Button ID="btn_Test" runat="server" Text="Click to Test" OnClick="btn_Test_Click" />
</asp:Content>

