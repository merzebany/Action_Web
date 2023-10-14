<%@ Page Title="Home Page" Language="VB" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.vb" Inherits="Action_Web._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">


    <div style ="background-color : forestgreen; color:white ;  font-size:xx-large;  height: 55px" align="center">

        Project Acivity Orgnization

        </div>
    <br />

     <table align="center" class="nav-justified">
         <tr>
             <td style="height: 35px; width: 283px;" >
                 <asp:Label ID="Label1" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium" Font-Strikeout="False" Text="Activity ID"></asp:Label>
             </td>
             <td style="height: 35px">
                 <asp:TextBox ID="TB_ActID" runat="server" Width="372px" BorderStyle="Solid" style="padding-top:3px; padding-bottom:3px;" ></asp:TextBox>
             </td>
         </tr>
         <tr>
             <td style="width: 283px; height: 35px;">
                 <asp:Label ID="Label2" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium" Font-Strikeout="False" Text="Activity"></asp:Label>
             </td>
             <td style="height: 35px">
                 <asp:TextBox ID="TB_Act" runat="server" Width="372px" BorderStyle="Solid" style="padding-top:3px; padding-bottom:3px;"></asp:TextBox>
             </td>
         </tr>
         <tr>
             <td style="height: 37px; width: 283px;">
                 <asp:Label ID="Label3" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium" Font-Strikeout="False" Text="Responspilty"></asp:Label>
             </td>
             <td style="height: 37px">
                 <asp:TextBox ID="TB_Resp" runat="server" Width="372px" BorderStyle="Solid" style="padding-top:3px; padding-bottom:3px;"></asp:TextBox>
             </td>
         </tr>
         <tr>
             <td style="width: 283px">
                 <asp:Label ID="Label4" runat="server" Font-Bold="True" Font-Names="Arial" Font-Size="Medium" Font-Strikeout="False" Text="End Date"></asp:Label>
             </td>
             <td style="height: 35px">
                 <asp:TextBox ID="TB_EndDate" runat="server" Width="372px" BorderStyle="Solid" TextMode="DateTime" style="padding-top:3px; padding-bottom:3px;"></asp:TextBox>
             </td>
         </tr>

         <tr>
             <td style="width: 283px">
              
             </td>
             <td style="height: 35px">
              
                 <asp:Button ID="BT_UpData" runat="server" BackColor="#009933" BorderColor="#009933" BorderStyle="Solid" Font-Bold="True" Font-Size="Small" ForeColor="White" Text="UpData" Width="150px" style="margin-right:20px;" />
                 <asp:Button ID="BT_Delete" runat="server" BackColor="#009933" BorderColor="#009933" BorderStyle="Solid" Font-Bold="True" Font-Size="Small" ForeColor="White" Text="Delete" Width="150px" style="margin-right:20px;" />
              
             </td>
         </tr>

    </table>

     <div align="center"  style="padding :15px ; margin-top:20px;">

         <hr />
         <asp:GridView ID="GridView1" runat="server" style="margin-left: 0px" Width="1006px" AllowSorting="True" Font-Bold="True" Height="184px" HorizontalAlign="Center">
             <HeaderStyle BackColor="#339966" Font-Size="Medium" ForeColor="White" HorizontalAlign="Center" VerticalAlign="Middle" />
             <RowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
             <SelectedRowStyle BackColor="#00FFCC" />
         </asp:GridView>


     </div>


</asp:Content>
