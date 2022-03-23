<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log14_ok.aspx.vb" Inherits="v_log14_ok" %>


<html xmlns="http://www.w3.org/1999/xhtml" >
<head id="Head1" runat="server">
    
</head>

<script type="text/javascript">

function choiceUser(fieldClientID,isMutilSelect){
window.open('UserSelector.aspx?fieldClientID='+fieldClientID.id+'&mutilSelect='+isMutilSelect,'','width=220,height=220','');
}
function choiceDate(fieldClientID){
window.open('dtPicker.aspx?fieldClientID='+fieldClientID.id,'','width=220,height=220','');
}

</script>

<body bgcolor="#99ccff" style="text-align: center; background-color: #ccffff;" topmargin="3">
 
    <form id="Form1" runat="server">
      
          <table style="width: 760px">
    <tr><td>
    <asp:Image ID="Image1" runat="server" Height="86px" Width="760px" ImageUrl="~/image/FAR_TOP2.JPG" />
    </td></tr>
     <tr><td colspan="3" style ="text-align: right ">  
            <asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="登出" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button>
        </td></tr>
    </table>
          
          <table cellpadding="0" cellspacing="0" style="width: 100%; border-top-style: none;
          border-right-style: none; border-left-style: none; border-bottom-style: none">
          <tr>
            <td align="center" colspan="3" style="background-color: #ccffff; ">
                <asp:SqlDataSource ID="SqlDataSource1" runat="server"
                    ProviderName="System.Data.OleDb" SelectCommand="SELECT * FROM [LogTrace] order by ActionTime ASC">
                </asp:SqlDataSource>
                &nbsp;<br />
                       <asp:Label ID="lblMDB" runat="server" Font-Size="Small" ForeColor="Blue" Text="lblMDB" BackColor="#C0FFFF"></asp:Label><br />
                <asp:Label ID="title" runat="server" Font-Bold="True" Text="【14.使用者&權限異動明細查詢】"></asp:Label>
              <P>
&nbsp;<asp:label id="result1" runat="server" ForeColor="#404040" BackColor="#C0FFFF" Font-Size="Larger">FROM XX TO XX</asp:label>
              </P>
                <p>
                  <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                      DataSourceID="SqlDataSource1" BackColor="White" BorderColor="Black" CellPadding="3" EnableSortingAndPagingCallbacks="True" EnableTheming="True" ForeColor="Black" EmptyDataText="查不到符合的資料" Width="90%" AllowSorting="True" DataKeyNames="AccName" PageSize="50">
                      <Columns>
                          <asp:BoundField DataField="bra01" HeaderText="分行代碼" SortExpression="bra01" />
                           <asp:BoundField DataField="AccName" HeaderText="使用者帳號" 
                               SortExpression="AccName" />
                          <asp:BoundField DataField="ActionTime" HeaderText="操作時間" 
                               SortExpression="ActionTime" />
                          <asp:BoundField DataField="ActionCode" HeaderText="操作" 
                               SortExpression="ActionCode" />
                          
                      </Columns>
                      <FooterStyle BackColor="#B5C7DE" ForeColor="#4A3C8C" />
                      <RowStyle BackColor="#C0C0FF" ForeColor="#4A3C8C" />
                      <SelectedRowStyle BackColor="#738A9C" Font-Bold="True" ForeColor="#F7F7F7" />
                      <PagerStyle BackColor="#E7E7FF" ForeColor="#4A3C8C" HorizontalAlign="Left" />
                      <HeaderStyle BackColor="#4A3C8C" Font-Bold="True" ForeColor="#F7F7F7" />
                      <AlternatingRowStyle BackColor="#F7F7F7" />
                      <EmptyDataRowStyle BackColor="#C0FFFF" BorderColor="#C0FFFF" Font-Bold="False" Font-Size="Medium" />
                      <PagerSettings Mode="NumericFirstLast" Position="TopAndBottom" />
                  </asp:GridView>
                    &nbsp;</p>
                <p>
                  <asp:Label ID="Lbl2" runat="server" ForeColor="Red" ></asp:Label>
                </p>
                <p>
                    <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" Width="95px" Height="25px" />&nbsp;&nbsp;<asp:button id="butMenu" runat="server" ForeColor="#0000C0" BackColor="#E0E0E0" BorderColor="#00AEEF" Text="回主選單" Width="95px" BorderStyle="Solid" Height="25px"></asp:button></p>
                
            </td>
          </tr>
          <tr style="font-size: 12pt">
            <td style="width: 100px">
            </td>
            <td style="width: 543px">
            </td>
            <td style="width: 43px">
            </td>
          </tr>
        </table>

 
 


      <center style="text-align: right">
        &nbsp;</center>
      <center style="text-align: right">
        &nbsp;</center>
      <center style="text-align: right">
        &nbsp;</center>
  </form>

</body>
</html>
