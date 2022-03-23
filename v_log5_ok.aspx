<%@ Page Language="VB" AutoEventWireup="false" CodeFile="v_log5_ok.aspx.vb" Inherits="v_log5_ok" %>


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
        <table style="width: 1600px">
    <tr><td style="width: 835px; height: 92px;">
    <asp:Image ID="Image2" runat="server" Height="86px" Width="1592px" ImageUrl="~/image/FAR_TOP5.JPG" />&nbsp;
    </td></tr>
     <tr><td colspan="3" style="text-align: center ; width: 844px;">  &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;&nbsp; &nbsp;
         &nbsp;&nbsp;&nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;
         &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp;<asp:Label ID="Label1" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button1" runat="server" Width="50px" height="25px" Text="登出" ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px"></asp:button><br />
         &nbsp;<br />
         &nbsp;<asp:Label ID="lblMDB" runat="server" BackColor="#C0FFFF" Font-Size="Medium"
             ForeColor="Blue" Text="lblMDB"></asp:Label><br />
         &nbsp;<asp:Label ID="title" runat="server" Font-Bold="True" Text="【5.各項服務明細查詢】."></asp:Label><P>
          &nbsp;<asp:label id="result1" runat="server" ForeColor="#404040" BackColor="#C0FFFF" Font-Size="Larger">FROM XX TO XX</asp:label>
                <asp:SqlDataSource ID="SqlDataSource1" runat="server"
                    ProviderName="System.Data.OleDb" SelectCommand="SELECT * FROM [tabLog]">
                </asp:SqlDataSource>
                </P>
                <p>
                  <asp:GridView ID="GridView1" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                      DataSourceID="SqlDataSource1" BackColor="White" BorderColor="Black" CellPadding="3" EnableTheming="True" ForeColor="Black" EmptyDataText="查不到符合的資料" Width="90%" AllowSorting="True" PageSize="50">
                      <Columns>
                           <asp:BoundField DataField="line_no" HeaderText="線路別" SortExpression="line_no" />
                          <asp:BoundField DataField="stime_in" HeaderText="進線時間" SortExpression="stime_in" />
                          <asp:BoundField DataField="etime_out" HeaderText="掛線時間" SortExpression="etime_out" />
                          <asp:BoundField DataField="lang" HeaderText="語言" SortExpression="lang" Visible="False" />
                          <asp:BoundField DataField="trid" HeaderText="交易代碼" SortExpression="trid" />
                          <asp:BoundField DataField="error" HeaderText="訊息代碼" SortExpression="error" />
                          <asp:BoundField DataField="bra01" HeaderText="分行代碼" SortExpression="bra01" />
                          <asp:BoundField DataField="bra05" HeaderText="區別名稱" SortExpression="bra05" />
                          <asp:BoundField DataField="ao_code" HeaderText="行員編號" SortExpression="ao_code" />
                          <asp:BoundField DataField="id_no" HeaderText="客戶ID後4碼" SortExpression="id_no" />
                          <asp:BoundField DataField="email" HeaderText="email" SortExpression="email" />
                          <asp:BoundField DataField="email_cc" HeaderText="email_cc" SortExpression="email_cc" />
                          <asp:BoundField DataField="sno" HeaderText="流水編號" SortExpression="sno" />
                          <asp:BoundField DataField="file_name" HeaderText="錄音檔名" ReadOnly="True" SortExpression="file_name" />
                          <asp:BoundField DataField="file_size" HeaderText="錄音大小" ReadOnly="True" SortExpression="file_size" />
                          <asp:ButtonField ButtonType="Button" CommandName="DownloadRec" Text="聽錄音" />
                          
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
             &nbsp;<asp:Label ID="Lbl2" runat="server" ForeColor="Red" Width="396px" ></asp:Label>&nbsp;<br />
                      <asp:Label ID="lblErr1" runat="server" Font-Bold="False" Font-Size="X-Small" Text="訊息代碼:  0000.成功   "></asp:Label><br />
                      <asp:Label ID="lblErr2" runat="server" Font-Bold="False" Font-Size="X-Small" Text="E001.帳號不存在     E002.首次變更密碼      E003.密碼被鎖定   E004.密碼逾時需變更  E005.密碼錯誤"></asp:Label><br />
                      <asp:Label ID="lblErr3" runat="server" Font-Bold="False" Font-Size="X-Small" Text="E011.和前3次密碼相同   E997.未輸入資料  E998.資料庫暫停服務   E999.連線程式暫停服務"></asp:Label><br />
                    <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" Width="95px" Height="25px" />&nbsp;
                         <asp:button id="butMenu" runat="server" ForeColor="#0000C0" BackColor="#E0E0E0" BorderColor="#00AEEF" Text="回主選單" Width="95px" BorderStyle="Solid" Height="25px"></asp:button>
         </p>
         <p>
             &nbsp;</p>
        </td></tr>
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
