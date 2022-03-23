<%@ Page Language="VB" AutoEventWireup="true" EnableEventValidation = "false" CodeFile="v_log11.aspx.vb" Inherits="v_log11" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>分行基本資料</title>
    <style type="text/css">
      .style1
      {
        text-align: left;
      }
      .style2
      {
        text-align: right;
      }
      .style3
      {
        text-align: center;
      }
    </style>
</head>
<body style="background-color: #ccffff; font-family: 標楷體;">

    <form id="form1" runat="server">
    <div class="style1">
      <div class="style3">
    <asp:Image ID="Image1" runat="server" Height="86px" Width="760px" 
          ImageUrl="~/image/FAR_TOP2.JPG" style="text-align: center" />
      </div>
    <table style="width: 100%; height: 100%" border="1">
        <tr>
     <td width="97%" border="1" height="4%" class="style2">
            <asp:Label ID="Label6" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button2" runat="server" Width="50px" height="25px" Text="登出" 
            ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px" 
            style="font-family: 標楷體"></asp:button>
          </td> </tr>
        <tr>
     <td width="97%" align="left" border="1" height="4%" 
            style="background-color: gainsboro">
          <center><b> 
              <asp:Label ID="title" runat="server" Text="【11.分行資料維護】"></asp:Label></b>&nbsp;</center></td> </tr>
       <tr><td style="text-align: left; height: 16%;"> 
              <asp:label id="Label1" runat="server" text="分行代碼" width="15%" Height="80%" style="vertical-align: middle; text-align: center" Font-Bold="True"></asp:label>
              <asp:textbox id="TextBox1" runat="server" width="15%" Height="80%" MaxLength="3"></asp:textbox>
              <asp:label id="Label2" runat="server" text="分行名稱" width="15%" Height="80%" style="vertical-align: middle; text-align: center" Font-Bold="True"></asp:label> 
              <asp:textbox id="TextBox2" runat="server" width="15%" Height="80%" MaxLength="40"></asp:textbox>                          
              <asp:label id="Label3" runat="server" text="單位主管" width="15%" Height="80%" style="vertical-align: middle; text-align: center" Font-Bold="True"></asp:label>
              <asp:textbox id="TextBox3" runat="server" width="15%" Height="80%" MaxLength="10"></asp:textbox>
              
        </tr>
        <tr><td height="5%" style="text-align: left"> 
              <asp:label id="Label4" runat="server" text="區別名稱" width="15%" Height="80%" style="vertical-align: middle; text-align: center" Font-Bold="True"></asp:label>
              <asp:textbox id="TextBox4" runat="server" width="15%" Height="80%" MaxLength="20"></asp:textbox>
              <asp:label id="Label5" runat="server" text="區別督導" width="15%" Height="80%" style="vertical-align: middle; text-align: center" Font-Bold="True"></asp:label> 
              <asp:textbox id="TextBox5" runat="server" width="15%" Height="80%" MaxLength="6"></asp:textbox>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            
        </tr>
        <tr><td style="background-color: gainsboro; text-align: center; height: 4%;">
                  <asp:button id="Button1" runat="server" Width="95px" Text="新增" 
                ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid"></asp:button>
            &nbsp;</tr>
    <tr>    
    <td align="center" style="width: 100%; height: 301px;"  valign="top">
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False" DataKeyNames="BRA01"
        
            DataSourceID="SqlDataSource1" Width="100%" OnRowCommand="CustomersGridView_RowCommand" AllowPaging="True" AllowSorting="True" PageSize="50">
            
            <Columns>
                
                <asp:BoundField DataField="BRA01" HeaderText="分行代碼" ReadOnly="True" SortExpression="BRA01" />
                <asp:BoundField DataField="BRA01_" HeaderText="分行代碼_覆核" ReadOnly="True" SortExpression="BRA01_" >
                    <ItemStyle ForeColor="Blue" />
                </asp:BoundField>

                <asp:BoundField DataField="BRA02" HeaderText="分行名稱" SortExpression="BRA02" />
                <asp:BoundField DataField="BRA02_" HeaderText="分行名稱_覆核" ReadOnly="True" SortExpression="BRA02_" >
                    <ItemStyle ForeColor="Blue" />
                </asp:BoundField>

                <asp:BoundField DataField="BRA03" HeaderText="分行經理" SortExpression="BRA03" />
                <asp:BoundField DataField="BRA03_" HeaderText="分行經理_覆核" ReadOnly="True" SortExpression="BRA03_" >
                    <ItemStyle ForeColor="Blue" />
                </asp:BoundField>

                <asp:BoundField DataField="BRA04" HeaderText="區域督導" SortExpression="BRA04" />
                <asp:BoundField DataField="BRA04_" HeaderText="區域督導_覆核" ReadOnly="True" SortExpression="BRA04_" >
                    <ItemStyle ForeColor="Blue" />
                </asp:BoundField>

                <asp:BoundField DataField="BRA05" HeaderText="區別名稱" SortExpression="BRA05" />
                <asp:BoundField DataField="BRA05_" HeaderText="區別名稱_覆核" ReadOnly="True" SortExpression="BRA05_" >
                    <ItemStyle ForeColor="Blue" />
                </asp:BoundField>

                <asp:BoundField DataField="BRA_CHECK_TYPE" HeaderText="覆核_型別" ReadOnly="True"  SortExpression="BRA_CHECK_TYPE" >
                    <ItemStyle ForeColor="Red" />
                </asp:BoundField>
                <asp:BoundField DataField="BRA_CHECK" HeaderText="覆核" ReadOnly="True" SortExpression="BRA_CHECK" >
                    <ItemStyle ForeColor="Red" />
                </asp:BoundField>

               
                <asp:TemplateField InsertVisible="False" ShowHeader="False">
                
                    <EditItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update"
                            Text="更新"  Width="40px" >
                        </asp:LinkButton>
                        
                        <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                            Text="取消"  Width="40px">
                        </asp:LinkButton>
                        
                    </EditItemTemplate>
                    
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Edit"
                            Text="編輯" Width="40px" > 
                        </asp:LinkButton>
                            
                    </ItemTemplate>
                    
                </asp:TemplateField>
                
                
                <asp:TemplateField InsertVisible="False" ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton3" runat="server" CausesValidation="False" CommandName="Delete"
                            Text="刪除" CommandArgument='<%#Eval("BRA01")%>' Width="40px" >  
                        </asp:LinkButton>
                        </ItemTemplate>
                </asp:TemplateField>
                
                <asp:ButtonField Text="覆核" CommandName="CHECK"  >
                   <ItemStyle Width="40px" />
                </asp:ButtonField>
                
                
            </Columns>
            <PagerSettings Position="TopAndBottom" />
            <PagerStyle HorizontalAlign="Center" />
            
        </asp:GridView>
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
        
            ConnectionString="<%$ ConnectionStrings:FARConnectionString1 %>"
        
            SelectCommand="SELECT * FROM [BRA]" ConflictDetection="CompareAllValues" 
            DeleteCommand="UPDATE BRA SET  BRA02_ = '' , BRA03_ = '' , BRA04_ = '' , BRA05_ = '' , BRA_CHECK_TYPE = '刪除' , BRA_CHECK = 'Y' WHERE (BRA01 = @original_BRA01)" 
            OldValuesParameterFormatString="original_{0}" 
            UpdateCommand="UPDATE BRA SET BRA02_ = @BRA02, BRA03_ = @BRA03, BRA04_ = @BRA04, BRA05_ = @BRA05, BRA_CHECK_TYPE = '修改' , BRA_CHECK = 'Y'  WHERE (BRA01 = @original_BRA01)" 
            InsertCommand="INSERT INTO [BRA] ([BRA01], [BRA02], [BRA03], [BRA04], [BRA05]) VALUES (@BRA01, @BRA02, @BRA03, @BRA04, @BRA05)" ProviderName="<%$ ConnectionStrings:FARConnectionString1.ProviderName %>">
            
            <DeleteParameters>
                <asp:Parameter Name="original_BRA01" Type="String" />
            </DeleteParameters>
            
            <UpdateParameters>
                <asp:Parameter Name="BRA02" Type="String" />
                <asp:Parameter Name="BRA03" Type="String" />
                <asp:Parameter Name="BRA04" Type="String" />
                <asp:Parameter Name="BRA05" Type="String" />
                <asp:Parameter Name="original_BRA01" Type="String" />
            </UpdateParameters>
            
            <InsertParameters>
                <asp:Parameter Name="BRA01" Type="String" />
                <asp:Parameter Name="BRA02" Type="String" />
                <asp:Parameter Name="BRA03" Type="String" />
                <asp:Parameter Name="BRA04" Type="String" />
                <asp:Parameter Name="BRA05" Type="String" />
            </InsertParameters>
            
        </asp:SqlDataSource>
        
        
         <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" 
          Width="95px" Height="25px" style="font-family: 標楷體" />
                    
        &nbsp; <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" 
          ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid" Height="25px" 
          style="font-family: 標楷體"></asp:button>
                &nbsp;&nbsp;
        </td>
    </table>
    </div>
    </form>
    
</body>
</html>
