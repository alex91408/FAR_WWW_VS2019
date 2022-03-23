<%@ Page Language="VB" AutoEventWireup="true"　EnableEventValidation = "false" CodeFile="v_log12.aspx.vb" Inherits="v_log12" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>系統使用者維護</title>
    <style type="text/css">
      .style1
      {
        height: 8%;
      }
      .style2
      {
        font-family: 標楷體;
      }
      .style3
      {
        text-align: center;
      }
    </style>
</head>
<body style="background-color: #ccffff; font-family: 標楷體;">

    <form id="form1" runat="server">
    <div style="text-align: center">
    <asp:Image ID="Image1" runat="server" Height="86px" Width="760px" ImageUrl="~/image/FAR_TOP2.JPG" />
    <br />
    <table style="width: 100%; height: 100%" border="1">
    
    <tr>
     <td width="97%" align="left" border="1" height="4%" style="text-align: right;">
            <asp:Label ID="Label6" runat="server" Text="Label" BackColor="#FFC080" BorderColor="Black" BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" height="23px"></asp:Label>
            <asp:button id="Button3" runat="server" Width="50px" height="25px" Text="登出" 
            ForeColor="Black" BorderColor="black" Font-Size="Medium"
                BackColor="#E0E0E0" BorderStyle="Solid" BorderWidth="1px" 
            style="font-family: 標楷體"></asp:button>
        </td> </tr>
    
    <tr>
     <td width="97%" align="left" border="1" height="4%" 
        style="background-color: gainsboro">
          <center><b> <span style="font-size: 12pt; font-family: 細明體; mso-bidi-font-family: 'Times New Roman';
                  mso-font-kerning: 1.0pt; mso-ansi-language: EN-US; mso-fareast-language: ZH-TW;
                  mso-bidi-language: AR-SA">
              <asp:Label ID="title" runat="server" Text="【12.系統使用者資料維護】" 
              style="font-family: 標楷體"></asp:Label></span></b></center></td> </tr>
       <tr><td height="5%" style="text-align: left"> 
              <asp:label id="Label1" runat="server" text="分行代碼" width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label>
           
           <asp:DropDownList ID="DropDownList1" runat="server"  width="15%">
               <asp:ListItem></asp:ListItem>
           </asp:DropDownList>
           
           <asp:label id="Label2" runat="server" text="行員編號"  width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label> 
              <asp:textbox id="TextBox2" runat="server" width="15%" Height="80%" 
                MaxLength="6" CssClass="style2"></asp:textbox>                          
              <asp:label id="Label3" runat="server" text="行員姓名"  width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label>
              <asp:textbox id="TextBox3" runat="server"  width="15%" Height="80%" 
                CssClass="style2"></asp:textbox>
              
        </tr>
        <tr><td style="text-align: left; height: 5%;"> 
              <asp:label id="Label4" runat="server" text="權限等級" width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label>
              <asp:DropDownList ID="DropDownList3" runat="server" width="15%" 
                CssClass="style2">
                  <asp:ListItem></asp:ListItem>
                <asp:ListItem Value="1">1.總行層級</asp:ListItem>
                <asp:ListItem Value="2">2.分行層級</asp:ListItem>
                <asp:ListItem Value="3">3.理專層級</asp:ListItem>
                <asp:ListItem Value="4">4.區域層級</asp:ListItem>
            </asp:DropDownList>
             <asp:label id="Label5" runat="server" text="功能群組"  width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label>
             
            <asp:DropDownList ID="DropDownList2" runat="server" width="15%" 
                DataSourceID="SqlDataSource6" DataTextField="GRP02" DataValueField="GRP01" 
                CssClass="style2">
            
            </asp:DropDownList>
            
             <asp:label id="Label7" runat="server" text="email"  width="15%" Height="80%" 
                style="vertical-align: middle; text-align: center" Font-Bold="True" 
                CssClass="style2"></asp:label> 
              <asp:textbox id="TextBox6" runat="server" width="15%" Height="80%" 
                CssClass="style2"></asp:textbox>
              
                   <tr>
        
        <td style="background-color: gainsboro; text-align: center; " class="style1">
                  <asp:button id="Button2" runat="server" Width="95px" Text="搜尋" 
                ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid" style="font-family: 標楷體"></asp:button>
                   <asp:button id="Button1" runat="server" Width="95px" Text="新增" 
                ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid" style="font-family: 標楷體"></asp:button>
                  <br />
            <asp:Label ID="lblStat" runat="server" ForeColor="Red" Height="22px" Width="87px" 
                    Visible="False">新增時</asp:Label>
            <asp:Label ID="lblStat2" runat="server" ForeColor="Red" Height="22px" Width="270px" 
                    Visible="False">，權限等級限定(3.理專層級)</asp:Label>
            <asp:Label ID="lblStat3" runat="server" ForeColor="Red" Height="22px" Width="296px" 
                    Visible="False">，功能群組限定(11.系統預設)</asp:Label>
        
        </tr>   
        
             </tr>
  
              
        <asp:GridView ID="GridView1" runat="server" AutoGenerateColumns="False"
            DataSourceID="SqlDataSource2" Width="100%" Height ="105%" DataKeyNames="SAL01" OnRowCommand="CustomersGridView_RowCommand" AllowPaging="True" AllowSorting="True" PageSize="50">
            <Columns>
             <asp:ButtonField CommandName="restore" Text="密碼重置" ButtonType="Button" >
                            <ControlStyle BackColor="White" BorderStyle="Solid" BorderWidth="1px" />
                        </asp:ButtonField>
                <asp:TemplateField HeaderText="分行代碼" SortExpression="SAL03">
                    <EditItemTemplate>
                    
                        <asp:DropDownList ID="DropDownList4" runat="server" DataSourceID="SqlDataSource1"
                            DataTextField="BRA02" DataValueField="BRA01" SelectedValue='<%# Bind("SAL03") %>'>
                        </asp:DropDownList>&nbsp;
                    
                    </EditItemTemplate>
                    
                    <ItemTemplate>
                        <asp:Label ID="Label1" runat="server" Text='<%# Bind("SAL03") %>'></asp:Label>
                    </ItemTemplate>
                    
                </asp:TemplateField>
                
                <asp:BoundField DataField="SAL03_" HeaderText="分行代碼覆核" ReadOnly="True"  SortExpression="SAL03_" />
                
                <asp:BoundField DataField="SAL01" HeaderText="行員編號" ReadOnly="True" SortExpression="SAL01" >
                    <ItemStyle Width="40px" />
                </asp:BoundField>
                
                <asp:BoundField DataField="SAL02" HeaderText="行員姓名" SortExpression="SAL02" >
                    <ItemStyle Width="50px" />
                </asp:BoundField>
                <asp:BoundField DataField="SAL02_" HeaderText="行員姓名覆核" ReadOnly="True" SortExpression="SAL02_" />
              
              
              <asp:TemplateField HeaderText="功能群組" SortExpression="SAL04">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList5" runat="server" DataSourceID="SqlDataSource3"
                            DataTextField="GRP02" DataValueField="GRP01" SelectedValue='<%# Bind("SAL04") %>'>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label3" runat="server" Text='<%# Bind("SAL04") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>

                <asp:BoundField DataField="SAL04_" HeaderText="功能群組覆核" ReadOnly="True" SortExpression="SAL04_" />


                <asp:TemplateField HeaderText="權限層級" SortExpression="SAL05">
                    <EditItemTemplate>
                        <asp:DropDownList ID="DropDownList6" runat="server" SelectedValue='<%# Bind("SAL05") %>'>
                            <asp:ListItem Value="1">1.總行層級</asp:ListItem>
                            <asp:ListItem Value="2">2.分行層級</asp:ListItem>
                            <asp:ListItem Value="3">3.理專層級</asp:ListItem>
                            <asp:ListItem Value="4">4.區域層級</asp:ListItem>
                        </asp:DropDownList>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("SAL05") %>'></asp:Label>
                    </ItemTemplate>
                </asp:TemplateField>
                
                <asp:BoundField DataField="SAL05_" HeaderText="權限層級覆核" ReadOnly="True" SortExpression="SAL05_" />

                
                <asp:BoundField DataField="SAL08" HeaderText="EMAIL" SortExpression="SAL08" />
                <asp:BoundField DataField="SAL08_" HeaderText="EMAIL覆核" ReadOnly="True" SortExpression="SAL08_" />
                
                <asp:BoundField DataField="SAL06" HeaderText="密碼錯誤次數" SortExpression="SAL06" ReadOnly="True"/>
                <asp:BoundField DataField="SAL07" HeaderText="錯誤時間" SortExpression="SAL07" ReadOnly="True"/>
               
               
               <asp:BoundField DataField="SAL_CHECK_TYPE" HeaderText="覆核型別" ReadOnly="True"  SortExpression="SAL_CHECK_TYPE" >
                     <ItemStyle Width="40px" ForeColor="Red" />
                </asp:BoundField>
                
                <asp:BoundField DataField="SAL_CHECK" HeaderText="覆核" ReadOnly="True" SortExpression="SAL_CHECK" >
                    <ItemStyle ForeColor="Red" />
                </asp:BoundField>
               
               
                <asp:TemplateField ShowHeader="False">
                    <EditItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update"
                            Text="更新"  Width="40px"  ></asp:LinkButton>
                        <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                            Text="取消"  Width="40px"  ></asp:LinkButton>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Edit"
                            Text="編輯"  Width="40px"  ></asp:LinkButton>
                     </ItemTemplate>
                </asp:TemplateField>
                
                <asp:TemplateField ShowHeader="False">
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton3" runat="server" CausesValidation="False" CommandName="Delete"
                            Text="刪除" CommandArgument='<%#Eval("sal01")%>'   Width="40px" ></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                
                <asp:ButtonField Text="覆核" CommandName="CHECK" > 
                   <ItemStyle Width="40px" />
                </asp:ButtonField >
        
                
                
            </Columns>
            <PagerSettings Mode="NumericFirstLast" Position="TopAndBottom" />
            <HeaderStyle HorizontalAlign="Center" VerticalAlign="Top" />
            <PagerStyle HorizontalAlign="Center" />
        </asp:GridView>
        
        
         

    <tr>    
    <td align="center" style="width: 100%; text-align: center;"  valign="top">
        
        
        <asp:SqlDataSource ID="SqlDataSource2" runat="server" 
        
             ConnectionString="<%$ ConnectionStrings:ConnectionString2 %> " 

             DeleteCommand="UPDATE SAL SET  SAL01_ ='', SAL02_ ='', SAL03_ ='', SAL04_='', SAL05_='', SAL08_='',  SAL09_='' , SAL_CHECK_TYPE = '刪除' , SAL_CHECK = 'Y'   WHERE (SAL01 = @original_SAL01)" 
             OldValuesParameterFormatString="original_{0}" 
             
             UpdateCommand="UPDATE SAL SET SAL02_ = @SAL02, SAL03_ = @SAL03, SAL04_ = @SAL04, SAL05_ = @SAL05, SAL08_ = @SAL08 , SAL_CHECK_TYPE = '修改' , SAL_CHECK = 'Y'  WHERE (SAL01 = @original_SAL01)" >
       
  

            <DeleteParameters>
                <asp:Parameter Name="original_SAL01" />
            </DeleteParameters>


            <UpdateParameters>
                <asp:Parameter Name="SAL02" />
                <asp:Parameter Name="SAL03" />
                <asp:Parameter Name="SAL04" />
                <asp:Parameter Name="SAL05" />
                <asp:Parameter Name="SAL08" />
                <asp:Parameter Name="original_SAL01" />
            </UpdateParameters>

        </asp:SqlDataSource>
        
        
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString3 %>"
            SelectCommand="SELECT * FROM [BRA] where BRA_CHECK_TYPE <> '新增'">
        </asp:SqlDataSource>
        
        <asp:SqlDataSource ID="SqlDataSource3" runat="server" ConnectionString="<%$ ConnectionStrings:ConnectionString2 %>"
            SelectCommand="SELECT * FROM [GRP]">
        </asp:SqlDataSource>
        
        
        <asp:SqlDataSource ID="SqlDataSource6" runat="server" ConnectionString="<%$ ConnectionStrings:FARConnectionString1 %>"
            SelectCommand="SELECT [GRP01], [GRP02] FROM [search_GRP]">
        </asp:SqlDataSource>
            
            
        </td>
    </table>
    </div>
    <p class="style3">
            
            
                   <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" Width="95px" Height="25px" />
        &nbsp; <asp:button id="butMenu" runat="server" Width="95px" Text="回主選單" ForeColor="#0000C0" BackColor="#E0E0E0"
                    BorderColor="#00AEEF" BorderStyle="Solid" Height="25px"></asp:button>
        </p>
    </form>
    
</body>
</html>
