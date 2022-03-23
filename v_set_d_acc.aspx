<%@ Page Language="VB" AutoEventWireup="true" EnableEventValidation = "false" CodeFile="v_set_d_acc.aspx.vb" Inherits="v_set_d_acc" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    
  <style type="text/css">
    .style1
    {}
    .style2
    {
      text-align: left;
    }
  </style>
    
</head>


<script type="text/javascript">

function choiceUser(fieldClientID,isMutilSelect){
window.open('UserSelector.aspx?fieldClientID='+fieldClientID.id+'&mutilSelect='+isMutilSelect,'','width=220,height=220','');
}
function choiceDate(fieldClientID){
window.open('dtPicker.aspx?fieldClientID='+fieldClientID.id,'','width=220,height=220','');
}

</script>





<body style="text-align: center; background-color: #ccffff;" bgcolor="#99ccff" leftmargin="0" rightmargin="0" topmargin="0">
    <form id="form1" runat="server">
    <div>
      <table style="height: 132px; left: 0px; top: 0px; text-align: center;">
        <tr>
          <td align="left" style="height: 118px">
              <asp:Image ID="Image2" runat="server" Height="86px" ImageUrl="~/image/FAR_TOP5.JPG"
                  Width="1213px" style="text-align: center" /></td>
        </tr>
        <tr style="font-size: 12pt">
          <td align="left">
            <p align="left">
                <div class="style2">
                &nbsp;</div>

                <%--20220223 old <asp:SqlDataSource ID="SqlDataSource1" runat="server" ProviderName="<%$ ConnectionStrings:FARConnectionString1.ProviderName %>" --%>
                <%--20220223 new --%>

                <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
                    
                   ConnectionString="<%$ ConnectionStrings:FARConnectionString1 %>"

                   SelectCommand="SELECT * FROM [GRP]" 
                   DeleteCommand="UPDATE GRP SET GRP_CHECK_TYPE = '刪除' , GRP_CHECK = 'Y' where (GRP01 ='999') "
                   OldValuesParameterFormatString="original_{0}"                  
                   
                   UpdateCommand="UPDATE GRP SET GRP_CHECK_TYPE = '修改' , GRP_CHECK = 'Y' , GRP02_ = @GRP02, v01_ = @v01, v01a_ = @v01a, v02_ = @v02, v02a_ =@v02a, v03_ =@v03, v03a_ =@v03a, v04_ =@v04, v04a_ =@v04a, v05_ =@v05, v05a_ =@v05a, v05b_ =@v05b, v11_ =@v11, v11a_ =@v11a, v11b_ =@v11b, v11c_ =@v11c, v11d_ =@v11d, v11e_ = @v11e, v12_ =@v12, v12a_ =@v12a, v12b_ =@v12b, v12c_ =@v12c, v12d_ =@v12d, v12e_ =@v12e, v12f_ =@v12f, v12g_ =@v12g, v12h_ =@v12h, v12i_ =@v12i, v13_ =@v13, v13a_ =@v13a, v13b_ =@v13b, v13c_ =@v13c, v13d_ =@v13d, v14_ =@v14, v14a_ =@v14a where GRP01 = '99A'  " 
                   
                   InsertCommand="INSERT INTO [GRP] ([GRP01]) VALUES (@GRP01) "  >
 
                   
                           
                    
                    <UpdateParameters>
                        <asp:Parameter Name="GRP01" Type="String" />
                    </UpdateParameters>
     
     
     
                   <DeleteParameters>
                     <asp:Parameter Name="GRP01"  Type="String" />
                   </DeleteParameters>
                     
                    
                </asp:SqlDataSource>
                </p>
              <p align="left" class="style2">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  &nbsp;<asp:Label ID="Label1" runat="server" BackColor="#FFC080" BorderColor="Black"
                    BorderStyle="Solid" BorderWidth="1px" Font-Size="Medium" Height="23px" Text="Label"></asp:Label>
                  &nbsp;
                  <asp:Button ID="Button1" runat="server" BackColor="#E0E0E0" BorderColor="black" BorderStyle="Solid"
                    BorderWidth="1px" Font-Size="Medium" ForeColor="Black" Height="25px" Text="登出"
                    Width="50px" />
                  &nbsp;
              </p>
              <p class="style2">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:Label ID="lblMDB" runat="server" Font-Size="Large" ForeColor="Blue" Text="lblMDB"></asp:Label><asp:Label ID="title" runat="server" Font-Bold="True" Text="【13.群組管理】"></asp:Label></p>
              <p class="style2">
                  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  群組代碼
                  <asp:TextBox ID="txtGRP01" runat="server" BackColor="#C0C0FF" BorderColor="White" MaxLength="2"
                      Style="z-index: 108; left: 255px; top: 570px" Width="82px"></asp:TextBox>
                  群組名稱
                  <asp:TextBox ID="txtGRP02" runat="server" BackColor="#C0C0FF" BorderColor="White"
                      MaxLength="40" Style="z-index: 108; left: 255px; top: 570px" Width="82px"></asp:TextBox>&nbsp; 
              <asp:Button ID="addGRP" runat="server" BorderColor="#00AEEF" BorderStyle="Solid"
                CausesValidation="False" ForeColor="#0000C0" Height="24" Style="z-index: 107; left: 192px;
                top: 569px" Text="新增" Width="51px" /></p>
            <p class="style2">
              &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <asp:Label ID="lblStat" runat="server" ForeColor="Blue" Height="22px" Style="z-index: 106;
                left: 116px; top: 603px" Visible="False" Width="534px">lblStat</asp:Label></p>
              <p class="style2">
                  &nbsp; &nbsp; &nbsp;&nbsp;<asp:GridView ID="GridView1" runat="server" AllowPaging="True" AllowSorting="True"  OnRowCommand="CustomersGridView_RowCommand"
                      AutoGenerateColumns="False" BackColor="White" BorderColor="#999999" CellPadding="3"
                      DataKeyNames="GRP01" 
                      DataSourceID="SqlDataSource1" 
                      EmptyDataText="查不到符合的資料"
                      EnableTheming="True" 
                      Width="6000px" 
                      BorderStyle="None" 
                      BorderWidth="1px" 
                      GridLines="Vertical" 
                      HorizontalAlign="Left"
                      PageSize="20">
                      
                      <FooterStyle BackColor="#CCCCCC" ForeColor="Black" />
                      
                      <RowStyle BackColor="#EEEEEE" ForeColor="Black" HorizontalAlign="Left" />
                      
                      <EmptyDataRowStyle BackColor="#C0FFFF" BorderColor="#C0FFFF" Font-Bold="False" Font-Size="Medium" />
              
                      <Columns>
      
    
                        <asp:ButtonField Text="覆核" CommandName="CHECK"  >
                          <ItemStyle Width="40px" />
                        </asp:ButtonField>
    
    
                        <asp:BoundField DataField="GRP_CHECK_TYPE" HeaderText="覆核型別" ReadOnly="True"  SortExpression="GRP_CHECK_TYPE" >
                             <ItemStyle Width="40px" ForeColor="Red" />
                        </asp:BoundField>
                        
                        <asp:BoundField DataField="GRP_CHECK" HeaderText="覆核" ReadOnly="True" SortExpression="GRP_CHECK" >
                            <ItemStyle Width="40px" ForeColor="Red" />
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
                                    Text="刪除" CommandArgument='<%#Eval("GRP01")%>' Width="40px" >  
                                </asp:LinkButton>
                                </ItemTemplate>
                        </asp:TemplateField>
                   
         
 
                    
                    
                     
                          <asp:BoundField DataField="GRP01" HeaderText="群組代碼" SortExpression="GRP01"  ReadOnly="True" >
                          <ItemStyle Width="40px" />
                          </asp:BoundField>

                           
                          <asp:BoundField DataField="GRP01_" HeaderText="群組代碼(覆核值)" SortExpression="GRP01_" ReadOnly="True" > 
                          <ItemStyle Width="40px" />
                          </asp:BoundField>
                    
                          <asp:BoundField DataField="GRP02" HeaderText="群組名稱" SortExpression="GRP02" >
                               <ItemStyle Width="150px" />
                               <HeaderStyle Wrap="True" />
                          </asp:BoundField>
 
                           <asp:BoundField DataField="GRP02_" HeaderText="群組名稱(覆核值)" SortExpression="GRP02_" ReadOnly="True">
                              <ItemStyle Width="150px" Wrap="True" />
                              <HeaderStyle Wrap="True" /> 
                          </asp:BoundField>
                   
                   
                          <asp:CheckBoxField DataField="v01" HeaderText="v01.電話錄音線路使用統計表" SortExpression="v01" />
                          <asp:CheckBoxField DataField="v01_" HeaderText="v01_.電話錄音線路使用統計表(覆核值)" SortExpression="v01_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v01a" HeaderText="v01a.儲存" SortExpression="v01a" />
                          <asp:CheckBoxField DataField="v01a_" HeaderText="v01a_.儲存(覆核值)" SortExpression="v01a_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v02" HeaderText="v02.電話錄音進線時段統計表" SortExpression="v02" />
                          <asp:CheckBoxField DataField="v02_" HeaderText="v02_.電話錄音進線時段統計表(覆核值)" SortExpression="v02_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v02a" HeaderText="v02a.儲存" SortExpression="v02a" />
                          <asp:CheckBoxField DataField="v02a_" HeaderText="v02a_.儲存(覆核值)" SortExpression="v02a_"  ReadOnly="True"/>
 
                          <asp:CheckBoxField DataField="v03" HeaderText="v03.各項服務統計表" SortExpression="v03" />
                          <asp:CheckBoxField DataField="v03_" HeaderText="v03_.各項服務統計表(覆核值)" SortExpression="v03_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v03a" HeaderText="v03a.儲存" SortExpression="v03a" />
                          <asp:CheckBoxField DataField="v03a_" HeaderText="v03a_.儲存(覆核值)" SortExpression="v03a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v04" HeaderText="v04. 各分行電話錄音統計表" SortExpression="v04" />
                          <asp:CheckBoxField DataField="v04_" HeaderText="v04_. 各分行電話錄音統計表(覆核值)" SortExpression="v04_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v04a" HeaderText="v04a.儲存" SortExpression="v04a" />
                          <asp:CheckBoxField DataField="v04a_" HeaderText="v04a_.儲存(覆核值)" SortExpression="v04a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05" HeaderText="v05.各項服務明細查詢" SortExpression="v05" />
                          <asp:CheckBoxField DataField="v05_" HeaderText="v05_.各項服務明細查詢(覆核值)" SortExpression="v05_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05a" HeaderText="v05a.儲存" SortExpression="v05a" />
                          <asp:CheckBoxField DataField="v05a_" HeaderText="v05a_.儲存(覆核值)" SortExpression="v05a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v05b" HeaderText="v05b.下載" SortExpression="v05b" />
                          <asp:CheckBoxField DataField="v05b_" HeaderText="v05b_.下載(覆核值)" SortExpression="v05b_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11" HeaderText="v11.分行資料維護" SortExpression="v11" />
                          <asp:CheckBoxField DataField="v11_" HeaderText="v11_.分行資料維護(覆核值)" SortExpression="v11_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11a" HeaderText="v11a.儲存" SortExpression="v11a" />
                          <asp:CheckBoxField DataField="v11a_" HeaderText="v11a_.儲存(覆核值)" SortExpression="v11a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11b" HeaderText="v11b.新增" SortExpression="v11b" />
                          <asp:CheckBoxField DataField="v11b_" HeaderText="v11b_.新增(覆核值)" SortExpression="v11b_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11c" HeaderText="v11c.刪除" SortExpression="v11c" />
                          <asp:CheckBoxField DataField="v11c_" HeaderText="v11c_.刪除(覆核值)" SortExpression="v11c_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v11d" HeaderText="v11d.編輯" SortExpression="v11d" />
                          <asp:CheckBoxField DataField="v11d_" HeaderText="v11d_.編輯(覆核值)" SortExpression="v11d_"  ReadOnly="True"/>
                          
                          <asp:CheckBoxField DataField="v11e" HeaderText="v11e.覆核" SortExpression="v11e" />
                          <asp:CheckBoxField DataField="v11e_" HeaderText="v11e_.覆核(覆核值)" SortExpression="v11e_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12" HeaderText="v12.系統使用者資料維護" SortExpression="v12" />
                          <asp:CheckBoxField DataField="v12_" HeaderText="v12_.系統使用者資料維護(覆核值)" SortExpression="v12_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12a" HeaderText="v12a.儲存" SortExpression="v12a" />
                          <asp:CheckBoxField DataField="v12a_" HeaderText="v12a_.儲存(覆核值)" SortExpression="v12a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12b" HeaderText="v12b.查詢" SortExpression="v12b" />
                          <asp:CheckBoxField DataField="v12b_" HeaderText="v12b_.查詢(覆核值)" SortExpression="v12b_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12c" HeaderText="v12c.新增" SortExpression="v12c" />
                          <asp:CheckBoxField DataField="v12c_" HeaderText="v12c_.新增(覆核值)" SortExpression="v12c_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12d" HeaderText="v12d.密碼重置" SortExpression="v12d" />
                          <asp:CheckBoxField DataField="v12d_" HeaderText="v12d_.密碼重置(覆核值)" SortExpression="v12d_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12e" HeaderText="v12e.編輯" SortExpression="v12e" />
                          <asp:CheckBoxField DataField="v12e_" HeaderText="v12e_.編輯(覆核值)" SortExpression="v12e_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12f" HeaderText="v12f.刪除" SortExpression="v12f" />
                          <asp:CheckBoxField DataField="v12f_" HeaderText="v12f_.刪除(覆核值)" SortExpression="v12f_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12g" HeaderText="v12g.功能群組" SortExpression="v12g" />
                          <asp:CheckBoxField DataField="v12g_" HeaderText="v12g_.功能群組(覆核值)" SortExpression="v12g_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v12h" HeaderText="v12h.權限等級" SortExpression="v12h" />
                          <asp:CheckBoxField DataField="v12h_" HeaderText="v12h_.權限等級(覆核值)" SortExpression="v12h_"  ReadOnly="True"/>
                          
                          <asp:CheckBoxField DataField="v12i" HeaderText="v12i.覆核" SortExpression="v12i" />
                          <asp:CheckBoxField DataField="v12i_" HeaderText="v12i_.覆核(覆核值)" SortExpression="v12i_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13" HeaderText="v13.群組維護" SortExpression="v13" />
                          <asp:CheckBoxField DataField="v13_" HeaderText="v13_.群組維護(覆核值)" SortExpression="v13_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13a" HeaderText="v13a.儲存" SortExpression="v13a" />
                          <asp:CheckBoxField DataField="v13a_" HeaderText="v13a_.儲存(覆核值)" SortExpression="v13a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13b" HeaderText="v13b.新增" SortExpression="v13b" />
                          <asp:CheckBoxField DataField="v13b_" HeaderText="v13b_.新增(覆核值)" SortExpression="v13b_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13c" HeaderText="v13c.刪除" SortExpression="v13c" />
                          <asp:CheckBoxField DataField="v13c_" HeaderText="v13c_.刪除(覆核值)" SortExpression="v13c_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13d" HeaderText="v13d.編輯" SortExpression="v13d" />
                          <asp:CheckBoxField DataField="v13d_" HeaderText="v13d_.編輯(覆核值)" SortExpression="v13d_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v13e" HeaderText="v13e.覆核" SortExpression="v13e" />
                          <asp:CheckBoxField DataField="v13e_" HeaderText="v13e_.覆核(覆核值)" SortExpression="v13e_" ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v14" HeaderText="v14.使用者&amp;權限異動明細查詢" SortExpression="v14" />
                          <asp:CheckBoxField DataField="v14_" HeaderText="v14_.使用者&amp;權限異動明細查詢(覆核值)" SortExpression="v14_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v14a" HeaderText="v14a.儲存" SortExpression="v14a" />
                          <asp:CheckBoxField DataField="v14a_" HeaderText="v14a_.儲存(覆核值)" SortExpression="v14a_"  ReadOnly="True"/>

                          <asp:CheckBoxField DataField="v0b" HeaderText="v0b" SortExpression="v0b" Visible="False" />

                          <asp:CheckBoxField DataField="vemail" HeaderText="email" SortExpression="vemail" Visible="False" />

                      </Columns>
                    
                    
                      <PagerStyle BackColor="#999999" ForeColor="Black" HorizontalAlign="Left" />
                      <SelectedRowStyle BackColor="Red" Font-Bold="True" ForeColor="Fuchsia" HorizontalAlign="Left" BorderColor="Yellow" BorderStyle="Double" />
                      <HeaderStyle BackColor="#8080FF" Font-Bold="False" ForeColor="White" VerticalAlign="Top" HorizontalAlign="Left" Font-Size="Medium" />
                      <AlternatingRowStyle BackColor="Gainsboro" />
                      <PagerSettings Position="TopAndBottom" Mode="NumericFirstLast" />
                      <EditRowStyle BackColor="#C0FFC0" BorderColor="Red" BorderStyle="Dotted" />
                      
                  </asp:GridView>
                  
              </p>
          </td>
        </tr>
        <tr style="font-size: 12pt">
          <td align="left">
            &nbsp;</td>
        </tr>
      </table>
        <div class="style1">
        <br />
                  <asp:Button ID="SaveAsExecl" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                        BorderStyle="Solid" ForeColor="#0000C0" Text="儲存Excel" Width="95px" Height="25px" />
        &nbsp;&nbsp;
                <asp:Button ID="butMenu" runat="server" BackColor="#E0E0E0" BorderColor="#00AEEF"
                BorderStyle="Solid" ForeColor="#0000C0" Text="回主選單" Width="95px" Height="25px" />
      </div>
    </div>
    <asp:TextBox ID="txtT" runat="server" Visible="False"></asp:TextBox>
    </form>
</body>
</html>
