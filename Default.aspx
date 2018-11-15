<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="FCPortal._Default" Debug="true" ValidateRequest="false" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>FC Portal,2013</title>
    <link href="fcportal.css" rel="stylesheet" />
    <link id="favicon" href="favicon.ico" rel="icon" type="image/x-icon" />     
    <script  type="text/javascript" src="data.ashx?js=fcportal.js" ></script><%--data.ashx?js=--%>
</head>
<body>
    <form id="Form1" runat="server">
        <ext:ResourceManager ID="ResourceManager1" runat="server" />      
        <ext:Viewport runat="server" Layout="BorderLayout">
            <Items>
                <ext:Panel runat="server" Region="Center" Layout="FitLayout">
                    <Items>
                        <ext:GridPanel runat="server" ID="gridMain"  AutoScroll="True">
                            <Store>
                                <ext:Store ID="storeMain" runat="server"  PageSize="50" RemoteSort="True" OnSubmitData="storeWinFile_Submit">
                                    <Proxy>
                                        <ext:AjaxProxy Url="data.ashx">
                                            <ActionMethods Read="GET" />
                                            <Reader>
                                                <ext:JsonReader Root="data" TotalProperty="total" />
                                            </Reader>
                                        </ext:AjaxProxy>
                                    </Proxy>
                                    <Model>
                                        <ext:Model ID="Model1" runat="server">
                                            <Fields>
                                                <ext:ModelField Name="NO" />
                                                <ext:ModelField Name="fcDescription" />
                                                <ext:ModelField Name="contractor"  />
                                                <ext:ModelField Name="bugdet" Type="Float"/>
                                                <ext:ModelField Name="netAmount" Type="Float"/>
                                                <ext:ModelField Name="unusedBudget" Type="Float"/>
                                                <ext:ModelField Name="unusedPercentage" Type="Float"/>
                                                <ext:ModelField Name="remainingPercentage" Type="Float"/>
                                                <ext:ModelField Name="score" Type="Float"/>
                                                <ext:ModelField Name="workCenter"/>
                                                <ext:ModelField Name="pricingScheme"/>
                                                <ext:ModelField Name="contractAdmin"/>
                                                <ext:ModelField Name="buyer"/>
                                                <ext:ModelField Name="mainCoordinator"/>
                                                <ext:ModelField Name="userRepresentative"/>
                                                <ext:ModelField Name="contact"/>
                                                <ext:ModelField Name="telephone"/>
                                                <ext:ModelField Name="validateDate" Type="Date"/>
                                                <ext:ModelField Name="expireDate" Type="Date"  />                                        
                                            </Fields>
                                        </ext:Model>
                                    </Model> 
                                    <Listeners>
                                        <BeforeLoad Handler="preLoad();"/>
                                        <Load Handler="App.RowSelectionModelMain.selectedData=[{'rowIndex':0}];App.RowSelectionModelMain.view.panel.initSelectionData();afterLoad();"></Load>
                                    </Listeners>                                             
                            
                                </ext:Store>                        
                            </Store>
                            <ColumnModel ID="ColumnModel1" runat="server">
                                <Columns>
                                    <ext:RowNumbererColumn ID="RowNumbererColumn1" runat="server" Width="30" />
                                    <ext:Column ID="Column1" runat="server" Text="Doc. NO." DataIndex="NO" Width="90" Align="Center" >
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column5" runat="server" Text="Description" Width="80" DataIndex="fcDescription" Flex="1">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column6" runat="server" Text="Contractor" Width="80" DataIndex="contractor"  Flex="1">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column16" runat="server" Text="Pricing Scheme" Width="120" DataIndex="pricingScheme">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column7" runat="server" Text="Budget" Width="80" DataIndex="bugdet" Align="Right">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column8" runat="server" Text="Checked Value" Width="100" DataIndex="netAmount" Align="Right">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column9" runat="server" Text="Remaining Budget" Width="100" DataIndex="unusedBudget" Align="Right">
                                        <Renderer Fn="moneyFormat" />
                                    </ext:Column>
                                    <ext:Column ID="Column10" runat="server" Text="Remaining Budget %" Width="120" DataIndex="unusedPercentage" Align="Right">
                                        <Renderer Fn="percentageFormat" />
                                    </ext:Column> 
                                    <ext:Column ID="Column33" runat="server" Text="Remaining Duration %" Width="120" DataIndex="remainingPercentage" Align="Right">
                                        <Renderer Fn="percentageFormat" />
                                    </ext:Column>                            
                                    <ext:ComponentColumn ID="ComponentColumn1" runat="server" Text="Score" DataIndex="score" Width="160">
                                        <Component>
                                            <ext:ProgressBar ID="ProgressBar1" runat="server" Text="Progress"/>
                                        </Component>
                                        <Listeners>
                                            <Bind Handler="cmp.updateProgress(record.get('score')/100,getStatus(record),false)" /> 
                                        </Listeners>
                                    </ext:ComponentColumn>                            
                                </Columns>                        
                            </ColumnModel>                    
                            <SelectionModel>
                                <ext:RowSelectionModel ID="RowSelectionModelMain" runat="server" Mode="Single" >
                                    <Listeners>
                                        <Select Handler="selectEvent(record,#{gridSub}.store);"></Select>
                                    </Listeners>                                                                                                               
                                </ext:RowSelectionModel>                        
                            </SelectionModel>
                            <View>
                                <ext:GridView ID="GridView1" runat="server" StripeRows="true" />                                           
                            </View>            
                            <BottomBar>
                                    <ext:PagingToolbar ID="PagingToolbar1" runat="server">
                                    <Items>
                                        <ext:ToolbarSeparator />
                                        <ext:Checkbox ID="cbValid" runat="server" Checked="true" FieldLabel="Valid Only" LabelWidth="60">
                                            <Listeners>
                                                <Change Handler="if(#{hiddenValue}.getValue()=='1'){return;};App.storeMain.load();if(#{cbValid}.getValue()==1){#{cbContractStatus}.setValue('Valid Contracts')}else{#{cbContractStatus}.setValue('All Contracts')};"></Change>
                                            </Listeners>
                                        </ext:Checkbox>
                                    </Items>
                                    <Plugins>
                                        <ext:ProgressBarPager ID="ProgressBarPager1" runat="server" />
                                    </Plugins>
                                </ext:PagingToolbar>
                            </BottomBar>
                            <TopBar>
                                <ext:Toolbar ID="Toolbar1" runat="server" StyleSpec="text-align:right;">
                                    <Items>
                                        <ext:ComboBox ID="cbContractStatus" runat="server" Icon="Magnifier" ForceSelection="true" AllowBlank="true">                                            
                                             <Listeners>
                                                <Select Handler="if(#{cbContractStatus}.getValue()=='Valid Contracts'){#{cbValid}.setValue(1);return;};#{cbValid}.setValue(0);"></Select>
                                            </Listeners>
                                        </ext:ComboBox>
                                        <ext:ComboBox ID="cbContractDescription" runat="server" Icon="Magnifier" ForceSelection="true" AllowBlank="true" EmptyText="Description">
                                            <Listeners>
                                                <Select Handler="#{cbContractor}.clearValue();#{searchKeyword}.setValue(#{cbContractDescription}.getValue());searchClick();"></Select>
                                            </Listeners>
                                        </ext:ComboBox>
                                        <ext:ComboBox ID="cbContractor" runat="server" Icon="Magnifier" ForceSelection="true" AllowBlank="true" EmptyText="Contractor">
                                            <Listeners>
                                                <Select Handler="#{cbContractDescription}.clearValue();#{searchKeyword}.setValue(#{cbContractor}.getValue());searchClick();"></Select>
                                            </Listeners>
                                        </ext:ComboBox>
                                        <%--<ext:HyperLink runat="server" Text="Instructions" Icon="Help" Target="_blank" NavigateUrl="http://10.137.1.73/portal/DownloadService?docId=091ea3c180222572 " />--%>
                                        <ext:ToolbarFill></ext:ToolbarFill>
                                        <ext:Label runat="server" Html="<U>Black:Ongoing/Maroon:will be Expired in 3 months/Gray:Expired</U>"></ext:Label>
                                        <ext:Button ID="Button1" runat="server" Text="Export" Icon="PageExcel">
                                            <DirectEvents>
                                                <Click OnEvent="exportExcel">
                                                    <EventMask ShowMask="True" Msg="Exporting...."/>
                                                    <ExtraParams>
                                                        <ext:Parameter Name="NO" Value="App.searchNO.getValue()" Mode="Raw"/>
                                                        <ext:Parameter Name="CA" Value="App.searchCA.getValue()" Mode="Raw"/>
                                                        <ext:Parameter Name="PS" Value="App.SearchPringScheme.getValue()" Mode="Raw"/>
                                                         <ext:Parameter Name="searchKey" Value="App.searchKeyword.getValue()" Mode="Raw"/>
                                                        <ext:Parameter Name="isValid" Value="App.cbValid.getValue()" Mode="Raw"/>
                                                    </ExtraParams>                                            
                                                </Click>
                                            </DirectEvents>                                    
                                        </ext:Button>
                                        <ext:Button ID="btnSearch" runat="server" Text="Search" Icon="Find">
                                            <Listeners>
                                                <Click Handler="#{winSearch}.show();" />
                                            </Listeners>
                                        </ext:Button>
                                        <ext:Button ID="btnUploadMain" runat="server" Text="Upload" Icon="DiskUpload">
                                            <Listeners>
                                                <Click Handler="if(checkLogin()){App.fileContractNO.setValue(App.RowSelectionModelMain.getSelection()[0].data.NO);#{winFile}.show();#{storeWinFile}.reload();}" />
                                            </Listeners>
                                        </ext:Button>                                        
                                    </Items>
                                </ext:Toolbar>
                            </TopBar>                    
                        </ext:GridPanel> 
                     </Items>  
                    <TopBar>
                        <ext:Toolbar ID="Toolbar2" runat="server">
                            <Items>
                                <%-- <ext:HyperLink ID="HyperLink1" runat="server" Text="Instructions" Icon="Help" Target="_blank" NavigateUrl="http://10.137.1.73/portal/DownloadService?docId=091ea3c180222572" />--%>
                                 <ext:HyperLink ID="HyperLink1" runat="server" Text="Instructions" Icon="Help" Target="_blank" NavigateUrl="https://ewebtop.basf-ypc.com.cn/portal/DownloadService?docId=091ea3c180222572" />
                                 <ext:ToolbarFill></ext:ToolbarFill>
                                <ext:SplitButton ID="btnLogin" runat="server" Text="Login" Icon="UserGo">
                                            <Listeners>
                                                <Click Handler="if(!checkLogin('True')){#{winLogin}.show();}" />
                                            </Listeners>
                                            <Menu>
                                                <ext:Menu ID="MenuLogin" runat="server">
                                                    <Items>                    
                                                        <ext:MenuItem ID="miLogin" runat="server" Text="Login" Icon="UserGo" Handler="#{winLogin}.show();"/>
                                                        <ext:MenuItem ID="miReport" runat="server" Text="DetailReport" Icon="ReportGo" >
                                                            <%--<DirectEvents>
                                                                <Click OnEvent="exportExcel">
                                                                    <EventMask ShowMask="True" Msg="Exporting...."/>
                                                                </Click>
                                                            </DirectEvents>--%>
                                                            <Listeners>
                                                                <Click Handler="if(checkLogin()){#{winReport}.show();};"/>                                                         
                                                            </Listeners>                                                    
                                                        </ext:MenuItem>
                                                        <ext:MenuItem ID="miKPI" runat="server" Text="KPI" Icon="ReportAdd" >
                                                            <%--<DirectEvents>
                                                                <Click OnEvent="exportExcel">
                                                                    <EventMask ShowMask="True" Msg="Exporting...."/>
                                                                </Click>
                                                            </DirectEvents>--%>
                                                            <Listeners>
                                                                <Click Handler="if(checkLogin()){#{winKPI}.show();};"/>                                                         
                                                            </Listeners>                                                    
                                                        </ext:MenuItem>
                                                    </Items>
                                                </ext:Menu>
                                            </Menu>
                                        </ext:SplitButton>
                            </Items>
                        </ext:Toolbar>
                    </TopBar>                 
                </ext:Panel>   
                <ext:Container runat="server" Region="South" Layout="BorderLayout" Split="true" Height="400">
                    <Items>
                        <ext:GridPanel runat="server" ID="gridSub" Region="North" Split="true" Collapsible="true" Height="75" Title="Contractor Information">
                            <Store>
                                <ext:Store ID="StoreSub" runat="server"  PageSize="50"  OnSubmitData="storeWinFile_Submit" >                                    
                                    <Model>
                                        <ext:Model ID="Model4" runat="server">
                                            <Fields>                                                
                                                <ext:ModelField Name="workCenter"/>
                                                <ext:ModelField Name="contractor" />
                                                <ext:ModelField Name="contractAdmin"/>
                                                <ext:ModelField Name="buyer"/>
                                                <ext:ModelField Name="mainCoordinator"/>
                                                <ext:ModelField Name="userRepresentative"/>
                                                <ext:ModelField Name="contact"/>
                                                <ext:ModelField Name="telephone"/>
                                                <ext:ModelField Name="validateDate" Type="Date" />
                                                <ext:ModelField Name="expireDate" Type="Date"/>   
                                            </Fields>
                                        </ext:Model>
                                    </Model>                  
                                </ext:Store>
                            </Store>
                            <ColumnModel ID="ColumnModel2" runat="server">
                                <Columns>
                                    <ext:Column ID="Column2" runat="server" Text="Work Center" DataIndex="workCenter" Width="100" />
                                    <ext:Column ID="Column3" runat="server" Text="Contractor" Width="80" DataIndex="contractor" Flex="1">
                                    </ext:Column>
                                     <ext:Column ID="Column13" runat="server" Text="Contract Admin" Width="120" DataIndex="contractAdmin">
                                    </ext:Column>
                                    <ext:Column ID="Column14" runat="server" Text="Buyer" Width="120" DataIndex="buyer">
                                    </ext:Column> 
                                     <ext:Column ID="Column4" runat="server" Text="Main Coordinator" Width="100" DataIndex="mainCoordinator">
                                    </ext:Column> 
                                    <ext:Column ID="Column15" runat="server" Text="User Representative" Width="120" DataIndex="userRepresentative">
                                    </ext:Column>                                                                      
                                    <ext:Column ID="Column11" runat="server" Text="Contact" Width="160" DataIndex="contact">
                                    </ext:Column>
                                    <ext:Column ID="Column12" runat="server" Text="Tel." Width="180" DataIndex="telephone">
                                    </ext:Column>
                                    <ext:DateColumn  runat="server" Text="Validate Date" Width="120" DataIndex="validateDate" Format="yyyy-MM-dd">
                                    </ext:DateColumn>
                                    <ext:DateColumn  runat="server" Text="Expire Date" Width="120" DataIndex="expireDate" Format="yyyy-MM-dd">
                                    </ext:DateColumn>                                                                                              
                                </Columns>
                            </ColumnModel>
                        </ext:GridPanel> 
                        <ext:TabPanel runat="server" ID="tp1" Region="Center">
                            <Items>
                                <ext:Panel ID="Panel2" Title="General" runat="server" Layout="FitLayout" Icon="ChartPie">
                                    <Items>
                                        <ext:Container runat="server" Title="General">
                                            <LayoutConfig>
                                                <ext:HBoxLayoutConfig Align="Stretch" DefaultMargins="2" />
                                            </LayoutConfig>
                                            <Items>
                                                <ext:TabPanel runat="server" Flex="1" TabPosition="Right" CollapseDirection="Right">
                                                    <Items>
                                                        <ext:Panel runat="server" ID="KPI1"  Title="KPI1" Layout="FitLayout" >
                                                             <Items>
                                                                 <ext:Container runat="server">
                                                                    <LayoutConfig>
                                                                        <ext:VBoxLayoutConfig Align="Stretch" />
                                                                    </LayoutConfig>
                                                                 <Items>
                                                                     <ext:Panel runat="server" ID="pContractor" Html="<center></center>" Height="30" Border="false" />
                                                                     <ext:Chart ID="Chart1"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Theme="Base:gradients" Flex="1">
                                                                        <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                                                        <Store>
                                                                            <ext:Store ID="storePie1" runat="server" AutoLoad="false">                           
                                                                                <Proxy>
                                                                                    <ext:AjaxProxy Url="data.ashx?object=pie1">
                                                                                        <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                                                    </ext:AjaxProxy>
                                                                                </Proxy>
                                                                                <Model>
                                                                                    <ext:Model ID="Model2" runat="server">
                                                                                        <Fields>
                                                                                            <ext:ModelField Name="Name" />
                                                                                            <ext:ModelField Name="Deduction Cost(RMB)" Mapping="Data1" />
                                                                                            <ext:ModelField Name="Deduction Rate(%)" Mapping="Data2"/>
                                                                                            <ext:ModelField Name="MOM" />
                                                                                        </Fields>
                                                                                    </ext:Model>
                                                                                </Model>
                                                                            </ext:Store>
                                                                        </Store>
                                                                        <Axes>
                                                                            <ext:NumericAxis Fields="Deduction Cost(RMB)" Title="" Grid="true" Position="Left">
                                                                                <Label>
                                                                                    <Renderer Handler="return moneyFormat2(value);"/>
                                                                                </Label>
                                                                            </ext:NumericAxis>
                                                                            <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                                                                <Label>
                                                                                    <Rotate Degrees="336" />
                                                                                </Label>
                                                                            </ext:CategoryAxis>
                                                                            <ext:NumericAxis Fields="Deduction Rate(%)" Title="" Grid="false" Position="Right">                                                                
                                                                            </ext:NumericAxis>
                                                                        </Axes>
                                                                        <Series>
                                                                            <ext:ColumnSeries Axis="Left" XField="Name" YField="Deduction Cost(RMB)" Highlight="true">
                                                                                <Tips runat="server" TrackMouse="true" Width="50" Height="28">
                                                                                    <Renderer Handler="this.setTitle(String(moneyFormat2(item.value[1])));" />
                                                                                </Tips>
                                                                                <Label Display="Outside" Field="MOM"/>
                                                                                <Listeners>
                                                                                    <ItemDblClick Handler="#{hiddenDateMonth}.setValue(item.value[0]);Ext.Msg.show({'title':'MOM Mark','buttons':Ext.Msg.OKCANCEL,'fn':showResultText,'minWidth':250,'msg':'Please enter your remark:','prompt':true});" />
                                                                                </Listeners>                                                    
                                                                            </ext:ColumnSeries>
                                                                            <ext:LineSeries Axis="Right" Smooth="3" XField="Name" YField="Deduction Rate(%)" Highlight="true">
                                                                                <%--<Tips runat="server" TrackMouse="true" Width="42" Height="28">
                                                                                    <Renderer Handler="this.setTitle(String(item.value[1])+'%');" />
                                                                                </Tips>--%>                                                                
                                                                                <Style Fill="#FF0033" Stroke="#FF6666" StrokeWidth="2" />
                                                                                <Label Display="Over" FontSize="12" >
                                                                                    <Renderer Handler="return(String(item.value[1])+'%');"/>
                                                                                </Label>                                                            
                                                                            </ext:LineSeries>
                                                                        </Series>                                                         
                                                                    </ext:Chart>
                                                                    </Items>
                                                                </ext:Container>
                                                             </Items>
                                                            <TopBar>
                                                                <ext:Toolbar runat="server">
                                                                    <Items>
                                                                        <ext:ToolbarFill></ext:ToolbarFill>
                                                                        <ext:DateField runat="server"  FieldLabel="Time Range" LabelWidth="72" ID="dfKPI1Start" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Label runat="server" Text="-"/>
                                                                        <ext:DateField runat="server" ID="dfKPI1End" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Button runat="server" Text="Submit">
                                                                            <Listeners>
                                                                                <Click Handler="#{storePie1}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI1Start}.getValue(),DateEnd: #{dfKPI1End}.getValue()}});" />
                                                                            </Listeners>
                                                                        </ext:Button>
                                                                    </Items>
                                                                </ext:Toolbar>
                                                            </TopBar>
                                                         </ext:Panel>                                                       
                                                        <ext:Panel runat="server" ID="KPI2" Title="KPI2">
                                                            <LayoutConfig>
                                                                <ext:VBoxLayoutConfig Align="Stretch" />
                                                            </LayoutConfig>
                                                            <Items>
                                                                <ext:Panel runat="server" ID="Panel1" Html="<br/><center>Deduction Rate Distribution</center>" Height="30" Border="false" />
                                                                <ext:Chart ID="Chart5"  runat="server" Flex="1" Animate="true" Shadow="true" InsetPadding="20">
                                                                    <Store>
                                                                        <ext:Store ID="storeKPI2Scatter" runat="server" AutoLoad="False">                           
                                                                            <Proxy>
                                                                                <ext:AjaxProxy Url="data.ashx?object=scatterKPI2">
                                                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                                                </ext:AjaxProxy>
                                                                            </Proxy>
                                                                            <Model>
                                                                                <ext:Model ID="Model11" runat="server">
                                                                                    <Fields>
                                                                                        <ext:ModelField Name="Data1" />
                                                                                        <ext:ModelField Name="Data2"/>
                                                                                        <ext:ModelField Name="Name"/>
                                                                                    </Fields>
                                                                                </ext:Model>
                                                                            </Model>
                                                                        </ext:Store>
                                                                    </Store>
                                                                    <Axes>
                                                                        <ext:NumericAxis Fields="Data1" Title="Contractor Quotation(RMB)" Grid="true" Position="Bottom">
                                                                            <Label>
                                                                                <Renderer Handler="return moneyFormat2(value);"/>
                                                                            </Label>
                                                                        </ext:NumericAxis>
                                                                        <ext:NumericAxis Fields="Data2" Title="Deduction Rate(%)" Grid="true" Position="Left">  
                                                                            <Label>
                                                                                <Renderer Handler="return value+'%';"/>
                                                                            </Label>                                                              
                                                                        </ext:NumericAxis>
                                                                    </Axes>
                                                                    <Series>
                                                                        <ext:ScatterSeries XField="Data1" YField="Data2" Highlight="True">
                                                                             <Tips runat="server" TrackMouse="true" Width="120" Height="39">
                                                                                <Renderer Handler="this.setTitle(String(item.storeItem.data.Name)+':<br/>('+String(item.value[0])+','+String(item.value[1])+'%)');" />
                                                                            </Tips>
                                                                            <%--<Label Field="Data2" Display="Over">
                                                                                <Renderer Handler="return '('+String(item.value[0])+'%,'+String(item.value[1])+')';" />
                                                                            </Label> --%>
                                                                        </ext:ScatterSeries>
                                                                    </Series>       
                                                                                                                                                                                                 
                                                                </ext:Chart>
                                                            </Items>
                                                            <TopBar>
                                                                <ext:Toolbar runat="server">
                                                                    <Items>
                                                                        <ext:ToolbarFill></ext:ToolbarFill>
                                                                        <ext:DateField runat="server" ID="dfKPI2Start"  FieldLabel="Time Range" LabelWidth="72" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Label runat="server" Text="-"/>
                                                                        <ext:DateField runat="server" ID="dfKPI2End" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Button runat="server" Text="Submit">
                                                                            <Listeners>
                                                                                <Click Handler="#{storeKPI2Scatter}.removeAll();#{storeKPI2Scatter}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI2Start}.getValue(),DateEnd: #{dfKPI2End}.getValue()}});" />
                                                                            </Listeners>
                                                                        </ext:Button>
                                                                    </Items>
                                                                </ext:Toolbar>
                                                            </TopBar>
                                                        </ext:Panel>
                                                        <ext:Panel runat="server" ID="KPI3"  Title="KPI3" Layout="FitLayout" >
                                                             <Items>
                                                                 <ext:Container runat="server">
                                                                    <LayoutConfig>
                                                                        <ext:VBoxLayoutConfig Align="Stretch" />
                                                                    </LayoutConfig>
                                                                 <Items>
                                                                     <ext:Panel runat="server" ID="Panel4" Html="<center>All Users Deduction</center>" Height="30" Border="false" />
                                                                     <ext:Chart ID="Chart6"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Theme="Base:gradients" Flex="1">
                                                                        <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                                                        <Store>
                                                                            <ext:Store ID="storeKPI3" runat="server" AutoLoad="false">                           
                                                                                <Proxy>
                                                                                    <ext:AjaxProxy Url="data.ashx?object=pie3">
                                                                                        <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                                                    </ext:AjaxProxy>
                                                                                </Proxy>
                                                                                <Model>
                                                                                    <ext:Model ID="Model12" runat="server">
                                                                                        <Fields>
                                                                                            <ext:ModelField Name="Name" />
                                                                                            <ext:ModelField Name="Cost(RMB)" Mapping="Data1" />
                                                                                            <ext:ModelField Name="Deduction Rate(%)" Mapping="Data2"/>
                                                                                            <ext:ModelField Name="MOM" />
                                                                                        </Fields>
                                                                                    </ext:Model>
                                                                                </Model>
                                                                            </ext:Store>
                                                                        </Store>
                                                                        <Axes>
                                                                            <ext:NumericAxis Fields="Cost(RMB)" Title="" Grid="true" Position="Left">
                                                                                <Label>
                                                                                    <Renderer Handler="return moneyFormat2(value);"/>
                                                                                </Label>
                                                                            </ext:NumericAxis>
                                                                            <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                                                                <Label>
                                                                                    <Rotate Degrees="270" />
                                                                                </Label>
                                                                            </ext:CategoryAxis>
                                                                            <ext:NumericAxis Fields="Deduction Rate(%)" Title="" Grid="false" Position="Right">                                                                
                                                                            </ext:NumericAxis>
                                                                        </Axes>
                                                                        <Series>
                                                                            <ext:ColumnSeries Axis="Left" XField="Name" YField="Cost(RMB)" Highlight="true">
                                                                                <Tips runat="server" TrackMouse="true" Width="80" Height="28">
                                                                                    <Renderer Handler="this.setTitle(String(moneyFormat2(item.value[1])));" />
                                                                                </Tips>                                                                                                                                   
                                                                            </ext:ColumnSeries>
                                                                            <ext:LineSeries Axis="Right" Smooth="3" XField="Name" YField="Deduction Rate(%)" Highlight="true">
                                                                                <%--<Tips runat="server" TrackMouse="true" Width="42" Height="28">
                                                                                    <Renderer Handler="this.setTitle(String(item.value[1])+'%');" />
                                                                                </Tips>--%>                                                                
                                                                                <Style Fill="#FF0033" Stroke="#FF6666" StrokeWidth="2" />
                                                                                <Label Display="Over" FontSize="12" >
                                                                                    <Renderer Handler="return(String(item.value[1])+'%');"/>
                                                                                </Label>                                                            
                                                                            </ext:LineSeries>
                                                                        </Series>                                                         
                                                                    </ext:Chart>
                                                                    </Items>
                                                                </ext:Container>
                                                             </Items>
                                                            <TopBar>
                                                                <ext:Toolbar runat="server">
                                                                    <Items>
                                                                        <ext:ToolbarFill></ext:ToolbarFill>
                                                                        <ext:DateField runat="server"  FieldLabel="Time Range" LabelWidth="72" ID="dfKPI3Start" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Label runat="server" Text="-"/>
                                                                        <ext:DateField runat="server" ID="dfKPI3End" Format="yyyy-MM"></ext:DateField>
                                                                        <ext:Button runat="server" Text="Submit">
                                                                            <Listeners>
                                                                                <Click Handler="#{storeKPI3}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI3Start}.getValue(),DateEnd: #{dfKPI3End}.getValue()}});" />
                                                                            </Listeners>
                                                                        </ext:Button>
                                                                    </Items>
                                                                </ext:Toolbar>
                                                            </TopBar>
                                                         </ext:Panel>
                                                    </Items>
                                                    <Listeners>
                                                        <TabChange Handler="if(newTab.getId()=='KPI2'){#{storeKPI2Scatter}.removeAll();#{storeKPI2Scatter}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI2Start}.getValue(),DateEnd: #{dfKPI2End}.getValue()}});}else if(newTab.getId()=='KPI1'){#{storePie1}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI1Start}.getValue(),DateEnd: #{dfKPI1End}.getValue()}});}else if(newTab.getId()=='KPI3'){#{storeKPI3}.reload({params: { NO:#{hiddenSelectedNO}.getValue(),DateStart: #{dfKPI3Start}.getValue(),DateEnd: #{dfKPI3End}.getValue()}});}"/>
                                                    </Listeners>
                                                </ext:TabPanel> 
                                                <ext:BoxSplitter runat="server" Collapsible="true" />                                            
                                                 <ext:Panel runat="server" Flex="1" Title="" Layout="FitLayout" CollapseDirection="Left">
                                                     <TopBar>
                                                         <ext:Toolbar runat="server">
                                                             <Items>
                                                                 <ext:ToolbarFill runat="server"/>
                                                                 <ext:DateField ID="DateField1" runat="server" Vtype="daterange" EndDateField="DateField2" EnableKeyEvents="true" Width="90" Format="yyyy-MM-dd">
                                                                     <Listeners>
                                                                          <KeyUp Fn="onKeyUp" />
                                                                     </Listeners>
                                                                 </ext:DateField>
                                                                 <ext:Label runat="server" Text="-"/>
                                                                 <ext:DateField ID="DateField2" runat="server" Vtype="daterange" StartDateField="DateField1"  EnableKeyEvents="true" Width="90" Format="yyyy-MM-dd">
                                                                     <Listeners>
                                                                          <KeyUp Fn="onKeyUp" />
                                                                     </Listeners>
                                                                 </ext:DateField>
                                                                <ext:SplitButton runat="server" ID="btnTimeRangeSubmit" Text="Submit TimeRange" Icon="ClockGo" ArrowAlign="Right" Handler="submitTimeRange">
                                                                    <Menu>
                                                                        <ext:Menu ID="Menu1" runat="server">
                                                                            <Items>
                                                                                <ext:MenuItem ID="MenuItem1" runat="server" Text="Submit Last 3 Months" Icon="ClockAdd" Handler="submitTimeRange"/>
                                                                                <ext:MenuItem ID="MenuItem2" runat="server" Text="Submit Last 6 Months" Icon="ClockAdd" Handler="submitTimeRange"/>
                                                                                <ext:MenuItem ID="MenuItem3" runat="server" Text="Submit Last 1 Year" Icon="ClockAdd" Handler="submitTimeRange"/>
                                                                            </Items>
                                                                        </ext:Menu>
                                                                    </Menu>
                                                                </ext:SplitButton>
                                                                 <ext:ToolbarSpacer ID="ToolbarSpacer1" runat="server" />
                                                                 <ext:ToolbarSeparator ID="ToolbarSeparator1" runat="server" />
                                                                 <ext:ToolbarSpacer ID="ToolbarSpacer2" runat="server" />
                                                                 <ext:SplitButton ID="btnEvaluationExport" runat="server" Text="Export Details" Icon="PageGo" ArrowAlign="Right" Handler="exportClick">
                                                                    <Menu>
                                                                        <ext:Menu ID="Menu2" runat="server">
                                                                            <Items>
                                                                                <ext:MenuItem ID="MenuItem4" runat="server" Text="Per Month" Icon="PageGo" Handler="exportClick"/>
                                                                                <ext:MenuItem ID="MenuItem5" runat="server" Text="Per Quarter" Icon="PageGo" Handler="exportClick"/>
                                                                                <ext:MenuItem ID="MenuItem6" runat="server" Text="Per Year" Icon="PageGo" Handler="exportClick"/>
                                                                                <ext:MenuItem ID="MenuItem7" runat="server" Text="Overall Statistics" Icon="PageGo" Handler="exportClick"/>
                                                                                <ext:MenuItem ID="MenuItem8" runat="server" Text="Month Statistics" Icon="PageGo" Handler="exportClick"/>
                                                                            </Items>
                                                                        </ext:Menu>
                                                                    </Menu>
                                                                 </ext:SplitButton>
                                                                 <ext:Button runat="server" Text="Save Chart" Icon="Disk" Handler="saveChart"/>
                                                             </Items>
                                                         </ext:Toolbar>
                                                     </TopBar>
                                                    <Items>
                                                        <ext:Chart ID="Chart2" runat="server" InsetPadding="12" Animate="True">
                                                            <Store>
                                                                <ext:Store ID="storePie2" runat="server" AutoLoad="false">
                                                                    <Proxy>
                                                                        <ext:AjaxProxy Url="data.ashx?object=pie2">
                                                                            <ActionMethods Read="GET" />
                                                                        </ext:AjaxProxy>
                                                                    </Proxy>
                                                                    <Parameters>
                                                                        <ext:StoreParameter Name="Months" Value="#{DateField2}.getValue().dateDiff('m',#{DateField1}.getValue())" Mode="Raw"/>
                                                                        <ext:StoreParameter Name="EndDate" Value="#{DateField2}.getValue().format('yyyy-MM-dd')" Mode="Raw"/>
                                                                    </Parameters>
                                                                    <Model>
                                                                        <ext:Model ID="Model3" runat="server">
                                                                            <Fields>
                                                                                <ext:ModelField Name="Month" />
                                                                                <ext:ModelField Name="Contract_NO" />
                                                                                <ext:ModelField Name="User" Mapping="userScore" />
                                                                                <ext:ModelField Name="Dep." Mapping="depScore" />
                                                                                <ext:ModelField Name="allScore" />
                                                                            </Fields>
                                                                        </ext:Model>
                                                                    </Model>
                                                                </ext:Store>
                                                            </Store>
                                                            <LegendConfig Position="Right" />

                                                            <Axes>
                                                                <ext:NumericAxis                             
                                                                    Fields="User,Dep."
                                                                    Position="Left"
                                                                    Title="Score"
                                                                    Grid="true" Minimum="0" Maximum="100">
                                                                </ext:NumericAxis>                            

                                                                <ext:CategoryAxis Fields="Month" Position="Bottom" />                            
                                                            </Axes>

                                                            <Series>
                                                                <ext:BarSeries 
                                                                    Axis="Left"
                                                                    Gutter="80"
                                                                    XField="Month" 
                                                                    YField="User,Dep."
                                                                    Stacked="true" Column="True" YPadding="0" XPadding="10">
                                                                    <Tips runat="server" TrackMouse="true" Width="36" Height="28">
                                                                        <Renderer Handler="this.setTitle(String(item.value[1]));" />
                                                                    </Tips>
                                                                </ext:BarSeries>
                                                                 <ext:LineSeries Axis="Left" Smooth="3" Fill="False" XField="Month" YField="allScore" ShowInLegend="False" />
                                                            </Series>
                                                        </ext:Chart>
                                                    </Items>
                                                 </ext:Panel>
                                            </Items>
                                        </ext:Container> 
                                    </Items>
                                </ext:Panel>
                                <ext:Panel ID="pEvaluation" Title="Evaluation" runat="server" Layout="FitLayout" Icon="ScriptGo">
                                    <Items>
                                        <ext:Container ID="subTabEvaluation" runat="server" Title="Evaluation">
                                            <LayoutConfig>
                                                <ext:HBoxLayoutConfig Align="Stretch" DefaultMargins="2" />
                                            </LayoutConfig>
                                            <Items>
                                                <ext:Panel ID="PanelEvaluationData" runat="server"  Flex="3" Layout="FitLayout" >
                                                    <Items>
                                                        <ext:GridPanel runat="server" ID="GridPanel1" >
                                                            <Store>
                                                                <ext:Store ID="storeEvaluation" runat="server" AutoLoad="false"> 
                                                                    <Proxy>
                                                                        <ext:AjaxProxy Url="data.ashx?object=evaluation">
                                                                            <ActionMethods Read="GET" />                                                                                                                                                                                                                                                                   
                                                                        </ext:AjaxProxy>                                                                
                                                                    </Proxy>
                                                                    <Parameters>
                                                                        <ext:StoreParameter Name="User" Value="#{hiddenUser}.value" Mode="Raw" />
                                                                        <ext:StoreParameter Name="isUser" Value="#{rdUser}.getValue()" Mode="Raw" />
                                                                    </Parameters>
                                                                    <Model>
                                                                        <ext:Model ID="Model5" runat="server">
                                                                            <Fields>                                                
                                                                                <ext:ModelField Name="SES_No"/>
                                                                                <ext:ModelField Name="Short_Descrption" />
                                                                                <ext:ModelField Name="Start_Date"/>
                                                                                <ext:ModelField Name="End_Date"/>
                                                                                <ext:ModelField Name="TECO_Format" Type="Date"/>
                                                                                <%--<ext:ModelField Name="CS_REC_Format" Type="Date"/>--%>
                                                                                <ext:ModelField Name="SES_CONF_Format" Type="Date"/>
                                                                                <ext:ModelField Name="Requisitioner"/>                                                                        
                                                                            </Fields>
                                                                        </ext:Model>
                                                                    </Model>                  
                                                                </ext:Store>
                                                            </Store>
                                                            <ColumnModel ID="ColumnModel3" runat="server">
                                                                <Columns>
                                                                    <ext:Column ID="Column17" runat="server" Text="SES NO." DataIndex="SES_No" Width="80" />
                                                                    <ext:Column ID="Column18" runat="server" Text="Description" Width="80" DataIndex="Short_Descrption" Flex="1">
                                                                    </ext:Column>
                                                                     <ext:Column ID="Column19" runat="server" Text="Start Date" Width="80" DataIndex="Start_Date">
                                                                    </ext:Column>
                                                                    <ext:Column ID="Column20" runat="server" Text="End Date" Width="80" DataIndex="End_Date">
                                                                    </ext:Column> 
                                                                    <ext:DateColumn ID="DateColumn1"  runat="server" Text="TECO Date" Width="80" DataIndex="TECO_Format" Format="yyyy-MM-dd">
                                                                    </ext:DateColumn>
                                                                    <ext:DateColumn ID="DateColumn2"  runat="server" Text="SES Confirm Date" Width="100" DataIndex="SES_CONF_Format" Format="yyyy-MM-dd">
                                                                    </ext:DateColumn>                                                                       
                                                                    <ext:Column ID="Column23" runat="server" Text="Requisitioner" Width="80" DataIndex="Requisitioner">
                                                                    </ext:Column>
                                                                                                                                                                                                                
                                                                </Columns>
                                                            </ColumnModel>
                                                        </ext:GridPanel>  
                                                    </Items>
                                                </ext:Panel>
                                                <ext:Panel ID="panelEvaluation" runat="server"  Flex="4" Layout="FitLayout" >
                                                    <Items>
                                                        <ext:GridPanel runat="server" ID="gridEvaluationContract" Layout="FitLayout">
                                                            <Store>
                                                                <ext:Store ID="storeEvaluationContract" runat="server" AutoLoad="false"> 
                                                                    <Proxy>
                                                                        <ext:AjaxProxy Url="data.ashx?object=evaluationContract">
                                                                            <ActionMethods Read="GET" />                                                                                                                                                                                                                                                                   
                                                                        </ext:AjaxProxy>                                                                                                                                
                                                                    </Proxy>
                                                                    <Parameters>
                                                                        <ext:StoreParameter Name="User" Value="#{hiddenUser}.value" Mode="Raw" />
                                                                        <ext:StoreParameter Name="isUser" Value="#{rdUser}.getValue()" Mode="Raw" />
                                                                    </Parameters>
                                                                    <Model>
                                                                        <ext:Model ID="Model6" runat="server">
                                                                            <Fields>
                                                                                <ext:ModelField Name="User"/>                                                
                                                                                <ext:ModelField Name="Contract_No"/>
                                                                                <ext:ModelField Name="Score1" Type="float"/>
                                                                                <ext:ModelField Name="Score2" Type="float"/>
                                                                                <ext:ModelField Name="Score3" Type="float"/>
                                                                                <ext:ModelField Name="Score4" Type="float"/>
                                                                                <ext:ModelField Name="Score5" Type="float"/>
                                                                                <ext:ModelField Name="Score6" Type="float"/>
                                                                                <ext:ModelField Name="Score7" Type="float"/>
                                                                                <ext:ModelField Name="Evaluated" Type="Boolean"/> 
                                                                                <ext:ModelField Name="FileID" />                                                                      
                                                                            </Fields>
                                                                        </ext:Model>
                                                                    </Model>                                                  
                                                                </ext:Store>
                                                            </Store>
                                                            <ColumnModel ID="ColumnModel4" runat="server">
                                                                <Columns>
                                                                    <ext:Column ID="Column21" runat="server" Text="合同号<br/>Contract NO." DataIndex="Contract_No" Width="80" Flex="1" Hideable="False" Align="Center">
                                                                        <Renderer Fn="evaluateRender"/>
                                                                    </ext:Column>
                                                                    <ext:RatingColumn ID="RatingColumn1" runat="server" Text="工作准备<br/>Preparation" DataIndex="Score1" RoundToTick="True" Editable="true" Hideable="False" /> 
                                                                    <ext:RatingColumn ID="RatingColumn2" runat="server" Text="工作表现<br/>Performance" DataIndex="Score2" RoundToTick="True" Editable="true" Hideable="False" /> 
                                                                    <ext:RatingColumn ID="RatingColumn3" runat="server" Text="EHSS管理<br/>EHSS" DataIndex="Score3" RoundToTick="True" Editable="true"  Hideable="False" /> 
                                                                    <ext:RatingColumn ID="RatingColumn4" runat="server" Text="质量控制<br/>Quality Control" DataIndex="Score4" RoundToTick="True" Editable="true"  Hideable="False" /> 
                                                                    <ext:RatingColumn ID="RatingColumn5" runat="server" Text="时间管理<br/>Timeline Management" DataIndex="Score5" RoundToTick="True" Editable="true" Hideable="False" /> 
                                                                    <ext:RatingColumn ID="RatingColumn6" runat="server" Text="文档管理<br/>Documentation" DataIndex="Score6" RoundToTick="True" Editable="true" Hideable="False" /> 
                                                                    <%--<ext:RatingColumn ID="RatingColumn7" runat="server" Text="诚实度<br/>Honesty" DataIndex="Score7" RoundToTick="True" Editable="false" />--%> 
                                                                    <ext:Column ID="clmDepScore" runat="server" Text="职能部门评分<br/>Function Departent Score" Width="200" Hidden="true" Hideable="False" >
                                                                        <Columns>
                                                                            <ext:Column runat="server" Text="CTS" DataIndex="Score1" Width="100" Hidden="true" Hideable="False" Align="Center" Sortable="True" >
                                                                                <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                            <ext:Column runat="server" Text="CHA" DataIndex="Score2" Width="100" Hidden="true" Hideable="False" Align="Center"  Sortable="True">
                                                                                <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                            <ext:Column runat="server" Text="Main Coordinator" DataIndex="Score3" Width="150" Hidden="true" Hideable="False" Align="Center"  Sortable="True">
                                                                                 <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                            <ext:Column runat="server" Text="User Representative" DataIndex="Score4" Width="150" Hidden="true" Hideable="False" Align="Center"  Sortable="True">
                                                                                 <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                            <ext:Column runat="server" Text="CTM/T" DataIndex="Score5" Width="100" Hidden="true" Hideable="False" Align="Center"  Sortable="True">
                                                                                 <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                            <ext:Column runat="server" Text="CTE/D" DataIndex="Score6" Width="100" Hidden="true" Hideable="False" Align="Center"  Sortable="True">
                                                                                 <Renderer Handler="if(value<0){return '';}else{return '<B>'+value+'</B>';}"/>
                                                                            </ext:Column>
                                                                        </Columns>
                                                                    </ext:Column>
                                                                                                                                                                       
                                                                    <ext:CommandColumn runat="server" ID="clmEvaluate" Text="评估<br/>Evaluate" Width="80" Hideable="False"><%--DataIndex="Evaluated"--%>
                                                                        <Commands>
                                                                            <ext:GridCommand Icon="ScriptGo" CommandName="Evaluate" Text="Evaluate" StandOut="True">
                                                                                <%--<ToolTip Text="点击提交测评<br/>Click to Evaluate" />  --%>                                                                                                                                              
                                                                            </ext:GridCommand>                                                                                                                                                                                                                                                                                
                                                                        </Commands>                                                                                                                                                                                                                                                             
                                                                        <Listeners>
                                                                            <Command Handler="if(record.data.Evaluated){showAlert('Error','You have already evaluated '+record.data.Contract_No+'!');record.reject();return false;}" />
                                                                        </Listeners>
                                                                        <DirectEvents>
                                                                            <Command OnEvent="Evaluation_Click">
                                                                                <ExtraParams>
                                                                                    <ext:Parameter Name="Contract_No" Value="record.data.Contract_No" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score1" Value="record.data.Score1*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score2" Value="record.data.Score2*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score3" Value="record.data.Score3*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score4" Value="record.data.Score4*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score5" Value="record.data.Score5*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="Score6" Value="record.data.Score6*2.0" Mode="Raw"/>
                                                                                    <ext:Parameter Name="User" Value="#{hiddenUser}.value" Mode="Raw"/>
                                                                                </ExtraParams>
                                                                            </Command>
                                                                        </DirectEvents>                                                                                                                          
                                                                    </ext:CommandColumn> 
                                                            
                                                                    <ext:CommandColumn runat="server" Text="文件上传<br/>Upload Files" Width="80" Hideable="False">
                                                                        <Commands>
                                                                            <ext:GridCommand Icon="DiskUpload" CommandName="Upload" Text="Upload" StandOut="True">
                                                                                <%--<ToolTip Text="上传文件(如果存在)<br/>Upload files if any"/>--%>
                                                                            </ext:GridCommand>
                                                                        </Commands>
                                                                        <Listeners>
                                                                            <Command Handler="#{fileContractNO}.setValue(record.data.Contract_No);#{winFile}.show();#{storeWinFile}.reload();"></Command>
                                                                        </Listeners>
                                                                    </ext:CommandColumn>                                                          
                                                                                                                                                                                            
                                                                </Columns>                                                                
                                                            </ColumnModel>
                                                            <Listeners>
                                                                <%--<CellClick Handler="if(record.data.Evaluated){record.reject();}" />--%>
                                                            </Listeners>
                                                            <SelectionModel>
                                                                <ext:RowSelectionModel runat="server" Mode="Single">
                                                                    <Listeners>
                                                                        <Select Handler="searchAutomatic(record.data.Contract_No);"></Select>
                                                                    </Listeners>
                                                                </ext:RowSelectionModel>
                                                            </SelectionModel>
                                                            <TopBar>
                                                                <ext:Toolbar runat="server">
                                                                    <Items>
                                                                        <ext:ToolbarFill runat="server"/>
                                                                        <ext:Button ID="btnRefresh2" runat="server" Icon="ArrowRefresh" Text="Refresh" Handler="#{storeMain}.reload();"/>                                                                            
                                                                        <ext:ToolbarSeparator />
                                                                        <ext:Button ID="btnDownloadTemplate" runat="server" Icon="PageWhiteExcel" Text="Download Template">
                                                                            <Listeners>
                                                                                <Click Handler="if(checkLogin()){#{gridMain}.submitData(false, {isUpload:true})};"/>
                                                                            </Listeners>
                                                                        </ext:Button>
                                                                        <ext:Button ID="btnBatchImport" runat="server" Icon="PageCopy" Text="Batch Import">
                                                                            <Listeners>
                                                                                <Click Handler="if(checkLogin()){#{winUpload}.show();};"/>
                                                                            </Listeners>
                                                                        </ext:Button>
                                                                    </Items>
                                                                </ext:Toolbar>
                                                            </TopBar>
                                                    
                                                        </ext:GridPanel>
                                                    </Items>    
                                                </ext:Panel>
                                            </Items>
                                        </ext:Container>
                                    </Items>
                                </ext:Panel>
                                <ext:Panel ID="pFiles" Title="Files Management" runat="server" Layout="FitLayout" Icon="PageWhiteStack">
                                    <Items>
                                        <ext:GridPanel runat="server" ID="GridPanelFile2" Height="220">
                                            <Store>
                                                <ext:Store ID="storeFiles" runat="server" OnSubmitData="storeWinFile_Submit" AutoLoad="false"> 
                                                    <Proxy>
                                                        <ext:AjaxProxy Url="data.ashx?object=getfiles">
                                                            <ActionMethods Read="GET" />                                                                                                                                                                                                                                                                   
                                                        </ext:AjaxProxy>                                                                
                                                    </Proxy>
                                                    <Model>
                                                        <ext:Model ID="Model8" runat="server">
                                                            <Fields>                                                
                                                                <ext:ModelField Name="FileID"/>
                                                                <ext:ModelField Name="FileName" />
                                                                <ext:ModelField Name="FileLength" Type="Int"/>
                                                                <ext:ModelField Name="FileLengthStr"/>
                                                                <ext:ModelField Name="UploadUser"/>
                                                                <ext:ModelField Name="DateIn" Type="Date"/>
                                                                <ext:ModelField Name="Remark"/> 
                                                                <ext:ModelField Name="Type"/>                                                                        
                                                            </Fields>
                                                        </ext:Model>
                                                    </Model>                  
                                                </ext:Store>
                                            </Store>
                                            <ColumnModel ID="ColumnModel6" runat="server">
                                                <Columns>
                                                    <ext:Column ID="Column25" runat="server" Text="FileID" DataIndex="FileID" Width="80" Hidden="True" />
                                                    <ext:Column ID="Column26" runat="server" Text="File Name" Width="80" DataIndex="FileName" Flex="1" />
                                                    <ext:Column ID="Column32" runat="server" Text="File Type" Width="80" DataIndex="Type" Flex="1" />
                                                    <ext:Column ID="Column29"  runat="server" Text="File Length" Width="80" DataIndex="FileLengthStr" Flex="1" />
                                                    <ext:Column ID="Column30"  runat="server" Text="Upload User" Width="80" DataIndex="UploadUser" Flex="1" />
                                                    <ext:DateColumn ID="DateColumn3"  runat="server" Text="Date" Width="80" DataIndex="DateIn" Format="yyyy-MM-dd" Flex="1">
                                                    </ext:DateColumn> 
                                                    <ext:Column ID="Column31" runat="server" Text="Remark" Width="80" DataIndex="Remark" Flex="1" />                                                                                                                                                                                                                
                                                </Columns>                        
                                            </ColumnModel> 
                                            <View>
                                                <ext:GridView ID="GridView3" runat="server" StripeRows="true" TrackOver="true" >
                                                    <Listeners>
                                                        <CellDblClick Handler="#{hiddenFileID}.setValue(record.data.FileID+'[@]'+record.data.FileName);#{GridPanelFile2}.submitData(false, {isUpload:true});" />
                                                    </Listeners>                             
                                                </ext:GridView>
                                            </View>                 
                                        </ext:GridPanel>

                                    </Items>
                                </ext:Panel>      
                            </Items>
                            <BottomBar>
                                <ext:StatusBar ID="statusBar1" runat="server">
                                    <Items>
                                        <ext:Label runat="server" Html='Content:<a  href="mailto:suhy@basf-ypc.com.cn">Su Hongying</a>&nbsp;&nbsp;Administration:<a href="mailto:Gang.Ji@basf-ypc.com.cn">Ji Gang</a>&nbsp;&nbsp;Copyright © 2013 BASF-YPC Company Limited'></ext:Label>
                                        <ext:ToolbarFill></ext:ToolbarFill>
                                    </Items>
                                </ext:StatusBar>
                            </BottomBar>
                            <Listeners>
                                <TabChange Handler="if(newTab.getId()=='pEvaluation'){loginTypeChange();#{storeEvaluationContract}.reload();#{storeEvaluation}.reload();}else if(newTab.getId()=='pFiles'){#{storeFiles}.reload();}"/>
                            </Listeners>
                        </ext:TabPanel>                                             
                    </Items>                                       
                </ext:Container>
            </Items>            
        </ext:Viewport>
        <ext:Window ID="winSearch" runat="server" Title="Search"  Icon="Find"  Width="350" BodyPaddingSummary="5 5 5" Modal="false" Layout="FormLayout" Hidden="True" Resizable="False" AnimateTarget="btnSearch">
            <Items>
                <ext:TextField ID="searchKeyword" runat="server" FieldLabel="Keyword" >                    
                </ext:TextField> 
                <ext:TextField ID="searchNO" runat="server" FieldLabel="Contract NO." >                    
                </ext:TextField> 
                <ext:ComboBox ID="searchCA"  runat="server" FieldLabel="Contract Admin." ForceSelection="True" AllowBlank="True">
                </ext:ComboBox>  
                <ext:ComboBox ID="SearchPringScheme"  runat="server" FieldLabel="Pricing Scheme" ForceSelection="True" AllowBlank="True">
                </ext:ComboBox>               
            </Items>
            <Buttons>
                <ext:Button ID="btnSearchOK" runat="server" Icon="Find" Text="Search" Handler="searchClick();" />
                <ext:Button ID="btnSeachCancel" runat="server" Icon="PageCancel" Text="Reset" Handler="clearSearch();" />                
            </Buttons>   
            <KeyMap runat="server">
                <Binding>
                    <ext:KeyBinding Handler="searchClick();">
                        <Keys>
                            <ext:Key Code="ENTER"></ext:Key>
                        </Keys>
                    </ext:KeyBinding>
                </Binding>
            </KeyMap>
            <Listeners>
                <Show Handler="#{searchKeyword}.focus();#{searchNO}.focus();#{searchCA}.focus();#{SearchPringScheme}.focus();"/><%--for IE8--%>
            </Listeners>         
        </ext:Window>

        <ext:Window ID="winLogin" runat="server" Resizable="false"  Height="150" Width="320" Icon="Lock"  Title="Login" BodyPadding="5" Layout="Form" Hidden="True" Modal="False" AnimateTarget="btnLogin">
            <Items>
                <ext:TextField ID="txtUsername"  runat="server"  FieldLabel="Username" AllowBlank="false"  BlankText="Your username is required." />
                <ext:TextField ID="txtPassword"  runat="server"  InputType="Password"  FieldLabel="Password"  AllowBlank="false" BlankText="Your password is required." />
                <ext:RadioGroup ID="RadioGroup1" runat="server" FieldLabel="UserType" AutomaticGrouping="false">                            
                    <Items>
                        <ext:Radio ID="rdUser" runat="server" Name="rating" InputValue="0" BoxLabel="User" Checked="True" /> 
                        <ext:Radio ID="rdDep" runat="server" Name="rating" InputValue="1" BoxLabel="Dep." />
                    </Items>
                    <Listeners>
                        <Change Handler="loginTypeChange();" />
                    </Listeners>
                </ext:RadioGroup>
            </Items>
            <Buttons>
                <ext:Button ID="btnLoginOK" runat="server" Text="Login" Icon="Accept">
                     <Listeners>
                        <Click Handler="
                            if (!#{txtUsername}.validate() || !#{txtPassword}.validate()) {
                                Ext.Msg.alert('Error','The Username and Password fields are both required');
                                // return false to prevent the btnLogin_Click Ajax Click event from firing.
                                return false; 
                            }" />
                    </Listeners>
                    <DirectEvents>
                        <Click OnEvent="btnLogin_Click" Success="#{winLogin}.close();">
                            <EventMask ShowMask="true" Msg="Verifying..." MinDelay="1000" />
                        </Click>
                    </DirectEvents>
                </ext:Button>
                <ext:Button ID="btnCancel" runat="server" Text="Cancel" Icon="Decline">
                    <Listeners>
                        <Click Handler="#{winLogin}.hide();" />
                    </Listeners>
                </ext:Button>
            </Buttons>
            <KeyMap ID="KeyMap1" runat="server">
                <Binding>
                    <ext:KeyBinding Handler="#{btnLoginOK}.fireEvent('click');">
                        <Keys>
                            <ext:Key Code="ENTER"></ext:Key>
                        </Keys>
                    </ext:KeyBinding>
                </Binding>
            </KeyMap> 
            <Listeners>
                <Show Handler="#{storeEvaluationContract}.removeAll();#{txtUsername}.focus();#{txtPassword}.focus();loginTypeChange();#{rdDep}.focus();#{rdUser}.focus();"/><%--for IE8--%>
            </Listeners>        
        </ext:Window>

        <ext:Window ID="winFile" runat="server" Width="360" Height="360" Hidden="True" Title="File Management" BodyPaddingSummary="5 5 5" Resizable="false" Icon="DiskUpload" Layout="VBoxLayout" Frame="True">
            <LayoutConfig>
                <ext:VBoxLayoutConfig Align="Stretch" />
            </LayoutConfig>
            <Items>
                <ext:FormPanel ID="BasicForm"  runat="server"  Frame="true">                
                    <Defaults>
                        <ext:Parameter Name="anchor" Value="95%" Mode="Value" />
                        <ext:Parameter Name="allowBlank" Value="false" Mode="Raw" />
                        <ext:Parameter Name="msgTarget" Value="side" Mode="Value" />
                    </Defaults>
                    <Items>
                        <ext:FileUploadField  ID="FileUploadField1" runat="server" EmptyText="Select a file" FieldLabel="File"  ButtonText=""  Icon="ArrowUp" LabelWidth="70" >                            
                        </ext:FileUploadField>
                        <ext:ComboBox ID="cbFileType"  runat="server" FieldLabel="Type" ForceSelection="True" AllowBlank="True" LabelWidth="70">
                            <Items>
                                <ext:ListItem Text="MOM"/>
                                <ext:ListItem Text="图片"/>
                                <ext:ListItem Text="违章记录/扣款记录"/>
                                <ext:ListItem Text="投诉信"/>
                                <ext:ListItem Text="NCR"/>
                                <ext:ListItem Text="事故分析报告"/>
                                <ext:ListItem Text="年审报告"/>
                                <ext:ListItem Text="其他"/>
                            </Items>
                        </ext:ComboBox>   
                        <ext:TextField ID="fileRemark" runat="server" FieldLabel="Remark" LabelWidth="70"/>
                    </Items>
                    <Listeners>
                        <ValidityChange Handler="#{SaveButton}.setDisabled(!valid);" />
                    </Listeners>
                    <Buttons>
                        <ext:Button ID="SaveButton" runat="server" Text="Save" Disabled="true">
                            <DirectEvents>
                                <Click 
                                    OnEvent="UploadClick"
                                    Before="if (!#{BasicForm}.getForm().isValid()) { return false; } 
                                        Ext.Msg.wait('Uploading your file...', 'Uploading');"
                                
                                    Failure="Ext.Msg.show({ 
                                        title   : 'Error', 
                                        msg     : 'Error during uploading', 
                                        minWidth: 200, 
                                        modal   : true, 
                                        icon    : Ext.Msg.ERROR, 
                                        buttons : Ext.Msg.OK 
                                    });">
                                </Click>
                            </DirectEvents>
                        </ext:Button>
                        <ext:Button ID="Button2" runat="server" Text="Reset">
                            <Listeners>
                                <Click Handler="#{BasicForm}.getForm().reset();" />
                            </Listeners>
                        </ext:Button>
                    </Buttons>
                </ext:FormPanel>
                <ext:GridPanel runat="server" ID="gridWinFile" Height="220">
                    <Store>
                        <ext:Store ID="storeWinFile" runat="server" OnSubmitData="storeWinFile_Submit"> 
                            <Proxy>
                                <ext:AjaxProxy Url="data.ashx?object=getfiles">
                                    <ActionMethods Read="GET" />                                                                                                                                                                                                                                                                   
                                </ext:AjaxProxy>                                                                
                            </Proxy>
                            <Parameters>
                                <ext:StoreParameter Name="NO" Value="#{fileContractNO}.getValue()" Mode="Raw"/>
                            </Parameters>
                            <Model>
                                <ext:Model ID="Model7" runat="server">
                                    <Fields>                                                
                                        <ext:ModelField Name="FileID"/>
                                        <ext:ModelField Name="FileName" />
                                        <ext:ModelField Name="FileLength" Type="Int"/>
                                        <ext:ModelField Name="UploadUser"/>
                                        <ext:ModelField Name="DateIn" Type="Date"/>
                                        <ext:ModelField Name="Remark"/>                                                                        
                                        <ext:ModelField Name="Type"/> 
                                    </Fields>
                                </ext:Model>
                            </Model>                  
                        </ext:Store>
                    </Store>
                    <ColumnModel ID="ColumnModel5" runat="server">
                        <Columns>
                            <ext:Column ID="Column22" runat="server" Text="FileID" DataIndex="FileID" Width="80" Hidden="True" />
                            <ext:Column ID="Column24" runat="server" Text="File Name" Width="80" DataIndex="FileName" Flex="1" />
                            <ext:Column ID="Column240"  runat="server" Text="File Length" Width="80" DataIndex="FileLength" Hidden="True" />
                            <ext:Column ID="Column28"  runat="server" Text="File Type" Width="80" DataIndex="Type" />
                            <ext:DateColumn ID="DateColumn4"  runat="server" Text="Date" Width="80" DataIndex="DateIn" Format="yyyy-MM-dd" />
                            <ext:Column ID="Column27" runat="server" Text="Remark" Width="80" DataIndex="Remark" />                                                                                                                                                                                                                
                        </Columns>                        
                    </ColumnModel> 
                    <View>
                        <ext:GridView ID="GridView2" runat="server" StripeRows="true" TrackOver="true" >
                            <Listeners>
                                <CellDblClick Handler="#{hiddenFileID}.setValue(record.data.FileID+'[@]'+record.data.FileName);#{gridWinFile}.submitData(false, {isUpload:true});" />
                            </Listeners>                             
                        </ext:GridView>
                    </View>                 
                </ext:GridPanel>
            </Items>
             <Listeners>
                <Show Handler="#{fileRemark}.focus();#{cbFileType}.focus();"/><%--for IE8--%>
            </Listeners> 
        </ext:Window>
        <ext:Window ID="winUpload" runat="server" Width="300" Height="100" Hidden="True" Title="Upload Template" BodyPaddingSummary="5 5 5" Resizable="false" Icon="DiskUpload" Layout="FitLayout" Frame="True" AnimateTarget="btnBatchImport">
            <Items>                
                <ext:FileUploadField runat="server" ID="fufFile">

                </ext:FileUploadField>
            </Items>
            <Buttons>
                <ext:Button runat="server" ID="btnUploadFile" Icon="DiskUpload" Text="Upload">
                    <DirectEvents>
                        <Click 
                            OnEvent="UploadClick"
                            Before="Ext.Msg.wait('Uploading your file...', 'Uploading');"
                                
                            Failure="Ext.Msg.show({ 
                                title   : 'Error', 
                                msg     : 'Error during uploading', 
                                minWidth: 200, 
                                modal   : true, 
                                icon    : Ext.Msg.ERROR, 
                                buttons : Ext.Msg.OK 
                            });">
                        </Click>
                    </DirectEvents>
                </ext:Button>
            </Buttons>
        </ext:Window>

        <ext:Window ID="winReport" runat="server" Title="Select Report Date"  Icon="Report"  Width="236" BodyPaddingSummary="5 5 5" Modal="false" Layout="FormLayout" Hidden="true" Resizable="False" AnimateTarget="btnLogin">
            <Items>
                <ext:Label runat="server" Text="Please select your report Date:"></ext:Label> 
                <ext:DateField runat="server" ID="DateReport"></ext:DateField>
            </Items>
            <Buttons>
                <ext:Button ID="Button3" runat="server" Icon="PageSave" Text="Save" Handler="Ext.Msg.show({'title':'Please Wait','animEl':'App.Button3','buttons':false,'closable':false,'msg':'Exporting...','progress':true,'progressText':'Initializing...','width':300});App.TaskManager1.startTask('Task1');#{StoreSub}.submitData(false, {isUpload:true});#{winReport}.hide();" />
            </Buttons>               
            <Listeners>
                <Show Handler="#{DateReport}.focus();"/><%--for IE8--%>
            </Listeners>         
        </ext:Window>
        <ext:Window ID="winKPI" runat="server" Title="KPI" Icon="Report"  Width="1000" Height="600" BodyPaddingSummary="5 5 5" Modal="false" Hidden="true" Maximizable="true" Resizable="True" AnimateTarget="btnLogin" Layout="FitLayout">
            <Items>                
                <ext:TabPanel runat="server" >
                    <Items>
                        <ext:Panel ID="Tab1" Title="KPI Line" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="pOverallKPI" Html="<br/><center>Overall KPI</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart3"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">
                                    <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                    <Store>
                                        <ext:Store ID="storeKPIAll" runat="server" AutoLoad="false">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=overallKPI">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model9" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Name" />
                                                        <ext:ModelField Name="Deduction Cost(RMB)" Mapping="Data1" />
                                                        <ext:ModelField Name="Deduction Rate(%)" Mapping="Data2"/>
                                                        <ext:ModelField Name="MOM" />
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis Fields="Deduction Cost(RMB)" Title="" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="return moneyFormat2(value);"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                            <Label>
                                                <Rotate Degrees="336" />
                                            </Label>
                                        </ext:CategoryAxis>
                                        <ext:NumericAxis Fields="Deduction Rate(%)" Title="" Grid="false" Position="Right">                                                                
                                        </ext:NumericAxis>
                                    </Axes>
                                    <Series>
                                        <ext:ColumnSeries Axis="Left" XField="Name" YField="Deduction Cost(RMB)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="70" Height="28">
                                                <Renderer Handler="this.setTitle(String(moneyFormat2(item.value[1])));" />
                                            </Tips>                                                                                        
                                        </ext:ColumnSeries>
                                        <ext:LineSeries Axis="Right" Smooth="3" XField="Name" YField="Deduction Rate(%)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="60" Height="28">
                                                <Renderer Handler="this.setTitle(String(item.value[1])+'%');" />
                                            </Tips>                                                                
                                            <Style Fill="#FF0033" Stroke="#FF6666" StrokeWidth="2" />
                                            <Label Display="Over" FontSize="12" >
                                                <Renderer Handler="return(String(item.value[1])+'%');"/>
                                            </Label>                                                            
                                        </ext:LineSeries>
                                    </Series>                                                         
                                </ext:Chart>
                            </Items> 
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                    <Items>
                                        <ext:TextField runat="server" ID="txtContracts" FieldLabel="Contracts" LabelWidth="72" Text="4911967031,4913576950" Width="360" />
                                        <ext:ToolbarFill />                                        
                                        <ext:Button ID="btnKPIRefresh" runat="server" Icon="ArrowRefresh" Text="ShowContractorStatus">
                                            <Listeners>
                                                <Click Handler="#{storeKPIAll}.reload({params: { NO: #{txtContracts}.getValue()}});#{pOverallKPI}.body.update('<br/><center>'+showQuarterText()+'</center>');" />
                                            </Listeners>
                                        </ext:Button>
                                        <ext:ToolbarSeparator />
                                        <ext:DateField runat="server" ID="dfKPILineStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfKPILineEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button ID="btnOverAll" runat="server" Icon="ReportGo" Text="OverAll Statistics">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeKPIAll},#{dfKPILineStart}.getValue(),#{dfKPILineEnd}.getValue(),'true');#{pOverallKPI}.body.update('<br/><center>Overall KPI</center>');" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>     

                        </ext:Panel>
                        <ext:Panel ID="Tab2" Title="KPI Scatter" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="pScatter" Html="<br/><center>Contractor Scatter</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart4"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">
                                    <Store>
                                        <ext:Store ID="storeKPIScatter" runat="server" AutoLoad="False">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=scatterKPI">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model10" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Data1" />
                                                        <ext:ModelField Name="Data2"/>
                                                        <ext:ModelField Name="Name" Type="String"/>
                                                        <ext:ModelField Name="MOM" Type="String"/>
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                            <Listeners>
                                                <DataChanged Handler="if(#{storeKPIScatter}.data.length>0){#{pScatter}.body.update('<center>Contractor Scatter<br/>'+#{storeKPIScatter}.data.items[0].data.MOM+'</center>');}" />
                                            </Listeners>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis AxisID="axisV" Fields="Data1" Title="Deduction Cost(RMB)" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="if(#{cbDistribution}.getValue()){return value+'%';}else{return moneyFormat2(value);}"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:NumericAxis Fields="Data2" Title="Deduction Rate(%)" Grid="true" Position="Bottom">  
                                            <Label>
                                                <Renderer Handler="return value+'%';"/>
                                            </Label>                                                              
                                        </ext:NumericAxis>
                                    </Axes>
                                    <Series>
                                        <ext:ScatterSeries XField="Data2" YField="Data1" Highlight="True">
                                             <Tips runat="server" TrackMouse="true" Width="300" Height="50">
                                                <Renderer Handler="if(#{cbDistribution}.getValue()){this.setTitle(String(item.storeItem.data.Name)+':<br/>('+String(item.value[0])+'%,'+String(item.value[1])+'%)');}else{this.setTitle(String(item.storeItem.data.Name)+':<br/>('+String(item.value[0])+'%,'+String(item.value[1])+')');}" />
                                            </Tips>
                                            <%--<Label Field="Data2" Display="Over">
                                                <Renderer Handler="return '('+String(item.value[0])+'%,'+String(item.value[1])+')';" />
                                            </Label> --%>
                                        </ext:ScatterSeries>
                                    </Series>                                                         
                                </ext:Chart>
                            </Items>
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                    <Items>
                                        <ext:ToolbarFill></ext:ToolbarFill>
                                        <ext:Checkbox runat="server" ID="cbDistribution" FieldLabel="Distribution?" LabelWidth="16">
                                            <Listeners>
                                                <Change Handler="if(#{cbDistribution}.getValue()){#{Chart4}.axes.items[0].setTitle('Overall Deduction Rate Distribution(%)');scatterStoreReload(#{storeKPIScatter},#{dfKPIStart}.getValue(),#{dfKPIEnd}.getValue(),true);}else{#{Chart4}.axes.items[0].setTitle('Deduction Cost(RMB)');scatterStoreReload(#{storeKPIScatter},#{dfKPIStart}.getValue(),#{dfKPIEnd}.getValue(),false);}" />
                                            </Listeners>
                                        </ext:Checkbox>
                                        <ext:DateField runat="server" ID="dfKPIStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfKPIEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button runat="server" Text="Submit">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeKPIScatter},#{dfKPIStart}.getValue(),#{dfKPIEnd}.getValue(),#{cbDistribution}.getValue());" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>
                        </ext:Panel>   
                        <ext:Panel ID="Tab3" Title="User KPI" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="pUserKPI" Html="<br/><center>User Deduction</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart7"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">
                                    <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                    <Store>
                                        <ext:Store ID="storeUserKPI" runat="server" AutoLoad="false">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=userKPI">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model13" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Name" />
                                                        <ext:ModelField Name="Deduction Cost(RMB)" Mapping="Data1" />
                                                        <ext:ModelField Name="Deduction Rate(%)" Mapping="Data2"/>
                                                        <ext:ModelField Name="MOM" />
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis Fields="Deduction Cost(RMB)" Title="" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="return moneyFormat2(value);"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                            <Label>
                                                <Rotate Degrees="270" />
                                            </Label>
                                        </ext:CategoryAxis>
                                        <ext:NumericAxis Fields="Deduction Rate(%)" Title="" Grid="false" Position="Right">                                                                
                                        </ext:NumericAxis>
                                    </Axes>
                                    <Series>
                                        <ext:ColumnSeries Axis="Left" XField="Name" YField="Deduction Cost(RMB)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="70" Height="28">
                                                <Renderer Handler="this.setTitle(String(moneyFormat2(item.value[1])));" />
                                            </Tips>                                                                                        
                                        </ext:ColumnSeries>
                                        <ext:LineSeries Axis="Right" Smooth="3" XField="Name" YField="Deduction Rate(%)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="60" Height="28">
                                                <Renderer Handler="this.setTitle(String(item.value[1])+'%');" />
                                            </Tips>                                                                
                                            <Style Fill="#FF0033" Stroke="#FF6666" StrokeWidth="2" />
                                            <Label Display="Over" FontSize="12" >
                                                <Renderer Handler="return(String(item.value[1])+'%');"/>
                                            </Label>                                                            
                                        </ext:LineSeries>
                                    </Series>                                                         
                                </ext:Chart>
                            </Items> 
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                   <Items>
                                        <ext:ToolbarFill></ext:ToolbarFill>                                        
                                        <ext:DateField runat="server" ID="dfUserKPIStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfUserKPIEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button runat="server" Text="Submit">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeUserKPI},#{dfUserKPIStart}.getValue(),#{dfUserKPIEnd}.getValue(),'true');" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>     

                        </ext:Panel>

                        <ext:Panel ID="Tab4" Title="Sample Analyse" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="pUserSample" Html="<br/><center>User Sample Analyse</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart8"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">
                                    <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                    <Store>
                                        <ext:Store ID="storeUserSample" runat="server" AutoLoad="false">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=userSample">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model14" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Name" />
                                                        <ext:ModelField Name="SSR Cost(RMB)" Mapping="Data1" />
                                                        <ext:ModelField Name="SSR Count" Mapping="Data2"/>
                                                        <ext:ModelField Name="MOM" />
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                            <Listeners>
                                                <DataChanged Handler="if(#{storeUserSample}.data.length>0){#{pUserSample}.body.update(#{storeUserSample}.data.items[0].data.MOM);}" />
                                            </Listeners>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis Fields="SSR Cost(RMB)" Title="Cost(RMB)" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="return moneyFormat2(value);"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                            <Label>
                                                <Rotate Degrees="270" />
                                            </Label>
                                        </ext:CategoryAxis>
                                        <ext:NumericAxis Fields="SSR Count" Title="Count" Grid="false" Position="Right">                                                                
                                        </ext:NumericAxis>
                                    </Axes>
                                    <Series>
                                        <ext:ColumnSeries Axis="Left" XField="Name" YField="SSR Cost(RMB)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="70" Height="28">
                                                <Renderer Handler="this.setTitle(String(moneyFormat2(item.value[1])));" />
                                            </Tips>                                                                                        
                                        </ext:ColumnSeries>
                                        <ext:LineSeries Axis="Right" Smooth="3" XField="Name" YField="SSR Count" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="60" Height="28">
                                                <Renderer Handler="this.setTitle(String(item.value[1]));" />
                                            </Tips>                                                                
                                            <Style Fill="#FF0033" Stroke="#FF6666" StrokeWidth="2" />
                                            <Label Display="Over" FontSize="12" >
                                                <Renderer Handler="return(String(item.value[1]));"/>
                                            </Label>                                                            
                                        </ext:LineSeries>
                                    </Series>                                                         
                                </ext:Chart>
                            </Items> 
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                   <Items>
                                        <ext:ToolbarFill></ext:ToolbarFill>                                        
                                        <ext:DateField runat="server" ID="dfUserSampleStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfUserSampleEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button runat="server" Text="Submit">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeUserSample},#{dfUserSampleStart}.getValue(),#{dfUserSampleEnd}.getValue(),'true');" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>     

                        </ext:Panel>

                        <ext:Panel ID="Tab5" Title="Contractor/Dis KPI" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="pContractorKPI" Html="<br/><center>Contractor Cost</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart9"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">                                    
                                    <Store>
                                        <ext:Store ID="storeContractorKPI" runat="server" AutoLoad="false">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=contractorKPI">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model15" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Name" />
                                                        <ext:ModelField Name="SSR Cost(RMB)" Mapping="Data1" /> 
                                                        <ext:ModelField Name="MOM" />
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                            <Listeners>
                                                <DataChanged Handler="if(#{storeContractorKPI}.data.length>0){#{pContractorKPI}.body.update(#{storeContractorKPI}.data.items[0].data.MOM);}" />
                                            </Listeners>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis Fields="SSR Cost(RMB)" Title="Cost(K RMB)" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="return moneyFormat2(value);"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:CategoryAxis Position="Bottom" Fields="Name" Title="">
                                            <Label>
                                                <Rotate Degrees="336" />
                                            </Label>
                                        </ext:CategoryAxis>                                        
                                    </Axes>
                                    <Series>
                                        <ext:ColumnSeries Axis="Left" XField="Name" YField="SSR Cost(RMB)" Highlight="true">
                                            <Tips runat="server" TrackMouse="true" Width="200" Height="30">
                                                <Renderer Handler="this.setTitle(storeItem.get('Name') + '<br/>'+String(moneyFormat2(item.value[1])));" />
                                            </Tips>   
                                            <Label Display="Outside" Field="SSR Cost(RMB)">     
                                                <Renderer Handler="return String(item.value[1])+'K';"/>
                                            </Label>
                                        </ext:ColumnSeries>                                        
                                    </Series>                                                         
                                </ext:Chart>
                            </Items> 
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                   <Items>
                                        <ext:ToolbarFill></ext:ToolbarFill>
                                        <ext:Checkbox runat="server" ID="cbContractorKPI" FieldLabel="Discipline?" LabelWidth="32">
                                            <Listeners>
                                                <Change Handler="if(#{cbContractorKPI}.getValue()){#{pContractorKPI}.body.update('<br/><center>Discipline Cost</center>');scatterStoreReload(#{storeContractorKPI},#{dfContractorKPIStart}.getValue(),#{dfContractorKPIEnd}.getValue(),#{cbContractorKPI}.getValue());}else{#{pContractorKPI}.body.update('<br/><center>Contractor Cost</center>');scatterStoreReload(#{storeContractorKPI},#{dfContractorKPIStart}.getValue(),#{dfContractorKPIEnd}.getValue(),#{cbContractorKPI}.getValue());}" />
                                            </Listeners>
                                        </ext:Checkbox>                                       
                                        <ext:DateField runat="server" ID="dfContractorKPIStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfContractorKPIEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button runat="server" Text="Submit">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeContractorKPI},#{dfContractorKPIStart}.getValue(),#{dfContractorKPIEnd}.getValue(),#{cbContractorKPI}.getValue());" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>     

                        </ext:Panel>

                        <ext:Panel ID="Tab6" Title="User Proportion" runat="server" Closable="false">
                            <LayoutConfig>
                                <ext:VBoxLayoutConfig Align="Stretch" />
                            </LayoutConfig>
                            <Items>
                                <ext:Panel runat="server" ID="Panel5" Html="<br/><center>Discipline Proportion of User</center>" Height="30" Border="false" />
                                <ext:Chart ID="Chart10"  runat="server" Animate="true" Shadow="true" InsetPadding="20" Flex="1">
                                    <LegendConfig Position="Right" BoxStrokeWidth="0" /> 
                                    <Store>
                                        <ext:Store ID="storeUserPro" runat="server" AutoLoad="false">                           
                                            <Proxy>
                                                <ext:AjaxProxy Url="data.ashx?object=userPro">
                                                    <ActionMethods Read="GET" />                                                                                                                                                                                                
                                                </ext:AjaxProxy>
                                            </Proxy>
                                            <Model>
                                                <ext:Model ID="Model16" runat="server">
                                                    <Fields>
                                                        <ext:ModelField Name="Section" />
                                                        <ext:ModelField Name="BCM" />
                                                        <ext:ModelField Name="CAD" />
                                                        <ext:ModelField Name="CAL" />
                                                        <ext:ModelField Name="CCS" />
                                                        <ext:ModelField Name="CIV" />
                                                        <ext:ModelField Name="CLE" />
                                                        <ext:ModelField Name="EHR" />
                                                        <ext:ModelField Name="EIC" />
                                                        <ext:ModelField Name="ENG" />
                                                        <ext:ModelField Name="EPC" />
                                                        <ext:ModelField Name="EPM" />
                                                        <ext:ModelField Name="EPR" />
                                                        <ext:ModelField Name="ERM" />
                                                        <ext:ModelField Name="FFS" />
                                                        <ext:ModelField Name="FSY" />
                                                        <ext:ModelField Name="HVA" />
                                                        <ext:ModelField Name="INS" />
                                                        <ext:ModelField Name="IPV" />
                                                        <ext:ModelField Name="IVR" />
                                                        <ext:ModelField Name="LES" />
                                                        <ext:ModelField Name="LIF" />
                                                        <ext:ModelField Name="MAS" />
                                                        <ext:ModelField Name="MOR" />
                                                        <ext:ModelField Name="NDE" />
                                                        <ext:ModelField Name="PAI" />
                                                        <ext:ModelField Name="PQS" />
                                                        <ext:ModelField Name="QSS" />
                                                        <ext:ModelField Name="RBC" />
                                                        <ext:ModelField Name="ROC" />
                                                        <ext:ModelField Name="SCA" />
                                                        <ext:ModelField Name="SLP" />
                                                        <ext:ModelField Name="STS" />
                                                        <ext:ModelField Name="TRA" />
                                                        <ext:ModelField Name="CFM" />
                                                        <ext:ModelField Name="HOT" />
                                                        <ext:ModelField Name="OEM" />                                                       
                                                    </Fields>
                                                </ext:Model>
                                            </Model>
                                        </ext:Store>
                                    </Store>
                                    <Axes>
                                        <ext:NumericAxis Fields="BCM,CAD,CAL,CCS,CIV,CLE,EHR,EIC,ENG,EPC,EPM,EPR,ERM,FFS,FSY,HVA,INS,IPV,IVR,LES,LIF,MAS,MOR,NDE,PAI,PQS,QSS,RBC,ROC,SCA,SLP,STS,TRA,CFM,HOT,OEM" Title="" Grid="true" Position="Left">
                                            <Label>
                                                <Renderer Handler="return moneyFormat2(value);"/>
                                            </Label>
                                        </ext:NumericAxis>
                                        <ext:CategoryAxis Position="Bottom" Fields="Section" Title="">  
                                              <Label>
                                                <Rotate Degrees="270" />
                                              </Label>                                        
                                        </ext:CategoryAxis>
                                    </Axes>
                                    <Series>
                                        <ext:ColumnSeries Axis="Left" XField="Section" YField="BCM,CAD,CAL,CCS,CIV,CLE,EHR,EIC,ENG,EPC,EPM,EPR,ERM,FFS,FSY,HVA,INS,IPV,IVR,LES,LIF,MAS,MOR,NDE,PAI,PQS,QSS,RBC,ROC,SCA,SLP,STS,TRA,CFM,HOT,OEM" Highlight="true" Stacked="true">
                                            <Tips runat="server" TrackMouse="true" Width="70" Height="36">
                                                <Renderer Handler="this.setTitle(String(item.yField)+'<br/>'+String(item.value[1])+'%');" />
                                            </Tips>                                                                                        
                                        </ext:ColumnSeries>                                        
                                    </Series>                                                         
                                </ext:Chart>
                            </Items> 
                            <BottomBar>
                                <ext:Toolbar runat="server">
                                   <Items>
                                        <ext:ToolbarFill></ext:ToolbarFill>
                                        <ext:ComboBox ID="cbDepart"  runat="server" FieldLabel="Department" ForceSelection="True" AllowBlank="False" LabelWidth="60" EmptyText="All">
                                            <Items>
                                                <ext:ListItem Text="All"/>
                                                <ext:ListItem Text="CTA"/>
                                                <ext:ListItem Text="CTE"/>
                                                <ext:ListItem Text="CTM"/>
                                            </Items>
                                        </ext:ComboBox>                                           
                                        <ext:DateField runat="server" ID="dfUserProStart" Format="yyyy-MM" FieldLabel="Time Range" LabelWidth="72"></ext:DateField>
                                        <ext:Label runat="server" Text="-"/>
                                        <ext:DateField runat="server" ID="dfUserProEnd" Format="yyyy-MM"></ext:DateField>
                                        <ext:Button runat="server" Text="Submit">
                                            <Listeners>
                                                <Click Handler="scatterStoreReload(#{storeUserPro},#{dfUserProStart}.getValue(),#{dfUserProEnd}.getValue(),#{cbDepart}.getValue());" />
                                            </Listeners>
                                        </ext:Button>
                                    </Items>
                                </ext:Toolbar>
                            </BottomBar>     

                        </ext:Panel>

                     </Items>
                     <Listeners>
                         <TabChange Handler="if(newTab.getId()=='Tab2'){scatterStoreReload(#{storeKPIScatter},#{dfKPIStart}.getValue(),#{dfKPIEnd}.getValue(),#{cbDistribution}.getValue());}else if(newTab.getId()=='Tab3'){scatterStoreReload(#{storeUserKPI},#{dfUserKPIStart}.getValue(),#{dfUserKPIEnd}.getValue(),'true');}else if(newTab.getId()=='Tab4'){scatterStoreReload(#{storeUserSample},#{dfUserSampleStart}.getValue(),#{dfUserSampleEnd}.getValue(),'true');}else if(newTab.getId()=='Tab5'){scatterStoreReload(#{storeContractorKPI},#{dfContractorKPIStart}.getValue(),#{dfContractorKPIEnd}.getValue(),#{cbContractorKPI}.getValue());}else if(newTab.getId()=='Tab6'){scatterStoreReload(#{storeUserPro},#{dfUserProStart}.getValue(),#{dfUserProEnd}.getValue(),#{cbDepart}.getValue());}" />
                     </Listeners>           
                </ext:TabPanel>                                 
            </Items>              
            <Listeners>
                <Show Handler="scatterStoreReload(#{storeKPIAll},#{dfKPILineStart}.getValue(),#{dfKPILineEnd}.getValue(),'true');#{pOverallKPI}.body.update('<br/><center>Overall KPI</center>');"/>
            </Listeners>         
        </ext:Window>
        <ext:TaskManager ID="TaskManager1" runat="server">
            <Tasks>
                <ext:Task 
                    TaskID="Task1"
                    Interval="1000" 
                    AutoRun="false">
                    <DirectEvents>
                        <Update OnEvent="RefreshProgress" />
                    </DirectEvents>                    
                </ext:Task>
            </Tasks>
        </ext:TaskManager>
        <ext:Hidden runat="server" ID="hiddenValue"></ext:Hidden>
        <ext:Hidden runat="server" ID="hiddenUser"></ext:Hidden>
        <ext:Hidden runat="server" ID="hiddenDateMonth"></ext:Hidden>
        <ext:Hidden runat="server" ID="fileContractNO">
            <Listeners>
                <Change Handler="#{fileRemark}.setValue(this.getValue());#{winFile}.setTitle('File Management('+this.getValue()+')')"/>
            </Listeners>
        </ext:Hidden>
        <ext:Hidden runat="server" ID="hiddenFileID"></ext:Hidden>
        <ext:Hidden runat="server" ID="hiddenSelectedNO"></ext:Hidden>
        <ext:Spotlight ID="Spot" runat="server" Easing="EaseInOut" Duration="6" />
        <ext:ToolTip ID="ToolTip1" 
            runat="server" 
            Target="={#{gridWinFile}.getView().el}"
            Delegate="={#{gridWinFile}.getView().itemSelector}"
            TrackMouse="true">
            <Listeners>
                <Show Handler="onShow(this, #{gridWinFile});" />
            </Listeners>
        </ext:ToolTip> 
        <ext:ToolTip ID="ToolTip2" 
            runat="server" 
            Target="={#{GridPanelFile2}.getView().el}"
            Delegate="={#{GridPanelFile2}.getView().itemSelector}"
            TrackMouse="true">
            <Listeners>
                <Show Handler="onShow(this, #{GridPanelFile2});" />
            </Listeners>
        </ext:ToolTip> 
        <ext:ToolTip ID="ToolTip3" 
            runat="server" 
            Target="={#{gridEvaluationContract}.getView().el}"
            Delegate=".x-grid-cell"
            TrackMouse="true">
            <Listeners>
                <Show Handler="onShowMain(this, #{gridEvaluationContract});" /> 
            </Listeners>
        </ext:ToolTip> 
    </form>
</body>
</html>
