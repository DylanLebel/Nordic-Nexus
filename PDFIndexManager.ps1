#==============================================================================
#  PDF / DXF / MODEL INDEX MANAGER - Professional Edition
#  "One-Click Drawing Collector" - Crawler & Index Tool
#==============================================================================
#  Features:
#    - WPF GUI with professional dark theme
#    - PDF Crawler tab + DXF Crawler tab + Model Crawler tab
#    - Section-by-section crawling (pick which folders to scan)
#    - Real-time progress with background threading
#    - Incremental indexing (append to existing, don't re-crawl everything)
#    - One-click deduplication with smart revision logic
#    - Collect tab: PDF only / DXF only / PDF+DXF toggle
#    - Model index output for SolidWorks assembly lookup
#    - Index viewer with search and stats
#    - Full logging
#==============================================================================

param(
    [string]$BOMFile = "",         # Pre-load a BOM file and open on Collect tab
    [string]$CollectOutput = ""    # Pre-set output folder for Collect tab
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

# ---- CONFIGURATION ----
$script:OutputDir      = "C:\Users\dlebel\Documents\PDFIndex"

# PDF index files
$script:RawCSV         = Join-Path $script:OutputDir "pdf_index_raw.csv"
$script:CleanCSV       = Join-Path $script:OutputDir "pdf_index_clean.csv"
$script:LogFile        = Join-Path $script:OutputDir "crawl_log.txt"
$script:StateFile      = Join-Path $script:OutputDir "crawl_state.json"

# DXF index files (same folder, separate files)
$script:DxfRawCSV      = Join-Path $script:OutputDir "dxf_index_raw.csv"
$script:DxfCleanCSV    = Join-Path $script:OutputDir "dxf_index_clean.csv"
$script:DxfStateFile   = Join-Path $script:OutputDir "dxf_crawl_state.json"

# Model index files (SLDASM + SLDPRT)
$script:ModelRawCSV    = Join-Path $script:OutputDir "model_index_raw.csv"
$script:ModelAllCSV    = Join-Path $script:OutputDir "model_index_all.csv"
$script:ModelCleanCSV  = Join-Path $script:OutputDir "model_index_clean.csv"
$script:ModelStateFile = Join-Path $script:OutputDir "model_crawl_state.json"

# Default roots - user can add/remove via UI
$script:DefaultRoots = @(
    "J:\",
    "Y:\",
    "J:\Epicor",
    "J:\MFL Jobs",
    "J:\NordicMinesteel",
    "C:\NMT_PDM"
)

# Ensure output dir exists
if (-not (Test-Path $script:OutputDir)) {
    New-Item -ItemType Directory -Path $script:OutputDir -Force | Out-Null
}

# ============================================================================
#  XAML UI DEFINITION
# ============================================================================

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Nordic Minesteel - PDF/DXF/Model Index Manager"
    Width="920" Height="730"
    MinWidth="750" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    Background="#1a1a2e">

    <Window.Resources>
        <Style x:Key="HeaderText" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,0,8"/>
        </Style>
        <Style x:Key="SubText" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="FontSize" Value="11"/>
        </Style>
        <Style x:Key="StatValue" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#2ea3f2"/>
            <Setter Property="FontSize" Value="22"/>
            <Setter Property="FontWeight" Value="Bold"/>
        </Style>
        <Style x:Key="StatLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="Margin" Value="0,2,0,0"/>
        </Style>
        <Style x:Key="ActionButton" TargetType="Button">
            <Setter Property="Background" Value="#2ea3f2"/>
            <Setter Property="Foreground" Value="#ffffff"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="20,10"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#1a8fd8"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#0f3460"/>
                                <Setter Property="Foreground" Value="#6c7086"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="DxfButton" TargetType="Button" BasedOn="{StaticResource ActionButton}">
            <Setter Property="Background" Value="#4caf50"/>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource ActionButton}">
            <Setter Property="Background" Value="#f44336"/>
        </Style>
        <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource ActionButton}">
            <Setter Property="Background" Value="#0f3460"/>
            <Setter Property="Foreground" Value="#e8e8e8"/>
        </Style>
        <Style x:Key="CardBorder" TargetType="Border">
            <Setter Property="Background" Value="#16213e"/>
            <Setter Property="CornerRadius" Value="8"/>
            <Setter Property="Padding" Value="16"/>
            <Setter Property="Margin" Value="0,0,0,12"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="Border" Background="Transparent"
                                Padding="{TemplateBinding Padding}"
                                BorderThickness="0,0,0,2" BorderBrush="Transparent">
                            <ContentPresenter ContentSource="Header"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="Border" Property="BorderBrush" Value="#2ea3f2"/>
                                <Setter Property="Foreground" Value="#e8e8e8"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Foreground" Value="#e8e8e8"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="Margin" Value="0,4"/>
        </Style>
        <Style TargetType="RadioButton">
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="Margin" Value="0,2"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
    </Window.Resources>

    <Grid Margin="20">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER with branding -->
        <StackPanel Grid.Row="0" Margin="0,0,0,16">
            <TextBlock Text="NORDIC MINESTEEL TECHNOLOGIES" Foreground="#2ea3f2"
                       FontSize="10" FontWeight="Bold" Margin="0,0,0,4"/>
            <TextBlock Text="Drawing Index Manager" Foreground="#e8e8e8"
                       FontSize="24" FontWeight="Bold"/>
            <TextBlock Text="Crawl, index, and collect PDF, DXF, and model files across network drives"
                       Foreground="#6c7086" FontSize="12" Margin="0,4,0,0"/>
        </StackPanel>

        <!-- MAIN TABS -->
        <TabControl Grid.Row="1" Background="Transparent" BorderThickness="0" Padding="0">

            <!-- TAB 1: PDF CRAWLER -->
            <TabItem Header="&#x1F4C4;  PDF Crawler">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,12,0,0">
                    <StackPanel>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Scan Sections" Style="{StaticResource HeaderText}"/>
                                <TextBlock Text="Select root folders to crawl for PDF files."
                                           Style="{StaticResource SubText}" Margin="0,0,0,12" TextWrapping="Wrap"/>
                                <StackPanel x:Name="FolderCheckboxPanel"/>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnAddFolder" Content="+ Add Folder"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnRemoveFolder" Content="- Remove Selected"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Crawl Controls" Style="{StaticResource HeaderText}"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock x:Name="TxtCrawlStatus" Text="Ready"
                                                   Foreground="#4caf50" FontSize="13" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="TxtCrawlDetail" Text="Select sections and click Start Crawl"
                                                   Style="{StaticResource SubText}" Margin="0,4,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                        <CheckBox x:Name="ChkParallel" Content="Parallel" IsChecked="True"
                                                  Foreground="#a0a0b8" FontSize="11" VerticalAlignment="Center"
                                                  Margin="0,0,12,0" ToolTip="Scan folders simultaneously using multiple threads"/>
                                        <Button x:Name="BtnSearchPDM" Content="&#x1F5C4;  Search PDM Vault"
                                                Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnStartCrawl" Content="&#x25B6;  Start PDF Crawl"
                                                Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnStopCrawl" Content="&#x25A0;  Stop"
                                                Style="{StaticResource DangerButton}" IsEnabled="False"/>
                                    </StackPanel>
                                </Grid>
                                <ProgressBar x:Name="CrawlProgress" Height="6"
                                             Background="#0f3460" Foreground="#2ea3f2"
                                             BorderThickness="0" Minimum="0" Maximum="100" Value="0"/>
                                <TextBlock x:Name="TxtProgressDetail" Text=""
                                           Style="{StaticResource SubText}" Margin="0,6,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="PDF Index Statistics" Style="{StaticResource HeaderText}"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatTotalPDFs" Text="--" Style="{StaticResource StatValue}"/>
                                        <TextBlock Text="Total PDFs" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatUniqueParts" Text="--" Style="{StaticResource StatValue}"/>
                                        <TextBlock Text="Unique Parts" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatDuplicates" Text="--" Style="{StaticResource StatValue}"/>
                                        <TextBlock Text="Duplicates" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatErrors" Text="--" Style="{StaticResource StatValue}"/>
                                        <TextBlock Text="Errors" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                </Grid>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnRunDedup" Content="Run Deduplication"
                                            Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenClean" Content="Open Clean CSV"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenRaw" Content="Open Raw CSV"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                                <TextBlock x:Name="TxtDedupStatus" Text="Run after crawling to generate the clean index."
                                           Style="{StaticResource SubText}" Margin="0,8,0,0"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- TAB 2: DXF CRAWLER -->
            <TabItem Header="&#x1F4D0;  DXF Crawler">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,12,0,0">
                    <StackPanel>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="DXF Scan Sections" Style="{StaticResource HeaderText}"/>
                                <TextBlock Text="Select root folders to crawl for DXF files. Uses the same folder list as PDF crawler."
                                           Style="{StaticResource SubText}" Margin="0,0,0,12" TextWrapping="Wrap"/>
                                <StackPanel x:Name="DxfFolderCheckboxPanel"/>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnDxfAddFolder" Content="+ Add Folder"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnDxfRemoveFolder" Content="- Remove Selected"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="DXF Crawl Controls" Style="{StaticResource HeaderText}"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock x:Name="TxtDxfCrawlStatus" Text="Ready"
                                                   Foreground="#4caf50" FontSize="13" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="TxtDxfCrawlDetail" Text="Select sections and click Start DXF Crawl"
                                                   Style="{StaticResource SubText}" Margin="0,4,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                        <CheckBox x:Name="ChkDxfParallel" Content="Parallel" IsChecked="True"
                                                  Foreground="#a0a0b8" FontSize="11" VerticalAlignment="Center"
                                                  Margin="0,0,12,0" ToolTip="Scan folders simultaneously using multiple threads"/>
                                        <Button x:Name="BtnStartDxfCrawl" Content="&#x25B6;  Start DXF Crawl"
                                                Style="{StaticResource DxfButton}" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnStopDxfCrawl" Content="&#x25A0;  Stop"
                                                Style="{StaticResource DangerButton}" IsEnabled="False"/>
                                    </StackPanel>
                                </Grid>
                                <ProgressBar x:Name="DxfCrawlProgress" Height="6"
                                             Background="#0f3460" Foreground="#4caf50"
                                             BorderThickness="0" Minimum="0" Maximum="100" Value="0"/>
                                <TextBlock x:Name="TxtDxfProgressDetail" Text=""
                                           Style="{StaticResource SubText}" Margin="0,6,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="DXF Index Statistics" Style="{StaticResource HeaderText}"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatTotalDXFs" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Total DXFs" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatUniqueDxfParts" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Unique Parts" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatDxfDuplicates" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Duplicates" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatDxfErrors" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Errors" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                </Grid>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnRunDxfDedup" Content="Run DXF Deduplication"
                                            Style="{StaticResource DxfButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenDxfClean" Content="Open Clean CSV"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenDxfRaw" Content="Open Raw CSV"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                                <TextBlock x:Name="TxtDxfDedupStatus" Text="Run after crawling to generate the clean DXF index."
                                           Style="{StaticResource SubText}" Margin="0,8,0,0"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- TAB 3: MODEL CRAWLER -->
            <TabItem Header="&#x1F5C2;  Model Crawler">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,12,0,0">
                    <StackPanel>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Model Scan Sections" Style="{StaticResource HeaderText}"/>
                                <TextBlock Text="Select root folders to crawl for SolidWorks model files (.SLDASM + .SLDPRT)."
                                           Style="{StaticResource SubText}" Margin="0,0,0,12" TextWrapping="Wrap"/>
                                <StackPanel x:Name="ModelFolderCheckboxPanel"/>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnModelAddFolder" Content="+ Add Folder"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnModelRemoveFolder" Content="- Remove Selected"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Model Crawl Controls" Style="{StaticResource HeaderText}"/>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock x:Name="TxtModelCrawlStatus" Text="Ready"
                                                   Foreground="#4caf50" FontSize="13" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="TxtModelCrawlDetail" Text="Select sections and click Start Model Crawl"
                                                   Style="{StaticResource SubText}" Margin="0,4,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                        <CheckBox x:Name="ChkModelParallel" Content="Parallel" IsChecked="True"
                                                  Foreground="#a0a0b8" FontSize="11" VerticalAlignment="Center"
                                                  Margin="0,0,12,0" ToolTip="Scan folders simultaneously using multiple threads"/>
                                        <Button x:Name="BtnStartModelCrawl" Content="&#x25B6;  Start Model Crawl"
                                                Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                        <Button x:Name="BtnStopModelCrawl" Content="&#x25A0;  Stop"
                                                Style="{StaticResource DangerButton}" IsEnabled="False"/>
                                    </StackPanel>
                                </Grid>
                                <ProgressBar x:Name="ModelCrawlProgress" Height="6"
                                             Background="#0f3460" Foreground="#2ea3f2"
                                             BorderThickness="0" Minimum="0" Maximum="100" Value="0"/>
                                <TextBlock x:Name="TxtModelProgressDetail" Text=""
                                           Style="{StaticResource SubText}" Margin="0,6,0,0"/>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Model Index Statistics" Style="{StaticResource HeaderText}"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatTotalModels" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Total Models" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatModelAssemblies" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Assemblies" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatModelParts" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Parts" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatModelErrors" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="Errors" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                </Grid>
                                <StackPanel Orientation="Horizontal" Margin="0,12,0,0">
                                    <Button x:Name="BtnRunModelDedup" Content="Run Model Deduplication"
                                            Style="{StaticResource ActionButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenModelAll" Content="Open All Models CSV"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnOpenModelClean" Content="Open Clean CSV"
                                            Style="{StaticResource SecondaryButton}"/>
                                </StackPanel>
                                <TextBlock x:Name="TxtModelDedupStatus" Text="Run after model crawl to generate clean model index."
                                           Style="{StaticResource SubText}" Margin="0,8,0,0"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- TAB 4: LOG -->
            <TabItem Header="&#x1F4CB;  Log">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,8">
                        <Button x:Name="BtnClearLog" Content="Clear Log"
                                Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                        <Button x:Name="BtnOpenLogFile" Content="Open Log File"
                                Style="{StaticResource SecondaryButton}"/>
                    </StackPanel>
                    <Border Grid.Row="1" Background="#16213e" CornerRadius="8" Padding="12">
                        <ScrollViewer x:Name="LogScroller" VerticalScrollBarVisibility="Auto">
                            <TextBlock x:Name="TxtLog" Foreground="#a0a0b8"
                                       FontFamily="Consolas" FontSize="11" TextWrapping="Wrap"/>
                        </ScrollViewer>
                    </Border>
                </Grid>
            </TabItem>

            <!-- TAB 5: COLLECT -->
            <TabItem x:Name="TabCollect" Header="&#x1F4E6;  Collect">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,12,0,0">
                    <StackPanel>

                        <!-- BOM Source Card -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Assembly BOM" Style="{StaticResource HeaderText}"/>
                                <TextBlock TextWrapping="Wrap" Style="{StaticResource SubText}" Margin="0,0,0,10">
                                    Path to the part-number list exported by the SolidWorks macro (or any text file with one part number per line).
                                </TextBlock>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="TxtBOMPath" Grid.Column="0"
                                             Background="#16213e" Foreground="#e8e8e8"
                                             BorderBrush="#0f3460" Padding="8,6"
                                             FontFamily="Consolas" FontSize="11"
                                             VerticalContentAlignment="Center"/>
                                    <Button x:Name="BtnBrowseBOM" Grid.Column="1"
                                            Content="Browse..." Style="{StaticResource SecondaryButton}"
                                            Margin="8,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- Output Folder Card -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Output Folder" Style="{StaticResource HeaderText}"/>
                                <TextBlock TextWrapping="Wrap" Style="{StaticResource SubText}" Margin="0,0,0,10">
                                    Matching files will be copied here. PDFs go into the root; DXFs go into a DXFs\ subfolder.
                                </TextBlock>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="TxtCollectOutput" Grid.Column="0"
                                             Background="#16213e" Foreground="#e8e8e8"
                                             BorderBrush="#0f3460" Padding="8,6"
                                             FontFamily="Consolas" FontSize="11"
                                             VerticalContentAlignment="Center"/>
                                    <Button x:Name="BtnBrowseCollectOutput" Grid.Column="1"
                                            Content="Browse..." Style="{StaticResource SecondaryButton}"
                                            Margin="8,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>

                        <!-- File Type Toggle Card -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Collect Mode" Style="{StaticResource HeaderText}"/>
                                <TextBlock TextWrapping="Wrap" Style="{StaticResource SubText}" Margin="0,0,0,10">
                                    Choose which file types to collect for each part number.
                                </TextBlock>
                                <StackPanel Orientation="Horizontal">
                                    <RadioButton x:Name="RbPdfOnly" Content="&#x1F4C4;  PDF Only"
                                                 IsChecked="True" GroupName="CollectMode" Margin="0,0,24,0"/>
                                    <RadioButton x:Name="RbDxfOnly" Content="&#x1F4D0;  DXF Only (Burn Profile)"
                                                 GroupName="CollectMode" Margin="0,0,24,0"/>
                                    <RadioButton x:Name="RbBoth" Content="&#x1F4C4;&#x2B;&#x1F4D0;  PDF + DXF"
                                                 GroupName="CollectMode"/>
                                </StackPanel>
                            </StackPanel>
                        </Border>

                        <!-- Collect Action Card -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0">
                                        <TextBlock x:Name="TxtCollectStatus" Text="Ready — set BOM path and output folder, then click Collect."
                                                   Foreground="#a0a0b8" FontSize="13" FontWeight="SemiBold" TextWrapping="Wrap"/>
                                        <TextBlock x:Name="TxtCollectDetail" Text=""
                                                   Style="{StaticResource SubText}" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                    </StackPanel>
                                    <Button x:Name="BtnCollect" Grid.Column="1"
                                            Content="&#x1F4CB;  Collect Files"
                                            Style="{StaticResource ActionButton}"
                                            Padding="24,12" Margin="16,0,0,0"/>
                                </Grid>
                                <ProgressBar x:Name="CollectProgress" Height="6"
                                             Background="#0f3460" Foreground="#4caf50"
                                             BorderThickness="0" Minimum="0" Maximum="100" Value="0"/>
                            </StackPanel>
                        </Border>

                        <!-- Results Card -->
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <Grid Margin="0,0,0,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="Results" Style="{StaticResource HeaderText}" Margin="0"/>
                                    <Button x:Name="BtnOpenCollectFolder" Grid.Column="1"
                                            Content="Open Folder" Style="{StaticResource SecondaryButton}"
                                            IsEnabled="False"/>
                                </Grid>
                                <Grid Margin="0,0,0,12">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatCollectTotal" Text="--" Style="{StaticResource StatValue}"/>
                                        <TextBlock Text="Parts in BOM" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatCollectPdfs" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="PDFs Copied" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatCollectDxfs" Text="--" Style="{StaticResource StatValue}" Foreground="#2ea3f2"/>
                                        <TextBlock Text="DXFs Copied" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="3" HorizontalAlignment="Center">
                                        <TextBlock x:Name="StatCollectMissing" Text="--" Style="{StaticResource StatValue}" Foreground="#f44336"/>
                                        <TextBlock Text="Not Found" Style="{StaticResource StatLabel}"/>
                                    </StackPanel>
                                </Grid>
                                <Border Background="#16213e" CornerRadius="6" Padding="10" Height="200">
                                    <ScrollViewer x:Name="CollectResultScroller" VerticalScrollBarVisibility="Auto">
                                        <TextBlock x:Name="TxtCollectResults"
                                                   Foreground="#a0a0b8" FontFamily="Consolas" FontSize="11"
                                                   TextWrapping="Wrap"/>
                                    </ScrollViewer>
                                </Border>
                            </StackPanel>
                        </Border>

                    </StackPanel>
                </ScrollViewer>
            </TabItem>

            <!-- TAB 5: SETTINGS -->
            <TabItem Header="&#x2699;  Settings">
                <ScrollViewer VerticalScrollBarVisibility="Auto" Margin="0,12,0,0">
                    <StackPanel>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="Output Directory" Style="{StaticResource HeaderText}"/>
                                <TextBlock TextWrapping="Wrap" Style="{StaticResource SubText}" Margin="0,0,0,8">
                                    All index CSV files are stored here (pdf_index_raw.csv, dxf_index_raw.csv, etc.)
                                </TextBlock>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBox x:Name="TxtOutputDir" Grid.Column="0"
                                             Background="#16213e" Foreground="#e8e8e8"
                                             BorderBrush="#0f3460" Padding="8,6"
                                             FontFamily="Consolas" FontSize="11"/>
                                    <Button x:Name="BtnBrowseOutput" Grid.Column="1"
                                            Content="Browse" Style="{StaticResource SecondaryButton}"
                                            Margin="8,0,0,0"/>
                                </Grid>
                            </StackPanel>
                        </Border>
                        <Border Style="{StaticResource CardBorder}">
                            <StackPanel>
                                <TextBlock Text="File Management" Style="{StaticResource HeaderText}"/>
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                    <Button x:Name="BtnOpenOutputDir" Content="Open Index Folder"
                                            Style="{StaticResource SecondaryButton}" Margin="0,0,8,0"/>
                                    <Button x:Name="BtnResetIndex" Content="Reset All Data"
                                            Style="{StaticResource DangerButton}"/>
                                </StackPanel>
                                <TextBlock Text="Reset will delete all raw/clean index files (PDF and DXF) and the log."
                                           Style="{StaticResource SubText}"/>
                            </StackPanel>
                        </Border>
                    </StackPanel>
                </ScrollViewer>
            </TabItem>
        </TabControl>

        <!-- FOOTER -->
        <Border Grid.Row="2" Margin="0,12,0,0">
            <Grid>
                <TextBlock Text="Drawing Index Manager v3.0 - Nordic Minesteel Technologies"
                           Foreground="#6c7086" FontSize="10" HorizontalAlignment="Left"/>
                <TextBlock x:Name="TxtLastCrawl" Text="Last crawl: Never"
                           Foreground="#6c7086" FontSize="10" HorizontalAlignment="Right"/>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# ============================================================================
#  LOAD XAML & FIND CONTROLS
# ============================================================================

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

# PDF Crawler controls
$FolderCheckboxPanel = $window.FindName("FolderCheckboxPanel")
$BtnAddFolder        = $window.FindName("BtnAddFolder")
$BtnRemoveFolder     = $window.FindName("BtnRemoveFolder")
$ChkParallel         = $window.FindName("ChkParallel")
$BtnSearchPDM        = $window.FindName("BtnSearchPDM")
$BtnStartCrawl       = $window.FindName("BtnStartCrawl")
$BtnStopCrawl        = $window.FindName("BtnStopCrawl")
$TxtCrawlStatus      = $window.FindName("TxtCrawlStatus")
$TxtCrawlDetail      = $window.FindName("TxtCrawlDetail")
$CrawlProgress       = $window.FindName("CrawlProgress")
$TxtProgressDetail   = $window.FindName("TxtProgressDetail")
$StatTotalPDFs       = $window.FindName("StatTotalPDFs")
$StatUniqueParts     = $window.FindName("StatUniqueParts")
$StatDuplicates      = $window.FindName("StatDuplicates")
$StatErrors          = $window.FindName("StatErrors")
$BtnRunDedup         = $window.FindName("BtnRunDedup")
$BtnOpenClean        = $window.FindName("BtnOpenClean")
$BtnOpenRaw          = $window.FindName("BtnOpenRaw")
$TxtDedupStatus      = $window.FindName("TxtDedupStatus")

# DXF Crawler controls
$DxfFolderCheckboxPanel = $window.FindName("DxfFolderCheckboxPanel")
$BtnDxfAddFolder        = $window.FindName("BtnDxfAddFolder")
$BtnDxfRemoveFolder     = $window.FindName("BtnDxfRemoveFolder")
$ChkDxfParallel         = $window.FindName("ChkDxfParallel")
$ChkModelParallel       = $window.FindName("ChkModelParallel")
$BtnStartDxfCrawl       = $window.FindName("BtnStartDxfCrawl")
$BtnStopDxfCrawl        = $window.FindName("BtnStopDxfCrawl")
$TxtDxfCrawlStatus      = $window.FindName("TxtDxfCrawlStatus")
$TxtDxfCrawlDetail      = $window.FindName("TxtDxfCrawlDetail")
$DxfCrawlProgress       = $window.FindName("DxfCrawlProgress")
$TxtDxfProgressDetail   = $window.FindName("TxtDxfProgressDetail")
$StatTotalDXFs          = $window.FindName("StatTotalDXFs")
$StatUniqueDxfParts     = $window.FindName("StatUniqueDxfParts")
$StatDxfDuplicates      = $window.FindName("StatDxfDuplicates")
$StatDxfErrors          = $window.FindName("StatDxfErrors")
$BtnRunDxfDedup         = $window.FindName("BtnRunDxfDedup")
$BtnOpenDxfClean        = $window.FindName("BtnOpenDxfClean")
$BtnOpenDxfRaw          = $window.FindName("BtnOpenDxfRaw")
$TxtDxfDedupStatus      = $window.FindName("TxtDxfDedupStatus")

# Model Crawler controls
$ModelFolderCheckboxPanel = $window.FindName("ModelFolderCheckboxPanel")
$BtnModelAddFolder        = $window.FindName("BtnModelAddFolder")
$BtnModelRemoveFolder     = $window.FindName("BtnModelRemoveFolder")
$BtnStartModelCrawl       = $window.FindName("BtnStartModelCrawl")
$BtnStopModelCrawl        = $window.FindName("BtnStopModelCrawl")
$TxtModelCrawlStatus      = $window.FindName("TxtModelCrawlStatus")
$TxtModelCrawlDetail      = $window.FindName("TxtModelCrawlDetail")
$ModelCrawlProgress       = $window.FindName("ModelCrawlProgress")
$TxtModelProgressDetail   = $window.FindName("TxtModelProgressDetail")
$StatTotalModels          = $window.FindName("StatTotalModels")
$StatModelAssemblies      = $window.FindName("StatModelAssemblies")
$StatModelParts           = $window.FindName("StatModelParts")
$StatModelErrors          = $window.FindName("StatModelErrors")
$BtnRunModelDedup         = $window.FindName("BtnRunModelDedup")
$BtnOpenModelAll          = $window.FindName("BtnOpenModelAll")
$BtnOpenModelClean        = $window.FindName("BtnOpenModelClean")
$TxtModelDedupStatus      = $window.FindName("TxtModelDedupStatus")

# Shared controls
$TxtLog              = $window.FindName("TxtLog")
$LogScroller         = $window.FindName("LogScroller")
$BtnClearLog         = $window.FindName("BtnClearLog")
$BtnOpenLogFile      = $window.FindName("BtnOpenLogFile")
$TxtOutputDir        = $window.FindName("TxtOutputDir")
$BtnBrowseOutput     = $window.FindName("BtnBrowseOutput")
$BtnOpenOutputDir    = $window.FindName("BtnOpenOutputDir")
$BtnResetIndex       = $window.FindName("BtnResetIndex")
$TxtLastCrawl        = $window.FindName("TxtLastCrawl")

# Collect tab controls
$TabCollect              = $window.FindName("TabCollect")
$TxtBOMPath              = $window.FindName("TxtBOMPath")
$BtnBrowseBOM            = $window.FindName("BtnBrowseBOM")
$TxtCollectOutput        = $window.FindName("TxtCollectOutput")
$BtnBrowseCollectOutput  = $window.FindName("BtnBrowseCollectOutput")
$BtnCollect              = $window.FindName("BtnCollect")
$TxtCollectStatus        = $window.FindName("TxtCollectStatus")
$TxtCollectDetail        = $window.FindName("TxtCollectDetail")
$CollectProgress         = $window.FindName("CollectProgress")
$TxtCollectResults       = $window.FindName("TxtCollectResults")
$StatCollectTotal        = $window.FindName("StatCollectTotal")
$StatCollectPdfs         = $window.FindName("StatCollectPdfs")
$StatCollectDxfs         = $window.FindName("StatCollectDxfs")
$StatCollectMissing      = $window.FindName("StatCollectMissing")
$BtnOpenCollectFolder    = $window.FindName("BtnOpenCollectFolder")
$RbPdfOnly               = $window.FindName("RbPdfOnly")
$RbDxfOnly               = $window.FindName("RbDxfOnly")
$RbBoth                  = $window.FindName("RbBoth")

# Initialize
$TxtOutputDir.Text = $script:OutputDir
$script:CrawlRunning    = $false
$script:DxfCrawlRunning = $false
$script:ModelCrawlRunning = $false
$script:CancelRequested = $false
$script:ParallelWorkers = $null
$script:DxfParallelWorkers = $null
$script:ModelBgPowerShell = $null
$script:ModelBgRunspace = $null
$script:BgPool = $null
$script:DxfBgPool = $null

# Synchronized hashtables
$script:SyncHash = [hashtable]::Synchronized(@{
    Status         = ""
    Detail         = ""
    Progress       = 0
    TotalPDFs      = "--"
    Errors         = "--"
    LogQueue       = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    IsRunning      = $false
    IsComplete     = $false
    FinalMsg       = ""
    FinalTime      = ""
    FinalTotal     = 0
    FinalErrors    = 0
    ParallelTotal  = 0
    ParallelDone   = 0
    ParallelCount  = 0
})

$script:DxfSyncHash = [hashtable]::Synchronized(@{
    Status         = ""
    Detail         = ""
    Progress       = 0
    TotalDXFs      = "--"
    Errors         = "--"
    LogQueue       = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    IsRunning      = $false
    IsComplete     = $false
    FinalTime      = ""
    FinalTotal     = 0
    FinalErrors    = 0
    ParallelTotal  = 0
    ParallelDone   = 0
    ParallelCount  = 0
})

$script:ModelSyncHash = [hashtable]::Synchronized(@{
    Status         = ""
    Detail         = ""
    Progress       = 0
    TotalModels    = "--"
    Assemblies     = "--"
    Parts          = "--"
    Errors         = "--"
    LogQueue       = [System.Collections.ArrayList]::Synchronized([System.Collections.ArrayList]::new())
    IsRunning      = $false
    IsComplete     = $false
    FinalTime      = ""
    FinalTotal     = 0
    FinalAssemblies= 0
    FinalParts     = 0
    FinalErrors    = 0
})

# PDF UI Timer
$script:UITimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:UITimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:UITimer.Add_Tick({
    $sh = $script:SyncHash
    while ($sh.LogQueue.Count -gt 0) {
        $entry = $sh.LogQueue[0]; $sh.LogQueue.RemoveAt(0)
        $TxtLog.Text += "$entry`r`n"; $LogScroller.ScrollToEnd()
    }
    if ($sh.IsRunning) {
        if ($sh.Status) { $TxtCrawlDetail.Text   = $sh.Status }
        if ($sh.Detail) { $TxtProgressDetail.Text = $sh.Detail }
        $CrawlProgress.Value = $sh.Progress
        $StatTotalPDFs.Text  = $sh.TotalPDFs
        $StatErrors.Text     = $sh.Errors
    }
    if ($sh.IsComplete) {
        $sh.IsComplete = $false; $sh.IsRunning = $false
        $script:CrawlRunning = $false
        $TxtCrawlStatus.Text = "Crawl Complete"
        $TxtCrawlStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
        $TxtCrawlDetail.Text = "$($sh.FinalTotal) PDFs indexed in $($sh.FinalTime)"
        $CrawlProgress.Value = 100
        $TxtProgressDetail.Text = "Done. Run Deduplication next."
        $StatTotalPDFs.Text = $sh.FinalTotal.ToString("N0")
        $StatErrors.Text = $sh.FinalErrors.ToString("N0")
        $BtnStartCrawl.IsEnabled = $true; $BtnStopCrawl.IsEnabled = $false
        $TxtLastCrawl.Text = "Last PDF crawl: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $script:UITimer.Stop(); Update-Stats
    }
})

# DXF UI Timer
$script:DxfUITimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:DxfUITimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:DxfUITimer.Add_Tick({
    $sh = $script:DxfSyncHash
    while ($sh.LogQueue.Count -gt 0) {
        $entry = $sh.LogQueue[0]; $sh.LogQueue.RemoveAt(0)
        $TxtLog.Text += "$entry`r`n"; $LogScroller.ScrollToEnd()
    }
    if ($sh.IsRunning) {
        if ($sh.Status) { $TxtDxfCrawlDetail.Text   = $sh.Status }
        if ($sh.Detail) { $TxtDxfProgressDetail.Text = $sh.Detail }
        $DxfCrawlProgress.Value = $sh.Progress
        $StatTotalDXFs.Text     = $sh.TotalDXFs
        $StatDxfErrors.Text     = $sh.Errors
    }
    if ($sh.IsComplete) {
        $sh.IsComplete = $false; $sh.IsRunning = $false
        $script:DxfCrawlRunning = $false
        $TxtDxfCrawlStatus.Text = "DXF Crawl Complete"
        $TxtDxfCrawlStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
        $TxtDxfCrawlDetail.Text = "$($sh.FinalTotal) DXFs indexed in $($sh.FinalTime)"
        $DxfCrawlProgress.Value = 100
        $TxtDxfProgressDetail.Text = "Done. Run DXF Deduplication next."
        $StatTotalDXFs.Text = $sh.FinalTotal.ToString("N0")
        $StatDxfErrors.Text = $sh.FinalErrors.ToString("N0")
        $BtnStartDxfCrawl.IsEnabled = $true; $BtnStopDxfCrawl.IsEnabled = $false
        $TxtLastCrawl.Text = "Last DXF crawl: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $script:DxfUITimer.Stop(); Update-DxfStats
    }
})

# Model UI Timer
$script:ModelUITimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:ModelUITimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:ModelUITimer.Add_Tick({
    $sh = $script:ModelSyncHash
    while ($sh.LogQueue.Count -gt 0) {
        $entry = $sh.LogQueue[0]; $sh.LogQueue.RemoveAt(0)
        $TxtLog.Text += "$entry`r`n"; $LogScroller.ScrollToEnd()
    }
    if ($sh.IsRunning) {
        if ($sh.Status) { $TxtModelCrawlDetail.Text = $sh.Status }
        if ($sh.Detail) { $TxtModelProgressDetail.Text = $sh.Detail }
        $ModelCrawlProgress.Value = $sh.Progress
        $StatTotalModels.Text = $sh.TotalModels
        $StatModelAssemblies.Text = $sh.Assemblies
        $StatModelParts.Text = $sh.Parts
        $StatModelErrors.Text = $sh.Errors
    }
    if ($sh.IsComplete) {
        $sh.IsComplete = $false; $sh.IsRunning = $false
        $script:ModelCrawlRunning = $false
        $TxtModelCrawlStatus.Text = "Model Crawl Complete"
        $TxtModelCrawlStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
        $TxtModelCrawlDetail.Text = "$($sh.FinalTotal) models indexed in $($sh.FinalTime)"
        $ModelCrawlProgress.Value = 100
        $TxtModelProgressDetail.Text = "Done. Run Model Deduplication next."
        $StatTotalModels.Text = $sh.FinalTotal.ToString("N0")
        $StatModelAssemblies.Text = $sh.FinalAssemblies.ToString("N0")
        $StatModelParts.Text = $sh.FinalParts.ToString("N0")
        $StatModelErrors.Text = $sh.FinalErrors.ToString("N0")
        $BtnStartModelCrawl.IsEnabled = $true; $BtnStopModelCrawl.IsEnabled = $false
        $TxtLastCrawl.Text = "Last model crawl: $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $script:ModelUITimer.Stop(); Update-ModelStats
    }
})

# ============================================================================
#  HELPER FUNCTIONS
# ============================================================================

function Add-LogEntry {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "HH:mm:ss"
    $prefix = switch ($Level) {
        "ERROR"   { "[!]" }; "WARN" { "[~]" }; "SUCCESS" { "[+]" }; default { "[ ]" }
    }
    $entry = "$timestamp $prefix $Message"
    $window.Dispatcher.Invoke([Action]{ $TxtLog.Text += "$entry`r`n"; $LogScroller.ScrollToEnd() })
    try { Add-Content -Path $script:LogFile -Value $entry -ErrorAction SilentlyContinue } catch {}
}

function Update-Stats {
    $window.Dispatcher.Invoke([Action]{
        if (Test-Path $script:RawCSV) {
            $rawCount = (Get-Content $script:RawCSV -ErrorAction SilentlyContinue | Measure-Object).Count - 1
            $StatTotalPDFs.Text = [math]::Max(0,$rawCount).ToString("N0")
        } else { $StatTotalPDFs.Text = "--" }
        if (Test-Path $script:CleanCSV) {
            $cleanData = Import-Csv $script:CleanCSV -ErrorAction SilentlyContinue
            if ($cleanData) {
                $StatUniqueParts.Text = $cleanData.Count.ToString("N0")
                $totalCandidates = ($cleanData | Measure-Object -Property Candidates -Sum).Sum
                $StatDuplicates.Text = ($totalCandidates - $cleanData.Count).ToString("N0")
            }
        } else { $StatUniqueParts.Text = "--"; $StatDuplicates.Text = "--" }
        if (Test-Path $script:StateFile) {
            try { $state = Get-Content $script:StateFile -Raw | ConvertFrom-Json; $TxtLastCrawl.Text = "Last PDF crawl: $($state.LastCrawlTime)" } catch {}
        }
    })
}

function Update-DxfStats {
    $window.Dispatcher.Invoke([Action]{
        if (Test-Path $script:DxfRawCSV) {
            $rawCount = (Get-Content $script:DxfRawCSV -ErrorAction SilentlyContinue | Measure-Object).Count - 1
            $StatTotalDXFs.Text = [math]::Max(0,$rawCount).ToString("N0")
        } else { $StatTotalDXFs.Text = "--" }
        if (Test-Path $script:DxfCleanCSV) {
            $cleanData = Import-Csv $script:DxfCleanCSV -ErrorAction SilentlyContinue
            if ($cleanData) {
                $StatUniqueDxfParts.Text = $cleanData.Count.ToString("N0")
                $totalCandidates = ($cleanData | Measure-Object -Property Candidates -Sum).Sum
                $StatDxfDuplicates.Text = ($totalCandidates - $cleanData.Count).ToString("N0")
            }
        } else { $StatUniqueDxfParts.Text = "--"; $StatDxfDuplicates.Text = "--" }
    })
}

function Update-ModelStats {
    $window.Dispatcher.Invoke([Action]{
        if (Test-Path $script:ModelAllCSV) {
            $rawCount = (Get-Content $script:ModelAllCSV -ErrorAction SilentlyContinue | Measure-Object).Count - 1
            $StatTotalModels.Text = [math]::Max(0, $rawCount).ToString("N0")
        } else { $StatTotalModels.Text = "--" }

        if (Test-Path $script:ModelStateFile) {
            try {
                $state = Get-Content $script:ModelStateFile -Raw | ConvertFrom-Json
                if ($state.TotalAssemblies -ne $null) { $StatModelAssemblies.Text = [int]$state.TotalAssemblies }
                else { $StatModelAssemblies.Text = "--" }
                if ($state.TotalParts -ne $null) { $StatModelParts.Text = [int]$state.TotalParts }
                else { $StatModelParts.Text = "--" }
                if ($state.TotalErrors -ne $null) { $StatModelErrors.Text = [int]$state.TotalErrors }
                else { $StatModelErrors.Text = "--" }
            } catch {
                $StatModelAssemblies.Text = "--"
                $StatModelParts.Text = "--"
                $StatModelErrors.Text = "--"
            }
        } else {
            $StatModelAssemblies.Text = "--"
            $StatModelParts.Text = "--"
            $StatModelErrors.Text = "--"
        }
    })
}

function Get-PartNumberInfo {
    param([string]$BaseName)
    if ($BaseName -match '^(.+?)[\s_\-]?[Rr][Ee][Vv][\.\-_]?([A-Za-z0-9]+)$') {
        $basePart = $Matches[1].Trim()
        $revRaw   = $Matches[2].Trim().ToUpper()
        $revValue = Get-RevisionValue -RevString $revRaw
        return @{ BasePart = $basePart.ToUpper(); RevRaw = $revRaw; RevValue = $revValue; HasRev = $true }
    } else {
        return @{ BasePart = $BaseName.Trim().ToUpper(); RevRaw = $null; RevValue = 0; HasRev = $false }
    }
}

function Get-RevisionValue {
    param([string]$RevString)
    if ($RevString -match '^\d+$') { return [int]$RevString }
    if ($RevString -match '^[A-Z]$') { return [int][char]$RevString - 64 }
    if ($RevString -match '^[A-Z]+$') {
        $value = 0
        for ($i = 0; $i -lt $RevString.Length; $i++) { $value = $value * 26 + ([int][char]$RevString[$i] - 64) }
        return $value
    }
    if ($RevString -match '([A-Z]+)(\d+)') {
        $lv = 0; $lp = $Matches[1]
        for ($i = 0; $i -lt $lp.Length; $i++) { $lv = $lv * 26 + ([int][char]$lp[$i] - 64) }
        return ($lv * 1000) + [int]$Matches[2]
    }
    return $RevString.GetHashCode()
}

# ============================================================================
#  Test if a file path is inside an Obsolete/Archive/Old/Deprecated folder
# ============================================================================
function Test-ObsoletePath {
    param([string]$FilePath)
    return ($FilePath -match '(?i)\\(Obsolete|Archive|Old|Deprecated)\\')
}

# ============================================================================
#  GENERIC CRAWL FUNCTION - used for both PDF and DXF
#  v2.0 - Optimized parallel with dynamic folder expansion, RunspacePool,
#          buffered I/O, and thread-safe atomic counters
# ============================================================================
function Start-FileCrawl {
    param(
        [string[]]   $SelectedFolders,
        [string]     $RawCSV,
        [string]     $StateFile,
        [string]     $FileExt,       # "*.pdf" or "*.dxf"
        [hashtable]  $SyncHash,
        [object]     $StartBtn,
        [object]     $StopBtn,
        [object]     $StatusTxt,
        [bool]       $UseParallel,
        [ref]        $WorkersRef,
        [ref]        $PoolRef,
        [object]     $UITimer
    )

    $SyncHash.ParallelTotal = 0
    $SyncHash.ParallelDone  = 0
    $SyncHash.ParallelCount = 0
    $SyncHash.FinalErrors   = 0
    $SyncHash.IsRunning     = $true
    $SyncHash.IsComplete    = $false

    $logFile = $script:LogFile

    if (-not $UseParallel) {
        # ---- SEQUENTIAL MODE (optimized with buffered writes) ----
        $runspace = [RunspaceFactory]::CreateRunspace()
        $runspace.Open()
        $runspace.SessionStateProxy.SetVariable("SyncHash",        $SyncHash)
        $runspace.SessionStateProxy.SetVariable("selectedFolders", $SelectedFolders)
        $runspace.SessionStateProxy.SetVariable("RawCSV",          $RawCSV)
        $runspace.SessionStateProxy.SetVariable("LogFile",         $logFile)
        $runspace.SessionStateProxy.SetVariable("StateFile",       $StateFile)
        $runspace.SessionStateProxy.SetVariable("FileExt",         $FileExt)

        $ps = [PowerShell]::Create(); $ps.Runspace = $runspace

        [void]$ps.AddScript({
            $sh = $SyncHash
            function BG-Log {
                param([string]$Msg,[string]$Level="INFO")
                $ts = Get-Date -Format "HH:mm:ss"
                $pfx = switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}}
                $entry = "$ts $pfx $Msg"
                $sh.LogQueue.Add($entry)|Out-Null
                try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}
            }
            $sh.IsRunning = $true
            $appendMode = Test-Path $RawCSV
            if ($appendMode) {
                try {
                    $f=[System.IO.File]::Open($RawCSV,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite)
                    if($f.Length-gt 0){$f.Seek(-1,[System.IO.SeekOrigin]::End)|Out-Null;$lb=$f.ReadByte();if($lb-ne 10){$f.Seek(0,[System.IO.SeekOrigin]::End)|Out-Null;$nl=[System.Text.Encoding]::UTF8.GetBytes([Environment]::NewLine);$f.Write($nl,0,$nl.Length)}}
                    $f.Close();$f.Dispose()
                } catch {}
            } else {
                [System.IO.File]::WriteAllText($RawCSV,'"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder"'+[Environment]::NewLine)
                BG-Log "Created new raw index" "INFO"
            }
            $stream=$null
            try {
                $stream=[System.IO.StreamWriter]::new($RawCSV,$true,[System.Text.Encoding]::UTF8,65536)
                $totalFiles=0;$totalErrors=0
                $sw=[System.Diagnostics.Stopwatch]::StartNew()
                $folderIdx=0;$lastUpdate=[DateTime]::MinValue
                $buf=[System.Text.StringBuilder]::new(8192)
                $bufCount=0
                foreach($root in $selectedFolders){
                    $folderIdx++
                    $pctBase=[math]::Floor(($folderIdx-1)/$selectedFolders.Count*100)
                    if(-not(Test-Path $root)){BG-Log "WARNING: Root not found: $root" "WARN";continue}
                    BG-Log "Scanning [$folderIdx/$($selectedFolders.Count)]: $root" "INFO"
                    $sh.Status="Scanning: $root";$sh.Progress=$pctBase;$rootCount=0
                    $rootEsc=$root.Replace('"','""')
                    $stack=[System.Collections.Generic.Stack[string]]::new();$stack.Push($root)
                    while($stack.Count-gt 0){
                        $cur=$stack.Pop()
                        if($sh.IsRunning-eq $false){break}
                        try {
                            foreach($fp in [System.IO.Directory]::EnumerateFiles($cur,$FileExt)){
                                try{
                                    $fn=[System.IO.Path]::GetFileName($fp)
                                    $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                    $fi=[System.IO.FileInfo]::new($fp)
                                    if($fi.Length-eq 0){continue}
                                    $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                    [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).AppendLine('"')
                                    $rootCount++;$totalFiles++;$bufCount++
                                    if($bufCount-ge 500){$stream.Write($buf.ToString());$buf.Clear();$bufCount=0}
                                    $now=[DateTime]::Now
                                    if(($now-$lastUpdate).TotalSeconds-ge 1){$lastUpdate=$now;$elapsed=$sw.Elapsed.ToString('hh\:mm\:ss');$sh.Status="Scanning: $root";$sh.Detail="$totalFiles files indexed | $elapsed elapsed";$sh.TotalDXFs=$totalFiles.ToString("N0");$sh.TotalPDFs=$totalFiles.ToString("N0");$sh.Errors=$totalErrors.ToString("N0")}
                                }catch{$totalErrors++}
                            }
                            foreach($sd in [System.IO.Directory]::EnumerateDirectories($cur)){$stack.Push($sd)}
                        }catch{$totalErrors++;if($_.Exception.Message-notmatch "denied"){BG-Log "Error in $cur : $($_.Exception.Message)" "ERROR"}}
                    }
                    if($bufCount-gt 0){$stream.Write($buf.ToString());$buf.Clear();$bufCount=0}
                    BG-Log "  Finished $root : $rootCount files" "SUCCESS"
                    $sh.Progress=[math]::Floor($folderIdx/$selectedFolders.Count*100)
                }
            } finally {if($bufCount-gt 0-and $stream){$stream.Write($buf.ToString())};if($stream){$stream.Flush();$stream.Close();$stream.Dispose()}}
            $sw.Stop();$elapsed=$sw.Elapsed.ToString('hh\:mm\:ss')
            $state=@{LastCrawlTime=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");TotalFiles=$totalFiles;TotalErrors=$totalErrors;ElapsedTime=$elapsed;FoldersCrawled=$selectedFolders}|ConvertTo-Json
            [System.IO.File]::WriteAllText($StateFile,$state)
            BG-Log "CRAWL COMPLETE: $totalFiles files, $totalErrors errors, $elapsed" "SUCCESS"
            $sh.FinalTotal=$totalFiles;$sh.FinalErrors=$totalErrors;$sh.FinalTime=$elapsed;$sh.IsComplete=$true
        })

        $UITimer.Start()
        $handle = $ps.BeginInvoke()
        $script:BgPowerShell = $ps; $script:BgHandle = $handle; $script:BgRunspace = $runspace

    } else {
        # ---- PARALLEL MODE v2.0 ----
        # Dynamic recursive folder expansion for balanced workload distribution.
        # Aggressively splits large roots (like J:\) into many small scan units
        # to maximize parallelism across all CPU cores.

        # Step 1: Expand folders dynamically to create many balanced scan units
        $scanUnits = [System.Collections.Generic.List[string]]::new()
        $expandQueue = [System.Collections.Generic.Queue[pscustomobject]]::new()

        # Seed the queue with all selected folders
        foreach ($root in $SelectedFolders) {
            if (Test-Path $root) {
                $expandQueue.Enqueue([pscustomobject]@{ Path=$root; Depth=0; OrigRoot=$root })
            }
        }

        # Target: at least 4x logical CPUs worth of units for good load balancing
        $cpuCount = [Environment]::ProcessorCount
        $maxWorkers = [math]::Max(4, $cpuCount)
        $targetUnits = $maxWorkers * 4   # e.g. 16 cores -> 64 units

        # Expansion loop: keep splitting folders until we have enough units.
        #
        # The unit list is a SET of non-overlapping folders. Each worker recurses
        # fully from its assigned root. So:
        #   - If a folder is NOT expanded, add it as a unit (worker covers everything below it).
        #   - If a folder IS expanded into children, DON'T add the parent — but DO add a
        #     "shallow" entry for it marked as files-only=true, so that files sitting
        #     directly in the parent (not in any subfolder) are still picked up.
        #   - Drive roots (J:\) never have files directly in them, so no shallow unit needed.
        #
        # scanUnits stores: [pscustomobject]@{ Path=...; ShallowOnly=$false/$true }
        # ShallowOnly=true  -> worker scans files in that folder only (no subdirs)
        # ShallowOnly=false -> worker recurses fully

        $scanUnits = [System.Collections.Generic.List[pscustomobject]]::new()
        $maxExpandDepth = 4

        while ($expandQueue.Count -gt 0) {
            $item = $expandQueue.Dequeue()
            $shouldExpand = $false

            if ($item.Depth -lt $maxExpandDepth) {
                $norm = $item.Path.TrimEnd('\') + '\'
                if ($norm -match '^[A-Za-z]:\\$') {
                    $shouldExpand = $true   # always expand bare drive roots
                } elseif (($scanUnits.Count + $expandQueue.Count) -lt $targetUnits) {
                    $shouldExpand = $true   # expand to reach target unit count
                }
            }

            if ($shouldExpand) {
                try {
                    $subs = [System.IO.Directory]::GetDirectories($item.Path)
                    if ($subs.Count -gt 0) {
                        foreach ($s in $subs) {
                            $expandQueue.Enqueue([pscustomobject]@{ Path=$s; Depth=($item.Depth+1); OrigRoot=$item.OrigRoot })
                        }
                        # Add this folder as a SHALLOW unit (files directly in it only).
                        # Skip for bare drive roots — they never have loose files.
                        $norm = $item.Path.TrimEnd('\') + '\'
                        if ($norm -notmatch '^[A-Za-z]:\\$') {
                            $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$true })
                        }
                    } else {
                        # No subfolders — add as full recursive unit
                        $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false })
                    }
                } catch {
                    $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false })
                }
            } else {
                # Not expanding — add as full recursive unit
                $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false })
            }
        }

        $unitCount  = $scanUnits.Count
        $shallowCnt = ($scanUnits | Where-Object { $_.ShallowOnly }).Count
        Add-LogEntry "Parallel crawl ($FileExt): $unitCount scan units ($shallowCnt shallow) across $maxWorkers max workers" "INFO"
        $SyncHash.ParallelCount = $unitCount

        $tempDir = Join-Path $script:OutputDir "crawl_tmp_$($FileExt -replace '\*\.','')"
        if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
        Get-ChildItem -Path $tempDir -Filter "chunk_*.csv" -EA SilentlyContinue | Remove-Item -Force -EA SilentlyContinue

        # Step 2: Create a RunspacePool instead of individual runspaces
        # This caps actual concurrent threads and reuses them efficiently.
        $pool = [RunspaceFactory]::CreateRunspacePool(1, $maxWorkers)
        $pool.Open()

        # Thread-safe counter using a locked integer array [total, errors, done]
        $counters = [hashtable]::Synchronized(@{ Total=0; Errors=0; Done=0 })

        $workers = [System.Collections.Generic.List[object]]::new()
        $WorkersRef.Value = $workers
        $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

        # Worker script block — shared by all units via the pool
        # ShallowOnly=true  -> only scan files directly in ScanRoot (no subdir recursion)
        # ShallowOnly=false -> full recursive scan from ScanRoot downward
        $workerScript = {
            param($SyncHash, $ScanRoot, $ShallowOnly, $ChunkFile, $LogFile, $FileExt, $Counters)
            $sh = $SyncHash
            function BG-Log {param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
            $mode=if($ShallowOnly){"[Shallow]"}else{"[Recurse]"}
            BG-Log "  [Worker] $mode $ScanRoot" "INFO"
            $localCount=0;$localErrors=0
            $stream=$null
            $rootEsc=$ScanRoot.Replace('"','""')
            $buf=[System.Text.StringBuilder]::new(16384)
            $bufLines=0
            try {
                $stream=[System.IO.StreamWriter]::new($ChunkFile,$false,[System.Text.Encoding]::UTF8,65536)
                if ($ShallowOnly) {
                    # Scan only files directly in this folder — no subdirectory recursion
                    try {
                        foreach($fp in [System.IO.Directory]::EnumerateFiles($ScanRoot,$FileExt)){
                            try{
                                $fn=[System.IO.Path]::GetFileName($fp)
                                $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                $fi=[System.IO.FileInfo]::new($fp)
                                if($fi.Length-eq 0){continue}
                                $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).AppendLine('"')
                                $localCount++;$bufLines++
                                if($bufLines-ge 1000){$stream.Write($buf.ToString());$buf.Clear();$bufLines=0}
                            }catch{$localErrors++}
                        }
                    }catch{$localErrors++;if($_.Exception.Message-notmatch "denied"){BG-Log "Error scanning $ScanRoot : $($_.Exception.Message)" "ERROR"}}
                } else {
                    # Full recursive scan
                    $stack=[System.Collections.Generic.Stack[string]]::new();$stack.Push($ScanRoot)
                    while($stack.Count-gt 0){
                        $cur=$stack.Pop()
                        if($sh.IsRunning-eq $false){break}
                        try {
                            foreach($fp in [System.IO.Directory]::EnumerateFiles($cur,$FileExt)){
                                try{
                                    $fn=[System.IO.Path]::GetFileName($fp)
                                    $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                    $fi=[System.IO.FileInfo]::new($fp)
                                    if($fi.Length-eq 0){continue}
                                    $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                    [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).AppendLine('"')
                                    $localCount++;$bufLines++
                                    if($bufLines-ge 1000){$stream.Write($buf.ToString());$buf.Clear();$bufLines=0}
                                }catch{$localErrors++}
                            }
                            foreach($sd in [System.IO.Directory]::EnumerateDirectories($cur)){$stack.Push($sd)}
                        }catch{$localErrors++;if($_.Exception.Message-notmatch "denied"){BG-Log "Error in $cur : $($_.Exception.Message)" "ERROR"}}
                    }
                }
            } finally {
                if($bufLines-gt 0-and $stream){$stream.Write($buf.ToString())}
                if($stream){$stream.Flush();$stream.Close();$stream.Dispose()}
            }
            BG-Log "  [Worker] Done: $ScanRoot - $localCount files" "SUCCESS"

            # Thread-safe counter updates using lock
            [System.Threading.Monitor]::Enter($Counters)
            try {
                $Counters.Total += $localCount
                $Counters.Errors += $localErrors
                $Counters.Done += 1
                $sh.ParallelTotal = $Counters.Total
                $sh.FinalErrors = $Counters.Errors
                $sh.ParallelDone = $Counters.Done
            } finally { [System.Threading.Monitor]::Exit($Counters) }

            $done=$sh.ParallelDone;$total=$sh.ParallelCount
            $sh.TotalDXFs=$sh.ParallelTotal.ToString("N0");$sh.TotalPDFs=$sh.ParallelTotal.ToString("N0")
            $sh.Progress=[math]::Floor($done/$total*100)
            $sh.Detail="$($sh.ParallelTotal.ToString('N0')) files found across $done/$total units"
            $sh.Errors=$sh.FinalErrors.ToString("N0")
        }

        # Step 3: Launch all workers into the pool
        foreach ($unit in $scanUnits) {
            $safeUnitName = ($unit.Path -replace '[\\:/\s]','_')
            if ($safeUnitName.Length -gt 80) { $safeUnitName = $safeUnitName.Substring(0,80) }
            $suffix = if ($unit.ShallowOnly) { "_SH" } else { "" }
            $chunkFile = Join-Path $tempDir ("chunk_" + $safeUnitName + $suffix + ".csv")

            $wPs = [PowerShell]::Create()
            $wPs.RunspacePool = $pool
            [void]$wPs.AddScript($workerScript)
            [void]$wPs.AddArgument($SyncHash)
            [void]$wPs.AddArgument($unit.Path)
            [void]$wPs.AddArgument($unit.ShallowOnly)
            [void]$wPs.AddArgument($chunkFile)
            [void]$wPs.AddArgument($logFile)
            [void]$wPs.AddArgument($FileExt)
            [void]$wPs.AddArgument($counters)

            $wHandle = $wPs.BeginInvoke()
            $workers.Add([pscustomobject]@{ PS=$wPs; Handle=$wHandle; Chunk=$chunkFile; Root=$unit.Path; Shallow=$unit.ShallowOnly })
        }

        # Step 4: Coordinator runspace — waits for pool workers, merges chunks
        $coordRS = [RunspaceFactory]::CreateRunspace(); $coordRS.Open()
        $coordRS.SessionStateProxy.SetVariable("SyncHash",        $SyncHash)
        $coordRS.SessionStateProxy.SetVariable("Workers",         $workers)
        $coordRS.SessionStateProxy.SetVariable("RawCSV",          $RawCSV)
        $coordRS.SessionStateProxy.SetVariable("LogFile",         $logFile)
        $coordRS.SessionStateProxy.SetVariable("StateFile",       $StateFile)
        $coordRS.SessionStateProxy.SetVariable("selectedFolders", $SelectedFolders)
        $coordRS.SessionStateProxy.SetVariable("Stopwatch",       $stopwatch)
        $coordRS.SessionStateProxy.SetVariable("TempDir",         $tempDir)
        $coordRS.SessionStateProxy.SetVariable("Pool",            $pool)

        $coordPs = [PowerShell]::Create(); $coordPs.Runspace = $coordRS
        [void]$coordPs.AddScript({
            $sh = $SyncHash
            function BG-Log {param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
            BG-Log "Coordinator: waiting for $($Workers.Count) workers (RunspacePool)..." "INFO"

            # Wait for all workers with a per-worker timeout.
            # If a worker hangs (e.g. frozen network path), we stop it and move on
            # so the crawl always completes rather than blocking forever.
            $workerTimeoutMs = 120000  # 2 minutes max per unit before forcing skip
            $timedOut = 0
            foreach($w in $Workers){
                $finished = $w.Handle.AsyncWaitHandle.WaitOne($workerTimeoutMs)
                if($finished){
                    try{$w.PS.EndInvoke($w.Handle)}catch{}
                } else {
                    $timedOut++
                    BG-Log "  [Worker] TIMEOUT on: $($w.Root) - skipping" "WARN"
                    try{$w.PS.Stop()}catch{}
                }
                try{$w.PS.Dispose()}catch{}
            }
            if($timedOut -gt 0){ BG-Log "$timedOut worker(s) timed out and were skipped" "WARN" }

            # Close the RunspacePool
            try { $Pool.Close(); $Pool.Dispose() } catch {}

            BG-Log "All $($Workers.Count) workers done. Merging chunks..." "INFO"
            $sh.Status = "Merging results..."

            # Prepare output CSV
            $appendMode = Test-Path $RawCSV
            if($appendMode){
                try{$f=[System.IO.File]::Open($RawCSV,[System.IO.FileMode]::Open,[System.IO.FileAccess]::ReadWrite);if($f.Length-gt 0){$f.Seek(-1,[System.IO.SeekOrigin]::End)|Out-Null;if($f.ReadByte()-ne 10){$f.Seek(0,[System.IO.SeekOrigin]::End)|Out-Null;$nl=[System.Text.Encoding]::UTF8.GetBytes([Environment]::NewLine);$f.Write($nl,0,$nl.Length)}};$f.Close();$f.Dispose()}catch{}
            } else {
                [System.IO.File]::WriteAllText($RawCSV,'"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder"'+[Environment]::NewLine)
            }

            # Merge all chunk files with buffered I/O
            $ms=$null;$totalMerged=0
            try{
                $ms=[System.IO.StreamWriter]::new($RawCSV,$true,[System.Text.Encoding]::UTF8,65536)
                foreach($w in $Workers){
                    if(Test-Path $w.Chunk){
                        $reader=$null
                        try {
                            $reader=[System.IO.StreamReader]::new($w.Chunk,[System.Text.Encoding]::UTF8,$true,65536)
                            while(-not $reader.EndOfStream){
                                $line=$reader.ReadLine()
                                if($line){$ms.WriteLine($line);$totalMerged++}
                            }
                        } finally { if($reader){$reader.Close();$reader.Dispose()} }
                        Remove-Item $w.Chunk -Force -EA SilentlyContinue
                    }
                }
            }finally{if($ms){$ms.Flush();$ms.Close();$ms.Dispose()}}

            BG-Log "Merged $totalMerged lines into $RawCSV" "INFO"
            try{if((Get-ChildItem $TempDir -EA SilentlyContinue|Measure-Object).Count-eq 0){Remove-Item $TempDir -Force -EA SilentlyContinue}}catch{}
            $Stopwatch.Stop();$elapsed=$Stopwatch.Elapsed.ToString('hh\:mm\:ss')
            $totalFiles=$sh.ParallelTotal;$totalErrors=$sh.FinalErrors
            $state=@{LastCrawlTime=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");TotalFiles=[int]$totalFiles;TotalErrors=[int]$totalErrors;ElapsedTime=$elapsed;FoldersCrawled=$selectedFolders;ScanUnits=$Workers.Count}|ConvertTo-Json
            [System.IO.File]::WriteAllText($StateFile,$state)
            BG-Log "CRAWL COMPLETE (PARALLEL): $totalFiles files, $totalErrors errors, $elapsed ($($Workers.Count) units)" "SUCCESS"
            $sh.FinalTotal=$totalFiles;$sh.FinalTime=$elapsed;$sh.IsComplete=$true
        })

        $UITimer.Start()
        $coordPs.BeginInvoke() | Out-Null
        $script:BgPowerShell = $coordPs; $script:BgRunspace = $coordRS
        $PoolRef.Value = $pool
    }
}

# ============================================================================
#  MODEL CRAWL FUNCTION (SLDASM + SLDPRT) - Parallel RunspacePool v2.0
# ============================================================================
function Start-ModelCrawl {
    param(
        [string[]]  $SelectedFolders,
        [string]    $RawCSV,
        [string]    $AllCSV,
        [string]    $StateFile,
        [hashtable] $SyncHash,
        [object]    $UITimer,
        [bool]      $UseParallel = $true
    )

    $SyncHash.IsRunning     = $true
    $SyncHash.IsComplete    = $false
    $SyncHash.Progress      = 0
    $SyncHash.TotalModels   = "0"
    $SyncHash.Assemblies    = "0"
    $SyncHash.Parts         = "0"
    $SyncHash.Errors        = "0"
    $SyncHash.ParallelTotal = 0
    $SyncHash.ParallelDone  = 0
    $SyncHash.ParallelCount = 0
    $SyncHash.FinalErrors   = 0

    $logFile = $script:LogFile

    if (-not $UseParallel) {
        # ---- SEQUENTIAL MODE (single background runspace) ----
        $runspace = [RunspaceFactory]::CreateRunspace(); $runspace.Open()
        $runspace.SessionStateProxy.SetVariable("SyncHash",        $SyncHash)
        $runspace.SessionStateProxy.SetVariable("selectedFolders", $SelectedFolders)
        $runspace.SessionStateProxy.SetVariable("RawCSV",          $RawCSV)
        $runspace.SessionStateProxy.SetVariable("AllCSV",          $AllCSV)
        $runspace.SessionStateProxy.SetVariable("StateFile",       $StateFile)
        $runspace.SessionStateProxy.SetVariable("LogFile",         $logFile)
        $ps = [PowerShell]::Create(); $ps.Runspace = $runspace
        [void]$ps.AddScript({
            $sh = $SyncHash
            function BG-Log {param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
            $patterns=@("*.sldasm","*.sldprt")
            [System.IO.File]::WriteAllText($RawCSV,'"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder","FileType"'+[Environment]::NewLine)
            $totalFiles=0;$assyCount=0;$partCount=0;$totalErrors=0
            $sw=[System.Diagnostics.Stopwatch]::StartNew();$stream=$null
            $buf=[System.Text.StringBuilder]::new(8192);$bufCount=0;$lastUpdate=[DateTime]::MinValue
            try {
                $stream=[System.IO.StreamWriter]::new($RawCSV,$true,[System.Text.Encoding]::UTF8,65536)
                $folderIdx=0
                foreach($root in $selectedFolders){
                    if($sh.IsRunning-eq $false){break}
                    $folderIdx++
                    if(-not(Test-Path $root)){BG-Log "WARNING: Root not found: $root" "WARN";continue}
                    BG-Log "Scanning models [$folderIdx/$($selectedFolders.Count)]: $root" "INFO"
                    $rootEsc=$root.Replace('"','""');$rootCount=0
                    $stack=[System.Collections.Generic.Stack[string]]::new();$stack.Push($root)
                    while($stack.Count-gt 0){
                        if($sh.IsRunning-eq $false){break};$cur=$stack.Pop()
                        try {
                            foreach($pat in $patterns){
                                foreach($fp in [System.IO.Directory]::EnumerateFiles($cur,$pat)){
                                    if($sh.IsRunning-eq $false){break}
                                    try{
                                        $fi=[System.IO.FileInfo]::new($fp);if($fi.Length-eq 0){continue}
                                        $fn=[System.IO.Path]::GetFileName($fp);$bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                        $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                        $ft=[System.IO.Path]::GetExtension($fp).TrimStart('.').ToUpperInvariant()
                                        [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).Append('","').Append($ft).AppendLine('"')
                                        $rootCount++;$totalFiles++;$bufCount++
                                        if($ft-eq"SLDASM"){$assyCount++}elseif($ft-eq"SLDPRT"){$partCount++}
                                        if($bufCount-ge 500){$stream.Write($buf.ToString());$buf.Clear();$bufCount=0}
                                        $now=[DateTime]::Now;if(($now-$lastUpdate).TotalSeconds-ge 1){$lastUpdate=$now;$sh.Status="Scanning: $root";$sh.Detail="$totalFiles models | Assemblies: $assyCount | Parts: $partCount | Elapsed: $($sw.Elapsed.ToString('hh\:mm\:ss'))";$sh.TotalModels=$totalFiles.ToString("N0");$sh.Assemblies=$assyCount.ToString("N0");$sh.Parts=$partCount.ToString("N0");$sh.Errors=$totalErrors.ToString("N0")}
                                    }catch{$totalErrors++}
                                }
                            }
                            foreach($sd in [System.IO.Directory]::EnumerateDirectories($cur)){$stack.Push($sd)}
                        }catch{$totalErrors++;if($_.Exception.Message-notmatch"denied"){BG-Log "Error in $cur : $($_.Exception.Message)" "ERROR"}}
                    }
                    if($bufCount-gt 0){$stream.Write($buf.ToString());$buf.Clear();$bufCount=0}
                    BG-Log "  Finished $root : $rootCount model files" "SUCCESS"
                    $sh.Progress=[math]::Floor($folderIdx/$selectedFolders.Count*100)
                }
            } finally {if($bufCount-gt 0-and $stream){$stream.Write($buf.ToString())};if($stream){$stream.Flush();$stream.Close();$stream.Dispose()}}
            $sw.Stop();$elapsed=$sw.Elapsed.ToString('hh\:mm\:ss')
            try{Copy-Item -Path $RawCSV -Destination $AllCSV -Force}catch{BG-Log "Failed copy to AllCSV: $($_.Exception.Message)" "ERROR"}
            try{$state=@{LastCrawlTime=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");TotalFiles=$totalFiles;TotalAssemblies=$assyCount;TotalParts=$partCount;TotalErrors=$totalErrors;ElapsedTime=$elapsed;FoldersCrawled=$selectedFolders;Patterns=@("*.sldasm","*.sldprt")}|ConvertTo-Json;[System.IO.File]::WriteAllText($StateFile,$state)}catch{BG-Log "Failed writing model state: $($_.Exception.Message)" "ERROR"}
            BG-Log "MODEL CRAWL COMPLETE: $totalFiles files ($assyCount assemblies, $partCount parts), $totalErrors errors, $elapsed" "SUCCESS"
            $sh.FinalTotal=$totalFiles;$sh.FinalAssemblies=$assyCount;$sh.FinalParts=$partCount;$sh.FinalErrors=$totalErrors;$sh.FinalTime=$elapsed;$sh.IsComplete=$true
        })
        $UITimer.Start()
        $handle = $ps.BeginInvoke()
        $script:ModelBgPowerShell = $ps; $script:ModelBgRunspace = $runspace; $script:ModelBgHandle = $handle
        return
    }

    # --- Step 1: Dynamic folder expansion (same as PDF parallel v2.0) ---
    $scanUnits   = [System.Collections.Generic.List[pscustomobject]]::new()
    $expandQueue = [System.Collections.Generic.Queue[pscustomobject]]::new()
    foreach ($root in $SelectedFolders) {
        if (Test-Path $root) { $expandQueue.Enqueue([pscustomobject]@{ Path=$root; Depth=0 }) }
    }
    $cpuCount       = [Environment]::ProcessorCount
    $maxWorkers     = [math]::Max(4, $cpuCount)
    $targetUnits    = $maxWorkers * 4
    $maxExpandDepth = 4

    while ($expandQueue.Count -gt 0) {
        $item = $expandQueue.Dequeue()
        $shouldExpand = $false
        if ($item.Depth -lt $maxExpandDepth) {
            $norm = $item.Path.TrimEnd('\') + '\'
            if    ($norm -match '^[A-Za-z]:\\$') { $shouldExpand = $true }
            elseif ($item.Depth -lt 2)            { $shouldExpand = $true }  # always expand first 2 levels regardless of unit count
            elseif (($scanUnits.Count + $expandQueue.Count) -lt $targetUnits) { $shouldExpand = $true }
        }
        if ($shouldExpand) {
            try {
                $subs = [System.IO.Directory]::GetDirectories($item.Path)
                if ($subs.Count -gt 0) {
                    foreach ($s in $subs) { $expandQueue.Enqueue([pscustomobject]@{ Path=$s; Depth=($item.Depth+1) }) }
                    $norm = $item.Path.TrimEnd('\') + '\'
                    if ($norm -notmatch '^[A-Za-z]:\\$') { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$true }) }
                } else { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
            } catch { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
        } else { $scanUnits.Add([pscustomobject]@{ Path=$item.Path; ShallowOnly=$false }) }
    }

    $unitCount = $scanUnits.Count
    Add-LogEntry "Parallel model crawl: $unitCount scan units across $maxWorkers max workers" "INFO"
    $SyncHash.ParallelCount = $unitCount

    $tempDir = Join-Path $script:OutputDir "crawl_tmp_model"
    if (-not (Test-Path $tempDir)) { New-Item -ItemType Directory -Path $tempDir -Force | Out-Null }
    Get-ChildItem -Path $tempDir -Filter "chunk_*.csv" -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

    # Write CSV header; workers write headerless chunk files; coordinator appends them
    [System.IO.File]::WriteAllText($RawCSV, '"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder","FileType"' + [Environment]::NewLine)

    # --- Step 2: RunspacePool ---
    $pool     = [RunspaceFactory]::CreateRunspacePool(1, $maxWorkers)
    $pool.Open()
    $counters  = [hashtable]::Synchronized(@{ Total=0; Errors=0; Done=0; Assemblies=0; Parts=0 })
    $workers   = [System.Collections.Generic.List[object]]::new()
    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $patterns  = @("*.sldasm", "*.sldprt")

    # Worker script - scans both SLDASM + SLDPRT, writes FileType column
    $workerScript = {
        param($SyncHash, $ScanRoot, $ShallowOnly, $ChunkFile, $LogFile, $Counters, $Patterns)
        $sh = $SyncHash
        function BG-Log {param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
        $mode = if($ShallowOnly){"[Shallow]"}else{"[Recurse]"}
        BG-Log "  [Worker] $mode $ScanRoot" "INFO"
        $localCount=0; $localAssy=0; $localParts=0; $localErrors=0
        $rootEsc=$ScanRoot.Replace('"','""')
        $buf=[System.Text.StringBuilder]::new(16384); $bufLines=0; $stream=$null
        try {
            $stream=[System.IO.StreamWriter]::new($ChunkFile,$false,[System.Text.Encoding]::UTF8,65536)
            if ($ShallowOnly) {
                foreach ($pat in $Patterns) {
                    try {
                        foreach ($fp in [System.IO.Directory]::EnumerateFiles($ScanRoot,$pat)) {
                            try {
                                $fi=[System.IO.FileInfo]::new($fp); if($fi.Length-eq 0){continue}
                                $fn=[System.IO.Path]::GetFileName($fp); $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                $ft=[System.IO.Path]::GetExtension($fp).TrimStart('.').ToUpperInvariant()
                                [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).Append('","').Append($ft).AppendLine('"')
                                $localCount++; $bufLines++
                                if($ft-eq"SLDASM"){$localAssy++}elseif($ft-eq"SLDPRT"){$localParts++}
                                if($bufLines-ge 1000){$stream.Write($buf.ToString());$buf.Clear();$bufLines=0}
                            }catch{$localErrors++}
                        }
                    }catch{$localErrors++;if($_.Exception.Message-notmatch"denied"){BG-Log "Error scanning $ScanRoot : $($_.Exception.Message)" "ERROR"}}
                }
            } else {
                $stack=[System.Collections.Generic.Stack[string]]::new(); $stack.Push($ScanRoot)
                while($stack.Count-gt 0){
                    $cur=$stack.Pop()
                    if($sh.IsRunning-eq $false){break}
                    try {
                        foreach ($pat in $Patterns) {
                            foreach($fp in [System.IO.Directory]::EnumerateFiles($cur,$pat)){
                                try{
                                    $fi=[System.IO.FileInfo]::new($fp); if($fi.Length-eq 0){continue}
                                    $fn=[System.IO.Path]::GetFileName($fp); $bn=[System.IO.Path]::GetFileNameWithoutExtension($fp)
                                    $lw=$fi.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                                    $ft=[System.IO.Path]::GetExtension($fp).TrimStart('.').ToUpperInvariant()
                                    [void]$buf.Append('"').Append($fn.Replace('"','""')).Append('","').Append($bn.Replace('"','""')).Append('","').Append($fp.Replace('"','""')).Append('","').Append($lw).Append('",').Append($fi.Length).Append(',"').Append($rootEsc).Append('","').Append($ft).AppendLine('"')
                                    $localCount++; $bufLines++
                                    if($ft-eq"SLDASM"){$localAssy++}elseif($ft-eq"SLDPRT"){$localParts++}
                                    if($bufLines-ge 1000){$stream.Write($buf.ToString());$buf.Clear();$bufLines=0}
                                }catch{$localErrors++}
                            }
                        }
                        foreach($sd in [System.IO.Directory]::EnumerateDirectories($cur)){$stack.Push($sd)}
                    }catch{$localErrors++;if($_.Exception.Message-notmatch"denied"){BG-Log "Error in $cur : $($_.Exception.Message)" "ERROR"}}
                }
            }
        } finally {
            if($bufLines-gt 0-and $stream){$stream.Write($buf.ToString())}
            if($stream){$stream.Flush();$stream.Close();$stream.Dispose()}
        }
        BG-Log "  [Worker] Done: $ScanRoot - $localCount model files" "SUCCESS"
        [System.Threading.Monitor]::Enter($Counters)
        try {
            $Counters.Total      += $localCount; $Counters.Errors += $localErrors
            $Counters.Done       += 1;           $Counters.Assemblies += $localAssy; $Counters.Parts += $localParts
            $sh.ParallelTotal     = $Counters.Total
            $sh.FinalErrors       = $Counters.Errors
            $sh.ParallelDone      = $Counters.Done
            $sh.TotalModels       = $Counters.Total.ToString("N0")
            $sh.Assemblies        = $Counters.Assemblies.ToString("N0")
            $sh.Parts             = $Counters.Parts.ToString("N0")
            $sh.Errors            = $Counters.Errors.ToString("N0")
        } finally { [System.Threading.Monitor]::Exit($Counters) }
        $done=$sh.ParallelDone; $total=$sh.ParallelCount
        $sh.Progress=[math]::Floor($done/$total*100)
        $sh.Detail="$($sh.TotalModels) models | Assemblies: $($sh.Assemblies) | Parts: $($sh.Parts) | Done $done/$total units"
    }

    # --- Step 3: Launch all workers ---
    foreach ($unit in $scanUnits) {
        $safeUnitName = ($unit.Path -replace '[\\:/\s]','_')
        if ($safeUnitName.Length -gt 80) { $safeUnitName = $safeUnitName.Substring(0,80) }
        $suffix    = if ($unit.ShallowOnly) { "_SH" } else { "" }
        $chunkFile = Join-Path $tempDir ("chunk_" + $safeUnitName + $suffix + ".csv")
        $wPs = [PowerShell]::Create(); $wPs.RunspacePool = $pool
        [void]$wPs.AddScript($workerScript)
        [void]$wPs.AddArgument($SyncHash); [void]$wPs.AddArgument($unit.Path)
        [void]$wPs.AddArgument($unit.ShallowOnly); [void]$wPs.AddArgument($chunkFile)
        [void]$wPs.AddArgument($logFile); [void]$wPs.AddArgument($counters); [void]$wPs.AddArgument($patterns)
        $wHandle = $wPs.BeginInvoke()
        $workers.Add([pscustomobject]@{ PS=$wPs; Handle=$wHandle; Chunk=$chunkFile; Root=$unit.Path; Shallow=$unit.ShallowOnly })
    }

    # --- Step 4: Coordinator runspace - waits, merges, writes state ---
    $coordRS = [RunspaceFactory]::CreateRunspace(); $coordRS.Open()
    $coordRS.SessionStateProxy.SetVariable("SyncHash",        $SyncHash)
    $coordRS.SessionStateProxy.SetVariable("Workers",         $workers)
    $coordRS.SessionStateProxy.SetVariable("RawCSV",          $RawCSV)
    $coordRS.SessionStateProxy.SetVariable("AllCSV",          $AllCSV)
    $coordRS.SessionStateProxy.SetVariable("LogFile",         $logFile)
    $coordRS.SessionStateProxy.SetVariable("StateFile",       $StateFile)
    $coordRS.SessionStateProxy.SetVariable("selectedFolders", $SelectedFolders)
    $coordRS.SessionStateProxy.SetVariable("Stopwatch",       $stopwatch)
    $coordRS.SessionStateProxy.SetVariable("TempDir",         $tempDir)
    $coordRS.SessionStateProxy.SetVariable("Pool",            $pool)
    $coordRS.SessionStateProxy.SetVariable("Counters",        $counters)
    $coordRS.SessionStateProxy.SetVariable("WorkerScript",    $workerScript)
    $coordRS.SessionStateProxy.SetVariable("Patterns",        $patterns)

    $coordPs = [PowerShell]::Create(); $coordPs.Runspace = $coordRS
    [void]$coordPs.AddScript({
        $sh = $SyncHash
        function BG-Log {param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
        BG-Log "Coordinator: waiting for $($Workers.Count) model workers (RunspacePool)..." "INFO"
        $workerTimeoutMs = 900000; $timedOut = 0
        $retryQueue = [System.Collections.Generic.List[object]]::new()
        foreach($w in $Workers){
            $finished=$w.Handle.AsyncWaitHandle.WaitOne($workerTimeoutMs)
            if($finished){try{$w.PS.EndInvoke($w.Handle)}catch{}}
            else{$timedOut++;BG-Log "  [Worker] TIMEOUT on: $($w.Root) - queuing for sequential retry" "WARN";try{$w.PS.Stop()}catch{};$retryQueue.Add($w)}
            try{$w.PS.Dispose()}catch{}
        }
        if($timedOut -gt 0){
            BG-Log "$timedOut worker(s) timed out - retrying sequentially (no time limit)..." "WARN"
            foreach($rw in $retryQueue){
                if(Test-Path $rw.Chunk){Remove-Item $rw.Chunk -Force -EA SilentlyContinue}
                BG-Log "  [Retry] $($rw.Root) (Shallow=$($rw.Shallow))" "INFO"
                try{
                    $rPs=[PowerShell]::Create();$rPs.RunspacePool=$Pool
                    [void]$rPs.AddScript($WorkerScript)
                    [void]$rPs.AddArgument($SyncHash);[void]$rPs.AddArgument($rw.Root)
                    [void]$rPs.AddArgument($rw.Shallow);[void]$rPs.AddArgument($rw.Chunk)
                    [void]$rPs.AddArgument($LogFile);[void]$rPs.AddArgument($Counters);[void]$rPs.AddArgument($Patterns)
                    $rh=$rPs.BeginInvoke()
                    $rd=$rh.AsyncWaitHandle.WaitOne(7200000)
                    if($rd){try{$rPs.EndInvoke($rh)}catch{};BG-Log "  [Retry] Done: $($rw.Root)" "SUCCESS"}
                    else{BG-Log "  [Retry] STILL timed out after 2h: $($rw.Root)" "WARN"}
                    try{$rPs.Dispose()}catch{}
                }catch{BG-Log "  [Retry] Error: $($rw.Root): $_" "WARN"}
            }
            BG-Log "Sequential retry complete." "INFO"
        }
        try { $Pool.Close(); $Pool.Dispose() } catch {}
        BG-Log "All model workers done. Merging chunks..." "INFO"
        $sh.Status = "Merging results..."

        # Append chunk files to RawCSV (header already written before workers started)
        $ms=$null; $totalMerged=0
        try{
            $ms=[System.IO.StreamWriter]::new($RawCSV,$true,[System.Text.Encoding]::UTF8,65536)
            foreach($w in $Workers){
                if(Test-Path $w.Chunk){
                    $reader=$null
                    try {
                        $reader=[System.IO.StreamReader]::new($w.Chunk,[System.Text.Encoding]::UTF8,$true,65536)
                        while(-not $reader.EndOfStream){$line=$reader.ReadLine();if($line){$ms.WriteLine($line);$totalMerged++}}
                    } finally { if($reader){$reader.Close();$reader.Dispose()} }
                    Remove-Item $w.Chunk -Force -ErrorAction SilentlyContinue
                }
            }
        }finally{if($ms){$ms.Flush();$ms.Close();$ms.Dispose()}}
        BG-Log "Merged $totalMerged model rows into $RawCSV" "INFO"
        try{if((Get-ChildItem $TempDir -ErrorAction SilentlyContinue|Measure-Object).Count-eq 0){Remove-Item $TempDir -Force -ErrorAction SilentlyContinue}}catch{}

        # Copy raw to AllCSV
        try { Copy-Item -Path $RawCSV -Destination $AllCSV -Force } catch { BG-Log "Failed copy to AllCSV: $($_.Exception.Message)" "ERROR" }

        $Stopwatch.Stop(); $elapsed=$Stopwatch.Elapsed.ToString('hh\:mm\:ss')
        $totalFiles=$Counters.Total; $totalErrors=$Counters.Errors; $assyCount=$Counters.Assemblies; $partCount=$Counters.Parts
        $state=@{
            LastCrawlTime=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss");TotalFiles=[int]$totalFiles
            TotalAssemblies=[int]$assyCount;TotalParts=[int]$partCount;TotalErrors=[int]$totalErrors
            ElapsedTime=$elapsed;FoldersCrawled=$selectedFolders;ScanUnits=$Workers.Count
            Patterns=@("*.sldasm","*.sldprt");Mode="Parallel"
        }|ConvertTo-Json
        [System.IO.File]::WriteAllText($StateFile,$state)
        BG-Log "MODEL CRAWL COMPLETE (PARALLEL): $totalFiles files ($assyCount assemblies, $partCount parts), $totalErrors errors, $elapsed ($($Workers.Count) units)" "SUCCESS"
        $sh.FinalTotal=$totalFiles; $sh.FinalAssemblies=$assyCount; $sh.FinalParts=$partCount
        $sh.FinalErrors=$totalErrors; $sh.FinalTime=$elapsed; $sh.IsComplete=$true
    })

    $UITimer.Start()
    $coordPs.BeginInvoke() | Out-Null
    $script:ModelBgPowerShell = $coordPs
    $script:ModelBgRunspace   = $coordRS
    $script:ModelBgHandle     = $null
}

# ============================================================================
#  GENERIC DEDUP FUNCTION - used for PDF, DXF, and Model
# ============================================================================
function Run-Dedup {
    param(
        [string]$RawCSV,
        [string]$CleanCSV,
        [object]$StatusTxt,
        [object]$StatUnique,
        [object]$StatDupes,
        [object]$StatTotal
    )

    if (-not (Test-Path $RawCSV)) {
        [System.Windows.MessageBox]::Show("No raw index found. Run the crawler first.", "No Data",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $StatusTxt.Text = "Processing..."
    $StatusTxt.Foreground = [System.Windows.Media.Brushes]::Gold
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $rawData = Import-Csv -Path $RawCSV
        Add-LogEntry "Dedup: loaded $($rawData.Count) rows from $RawCSV"

        $partGroups = @{}
        foreach ($row in $rawData) {
            try {
                $info = Get-PartNumberInfo -BaseName $row.BaseName
                $key  = $info.BasePart
                if (-not $partGroups.ContainsKey($key)) { $partGroups[$key] = [System.Collections.ArrayList]::new() }
                [void]$partGroups[$key].Add(@{
                    FileName=$row.FileName; BaseName=$row.BaseName; FullPath=$row.FullPath
                    LastWriteTime=[datetime]$row.LastWriteTime; SizeBytes=[long]$row.SizeBytes
                    RootFolder=$row.RootFolder; BasePart=$info.BasePart; RevRaw=$info.RevRaw
                    RevValue=$info.RevValue; HasRev=$info.HasRev
                    IsObsolete=(Test-ObsoletePath $row.FullPath)
                })
            } catch { continue }
        }

        $cleanResults = [System.Collections.ArrayList]::new()
        foreach ($key in $partGroups.Keys) {
            $candidates = $partGroups[$key]
            if ($candidates.Count -eq 1) { $winner = $candidates[0] }
            else {
                # Sort priority: non-obsolete first, then highest revision, then newest date
                $withRev = @($candidates | Where-Object { $_.HasRev -eq $true })
                if ($withRev.Count -gt 0) {
                    $winner = $withRev | Sort-Object @{Expression={$_.IsObsolete};Ascending=$true},@{Expression={$_.RevValue};Descending=$true},@{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
                } else {
                    $winner = $candidates | Sort-Object @{Expression={$_.IsObsolete};Ascending=$true},@{Expression={$_.LastWriteTime};Descending=$true} | Select-Object -First 1
                }
            }
            $revDisplay = if ($winner.RevRaw) { $winner.RevRaw } else { "(none)" }
            $isObs = if ($winner.IsObsolete) { "Yes" } else { "No" }
            [void]$cleanResults.Add([PSCustomObject]@{
                BasePart=$winner.BasePart; Rev=$revDisplay; FileName=$winner.FileName
                FullPath=$winner.FullPath; LastWriteTime=$winner.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
                SizeKB=[math]::Round($winner.SizeBytes/1024,1); RootFolder=$winner.RootFolder; Candidates=$candidates.Count
                IsObsolete=$isObs
            })
        }

        $cleanResults = $cleanResults | Sort-Object BasePart
        $cleanResults | Export-Csv -Path $CleanCSV -NoTypeInformation -Encoding UTF8

        $dupes = $rawData.Count - $cleanResults.Count
        $obsoleteCount = @($cleanResults | Where-Object { $_.IsObsolete -eq "Yes" }).Count
        $StatUnique.Text = $cleanResults.Count.ToString("N0")
        $StatDupes.Text  = $dupes.ToString("N0")
        if ($StatTotal) { $StatTotal.Text = $rawData.Count.ToString("N0") }

        $statusMsg = "Complete: $($cleanResults.Count) unique from $($rawData.Count) total ($dupes duplicates resolved)"
        if ($obsoleteCount -gt 0) {
            $statusMsg += " | $obsoleteCount from obsolete folders"
        }
        $StatusTxt.Text = $statusMsg
        $StatusTxt.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
        Add-LogEntry "Dedup complete: $($cleanResults.Count) unique, $dupes duplicates, $obsoleteCount obsolete" "SUCCESS"

    } catch {
        $StatusTxt.Text = "Error: $($_.Exception.Message)"
        $StatusTxt.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        Add-LogEntry "Dedup error: $($_.Exception.Message)" "ERROR"
    }
}

# ============================================================================
#  FOLDER MANAGEMENT - PDF TAB
# ============================================================================

function Add-FolderCheckbox {
    param([object]$Panel, [string]$Path, [bool]$Checked = $true)
    $cb = [System.Windows.Controls.CheckBox]::new()
    $cb.Content = $Path; $cb.IsChecked = $Checked; $cb.Tag = $Path
    $cb.FontFamily = [System.Windows.Media.FontFamily]::new("Consolas"); $cb.FontSize = 11
    $Panel.Children.Add($cb) | Out-Null
}

foreach ($root in $script:DefaultRoots) {
    Add-FolderCheckbox -Panel $FolderCheckboxPanel    -Path $root -Checked $true
    Add-FolderCheckbox -Panel $DxfFolderCheckboxPanel -Path $root -Checked $true
    Add-FolderCheckbox -Panel $ModelFolderCheckboxPanel -Path $root -Checked $true
}

$BtnAddFolder.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select a root folder to scan for PDFs"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FolderCheckbox -Panel $FolderCheckboxPanel -Path $dialog.SelectedPath -Checked $true
        Add-LogEntry "Added PDF scan folder: $($dialog.SelectedPath)"
    }
})

$BtnRemoveFolder.Add_Click({
    $toRemove = @($FolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and -not $_.IsChecked })
    foreach ($cb in $toRemove) { $FolderCheckboxPanel.Children.Remove($cb); Add-LogEntry "Removed folder: $($cb.Tag)" }
    if ($toRemove.Count -eq 0) { [System.Windows.MessageBox]::Show("Uncheck folders to remove, then click Remove.", "No Folders Selected") }
})

$BtnDxfAddFolder.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select a root folder to scan for DXF files"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FolderCheckbox -Panel $DxfFolderCheckboxPanel -Path $dialog.SelectedPath -Checked $true
        Add-LogEntry "Added DXF scan folder: $($dialog.SelectedPath)"
    }
})

$BtnDxfRemoveFolder.Add_Click({
    $toRemove = @($DxfFolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and -not $_.IsChecked })
    foreach ($cb in $toRemove) { $DxfFolderCheckboxPanel.Children.Remove($cb); Add-LogEntry "Removed DXF folder: $($cb.Tag)" }
    if ($toRemove.Count -eq 0) { [System.Windows.MessageBox]::Show("Uncheck folders to remove, then click Remove.", "No Folders Selected") }
})

$BtnModelAddFolder.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select a root folder to scan for model files"
    $dialog.ShowNewFolderButton = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        Add-FolderCheckbox -Panel $ModelFolderCheckboxPanel -Path $dialog.SelectedPath -Checked $true
        Add-LogEntry "Added model scan folder: $($dialog.SelectedPath)"
    }
})

$BtnModelRemoveFolder.Add_Click({
    $toRemove = @($ModelFolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and -not $_.IsChecked })
    foreach ($cb in $toRemove) { $ModelFolderCheckboxPanel.Children.Remove($cb); Add-LogEntry "Removed model folder: $($cb.Tag)" }
    if ($toRemove.Count -eq 0) { [System.Windows.MessageBox]::Show("Uncheck folders to remove, then click Remove.", "No Folders Selected") }
})

# ============================================================================
#  PDF CRAWL BUTTON
# ============================================================================

$BtnStartCrawl.Add_Click({
    if ($script:CrawlRunning) { return }
    $selectedFolders = @($FolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked } | ForEach-Object { $_.Tag })
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please check at least one folder.", "No Folders Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    $script:CrawlRunning = $true
    $BtnStartCrawl.IsEnabled = $false; $BtnStopCrawl.IsEnabled = $true
    $TxtCrawlStatus.Text = "Crawling PDFs..."
    $TxtCrawlStatus.Foreground = [System.Windows.Media.Brushes]::Gold
    $CrawlProgress.Value = 0
    Add-LogEntry "PDF Crawl starting ($($selectedFolders.Count) sections)" "INFO"

    Start-FileCrawl -SelectedFolders $selectedFolders `
        -RawCSV $script:RawCSV -StateFile $script:StateFile `
        -FileExt "*.pdf" -SyncHash $script:SyncHash `
        -StartBtn $BtnStartCrawl -StopBtn $BtnStopCrawl -StatusTxt $TxtCrawlStatus `
        -UseParallel ($ChkParallel.IsChecked -eq $true) `
        -WorkersRef ([ref]$script:ParallelWorkers) `
        -PoolRef ([ref]$script:BgPool) `
        -UITimer $script:UITimer
})

$BtnStopCrawl.Add_Click({
    $script:SyncHash.IsRunning = $false; $script:UITimer.Stop()
    if ($script:ParallelWorkers) {
        foreach ($w in $script:ParallelWorkers) { try{$w.PS.Stop()}catch{}; try{$w.PS.Dispose()}catch{} }
        $script:ParallelWorkers = $null
    }
    if ($script:BgPool) { try { $script:BgPool.Close(); $script:BgPool.Dispose() } catch {}; $script:BgPool = $null }
    if ($script:BgPowerShell) { try { $script:BgPowerShell.Stop() } catch {} }
    Add-LogEntry "PDF crawl stopped by user" "WARN"
    $TxtCrawlStatus.Text = "Stopped"; $TxtCrawlStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
    $BtnStartCrawl.IsEnabled = $true; $BtnStopCrawl.IsEnabled = $false; $script:CrawlRunning = $false
})

# ============================================================================
#  DXF CRAWL BUTTON
# ============================================================================

$BtnStartDxfCrawl.Add_Click({
    if ($script:DxfCrawlRunning) { return }
    $selectedFolders = @($DxfFolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked } | ForEach-Object { $_.Tag })
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please check at least one folder.", "No Folders Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    $script:DxfCrawlRunning = $true
    $BtnStartDxfCrawl.IsEnabled = $false; $BtnStopDxfCrawl.IsEnabled = $true
    $TxtDxfCrawlStatus.Text = "Crawling DXFs..."
    $TxtDxfCrawlStatus.Foreground = [System.Windows.Media.Brushes]::Gold
    $DxfCrawlProgress.Value = 0
    Add-LogEntry "DXF Crawl starting ($($selectedFolders.Count) sections)" "INFO"

    Start-FileCrawl -SelectedFolders $selectedFolders `
        -RawCSV $script:DxfRawCSV -StateFile $script:DxfStateFile `
        -FileExt "*.dxf" -SyncHash $script:DxfSyncHash `
        -StartBtn $BtnStartDxfCrawl -StopBtn $BtnStopDxfCrawl -StatusTxt $TxtDxfCrawlStatus `
        -UseParallel ($ChkDxfParallel.IsChecked -eq $true) `
        -WorkersRef ([ref]$script:DxfParallelWorkers) `
        -PoolRef ([ref]$script:DxfBgPool) `
        -UITimer $script:DxfUITimer
})

$BtnStopDxfCrawl.Add_Click({
    $script:DxfSyncHash.IsRunning = $false; $script:DxfUITimer.Stop()
    if ($script:DxfParallelWorkers) {
        foreach ($w in $script:DxfParallelWorkers) { try{$w.PS.Stop()}catch{}; try{$w.PS.Dispose()}catch{} }
        $script:DxfParallelWorkers = $null
    }
    if ($script:DxfBgPool) { try { $script:DxfBgPool.Close(); $script:DxfBgPool.Dispose() } catch {}; $script:DxfBgPool = $null }
    Add-LogEntry "DXF crawl stopped by user" "WARN"
    $TxtDxfCrawlStatus.Text = "Stopped"; $TxtDxfCrawlStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
    $BtnStartDxfCrawl.IsEnabled = $true; $BtnStopDxfCrawl.IsEnabled = $false; $script:DxfCrawlRunning = $false
})

# ============================================================================
#  MODEL CRAWL BUTTON
# ============================================================================

$BtnStartModelCrawl.Add_Click({
    if ($script:ModelCrawlRunning) { return }
    $selectedFolders = @($ModelFolderCheckboxPanel.Children | Where-Object { $_ -is [System.Windows.Controls.CheckBox] -and $_.IsChecked } | ForEach-Object { $_.Tag })
    if ($selectedFolders.Count -eq 0) {
        [System.Windows.MessageBox]::Show("Please check at least one folder.", "No Folders Selected", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $script:ModelCrawlRunning = $true
    $BtnStartModelCrawl.IsEnabled = $false
    $BtnStopModelCrawl.IsEnabled = $true
    $TxtModelCrawlStatus.Text = "Crawling model files..."
    $TxtModelCrawlStatus.Foreground = [System.Windows.Media.Brushes]::Gold
    $ModelCrawlProgress.Value = 0
    Add-LogEntry "Model Crawl starting ($($selectedFolders.Count) sections)" "INFO"

    Start-ModelCrawl -SelectedFolders $selectedFolders `
        -RawCSV $script:ModelRawCSV -AllCSV $script:ModelAllCSV -StateFile $script:ModelStateFile `
        -SyncHash $script:ModelSyncHash `
        -UITimer $script:ModelUITimer `
        -UseParallel ($ChkModelParallel.IsChecked -eq $true)
})

$BtnStopModelCrawl.Add_Click({
    $script:ModelSyncHash.IsRunning = $false
    $script:ModelUITimer.Stop()
    if ($script:ModelBgPowerShell) {
        try { $script:ModelBgPowerShell.Stop() } catch {}
        try { $script:ModelBgPowerShell.Dispose() } catch {}
        $script:ModelBgPowerShell = $null
    }
    if ($script:ModelBgRunspace) {
        try { $script:ModelBgRunspace.Close() } catch {}
        try { $script:ModelBgRunspace.Dispose() } catch {}
        $script:ModelBgRunspace = $null
    }
    Add-LogEntry "Model crawl stopped by user" "WARN"
    $TxtModelCrawlStatus.Text = "Stopped"
    $TxtModelCrawlStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
    $BtnStartModelCrawl.IsEnabled = $true
    $BtnStopModelCrawl.IsEnabled = $false
    $script:ModelCrawlRunning = $false
})

# ============================================================================
#  DEDUP BUTTONS
# ============================================================================

$BtnRunDedup.Add_Click({
    Run-Dedup -RawCSV $script:RawCSV -CleanCSV $script:CleanCSV `
        -StatusTxt $TxtDedupStatus -StatUnique $StatUniqueParts -StatDupes $StatDuplicates -StatTotal $StatTotalPDFs
})

$BtnRunDxfDedup.Add_Click({
    Run-Dedup -RawCSV $script:DxfRawCSV -CleanCSV $script:DxfCleanCSV `
        -StatusTxt $TxtDxfDedupStatus -StatUnique $StatUniqueDxfParts -StatDupes $StatDxfDuplicates -StatTotal $StatTotalDXFs
})

$BtnRunModelDedup.Add_Click({
    Run-Dedup -RawCSV $script:ModelRawCSV -CleanCSV $script:ModelCleanCSV `
        -StatusTxt $TxtModelDedupStatus -StatUnique $StatModelAssemblies -StatDupes $StatModelParts -StatTotal $StatTotalModels
    Update-ModelStats
})

$BtnOpenClean.Add_Click({
    if (Test-Path $script:CleanCSV) { Start-Process $script:CleanCSV }
    else { [System.Windows.MessageBox]::Show("Clean CSV not found. Run deduplication first.", "Not Found") }
})
$BtnOpenRaw.Add_Click({
    if (Test-Path $script:RawCSV) { Start-Process $script:RawCSV }
    else { [System.Windows.MessageBox]::Show("Raw CSV not found. Run the crawler first.", "Not Found") }
})
$BtnOpenDxfClean.Add_Click({
    if (Test-Path $script:DxfCleanCSV) { Start-Process $script:DxfCleanCSV }
    else { [System.Windows.MessageBox]::Show("DXF clean CSV not found. Run DXF deduplication first.", "Not Found") }
})
$BtnOpenDxfRaw.Add_Click({
    if (Test-Path $script:DxfRawCSV) { Start-Process $script:DxfRawCSV }
    else { [System.Windows.MessageBox]::Show("DXF raw CSV not found. Run the DXF crawler first.", "Not Found") }
})
$BtnOpenModelAll.Add_Click({
    if (Test-Path $script:ModelAllCSV) { Start-Process $script:ModelAllCSV }
    else { [System.Windows.MessageBox]::Show("Model all CSV not found. Run the model crawler first.", "Not Found") }
})
$BtnOpenModelClean.Add_Click({
    if (Test-Path $script:ModelCleanCSV) { Start-Process $script:ModelCleanCSV }
    else { [System.Windows.MessageBox]::Show("Model clean CSV not found. Run model deduplication first.", "Not Found") }
})

# ============================================================================
#  PDM VAULT SEARCH (PDF only — DXFs not typically in PDM)
# ============================================================================

$BtnSearchPDM.Add_Click({
    if ($script:CrawlRunning) {
        [System.Windows.MessageBox]::Show("A crawl is already running. Please wait.", "Crawl In Progress",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        return
    }
    $pdmScript = "C:\Users\dlebel\Downloads\feb13\PDMCrawler.ps1"
    if (-not (Test-Path $pdmScript)) {
        [System.Windows.MessageBox]::Show("PDMCrawler.ps1 not found at:`n$pdmScript", "Script Not Found",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }
    Add-LogEntry "PDM VAULT SEARCH STARTING" "INFO"
    $TxtCrawlStatus.Text = "Searching PDM Vault..."; $TxtCrawlStatus.Foreground = [System.Windows.Media.Brushes]::Gold
    $CrawlProgress.IsIndeterminate = $true; $BtnSearchPDM.IsEnabled = $false; $BtnStartCrawl.IsEnabled = $false
    $script:CrawlRunning = $true

    $pdmRS = [RunspaceFactory]::CreateRunspace(); $pdmRS.Open()
    $pdmRS.SessionStateProxy.SetVariable("SyncHash",  $script:SyncHash)
    $pdmRS.SessionStateProxy.SetVariable("OutputCSV", $script:RawCSV)
    $pdmRS.SessionStateProxy.SetVariable("LogFile",   $script:LogFile)

    $pdmPs = [PowerShell]::Create(); $pdmPs.Runspace = $pdmRS
    [void]$pdmPs.AddScript({
        $sh=$SyncHash
        function BG-Log{param([string]$Msg,[string]$Level="INFO");$ts=Get-Date -Format "HH:mm:ss";$pfx=switch($Level){"ERROR"{"[!]"}"WARN"{"[~]"}"SUCCESS"{"[+]"}default{"[ ]"}};$entry="$ts $pfx $Msg";$sh.LogQueue.Add($entry)|Out-Null;try{Add-Content -Path $LogFile -Value $entry -EA SilentlyContinue}catch{}}
        $sh.IsRunning=$true;$totalFound=0;$totalErrors=0
        try{
            BG-Log "Connecting to PDM vault: NMT_PDM..." "INFO"
            $vault=New-Object -ComObject "ConisioLib.EdmVault"
            $vault.LoginAuto("NMT_PDM",0)
            BG-Log "Connected." "INFO"
            $rootID=1;try{$rootID=$vault.RootFolder.ID}catch{}
            $search=$vault.CreateSearch();$search.FileName="%.pdf";$search.Recursive=$true;$search.FindFiles=$true;$search.StartFolderID=$rootID
            BG-Log "Searching vault for all PDFs..." "INFO"
            $rows=[System.Collections.Generic.List[string]]::new()
            $pos=$search.GetFirstResult()
            while($pos-ne $null){
                $fn=$null;$fp="";$fs=0;$fd=""
                try{$fn=$pos.Name}catch{};try{$fp=$pos.Path}catch{};try{$fs=$pos.FileSize}catch{}
                try{$rawDate=$pos.FileDate;if($rawDate-is[datetime]){$fd=$rawDate.ToString("yyyy-MM-dd HH:mm:ss")}else{$fd=[datetime]::FromOADate([double]$rawDate).ToString("yyyy-MM-dd HH:mm:ss")}}catch{}
                if($fn){
                    $bn=[System.IO.Path]::GetFileNameWithoutExtension($fn).Replace('"','""')
                    $row='"'+$fn.Replace('"','""')+'","'+$bn+'","'+$fp.Replace('"','""')+'","'+$fd+'",'+$fs+',"PDM:NMT_PDM"'
                    $rows.Add($row);$totalFound++
                    if($totalFound%500-eq 0){BG-Log "  ... $totalFound PDFs found" "INFO"}
                }
                try{$pos=$search.GetNextResult()}catch{break}
            }
            BG-Log "Search complete: $totalFound PDFs found." "SUCCESS"
            if($rows.Count-gt 0){
                $outDir=Split-Path $OutputCSV -Parent
                if(-not(Test-Path $outDir)){New-Item -ItemType Directory -Path $outDir -Force|Out-Null}
                if(-not(Test-Path $OutputCSV)){'"FileName","BaseName","FullPath","LastWriteTime","SizeBytes","RootFolder"'|Out-File -FilePath $OutputCSV -Encoding UTF8 -Force}
                else{$bytes=[System.IO.File]::ReadAllBytes($OutputCSV);if($bytes.Length-gt 0-and $bytes[-1]-ne 10){[System.IO.File]::AppendAllText($OutputCSV,[Environment]::NewLine)}}
                $rows|Out-File -FilePath $OutputCSV -Encoding UTF8 -Append
                BG-Log "Done. $totalFound PDFs appended to index." "SUCCESS"
            }
        }catch{BG-Log "ERROR: $($_.Exception.Message)" "ERROR";$totalErrors++}
        $sh.FinalTotal=$totalFound;$sh.FinalErrors=$totalErrors;$sh.FinalTime="N/A";$sh.IsComplete=$true
    })

    $script:PDMTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:PDMTimer.Interval = [TimeSpan]::FromMilliseconds(500)
    $script:PDMTimer.Add_Tick({
        $sh=$script:SyncHash
        while($sh.LogQueue.Count-gt 0){$entry=$sh.LogQueue[0];$sh.LogQueue.RemoveAt(0);$TxtLog.Text+="$entry`r`n";$LogScroller.ScrollToEnd()}
        if($sh.IsComplete){
            $sh.IsComplete=$false;$sh.IsRunning=$false;$script:CrawlRunning=$false;$script:PDMTimer.Stop()
            $CrawlProgress.IsIndeterminate=$false;$CrawlProgress.Value=100
            $TxtCrawlStatus.Text="PDM Search Complete";$TxtCrawlStatus.Foreground=[System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
            $TxtCrawlDetail.Text="PDM vault results appended. Run Deduplication next."
            $BtnSearchPDM.IsEnabled=$true;$BtnStartCrawl.IsEnabled=$true;$BtnStopCrawl.IsEnabled=$false
            Update-Stats;Add-LogEntry "PDM search complete." "SUCCESS"
        }
    })
    $script:PDMTimer.Start();$pdmPs.BeginInvoke()|Out-Null
    $script:BgPDMPowerShell=$pdmPs;$script:BgPDMRunspace=$pdmRS
})

# ============================================================================
#  COLLECT TAB
# ============================================================================

$BtnBrowseBOM.Add_Click({
    $dialog = [System.Windows.Forms.OpenFileDialog]::new()
    $dialog.Title = "Select BOM / Part Number List"
    $dialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtBOMPath.Text = $dialog.FileName }
})

$BtnBrowseCollectOutput.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.Description = "Select output folder for collected files"
    $dialog.ShowNewFolderButton = $true
    if ($TxtCollectOutput.Text -ne "" -and (Test-Path $TxtCollectOutput.Text)) { $dialog.SelectedPath = $TxtCollectOutput.Text }
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $TxtCollectOutput.Text = $dialog.SelectedPath }
})

$BtnOpenCollectFolder.Add_Click({
    $folder = $TxtCollectOutput.Text
    if ($folder -ne "" -and (Test-Path $folder)) { Start-Process explorer.exe $folder }
})

$BtnCollect.Add_Click({
    $bomPath   = $TxtBOMPath.Text.Trim()
    $outFolder = $TxtCollectOutput.Text.Trim()

    if ($bomPath -eq "" -or -not (Test-Path $bomPath)) {
        [System.Windows.MessageBox]::Show("Please select a valid BOM file first.", "BOM Required",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
    }
    if ($outFolder -eq "") {
        [System.Windows.MessageBox]::Show("Please set an output folder first.", "Output Folder Required",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
    }

    # Determine mode
    $collectPDFs = ($RbPdfOnly.IsChecked -or $RbBoth.IsChecked)
    $collectDXFs = ($RbDxfOnly.IsChecked -or $RbBoth.IsChecked)

    if ($collectPDFs -and -not (Test-Path $script:CleanCSV)) {
        [System.Windows.MessageBox]::Show("PDF clean index not found.`nRun PDF Crawler + Deduplication first.", "PDF Index Missing",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
    }
    if ($collectDXFs -and -not (Test-Path $script:DxfCleanCSV)) {
        [System.Windows.MessageBox]::Show("DXF clean index not found.`nRun DXF Crawler + Deduplication first.", "DXF Index Missing",
            [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning); return
    }

    $BtnCollect.IsEnabled = $false
    $CollectProgress.Value = 0
    $TxtCollectResults.Text = ""
    $TxtCollectStatus.Foreground = [System.Windows.Media.Brushes]::Gold
    $TxtCollectStatus.Text = "Loading index..."
    $TxtCollectDetail.Text = ""
    $StatCollectTotal.Text = "--"; $StatCollectPdfs.Text = "--"; $StatCollectDxfs.Text = "--"; $StatCollectMissing.Text = "--"
    $BtnOpenCollectFolder.IsEnabled = $false
    [System.Windows.Forms.Application]::DoEvents()

    try {
        # Load indexes
        $pdfIndex = $null; $dxfIndex = $null
        if ($collectPDFs) { $pdfIndex = Import-Csv -Path $script:CleanCSV }
        if ($collectDXFs) { $dxfIndex = Import-Csv -Path $script:DxfCleanCSV }

        $TxtCollectStatus.Text = "Reading BOM..."
        [System.Windows.Forms.Application]::DoEvents()

        # Read BOM
        $lines = Get-Content $bomPath
        $partNumbers = @()
        foreach ($line in $lines) {
            if ($line -match '^\s*(Part\s+Number|Item\s+No|={2,}|#)' -or $line.Trim() -eq '') { continue }
            $clean = $line.Trim().ToUpper()
            if ($clean -ne '') { $partNumbers += $clean }
        }
        $partNumbers = $partNumbers | Select-Object -Unique
        $StatCollectTotal.Text = $partNumbers.Count
        $TxtCollectStatus.Text = "Collecting $($partNumbers.Count) parts..."
        [System.Windows.Forms.Application]::DoEvents()

        # Create output folders
        if (-not (Test-Path $outFolder)) { New-Item -ItemType Directory -Path $outFolder -Force | Out-Null }
        $dxfFolder = Join-Path $outFolder "DXFs"
        if ($collectDXFs -and -not (Test-Path $dxfFolder)) { New-Item -ItemType Directory -Path $dxfFolder -Force | Out-Null }

        $pdfFound = 0; $dxfFound = 0; $missing = @()
        $log = [System.Text.StringBuilder]::new()
        $total = $partNumbers.Count; $idx = 0

        foreach ($partNum in $partNumbers) {
            $idx++
            $CollectProgress.Value = [math]::Round($idx / $total * 100)
            [System.Windows.Forms.Application]::DoEvents()

            $pdfOk = $false; $dxfOk = $false

            # Strip revision suffix from BOM part number so we can match against BasePart.
            # e.g. "25347-P24_REV1" -> "25347-P24", "25347-W02-05_REV2" -> "25347-W02-05"
            $partNumBase = $partNum -replace '_REV\w+$','' -replace '-REV\w+$','' -replace '\s+REV\w+$',''

            # --- Collect PDF ---
            if ($collectPDFs) {
                $match = $pdfIndex | Where-Object { $_.BasePart -eq $partNum -or $_.BasePart -eq $partNumBase -or $_.FileName -like "*$partNum*" } | Select-Object -First 1
                if ($match -and (Test-Path $match.FullPath)) {
                    $dest = Join-Path $outFolder $match.FileName
                    $counter = 1
                    while (Test-Path $dest) {
                        $bn = [System.IO.Path]::GetFileNameWithoutExtension($match.FileName)
                        $ext = [System.IO.Path]::GetExtension($match.FileName)
                        $dest = Join-Path $outFolder "${bn}_${counter}${ext}"; $counter++
                    }
                    Copy-Item -Path $match.FullPath -Destination $dest -Force
                    $pdfFound++; $pdfOk = $true
                    [void]$log.AppendLine("[PDF-OK]  $partNum  ->  $($match.FileName)")
                } elseif ($collectPDFs -and -not $collectDXFs) {
                    [void]$log.AppendLine("[PDF-X]   $partNum  ->  not found in PDF index")
                }
            }

            # --- Collect DXF ---
            # For each BOM part number we collect:
            #   1. Exact BasePart match (e.g. 25347-W01)
            #   2. Sub-part matches  (e.g. 25347-W01-02, 25347-W01-03 ...)
            #      These are weldment cutlist plates that only exist as DXFs,
            #      never as individual SolidWorks parts in the assembly tree.
            #   3. Rev-stripped match (e.g. BOM has 25347-P24_Rev1 -> match 25347-P24)
            if ($collectDXFs) {
                $dxfMatches = $dxfIndex | Where-Object {
                    $_.BasePart -eq $partNum -or
                    $_.BasePart -eq $partNumBase -or
                    $_.BasePart -like "$partNumBase-*" -or
                    $_.FileName -like "*$partNum*"
                }
                if ($dxfMatches) {
                    foreach ($match in $dxfMatches) {
                        if (-not (Test-Path $match.FullPath)) {
                            [void]$log.AppendLine("[DXF-MISSING-FILE]  $partNum  ->  $($match.FullPath) not on disk")
                            continue
                        }
                        $dest = Join-Path $dxfFolder $match.FileName
                        $counter = 1
                        while (Test-Path $dest) {
                            $bn  = [System.IO.Path]::GetFileNameWithoutExtension($match.FileName)
                            $ext = [System.IO.Path]::GetExtension($match.FileName)
                            $dest = Join-Path $dxfFolder "${bn}_${counter}${ext}"; $counter++
                        }
                        Copy-Item -Path $match.FullPath -Destination $dest -Force
                        $dxfFound++; $dxfOk = $true
                        [void]$log.AppendLine("[DXF-OK]  $partNum  ->  $($match.FileName)")
                    }
                } else {
                    [void]$log.AppendLine("[DXF-X]   $partNum  ->  not found in DXF index")
                }
            }

            # Track as missing if NOTHING was found for this part
            if ((-not $pdfOk -and $collectPDFs) -or (-not $dxfOk -and $collectDXFs)) {
                if (-not $pdfOk -and -not $dxfOk) { $missing += $partNum }
            }
        }

        # Done
        $CollectProgress.Value = 100
        $StatCollectPdfs.Text    = if ($collectPDFs) { $pdfFound } else { "N/A" }
        $StatCollectDxfs.Text    = if ($collectDXFs) { $dxfFound } else { "N/A" }
        $StatCollectMissing.Text = $missing.Count

        $modeStr = if ($RbPdfOnly.IsChecked) { "PDF Only" } elseif ($RbDxfOnly.IsChecked) { "DXF Only" } else { "PDF + DXF" }

        if ($missing.Count -eq 0) {
            $TxtCollectStatus.Text = "Complete! [$modeStr]  PDFs: $pdfFound  DXFs: $dxfFound"
            $TxtCollectStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#4caf50")
        } else {
            $TxtCollectStatus.Text = "Complete [$modeStr]  PDFs: $pdfFound  DXFs: $dxfFound  Missing: $($missing.Count)"
            $TxtCollectStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#ffc107")
        }

        if ($missing.Count -gt 0) {
            [void]$log.AppendLine(""); [void]$log.AppendLine("--- NOT FOUND ---")
            $missing | ForEach-Object { [void]$log.AppendLine("  - $_") }
        }

        $TxtCollectResults.Text = $log.ToString()
        $TxtCollectDetail.Text  = "Output: $outFolder" + $(if ($collectDXFs) { "  |  DXFs: $dxfFolder" } else { "" })
        $BtnOpenCollectFolder.IsEnabled = $true
        Add-LogEntry "Collect complete [$modeStr]: PDFs=$pdfFound DXFs=$dxfFound Missing=$($missing.Count)" "SUCCESS"
        $CollectResultScroller.ScrollToEnd()

    } catch {
        $TxtCollectStatus.Text = "Error: $($_.Exception.Message)"
        $TxtCollectStatus.Foreground = [System.Windows.Media.Brushes]::OrangeRed
        Add-LogEntry "Collect error: $($_.Exception.Message)" "ERROR"
    } finally {
        $BtnCollect.IsEnabled = $true
    }
})

# ============================================================================
#  UTILITY / SETTINGS BUTTONS
# ============================================================================

$BtnClearLog.Add_Click({ $TxtLog.Text = ""; Add-LogEntry "Log cleared" })
$BtnOpenLogFile.Add_Click({ if (Test-Path $script:LogFile) { Start-Process notepad.exe $script:LogFile } })

$BtnBrowseOutput.Add_Click({
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    $dialog.SelectedPath = $TxtOutputDir.Text
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:OutputDir   = $dialog.SelectedPath
        $script:RawCSV      = Join-Path $script:OutputDir "pdf_index_raw.csv"
        $script:CleanCSV    = Join-Path $script:OutputDir "pdf_index_clean.csv"
        $script:LogFile     = Join-Path $script:OutputDir "crawl_log.txt"
        $script:StateFile   = Join-Path $script:OutputDir "crawl_state.json"
        $script:DxfRawCSV   = Join-Path $script:OutputDir "dxf_index_raw.csv"
        $script:DxfCleanCSV = Join-Path $script:OutputDir "dxf_index_clean.csv"
        $script:DxfStateFile= Join-Path $script:OutputDir "dxf_crawl_state.json"
        $script:ModelRawCSV = Join-Path $script:OutputDir "model_index_raw.csv"
        $script:ModelAllCSV = Join-Path $script:OutputDir "model_index_all.csv"
        $script:ModelCleanCSV = Join-Path $script:OutputDir "model_index_clean.csv"
        $script:ModelStateFile = Join-Path $script:OutputDir "model_crawl_state.json"
        $TxtOutputDir.Text  = $script:OutputDir
        Add-LogEntry "Output directory changed to: $($script:OutputDir)"
    }
})

$BtnOpenOutputDir.Add_Click({
    if (Test-Path $script:OutputDir) { Start-Process explorer.exe $script:OutputDir }
})

$BtnResetIndex.Add_Click({
    $result = [System.Windows.MessageBox]::Show(
        "This will delete all PDF, DXF, and MODEL index files (raw CSVs, clean CSVs, log, state files). Continue?",
        "Reset All Data", [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Warning)
    if ($result -eq [System.Windows.MessageBoxResult]::Yes) {
        @($script:RawCSV, $script:CleanCSV, $script:LogFile, $script:StateFile,
          $script:DxfRawCSV, $script:DxfCleanCSV, $script:DxfStateFile,
          $script:ModelRawCSV, $script:ModelAllCSV, $script:ModelCleanCSV, $script:ModelStateFile) | ForEach-Object {
            if (Test-Path $_) { Remove-Item $_ -Force }
        }
        $StatTotalPDFs.Text="--";$StatUniqueParts.Text="--";$StatDuplicates.Text="--";$StatErrors.Text="--"
        $StatTotalDXFs.Text="--";$StatUniqueDxfParts.Text="--";$StatDxfDuplicates.Text="--";$StatDxfErrors.Text="--"
        $StatTotalModels.Text="--";$StatModelAssemblies.Text="--";$StatModelParts.Text="--";$StatModelErrors.Text="--"
        $TxtLog.Text=""; $TxtLastCrawl.Text="Last crawl: Never"
        Add-LogEntry "All index data has been reset" "WARN"
    }
})

# ============================================================================
#  STARTUP
# ============================================================================

Update-Stats
Update-DxfStats
Update-ModelStats
Add-LogEntry "Drawing Index Manager started" "INFO"
Add-LogEntry "PDF index: $($script:CleanCSV)" "INFO"
Add-LogEntry "DXF index: $($script:DxfCleanCSV)" "INFO"
Add-LogEntry "MODEL index: $($script:ModelCleanCSV)" "INFO"

if ($BOMFile -ne "" -and (Test-Path $BOMFile)) {
    $TxtBOMPath.Text = $BOMFile; Add-LogEntry "Pre-loaded BOM: $BOMFile" "INFO"
    $TabCollect.IsSelected = $true
}
if ($CollectOutput -ne "") { $TxtCollectOutput.Text = $CollectOutput }

$window.ShowDialog() | Out-Null

# Cleanup
foreach ($item in @($script:BgPowerShell, $script:BgPDMPowerShell, $script:ModelBgPowerShell)) { if($item){try{$item.Stop()}catch{};try{$item.Dispose()}catch{}} }
foreach ($item in @($script:BgRunspace, $script:BgPDMRunspace, $script:ModelBgRunspace)) { if($item){try{$item.Close()}catch{};try{$item.Dispose()}catch{}} }
foreach ($p in @($script:BgPool, $script:DxfBgPool)) { if($p){try{$p.Close()}catch{};try{$p.Dispose()}catch{}} }
foreach ($list in @($script:ParallelWorkers, $script:DxfParallelWorkers)) {
    if ($list) { foreach ($w in $list) { try{$w.PS.Stop()}catch{};try{$w.PS.Dispose()}catch{} } }
}
