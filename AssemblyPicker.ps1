# ============================================================================
#  AssemblyPicker.ps1  v3.0
#  WPF tree-view picker for SolidWorks assembly components
#  Nordic Minesteel Technologies
#
#  Args:
#    $ListFile     - node list: first line ROOT|Title, then ID|ParentID|Depth|Label|Path
#    $ResultFile   - output: MODE, JOBNAME, then selected instance paths
#    $PdfIndexPath - (optional) path to pdf_index_clean.csv for status display
#    $DxfIndexPath - (optional) path to dxf_index_clean.csv for status display
#    $ManagerPath  - (optional) path to PDFIndexManager.ps1 for Update Index button
# ============================================================================

param(
    [string]$ListFile,
    [string]$ResultFile,
    [string]$PdfIndexPath = "",
    [string]$DxfIndexPath = "",
    [string]$ManagerPath  = ""
)

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

# --- Read node list ---
if (-not (Test-Path $ListFile)) {
    [System.Windows.MessageBox]::Show("Node list file not found:`n$ListFile", "Error")
    exit 1
}

$rootTitle = "Assembly"
$nodes = [System.Collections.Generic.List[hashtable]]::new()

$lineIndex = 0
foreach ($line in [System.IO.File]::ReadLines($ListFile)) {
    $line = $line.Trim()
    if ($line -eq "") { continue }
    $lineIndex++

    if ($lineIndex -eq 1) {
        # ROOT|Title
        $parts = $line -split '\|', 2
        if ($parts.Count -eq 2) { $rootTitle = $parts[1] }
        continue
    }

    # ID|ParentID|Depth|Label|Path
    $parts = $line -split '\|', 5
    if ($parts.Count -eq 5) {
        $nodes.Add(@{
            ID       = [int]$parts[0]
            ParentID = [int]$parts[1]
            Depth    = [int]$parts[2]
            Label    = $parts[3]
            Path     = $parts[4]
        })
    }
}

if ($nodes.Count -eq 0) {
    [System.Windows.MessageBox]::Show("No components found in assembly.", "Error")
    exit 1
}

# --- Build XAML (Nordic Minesteel Technologies branding) ---
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Nordic Minesteel - $rootTitle"
    Width="560" Height="720"
    MinWidth="400" MinHeight="400"
    WindowStartupLocation="CenterScreen"
    Background="#1a1a2e">

    <Window.Resources>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#e8e8e8"/>
        </Style>
        <Style x:Key="Brand" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#2ea3f2"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="FontWeight" Value="Bold"/>
            <Setter Property="Margin" Value="0,0,0,6"/>
        </Style>
        <Style x:Key="Header" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="FontSize" Value="15"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,0,0,2"/>
        </Style>
        <Style x:Key="Sub" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
        </Style>
        <Style x:Key="SectionLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,10,0,4"/>
        </Style>
        <Style x:Key="PrimaryBtn" TargetType="Button">
            <Setter Property="Background" Value="#2ea3f2"/>
            <Setter Property="Foreground" Value="#ffffff"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Padding" Value="24,9"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="5" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="SecBtn" TargetType="Button">
            <Setter Property="Background" Value="#0f3460"/>
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ModeBtn" TargetType="ToggleButton">
            <Setter Property="Background" Value="#16213e"/>
            <Setter Property="Foreground" Value="#a0a0b8"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ToggleButton">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter Property="Background" Value="#2ea3f2"/>
                                <Setter Property="Foreground" Value="#ffffff"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#16213e"/>
            <Setter Property="Foreground" Value="#e8e8e8"/>
            <Setter Property="BorderBrush" Value="#0f3460"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="CaretBrush" Value="#e8e8e8"/>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="#16213e"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header with branding -->
        <Border Grid.Row="0" Padding="20,16,20,0">
            <StackPanel>
                <TextBlock Text="NORDIC MINESTEEL TECHNOLOGIES" Style="{StaticResource Brand}"/>
                <TextBlock x:Name="TxtRootTitle" Text="Assembly" Style="{StaticResource Header}"/>
                <TextBlock Text="Check the sub-assemblies to include. Checking a sub-assembly includes all its parts." Style="{StaticResource Sub}"/>
                <Grid Margin="0,0,0,8">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="BtnSelectAll"  Grid.Column="0" Content="Select All"  Style="{StaticResource SecBtn}" Margin="0,0,6,0"/>
                    <Button x:Name="BtnSelectNone" Grid.Column="1" Content="Select None" Style="{StaticResource SecBtn}"/>
                    <TextBox x:Name="TxtSearch" Grid.Column="2" Margin="12,0,0,0"
                             FontSize="11" Padding="6,5"
                             Tag="Search components..."/>
                </Grid>
            </StackPanel>
        </Border>

        <!-- Tree area -->
        <Border Grid.Row="1" Margin="20,8,20,0" Background="#16213e" CornerRadius="6">
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Padding="4">
                <StackPanel x:Name="TreePanel" Margin="8,8,8,8"/>
            </ScrollViewer>
        </Border>

        <!-- Footer controls -->
        <Border Grid.Row="2" Padding="20,12,20,18">
            <StackPanel>
                <TextBlock Text="OUTPUT FOLDER NAME" Style="{StaticResource SectionLabel}"/>
                <TextBox x:Name="TxtJobName" Margin="0,0,0,10"/>

                <TextBlock Text="FILE TYPE" Style="{StaticResource SectionLabel}"/>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                    <ToggleButton x:Name="BtnPdfOnly" Content="&#x1F4C4; PDF Only"        IsChecked="True" Style="{StaticResource ModeBtn}" Margin="0,0,6,0"/>
                    <ToggleButton x:Name="BtnDxfOnly" Content="&#x1F4D0; DXF Only"        Style="{StaticResource ModeBtn}" Margin="0,0,6,0"/>
                    <ToggleButton x:Name="BtnBothFiles" Content="&#x1F4C4;+&#x1F4D0; Both" Style="{StaticResource ModeBtn}"/>
                </StackPanel>

                <TextBlock Text="INDEX STATUS" Style="{StaticResource SectionLabel}"/>
                <StackPanel Margin="0,0,0,6">
                    <TextBlock x:Name="TxtPdfIndexStatus" Text="PDF Index: checking..." Foreground="#6c7086" FontSize="10"/>
                    <TextBlock x:Name="TxtDxfIndexStatus" Text="DXF Index: checking..." Foreground="#6c7086" FontSize="10" Margin="0,2,0,0"/>
                </StackPanel>

                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,4,0,0">
                    <Button x:Name="BtnUpdateIndex" Content="&#x1F504; Update Index" Style="{StaticResource SecBtn}" Margin="0,0,8,0"/>
                    <Button x:Name="BtnCancel" Content="Cancel"          Style="{StaticResource SecBtn}"    Margin="0,0,8,0"/>
                    <Button x:Name="BtnOK"     Content="Collect Drawings" Style="{StaticResource PrimaryBtn}"/>
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

$reader = [System.Xml.XmlNodeReader]::new($xaml)
$window = [System.Windows.Markup.XamlReader]::Load($reader)

$TxtRootTitle      = $window.FindName("TxtRootTitle")
$TreePanel         = $window.FindName("TreePanel")
$TxtJobName        = $window.FindName("TxtJobName")
$TxtSearch         = $window.FindName("TxtSearch")
$BtnPdfOnly        = $window.FindName("BtnPdfOnly")
$BtnDxfOnly        = $window.FindName("BtnDxfOnly")
$BtnBothFiles      = $window.FindName("BtnBothFiles")
$BtnOK             = $window.FindName("BtnOK")
$BtnCancel         = $window.FindName("BtnCancel")
$BtnSelectAll      = $window.FindName("BtnSelectAll")
$BtnSelectNone     = $window.FindName("BtnSelectNone")
$BtnUpdateIndex    = $window.FindName("BtnUpdateIndex")
$TxtPdfIndexStatus = $window.FindName("TxtPdfIndexStatus")
$TxtDxfIndexStatus = $window.FindName("TxtDxfIndexStatus")

$TxtRootTitle.Text = $rootTitle
$TxtJobName.Text   = $rootTitle

# --- Search box placeholder text ---
$TxtSearch.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#6c7086")
$TxtSearch.Text = $TxtSearch.Tag
$TxtSearch.Add_GotFocus({
    if ($TxtSearch.Text -eq $TxtSearch.Tag) {
        $TxtSearch.Text = ""
        $TxtSearch.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#e8e8e8")
    }
})
$TxtSearch.Add_LostFocus({
    if ($TxtSearch.Text -eq "") {
        $TxtSearch.Text = $TxtSearch.Tag
        $TxtSearch.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom("#6c7086")
    }
})

# --- Check index status and show it to the user ---
function Get-IndexStatus {
    param([string]$IndexPath, [string]$Label)
    if ($IndexPath -eq "" -or -not (Test-Path $IndexPath)) {
        return @{ Text = "$Label  Not found - run Index Manager"; Color = "#f44336"; OK = $false }
    }
    $file = Get-Item $IndexPath
    $lineCount = 0
    try { $lineCount = (Get-Content $IndexPath -ReadCount 0).Count - 1 } catch { $lineCount = 0 }
    $age = (Get-Date) - $file.LastWriteTime
    $dateStr = $file.LastWriteTime.ToString("MMM d, yyyy h:mm tt")
    if ($age.TotalDays -gt 30) {
        return @{ Text = "$Label  $($lineCount.ToString('N0')) entries (crawled $dateStr) - STALE"; Color = "#ffc107"; OK = $true }
    }
    return @{ Text = "$Label  $($lineCount.ToString('N0')) entries (crawled $dateStr)"; Color = "#4caf50"; OK = $true }
}

$pdfStatus = Get-IndexStatus $PdfIndexPath "PDF Index:"
$dxfStatus = Get-IndexStatus $DxfIndexPath "DXF Index:"

$TxtPdfIndexStatus.Text = $pdfStatus.Text
$TxtPdfIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($pdfStatus.Color)
$TxtDxfIndexStatus.Text = $dxfStatus.Text
$TxtDxfIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($dxfStatus.Color)

# Hide Update Index button if no manager path provided
if ($ManagerPath -eq "" -or -not (Test-Path $ManagerPath)) {
    $BtnUpdateIndex.Visibility = [System.Windows.Visibility]::Collapsed
}

# --- Build a lookup: ID -> node ---
$nodeMap = @{}
foreach ($nd in $nodes) { $nodeMap[$nd.ID] = $nd }

# --- Build a lookup: ID -> list of child IDs ---
$childMap = @{}
foreach ($nd in $nodes) {
    $parentId = $nd.ParentID
    if (-not $childMap.ContainsKey($parentId)) { $childMap[$parentId] = [System.Collections.Generic.List[int]]::new() }
    $childMap[$parentId].Add($nd.ID)
}

# --- Create checkbox controls, store in map ---
$cbMap  = @{}   # nodeID -> CheckBox
$rowMap = @{}   # nodeID -> Grid row (for search filtering)

# Cascade-down helper: check/uncheck all descendants
$script:Cascading = $false

function Set-ChildrenChecked {
    param([int]$parentID, [bool]$isChecked)
    if (-not $childMap.ContainsKey($parentID)) { return }
    foreach ($cid in $childMap[$parentID]) {
        if ($cbMap.ContainsKey($cid)) {
            $cbMap[$cid].IsChecked = $isChecked
        }
        Set-ChildrenChecked -parentID $cid -isChecked $isChecked
    }
}

# Build rows for all nodes, indented by depth
foreach ($nd in $nodes) {
    $depth  = $nd.Depth
    $indent = $depth * 20   # 20px per level

    # Determine if this node is a sub-assembly (has children)
    $hasChildren = $childMap.ContainsKey($nd.ID) -and $childMap[$nd.ID].Count -gt 0

    $row = [System.Windows.Controls.Grid]::new()
    $col0 = [System.Windows.Controls.ColumnDefinition]::new()
    $col0.Width = [System.Windows.GridLength]::new($indent)
    $col1 = [System.Windows.Controls.ColumnDefinition]::new()
    $col1.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star)
    $row.ColumnDefinitions.Add($col0)
    $row.ColumnDefinitions.Add($col1)
    $row.Margin = [System.Windows.Thickness]::new(0, 1, 0, 1)

    # Checkbox in column 1
    $cb = [System.Windows.Controls.CheckBox]::new()
    $cb.IsChecked = $false
    $cb.Tag = $nd.Path

    # Label with icon (BMP-safe characters for PS 5.1)
    $icon = if ($hasChildren) { "$([char]0x25B8) " } else { "  " }
    $label = [System.Windows.Controls.TextBlock]::new()
    $label.Text = $icon + $nd.Label
    $label.Foreground = if ($hasChildren) {
        [System.Windows.Media.BrushConverter]::new().ConvertFrom("#e8e8e8")
    } else {
        [System.Windows.Media.BrushConverter]::new().ConvertFrom("#a0a0b8")
    }
    $label.FontSize   = if ($hasChildren) { 12 } else { 11 }
    $label.FontWeight = if ($hasChildren) {
        [System.Windows.FontWeights]::SemiBold
    } else {
        [System.Windows.FontWeights]::Normal
    }
    $label.VerticalAlignment = [System.Windows.VerticalAlignment]::Center
    $label.Margin = [System.Windows.Thickness]::new(4, 0, 0, 0)

    # Tooltip showing full file path
    $label.ToolTip = $nd.Path

    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $sp.Children.Add($cb)      | Out-Null
    $sp.Children.Add($label)   | Out-Null

    [System.Windows.Controls.Grid]::SetColumn($sp, 1)
    $row.Children.Add($sp) | Out-Null

    $TreePanel.Children.Add($row) | Out-Null

    $cbMap[$nd.ID]  = $cb
    $rowMap[$nd.ID] = $row

    # Wire cascade: use GetNewClosure() to capture $nodeID by value at this iteration
    $nodeID = $nd.ID
    $checkedBlock = {
        if (-not $script:Cascading) {
            $script:Cascading = $true
            Set-ChildrenChecked -parentID $nodeID -isChecked $true
            $script:Cascading = $false
        }
    }.GetNewClosure()

    $uncheckedBlock = {
        if (-not $script:Cascading) {
            $script:Cascading = $true
            Set-ChildrenChecked -parentID $nodeID -isChecked $false
            $script:Cascading = $false
        }
    }.GetNewClosure()

    $cb.Add_Checked($checkedBlock)
    $cb.Add_Unchecked($uncheckedBlock)
}

# --- Search / filter tree ---
$TxtSearch.Add_TextChanged({
    $query = $TxtSearch.Text
    if ($query -eq $TxtSearch.Tag -or $query -eq "") {
        # Show all rows
        foreach ($nd in $nodes) {
            if ($rowMap.ContainsKey($nd.ID)) {
                $rowMap[$nd.ID].Visibility = [System.Windows.Visibility]::Visible
            }
        }
        return
    }
    foreach ($nd in $nodes) {
        if ($rowMap.ContainsKey($nd.ID)) {
            if ($nd.Label -like "*$query*") {
                $rowMap[$nd.ID].Visibility = [System.Windows.Visibility]::Visible
            } else {
                $rowMap[$nd.ID].Visibility = [System.Windows.Visibility]::Collapsed
            }
        }
    }
})

# --- Select All / None ---
$BtnSelectAll.Add_Click({
    foreach ($cb in $cbMap.Values) { $cb.IsChecked = $true }
})
$BtnSelectNone.Add_Click({
    foreach ($cb in $cbMap.Values) { $cb.IsChecked = $false }
})

# --- File type toggle (PDF / DXF / Both) - mutually exclusive ---
$BtnPdfOnly.Add_Checked({   $BtnDxfOnly.IsChecked = $false; $BtnBothFiles.IsChecked = $false })
$BtnDxfOnly.Add_Checked({   $BtnPdfOnly.IsChecked = $false; $BtnBothFiles.IsChecked = $false })
$BtnBothFiles.Add_Checked({ $BtnPdfOnly.IsChecked = $false; $BtnDxfOnly.IsChecked   = $false })
# Prevent un-checking all: if the user unchecks the active one, keep it checked
$BtnPdfOnly.Add_Unchecked({   if (-not $BtnDxfOnly.IsChecked -and -not $BtnBothFiles.IsChecked) { $BtnPdfOnly.IsChecked  = $true } })
$BtnDxfOnly.Add_Unchecked({   if (-not $BtnPdfOnly.IsChecked -and -not $BtnBothFiles.IsChecked) { $BtnDxfOnly.IsChecked  = $true } })
$BtnBothFiles.Add_Unchecked({ if (-not $BtnPdfOnly.IsChecked -and -not $BtnDxfOnly.IsChecked)   { $BtnBothFiles.IsChecked = $true } })

# --- Update Index: launch manager, refresh status when it closes ---
$BtnUpdateIndex.Add_Click({
    $window.Title = "Nordic Minesteel - Waiting for Index Manager..."
    $window.IsEnabled = $false
    $window.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::Render)
    Start-Process "PowerShell.exe" -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$ManagerPath`"" -WindowStyle Normal -Wait
    # Refresh index status after manager closes
    $pdfStatus = Get-IndexStatus $PdfIndexPath "PDF Index:"
    $dxfStatus = Get-IndexStatus $DxfIndexPath "DXF Index:"
    $TxtPdfIndexStatus.Text = $pdfStatus.Text
    $TxtPdfIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($pdfStatus.Color)
    $TxtDxfIndexStatus.Text = $dxfStatus.Text
    $TxtDxfIndexStatus.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFrom($dxfStatus.Color)
    $window.Title = "Nordic Minesteel - $rootTitle"
    $window.IsEnabled = $true
}.GetNewClosure())

# --- Keyboard shortcuts: Enter = Collect, Escape = Cancel ---
$window.Add_KeyDown({
    param($sender, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Escape) {
        $window.Close()
    }
    elseif ($e.Key -eq [System.Windows.Input.Key]::Return -and -not $TxtJobName.IsFocused -and -not $TxtSearch.IsFocused) {
        $BtnOK.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent))
    }
})

# --- Cancel ---
$BtnCancel.Add_Click({ $window.Close() })

# --- OK ---
$BtnOK.Add_Click({
    # Collect only TOP-LEVEL checked paths (skip children whose ancestor is also checked)
    # This avoids flooding the result with every individual part when a parent assembly
    # is checked - the macro recurses into children automatically.
    $selected = @()
    foreach ($nd in $nodes) {
        if ($cbMap.ContainsKey($nd.ID) -and $cbMap[$nd.ID].IsChecked -eq $true) {
            # Walk up the parent chain - if any ancestor is also checked, skip this node
            $ancestorChecked = $false
            $checkParentId = $nd.ParentID
            while ($checkParentId -ne 0) {
                if ($cbMap.ContainsKey($checkParentId) -and $cbMap[$checkParentId].IsChecked -eq $true) {
                    $ancestorChecked = $true
                    break
                }
                if ($nodeMap.ContainsKey($checkParentId)) {
                    $checkParentId = $nodeMap[$checkParentId].ParentID
                } else {
                    break
                }
            }
            if (-not $ancestorChecked) {
                $selected += $nd.Path
            }
        }
    }

    if ($selected.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "Please select at least one component.",
            "Nothing Selected",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Warning)
        return
    }

    $collectMode = "QUICK"
    $fileType    = if ($BtnDxfOnly.IsChecked)   { "DXF" } `
                   elseif ($BtnBothFiles.IsChecked) { "BOTH" } `
                   else { "PDF" }
    $jobName = $TxtJobName.Text.Trim()
    if ($jobName -eq "") { $jobName = $TxtRootTitle.Text }

    # Result file format:
    #   Line 1: COLLECT MODE  (QUICK or FULL)
    #   Line 2: FILE TYPE     (PDF, DXF, or BOTH)
    #   Line 3: JOB NAME
    #   Lines 4+: selected component file paths
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add($collectMode)
    $lines.Add($fileType)
    $lines.Add($jobName)
    foreach ($p in $selected) { $lines.Add($p) }

    [System.IO.File]::WriteAllLines($ResultFile, $lines)
    $window.Close()
})

$window.ShowDialog() | Out-Null
