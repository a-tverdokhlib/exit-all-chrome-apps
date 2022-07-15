<#
    .SYNOPSIS
        UI that will display the history of clipboard items

    .DESCRIPTION
        UI that will display the history of clipboard items. Options include filtering for text by
        typing into the filter textbox, context menu for removing and copying text as well as a menu to 
        clear all entries in the clipboard and clipboard history viewer.

        Use keyboard shortcuts to run common commands:

        Ctrl + C -> Copy selected text from viewer
        Ctrl + R -> Remove selected text from viewer


        
        Ctrl + E -> Exit the clipboard viewer

    .NOTES
        Author: Boe Prox
        Created: 10 July 2014
        Version History:
            1.0 - Boe Prox - 10 July 2014
                -Initial Version
            1.1 - Boe Prox - 24 July 2014
                -Moved Filter from timer to TextChanged Event
                -Add capability to select multiple items to remove or add to clipboard
                -Able to now use mouse scroll wheel to scroll when over listbox
                - Added Keyboard shortcuts for common operations (copy, remove and exit)
#>
#Requires -Version 3.0
$Runspacehash = [hashtable]::Synchronized(@{})
$Runspacehash.Host = $Host
$Runspacehash.runspace = [RunspaceFactory]::CreateRunspace()
$Runspacehash.runspace.ApartmentState = "STA"
$Runspacehash.runspace.Open()
$Runspacehash.runspace.SessionStateProxy.SetVariable("Runspacehash",$Runspacehash)
$Runspacehash.PowerShell = {Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase}.GetPowerShell()
$Runspacehash.PowerShell.Runspace = $Runspacehash.runspace
$Runspacehash.Handle = $Runspacehash.PowerShell.AddScript({
    Function Get-ClipBoard {
        [Windows.Clipboard]::GetText()
    }
    Function Set-ClipBoard {
        $Script:CopiedText = @"
$($listbox.SelectedItems | Out-String)
"@
        [Windows.Clipboard]::SetText($Script:CopiedText)
                
    }
    Function Clear-Viewer {
        [void]$Script:ObservableCollection.Clear()
        [Windows.Clipboard]::Clear()
        write-output "Hrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrrr"
    }
    
    #Build the GUI
    [xml]$xaml = @"
    <Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        x:Name="Window" Title="Clipboard History Viewer" WindowStartupLocation = "CenterScreen" 
        Width = "350" Height = "425" ShowInTaskbar = "True" Background = "White">
        <Grid >
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto" />
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
            </Grid.RowDefinitions>
            <Grid.Resources>
                <Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
                    <Setter Property="Background" Value="LightGray"/>
                    <Setter Property="Foreground" Value="Black"/>
                    <Style.Triggers>
                        <Trigger Property="ItemsControl.AlternationIndex" Value="1">
                            <Setter Property="Background" Value="White"/>
                            <Setter Property="Foreground" Value="Black"/>
                        </Trigger>
                    </Style.Triggers>
                </Style>
            </Grid.Resources>
            <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
                <Menu.Background>
                    <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                        <LinearGradientBrush.GradientStops> 
                        <GradientStop Color='#C4CBD8' Offset='0' /> 
                        <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                        <GradientStop Color='#CFD7E2' Offset='0.9' /> 
                        <GradientStop Color='#C4CBD8' Offset='1' /> 
                        </LinearGradientBrush.GradientStops>
                    </LinearGradientBrush>
                </Menu.Background>
                <MenuItem x:Name = 'FileMenu' Header = '_Tools'>
                    <MenuItem x:Name = 'Clear_Menu' Header = '_Clear' />
                    <MenuItem x:Name = 'Save_Menu'  Header = '_Save As File'/>
                    <MenuItem x:Name = 'StayTop_Menu' Header = '_Stay On Top' IsCheckable="true"/>
                    <MenuItem x:Name = 'AddTime_Menu' Header = '_Add Time Stamp' IsCheckable="true"/>
                    <MenuItem x:Name = 'Pause_Menu' Header = '_Pause' IsCheckable="true"/>
                </MenuItem>
            </Menu>
            <GroupBox Header = "Filter"  Grid.Row = '2' Background = "White">
                <TextBox x:Name="InputBox" Height = "25" Grid.Row="2" />
            </GroupBox>
            <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" Grid.Row="4" Height = "Auto">                
                <ListBox x:Name="listbox" AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}" SelectionMode='Extended'>                
                    <ListBox.Template>
                        <ControlTemplate TargetType="ListBox">
                            <Border BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderBrush}">
                                <ItemsPresenter/>
                            </Border>
                        </ControlTemplate>
                    </ListBox.Template>
                    <ListBox.ContextMenu>
                        <ContextMenu x:Name = 'ClipboardMenu'>
                            <MenuItem x:Name = 'Copy_Menu' Header = 'Copy'/>
                            <MenuItem x:Name = 'Edit_Menu' Header = 'Edit'/>
                            <MenuItem x:Name = 'Remove_Menu' Header = 'Remove'/>
                            <MenuItem x:Name = 'SelectAll_Menu' Header = 'Select All'/>                              
                        </ContextMenu>
                    </ListBox.ContextMenu>
                </ListBox>
            </ScrollViewer>  
            <TextBox x:Name="editBox" Height = "20" Margin = "0, -300, 0, 0" Grid.Row="4" Visibility="hidden"/>
            <TextBox x:Name="indexBox" Height = "10" Grid.Row="5" Visibility="hidden"/>          
        </Grid>
        
    </Window>
"@
 
 # 3/29

    $reader=(New-Object System.Xml.XmlNodeReader $xaml)
    $Window=[Windows.Markup.XamlReader]::Load( $reader )

    #Connect to Controls
    $listbox = $Window.FindName('listbox')
    $InputBox = $Window.FindName('InputBox')
    $ItemEdit = $Window.FindName('editBox')
    $indexBox = $Window.FindName('indexBox')
    $Copy_Menu = $Window.FindName('Copy_Menu')
    $Edit_Menu = $Window.FindName('Edit_Menu')
    $Remove_Menu = $Window.FindName('Remove_Menu')
    $Clear_Menu = $Window.FindName('Clear_Menu')
    $SelectAll_Menu = $Window.FindName('SelectAll_Menu')
    $Save_Menu = $Window.FindName('Save_Menu')
    $StayTop_Menu = $Window.FindName('StayTop_Menu')
    $AddTime_Menu = $Window.FindName('AddTime_Menu')
    $Pause_Menu = $Window.FindName('Pause_Menu')
       

    #Events
    $Clear_Menu.Add_Click({
        Clear-Viewer
    })

    $SelectAll_Menu.Add_Click({
        $listbox.SelectAll()
    })

    
    $Edit_Menu.Add_Click({
        $pos = $listbox.SelectedIndex;
        $indexBox.Text = $pos
        $ItemEdit.Text = $listbox.SelectedItems[0]
        $ItemEdit.Visibility = "visible"
        $marginText = "0, " + ($pos * 40 - 300) + ", 0, 0"
        $ItemEdit.Margin = $marginText
        $ItemEdit.Focus()
    })

    $Remove_Menu.Add_Click({
        @($listbox.SelectedItems) | ForEach {
            [void]$Script:ObservableCollection.Remove($_)
        }
    })
    
    $Copy_Menu.Add_Click({
        Set-ClipBoard        
    })
    
    $Save_Menu.Add_Click({
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
        $Dialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.initialDirectory = 'D:'
        $NameFilter = "Text Files (*.txt)|*.txt"
        $Dialog.Filter = $NameFilter
        [void]($Dialog.ShowDialog())
        # if($Dialog.ShowDialog() -eq 'Ok'){
        #     $save_data = $listbox.Items
        #     $save_data | Out-File $Dialog.FileName
        # }
        
    })

    $StayTop_Menu.Add_Click({
        if ($Window.Topmost -eq $True) {
            $Window.Topmost = $False
        } else {
            $Window.Topmost = $True
        }
    })
    
    $Window.Add_Activated({
        $InputBox.Focus()
    })

    $Window.Add_SourceInitialized({
        #Create observable collection
        $Script:ObservableCollection = New-Object System.Collections.ObjectModel.ObservableCollection[string]
        
        $Listbox.ItemsSource = $Script:ObservableCollection
        
        #Create Timer object
        $Script:timer = new-object System.Windows.Threading.DispatcherTimer 
        $timer.Interval = [TimeSpan]"0:0:.1"

        #Add event per tick
        $timer.Add_Tick({
            $text =  Get-Clipboard
            If (($Script:Previous -ne $Text -AND $Script:CopiedText -ne $Text) -AND $text.length -gt 0) {
                #Add to collection
                
                if ($Pause_Menu.IsChecked -eq $False) {
                    if ($AddTime_Menu.IsChecked -eq $False) {
                        [void]$Script:ObservableCollection.Add($text)
                    } else {
                        [void]$Script:ObservableCollection.Add($text + " : " + (Get-Date -Format "dddd MM/dd/yyyy HH:mm:ss"))
                    }                    
                }
                $Script:Previous = $text

            }     
        })
        $timer.Start()
        If (-NOT $timer.IsEnabled) {
            $Window.Close()
        }
    })

    $Window.Add_Closed({
        $Script:timer.Stop()
        $Script:ObservableCollection.Clear()
        $Runspacehash.PowerShell.Dispose()
    })

    $InputBox.Add_TextChanged({
        [System.Windows.Data.CollectionViewSource]::GetDefaultView($Listbox.ItemsSource).Filter = [Predicate[Object]]{             
            Try {
                $args[0] -match [regex]::Escape($InputBox.Text)
            } Catch {
                $True
            }
        }    
    })
    
    $listbox.Add_MouseRightButtonUp({
        If ($Script:ObservableCollection.Count -gt 0) {
            $Remove_Menu.IsEnabled = $True
            $Copy_Menu.IsEnabled = $True
        } Else {
            $Remove_Menu.IsEnabled = $False
            $Copy_Menu.IsEnabled = $False
        }
    })

    $Window.Add_KeyDown({ 
        $key = $_.Key  
        If ([System.Windows.Input.Keyboard]::IsKeyDown("RightCtrl") -OR [System.Windows.Input.Keyboard]::IsKeyDown("LeftCtrl")) {
            Switch ($Key) {
            "C" {
                Set-ClipBoard                                          
            }
            "R" {
                @($listbox.SelectedItems) | ForEach {
                    [void]$Script:ObservableCollection.Remove($_)
                }            
            }
            "E" {
                $This.Close()
            }
            Default {$Null}
            }
        }
        If ([System.Windows.Input.Keyboard]::IsKeyDown("Enter")) {
            
            [void]$Script:ObservableCollection.RemoveAt($indexBox.Text)
            [void]$Script:ObservableCollection.Insert($indexBox.Text, $ItemEdit.Text)
            $ItemEdit.Visibility = "hidden"
        }
    })

    [void]$Window.ShowDialog()
}).BeginInvoke()