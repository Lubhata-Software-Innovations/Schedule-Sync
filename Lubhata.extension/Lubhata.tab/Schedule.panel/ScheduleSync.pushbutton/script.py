# -*- coding: utf-8 -*-
"""
Lubhata Schedule Sync
=====================
Professional Data Synchronization Tool

BRIDGE THE GAP BETWEEN REVIT & EXCEL:
1. EXPORT: Select any schedule. The tool intelligently extracts Element IDs (even if hidden) to ensure a stable data link.
2. EDIT: Open the CSV in Excel. Modify text, numbers, and parameters freely. (Note: Do NOT modify the 'ElementId' column).
3. SYNC: Select the CSV. The tool maps data back to your elements, validating data types automatically.

Copyright (c) 2026 Lubhata Software & Innovations
"""
__title__ = "Schedule Sync"
__author__ = "Lubhata Software & Innovations"
__helpurl__ = "https://lubhata.com"
__min_revit_ver__ = 2019
__clean_engine__ = True

import clr
import os
import csv
import webbrowser  # For opening links
from pyrevit import revit, DB, forms

# Load WPF Assemblies
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

# --- EMBEDDED XAML UI ---
XAML_SOURCE = """
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Lubhata Schedule Sync" Height="650" Width="600"
        WindowStartupLocation="CenterScreen"
        Background="#1e1e1e">

    <Window.Resources>
        <!-- Modern Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#007ACC"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" CornerRadius="4">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#009BE0"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Link Button Style (Look like hyperlinks) -->
        <Style x:Key="LinkButton" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="#007ACC"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="FontSize" Value="11"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Foreground" Value="#4da6ff"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Text Style -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#dddddd"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
    </Window.Resources>

    <Grid Margin="25">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- Header with Vector Logo -->
        <StackPanel Grid.Row="0" Margin="0,0,0,20" Orientation="Horizontal" VerticalAlignment="Center">
            
            <!-- Embedded 'Schedule + Sync' Logo -->
            <Viewbox Width="50" Height="50" Margin="0,0,20,0">
                <Canvas Width="100" Height="100">
                     <!-- Sync Ring (Dark Arrow) -->
                    <Path Data="M50 10C72.5 10 90 28.5 90 51C90 60.5 86.5 69 80.5 75.5" Stroke="#333333" StrokeThickness="5" StrokeStartLineCap="Round"/>
                    <Path Data="M80.5 75.5L87 67M80.5 75.5L72 69" Stroke="#333333" StrokeThickness="5" StrokeStartLineCap="Round" StrokeLineJoin="Round"/>

                    <!-- Sync Ring (Blue Arrow) -->
                    <Path Data="M50 90C27.5 90 10 71.5 10 49C10 39.5 13.5 31 19.5 24.5" Stroke="#4da6ff" StrokeThickness="5" StrokeStartLineCap="Round"/>
                    <Path Data="M19.5 24.5L13 33M19.5 24.5L28 31" Stroke="#4da6ff" StrokeThickness="5" StrokeStartLineCap="Round" StrokeLineJoin="Round"/>

                    <!-- Central Paper Icon -->
                    <Rectangle Canvas.Left="32" Canvas.Top="28" Width="36" Height="44" RadiusX="4" RadiusY="4" Fill="White"/>
                    <!-- Header Strip -->
                    <Path Data="M32 32C32 29.79 33.79 28 36 28H64C66.21 28 68 29.79 68 32V38H32V32Z" Fill="#007ACC"/>
                    <!-- Data Lines -->
                    <Rectangle Canvas.Left="36" Canvas.Top="44" Width="28" Height="3" RadiusX="1.5" RadiusY="1.5" Fill="#E1E1E1"/>
                    <Rectangle Canvas.Left="36" Canvas.Top="52" Width="28" Height="3" RadiusX="1.5" RadiusY="1.5" Fill="#E1E1E1"/>
                    <Rectangle Canvas.Left="36" Canvas.Top="60" Width="20" Height="3" RadiusX="1.5" RadiusY="1.5" Fill="#E1E1E1"/>
                    <!-- Active Cell -->
                    <Rectangle Canvas.Left="58" Canvas.Top="58" Width="6" Height="6" RadiusX="1" RadiusY="1" Fill="#4da6ff"/>
                </Canvas>
            </Viewbox>

            <StackPanel VerticalAlignment="Center">
                <TextBlock Text="Lubhata Schedule Sync" FontSize="26" FontWeight="Bold" Foreground="White" Margin="0,0,0,2"/>
                <TextBlock Text="Professional Revit Data Management" FontSize="13" Foreground="#888888"/>
            </StackPanel>
        </StackPanel>

        <!-- Main Content Area -->
        <TabControl Grid.Row="1" Background="Transparent" BorderThickness="0">
            <TabControl.Resources>
                <Style TargetType="TabItem">
                    <Setter Property="Template">
                        <Setter.Value>
                            <ControlTemplate TargetType="TabItem">
                                <Border Name="Border" BorderThickness="0,0,0,3" BorderBrush="Transparent" Margin="0,0,15,0" Padding="12,8">
                                    <ContentPresenter x:Name="ContentSite" VerticalAlignment="Center" HorizontalAlignment="Center" ContentSource="Header"/>
                                </Border>
                                <ControlTemplate.Triggers>
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter TargetName="Border" Property="BorderBrush" Value="#007ACC" />
                                        <Setter Property="Foreground" Value="#007ACC"/>
                                        <Setter Property="FontWeight" Value="Bold"/>
                                    </Trigger>
                                    <Trigger Property="IsSelected" Value="False">
                                        <Setter Property="Foreground" Value="#888888"/>
                                    </Trigger>
                                </ControlTemplate.Triggers>
                            </ControlTemplate>
                        </Setter.Value>
                    </Setter>
                </Style>
            </TabControl.Resources>

            <!-- EXPORT TAB -->
            <TabItem Header="EXPORT">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Select Schedule to Export:" Foreground="#CCCCCC" Margin="0,0,0,10"/>
                    
                    <Border Grid.Row="1" BorderThickness="1" BorderBrush="#333" CornerRadius="4">
                        <ListBox x:Name="schedule_list" Background="#252526" BorderThickness="0" Foreground="White" SelectionMode="Single">
                            <ListBox.ItemContainerStyle>
                                <Style TargetType="ListBoxItem">
                                    <Setter Property="Padding" Value="10,6"/>
                                    <Setter Property="Template">
                                        <Setter.Value>
                                            <ControlTemplate TargetType="ListBoxItem">
                                                <Border Name="Border" Background="Transparent" Padding="{TemplateBinding Padding}">
                                                    <ContentPresenter/>
                                                </Border>
                                                <ControlTemplate.Triggers>
                                                    <Trigger Property="IsSelected" Value="True">
                                                        <Setter TargetName="Border" Property="Background" Value="#333"/>
                                                        <Setter Property="Foreground" Value="#4da6ff"/>
                                                    </Trigger>
                                                    <Trigger Property="IsMouseOver" Value="True">
                                                        <Setter TargetName="Border" Property="Background" Value="#2D2D30"/>
                                                    </Trigger>
                                                </ControlTemplate.Triggers>
                                            </ControlTemplate>
                                        </Setter.Value>
                                    </Setter>
                                </Style>
                            </ListBox.ItemContainerStyle>
                        </ListBox>
                    </Border>

                    <Button x:Name="btn_export" Grid.Row="2" Content="EXPORT TO CSV" Height="45" Margin="0,20,0,0" FontSize="14"/>
                </Grid>
            </TabItem>

            <!-- IMPORT TAB -->
            <TabItem Header="IMPORT">
                <Grid Margin="0,20,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <TextBlock Grid.Row="0" Text="Select Modified CSV File:" Foreground="#CCCCCC" Margin="0,0,0,10"/>

                    <Grid Grid.Row="1">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="90"/>
                        </Grid.ColumnDefinitions>
                        <TextBox x:Name="txt_filepath" Grid.Column="0" Height="35" VerticalContentAlignment="Center" Background="#252526" Foreground="White" BorderBrush="#3e3e42" Padding="8,0"/>
                        <Button x:Name="btn_browse" Grid.Column="1" Content="BROWSE" Margin="10,0,0,0" FontSize="11"/>
                    </Grid>

                    <Border Grid.Row="2" Background="#252526" Margin="0,20,0,0" CornerRadius="4" Padding="20">
                        <StackPanel>
                            <TextBlock Text="SYNC INSTRUCTIONS" FontWeight="Bold" Foreground="#666" FontSize="11" Margin="0,0,0,15"/>
                            
                            <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="1." FontWeight="Bold" Width="20" Foreground="#007ACC"/>
                                <TextBlock Text="Ensure CSV retains the 'ElementId' column." TextWrapping="Wrap" Width="450"/>
                            </StackPanel>
                             <StackPanel Orientation="Horizontal" Margin="0,0,0,10">
                                <TextBlock Text="2." FontWeight="Bold" Width="20" Foreground="#007ACC"/>
                                <TextBlock Text="Modify values in Excel. Do not rename headers." TextWrapping="Wrap" Width="450"/>
                            </StackPanel>
                             <StackPanel Orientation="Horizontal">
                                <TextBlock Text="3." FontWeight="Bold" Width="20" Foreground="#007ACC"/>
                                <TextBlock Text="Calculated values &amp; read-only params are skipped." TextWrapping="Wrap" Width="450" Foreground="#999"/>
                            </StackPanel>
                        </StackPanel>
                    </Border>

                    <Button x:Name="btn_import" Grid.Row="3" Content="SYNC DATA TO REVIT" Height="45" Margin="0,20,0,0" Background="#2e7d32" FontSize="14"/>
                </Grid>
            </TabItem>
        </TabControl>

        <!-- Footer / Legal -->
        <Border Grid.Row="2" Margin="0,20,0,0" BorderThickness="0,1,0,0" BorderBrush="#333" Padding="0,15,0,0">
            <StackPanel HorizontalAlignment="Center">
                <TextBlock Text="© 2026 Lubhata Software &amp; Innovations" HorizontalAlignment="Center" FontWeight="SemiBold" Foreground="#888" FontSize="11"/>
                <TextBlock
                    Text="Permission is granted for personal or internal organizational use only. Redistribution, modification for redistribution, sublicensing, resale, or disclosure of the source code or related materials is expressly prohibited."
                    TextWrapping="Wrap"
                    TextAlignment="Center"
                    HorizontalAlignment="Center"
                    Foreground="#FF0000"
                    FontSize="10"
                    Margin="0,2,0,10"
                    MaxWidth="400"/>
                
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Center">
                    <Button x:Name="btn_web" Content="lubhata.com" Style="{StaticResource LinkButton}" />
                    <TextBlock Text=" • " Foreground="#444" Margin="10,0" VerticalAlignment="Center"/>
                    <Button x:Name="btn_linkedin" Content="LinkedIn" Style="{StaticResource LinkButton}" />
                    <TextBlock Text=" • " Foreground="#444" Margin="10,0" VerticalAlignment="Center"/>
                    <Button x:Name="btn_email" Content="Contact Support" Style="{StaticResource LinkButton}" />
                </StackPanel>
            </StackPanel>
        </Border>
    </Grid>
</Window>
"""

class ScheduleSyncWindow(forms.WPFWindow):
    def __init__(self):
        try:
            # Initialize with literal_string=True to use the embedded variable
            forms.WPFWindow.__init__(self, XAML_SOURCE, literal_string=True)
            
            # Wire up Events manually
            self.btn_browse.Click += self.browse_click
            self.btn_export.Click += self.export_click
            self.btn_import.Click += self.import_click
            
            # Wire up Footer Links
            self.btn_web.Click += self.website_click
            self.btn_linkedin.Click += self.linkedin_click
            self.btn_email.Click += self.email_click
            
            # Revit Data
            self.doc = revit.doc
            self.uidoc = revit.uidoc
            
            # Populate Schedule List
            self.schedules = self.get_schedules()
            self.schedule_list.ItemsSource = [s.Name for s in self.schedules]
            
            # Default CSV path
            self.csv_path = None
        except Exception as e:
            forms.alert("Initialization Error: \n" + str(e))

    def get_schedules(self):
        """Collects all Schedule Views"""
        col = DB.FilteredElementCollector(self.doc)\
                .OfClass(DB.ViewSchedule)\
                .WhereElementIsNotElementType()
        
        valid_schedules = []
        for s in col:
            if not s.IsTemplate:
                # Filter out Revision Schedules or other internal types based on name pattern
                if "<" not in s.Name: 
                    valid_schedules.append(s)
                    
        return sorted(valid_schedules, key=lambda x: x.Name)

    def get_safe_id(self, element_id):
        """Extracts integer ID safely for Revit 2023 and 2024+"""
        if hasattr(element_id, "Value"):
            return str(element_id.Value)
        return str(element_id.IntegerValue)

    # --- UI EVENTS ---

    def website_click(self, sender, args):
        webbrowser.open("https://lubhata.com")

    def linkedin_click(self, sender, args):
        webbrowser.open("https://www.linkedin.com/company/lubhata-software-and-innovations/?viewAsMember=true")

    def email_click(self, sender, args):
        webbrowser.open("mailto:info@lubhata.com")

    def browse_click(self, sender, args):
        path = forms.pick_file(file_ext='csv')
        if path:
            self.csv_path = path
            self.txt_filepath.Text = path

    def export_click(self, sender, args):
        selected_name = self.schedule_list.SelectedItem
        if not selected_name:
            forms.alert("Please select a schedule to export.")
            return

        schedule_view = next((s for s in self.schedules if s.Name == selected_name), None)
        
        if schedule_view:
            save_path = forms.save_file(file_ext='csv', default_name=selected_name)
            if save_path:
                self.run_smart_export(schedule_view, save_path)
                self.Close()

    def import_click(self, sender, args):
        if not self.csv_path or not os.path.exists(self.csv_path):
            forms.alert("Please select a valid CSV file.")
            return
            
        self.run_smart_import(self.csv_path)
        self.Close()

    # --- CORE LOGIC ---

    def run_smart_export(self, view, path):
        try:
            elements = DB.FilteredElementCollector(self.doc, view.Id).ToElements()
            definition = view.Definition
            field_order = definition.GetFieldOrder()
            
            headers = ["ElementId"]
            fields_data = [] 
            
            for field_id in field_order:
                field = definition.GetField(field_id)
                if not field.IsHidden:
                    headers.append(field.ColumnHeading or field.GetName())
                    fields_data.append((field.ParameterId, field.GetName()))
            
            with open(path, 'w') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                
                for el in elements:
                    row = [self.get_safe_id(el.Id)]
                    for pid, name in fields_data:
                        val = ""
                        if pid.IntegerValue < 0:
                            param = el.get_Parameter(DB.BuiltInParameter(pid.IntegerValue))
                        else:
                            param = el.LookupParameter(name)
                            
                        if param:
                            if param.StorageType == DB.StorageType.String:
                                val = param.AsString()
                            else:
                                val = param.AsValueString()
                        
                        row.append(val if val else "")
                    writer.writerow(row)
            
            forms.alert("Export Successful!\nSaved to: " + path)
            
        except Exception as e:
            forms.alert("Export Failed:\n" + str(e))

    def run_smart_import(self, path):
        try:
            with open(path, 'r') as f:
                reader = csv.DictReader(f)
                if 'ElementId' not in reader.fieldnames:
                    forms.alert("Error: CSV must contain 'ElementId' column.")
                    return

                with revit.Transaction("Lubhata Schedule Sync"):
                    count = 0
                    errors = 0
                    for row in reader:
                        eid_str = row['ElementId']
                        if not eid_str: continue

                        try:
                            eid = DB.ElementId(int(eid_str))
                            el = self.doc.GetElement(eid)
                            
                            if el:
                                for header, value in row.items():
                                    if header == 'ElementId': continue
                                    param = el.LookupParameter(header)
                                    if param and not param.IsReadOnly:
                                        if param.StorageType == DB.StorageType.String:
                                            param.Set(str(value))
                                        elif param.StorageType == DB.StorageType.Double:
                                            try:
                                                param.SetValueString(value)
                                            except:
                                                pass
                                        elif param.StorageType == DB.StorageType.Integer:
                                            if value and value.isdigit():
                                                param.Set(int(value))
                                        count += 1
                        except:
                            errors += 1
                            
                    forms.alert("Sync Complete.\nUpdated {} parameters.".format(count))

        except Exception as e:
            forms.alert("Import Failed:\n" + str(e))

# Run the window
try:
    ScheduleSyncWindow().ShowDialog()
except Exception as e:
    pass