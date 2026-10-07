# -*- coding: utf-8 -*-
__doc__ = "Select rebar by Schedule Mark using a grouped checklist dialog."

from pyrevit import revit, DB, UI, forms, script
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Structure import Rebar
from System.Collections.Generic import List
from collections import OrderedDict
from collections import defaultdict

doc = revit.doc
uidoc = revit.uidoc

# ###### Collect all rebar ##########################################################################################################
all_rebar = FilteredElementCollector(doc).OfClass(Rebar).WhereElementIsNotElementType().ToElements()

if not all_rebar:
    forms.alert("No structural rebar found in the document.", exitscript=True)

# ###### Build dict: bar_mark -> { schedule_mark -> [ElementId, ...] } #######################################################

bar_mark_groups = defaultdict(lambda: defaultdict(list))

for bar in all_rebar:
    # Schedule Mark
    p_sched = bar.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK)
    if p_sched is None:
        p_sched = bar.LookupParameter("Schedule Mark")
    sched_mark = (p_sched.AsString() if (p_sched and p_sched.AsString()) else "<No Schedule Mark>")

    # Bar Mark — used as the group header
    p_bar = bar.LookupParameter("Bar Mark")
    if p_bar is None:
        p_bar = bar.LookupParameter("Mark")
    bar_mark = (p_bar.AsString() if (p_bar and p_bar.AsString()) else "<No Bar Mark>")

    bar_mark_groups[bar_mark][sched_mark].append(bar.Id)

display_to_ids = {}  

# Sort group headers; push <No Bar Mark> to the end########################################################################################
sorted_bar_marks = sorted(
    bar_mark_groups.keys(),
    key=lambda m: (m.startswith("<"), m)
)

grouped_items = OrderedDict()

for bar_mark in sorted_bar_marks:
    sched_marks = bar_mark_groups[bar_mark]

    # Sort schedule marks within each group; push <No Schedule Mark> to end
    sorted_sched_marks = sorted(
        sched_marks.keys(),
        key=lambda m: (m.startswith("<"), m)
    )

    display_list = []
    for sched_mark in sorted_sched_marks:
        ids = sched_marks[sched_mark]
        count = len(ids)
        display = "{} ({} bar{})".format(sched_mark, count, "s" if count != 1 else "")
        display_list.append(display)
        display_to_ids[display] = ids

    grouped_items[bar_mark] = display_list

# ######  Show grouped checklist ##############################################################################################
selected_display = forms.SelectFromList.show(
    grouped_items,
    title="Select Rebar by Schedule Mark",
    button_name="Select",
    multiselect=True,
    message="Tick the schedule marks to select:",
    group_selector_title="Bar Mark"
)

if not selected_display:
    script.exit()

# ###### Collect ElementIds for all ticked items #############################################################################
ids_to_select = []
for disp in selected_display:
    ids_to_select.extend(display_to_ids[disp])

# ######  Set selection #######################################################################################################
id_collection = List[ElementId](ids_to_select)
uidoc.Selection.SetElementIds(id_collection)
uidoc.ShowElements(id_collection)








# # -*- coding: utf-8 -*-
# __title__ = "Select\nBy Mark"
# __doc__ = "Select rebar by Schedule Mark using a multi-column WPF grid."

# import clr
# clr.AddReference('PresentationFramework')
# clr.AddReference('PresentationCore')
# clr.AddReference('WindowsBase')

# clr.AddReference('System.Xml')
# from System.IO import MemoryStream, Stream
# from System.Text import Encoding
# from System.Windows.Markup import XamlReader

# from pyrevit import revit, DB, forms, script
# from Autodesk.Revit.DB import *
# from Autodesk.Revit.DB.Structure import Rebar
# from System.Collections.Generic import List
# from System.Collections.ObjectModel import ObservableCollection
# from collections import defaultdict

# # System.Windows — layout primitives
# from System.Windows import Window, Thickness, GridLength

# # System.Windows.Controls — UI controls only
# from System.Windows.Controls import (
#     CheckBox, TextBlock, DataGrid,
#     DataGridTextColumn, DataGridTemplateColumn,
#     StackPanel, Button, ScrollViewer,
#     DataGridRow, Separator
# )

# # XAML parser
# from System.Windows.Markup import XamlReader
# from System.IO import MemoryStream
# from System.Text import Encoding

# doc = revit.doc
# uidoc = revit.uidoc

# # ── XAML layout ───────────────────────────────────────────────────────────────
# XAML = """
# <Window
#     xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
#     xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
#     Title="Select Rebar by Schedule Mark"
#     Width="680" Height="600"
#     WindowStartupLocation="CenterScreen"
#     Background="#f0f0f0">

#     <Window.Resources>

#         <!-- Row alternating -->
#         <Style x:Key="AltRow" TargetType="DataGridRow">
#             <Setter Property="Background" Value="#ffffff"/>
#             <Style.Triggers>
#                 <Trigger Property="AlternationIndex" Value="1">
#                     <Setter Property="Background" Value="#f5f5f5"/>
#                 </Trigger>
#                 <Trigger Property="IsMouseOver" Value="True">
#                     <Setter Property="Background" Value="#dce8f7"/>
#                 </Trigger>
#             </Style.Triggers>
#         </Style>

#         <!-- Column header — matches Revit 2023 ribbon header grey -->
#         <Style TargetType="DataGridColumnHeader">
#             <Setter Property="Background" Value="#e0e0e0"/>
#             <Setter Property="Foreground" Value="#1a1a1a"/>
#             <Setter Property="FontWeight" Value="SemiBold"/>
#             <Setter Property="FontSize" Value="12"/>
#             <Setter Property="FontFamily" Value="Segoe UI"/>
#             <Setter Property="Padding" Value="8,6"/>
#             <Setter Property="BorderBrush" Value="#c8c8c8"/>
#             <Setter Property="BorderThickness" Value="0,0,1,1"/>
#         </Style>

#         <!-- Cell -->
#         <Style TargetType="DataGridCell">
#             <Setter Property="Foreground" Value="#1a1a1a"/>
#             <Setter Property="FontFamily" Value="Segoe UI"/>
#             <Setter Property="FontSize" Value="12"/>
#             <Setter Property="BorderThickness" Value="0"/>
#             <Setter Property="Padding" Value="6,4"/>
#             <Style.Triggers>
#                 <Trigger Property="IsSelected" Value="True">
#                     <Setter Property="Background" Value="Transparent"/>
#                     <Setter Property="BorderBrush" Value="Transparent"/>
#                     <Setter Property="Foreground" Value="#1a1a1a"/>
#                 </Trigger>
#             </Style.Triggers>
#         </Style>

#         <!-- CheckBox -->
#         <Style TargetType="CheckBox">
#             <Setter Property="HorizontalAlignment" Value="Center"/>
#             <Setter Property="VerticalAlignment" Value="Center"/>
#             <Setter Property="Foreground" Value="#1a1a1a"/>
#         </Style>

#         <!-- TextBox (search) -->
#         <Style TargetType="TextBox">
#             <Setter Property="Background" Value="#ffffff"/>
#             <Setter Property="Foreground" Value="#1a1a1a"/>
#             <Setter Property="BorderBrush" Value="#aaaaaa"/>
#             <Setter Property="BorderThickness" Value="1"/>
#             <Setter Property="FontFamily" Value="Segoe UI"/>
#             <Setter Property="FontSize" Value="12"/>
#             <Setter Property="Padding" Value="6,4"/>
#             <Setter Property="VerticalContentAlignment" Value="Center"/>
#         </Style>

#         <!-- Buttons — Revit-style flat grey with blue accent on OK -->
#         <Style x:Key="ActionBtn" TargetType="Button">
#             <Setter Property="FontFamily" Value="Segoe UI"/>
#             <Setter Property="FontSize" Value="12"/>
#             <Setter Property="FontWeight" Value="SemiBold"/>
#             <Setter Property="Foreground" Value="#1a1a1a"/>
#             <Setter Property="Padding" Value="20,6"/>
#             <Setter Property="BorderThickness" Value="1"/>
#             <Setter Property="BorderBrush" Value="#aaaaaa"/>
#             <Setter Property="Cursor" Value="Hand"/>
#             <Setter Property="Template">
#                 <Setter.Value>
#                     <ControlTemplate TargetType="Button">
#                         <Border x:Name="Bd"
#                                 Background="{TemplateBinding Background}"
#                                 BorderBrush="{TemplateBinding BorderBrush}"
#                                 BorderThickness="{TemplateBinding BorderThickness}"
#                                 CornerRadius="2"
#                                 Padding="{TemplateBinding Padding}">
#                             <ContentPresenter HorizontalAlignment="Center"
#                                              VerticalAlignment="Center"/>
#                         </Border>
#                         <ControlTemplate.Triggers>
#                             <Trigger Property="IsMouseOver" Value="True">
#                                 <Setter TargetName="Bd" Property="Background" Value="#d6e8fa"/>
#                                 <Setter TargetName="Bd" Property="BorderBrush" Value="#5b9bd5"/>
#                             </Trigger>
#                             <Trigger Property="IsPressed" Value="True">
#                                 <Setter TargetName="Bd" Property="Background" Value="#b8d4f0"/>
#                             </Trigger>
#                         </ControlTemplate.Triggers>
#                     </ControlTemplate>
#                 </Setter.Value>
#             </Setter>
#         </Style>

#         <Style x:Key="OkBtn" TargetType="Button" BasedOn="{StaticResource ActionBtn}">
#             <Setter Property="Foreground" Value="#ffffff"/>
#             <Setter Property="BorderBrush" Value="#2e75b6"/>
#             <Setter Property="Template">
#                 <Setter.Value>
#                     <ControlTemplate TargetType="Button">
#                         <Border x:Name="Bd"
#                                 Background="{TemplateBinding Background}"
#                                 BorderBrush="{TemplateBinding BorderBrush}"
#                                 BorderThickness="1"
#                                 CornerRadius="2"
#                                 Padding="{TemplateBinding Padding}">
#                             <ContentPresenter HorizontalAlignment="Center"
#                                              VerticalAlignment="Center"/>
#                         </Border>
#                         <ControlTemplate.Triggers>
#                             <Trigger Property="IsMouseOver" Value="True">
#                                 <Setter TargetName="Bd" Property="Background" Value="#2e75b6"/>
#                             </Trigger>
#                             <Trigger Property="IsPressed" Value="True">
#                                 <Setter TargetName="Bd" Property="Background" Value="#1a5c9e"/>
#                             </Trigger>
#                         </ControlTemplate.Triggers>
#                     </ControlTemplate>
#                 </Setter.Value>
#             </Setter>
#         </Style>

#     </Window.Resources>

#     <Border Background="#f0f0f0" Padding="12">
#         <DockPanel>

#             <!-- Top toolbar -->
#             <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,8">
#                 <TextBlock Text="🔍" Foreground="#666666" VerticalAlignment="Center"
#                            FontSize="13" Margin="0,0,6,0"/>
#                 <TextBox x:Name="SearchBox" Width="200" Height="26"/>
#                 <TextBlock Width="16"/>
#                 <CheckBox x:Name="SelectAllBox" Content="Select All"
#                           Foreground="#333333" VerticalAlignment="Center"
#                           FontFamily="Segoe UI" FontSize="12"/>
#                 <TextBlock x:Name="CountLabel" Foreground="#666666"
#                            VerticalAlignment="Center" FontSize="11"
#                            FontFamily="Segoe UI" Margin="16,0,0,0"/>
#             </StackPanel>

#             <!-- Bottom buttons -->
#             <StackPanel DockPanel.Dock="Bottom" Orientation="Horizontal"
#                         HorizontalAlignment="Right" Margin="0,10,0,0">
#                 <Button x:Name="CancelBtn" Content="Cancel"
#                         Style="{StaticResource ActionBtn}"
#                         Background="#e0e0e0" Margin="0,0,8,0"/>
#                 <Button x:Name="OkBtn" Content="Select"
#                         Style="{StaticResource OkBtn}"
#                         Background="#3a86c8"/>
#             </StackPanel>

#             <!-- Thin separator line above buttons -->
#             <Separator DockPanel.Dock="Bottom" Background="#cccccc" Margin="0,8,0,0"/>

#             <!-- Data grid -->
#             <DataGrid x:Name="RebarGrid"
#                       AutoGenerateColumns="False"
#                       CanUserAddRows="False"
#                       CanUserDeleteRows="False"
#                       CanUserReorderColumns="False"
#                       CanUserResizeRows="False"
#                       IsReadOnly="False"
#                       SelectionMode="Single"
#                       SelectionUnit="Cell"
#                       HeadersVisibility="Column"
#                       GridLinesVisibility="Horizontal"
#                       HorizontalGridLinesBrush="#e0e0e0"
#                       AlternationCount="2"
#                       RowStyle="{StaticResource AltRow}"
#                       Background="#ffffff"
#                       RowBackground="#ffffff"
#                       BorderBrush="#c8c8c8"
#                       BorderThickness="1"
#                       ColumnHeaderHeight="34"
#                       RowHeight="30">

#                 <DataGrid.Columns>
#                     <DataGridTemplateColumn Header="" Width="40" CanUserSort="False">
#                         <DataGridTemplateColumn.CellTemplate>
#                             <DataTemplate>
#                                 <CheckBox IsChecked="{Binding IsSelected, Mode=TwoWay,
#                                           UpdateSourceTrigger=PropertyChanged}"
#                                           HorizontalAlignment="Center"
#                                           VerticalAlignment="Center"/>
#                             </DataTemplate>
#                         </DataGridTemplateColumn.CellTemplate>
#                     </DataGridTemplateColumn>

#                     <DataGridTextColumn Header="Bar Mark"
#                                         Binding="{Binding BarMark}"
#                                         Width="150" IsReadOnly="True"/>

#                     <DataGridTextColumn Header="Schedule Mark"
#                                         Binding="{Binding ScheduleMark}"
#                                         Width="*" IsReadOnly="True"/>

#                     <DataGridTextColumn Header="Count"
#                                         Binding="{Binding Count}"
#                                         Width="90" IsReadOnly="True"/>
#                 </DataGrid.Columns>
#             </DataGrid>

#         </DockPanel>
#     </Border>
# </Window>
# """

# clr.AddReference('System.Xml')
# from System.Xml import XmlReader
# from System.IO import StringReader
# from System.Windows.Markup import XamlReader as XamlLoad

# xml_reader = XmlReader.Create(StringReader(XAML))
# window = XamlLoad.Load(xml_reader)

# # ── Row view-model ─────────────────────────────────────────────────────────────
# from System.ComponentModel import INotifyPropertyChanged
# from System.Windows import DependencyObject, DependencyProperty

# class RebarRow(object):
#     """One row in the grid: one Schedule Mark entry."""
#     def __init__(self, bar_mark, sched_mark, ids):
#         self._selected = False
#         self.BarMark = bar_mark
#         self.ScheduleMark = sched_mark
#         self.Count = "{} bar{}".format(len(ids), "s" if len(ids) != 1 else "")
#         self.ElementIds = ids

#     @property
#     def IsSelected(self):
#         return self._selected

#     @IsSelected.setter
#     def IsSelected(self, value):
#         self._selected = value

# # ── Collect rebar data ─────────────────────────────────────────────────────────
# all_rebar = FilteredElementCollector(doc)\
#     .OfClass(Rebar)\
#     .WhereElementIsNotElementType()\
#     .ToElements()

# if not all_rebar:
#     forms.alert("No structural rebar found in the document.", exitscript=True)

# bar_mark_groups = defaultdict(lambda: defaultdict(list))

# for bar in all_rebar:
#     # Schedule Mark
#     p_sched = bar.get_Parameter(BuiltInParameter.REBAR_ELEM_SCHEDULE_MARK)
#     if p_sched is None:
#         p_sched = bar.LookupParameter("Schedule Mark")
#     sched_mark = (p_sched.AsString() if (p_sched and p_sched.AsString()) else "<No Schedule Mark>")

#     # Bar Mark — no dedicated BuiltInParameter, use LookupParameter
#     bar_mark = "<No Bar Mark>"
#     for param_name in ["Bar Mark", "Mark", "Type Mark"]:
#         p_bar = bar.LookupParameter(param_name)
#         if p_bar and p_bar.AsString():
#             bar_mark = p_bar.AsString()
#             break

#     bar_mark_groups[bar_mark][sched_mark].append(bar.Id)

# # Build flat sorted list of rows
# all_rows = []
# for bar_mark in sorted(bar_mark_groups.keys(), key=lambda m: (m.startswith("<"), m)):
#     for sched_mark in sorted(bar_mark_groups[bar_mark].keys(), key=lambda m: (m.startswith("<"), m)):
#         all_rows.append(RebarRow(bar_mark, sched_mark, bar_mark_groups[bar_mark][sched_mark]))

# # ── Build and show WPF window ──────────────────────────────────────────────────
# from System.Windows.Markup import XamlReader
# from System.IO import StringReader
# from System.Windows.Controls import DataGridRow
# import System.Windows.Data as WPFData

# # Parse XAML properly
# from System.Text import Encoding
# from System.IO import MemoryStream
# xaml_bytes = Encoding.UTF8.GetBytes(XAML)
# xaml_stream = MemoryStream(xaml_bytes)
# window = XamlReader.Load(xaml_stream)

# grid      = window.FindName("RebarGrid")
# search    = window.FindName("SearchBox")
# sel_all   = window.FindName("SelectAllBox")
# ok_btn    = window.FindName("OkBtn")
# cancel_btn= window.FindName("CancelBtn")
# count_lbl = window.FindName("CountLabel")

# # Load rows into grid
# items = ObservableCollection[object]()
# for row in all_rows:
#     items.Add(row)
# grid.ItemsSource = items

# def update_count():
#     n = sum(1 for r in all_rows if r.IsSelected)
#     count_lbl.Text = "{} mark{} selected".format(n, "s" if n != 1 else "")

# update_count()

# # Search filter
# def on_search(sender, e):
#     query = search.Text.strip().lower()
#     filtered = ObservableCollection[object]()
#     for row in all_rows:
#         if query in row.BarMark.lower() or query in row.ScheduleMark.lower():
#             filtered.Add(row)
#     grid.ItemsSource = filtered

# search.TextChanged += on_search

# # Select All toggle
# def on_select_all(sender, e):
#     state = sel_all.IsChecked
#     for row in all_rows:
#         row.IsSelected = state
#     # refresh grid
#     grid.ItemsSource = None
#     src = ObservableCollection[object]()
#     query = search.Text.strip().lower()
#     for row in all_rows:
#         if not query or query in row.BarMark.lower() or query in row.ScheduleMark.lower():
#             src.Add(row)
#     grid.ItemsSource = src
#     update_count()

# sel_all.Checked   += on_select_all
# sel_all.Unchecked += on_select_all

# # Clicking a row toggles its checkbox
# def on_cell_click(sender, e):
#     row_obj = e.Row.Item if hasattr(e, 'Row') and e.Row else None
#     if row_obj:
#         row_obj.IsSelected = not row_obj.IsSelected
#         grid.Items.Refresh()
#         update_count()

# grid.SelectedCellsChanged += on_cell_click

# # OK / Cancel
# result_ids = []

# def on_ok(sender, e):
#     for row in all_rows:
#         if row.IsSelected:
#             result_ids.extend(row.ElementIds)
#     window.DialogResult = True
#     window.Close()

# def on_cancel(sender, e):
#     window.Close()

# ok_btn.Click     += on_ok
# cancel_btn.Click += on_cancel

# window.ShowDialog()

# # ── Apply selection ────────────────────────────────────────────────────────────
# if not result_ids:
#     script.exit()

# id_collection = List[ElementId](result_ids)
# uidoc.Selection.SetElementIds(id_collection)
# uidoc.ShowElements(id_collection)

# out = script.get_output()
# out.print_md("**Selected {} rebar elements.**".format(len(result_ids)))