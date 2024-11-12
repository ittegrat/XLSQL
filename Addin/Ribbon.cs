using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace XLSQL
{
  [ComVisible(true)]
  [ProgId("XLSQL.Ribbon")]
  [Guid("0AF7D383-0E8E-48B8-8ECA-95719ED102FB")]
  [ExcelCommand(Prefix = "XLSQL.")]
  public class Ribbon : ExcelRibbon
  {

    static IRibbonUI ribbon;
    static bool hidden = false;

    [ExcelCommand("Hide the XLSQL ribbon.")]
    public static void HideTab() {
      hidden = true;
      ribbon?.Invalidate();
    }
    [ExcelCommand("Show the XLSQL ribbon.")]
    public static void ShowTab() {
      hidden = false;
      ribbon?.Invalidate();
    }

    public Ribbon() {
      var asm = typeof(Ribbon).Assembly;
      FriendlyName = asm.GetCustomAttribute<AssemblyTitleAttribute>().Title;
      Description = asm.GetCustomAttribute<AssemblyDescriptionAttribute>().Description;
      hidden = Configuration.HiddenRibbonTab;
    }

    public override string GetCustomUI(string RibbonID) {
      var asm = typeof(Ribbon).Assembly;
      var rn = asm.GetManifestResourceNames().Single(s => s.EndsWith(".Ribbon.xml"));
      using (var stm = asm.GetManifestResourceStream(rn)) {
        using (var sr = new StreamReader(stm)) {
          return sr.ReadToEnd();
        }
      }
    }

    public void OnRibbonLoad(IRibbonUI rui) {
      ribbon = rui;
    }
    public bool TabGetVisible(IRibbonControl control) {
      return !hidden;
    }

    public void OnCnxButton(IRibbonControl control) {
      var action = control.Id.Split('.').Last();
      string macro = null;
      switch (action) {
        case "Open":
          macro = "OpenFileDB"; break;
        case "NewMem":
          macro = "NewMemoryDB"; break;
        case "NewFile":
          macro = "NewFileDB"; break;
        case "Close":
          macro = "CloseDB"; break;
      }
      (ExcelDnaUtil.Application as Excel.Application).Run($"XLSQL.{macro}");
    }
    public void OnDbButton(IRibbonControl control) {
      var action = control.Id.Split('.').Last();
      string macro = null;
      switch (action) {
        case "NewTable":
          macro = "NewRangeTable"; break;
        case "FreezeTbl":
          macro = "FreezeRangeTable"; break;
        case "UnfreezeTbl":
          macro = "UnfreezeRangeTable"; break;
        case "RefreshTbl":
          macro = "RefreshRangeTable"; break;
        case "QueryEd":
          QueryEditor();
          return;
        case "SqlHelp":
          System.Diagnostics.Process.Start("https://www.sqlite.org/lang.html");
          return;
      }
      (ExcelDnaUtil.Application as Excel.Application).Run($"XLSQL.{macro}");
    }
    public void OnAbout(IRibbonControl control) {
      (ExcelDnaUtil.Application as Excel.Application).Run($"XLSQL.About");
    }

    void QueryEditor() {

      const string xaml = @"
<Window xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
        xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
        xmlns:AvalonEdit='clr-namespace:ICSharpCode.AvalonEdit;assembly=ICSharpCode.AvalonEdit'
        MinWidth='500' MinHeight='300'
        SizeToContent='WidthAndHeight'
>
  <Grid>
    <Grid.ColumnDefinitions>
      <ColumnDefinition />
      <ColumnDefinition />
    </Grid.ColumnDefinitions>
    <Grid.RowDefinitions>
      <RowDefinition Height='*' />
      <RowDefinition Height='Auto' />
    </Grid.RowDefinitions>
    <Border Grid.Row='0' Grid.ColumnSpan='2'
            BorderThickness='1' Margin='5,5,5,0' BorderBrush='{x:Static SystemColors.ControlDarkBrush}'
    >
      <AvalonEdit:TextEditor FontFamily='Consolas' FontSize='10pt'
            SyntaxHighlighting='TSQL'
      />
    </Border>
    <ComboBox Grid.Row='1' Margin='5,8,5,8' />
    <Button Grid.Row='1' Grid.Column='1' Margin='5,8,5,8' Content='Execute' />
  </Grid>
</Window>
";

      var title = "XLSQL - Query Editor";

      if (!DbPool.DbNames.Any()) {
        MessageBox.Show("No database open.", title, MessageBoxButton.OK, MessageBoxImage.Warning);
        return;
      }

      var window = XamlReader.Parse(xaml) as Window;
      window.Title = title;

      var editor = ((window.Content as Grid).Children[0] as Border).Child as ICSharpCode.AvalonEdit.TextEditor;
      var db = (window.Content as Grid).Children[1] as ComboBox;
      var execb = (window.Content as Grid).Children[2] as Button;

      db.ItemsSource = DbPool.DbNames;
      db.SelectedIndex = 0;

      string dbName = null;
      string qry = null;

      execb.Click += (s, e) => {
        dbName = db.SelectedItem as string;
        qry = editor.Text;
        window.Close();
      };

      window.ShowDialog();
      if (String.IsNullOrWhiteSpace(qry))
        return;

      var ans = SheetFunctions.Execute(dbName, qry, null, null, null);
      if (SheetFunctions.IsError(ans))
        MessageBox.Show((string)ans, title, MessageBoxButton.OK, MessageBoxImage.Error);

    }

    public string OnGetTitle(IRibbonControl control) {
      var key = control.Id.Substring(6);
      return titles[key];
    }
    public string OnGetDescription(IRibbonControl control) {
      var key = control.Id.Substring(6);
      return descriptions[key];
    }

    static readonly Dictionary<string, string> titles = new Dictionary<string, string> {
      {"Connections.Open","Open Connection"},
      {"Connections.NewMem","New Memory Database"},
      {"Connections.NewFile","New File Database"},
      {"Connections.Close","Close Connection"},
      {"Database.NewTable","New Range Table"},
      {"Database.FreezeTbl","Freeze Range Table"},
      {"Database.UnfreezeTbl","Unfreeze Range Table"},
      {"Database.RefreshTbl","Refresh Range Table"},
      {"Database.QueryEd","Query Editor"},
      {"Database.SqlHelp","SQL Help"},
    };

    static readonly Dictionary<string, string> descriptions = new Dictionary<string, string> {
      {"Connections.Open","Opens a connection to an existent SQLite database."},
      {"Connections.NewMem","Creates a new memory database and opens a connection to it."},
      {"Connections.NewFile","Creates a new file database and opens a connection to it."},
      {"Connections.Close","Closes one or more SQLite database connections."},
      {"Database.NewTable","Creates a virtual table backed by a range in 'temp' schema."},
      {"Database.FreezeTbl","Freezes the values in a 'Range' table."},
      {"Database.UnfreezeTbl","Unfreezes the values in a 'Range' table."},
      {"Database.RefreshTbl","Updates the values in a frozen 'Range' table."},
      {"Database.QueryEd","Opens the Query Editor"},
      {"Database.SqlHelp","Opens the SQL Syntax page on the SQLite web site."},
    };


  }
}
