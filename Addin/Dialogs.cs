using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers.DialogBox;
using WinForms = System.Windows.Forms;

namespace XLSQL
{

  [ExcelCommand(Prefix = "XLSQL.")]
  public static class Dialogs
  {

    enum OpenDBType { NewFile, NewMemory, OpenFile }
    enum UpdateTblType { Freeze, Unfreeze, Refresh }

    public static void NewMemoryDB() {
      OpenDB(OpenDBType.NewMemory);
    }
    public static void NewFileDB() {
      OpenDB(OpenDBType.NewFile);
    }
    public static void OpenFileDB() {
      OpenDB(OpenDBType.OpenFile);
    }
    public static void CloseDB() {

      var title = "XLSQL - Close database";

      if (!DbPool.DbNames.Any()) {
        WinForms.MessageBox.Show("No database open.", title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
        return;
      }

      var dialog = new Dialog() { Title = title, }
        .Add(new StaticText() { Left = 10, Top = 6, Height = 94, Text = "Databases:", })
        .Add(new ListBox("db") { Left = 10, Top = 24, Width = 185, Height = 72, }.AddItemRange(DbPool.DbNames))
        .Add(new OkButton("one") { Left = 205, Top = 24, Width = 88, Text = "Close", })
        .Add(new OkButton("all") { Left = 205, Top = 78, Width = 88, Text = "Close all", })
      ;

      if (!dialog.Show() || dialog.TriggerId == null)
        return;

      var items = dialog.TriggerId == "one"
        ? new object[] { dialog.GetValue<string>("db") }
        : dialog.GetControl<ListBox>("db").Items
      ;

      foreach (var dbName in items) {
        var ans = SheetFunctions.CloseConnection((string)dbName, null);
        if (SheetFunctions.IsError(ans)) {
          WinForms.MessageBox.Show((string)ans, title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
          return;
        }
      }

    }
    //[ExcelCommand(MenuText = "SaveAs")]
    //public static void SaveAsDB() { }

    //** [ExcelCommand(MenuText = "CloneDB")]
    //** public static void ....() { }

    public static void NewRangeTable() {

      var title = "XLSQL - New table";

      if (!DbPool.DbNames.Any()) {
        WinForms.MessageBox.Show("No database open.", title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
        return;
      }

      var dialog = new Dialog() { Title = title, }
        .Add(new StaticText() { Left = 6, Top = 9, Text = "Range:", })
        .Add(new RefEditBox("range") { Left = 57, Top = 7, Width = 310, })
        .Add(new StaticText() { Left = 12, Top = 36, Text = "Table:", })
        .Add(new TextEditBox("name") { Left = 57, Top = 33, Width = 200, })
        .Add(new StaticText() { Left = 30, Top = 63, Text = "DB:", })
        .Add(new DropDown("db") { Left = 57, Top = 60, Width = 200, }.AddItemRange(DbPool.DbNames))
        .Add(new CheckBox("ovr") { Left = 276, Top = 31, Width = 90, Text = "Overwrite", Value = false, })
        .Add(new CheckBox("head") { Left = 276, Top = 48, Width = 90, Text = "Headers", Value = false, })
        .Add(new CheckBox("frz") { Left = 276, Top = 65, Width = 90, Text = "Freeze", Value = false, })
        .Add(new OkButtonDef() { Left = 176, Top = 88, Width = 88, })
        .Add(new CancelButton() { Left = 276, Top = 88, Width = 88, })
      ;

      var selected = String.Empty;
      var selection = XlCall.Excel(XlCall.xlfSelection);
      if (selection is ExcelReference sel && (sel.RowLast - sel.RowFirst + 1) * (sel.ColumnLast - sel.ColumnFirst + 1) > 1) {
        var r1c1 = (bool)XlCall.Excel(XlCall.xlfGetWorkspace, 4);
        selected = XlCall.Excel(XlCall.xlfReftext, selection, !r1c1) as string;
      }
      dialog.SetValue("range", selected);
      dialog.SetFocus(String.IsNullOrEmpty(selected) ? "range" : "name");

      bool Handled(Dialog d) {
        var range = d.GetValue<string>("range");
        if (String.IsNullOrWhiteSpace(range)) {
          XlCall.Excel(XlCall.xlcAlert, "Reference is not valid.", 2);
          d.SetFocus("range");
          return false;
        }
        var name = d.GetValue<string>("name");
        if (String.IsNullOrWhiteSpace(name)) {
          XlCall.Excel(XlCall.xlcAlert, "Table name is not valid.", 2);
          d.SetFocus("name");
          return false;
        }
        return true;
      }

      if (!dialog.Show(Handled))
        return;

      var dbName = dialog.GetValue<string>("db");
      var tbName = dialog.GetValue<string>("name");

      selected = dialog.GetValue<string>("range");
      ExcelReference xlref;
      if (!selected.Contains("!")) {
        xlref = new ExcelReference(0, 0, 0, 0, IntPtr.Zero);
        var sheetNm = XlCall.Excel(XlCall.xlSheetNm, xlref);
        selected = $"'{sheetNm}'!{selected}";
      }
      xlref = XlCall.Excel(XlCall.xlfTextref, selected, false) as ExcelReference;

      var headers = dialog.GetValue<bool>("head");
      var freeze = dialog.GetValue<bool>("frz");
      var over = dialog.GetValue<bool>("ovr");

      var ans = SheetFunctions.CreateTable(dbName, tbName, xlref, null, headers, freeze, over);
      if (SheetFunctions.IsError(ans))
        WinForms.MessageBox.Show((string)ans, dialog.Title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);

    }
    public static void FreezeRangeTable() {
      UpdateTable(UpdateTblType.Freeze);
    }
    public static void UnfreezeRangeTable() {
      UpdateTable(UpdateTblType.Unfreeze);
    }
    public static void RefreshRangeTable() {
      UpdateTable(UpdateTblType.Refresh);
    }

    public static void About() {

      Func<System.Reflection.Assembly, string> row = (asm) => $"{asm.GetName().Name}: {System.Reflection.CustomAttributeExtensions.GetCustomAttribute<System.Reflection.AssemblyFileVersionAttribute>(asm).Version}";

      var rows = new List<object> {
        row(typeof(Dialogs).Assembly),
        row(typeof(DnaLibrary).Assembly),
        row(typeof(System.Data.SQLite.SQLiteConnection).Assembly),
        row(typeof(ICSharpCode.AvalonEdit.TextEditor).Assembly),
        row(typeof(NLog.Logger).Assembly),
      };

      var lb = new ListBox(rows) { Left = 20, Top = 22, Width = 280, Height = 110, };
      lb.SelectedIndex = null;

      var dialog = new Dialog { Width = 320, Height = 186, Title = "About XLSQL", }
        .Add(new GroupBox { Left = 10, Top = 5, Width = 300, Height = 140, Text = "Assembly File Versions", })
        .Add(lb)
        .Add(new OkButtonDef { Left = 120, Top = 152, Width = 80, Height = 24, Text = "&OK", })
      ;

      dialog.Show();

    }

    static void OpenDB(OpenDBType type) {

      var ext = Configuration.ExtensionsEnabled;

      var dialog = new Dialog()
        .Add(new OkButton("pathB") { Left = 6, Top = 7, Width = 67, Text = "Path:", })
        .Add(new TextEditBox("path") { Left = 78, Top = 9, Width = 346, })
        .Add(new StaticText() { Left = 8, Top = 36, Text = "DB Name:", })
        .Add(new TextEditBox("name") { Left = 78, Top = 34, Width = 250, })
        .Add(new CheckBox("ro") { Left = 8, Top = 58, Text = "Read-only", Value = false, Enabled = false })
        .Add(new CheckBox("ext") { Left = 126, Top = 58, Text = "Load Extensions", Value = ext, Enabled = ext })
        .Add(new CancelButton() { Left = 338, Top = 33, Width = 88, })
        .Add(new OkButtonDef("ok") { Left = 338, Top = 58, Width = 88, })
      ;
      dialog.SetFocus("path");

      switch (type) {

        case OpenDBType.NewFile:
          dialog.Title = "XLSQL - New database";
          break;

        case OpenDBType.NewMemory:
          dialog.Title = "XLSQL - New memory database";
          dialog.Disable("pathB");
          dialog.Disable("path");
          dialog.SetValue("path", ":memory:");
          dialog.SetFocus("name");
          break;

        case OpenDBType.OpenFile:
          dialog.Title = "XLSQL - Open database";
          dialog.Enable("ro");
          break;

      }

      bool fileOverwrite = false;
      bool connOverwrite = false;

      bool Handled(Dialog d) {

        string path, name;

        if (d.TriggerId == "pathB") {

          WinForms.FileDialog fd;
          if (type == OpenDBType.OpenFile) {
            fd = new WinForms.OpenFileDialog();
            fileOverwrite = true;
          }
          else {
            fd = new WinForms.SaveFileDialog { OverwritePrompt = false };
          }

          fd.Title = d.Title;
          fd.Filter = "Databases|*.sqlite;*.sqlite3;*.db;*.db3|All files|*.*";
          fd.FilterIndex = 1;
          fd.RestoreDirectory = true;

          path = d.GetValue<string>("path");
          if (!String.IsNullOrWhiteSpace(path)) {
            fd.InitialDirectory = Path.GetDirectoryName(path);
            fd.FileName = Path.GetFileName(path);
            if (!fd.Filter.Contains(Path.GetExtension(path)))
              fd.FilterIndex = 2;
          }

          if (fd.ShowDialog() == WinForms.DialogResult.OK) {
            path = fd.FileName;
            if (!Path.HasExtension(path))
              path += ".sqlite";
            d.SetValue("path", path);
            name = Path.GetFileNameWithoutExtension(path);
            if (DbPool.ContainsName(name)) {
              d.SetFocus("name");
            }
            else {
              d.SetValue("name", name);
              d.SetFocus("ok");
            }

          }

          fd.Reset();
          return false;

        }

        path = d.GetValue<string>("path");
        if (String.IsNullOrWhiteSpace(path)) {
          WinForms.MessageBox.Show("Invalid file path.", d.Title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
          d.SetFocus("path");
          return false;
        }
        if (!fileOverwrite && File.Exists(path)) {
          var ans = WinForms.MessageBox.Show(
            $"The file '{Path.GetFileName(path)}' already exists.\nDo you want to replace it?", d.Title,
            WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Warning, WinForms.MessageBoxDefaultButton.Button2
          );
          if (ans != WinForms.DialogResult.Yes)
            return false;
          fileOverwrite = true;
        }

        name = d.GetValue<string>("name");
        if (String.IsNullOrWhiteSpace(name)) {
          WinForms.MessageBox.Show("Invalid db name.", d.Title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);
          d.SetFocus("name");
          return false;
        }
        if (!connOverwrite && DbPool.ContainsName(name)) {
          var ans = WinForms.MessageBox.Show(
            $"Connection '{name}' already exists.\nDo you want to replace it?", d.Title,
            WinForms.MessageBoxButtons.YesNo, WinForms.MessageBoxIcon.Warning, WinForms.MessageBoxDefaultButton.Button2
          );
          if (ans != WinForms.DialogResult.Yes)
            return false;
          connOverwrite = true;
        }

        return true;

      }

      if (!dialog.Show(Handled))
        return;

      var dbName = dialog.GetValue<string>("name");
      var dbFile = type == OpenDBType.NewMemory
        ? null
        : dialog.GetValue<string>("path")
      ;
      var ro = dialog.GetValue<bool>("ro");
      ext = dialog.GetValue<bool>("ext");

      object ret;
      if (type == OpenDBType.NewFile)
        ret = Commands.CreateDatabase(dbFile, dbName, fileOverwrite, connOverwrite, !ext);
      else
        ret = SheetFunctions.OpenConnection(dbName, dbFile, null, ro, !ext, connOverwrite);
      if (SheetFunctions.IsError(ret))
        WinForms.MessageBox.Show((string)ret, dialog.Title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);

    }
    static void UpdateTable(UpdateTblType type) {

      var command = type.ToString();

      var title = $"XLSQL - {command} table";

      if (!DbPool.DbNames.Any()) {
        WinForms.MessageBox.Show("No database open.", title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
        return;
      }

      var dialog = new Dialog() { Title = title, }
        .Add(new StaticText() { Left = 6, Top = 9, Text = "DB:", })
        .Add(new DropDown("db") { Left = 50, Top = 6, Width = 180, IsTrigger = true, })
        .Add(new StaticText() { Left = 6, Top = 35, Text = "Table:", })
        .Add(new DropDown("tbl") { Left = 50, Top = 32, Width = 180, })
        .Add(new OkButton("exec") { Left = 26, Top = 57, Width = 88, Text = command, })
        .Add(new CancelButtonDef() { Left = 126, Top = 57, Width = 88, })
      ;

      var ddConn = dialog.GetControl<DropDown>("db");
      var ddTbl = dialog.GetControl<DropDown>("tbl");

      var tbls = new List<string[]>();
      foreach (var name in DbPool.DbNames) {
        var db = DbPool.Get(name);
        var tl = db.ListTables(null);
        if (tl.Count > 0) {
          ddConn.AddItem(name);
          tbls.Add(tl.Select(r => (string)r[0]).ToArray());
        }
      }

      if (tbls.Count == 0) {
        WinForms.MessageBox.Show("No virtual table found.", title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Warning);
        return;
      }

      ddTbl.AddItemRange(tbls[0]);
      ddTbl.SelectedIndex = 0;
      ddConn.SelectedIndex = 0;
      var dbName = ddConn.SelectedItem as string;

      bool Handled(Dialog d) {

        if (d.TriggerId == "db") {

          var name = d.GetValue<string>("db");
          if (name != dbName) {
            ddTbl.ClearItems().AddItemRange(tbls[(int)ddConn.SelectedIndex]);
            dbName = name;
          }

          return false;

        }

        return true;

      }

      if (!dialog.Show(Handled))
        return;

      object ans = null;
      switch (type) {
        case UpdateTblType.Freeze:
          ans = SheetFunctions.FreezeTable(ddConn.SelectedItem as string, ddTbl.SelectedItem as string, null);
          break;
        case UpdateTblType.Unfreeze:
          ans = SheetFunctions.UnfreezeTable(ddConn.SelectedItem as string, ddTbl.SelectedItem as string, null);
          break;
        case UpdateTblType.Refresh:
          ans = SheetFunctions.RefreshTable(ddConn.SelectedItem as string, ddTbl.SelectedItem as string, null);
          break;
      }
      if (SheetFunctions.IsError(ans))
        WinForms.MessageBox.Show((string)ans, dialog.Title, WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error);

    }

  }

}
