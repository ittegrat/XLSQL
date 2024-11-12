using System;
using System.Text;
using ExcelDna.Integration;
using System.Data.SQLite;

namespace XLSQL
{
  internal class XLRefTable : SQLiteVirtualTable
  {

    bool disposed = false;

    readonly int rowFirst;
    readonly int rowLast;
    readonly int columnFirst;
    readonly int columnLast;
    readonly IntPtr sheetId = IntPtr.Zero;

    readonly object[,] head = null;
    object[,] data = null;

    public static string GetCreateStatement(string mName, string tbName, ExcelReference data, bool headers, bool freezeData) {
      var sb = new StringBuilder($"CREATE VIRTUAL TABLE temp.{tbName} USING {mName}(")
        .Append(data.RowFirst).Append(',')
        .Append(data.RowLast).Append(',')
        .Append(data.ColumnFirst).Append(',')
        .Append(data.ColumnLast).Append(',')
        .Append(data.SheetId.ToInt64()).Append(',')
        .Append(headers).Append(',')
        .Append(freezeData).Append(')')
      ;
      return sb.ToString();
    }

    public XLRefTable(string[] args) : base(args) {

      rowFirst = int.Parse(args[3]);
      rowLast = int.Parse(args[4]);
      columnFirst = int.Parse(args[5]);
      columnLast = int.Parse(args[6]);
      sheetId = new IntPtr(long.Parse(args[7]));

      var headers = bool.Parse(args[8]);
      if (headers) {
        var xlref = new ExcelReference(rowFirst, rowFirst, columnFirst, columnLast, sheetId);
        head = (object[,])xlref.GetValue();
        ++rowFirst;
      }

      var frozen = bool.Parse(args[9]);
      if (frozen)
        Update(XLRefModule.UpdateCommand.Freeze);

    }
    public virtual string GetSchema() {

      CheckDisposed();

      var headers = !(head is null);
      var n = headers ? head.GetLength(1) : 1 + columnLast - columnFirst;

      var sb = new StringBuilder("CREATE TABLE X( ");
      for (var i = 0; i < n; ++i) {
        if (i > 0) sb.Append(" , ");
        if (headers && head[0, i] is string str && !String.IsNullOrWhiteSpace(str))
          sb.Append(str);
        else
          // sb.Append($"C{1 + i}");
          sb.Append($"'{ToColumn(1 + columnFirst + i)}'");
      }
      sb.Append(" );");
      //*$* sb.Append(" , cmd HIDDEN );");

      return sb.ToString();

    }
    public virtual XLRefTableCursor GetCursor() {
      CheckDisposed();
      return new XLRefTableCursor(this,
        data ??
        (object[,])new ExcelReference(rowFirst, rowLast, columnFirst, columnLast, sheetId).GetValue()
      );
    }
    public virtual void Update(XLRefModule.UpdateCommand uc) {

      if (uc == XLRefModule.UpdateCommand.Refresh && data is null)
        return;
      else if (uc == XLRefModule.UpdateCommand.Freeze && data != null)
        return;
      else if (uc == XLRefModule.UpdateCommand.Unfreeze) {
        data = null;
        return;
      }

      var xlref = new ExcelReference(rowFirst, rowLast, columnFirst, columnLast, sheetId);
      if (rowFirst == rowLast && columnFirst == columnLast)
        data = new object[,] { { xlref.GetValue() } };
      else
        data = (object[,])xlref.GetValue();

    }

    protected override void Dispose(bool disposing) {
      base.Dispose(disposing);
      disposed = true;
    }

    void CheckDisposed() {
      if (disposed)
        throw new ObjectDisposedException(GetType().Name);
    }
    string ToColumn(int index) {

      if (index < 1 || index > 16384)
        throw new ArgumentException("Invalid column index.");

      var pos = 3;
      var chars = new char[pos];

      while (index > 0) {
        var mod = (index - 1) % 26;
        chars[--pos] = (char)(65 + mod);
        index = (index - mod) / 26;
      }

      return new String(chars, pos, 3 - pos);

    }

  }
}
