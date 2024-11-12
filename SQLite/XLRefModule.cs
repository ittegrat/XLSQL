using System;
using ExcelDna.Integration;
using System.Data.SQLite;

namespace XLSQL
{
  internal class XLRefModule : SQLiteModule
  {

    public const string MNAME = "XLREF";

    public enum UpdateCommand : long {
      Freeze = -1,
      Unfreeze = -2,
      Refresh = -3,
    }

    bool disposed = false;

    public static string GetCreateStmt(string tbName, ExcelReference data, bool headers, bool freezeData) {
      return XLRefTable.GetCreateStatement(MNAME, tbName, data, headers, freezeData);
    }

    public XLRefModule() : base(MNAME) { }

    public override SQLiteErrorCode Create(SQLiteConnection connection, IntPtr pClientData, string[] args,
      ref SQLiteVirtualTable table, ref string error
    ) {
      return Connect(connection, pClientData, args, ref table, ref error);
    }
    public override SQLiteErrorCode Connect(SQLiteConnection connection, IntPtr pClientData, string[] args,
      ref SQLiteVirtualTable table, ref string error
    ) {

      CheckDisposed();

      var tbl = new XLRefTable(args);
      var schema = tbl.GetSchema();

      var rc = DeclareTable(connection, schema, ref error);

      if (rc == SQLiteErrorCode.Ok)
        table = tbl;

      return rc;

    }
    public override SQLiteErrorCode Disconnect(SQLiteVirtualTable table) {
      CheckDisposed();
      table.Dispose();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Destroy(SQLiteVirtualTable table) {
      return Disconnect(table);
    }

    public override SQLiteErrorCode Open(SQLiteVirtualTable table, ref SQLiteVirtualTableCursor cursor) {

      CheckDisposed();

      if (!(table is XLRefTable tbl)) {
        SetTableError(table, $"Type mismatch: table type is '{table.GetType().Name}', expected '{typeof(XLRefTable).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      try {
        cursor = tbl.GetCursor();
      }
      catch {
        SetTableError(table, $"Failed to get cursor from table '{tbl.TableName}'.");
        return SQLiteErrorCode.Internal;
      }

      return SQLiteErrorCode.Ok;

    }
    public override SQLiteErrorCode Filter(SQLiteVirtualTableCursor cursor, int idxNum, string idxStr, SQLiteValue[] args) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      cs.Filter(idxNum, idxStr, args);
      cs.Reset();
      cs.Next();

      return SQLiteErrorCode.Ok;

    }
    public override SQLiteErrorCode Next(SQLiteVirtualTableCursor cursor) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      if (cs.Eof) {
        SetCursorError(cursor, "Already hit end of table.");
        return SQLiteErrorCode.Error;
      }

      cs.Next();

      return SQLiteErrorCode.Ok;

    }
    public override bool Eof(SQLiteVirtualTableCursor cursor) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return true;
      }

      return cs.Eof;

    }
    public override SQLiteErrorCode RowId(SQLiteVirtualTableCursor cursor, ref long rowId) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      if (cs.Eof) {
        SetCursorError(cursor, "Already hit end of table.");
        return SQLiteErrorCode.Error;
      }

      rowId = cs.RowId;

      return SQLiteErrorCode.Ok;

    }
    public override SQLiteErrorCode Column(SQLiteVirtualTableCursor cursor, SQLiteContext context, int idx) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      if (cs.Eof) {
        SetCursorError(cursor, "Already hit end of table.");
        return SQLiteErrorCode.Error;
      }

      var value = cs.Current(idx);

      switch (value) {
        case string v:
          context.SetString(v);
          break;
        case double v:
          context.SetDouble(v);
          break;
        case bool v:
          context.SetInt(v ? 1 : 0);
          break;
        case ExcelEmpty _:
        case ExcelError _:
        case ExcelMissing _:
        case null:
          context.SetNull();
          break;
        case int v: // Should never occur
          context.SetInt(v);
          break;
        case long v: // Should never occur
          context.SetInt64(v);
          break;
        default:
          context.SetError("Unmapped type");
          break;
      }

      return SQLiteErrorCode.Ok;

    }
    public override SQLiteErrorCode Close(SQLiteVirtualTableCursor cursor) {

      CheckDisposed();

      if (!(cursor is XLRefTableCursor cs)) {
        SetCursorError(cursor, $"Type mismatch: cursor type is '{cursor.GetType().Name}', expected '{typeof(XLRefTableCursor).Name}'.");
        return SQLiteErrorCode.Internal;
      }

      cs.Close();

      return SQLiteErrorCode.Ok;

    }

    public override SQLiteErrorCode BestIndex(SQLiteVirtualTable table, SQLiteIndex index) {
      CheckDisposed();
      if (table.BestIndex(index))
        return SQLiteErrorCode.Ok;
      SetTableError(table, $"Failed to select best index for virtual table '{table.TableName}'.");
      return SQLiteErrorCode.Internal;
    }
    public override SQLiteErrorCode Rename(SQLiteVirtualTable table, string newName) {
      CheckDisposed();
      if (table.Rename(newName))
        return SQLiteErrorCode.Ok;
      SetTableError(table, $"Failed to rename virtual table from '{table.TableName}' to '{newName}'.");
      return SQLiteErrorCode.Internal;
    }
    public override SQLiteErrorCode Update(SQLiteVirtualTable table, SQLiteValue[] values, ref long rowId) {
      CheckDisposed();
      if (values.Length > 1
        && values[0].GetTypeAffinity() == TypeAffinity.Null
        && values[1].GetTypeAffinity() == TypeAffinity.Int64
      ) {
        var uc = values[1].GetInt64();
        if (Enum.IsDefined(typeof(UpdateCommand), uc))
          (table as XLRefTable).Update((UpdateCommand)uc);
        else {
          SetTableError(table, $"Invalid UpdateCommand code '{uc}', table '{table.TableName}'.");
          return SQLiteErrorCode.Mismatch;
        }
      }
      SetTableError(table, $"Virtual table '{table.TableName}' is read-only.");
      return SQLiteErrorCode.Perm;
    }

    public override bool FindFunction(SQLiteVirtualTable table, int argCount, string name,
      ref SQLiteFunction function, ref IntPtr pClientData
    ) {
      CheckDisposed();
      return true;
    }
    public override SQLiteErrorCode Begin(SQLiteVirtualTable table) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Sync(SQLiteVirtualTable table) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Commit(SQLiteVirtualTable table) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Rollback(SQLiteVirtualTable table) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Savepoint(SQLiteVirtualTable table, int savepoint) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode Release(SQLiteVirtualTable table, int savepoint) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }
    public override SQLiteErrorCode RollbackTo(SQLiteVirtualTable table, int savepoint) {
      CheckDisposed();
      return SQLiteErrorCode.Ok;
    }

    protected override void Dispose(bool disposing) {
      base.Dispose(disposing);
      disposed = true;
    }

    void CheckDisposed() {
      if (disposed)
        throw new ObjectDisposedException(GetType().Name);
    }

  }
}
