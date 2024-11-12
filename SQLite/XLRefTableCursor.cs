using System;
using System.Data.SQLite;

namespace XLSQL
{
  internal class XLRefTableCursor : SQLiteVirtualTableCursor
  {

    bool disposed = false;
    bool closed = false;

    readonly object[,] data;
    readonly long length;
    long rowId;

    public virtual bool Eof {
      get { CheckClosed(); return rowId >= length; }
    }
    public virtual long RowId {
      get { CheckClosed(); return rowId; }
    }

    public XLRefTableCursor(XLRefTable table, object[,] data) : base(table) {
      this.data = data;
      length = data.GetLongLength(0);
      Reset();
    }

    public virtual void Close() {
      closed = true;
    }
    public virtual object Current(int idx) {
      CheckClosed();
      return data[rowId, idx];
    }
    public virtual void Reset() {
      CheckClosed();
      rowId = -1;
    }
    public virtual void Next() {
      CheckClosed();
      ++rowId;
      if (rowId < length)
        NextRowIndex();
    }

    protected override void Dispose(bool disposing) {
      try {
        if (!closed)
          Close();
      }
      finally {
        base.Dispose(disposing);
        disposed = true;
      }
    }

    void CheckClosed() {
      CheckDisposed();
      if (closed)
        throw new InvalidOperationException("Virtual table cursor is closed.");
    }
    void CheckDisposed() {
      if (disposed)
        throw new ObjectDisposedException(GetType().Name);
    }

  }
}
