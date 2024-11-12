using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using System.Data.SQLite;

namespace XLSQL
{
  internal class DbAdapter : IDisposable
  {

    readonly SQLiteConnection db;
    readonly Dictionary<string, SQLiteCommand> commands = new Dictionary<string, SQLiteCommand>();
    bool disposed = false;

    public DbAdapter(string dbFile, bool readOnly, bool loadExt) {

      string connStr;

      dbFile = dbFile?.Trim();
      if (String.IsNullOrEmpty(dbFile)) {
        connStr = "Data Source=:memory:";
      }
      else {
        if (Path.IsPathRooted(dbFile) && Path.GetPathRoot(dbFile).Length > 1 && File.Exists(dbFile)) {
          if (dbFile.StartsWith(@"\\") && !dbFile.StartsWith(@"\\\\"))
            dbFile = @"\\" + dbFile;
          connStr = $"Data Source={dbFile};FailIfMissing=True;ReadOnly={readOnly}";
        }
        else throw new ArgumentException(Strings.INVALID_DBFILE);
      }

      db = new SQLiteConnection(connStr);
      db.Open();
      db.CreateModule(new XLRefModule());
      if (loadExt && Configuration.ExtensionsEnabled) {
        db.EnableExtensions(true);
        foreach (var e in Configuration.Extensions)
          db.LoadExtension("SQLite.Interop.dll", e);
      }

    }

    public void CreateQuery(string qryName, string query, object[] paramNames, bool overwrite) {

      CheckDisposed();

      if (commands.TryGetValue(qryName, out var other)) {
        if (overwrite)
          other.Dispose();
        else
          throw new DuplicateNameException(Strings.ALREADY_EXISTS);
      }

      var cmd = db.CreateCommand();
      cmd.CommandText = query;

      if (paramNames != null) {

        var pc = cmd.Parameters;
        pc.NoCase = true;

        var np = paramNames.Length;

        if (np == 1 && paramNames[0] is int k) {
          if (k <= 0)
            throw new ArgumentException($"Invalid number of positional parameters (ParamNames[0]={k}).");
          for (var i = 0; i < k; ++i)
            pc.Add(new SQLiteParameter());
        }
        else {
          for (var i = 0; i < np; ++i) {
            var name = (paramNames[i] as string)?.Trim();
            if (String.IsNullOrEmpty(name))
              throw new ArgumentException($"Invalid parameter name '{paramNames[i]}' (index={i}).");
            pc.Add(new SQLiteParameter(name));
          }
        }

      }

      commands[qryName] = cmd;

    }
    public void CreateTable(string tblName, ExcelReference data, bool headers, bool freezeData, bool overwrite) {

      CheckDisposed();

      if (overwrite) {
        using (var cmd = db.CreateCommand()) {
          cmd.CommandText = $"DROP TABLE IF EXISTS {tblName}";
          cmd.ExecuteNonQuery();
        }
      }

      using (var cmd = db.CreateCommand()) {
        cmd.CommandText = XLRefModule.GetCreateStmt(tblName, data, headers, freezeData).ToString();
        cmd.ExecuteNonQuery();
      }

    }
    public void DeleteQuery(string qryName) {
      if (!commands.TryGetValue(qryName, out var cmd))
        throw new KeyNotFoundException(Strings.QRY_NOTFOUND);
      cmd.Dispose();
      commands.Remove(qryName);
    }
    public int Execute(string query, object[] paramValues, object[] paramNames) {

      CheckDisposed();

      SQLiteCommand cmd = null;
      bool dispose = false;

      try {

        (cmd, dispose) = GetCommand(query, paramValues, paramNames);

        var ans = cmd.ExecuteNonQuery();

        bool affected = false;
        affected |= query.IndexOf("DELETE", StringComparison.OrdinalIgnoreCase) >= 0;
        affected |= query.IndexOf("INSERT", StringComparison.OrdinalIgnoreCase) >= 0;
        affected |= query.IndexOf("REPLACE", StringComparison.OrdinalIgnoreCase) >= 0;
        affected |= query.IndexOf("UPDATE", StringComparison.OrdinalIgnoreCase) >= 0;

        return affected ? ans : -1;

      }
      finally {
        if (dispose)
          cmd.Dispose();
      }

    }
    public bool ExistsTable(string tbName) {
      CheckDisposed();
      using (var cmd = db.CreateCommand()) {
        cmd.CommandText = $@"
          SELECT name
          FROM temp.sqlite_schema
          WHERE name = '{tbName}'
            AND sql like '%CREATE%VIRTUAL%USING%{XLRefModule.MNAME}%'
        ";
        var ans = cmd.ExecuteScalar();
        return !(ans is null);
      }
    }
    public List<object[]> ListQueries(string nameFilter, string sqlFilter) {
      CheckDisposed();
      nameFilter = nameFilter?.Trim();
      if (String.IsNullOrEmpty(nameFilter)) nameFilter = ".*";
      sqlFilter = sqlFilter?.Trim();
      if (String.IsNullOrEmpty(sqlFilter)) sqlFilter = ".*";
      var nameRegex = new Regex(nameFilter, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
      var sqlRegex = new Regex(sqlFilter, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
      var queries = commands
        .Where(q => nameRegex.Match(q.Key).Success)
        .Where(q => sqlRegex.Match(q.Value.CommandText).Success)
        .OrderBy(q => q.Key)
        .Select(q => new object[] { q.Key, q.Value.CommandText })
      ;
      return queries.ToList();
    }
    public List<object[]> ListTables(string nameFilter) {
      CheckDisposed();

      var qry = $@"
        SELECT name, sql
        FROM temp.sqlite_schema
        WHERE sql like '%CREATE%VIRTUAL%USING%{XLRefModule.MNAME}%'
      ";
      var rows = Query(qry, null, null, false);
      if (rows.Count == 0)
        return rows;

      nameFilter = nameFilter?.Trim();
      if (String.IsNullOrEmpty(nameFilter)) nameFilter = ".*";
      var nameRegex = new Regex(nameFilter, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
      var addrRegex = new Regex($@"{XLRefModule.MNAME}\((\d+),(\d+),(\d+),(\d+),(\d+),");
      string GetAddress(string sql) {
        var match = addrRegex.Match(sql);
        var rowFirst = Int32.Parse(match.Groups[1].Value);
        var rowLast = Int32.Parse(match.Groups[2].Value);
        var columnFirst = Int32.Parse(match.Groups[3].Value);
        var columnLast = Int32.Parse(match.Groups[4].Value);
        var sheetId = new IntPtr(Int64.Parse(match.Groups[5].Value));
        var sheet = XlCall.Excel(XlCall.xlSheetNm, new ExcelReference(rowFirst, rowLast, columnFirst, columnLast, sheetId));
        var address = (string)XlCall.Excel(XlCall.xlfAddress, 1 + rowFirst, 1 + columnFirst, 1, true, sheet);
        if (rowLast != rowFirst || columnLast != columnFirst) {
          var lr = XlCall.Excel(XlCall.xlfAddress, 1 + rowLast, 1 + columnLast, 1, true);
          address = $"{address}:{lr}";
        }
        return address;
      }
      var tables = rows
        .Where(row => nameRegex.Match((string)row[0]).Success)
        .OrderBy(row => row[0])
        .Select(row => new object[] { row[0], GetAddress((string)row[1]) })
      ;
      return tables.ToList();
    }
    public List<object[]> Query(string query, object[] paramValues, object[] paramNames, bool headings) {

      CheckDisposed();

      SQLiteCommand cmd = null;
      bool dispose = false;

      try {

        (cmd, dispose) = GetCommand(query, paramValues, paramNames);

        using (var dr = cmd.ExecuteReader()) {

          var records = new List<object[]>();

          var fc = dr.FieldCount;

          if (headings) {
            var head = new object[fc];
            for (var i = 0; i < fc; ++i)
              head[i] = dr.GetName(i);
            records.Add(head);
          }

          while (dr.Read()) {
            var rec = new object[fc];
            dr.GetValues(rec);
            records.Add(rec);
          }

          dr.Close();

          return records;

        }

      }
      finally {
        if (dispose)
          cmd.Dispose();
      }

    }
    public void UpdateTable(string tblName, XLRefModule.UpdateCommand uc) {

      CheckDisposed();

      if (!ExistsTable(tblName))
        throw new ArgumentException($"Virtual table {tblName} does not exist.");

      using (var cmd = db.CreateCommand()) {
        cmd.CommandText = $"INSERT INTO {tblName} (rowid) VALUES ({(long)uc})";
        try {
          cmd.ExecuteNonQuery();
        }
        catch (SQLiteException sex) {
          var msg = sex.Message;
        }
      }
    }

    public void Dispose() {
      if (disposed)
        return;
      foreach (var cmd in commands.Values)
        cmd.Dispose();
      db.Close();
      db.Dispose();
      disposed = true;
    }
    private void CheckDisposed() {
      if (disposed)
        throw new ObjectDisposedException(GetType().Name);
    }

    (SQLiteCommand, bool) GetCommand(string query, object[] paramValues, object[] paramNames) {

      SQLiteCommand cmd;
      var dispose = false;

      if (query[0] == '$') {
        if (!commands.TryGetValue(query, out cmd))
          throw new KeyNotFoundException(Strings.QRY_NOTFOUND);
      }
      else {
        cmd = db.CreateCommand();
        cmd.CommandText = query;
        cmd.Parameters.NoCase = true;
        dispose = true;
      }

      if (paramValues is null)
        return (cmd, dispose);

      var pc = cmd.Parameters;
      var np = paramValues.Length;
      var nm = paramNames != null;

      if (!(dispose || pc.Count == np))
        throw new ArgumentException("Invalid number of parameters.");

      if (nm && paramNames.Length != np)
        throw new ArgumentException("Invalid number of parameter names.");

      for (var i = 0; i < np; ++i) {
        var name = (paramNames?[i] as string)?.Trim();
        if (nm & String.IsNullOrEmpty(name))
          throw new ArgumentException($"Invalid parameter name '{paramNames[i]}' (index={i}).");
        if (dispose)
          pc.AddWithValue(name, paramValues[i]);
        else if (nm)
          pc[name].Value = paramValues[i];
        else
          pc[i].Value = paramValues[i];
      }

      return (cmd, dispose);

    }

  }
}
