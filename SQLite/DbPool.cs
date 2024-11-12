using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace XLSQL
{
  internal static class DbPool
  {

    static readonly SortedDictionary<string, DbAdapter> dbPool = new SortedDictionary<string, DbAdapter>();
    static readonly Regex cnameCheck = new Regex(@"^\w{1,31}$", RegexOptions.Compiled);
    static readonly Regex qnameCheck = new Regex(@"^\$\w{1,30}$", RegexOptions.Compiled);

    public static IEnumerable<string> DbNames => dbPool.Keys;

    public static void Open(string cName, string dbFile, bool readOnly, bool loadExt, bool overwrite) {
      if (dbPool.ContainsKey(cName)) {
        if (overwrite)
          Close(cName);
        else
          throw new System.Data.DuplicateNameException(Strings.ALREADY_EXISTS);
      }
      dbPool[cName] = new DbAdapter(dbFile, readOnly, loadExt);
    }
    public static void Close(string cName) {
      var db = Get(cName);
      db.Dispose();
      dbPool.Remove(cName);
    }
    public static void Create(string dbFile, string cName, bool fileOverwrite, bool connOverwrite, bool loadExt) {

      if (!Path.IsPathRooted(dbFile))
        throw new ArgumentException($"File '{dbFile}' does not contain a root.");
      if (Path.GetPathRoot(dbFile).Length < 2)
        throw new ArgumentException($"File '{dbFile}' does not contain an absolute root.");
      if (!Path.HasExtension(dbFile))
        dbFile += ".sqlite";
      if (!fileOverwrite && File.Exists(dbFile))
        throw new ArgumentException($"File '{dbFile}' already exists.");

      if (!connOverwrite && dbPool.ContainsKey(cName))
        throw new System.Data.DuplicateNameException(Strings.ALREADY_EXISTS);

      // See 'System.Data.SQLite.SQLiteConnection.CreateFile'
      File.Create(dbFile).Close();
      Open(cName, dbFile, false, loadExt, connOverwrite);

    }
    public static DbAdapter Get(string cName) {
      if (dbPool.TryGetValue(cName, out var db))
        return db;
      throw new KeyNotFoundException(Strings.DB_NOTFOUND);
    }

    public static bool ContainsName(string cName) { return dbPool.ContainsKey(cName); }
    public static bool InvalidName(ref string name) {
      name = name?.Trim().ToLowerInvariant();
      return String.IsNullOrEmpty(name) || !cnameCheck.Match(name).Success;
    }
    public static bool InvalidQuery(ref string name) {
      name = name?.Trim();
      if (String.IsNullOrEmpty(name))
        return true;
      if (name[0] != '$')
        return false;
      name = name.ToLowerInvariant();
      return !qnameCheck.Match(name).Success;
    }

  }
}
