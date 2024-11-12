using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelDna.Integration.Helpers;

namespace XLSQL
{
  [ExcelFunction(Prefix = "SQL.")]
  public static class SheetFunctions
  {

    static readonly Regex dtFormat = new Regex(@"^\d{4}-\d{2}-\d{2}", RegexOptions.Compiled);
    internal static bool IsError(object ans) {
      if (ans is string s && s != null)
        return !dtFormat.Match(s).Success;
      return true;
    }

    [ExcelFunction("Open a connection to a SQLite database.")]
    public static object OpenConnection(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Database file path (the file must exist). If omitted, a memory database is opened.")] string DbFile,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("Read only flag. Default: FALSE.")] bool ReadOnly,
      [ExcelArgument("Don't load sqlite extensions. Default: FALSE.")] bool NoExt,
      [ExcelArgument("Close and replace the connection if it already exists. Default: FALSE.")] bool Overwrite
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        DbPool.Open(CName, DbFile, ReadOnly, !NoExt, Overwrite);
        return DateTime.Now.ToString(Configuration.DateFormat);
      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("Close a SQLite database connection.")]
    public static object CloseConnection(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        DbPool.Close(CName);
        return DateTime.Now.ToString(Configuration.DateFormat);
      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("List open connections whose name matches the regex pattern.")]
    public static object ListConnections(
      [ExcelArgument("Regex pattern filter. Default: '.*'.")] string Regex,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("Disable the output size check. Default: FALSE.")] bool NoCheckSize,
      [ExcelArgument("Ensure the result is bidimensional. Default: FALSE.")] bool Ensure2d
    ) {

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        Regex = Regex?.Trim();
        if (String.IsNullOrEmpty(Regex)) Regex = ".*";
        var regex = new Regex(Regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        var dbNames = DbPool.DbNames
          .Where(s => regex.Match(s).Success)
          .OrderBy(s => s)
          .Select(s => new object[] { s })
          .ToList()
        ;
        return ToResult(new Caller(), dbNames, ExcelMissing.Value, !NoCheckSize, Ensure2d);
      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }

    [ExcelFunction("Create a virtual table backed by a range in 'temp' schema.")]
    public static object CreateTable(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Table name.")] string TblName,
      [ExcelArgument(AllowReference = true, Description = "A range containing the underlying data.")] object Data,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("If true, the first row of Data is used for column names. If false, A1 column names are used. Default: FALSE.")] bool Headers,
      [ExcelArgument("If true, the underlying data are frozen. If false, the underlying range is evaluated each time a query is run. Default: FALSE.")] bool Freeze,
      [ExcelArgument("Drop and replace the table if it already exists. Default: FALSE.")] bool Overwrite
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidName(ref TblName))
        return Strings.INVALID_TBNAME;

      if (!(Data is ExcelReference data && (!Headers || data.RowLast > data.RowFirst)))
        return Strings.INVALID_DATA;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        var db = DbPool.Get(CName);
        db.CreateTable(TblName, data, Headers, Freeze, Overwrite);
        return DateTime.Now.ToString(Configuration.DateFormat);
      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("Freeze a virtual table set up using 'CreateTable'.")]
    public static object FreezeTable(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Table name.")] string TblName,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger
    ) {
      return UpdateTable(XLRefModule.UpdateCommand.Freeze, CName, TblName, Trigger);
    }
    [ExcelFunction("Refresh a frozen virtual table set up using 'CreateTable'.")]
    public static object RefreshTable(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Table name.")] string TblName,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger
    ) {
      return UpdateTable(XLRefModule.UpdateCommand.Refresh, CName, TblName, Trigger);
    }
    [ExcelFunction("Unfreeze a virtual table set up using 'CreateTable'.")]
    public static object UnfreezeTable(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Table name.")] string TblName,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger
    ) {
      return UpdateTable(XLRefModule.UpdateCommand.Unfreeze, CName, TblName, Trigger);
    }
    [ExcelFunction("List virtual tables whose name matches the regex pattern.")]
    public static object ListTables(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Regex pattern for name filter. Default: '.*'.")] string NameRegex,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("Disable the output size check. Default: FALSE.")] bool NoCheckSize,
      [ExcelArgument("Ensure the result is bidimensional. Default: FALSE.")] bool Ensure2d
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        var db = DbPool.Get(CName);
        var tables = db.ListTables(NameRegex);
        return ToResult(new Caller(), tables, ExcelMissing.Value, !NoCheckSize, Ensure2d);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    static object UpdateTable(XLRefModule.UpdateCommand command, string CName, string TblName, object Trigger) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidName(ref TblName))
        return Strings.INVALID_TBNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        var db = DbPool.Get(CName);
        db.UpdateTable(TblName, command);
        return DateTime.Now.ToString(Configuration.DateFormat);
      }
      catch (Exception ex) {
        return ToResult(ex);
      }
    }

    [ExcelFunction("Create a reusable, possibly parameterized, query.")]
    public static object CreateQuery(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Query name.")] string QName,
      [ExcelArgument("The query statement.")] string Query,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("The number of query parameters or a one-dimensional array of parameter names. Default: MISSING.")] object[] ParamNames,
      [ExcelArgument("Dispose and replace the query if it already exists. Default: FALSE.")] bool Overwrite
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidQuery(ref QName))
        return Strings.INVALID_QNAME;

      if (DbPool.InvalidQuery(ref Query))
        return Strings.INVALID_QUERY;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        if (ParamNames?[0] is ExcelMissing)
          ParamNames = null;
        else if (ParamNames?[0] is double d)
          ParamNames[0] = (int)d;

        var db = DbPool.Get(CName);
        db.CreateQuery(QName, Query, ParamNames, Overwrite);
        return DateTime.Now.ToString(Configuration.DateFormat);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("Delete a reusable query.")]
    public static object DeleteQuery(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Query name.")] string QName,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidQuery(ref QName))
        return Strings.INVALID_QNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        var db = DbPool.Get(CName);
        db.DeleteQuery(QName);
        return DateTime.Now.ToString(Configuration.DateFormat);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("List reusable queries whose name or sql text matches the corresponding regex pattern.")]
    public static object ListQueries(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Regex pattern for name filter. Default: '.*'.")] string NameRegex,
      [ExcelArgument("Regex pattern for sql filter. Default: '.*'.")] string SqlRegex,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("Disable the output size check. Default: FALSE.")] bool NoCheckSize,
      [ExcelArgument("Ensure the result is bidimensional. Default: FALSE.")] bool Ensure2d
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        var db = DbPool.Get(CName);
        var queries = db.ListQueries(NameRegex, SqlRegex);
        return ToResult(new Caller(), queries, ExcelMissing.Value, !NoCheckSize, Ensure2d);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }

    [ExcelFunction("Execute a non-query statement and return the number of rows affected by it.")]
    public static object Execute(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("The non-query statement to be executed.")] string Query,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("A one-dimensional array of query parameter values. Default: MISSING.")] object[] ParamValues,
      [ExcelArgument("A one-dimensional array of query parameter names. Default: MISSING.")] object[] ParamNames
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidQuery(ref Query))
        return Strings.INVALID_QUERY;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        if (ParamValues?[0] == ExcelMissing.Value)
          ParamValues = null;

        if (ParamNames?[0] == ExcelMissing.Value)
          ParamNames = null;

        var db = DbPool.Get(CName);
        var ans = db.Execute(Query, ParamValues, ParamNames);
        if (ans >= 0)
          return ans;
        return DateTime.Now.ToString(Configuration.DateFormat);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("Run a query and return a result set.")]
    public static object Query(
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("The query statement to be run.")] string Query,
      [ExcelArgument("Calculation trigger. Default: MISSING.")] object Trigger,
      [ExcelArgument("A one-dimensional array of query parameter values. Default: MISSING.")] object[] ParamValues,
      [ExcelArgument("A one-dimensional array of query parameter names. Default: MISSING.")] object[] ParamNames,
      [ExcelArgument("If true, a row with column names is prepended to the result set. Default: FALSE.")] bool Headings,
      [ExcelArgument("Replacement value for NULLs. Default: #N/A Excel error.")] object IfNull,
      [ExcelArgument("Disable the output size check. Default: FALSE.")] bool NoCheckSize,
      [ExcelArgument("Ensure the result is bidimensional. Default: FALSE.")] bool Ensure2d
    ) {

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (DbPool.InvalidQuery(ref Query))
        return Strings.INVALID_QUERY;

      if (Trigger is ExcelError)
        return Trigger;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {

        if (ParamValues?[0] == ExcelMissing.Value)
          ParamValues = null;

        if (ParamNames?[0] == ExcelMissing.Value)
          ParamNames = null;

        var db = DbPool.Get(CName);
        var records = db.Query(Query, ParamValues, ParamNames, Headings);
        return ToResult(new Caller(), records, IfNull, !NoCheckSize, Ensure2d);

      }
      catch (Exception ex) {
        return ToResult(ex);
      }

    }
    [ExcelFunction("Run a query on multiple ranges and return a result set.\nAt most 9 ranges are supported. Tables are named T1, T2 ...")]
    public static object QueryRange(
      [ExcelArgument("The query statement to be run.")] string Query,
      [ExcelArgument("A table of function options. Default: MISSING.")] object[,] Options,
      [ExcelArgument("A one-dimensional array of query parameter values. Default: MISSING.")] object[] ParamValues,
      [ExcelArgument("A one-dimensional array of query parameter names. Default: MISSING.")] object[] ParamNames,
      [ExcelArgument(AllowReference = true, Description = "The range backing T1.")] object Table1,
      [ExcelArgument(AllowReference = true, Description = "The range backing T2.", Name = "...")] object Table2,
      [ExcelArgument(AllowReference = true, Description = "The range backing T3.")] object Table3,
      [ExcelArgument(AllowReference = true, Description = "The range backing T4.")] object Table4,
      [ExcelArgument(AllowReference = true, Description = "The range backing T5.")] object Table5,
      [ExcelArgument(AllowReference = true, Description = "The range backing T6.")] object Table6,
      [ExcelArgument(AllowReference = true, Description = "The range backing T7.")] object Table7,
      [ExcelArgument(AllowReference = true, Description = "The range backing T8.")] object Table8,
      [ExcelArgument(AllowReference = true, Description = "The range backing T9.")] object Table9
    ) {

      if (String.IsNullOrWhiteSpace(Query))
        return Strings.INVALID_QUERY;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      // Parse options
      object ifnull = ExcelMissing.Value;
      bool headers = false, headings = false, checksize = true, e2d = false;
      if (Options.Length > 1) {
        if (Options.GetLength(1) != 2)
          return "#INVALID_OPTABLE";
        for (var i = 0; i < Options.GetLength(0); ++i) {
          if (Options[i, 0] is string k) {
            try {
              switch (k.Trim().ToLowerInvariant()) {
                case "headers": headers = Convert.ToBoolean(Options[i, 1]); break;
                case "headings": headings = Convert.ToBoolean(Options[i, 1]); break;
                case "nochecksize": checksize = !Convert.ToBoolean(Options[i, 1]); break;
                case "ensure2d": e2d = Convert.ToBoolean(Options[i, 1]); break;
                case "ifnull": ifnull = Options[i, 1]; break;
                default: return $"#INVALID_OPKEY{{{k}}}";
              }
            }
            catch /*(Exception ex)*/ {
              return $"#INVALID_OPVAL{{{k}: {Options[i, 1]}}}";
            }
          }
          else return $"#INVALID_OPKEY{{{Options[i, 0]}}}";
        }
      }

      var tables = new object[] {
        Table1, Table2, Table3, Table4, Table5, Table6, Table7, Table8, Table9
      };

      try {

        if (ParamValues?[0] == ExcelMissing.Value)
          ParamValues = null;

        if (ParamNames?[0] == ExcelMissing.Value)
          ParamNames = null;

        using (var db = new DbAdapter(null, false, true)) {

          for (var i = 0; i < tables.Length; ++i) {
            if (tables[i] == ExcelMissing.Value)
              continue;
            if (!(tables[i] is ExcelReference data && (!headers || data.RowLast > data.RowFirst)))
              continue;
            db.CreateTable($"T{1 + i}", data, headers, true, false);
          }

          var records = db.Query(Query, ParamValues, ParamNames, headings);
          return ToResult(new Caller(), records, ifnull, checksize, e2d);

        }

      }
      catch (Exception ex) {
        return ToResult(ex);
      }
    }

    static object ToResult(Caller caller, List<object[]> rows, object IfNull, bool checkSize, bool Ensure2d) {

      var r = rows.Count;

      if (r == 0)
        return ExcelError.ExcelErrorNull;

      var c = rows[0].Length;

      if (checkSize && (caller.Rows * caller.Columns > 1)) {
        if (caller.TooSmall(false, r, out var msg)) return msg;
        if (caller.TooSmall(true, c, out msg)) return msg;
      }

      if (caller.IsRange && (IfNull == ExcelMissing.Value || IfNull == ExcelEmpty.Value))
        IfNull = ExcelError.ExcelErrorNA;

      var rr = Ensure2d && r == 1;
      var cc = Ensure2d && c == 1;

      var ans = new object[r + (rr ? 1 : 0), c + (cc ? 1 : 0)];

      for (var i = 0; i < r; ++i) {
        var row = rows[i];
        for (var j = 0; j < c; ++j)
          ans[i, j] = row[j] == DBNull.Value ? IfNull : row[j];
        if (cc) ans[i, c] = ExcelError.ExcelErrorNA;
      }
      if (rr) {
        if (cc) ++c;
        for (var j = 0; j < c; ++j)
          ans[r, j] = ExcelError.ExcelErrorNA;
      }

      return ans;

    }
    static string ToResult(Exception ex) {
      if (ex is System.Data.SQLite.SQLiteException sex) {
        var msg = sex.Message;
        return $"{Char.ToUpper(msg[0])}{msg.Substring(1).Replace(Environment.NewLine, ". ")}";
      }
      return ex.Message;
    }

  }
}
