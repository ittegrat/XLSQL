using System;
using ExcelDna.Integration;

namespace XLSQL
{
  [ExcelCommand(Prefix = "XLSQL.")]
  public static class Commands
  {

    [ExcelCommand("Create a SQLite database file and open a connection to it.")]
    public static object CreateDatabase(
      [ExcelArgument("Database file path.")] string DbFile,
      [ExcelArgument("Connection name.")] string CName,
      [ExcelArgument("Replace the file if it already exists. Default: FALSE.")] bool FileOverwrite,
      [ExcelArgument("Close and replace the connection if it already exists. Default: FALSE.")] bool ConnOverwrite,
      [ExcelArgument("Don't load sqlite extensions. Default: FALSE.")] bool NoExt
    ) {

      DbFile = DbFile?.Trim();
      if (String.IsNullOrEmpty(DbFile))
        return Strings.INVALID_DBFILE;

      if (DbPool.InvalidName(ref CName))
        return Strings.INVALID_CNAME;

      if (ExcelDnaUtil.IsInFunctionWizard())
        return Strings.FUNC_WIZARD;

      try {
        DbPool.Create(DbFile, CName, FileOverwrite, ConnOverwrite, !NoExt);
        return DateTime.Now.ToString(Configuration.DateFormat);
      }
      catch (Exception ex) {
        return ex.Message;
      }

    }

  }
}
