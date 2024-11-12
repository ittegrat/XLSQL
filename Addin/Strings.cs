
namespace XLSQL
{
  internal static class Strings
  {

    public const string ALREADY_EXISTS = "#ALREADY_EXISTS";
    public const string DB_NOTFOUND = "#DB_NOT_FOUND";
    public const string FUNC_WIZARD = "#FUNC_WIZARD";
    public const string INVALID_DATA = "#INVALID_DATA";
    public const string INVALID_DBFILE = "#INVALID_DBFILE";
    public const string INVALID_CNAME = "#CNX_NAME!";
    public const string INVALID_QUERY = "#INVALID_QUERY";
    public const string INVALID_QNAME = "#QRY_NAME!";
    public const string INVALID_TBNAME = "#TBL_NAME!";
    public const string QRY_NOTFOUND = "#QRY_NOT_FOUND";

    public static string ERR(string s) {
      return $"#ERR{{{s}}}";
    }

  }
}
