using System;
using System.Collections.Specialized;
using System.Configuration;
using System.Linq;

namespace XLSQL
{
  public static class Configuration
  {

    const string XLSQL = "xlsql";

    public static string DateFormat { get; } = "yyyy-MM-dd HH:mm:ss.fff";
    public static object[] Missing { get; } = new object[] { };

    public static string RefreshError { get; private set; }

    public static bool ExtensionsEnabled { get; set; }
    public static bool HiddenRibbonTab { get; set; }
    public static string[] Extensions { get; set; }

    public static void Refresh() {
      try {

        ConfigurationManager.RefreshSection(XLSQL);
        var xlsql = ConfigurationManager.GetSection(XLSQL) as NameValueCollection;

        T GetValue<T>(string key, T @default) {
          var value = xlsql[key];
          return value is null ? @default : (T)Convert.ChangeType(value, typeof(T));
        }

        ExtensionsEnabled = GetValue("sqlite.extensions.enable", true);
        HiddenRibbonTab = GetValue("ribbon.tab.hidden", false);

        var extensions = xlsql["sqlite.extensions"] ?? String.Empty;
        Extensions = extensions.Split(',').Select(e => "sqlite3_" + e + "_init").ToArray();

      }
      catch (Exception ex) {
        RefreshError = ex.ToString();
      }
    }

    static Configuration() { Refresh(); }

  }
}
