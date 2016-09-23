using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

[assembly: AssemblyTitle("tabler")]
[assembly: AssemblyProduct("tabler")]
[assembly: AssemblyCopyright("Copyright © Ryan Parnas 2016")]

namespace tabler
{
  class Program
  {
    [STAThread]
    static void Main(string[] args)
    {
      var conversionFunctions = new Dictionary<string, Func<DataTable, string>>
      {
        { "md", ConvertToMD },
        { "html", ConvertToHTML }
      };

      if (args.Length != 3 || !conversionFunctions.ContainsKey(args[0]))
      {
        Console.WriteLine("Usage: tabler (md | html) <file> <worksheet>");
        return;
      }

      var outputFormat = args[0];
      var path = args[1];
      var sheetName = args[2];
      var table = GetXLSX(path, sheetName);
      Clipboard.SetText(conversionFunctions[outputFormat](table));
    }

    /// <summary>Converts the given datatable to GitHub Markdown.
    static string ConvertToMD(DataTable table)
    {
      Func<string, string> escapeText = s =>
      {
        var trimmed = s.Trim().Replace("|", "&#124;");
        return string.IsNullOrWhiteSpace(trimmed) ? "&nbsp;" : trimmed;
      };

      var sb = new StringBuilder();

      // Header row
      for (int col = 0; col < table.Columns.Count; col++)
      {
        sb.Append("|" + escapeText(table.Columns[col].ColumnName));
      }
      sb.AppendLine("|");

      // Header row border
      for (int col = 0; col < table.Columns.Count; col++)
      {
        sb.Append("|---");
      }
      sb.AppendLine("|");

      // Other rows
      for (int row = 0; row < table.Rows.Count; row++)
      {
        for (int col = 0; col < table.Columns.Count; col++)
        {
          sb.Append("|" + escapeText(table.Rows[row][col].ToString()));
        }
        sb.AppendLine("|");
      }

      return sb.ToString();
    }

    /// <summary>Converts the given datatable to HTML.</summary>
    static string ConvertToHTML(DataTable table)
    {
      Func<string, string> escapeText = s => s.Trim();

      var sb = new StringBuilder();
      sb.AppendLine("<table>");

      // Header row
      sb.AppendLine("  <tr>");
      for (int col = 0; col < table.Columns.Count; col++)
      {
        sb.AppendLine("    <th>" + escapeText(table.Columns[col].ColumnName) + "</th>");
      }
      sb.AppendLine("  </tr>");

      // Other Rows
      for (int row = 0; row < table.Rows.Count; row++)
      {
        sb.Append("  <tr>");
        for (int col = 0; col < table.Columns.Count; col++)
        {
          sb.Append("<td>" + escapeText(table.Rows[row][col].ToString().Trim()) + "</td>");
        }
        sb.AppendLine("  </tr>");
      }

      sb.AppendLine("</table>");
      return sb.ToString();
    }

    /// <summary>Retrieves an excel worksheet as a DataTable.</summary>
    static DataTable GetXLSX(string path, string sheetName)
    {
      var ep = new ExcelPackage(new FileInfo(path));
      foreach (var worksheet in ep.Workbook.Worksheets)
      {
        if (worksheet.Name == sheetName && worksheet.Dimension != null)
          return SheetToTable(worksheet);
      }
      return null;
    }

    /// <summary>Converts an excel worksheet to a DataTable.</summary>
    static DataTable SheetToTable(ExcelWorksheet ws)
    {
      var ret = new DataTable(ws.Name);

      for (int col = 1; col <= ws.Dimension.End.Column; col++)
      {
        var cell = ws.Cells[1, col];
        var text = cell.Text.Replace('\n', '_').Replace('\r', '_');
        ret.Columns.Add(string.IsNullOrWhiteSpace(text) ? cell.Address : text);
      }

      for (int row = 2; row <= ws.Dimension.End.Row; row++)
      {
        var newRow = ret.NewRow();

        for (int col = 1; col <= ws.Dimension.End.Column; col++)
          newRow[col - 1] = ws.Cells[row, col].Text;

        ret.Rows.Add(newRow);
      }

      return ret;
    }
  }
}
