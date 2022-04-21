using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace JsonToXlsx
{
  class Config
  {
    [DefaultValue("")]
    public string OutputFile { get; set; }

    [DefaultValue("Sheet1")]
    public string SheetName { get; set; }

    [DefaultValue(0)]
    public int FreezeRows { get; set; }

    [DefaultValue(0)]
    public int FreezeColumns { get; set; }

    [DefaultValue(0)]
    public int HeaderHeight { get; set; }

    [DefaultValue(false)]
    public bool SubHeader { get; set; }

    [DefaultValue(0)]
    public int SubHeaderHeight { get; set; }

    [DefaultValue(0)]
    public int SubHeaderRotation { get; set; }

    [DefaultValue("Center")]
    public string SubHeaderAlignmentH { get; set; }

    [DefaultValue("Center")]
    public string SubHeaderAlignmentV { get; set; }

    public IList<ColumnConfig> Columns { get; set; }

    [JsonIgnore]
    public int RowCount { get; set; }

    public XLWorkbook CreateWorkbook()
    {
      var wb = new XLWorkbook(XLEventTracking.Disabled);
      var ws = wb.AddWorksheet(SheetName);

      var columnRow = ws.Row(1);

      columnRow.Style.Font.Bold = true;
      columnRow.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;

      if (HeaderHeight > 0)
      {
        columnRow.Height = HeaderHeight;
      }

      foreach (var columnConfig in Columns)
      {
        var column = ws.Column(columnConfig.Index);

        columnConfig.ApplyStyle(column);

        var cell = column.Cell(1).SetValue<string>(string.IsNullOrWhiteSpace(columnConfig.Caption) ? "" : columnConfig.Caption);

        columnConfig.ApplyStyle(cell);

        if (columnConfig.MergeH > 0)
        {
          ws.Range(1, columnConfig.Index, 1, columnConfig.Index + columnConfig.MergeH).Row(1).Merge();
        }

        if (columnConfig.MergeV > 0)
        {
          ws.Range(1, columnConfig.Index, 1 + columnConfig.MergeV, columnConfig.Index).Column(1).Merge();
        }
      }

      if (SubHeader)
      {
        var subHeaderRow = ws.Row(2);

        if (SubHeaderRotation >= 0)
        {
          subHeaderRow.Style.Alignment.SetTextRotation(SubHeaderRotation);
        }

        if (SubHeaderHeight > 0)
        {
          subHeaderRow.Height = SubHeaderHeight;
        }

        if (!string.IsNullOrWhiteSpace(SubHeaderAlignmentH))
        {
          subHeaderRow.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)Enum.Parse(typeof(XLAlignmentHorizontalValues), SubHeaderAlignmentH));
        }

        if (!string.IsNullOrWhiteSpace(SubHeaderAlignmentV))
        {
          subHeaderRow.Style.Alignment.SetVertical((XLAlignmentVerticalValues)Enum.Parse(typeof(XLAlignmentVerticalValues), SubHeaderAlignmentV));
        }

        subHeaderRow.SetDataType(XLDataType.Text);
      }

      ws.SheetView.FreezeColumns(FreezeColumns);
      ws.SheetView.FreezeRows(FreezeRows);

      return wb;
    }
  }

  class ColumnConfig
  {
    private static DateTime UNIX_EPOCH_START = new DateTime(1970, 1, 1, 0, 0, 0, 0, DateTimeKind.Utc);

    [JsonIgnore]
    public int Index { get; set; }

    public string Name { get; set; }

    [DefaultValue("")]
    public string Caption { get; set; }

    [DefaultValue("string")]
    public string Type { get; set; }

    [DefaultValue(0)]
    public int Width { get; set; }

    [DefaultValue(0)]
    public int HeaderRotation { get; set; }

    [DefaultValue("Left")]
    public string HeaderAlignmentH { get; set; }

    [DefaultValue("Top")]
    public string HeaderAlignmentV { get; set; }

    [DefaultValue(0)]
    public int MergeH { get; set; }

    [DefaultValue(0)]
    public int MergeV { get; set; }

    [DefaultValue("")]
    public string FontColor { get; set; }

    public void ApplyStyle(IXLColumn column)
    {
      switch (Type)
      {
        case "percent":
          column.Style.NumberFormat.Format = "0%";
          column.Width = 7;
          break;

        case "integer":
          column.Style.NumberFormat.Format = "# ##0;-# ##0;0";
          column.Width = 10;
          break;

        case "decimal":
          column.Style.NumberFormat.Format = "# ##0.0##;_-# ##0.0##;0";
          column.Width = 10;
          break;

        case "datetime":
        case "datetime+utc":
          column.Style.DateFormat.Format = "dd.mm.yyyy hh:mm:ss";
          column.Width = 20;
          break;

        case "date":
        case "date+utc":
          column.Style.DateFormat.Format = "dd.mm.yyyy";
          column.Width = 10;
          break;

        case "time":
        case "time+utc":
          column.Style.DateFormat.Format = "hh:mm:ss";
          column.Width = 10;
          break;
      }

      if (Width > 0)
      {
        column.Width = Width;
      }

      if (!string.IsNullOrWhiteSpace(FontColor))
      {
        column.Style.Font.FontColor = FontColor.StartsWith("#") ? XLColor.FromHtml(FontColor) : XLColor.FromName(FontColor);
      }
    }

    public void ApplyStyle(IXLCell cell)
    {
      if (HeaderRotation >= 0)
      {
        cell.Style.Alignment.SetTextRotation(HeaderRotation);
      }

      if (!string.IsNullOrWhiteSpace(HeaderAlignmentH))
      {
        cell.Style.Alignment.SetHorizontal((XLAlignmentHorizontalValues)Enum.Parse(typeof(XLAlignmentHorizontalValues), HeaderAlignmentH));
      }

      if (!string.IsNullOrWhiteSpace(HeaderAlignmentV))
      {
        cell.Style.Alignment.SetVertical((XLAlignmentVerticalValues)Enum.Parse(typeof(XLAlignmentVerticalValues), HeaderAlignmentV));
      }

      if ((MergeH > 0 || MergeV > 0) && !string.IsNullOrWhiteSpace(FontColor))
      {
        cell.Style.Font.FontColor = XLColor.Black;
      }
    }

    public void CreateCell(IXLCell cell, object value, string type)
    {
      if (value == null)
      {
        return;
      }

      try
      {
        switch (string.IsNullOrWhiteSpace(type) ? Type : type)
        {
          case "integer":
            cell.SetValue<int>(Convert.ToInt32(value));
            break;

          case "decimal":
          case "percent":
            cell.SetValue<decimal>(Convert.ToDecimal(value));
            break;

          case "datetime":
          case "date":
          case "time":
            cell.SetValue<DateTime>(value is long ? FromUnixTimestamp(value) : (DateTime)value);
            break;

          case "datetime+utc":
          case "date+utc":
          case "time+utc":
            cell.SetValue<DateTime>(((value is long ? FromUnixTimestamp(value) : (DateTime)value)).ToUniversalTime());
            break;

          case "boolean":
            cell.SetValue<bool>(Convert.ToBoolean(value));
            break;

          default:
            cell.SetValue<string>(value.ToString());
            break;
        }
      }
      catch { }
    }

    public DateTime FromUnixTimestamp(object value)
    {
      return UNIX_EPOCH_START.AddMilliseconds((long)value).ToLocalTime();
    }
  }

  class Program
  {
    private static int RowI = 2;

    static void Main(string[] args)
    {
      Console.InputEncoding = Encoding.UTF8;

      Config config = null;
      XLWorkbook wb = null;
      IXLWorksheet ws = null;

      var settings = new JsonSerializerSettings()
      {
        DateTimeZoneHandling = DateTimeZoneHandling.Local
      };

      do
      {
        var line = Console.ReadLine();

        if (string.IsNullOrWhiteSpace(line))
        {
          break;
        }

        if (config == null)
        {
          config = TryParseConfig(line);

          if (config != null)
          {
            wb = config.CreateWorkbook();
            ws = wb.Worksheet(1);
          }
        }
        else
        {
          TryParseRow(config, ws, settings, line);
        }
      }
      while (true);

      if (wb == null)
      {
        wb = new XLWorkbook();
        wb.AddWorksheet("Sheet1").Column(0).Value = "Column1";
      }

      if (string.IsNullOrWhiteSpace(config.OutputFile))
      {
        wb.SaveAs(Console.OpenStandardOutput(), false, false);
      }
      else
      {
        wb.SaveAs(config.OutputFile, false, false);

        Console.WriteLine(config.OutputFile);
      }

      Environment.Exit(0);
    }

    private static Config TryParseConfig(string json)
    {
      try
      {
        var config = JsonConvert.DeserializeObject<Config>(json);

        for (var i = 0; i < config.Columns.Count; ++i)
        {
          config.Columns[i].Index = i + 1;
        }

        return config;
      }
      catch (Exception x)
      {
        Console.Error.WriteLine(x.ToString());

        return null;
      }
    }

    private static void TryParseRow(Config config, IXLWorksheet ws, JsonSerializerSettings settings, string json)
    {
      try
      {
        var data = JsonConvert.DeserializeObject<Dictionary<string, object>>(json, settings);
        var row = ws.Row(RowI);

        for (var i = 0; i < config.Columns.Count; ++i)
        {
          var columnI = i + 1;
          var column = config.Columns[i];
          var value = data.ContainsKey(column.Name) ? data[column.Name] : null;

          column.CreateCell(row.Cell(columnI), value, config.SubHeader && RowI == 2 ? "string" : null);
        }

        RowI += 1;
      }
      catch (Exception x)
      {
        Console.Error.WriteLine(x.ToString());
      }
    }
  }
}
