using System;
using System.IO;
using System.Text;
using System.Drawing;
using FastReport.Utils;
using FastReport.Table;

namespace FastReport.Export.Json
{
    /// <summary>
    /// Exports a prepared report to a JSON description suitable for receipt printers.
    /// </summary>
    public class JsonExport : ExportBase
    {
        private StreamWriter writer;
        private bool hasPage;
        private bool hasObject;
        private int pageIndex;

        /// <inheritdoc/>
        protected override string GetFileFilter()
        {
            return "Json file (*.json)|*.json";
        }

        /// <inheritdoc/>
        protected override void Start()
        {
            base.Start();
            writer = new StreamWriter(Stream, new UTF8Encoding(false), 4096, true);
            writer.Write("{\"unit\":\"mm\",\"pages\":[");
            hasPage = false;
            pageIndex = 0;
        }

        /// <inheritdoc/>
        protected override void ExportPageBegin(ReportPage page)
        {
            base.ExportPageBegin(page);

            if (hasPage)
                writer.Write(',');

            hasPage = true;
            hasObject = false;
            pageIndex++;

            float pageWidth = ExportUtils.GetPageWidth(page);
            float pageHeight = ExportUtils.GetPageHeight(page);

            writer.Write("{\"index\":");
            writer.Write(pageIndex);
            writer.Write(",\"width\":");
            writer.Write(ExportUtils.FloatToString(pageWidth));
            writer.Write(",\"height\":");
            writer.Write(ExportUtils.FloatToString(pageHeight));
            writer.Write(",\"margins\":{");
            writer.Write("\"left\":");
            writer.Write(ExportUtils.FloatToString(page.LeftMargin));
            writer.Write(",\"top\":");
            writer.Write(ExportUtils.FloatToString(page.TopMargin));
            writer.Write(",\"right\":");
            writer.Write(ExportUtils.FloatToString(page.RightMargin));
            writer.Write(",\"bottom\":");
            writer.Write(ExportUtils.FloatToString(page.BottomMargin));
            writer.Write("},\"objects\":[");
        }

        /// <inheritdoc/>
        protected override void ExportBand(BandBase band)
        {
            base.ExportBand(band);

            if (band == null)
                return;

            foreach (ReportComponentBase component in band.Objects)
            {
                ExportComponent(component);
            }
        }

        /// <inheritdoc/>
        protected override void ExportPageEnd(ReportPage page)
        {
            if (writer == null)
                return;

            writer.Write("]}");
        }

        /// <inheritdoc/>
        protected override void Finish()
        {
            if (writer == null)
                return;

            writer.Write("]}");
            writer.Flush();
            writer = null;
        }

        private void ExportComponent(ReportComponentBase component)
        {
            if (component == null)
                return;

            if (!component.Exportable)
                return;

            WriteObject(component);

            if (component is IParent parent)
            {
                ObjectCollection children = new ObjectCollection();
                parent.GetChildObjects(children);
                foreach (Base child in children)
                {
                    if (child is ReportComponentBase childComponent)
                        ExportComponent(childComponent);
                }
            }
        }

        private void WriteObject(ReportComponentBase component)
        {
            if (hasObject)
                writer.Write(',');

            hasObject = true;

            writer.Write("{");
            bool hasProperty = false;
            WriteStringProperty("name", component.Name, ref hasProperty);
            WriteStringProperty("type", component.GetType().Name, ref hasProperty);
            WriteNumberProperty("left", ConvertToMillimeters(component.AbsLeft), ref hasProperty);
            WriteNumberProperty("top", ConvertToMillimeters(component.AbsTop), ref hasProperty);
            WriteNumberProperty("width", ConvertToMillimeters(component.Width), ref hasProperty);
            WriteNumberProperty("height", ConvertToMillimeters(component.Height), ref hasProperty);

            if (component is TextObject textObject)
            {
                WriteTextProperties(textObject, ref hasProperty);
            }

            if (component is TableObject tableObject)
            {
                WriteTableObject(tableObject, ref hasProperty);
            }

            if (component.Border != null)
            {
                WriteBorder(component.Border, ref hasProperty);
            }

            writer.Write("}");
        }

        private void WriteBorder(Border border, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write("\"border\":{");
            bool hasBorderProperty = false;
            WriteNumberProperty("left", ConvertToMillimeters(border.LeftLine.Width), ref hasBorderProperty);
            WriteNumberProperty("top", ConvertToMillimeters(border.TopLine.Width), ref hasBorderProperty);
            WriteNumberProperty("right", ConvertToMillimeters(border.RightLine.Width), ref hasBorderProperty);
            WriteNumberProperty("bottom", ConvertToMillimeters(border.BottomLine.Width), ref hasBorderProperty);
            WriteStringProperty("color", ToColorString(border.Color), ref hasBorderProperty);
            writer.Write("}");
        }

        private void WriteTextProperties(TextObject textObject, ref bool hasProperty)
        {
            WriteStringProperty("text", textObject.Text, ref hasProperty);
            WriteStringProperty("horzAlign", textObject.HorzAlign.ToString(), ref hasProperty);
            WriteStringProperty("vertAlign", textObject.VertAlign.ToString(), ref hasProperty);
            WriteStringProperty("fontName", textObject.Font.Name, ref hasProperty);
            WriteNumberProperty("fontSize", textObject.Font.Size, ref hasProperty);
            WriteStringProperty("fontStyle", textObject.Font.Style.ToString(), ref hasProperty);
            WriteStringProperty("textColor", ToColorString(textObject.TextColor), ref hasProperty);
            WriteStringProperty("fillColor", ToColorString(textObject.FillColor), ref hasProperty);
        }

        private void WriteTableObject(TableObject tableObject, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write("\"table\":{");
            bool hasTableProperty = false;
            WriteIntProperty("rowCount", tableObject.RowCount, ref hasTableProperty);
            WriteIntProperty("columnCount", tableObject.ColumnCount, ref hasTableProperty);
            WriteTableRows(tableObject, ref hasTableProperty);
            WriteTableColumns(tableObject, ref hasTableProperty);
            WriteTableCells(tableObject, ref hasTableProperty);
            writer.Write("}");
        }

        private void WriteTableRows(TableObject tableObject, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write("\"rows\":[");
            bool hasRow = false;
            for (int i = 0; i < tableObject.Rows.Count; i++)
            {
                TableRow row = tableObject.Rows[i];
                if (!row.Visible)
                    continue;

                if (hasRow)
                    writer.Write(',');

                hasRow = true;
                writer.Write("{");
                bool hasRowProperty = false;
                WriteIntProperty("index", i, ref hasRowProperty);
                WriteNumberProperty("height", ConvertToMillimeters(row.Height), ref hasRowProperty);
                writer.Write("}");
            }
            writer.Write("]");
        }

        private void WriteTableColumns(TableObject tableObject, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write("\"columns\":[");
            bool hasColumn = false;
            for (int i = 0; i < tableObject.Columns.Count; i++)
            {
                TableColumn column = tableObject.Columns[i];
                if (!column.Visible)
                    continue;

                if (hasColumn)
                    writer.Write(',');

                hasColumn = true;
                writer.Write("{");
                bool hasColumnProperty = false;
                WriteIntProperty("index", i, ref hasColumnProperty);
                WriteNumberProperty("width", ConvertToMillimeters(column.Width), ref hasColumnProperty);
                writer.Write("}");
            }
            writer.Write("]");
        }

        private void WriteTableCells(TableObject tableObject, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write("\"cells\":[");
            bool hasCell = false;
            for (int rowIndex = 0; rowIndex < tableObject.Rows.Count; rowIndex++)
            {
                TableRow row = tableObject.Rows[rowIndex];
                if (!row.Visible)
                    continue;

                for (int columnIndex = 0; columnIndex < tableObject.Columns.Count; columnIndex++)
                {
                    TableColumn column = tableObject.Columns[columnIndex];
                    if (!column.Visible)
                        continue;

                    TableCell cell = tableObject[columnIndex, rowIndex];
                    if (cell == null || !cell.Exportable)
                        continue;

                    if (cell.Address.X != columnIndex || cell.Address.Y != rowIndex)
                        continue;

                    if (hasCell)
                        writer.Write(',');

                    hasCell = true;
                    WriteTableCell(cell, columnIndex, rowIndex);
                }
            }
            writer.Write("]");
        }

        private void WriteTableCell(TableCell cell, int columnIndex, int rowIndex)
        {
            writer.Write("{");
            bool hasProperty = false;
            WriteIntProperty("row", rowIndex, ref hasProperty);
            WriteIntProperty("column", columnIndex, ref hasProperty);
            WriteIntProperty("rowSpan", cell.RowSpan, ref hasProperty);
            WriteIntProperty("colSpan", cell.ColSpan, ref hasProperty);
            WriteNumberProperty("left", ConvertToMillimeters(cell.AbsLeft), ref hasProperty);
            WriteNumberProperty("top", ConvertToMillimeters(cell.AbsTop), ref hasProperty);
            WriteNumberProperty("width", ConvertToMillimeters(cell.Width), ref hasProperty);
            WriteNumberProperty("height", ConvertToMillimeters(cell.Height), ref hasProperty);
            WriteTextProperties(cell, ref hasProperty);

            if (cell.Border != null)
                WriteBorder(cell.Border, ref hasProperty);

            writer.Write("}");
        }

        private void WriteStringProperty(string name, string value, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write('"');
            writer.Write(name);
            writer.Write("\":");
            writer.Write('"');
            writer.Write(EscapeJson(value ?? string.Empty));
            writer.Write('"');
        }

        private void WriteIntProperty(string name, int value, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write('"');
            writer.Write(name);
            writer.Write("\":");
            writer.Write(value);
        }

        private void WriteNumberProperty(string name, float value, ref bool hasProperty)
        {
            WritePropertySeparator(ref hasProperty);
            writer.Write('"');
            writer.Write(name);
            writer.Write("\":");
            writer.Write(ExportUtils.FloatToString(value));
        }

        private void WritePropertySeparator(ref bool hasProperty)
        {
            if (hasProperty)
                writer.Write(',');

            hasProperty = true;
        }

        private static float ConvertToMillimeters(float value)
        {
            return value / Units.Millimeters;
        }

        private static string ToColorString(Color color)
        {
            return $"#{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
        }

        private static string EscapeJson(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            StringBuilder builder = new StringBuilder(value.Length);
            foreach (char ch in value)
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\b':
                        builder.Append("\\b");
                        break;
                    case '\f':
                        builder.Append("\\f");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        if (char.IsControl(ch))
                        {
                            builder.Append("\\u");
                            builder.Append(((int)ch).ToString("x4"));
                        }
                        else
                        {
                            builder.Append(ch);
                        }
                        break;
                }
            }

            return builder.ToString();
        }
    }
}
