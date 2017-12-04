namespace RobsonRocha.Exemplos.OpenXml
{
    /// <summary>
    /// Contains information about a sheet cell
    /// </summary>
    public class CellInfo
    {
        /// <summary>
        /// The row which this cell belongs
        /// </summary>
        public RowInfo RowInfo { get; }

        /// <summary>
        /// The cell reference within the sheet (i.e.: A42)
        /// </summary>
        public string Reference { get; }

        /// <summary>
        /// The column name which this cell belongs, if identified
        /// </summary>
        public string Column { get; }

        /// <summary>
        /// The cell value
        /// </summary>
        public string Value { get; }

        internal CellInfo(RowInfo rowInfo, string reference, string column, string value)
        {
            RowInfo = rowInfo;
            Reference = reference;
            Column = column;
            Value = value;
        }

        public override string ToString() => Value;
    }
}
