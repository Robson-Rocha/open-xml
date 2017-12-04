namespace RobsonRocha.Exemplos.OpenXml
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Contains information about a workbook sheet
    /// </summary>
    public class SheetInfo
    {
        private List<RowInfo> _rows = new List<RowInfo>();
        private Lazy<List<RowInfo>> _rowsWithHeader;

        internal Dictionary<string, CellInfo> RefDict = new Dictionary<string, CellInfo>();
        internal List<string> ColumnList = new List<string>();

        /// <summary>
        /// The name of the sheet
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// List of all the rows imported for the sheet, including the header (if set) and the lines above and below it
        /// </summary>
        public IReadOnlyList<RowInfo> AllRows => _rows;

        /// <summary>
        /// List of all the rows imported for the sheet that were below the header (if it was set)
        /// </summary>
        public IReadOnlyList<RowInfo> Rows => _rowsWithHeader.Value;

        /// <summary>
        /// List of all column names for this sheet
        /// </summary>
        public IReadOnlyList<string> Columns => ColumnList;

        /// <summary>
        /// Get the <see cref="CellInfo"/> of a cell, given the <paramref name="reference"/> that represents it
        /// </summary>
        /// <param name="reference">The reference (i.e.: A42) of the cell</param>
        /// <returns>If a cell with the given <paramref name="reference"/> exists among the rows of this sheet, returns the <see cref="CellInfo"/> for the cell; otherwise, returns null</returns>
        public CellInfo GetCell(string reference)
        {
            return RefDict.TryGetValue(reference, out CellInfo colData) ? colData : null;
        }

        internal SheetInfo(string name)
        {
            Name = name;
            _rowsWithHeader = new Lazy<List<RowInfo>>(() => _rows?.Where(r => r.Any(c => c.Column != null))?.Skip(1)?.ToList());
        }

        internal RowInfo AddRow()
        {
            RowInfo newRow = new RowInfo(this);
            _rows.Add(newRow);
            return newRow;
        }
    }
}
