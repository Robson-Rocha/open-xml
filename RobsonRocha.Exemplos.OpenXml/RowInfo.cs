namespace RobsonRocha.Exemplos.OpenXml
{
    using System.Collections.Generic;
    using System.Collections;
    using System.Dynamic;

    /// <summary>
    /// Contains information about a sheet row
    /// </summary>
    public class RowInfo : DynamicObject, IEnumerable<CellInfo>
    {
        private readonly Dictionary<string, CellInfo> _colDict = new Dictionary<string, CellInfo>();
        private readonly List<CellInfo> _colList = new List<CellInfo>();
        private readonly SheetInfo _sheetInfo;

        /// <summary>
        /// Returns the <see cref="CellInfo"/> about a column within the row, given her name
        /// </summary>
        /// <param name="columnName">The name of the column which value must be retrieved</param>
        /// <returns>If a column with the given <paramref name="columnName"/> exists among the cells of this row, returns the <see cref="CellInfo"/> for the cell; otherwise, returns null</returns>
        public CellInfo this[string columnName] => _colDict.TryGetValue(columnName, out CellInfo value) ? value : null;

        /// <summary>
        /// Returns the <see cref="CellInfo"/> about a column within the row, given her index
        /// </summary>
        /// <param name="index">The of the column which value must be retrieved</param>
        /// <returns>If a column with the given <paramref name="index"/> exists among the cells of this row, returns the <see cref="CellInfo"/> for the cell; otherwise, throws a <see cref="ArgumentOutOfRangeException"/></returns>
        public CellInfo this[int index] => _colList[index];

        internal RowInfo(SheetInfo sheetInfo)
        {
            _sheetInfo = sheetInfo;
        }

        internal void AddColumn(string reference, string column, string data)
        {
            CellInfo colData = new CellInfo(this, reference, column, data);
            _sheetInfo.RefDict.Add(reference, colData);
            if(!string.IsNullOrWhiteSpace(column))
                _colDict.Add(column, colData);
            _colList.Add(colData);

            //Debug.WriteLine($"Sheet: '{_sheetInfo.Name}', Cell: '{reference}', Column: '{column}', Data: '{data}'");
        }

        public override bool TryGetMember(GetMemberBinder binder, out object result)
        {
            bool found = _colDict.TryGetValue(binder.Name, out CellInfo cell);
            result = found ? cell.Value : null;
            return found;
        }

        public IEnumerator<CellInfo> GetEnumerator()
        {
            foreach (CellInfo colData in _colList)
                yield return colData;
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
