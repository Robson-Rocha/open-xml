namespace RobsonRocha.Exemplos.OpenXml
{
    using System.Collections.Generic;

    /// <summary>
    /// Contains options for configuring <see cref="XlsxReader"/> behavior
    /// </summary>
    public class ReadXlsxOptions
    {
        /// <summary>
        /// The name of the sheet which these options must be applied
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// If true, the sheet specified in <see cref="SheetName"/> will not be read
        /// </summary>
        public bool Ignore { get; set; }

        /// <summary>
        /// If set to an value greater than 0, indicates the row which defines the header names.
        /// If the header names are set using <see cref="Headers"/>, this property value must be 0 to avoid overriding
        /// </summary>
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// If set to an value greater than 0, indicates the row which starts the columnar data.
        /// The value must be greater than or equal to <see cref="HeaderRowIndex"/>
        /// </summary>
        public int StartingRowIndex { get; set; }

        /// <summary>
        /// Defines the header names to identify column data, and, if <see cref="HeaderRowIndex"/> is used to auto identify header names, contains the identified header names after the import
        /// </summary>
        public IEnumerable<string> Headers { get; set; }
    }
}
