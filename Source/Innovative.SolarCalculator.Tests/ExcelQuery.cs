// ***
// *** Copyright (C) 2013, Daniel M. Porrey.  All rights reserved.
// *** Written By Daniel M. Porrey
// ***
// *** This software is provided "AS IS," without a warranty of any kind. ALL EXPRESS OR IMPLIED CONDITIONS, REPRESENTATIONS AND WARRANTIES, 
// *** INCLUDING ANY IMPLIED WARRANTY OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR NON-INFRINGEMENT, ARE HEREBY EXCLUDED. DANIEL M PORREY 
// *** AND ITS LICENSORS SHALL NOT BE LIABLE FOR ANY DAMAGES SUFFERED BY LICENSEE AS A RESULT OF USING, MODIFYING OR DISTRIBUTING THIS SOFTWARE 
// *** OR ITS DERIVATIVES. IN NO EVENT WILL DANIEL M PORREY OR ITS LICENSORS BE LIABLE FOR ANY LOST REVENUE, PROFIT OR DATA, OR FOR DIRECT, INDIRECT, 
// *** SPECIAL, CONSEQUENTIAL, INCIDENTAL OR PUNITIVE DAMAGES, HOWEVER CAUSED AND REGARDLESS OF THE THEORY OF LIABILITY, ARISING OUT OF THE USE OF 
// *** OR INABILITY TO USE THIS SOFTWARE, EVEN IF DANIEL M PORREY HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES. 
// ***
// *** Licensed under Microsoft Reciprocal License (Ms-RL)
// *** This license governs use of the accompanying software. If you use the software, you accept this license. If you do not accept the license, 
// *** do not use the software. Full license details can be found at https://solarcalculator.codeplex.com/license.
// ***
using System.Data.OleDb;
using System.Collections.Generic;
using System.Linq;

namespace System.Data
{
    /// <summary>
    /// This is a simple utility class for performing some basic operations
    /// against an Excel spreadsheet.
    /// </summary>
    public class ExcelQuery
    {
        /// <summary>
        /// Represents a cell within a worksheet.
        /// </summary>
        public class Cell
        {
            private string _column = string.Empty;
            private int _row = 0;
            private OleDbType _dataType = OleDbType.BSTR;

            /// <summary>
            /// Creates a default instance of a Cell
            /// </summary>
            public Cell()
            {
            }

            /// <summary>
            /// Creates an instance of a Cell with the specified column and row.
            /// </summary>
            public Cell(string column, int row)
            {
                this.Column = column;
                this.Row = row;
            }

            /// <summary>
            /// Creates an instance of a Cell with the specified column, row and datatype.
            /// </summary>
            public Cell(string column, int row, OleDbType dataType)
            {
                this.Column = column;
                this.Row = row;
                this.DataType = dataType;
            }

            /// <summary>
            /// Gets/sets the name of the column in this Cell instance.
            /// </summary>
            public string Column
            {
                get
                {
                    return _column;
                }
                set
                {
                    _column = value;
                }
            }

            /// <summary>
            /// Gets/sets the row index for this Cell instance.
            /// </summary>
            public int Row
            {
                get
                {
                    return _row;
                }
                set
                {
                    _row = value;
                }
            }

            /// <summary>
            /// Gets/sets the Data Type for this Cell instance.
            /// </summary>
            public OleDbType DataType
            {
                get
                {
                    return _dataType;
                }
                set
                {
                    _dataType = value;
                }
            }
        }

        public static class DataProviders
        {
            public static string Jet12 = "Microsoft.ACE.OLEDB.12.0";
            public static string Jet4 = "Microsoft.Jet.OLEDB.4.0";
        }

        private static object _lockObject = new object();
        private string _sheetName = string.Empty;
        private string _excelFilePath = string.Empty;
        private bool _headerHasFieldName = true;
        private string _dataProvider = DataProviders.Jet12;

        /// <summary>
        /// Gets/sets whether or not the first row in the spreadsheet (or range) contains
        /// names that will represent the column names of the return data.
        /// </summary>
        public bool HeaderHasFieldNames
        {
            get
            {
                return _headerHasFieldName;
            }
            set
            {
                _headerHasFieldName = value;
            }
        }

        /// <summary>
        /// Creates a default instance of the ExcelQuery object.
        /// </summary>
        public ExcelQuery()
        {
        }

        /// <summary>
        /// Creates an instance of the ExcelQuery object using the specified file path.
        /// </summary>
        /// <param name="excelFilePath">The full path to a valid Excel spreadsheet./</param>
        public ExcelQuery(string excelFilePath)
        {
            this.ExcelFilePath = excelFilePath;
        }

        /// <summary>
        /// Creates an instance of the ExcelQuery object using the specified file path and sheet name.
        /// </summary>
        /// <param name="excelFilePath">The full path to a valid Excel spreadsheet.</param>
        /// <param name="sheetName">The name of the sheet containing the data.</param>
        public ExcelQuery(string excelFilePath, string sheetName)
        {
            this.ExcelFilePath = excelFilePath;
            this.SheetName = sheetName;
        }

        /// <summary>
        /// Creates an instance of the ExcelQuery object using the specified file path and sheet name
        /// specifying whether or not the first row has column names.
        /// </summary>
        /// <param name="excelFilePath">The full path to a valid Excel spreadsheet.</param>
        /// <param name="sheetName">The name of the sheet containing the data.</param>
        /// <param name="headerHasFieldName">Specifies whether or not the first row in the spreadsheet (or range) contains
        /// names that will represent the column names of the return data.</param>
        public ExcelQuery(string excelFilePath, string sheetName, bool headerHasFieldName)
        {
            this.ExcelFilePath = excelFilePath;
            this.SheetName = sheetName;
            this.HeaderHasFieldNames = headerHasFieldName;
        }

        /// <summary>
        /// Executes a query that will return a DataSet.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>
        /// <param name="tableName">The name to give the table in the return DataSet.</param>
        /// <returns>A new DataSet object containing the results of the query.</returns>
        public DataSet ExecuteDataSet(string sql, string tableName)
        {
            DataSet returnValue = new DataSet();

            lock (_lockObject)
            {
                using (OleDbConnection connection = new OleDbConnection())
                {
                    connection.ConnectionString = this.ConnectionString();

                    using (OleDbDataAdapter adp = new OleDbDataAdapter(sql, connection))
                    {
                        if (adp.Fill(returnValue) > 0)
                        {
                            returnValue.Tables[0].TableName = tableName;
                        }
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Executes a INSERT, UPDATE or DELETE query against the spreadsheet.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>
        /// <returns>The count of affected rows.</returns>
        public int ExecuteNonQuery(string sql)
        {
            int returnValue = 0;

            lock (_lockObject)
            {
                OleDbConnection connection = null;

                try
                {
                    using (connection = new OleDbConnection())
                    {
                        connection.ConnectionString = this.ConnectionString();
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        connection.Open();
                        returnValue = cmd.ExecuteNonQuery();
                    }
                }
                finally
                {
                    if (connection != null && connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Executes a parameterized INSERT, UPDATE or DELETE query against the spreadsheet.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>		
        /// <param name="parameters">An array of parameter values that matches the number of parameters in the SQL statement.</param>
        /// <returns>The count of affected rows.</returns>
        public int ExecuteNonQuery(string sql, OleDbParameter[] parameters)
        {
            int returnValue = 0;

            lock (_lockObject)
            {
                OleDbConnection connection = null;

                try
                {
                    using (connection = new OleDbConnection())
                    {
                        connection.ConnectionString = this.ConnectionString();
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        cmd.Parameters.AddRange(parameters);
                        connection.Open();
                        returnValue = cmd.ExecuteNonQuery();
                    }
                }
                finally
                {
                    if (connection != null && connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Executes a query that returns a single value.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>
        /// <returns>Returns an object containing the result of the SQL query.</returns>
        public object ExecuteScalar(string sql)
        {
            object returnValue = 0;

            lock (_lockObject)
            {
                OleDbConnection connection = null;

                try
                {
                    using (connection = new OleDbConnection())
                    {
                        connection.ConnectionString = this.ConnectionString();
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        connection.Open();
                        returnValue = cmd.ExecuteScalar();
                    }
                }
                finally
                {
                    if (connection != null && connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Executes a parameterized query that returns a single value.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>
        /// <param name="parameters">An array of parameter values that matches the number of parameters in the SQL statement.</param>
        /// <returns>Returns an object containing the result of the SQL query.</returns>
        public object ExecuteScalar(string sql, OleDbParameter[] parameters)
        {
            object returnValue = 0;

            lock (_lockObject)
            {
                OleDbConnection connection = null;

                try
                {
                    using (connection = new OleDbConnection())
                    {
                        connection.ConnectionString = this.ConnectionString();
                        OleDbCommand cmd = new OleDbCommand(sql, connection);
                        cmd.Parameters.AddRange(parameters);
                        connection.Open();
                        returnValue = cmd.ExecuteScalar();
                    }
                }
                finally
                {
                    if (connection != null && connection.State == ConnectionState.Open)
                    {
                        connection.Close();
                    }
                }
            }

            return returnValue;
        }

        /// <summary>
        /// Executes a SQL statement and opens a reader on the spreadsheet.
        /// </summary>
        /// <param name="sql">The SQL query to execute.</param>
        /// <returns>A reference to the reader that has been opened.</returns>
        public OleDbDataReader ExecuteReader(string sql)
        {
            OleDbDataReader returnValue = null;

            lock (_lockObject)
            {
                OleDbConnection connection = new OleDbConnection();
                connection.ConnectionString = this.ConnectionString();
                OleDbCommand cmd = new OleDbCommand(sql, connection);
                connection.Open();
                returnValue = cmd.ExecuteReader();
            }

            return returnValue;
        }

        /// <summary>
        /// Gets/sets the full path of the Excel spreadsheet.
        /// </summary>
        public string ExcelFilePath
        {
            get
            {
                return _excelFilePath;
            }
            set
            {
                _excelFilePath = value;
            }
        }

        /// <summary>
        /// Get/set the name of the sheet to retrieve data from.
        /// </summary>
        public string SheetName
        {
            get
            {
                return _sheetName;
            }
            set
            {
                _sheetName = value;

                if (!_sheetName.Substring(_sheetName.Length - 1, 1).Equals("$"))
                {
                    _sheetName = string.Format(@"{0}$", _sheetName);
                }
            }
        }

        /// <summary>
        /// Gets the connection string used to perform queries against the spreadsheet.
        /// </summary>
        /// <returns></returns>
        public string ConnectionString()
        {
            string returnValue = string.Empty;

            if (Convert.ToBoolean(this.HeaderHasFieldNames))
            {
                returnValue = string.Format("Provider={0};Data Source={1};Extended Properties=\"Excel 12.0;HDR=YES\"", this.DataProvider, this.ExcelFilePath);
            }
            else
            {
                returnValue = string.Format("Provider={0};Data Source={1};Extended Properties=\"Excel 12.0;HDR=NO\"", this.DataProvider, this.ExcelFilePath);
            }

            return returnValue;
        }

        /// <summary>
        /// Gets/sets the Access Database engine that should be used to read the Excel file.
        /// </summary>
        public string DataProvider
        {
            get { return _dataProvider; }
            set { _dataProvider = value; }
        }

        /// <summary>
        /// Returns a DataTable object describing the schema data requested.by Schema 
        /// Type. Use the OleDbSchemaGuid values to specify the type of data to 
        /// retrieve. For example, OleDbSchemaGuid.Tables can be used to get a 
        /// list of tables from the spreadsheet.
        /// </summary>
        /// <param name="schemaTypeGuid">A value from System.Data.OleDb.OleDbSchemaGuid.</param>
        /// <returns>A DataTable object containing the schema description requested.</returns>
        public DataTable GetSchema(Guid schemaTypeGuid)
        {
            DataTable returnValue = null;

            using (OleDbConnection connection = new OleDbConnection(this.ConnectionString()))
            {
                connection.Open();
                returnValue = connection.GetOleDbSchemaTable(schemaTypeGuid, new object[] { null, null, null, "TABLE" });
                connection.Close();
            }

            return returnValue;
        }

        /// <summary>
        /// Gets a list of sheets contained within the spreadsheet.
        /// </summary>
        /// <returns>Gets a list of sheet names contained in the spreadsheet.</returns>
        public IEnumerable<string> GetSheets()
        {
            List<string> returnValue = new List<string>();

            DataTable schema = this.GetSchema(OleDbSchemaGuid.Tables);

            returnValue = (from tbl in schema.AsEnumerable()
                           where tbl.Field<string>("TABLE_TYPE") == "TABLE"
                           select tbl.Field<string>("TABLE_NAME")).ToList();

            return returnValue;
        }

        /// <summary>
        /// Updates the contents of a specific cell in the spreadsheet.
        /// </summary>
        /// <param name="cell">The Cell to update.</param>
        /// <param name="value">The value to place in the cell.</param>
        /// <returns>True if successful, False otherwise.</returns>
        public bool UpdateCell(Cell cell, object value)
        {
            return this.UpdateCell(cell.Column, cell.Row, cell.DataType, value);
        }

        /// <summary>
        /// Updates the contents of a specific cell in the spreadsheet.
        /// </summary>
        /// <param name="column">The name of the column to update.</param>
        /// <param name="row">The row to update.</param>
        /// <param name="dataType">The type of data to be placed into the cell.</param>
        /// <param name="value">The value to be placed into the cell.</param>
        /// <returns>True if successful, False otherwise.</returns>
        public bool UpdateCell(string column, int row, OleDbType dataType, object value)
        {
            bool returnValue = false;

            string sql = string.Format("UPDATE [{0}{1}{2}:{1}{2}] SET F1 = ?", this.SheetName, column, row.ToString());
            OleDbParameter[] parameters = new OleDbParameter[1];
            parameters[0] = new OleDbParameter("F1", dataType) { Value = value };
            returnValue = (this.ExecuteNonQuery(sql, parameters) > 0);

            return returnValue;
        }
    }
}
