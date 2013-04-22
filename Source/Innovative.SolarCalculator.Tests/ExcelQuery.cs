using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;

namespace System.Data
{
	public class ExcelQuery
	{
		public class Cell
		{
			private string _column = string.Empty;
			private int _row = 0;
			private OleDbType _dataType = OleDbType.BSTR;

			public Cell()
			{
			}

			public Cell(string column, int row)
			{
				this.Column = column;
				this.Row = row;
			}

			public Cell(string column, int row, OleDbType dataType)
			{
				this.Column = column;
				this.Row = row;
				this.DataType = dataType;
			}

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

		private static object _lockObject = new object();
		private string _sheetName = string.Empty;
		private string _excelFilePath = string.Empty;
		private bool _headerHasFieldName = true;

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

		public ExcelQuery()
		{
		}

		public ExcelQuery(string excelFilePath)
		{
			this.ExcelFilePath = excelFilePath;
		}

		public ExcelQuery(string excelFilePath, string sheetName)
		{
			this.ExcelFilePath = excelFilePath;
			this.SheetName = sheetName;
		}

		public ExcelQuery(string excelFilePath, string sheetName, bool headerHasFieldName)
		{
			this.ExcelFilePath = excelFilePath;
			this.SheetName = sheetName;
			this.HeaderHasFieldNames = headerHasFieldName;
		}

		public DataSet ExecuteDataSetQuery(string sql)
		{
			DataSet returnValue = new DataSet();

			lock (_lockObject)
			{
				using (OleDbConnection connection = new OleDbConnection())
				{
					connection.ConnectionString = this.ConnectionString();

					using (OleDbDataAdapter adp = new OleDbDataAdapter(sql, connection))
					{
						adp.Fill(returnValue);
					}
				}
			}

			return returnValue;
		}

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

		public string ConnectionString()
		{
			string returnValue = string.Empty;

			if (Convert.ToBoolean(this.HeaderHasFieldNames))
			{
				returnValue = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\"", _excelFilePath);
			}
			else
			{
				returnValue = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=NO\"", _excelFilePath);
			}

			return returnValue;
		}

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

		public DataTable GetSheets()
		{
			DataTable returnValue = null;

			returnValue = this.GetSchema(OleDbSchemaGuid.Tables);

			return returnValue;
		}

		public bool UpdateCell(Cell cell, object value)
		{
			return this.UpdateCell(cell.Column, cell.Row, cell.DataType, value);
		}

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
