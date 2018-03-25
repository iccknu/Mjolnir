using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Text;

namespace Databases.Connectors
{
    public class MsSQL : Connector
    {
        private static class Syntax
        {
            public const string CreateSchema = "CREATE SCHEMA $schema;";
            public const string CreateTable = "CREATE TABLE $tableName ($fields)";
            public const string CreateTableId = "CREATE TABLE $tableName (Id int IDENTITY(1,1),$fields)";

            public const string UpdateMainTable = "UPDATE $MainTable SET $ColumnId=(SELECT Id FROM $TableSourceName WHERE $ColumnSourceName=N'$ColumnSourceValue') WHERE $ColumnName=$ColumnNameValue;";

            public const string Insert = "INSERT INTO $tableName ($fields) VALUES ";
            public const string InsertValues = "($values)";

            public const string NVARCHAR = "NVARCHAR($length)";

            public const string ParamSymbol = "@";
        }

        private SqlConnection conn = null;
        private string connectionString = null;

        public override string IntType { get { return "int"; } }
        public override string FloatType { get { return "float"; } }
        public override string DoubleType { get { return "real"; } }
        public override int bulkLimit { get { return 999; } }

        public override string NVARCHARType(int length)
        {
            return "NVARCHAR(" + length.ToString() + ")";
        }

        public MsSQL(string connectionString)
        {
            this.connectionString = connectionString;
        }

        private SqlConnection SQLConnect()
        {
            if (this.conn == null || this.conn.State != System.Data.ConnectionState.Open)
            {
                SqlConnection conn = new SqlConnection(this.connectionString);

                try
                {
                    conn.Open();
                }
                catch
                {
                    conn = null;
                }

                this.conn = conn;
            }

            return this.conn;
        }

        public override void CreateSchema(string schema)
        {
            Execute(base.CreateByTemplate(Syntax.CreateSchema, new Dictionary<string, string>() { { "schema", schema } }));
        }

        public override void CreateTable(string tableName, Dictionary<string, string> fields, string schema = "CSV")
        {
            throw new NotImplementedException();
        }

        public override bool Execute(string query)
        {
            bool status = false;

            try
            {
                SqlCommand com = new SqlCommand();
                com.CommandText = query;
                com.Connection = SQLConnect();
                int result = com.ExecuteNonQuery();
                if (result!=-1)
                {
                    status = true;
                }
            }
            catch
            {
                status = false;
            }

            return status;
        }

        public override DataTable Select(string query)
        {
            SqlConnection con = SQLConnect();
            SqlCommand com = new SqlCommand(query, con);
            SqlDataAdapter adapter = new SqlDataAdapter(com);

            DataTable data = new DataTable();

            adapter.Fill(data);

            return data;
        }

        private string CreateParam(string name)
        {
            return Syntax.ParamSymbol + name;
        }

        public override void BulkInsert(string tableName, string[] fields, string[][] source, int[] filter = null)
        {
            SqlCommand com = new SqlCommand();

            int rowsCount = source[0].Length;
            int columnsCount = fields.Length;

            int ourBulkCount = 0;
            int ourBulkLimit = this.bulkLimit / source.Length;

            string[] ourFields = fields;

            if (filter != null)
            {
                ourBulkLimit = this.bulkLimit / filter.Length;
                columnsCount = filter.Length;

                List<string> tempFields = new List<string>();

                for (int i = 0; i < filter.Length; i++)
                {
                    tempFields.Add(fields[filter[i]]);
                }


                ourFields = tempFields.ToArray();
            }

            string bulkInsert = Insert(tableName, ourFields);

            for (int i = 0; i < rowsCount; i++)
            {
                string tempInsert = "";

                for (int u = 0; u < columnsCount; u++)
                {
                    string paramName = CreateParam(i.ToString() + "_" + u.ToString());

                    tempInsert += paramName + ",";
                    com.Parameters.AddWithValue(paramName, source[u][i]);
                }

                bulkInsert += "(" + tempInsert.Substring(0, tempInsert.Length - 1) + "),";
                tempInsert = "";

                if (ourBulkCount == ourBulkLimit)
                {

                    com.CommandText = bulkInsert.Substring(0, bulkInsert.Length - 1);
                    com.Connection = SQLConnect();
                    com.ExecuteNonQuery();

                    ourBulkCount = 0;

                    bulkInsert = Insert(tableName, ourFields);

                    com = new SqlCommand();
                }

                ourBulkCount++;
            }
        }

        public override void Update(string tableName, Dictionary<string, string> setFields)
        {
            throw new NotImplementedException();
        }

        public override string Insert(string tableName, string[] fields)
        {
            string query = "INSERT INTO " + tableName + " ";
            string insertFields = String.Join(',', fields);

            insertFields = "(" + insertFields + ") VALUES";

            query = query + insertFields;

            return query;
        }
    }
}
