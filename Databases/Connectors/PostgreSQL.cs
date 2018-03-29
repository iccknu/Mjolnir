using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Databases.Connectors
{
    public class PostgreSQL:Connector
    {
        public override string IntType { get { return "int"; } }
        public override string FloatType { get { return "float"; } }
        public override string DoubleType { get { return "real"; } }
        public override int bulkLimit { get { return 999; } }

        public override void BulkInsert(string tableName, string[] fields, string[][] source, int[] filter = null)
        {
            throw new NotImplementedException();
        }

        public override void CreateSchema(string schema)
        {
            throw new NotImplementedException();
        }

        public override void CreateTable(string tableName, Dictionary<string, string> fields, string schema = "CSV")
        {
            throw new NotImplementedException();
        }

        public override bool Execute(string query)
        {
            throw new NotImplementedException();
        }

        public override string Insert(string tableName, string[] fields)
        {
            throw new NotImplementedException();
        }

        public override string NVARCHARType(int length)
        {
            throw new NotImplementedException();
        }

        public override DataTable Select(string query)
        {
            throw new NotImplementedException();
        }

        public override void Update(string tableName, Dictionary<string, string> setFields)
        {
            throw new NotImplementedException();
        }
    }
}
