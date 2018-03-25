using System;
using System.Collections.Generic;
using System.Data;

namespace Databases
{
    public abstract class Connector
    {
        public abstract string IntType { get; }
        public abstract string FloatType { get; }
        public abstract string DoubleType { get; }
        public abstract int bulkLimit { get; }
        public abstract string NVARCHARType(int length);
        public abstract void CreateSchema(string schema);
        public abstract void CreateTable(string tableName, Dictionary<string, string> fields, string schema = "CSV");
        public abstract bool Execute(string query);
        public abstract string Insert(string tableName, string[] fields);
        public abstract void Update(string tableName, Dictionary<string, string> setFields);
        public abstract void BulkInsert(string tableName, string[] fields, string[][] source, int[] filter = null);
        public abstract DataTable Select(string query);
        public string CreateByTemplate(string template, Dictionary<string, string> vars)
        {
            string result = template;

            foreach (KeyValuePair<string, string> key in vars)
            {
                result = result.Replace("$" + key.Key, key.Value);
            }

            return result;
        }

    }
}
