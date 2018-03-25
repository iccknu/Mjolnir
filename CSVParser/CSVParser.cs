using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CSVParser
{
    public class CSVParser
    {

        private string path;
        private string symbol;

        private string[] fields;
        private string[][] tables;

        private string[,] tableCSV;

        private uint[,] __indexTable;//Unsigned, because no need negative

        private float maxPercent = 99F;

        private int rowsCount = 0;
        private int fieldsCount = 0;
        private int[] __index;
        private int[] __indexFull;
        /// <summary>
        /// Initialization
        /// </summary>
        /// <param name="path">Path to CSV file</param>
        /// <param name="symbol">Delimiter character. By default is ;</param>
        public CSVParser(string path, string symbol = ";")
        {
            this.path = path;
            this.symbol = symbol;
        }
        /// <summary>
        /// Get rows
        /// </summary>
        /// <returns></returns>
        public int GetRows()
        {
            return this.rowsCount;
        }
        /// <summary>
        /// Get index table
        /// </summary>
        /// <returns>uint array with indexes</returns>
        public uint[,] GetIndexTable()
        {
            return this.__indexTable;
        }

        /// <summary>
        /// Return string array with values based on row id
        /// </summary>
        /// <param name="rowId">RowId</param>
        /// <returns>String array</returns>
        public string[] GetRowValues(int rowId)
        {
            int columns = fields.Length;
            string[] row = new string[columns];

            for (int i = 0; i < columns; i++)
            {
                row[i] = tables[i][__indexTable[i, rowId]];
            }

            return row;
        }
        /// <summary>
        /// Return field value as a string
        /// </summary>
        /// <param name="columnId">Column id</param>
        /// <param name="rowId">Row id</param>
        /// <returns></returns>
        public string GetFieldValue(int columnId, int rowId)
        {
            return tables[columnId][__indexTable[columnId, rowId]];
        }
        /// <summary>
        /// Return string array with all fields
        /// </summary>
        /// <returns></returns>
        public string[] GetFields()
        {
            return this.fields;
        }

        /// <summary>
        /// Return string array with all tables
        /// </summary>
        /// <returns></returns>
        public string[][] GetTables()
        {
            return this.tables;
        }
        public string[][] GetTable()
        {
            string[][] temp = new string[fieldsCount][];

            foreach (int index in __index)
            {
                temp[index] = new string[rowsCount];
            }

            for (int i = 0; i < rowsCount; i++)
            {
                foreach (int index in __index)
                {
                    temp[index][i] = tables[index][__indexTable[index, i]];
                }
            }


            foreach (int index in __indexFull)
            {
                temp[index] = tables[index];
            }

            return temp;
        }

        public string GetFieldValueFiltered(string mainField, string searchField, string searchValue)
        {
            int searchFieldIndex = Array.IndexOf(fields, searchField);
            int mainFieldIndex = Array.IndexOf(fields, mainField);

            int findIndex = Array.IndexOf(tables[searchFieldIndex], searchValue);

            string findValue = tables[mainFieldIndex][__indexTable[mainFieldIndex, findIndex]];

            return findValue;
        }

        /// <summary>
        /// Return string array with values for fields
        /// </summary>
        /// <param name="fieldName">Field name</param>
        /// <returns></returns>
        public string[] GetFieldValues(string fieldName)
        {
            int index = Array.IndexOf(fields, fieldName);

            int length = this.tables[index].Length;

            if (length == rowsCount)
            {
                return tables[index];
            }
            else
            {
                string[] temp = new string[rowsCount];

                for (int i = 0; i < rowsCount; i++)
                {
                    temp[i] = tables[index][__indexTable[index, i]];
                }

                return temp;
            }
        }


        private void Dest()
        {
            this.fields = null;

            this.tableCSV = null;
        }
        /// <summary>
        /// Parser. For low memory consumption used indexes
        /// </summary>
        /// <returns></returns>
        public bool Parse()
        {

            if (File.Exists(this.path))
            {
                string[] sourceCSV = File.ReadAllLines(this.path);
                this.fields = sourceCSV[0].Split(this.symbol);

                int fieldsCount = this.fields.Length;
                int rowsCount = sourceCSV.Length - 1;//Because first row in csv - is a fields

                this.fieldsCount = fieldsCount;
                this.rowsCount = rowsCount;


                string[][] test = new string[this.fields.Length][];

                List<int> errorsIndex = new List<int>();

                for (int i = 1, virtualIndex = 0; virtualIndex < rowsCount; i++, virtualIndex++)
                {

                    string[] temp = sourceCSV[i].Split(this.symbol);

                    if (temp.Length != fieldsCount)//If invalid string
                    {
                        errorsIndex.Add(i);
                    }

                }

                if (errorsIndex.Count > 0)
                {
                    List<string> tempString = sourceCSV.ToList();

                    errorsIndex.Reverse(); //Need reverse list for correct indexes

                    //Delete erroneous string indexes
                    foreach (int index in errorsIndex)
                    {
                        tempString.RemoveAt(index);
                    }

                    sourceCSV = tempString.ToArray();
                }


                rowsCount = rowsCount - errorsIndex.Count;

                this.rowsCount = rowsCount;

                for (int y = 0; y < fieldsCount; y++)
                {
                    test[y] = new string[rowsCount];
                }

                int[] statistics = new int[fields.Length];

                int errors = 0;

                for (int i = 1, virtualIndex = 0; virtualIndex < rowsCount; i++, virtualIndex++)
                {

                    string[] temp = sourceCSV[i].Split(this.symbol);

                    if (temp.Length == fieldsCount)//If field is valid. Old version
                    {

                        for (int u = 0; u < temp.Length; u++)
                        {
                            if (u < fieldsCount && temp.Length <= fieldsCount) //temp.Length >= fieldsCount
                            {

                                if (temp[u] == "")
                                {
                                    temp[u] = "EMPTY";
                                }

                                test[u][virtualIndex] = temp[u];

                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                    else
                    {
                        errors++;
                    }

                }

                this.tables = new string[fieldsCount][];

                List<int> indexes = new List<int>();
                List<int> indexesFull = new List<int>();

                for (int i = 0; i < fieldsCount; i++)
                {
                    statistics[i] = test[i].Distinct().Count();

                    float percent = (statistics[i] * 1F) / (sourceCSV.Length * 1F) * 100F;

                    if (percent < maxPercent)
                    {
                        tables[i] = test[i].Distinct().ToArray();
                        indexes.Add(i);
                    }
                    else
                    {
                        indexesFull.Add(i);
                        tables[i] = test[i];
                    }
                }

                __index = indexes.ToArray();
                __indexFull = indexesFull.ToArray();
                __indexTable = new uint[fieldsCount, rowsCount];

                for (int virtualIndex = 0; virtualIndex < rowsCount; virtualIndex++)
                {
                    for (int u = 0; u < fieldsCount; u++)
                    {
                        if (tables[u].Length != rowsCount)
                        {
                            string source = test[u][virtualIndex];
                            int index = Array.IndexOf(tables[u], source);
                            __indexTable[u, virtualIndex] = (uint)index;
                        }
                        else
                        {
                            __indexTable[u, virtualIndex] = (uint)virtualIndex;
                        }
                    }
                }

            }
            else
            {
                //In future change to exception
                return false;
            }

            return true;
        }
    }
}
