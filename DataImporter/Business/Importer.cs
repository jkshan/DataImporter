using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using DataImporter.Business.Interface;
using DataImporter.Business.Validator;
using DataImporter.Models;

namespace DataImporter.Business
{
    public class Importer
    {
        IParser _parser;
        private string _targetTableName;
        private string _connectionString;
        private int _batchCount;

        public Importer(IParser Parser, string ConnectionString, string TargetTableName, int BatchCount)
        {
            _parser = Parser;
            _connectionString = ConnectionString;
            _targetTableName = TargetTableName;
            _batchCount = BatchCount;
        }

        /// <summary>
        /// Created the SQL Bulk Copy Object with Column Mappings
        /// </summary>
        /// <param name="sqlConnection">SQL Connection On which the Bulk Insert will happen</param>
        /// <returns>SqlBUlkCopy Object with Column Mappings</returns>
        private SqlBulkCopy CreateSqlBulkCopy(SqlConnection sqlConnection, string TargetTabelName)
        {
            // Create the bulk copy object
            var sqlBulkCopy = new SqlBulkCopy(sqlConnection)
            {
                DestinationTableName = TargetTabelName
            };

            // Setup the column mappings, anything ommitted is skipped
            sqlBulkCopy.ColumnMappings.Add("Account", "Account");
            sqlBulkCopy.ColumnMappings.Add("Description", "Description");
            sqlBulkCopy.ColumnMappings.Add("CurrencyCode", "CurrencyCode");
            sqlBulkCopy.ColumnMappings.Add("Amount", "Amount");
            return sqlBulkCopy;
        }

        /// <summary>
        /// Create the Datatable Based on the SQL Table
        /// </summary>
        /// <param name="TargetTabelName">Name of the table for DataTable</param>
        /// <returns></returns>
        private DataTable CreateDataTable(string TargetTabelName)
        {
            var dataTable = new DataTable(TargetTabelName);

            dataTable.Columns.Add("Account");
            dataTable.Columns.Add("Description");
            dataTable.Columns.Add("CurrencyCode");
            dataTable.Columns.Add("Amount", typeof(decimal));

            return dataTable;
        }

        /// <summary>
        /// Pushes the Data to SQL
        /// </summary>        
        /// <param name="sqlBulkCopy">Sql Bulk Copy object to which the Datas to be pushed</param>
        /// <param name="dataTable">Datatable containing the Data to be pushed using SQL Bulk Copy</param>
        /// <returns></returns>
        private async Task InsertDataTable(SqlBulkCopy sqlBulkCopy, DataTable dataTable)
        {
            await sqlBulkCopy.WriteToServerAsync(dataTable);

            dataTable.Rows.Clear();
        }

        /// <summary>
        /// Process the Records in batch Manner
        /// </summary>
        /// <param name="SqlConnection">Target Sql Database's Connection String</param>
        /// <param name="TargetTableName">Table Name in the Target Database</param>
        /// <returns></returns>
        public async Task<List<Record>> Process()
        {
            var ErrorList = new List<Record>();
            int createdCount = 0, failedCount = 0;
            var dataTable = CreateDataTable(_targetTableName);
            using(SqlConnection sqlConnection = new SqlConnection(_connectionString))
            {
                sqlConnection.Open();

                var sqlBulkCopy = CreateSqlBulkCopy(sqlConnection, _targetTableName);

                foreach(var item in _parser.Readline())
                {
                    Record row = new Record();
                    try
                    {
                        if(item.Count() != 4)
                        {
                            throw new ApplicationException();
                        }
                        row = new Record()
                        {
                            Account = item[0],
                            Description = item[1],
                            CurrencyCode = item[2],
                            Amount = item[3]
                        };
                    }
                    catch(Exception)
                    {
                        row = new Record()
                        {
                            ErrorMessages = string.Format("Error Processing data @ Row: {0}", createdCount + failedCount)
                        };
                        ErrorList.Add(row);
                        failedCount++;
                        continue;
                    }
                    if(DataValidator.ValidateData(row))
                    {
                        dataTable.Rows.Add(item.ToArray());
                    }
                    else
                    {
                        row.ErrorMessages += ", Row No: " + (createdCount + failedCount);
                        ErrorList.Add(row);
                        failedCount++;
                        continue;
                    }

                    createdCount++;

                    if(createdCount % _batchCount == 0)
                    {
                        await InsertDataTable(sqlBulkCopy, dataTable);
                    }
                }

                if(dataTable.Rows.Count > 0)
                {
                    await InsertDataTable(sqlBulkCopy, dataTable);
                }
            }
            return ErrorList;
        }
    }
}