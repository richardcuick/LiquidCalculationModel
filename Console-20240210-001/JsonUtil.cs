using Newtonsoft.Json;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Kunsheng.Utility
{
    public static class JsonUtil
    {
        public static string ToJson(DataRow row)
        {
            Dictionary<string, object> rowDictionary = row.Table.Columns
            .Cast<DataColumn>()
            .ToDictionary(col => col.ColumnName, col => row[col]);

            string json = JsonConvert.SerializeObject(rowDictionary, Formatting.Indented);

            return json;
        }

        public static string ToJson(DataTable table)
        {
            IList<Dictionary<string, object>> tableList = new List<Dictionary<string, object>>();

            foreach (DataRow row in table.Rows)
            {
                Dictionary<string, object> rowDictionary = table.Columns
               .Cast<DataColumn>()
               .ToDictionary(col => col.ColumnName, col => row[col]);

                tableList.Add(rowDictionary);
            }

            string json = JsonConvert.SerializeObject(tableList, Formatting.Indented);

            return json;
        }

        public static DataRow ToDataRow(string json, DataTable table)
        {
            var dictionary = JsonConvert.DeserializeObject<Dictionary<string, object>>(json);

            DataRow row = table.NewRow();
            foreach (var keyValuePair in dictionary!)
            {
                row[keyValuePair.Key] = Convert.ChangeType(keyValuePair.Value, table.Columns[keyValuePair.Key]!.DataType);
            }
            return row;
        }

        public static DataTable ToDataTable(string json, DataTable table)
        {
            var dictList = JsonConvert.DeserializeObject<IList<Dictionary<string, object>>>(json);

            foreach(Dictionary<string,object> dict in dictList)
            {
                DataRow row = table.NewRow();
                foreach (var keyValuePair in dict!)
                {
                    row[keyValuePair.Key] = Convert.ChangeType(keyValuePair.Value, table.Columns[keyValuePair.Key]!.DataType);
                }
                table.Rows.Add(row);
            }

            return table;
        }
    }
}
