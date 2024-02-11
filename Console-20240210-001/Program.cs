using Newtonsoft.Json;
using System.Data;
using Formatting = Newtonsoft.Json.Formatting;
using Kunsheng.Utility;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Console_20240210_001;

namespace Console20240210001
{
    public class Program
    {
        private static bool _updating;
        public static void Main(string[] args)
        {
            Property property_length = new Property() { Name = "示例.L", Type = typeof(decimal), Unit = Unit.Meter, Expression = null };
            Property property_width = new Property() { Name = "示例.W", Type = typeof(decimal), Unit = Unit.Meter, Expression = null };
            Property property_height = new Property() { Name = "示例.H", Type = typeof(decimal), Unit = Unit.Meter, Expression = null };
            Property property_k = new Property() { Name = "示例.K", Type = typeof(decimal), Unit = Unit.Meter, Expression = null };
            Property property_capacity = new Property() { Name = "示例.C", Type = typeof(decimal), Unit = Unit.Centimeter, Expression = "示例.L*示例.W*示例.H*示例.K" };
            Property property_h = new Property() { Name = "示例.h", Type = typeof(decimal), Unit = Unit.Meter };
            Property property_h0 = new Property() { Name = "示例.h0", Type = typeof(decimal), Unit = Unit.Meter };
            Property property_volume = new Property() { Name = "示例.v", Type = typeof(decimal), Unit = Unit.Centimeter,Expression= "示例.L*示例.W*(示例.H-示例.h+示例.h0)*示例.K" };
            Property property_volume_average = new Property() { Name = "示例.va", Type = typeof(decimal), Unit = Unit.Centimeter, Expression = null };
            Property property_volume_delta = new Property() { Name = "示例.vd", Type = typeof(decimal), Unit = Unit.Centimeter, Expression = null };
            Property property_trend = new Property() { Name = "示例.Trend", Type = typeof(string), Unit = Unit.None, Expression = null };
            Property property_trend_MannKendall = new Property() { Name = "示例.TrendM", Type = typeof(string), Unit = Unit.None, Expression = null };
            Property property_warn = new Property() { Name = "示例.Warn", Type = typeof(string), Unit = Unit.None, Expression = null };


            List<Property> properties = new List<Property>();
            properties.Add(property_length);
            properties.Add(property_width);
            properties.Add(property_height);
            properties.Add(property_k);
            properties.Add(property_capacity);
            properties.Add(property_h);
            properties.Add(property_h0);
            properties.Add(property_volume);
            properties.Add(property_volume_average);
            properties.Add(property_volume_delta);
            properties.Add(property_trend);
            properties.Add(property_trend_MannKendall);
            properties.Add(property_warn);

            DataTable dataTable = new DataTable(); 

            IList<DataColumn> columnsImperial = new List<DataColumn>();

            foreach(Property property in properties)
            {
                DataColumn column = new DataColumn(property.Name, property.Type!, property.Expression!);

                dataTable.Columns.Add(column);

                DataColumn? column2 = null;
                switch (property.Unit)
                {
                    case Unit.Meter:
                        column2 = new DataColumn(string.Format("{0}'",property.Name), property.Type!, string.Format("{0}* 3.281",property.Name));
                        break;
                    case Unit.Centimeter:
                        column2 = new DataColumn(string.Format("{0}'",property.Name), property.Type!, string.Format("{0}* 10.764", property.Name));
                        break;
                }
                if(column2!=null) columnsImperial.Add(column2);
            }

            dataTable.Columns.AddRange(columnsImperial.ToArray());


            int nPoints = 1800;
            int monitorPoints = 120;

            Random random = new Random();

            Dictionary<string, decimal> globalVars=new Dictionary<string, decimal>();
            globalVars.Add("示例.体积总和", 0m);
            globalVars.Add("示例.变化阈值", 3m);
            globalVars.Add("示例.报警阈值", 8m);
            globalVars.Add("示例.L", 6m);
            globalVars.Add("示例.W", 3m);
            globalVars.Add("示例.H", 3m);
            globalVars.Add("示例.K", 1m);
            globalVars.Add("示例.h0", 0.1m);

            // 生成数据
            for (int i = 0; i < nPoints- monitorPoints+1; i++)
            {
                // 创建上升下降趋势
                decimal x =Convert.ToDecimal( (4 * Math.PI) * i / nPoints);
                decimal y =Convert.ToDecimal( Math.Sin(Convert.ToDouble(x)) * 100);

                // 添加噪音
                decimal noise = Convert.ToDecimal( random.NextDouble() * 10 - 5); // 生成-5到5之间的噪音

                decimal h = Math.Round(Convert.ToDecimal((y + noise) / 100 + 0.1m), 3);

                dataTable.Rows.Add(new object[] { globalVars["示例.L"], globalVars["示例.W"], globalVars["示例.H"], globalVars["示例.K"], 0, h, globalVars["示例.h0"], 0, 0, 0 });

                globalVars["示例.体积总和"] += ConvertUtil.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.v"]);

                if (dataTable.Rows.Count==1)
                {
                    for(int j=0;j< monitorPoints-1;j++)
                    {
                        dataTable.Rows.Add(dataTable.Rows[0].ItemArray);
                    }

                    globalVars["示例.体积总和"] = ConvertUtil.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.v"]) * monitorPoints;
                }

                int period = Convert.ToInt32(Math.Floor(Convert.ToDecimal(dataTable.Rows.Count / monitorPoints)));

                if (dataTable.Rows.Count==monitorPoints * period )
                {
                    decimal[] data = dataTable.AsEnumerable()
                                    .Skip((period - 1) * monitorPoints)
                                    .Take(monitorPoints)
                                    .Select(row => row.Field<decimal>("示例.v"))
                                    .ToArray();

                    int s = TrendUtil.MannKendallTest(data);

                    for (int j = (period-1)* monitorPoints; j < period*monitorPoints; j++)
                    {
                       dataTable.Rows[j]["示例.va"] = globalVars["示例.体积总和"] / monitorPoints;
                        if(s==0) dataTable.Rows[j]["示例.TrendM"] = "平稳";
                        if(s>0) dataTable.Rows[j]["示例.TrendM"] = "上升";
                        if (s < 0) dataTable.Rows[j]["示例.TrendM"] = "下降";
                    }

                    globalVars["示例.体积总和"] = 0;
                   
                }

                if(dataTable.Rows.Count > monitorPoints)
                {
                    dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"] = Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.v"]) - Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1 - monitorPoints]["示例.va"]);

                    if (Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"]) > globalVars["示例.变化阈值"])
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Trend"] = "上升";
                    }
                    else if (Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"]) <-globalVars["示例.变化阈值"])
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Trend"] = "下降";
                    }
                    else
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Trend"] = "平稳";
                    }

                    if (Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"]) > globalVars["示例.报警阈值"])
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Warn"] = "溢出";
                    }
                    else if (Convert.ToDecimal(dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"]) < -globalVars["示例.报警阈值"])
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Warn"] = "渗漏";
                    }
                    else
                    {
                        dataTable.Rows[dataTable.Rows.Count - 1]["示例.Warn"] = "正常";
                    }
                }
                else
                {
                    dataTable.Rows[dataTable.Rows.Count - 1]["示例.vd"] = 0m;
                }
            }

            string json = JsonUtil.ToJson(dataTable);

            Console.WriteLine(json);

            Console.WriteLine(dataTable.Rows.Count);

            DataTable dataTable2 = dataTable.Clone();

            dataTable2 = JsonUtil.ToDataTable(json, dataTable2);

            //// 增加120行
            //for(int i=0;i<monitorPoints; i++)
            //{
            //    dataTable2.Rows.Add(dataTable2.Rows[dataTable2.Rows.Count - 1].ItemArray);
            //}
            //
            //// 删除前面120行
            //if (dataTable2.Rows.Count > 120)
            //{
            //    for (int i = 0; i < 120; i++)
            //    {
            //        // 删除最前面的行
            //        dataTable2.Rows[0].Delete();
            //    }
            //    // 确认更改
            //    dataTable2.AcceptChanges();
            //}

            ExcelUtil.ExportDataTableToExcel(dataTable2, "MyWorkbook.xlsx", null);
        }
    }
}