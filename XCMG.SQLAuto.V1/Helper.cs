using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using RekTec.Crm.Common.Helper;
using System;
using System.Data;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using static NPOI.HSSF.Util.HSSFColor;
using static System.Net.Mime.MediaTypeNames;

namespace XCMG.SQLAuto.V1
{
    public static class Helper
    {
        static string sheetName = string.Empty;

        public static void ImportExcel(string filePath)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("新系统字段标签名");
            dt.Columns.Add("新系统字段名");
            dt.Columns.Add("是否必填");  
            dt.Columns.Add("字段类型");
            dt.Columns.Add("新系统关联到");
            dt.Columns.Add("字段值选项列表");
            dt.Columns.Add("老系统表名");
            dt.Columns.Add("老系统字段名");
            dt.Columns.Add("老系统字段类型");
            dt.Columns.Add("老系统关联到");
            dt.Columns.Add("匹配逻辑补充说明");
            dt.Columns.Add("备注");
            dt.Columns.Add("新系统关联到字段");
            dt.Columns.Add("老系统关联到字段");
            dt.Columns.Add("新系统表名");
            dt.Columns.Add("数据库地址");
            dt.Columns.Add("销售组织");
            dt.Columns.Add("映射关系");
            dt.Columns.Add("销售组织管理员");
            dt.Columns.Add("老系统表名简化");

            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fs); // HSSFWorkbook 用于 .xls
                ISheet sheet = workbook.GetSheetAt(0);
                int sheetIndex = workbook.GetSheetIndex(sheet);
                sheetName = workbook.GetSheetName(sheetIndex);   // 新系统表名

                for (int i = 1; i <= sheet.LastRowNum; i++) // 跳过表头
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; // 跳过空行
                    DataRow dataRow = dt.NewRow();
                    dataRow["新系统字段标签名"] = row.GetCell(0)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["新系统字段名"] = row.GetCell(1)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["是否必填"] = row.GetCell(2)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["字段类型"] = row.GetCell(3)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["新系统关联到"] = row.GetCell(4)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["字段值选项列表"] = row.GetCell(5)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统表名"] = sheet.GetRow(1).GetCell(6)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统字段名"] = row.GetCell(7)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统字段类型"] = row.GetCell(8)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统关联到"] = row.GetCell(9)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["匹配逻辑补充说明"] = row.GetCell(10)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["备注"] = row.GetCell(11)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["新系统关联到字段"] = GetRelationshipName(row.GetCell(10)?.ToString()?.Trim() ?? string.Empty, 1);
                    dataRow["老系统关联到字段"] = GetRelationshipName(row.GetCell(10)?.ToString()?.Trim() ?? string.Empty, 0);
                    dataRow["新系统表名"] = sheetName;   // sheet.GetRow(1).GetCell(14)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["数据库地址"] = sheet.GetRow(1).GetCell(15)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["销售组织"] = sheet.GetRow(1).GetCell(16)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["映射关系"] = row.GetCell(10)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["销售组织管理员"] = sheet.GetRow(1).GetCell(18)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统表名简化"] = GetTableName(sheet.GetRow(1).GetCell(6)?.ToString()?.Trim() ?? string.Empty);

                    dt.Rows.Add(dataRow);
                }
                
                BuildSQLFromExcel(dt);
            }
        }

        public static void BuildSQLFromExcel(DataTable baseDt)
        {
            try
            {
                DataRow rowFirst = baseDt.Rows[0];
                string dbName = Cast.ConToString(rowFirst["数据库地址"]);
                string orgName = Cast.ConToString(rowFirst["销售组织"]);
                string adminName = Cast.ConToString(rowFirst["销售组织管理员"]);

                StringBuilder builderBody = new StringBuilder();
                StringBuilder endBuilder = new StringBuilder();
                string main = Cast.ConToString(rowFirst["老系统表名简化"]);
                endBuilder.Append(@$"       {main}.ownerid,
       {main}.ModifiedBy,
       {main}.ModifiedOn,
       {main}.CreatedBy,
       {main}.CreatedOn,
       {main}.statecode
    FROM {dbName}.{Cast.ConToString(rowFirst["老系统表名"])}Base AS {GetTableName(Cast.ConToString(rowFirst["老系统表名"]))}
    LEFT JOIN {dbName}.Systemuser AS owner ON owner.systemuserid = {GetTableName(Cast.ConToString(rowFirst["老系统表名"]))}.ownerid
");
                LookupEntityModels models = new LookupEntityModels();
                StringBuilder bodyBuilder_New = new StringBuilder();
                StringBuilder endBuilder_New = new StringBuilder();
                endBuilder_New.Append(@$"FROM {orgName}_{Cast.ConToString(rowFirst["老系统表名"])} main
    INNER JOIN businessunit AS businessunit ON businessunit.name = '{GetOrgName(orgName)}' AND businessunit.isdisabled = 0
    LEFT JOIN Systemuser AS owner ON owner.address1_telephone1 = main.ownerid_address1_telephone1 AND owner.isdisabled = 0 AND owner.new_organization_id = businessunit.businessunitid
");

                StringBuilder updataBuilder = new StringBuilder();
                StringBuilder insertBuilder = new StringBuilder();
                StringBuilder insertBuilder_New = new StringBuilder();

                foreach (DataRow row in baseDt.Rows)
                {
                    switch (Cast.ConToString(row["字段类型"]).ToLower())
                    {
                        case "string":
                            builderBody.Append(GetFieldIsStringOrInter(row));
                            break;

                        case "memo":
                            builderBody.Append(GetFieldIsStringOrInter(row));
                            break;

                        case "integer":
                            builderBody.Append(GetFieldIsStringOrInter(row));
                            break;

                        case "double":
                            builderBody.Append(GetFieldIsDoubleOrDecimal(row));
                            break;

                        case "decimal":
                            builderBody.Append(GetFieldIsDoubleOrDecimal(row));
                            break;
                        
                        case "boolean":
                            builderBody.Append(GetFieldIsBoolean(row));
                            break;

                        case "datetime":
                            builderBody.Append(GetFieldIsDatetime(row));
                            break;

                        case "date":
                            builderBody.Append(GetFieldIsDatetime(row));
                            break;

                        case "picklist":
                            builderBody.Append(GetFieldIsPicklist(row));
                            break;

                        case "lookup":
                            GetFieldIsLookup(row, dbName, builderBody, endBuilder, bodyBuilder_New, endBuilder_New, models);
                            break;
                        default:
                            builderBody.Append($"--{Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n");
                            break;
                    }
                    
                    updataBuilder.Append($"   {Cast.ConToString(row["新系统字段名"])} = t2.{Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n");
                    insertBuilder.Append($"   {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n");
                    insertBuilder_New.Append($"   t2.{Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n");
                }

                string SQL = $@"
SELECT * INTO {orgName}_{Cast.ConToString(rowFirst["老系统表名"])} FROM(
    SELECT
       {GetTableName(Cast.ConToString(rowFirst["老系统表名"]))}.{Cast.ConToString(rowFirst["老系统表名"])}id AS new_oldid,
       owner.address1_telephone1 AS ownerid_address1_telephone1,
{builderBody.ToString()}{endBuilder.ToString()}    WHERE {GetTableName(Cast.ConToString(rowFirst["老系统表名"]))}.statecode = 0
)t;";

                string SQL_new = $@"
MERGE INTO {Cast.ConToString(rowFirst["新系统表名"])}Base t1
USING(
    SELECT
{bodyBuilder_New.ToString()}       main.*,
       owner.systemuserid AS new_systemuser_id
    {endBuilder_New.ToString()}
) t2
ON(t1.new_oldid = t2.new_oldid)
WHEN MATCHED THEN
UPDATE SET
{updataBuilder.ToString()}
   statecode = t2.statecode,
   CreatedOn = t2.CreatedOn,
   ModifiedOn = t2.ModifiedOn,
   ModifiedBy = isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   CreatedBy = isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   OwnerIdType = 8,
   ownerid = isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   OwningBusinessUnit = isnull((SELECT businessunitid FROM SystemUser WHERE systemuserid = t2.new_systemuser_id),(SELECT businessunitid FROM SystemUser WHERE fullname = '{adminName}'))
WHEN NOT MATCHED THEN
INSERT
(
{insertBuilder.ToString()}
   {Cast.ConToString(rowFirst["新系统表名"])}id,
   new_oldid,
   statecode,
   CreatedOn,
   ModifiedOn,
   ModifiedBy,
   CreatedBy,
   OwnerIdType,
   ownerid,
   OwningBusinessUnit
)
VALUES
(
{insertBuilder_New.ToString()}
   newid(),
   t2.new_oldid,
   t2.statecode,
   t2.CreatedOn,
   t2.ModifiedOn,
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   8,
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   isnull((SELECT businessunitid FROM SystemUser WHERE systemuserid = t2.new_systemuser_id),(SELECT businessunitid FROM SystemUser WHERE fullname = '{adminName}'))
);";

                //string outputPath = "C:\\Mac\\Home\\Desktop\\jimmyli\\Import\\Output\\SQL_0605.txt";
                Console.WriteLine("请在下方输入文件名：");
                var outputPath = "\\\\Mac\\Home\\Downloads\\{0}.txt";

                var outputPathName = Console.ReadLine();
                outputPath = string.Format(outputPath, outputPathName);
                SaveToTxt(SQL + "\n" + SQL_new, outputPath);
            }
            catch (Exception ex)
            {
                throw new Exception("生成sql异常：" + ex.Message);
            }
        }

        private static string GetFieldIsStringOrInter(DataRow row)
        {
            return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n";
        }

        private static string GetFieldIsDoubleOrDecimal(DataRow row)
        {
            if (Cast.ConToString(row["新系统字段名"]).Contains('万'))
            {
                return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])}/10000.00 AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}(转换成万元)\n";
            }
            else
            {
                return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n";
            }
        }

        private static string GetFieldIsBoolean(DataRow row)
        {
            return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n";
        }

        private static string GetFieldIsDatetime(DataRow row)
        {
            return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n";
        }

        private static string GetFieldIsPicklist(DataRow row)
        {
            string mapping = Cast.ConToString(row["映射关系"]);
            if (string.IsNullOrWhiteSpace(mapping))
            {
                return $"       {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n";
            }

            var mappingList = mapping.Replace("；", ";").Trim().Split(';');
            if(mappingList.Length == 0)
            {
                return string.Empty;
            }

            StringBuilder builder = new StringBuilder();
            int index = 0;
            foreach (string mappingTwo in mappingList)
            {
                if (!string.IsNullOrWhiteSpace(mappingTwo))
                {
                    var mappingArray = mappingTwo.Split("=");
                    if (mappingArray.Length == 2)
                    {
                        if (!mappingArray[0].Contains('/'))
                        {
                            if(index == 0)
                            {
                                builder.Append($"       CASE WHEN {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} = {mappingArray[0]} THEN {mappingArray[1]}\n");
                            }
                            else
                            {
                                builder.Append($"           WHEN {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} = {mappingArray[0]} THEN {mappingArray[1]}\n");
                            }
                        }
                        else
                        {
                            var picklistArr = mappingArray[0].Split('/');
                            if (index == 0)
                            {
                                builder.Append($"       CASE WHEN {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} IN ({string.Join(',', picklistArr)}) THEN {mappingArray[1]}\n");
                            }
                            else
                            {
                                builder.Append($"           WHEN {Cast.ConToString(row["老系统表名简化"])}.{Cast.ConToString(row["老系统字段名"])} IN ({string.Join(',', picklistArr)}) THEN {mappingArray[1]}\n");
                            }
                        }
                    }
                }
                index++;
            }
            builder.Append($"           ELSE NULL END AS {Cast.ConToString(row["新系统字段名"])},  --{Cast.ConToString(row["新系统字段标签名"])}\n");
            return builder.ToString();
        }

        private static void GetFieldIsLookup(DataRow row, string dbName, StringBuilder bodyBuilder_old, StringBuilder endBuilder_old, StringBuilder bodyBuilder_new, StringBuilder endBuilder_new, LookupEntityModels models)
        {
            int oldcount = 0;
            string oldTableName = Cast.ConToString(row["老系统关联到"]) + "Base";
            if (!string.IsNullOrWhiteSpace(endBuilder_old.ToString()))
            {
                oldcount = Regex.Matches(endBuilder_old.ToString(), Regex.Escape(oldTableName)).Count;
            }
            string oldTableNameJX = GetTableName(Cast.ConToString(row["老系统关联到"]));
            if (oldcount > 0)
            {
                oldTableNameJX = oldTableNameJX + oldcount.ToString();
            }
            bodyBuilder_old.Append($"       {oldTableNameJX}.{Cast.ConToString(row["老系统关联到字段"])} AS {Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
            endBuilder_old.Append($"    LEFT JOIN {dbName}.{Cast.ConToString(row["老系统关联到"])}Base AS {oldTableNameJX} ON {oldTableNameJX}.{Cast.ConToString(row["老系统关联到"])}id = {oldTableNameJX}.{Cast.ConToString(row["老系统字段名"])}\n");


            int newcount = 0;
            string newTableName = Cast.ConToString(row["新系统关联到"]) + "Base";
            if (!string.IsNullOrWhiteSpace(endBuilder_new.ToString()))
            {
                newcount = Regex.Matches(endBuilder_new.ToString(), Regex.Escape(newTableName)).Count;
            }
            Console.WriteLine($"newTableName:{newTableName}, newcount:{newcount}");
            string newTableNameJX = GetTableName(Cast.ConToString(row["新系统关联到"]));
            if (newcount > 0)
            {
                newTableNameJX = newTableNameJX + newcount.ToString();
            }
            bodyBuilder_new.Append($"       {newTableNameJX}.{Cast.ConToString(row["新系统关联到"])}id AS {Cast.ConToString(row["新系统字段名"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
            bool isNeedOrg = Cast.ConToString(row["新系统关联到"]) == "new_srv_station" || Cast.ConToString(row["新系统关联到"]) == "new_srv_worker" || Cast.ConToString(row["新系统关联到"]) == "new_accountstaff" || Cast.ConToString(row["新系统关联到"]) == "new_dot_conditiont";
            endBuilder_new.Append($"    LEFT JOIN {Cast.ConToString(row["新系统关联到"])}Base AS {newTableNameJX} ON {newTableNameJX}.{Cast.ConToString(row["新系统关联到字段"])} = main.{Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])} {(isNeedOrg ? ("AND " + newTableNameJX + ".new_organization_id = businessunit.businessunitid ") : "")}AND {newTableNameJX}.statecode = 0\n");
        }

        public static void SaveToTxt(string content, string filePath, bool append = false)
        {
            string dir = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            if (append)
                File.AppendAllText(filePath, content + Environment.NewLine);
            else
                File.WriteAllText(filePath, content);

            Console.WriteLine("成功!文件输出地址：" + filePath);
        }

        private static string GetTableName(string tableName)
        {
            int lastUnderscoreIndex = tableName.LastIndexOf('_');

            string result = lastUnderscoreIndex >= 0 && lastUnderscoreIndex < tableName.Length - 1
                ? tableName.Substring(lastUnderscoreIndex + 1)
                : tableName;

            return result;
        }

        private static string GetRelationshipName(string relationshipName, int index = 0)
        {
            var mappingArray = relationshipName.Split("=");
            if(mappingArray.Length == 2)
            {
                return mappingArray[index].Trim().ToLower();
            }

            return "new_code";
        }

        private static string GetOrgName(string orgName)
        {
            if (orgName.Contains("KAJ"))
            {
                return "徐工矿机";
            }
            else if (orgName.Contains("DL"))
            {
                return "徐工道路";
            }
            else if (orgName.Contains("CQ"))
            {
                return "徐工重庆";
            }
            else
            {
                return orgName;
            }
        }
    }
}
