﻿using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using RekTec.Crm.Common.Helper;
using System;
using System.Data;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml.Linq;
using static NPOI.HSSF.Util.HSSFColor;

namespace XCMG.SQLAuto.V1
{
    public static class Helper
    {
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
                    dataRow["新系统关联到字段"] = row.GetCell(12)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["老系统关联到字段"] = row.GetCell(13)?.ToString()?.Trim() ?? string.Empty;
                    dataRow["新系统表名"] = sheet.GetRow(1).GetCell(14)?.ToString()?.Trim() ?? string.Empty;
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
                endBuilder_New.Append($"FROM {orgName}_{Cast.ConToString(rowFirst["老系统表名"])} main\n");

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
       {Cast.ConToString(rowFirst["老系统表名"])}id AS new_oldid,
       owner.address1_telephone1 AS ownerid_address1_telephone1,
{builderBody.ToString()}{endBuilder.ToString()})t";

                string SQL_new = $@"
MERGE INTO {Cast.ConToString(rowFirst["新系统表名"])}Base t1
USING(
    SELECT
{bodyBuilder_New.ToString()}       main.*,
       owner.systemuserid AS new_systemuser_id
    {endBuilder_New.ToString()}    LEFT JOIN Systemuser AS owner ON owner.address1_telephone1 = main.ownerid_address1_telephone1 AND owner.isdisabled = 0
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
   t2.statecode,
   t2.CreatedOn,
   t2.ModifiedOn,
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   8,
   isnull(t2.new_systemuser_id,(SELECT SystemUserid FROM SystemUser WHERE fullname = '{adminName}')),
   isnull((SELECT businessunitid FROM SystemUser WHERE systemuserid = t2.new_systemuser_id),(SELECT businessunitid FROM SystemUser WHERE fullname = '{adminName}'))
)";

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
            bodyBuilder_old.Append($"       {GetTableName(Cast.ConToString(row["老系统关联到"]))}.{Cast.ConToString(row["老系统关联到字段"])} AS {Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
            endBuilder_old.Append($"    LEFT JOIN {dbName}.{Cast.ConToString(row["老系统关联到"])}Base AS {GetTableName(Cast.ConToString(row["老系统关联到"]))} ON {GetTableName(Cast.ConToString(row["老系统关联到"]))}.{Cast.ConToString(row["老系统关联到"])}id = {GetTableName(Cast.ConToString(row["老系统表名"]))}.{Cast.ConToString(row["老系统字段名"])}\n");


            //var oldModels = models?.OldModels?.Where(x => x.EntityName == Cast.ConToString(row["老系统关联到"])).ToList();//&& (x.Key != Cast.ConToString(row["老系统关联到字段"]) || x.Value != Cast.ConToString(row["老系统字段名"]))
            //var lists = (from model in models?.OldModels
            //            where model.EntityName == Cast.ConToString(row["老系统关联到"]) && model.Key != Cast.ConToString(row["老系统关联到字段"])
            //            select model).ToList();

            //if (oldModels?.Count == 0 )
            //{
            //    bodyBuilder_old.Append($"       {Cast.ConToString(row["老系统关联到"])}.{Cast.ConToString(row["老系统关联到字段"])} AS {Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
            //    endBuilder_old.Append($"    LEFT JOIN {dbName}.{Cast.ConToString(row["老系统关联到"])}Base AS {Cast.ConToString(row["老系统关联到"])} ON {Cast.ConToString(row["老系统关联到"])}id = {Cast.ConToString(row["老系统表名"])}.{Cast.ConToString(row["老系统字段名"])}\n");

                       //    var oldList = new List<LookupEntityModel>();
                       //    var oldModel = new LookupEntityModel
                       //    {
                       //        EntityName = Cast.ConToString(row["老系统关联到"]),
                       //        Key = Cast.ConToString(row["老系统关联到字段"]),
                       //        Value = Cast.ConToString(row["老系统字段名"]),
                       //    };

                       //    models?.OldModels.Add(oldModel);
                       //}
                       //else
                       //{
                       //    foreach ( var oldModel in oldModels)
                       //    {
                       //        if (oldModel.Key != Cast.ConToString(row["老系统关联到字段"]))
                       //        {
                       //            bodyBuilder_old.Append($"       {Cast.ConToString(row["老系统关联到"])}.{Cast.ConToString(row["老系统关联到字段"])} AS {Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
                       //        }
                       //        else if (oldModel.Value != Cast.ConToString(row["老系统字段名"]))
                       //        {
                       //            bodyBuilder_old.Append($"       {Cast.ConToString(row["老系统关联到"])}.{Cast.ConToString(row["老系统关联到字段"])} AS {Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
                       //            endBuilder_old.Append($"    LEFT JOIN {dbName}.{Cast.ConToString(row["老系统关联到"])}Base AS {Cast.ConToString(row["老系统关联到"])} ON {Cast.ConToString(row["老系统关联到"])}id = {Cast.ConToString(row["老系统表名"])}.{Cast.ConToString(row["老系统字段名"])}\n");
                       //        }
                       //        else
                       //        {

                       //        }
                       //    }
                       //}

            bodyBuilder_new.Append($"       {GetTableName(Cast.ConToString(row["新系统关联到"]))}.{Cast.ConToString(row["新系统关联到"])}id AS {Cast.ConToString(row["新系统字段名"])},    --{Cast.ConToString(row["新系统字段标签名"])}\n");
            endBuilder_new.Append($"    LEFT JOIN {Cast.ConToString(row["新系统关联到"])}Base AS {GetTableName(Cast.ConToString(row["新系统关联到"]))} ON {GetTableName(Cast.ConToString(row["新系统关联到"]))}.{Cast.ConToString(row["新系统关联到字段"])} = main.{Cast.ConToString(row["老系统字段名"])}_{Cast.ConToString(row["老系统关联到字段"])} AND {GetTableName(Cast.ConToString(row["新系统关联到"]))}.statecode = 0\n");
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

    }
}
