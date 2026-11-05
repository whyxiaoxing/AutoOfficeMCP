using System;
using System.Collections.Generic;
using System.Globalization;
using Newtonsoft.Json.Linq;
using Office;

namespace Xlsx
{
    /// <summary>
    /// Excel 函数目录的版本元数据、旧版降级建议与公式模板。
    /// 目录主体仍在 Tools\excel\resources\excel_functions.json；本类只补充可靠性与兼容性信息。
    /// </summary>
    internal static class ExcelFunctionMetadata
    {
        private static readonly string[] Baseline = new string[] { "2016", "2019", "365", "WPS" };
        private static readonly string[] Since2019 = new string[] { "2019", "365", "WPS" };
        private static readonly string[] Since365 = new string[] { "365", "WPS" };

        private static readonly Dictionary<string, string> MinimumVersions =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "IFS", "2019" },
                { "SWITCH", "2019" },
                { "TEXTJOIN", "2019" },
                { "CONCAT", "2019" },
                { "MAXIFS", "2019" },
                { "MINIFS", "2019" },
                { "XLOOKUP", "365" },
                { "XMATCH", "365" },
                { "FILTER", "365" },
                { "UNIQUE", "365" },
                { "SORT", "365" },
                { "SORTBY", "365" },
                { "SEQUENCE", "365" },
                { "LET", "365" },
                { "LAMBDA", "365" },
                { "RANDARRAY", "365" },
                { "TAKE", "365" },
                { "DROP", "365" },
                { "CHOOSECOLS", "365" },
                { "CHOOSEROWS", "365" },
                { "HSTACK", "365" },
                { "VSTACK", "365" },
                { "TOCOL", "365" },
                { "TOROW", "365" },
                { "WRAPROWS", "365" },
                { "WRAPCOLS", "365" },
                { "EXPAND", "365" },
                { "MAP", "365" },
                { "REDUCE", "365" },
                { "SCAN", "365" },
                { "BYROW", "365" },
                { "BYCOL", "365" },
                { "TEXTSPLIT", "365" },
                { "TEXTBEFORE", "365" },
                { "TEXTAFTER", "365" },
                { "REGEXTEST", "365" },
                { "REGEXEXTRACT", "365" },
                { "REGEXREPLACE", "365" }
            };

        private static readonly HashSet<string> Since2019Names = new HashSet<string>(
            new string[]
            {
                "IFS", "SWITCH", "TEXTJOIN", "CONCAT", "MAXIFS", "MINIFS"
            },
            StringComparer.OrdinalIgnoreCase);

        public static JObject Enrich(JObject function)
        {
            if (function == null)
            {
                return new JObject();
            }
            JObject enriched = (JObject)function.DeepClone();
            string name = enriched.Value<string>("name") ?? string.Empty;
            string minimumVersion;
            string[] availability;
            if (MinimumVersions.TryGetValue(name, out minimumVersion))
            {
                availability = minimumVersion == "2019" ? Since2019 : Since365;
            }
            else
            {
                minimumVersion = "2016";
                availability = Baseline;
            }
            enriched["minVersion"] = minimumVersion;
            JArray versions = new JArray();
            foreach (string version in availability)
            {
                versions.Add(version);
            }
            enriched["availability"] = versions;
            enriched["wpsAvailable"] = Contains(availability, "WPS");
            enriched["versionNote"] = "minVersion=" + minimumVersion +
                                      "；可用版本：" + string.Join("/", availability) +
                                      "。静态目录标注，实际以本机 Excel/WPS 版本为准。";

            JObject fallback = FindFallback(name);
            if (fallback != null)
            {
                enriched["fallback"] = fallback;
            }
            return enriched;
        }

        public static bool IsAvailableInVersion(JObject function, string version)
        {
            if (function == null || string.IsNullOrWhiteSpace(version))
            {
                return true;
            }
            string name = function.Value<string>("name") ?? string.Empty;
            string minimumVersion;
            string[] availability;
            if (MinimumVersions.TryGetValue(name, out minimumVersion))
            {
                availability = minimumVersion == "2019" ? Since2019 : Since365;
            }
            else
            {
                availability = Baseline;
            }
            return Contains(availability, version.Trim());
        }

        public static OperationResult QueryFallback(JObject arguments)
        {
            string keyword = ExcelArgumentReader.GetTrimmedString(arguments, "keyword", "query", "search", "q");
            string[] terms = SplitTerms(keyword);
            int limit = ExcelArgumentReader.GetInt(arguments, 50, "limit");
            limit = Math.Max(1, Math.Min(limit, 500));

            JArray all = BuildFallbacks();
            JArray matches = new JArray();
            int total = 0;
            foreach (JToken item in all)
            {
                if (!Matches(item, terms))
                {
                    continue;
                }
                total++;
                if (matches.Count < limit)
                {
                    matches.Add(item.DeepClone());
                }
            }

            JObject data = new JObject();
            data["mode"] = "fallback";
            data["keyword"] = keyword;
            data["totalMatches"] = total;
            data["returned"] = matches.Count;
            data["truncated"] = total > matches.Count;
            data["fallbacks"] = matches;
            data["note"] = "降级建议按兼容性给出；旧版 Excel 使用数组公式时请按 Ctrl+Shift+Enter（365 除外）。";
            string message = keyword.Length == 0
                ? "函数降级建议共 " + total + " 条，返回 " + matches.Count + " 条。"
                : "函数降级建议匹配 " + total + " 条，返回 " + matches.Count + " 条。";
            return OperationResult.Succeed(message, data);
        }

        public static OperationResult QueryTemplates(JObject arguments)
        {
            string keyword = ExcelArgumentReader.GetTrimmedString(arguments, "keyword", "query", "search", "q");
            string[] terms = SplitTerms(keyword);
            int limit = ExcelArgumentReader.GetInt(arguments, 50, "limit");
            limit = Math.Max(1, Math.Min(limit, 50));

            JArray all = BuildTemplates();
            JArray matches = new JArray();
            int total = 0;
            foreach (JToken item in all)
            {
                if (!Matches(item, terms))
                {
                    continue;
                }
                total++;
                if (matches.Count < limit)
                {
                    matches.Add(item.DeepClone());
                }
            }

            JObject data = new JObject();
            data["mode"] = "templates";
            data["keyword"] = keyword;
            data["totalMatches"] = total;
            data["returned"] = matches.Count;
            data["templates"] = matches;
            data["note"] = "模板通过 excel_apply_formula / excel_evaluate 落地；不新增独立模板工具。";
            string message = keyword.Length == 0
                ? "公式模板共 " + total + " 条，返回 " + matches.Count + " 条。"
                : "公式模板匹配 " + total + " 条，返回 " + matches.Count + " 条。";
            return OperationResult.Succeed(message, data);
        }

        private static JObject FindFallback(string functionName)
        {
            if (string.IsNullOrWhiteSpace(functionName))
            {
                return null;
            }
            foreach (JToken item in BuildFallbacks())
            {
                JObject fallback = item as JObject;
                if (fallback != null &&
                    string.Equals(
                        fallback.Value<string>("function"),
                        functionName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return (JObject)fallback.DeepClone();
                }
            }
            return null;
        }

        private static JArray BuildFallbacks()
        {
            JArray items = new JArray();
            items.Add(BuildFallback(
                "XLOOKUP",
                "INDEX+MATCH",
                "=INDEX(返回列,MATCH(查找值,查找列,0))",
                "INDEX+MATCH 在所有 Excel 2016+ 版本可用；需要未命中时返回空值，请外包 IFERROR。",
                "=IFERROR(INDEX(返回列,MATCH(查找值,查找列,0)),\"\")"));
            items.Add(BuildFallback(
                "FILTER",
                "辅助列+IF",
                "=IF(条件,原值,\"\")",
                "旧版无动态数组。先用辅助列逐行判断，再用筛选/排序或 INDEX+SMALL+IF 数组公式收集结果。",
                "=IFERROR(INDEX(明细!A:A,SMALL(IF(明细!$C$2:$C$1000=\"华东\",ROW($2:$1000)),ROW(A1))),\"\")"));
            items.Add(BuildFallback(
                "SORT",
                "辅助列+排名/Excel 排序",
                "=RANK.EQ(排序键,排序键区域,0)+COUNTIF($排序键$2:当前键,当前键)-1",
                "旧版可用辅助排名列，或直接用 excel_sort_data 多键排序。",
                "=RANK.EQ(C2,$C$2:$C$1000,1)+COUNTIF($C$2:C2,C2)-1"));
            items.Add(BuildFallback(
                "SORTBY",
                "辅助列+多键排序",
                "=RANK.EQ(主键,主键区域,0)+COUNTIF($主键$2:当前主键,当前主键)-1",
                "多键排序建议用辅助排名列或 excel_sort_data 的 sortKeys。",
                ""));
            items.Add(BuildFallback(
                "UNIQUE",
                "COUNTIF+IF",
                "=IF(COUNTIF($A$2:A2,A2)=1,A2,\"\")",
                "旧版逐行标记首次出现，再筛选非空；去重也可用 excel_remove_duplicates。",
                "=IF(COUNTIF($A$2:A2,A2)=1,A2,\"\")"));
            items.Add(BuildFallback(
                "TEXTJOIN",
                "CONCATENATE/&",
                "=A2&分隔符&B2",
                "旧版用 & 或 CONCATENATE 拼接；批量拼接可在辅助列累积。",
                "=TRIM(A2&\" \"&B2)"));
            items.Add(BuildFallback(
                "CONCAT",
                "CONCATENATE/&",
                "=CONCATENATE(A2,B2)",
                "CONCATENATE 在所有版本可用；区域拼接需要辅助列。",
                "=A2&B2"));
            items.Add(BuildFallback(
                "IFS",
                "嵌套 IF",
                "=IF(条件1,值1,IF(条件2,值2,默认值))",
                "嵌套 IF 兼容所有版本；超过 7 层时可拆分辅助列。",
                "=IF(A2>=90,\"A\",IF(A2>=60,\"B\",\"C\"))"));
            items.Add(BuildFallback(
                "SWITCH",
                "嵌套 IF",
                "=IF(A2=\"甲\",1,IF(A2=\"乙\",2,0))",
                "用嵌套 IF 或 VLOOKUP 映射表替代。",
                "=IF(A2=\"甲\",1,IF(A2=\"乙\",2,0))"));
            items.Add(BuildFallback(
                "MAXIFS",
                "数组公式 MAX(IF(...))",
                "=MAX(IF(条件区域=条件,值区域))",
                "旧版按 Ctrl+Shift+Enter 输入数组公式。",
                "=MAX(IF($A$2:$A$1000=\"华东\",$D$2:$D$1000))"));
            items.Add(BuildFallback(
                "MINIFS",
                "数组公式 MIN(IF(...))",
                "=MIN(IF(条件区域=条件,值区域))",
                "旧版按 Ctrl+Shift+Enter 输入数组公式。",
                "=MIN(IF($A$2:$A$1000=\"华东\",$D$2:$D$1000))"));
            items.Add(BuildFallback(
                "TEXTSPLIT",
                "分列/LEFT+MID+FIND",
                "=LEFT(A2,FIND(\",\",A2)-1)",
                "旧版用“数据-分列”或 LEFT/MID/FIND 组合拆分。",
                "=MID(A2,FIND(\",\",A2)+1,255)"));
            items.Add(BuildFallback(
                "TEXTBEFORE",
                "LEFT+FIND",
                "=LEFT(A2,FIND(分隔符,A2)-1)",
                "找不到分隔符时外包 IFERROR。",
                "=IFERROR(LEFT(A2,FIND(\",\",A2)-1),A2)"));
            items.Add(BuildFallback(
                "TEXTAFTER",
                "MID+FIND",
                "=MID(A2,FIND(分隔符,A2)+1,255)",
                "找不到分隔符时外包 IFERROR。",
                "=IFERROR(MID(A2,FIND(\",\",A2)+1,255),\"\")"));
            items.Add(BuildFallback(
                "LET",
                "辅助单元格/名称管理器",
                "把中间结果写入辅助列，再用普通公式引用",
                "旧版没有 LET；辅助列或定义名称可读性更好。",
                ""));
            items.Add(BuildFallback(
                "LAMBDA",
                "辅助列/VBA 自定义函数",
                "把计算拆成辅助列，或使用受控 VBA",
                "旧版没有 LAMBDA；VBA 需 config.json 放行 security.allowVba。",
                ""));
            items.Add(BuildFallback(
                "SEQUENCE",
                "ROW(A1) 下拉",
                "=ROW(A1)",
                "用 ROW(A1) 下拉生成 1..N 序列。",
                "=ROW(A1)"));
            items.Add(BuildFallback(
                "RANDARRAY",
                "RAND",
                "=RAND()",
                "逐单元格 RAND 后复制为值。",
                "=RAND()"));
            items.Add(BuildFallback(
                "REGEXTEST",
                "WPS 正则/VBA RegExp",
                "旧版 Excel 无原生正则；WPS 或 VBA RegExp 替代",
                "项目已提供 excel_regex_extract 工具，可先抽取再判断。",
                ""));
            items.Add(BuildFallback(
                "REGEXEXTRACT",
                "excel_regex_extract 工具",
                "使用 excel_regex_extract 批量抽取后回写",
                "旧版 Excel 没有 REGEXEXTRACT；WPS 新版支持。",
                ""));
            items.Add(BuildFallback(
                "REGEXREPLACE",
                "WPS 正则/VBA RegExp",
                "旧版用 SUBSTITUTE 组合或 VBA RegExp.Replace",
                "复杂替换建议先在辅助列生成，再复制为值。",
                ""));
            items.Add(BuildFallback(
                "TAKE",
                "INDEX/OFFSET",
                "=INDEX(区域,开始行,开始列):INDEX(区域,结束行,结束列)",
                "用 INDEX 构造起始/结束单元格。",
                "=INDEX(A:A,2):INDEX(A:A,10)"));
            items.Add(BuildFallback(
                "DROP",
                "INDEX/OFFSET",
                "=INDEX(区域,跳过行数+1,1)",
                "用 INDEX 从指定行开始取子区域。",
                "=INDEX(A:A,3):INDEX(A:A,100)"));
            return items;
        }

        private static JObject BuildFallback(
            string function,
            string replacement,
            string formula,
            string note,
            string example)
        {
            JObject item = new JObject();
            item["function"] = function;
            item["replacement"] = replacement;
            item["formula"] = formula;
            item["note"] = note;
            if (!string.IsNullOrWhiteSpace(example))
            {
                item["example"] = example;
            }
            return item;
        }

        private static JArray BuildTemplates()
        {
            JArray templates = new JArray();

            JObject reconcile = new JObject();
            reconcile["id"] = "reconcile";
            reconcile["name"] = "对账";
            reconcile["purpose"] = "明细与汇总按业务键核对金额差额，定位未达账/重复/金额错误。";
            reconcile["requiredVersion"] = "2016+";
            JArray reconcileFormulas = new JArray();
            reconcileFormulas.Add(BuildFormula("按业务键汇总明细", "=SUMIFS(明细!$D$2:$D$100000,明细!$A$2:$A$100000,$A2)"));
            reconcileFormulas.Add(BuildFormula("计算差额", "=ROUND(B2-C2,2)"));
            reconcileFormulas.Add(BuildFormula("标记差额超阈值", "=IF(ABS(D2)>0.005,\"差异\",\"一致\")"));
            reconcileFormulas.Add(BuildFormula("查找明细中不存在的汇总键", "=IF(COUNTIF(明细!$A:$A,$A2)=0,\"缺明细\",\"\")"));
            reconcile["formulas"] = reconcileFormulas;
            reconcile["steps"] = new JArray(
                "先清洗明细键与汇总键（TRIM/ASC），确保键完全一致。",
                "用 SUMIFS 按键汇总明细金额，避免直接 VLOOKUP 只取首行。",
                "差额列用 ROUND(...,2) 固化两位小数，再按 ABS(差额)>tolerance 筛选。",
                "对差额非零的键回查重复键、方向相反记录和跨期记录。");
            reconcile["notes"] = "SUMIFS 的求和列与条件列必须使用绝对引用；复制到 500 行后建议用 excel_formula_reference_audit 审计 $ 引用。";
            templates.Add(reconcile);

            JObject groupKey = new JObject();
            groupKey["id"] = "group_key_check";
            groupKey["name"] = "分组键脏值";
            groupKey["purpose"] = "检查前后空格、全半角、不可见字符、尾随空格造成的同键不同值。";
            groupKey["requiredVersion"] = "2016+";
            JArray keyFormulas = new JArray();
            keyFormulas.Add(BuildFormula("前后空格", "=LEN(A2)<>LEN(TRIM(A2))"));
            keyFormulas.Add(BuildFormula("尾随空格导致重复", "=COUNTIF($A:$A,A2&\" \")>0"));
            keyFormulas.Add(BuildFormula("全角字符", "=LEN(A2)<>LENB(A2)"));
            keyFormulas.Add(BuildFormula("全角字母数字", "=ASC(A2)<>A2"));
            keyFormulas.Add(BuildFormula("去不可见字符后的标准键", "=ASC(TRIM(SUBSTITUTE(SUBSTITUTE(A2,CHAR(9),\"\"),CHAR(160),\"\")))"));
            keyFormulas.Add(BuildFormula("键是否重复", "=COUNTIF($A:$A,$A2)>1"));
            groupKey["formulas"] = keyFormulas;
            groupKey["steps"] = new JArray(
                "在辅助列生成标准键，原键保留用于追溯。",
                "分别统计各类脏值数量，抽样查看原文。",
                "确认规则后用标准键替换分组键，并复制为值。",
                "再次运行重复/脏值检查，确认归零。");
            groupKey["notes"] = "LENB 在部分新版函数中行为变化，仅用于辅助判断；核心清洗仍以 TRIM/ASC/SUBSTITUTE 为准。";
            templates.Add(groupKey);

            JObject filterRows = new JObject();
            filterRows["id"] = "filter_rows";
            filterRows["name"] = "条件取行";
            filterRows["purpose"] = "按条件抽取明细行；365 用 FILTER，旧版用辅助列+INDEX/SMALL+IF。";
            filterRows["requiredVersion"] = "365（FILTER）/ 2016+（辅助列兜底）";
            JArray filterFormulas = new JArray();
            filterFormulas.Add(BuildFormula(
                "365 FILTER",
                "=FILTER(明细!$A$2:$D$1000,明细!$C$2:$C$1000=\"华东\",\"无匹配\")"));
            filterFormulas.Add(BuildFormula(
                "旧版辅助列",
                "=IF(明细!$C2=\"华东\",ROW(),\"\")"));
            filterFormulas.Add(BuildFormula(
                "旧版 INDEX+SMALL+IF（数组公式）",
                "=IFERROR(INDEX(明细!A:A,SMALL(IF(明细!$C$2:$C$1000=\"华东\",ROW($2:$1000)),ROW(A1))),\"\")"));
            filterFormulas.Add(BuildFormula(
                "多条件辅助列",
                "=IF(AND(明细!$C2=\"华东\",明细!$D2>=1000),ROW(),\"\")"));
            filterRows["formulas"] = filterFormulas;
            filterRows["steps"] = new JArray(
                "确认目标 Excel 版本；365 优先使用 FILTER。",
                "旧版把条件写入辅助列，再用 SMALL/ROW 取第 n 个匹配行号。",
                "用 IFERROR 包裹，避免下拉到底部出现 #NUM!。",
                "结果复制为值后再删除辅助列。");
            filterRows["notes"] = "FILTER 是动态数组，溢出区域不能有内容；旧版数组公式需 Ctrl+Shift+Enter。";
            templates.Add(filterRows);

            JObject textDate = new JObject();
            textDate["id"] = "text_date_normalize";
            textDate["name"] = "文本/日期规范化";
            textDate["purpose"] = "把混合分隔符、中文日期、全角字符规范为可计算日期/文本。";
            textDate["requiredVersion"] = "2016+";
            JArray dateFormulas = new JArray();
            dateFormulas.Add(BuildFormula(
                "中文日期转日期",
                "=TEXT(DATEVALUE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(TRIM(A2),\"年\",\"-\"),\"月\",\"-\"),\"日\",\"\")),\"yyyy-mm-dd\")"));
            dateFormulas.Add(BuildFormula(
                "纯数字日期 YYYYMMDD",
                "=DATE(LEFT(A2,4),MID(A2,5,2),RIGHT(A2,2))"));
            dateFormulas.Add(BuildFormula(
                "日期序列转文本",
                "=TEXT(A2,\"yyyy-mm-dd\")"));
            dateFormulas.Add(BuildFormula(
                "全角转半角并去空格",
                "=ASC(TRIM(SUBSTITUTE(A2,CHAR(160),\"\")))"));
            dateFormulas.Add(BuildFormula(
                "统一分隔符",
                "=SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2,\".\",\"-\"),\"/\",\"-\"),\"年\",\"-\")"));
            textDate["formulas"] = dateFormulas;
            textDate["steps"] = new JArray(
                "先抽样识别日期形态：文本、序列值、中文年月日、斜杠/点号分隔。",
                "生成规范化辅助列，用 DATEVALUE/DATE 转成真日期。",
                "日期列设置 yyyy-mm-dd 格式并复制为值。",
                "用 excel_precision_check 或抽样 evaluate 验证转换结果。");
            textDate["notes"] = "复杂正则抽取请用 excel_regex_extract；DATEVALUE 受系统区域设置影响，跨区域环境优先用 DATE 拆分。";
            templates.Add(textDate);

            return templates;
        }

        private static JObject BuildFormula(string label, string formula)
        {
            JObject item = new JObject();
            item["label"] = label;
            item["formula"] = formula;
            return item;
        }

        private static string[] SplitTerms(string keyword)
        {
            return (keyword ?? string.Empty).Split(
                new char[] { ' ', '\t', '\r', '\n', ',', '，' },
                StringSplitOptions.RemoveEmptyEntries);
        }

        private static bool Matches(JToken item, string[] terms)
        {
            if (terms == null || terms.Length == 0)
            {
                return true;
            }
            string haystack = item.ToString();
            foreach (string term in terms)
            {
                if (haystack.IndexOf(term, StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }
            return true;
        }

        private static bool Contains(string[] values, string target)
        {
            foreach (string value in values)
            {
                if (string.Equals(value, target, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }
}
