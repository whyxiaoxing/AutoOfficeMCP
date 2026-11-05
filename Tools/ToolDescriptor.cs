using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace AutoOfficeMCP.Tools
{
    /// <summary>描述符中的单个工具参数。</summary>
    public sealed class ToolParameter
    {
        public string Name { get; set; }

        /// <summary>JSON Schema 基础类型，限定为 string|integer|number|boolean|array|object。</summary>
        public string Type { get; set; }

        public bool Required { get; set; }

        public string Description { get; set; }
    }

    /// <summary>
    /// 工具描述符（Tools\**\*.json）。
    /// 描述符只描述工具，不包含实现；执行由 IToolExecutor 认领。
    /// </summary>
    public sealed class ToolDescriptor
    {
        private static readonly HashSet<string> AllowedTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "string", "integer", "number", "boolean", "array", "object"
        };

        public string Id { get; set; }

        public string Name { get; set; }

        public string Description { get; set; }

        public string AppType { get; set; }

        public string Category { get; set; }

        public string RiskLevel { get; set; }

        public List<ToolParameter> Parameters { get; private set; }

        public string SourceFile { get; set; }

        public ToolDescriptor()
        {
            Parameters = new List<ToolParameter>();
        }

        /// <summary>MCP tools/call 使用的稳定工具名：优先 id，缺失时退回 name。</summary>
        public string ToolName
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(Id))
                {
                    return Id;
                }
                return Name;
            }
        }

        public bool IsHighRisk
        {
            get { return string.Equals(RiskLevel, "high", StringComparison.OrdinalIgnoreCase); }
        }

        /// <summary>
        /// 从单个描述符 JSON 构建模型。个别参数损坏时跳过并写 stderr，不阻断其它工具加载。
        /// </summary>
        public static ToolDescriptor FromJson(JObject root, string sourceFile, out string error)
        {
            error = null;
            if (root == null)
            {
                error = "描述符 JSON 不是对象。";
                return null;
            }

            string id = GetString(root, "id");
            string name = GetString(root, "name");
            if (string.IsNullOrWhiteSpace(id) && string.IsNullOrWhiteSpace(name))
            {
                error = "描述符缺少 id 与 name。";
                return null;
            }

            ToolDescriptor descriptor = new ToolDescriptor();
            descriptor.Id = string.IsNullOrWhiteSpace(id) ? name.Trim() : id.Trim();
            descriptor.Name = string.IsNullOrWhiteSpace(name) ? descriptor.Id : name.Trim();
            descriptor.Description = GetString(root, "description");
            if (string.IsNullOrWhiteSpace(descriptor.Description))
            {
                descriptor.Description = descriptor.Name;
            }

            string appType = GetString(root, "appType");
            descriptor.AppType = string.IsNullOrWhiteSpace(appType) ? "common" : appType.Trim().ToLowerInvariant();

            string category = GetString(root, "category");
            descriptor.Category = string.IsNullOrWhiteSpace(category) ? "未分类" : category.Trim();

            string riskLevel = GetString(root, "riskLevel");
            descriptor.RiskLevel = NormalizeRiskLevel(riskLevel, descriptor.ToolName);
            descriptor.SourceFile = sourceFile;

            JToken parametersToken = root["parameters"];
            if (parametersToken == null || parametersToken.Type == JTokenType.Null)
            {
                return descriptor;
            }

            JArray parametersArray = parametersToken as JArray;
            if (parametersArray == null)
            {
                error = "parameters 必须是数组。";
                return null;
            }

            HashSet<string> seenNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (JToken item in parametersArray)
            {
                JObject parameterObject = item as JObject;
                if (parameterObject == null)
                {
                    Console.Error.WriteLine("工具 " + descriptor.ToolName + " 的参数项不是对象，已跳过。");
                    continue;
                }

                string parameterName = GetString(parameterObject, "name");
                if (string.IsNullOrWhiteSpace(parameterName))
                {
                    Console.Error.WriteLine("工具 " + descriptor.ToolName + " 存在无名参数，已跳过。");
                    continue;
                }
                parameterName = parameterName.Trim();
                if (!seenNames.Add(parameterName))
                {
                    Console.Error.WriteLine("工具 " + descriptor.ToolName + " 参数重名：" +
                                            parameterName + "，已跳过重复项。");
                    continue;
                }

                ToolParameter parameter = new ToolParameter();
                parameter.Name = parameterName;
                parameter.Type = NormalizeParameterType(GetString(parameterObject, "type"), descriptor.ToolName, parameterName);
                parameter.Required = GetBool(parameterObject, "required");
                parameter.Description = GetString(parameterObject, "description");
                descriptor.Parameters.Add(parameter);
            }

            return descriptor;
        }

        /// <summary>按 MCP inputSchema 约定生成 properties / required。</summary>
        public JObject BuildInputSchema()
        {
            JObject properties = new JObject();
            JArray required = new JArray();
            foreach (ToolParameter parameter in Parameters)
            {
                JObject property = new JObject();
                property["type"] = parameter.Type;
                if (!string.IsNullOrWhiteSpace(parameter.Description))
                {
                    property["description"] = parameter.Description;
                }
                properties[parameter.Name] = property;
                if (parameter.Required)
                {
                    required.Add(parameter.Name);
                }
            }

            JObject schema = new JObject();
            schema["type"] = "object";
            schema["properties"] = properties;
            schema["required"] = required;
            return schema;
        }

        /// <summary>tools/list 的单项形状，只暴露参考客户端会读取的字段。</summary>
        public JObject BuildToolListItem()
        {
            JObject item = new JObject();
            item["name"] = ToolName;
            item["description"] = Description;
            item["inputSchema"] = BuildInputSchema();
            return item;
        }

        private static string NormalizeParameterType(string rawType, string toolName, string parameterName)
        {
            string normalized = (rawType ?? string.Empty).Trim().ToLowerInvariant();
            if (AllowedTypes.Contains(normalized))
            {
                return normalized;
            }

            Console.Error.WriteLine("工具 " + toolName + " 参数 " + parameterName + " 的类型 \"" + rawType
                                    + "\" 不在 MCP 白名单内，已按 string 处理。");
            return "string";
        }

        /// <summary>
        /// riskLevel 兼容参考项目的 safe，并统一为 low|medium|high。
        /// 未知值按 medium 处理，避免误把高危工具静默降级。
        /// </summary>
        private static string NormalizeRiskLevel(string rawRiskLevel, string toolName)
        {
            string normalized = (rawRiskLevel ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized == "safe")
            {
                return "low";
            }
            if (normalized == "low" || normalized == "medium" || normalized == "high")
            {
                return normalized;
            }

            Console.Error.WriteLine("工具 " + toolName + " 的 riskLevel \"" + rawRiskLevel +
                                    "\" 无法识别，按 medium 处理。");
            return "medium";
        }

        private static string GetString(JObject root, string propertyName)
        {
            JToken token = root[propertyName];
            if (token == null || token.Type == JTokenType.Null)
            {
                return string.Empty;
            }
            return token.Type == JTokenType.String ? token.Value<string>() : token.ToString();
        }

        private static bool GetBool(JObject root, string propertyName)
        {
            JToken token = root[propertyName];
            if (token == null)
            {
                return false;
            }
            try
            {
                return token.Value<bool>();
            }
            catch
            {
                return false;
            }
        }
    }
}
