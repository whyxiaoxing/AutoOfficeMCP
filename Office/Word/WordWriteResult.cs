using System;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;

namespace Office.WordTools
{
    /// <summary>
    /// 单个 Word 写处理器返回的真实结果。
    /// Changed=null 表示让观察层用文档指纹/目标快照自动比对；
    /// 设置 Changed 时必须基于 Word 回读或明确的副作用，不得凭"语句未抛异常"。
    /// </summary>
    internal sealed class WordWriteResult
    {
        public bool Success { get; set; }

        public string Message { get; set; }

        public OfficeErrorCode ErrorCode { get; set; }

        public JObject Data { get; set; }

        public bool? Changed { get; set; }

        public int? AfterStart { get; set; }

        public int? AfterEnd { get; set; }

        public bool Structural { get; set; }

        public string UndoReason { get; set; }

        public List<string> TargetRefs { get; private set; }

        public WordWriteResult()
        {
            Message = string.Empty;
            ErrorCode = OfficeErrorCode.INTERNAL_ERROR;
            TargetRefs = new List<string>();
        }

        public static WordWriteResult Ok(string message, JObject data)
        {
            return Ok(message, data, null);
        }

        public static WordWriteResult Ok(string message, JObject data, bool? changed)
        {
            WordWriteResult result = new WordWriteResult();
            result.Success = true;
            result.Message = message ?? string.Empty;
            result.ErrorCode = OfficeErrorCode.INTERNAL_ERROR;
            result.Data = data;
            result.Changed = changed;
            return result;
        }

        public static WordWriteResult Fail(OfficeErrorCode errorCode, string message, JObject data)
        {
            return Fail(errorCode, message, data, null);
        }

        public static WordWriteResult Fail(
            OfficeErrorCode errorCode,
            string message,
            JObject data,
            bool? changed)
        {
            WordWriteResult result = new WordWriteResult();
            result.Success = false;
            result.Message = message ?? string.Empty;
            result.ErrorCode = errorCode;
            result.Data = data;
            result.Changed = changed;
            return result;
        }

        public WordWriteResult AddTargetRef(string targetRef)
        {
            if (!string.IsNullOrWhiteSpace(targetRef))
            {
                TargetRefs.Add(targetRef);
            }
            return this;
        }

        public WordWriteResult WithAfterRange(int start, int end)
        {
            AfterStart = start;
            AfterEnd = end;
            return this;
        }

        public WordWriteResult AsStructural(string undoReason)
        {
            Structural = true;
            UndoReason = undoReason;
            return this;
        }
    }
}
