using System;
using Word = Microsoft.Office.Interop.Word;

namespace Office.WordTools
{
    /// <summary>
    /// 一次 Word 写入的 COM 纪律外壳：
    /// 1) ScreenUpdating=false / DisplayAlerts=wdAlertsNone，并在 finally 恢复；
    /// 2) UndoRecord.StartCustomRecord("AutoOfficeMCP/工具名") / EndCustomRecord；
    /// 3) 不持有 Application，不在此类里调用 Undo（回滚必须核对指纹后由 T02 引擎执行）。
    /// </summary>
    internal sealed class WordWriteSession : IDisposable
    {
        private Word.Application _application;
        private Word.UndoRecord _undoRecord;
        private readonly bool _useCustomUndoRecord;
        private bool _screenUpdatingCaptured;
        private bool _screenUpdatingOriginal;
        private bool _displayAlertsCaptured;
        private Word.WdAlertLevel _displayAlertsOriginal;
        private bool _customRecordStarted;

        public bool UndoRecordStarted
        {
            get { return _customRecordStarted; }
        }

        public string CustomRecordName { get; private set; }

        /// <summary>true 表示本轮刻意不开启自定义撤销记录，依赖 Word 原生撤销。</summary>
        public bool NativeUndoEnabled
        {
            get { return !_useCustomUndoRecord; }
        }

        private WordWriteSession(Word.Application application, string toolName, bool useCustomUndoRecord)
        {
            _application = application;
            _useCustomUndoRecord = useCustomUndoRecord;
            CustomRecordName = "AutoOfficeMCP/" + toolName;
        }

        public static WordWriteSession Start(Word.Application application, string toolName)
        {
            return Start(application, toolName, true);
        }

        /// <summary>
        /// useCustomUndoRecord=false 时跳过 UndoRecord.StartCustomRecord：
        /// Office 2024 上「自定义 UndoRecord 开启期间执行 Find.Execute(wdReplaceAll)」会直接
        /// 杀死 WINWORD（0x800706BE），查找替换类工具必须走 Word 原生撤销。
        /// </summary>
        public static WordWriteSession Start(
            Word.Application application,
            string toolName,
            bool useCustomUndoRecord)
        {
            if (application == null)
            {
                throw new OfficeComException(OfficeErrorCode.APP_NOT_RUNNING, "Word 应用不可用，无法开始写入。");
            }

            WordWriteSession session = new WordWriteSession(application, toolName, useCustomUndoRecord);
            session.ApplyTemporarySettings();
            session.StartCustomRecord();
            return session;
        }

        public void Dispose()
        {
            if (_application == null)
            {
                return;
            }

            try
            {
                EndCustomRecord();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 结束自定义撤销记录失败：" + ex.Message);
            }
            finally
            {
                RestoreTemporarySettings();
                ComHelper.Release(_undoRecord);
                _undoRecord = null;
                _application = null;
            }
        }

        private void ApplyTemporarySettings()
        {
            try
            {
                _screenUpdatingOriginal = _application.ScreenUpdating;
                _screenUpdatingCaptured = true;
                _application.ScreenUpdating = false;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 关闭 ScreenUpdating 失败：" + ex.Message);
            }

            try
            {
                _displayAlertsOriginal = _application.DisplayAlerts;
                _displayAlertsCaptured = true;
                _application.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 关闭 DisplayAlerts 失败：" + ex.Message);
            }
        }

        private void StartCustomRecord()
        {
            if (!_useCustomUndoRecord)
            {
                Console.Error.WriteLine("[word] 本轮不使用自定义撤销记录（Word 原生撤销），避免 Office 2024 " +
                                        "UndoRecord + Find.Execute 崩溃。");
                return;
            }

            try
            {
                _undoRecord = _application.UndoRecord;
                if (_undoRecord == null)
                {
                    Console.Error.WriteLine("[word] Word 未提供 UndoRecord，无法开启自定义撤销记录。");
                    return;
                }

                // 已在录制其他自定义记录时绝不抢占；否则 EndCustomRecord 会把用户/上层的记录提前截断。
                if (_undoRecord.IsRecordingCustomRecord)
                {
                    Console.Error.WriteLine("[word] 当前已有自定义撤销记录「" +
                                            _undoRecord.CustomRecordName + "」，本次不嵌套开启。");
                    return;
                }

                _undoRecord.StartCustomRecord(CustomRecordName);
                _customRecordStarted = true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("[word] 开启自定义撤销记录失败：" + ex.Message);
            }
        }

        private void EndCustomRecord()
        {
            if (!_customRecordStarted || _undoRecord == null)
            {
                return;
            }
            _undoRecord.EndCustomRecord();
            _customRecordStarted = false;
        }

        private void RestoreTemporarySettings()
        {
            if (_displayAlertsCaptured)
            {
                try
                {
                    _application.DisplayAlerts = _displayAlertsOriginal;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[word] 恢复 DisplayAlerts 失败：" + ex.Message);
                }
            }

            if (_screenUpdatingCaptured)
            {
                try
                {
                    _application.ScreenUpdating = _screenUpdatingOriginal;
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine("[word] 恢复 ScreenUpdating 失败：" + ex.Message);
                }
            }
        }
    }
}
