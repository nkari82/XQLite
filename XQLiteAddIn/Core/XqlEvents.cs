// XqlEvents.cs
using System;

namespace XQLite.AddIn
{
    /// <summary>
    /// UI/백엔드 간 신호를 느슨하게 연결하는 경량 이벤트 허브.
    /// Ribbon, ExcelInterop, Sync 등에서 서로 직접 참조 없이 구독/발행.
    /// </summary>
    internal static class XqlEvents
    {
        public static event Action? SchemaChanged;   // 헤더(스키마) 편집/변경 감지
        public static event Action? RequestReevalCommit; // 커밋 버튼 즉시 재평가

        public static void RaiseSchemaChanged()
        {
            try { SchemaChanged?.Invoke(); } catch { /* ignore */ }
            try { RequestReevalCommit?.Invoke(); } catch { /* ignore */ }
        }

        public static void RaiseReevalCommit()
        {
            try { RequestReevalCommit?.Invoke(); } catch { /* ignore */ }
        }
    }
}
