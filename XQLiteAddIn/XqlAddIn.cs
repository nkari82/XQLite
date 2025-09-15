using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;


namespace XQLite.AddIn;


public sealed class XqlAddIn : IExcelAddIn
{
    public static Excel.Application App => (Excel.Application)ExcelDnaUtil.Application;
    private static XqlConfig? _cfg;


    public void AutoOpen()
    {
        // 외부 컨피그만 사용 (시트 읽지 않음)
        _cfg = XqlConfig.Load();


        // GraphQL 클라이언트 초기화
        XqlGraphQLClient.Init(_cfg);


        // 타이머 기반 서비스 시작 (시트 반영 없이 왕복만)
        XqlPresenceService.Start(_cfg);
        XqlSyncService.Start(_cfg);


        // 옵션: 데모용 테스트 행 하나 큐에 넣기 (원하면 주석 해제)
        // XqlSyncService.QueueUpsert("items", new Dictionary<string, object?>{ {"id", 1}, {"name", "Sword"}, {"deleted", 0} });
    }


    public void AutoClose()
    {
        XqlPresenceService.Stop();
        XqlSyncService.Stop();
    }
}