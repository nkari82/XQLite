using System;

namespace XQLite.AddIn
{
    public static class XqlAdaptiveBatcher
    {
        // 대략적인 JSON 직렬화 크기를 추정해서 batch 크기를 조절 (보수적으로)
        public static int PickBatchSize(int rowCount, int min = 200, int max = 2000)
        {
            // 행이 매우 많은 경우 작은 배치로 시작, 적으면 크게
            if (rowCount > 20000) return min;  // 1k 미만 추천
            if (rowCount > 5000) return Math.Min(1000, max);
            return Math.Min(max, 1500);
        }
    }
}