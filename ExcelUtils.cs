using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MyApp
{
    public static class ExcelUtils
    {
        // 컴파일된 정규식으로 성능 미세개선
        private static readonly Regex WsRegex = new(@"\s+", RegexOptions.Compiled);

        public static string Clean(string input) =>
            string.IsNullOrWhiteSpace(input) ? "" : WsRegex.Replace(input, " ").Trim();

        public static string Normalize(string input) => Clean(input).ToUpperInvariant();

        /// <summary>
        /// 지정한 앵커 텍스트(기본: "MCC FEEDER")를 기준으로 블록 범위를 찾는다.
        /// </summary>
        public static List<(string name, int startRow, int endRow)> ExtractBlocks(
            string[,] df, string anchor = "MCC FEEDER")
        {
            if (df == null) throw new ArgumentNullException(nameof(df));

            var blocks = new List<(string name, int startRow, int endRow)>();
            int rows = df.GetLength(0);
            int cols = df.GetLength(1);

            if (rows == 0 || cols == 0) return blocks;

            var anchors = new List<(int row, int col, string name)>();

            // 1) 앵커 수집
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var cell = df[r, c] ?? string.Empty;
                    if (Normalize(cell) == Normalize(anchor))
                    {
                        // 앵커 바로 아래 셀을 블록명 후보로 사용
                        string name = (r + 1 < rows) ? Clean(df[r + 1, c] ?? string.Empty) : $"Block_{anchors.Count + 1}";
                        if (string.IsNullOrEmpty(name)) name = $"Block_{anchors.Count + 1}";
                        anchors.Add((r, c, name));
                    }
                }
            }

            if (anchors.Count == 0) return blocks;

            // 2) 범위 결정
            for (int i = 0; i < anchors.Count; i++)
            {
                int start = anchors[i].row;
                int end = (i + 1 < anchors.Count) ? anchors[i + 1].row - 2 : rows - 1;

                if (i + 1 == anchors.Count)
                {
                    // 마지막 블록: 실제 데이터가 있는 마지막 행까지 잡기
                    for (int r = rows - 1; r >= start; r--)
                    {
                        bool hasData = false;
                        for (int c = 0; c < cols; c++)
                        {
                            if (!string.IsNullOrWhiteSpace(Clean(df[r, c] ?? string.Empty)))
                            { hasData = true; break; }
                        }
                        if (hasData) { end = r; break; }
                    }
                }

                if (end < start) end = start; // 안전장치
                blocks.Add((anchors[i].name, start, end));
            }

            return blocks;
        }
    }
}
