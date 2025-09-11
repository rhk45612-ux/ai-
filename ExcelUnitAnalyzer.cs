// ExcelUnitAnalyzer.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace MyApp
{
    /// <summary>
    /// 엑셀 "MCC 계산서" 시트를 분석해 (통과/불일치/ELB) 결과 문자열을 반환.
    /// </summary>
    public static partial class ExcelUnitAnalyzer
    {
        public static (string pass, string mismatch, string elb) Analyze(string filePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(filePath));
            var ws = package.Workbook.Worksheets.FirstOrDefault(s => s.Name == "MCC 계산서");
            if (ws == null) return ("", "시트 'MCC 계산서'를 찾을 수 없습니다.", "");
            if (ws.Dimension == null) return ("", "시트가 비어 있습니다.", "");

            int rows = ws.Dimension.End.Row;
            int cols = ws.Dimension.End.Column;

            // 2차원 텍스트 버퍼 구성
            var df = new string[rows, cols];
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    df[r - 1, c - 1] = ExcelUtils.Clean(ws.Cells[r, c].Text);

            // 블록 추출
            var blocks = ExcelUtils.ExtractBlocks(df);
            if (blocks.Count == 0)
                return ("", "'MCC FEEDER' 앵커를 찾지 못했습니다.", "");

            var passSections = new List<string>();
            var mismatchSections = new List<string>();
            var elbSections = new List<string>();

            foreach (var (name, startRow, endRow) in blocks)
            {
                var ctx = new BlockContext(name, startRow, endRow, cols);

                // 1) 컬럼 탐지
                ctx.Columns = DetectColumns(df, ctx);

                // 2) 스케줄 면수 추출
                ctx.ScheduleMyeon = ExtractScheduleMyeon(df, ctx);

                // 3) 유니트 집계(일반/인버터) + 폭합
                var unitAgg = CountUnits(df, ctx);

                // 4) +600 규칙 판단
                var plus600 = DecidePlus600(df, ctx, unitAgg);

                // 5) CT 규칙 적용
                var ct = ApplyCtRule(unitAgg, out int remainingWidthAfterCt);

                // 6) 최종 면수 산정
                var faces = ComputeFaces(unitAgg, plus600, ct, remainingWidthAfterCt);

                // 7) 섹션 문자열 빌드 + 분류(pass/mismatch/elb)
                BuildSection(
                    df, ctx, unitAgg, plus600, ct, faces, remainingWidthAfterCt,
                    out var sectionLines, out var classification, out var elbDetailLines
                );

                switch (classification)
                {
                    case SectionClass.Elb:
                        elbSections.Add(string.Join(Environment.NewLine, sectionLines));
                        if (elbDetailLines.Count > 0) elbSections.AddRange(elbDetailLines);
                        elbSections.Add("");
                        break;

                    case SectionClass.Pass:
                        passSections.Add(string.Join(Environment.NewLine, sectionLines));
                        passSections.Add("");
                        break;

                    case SectionClass.Mismatch:
                        mismatchSections.Add(string.Join(Environment.NewLine, sectionLines));
                        mismatchSections.Add("");
                        break;
                }
            }

            return (
                string.Join(Environment.NewLine, passSections),
                string.Join(Environment.NewLine, mismatchSections),
                string.Join(Environment.NewLine, elbSections)
            );
        }
    }
}
