// ExcelUnitAnalyzer.Internal.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace MyApp
{
    // ===== 내부 모델 =====
    internal enum SectionClass { Pass, Mismatch, Elb }

    internal sealed class ColumnMap
    {
        public int UnitSizeCol = -1;
        public int TypeCol = -1;
        public int MyeonCol = -1;
        public int ElbCol = -1;
    }

    internal sealed class BlockContext
    {
        public string Name { get; }
        public int StartRow { get; }
        public int EndRow { get; }
        public int Cols { get; }
        public ColumnMap Columns { get; set; } = new();
        public int? ScheduleMyeon { get; set; }

        public BlockContext(string name, int startRow, int endRow, int cols)
        {
            Name = name;
            StartRow = startRow;
            EndRow = endRow;
            Cols = cols;
        }
    }

    internal sealed class UnitAggregate
    {
        // 일반 유니트 개수
        public Dictionary<string, int> NormalCounts { get; } =
            new() { ["600"] = 0, ["800"] = 0, ["900"] = 0, ["1200"] = 0 };

        public int NormalWidthSum { get; set; } = 0;

        // 인버터
        public int InverterW600Count { get; set; } = 0;
        public int InverterW800Count { get; set; } = 0;
        public int InverterW600ExtraMyeon { get; set; } = 0;

        public int MccbCount => NormalCounts.Values.Sum(); // 인버터 제외
    }

    internal sealed class Plus600Decision
    {
        public bool FirstMyeonBlank { get; set; }
        public bool HasBlankMyeonOnAny600 { get; set; }
        public bool MultipleOf3_600 { get; set; }
        public bool ForcePlus600BySpare { get; set; }
        public bool UsePlus600 => FirstMyeonBlank || MultipleOf3_600 || ForcePlus600BySpare;
    }

    internal sealed class CtResult
    {
        public bool Configured { get; set; }
        public int CtFaces { get; set; }
        public int CtWidth { get; set; }
        public string Label { get; set; } = "";
    }

    internal sealed class FaceSummary
    {
        public int NormalFaces { get; set; }
        public int InverterFaces { get; set; }
        public int CtFaces { get; set; }
        public int Total => NormalFaces + InverterFaces + CtFaces;
    }

    // ===== 내부 로직 =====
    internal static partial class ExcelUnitAnalyzer
    {
        internal static ColumnMap DetectColumns(string[,] df, BlockContext ctx)
        {
            var map = new ColumnMap();
            for (int r = ctx.StartRow; r <= ctx.EndRow; r++)
            {
                for (int c = 0; c < ctx.Cols; c++)
                {
                    string val = ExcelUtils.Normalize(df[r, c]);
                    if (val == "UNIT SIZE") map.UnitSizeCol = c;
                    if (val == "TYPE") map.TypeCol = c;
                    if (val.Replace(" ", "") == "면수") map.MyeonCol = c;
                    if (val.Replace(" ", "") == "ELB(AF/AT)") map.ElbCol = c;
                }
            }
            return map;
        }

        internal static int? ExtractScheduleMyeon(string[,] df, BlockContext ctx)
        {
            if (ctx.Columns.MyeonCol == -1) return null;

            var nums = new List<int>();
            for (int r = ctx.StartRow + 1; r <= ctx.EndRow; r++)
            {
                string s = ExcelUtils.Clean(df[r, ctx.Columns.MyeonCol]);
                foreach (Match m in Regex.Matches(s, @"\d+"))
                    if (int.TryParse(m.Value, out int n)) nums.Add(n);
            }
            return nums.Count > 0 ? nums.Last() : (int?)null;
        }

        internal static UnitAggregate CountUnits(string[,] df, BlockContext ctx)
        {
            var agg = new UnitAggregate();
            int uCol = ctx.Columns.UnitSizeCol;
            int tCol = ctx.Columns.TypeCol;

            if (uCol == -1) return agg;

            for (int r = ctx.StartRow + 1; r <= ctx.EndRow; r++)
            {
                string raw = ExcelUtils.Clean(df[r, uCol]);
                if (string.IsNullOrEmpty(raw)) continue;

                string up = raw.ToUpperInvariant();
                if (up.Contains("W:800"))
                {
                    agg.InverterW800Count++;
                    continue;
                }
                if (up.Contains("W:600"))
                {
                    agg.InverterW600Count++;
                    if (tCol != -1)
                    {
                        string t = ExcelUtils.Clean(df[r, tCol]);
                        if (t == "RI3S12O7L1G-1") agg.InverterW600ExtraMyeon += 1;
                        else if (t == "RI3S12O7L1G-2") agg.InverterW600ExtraMyeon += 2;
                    }
                    continue;
                }

                if (agg.NormalCounts.ContainsKey(raw))
                {
                    agg.NormalCounts[raw]++;
                    agg.NormalWidthSum += raw switch
                    {
                        "600" => 600,
                        "800" => 800,
                        "900" => 900,
                        "1200" => 1200,
                        _ => 0
                    };
                }
            }

            return agg;
        }

        internal static Plus600Decision DecidePlus600(string[,] df, BlockContext ctx, UnitAggregate agg)
        {
            var d = new Plus600Decision();

            // 600 중 '면수' 공란이 있는지
            if (ctx.Columns.MyeonCol != -1 && ctx.Columns.UnitSizeCol != -1)
            {
                for (int r = ctx.StartRow + 1; r <= ctx.EndRow; r++)
                {
                    string u = ExcelUtils.Clean(df[r, ctx.Columns.UnitSizeCol]).ToUpperInvariant();
                    if (u == "600" && string.IsNullOrWhiteSpace((df[r, ctx.Columns.MyeonCol] ?? "").Trim()))
                    { d.HasBlankMyeonOnAny600 = true; break; }
                }
            }

            // 첫 데이터 행의 면수 공란 여부
            d.FirstMyeonBlank = true;
            if (ctx.Columns.MyeonCol != -1 && ctx.Columns.UnitSizeCol != -1)
            {
                int firstDataRow = -1;
                for (int r = ctx.StartRow + 1; r <= ctx.EndRow; r++)
                {
                    if (!string.IsNullOrWhiteSpace(ExcelUtils.Clean(df[r, ctx.Columns.UnitSizeCol])))
                    { firstDataRow = r; break; }
                }
                if (firstDataRow != -1)
                    d.FirstMyeonBlank = string.IsNullOrWhiteSpace((df[firstDataRow, ctx.Columns.MyeonCol] ?? "").Trim());
            }

            // 600 개수가 3의 배수인가
            agg.NormalCounts.TryGetValue("600", out int cnt600);
            d.MultipleOf3_600 = (cnt600 > 0 && (cnt600 % 3 == 0));

            // 스페어 흡수 규칙
            d.ForcePlus600BySpare =
                (cnt600 > 0 && !d.MultipleOf3_600 && d.HasBlankMyeonOnAny600)
                && (agg.InverterW600Count > 0 || agg.InverterW800Count > 0);

            return d;
        }

        internal static CtResult ApplyCtRule(UnitAggregate agg, Plus600Decision plus600, out int remainingWidthAfterCt)
        {
            remainingWidthAfterCt = agg.NormalWidthSum;
            var ret = new CtResult { Configured = false, CtFaces = 0, CtWidth = 0, Label = "" };

            agg.NormalCounts.TryGetValue("600", out int cnt600Normal);
            if (cnt600Normal < 1) return ret; // 메인600 없으면 CT 미적용

            int mccb = agg.MccbCount;

            string label = "";
            int ctWidth = 0;

            // 메인 600 제거
            agg.NormalCounts["600"] -= 1;
            remainingWidthAfterCt -= 600;

            if (mccb <= 4)
            {
                // 유니트 4개 이하
                ctWidth = 600 + 300 + 300 + 600; // 메인600+CT300+CT300+유닛600
                label = "(유니트≤4: 메인600 + CT300 + CT300 + 유닛600)";
                if (agg.NormalCounts["600"] > 0)
                {
                    agg.NormalCounts["600"] -= 1;
                    remainingWidthAfterCt -= 600;
                }
            }
            else if (mccb >= 5 && mccb <= 11)
            {
                ctWidth = 600 + 600 + 300;  // 메인600 + CT600 + CT300
                label = "(MCCB 5~11: 메인600 + CT600 + CT300)";
            }
            else
            {
                ctWidth = 600 + 900 + 300;  // 메인600 + CT900 + CT300
                label = "(MCCB 12~: 메인600 + CT900 + CT300)";
            }

            ret.Configured = true;
            ret.CtWidth = ctWidth;
            ret.CtFaces = (int)Math.Ceiling(ctWidth / 1800.0);
            ret.Label = label;
            return ret;
        }

        internal static FaceSummary ComputeFaces(UnitAggregate agg, Plus600Decision plus600, CtResult ct, int remainingWidthAfterCt)
        {
            int adjustedWidth = remainingWidthAfterCt + (plus600.UsePlus600 ? 600 : 0);
            int normalFaces = adjustedWidth > 0 ? (int)Math.Ceiling(adjustedWidth / 1800.0) : 0;
            int inverterFaces = agg.InverterW600ExtraMyeon + (agg.InverterW800Count * 2);

            return new FaceSummary
            {
                NormalFaces = normalFaces,
                InverterFaces = inverterFaces,
                CtFaces = ct.CtFaces
            };
        }

        internal static void BuildSection(
            string[,] df,
            BlockContext ctx,
            UnitAggregate agg,
            Plus600Decision plus600,
            CtResult ct,
            FaceSummary faces,
            out List<string> sectionLines,
            out SectionClass classification,
            out List<string> elbDetailLines)
        {
            sectionLines = new List<string> { $"[{ctx.Name}]" };

            // 스케줄 면수 라인
            sectionLines.Add($"→ 스케줄 표 면수: {(ctx.ScheduleMyeon?.ToString() ?? "없음")}");

            // 설명 파트
            var parts = new List<string>();
            foreach (var kv in agg.NormalCounts)
                if (kv.Value > 0) parts.Add($"{kv.Key}: {kv.Value}ea");

            if (agg.InverterW600Count > 0)
                parts.Add($"인버터 W:600 {agg.InverterW600Count}ea(가산 {agg.InverterW600ExtraMyeon}면)");
            if (agg.InverterW800Count > 0)
                parts.Add($"인버터 W:800 {agg.InverterW800Count}ea(가산 {agg.InverterW800Count * 2}면)");

            if (agg.NormalWidthSum > 0)
                parts.Add($"일반폭 합계 {(agg.NormalWidthSum - (ct.Configured ? 600 : 0) - (ct.Configured && ctx.Columns.UnitSizeCol != -1 ? 0 : 0))}" +
                          $"{(plus600.UsePlus600 ? " +600" : "")} → {faces.NormalFaces}면(@1800)");

            if (ct.Configured) parts.Add($"CT {faces.CtFaces}면 포함");
            else parts.Add(plus600.UsePlus600
                                         ? (plus600.ForcePlus600BySpare ? "+600 규칙(남은 600 SPARE 흡수) 적용" : "+600 규칙 적용")
                                         : "+1면 규칙 적용");

            sectionLines.Add($"→ {string.Join(", ", parts)} / 총 계산 면수: {faces.Total}면");

            // 스케줄 비교 → 분류
            if (ctx.ScheduleMyeon.HasValue)
            {
                if (ctx.ScheduleMyeon.Value == faces.Total)
                    sectionLines.Add("→ ✅ 계산 면수와 스케줄 면수 일치");
                else
                    sectionLines.Add($"→ ❌ 계산 면수와 스케줄 면수 불일치 (스케줄: {ctx.ScheduleMyeon}면)");
            }
            else
            {
                sectionLines.Add("→ ⚠ 스케줄 면수 없음(비교 불가)");
            }

            // ELB 정리
            elbDetailLines = new List<string>();
            bool hasElbData = false;
            if (ctx.Columns.ElbCol != -1)
            {
                for (int r = ctx.StartRow + 1; r <= ctx.EndRow; r++)
                {
                    string val = ExcelUtils.Clean(df[r, ctx.Columns.ElbCol]);
                    if (!string.IsNullOrWhiteSpace(val))
                    {
                        hasElbData = true;
                        string unit = ctx.Columns.UnitSizeCol >= 0 ? ExcelUtils.Clean(df[r, ctx.Columns.UnitSizeCol]) : "";
                        string type = ctx.Columns.TypeCol >= 0 ? ExcelUtils.Clean(df[r, ctx.Columns.TypeCol]) : "";
                        elbDetailLines.Add($"→ R{r + 1}: ELB={val}"
                                           + (string.IsNullOrEmpty(unit) ? "" : $", UNIT={unit}")
                                           + (string.IsNullOrEmpty(type) ? "" : $", TYPE={type}"));
                    }
                }
            }

            if (hasElbData) classification = SectionClass.Elb;
            else if (ctx.ScheduleMyeon.HasValue && ctx.ScheduleMyeon.Value != faces.Total)
                classification = SectionClass.Mismatch;
            else
                classification = SectionClass.Pass;
        }
    }
}
