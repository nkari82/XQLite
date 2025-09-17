// TableMeta.cs
using System;
using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    /// <summary>
    /// 단일 컬럼 메타 정보. 타입/널/락/기본값 등 최소 속성 제공.
    /// </summary>
    public sealed class ColMeta
    {
        [JsonProperty("name")] public string Name { get; set; } = "";          // DB 컬럼명 (정규화)
        [JsonProperty("header")] public string OriginalHeader { get; set; } = ""; // 엑셀 헤더 원문
        [JsonProperty("type")] public string Type { get; set; } = "TEXT";       // SQLite 타입(가벼운 문자열)
        [JsonProperty("nullable")] public bool Nullable { get; set; } = true;
        [JsonProperty("locked")] public bool Locked { get; set; } = false;        // 컬럼 락(편집 금지)
        [JsonProperty("is_meta")] public bool IsMeta { get; set; } = false;        // 메타 컬럼 여부(id/row_version/…)
        [JsonProperty("default")] public string? Default { get; set; }             // DEFAULT expr
        [JsonProperty("check")] public string? Check { get; set; }               // CHECK expr
        [JsonProperty("ref_table")] public string? RefTable { get; set; }            // FK: 참조 테이블
        [JsonProperty("ref_column")] public string? RefColumn { get; set; }           // FK: 참조 컬럼
        [JsonProperty("ordinal")] public int Ordinal { get; set; }                 // 1-base 열 순서
        [JsonProperty("max_len")] public int? MaxLen { get; set; }                 // TEXT 길이 힌트(검증용)

        public override string ToString() => $"{Name} {Type}{(Nullable ? "" : " NOT NULL")}{(Default != null ? " DEFAULT " + Default : "")}";
    }

    /// <summary>
    /// 테이블(엑셀 ListObject) 메타. 엑셀에서 유추하거나 외부 설정에서 로드.
    /// </summary>
    public sealed class TableMeta
    {
        [JsonProperty("name")] public string Name { get; set; } = "";           // DB 테이블명: "Sheet.ListObject"
        [JsonProperty("display_name")] public string? DisplayName { get; set; }         // UI 표시명
        [JsonProperty("worksheet")] public string WorksheetName { get; set; } = "";
        [JsonProperty("list_object")] public string ListObjectName { get; set; } = "";
        [JsonProperty("primary_key")] public string? PrimaryKey { get; set; }          // ex) "id"
        [JsonProperty("unique_key")] public string? UniqueKey { get; set; }
        [JsonProperty("columns")] public List<ColMeta> Columns { get; set; } = new();

        [JsonIgnore]
        public Dictionary<string, ColMeta> ByName =>
            Columns.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);

        public bool TryGetColumn(string name, out ColMeta col) =>
            ByName.TryGetValue(name, out col!);

        /// <summary>
        /// 엑셀 테이블에서 헤더를 읽어 TableMeta를 생성. 메타 컬럼(id/row_version/updated_at/deleted) 자동 보강.
        /// </summary>
        public static TableMeta FromListObject(Excel.Worksheet ws, Excel.ListObject lo, string? explicitName = null)
        {
            if (ws == null || lo == null) throw new ArgumentNullException();
            var header = lo.HeaderRowRange ?? throw new InvalidOperationException("Header row not found.");
            int colCount = header.Columns.Count;

            string tableName = explicitName ?? $"{ws.Name}.{lo.Name}";
            var tm = new TableMeta
            {
                Name = tableName,
                WorksheetName = ws.Name,
                ListObjectName = lo.Name,
                DisplayName = lo.Name
            };

            // 헤더 읽기
            var headArr = (object[,])header.Value2;
            for (int c = 1; c <= colCount; c++)
            {
                string raw = Convert.ToString(headArr[1, c]) ?? $"C{c}";
                string normalized = NormalizeToDbIdent(raw);
                tm.Columns.Add(new ColMeta
                {
                    Name = normalized,
                    OriginalHeader = raw,
                    Type = GuessTypeByHeader(raw),
                    Nullable = true,
                    Locked = false,
                    IsMeta = IsMetaColumnName(normalized),
                    Ordinal = c
                });
            }

            // 기본 메타 컬럼이 없으면 추가
            EnsureDefaultMetaColumns(tm);

            // PK 기본값
            if (string.IsNullOrWhiteSpace(tm.PrimaryKey))
                tm.PrimaryKey = tm.Columns.FirstOrDefault(c => c.Name.Equals("id", StringComparison.OrdinalIgnoreCase))?.Name;

            return tm;
        }

        /// <summary>
        /// 외부 설정(JSON 등)에서 로드했을 때도, 메타 컬럼 보강을 한 번 더 적용.
        /// </summary>
        public static void EnsureDefaultMetaColumns(TableMeta tm)
        {
            // id
            if (!tm.Columns.Any(c => c.Name.Equals("id", StringComparison.OrdinalIgnoreCase)))
                tm.Columns.Insert(0, new ColMeta { Name = "id", OriginalHeader = "id", Type = "INTEGER", Nullable = false, IsMeta = true, Ordinal = 0, Default = null });

            // row_version
            if (!tm.Columns.Any(c => c.Name.Equals("row_version", StringComparison.OrdinalIgnoreCase)))
                tm.Columns.Add(new ColMeta { Name = "row_version", OriginalHeader = "row_version", Type = "INTEGER", Nullable = false, IsMeta = true });

            // updated_at
            if (!tm.Columns.Any(c => c.Name.Equals("updated_at", StringComparison.OrdinalIgnoreCase)))
                tm.Columns.Add(new ColMeta { Name = "updated_at", OriginalHeader = "updated_at", Type = "TEXT", Nullable = false, IsMeta = true, Default = "CURRENT_TIMESTAMP" });

            // deleted (soft-delete 플래그)
            if (!tm.Columns.Any(c => c.Name.Equals("deleted", StringComparison.OrdinalIgnoreCase)))
                tm.Columns.Add(new ColMeta { Name = "deleted", OriginalHeader = "deleted", Type = "INTEGER", Nullable = false, IsMeta = true, Default = "0" });
        }

        private static string NormalizeToDbIdent(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "col";
            // 공백/특수문자 → '_' 대체, 앞뒤 트림, 연속 '_' 축약
            var sb = new System.Text.StringBuilder();
            foreach (var ch in s.Trim())
            {
                if (char.IsLetterOrDigit(ch)) sb.Append(char.ToLowerInvariant(ch));
                else sb.Append('_');
            }
            var ident = sb.ToString().Trim('_');
            while (ident.Contains("__")) ident = ident.Replace("__", "_");
            if (string.IsNullOrEmpty(ident)) ident = "col";
            // 예약 메타 컬럼과 충돌 시 접미사
            if (IsMetaColumnName(ident) && !s.Equals(ident, StringComparison.OrdinalIgnoreCase))
                ident += "_1";
            return ident;
        }

        private static bool IsMetaColumnName(string name)
        {
            return name.Equals("id", StringComparison.OrdinalIgnoreCase)
                || name.Equals("row_version", StringComparison.OrdinalIgnoreCase)
                || name.Equals("updated_at", StringComparison.OrdinalIgnoreCase)
                || name.Equals("deleted", StringComparison.OrdinalIgnoreCase);
        }

        private static string GuessTypeByHeader(string header)
        {
            // 아주 간단한 휴리스틱(원하면 강화 가능)
            var h = header.ToLowerInvariant();
            if (h.Contains("id")) return "INTEGER";
            if (h.Contains("date") || h.Contains("time") || h.Contains("_at")) return "TEXT";
            if (h.Contains("rate") || h.Contains("price") || h.Contains("amount")) return "REAL";
            if (h.Contains("count") || h.Contains("num") || h.Contains("qty")) return "INTEGER";
            return "TEXT";
        }
    }

    /// <summary>
    /// 간단한 런타임 레지스트리: 테이블 메타 보관/조회.
    /// 외부 JSON 로드나 엑셀에서 유추한 메타를 등록해두고 재사용.
    /// </summary>
    public static class XqlSchemaRegistry
    {
        private static readonly Dictionary<string, TableMeta> _byName = new(StringComparer.OrdinalIgnoreCase);

        public static void Register(TableMeta tm)
        {
            if (tm == null || string.IsNullOrWhiteSpace(tm.Name))
                throw new ArgumentException("invalid table meta");
            TableMeta.EnsureDefaultMetaColumns(tm);
            _byName[tm.Name] = tm;
        }

        public static bool TryGet(string tableName, out TableMeta tm) => _byName.TryGetValue(tableName, out tm!);

        public static TableMeta GetOrAddFromListObject(Excel.Worksheet ws, Excel.ListObject lo)
        {
            var name = $"{ws.Name}.{lo.Name}";
            if (_byName.TryGetValue(name, out var found)) return found;
            var tm = TableMeta.FromListObject(ws, lo, name);
            _byName[name] = tm;
            return tm;
        }

        public static IEnumerable<TableMeta> All() => _byName.Values;
        internal static void Clear() => _byName.Clear();
    }
}
