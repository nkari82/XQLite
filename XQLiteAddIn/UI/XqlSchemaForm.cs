using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlSchemaForm : Form
    {
        /// <summary>
        /// 단일 컬럼 메타 정보. 타입/널/락/기본값 등 최소 속성 제공.
        /// </summary>
        internal sealed class ColMeta
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
        internal sealed class TableMeta
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


        private static XqlSchemaForm? _inst;
        internal static void ShowSingleton()
        {
            if (_inst == null || _inst.IsDisposed) _inst = new XqlSchemaForm();
            _inst.Show(); _inst.BringToFront();
        }

        private readonly ListView lvTables = new() { View = View.Details, Dock = DockStyle.Left, Width = 260, FullRowSelect = true };
        private readonly ListView lvCols = new() { View = View.Details, Dock = DockStyle.Fill, FullRowSelect = true };
        private readonly SplitContainer sp = new() { Dock = DockStyle.Fill, Orientation = Orientation.Vertical, SplitterDistance = 260 };
        private readonly Timer auto = new() { Interval = 5000 };

        public XqlSchemaForm()
        {
            Text = "XQLite Schema";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 820; Height = 480;

            lvTables.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "Table", Width = 220 },
                new ColumnHeader { Text = "Rows?", Width = 60 }
            });
            lvCols.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "Column", Width = 220 },
                new ColumnHeader { Text = "Type", Width = 120 },
                new ColumnHeader { Text = "NotNull", Width = 80 },
                new ColumnHeader { Text = "Default", Width = 200 }
            });

            sp.Panel1.Controls.Add(lvTables);
            sp.Panel2.Controls.Add(lvCols);
            Controls.Add(sp);

            lvTables.SelectedIndexChanged += async (_, __) => await LoadColumnsAsync();
            Load += async (_, __) => await LoadTablesAsync();
            auto.Tick += async (_, __) => await LoadTablesAsync();
            auto.Start();
        }

        // ─────────────────────────────────────────────────────────────────────
        // 그래프QL → UI 바인딩
        // ─────────────────────────────────────────────────────────────────────

        private async Task LoadTablesAsync()
        {
            try
            {
                // 1차: 서버가 schema 쿼리를 제공하는 경우
                const string q1 = "query{ schema{ tables{ name rowCount columns{ name type notNull default } } } }";
                try
                {
                    // #FIXME
#if false
                    var r = await XqlGraphQLClient.QueryAsync<DbSchemaResp>(q1, null);
                    var srvTables = r.Data?.schema?.tables;
                    if (srvTables != null)
                    {
                        var items = srvTables.Select(t =>
                        {
                            var tm = new TableMeta
                            {
                                Name = t.name,
                                WorksheetName = "",      // 알 수 없으므로 비움
                                ListObjectName = t.name, // 표시는 name로
                                DisplayName = t.name,
                                Columns = [.. (t.columns ?? Array.Empty<ColDecl>())
                                    .Select(c => new ColMeta
                                    {
                                        Name = c.name,
                                        OriginalHeader = c.name,
                                        Type = c.type ?? "TEXT",
                                        Nullable = !c.notNull,
                                        Default = c.@default,
                                        IsMeta = IsMetaColumnName(c.name),
                                    })]
                            };
                            TableMeta.EnsureDefaultMetaColumns(tm);
                            return new TableItem { Meta = tm, RowCount = t.rowCount };
                        }).ToArray();

                        BindTables(items);
                        return;
                    }
#endif
                }
                catch
                {
                    // schema 쿼리가 없는 서버일 수도 있으니 fallback 진행
                }

                // 2차: 인트로스펙션으로 'rows' 필드 갖는 타입을 테이블로 추정
#if false
                const string q2 = "query{ __schema{ types{ name kind fields{ name } } } }";
                var ir = await XqlGraphQLClient.QueryAsync<IntrospectResp>(q2, null);
                var types = ir.Data?.__schema?.types ?? Array.Empty<IntrospectType>();
                var approx = types
                    .Where(t => t.fields?.Any(f => string.Equals(f.name, "rows", StringComparison.OrdinalIgnoreCase)) == true)
                    .Select(t => new TableItem
                    {
                        Meta = new TableMeta
                        {
                            Name = t.name,
                            WorksheetName = "",
                            ListObjectName = t.name,
                            DisplayName = t.name,
                            Columns = new List<ColMeta>() // 컬럼은 선택 시 rows 쿼리로 유추
                        },
                        RowCount = null
                    })
                    .ToArray();

                BindTables(approx);
#endif
            }
            catch
            {
                // 조용히 실패 (UI 깜빡임 방지)
            }
        }

        private void BindTables(TableItem[] tables)
        {
            lvTables.BeginUpdate(); lvTables.Items.Clear();
            foreach (var t in tables)
            {
                var item = new ListViewItem(new[]
                {
                    t.Meta.Name,
                    t.RowCount?.ToString() ?? ""
                })
                { Tag = t };
                lvTables.Items.Add(item);
            }
            lvTables.EndUpdate();
            lvCols.Items.Clear();
        }

        private async Task LoadColumnsAsync()
        {
            if (lvTables.SelectedItems.Count == 0) { lvCols.Items.Clear(); return; }
            var ti = (TableItem?)lvTables.SelectedItems[0].Tag;
            if (ti == null) { lvCols.Items.Clear(); return; }
            var meta = ti.Meta;

            // 이미 메타가 컬럼을 갖고 있으면 그대로 사용
            if (meta.Columns != null && meta.Columns.Count > 0)
            {
                BindCols(meta.Columns.ToArray());
                return;
            }

            // 없으면 rows(table, limit:1)로 추정
            try
            {
                // #FIXME
#if false
                const string q = "query($t:String!){ rows(table:$t, limit:1){ rows } }";
                var r = await XqlGraphQLClient.QueryAsync<RowsOnlyResp>(q, new { t = meta.Name });
                var rows = r.Data?.rows?.FirstOrDefault()?.rows ?? Array.Empty<Dictionary<string, object?>>();
                var headers = rows.SelectMany(x => x.Keys).Distinct();

                var cols = headers.Select(k => new ColMeta
                {
                    Name = k,
                    OriginalHeader = k,
                    Type = "(unknown)",
                    Nullable = true,
                    Default = null
                }).ToArray();

                BindCols(cols);
#endif
            }
            catch
            {
                lvCols.Items.Clear();
            }
        }

        private void BindCols(ColMeta[] cols)
        {
            lvCols.BeginUpdate(); lvCols.Items.Clear();
            foreach (var c in cols)
            {
                // our model: Nullable=true면 NotNull=""
                string notNull = (c.Nullable ? "" : "YES");
                lvCols.Items.Add(new ListViewItem(new[]
                {
                    c.Name,
                    c.Type ?? "",
                    notNull,
                    c.Default ?? ""
                }));
            }
            lvCols.EndUpdate();
        }

        private static bool IsMetaColumnName(string name)
            => name.Equals("id", StringComparison.OrdinalIgnoreCase)
            || name.Equals("row_version", StringComparison.OrdinalIgnoreCase)
            || name.Equals("updated_at", StringComparison.OrdinalIgnoreCase)
            || name.Equals("deleted", StringComparison.OrdinalIgnoreCase);

        // ─────────────────────────────────────────────────────────────────────
        // 내부 DTO (GraphQL 응답용)
        // ─────────────────────────────────────────────────────────────────────
        private sealed class DbSchemaResp { public DbSchema? schema { get; set; } }
        private sealed class DbSchema { public TableDecl[]? tables { get; set; } }
        private sealed class TableDecl { public string name { get; set; } = string.Empty; public int? rowCount { get; set; } public ColDecl[]? columns { get; set; } }
        private sealed class ColDecl { public string name { get; set; } = string.Empty; public string? type { get; set; } public bool notNull { get; set; } public string? @default { get; set; } }

        private sealed class IntrospectResp { public SchemaNode? __schema { get; set; } }
        private sealed class SchemaNode { public IntrospectType[]? types { get; set; } }
        private sealed class IntrospectType { public string name { get; set; } = string.Empty; public string kind { get; set; } = string.Empty; public Field[]? fields { get; set; } }
        private sealed class Field { public string name { get; set; } = string.Empty; }

        private sealed class RowsOnlyResp { public Block[]? rows { get; set; } }
        private sealed class Block { public Dictionary<string, object?>[]? rows { get; set; } }

        // ListView.Tag에 담아둘 UI용 래퍼
        private sealed class TableItem
        {
            public TableMeta Meta { get; set; } = new TableMeta();
            public int? RowCount { get; set; }
        }
    }
}
