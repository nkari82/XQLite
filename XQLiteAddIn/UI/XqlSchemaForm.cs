using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace XQLite.AddIn
{
    public sealed class XqlSchemaForm : Form
    {
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
                                Columns = (t.columns ?? Array.Empty<ColDecl>())
                                    .Select(c => new ColMeta
                                    {
                                        Name = c.name,
                                        OriginalHeader = c.name,
                                        Type = c.type ?? "TEXT",
                                        Nullable = !c.notNull,
                                        Default = c.@default,
                                        IsMeta = IsMetaColumnName(c.name),
                                    }).ToList()
                            };
                            TableMeta.EnsureDefaultMetaColumns(tm);
                            return new TableItem { Meta = tm, RowCount = t.rowCount };
                        }).ToArray();

                        BindTables(items);
                        return;
                    }
                }
                catch
                {
                    // schema 쿼리가 없는 서버일 수도 있으니 fallback 진행
                }

                // 2차: 인트로스펙션으로 'rows' 필드 갖는 타입을 테이블로 추정
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
