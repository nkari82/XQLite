// 교체본: XqlSchemaForm.cs
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace XQLite.AddIn
{
    public sealed class XqlSchemaForm : Form
    {
        internal sealed class ColMeta
        {
            [JsonProperty("name")] public string Name { get; set; } = "";
            [JsonProperty("header")] public string OriginalHeader { get; set; } = "";
            [JsonProperty("type")] public string Type { get; set; } = "TEXT";
            [JsonProperty("nullable")] public bool Nullable { get; set; } = true;
            [JsonProperty("locked")] public bool Locked { get; set; } = false;
            [JsonProperty("is_meta")] public bool IsMeta { get; set; } = false;
            [JsonProperty("default")] public string? Default { get; set; }
            [JsonProperty("check")] public string? Check { get; set; }
            [JsonProperty("ref_table")] public string? RefTable { get; set; }
            [JsonProperty("ref_column")] public string? RefColumn { get; set; }
            [JsonProperty("ordinal")] public int Ordinal { get; set; }
            [JsonProperty("max_len")] public int? MaxLen { get; set; }

            public override string ToString()
                => $"{Name} {Type}{(Nullable ? "" : " NOT NULL")}{(Default != null ? " DEFAULT " + Default : "")}";
        }

        internal sealed class TableMeta
        {
            [JsonProperty("name")] public string Name { get; set; } = "";
            [JsonProperty("display_name")] public string? DisplayName { get; set; }
            [JsonProperty("worksheet")] public string WorksheetName { get; set; } = "";
            [JsonProperty("list_object")] public string ListObjectName { get; set; } = "";
            [JsonProperty("primary_key")] public string? PrimaryKey { get; set; }
            [JsonProperty("unique_key")] public string? UniqueKey { get; set; }
            [JsonProperty("columns")] public List<ColMeta> Columns { get; set; } = new();

            [JsonIgnore]
            public Dictionary<string, ColMeta> ByName =>
                Columns.ToDictionary(c => c.Name, c => c, StringComparer.OrdinalIgnoreCase);

            public bool TryGetColumn(string name, out ColMeta col)
                => ByName.TryGetValue(name, out col!);

            public static void EnsureDefaultMetaColumns(TableMeta tm)
            {
                if (!tm.Columns.Any(c => c.Name.Equals("id", StringComparison.OrdinalIgnoreCase)))
                    tm.Columns.Insert(0, new ColMeta { Name = "id", OriginalHeader = "id", Type = "INTEGER", Nullable = false, IsMeta = true, Ordinal = 0 });

                if (!tm.Columns.Any(c => c.Name.Equals("row_version", StringComparison.OrdinalIgnoreCase)))
                    tm.Columns.Add(new ColMeta { Name = "row_version", OriginalHeader = "row_version", Type = "INTEGER", Nullable = false, IsMeta = true });

                if (!tm.Columns.Any(c => c.Name.Equals("updated_at", StringComparison.OrdinalIgnoreCase)))
                    tm.Columns.Add(new ColMeta { Name = "updated_at", OriginalHeader = "updated_at", Type = "TEXT", Nullable = false, IsMeta = true, Default = "CURRENT_TIMESTAMP" });

                if (!tm.Columns.Any(c => c.Name.Equals("deleted", StringComparison.OrdinalIgnoreCase)))
                    tm.Columns.Add(new ColMeta { Name = "deleted", OriginalHeader = "deleted", Type = "INTEGER", Nullable = false, IsMeta = true, Default = "0" });
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

        private TableItem[] _tables = Array.Empty<TableItem>();

        public XqlSchemaForm()
        {
            Text = "XQLite Schema";
            StartPosition = FormStartPosition.CenterScreen;
            Width = 820; Height = 480;

            lvTables.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "Table", Width = 220 },
                new ColumnHeader { Text = "Rows",  Width = 60 }
            });
            lvCols.Columns.AddRange(new[]
            {
                new ColumnHeader { Text = "Column",  Width = 220 },
                new ColumnHeader { Text = "Type",    Width = 120 },
                new ColumnHeader { Text = "NotNull", Width = 80 },
                new ColumnHeader { Text = "Default", Width = 200 }
            });

            sp.Panel1.Controls.Add(lvTables);
            sp.Panel2.Controls.Add(lvCols);
            Controls.Add(sp);

            lvTables.SelectedIndexChanged += async (_, __) => await LoadColumns();
            Load += async (_, __) => await LoadTables();
            auto.Tick += async (_, __) => await LoadTables();
            auto.Start();
        }

        private async Task LoadTables()
        {
            try
            {
                if (XqlAddIn.Backend == null) return;

                // 1) 서버 메타 조회
                var meta = await XqlAddIn.Backend.TryFetchServerMeta().ConfigureAwait(true);
                if (meta is not JObject m) return;

                var tables = new List<TableItem>();
                var arr = m["tables"] as JArray ?? new JArray();

                foreach (var t in arr.OfType<JObject>())
                {
                    var tname = t["name"]?.ToString() ?? "";
                    if (string.IsNullOrWhiteSpace(tname)) continue;

                    var tm = new TableMeta
                    {
                        Name = tname,
                        WorksheetName = "",
                        ListObjectName = tname,
                        DisplayName = tname,
                        Columns = new List<ColMeta>()
                    };

                    var cols = t["cols"] as JArray ?? new JArray();
                    int ord = 1;
                    foreach (var c in cols.OfType<JObject>())
                    {
                        var cname = c["name"]?.ToString() ?? "";
                        if (string.IsNullOrWhiteSpace(cname)) continue;

                        var kind = c["kind"]?.ToString() ?? "TEXT";
                        var notnull = (bool?)c["notnull"] ?? false;

                        tm.Columns.Add(new ColMeta
                        {
                            Name = cname,
                            OriginalHeader = cname,
                            Type = kind.ToUpperInvariant(),
                            Nullable = !notnull,
                            IsMeta = IsMetaColumnName(cname),
                            Ordinal = ord++
                        });
                    }

                    TableMeta.EnsureDefaultMetaColumns(tm);
                    tables.Add(new TableItem { Meta = tm, RowCount = null });
                }

                _tables = tables.ToArray();

                // 2) row count 보강 (PullRows(0) 스냅샷으로 라이브 행수 계산)
                await EnrichRowCounts(_tables).ConfigureAwait(true);

                BindTables(_tables);
            }
            catch
            {
                // 조용히 무시
            }
        }

        private async Task EnrichRowCounts(TableItem[] tables)
        {
            try
            {
                if (XqlAddIn.Backend == null || tables.Length == 0) return;

                var pr = await XqlAddIn.Backend.PullRows(0).ConfigureAwait(true);
                // 테이블별 (row_key -> 마지막 deleted 플래그) 판단
                var byTable = new Dictionary<string, Dictionary<object, bool>>(StringComparer.Ordinal);
                foreach (var p in pr.Patches ?? new List<RowPatch>())
                {
                    if (!byTable.TryGetValue(p.Table, out var map))
                    {
                        map = new Dictionary<object, bool>();
                        byTable[p.Table] = map;
                    }
                    // 마지막 상태만 남긴다 (덮어쓰기)
                    map[p.RowKey] = p.Deleted;
                }

                foreach (var ti in tables)
                {
                    if (byTable.TryGetValue(ti.Meta.Name, out var map))
                        ti.RowCount = map.Values.Count(d => d == false);
                    else
                        ti.RowCount = 0;
                }
            }
            catch
            {
                // 실패 시 카운트는 그대로 둔다
            }
        }

        private void BindTables(TableItem[] tables)
        {
            lvTables.BeginUpdate();
            lvTables.Items.Clear();
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

        private async Task LoadColumns()
        {
            if (lvTables.SelectedItems.Count == 0) { lvCols.Items.Clear(); return; }
            if (lvTables.SelectedItems[0].Tag is not TableItem ti) { lvCols.Items.Clear(); return; }

            if (ti.Meta.Columns is { Count: > 0 })
            {
                BindCols(ti.Meta.Columns.ToArray());
                return;
            }

            await Task.CompletedTask;
            lvCols.Items.Clear();
        }

        private void BindCols(ColMeta[] cols)
        {
            lvCols.BeginUpdate();
            lvCols.Items.Clear();

            foreach (var c in cols)
            {
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

        private sealed class TableItem
        {
            public TableMeta Meta { get; set; } = new TableMeta();
            public int? RowCount { get; set; }
        }
    }
}
