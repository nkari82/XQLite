// XqlMetaRegistry.cs
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace XQLite.AddIn
{
    /// <summary>
    /// 시트(=테이블) 메타와 컬럼 타입/제약을 보관하고, 셀 값에 대한 유형/제약 검증을 수행한다.
    /// - 파싱은 안전(fail-safe)하게 동작하며 예외를 던지지 않는다.
    /// - 문자열은 NFC 정규화한다(스키마/데이터 모두 UTF-8 + NFC 권장).
    /// - 지원 타입: Int, Real, Text, Bool, Json, Date
    /// - 추가 제약: NotNull, Min/Max(수치), Regex(텍스트), CustomCheck(호출자 주입)
    /// </summary>
    internal sealed class XqlMetaRegistry
    {
        private readonly Dictionary<string, SheetMeta> _sheets = new(StringComparer.Ordinal);

        // ========= 시트 등록/조회 =========

        public void RegisterSheet(string sheet, SheetMeta meta)
        {
            if (string.IsNullOrWhiteSpace(sheet)) return;
            _sheets[sheet] = meta ?? throw new ArgumentNullException(nameof(meta));
        }

        public bool TryGetSheet(string sheet, out SheetMeta meta) =>
            _sheets.TryGetValue(sheet, out meta!);

        /// <summary>
        /// 리본/메뉴에서 툴팁을 뿌리기 위한 "컬럼명 → 요약 문자열" 맵 생성.
        /// </summary>
        public IReadOnlyDictionary<string, string> BuildTooltipsForSheet(string sheet)
        {
            var dict = new Dictionary<string, string>(StringComparer.Ordinal);
            if (!_sheets.TryGetValue(sheet, out var sm)) return dict;

            foreach (var kv in sm.Columns)
            {
                dict[kv.Key] = kv.Value.ToTooltip();
            }
            return dict;
        }

        // ========= 검증 엔진 =========

        /// <summary>
        /// 셀 값 검증(형/제약). 존재하지 않는 시트/컬럼이면 통과(보수적으로 허용).
        /// </summary>
        public ValidationResult ValidateCell(string sheet, string col, object? value)
        {
            if (!_sheets.TryGetValue(sheet, out var sm)) return ValidationResult.Ok();
            if (!sm.Columns.TryGetValue(col, out var ct)) return ValidationResult.Ok();

            // NotNull
            if (!ct.Nullable && IsNullish(value))
                return ValidationResult.Fail(ErrCode.E_NULL_NOT_ALLOWED, "Null/empty is not allowed.");

            // 타입별 검증
            switch (ct.Kind)
            {
                case ColumnKind.Int:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                        if (!TryToInt64(value, out var iv))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect INT.");
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.
                        if (ct.Min.HasValue && iv < (long)ct.Min.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"INT < Min({ct.Min})");
                        if (ct.Max.HasValue && iv > (long)ct.Max.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"INT > Max({ct.Max})");
                        break;
                    }
                case ColumnKind.Real:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
                        if (!TryToDouble(value, out var dv))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect REAL.");
                        if (ct.Min.HasValue && dv < ct.Min.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"REAL < Min({ct.Min})");
                        if (ct.Max.HasValue && dv > ct.Max.Value)
                            return ValidationResult.Fail(ErrCode.E_RANGE, $"REAL > Max({ct.Max})");
                        break;
                    }
                case ColumnKind.Bool:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
                        if (!TryToBool(value, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect BOOL.");
                        break;
                    }
                case ColumnKind.Text:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
                        var s = NormalizeToString(value);
                        if (ct.Regex != null && !ct.Regex.IsMatch(s))
                            return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, "TEXT regex mismatch.");
                        break;
                    }
                case ColumnKind.Json:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
#pragma warning disable CS8604 // 가능한 null 참조 인수입니다.
                        var s = NormalizeToString(value);
#pragma warning restore CS8604 // 가능한 null 참조 인수입니다.
                        try
                        {
                            var _ = JToken.Parse(s);
                        }
                        catch (Exception ex)
                        {
                            return ValidationResult.Fail(ErrCode.E_JSON_PARSE, $"JSON parse error: {ex.Message}");
                        }
                        break;
                    }
                case ColumnKind.Date:
                    {
                        if (IsNullish(value)) return ValidationResult.Ok();
                        if (!TryToDate(value, out _))
                            return ValidationResult.Fail(ErrCode.E_TYPE_MISMATCH, "Expect DATE.");
                        break;
                    }
                default:
                    return ValidationResult.Fail(ErrCode.E_UNSUPPORTED, $"Unsupported type: {ct.Kind}");
            }

            // 사용자 정의 체크
            if (ct.CustomCheck != null)
            {
                try
                {
                    if (!ct.CustomCheck(value))
                        return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, ct.CustomCheckDescription ?? "Custom check failed.");
                }
                catch (Exception ex)
                {
                    return ValidationResult.Fail(ErrCode.E_CHECK_FAIL, $"Custom check error: {ex.Message}");
                }
            }

            return ValidationResult.Ok();
        }

        // ========= 타입 파싱(문자열 → ColumnType) =========

        /// <summary>
        /// 예: "int notnull min=0 max=999", "real", "text regex=^[A-Z]+$", "json", "bool", "date"
        /// - 공백 구분, 대소문자 무시. 미지원 표기(예: title[..32])는 실패를 반환.
        /// - 파싱 실패 시 false와 에러 메시지를 반환(예외 없음).
        /// </summary>
        public static bool TryParseColumnType(string spec, out ColumnType type, out string? error)
        {
            type = new ColumnType { Kind = ColumnKind.Text, Nullable = true }; // 기본값: Text nullable
            error = null;

            if (string.IsNullOrWhiteSpace(spec)) return true;

            // 미지원 문법 방지: "name[..]" 류 패턴은 명시적으로 거부
            if (spec.Contains("[") || spec.Contains("]"))
            {
                error = "Bracket-length style is not supported (e.g., title[..32]).";
                return false;
            }

            var tokens = spec.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length == 0) return true;

            // 첫 토큰: 타입
            var t0 = tokens[0].Trim().ToLowerInvariant();
            switch (t0)
            {
                case "int":
                case "integer":
                    type.Kind = ColumnKind.Int; break;
                case "real":
                case "float":
                case "double":
                    type.Kind = ColumnKind.Real; break;
                case "text":
                case "string":
                case "str":
                    type.Kind = ColumnKind.Text; break;
                case "bool":
                case "boolean":
                    type.Kind = ColumnKind.Bool; break;
                case "json":
                    type.Kind = ColumnKind.Json; break;
                case "date":
                case "datetime":
                    type.Kind = ColumnKind.Date; break;
                default:
                    error = $"Unknown type: {t0}";
                    return false;
            }

            // 나머지 토큰: 옵션
            for (int i = 1; i < tokens.Length; i++)
            {
                var tok = tokens[i].Trim();
                if (tok.Equals("notnull", StringComparison.OrdinalIgnoreCase))
                {
                    type.Nullable = false;
                    continue;
                }

                // key=value
                var eq = tok.IndexOf('=');
                if (eq <= 0 || eq >= tok.Length - 1)
                {
                    error = $"Invalid token: {tok}";
                    return false;
                }

                var key = tok.Substring(0, eq).Trim().ToLowerInvariant();
                var val = tok.Substring(eq + 1).Trim();

                switch (key)
                {
                    case "min":
                        if (type.Kind == ColumnKind.Int)
                        {
                            if (!long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l))
                            { error = "Invalid min for INT."; return false; }
                            type.Min = l;
                        }
                        else if (type.Kind == ColumnKind.Real)
                        {
                            if (!double.TryParse(val, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d))
                            { error = "Invalid min for REAL."; return false; }
                            type.Min = d;
                        }
                        else { error = "min is valid only for INT/REAL."; return false; }
                        break;

                    case "max":
                        if (type.Kind == ColumnKind.Int)
                        {
                            if (!long.TryParse(val, NumberStyles.Integer, CultureInfo.InvariantCulture, out var l))
                            { error = "Invalid max for INT."; return false; }
                            type.Max = l;
                        }
                        else if (type.Kind == ColumnKind.Real)
                        {
                            if (!double.TryParse(val, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var d))
                            { error = "Invalid max for REAL."; return false; }
                            type.Max = d;
                        }
                        else { error = "max is valid only for INT/REAL."; return false; }
                        break;

                    case "regex":
                        if (type.Kind != ColumnKind.Text) { error = "regex is valid only for TEXT."; return false; }
                        try
                        {
                            type.Regex = new Regex(val, RegexOptions.Compiled);
                        }
                        catch (Exception ex)
                        {
                            error = $"Invalid regex: {ex.Message}";
                            return false;
                        }
                        break;

                    default:
                        error = $"Unknown option: {key}";
                        return false;
                }
            }

            return true;
        }

        // ========= 유틸 =========

        private static bool IsNullish(object? v)
        {
            if (v is null) return true;
            if (v is string s) return string.IsNullOrWhiteSpace(s);
            return false;
        }

        private static bool TryToInt64(object v, out long value)
        {
            try
            {
                switch (v)
                {
                    case sbyte sb: value = sb; return true;
                    case byte b: value = b; return true;
                    case short s: value = s; return true;
                    case ushort us: value = us; return true;
                    case int i: value = i; return true;
                    case uint ui: value = ui; return true;
                    case long l: value = l; return true;
                    case ulong ul:
                        if (ul <= long.MaxValue) { value = (long)ul; return true; }
                        break;
                    case float f:
                        value = (long)f; return true;
                    case double d:
                        // Excel의 정수도 double로 들어올 수 있음
                        if (Math.Abs(d % 1.0) < 1e-9) { value = (long)d; return true; }
                        break;
                    case decimal m:
                        if (m == decimal.Truncate(m)) { value = (long)m; return true; }
                        break;
                    case string s:
                        if (long.TryParse(s.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out var li))
                        { value = li; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = 0;
            return false;
        }

        private static bool TryToDouble(object v, out double value)
        {
            try
            {
                switch (v)
                {
                    case sbyte sb: value = sb; return true;
                    case byte b: value = b; return true;
                    case short s: value = s; return true;
                    case ushort us: value = us; return true;
                    case int i: value = i; return true;
                    case uint ui: value = ui; return true;
                    case long l: value = l; return true;
                    case ulong ul: value = ul; return true;
                    case float f: value = f; return true;
                    case double d: value = d; return true;
                    case decimal m: value = (double)m; return true;
                    case string s:
                        if (double.TryParse(s.Trim(), NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var dd))
                        { value = dd; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = 0;
            return false;
        }

        private static bool TryToBool(object v, out bool value)
        {
            try
            {
                switch (v)
                {
                    case bool b: value = b; return true;
                    case sbyte sb: value = sb != 0; return true;
                    case byte by: value = by != 0; return true;
                    case short s: value = s != 0; return true;
                    case ushort us: value = us != 0; return true;
                    case int i: value = i != 0; return true;
                    case uint ui: value = ui != 0; return true;
                    case long l: value = l != 0; return true;
                    case ulong ul: value = ul != 0; return true;
                    case string str:
                        var t = str.Trim().ToLowerInvariant();
                        if (t is "1" or "true" or "t" or "y" or "yes") { value = true; return true; }
                        if (t is "0" or "false" or "f" or "n" or "no") { value = false; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = false;
            return false;
        }

        private static bool TryToDate(object v, out DateTime value)
        {
            try
            {
                switch (v)
                {
                    case DateTime dt: value = dt; return true;
                    case double oa:   // Excel OADate
                        value = DateTime.FromOADate(oa);
                        return true;
                    case string s:
                        if (DateTime.TryParse(s.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out var d))
                        { value = d; return true; }
                        break;
                }
            }
            catch { /* ignore */ }

            value = default;
            return false;
        }

        /// <summary>
        /// 문자열로 변환하며 NFC 정규화.
        /// </summary>
        private static string NormalizeToString(object v)
        {
            var s = v switch
            {
                string ss => ss,
                _ => Convert.ToString(v, CultureInfo.InvariantCulture) ?? string.Empty,
            };
            return s.Normalize(NormalizationForm.FormC);
        }
    }

    // ========= 모델 =========

    internal sealed class SheetMeta
    {
        public string TableName { get; init; } = "";
        public string KeyColumn { get; init; } = "id";

        /// <summary>컬럼명 → 타입/제약</summary>
        public Dictionary<string, ColumnType> Columns { get; } = new(StringComparer.Ordinal);

        public void SetColumn(string name, ColumnType type)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            Columns[name] = type;
        }
    }

    internal enum ColumnKind
    {
        Int,
        Real,
        Text,
        Bool,
        Json,
        Date,
    }

    internal sealed class ColumnType
    {
        public ColumnKind Kind;
        public bool Nullable = true;

        /// <summary>수치형에만 사용(Int=long, Real=double)</summary>
        public double? Min;
        public double? Max;

        /// <summary>Text 전용</summary>
        public Regex? Regex;

        /// <summary>사용자 정의 체크(선택)</summary>
        public Func<object?, bool>? CustomCheck;
        public string? CustomCheckDescription;

        /// <summary>툴팁용 요약 문자열</summary>
        public string ToTooltip()
        {
            var sb = new StringBuilder();
            sb.Append(Kind.ToString().ToUpperInvariant());
            if (!Nullable) sb.Append(" NOTNULL");

            if (Kind is ColumnKind.Int or ColumnKind.Real)
            {
                if (Min.HasValue) sb.Append($" MIN={Min}");
                if (Max.HasValue) sb.Append($" MAX={Max}");
            }
            if (Kind == ColumnKind.Text && Regex != null)
            {
                sb.Append(" REGEX=");
                sb.Append(Regex.ToString());
            }
            if (!string.IsNullOrWhiteSpace(CustomCheckDescription))
            {
                sb.Append(" CHECK=");
                sb.Append(CustomCheckDescription);
            }
            return sb.ToString();
        }
    }

    // ========= 검증 결과 =========

    internal enum ErrCode
    {
        None = 0,
        E_TYPE_MISMATCH,
        E_RANGE,
        E_CHECK_FAIL,
        E_JSON_PARSE,
        E_NULL_NOT_ALLOWED,
        E_UNSUPPORTED,
    }

    internal readonly struct ValidationResult
    {
        public bool IsOk { get; }
        public ErrCode Code { get; }
        public string Message { get; }

        private ValidationResult(bool ok, ErrCode code, string msg)
        {
            IsOk = ok; Code = code; Message = msg;
        }

        public static ValidationResult Ok() => new(true, ErrCode.None, "");
        public static ValidationResult Fail(ErrCode code, string msg) => new(false, code, msg);
    }
}
