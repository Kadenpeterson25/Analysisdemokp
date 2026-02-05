from __future__ import annotations

import argparse
import os
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from statistics import mean


@dataclass
class SheetData:
    headers: list[str]
    rows: list[list[object]]


def _excel_col_to_num(col: str) -> int:
    num = 0
    for ch in col:
        num = num * 26 + (ord(ch.upper()) - 64)
    return num


def _read_shared_strings(z: zipfile.ZipFile) -> list[str]:
    if 'xl/sharedStrings.xml' not in z.namelist():
        return []
    tree = ET.fromstring(z.read('xl/sharedStrings.xml'))
    ns = {'a': tree.tag.split('}')[0].strip('{')}
    strings = []
    for si in tree.findall('.//a:si', ns):
        texts = [t.text or '' for t in si.findall('.//a:t', ns)]
        strings.append(''.join(texts))
    return strings


def _read_date_styles(z: zipfile.ZipFile) -> set[int]:
    if 'xl/styles.xml' not in z.namelist():
        return set()
    st = ET.fromstring(z.read('xl/styles.xml'))
    ns = {'a': st.tag.split('}')[0].strip('{')}
    num_fmt_by_id: dict[int, str] = {}
    for num_fmt in st.findall('.//a:numFmts/a:numFmt', ns):
        num_fmt_by_id[int(num_fmt.attrib['numFmtId'])] = num_fmt.attrib.get('formatCode', '')

    def is_date_fmt(code: str) -> bool:
        lower = code.lower()
        return any(ch in lower for ch in ['y', 'm', 'd', 'h', 's'])

    date_numfmt_ids = {14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47}
    for fid, code in num_fmt_by_id.items():
        if is_date_fmt(code):
            date_numfmt_ids.add(fid)

    date_style_ids: set[int] = set()
    cell_xfs = st.find('.//a:cellXfs', ns)
    if cell_xfs is not None:
        for idx, xf in enumerate(cell_xfs.findall('a:xf', ns)):
            num_fmt_id = int(xf.attrib.get('numFmtId', 0))
            if num_fmt_id in date_numfmt_ids:
                date_style_ids.add(idx)

    return date_style_ids


def _read_sheet(z: zipfile.ZipFile, date_styles: set[int]) -> SheetData:
    shared_strings = _read_shared_strings(z)
    sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
    ns = {'a': sheet.tag.split('}')[0].strip('{')}
    rows: list[dict[str, object]] = []
    for row in sheet.findall('.//a:sheetData/a:row', ns):
        row_dict: dict[str, object] = {}
        for c in row.findall('a:c', ns):
            cell_ref = c.attrib.get('r')
            col = ''.join([ch for ch in cell_ref if ch.isalpha()]) if cell_ref else None
            cell_type = c.attrib.get('t')
            style_idx = int(c.attrib.get('s', 0))
            v = c.find('a:v', ns)
            if v is None:
                value = None
            else:
                raw = v.text
                if cell_type == 's':
                    value = shared_strings[int(raw)] if raw is not None else None
                elif cell_type == 'b':
                    value = raw == '1'
                else:
                    if style_idx in date_styles and raw is not None:
                        try:
                            value = datetime(1899, 12, 30) + timedelta(days=float(raw))
                        except ValueError:
                            value = raw
                    else:
                        try:
                            if raw is None:
                                value = None
                            elif '.' in raw:
                                value = float(raw)
                            else:
                                value = int(raw)
                        except ValueError:
                            value = raw
            if col:
                row_dict[col] = value
        rows.append(row_dict)

    header_row = rows[0]
    sorted_cols = sorted(header_row.keys(), key=_excel_col_to_num)
    headers = [header_row[col] for col in sorted_cols]
    data_rows: list[list[object]] = []
    for row in rows[1:]:
        data_rows.append([row.get(col) for col in sorted_cols])
    return SheetData(headers=headers, rows=data_rows)


def _describe_columns(sheet: SheetData) -> dict[str, dict[str, object]]:
    col_stats: dict[str, dict[str, object]] = {}
    for idx, name in enumerate(sheet.headers):
        values = [row[idx] for row in sheet.rows]
        non_null = [v for v in values if v not in (None, '')]
        null_count = len(values) - len(non_null)
        unique_count = len(set(non_null))
        sample_values = non_null[:5]
        col_stats[name] = {
            'total': len(values),
            'non_null': len(non_null),
            'nulls': null_count,
            'unique': unique_count,
            'sample_values': sample_values,
        }
    return col_stats


def _numeric_stats(sheet: SheetData) -> dict[str, dict[str, float]]:
    numeric_columns: dict[str, dict[str, float]] = {}
    for idx, name in enumerate(sheet.headers):
        values = [row[idx] for row in sheet.rows]
        numeric_vals = [v for v in values if isinstance(v, (int, float)) and not isinstance(v, bool)]
        if len(numeric_vals) < 10:
            continue
        numeric_columns[name] = {
            'count': len(numeric_vals),
            'min': float(min(numeric_vals)),
            'max': float(max(numeric_vals)),
            'mean': float(mean(numeric_vals)),
            'sum': float(sum(numeric_vals)),
        }
    return numeric_columns


def _date_range(values: list[object]) -> tuple[datetime | None, datetime | None]:
    date_values = [v for v in values if isinstance(v, datetime)]
    if not date_values:
        return None, None
    return min(date_values), max(date_values)


def _write_svg_bar_chart(path: str, title: str, labels: list[str], values: list[float]) -> None:
    width = 900
    height = 500
    margin = 60
    bar_gap = 10
    if not values:
        return
    max_value = max(values)
    bar_width = (width - 2 * margin - bar_gap * (len(values) - 1)) / len(values)
    svg_lines = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        '<style>text{font-family:Arial, sans-serif; font-size:12px;}</style>',
        f'<text x="{width/2}" y="30" text-anchor="middle" font-size="18">{title}</text>',
    ]
    for i, (label, value) in enumerate(zip(labels, values)):
        bar_height = 0 if max_value == 0 else (value / max_value) * (height - 2 * margin)
        x = margin + i * (bar_width + bar_gap)
        y = height - margin - bar_height
        svg_lines.append(f'<rect x="{x}" y="{y}" width="{bar_width}" height="{bar_height}" fill="#4C78A8" />')
        svg_lines.append(f'<text x="{x + bar_width/2}" y="{height - margin + 15}" text-anchor="middle">{label}</text>')
        svg_lines.append(
            f'<text x="{x + bar_width/2}" y="{y - 5}" text-anchor="middle">{value:,.0f}</text>'
        )
    svg_lines.append('</svg>')
    with open(path, 'w', encoding='utf-8') as f:
        f.write('\n'.join(svg_lines))

def _write_column_profile_csv(path: str, headers: list[str], col_stats: dict[str, dict[str, object]]) -> None:
    with open(path, 'w', encoding='utf-8') as f:
        f.write('column,total,non_null,nulls,unique,sample_values\n')
        for name in headers:
            stats = col_stats[name]
            samples = '|'.join(str(v) for v in stats['sample_values'])
            f.write(
                f"{name},{stats['total']},{stats['non_null']},{stats['nulls']},{stats['unique']},{samples}\n"
            )


def _build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description='Generate JE analysis outputs from an XLSX file.')
    parser.add_argument('--input', default='je_samples.xlsx', help='Path to the JE XLSX file.')
    parser.add_argument('--output', default='analysis_outputs', help='Directory for output artifacts.')
    parser.add_argument('--max-bars', type=int, default=10, help='Max bars in top-N charts.')
    return parser


def main() -> None:
    parser = _build_argument_parser()
    args = parser.parse_args()

    input_path = args.input
    output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    with zipfile.ZipFile(input_path) as z:
        date_styles = _read_date_styles(z)
        sheet = _read_sheet(z, date_styles)

    col_stats = _describe_columns(sheet)
    numeric_stats = _numeric_stats(sheet)

    headers = sheet.headers
    data_rows = sheet.rows

    def get_column(name: str) -> list[object]:
        idx = headers.index(name)
        return [row[idx] for row in data_rows]

    effective_dates = get_column('EffectiveDate') if 'EffectiveDate' in headers else []
    entry_dates = get_column('EntryDate') if 'EntryDate' in headers else []
    effective_min, effective_max = _date_range(effective_dates)
    entry_min, entry_max = _date_range(entry_dates)

    analysis_path = os.path.join(output_dir, 'je_samples_analysis.txt')
    with open(analysis_path, 'w', encoding='utf-8') as f:
        f.write('JE Samples Basic Analysis\n')
        f.write('=========================\n\n')
        f.write(f'Total rows (excluding header): {len(data_rows):,}\n')
        f.write(f'Total columns: {len(headers)}\n\n')
        f.write('Date Ranges\n')
        f.write('-----------\n')
        if effective_min and effective_max:
            f.write(f'EffectiveDate: {effective_min.date()} to {effective_max.date()}\n')
        if entry_min and entry_max:
            f.write(f'EntryDate: {entry_min.date()} to {entry_max.date()}\n')
        f.write('\n')

        f.write('Column Summary\n')
        f.write('--------------\n')
        for name in headers:
            stats = col_stats[name]
            f.write(
                f"{name}: non-null={stats['non_null']:,}, nulls={stats['nulls']:,}, unique={stats['unique']:,}"
            )
            if stats['sample_values']:
                sample = ', '.join(str(v) for v in stats['sample_values'])
                f.write(f" | samples: {sample}")
            f.write('\n')
        f.write('\n')

        f.write('Numeric Column Statistics\n')
        f.write('-------------------------\n')
        for name, stats in numeric_stats.items():
            f.write(
                f"{name}: count={stats['count']:,}, min={stats['min']:,.2f}, max={stats['max']:,.2f}, "
                f"mean={stats['mean']:,.2f}, sum={stats['sum']:,.2f}\n"
            )

    _write_column_profile_csv(os.path.join(output_dir, 'column_profile.csv'), headers, col_stats)

    # charts
    if 'AccountType' in headers:
        account_vals = [v for v in get_column('AccountType') if v not in (None, '')]
        account_counts = Counter(account_vals).most_common(args.max_bars)
        labels, values = zip(*account_counts) if account_counts else ([], [])
        _write_svg_bar_chart(
            os.path.join(output_dir, 'account_type_counts.svg'),
            'Top Account Types (Count)',
            [str(l) for l in labels],
            list(values),
        )

    if 'BusinessUnit' in headers:
        unit_vals = [v for v in get_column('BusinessUnit') if v not in (None, '')]
        unit_counts = Counter(unit_vals).most_common(args.max_bars)
        labels, values = zip(*unit_counts) if unit_counts else ([], [])
        _write_svg_bar_chart(
            os.path.join(output_dir, 'business_unit_counts.svg'),
            'Top Business Units (Count)',
            [str(l) for l in labels],
            list(values),
        )

    if 'EffectiveDate' in headers and 'Amount' in headers:
        date_vals = get_column('EffectiveDate')
        amount_vals = get_column('Amount')
        monthly_totals: dict[str, float] = defaultdict(float)
        for date_val, amount in zip(date_vals, amount_vals):
            if isinstance(date_val, datetime) and isinstance(amount, (int, float)):
                key = date_val.strftime('%Y-%m')
                monthly_totals[key] += abs(float(amount))
        sorted_months = sorted(monthly_totals.items())
        if sorted_months:
            labels = [m for m, _ in sorted_months[:12]]
            values = [v for _, v in sorted_months[:12]]
            _write_svg_bar_chart(
                os.path.join(output_dir, 'monthly_absolute_amounts.svg'),
                'Monthly Absolute Amount Totals (First 12 Months)',
                labels,
                values,
            )


if __name__ == '__main__':
    main()
