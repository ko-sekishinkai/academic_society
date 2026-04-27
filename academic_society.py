

# -*- coding: utf-8 -*-
# 年度×診療科 検索アプリ（HTML）を Excel から再生成
# チェックボックス入りドロップダウンで複数選択対応（安全版）


import os
import sys
import json
from datetime import datetime
import pandas as pd


# ===== 設定（必要に応じて変更） =====
EXCEL_FILE = "学会発表・論文.xlsx"            # 実ファイル名に合わせてください
HTML_FILE  = "index.html"


def main():
    try:
        print("=== 実行開始 ===")
        print(f"カレントディレクトリ: {os.getcwd()}")
        print(f"Excelファイル設定    : {EXCEL_FILE}")


        # ---- Excel存在チェック ----
        if not os.path.isfile(EXCEL_FILE):
            print("⚠ Excelファイルが見つかりません。以下をご確認ください：")
            print("  - スクリプトと同じフォルダにExcelがあるか？")
            print("  - ファイル名・拡張子が完全一致か？（例：学会発表・論文.xlsx）")
            print("  - 実行時のカレントディレクトリが正しいか？")
            sys.exit(1)


        # ===== Excel読込 =====
        xl = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")


        # ===== カテゴリ別レコード格納 =====
        records = {"学会発表": [], "論文・教科書執筆等": []}


        # ===== 表示順（存在しない列は自動スキップ） =====
        columns_order = {
            "学会発表": [
                "年度","日付","事業所","診療科","発表者","カテゴリ","分類",
                "学会名","演題名","開催地","開催形式","発表形態"
            ],
            "論文・教科書執筆等": [
                "年度","事業所","診療科","発表者","カテゴリ","分類",
                "誌名","タイトル"
            ]
        }


        # ===== すべてのシートからカテゴリ判定して抽出 =====
        for sh in xl.sheet_names:
            df = xl.parse(sh)
            df.columns = [str(c).strip() for c in df.columns]
            if "カテゴリ" not in df.columns:
                continue
            unique_cats = set(df["カテゴリ"].dropna().astype(str))
            for cat in ["学会発表", "論文・教科書執筆等"]:
                if cat in unique_cats:
                    sub = df[df["カテゴリ"].astype(str) == cat].copy()
                    if "年度" in sub.columns and "診療科" in sub.columns:
                        sub = sub[~sub["年度"].isna() & ~sub["診療科"].isna()].copy()
                    else:
                        # 必須列が無いシートはスキップ
                        continue
                    for col in sub.columns:
                        sub[col] = sub[col].apply(lambda x: "" if pd.isna(x) else str(x))
                    order = [c for c in columns_order[cat] if c in sub.columns]
                    sub = sub[order]
                    records[cat].extend(sub.to_dict(orient="records"))


# ===== 選択肢生成（診療科は年度非依存） =====
        # 変更点：年度の選択肢を降順に。Excel由来の 2024.0 / NaN / 空文字 / "2024年度" 等も安全に処理。
        import math

        def _parse_year(y):
            """
            年度文字列/数値を並び替え用の整数に変換。
            - 例: "2024年度" → 2024, "2024" → 2024, 2024.0 → 2024
            - 変換不能（例: "", None, NaN, "不明", "令和6" など）は None
            """
            if y is None:
                return None
            # NaN 対応（float型 NaN）
            if isinstance(y, float) and math.isnan(y):
                return None

            s = str(y).strip()
            if not s:
                return None

            # "2024年度" の「年度」を落とす
            if s.endswith("年度"):
                s = s[:-2].strip()

            # 小数 "2024.0" や "2024.00" に対応
            # まずは整数として読めるか確認
            if s.isdigit():
                try:
                    return int(s)
                except Exception:
                    return None

            # 小数表現なら整数化トライ
            try:
                f = float(s)
                if not math.isnan(f) and f.is_integer():
                    return int(f)
            except Exception:
                pass

            # ここまで来たら数値化不可
            return None

        # 一意な年度の収集
        unique_years = list({
            r.get("年度", "")
            for cat in records
            for r in records[cat]
            if r.get("年度", "") is not None
        })

        # 並び順ポリシー：
        #   1) 数値化できる年度を「数値の降順」
        #   2) 数値化できない年度は「文字列の降順」で後ろ側
        all_years = sorted(
            unique_years,
            key=lambda y: (
                _parse_year(y) is not None,                           # 数値化できるものを前に
                _parse_year(y) if _parse_year(y) is not None else str(y)
            ),
            reverse=True  # 降順
        )

        # 診療科（年度非依存）
        all_depts_set = set()
        for cat in records:
            for r in records[cat]:
                d = r.get("診療科", "")
                if d:
                    all_depts_set.add(d)
        all_depts = sorted(list(all_depts_set))

        choices = {"年度": all_years, "診療科": all_depts}


        # ===== 見た目（CSS） =====
        css = """
* { box-sizing: border-box; }
body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', 'Yu Gothic', Meiryo, sans-serif; margin: 24px; }
h1 { font-size: 1.6rem; margin: 0 0 12px; }
header .meta { color: #666; font-size: .9rem; margin-bottom: 16px; }
.controls { display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0 12px; }
.controls label { font-weight: 600; font-size: .95rem; }
.app-container { display: flex; flex-direction: column; gap: 16px; width: 100%; }
.card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin: 0; width: 100%;  margin: 12px 0; }
.card h2 { font-size: 1.2rem; margin: 0 0 8px; }
.count { color: #333; font-size: .95rem; margin-bottom: 8px; }
.tablewrap { overflow-x: auto; border: 1px solid #eee; border-radius: 6px; }
table { border-collapse: collapse; width: max-content; min-width: 920px; }
th, td { padding: 8px 10px; border-bottom: 1px solid #eee; text-align: left; white-space: nowrap; word-break: keep-all; }
th { background: #f8f9fb; position: sticky; top: 0; z-index: 1; }
tr:nth-child(even) td { background: #fcfcff; }
.empty { color: #666; padding: 12px; }
.footer { margin-top: 18px; color: #555; font-size: .9rem; }
button { padding: 8px 12px; font-size: .9rem; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; background: #fff; }
button:hover { background: #f4f5f7; }
.note { color: #777; font-size: .85rem; }
/* 追加（選択中項目）：01/14 */
.badge { display:inline-block; padding: 2px 8px; background: #eef2ff; border:1px solid #c7d2fe; border-radius: 999px; font-size: .8rem; color:#1e40af; }


/* チェックボックス：縦並び */
.checkbox-list { display: grid; grid-template-columns: 1fr; gap: 6px; max-height: 240px; overflow-y: auto; border: 1px solid #eee; padding: 8px; border-radius: 6px; background: #fff; }
.chk { display: flex; align-items: center; gap: 8px; font-size: .95rem; }


/* ドロップダウン */
.dropdown { position: relative; display: inline-block; }
.dropdown-toggle {
  padding: 8px 12px; font-size: .95rem; border: 1px solid #ccc; border-radius: 6px; background: #fff; cursor: pointer;
}
.dropdown-toggle[aria-expanded="true"] { background: #f4f5f7; }
.dropdown-panel {
  position: absolute; z-index: 10; min-width: 280px; margin-top: 6px;
  background: #fff; border: 1px solid #ddd; border-radius: 8px; box-shadow: 0 6px 18px rgba(0,0,0,.08);
  padding: 10px; display: none;
}
.dropdown-panel.open { display: block; }
.dropdown-actions { display: flex; gap: 8px; justify-content: flex-end; margin-bottom: 8px; }
.dropdown-actions button {
  padding: 6px 10px; font-size: .85rem; border: 1px solid #ccc; border-radius: 6px; background: #fff; cursor: pointer;
}
dropdown-actions button:hover { background: #f4f5f7; }
"""


        # ===== 動き（JavaScript） =====
        js = """
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;


// DOM
const yearList  = document.getElementById('year_list');
const deptList  = document.getElementById('dept_list');
const exportBtn = document.getElementById('export');


const ddYearBtn   = document.getElementById('dd-year-btn');
const ddYearPanel = document.getElementById('dd-year-panel');
const ddDeptBtn   = document.getElementById('dd-dept-btn');
const ddDeptPanel = document.getElementById('dd-dept-panel');


const yearSelectAllBtn = document.getElementById('year_select_all');
const yearClearAllBtn  = document.getElementById('year_clear_all');
const deptSelectAllBtn = document.getElementById('dept_select_all');
const deptClearAllBtn  = document.getElementById('dept_clear_all');


// --- チェックボックス生成 ---
function renderYearChoices() {
  yearList.innerHTML = CHOICES['年度']
    .map(y => `<label class="chk"><input type="checkbox" name="year" value="${y}">${y}</label>`)
    .join('');
}
function updateDeptChoices() {
  const list = CHOICES['診療科']; // 年度非依存
  deptList.innerHTML = list
    .map(d => `<label class="chk"><input type="checkbox" name="dept" value="${d}">${d}</label>`)
    .join('');
}


// --- 選択値取得 ---
function getChecked(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(el => el.value);
}


// --- フィルタ ---
function getFiltered() {
  const years = getChecked('year');
  const depts = getChecked('dept');
  const result = { '学会発表': [], '論文・教科書執筆等': [] };
  for (const cat of Object.keys(DATA)) {
    result[cat] = DATA[cat].filter(r =>
      (years.length === 0 || years.includes(r['年度'])) &&
      (depts.length === 0 || depts.includes(r['診療科']))
    );
  }
  return result;
}


// --- テーブル描画 ---
function makeTable(containerId, cat, rows) {
  const wrap = document.getElementById(containerId);
  wrap.innerHTML = '';
  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class="empty">該当するデータがありません。</div>';
    return;
  }
  const cols = COLS[cat].filter(c => rows.some(r => c in r));
  const thead = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r => '<tr>' + cols.map(c => `<td>${(r[c] ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  const html = '<div class="tablewrap"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}


// ★追加：選択中の年度／診療科をバッジ表示
function renderBadges() {
  const years = getChecked('year');  // 既存：チェック済みの年度配列
  const depts = getChecked('dept');  // 既存：チェック済みの診療科配列

  const ytxt = (years.length ? years.join(', ') : '未選択');
  const dtxt = (depts.length ? depts.join(', ') : '未選択');

  const yEl = document.getElementById('badge_year');
  const dEl = document.getElementById('badge_dept');

  if (yEl) yEl.textContent = ytxt;
  if (dEl) dEl.textContent = dtxt;
}


// --- 件数＆表の更新 ---
function renderCards() {
  const filtered = getFiltered();
  document.getElementById('count_gakkai').textContent = `${filtered['学会発表'].length} 件`;
  document.getElementById('count_ronbun').textContent = `${filtered['論文・教科書執筆等'].length} 件`;
  makeTable('tbl_gakkai', '学会発表', filtered['学会発表']);
  makeTable('tbl_ronbun', '論文・教科書執筆等', filtered['論文・教科書執筆等']);

// ★追加：選択バッジの更新
  renderBadges();

}


// --- CSVエクスポート ---
function exportCSV() {
  const filtered = getFiltered();
  const merged = filtered['学会発表'].map(r => ({ 'カテゴリ': '学会発表', ...r }))
                .concat(filtered['論文・教科書執筆等'].map(r => ({ 'カテゴリ': '論文・教科書執筆等', ...r })));
  if (merged.length === 0) { alert('出力対象がありません。'); return; }
  const cols = Array.from(new Set(merged.flatMap(r => Object.keys(r))));
  const rows = [cols.join(',')].concat(merged.map(r =>
    cols.map(c => String(r[c] ?? '').replaceAll('"', '""'))
        .map(v => /[",\\n]/.test(v) ? `"${v}"` : v)
        .join(',')
  ));
  const blob = new Blob([rows.join('\\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'filtered_export.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}


// --- ドロップダウン開閉 ---
function toggleDropdown(btn, panel, open) {
  const isOpen = (open != null) ? open : !(panel.classList.contains('open'));
  panel.classList.toggle('open', isOpen);
  btn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
}


// 年度
ddYearBtn.addEventListener('click', () => toggleDropdown(ddYearBtn, ddYearPanel));
// 診療科
ddDeptBtn.addEventListener('click', () => toggleDropdown(ddDeptBtn, ddDeptPanel));


// 外側クリックで閉じる
document.addEventListener('click', (e) => {
  const withinYear = ddYearBtn.contains(e.target) || ddYearPanel.contains(e.target);
  const withinDept = ddDeptBtn.contains(e.target) || ddDeptPanel.contains(e.target);
  if (!withinYear) toggleDropdown(ddYearBtn, ddYearPanel, false);
  if (!withinDept) toggleDropdown(ddDeptBtn, ddDeptPanel, false);
});


// Escキーで閉じる
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    toggleDropdown(ddYearBtn, ddYearPanel, false);
    toggleDropdown(ddDeptBtn, ddDeptPanel, false);
  }
});


// --- すべて選択／選択解除 ---
yearSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = true);
  renderCards();
});
yearClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = false);
  renderCards();
});
deptSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="dept"]').forEach(el => el.checked = true);
  renderCards();
});
deptClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="dept"]').forEach(el => el.checked = false);
  renderCards();
});


// --- チェック変更で即時反映 ---
document.addEventListener('change', (e) => {
  if (e.target && (e.target.name === 'year' || e.target.name === 'dept')) {
    renderCards();
  }
});


// 初期化
renderYearChoices();
updateDeptChoices();
renderCards();
exportBtn.addEventListener('click', exportCSV);
"""


        # ===== HTMLテンプレート =====
        html_template = """
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>学会発表／論文・教科書執筆等</title>
<style>[[CSS]]</style>
</head>
<body>
  <header>
    <h1>学会発表／論文・教科書執筆等</h1>
    <div class="controls">
      <!-- 年度（チェックボックス入りドロップダウン） -->
      <div class="dropdown">
        <button class="dropdown-toggle" id="dd-year-btn" aria-expanded="false" aria-controls="dd-year-panel">
          年度を選択（複数可）
        </button>
        <div class="dropdown-panel" id="dd-year-panel" role="listbox" aria-labelledby="dd-year-btn">
          <div class="dropdown-actions">
            <button type="button" id="year_select_all">すべて選択</button>
            <button type="button" id="year_clear_all">選択解除</button>
          </div>
          <div id="year_list" class="checkbox-list" aria-label="年度選択"></div>
        </div>
      </div>


      <!-- 診療科（チェックボックス入りドロップダウン） -->
      <div class="dropdown">
        <button class="dropdown-toggle" id="dd-dept-btn" aria-expanded="false" aria-controls="dd-dept-panel">
          診療科を選択（複数可）
        </button>
        <div class="dropdown-panel" id="dd-dept-panel" role="listbox" aria-labelledby="dd-dept-btn">
          <div class="dropdown-actions">
            <button type="button" id="dept_select_all">すべて選択</button>
            <button type="button" id="dept_clear_all">選択解除</button>
          </div>
          <div id="dept_list" class="checkbox-list" aria-label="診療科選択"></div>
        </div>
      </div>


      <!-- CSVダウンロード -->
      <div style="align-self: end;">
        <button id="export" title="現在の抽出結果をCSVで保存">CSVダウンロード</button>
        <div class="note">※年度・診療科は複数選択できます（未選択の場合は全件）</div>
      </div>
    </div>

    <div class="note">
      選択中 → 年度: <span class="badge" id="badge_year"></span>
      ／ 診療科: <span class="badge" id="badge_dept"></span>
      </div>

  </header>


  <div class="app-container">
    <section class="card">
      <h2>学会発表－検索結果</h2>
      <div class="count" id="count_gakkai"></div>
      <div id="tbl_gakkai"></div>
    </section>
    <section class="card">
      <h2>論文・教科書執筆等－検索結果</h2>
      <div class="count" id="count_ronbun"></div>
      <div id="tbl_ronbun"></div>
    </section>
  </div>


<script>
[[JS]]
</script>
</body>
</html>
"""


        # ===== HTML生成 =====
        html = html_template.replace("[[CSS]]", css)\
            .replace("[[SRCFILE]]", os.path.basename(EXCEL_FILE))\
            .replace("[[TIMESTAMP]]", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))


        js_filled = js.replace("__DATA__", json.dumps(records, ensure_ascii=False))\
            .replace("__CHOICES__", json.dumps(choices, ensure_ascii=False))\
            .replace("__COLS__", json.dumps(columns_order, ensure_ascii=False))


        html = html.replace("[[JS]]", js_filled)


        out_path = os.path.abspath(HTML_FILE)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(html)


        print(f"✓ 生成完了: {out_path}")
        print("=== 正常終了 ===")


    except Exception as e:
        print("✗ エラーが発生しました。詳細：")
        print(f"  種別: {type(e).__name__}")
        print(f"  内容: {e}")
        # 例外のトレースが必要であれば以下を解除
        # import traceback; traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
    