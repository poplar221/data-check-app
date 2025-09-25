import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
from zoneinfo import ZoneInfo
import numpy as np

def find_header_and_read_excel(uploaded_file, sheet_name, keywords):
    """
    Excelファイルからキーワードを含む行をヘッダーとして特定し、データを読み込む関数。
    """
    if uploaded_file:
        uploaded_file.seek(0)
    try:
        # .xlsと.xlsxの両方に対応するため、engineを自動選択させる
        df_no_header = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None, engine=None)
        header_row_index = -1
        for i, row in df_no_header.iterrows():
            row_str = ''.join(map(str, row.dropna().values))
            if all(keyword in row_str for keyword in keywords):
                header_row_index = i
                break
        
        if header_row_index == -1:
            st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' でヘッダー行(キーワード: {keywords})が見つかりませんでした。")
            return None
        
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_index, engine=None)
        return df

    except Exception as e:
        st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' 読込中にエラー: {e}")
        return None

def main():
    """
    アプリケーションのメイン関数
    """
    st.set_page_config(layout="wide")

    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
        st.session_state.summary_metrics = {}
        st.session_state.processed_data = None

    st.title("退職給付債務計算のための従業員データチェッカー")
    try:
        mod_time = os.path.getmtime(__file__)
        jst_time = datetime.fromtimestamp(mod_time, tz=ZoneInfo("Asia/Tokyo"))
        last_updated = jst_time.strftime('%Y年%m月%d日 %H:%M:%S JST')
        st.caption(f"最終更新日時: {last_updated}")
    except Exception:
        pass
    
    st.write("前期末、当期末、退職者の従業員データ（Excelファイル）をアップロードして、データの整合性チェックを行います。")

    st.subheader("📁 ファイルのアップロードと各種設定")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("##### 1. 前期末従業員データ (必須)")
        file_prev = st.file_uploader("アップロード", type=['xlsx', 'xls'], key="up_prev", label_visibility="collapsed")
        st.markdown("###### シート名")
        if file_prev:
            try:
                sheets = pd.ExcelFile(file_prev).sheet_names
                default_sheet = "従業員データフォーマット"
                index = sheets.index(default_sheet) if default_sheet in sheets else 0
                sheet_prev = st.selectbox("シートを選択", options=sheets, index=index, key="sheet_prev", label_visibility="collapsed")
            except Exception:
                sheet_prev = st.text_input("シート名を入力", "従業員データフォーマット", key="sheet_prev", label_visibility="collapsed")
        else:
            sheet_prev = st.text_input("シート名を入力", "従業員データフォーマット", key="sheet_prev", label_visibility="collapsed")
        st.markdown("###### ヘッダー行 特定キーワード")
        keyword_prev_1 = st.text_input("キーワード1", "入社", key="kw_p1")
        keyword_prev_2 = st.text_input("キーワード2", "生年", key="kw_p2")

    with col2:
        st.markdown("##### 2. 当期末従業員データ (必須)")
        file_curr = st.file_uploader("アップロード", type=['xlsx', 'xls'], key="up_curr", label_visibility="collapsed")
        st.markdown("###### シート名")
        if file_curr:
            try:
                sheets = pd.ExcelFile(file_curr).sheet_names
                default_sheet = "従業員データフォーマット"
                index = sheets.index(default_sheet) if default_sheet in sheets else 0
                sheet_curr = st.selectbox("シートを選択", options=sheets, index=index, key="sheet_curr", label_visibility="collapsed")
            except Exception:
                sheet_curr = st.text_input("シート名を入力", "従業員データフォーマット", key="sheet_curr", label_visibility="collapsed")
        else:
            sheet_curr = st.text_input("シート名を入力", "従業員データフォーマット", key="sheet_curr", label_visibility="collapsed")
        st.markdown("###### ヘッダー行 特定キーワード")
        keyword_curr_1 = st.text_input("キーワード1", "入社", key="kw_c1")
        keyword_curr_2 = st.text_input("キーワード2", "生年", key="kw_c2")

    keywords_prev = [k for k in [keyword_prev_1, keyword_prev_2] if k]
    keywords_curr = [k for k in [keyword_curr_1, keyword_curr_2] if k]
    
    with st.expander("列名設定を展開/折りたたみ", expanded=True):
        NONE_OPTION = "(選択しない)"
        columns_prev, columns_curr, columns_retire = [], [], []
        if file_prev and sheet_prev:
            df_cols = find_header_and_read_excel(file_prev, sheet_prev, keywords=keywords_prev)
            if df_cols is not None: columns_prev = df_cols.columns.tolist()
        if file_curr and sheet_curr:
            df_cols = find_header_and_read_excel(file_curr, sheet_curr, keywords=keywords_curr)
            if df_cols is not None: columns_curr = df_cols.columns.tolist()
        
        def create_column_selector(label, default_name, columns, key, disabled=False):
            if columns:
                options = [NONE_OPTION] + columns
                index = options.index(default_name) if default_name in options else 0
                return st.selectbox(label, options=options, index=index, key=key, disabled=disabled)
            else:
                return st.text_input(label, default_name, key=key, disabled=disabled)

        st.info("ファイルをアップロードしシートを選択すると、下のドロップダウンに列名が自動で表示されます。")
        map_col1, map_col2, map_col3 = st.columns(3)
        with map_col1:
            st.markdown("<h6>① 前期末データ</h6>", unsafe_allow_html=True)
            col_emp_id_prev = create_column_selector("従業員番号", "従業員番号", columns_prev, "emp_id_prev")
            col_hire_date_prev = create_column_selector("入社年月日", "入社年月日", columns_prev, "hire_date_prev")
            col_enroll_date_prev = create_column_selector("加入年月日", "加入年月日", columns_prev, "enroll_date_prev")
            col_birth_date_prev = create_column_selector("生年月日", "生年月日", columns_prev, "birth_date_prev")
            st.text_input("退職日", value="－ (対象外) －", key="retire_date_prev_dummy", disabled=True)
            st.markdown("---")
            col_salary1_prev = create_column_selector("給与1", "給与1", columns_prev, "salary1_prev")
            salary1_prev_selected = (col_salary1_prev != NONE_OPTION)
            col_salary2_prev = create_column_selector("給与2", "給与2", columns_prev, "salary2_prev", disabled=not salary1_prev_selected)
            st.markdown("---")
            col_salary3_prev = create_column_selector("給与3", "給与3", columns_prev, "salary3_prev")
            salary3_prev_selected = (col_salary3_prev != NONE_OPTION)
            col_salary4_prev = create_column_selector("給与4", "給与4", columns_prev, "salary4_prev", disabled=not salary3_prev_selected)
        
        with map_col2:
            st.markdown("<h6>② 当期末データ</h6>", unsafe_allow_html=True)
            col_emp_id_curr = create_column_selector("従業員番号", "従業員番号", columns_curr, "emp_id_curr")
            col_hire_date_curr = create_column_selector("入社年月日", "入社年月日", columns_curr, "hire_date_curr")
            col_enroll_date_curr = create_column_selector("加入年月日", "加入年月日", columns_curr, "enroll_date_curr")
            col_birth_date_curr = create_column_selector("生年月日", "生年月日", columns_curr, "birth_date_curr")
            col_retire_date_curr = create_column_selector("退職日", "退職年月日", columns_curr, "retire_date_curr")
            st.markdown("---")
            col_salary1_curr = create_column_selector("給与1", "給与1", columns_curr, "salary1_curr")
            salary1_curr_selected = (col_salary1_curr != NONE_OPTION)
            col_salary2_curr = create_column_selector("給与2", "給与2", columns_curr, "salary2_curr", disabled=not salary1_curr_selected)
            st.markdown("---")
            col_salary3_curr = create_column_selector("給与3", "給与3", columns_curr, "salary3_curr")
            salary3_curr_selected = (col_salary3_curr != NONE_OPTION)
            col_salary4_curr = create_column_selector("給与4", "給与4", columns_curr, "salary4_curr", disabled=not salary3_curr_selected)
        
        retire_file_is_used = (col_retire_date_curr == NONE_OPTION)
        
        with col3:
            st.markdown("##### 3. 当期退職者データ (任意)")
            file_retire = st.file_uploader("アップロード", type=['xlsx', 'xls'], disabled=not retire_file_is_used, help="メイン画面の「列名設定」で「退職日」列を指定した場合、このアップローダーは無効になります。", key="up_retire", label_visibility="collapsed")
            st.markdown("###### シート名")
            if file_retire:
                try:
                    sheets = pd.ExcelFile(file_retire).sheet_names
                    default_sheet = "退職者データフォーマット"
                    index = sheets.index(default_sheet) if default_sheet in sheets else 0
                    sheet_retire = st.selectbox("シートを選択", options=sheets, index=index, key="sheet_retire", label_visibility="collapsed", disabled=not retire_file_is_used)
                except Exception:
                    sheet_retire = st.text_input("シート名を入力", "退職者データフォーマット", key="sheet_retire", label_visibility="collapsed", disabled=not retire_file_is_used)
            else:
                sheet_retire = st.text_input("シート名を入力", "退職者データフォーマット", key="sheet_retire", label_visibility="collapsed", disabled=not retire_file_is_used)
            st.markdown("###### ヘッダー行 特定キーワード")
            keyword_retire_1 = st.text_input("キーワード1", "退職", key="kw_r1", disabled=not retire_file_is_used)
            keyword_retire_2 = st.text_input("キーワード2", "生年", key="kw_r2", disabled=not retire_file_is_used)
        keywords_retire = [k for k in [keyword_retire_1, keyword_retire_2] if k]

        if file_retire and sheet_retire and retire_file_is_used:
            df_cols = find_header_and_read_excel(file_retire, sheet_retire, keywords=keywords_retire)
            if df_cols is not None:
                columns_retire = df_cols.columns.tolist()

        with map_col3:
            st.markdown("<h6>③ 退職者データ</h6>", unsafe_allow_html=True)
            if retire_file_is_used:
                col_emp_id_retire = create_column_selector("従業員番号", "従業員番号", columns_retire, "emp_id_retire")
                col_hire_date_retire = create_column_selector("入社年月日", "入社年月日", columns_retire, "hire_date_retire")
                col_enroll_date_retire = create_column_selector("加入年月日", "加入年月日", columns_retire, "enroll_date_retire")
                col_birth_date_retire = create_column_selector("生年月日", "生年月日", columns_retire, "birth_date_retire")
                col_retire_date_retire = create_column_selector("退職日", "退職年月日", columns_retire, "retire_date_retire")
            else:
                st.warning("「当期末データ」の「退職日」列が指定されているため、退職者ファイルは使用されません。")
    
    with st.sidebar:
        st.header("⚙️ データ指定設定")
        base_date = st.date_input("計算基準日（当期末）", value=datetime.now(), help="チェックの基準となる当期末の日付を指定します。")
        st.markdown("---")
        st.header("✔️ 追加エラーチェック設定")
        
        st.markdown("##### 給与1, 2のチェック")
        cumulative_checks_disabled_12 = (col_salary2_prev == NONE_OPTION or col_salary2_curr == NONE_OPTION)
        check_salary_decrease_1 = st.checkbox("給与減額チェック(1)", value=True, help="在籍者のうち、当期末の給与1が前期末の給与1よりも減少している従業員を検出します。")
        check_salary_increase_1 = st.checkbox("給与増加率チェック(1)", value=True, help="在籍者のうち、当期末の給与1が前期末の給与1に比べて、指定した増加率（x%）以上に増加している従業員を検出します。")
        increase_rate_x1 = st.text_input("増加率(x1)%", value="5")
        check_cumulative_salary_1_ui = st.checkbox("累計給与チェック(1-1)", value=True, help="在籍者のうち、当期末の累計給与2が「前期末の累計給与2 + 前期末の給与1 × 月数(y)」の計算結果よりも少ない従業員を検出します。", disabled=cumulative_checks_disabled_12)
        months_y1 = st.selectbox("月数(y1)", ("1", "12"), index=0, disabled=cumulative_checks_disabled_12)
        check_cumulative_salary_2_ui = st.checkbox("累計給与チェック(1-2)", value=True, help="在籍者のうち、当期末の累計給与2が「(前期末の累計給与2 + 前期末の給与1 × 月数(y)) × (1 + 許容率(z)%))」の計算結果よりも多い従業員を検出します。", disabled=cumulative_checks_disabled_12)
        allowance_rate_z1 = st.text_input("許容率(z1)%", value="0", disabled=cumulative_checks_disabled_12)
        if cumulative_checks_disabled_12:
            check_cumulative_salary_1, check_cumulative_salary_2 = False, False
        else:
            check_cumulative_salary_1, check_cumulative_salary_2 = check_cumulative_salary_1_ui, check_cumulative_salary_2_ui

        st.markdown("---")
        st.markdown("##### 給与3, 4のチェック")
        salary3_checks_disabled = (col_salary3_prev == NONE_OPTION or col_salary3_curr == NONE_OPTION)
        cumulative_checks_disabled_34 = (col_salary4_prev == NONE_OPTION or col_salary4_curr == NONE_OPTION) or salary3_checks_disabled
        check_salary_decrease_3_ui = st.checkbox("給与減額チェック(3)", value=True, help="在籍者のうち、当期末の給与3が前期末の給与3よりも減少している従業員を検出します。", disabled=salary3_checks_disabled)
        check_salary_increase_3_ui = st.checkbox("給与増加率チェック(3)", value=True, help="在籍者のうち、当期末の給与3が前期末の給与3に比べて、指定した増加率（x%）以上に増加している従業員を検出します。", disabled=salary3_checks_disabled)
        increase_rate_x3 = st.text_input("増加率(x3)%", value="5", disabled=salary3_checks_disabled)
        check_cumulative_salary_3_ui = st.checkbox("累計給与チェック(3-1)", value=True, help="在籍者のうち、当期末の累計給与4が「前期末の累計給与4 + 前期末の給与3 × 月数(y)」の計算結果よりも少ない従業員を検出します。", disabled=cumulative_checks_disabled_34)
        months_y3 = st.selectbox("月数(y3)", ("1", "12"), index=0, disabled=cumulative_checks_disabled_34)
        check_cumulative_salary_4_ui = st.checkbox("累計給与チェック(3-2)", value=True, help="在籍者のうち、当期末の累計給与4が「(前期末の累計給与4 + 前期末の給与3 × 月数(y)) × (1 + 許容率(z)%))」の計算結果よりも多い従業員を検出します。", disabled=cumulative_checks_disabled_34)
        allowance_rate_z3 = st.text_input("許容率(z3)%", value="0", disabled=cumulative_checks_disabled_34)
        if salary3_checks_disabled:
            check_salary_decrease_3, check_salary_increase_3 = False, False
        else:
            check_salary_decrease_3, check_salary_increase_3 = check_salary_decrease_3_ui, check_salary_increase_3_ui
        if cumulative_checks_disabled_34:
            check_cumulative_salary_3, check_cumulative_salary_4 = False, False
        else:
            check_cumulative_salary_3, check_cumulative_salary_4 = check_cumulative_salary_3_ui, check_cumulative_salary_4_ui

    if st.button("チェック開始", use_container_width=True, type="primary"):
        st.session_state.processing_complete = False
        if file_prev and file_curr:
            with st.spinner('データチェックを実行中です...'):
                try:
                    base_date_ts = pd.Timestamp(base_date)
                    prev_period_end_date_ts = base_date_ts - pd.DateOffset(years=1)
                    INTERNAL_COLS = {
                        "emp_id": "_emp_id", "hire_date": "_hire_date", "enroll_date": "_enroll_date", "birth_date": "_birth_date", 
                        "retire_date": "_retire_date", "salary1": "_salary1", "salary2": "_salary2",
                        "salary3": "_salary3", "salary4": "_salary4"
                    }
                    selections_prev = { "emp_id": col_emp_id_prev, "hire_date": col_hire_date_prev, "enroll_date": col_enroll_date_prev, "birth_date": col_birth_date_prev, "salary1": col_salary1_prev, "salary2": col_salary2_prev, "salary3": col_salary3_prev, "salary4": col_salary4_prev }
                    selections_curr = { "emp_id": col_emp_id_curr, "hire_date": col_hire_date_curr, "enroll_date": col_enroll_date_curr, "birth_date": col_birth_date_curr, "retire_date": col_retire_date_curr, "salary1": col_salary1_curr, "salary2": col_salary2_curr, "salary3": col_salary3_curr, "salary4": col_salary4_curr }
                    if retire_file_is_used: selections_retire = { "emp_id": col_emp_id_retire, "hire_date": col_hire_date_retire, "enroll_date": col_enroll_date_retire, "birth_date": col_birth_date_retire, "retire_date": col_retire_date_retire }
                    def rename_df_columns(df, selections):
                        rename_map = {v: INTERNAL_COLS[k] for k, v in selections.items() if v != NONE_OPTION and v in df.columns}
                        return df.rename(columns=rename_map)

                    st.info("ステップ1/7: Excelファイルを読み込み、列名を標準化しています...")
                    df_prev = find_header_and_read_excel(file_prev, sheet_prev, keywords=keywords_prev); df_curr = find_header_and_read_excel(file_curr, sheet_curr, keywords=keywords_curr); df_retire = None
                    if df_prev is None or df_curr is None:
                        st.error("🚫 **処理停止: 必須ファイルが読み込めませんでした。**", icon="🚨"); st.warning("ファイル設定やヘッダーキーワードが正しいか確認してください。"); st.stop()
                    
                    df_prev = rename_df_columns(df_prev, selections_prev); df_curr = rename_df_columns(df_curr, selections_curr)

                    if col_retire_date_curr != NONE_OPTION and INTERNAL_COLS["retire_date"] in df_curr.columns:
                        st.info(f"ステップ1.5/7: 当期末データから退職者を抽出...")
                        retiree_mask = df_curr[INTERNAL_COLS["retire_date"]].notna()
                        df_retire = df_curr[retiree_mask].copy(); df_curr = df_curr[~retiree_mask].copy()
                        if not df_retire.empty: st.success(f"{len(df_retire)}名の退職者を当期末データから抽出し、在籍者から除外しました。")
                    elif file_retire:
                        df_retire = find_header_and_read_excel(file_retire, sheet_retire, keywords=keywords_retire)
                        if df_retire is not None: df_retire = rename_df_columns(df_retire, selections_retire)

                    st.info("ステップ1.8/7: 日付列を日付形式に変換しています...")
                    date_cols_to_convert = [INTERNAL_COLS["hire_date"], INTERNAL_COLS["enroll_date"], INTERNAL_COLS["birth_date"], INTERNAL_COLS["retire_date"]]
                    for df in [df_prev, df_curr, df_retire]:
                        if df is not None:
                            for col in date_cols_to_convert:
                                if col in df.columns: df[col] = pd.to_datetime(df[col].astype(str), errors='coerce')
                            if INTERNAL_COLS["enroll_date"] in df.columns:
                                if INTERNAL_COLS["hire_date"] in df.columns:
                                    df[INTERNAL_COLS["hire_date"]].fillna(df[INTERNAL_COLS["enroll_date"]], inplace=True)
                                else:
                                    df.rename(columns={INTERNAL_COLS["enroll_date"]: INTERNAL_COLS["hire_date"]}, inplace=True)
                                df.drop(columns=[INTERNAL_COLS["enroll_date"]], inplace=True, errors='ignore')

                    st.info("ステップ2/7: マッチングキーを決定しています...")
                    use_emp_id_key = (INTERNAL_COLS["emp_id"] in df_prev.columns and INTERNAL_COLS["emp_id"] in df_curr.columns)
                    dataframes = {'前期末': df_prev, '当期末': df_curr}
                    if df_retire is not None:
                        use_emp_id_key = use_emp_id_key and (INTERNAL_COLS["emp_id"] in df_retire.columns); dataframes['退職者'] = df_retire
                    
                    key_col_name = '_key'
                    for name, df in dataframes.items():
                        if not use_emp_id_key:
                             if not {INTERNAL_COLS["hire_date"], INTERNAL_COLS["birth_date"]}.issubset(df.columns):
                                st.error(f"🚫 **処理停止: 代替キーに必要な列が見つかりませんでした。**", icon="🚨"); st.warning(f"「{name}」データで、代替キーの列マッピングが正しく行われているか確認してください。"); st.stop()
                             df[key_col_name] = df[INTERNAL_COLS["hire_date"]].dt.strftime('%Y%m%d').fillna('NODATE') + '_' + df[INTERNAL_COLS["birth_date"]].dt.strftime('%Y%m%d').fillna('NODATE')
                        else: df[key_col_name] = df[INTERNAL_COLS["emp_id"]].astype(str)
                    key_type = "従業員番号" if use_emp_id_key else "入社年月日/加入年月日 + 生年月日"; st.success(f"マッチングキーとして '{key_type}' を使用します。")
                    
                    results = {}; st.info("ステップ3/7: 基本エラーチェック...")
                    for name, df in dataframes.items():
                        duplicates = df[df[key_col_name].duplicated(keep=False)]; results[f'キー重複_{name}'] = duplicates.sort_values(by=key_col_name)
                    
                    for name, df, relevant_date, date_type in [('前期末', df_prev, prev_period_end_date_ts, '前期末日'), ('当期末', df_curr, base_date_ts, '計算基準日')]:
                        if df is None: continue
                        temp_errors = []
                        if {INTERNAL_COLS["hire_date"], INTERNAL_COLS["birth_date"]}.issubset(df.columns):
                            df_copy = df.copy(); valid_dates = df_copy.dropna(subset=[INTERNAL_COLS["hire_date"], INTERNAL_COLS["birth_date"]])
                            if not valid_dates.empty:
                                age = (valid_dates[INTERNAL_COLS["hire_date"]] - valid_dates[INTERNAL_COLS["birth_date"]]).dt.days / 365.25
                                invalid_age_df = df.loc[valid_dates[(age < 15) | (age >= 90)].index].copy()
                                if not invalid_age_df.empty:
                                    invalid_age_df['エラー理由'] = '入社時年齢が15歳未満または90歳以上'; temp_errors.append(invalid_age_df)
                        if INTERNAL_COLS["hire_date"] in df.columns:
                             invalid_hire_date_df = df[df[INTERNAL_COLS["hire_date"]] > relevant_date].copy()
                             if not invalid_hire_date_df.empty:
                                 invalid_hire_date_df['エラー理由'] = f'入社日が{date_type}({relevant_date.date()})より後'; temp_errors.append(invalid_hire_date_df)
                        if temp_errors:
                            df_with_reasons = pd.concat(temp_errors).drop_duplicates(subset=[key_col_name]); results[f'日付妥当性エラー_{name}'] = df_with_reasons
                    
                    if df_retire is not None and INTERNAL_COLS["retire_date"] in df_retire.columns:
                        temp_errors_retire = []
                        invalid_retire1 = df_retire[df_retire[INTERNAL_COLS["retire_date"]] <= prev_period_end_date_ts].copy()
                        if not invalid_retire1.empty:
                            invalid_retire1['エラー理由'] = f'退職日が前期末日({prev_period_end_date_ts.date()})以前'; temp_errors_retire.append(invalid_retire1)
                        invalid_retire2 = df_retire[df_retire[INTERNAL_COLS["retire_date"]] > base_date_ts].copy()
                        if not invalid_retire2.empty:
                            invalid_retire2['エラー理由'] = f'退職日が計算基準日({base_date_ts.date()})より後'; temp_errors_retire.append(invalid_retire2)
                        if temp_errors_retire:
                            results['日付妥当性エラー_退職者'] = pd.concat(temp_errors_retire).drop_duplicates(subset=[key_col_name])
                    
                    st.info("ステップ4/7: 在籍者・退職者・入社者の照合..."); merged_st = pd.merge(df_prev, df_curr, on=key_col_name, how='outer', suffixes=('_前期', '_当期'), indicator=True)
                    retiree_candidates = merged_st[merged_st['_merge'] == 'left_only'].copy(); new_hires = merged_st[merged_st['_merge'] == 'right_only'].copy(); continuing_employees = merged_st[merged_st['_merge'] == 'both'].copy()
                    results['入社者候補'] = new_hires
                    
                    st.info("ステップ4.5/7: 在籍者の基本情報変更チェック...")
                    bdate_prev, bdate_curr = f'{INTERNAL_COLS["birth_date"]}_前期', f'{INTERNAL_COLS["birth_date"]}_当期'; hdate_prev, hdate_curr = f'{INTERNAL_COLS["hire_date"]}_前期', f'{INTERNAL_COLS["hire_date"]}_当期'
                    if all(c in continuing_employees.columns for c in [bdate_prev, bdate_curr, hdate_prev, hdate_curr]):
                        changed_birth_date = continuing_employees[bdate_prev] != continuing_employees[bdate_curr]
                        changed_hire_date = continuing_employees[hdate_prev] != continuing_employees[hdate_curr]
                        changed_df = continuing_employees[changed_birth_date | changed_hire_date].copy()
                        changed_df['エラー理由'] = '前期と当期で基本情報(生年月日 or 入社日)が不一致'
                        results['基本情報変更エラー'] = changed_df
                    else: st.warning("生年月日または入社年月日の列が揃っていないため、基本情報変更チェックはスキップされました。")
                    
                    if df_retire is not None:
                        st.info("ステップ4.8/7: 退職者データの照合...")
                        merged_retire = pd.merge(retiree_candidates[[key_col_name]], df_retire, on=key_col_name, how='outer', indicator='retire_merge')
                        results['退職者候補（退職者データ不突合）'] = retiree_candidates[retiree_candidates[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'left_only'][key_col_name])]
                        results['退職者データ過剰（前期末データ不突合）'] = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'right_only'][key_col_name])]
                        results['マッチした退職者'] = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'both'][key_col_name])]
                    else: results['退職者候補'] = retiree_candidates
                    results['在籍者'] = continuing_employees
                    
                    st.info("ステップ5/7: 追加エラーチェック...")
                    sal1_int, sal2_int = INTERNAL_COLS["salary1"], INTERNAL_COLS["salary2"]
                    required_salary1_cols = {f'{sal1_int}_前期', f'{sal1_int}_当期'}
                    if required_salary1_cols.issubset(continuing_employees.columns):
                        check_df_sal1 = continuing_employees.copy(); [check_df_sal1.update({col: pd.to_numeric(check_df_sal1[col], errors='coerce')}) for col in required_salary1_cols]; check_df_sal1.dropna(subset=required_salary1_cols, inplace=True)
                        if check_salary_decrease_1: results['給与減額エラー(1)'] = check_df_sal1[check_df_sal1[f'{sal1_int}_当期'] < check_df_sal1[f'{sal1_int}_前期']]
                        if check_salary_increase_1:
                            try: x1 = float(increase_rate_x1); results['給与増加率エラー(1)'] = check_df_sal1[check_df_sal1[f'{sal1_int}_当期'] >= check_df_sal1[f'{sal1_int}_前期'] * (1 + x1 / 100)]
                            except ValueError: st.warning("給与増加率(x1)が無効な数値のためスキップしました。")
                        required_salary2_cols = {f'{sal2_int}_前期', f'{sal2_int}_当期'}
                        if not cumulative_checks_disabled_12 and required_salary2_cols.issubset(check_df_sal1.columns):
                            check_df_sal2 = check_df_sal1.copy(); [check_df_sal2.update({col: pd.to_numeric(check_df_sal2[col], errors='coerce')}) for col in required_salary2_cols]; check_df_sal2.dropna(subset=required_salary2_cols, inplace=True)
                            if check_cumulative_salary_1:
                                try: y1 = int(months_y1); results['累計給与エラー(1-1)'] = check_df_sal2[check_df_sal2[f'{sal2_int}_当期'] < check_df_sal2[f'{sal2_int}_前期'] + check_df_sal2[f'{sal1_int}_前期'] * y1]
                                except ValueError: st.warning("月数(y1)が無効な数値のためスキップしました。")
                            if check_cumulative_salary_2:
                                try: y1 = int(months_y1); z1 = float(allowance_rate_z1); upper_limit = (check_df_sal2[f'{sal2_int}_前期'] + check_df_sal2[f'{sal1_int}_前期'] * y1) * (1 + z1 / 100); results['累計給与エラー(1-2)'] = check_df_sal2[check_df_sal2[f'{sal2_int}_当期'] > upper_limit]
                                except ValueError: st.warning("月数(y1)または許容率(z1)が無効な数値のためスキップしました。")
                        elif not cumulative_checks_disabled_12: st.warning(f"「給与2」の列が指定/存在しないため、累計給与チェック(1)はスキップされました。")
                    else: st.warning(f"「給与1」の列が指定/存在しないため、給与1,2のチェックはスキップされました。")

                    sal3_int, sal4_int = INTERNAL_COLS["salary3"], INTERNAL_COLS["salary4"]
                    required_salary3_cols = {f'{sal3_int}_前期', f'{sal3_int}_当期'}
                    if required_salary3_cols.issubset(continuing_employees.columns):
                        check_df_sal3 = continuing_employees.copy(); [check_df_sal3.update({col: pd.to_numeric(check_df_sal3[col], errors='coerce')}) for col in required_salary3_cols]; check_df_sal3.dropna(subset=required_salary3_cols, inplace=True)
                        if check_salary_decrease_3: results['給与減額エラー(3)'] = check_df_sal3[check_df_sal3[f'{sal3_int}_当期'] < check_df_sal3[f'{sal3_int}_前期']]
                        if check_salary_increase_3:
                            try: x3 = float(increase_rate_x3); results['給与増加率エラー(3)'] = check_df_sal3[check_df_sal3[f'{sal3_int}_当期'] >= check_df_sal3[f'{sal3_int}_前期'] * (1 + x3 / 100)]
                            except ValueError: st.warning("給与増加率(x3)が無効な数値のためスキップしました。")
                        required_salary4_cols = {f'{sal4_int}_前期', f'{sal4_int}_当期'}
                        if not cumulative_checks_disabled_34 and required_salary4_cols.issubset(check_df_sal3.columns):
                            check_df_sal4 = check_df_sal3.copy(); [check_df_sal4.update({col: pd.to_numeric(check_df_sal4[col], errors='coerce')}) for col in required_salary4_cols]; check_df_sal4.dropna(subset=required_salary4_cols, inplace=True)
                            if check_cumulative_salary_3:
                                try: y3 = int(months_y3); results['累計給与エラー(3-1)'] = check_df_sal4[check_df_sal4[f'{sal4_int}_当期'] < check_df_sal4[f'{sal4_int}_前期'] + check_df_sal4[f'{sal3_int}_前期'] * y3]
                                except ValueError: st.warning("月数(y3)が無効な数値のためスキップしました。")
                            if check_cumulative_salary_4:
                                try: y3 = int(months_y3); z3 = float(allowance_rate_z3); upper_limit = (check_df_sal4[f'{sal4_int}_前期'] + check_df_sal4[f'{sal3_int}_前期'] * y3) * (1 + z3 / 100); results['累計給与エラー(3-2)'] = check_df_sal4[check_df_sal4[f'{sal4_int}_当期'] > upper_limit]
                                except ValueError: st.warning("月数(y3)または許容率(z3)が無効な数値のためスキップしました。")
                        elif not cumulative_checks_disabled_34: st.warning(f"「給与4」の列が指定/存在しないため、累計給与チェック(3)はスキップされました。")
                    else: st.warning(f"「給与3」の列が指定/存在しないため、給与3,4のチェックはスキップされました。")
                    
                    st.info("ステップ6/7: 結果をExcelファイルにまとめています...")
                    summary_info = {
                        "前期従業員データ数": len(df_prev), 
                        "当期従業員データ数": len(df_curr),
                        "在籍者数": len(results.get('在籍者', [])),
                        "マッチした退職者": len(results.get('マッチした退職者', []))
                        }
                    if df_retire is not None: 
                        summary_info["当期退職者データ数"] = len(df_retire)
                    
                    summary_errors = {
                        "キー重複": sum(len(df) for name, df in results.items() if 'キー重複' in name), 
                        "日付妥当性エラー": sum(len(df) for name, df in results.items() if '日付妥当性' in name), 
                        "基本情報変更エラー": len(results.get('基本情報変更エラー', [])), 
                        "入社者候補": len(results.get('入社者候補', [])), 
                        "給与減額エラー(1)": len(results.get('給与減額エラー(1)', [])), "給与増加率エラー(1)": len(results.get('給与増加率エラー(1)', [])), 
                        "累計給与エラー(1-1)": len(results.get('累計給与エラー(1-1)', [])), "累計給与エラー(1-2)": len(results.get('累計給与エラー(1-2)', [])),
                        "給与減額エラー(3)": len(results.get('給与減額エラー(3)', [])), "給与増加率エラー(3)": len(results.get('給与増加率エラー(3)', [])),
                        "累計給与エラー(3-1)": len(results.get('累計給与エラー(3-1)', [])), "累計給与エラー(3-2)": len(results.get('累計給与エラー(3-2)', []))
                    }
                    if df_retire is not None:
                        summary_errors["退職者候補（不突合）"] = len(results.get('退職者候補（退職者データ不突合）', []))
                        summary_errors["退職者データ過剰"] = len(results.get('退職者データ過剰（前期末データ不突合）', []))
                    else: 
                        summary_errors["退職者候補"] = len(results.get('退職者候補', []))
                    
                    st.session_state.summary_metrics = {**summary_info, **summary_errors}
                    
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy/mm/dd') as writer:
                        summary_list = []
                        app_title = "退職給付債務計算のための従業員データチェッカー"
                        work_time = datetime.now(tz=ZoneInfo("Asia/Tokyo")).strftime('%Y年%m月%d日 %H:%M:%S JST')
                        summary_list.extend([('アプリタイトル', app_title), ('アプリ最終更新日時', last_updated), ('作業日時', work_time), ('', '')])
                        summary_list.extend([('--- アップロードファイル ---', ''), ('前期末従業員データ', file_prev.name), ('当期末従業員データ', file_curr.name)])
                        if file_retire and retire_file_is_used: summary_list.append(('当期退職者データ', file_retire.name))
                        summary_list.append(('', ''))
                        # ... (中略: ファイル設定、列名設定などのサマリーは変更なし) ...
                        summary_list.append(('--- ファイル設定 ---', ''))
                        summary_list.append(('計算基準日', base_date.strftime('%Y/%m/%d')))
                        summary_list.extend([('前期末ヘッダーキーワード1', keyword_prev_1), ('前期末ヘッダーキーワード2', keyword_prev_2)])
                        summary_list.extend([('当期末ヘッダーキーワード1', keyword_curr_1), ('当期末ヘッダーキーワード2', keyword_curr_2)])
                        if retire_file_is_used:
                            summary_list.extend([('退職者ヘッダーキーワード1', keyword_retire_1), ('退職者ヘッダーキーワード2', keyword_retire_2)])
                        summary_list.extend([('前期末データのシート名', sheet_prev), ('当期末データのシート名', sheet_curr)])
                        if retire_file_is_used:
                            summary_list.append(('退職者データのシート名', sheet_retire))
                        summary_list.append(('', ''))
                        summary_list.append(('--- 列名設定：前期末 ---', '')); summary_list.extend([('従業員番号', col_emp_id_prev), ('入社年月日', col_hire_date_prev), ('加入年月日', col_enroll_date_prev), ('生年月日', col_birth_date_prev), ('給与1', col_salary1_prev), ('給与2', col_salary2_prev), ('給与3', col_salary3_prev), ('給与4', col_salary4_prev)])
                        summary_list.append(('--- 列名設定：当期末 ---', '')); summary_list.extend([('従業員番号', col_emp_id_curr), ('入社年月日', col_hire_date_curr), ('加入年月日', col_enroll_date_curr), ('生年月日', col_birth_date_curr), ('退職日', col_retire_date_curr), ('給与1', col_salary1_curr), ('給与2', col_salary2_curr), ('給与3', col_salary3_curr), ('給与4', col_salary4_curr)])
                        if retire_file_is_used:
                            summary_list.append(('--- 列名設定：退職者 ---', '')); summary_list.extend([('従業員番号', col_emp_id_retire), ('入社年月日', col_hire_date_retire), ('加入年月日', col_enroll_date_retire), ('生年月日', col_birth_date_retire), ('退職日', col_retire_date_retire)])
                        summary_list.append(('', ''))
                        summary_list.append(('--- 追加エラーチェック設定 (給与1,2) ---', '')); summary_list.extend([('給я減額チェック(1)', '有効' if check_salary_decrease_1 else '無効'), ('給与増加率チェック(1)', '有効' if check_salary_increase_1 else '無効')])
                        if check_salary_increase_1: summary_list.append(('└ 増加率(x1)%', increase_rate_x1))
                        summary_list.extend([('累計給与チェック(1-1)', '有効' if check_cumulative_salary_1 else '無効'), ('累計給与チェック(1-2)', '有効' if check_cumulative_salary_2 else '無効')])
                        if check_cumulative_salary_1 or check_cumulative_salary_2: summary_list.extend([('└ 月数(y1)', months_y1), ('└ 許容率(z1)%', allowance_rate_z1)])
                        summary_list.append(('', ''))
                        summary_list.append(('--- 追加エラーチェック設定 (給与3,4) ---', '')); summary_list.extend([('給与減額チェック(3)', '有効' if check_salary_decrease_3 else '無効'), ('給与増加率チェック(3)', '有効' if check_salary_increase_3 else '無効')])
                        if check_salary_increase_3: summary_list.append(('└ 増加率(x3)%', increase_rate_x3))
                        summary_list.extend([('累計給与チェック(3-1)', '有効' if check_cumulative_salary_3 else '無効'), ('累計給与チェック(3-2)', '有効' if check_cumulative_salary_4 else '無効')])
                        if check_cumulative_salary_3 or check_cumulative_salary_4: summary_list.extend([('└ 月数(y3)', months_y3), ('└ 許容率(z3)%', allowance_rate_z3)])
                        summary_list.append(('', ''))
                        
                        # --- Excelサマリーシートの表示項目を修正 ---
                        summary_list.append(('--- チェック結果サマリー ---', ''))
                        
                        # 表示順とラベルを定義
                        summary_order = [
                            ('前期従業員データ数', '前期従業員データ数', '人'), ('当期従業員データ数', '当期従業員データ数', '人'),
                            ('当期退職者データ数', '当期退職者データ数', '人'), ('キー重複', 'キー重複', '件'),
                            ('基本情報変更エラー', '基本情報変更エラー', '件'), ('日付妥当性エラー', '日付妥当性エラー', '件'),
                            ('在籍者数（凸合）', '在籍者数', '人'), 
                            ('退職者候補（不凸合＝前期のみ）', '退職者候補（不突合）' if df_retire is not None else '退職者候補', '人'),
                            ('入社者候補（不凸合＝当期のみ）', '入社者候補', '人'),
                            ('退職者データ過剰（不凸合＝前期なし）', '退職者データ過剰', '人'),
                            ('マッチした退職者（凸合）', 'マッチした退職者', '人'),
                            ('給与減額エラー(1)', '給与減額エラー(1)', '件'), ('給与増加率エラー(1)', '給与増加率エラー(1)', '件'),
                            ('累計給与エラー(1-1)', '累計給与エラー(1-1)', '件'), ('累計給与エラー(1-2)', '累計給与エラー(1-2)', '件'),
                            ('給与減額エラー(3)', '給与減額エラー(3)', '件'), ('給与増加率エラー(3)', '給与増加率エラー(3)', '件'),
                            ('累計給与エラー(3-1)', '累計給与エラー(3-1)', '件'), ('累計給与エラー(3-2)', '累計給与エラー(3-2)', '件')
                        ]

                        for label, key, unit in summary_order:
                            value = st.session_state.summary_metrics.get(key)
                            if value is not None:
                                summary_list.append((label, f"{value} {unit}"))
                        
                        df_summary = pd.DataFrame(summary_list, columns=['項目', '設定・結果'])
                        df_summary.to_excel(writer, sheet_name='サマリー', index=False)
                        summary_worksheet = writer.sheets['サマリー']; summary_worksheet.set_column('A:A', 35); summary_worksheet.set_column('B:B', 30)
                        
                        for sheet_name, df_result in results.items():
                            if not df_result.empty:
                                df_to_write = df_result.copy()
                                retiree_sheets = ['マッチした退職者', '退職者データ過剰（前期末データ不突合）']
                                sheets_to_keep_all_cols = retiree_sheets + ['基本情報変更エラー']
                                if sheet_name.startswith("日付妥当性エラー"): sheets_to_keep_all_cols.append(sheet_name)
                                
                                cols_to_drop = [c for c in ['_merge', 'retire_merge', key_col_name] if c in df_to_write.columns]
                                if sheet_name not in sheets_to_keep_all_cols:
                                    internal_cols_to_drop = [c for c in INTERNAL_COLS.values() if c in df_to_write.columns]
                                    cols_to_drop.extend(internal_cols_to_drop)
                                if cols_to_drop: df_to_write.drop(columns=cols_to_drop, inplace=True)
                                
                                df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
                                worksheet = writer.sheets[sheet_name]
                                date_col_width = 12
                                for idx, col in enumerate(df_to_write.columns):
                                    if pd.api.types.is_datetime64_any_dtype(df_to_write[col]):
                                        worksheet.set_column(idx, idx, date_col_width)
                    st.session_state.processed_data = output.getvalue()
                    st.info("ステップ7/7: 処理が完了しました。")
                    st.session_state.processing_complete = True

                except Exception as e:
                    st.error(f"処理中に予期せぬエラーが発生しました: {e}"); st.exception(e); st.stop()
        else:
            st.warning("必須項目である「前期末従業員データ」と「当期末従業員データ」をアップロードしてください。")

    if st.session_state.processing_complete:
        st.success("✅ データチェックが完了しました。")
        st.header("📊 チェック結果サマリー")
        
        # --- 画面サマリー表示を修正 ---
        summary_df_list = []
        
        # 表示順とラベル、内部キーのマッピングを定義
        summary_display_order = [
            ('前期従業員データ数', '前期従業員データ数', '人'),
            ('当期従業員データ数', '当期従業員データ数', '人'),
            ('当期退職者データ数', '当期退職者データ数', '人'),
            ('キー重複', 'キー重複', '件'),
            ('基本情報変更エラー', '基本情報変更エラー', '件'),
            ('日付妥当性エラー', '日付妥当性エラー', '件'),
            ('在籍者数（凸合）', '在籍者数', '人'),
            ('退職者候補（不凸合＝前期のみ）', '退職者候補（不突合）' if st.session_state.summary_metrics.get('退職者候補（不突合）') is not None else '退職者候補', '人'),
            ('入社者候補（不凸合＝当期のみ）', '入社者候補', '人'),
            ('退職者データ過剰（不凸合＝前期なし）', '退職者データ過剰', '人'),
            ('マッチした退職者（凸合）', 'マッチした退職者', '人'),
            ('給与減額エラー(1)', '給与減額エラー(1)', '件'),
            ('給与増加率エラー(1)', '給与増加率エラー(1)', '件'),
            ('累計給与エラー(1-1)', '累計給与エラー(1-1)', '件'),
            ('累計給与エラー(1-2)', '累計給与エラー(1-2)', '件'),
            ('給与減額エラー(3)', '給与減額エラー(3)', '件'),
            ('給与増加率エラー(3)', '給与増加率エラー(3)', '件'),
            ('累計給与エラー(3-1)', '累計給与エラー(3-1)', '件'),
            ('累計給与エラー(3-2)', '累計給与エラー(3-2)', '件'),
        ]

        for label, key, unit in summary_display_order:
            value = st.session_state.summary_metrics.get(key)
            # 値が存在する場合（0を含む）のみリストに追加
            if value is not None:
                summary_df_list.append({"項目": label, "件数/人数": f"{value} {unit}"})

        if summary_df_list:
            df_summary_display = pd.DataFrame(summary_df_list)
            st.table(df_summary_display)
            
        if st.session_state.processed_data:
            st.download_button(label="📥 チェック結果（Excelファイル）をダウンロード", data=st.session_state.processed_data, file_name="check_result.xlsx", mime="application/vnd.openxmlformats-officedocument-spreadsheetml.sheet", use_container_width=True)

if __name__ == "__main__":
    main()