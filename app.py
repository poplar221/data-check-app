import streamlit as st
import pandas as pd
import io
from datetime import datetime
import os
from zoneinfo import ZoneInfo

def find_header_and_read_excel(uploaded_file, sheet_name, keywords=['入社', '生年']):
    """
    Excelファイルからキーワードを含む行をヘッダーとして特定し、データを読み込む関数。
    """
    if uploaded_file:
        uploaded_file.seek(0)
    try:
        df_no_header = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        header_row_index = -1
        for i, row in df_no_header.iterrows():
            row_str = ''.join(map(str, row.dropna().values))
            if all(keyword in row_str for keyword in keywords):
                header_row_index = i
                break
        
        if header_row_index == -1:
            st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' でヘッダー行が見つかりませんでした。")
            return None
        
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_index)
        return df

    except Exception as e:
        st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' 読込中にエラー: {e}")
        return None

def main():
    """
    アプリケーションのメイン関数
    """
    st.set_page_config(layout="wide")

    st.title("退職給付債務計算のための従業員データチェッカー")
    try:
        mod_time = os.path.getmtime(__file__)
        jst_time = datetime.fromtimestamp(mod_time, tz=ZoneInfo("Asia/Tokyo"))
        last_updated = jst_time.strftime('%Y年%m月%d日 %H:%M:%S JST')
        st.caption(f"最終更新日時: {last_updated}")
    except Exception:
        pass
    
    st.write("前期末、当期末、退職者の従業員データ（Excelファイル）をアップロードして、データの整合性チェックを行います。")

    # --- メイン画面で先にファイルをアップロード ---
    st.subheader("📁 ファイルのアップロード")
    col1, col2, col3 = st.columns(3)
    with col1:
        file_prev = st.file_uploader("1. 前期末従業員データ (必須)", type=['xlsx'])
    with col2:
        file_curr = st.file_uploader("2. 当期末従業員データ (必須)", type=['xlsx'])
    
    # --- サイドバーで設定 ---
    with st.sidebar:
        st.header("⚙️ データ指定設定")
        st.subheader("ファイル設定")
        # シート名選択
        if file_prev:
            try:
                sheets = pd.ExcelFile(file_prev).sheet_names
                default_sheet = "従業員データフォーマット"
                index = sheets.index(default_sheet) if default_sheet in sheets else 0
                sheet_prev = st.selectbox("前期末データのシート名を選択", options=sheets, index=index)
            except Exception:
                sheet_prev = st.text_input("前期末データのシート名", "従業員データフォーマット")
        else:
            sheet_prev = st.text_input("前期末データのシート名", "従業員データフォーマット")
        
        if file_curr:
            try:
                sheets = pd.ExcelFile(file_curr).sheet_names
                default_sheet = "従業員データフォーマット"
                index = sheets.index(default_sheet) if default_sheet in sheets else 0
                sheet_curr = st.selectbox("当期末データのシート名を選択", options=sheets, index=index)
            except Exception:
                sheet_curr = st.text_input("当期末データのシート名", "従業員データフォーマット")
        else:
            sheet_curr = st.text_input("当期末データのシート名", "従業員データフォーマット")
        
        # --- ▼▼▼ ここから列名設定のロジックを変更 ▼▼▼ ---
        st.subheader("列名設定")
        NONE_OPTION = "(選択しない)"
        
        columns_prev, columns_curr = [], []
        # 前期末ファイルの列名を取得
        if file_prev and sheet_prev:
            try:
                df_cols = find_header_and_read_excel(file_prev, sheet_prev)
                if df_cols is not None:
                    columns_prev = df_cols.columns.tolist()
            except Exception:
                pass
        
        # 当期末ファイルの列名を取得
        if file_curr and sheet_curr:
            try:
                df_cols = find_header_and_read_excel(file_curr, sheet_curr)
                if df_cols is not None:
                    columns_curr = df_cols.columns.tolist()
            except Exception:
                pass
        
        # 共通のヘルパー関数
        def create_column_selector(label, default_name, columns, key):
            if columns:
                options = [NONE_OPTION] + columns
                index = options.index(default_name) if default_name in options else 0
                return st.selectbox(label, options=options, index=index, key=key)
            else:
                return st.text_input(label, default_name, key=key)

        st.info("前期末・当期末ファイルをアップロードすると、列名をドロップダウンで選択できます。")

        col_emp_id = create_column_selector("従業員番号の列", "従業員番号", columns_prev, "emp_id")
        col_hire_date = create_column_selector("入社年月日の列", "入社年月日", columns_prev, "hire_date")
        col_birth_date = create_column_selector("生年月日の列", "生年月日", columns_prev, "birth_date")
        col_salary1 = create_column_selector("給与1の列", "給与1", columns_prev, "salary1")
        col_salary2 = create_column_selector("給与2の列", "給与2", columns_prev, "salary2")
        
        # 当期末データの列名から「退職日」を選択
        col_retire_date = create_column_selector("退職日の列（当期末データ内）", "退職日", columns_curr, "retire_date")
        
        # --- ▲▲▲ 列名設定のロジック変更ここまで ▲▲▲ ---

        # 退職者データのシート名選択（従来通り）
        sheet_retire = st.text_input("退職者データのシート名", "退職者データフォーマット")
        
        st.header("✔️ 追加エラーチェック設定")
        check_salary_decrease = st.checkbox("給与減額チェックを有効にする", True)
        # (以下、変更なし)
        check_salary_increase = st.checkbox("給与増加率チェックを有効にする", True)
        increase_rate_x = st.text_input("増加率(x)%", value="5")
        check_cumulative_salary = st.checkbox("累計給与チェック1を有効にする", True)
        months_y = st.selectbox("月数(y)", ("1", "12"), index=0)
        check_cumulative_salary2 = st.checkbox("累計給与チェック2を有効にする", True)
        allowance_rate_z = st.text_input("許容率(z)%", value="0")

    # --- 退職者ファイルアップローダーを、退職日列の選択状態に応じて無効化 ---
    retire_uploader_disabled = (col_retire_date != NONE_OPTION)
    with col3:
        file_retire = st.file_uploader(
            "3. 当期退職者データ (任意)", type=['xlsx'],
            disabled=retire_uploader_disabled,
            help="サイドバーで「退職日」列を指定した場合、このアップローダーは無効になります。"
        )

    if st.button("チェック開始", use_container_width=True, type="primary"):
        if file_prev and file_curr:
            with st.spinner('データチェックを実行中です...'):
                try:
                    st.info("ステップ1/7: Excelファイルを読み込んでいます...")
                    df_prev = find_header_and_read_excel(file_prev, sheet_prev)
                    df_curr = find_header_and_read_excel(file_curr, sheet_curr)
                    df_retire = None
                    if df_prev is None or df_curr is None:
                        st.error("必須ファイル（前期末・当期末）の読み込みに失敗しました。")
                        st.stop()

                    # --- ▼▼▼ 退職者の決定ロジックを変更 ▼▼▼ ---
                    # 「退職日」列が指定されていれば、当期末データから退職者を抽出
                    if col_retire_date != NONE_OPTION and col_retire_date in df_curr.columns:
                        st.info(f"ステップ1.5/7: 当期末データから「{col_retire_date}」列を基に退職者を抽出...")
                        df_curr[col_retire_date] = pd.to_datetime(df_curr[col_retire_date].astype(str), errors='coerce')
                        retiree_mask = df_curr[col_retire_date].notna()
                        df_retire = df_curr[retiree_mask].copy() # 退職者データを生成
                        df_curr = df_curr[~retiree_mask].copy() # 元のデータからは削除
                        if not df_retire.empty:
                             st.success(f"{len(df_retire)}名の退職者を当期末データから抽出し、在籍者から除外しました。")
                        else:
                             st.warning(f"「{col_retire_date}」列に有効な日付が見つかりませんでした。")
                    # 「退職日」列が指定されていない場合、従来通りアップロードされたファイルを使用
                    elif file_retire:
                        df_retire = find_header_and_read_excel(file_retire, sheet_retire)
                    # --- ▲▲▲ 退職者の決定ロジック変更ここまで ▲▲▲ ---

                    st.info("ステップ2/7: マッチングキーを決定しています...")
                    # (これ以降のメイン処理ロジックは、ほぼ変更なし)
                    # (エラー処理の可読性のため、一部リファクタリング)
                    use_emp_id_key = (col_emp_id in df_prev.columns and col_emp_id in df_curr.columns)
                    dataframes = {'前期末': df_prev, '当期末': df_curr}
                    if df_retire is not None:
                        use_emp_id_key = use_emp_id_key and (col_emp_id in df_retire.columns)
                        dataframes['退職者'] = df_retire
                    key_col_name = '_key'
                    for name, df in dataframes.items():
                        if not use_emp_id_key:
                             if not {col_hire_date, col_birth_date}.issubset(df.columns):
                                st.error(f"代替キー（{col_hire_date}, {col_birth_date}）が'{name}'データに存在しません。")
                                st.stop()
                             df[col_hire_date] = pd.to_datetime(df[col_hire_date].astype(str), errors='coerce')
                             df[col_birth_date] = pd.to_datetime(df[col_birth_date].astype(str), errors='coerce')
                             df[key_col_name] = df[col_hire_date].dt.strftime('%Y%m%d').fillna('NODATE') + '_' + df[col_birth_date].dt.strftime('%Y%m%d').fillna('NODATE')
                        else:
                             df[key_col_name] = df[col_emp_id].astype(str)
                    key_type = "従業員番号" if use_emp_id_key else "入社年月日 + 生年月日"
                    st.success(f"マッチングキーとして '{key_type}' を使用します。")

                    results = {}
                    st.info("ステップ3/7: 基本エラーチェック...")
                    # (エラーチェック、マッチング、Excel出力等のロジックは変更なし)
                    for name, df in dataframes.items():
                        duplicates = df[df[key_col_name].duplicated(keep=False)]
                        results[f'キー重複_{name}'] = duplicates.sort_values(by=key_col_name)
                    for name, df in {'前期末': df_prev, '当期末': df_curr}.items():
                        if col_hire_date in df.columns and col_birth_date in df.columns:
                            df_copy = df.copy()
                            df_copy[col_hire_date] = pd.to_datetime(df_copy[col_hire_date].astype(str), errors='coerce')
                            df_copy[col_birth_date] = pd.to_datetime(df_copy[col_birth_date].astype(str), errors='coerce')
                            valid_dates = df_copy.dropna(subset=[col_hire_date, col_birth_date])
                            if not valid_dates.empty:
                                age = (valid_dates[col_hire_date] - valid_dates[col_birth_date]).dt.days / 365.25
                                invalid_age = valid_dates[(age < 15) | (age >= 90)]
                                results[f'日付妥当性エラー_{name}'] = df.loc[invalid_age.index]

                    st.info("ステップ4/7: 在籍者・退職者・入社者の照合...")
                    merged_st = pd.merge(df_prev, df_curr, on=key_col_name, how='outer', suffixes=('_前期', '_当期'), indicator=True)
                    retiree_candidates = merged_st[merged_st['_merge'] == 'left_only'].copy()
                    new_hires = merged_st[merged_st['_merge'] == 'right_only'].copy()
                    continuing_employees = merged_st[merged_st['_merge'] == 'both'].copy()
                    results['入社者候補'] = new_hires
                    if df_retire is not None:
                        st.info("ステップ4.5/7: 退職者データの照合...")
                        merged_retire = pd.merge(retiree_candidates[[key_col_name]], df_retire, on=key_col_name, how='outer', indicator='retire_merge')
                        retire_unmatched = retiree_candidates[retiree_candidates[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'left_only'][key_col_name])]
                        retire_extra = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'right_only'][key_col_name])]
                        retire_matched = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'both'][key_col_name])]
                        results['退職者候補（退職者データ不突合）'] = retire_unmatched
                        results['退職者データ過剰（前期末データ不突合）'] = retire_extra
                        results['マッチした退職者'] = retire_matched
                    else:
                        results['退職者候補'] = retiree_candidates
                    results['在籍者'] = continuing_employees

                    st.info("ステップ5/7: 追加エラーチェック...")
                    required_salary_cols = {f'{col_salary1}_前期', f'{col_salary1}_当期', f'{col_salary2}_前期', f'{col_salary2}_当期'}
                    if not required_salary_cols.issubset(continuing_employees.columns):
                        st.warning(f"給与列（{col_salary1}, {col_salary2}）がないため追加チェックはスキップ。")
                    else:
                        for col in required_salary_cols:
                            continuing_employees[col] = pd.to_numeric(continuing_employees[col], errors='coerce')
                        check_df = continuing_employees.dropna(subset=required_salary_cols).copy()
                        if check_salary_decrease: results['給与減額エラー'] = check_df[check_df[f'{col_salary1}_当期'] < check_df[f'{col_salary1}_前期']]
                        if check_salary_increase:
                            try:
                                x = float(increase_rate_x)
                                results['給与増加率エラー'] = check_df[check_df[f'{col_salary1}_当期'] >= check_df[f'{col_salary1}_前期'] * (1 + x / 100)]
                            except ValueError: st.warning("給与増加率(x)が無効な数値のためスキップしました。")
                        if check_cumulative_salary:
                            try:
                                y = int(months_y)
                                results['累計給与エラー1'] = check_df[check_df[f'{col_salary2}_当期'] < check_df[f'{col_salary2}_前期'] + check_df[f'{col_salary1}_前期'] * y]
                            except ValueError: st.warning("月数(y)が無効な数値のためスキップしました。")
                        if check_cumulative_salary2:
                            try:
                                y = int(months_y)
                                z = float(allowance_rate_z)
                                upper_limit = (check_df[f'{col_salary2}_前期'] + check_df[f'{col_salary1}_前期'] * y) * (1 + z / 100)
                                results['累計給与エラー2'] = check_df[check_df[f'{col_salary2}_当期'] > upper_limit]
                            except ValueError: st.warning("月数(y)または許容率(z)が無効な数値のためスキップしました。")
                    
                    summary_info = {"前期末従業員数": len(df_prev), "当期末従業員数": len(df_curr), "在籍者数": len(results.get('在籍者', []))}
                    if df_retire is not None:
                        summary_info["当期退職者数"] = len(df_retire)
                    summary_errors = {"キー重複": sum(len(df) for name, df in results.items() if 'キー重複' in name), "日付妥当性エラー": sum(len(df) for name, df in results.items() if '日付妥当性' in name), "入社者候補": len(results.get('入社者候補', [])), "給与減額エラー": len(results.get('給与減額エラー', [])), "給与増加率エラー": len(results.get('給与増加率エラー', [])), "累計給与エラー1": len(results.get('累計給与エラー1', [])), "累計給与エラー2": len(results.get('累計給与エラー2', []))}
                    if df_retire is not None:
                        summary_errors["退職者候補（不突合）"] = len(results.get('退職者候補（退職者データ不突合）', []))
                        summary_errors["退職者データ過剰"] = len(results.get('退職者データ過剰（前期末データ不突合）', []))
                    else:
                        summary_errors["退職者候補"] = len(results.get('退職者候補', []))
                    summary_metrics = {**summary_info, **summary_errors}

                    st.info("ステップ6/7: 結果をExcelファイルにまとめています...")
                    # ... Excel出力、ダウンロードボタンのロジックは変更なし ...
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy/mm/dd') as writer:
                        # (サマリーシート作成)
                        summary_list = []
                        app_title = "退職給付債務計算のための従業員データチェッカー"
                        work_time = datetime.now(tz=ZoneInfo("Asia/Tokyo")).strftime('%Y年%m月%d日 %H:%M:%S JST')
                        summary_list.append(('アプリタイトル', app_title)); summary_list.append(('アプリ最終更新日時', last_updated)); summary_list.append(('作業日時', work_time)); summary_list.append(('', ''))
                        summary_list.append(('--- アップロードファイル ---', '')); summary_list.append(('前期末従業員データ', file_prev.name)); summary_list.append(('当期末従業員データ', file_curr.name))
                        if file_retire: summary_list.append(('当期退職者データ', file_retire.name))
                        summary_list.append(('', ''))
                        summary_list.append(('--- ファイル設定 ---', '')); summary_list.append(('前期末データのシート名', sheet_prev)); summary_list.append(('当期末データのシート名', sheet_curr)); summary_list.append(('退職者データのシート名', sheet_retire)); summary_list.append(('', ''))
                        summary_list.append(('--- 列名設定 ---', '')); summary_list.append(('従業員番号の列名', col_emp_id)); summary_list.append(('入社年月日の列名', col_hire_date)); summary_list.append(('生年月日の列名', col_birth_date)); summary_list.append(('給与1の列名', col_salary1)); summary_list.append(('給与2の列名', col_salary2)); summary_list.append(('退職日の列名', col_retire_date)); summary_list.append(('', ''))
                        summary_list.append(('--- 追加エラーチェック設定 ---', '')); summary_list.append(('給与減額チェック', '有効' if check_salary_decrease else '無効')); summary_list.append(('給与増加率チェック', '有効' if check_salary_increase else '無効'))
                        if check_salary_increase: summary_list.append(('└ 増加率(x)%', increase_rate_x))
                        summary_list.append(('累計給与チェック1', '有効' if check_cumulative_salary else '無効'))
                        if check_cumulative_salary: summary_list.append(('└ 月数(y)', months_y))
                        summary_list.append(('累計給与チェック2', '有効' if check_cumulative_salary2 else '無効'))
                        if check_cumulative_salary2: summary_list.append(('└ 許容率(z)%', allowance_rate_z))
                        summary_list.append(('', ''))
                        summary_list.append(('--- チェック結果サマリー ---', ''))
                        info_labels = ["前期末従業員数", "当期末従業員数", "在籍者数", "当期退職者数"]
                        def format_value(label, value):
                            unit = "人" if label in info_labels else "件"
                            return f"{value} {unit}"
                        summary_list.append(('前期末従業員数', format_value('前期末従業員数', summary_metrics.get('前期末従業員数', 0)))); summary_list.append(('当期末従業員数', format_value('当期末従業員数', summary_metrics.get('当期末従業員数', 0)))); summary_list.append(('在籍者数', format_value('在籍者数', summary_metrics.get('在籍者数', 0))))
                        if df_retire is not None: summary_list.append(('退職者候補（不突合）', format_value('退職者候補（不突合）', summary_metrics.get('退職者候補（不突合）', 0))))
                        else: summary_list.append(('退職者候補', format_value('退職者候補', summary_metrics.get('退職者候補', 0))))
                        summary_list.append(('入社者候補', format_value('入社者候補', summary_metrics.get('入社者候補', 0))))
                        if df_retire is not None:
                            summary_list.append(('当期退職者数', format_value('当期退職者数', summary_metrics.get('当期退職者数', 0))))
                            summary_list.append(('退職者データ過剰', format_value('退職者データ過剰', summary_metrics.get('退職者データ過剰', 0))))
                        summary_list.append(('キー重複', format_value('キー重複', summary_metrics.get('キー重複', 0)))); summary_list.append(('日付妥当性エラー', format_value('日付妥当性エラー', summary_metrics.get('日付妥当性エラー', 0)))); summary_list.append(('給与減額エラー', format_value('給与減額エラー', summary_metrics.get('給与減額エラー', 0)))); summary_list.append(('給与増加率エラー', format_value('給与増加率エラー', summary_metrics.get('給与増加率エラー', 0)))); summary_list.append(('累計給与エラー1', format_value('累計給与エラー1', summary_metrics.get('累計給与エラー1', 0)))); summary_list.append(('累計給与エラー2', format_value('累計給与エラー2', summary_metrics.get('累計給与エラー2', 0))))
                        df_summary = pd.DataFrame(summary_list, columns=['項目', '設定・結果'])
                        
                        df_summary.to_excel(writer, sheet_name='サマリー', index=False)
                        summary_worksheet = writer.sheets['サマリー']
                        summary_worksheet.set_column('A:A', 35)
                        summary_worksheet.set_column('B:B', 30)
                        for sheet_name, df_result in results.items():
                            if not df_result.empty:
                                df_to_write = df_result.copy()
                                cols_to_drop = [c for c in ['_merge', 'retire_merge', key_col_name] if c in df_to_write.columns]
                                if cols_to_drop:
                                    df_to_write.drop(columns=cols_to_drop, inplace=True)
                                df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
                                worksheet = writer.sheets[sheet_name]
                                date_col_width = 12
                                for idx, col_name in enumerate(df_to_write.columns):
                                    if col_hire_date in col_name or col_birth_date in col_name or (col_retire_date != NONE_OPTION and col_retire_date in col_name):
                                        worksheet.set_column(idx, idx, date_col_width)
                    processed_data = output.getvalue()
                    st.info("ステップ7/7: 処理が完了しました。")

                except Exception as e:
                    st.error(f"処理中に予期せぬエラーが発生しました: {e}")
                    st.stop()

            st.success("✅ データチェックが完了しました。")
            st.header("📊 チェック結果サマリー")
            cols = st.columns(3)
            col_idx = 0
            info_labels = ["前期末従業員数", "当期末従業員数", "在籍者数", "当期退職者数"]
            for label, value in summary_metrics.items():
                if label in info_labels:
                    cols[col_idx].metric(label, f"{value} 人")
                elif value > 0:
                    cols[col_idx].metric(label, f"{value} 件", delta=f"{value} 件のエラー", delta_color="inverse")
                else:
                    cols[col_idx].metric(label, f"{value} 件")
                col_idx = (col_idx + 1) % 3
            st.download_button(label="📥 チェック結果（Excelファイル）をダウンロード", data=processed_data, file_name="check_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        else:
            st.warning("必須項目である「前期末従業員データ」と「当期末従業員データ」をアップロードしてください。")

if __name__ == "__main__":
    main()