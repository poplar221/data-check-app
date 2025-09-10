import streamlit as st
import pandas as pd
import io
from datetime import datetime

def find_header_and_read_excel(uploaded_file, sheet_name, keywords=['入社', '生年']):
    """
    Excelファイルからキーワードを含む行をヘッダーとして特定し、データを読み込む関数。

    Args:
        uploaded_file: st.file_uploaderからアップロードされたファイルオブジェクト。
        sheet_name (str): 読み込むシート名。
        keywords (list): ヘッダー行に含まれるべきキーワードのリスト。

    Returns:
        pandas.DataFrame: 読み込まれたデータフレーム。ヘッダーが見つからない場合はNone。
    """
    try:
        # ヘッダーなしで一度読み込み、ヘッダー行を探す
        df_no_header = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
        header_row_index = -1
        for i, row in df_no_header.iterrows():
            # 行の値を文字列として結合し、キーワードが含まれているかチェック
            row_str = ''.join(map(str, row.dropna().values))
            if all(keyword in row_str for keyword in keywords):
                header_row_index = i
                break
        
        if header_row_index == -1:
            st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' でヘッダー行（キーワード: {keywords}）が見つかりませんでした。")
            return None
        
        # 特定したヘッダー行を使って再度データを読み込む
        # seek(0)でファイルの読み取り位置を先頭に戻す
        uploaded_file.seek(0)
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=header_row_index)
        return df

    except Exception as e:
        st.error(f"ファイル '{uploaded_file.name}' のシート '{sheet_name}' の読み込み中にエラーが発生しました: {e}")
        return None

def main():
    """
    アプリケーションのメイン関数
    """
    # 1. UI（ユーザーインターフェース）の仕様
    # -------------------------------------------------------------------------
    st.set_page_config(layout="wide")

    st.title("退職給付債務計算のための従業員データチェッカー")
    st.write("前期末、当期末、退職者の従業員データ（Excelファイル）をアップロードして、データの整合性チェックを行います。")

    # --- サイドバーの設定 ---
    with st.sidebar:
        st.header("⚙️ データ指定設定")

        st.subheader("ファイル設定")
        sheet_prev = st.text_input("前期末データのシート名", "従業員データフォーマット")
        sheet_curr = st.text_input("当期末データのシート名", "従業員データフォーマット")
        sheet_retire = st.text_input("退職者データのシート名", "退職者データフォーマット")

        st.subheader("列名設定")
        col_emp_id = st.text_input("従業員番号の列名", "従業員番号")
        col_hire_date = st.text_input("入社年月日の列名", "入社年月日")
        col_birth_date = st.text_input("生年月日の列名", "生年月日")
        col_salary1 = st.text_input("給与1の列名（当期・前期比較用）", "給与１")
        col_salary2 = st.text_input("給与2の列名（累計チェック用）", "給与２")

        st.header("✔️ 追加エラーチェック設定")
        
        check_salary_decrease = st.checkbox("給与減額チェックを有効にする", True)
        
        check_salary_increase = st.checkbox("給与増加率チェックを有効にする", True)
        increase_rate_x = st.text_input("増加率(x)%", value="5")
        
        check_cumulative_salary = st.checkbox("累計給与チェックを有効にする", True)
        months_y = st.selectbox("月数(y)", ("1", "12"), index=0)
        
        check_cumulative_salary2 = st.checkbox("累計給与チェック2を有効にする", True)
        allowance_rate_z = st.text_input("許容率(z)%", value="0")


    # --- メイン画面のファイルアップロード ---
    st.subheader("📁 ファイルのアップロード")
    col1, col2, col3 = st.columns(3)
    with col1:
        file_prev = st.file_uploader("1. 前期末従業員データ", type=['xlsx'])
    with col2:
        file_curr = st.file_uploader("2. 当期末従業員データ", type=['xlsx'])
    with col3:
        file_retire = st.file_uploader("3. 当期退職者データ", type=['xlsx'])

    # --- 処理開始ボタン ---
    if st.button("チェック開始", use_container_width=True, type="primary"):
        if file_prev and file_curr and file_retire:
            
            with st.spinner('データチェックを実行中です...'):
                try:
                    # 2. データ処理の機能要件
                    # -------------------------------------------------------------------------
                    
                    # --- データ読み込み ---
                    st.info("ステップ1/7: Excelファイルを読み込んでいます...")
                    df_prev = find_header_and_read_excel(file_prev, sheet_prev)
                    df_curr = find_header_and_read_excel(file_curr, sheet_curr)
                    df_retire = find_header_and_read_excel(file_retire, sheet_retire)

                    if df_prev is None or df_curr is None or df_retire is None:
                        st.error("ファイルの読み込みに失敗しました。処理を中断します。")
                        st.stop()
                        
                    # --- マッチングキーの採用 ---
                    st.info("ステップ2/7: マッチングキーを決定しています...")
                    use_emp_id_key = (col_emp_id in df_prev.columns and
                                      col_emp_id in df_curr.columns and
                                      col_emp_id in df_retire.columns)

                    key_col_name = '_key'
                    dataframes = {'前期末': df_prev, '当期末': df_curr, '退職者': df_retire}
                    
                    for name, df in dataframes.items():
                        # 日付列などをチェックする前に必須列の存在を確認
                        required_cols_base = {col_hire_date, col_birth_date}
                        if not use_emp_id_key: # 従業員番号を使わない場合は必須
                            if not required_cols_base.issubset(df.columns):
                                st.error(f"代替キー（{col_hire_date}, {col_birth_date}）が'{name}'データに存在しないため、処理を中断します。")
                                st.stop()
                        
                        if use_emp_id_key:
                            df[key_col_name] = df[col_emp_id].astype(str)
                        else:
                            # 修正箇所1: YYYYMMDD形式を正しく読み込むために format を指定
                            df[col_hire_date] = pd.to_datetime(df[col_hire_date], format='%Y%m%d', errors='coerce')
                            df[col_birth_date] = pd.to_datetime(df[col_birth_date], format='%Y%m%d', errors='coerce')
                            df[key_col_name] = df[col_hire_date].dt.strftime('%Y%m%d').fillna('NODATE') + '_' + df[col_birth_date].dt.strftime('%Y%m%d').fillna('NODATE')
                    
                    key_type = "従業員番号" if use_emp_id_key else "入社年月日 + 生年月日"
                    st.success(f"マッチングキーとして '{key_type}' を使用します。")

                    # --- エラーチェック項目の実行 ---
                    results = {}
                    st.info("ステップ3/7: 基本エラーチェック（キー重複・日付妥当性）を実行しています...")

                    # キーの重複チェック
                    for name, df in dataframes.items():
                        duplicates = df[df[key_col_name].duplicated(keep=False)]
                        results[f'キー重複_{name}'] = duplicates.sort_values(by=key_col_name)

                    # 日付の妥当性チェック
                    for name, df in {'前期末': df_prev, '当期末': df_curr}.items():
                        if col_hire_date in df.columns and col_birth_date in df.columns:
                            df_copy = df.copy()
                            # 修正箇所2: YYYYMMDD形式を正しく読み込むために format を指定
                            df_copy[col_hire_date] = pd.to_datetime(df_copy[col_hire_date], format='%Y%m%d', errors='coerce')
                            df_copy[col_birth_date] = pd.to_datetime(df_copy[col_birth_date], format='%Y%m%d', errors='coerce')
                            
                            # NaT（無効な日付）を除外
                            valid_dates = df_copy.dropna(subset=[col_hire_date, col_birth_date])
                            
                            if not valid_dates.empty:
                                age = (valid_dates[col_hire_date] - valid_dates[col_birth_date]).dt.days / 365.25
                                invalid_age = valid_dates[(age < 15) | (age >= 90)]
                                results[f'日付妥当性エラー_{name}'] = df.loc[invalid_age.index]

                    # --- マッチングと分類 ---
                    st.info("ステップ4/7: 在籍者・退職者・入社者の照合を行っています...")
                    
                    # 在籍者照合
                    merged_st = pd.merge(
                        df_prev, df_curr, on=key_col_name, how='outer', 
                        suffixes=('_前期', '_当期'), indicator=True
                    )
                    
                    retiree_candidates = merged_st[merged_st['_merge'] == 'left_only'].copy()
                    new_hires = merged_st[merged_st['_merge'] == 'right_only'].copy()
                    continuing_employees = merged_st[merged_st['_merge'] == 'both'].copy()

                    # 退職者照合
                    merged_retire = pd.merge(
                        retiree_candidates[[key_col_name]], df_retire, on=key_col_name, 
                        how='outer', indicator='retire_merge'
                    )

                    # 分類
                    retire_unmatched = retiree_candidates[retiree_candidates[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'left_only'][key_col_name])]
                    retire_extra = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'right_only'][key_col_name])]
                    retire_matched = df_retire[df_retire[key_col_name].isin(merged_retire[merged_retire['retire_merge'] == 'both'][key_col_name])]
                    
                    results['入社者候補'] = new_hires
                    results['退職者候補（退職者データ不突合）'] = retire_unmatched
                    results['退職者データ過剰（前期末データ不突合）'] = retire_extra
                    results['マッチした退職者'] = retire_matched
                    results['在籍者'] = continuing_employees


                    # --- 追加エラーチェック項目の実行 ---
                    st.info("ステップ5/7: 追加エラーチェック（給与関連）を実行しています...")

                    # チェックに必要な列が存在するか確認
                    required_salary_cols = {f'{col_salary1}_前期', f'{col_salary1}_当期', 
                                            f'{col_salary2}_前期', f'{col_salary2}_当期'}
                    
                    if not required_salary_cols.issubset(continuing_employees.columns):
                        st.warning(f"給与列（{col_salary1}, {col_salary2}）が前期・当期データにないため、追加エラーチェックはスキップされます。")
                    else:
                        # 給与列を数値型に変換（エラーはNaNにする）
                        for col in required_salary_cols:
                            continuing_employees[col] = pd.to_numeric(continuing_employees[col], errors='coerce')
                        
                        # NaNを除外したデータでチェック
                        check_df = continuing_employees.dropna(subset=required_salary_cols).copy()

                        if check_salary_decrease:
                            results['給与減額エラー'] = check_df[check_df[f'{col_salary1}_当期'] < check_df[f'{col_salary1}_前期']]

                        if check_salary_increase:
                            try:
                                x = float(increase_rate_x)
                                results['給与増加率エラー'] = check_df[check_df[f'{col_salary1}_当期'] >= check_df[f'{col_salary1}_前期'] * (1 + x / 100)]
                            except ValueError:
                                st.warning("給与増加率(x)が無効な数値です。このチェックはスキップされました。")
                        
                        if check_cumulative_salary:
                            try:
                                y = int(months_y)
                                results['累計給与エラー1'] = check_df[check_df[f'{col_salary2}_当期'] < check_df[f'{col_salary2}_前期'] + check_df[f'{col_salary1}_前期'] * y]
                            except ValueError:
                                st.warning("月数(y)が無効な数値です。このチェックはスキップされました。")
                        
                        if check_cumulative_salary2:
                            try:
                                y = int(months_y)
                                z = float(allowance_rate_z)
                                upper_limit = (check_df[f'{col_salary2}_前期'] + check_df[f'{col_salary1}_前期'] * y) * (1 + z / 100)
                                results['累計給与エラー2'] = check_df[check_df[f'{col_salary2}_当期'] > upper_limit]
                            except ValueError:
                                st.warning("月数(y)または許容率(z)が無効な数値です。このチェックはスキップされました。")

                    # --- 結果の出力準備 ---
                    st.info("ステップ6/7: 結果をExcelファイルにまとめています...")
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        for sheet_name, df in results.items():
                            if not df.empty:
                                # 不要な列(_merge, _keyなど)を削除して出力
                                df_to_write = df.copy()
                                cols_to_drop = [c for c in ['_merge', 'retire_merge', key_col_name] if c in df_to_write.columns]
                                if cols_to_drop:
                                    df_to_write.drop(columns=cols_to_drop, inplace=True)
                                df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    processed_data = output.getvalue()
                    st.info("ステップ7/7: 処理が完了しました。")

                except Exception as e:
                    st.error(f"処理中に予期せぬエラーが発生しました: {e}")
                    st.stop()


            # 3. 実行中のフィードバックと結果表示
            # -------------------------------------------------------------------------
            st.success("✅ データチェックが完了しました。")
            
            # --- サマリー表示 ---
            st.header("📊 チェック結果サマリー")
            
            summary_metrics = {
                "キー重複": sum(len(df) for name, df in results.items() if 'キー重複' in name),
                "日付妥当性エラー": sum(len(df) for name, df in results.items() if '日付妥当性' in name),
                "退職者候補（不突合）": len(results.get('退職者候補（退職者データ不突合）', [])),
                "入社者候補": len(results.get('入社者候補', [])),
                "退職者データ過剰": len(results.get('退職者データ過剰（前期末データ不突合）', [])),
                "給与減額エラー": len(results.get('給与減額エラー', [])),
                "給与増加率エラー": len(results.get('給与増加率エラー', [])),
                "累計給与エラー1": len(results.get('累計給与エラー1', [])),
                "累計給与エラー2": len(results.get('累計給与エラー2', [])),
            }

            # 3列でメトリクスを表示
            cols = st.columns(3)
            col_idx = 0
            for label, value in summary_metrics.items():
                if value > 0:
                    cols[col_idx].metric(label, f"{value} 件", delta=f"{value} 件のエラー", delta_color="inverse")
                else:
                    cols[col_idx].metric(label, f"{value} 件")
                col_idx = (col_idx + 1) % 3

            # --- ダウンロードボタン ---
            st.download_button(
                label="📥 チェック結果（Excelファイル）をダウンロード",
                data=processed_data,
                file_name="check_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        else:
            st.warning("3つのファイルをすべてアップロードしてください。")

if __name__ == "__main__":
    main()