# app.py

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# ------------------------------------------------------------------------------------
# データ処理関数
# ------------------------------------------------------------------------------------

def find_header_and_read_excel(file_path, sheet_name, keywords):
    """
    Excelファイルを読み込み、指定されたキーワードが含まれる行をヘッダーとして特定し、
    その行からデータを読み込んでDataFrameとして返します。
    """
    if file_path is None:
        return None
    try:
        xls = pd.ExcelFile(file_path)
        if sheet_name not in xls.sheet_names:
            st.error(f"⚠️ '{file_path.name}' に '{sheet_name}' という名前のシートが見つかりません。")
            return None
            
        df_temp = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row_index = -1
        for i, row in df_temp.iterrows():
            row_str = ''.join(map(str, row.dropna().values))
            if all(keyword in row_str for keyword in keywords):
                header_row_index = i
                break
        
        if header_row_index != -1:
            st.info(f"📄 '{file_path.name}' の '{sheet_name}' シートからヘッダーを {header_row_index + 1} 行目で発見しました。")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_index)
            return df
        else:
            st.error(f"⚠️ '{file_path.name}' の '{sheet_name}' シートでキーワードを含むヘッダー行が見つかりませんでした。")
            return None
    except Exception as e:
        st.error(f"❌エラー: '{file_path.name}' の '{sheet_name}' シート読み込み中に問題が発生しました: {e}")
        return None

# ★★★ 修正箇所: 日付を「文字列」としてフォーマットするように変更 ★★★
def format_date_columns(df, date_cols):
    """
    指定された日付列をYYYYMMDD形式の文字列に変換する
    """
    df_copy = df.copy()
    for col in date_cols:
        if col in df_copy.columns and pd.api.types.is_datetime64_any_dtype(df_copy[col]):
            # NaTは空文字に、それ以外はYYYYMMDD形式の文字列に変換
            df_copy[col] = df_copy[col].dt.strftime('%Y%m%d').fillna('')
    return df_copy


def data_check_and_matching(df_zenki, df_touki, df_taishoku, col_employee_id, col_nyusha, col_seinengappi, col_salary1, col_salary2, error_check_config):
    """
    データチェックとマッチング処理を行い、結果をExcelファイルとしてメモリ上に保存し、
    サマリー情報も一緒に返します。
    """
    summary = {}
    
    st.write("---")
    st.subheader("処理状況")
    
    with st.spinner("STEP 1: 前処理とキー作成を実行中..."):
        dfz = df_zenki.copy()
        dft = df_touki.copy()
        dftai = df_taishoku.copy()
        
        all_dfs = [dfz, dft, dftai]
        
        for df in [dfz, dft]:
            if col_salary1 in df.columns:
                df[col_salary1] = pd.to_numeric(df[col_salary1], errors='coerce')
            if col_salary2 in df.columns:
                df[col_salary2] = pd.to_numeric(df[col_salary2], errors='coerce')

        for df in all_dfs:
            if col_nyusha in df.columns:
                df[col_nyusha] = pd.to_datetime(df[col_nyusha].astype(str), errors='coerce')
            if col_seinengappi in df.columns:
                df[col_seinengappi] = pd.to_datetime(df[col_seinengappi].astype(str), errors='coerce')
            if '退職年月日' in df.columns:
                 df['退職年月日'] = pd.to_datetime(df['退職年月日'].astype(str), errors='coerce')
            if '支給日' in df.columns:
                 df['支給日'] = pd.to_datetime(df['支給日'].astype(str), errors='coerce')

        if col_employee_id in dfz.columns and col_employee_id in dft.columns and col_employee_id in dftai.columns:
            st.info(f"🔑 「{col_employee_id}」をキーとして使用します。")
            for df in all_dfs:
                if 'key' not in df.columns:
                    df['key'] = df[col_employee_id].astype(str)
        else:
            st.info(f"🔑 「入社年月日」と「生年月日」の連結文字列をキーとして使用します。")
            for df in all_dfs:
                if 'key' not in df.columns and col_nyusha in df.columns and col_seinengappi in df.columns:
                    df['key'] = df[col_nyusha].dt.strftime('%Y%m%d').astype(str).str.cat(df[col_seinengappi].dt.strftime('%Y%m%d').astype(str), sep='_')
        st.success("✅ キーの作成が完了しました。")

    with st.spinner("STEP 2: 基本エラーチェックを実行中..."):
        zenki_duplicates = dfz[dfz.duplicated(subset=['key'], keep=False)]
        touki_duplicates = dft[dft.duplicated(subset=['key'], keep=False)]
        summary['zenki_duplicates'] = len(zenki_duplicates)
        summary['touki_duplicates'] = len(touki_duplicates)
        
        zenki_age_errors = pd.DataFrame()
        touki_age_errors = pd.DataFrame()
        if col_nyusha in dfz.columns and col_seinengappi in dfz.columns:
            valid_dates_z = dfz.dropna(subset=[col_nyusha, col_seinengappi])
            if not valid_dates_z.empty:
                days_diff_z = (valid_dates_z[col_nyusha] - valid_dates_z[col_seinengappi]).dt.days
                dfz.loc[valid_dates_z.index, 'age_at_hire'] = (days_diff_z / 365.25)
                zenki_age_errors = dfz[(dfz['age_at_hire'] < 15) | (dfz['age_at_hire'] >= 90)]

        if col_nyusha in dft.columns and col_seinengappi in dft.columns:
            valid_dates_t = dft.dropna(subset=[col_nyusha, col_seinengappi])
            if not valid_dates_t.empty:
                days_diff_t = (valid_dates_t[col_nyusha] - valid_dates_t[col_seinengappi]).dt.days
                dft.loc[valid_dates_t.index, 'age_at_hire'] = (days_diff_t / 365.25)
                touki_age_errors = dft[(dft['age_at_hire'] < 15) | (dft['age_at_hire'] >= 90)]
        summary['zenki_age_errors'] = len(zenki_age_errors)
        summary['touki_age_errors'] = len(touki_age_errors)
        st.success("✅ 基本エラーチェックが完了しました。")
    
    with st.spinner("STEP 3: 在籍者・退職者マッチングを実行中..."):
        merged_df = pd.merge(dfz, dft, on='key', how='outer', indicator=True, suffixes=('_zenki', '_touki'))
        
        only_zenki_full = merged_df[merged_df['_merge'] == 'left_only']
        only_touki_full = merged_df[merged_df['_merge'] == 'right_only']
        both_full = merged_df[merged_df['_merge'] == 'both']

        summary['only_zenki'] = len(only_zenki_full)
        summary['only_touki'] = len(only_touki_full)
        summary['both'] = len(both_full)

        retiree_merged = pd.merge(only_zenki_full[['key']], dftai[['key']], on='key', how='outer', indicator='retiree_check')
        retiree_missing_keys = retiree_merged[retiree_merged['retiree_check'] == 'left_only']
        retiree_not_candidate_keys = retiree_merged[retiree_merged['retiree_check'] == 'right_only']
        retiree_correct_keys = retiree_merged[retiree_merged['retiree_check'] == 'both']

        retiree_missing_full = pd.merge(retiree_missing_keys, dfz, on='key', how='left')
        retiree_not_candidate_full = pd.merge(retiree_not_candidate_keys, dftai, on='key', how='left')
        retiree_correct_full = pd.merge(retiree_correct_keys, dftai, on='key', how='left')
        summary['retiree_missing'] = len(retiree_missing_full)
        summary['retiree_not_candidate'] = len(retiree_not_candidate_full)
        summary['retiree_correct'] = len(retiree_correct_full)
        st.success("✅ マッチングが完了しました。")

    with st.spinner("STEP 4: 追加エラーチェックを実行中..."):
        salary_decrease = pd.DataFrame()
        salary_increase_rate = pd.DataFrame()
        cumulative_salary_check = pd.DataFrame()
        cumulative_salary_check2 = pd.DataFrame()

        sal1_zenki, sal1_touki = f"{col_salary1}_zenki", f"{col_salary1}_touki"
        sal2_zenki, sal2_touki = f"{col_salary2}_zenki", f"{col_salary2}_touki"

        if error_check_config['salary_decrease_check'] and sal1_zenki in both_full.columns and sal1_touki in both_full.columns:
            temp_df = both_full.dropna(subset=[sal1_touki, sal1_zenki])
            salary_decrease = temp_df[temp_df[sal1_touki] < temp_df[sal1_zenki]]
        summary['salary_decrease'] = len(salary_decrease)

        if error_check_config['salary_increase_rate_check'] and sal1_zenki in both_full.columns and sal1_touki in both_full.columns:
            x = error_check_config['x_rate']
            temp_df = both_full.dropna(subset=[sal1_touki, sal1_zenki])
            salary_increase_rate = temp_df[temp_df[sal1_touki] >= temp_df[sal1_zenki] * (1 + x / 100)]
        summary['salary_increase_rate'] = len(salary_increase_rate)

        if error_check_config['cumulative_salary_check'] and sal1_zenki in both_full.columns and sal2_zenki in both_full.columns and sal2_touki in both_full.columns:
            y = error_check_config['y_months']
            temp_df = both_full.dropna(subset=[sal2_touki, sal2_zenki, sal1_zenki])
            cumulative_salary_check = temp_df[temp_df[sal2_touki] < temp_df[sal2_zenki] + temp_df[sal1_zenki] * y]
        summary['cumulative_salary_check'] = len(cumulative_salary_check)

        if error_check_config['cumulative_salary_check2'] and sal1_zenki in both_full.columns and sal2_zenki in both_full.columns and sal2_touki in both_full.columns:
            y = error_check_config['y_months']
            z = error_check_config['z_rate']
            temp_df = both_full.dropna(subset=[sal2_touki, sal2_zenki, sal1_zenki])
            cumulative_salary_check2 = temp_df[temp_df[sal2_touki] > (temp_df[sal2_zenki] + temp_df[sal1_zenki] * y) * (1 + z / 100)]
        summary['cumulative_salary_check2'] = len(cumulative_salary_check2)
        st.success("✅ 追加エラーチェックが完了しました。")

    with st.spinner("STEP 5: 結果をExcelファイルに出力中..."):
        date_columns_to_format = [
            col_nyusha, col_seinengappi, "退職年月日", "支給日",
            f"{col_nyusha}_zenki", f"{col_nyusha}_touki",
            f"{col_seinengappi}_zenki", f"{col_seinengappi}_touki"
        ]
        
        result_dfs = {
            "前期末_キー重複": zenki_duplicates,
            "当期末_キー重複": touki_duplicates,
            "前期末_日付エラー": zenki_age_errors,
            "当期末_日付エラー": touki_age_errors,
            "在籍照合_前期のみ(退職者候補)": only_zenki_full,
            "在籍照合_当期のみ(入社者)": only_touki_full,
            "在籍照合_両方(在籍者)": both_full,
            "退職者照合_データ不在": retiree_missing_full,
            "退職者照合_候補でない": retiree_not_candidate_full,
            "退職者照合_一致": retiree_correct_full,
            "給与減額チェック": salary_decrease,
            "給与増加率チェック": salary_increase_rate,
            "累計給与チェック": cumulative_salary_check,
            "累計給与チェック2": cumulative_salary_check2
        }

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df_to_write in result_dfs.items():
                formatted_df = format_date_columns(df_to_write, date_columns_to_format)
                formatted_df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        processed_data = output.getvalue()
        st.success("✅ Excelファイルの出力準備が完了しました！")
    
    return processed_data, summary

# ------------------------------------------------------------------------------------
# Streamlit UI部分
# ------------------------------------------------------------------------------------

st.set_page_config(page_title="退職給付債務データチェック", layout="wide")

st.title('退職給付債務データチェックアプリ 📊')

st.write("""
このアプリは、前期末・当期末・退職者のExcelデータをアップロードすることで、
データの不整合やエラーを自動でチェックし、結果をExcelファイルとして出力します。
""")

with st.sidebar:
    st.header("⚙️ データ指定設定")
    
    st.subheader("1. ファイル設定")
    sheet_zenki = st.text_input("① 前期末データのシート名", "従業員データフォーマット")
    sheet_touki = st.text_input("② 当期末データのシート名", "従業員データフォーマット")
    sheet_taishoku = st.text_input("③ 退職者データのシート名", "退職者データフォーマット")
    
    st.subheader("2. 列名設定")
    col_employee_id = st.text_input("従業員番号の列名", "従業員番号")
    col_nyusha = st.text_input("入社年月日の列名", "入社年月日")
    col_seinengappi = st.text_input("生年月日の列名", "生年月日")
    col_salary1 = st.text_input("給与1の列名", "給与１")
    col_salary2 = st.text_input("給与2の列名", "給与２")
    
    st.header("🔍 エラーチェック設定")
    
    st.subheader("3. 追加エラーチェック項目")
    salary_decrease_check = st.checkbox("給与減額チェック", value=True)
    salary_increase_rate_check = st.checkbox("給与増加率チェック", value=True)
    x_rate = st.number_input("増加率(x) %", min_value=0.0, value=5.0, step=1.0, format="%.1f")
    cumulative_salary_check = st.checkbox("累計給与チェック", value=True)
    y_months = st.selectbox("月数(y)", [1, 12], index=0)
    cumulative_salary_check2 = st.checkbox("累計給与チェック2", value=True)
    z_rate = st.number_input("超過率(z) %", min_value=0.0, value=0.0, step=1.0, format="%.1f")
    
    error_check_config = {
        "salary_decrease_check": salary_decrease_check,
        "salary_increase_rate_check": salary_increase_rate_check,
        "x_rate": x_rate,
        "cumulative_salary_check": cumulative_salary_check,
        "y_months": y_months,
        "cumulative_salary_check2": cumulative_salary_check2,
        "z_rate": z_rate,
    }

st.header("📂 Excelファイルをアップロード")
uploaded_zenki = st.file_uploader("① 前期末従業員データ", type=['xlsx', 'xls'])
uploaded_touki = st.file_uploader("② 当期末従業員データ", type=['xlsx', 'xls'])
uploaded_taishoku = st.file_uploader("③ 当期退職者データ", type=['xlsx', 'xls'])

if st.button('チェック開始', type="primary", use_container_width=True):
    if uploaded_zenki and uploaded_touki and uploaded_taishoku:
        
        df_zenki = find_header_and_read_excel(uploaded_zenki, sheet_zenki, ['従業員番号', '生年月日', '給与'])
        df_touki = find_header_and_read_excel(uploaded_touki, sheet_touki, ['従業員番号', '生年月日', '給与'])
        df_taishoku = find_header_and_read_excel(uploaded_taishoku, sheet_taishoku, ['従業員番号', '生年月日', '退職'])
        
        if df_zenki is not None and df_touki is not None and df_taishoku is not None:
            result_excel, summary = data_check_and_matching(
                df_zenki, df_touki, df_taishoku,
                col_employee_id, col_nyusha, col_seinengappi,
                col_salary1, col_salary2,
                error_check_config
            )
            
            st.header("🎉 処理が完了しました！")

            st.subheader("チェック結果の概要")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("キー重複(前期)", f"{summary.get('zenki_duplicates', 0)} 件")
                st.metric("キー重複(当期)", f"{summary.get('touki_duplicates', 0)} 件")
                st.metric("在籍者数", f"{summary.get('both', 0)} 人")
            with col2:
                st.metric("日付エラー(前期)", f"{summary.get('zenki_age_errors', 0)} 件")
                st.metric("日付エラー(当期)", f"{summary.get('touki_age_errors', 0)} 件")
                st.metric("入社者数", f"{summary.get('only_touki', 0)} 人")
            with col3:
                st.metric("退職者照合(データ不在)", f"{summary.get('retiree_missing', 0)} 件")
                st.metric("退職者照合(候補でない)", f"{summary.get('retiree_not_candidate', 0)} 件")
                st.metric("退職者数", f"{summary.get('only_zenki', 0)} 人")
            with col4:
                st.metric("給与減額", f"{summary.get('salary_decrease', 0)} 件", delta_color="inverse")
                st.metric("給与増加率(x%)", f"{summary.get('salary_increase_rate', 0)} 件", delta_color="inverse")
                st.metric("累計給与", f"{summary.get('cumulative_salary_check', 0)} 件", delta_color="inverse")
                st.metric("累計給与2", f"{summary.get('cumulative_salary_check2', 0)} 件", delta_color="inverse")

            st.download_button(
                label="📥 結果をExcelファイルでダウンロード",
                data=result_excel,
                file_name="check_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.warning("⚠️ 3つのExcelファイルをすべてアップロードしてください。")