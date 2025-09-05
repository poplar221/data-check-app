import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io

# ------------------------------------------------------------------------------------
# STEP 1: データ処理関数 (この部分は変更ありません)
# ------------------------------------------------------------------------------------

def find_header_and_read_excel(file_path, sheet_name, keywords):
    """
    Excelファイルを読み込み、指定されたキーワードが含まれる行をヘッダーとして特定し、
    その行からデータを読み込んでDataFrameとして返します。
    """
    try:
        df_temp = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        header_row_index = -1
        for i, row in df_temp.iterrows():
            row_str = ''.join(map(str, row.values))
            if all(keyword in row_str for keyword in keywords):
                header_row_index = i
                break
        
        if header_row_index != -1:
            st.info(f"📄 '{file_path.name}' の '{sheet_name}' シートからヘッダーを {header_row_index + 1} 行目で発見しました。")
            df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row_index)
            return df
        else:
            st.error(f"⚠️ '{file_path.name}' の '{sheet_name}' シートでヘッダー行が見つかりませんでした。")
            return None
    except Exception as e:
        st.error(f"❌エラー: '{file_path.name}' の '{sheet_name}' シート読み込み中に問題が発生しました: {e}")
        return None

def data_check_and_matching(df_zenki, df_touki, df_taishoku, col_nyusha, col_seinengappi, col_employee_id):
    """
    データチェックとマッチング処理を行い、結果をExcelファイルとしてメモリ上に保存します。
    """
    st.write("---")
    st.subheader("処理状況")
    
    with st.spinner("STEP 1: 前処理とキー作成を実行中..."):
        dfz = df_zenki.copy()
        dft = df_touki.copy()
        dftai = df_taishoku.copy()
        
        for df in [dfz, dft, dftai]:
            if col_nyusha in df.columns and col_seinengappi in df.columns:
                df[col_nyusha] = pd.to_datetime(df[col_nyusha].astype(str), errors='coerce')
                df[col_seinengappi] = pd.to_datetime(df[col_seinengappi].astype(str), errors='coerce')

        if col_employee_id in dfz.columns and col_employee_id in dft.columns:
            st.info(f"🔑 「{col_employee_id}」をキーとして使用します。")
            for df in [dfz, dft, dftai]:
                df['key'] = df[col_employee_id].astype(str)
        else:
            st.info(f"🔑 「入社年月日」と「生年月日」の連結文字列をキーとして使用します。")
            for df in [dfz, dft, dftai]:
                df['key'] = df[col_nyusha].dt.strftime('%Y%m%d') + '_' + df[col_seinengappi].dt.strftime('%Y%m%d')
        st.success("✅ キーの作成が完了しました。")

    with st.spinner("STEP 2: エラーチェックを実行中..."):
        zenki_duplicates = dfz[dfz.duplicated(subset=['key'], keep=False)]
        touki_duplicates = dft[dft.duplicated(subset=['key'], keep=False)]
        
        zenki_age_errors = pd.DataFrame()
        touki_age_errors = pd.DataFrame()
        if col_nyusha in dfz.columns and col_seinengappi in dfz.columns:
            days_diff_z = (dfz[col_nyusha] - dfz[col_seinengappi]).dt.days
            dfz['age_at_hire'] = (days_diff_z / 365.25).astype(int)
            zenki_age_errors = dfz[(dfz['age_at_hire'] < 15) | (dfz['age_at_hire'] >= 90)]

        if col_nyusha in dft.columns and col_seinengappi in dft.columns:
            days_diff_t = (dft[col_nyusha] - dft[col_seinengappi]).dt.days
            dft['age_at_hire'] = (days_diff_t / 365.25).astype(int)
            touki_age_errors = dft[(dft['age_at_hire'] < 15) | (dft['age_at_hire'] >= 90)]
        st.success("✅ エラーチェックが完了しました。")
    
    with st.spinner("STEP 3 & 4: マッチングと退職者照合を実行中..."):
        merged_df = pd.merge(dfz[['key']], dft[['key']], on='key', how='outer', indicator=True)
        only_zenki_keys = merged_df[merged_df['_merge'] == 'left_only']
        only_touki_keys = merged_df[merged_df['_merge'] == 'right_only']
        only_zenki_full = pd.merge(only_zenki_keys, dfz, on='key', how='left')
        only_touki_full = pd.merge(only_touki_keys, dft, on='key', how='left')

        retiree_merged = pd.merge(only_zenki_keys[['key']], dftai[['key']], on='key', how='outer', indicator='retiree_check')
        retiree_missing_keys = retiree_merged[retiree_merged['retiree_check'] == 'left_only']
        retiree_not_candidate_keys = retiree_merged[retiree_merged['retiree_check'] == 'right_only']
        retiree_correct_keys = retiree_merged[retiree_merged['retiree_check'] == 'both']

        retiree_missing_full = pd.merge(retiree_missing_keys, dfz, on='key', how='left')
        retiree_not_candidate_full = pd.merge(retiree_not_candidate_keys, dftai, on='key', how='left')
        retiree_correct_full = pd.merge(retiree_correct_keys, dftai, on='key', how='left')
        st.success("✅ マッチングと退職者照合が完了しました。")

    with st.spinner("STEP 5: 結果をExcelファイルに出力中..."):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            zenki_duplicates.to_excel(writer, sheet_name='前期末_キー重複', index=False)
            touki_duplicates.to_excel(writer, sheet_name='当期末_キー重複', index=False)
            zenki_age_errors.to_excel(writer, sheet_name='前期末_日付エラー', index=False)
            touki_age_errors.to_excel(writer, sheet_name='当期末_日付エラー', index=False)
            only_zenki_full.to_excel(writer, sheet_name='在籍照合_前期のみ', index=False)
            only_touki_full.to_excel(writer, sheet_name='在籍照合_当期のみ', index=False)
            retiree_missing_full.to_excel(writer, sheet_name='退職者照合_データ不在', index=False)
            retiree_not_candidate_full.to_excel(writer, sheet_name='退職者照合_候補でない', index=False)
            retiree_correct_full.to_excel(writer, sheet_name='退職者照合_一致', index=False)
        
        processed_data = output.getvalue()
        st.success("✅ Excelファイルの出力準備が完了しました！")
    
    return processed_data

# ------------------------------------------------------------------------------------
# STEP 2: StreamlitのUI部分
# ------------------------------------------------------------------------------------

st.set_page_config(page_title="退職給付債務データチェック", layout="wide")

st.title('退職給付債務データチェックアプリ')

st.write("""
このアプリは、前期末・当期末・退職者のExcelデータをアップロードすることで、\n
データの不整合やエラーを自動でチェックし、結果をExcelファイルとして出力します。
""")

# --- サイドバーに設定項目を作成 ---
with st.sidebar:
    st.header("1. ファイル設定")
    # ★★★ 修正箇所①: シート名入力を3つに分割 ★★★
    sheet_zenki = st.text_input("① 前期末データのシート名", "data")
    sheet_touki = st.text_input("② 当期末データのシート名", "data")
    sheet_taishoku = st.text_input("③ 退職者データのシート名", "data")
    
    st.header("2. 列名設定")
    col_employee_id = st.text_input("従業員番号の列名", "従業員番号")
    col_nyusha = st.text_input("入社年月日の列名", "入社年月日")
    col_seinengappi = st.text_input("生年月日の列名", "生年月日")

# --- メイン画面にファイルアップローダーを設置 ---
st.header("3. Excelファイルをアップロード")
uploaded_zenki = st.file_uploader("① 前期末従業員データ", type=['xlsx'])
uploaded_touki = st.file_uploader("② 当期末従業員データ", type=['xlsx'])
uploaded_taishoku = st.file_uploader("③ 当期退職者データ", type=['xlsx'])


# --- チェック開始ボタンと処理 ---
if st.button('チェック開始', type="primary"):
    if uploaded_zenki and uploaded_touki and uploaded_taishoku:
        
        # ★★★ 修正箇所②: それぞれのシート名を使ってファイルを読み込む ★★★
        df_zenki = find_header_and_read_excel(uploaded_zenki, sheet_zenki, ['入社', '生年', '給与'])
        df_touki = find_header_and_read_excel(uploaded_touki, sheet_touki, ['入社', '生年', '給与'])
        df_taishoku = find_header_and_read_excel(uploaded_taishoku, sheet_taishoku, ['入社', '生年'])
        
        if df_zenki is not None and df_touki is not None and df_taishoku is not None:
            result_excel = data_check_and_matching(
                df_zenki, df_touki, df_taishoku,
                col_nyusha, col_seinengappi, col_employee_id
            )
            
            st.header("🎉 処理が完了しました！")
            
            st.download_button(
                label="📥 結果をExcelファイルでダウンロード",
                data=result_excel,
                file_name="check_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("⚠️ 3つのExcelファイルをすべてアップロードしてください。")