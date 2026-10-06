import streamlit as st
import pandas as pd
import random
import io
import re

# --- セッションステートの初期化 ---
if "lottery_map" not in st.session_state:
    st.session_state.lottery_map = None
if 'export_df' not in st.session_state: 
    st.session_state.export_df = None

# ページの初期設定
st.set_page_config(page_title='ホール座席抽選システム', layout='wide')

# 特殊な空白文字エラーを絶対に起こさないためのCSS1行結合
css_data = '<style>.stage-line { border-top: 8px solid #555555; text-align: center; margin: 10px auto 30px auto; padding-top: 5px; font-weight: bold; color: #333333; letter-spacing: 10px; width: 60%; background-color: #f7e7d0; padding: 10px; border-radius: 4px; }.grid-container { display: flex; flex-direction: column; align-items: center; margin-bottom: 30px; }.row-container { display: flex; justify-content: center; align-items: center; margin-bottom: 6px; }.row-label { width: 40px; font-weight: bold; text-align: center; color: #333; font-size: 14px; }.block-left { display: flex; width: 165px; justify-content: flex-end; }.block-center { display: flex; width: 440px; justify-content: center; }.block-right { display: flex; width: 165px; justify-content: flex-start; }.seat-box { width: 49px; height: 38px; margin: 2px; border-radius: 4px; display: inline-flex; flex-direction: column; align-items: center; justify-content: center; font-size: 11px; font-weight: bold; color: #333333; background-color: #ffffff; border: 1px solid #333333; }.seat-assigned { background-color: #4a90e2; color: white; border: 1px solid #357abd; box-shadow: 0px 2px 4px rgba(0,0,0,0.1); }.seat-id { font-size: 9px; opacity: 0.8; }.member-name { font-size: 10px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 44px; }.seat-disabled { background-color: #e8e8e8; color: #d9534f; border: 1px solid #bbbbbb; border-style: dashed; font-size: 16px; }.aisle { width: 40px; }.staff-container { display: flex; justify-content: center; width: 440px; border: 2px solid #b22222; padding: 4px; border-radius: 4px; background-color: #fff0f0; margin-top: 10px; }.staff-box { width: 49px; height: 38px; margin: 2px; border: 1px solid #b22222; background-color: #ffffff; color: #b22222; display: inline-flex; align-items: center; justify-content: center; font-size: 10px; font-weight: bold; border-radius: 2px; }.booth-container { display: flex; justify-content: center; width: 440px; border: 2px solid #333333; padding: 4px; background-color: #f0f0f0; margin-top: 10px; }.booth-box { width: 210px; height: 50px; border: 1px solid #333333; background-color: #ffffff; display: inline-flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; margin: 2px; }</style>'
st.markdown(css_data, unsafe_allow_html=True)

# --- 座席マスタの定義 ---
layout = {
    'A': (['X','X','X'], [1,2,3,4,5,6,7,8], ['X','X','X']),
    'B': (['X','X','X'], [1,2,3,4,5,6,7,8], [9,10,11]),
    'C': ([1,2,3], [4,5,6,7,8,9,10,11], [12,13,14]),
    'D': ([1,2,3], [4,5,6,7,8,9,10,11], [12,13,14]),
    'E': (['X','X','X'], [1,2,3,4,5,6,7,8], ['X','X','X']),
    'F': (['X','X','X'], [1,2,3,4,5,6,7,8], ['X','X','X']),
    'G': ([1,2,3], [], [None,None,None]),
    'H': ([1,2,3], [4,5,6,7,8,9,10,11], [12,13,14]),
    'I': ([1,2,3], [4,5,6,7,8,9,10,11], [12,13,14]),
    'J': (['X','X','X'], [4,5,6,7,8,9,10,11], [12,13,14]),
    'K': ([None,None,None], [4,5,6,7,8,9,10,11], [None,None,None])
}
ordered_valid_seats = []
for row in ['A','B','C','D','E','F','G','H','I','J','K']:
    left, center, right = layout[row]
    for s in left + center + right:
        if type(s) is int: ordered_valid_seats.append(f'{row}-{s}')

# =====================================================================
# 座席抽選メイン処理
# =====================================================================
st.title('💺 ホール座席抽選システム（連席優先・図面完全再現版）')

uploaded_file = st.sidebar.file_uploader('『チケット抽選_抽選結果.xlsx』をアップロード', type=['xlsx', 'xls'])
st.sidebar.info(f'現在の総利用可能座席数: {len(ordered_valid_seats)} 席')

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # カラム名の自動検出（部分一致で安全に探す）
        result_col = next((c for c in df.columns if '当選' in str(c)), None)
        num_col = next((c for c in df.columns if '人数' in str(c)), None)
        mail_col = next((c for c in df.columns if 'メール' in str(c) or '@' in str(c)), None)
        
        rep_name_col = next((c for c in df.columns if '代表者氏名' in str(c) or '名前' in str(c)), None)
        rep_id_col = next((c for c in df.columns if '代表者学籍番号' in str(c)), None)
        
        c1_name_col = next((c for c in df.columns if '同行者1 氏名' in str(c)), None)
        c1_id_col = next((c for c in df.columns if '同行者1学籍番号' in str(c)), None)
        
        c2_name_col = next((c for c in df.columns if '同行者2 氏名' in str(c)), None)
        c2_id_col = next((c for c in df.columns if '同行者2学籍番号' in str(c)), None)
        
        c3_name_col = next((c for c in df.columns if '同行者3 氏名' in str(c)), None)
        c3_id_col = next((c for c in df.columns if '同行者3学籍番号' in str(c)), None)
        
        if not result_col or not rep_name_col:
            st.error('エラー: エクセルファイルに「当選」または「代表者氏名」の列が見つかりません。')
        else:
            winners_df = df[df[result_col] == '当選'].copy()
            groups = []
            total_required_seats = 0
            
            for _, row in winners_df.iterrows():
                # 人数の数値化
                raw_count = str(row[num_col]) if (num_col and pd.notna(row[num_col])) else "1"
                numbers_only = re.sub(r'\D', '', raw_count)
                count = int(numbers_only) if numbers_only != "" else 1
                
                email = str(row[mail_col]).strip() if (mail_col and pd.notna(row[mail_col])) else '-'
                rep_name = str(row[rep_name_col]).strip()
                token = f"TOKEN-{abs(hash(email + rep_name)) % 1000000:06d}"
                
                # グループメンバー（名前と学籍番号のペア）を人数分だけ抽出
                members = []
                
                # 1人目：代表者
                if pd.notna(row[rep_name_col]):
                    m_name = str(row[rep_name_col]).strip()
                    m_id = str(row[rep_id_col]).strip() if (rep_id_col and pd.notna(row[rep_id_col])) else '-'
                    members.append({'name': m_name, 'student_id': m_id})
                
                # 2人目：同行者1
                if count >= 2 and c1_name_col and pd.notna(row.get(c1_name_col)):
                    m_name = str(row[c1_name_col]).strip()
                    m_id = str(row[c1_id_col]).strip() if (c1_id_col and pd.notna(row.get(c1_id_col))) else '-'
                    members.append({'name': m_name, 'student_id': m_id})
                
                # 3人目：同行者2
                if count >= 3 and c2_name_col and pd.notna(row.get(c2_name_col)):
                    m_name = str(row[c2_name_col]).strip()
                    m_id = str(row[c2_id_col]).strip() if (c2_id_col and pd.notna(row.get(c2_id_col))) else '-'
                    members.append({'name': m_name, 'student_id': m_id})
                
                # 4人目：同行者3
                if count >= 4 and c3_name_col and pd.notna(row.get(c3_name_col)):
                    m_name = str(row[c3_name_col]).strip()
                    m_id = str(row[c3_id_col]).strip() if (c3_id_col and pd.notna(row.get(c3_id_col))) else '-'
                    members.append({'name': m_name, 'student_id': m_id})
                
                # 実際のメンバー数に補正
                actual_count = len(members)
                total_required_seats += actual_count
                
                groups.append({
                    'rep_name': rep_name,
                    'count': actual_count,
                    'email': email,
                    'token': token,
                    'members': members
                })
            
            st.sidebar.success(f'名簿読み込み成功！ 必要総座席数: {total_required_seats}席')
            
            if total_required_seats > len(ordered_valid_seats):
                st.error(f'エラー: 必要座席数 ({total_required_seats}席) が全席数 ({len(ordered_valid_seats)}席) を超えています！')
            elif st.sidebar.button('🎉 当選者の座席抽選を実行'):
                seat_mapping = {}
                available_seats = ordered_valid_seats.copy()
                export_data = []
                
                random.shuffle(groups)
                
                # 連席を考慮したグループごとの座席割り当て
                for g in groups:
                    c = g['count']
                    assigned = False
                    
                    if c > 1:
                        start_indices = list(range(len(available_seats) - c + 1))
                        random.shuffle(start_indices)
                        for start_idx in start_indices:
                            candidate_seats = available_seats[start_idx : start_idx + c]
                            rows = [s.split('-')[0] for s in candidate_seats]
                            if len(set(rows)) == 1: # 同一列内での連席
                                for i, seat in enumerate(candidate_seats):
                                    member = g['members'][i]
                                    seat_mapping[seat] = member['name']
                                    available_seats.remove(seat)
                                    export_data.append({
                                        '座席番号': seat,
                                        '氏名': member['name'],
                                        '学籍番号': member['student_id'],
                                        '代表者名': g['rep_name'],
                                        'グループ人数': c,
                                        '代表者メールアドレス': g['email'],
                                        '受付キー': g['token']
                                    })
                                assigned = True
                                break
                    
                    if not assigned:
                        # 連席が取れなかった場合、または1人応募の場合
                        for i in range(c):
                            seat = available_seats.pop(0)
                            member = g['members'][i]
                            seat_mapping[seat] = member['name']
                            export_data.append({
                                '座席番号': seat,
                                '氏名': member['name'],
                                '学籍番号': member['student_id'],
                                '代表者名': g['rep_name'],
                                'グループ人数': c,
                                '代表者メールアドレス': g['email'],
                                '受付キー': g['token']
                            })
                
                st.session_state.lottery_map = seat_mapping
                export_df = pd.DataFrame(export_data).sort_values(by='座席番号').reset_index(drop=True)
                st.session_state.export_df = export_df
                st.balloons()
    except Exception as e:
        st.error(f'ファイル読み込みエラー: {e}')

# --- 図面のビジュアルレンダリング ---
if st.session_state.lottery_map is not None:
    results = st.session_state.lottery_map
    st.subheader('🎬 会場座席レイアウト（当選者配置結果）')
    st.markdown("<div class='stage-line'>舞台 STAGE (10.05m)</div>", unsafe_allow_html=True)
    
    grid_html = "<div class='grid-container'>"
    
    for row_label, (left_range, center_range, right_range) in layout.items():
        grid_html += "<div class='row-container'>"
        grid_html += f"<div class='row-label'>{row_label}</div>"
        
        # 左ブロック
        grid_html += "<div class='block-left'>"
        for s in left_range:
            if s == 'X':
                grid_html += "<div class='seat-box seat-disabled'>✕</div>"
            elif type(s) is int:
                seat_code = f'{row_label}-{s}'
                if seat_code in results:
                    grid_html += f"<div class='seat-box seat-assigned'><div class='seat-id'>{seat_code}</div><div class='member-name'>{results[seat_code]}</div></div>"
                else:
                    grid_html += f"<div class='seat-box'><div class='seat-id'>{seat_code}</div><div class='member-name'>-</div></div>"
            else:
                grid_html += "<div class='seat-box' style='visibility: hidden; border: none;'></div>"
        grid_html += "</div>"
        
        grid_html += "<div class='aisle'></div>"
        
        # 中央ブロック
        if row_label == 'G':
            grid_html += "<div class='block-center' style='color: #333; font-weight: bold; font-size: 15px; letter-spacing: 4px; align-items: center;'>座席中央通路（G行）</div>"
        else:
            grid_html += "<div class='block-center'>"
            target_center = [1,2,3,4,5,6,7,8] if row_label in ['A','B','E','F'] else [4,5,6,7,8,9,10,11]
            for s in target_center:
                if s in center_range:
                    seat_code = f'{row_label}-{s}'
                    if seat_code in results:
                        grid_html += f"<div class='seat-box seat-assigned'><div class='seat-id'>{seat_code}</div><div class='member-name'>{results[seat_code]}</div></div>"
                    else:
                        grid_html += f"<div class='seat-box'><div class='seat-id'>{seat_code}</div><div class='member-name'>-</div></div>"
            grid_html += "</div>"
        
        grid_html += "<div class='aisle'></div>"
        
        # 右ブロック
        grid_html += "<div class='block-right'>"
        for s in right_range:
            if s == 'X':
                grid_html += "<div class='seat-box seat-disabled'>✕</div>"
            elif type(s) is int:
                seat_code = f'{row_label}-{s}'
                if seat_code in results:
                    grid_html += f"<div class='seat-box seat-assigned'><div class='seat-id'>{seat_code}</div><div class='member-name'>{results[seat_code]}</div></div>"
                else:
                    grid_html += f"<div class='seat-box'><div class='seat-id'>{seat_code}</div><div class='member-name'>-</div></div>"
            else:
                grid_html += "<div class='seat-box' style='visibility: hidden; border: none;'></div>"
        grid_html += "</div>"
        
        grid_html += f"<div class='row-label'>{row_label}</div>"
        grid_html += "</div>"
        
    # 関係者席
    grid_html += "<div class='row-container'>"
    grid_html += "<div class='row-label'></div><div class='block-left'></div><div class='aisle'></div>"
    grid_html += "<div class='staff-container'>"
    for _ in range(8): grid_html += "<div class='staff-box'>関係者</div>"
    grid_html += "</div>"
    grid_html += "<div class='aisle'></div><div class='block-right'></div>"
    grid_html += "<div class='row-label' style='font-size:10px; color:#b22222; width:80px; text-align:left; margin-left:10px;'>←関係者席</div>"
    grid_html += "</div>"
    
    # 音響ブース
    grid_html += "<div class='row-container'><div class='booth-container'><div class='booth-box'>照明</div><div class='booth-box'>音</div></div></div></div>"
    st.markdown(grid_html, unsafe_allow_html=True)
    
    # データ一覧・ダウンロードエリア
    with st.expander("📊 座席決定の一覧表を確認（Excelダウンロードはこちら）", expanded=True):
        st.dataframe(st.session_state.export_df, use_container_width=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.export_df.to_excel(writer, index=False, sheet_name='座席抽選結果')
        buffer.seek(0)
        st.download_button(label="📥 stage_output_seats.xlsx をダウンロード", data=buffer, file_name="stage_output_seats.xlsx")
