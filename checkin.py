import streamlit as st
import pandas as pd
import cv2
from PIL import Image
import re
import numpy as np
import time

st.set_page_config(page_title='当日受付チェックインシステム', layout='wide')

# --- セッション状態の初期化 ---
if 'checked_in_tokens' not in st.session_state:
    st.session_state.checked_in_tokens = set()
if 'last_detected_key' not in st.session_state:
    st.session_state.last_detected_key = ""
if 'captured_frame' not in st.session_state:
    st.session_state.captured_frame = None

st.title('🎟️ 当日受付チェックイン専用システム（撮影してQR読み取り）')

st.sidebar.header("📁 ステップ1: 名簿読み込み")

master_file = st.sidebar.file_uploader(
    "『stage_output_seats.xlsx』をアップロードしてください",
    type=["xlsx"]
)

def extract_token(raw_key: str) -> str:
    raw_key = str(raw_key).strip()
    if not raw_key:
        return ""
    m = re.match(r'(TOKEN-\d+)', raw_key, re.IGNORECASE)
    if m:
        return m.group(1)
    return raw_key


if master_file is not None:

    df_master = pd.read_excel(master_file)
    df_master.columns = df_master.columns.str.strip()

    required_cols = ['受付キー', '代表者名', '座席番号']
    missing = [c for c in required_cols if c not in df_master.columns]
    if missing:
        st.error(f"Excelに必要な列がありません: {missing}")
        st.stop()

    email_col = None
    for cand in ['代表者メールアドレス', '応募者のメールアドレス', 'メールアドレス']:
        if cand in df_master.columns:
            email_col = cand
            break

    count_col = None
    for cand in ['グループ人数', '人数(同行者)', '人数']:
        if cand in df_master.columns:
            count_col = cand
            break

    df_master['受付キー'] = df_master['受付キー'].astype(str).str.strip()
    if email_col:
        df_master[email_col] = df_master[email_col].astype(str).str.strip()

    total_groups = df_master['受付キー'].nunique()
    checked_in_count = len(st.session_state.checked_in_tokens)

    st.sidebar.success("名簿データの読み込みに成功しました！")
    st.sidebar.metric("総当選グループ数", f"{total_groups} 組")
    st.sidebar.metric("受付済み", f"{checked_in_count} 組")

    col_cam, col_result = st.columns([1, 1])

    with col_cam:
        st.subheader("📷 カメラでQRを撮影")

        st.info("USBでiPhoneを繋いでいる場合は、カメラ番号を 1 や 2 に変えてみてください。")

        camera_index = st.number_input(
            "カメラ番号",
            min_value=0,
            max_value=10,
            value=0,
            step=1
        )

        if st.button("📸 撮影してQRを読む", type="primary", use_container_width=True):
            cap = cv2.VideoCapture(int(camera_index))
            cap.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
            cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

            if not cap.isOpened():
                st.error("カメラを開けませんでした。カメラ番号を変えて再試行してください。")
            else:
                # ===== 5秒間プレビュー =====
                preview_placeholder = st.empty()
                status_text = st.empty()

                for i in range(50):  # 約5秒（0.1秒 × 50回）
                    ret, frame = cap.read()
                    if ret:
                        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                        preview_placeholder.image(rgb, channels="RGB", use_container_width=True)
                        remaining = 5 - (i // 10)
                        status_text.info(f"準備中... あと {remaining} 秒")
                    time.sleep(0.1)

                status_text.success("撮影しました！")
                # ===== プレビュー終了 =====

                # 最後のフレームを正式に撮影
                ret, frame = cap.read()
                cap.release()

                if ret:
                    st.session_state.captured_frame = frame

                    # QR検出
                    qr_detector = cv2.QRCodeDetector()
                    data, points, _ = qr_detector.detectAndDecode(frame)

                    if data:
                        detected_raw = data.strip()
                        st.session_state.last_detected_key = detected_raw
                        st.success(f"✅ QRを検出しました: {detected_raw}")
                    else:
                        st.warning("QRコードを検出できませんでした。もう一度撮影してください。")
                else:
                    st.error("映像を取得できませんでした。")

        # 撮影した画像を表示
        if st.session_state.captured_frame is not None:
            rgb = cv2.cvtColor(st.session_state.captured_frame, cv2.COLOR_BGR2RGB)
            st.image(rgb, caption="撮影した画像", use_container_width=True)

    with col_result:
        st.subheader("🔍 照合結果・本人確認")

        manual_key = st.text_input(
            "👉 手動入力・修正用（受付キー / メールアドレス / QRの全文）",
            value=st.session_state.last_detected_key
        )

        current_raw = manual_key.strip() if manual_key.strip() else st.session_state.last_detected_key
        current_key = extract_token(current_raw)

        if current_key:
            conditions = (df_master['受付キー'] == current_key)
            if email_col:
                conditions = conditions | (df_master[email_col] == current_key)

            match_rows = df_master[conditions]

            if not match_rows.empty:
                first = match_rows.iloc[0]
                actual_token = str(first['受付キー']).strip()
                rep_name = str(first['代表者名']).strip()
                email = str(first[email_col]).strip() if email_col else "-"
                count = first[count_col] if count_col else len(match_rows)

                seats_list = match_rows['座席番号'].astype(str).tolist()
                seats_text = ", ".join(seats_list)

                display_cols = ['座席番号', '氏名', '学籍番号']
                available_cols = [c for c in display_cols if c in match_rows.columns]
                members_df = match_rows[available_cols].copy()

                if actual_token in st.session_state.checked_in_tokens:
                    st.error(f"⚠️ 二重受付\n{rep_name} 様のグループはすでに受付済みです")
                    st.dataframe(members_df, use_container_width=True, hide_index=True)
                else:
                    st.success("⭕ データ照合成功 — 学生証で本人確認してください")

                    st.markdown(
                        f"""
<div style="background:#f0fdf4;padding:16px;border-radius:8px;border:2px solid #16a34a;margin-bottom:12px;">
  <h2 style="margin:0 0 8px 0;">👤 代表者: {rep_name} 様</h2>
  <p style="margin:4px 0;">✉️ {email}</p>
  <p style="margin:4px 0;">👥 グループ人数: <b>{count}</b> 名</p>
  <p style="margin:4px 0;">💺 座席: <b style="font-size:1.2em;">{seats_text}</b></p>
  <p style="margin:8px 0 0 0;color:#b45309;font-weight:bold;">
    ※ 必ず全員の学生証を確認してから「入場確定」を押してください
  </p>
</div>
""",
                        unsafe_allow_html=True
                    )

                    st.markdown("**グループメンバー詳細**")
                    st.dataframe(members_df, use_container_width=True, hide_index=True)

                    if st.button("✅ 入場確定", use_container_width=True, type="primary"):
                        st.session_state.checked_in_tokens.add(actual_token)
                        st.session_state.last_detected_key = ""
                        st.session_state.captured_frame = None
                        st.success(f"入場確定しました: {rep_name} 様グループ")
                        st.rerun()
            else:
                st.error(f"❌ 該当データなし\n入力値: {current_raw}")
                if current_key != current_raw:
                    st.caption(f"（抽出した受付キー: {current_key}）")

    # --- 受付済み一覧 ---
    st.markdown("---")
    st.subheader("📋 本日の受付済みグループ一覧")

    if st.session_state.checked_in_tokens:
        df_checked = df_master[df_master['受付キー'].isin(st.session_state.checked_in_tokens)]

        agg_dict = {
            '代表者名': 'first',
            '座席番号': lambda x: ", ".join(x.astype(str))
        }
        if email_col:
            agg_dict[email_col] = 'first'
        if count_col:
            agg_dict[count_col] = 'first'

        df_display = (
            df_checked
            .groupby('受付キー', as_index=False)
            .agg(agg_dict)
        )

        rename_map = {
            '受付キー': '受付キー',
            '代表者名': '代表者名',
            email_col: 'メールアドレス' if email_col else None,
            count_col: '人数' if count_col else None,
            '座席番号': '座席番号'
        }
        rename_map = {k: v for k, v in rename_map.items() if k in df_display.columns and v}
        df_display = df_display.rename(columns=rename_map)

        st.dataframe(df_display, use_container_width=True, hide_index=True)

        csv_data = df_display.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 受付済み名簿をCSVとしてダウンロード",
            data=csv_data,
            file_name="checked_in_list.csv",
            mime="text/csv",
            use_container_width=True
        )
    else:
        st.info("現在、受付済みのグループはありません。")

else:
    st.info("👈 左のサイドバーから stage_output_seats.xlsx を読み込んでください")
