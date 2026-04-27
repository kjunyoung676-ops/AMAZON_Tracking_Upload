"""
app.py — Tracking Upload Generator
Streamlit Cloud 배포용

OMS 배송완료 엑셀 + 아마존 템플릿 .xls → 배송확인 .xls 생성
xlutils.copy 방식으로 원본 서식/구조 완전 보존
"""

import io
import os
from datetime import datetime, date

import pandas as pd
import streamlit as st
import xlrd
import xlwt
from xlutils.copy import copy

# ── 페이지 설정 ───────────────────────────────────────────────
st.set_page_config(
    page_title="Tracking Upload Generator",
    page_icon="📦",
    layout="centered"
)

st.markdown("""
<style>
  @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&display=swap');
  .block-container { max-width: 680px; padding-top: 2rem; }
  .title  { font-family: 'IBM Plex Mono', monospace; color: #00ff88;
            font-size: 1.1rem; font-weight: 700; margin-bottom: 0; }
  .sub    { color: #888; font-size: 0.78rem; margin-bottom: 1.5rem; }
  .log-ok   { color: #00cc66; font-size: 0.8rem; }
  .log-warn { color: #ff6b35; font-size: 0.8rem; }
  .log-err  { color: #ff4444; font-size: 0.8rem; }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="title">// TRACKING_UPLOAD_GENERATOR</div>', unsafe_allow_html=True)
st.markdown('<div class="sub">CJ OMS 배송완료 엑셀 → Flat.File.ShippingConfirmation.jp 생성</div>', unsafe_allow_html=True)

# ── STEP 0: 아마존 템플릿 ─────────────────────────────────────
st.markdown("#### STEP 0 — 아마존 배송확인 템플릿")
st.caption("Flat.File.ShippingConfirmation.jp (1).xls — 원본 그대로 유지, 데이터만 삽입")

# 내장 템플릿 경로 (GitHub repo에 넣어둔 경우 자동 사용)
BUILTIN_TPL = os.path.join(os.path.dirname(__file__), "template", "Flat.File.ShippingConfirmation.jp.xls")

tpl_file = st.file_uploader(
    "템플릿 업로드 (repo에 내장된 경우 생략 가능)",
    type=["xls", "xlsx"],
    key="tpl"
)

tpl_bytes = None
tpl_source = ""

if tpl_file:
    tpl_bytes  = tpl_file.read()
    tpl_source = tpl_file.name
    st.success(f"✓ 업로드됨: {tpl_file.name}")
elif os.path.exists(BUILTIN_TPL):
    with open(BUILTIN_TPL, "rb") as f:
        tpl_bytes = f.read()
    tpl_source = "내장 템플릿"
    st.info(f"ℹ 내장 템플릿 사용: {os.path.basename(BUILTIN_TPL)}")
else:
    st.warning("⚠ 템플릿 파일을 업로드하거나 repo의 template/ 폴더에 넣어주세요.")

st.divider()

# ── STEP 1: OMS ───────────────────────────────────────────────
st.markdown("#### STEP 1 — OMS 배송완료 엑셀")
st.caption("dlvrComptList_*.xlsx")

oms_file = st.file_uploader("OMS 파일 업로드", type=["xlsx", "xls"], key="oms")

st.divider()

# ── STEP 2: 미배송 리포트 ─────────────────────────────────────
st.markdown("#### STEP 2 — 미배송 주문 리포트 _(선택)_")
st.caption("22******.txt — order-item-id 매핑용. 없으면 공백 처리")

rpt_file = st.file_uploader("리포트 업로드 (선택)", type=["txt"], key="rpt")

st.divider()

# ── ship-date ─────────────────────────────────────────────────
st.markdown("#### ship-date")
ship_date_val = st.date_input("출하일 (JST)", value=date.today())
ship_date_str = ship_date_val.strftime('%Y-%m-%d') + 'T00:00:00+09:00'
st.caption(f"→ `{ship_date_str}`")

st.divider()

# ── 생성 버튼 ─────────────────────────────────────────────────
ready = tpl_bytes is not None and oms_file is not None
btn   = st.button("⚡ 배송확인 파일 생성", type="primary", disabled=not ready, use_container_width=True)

if btn:
    logs = []

    with st.spinner("처리 중..."):
        try:
            # ── order-item-id 매핑 ────────────────────────────
            id_map = {}
            if rpt_file:
                try:
                    rpt_bytes = rpt_file.read()
                    rpt_df = pd.read_csv(io.BytesIO(rpt_bytes), sep='\t', encoding='cp932', dtype=str)
                    if 'order-id' in rpt_df.columns and 'order-item-id' in rpt_df.columns:
                        id_map = dict(zip(
                            rpt_df['order-id'].str.strip(),
                            rpt_df['order-item-id'].str.strip()
                        ))
                        logs.append(('ok', f"리포트: {len(id_map)}건 order-item-id 매핑"))
                    else:
                        logs.append(('warn', "리포트에 필요 컬럼 없음 → order-item-id 공백"))
                except Exception as ex:
                    logs.append(('warn', f"리포트 파싱 오류: {ex} → order-item-id 공백"))
            else:
                logs.append(('warn', "리포트 없음 → order-item-id 공백 처리"))

            # ── OMS 파싱 ──────────────────────────────────────
            oms_bytes = oms_file.read()
            oms_df    = pd.read_excel(io.BytesIO(oms_bytes), dtype=str)
            cols      = oms_df.columns.tolist()

            order_col = next(
                (c for c in cols if '쇼핑몰' in c and '주문번호' in c), None
            ) or next(
                (c for c in cols if '주문번호' in c and '상품' not in c), None
            )
            track_col = next(
                (c for c in cols if '주문송장번호' in c or '송장번호' in c), None
            )

            if not order_col:
                raise ValueError(f"주문번호 컬럼을 찾을 수 없습니다.\n컬럼 목록: {', '.join(cols)}")
            if not track_col:
                raise ValueError(f"송장번호 컬럼을 찾을 수 없습니다.\n컬럼 목록: {', '.join(cols)}")

            logs.append(('ok', f"OMS: {len(oms_df)}건 로드"))
            logs.append(('ok', f"[{order_col}] → order-id  /  [{track_col}] → tracking-number"))
            logs.append(('ok', f"ship-date: {ship_date_str}"))

            # ── 템플릿에 데이터 삽입 ──────────────────────────
            # xlutils.copy → 원본 .xls 서식/구조/다른 시트 완전 보존
            rb = xlrd.open_workbook(file_contents=tpl_bytes, formatting_info=True)
            wb = copy(rb)
            ws = wb.get_sheet(1)  # 시트 인덱스 1 = 出荷通知テンプレート_Template

            # 행 0: TemplateType 헤더  (건드리지 않음)
            # 행 1: 일본어 라벨        (건드리지 않음)
            # 행 2: 영어 컬럼명        (건드리지 않음)
            # 행 3~: 데이터 삽입
            ROW_START = 3
            row_idx   = ROW_START
            total     = 0
            multi     = 0

            for _, r in oms_df.iterrows():
                order_id      = str(r[order_col]).strip()
                raw_track     = str(r[track_col]).strip()
                order_item_id = id_map.get(order_id, '')
                track_list    = [t.strip() for t in raw_track.split(',') if t.strip()]

                if len(track_list) > 1:
                    multi += 1
                    logs.append(('warn', f"복수 송장: {order_id} → {len(track_list)}행 분리"))

                for trk in track_list:
                    ws.write(row_idx, 0, order_id)         # order-id
                    ws.write(row_idx, 1, order_item_id)    # order-item-id
                    ws.write(row_idx, 2, 1)                # quantity
                    ws.write(row_idx, 3, ship_date_str)    # ship-date
                    ws.write(row_idx, 4, 'SAGAWA')         # carrier-code
                    ws.write(row_idx, 5, '')               # carrier-name
                    ws.write(row_idx, 6, trk)              # tracking-number
                    ws.write(row_idx, 7, '宅配便')          # ship-method
                    row_idx += 1
                    total   += 1

            logs.append(('ok', f"총 {total}행 삽입 완료" + (f" (복수 송장 분리 {multi}건)" if multi else "")))

            # ── 저장 ──────────────────────────────────────────
            out_buf = io.BytesIO()
            wb.save(out_buf)
            out_buf.seek(0)

            file_date = ship_date_val.strftime('%Y%m%d')
            out_name  = f"tracking_upload_{file_date}.xls"

            logs.append(('ok', f"저장: {out_name}"))

        except Exception as ex:
            logs.append(('err', f"오류: {ex}"))
            out_buf  = None
            out_name = None

    # ── 로그 출력 ─────────────────────────────────────────────
    for ltype, msg in logs:
        icon = {"ok": "✓", "warn": "⚠", "err": "✗"}[ltype]
        color = {"ok": "#00cc66", "warn": "#ff6b35", "err": "#ff4444"}[ltype]
        st.markdown(f'<span style="color:{color};font-family:monospace;font-size:0.82rem">{icon} {msg}</span>', unsafe_allow_html=True)

    # ── 다운로드 버튼 ─────────────────────────────────────────
    if out_buf and out_name:
        st.success(f"✅ 생성 완료 — {out_name}")
        st.download_button(
            label=f"⬇ {out_name} 다운로드",
            data=out_buf.getvalue(),
            file_name=out_name,
            mime="application/vnd.ms-excel",
            use_container_width=True,
            type="primary"
        )
