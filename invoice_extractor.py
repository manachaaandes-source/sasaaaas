"""
請求書抽出ツール（MVP / 1号機）
--------------------------------
PDF請求書をアップロードすると、Claude APIで主要項目を抽出し、
CSV / Excel でダウンロードできるStreamlitアプリ。

【配布の考え方】
- Streamlit Cloud にデプロイして「URL + 個別パスワード」を購入者に渡す運用。
- パスワードは購入者ごとに固有（generate_licenses.pyで発行）。
- 1パスワード＝1端末。最初に使った端末に紐付き、別端末では使えない。
- APIキーは販売者側が st.secrets に入れる → 購入者はキー取得不要。

【最小構成の原則】
機能を足したくなったら、まず1本売れてからにすること。
"""

import io
import json
import base64
import datetime
import secrets as pysecrets

import streamlit as st
import pandas as pd
from anthropic import Anthropic
from supabase import create_client
from streamlit_local_storage import LocalStorage

# ============================================================
# 設定
# ============================================================
MODEL = "claude-sonnet-4-6"  # 精度重視。コストを下げたいなら "claude-haiku-4-5-20251001"
MAX_TOKENS = 2000

# 抽出させたい項目（ここを変えるだけで対応書類を増やせる = 横展開の起点）
EXTRACTION_PROMPT = """あなたは日本の請求書を読み取る経理アシスタントです。
添付されたPDFの請求書から以下の項目を抽出し、JSONのみを出力してください。
前置き・説明・```などのコードフェンスは一切付けないこと。

出力するJSONの形式:
{
  "invoice_number": "請求書番号（なければ空文字）",
  "issue_date": "発行日 YYYY-MM-DD形式（なければ空文字）",
  "due_date": "支払期日 YYYY-MM-DD形式（なければ空文字）",
  "vendor": "発行元（請求する側）の会社名",
  "customer": "請求先（支払う側）の会社名",
  "subtotal": 小計（数値のみ。不明ならnull）,
  "tax": 消費税額（数値のみ。不明ならnull）,
  "total": 合計金額（数値のみ。不明ならnull）,
  "line_items": [
    {"description": "明細の品目", "quantity": 数量(数値), "unit_price": 単価(数値), "amount": 金額(数値)}
  ]
}

金額は円記号やカンマを除いた数値だけにすること。
読み取れない項目は空文字またはnullにし、勝手に推測で埋めないこと。"""


# ============================================================
# パスワード認証 ＋ デバイス紐付け（1パスワード＝1端末）
# ============================================================
# 仕組み：
# 1. 購入者ごとに固有パスワードを発行（generate_licenses.pyで作る）
# 2. 初めて入力された端末に、そのパスワードを紐付ける（Supabaseに記録）
# 3. 別の端末で同じパスワードを使おうとすると拒否する
#
# 端末の識別：ブラウザのlocalStorageに見えないデバイスIDを保存して判別する。
# 完璧なデバイスロックではない（ブラウザデータを消されると回避されうる）が、
# 友達にパスワードを渡すようなカジュアルな共有はこれでほぼ防げる。
#
# 【事前準備（販売者）】st.secretsに以下を設定：
#   SUPABASE_URL          = "https://xxxx.supabase.co"
#   SUPABASE_SERVICE_KEY  = "sb_secret_..."   ← サーバー用キー。絶対に外に出さない
def get_device_id() -> str:
    """このブラウザ固有のデバイスIDを取得（無ければ発行して保存）。"""
    local_storage = LocalStorage()
    did = local_storage.getItem("invoice_tool_device_id")
    if not did:
        did = pysecrets.token_hex(16)
        local_storage.setItem("invoice_tool_device_id", did, key="invoice_tool_set_did")
    return did


def get_supabase():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_SERVICE_KEY", "")
    if not url or not key:
        return None
    return create_client(url, key)


def authenticate(password: str, device_id: str):
    """(OK?, エラーメッセージ) を返す。"""
    sb = get_supabase()
    if sb is None:
        return False, "サーバー設定エラーです（管理者: SUPABASE設定を確認してください）。"
    if not password:
        return False, "パスワードを入力してください。"
    if not device_id:
        return False, "端末を識別できませんでした。ページを再読み込みして再試行してください。"

    try:
        res = sb.table("licenses").select("*").eq("password", password).execute()
    except Exception:
        return False, "認証サーバーに接続できませんでした。少し待って再試行してください。"

    rows = res.data or []
    if not rows:
        return False, "パスワードが正しくありません。"

    bound = rows[0].get("device_id")
    if not bound:
        # 初回：この端末に紐付ける
        try:
            sb.table("licenses").update({
                "device_id": device_id,
                "bound_at": datetime.datetime.now(datetime.timezone.utc).isoformat(),
            }).eq("password", password).execute()
        except Exception:
            return False, "端末の登録に失敗しました。もう一度お試しください。"
        return True, ""

    if bound == device_id:
        return True, ""  # 同じ端末なのでOK

    return False, "このパスワードは別の端末で使用中です。1つのパスワードにつき1台のみご利用いただけます。"


def check_password() -> bool:
    """認証が通れば True。一度通ればセッション中は再認証不要。"""
    if st.session_state.get("authenticated"):
        return True

    device_id = get_device_id()

    st.title("🔒 請求書抽出ツール")
    st.caption("購入時にお渡ししたパスワードを入力してください。")
    pw = st.text_input("パスワード", type="password")
    if st.button("認証する"):
        ok, msg = authenticate(pw.strip(), device_id)
        if ok:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error(msg)
    return False


# ============================================================
# 抽出ロジック
# ============================================================
def extract_invoice(client: Anthropic, file_bytes: bytes) -> dict:
    """PDF1枚分のbytesを受け取り、抽出結果のdictを返す。"""
    b64 = base64.standard_b64encode(file_bytes).decode("utf-8")

    response = client.messages.create(
        model=MODEL,
        max_tokens=MAX_TOKENS,
        messages=[
            {
                "role": "user",
                "content": [
                    {
                        "type": "document",
                        "source": {
                            "type": "base64",
                            "media_type": "application/pdf",
                            "data": b64,
                        },
                    },
                    {"type": "text", "text": EXTRACTION_PROMPT},
                ],
            }
        ],
    )

    # テキストブロックだけ連結
    text = "".join(
        block.text for block in response.content if block.type == "text"
    ).strip()

    # 念のためコードフェンスを除去してからパース
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)


def to_summary_row(data: dict, filename: str) -> dict:
    """1請求書 = サマリー1行に整形。"""
    return {
        "ファイル名": filename,
        "請求書番号": data.get("invoice_number", ""),
        "発行日": data.get("issue_date", ""),
        "支払期日": data.get("due_date", ""),
        "発行元": data.get("vendor", ""),
        "請求先": data.get("customer", ""),
        "小計": data.get("subtotal"),
        "消費税": data.get("tax"),
        "合計": data.get("total"),
    }


def to_detail_rows(data: dict, filename: str) -> list:
    """1請求書の明細を複数行に展開（請求書番号で紐付け）。"""
    rows = []
    for item in data.get("line_items", []) or []:
        rows.append(
            {
                "ファイル名": filename,
                "請求書番号": data.get("invoice_number", ""),
                "品目": item.get("description", ""),
                "数量": item.get("quantity"),
                "単価": item.get("unit_price"),
                "金額": item.get("amount"),
            }
        )
    return rows


def build_excel(summary_df: pd.DataFrame, detail_df: pd.DataFrame) -> bytes:
    """サマリー＋明細の2シートExcelをbytesで返す。"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="サマリー", index=False)
        detail_df.to_excel(writer, sheet_name="明細", index=False)
    return buffer.getvalue()


# ============================================================
# メイン画面
# ============================================================
def main():
    st.set_page_config(page_title="請求書抽出ツール", page_icon="📄")
    st.title("📄 請求書抽出ツール")
    st.write("PDFの請求書をアップロードすると、項目を読み取ってCSV / Excelで出力します。")

    api_key = st.secrets.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        st.error("APIキーが設定されていません（管理者向け: secretsにANTHROPIC_API_KEYを設定）。")
        return

    client = Anthropic(api_key=api_key)

    files = st.file_uploader(
        "請求書PDF（複数選択可）",
        type=["pdf"],
        accept_multiple_files=True,
    )

    if not files:
        return

    if st.button(f"{len(files)}件を抽出する", type="primary"):
        summary_rows, detail_rows = [], []
        progress = st.progress(0.0)

        for i, f in enumerate(files):
            with st.spinner(f"{f.name} を読み取り中..."):
                try:
                    data = extract_invoice(client, f.read())
                    summary_rows.append(to_summary_row(data, f.name))
                    detail_rows.extend(to_detail_rows(data, f.name))
                except json.JSONDecodeError:
                    st.warning(f"⚠️ {f.name}: 読み取り結果をうまく解析できませんでした。")
                except Exception as e:
                    st.warning(f"⚠️ {f.name}: 処理に失敗しました（{e}）。")
            progress.progress((i + 1) / len(files))

        if not summary_rows:
            st.error("抽出できた請求書がありませんでした。")
            return

        summary_df = pd.DataFrame(summary_rows)
        detail_df = pd.DataFrame(detail_rows)

        st.success(f"{len(summary_rows)}件を抽出しました。")

        st.subheader("サマリー")
        st.dataframe(summary_df, use_container_width=True)

        if not detail_df.empty:
            st.subheader("明細")
            st.dataframe(detail_df, use_container_width=True)

        # ダウンロード
        st.download_button(
            "📥 サマリーCSVをダウンロード",
            summary_df.to_csv(index=False).encode("utf-8-sig"),
            file_name="invoices_summary.csv",
            mime="text/csv",
        )
        st.download_button(
            "📥 Excel（サマリー＋明細）をダウンロード",
            build_excel(summary_df, detail_df),
            file_name="invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    if check_password():
        main()
