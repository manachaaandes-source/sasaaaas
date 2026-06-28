"""
請求書抽出ツール（MVP / 1号機）
--------------------------------
PDF請求書をアップロードすると、Claude APIで主要項目を抽出し、
CSV / Excel でダウンロードできるStreamlitアプリ。

【配布の考え方】
- Streamlit Cloud にデプロイして「URL + パスワード」を購入者に渡す運用を想定。
- APIキーはあなた（販売者）側が st.secrets に入れる → 購入者はキー取得不要。
- 請求書1枚あたりのAPIコストは数円なので、買い切り価格に吸収できる。

【最小構成の原則】
ここに機能を足したくなったら、まず Gumroad で1本売れてからにすること。
複数フォーマット対応・画像対応・項目カスタマイズは全部「2号機以降」。
"""

import json
import base64
import io

import requests
import streamlit as st
import pandas as pd
from anthropic import Anthropic

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
# ライセンスキー認証（Gumroad購入者向け）
# ============================================================
# 共通パスワードではなく、購入者ごとに発行される固有のライセンスキーを
# Gumroadに照会して本人確認する。これで「キーの使い回し」を捕捉できる土台になる。
#
# 【事前準備（販売者）】
# 1. Gumroadの商品設定で "Generate a unique license key per sale" をONにする
#    （商品編集画面 → Settings → Content → License keys）
# 2. 商品IDを secrets の GUMROAD_PRODUCT_ID に入れる
#    （商品の "Product ID" は商品編集URLや共有設定から確認できる）
# 購入者には、Gumroadが購入時に自動でメールするライセンスキーを使ってもらう。
def verify_license(key: str):
    """Gumroadにキーを照会。(OK?, エラーメッセージ) を返す。"""
    product_id = st.secrets.get("GUMROAD_PRODUCT_ID", "")
    if not product_id:
        return False, "商品IDが未設定です（管理者向け: secretsにGUMROAD_PRODUCT_IDを設定）。"
    if not key:
        return False, "ライセンスキーを入力してください。"

    try:
        r = requests.post(
            "https://api.gumroad.com/v2/licenses/verify",
            data={
                "product_id": product_id,
                "license_key": key,
                # 検証のたびにカウントを増やさない（純粋に有効性だけ確認）
                "increment_uses_count": "false",
            },
            timeout=15,
        )
        data = r.json()
    except Exception:
        return False, "認証サーバーに接続できませんでした。少し待って再試行してください。"

    if not data.get("success"):
        return False, "ライセンスキーが正しくありません。購入時のメールをご確認ください。"

    purchase = data.get("purchase", {}) or {}
    # 返金・チャージバックされた購入は無効化
    if purchase.get("refunded") or purchase.get("chargebacked") or purchase.get("disputed"):
        return False, "この購入は無効化されています（返金等）。"
    # サブスク商品の場合、解約・失効していたら無効（買い切りなら全てNoneで素通り）
    if (purchase.get("subscription_cancelled_at")
            or purchase.get("subscription_ended_at")
            or purchase.get("subscription_failed_at")):
        return False, "サブスクリプションが有効ではありません。"

    # --- 共有検知をしたくなったら、ここを有効化する（今はオフ）---
    # uses = data.get("uses", 0)            # このキーが何回認証されたか
    # if uses > 300:                        # 閾値は運用しながら調整
    #     return False, "このキーは利用上限を超えています。お問い合わせください。"

    return True, ""


def check_license() -> bool:
    """ライセンスキーが有効なら True。一度通ればセッション中は再認証不要。"""
    if st.session_state.get("authenticated"):
        return True

    st.title("🔒 請求書抽出ツール")
    st.caption("購入時にメールで届いたライセンスキーを入力してください。")
    key = st.text_input("ライセンスキー", placeholder="XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX")
    if st.button("認証する"):
        ok, msg = verify_license(key.strip())
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
    if check_license():
        main()
