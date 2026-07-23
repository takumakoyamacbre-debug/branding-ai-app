import asyncio
import json
import os
import re
import time
import urllib.parse
from google import genai
from google.genai import types
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field
import streamlit as st

# --------------------------------------------------
# Pydantic スキーマ定義
# --------------------------------------------------
class ColorPalette(BaseModel):
  main_color: str = Field(description="メインカラー (HEX)")
  sub_color: str = Field(description="サブカラー (HEX)")
  accent_color: str = Field(description="アクセントカラー (HEX)")
  chart_colors: list[str] = Field(description="グラフ用5色のHEX")


class DesignSystem(BaseModel):
  color_palette: ColorPalette
  font_impression: str = Field(description="フォントの印象")
  tone_and_manner: str = Field(description="トーン＆マナー")


class CompanyInfo(BaseModel):
  company_name: str = Field(description="企業名")
  mission_vision: str = Field(description="理念・ビジョン")
  medium_term_plan: str = Field(
      description="中期経営計画要点（PDFが無い場合はHPからの推定）"
  )


class BrandingMasterData(BaseModel):
  company_info: CompanyInfo
  design_system: DesignSystem


USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML,"
    " like Gecko) Chrome/120.0.0.0 Safari/537.36"
)


# --------------------------------------------------
# Step 1: データ収集処理 (Playwright強化版 + 自動PDF探索)
# --------------------------------------------------
async def collect_assets_flexible(
    hp_url: str,
    pdf_url: str = None,
    uploaded_pdf_bytes: bytes = None,
    output_image="hp_screenshot.png",
    output_pdf="ir_document.pdf",
):
  """HPスクショ撮影および、各種パターン（手動ファイル・直リンク・HP自動検索）でのPDF取得"""
  pdf_saved = False

  # パターンA: ユーザーがファイルを直接アップロードした場合
  if uploaded_pdf_bytes:
    with open(output_pdf, "wb") as f:
      f.write(uploaded_pdf_bytes)
    pdf_saved = True

  async with async_playwright() as p:
    browser = await p.chromium.launch(headless=True)
    context = await browser.new_context(
        user_agent=USER_AGENT,
        viewport={"width": 1920, "height": 1080},
        extra_http_headers={
            "Accept-Language": "ja,en-US;q=0.9,en;q=0.8",
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8"
            ),
        },
    )

    page = await context.new_page()

    # 1. HPスクショ撮影
    await page.goto(hp_url, wait_until="domcontentloaded", timeout=40000)
    await page.wait_for_timeout(2000)
    await page.evaluate("window.scrollTo(0, 500)")
    await page.wait_for_timeout(1000)
    await page.evaluate("window.scrollTo(0, 0)")
    await page.wait_for_timeout(1000)
    await page.screenshot(path=output_image)

    # パターンB: PDF直URLが指定されている場合
    if not pdf_saved and pdf_url:
      try:
        response = await context.request.get(pdf_url)
        if response.ok:
          pdf_bytes = await response.body()
          with open(output_pdf, "wb") as f:
            f.write(pdf_bytes)
          pdf_saved = True
      except Exception as e:
        st.warning(
            f"指定されたPDF直リンクからのダウンロードに失敗しました: {e}"
        )

    # パターンC: PDF指定がなく、HPページ内からPDFリンクを自動探索する場合
    if not pdf_saved and not pdf_url and not uploaded_pdf_bytes:
      try:
        hrefs = await page.eval_on_selector_all(
            "a[href]", "elements => elements.map(e => e.href)"
        )
        pdf_links = [
            h
            for h in hrefs
            if h.lower().endswith(".pdf")
            or "ir" in h.lower()
            and ".pdf" in h.lower()
        ]

        if pdf_links:
          target_pdf_url = pdf_links[0]
          response = await context.request.get(target_pdf_url)
          if response.ok:
            pdf_bytes = await response.body()
            with open(output_pdf, "wb") as f:
              f.write(pdf_bytes)
            pdf_saved = True
      except Exception as e:
        pass

    await browser.close()

  return pdf_saved


# --------------------------------------------------
# Step 2: Gemini マルチモーダル解析 (PDF有無に自動対応)
# --------------------------------------------------
def analyze_branding(api_key: str, has_pdf: bool):
  client = genai.Client(api_key=api_key)

  with open("hp_screenshot.png", "rb") as f:
    image_bytes = f.read()

  contents = [types.Part.from_bytes(data=image_bytes, mime_type="image/png")]

  if has_pdf and os.path.exists("ir_document.pdf"):
    uploaded_pdf = client.files.upload(file="ir_document.pdf")
    contents.append(uploaded_pdf)
    pdf_instruction = (
        "添付の「企業HPスクリーンショット」および「IR資料（PDF）」を統合分析し、"
    )
  else:
    pdf_instruction = (
        "添付の「企業HPスクリーンショット」を重点的に分析し（IR資料PDFは無しためHP情報中心）、"
    )

  prompt = f"""
    あなたはブランディングデザイナー兼経営アナリストです。
    {pdf_instruction}
    1. 企業のビジョン・理念、中期経営計画（または事業概要）の要点
    2. ブランドカラー（メイン、サブ、アクセント、グラフ用5色）
    3. トーン＆マナーとフォントの印象
    を抽出してJSONで出力してください。
    """

  contents.append(prompt)

  response = client.models.generate_content(
      model="models/gemini-3.5-flash",
      contents=contents,
      config=types.GenerateContentConfig(
          response_mime_type="application/json",
          response_schema=BrandingMasterData,
      ),
  )

  return json.loads(response.text)


# --------------------------------------------------
# Step 3: PowerPoint自動作成 (PermissionError対策付き)
# --------------------------------------------------
def hex_to_rgb(hex_str: str) -> RGBColor:
  hex_str = hex_str.lstrip("#")
  if len(hex_str) != 6:
    hex_str = "000000"
  return RGBColor(*(int(hex_str[i : i + 2], 16) for i in (0, 2, 4)))


def generate_pptx(data: dict, output_path="branding_report.pptx"):
  company = data["company_info"]
  design = data["design_system"]
  palette = design["color_palette"]

  main_rgb = hex_to_rgb(palette["main_color"])
  sub_rgb = hex_to_rgb(palette["sub_color"])

  prs = Presentation()
  prs.slide_width = Inches(13.333)
  prs.slide_height = Inches(7.5)
  blank_layout = prs.slide_layouts[6]

  # --- Slide 1: 表紙 ---
  slide1 = prs.slides.add_slide(blank_layout)
  bar = slide1.shapes.add_shape(
      MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.3)
  )
  bar.fill.solid()
  bar.fill.fore_color.rgb = main_rgb
  bar.line.fill.background()

  tx = slide1.shapes.add_textbox(
      Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0)
  )
  tf = tx.text_frame
  tf.word_wrap = True
  p = tf.paragraphs[0]
  p.text = f"{company['company_name']}\nブランディング & 経営分析レポート"
  p.font.size = Pt(40)
  p.font.bold = True
  p.font.color.rgb = sub_rgb

  # --- Slide 2: ビジョン & 成長戦略 ---
  slide2 = prs.slides.add_slide(blank_layout)
  t_box = slide2.shapes.add_textbox(
      Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.0)
  )
  p = t_box.text_frame.paragraphs[0]
  p.text = "01. 企業ビジョン ＆ 成長戦略"
  p.font.size = Pt(28)
  p.font.bold = True
  p.font.color.rgb = main_rgb

  c1 = slide2.shapes.add_shape(
      MSO_SHAPE.ROUNDED_RECTANGLE,
      Inches(0.8),
      Inches(1.8),
      Inches(5.6),
      Inches(4.8),
  )
  c1.fill.solid()
  c1.fill.fore_color.rgb = RGBColor(245, 247, 250)
  tf1 = c1.text_frame
  tf1.word_wrap = True
  tf1.margin_left = tf1.margin_right = tf1.margin_top = Inches(0.3)
  p = tf1.paragraphs[0]
  p.text = "【ミッション・ビジョン】\n"
  p.font.bold = True
  p.font.size = Pt(18)
  p.add_run().text = company["mission_vision"]

  c2 = slide2.shapes.add_shape(
      MSO_SHAPE.ROUNDED_RECTANGLE,
      Inches(6.8),
      Inches(1.8),
      Inches(5.7),
      Inches(4.8),
  )
  c2.fill.solid()
  c2.fill.fore_color.rgb = RGBColor(245, 247, 250)
  tf2 = c2.text_frame
  tf2.word_wrap = True
  tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.3)
  p = tf2.paragraphs[0]
  p.text = "【事業概要 / 中期経営計画】\n"
  p.font.bold = True
  p.font.size = Pt(18)
  p.add_run().text = company["medium_term_plan"]

  # --- Slide 3: デザインシステム ---
  slide3 = prs.slides.add_slide(blank_layout)
  t_box = slide3.shapes.add_textbox(
      Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.0)
  )
  p = t_box.text_frame.paragraphs[0]
  p.text = "02. ブランドデザインシステム"
  p.font.size = Pt(28)
  p.font.bold = True
  p.font.color.rgb = main_rgb

  for i, hex_code in enumerate(palette["chart_colors"]):
    left = Inches(0.8 + i * 2.3)
    rect = slide3.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.2), Inches(2.0), Inches(1.3)
    )
    rect.fill.solid()
    rect.fill.fore_color.rgb = hex_to_rgb(hex_code)
    rect.line.fill.background()

    lbl = slide3.shapes.add_textbox(
        left, Inches(3.6), Inches(2.0), Inches(0.5)
    )
    p = lbl.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = hex_code
    p.font.bold = True

  # PowerPointが開かれている場合の上書き拒否エラー（PermissionError）を防止
  try:
    prs.save(output_path)
  except PermissionError:
    timestamp = int(time.time())
    output_path = f"branding_report_{timestamp}.pptx"
    prs.save(output_path)

  return output_path


# --------------------------------------------------
# Streamlit UI 画面設計
# --------------------------------------------------
st.set_page_config(
    page_title="ブランド＆経営分析 AIエージェント", page_icon="🎨", layout="wide"
)

st.title("🎨 企業ブランド＆経営分析 AIレポートジェネレーター")
st.caption(
    "企業HPのURLを入力するだけで、Gemini"
    " AIがブランディングを自動解析しPowerPoint資料を生成します。"
)

# サイドバー設定
with st.sidebar:
  st.header("⚙️ 設定")
  default_api_key = os.environ.get("GEMINI_API_KEY", "")
  api_key = st.text_input(
      "Gemini API Key", value=default_api_key, type="password"
  )

st.divider()

# メインフォーム
hp_url = st.text_input(
    "1. 企業HP URL（必須）", placeholder="https://www.sony.com/ja/"
)

st.write("---")
st.subheader("📄 参照資料 (IR・提案書等) の指定（任意）")
st.info(
    "※"
    " PDFを指定しない場合、HP上のリンクから自動探索するか、HP情報のみで分析を実行します。"
)

pdf_tab1, pdf_tab2 = st.tabs(["ファイルを直接アップロード", "PDF直リンクURLを指定"])

with pdf_tab1:
  uploaded_pdf = st.file_uploader(
      "IR資料や提案書PDFをアップロード", type=["pdf"]
  )

with pdf_tab2:
  pdf_url = st.text_input(
      "IR資料 (PDF) の直リンクURL",
      placeholder=(
          "https://www.sony.com/ja/SonyInfo/IR/library/presen/er/pdf/23q4_sony.pdf"
      ),
  )

run_button = st.button(
    "🚀 分析＆スライド生成を実行", type="primary", use_container_width=True
)

# 実行ロジック
if run_button:
  if not api_key:
    st.error("サイドバーから Gemini API Key を入力してください。")
  elif not hp_url:
    st.warning("企業HPのURLを入力してください。")
  else:
    status = st.status("分析プロセスを実行中...", expanded=True)

    try:
      # Step 1: 収集
      status.write("1/3 Webサイトのキャプチャと参照資料を取得中...")

      if os.path.exists("ir_document.pdf"):
        try:
          os.remove("ir_document.pdf")
        except:
          pass

      uploaded_bytes = uploaded_pdf.read() if uploaded_pdf else None
      has_pdf = asyncio.run(
          collect_assets_flexible(
              hp_url=hp_url,
              pdf_url=pdf_url if pdf_url else None,
              uploaded_pdf_bytes=uploaded_bytes,
          )
      )

      if has_pdf:
        status.write("   ✓ HPスクリーンショット ＆ 参照PDFの準備が完了しました！")
      else:
        status.write(
            "   ✓ HPスクリーンショットの準備が完了しました（※参照PDFなしで進行）"
        )

      # Step 2: Gemini解析
      status.write("2/3 Gemini 3.5 Flash による統合分析を実施中...")
      data = analyze_branding(api_key, has_pdf=has_pdf)
      status.write("   ✓ 解析＆カラーパレット抽出完了！")

      # Step 3: PPT作成
      status.write("3/3 PowerPoint スライドを自動構築中...")
      pptx_file = generate_pptx(data)
      status.write("   ✓ スライド生成完了！")

      status.update(label="🎉 すべての処理が正常に完了しました！", state="complete")

      st.success("レポートが作成されました！以下から確認・ダウンロードできます。")

      # プレビュー表示
      res_col1, res_col2 = st.columns(2)
      with res_col1:
        st.subheader("📸 キャプチャ画像")
        st.image("hp_screenshot.png", use_container_width=True)

      with res_col2:
        st.subheader("📊 抽出されたブランドデータ")
        st.json(data)

      # ダウンロードボタン
      with open(pptx_file, "rb") as f:
        st.download_button(
            label="📥 完成したPowerPoint資料(.pptx)をダウンロード",
            data=f,
            file_name=f"{data['company_info']['company_name']}_ブランディングレポート.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )

    except Exception as e:
      status.update(label="✗ エラーが発生しました", state="error")
      st.error(f"エラー詳細: {e}")