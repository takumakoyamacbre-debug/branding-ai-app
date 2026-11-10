import asyncio
import json
import os
import time
import urllib.parse
from google import genai
from google.genai import types
from playwright.async_api import async_playwright
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches, Pt
from pydantic import BaseModel, Field
import streamlit as st

# --------------------------------------------------
# Pydantic スキーマ定義（12色テーマカラー対応）
# --------------------------------------------------
class ThemeColorPalette(BaseModel):
  palette_name: str = Field(description="パレットのコンセプト名（例: モダンクラシック）")
  dark_1: str = Field(description="濃色1 (HEX) - 主にテキストや重要要素")
  dark_2: str = Field(description="濃色2 (HEX) - サブテキストや枠線")
  light_1: str = Field(description="淡色1 (HEX) - 背景やカードのベース")
  light_2: str = Field(description="淡色2 (HEX) - アクセント背景・薄い塗りつぶし")
  accent_1: str = Field(description="アクセント1 (HEX) - メインブランドカラー")
  accent_2: str = Field(description="アクセント2 (HEX)")
  accent_3: str = Field(description="アクセント3 (HEX)")
  accent_4: str = Field(description="アクセント4 (HEX)")
  accent_5: str = Field(description="アクセント5 (HEX)")
  accent_6: str = Field(description="アクセント6 (HEX)")
  hyperlink: str = Field(description="ハイパーリンクカラー (HEX)")
  followed_hyperlink: str = Field(description="表示済みハイパーリンクカラー (HEX)")


class DesignSystemOptions(BaseModel):
  color_palettes: list[ThemeColorPalette] = Field(description="異なるトーンの3つのカラーパレット案")
  font_impression: str = Field(description="フォントの印象")
  tone_and_manner: str = Field(description="トーン＆マナー")


class CompanyInfo(BaseModel):
  company_name: str = Field(description="企業名")
  mission_vision: str = Field(description="理念・ビジョン")
  medium_term_plan: str = Field(description="中期経営計画要点")


class BrandingMasterData(BaseModel):
  company_info: CompanyInfo
  design_system: DesignSystemOptions


USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"


# --------------------------------------------------
# Step 1: データ収集処理
# --------------------------------------------------
async def collect_assets_flexible(hp_url: str, pdf_url: str = None, uploaded_pdf_bytes: bytes = None, output_image="hp_screenshot.png", output_pdf="ir_document.pdf"):
  pdf_saved = False
  if uploaded_pdf_bytes:
    with open(output_pdf, "wb") as f:
      f.write(uploaded_pdf_bytes)
    pdf_saved = True

  async with async_playwright() as p:
    browser = await p.chromium.launch(headless=True, args=["--no-sandbox", "--disable-setuid-sandbox"])
    context = await browser.new_context(user_agent=USER_AGENT, viewport={"width": 1920, "height": 1080})
    page = await context.new_page()

    await page.goto(hp_url, wait_until="domcontentloaded", timeout=40000)
    await page.wait_for_timeout(2000)
    await page.screenshot(path=output_image)

    if not pdf_saved and pdf_url:
      try:
        response = await context.request.get(pdf_url)
        if response.ok:
          with open(output_pdf, "wb") as f:
            f.write(await response.body())
          pdf_saved = True
      except:
        pass

    if not pdf_saved and not pdf_url and not uploaded_pdf_bytes:
      try:
        hrefs = await page.eval_on_selector_all("a[href]", "elements => elements.map(e => e.href)")
        pdf_links = [h for h in hrefs if h.lower().endswith(".pdf") or "ir" in h.lower() and ".pdf" in h.lower()]
        if pdf_links:
          response = await context.request.get(pdf_links[0])
          if response.ok:
            with open(output_pdf, "wb") as f:
              f.write(await response.body())
            pdf_saved = True
      except:
        pass

    await browser.close()
  return pdf_saved


# --------------------------------------------------
# Step 2: Gemini マルチモーダル解析 (複数パレット生成)
# --------------------------------------------------
def analyze_branding_multi_palettes(api_key: str, has_pdf: bool):
  client = genai.Client(api_key=api_key)

  with open("hp_screenshot.png", "rb") as f:
    image_bytes = f.read()

  contents = [types.Part.from_bytes(data=image_bytes, mime_type="image/png")]
  if has_pdf and os.path.exists("ir_document.pdf"):
    uploaded_pdf = client.files.upload(file="ir_document.pdf")
    contents.append(uploaded_pdf)
    instruction = "添付の「企業HPスクリーンショット」および「IR資料（PDF）」を統合分析し、"
  else:
    instruction = "添付の「企業HPスクリーンショット」を重点的に分析し、"

  prompt = f"""
    あなたはプロのブランディングデザイナーです。
    {instruction}
    1. 企業のビジョン・理念、中期経営計画（事業概要）の要点
    2. PowerPointのマスター設定を想定した「12色（濃色2色、淡色2色、アクセント6色、ハイパーリンク1色、訪問済みハイパーリンク1色）」からなるカラーパレット案を、方向性が異なる3パターン作成してください。
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
# Step 3: PowerPoint自動作成 (選択された12色パレット反映)
# --------------------------------------------------
def hex_to_rgb(hex_str: str) -> RGBColor:
  hex_str = hex_str.lstrip("#")
  if len(hex_str) != 6:
    hex_str = "000000"
  return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

def generate_pptx_with_palette(data: dict, selected_palette_index: int, output_path="branding_report.pptx"):
  company = data["company_info"]
  design = data["design_system"]
  palette = design["color_palettes"][selected_palette_index]

  dark_1 = hex_to_rgb(palette["dark_1"])
  light_1 = hex_to_rgb(palette["light_1"])
  accent_1 = hex_to_rgb(palette["accent_1"])

  prs = Presentation()
  prs.slide_width = Inches(13.333)
  prs.slide_height = Inches(7.5)
  blank_layout = prs.slide_layouts[6]

  # --- Slide 1: 表紙 ---
  slide1 = prs.slides.add_slide(blank_layout)
  bar = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.3))
  bar.fill.solid()
  bar.fill.fore_color.rgb = accent_1
  bar.line.fill.background()

  tx = slide1.shapes.add_textbox(Inches(1.0), Inches(2.2), Inches(11.333), Inches(3.0))
  tf = tx.text_frame
  tf.word_wrap = True
  p = tf.paragraphs[0]
  p.text = f"{company['company_name']}\nブランディング & 経営分析レポート"
  p.font.size = Pt(40)
  p.font.bold = True
  p.font.color.rgb = dark_1

  # --- Slide 2: ビジョン & 成長戦略 ---
  slide2 = prs.slides.add_slide(blank_layout)
  t_box = slide2.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.0))
  p = t_box.text_frame.paragraphs[0]
  p.text = "01. 企業ビジョン ＆ 成長戦略"
  p.font.size = Pt(28)
  p.font.bold = True
  p.font.color.rgb = dark_1

  c1 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
  c1.fill.solid()
  c1.fill.fore_color.rgb = light_1
  tf1 = c1.text_frame
  tf1.word_wrap = True
  tf1.margin_left = tf1.margin_right = tf1.margin_top = Inches(0.3)
  p = tf1.paragraphs[0]
  p.text = "【ミッション・ビジョン】\n"
  p.font.bold = True
  p.font.size = Pt(18)
  p.font.color.rgb = dark_1
  p.add_run().text = company["mission_vision"]

  c2 = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.8), Inches(1.8), Inches(5.7), Inches(4.8))
  c2.fill.solid()
  c2.fill.fore_color.rgb = light_1
  tf2 = c2.text_frame
  tf2.word_wrap = True
  tf2.margin_left = tf2.margin_right = tf2.margin_top = Inches(0.3)
  p = tf2.paragraphs[0]
  p.text = "【事業概要 / 中期経営計画】\n"
  p.font.bold = True
  p.font.size = Pt(18)
  p.font.color.rgb = dark_1
  p.add_run().text = company["medium_term_plan"]

  # --- Slide 3: 選択されたカラーパレット一覧確認 ---
  slide3 = prs.slides.add_slide(blank_layout)
  t_box = slide3.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.5), Inches(1.0))
  p = t_box.text_frame.paragraphs[0]
  p.text = f"02. 採用されたテーマカラー ({palette['palette_name']})"
  p.font.size = Pt(28)
  p.font.bold = True
  p.font.color.rgb = dark_1

  accents = [palette["accent_1"], palette["accent_2"], palette["accent_3"], palette["accent_4"], palette["accent_5"], palette["accent_6"]]
  for i, hex_code in enumerate(accents):
    left = Inches(0.8 + i * 1.9)
    rect = slide3.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, Inches(2.2), Inches(1.7), Inches(1.5))
    rect.fill.solid()
    rect.fill.fore_color.rgb = hex_to_rgb(hex_code)
    rect.line.fill.background()

    lbl = slide3.shapes.add_textbox(left, Inches(3.8), Inches(1.7), Inches(0.5))
    p = lbl.text_frame.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    p.text = f"Acc {i+1}\n{hex_code}"
    p.font.size = Pt(12)
    p.font.bold = True

  try:
    prs.save(output_path)
  except PermissionError:
    timestamp = int(time.time())
    output_path = f"branding_report_{timestamp}.pptx"
    prs.save(output_path)

  return output_path


# --------------------------------------------------
# Streamlit UI 画面設計 (セッション状態を使った2段階フロー)
# --------------------------------------------------
st.set_page_config(page_title="ブランド＆経営分析 AIエージェント", page_icon="🎨", layout="wide")

st.title("🎨 企業ブランド＆経営分析 AIレポートジェネレーター")
st.caption("企業HPを入力すると、Geminiが複数パターンのマスターカラー案を提案し、選んだ色でスライドを生成します。")

with st.sidebar:
  st.header("⚙️ 設定")
  default_api_key = os.environ.get("GEMINI_API_KEY", "")
  api_key = st.text_input("Gemini API Key", value=default_api_key, type="password")

st.divider()

# セッション状態の初期化
if "analyzed_data" not in st.session_state:
  st.session_state.analyzed_data = None

hp_url = st.text_input("1. 企業HP URL（必須）", placeholder="https://www.sony.com/ja/")

st.write("---")
st.subheader("📄 参照資料 (IR・提案書等) の指定（任意）")
pdf_tab1, pdf_tab2 = st.tabs(["ファイルを直接アップロード", "PDF直リンクURLを指定"])
with pdf_tab1:
  uploaded_pdf = st.file_uploader("IR資料や提案書PDFをアップロード", type=["pdf"])
with pdf_tab2:
  pdf_url = st.text_input("IR資料 (PDF) の直リンクURL")

# --- Step 1: 分析実行ボタン ---
if st.button("🔍 カラーパレット案を分析・生成する", type="primary", use_container_width=True):
  if not api_key:
    st.error("サイドバーから Gemini API Key を入力してください。")
  elif not hp_url:
    st.warning("企業HPのURLを入力してください。")
  else:
    with st.status("分析プロセスを実行中...", expanded=True) as status:
      try:
        status.write("1/2 Webサイトのキャプチャと参照資料を取得中...")
        uploaded_bytes = uploaded_pdf.read() if uploaded_pdf else None
        has_pdf = asyncio.run(collect_assets_flexible(hp_url=hp_url, pdf_url=pdf_url if pdf_url else None, uploaded_pdf_bytes=uploaded_bytes))

        status.write("2/2 Geminiによる3つのカラーパレット案を解析中...")
        data = analyze_branding_multi_palettes(api_key, has_pdf=has_pdf)
        
        st.session_state.analyzed_data = data
        status.update(label="✨ カラーパレット案の生成が完了しました！", state="complete")
      except Exception as e:
        status.update(label="✗ エラーが発生しました", state="error")
        st.error(f"エラー詳細: {e}")

# --- Step 2: パレット選択 & スライド生成セクション ---
if st.session_state.analyzed_data:
  st.divider()
  st.subheader("🎨 生成されたカラーパレット案を選択してください")
  
  data = st.session_state.analyzed_data
  palettes = data["design_system"]["color_palettes"]

  # ラジオボタン用に選択肢を作成
  palette_options = [f"【{p['palette_name']}】 (アクセント主色: {p['accent_1']})" for p in palettes]
  selected_label = st.radio("パレット案", options=palette_options)
  selected_index = palette_options.index(selected_label)

  # 選択されたパレットの詳細を視覚的に表示
  chosen_p = palettes[selected_index]
  cols = st.columns(6)
  accents = [chosen_p["accent_1"], chosen_p["accent_2"], chosen_p["accent_3"], chosen_p["accent_4"], chosen_p["accent_5"], chosen_p["accent_6"]]
  for idx, col in enumerate(cols):
    with col:
      st.markdown(f"<div style='background-color:{accents[idx]}; height:40px; border-radius:5px;'></div>", unsafe_allow_html=True)
      st.caption(f"Acc {idx+1}")

  st.write("")
  if st.button("🚀 このパレットでPowerPointスライドを生成する", type="primary", use_container_width=True):
    with st.spinner("PowerPointスライドを構築中..."):
      pptx_file = generate_pptx_with_palette(data, selected_index)
      st.success("スライドが正常に生成されました！")

      with open(pptx_file, "rb") as f:
        st.download_button(
            label="📥 完成したPowerPoint資料(.pptx)をダウンロード",
            data=f,
            file_name=f"{data['company_info']['company_name']}_ブランディングレポート.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            use_container_width=True,
        )
