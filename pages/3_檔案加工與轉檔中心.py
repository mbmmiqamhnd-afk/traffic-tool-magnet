import io
import os
import zipfile
import streamlit as st
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import white, black
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from pdf2image import convert_from_bytes
from menu import show_sidebar

st.set_page_config(page_title="檔案加工與轉檔中心", page_icon="🗂️", layout="wide")
show_sidebar()

st.header("🗂️ 綜合檔案加工與轉檔中心（通用版）")
st.write("支援：自訂單位署名角標遮蓋與頁碼添加、PDF 轉 PPTX／圖片，以及多圖合併轉 PDF。")


def get_font_path():
    possible_paths = [
        "NotoSansTC-Regular.ttf",
        "fonts/NotoSansTC-Regular.ttf",
        "kaiu.ttf",
        "font.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf",
        "C:/Windows/Fonts/kaiu.ttf",
    ]
    for p in possible_paths:
        if os.path.exists(p):
            return p
    return None


font_path = get_font_path()


def set_east_asian_font(run, font_name):
    rPr = run._r.get_or_add_rPr()
    ea = rPr.find(qn("a:ea"))
    if ea is None:
        ea = rPr.makeelement(qn("a:ea"), {})
        rPr.append(ea)
    ea.set("typeface", font_name)


def create_pdf_overlay(page_width, page_height, page_num, current_font, label_prefix, show_page_num):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_width, page_height))

    parts = []
    if label_prefix.strip():
        parts.append(label_prefix.strip())
    if show_page_num:
        parts.append(f"第 {page_num} 頁")
    text = " - ".join(parts)

    font_size = 14
    text_width = c.stringWidth(text, current_font, font_size) if text else 80
    box_width = max(text_width + 18, 80)
    box_height = 22

    c.setFillColor(white)
    c.rect(page_width - box_width, 0, box_width, box_height, fill=1, stroke=0)

    if text:
        c.setFillColor(black)
        c.setFont(current_font, font_size)
        c.drawRightString(page_width - 8, 5, text)

    c.save()
    packet.seek(0)
    return packet


def process_image(image_file, font_p, label_prefix):
    img = Image.open(image_file).convert("RGB")
    draw = ImageDraw.Draw(img)
    width, height = img.size

    box_w, box_h = 260, 50
    rect_x0, rect_y0 = width - box_w, height - box_h
    draw.rectangle([rect_x0, rect_y0, width, height], fill="white")

    if label_prefix.strip():
        try:
            font = ImageFont.truetype(font_p, 24) if font_p else ImageFont.load_default()
        except Exception:
            font = ImageFont.load_default()
        draw.text((rect_x0 + 10, rect_y0 + 10), label_prefix.strip(), fill="black", font=font)

    img_byte_arr = io.BytesIO()
    img.save(img_byte_arr, format="JPEG", quality=95)
    return img_byte_arr.getvalue()


def process_pptx(pptx_file, font_p, label_prefix, show_page_num):
    prs = Presentation(pptx_file)
    slide_width = prs.slide_width
    slide_height = prs.slide_height
    font_size_pt = 14
    padding_pt = 8
    box_height_pt = 22

    def measure_text_width_pt(text, size_pt):
        try:
            px_size = int(size_pt * 96 / 72)
            font = ImageFont.truetype(font_p, px_size) if font_p else ImageFont.load_default()
            return font.getlength(text) * 72 / 96
        except Exception:
            return len(text) * size_pt

    for i, slide in enumerate(prs.slides):
        parts = []
        if label_prefix.strip():
            parts.append(label_prefix.strip())
        if show_page_num:
            parts.append(f"第 {i + 1} 頁")
        text = " - ".join(parts)

        text_width_pt = measure_text_width_pt(text, font_size_pt) if text else 80
        box_width = Pt(max(text_width_pt + padding_pt * 2, 80))
        box_height = Pt(box_height_pt)

        shape = slide.shapes.add_shape(1, slide_width - box_width, slide_height - box_height, box_width, box_height)
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shape.line.fill.background()
        shape.shadow.inherit = False

        if text:
            tf = shape.text_frame
            tf.margin_left = Pt(0)
            tf.margin_right = Pt(padding_pt)
            tf.margin_top = Pt(0)
            tf.margin_bottom = Pt(0)
            tf.word_wrap = False
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.RIGHT
            run = p.add_run()
            run.text = text
            run.font.size = Pt(font_size_pt)
            run.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
            run.font.name = "Microsoft JhengHei"
            set_east_asian_font(run, "Microsoft JhengHei")

    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    return out.getvalue()


tab1, tab2, tab3 = st.tabs(["📝 自訂署名遮蓋與頁碼", "📄 PDF 轉 PPTX / 圖片", "🖼️ 多圖合併轉 PDF"])

with tab1:
    st.subheader("檔案右下角商標遮蓋與自訂頁碼添加")
    c_lbl, c_pg = st.columns([2, 1])
    with c_lbl:
        custom_label = st.text_input("右下角署名文字（留空則僅做純白遮蓋）", value="交通組製")
    with c_pg:
        st.write("")
        st.write("")
        include_page_num = st.checkbox("加上頁碼（僅 PDF / PPTX 適用）", value=True)

    watermark_file = st.file_uploader("上傳 PDF、PPTX 或圖片 (JPG/PNG)", type=["pdf", "pptx", "jpg", "jpeg", "png"], key="watermark")

    if watermark_file and st.button("🚀 開始加工", type="primary"):
        file_ext = watermark_file.name.split(".")[-1].lower()
        try:
            if file_ext == "pdf":
                reader = PdfReader(watermark_file)
                writer = PdfWriter()
                pdf_font = "Helvetica"
                if font_path and font_path.endswith(".ttf"):
                    try:
                        pdfmetrics.registerFont(TTFont("CustomFont", font_path))
                        pdf_font = "CustomFont"
                    except Exception:
                        pass
                for i, page in enumerate(reader.pages):
                    w, h = float(page.mediabox.width), float(page.mediabox.height)
                    overlay = create_pdf_overlay(w, h, i + 1, pdf_font, custom_label, include_page_num)
                    page.merge_page(PdfReader(overlay).pages[0])
                    writer.add_page(page)
                out = io.BytesIO()
                writer.write(out)
                st.success("🎉 PDF 加工完成！")
                st.download_button("📥 下載加工版 PDF", out.getvalue(), f"加工版_{watermark_file.name}", "application/pdf", use_container_width=True)

            elif file_ext == "pptx":
                result_pptx = process_pptx(watermark_file, font_path, custom_label, include_page_num)
                st.success("🎉 PPTX 加工完成！")
                st.download_button(
                    "📥 下載加工版 PPTX",
                    result_pptx,
                    f"加工版_{watermark_file.name}",
                    "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
            else:
                result_img = process_image(watermark_file, font_path, custom_label)
                st.image(result_img, caption="預覽加工後的圖片")
                st.success("🎉 圖片加工完成！")
                st.download_button("📥 下載加工版圖片", result_img, f"加工版_{watermark_file.name}", f"image/{file_ext}", use_container_width=True)
        except Exception as e:
            st.error(f"發生錯誤：{e}")

with tab2:
    st.subheader("將 PDF 轉為 PPTX 或圖片壓縮檔")
    pdf_convert_file = st.file_uploader("請上傳 PDF 檔案", type=["pdf"], key="pdf_convert")
    if pdf_convert_file:
        c_opt, c_dpi = st.columns(2)
        option = c_opt.radio("轉換格式：", ("轉成 PPTX", "轉成圖片 (ZIP壓縮檔)"))
        dpi_val = c_dpi.selectbox("輸出解析度 (DPI)", [150, 200], index=0)
        if st.button(f"🚀 開始{option}", type="primary"):
            with st.spinner("正在解析 PDF..."):
                try:
                    images = convert_from_bytes(pdf_convert_file.read(), dpi=dpi_val)
                    if images:
                        if option == "轉成 PPTX":
                            prs = Presentation()
                            w, h = images[0].size
                            prs.slide_width = int(w * 914400 / dpi_val)
                            prs.slide_height = int(h * 914400 / dpi_val)
                            for img in images:
                                slide = prs.slides.add_slide(prs.slide_layouts[6])
                                s = io.BytesIO()
                                img.save(s, format="JPEG", quality=88)
                                s.seek(0)
                                slide.shapes.add_picture(s, 0, 0, width=prs.slide_width, height=prs.slide_height)
                            pptx_out = io.BytesIO()
                            prs.save(pptx_out)
                            st.success(f"✅ PPTX 轉換成功（共 {len(images)} 頁）")
                            st.download_button(
                                "📥 下載 PPTX",
                                pptx_out.getvalue(),
                                f"{pdf_convert_file.name.rsplit('.', 1)[0]}.pptx",
                                "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                use_container_width=True,
                            )
                        else:
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "a", zipfile.ZIP_DEFLATED, False) as zf:
                                for i, img in enumerate(images):
                                    s = io.BytesIO()
                                    img.save(s, format="JPEG", quality=95)
                                    zf.writestr(f"page_{i + 1}.jpg", s.getvalue())
                            st.success(f"✅ 圖片轉換成功（共 {len(images)} 張）")
                            st.download_button(
                                "📥 下載圖片 ZIP",
                                zip_buf.getvalue(),
                                f"{pdf_convert_file.name.rsplit('.', 1)[0]}_images.zip",
                                "application/zip",
                                use_container_width=True,
                            )
                except Exception as e:
                    st.error(f"❌ 發生錯誤：{e}")

with tab3:
    st.subheader("將多張圖片合併為單一 PDF")
    img_files = st.file_uploader("請上傳圖片（支援多選，依上傳順序排列）", type=["jpg", "jpeg", "png", "bmp"], accept_multiple_files=True, key="img_to_pdf")
    if img_files:
        st.info(f"已選擇 {len(img_files)} 張圖片。")
        if st.button("🚀 開始合併為 PDF", type="primary"):
            try:
                imgs = [Image.open(f).convert("RGB") for f in img_files]
                pdf_buf = io.BytesIO()
                imgs[0].save(pdf_buf, format="PDF", save_all=True, append_images=imgs[1:], resolution=100.0)
                st.success(f"✅ PDF 建立成功（共 {len(imgs)} 頁）")
                st.download_button("📥 下載合併後的 PDF", pdf_buf.getvalue(), "merged_images.pdf", "application/pdf", use_container_width=True)
            except Exception as e:
                st.error(f"❌ 發生錯誤：{e}")
