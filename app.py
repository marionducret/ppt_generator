import copy
import io
import os
import uuid
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches

TEMPLATE_PATH = "templates.pptx"
PROTOTYPE_COUNT = 3
SLIDE_SIZE = (1280, 720)
CONTENT_BOUNDS = {
    "x_min": 0.6,
    "x_max": 9.2,
    "y_min": 1.15,
    "y_max": 6.7,
}

PLACEHOLDERS = {
    "Titre général": {"prototype_index": 0, "token": "{{TITLE_GENERAL}}"},
    "Titre intermédiaire": {"prototype_index": 1, "token": "{{TITLE_SECTION}}"},
    "Contenu": {"prototype_index": 2, "token": "{{CONTENT_TITLE}}"},
}


# ---------- Helpers PPT ----------
def duplicate_slide(prs: Presentation, source_slide_index: int):
    source = prs.slides[source_slide_index]
    layout = prs.slide_layouts[0] if prs.slide_layouts else prs.slide_master.slide_layouts[0]
    new_slide = prs.slides.add_slide(layout)

    # remove default placeholders from added slide
    for shape in list(new_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)

    for shape in source.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide


def replace_token_in_slide(slide, token: str, value: str) -> bool:
    replaced = False
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        if token not in shape.text:
            continue

        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if token in run.text:
                    run.text = run.text.replace(token, value)
                    replaced = True
        if not replaced:
            shape.text = shape.text.replace(token, value)
            replaced = True
    return replaced


def add_textbox(slide, item):
    box = slide.shapes.add_textbox(
        Inches(item["x"]), Inches(item["y"]), Inches(item["w"]), Inches(item["h"])
    )
    tf = box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = item["text"]

    run = p.runs[0]
    run.font.size = item["font_size_pt"]
    run.font.bold = item["bold"]
    run.font.name = item["font_name"]


def add_image(slide, item):
    slide.shapes.add_picture(
        item["path"],
        Inches(item["x"]),
        Inches(item["y"]),
        width=Inches(item["w"]),
        height=Inches(item["h"]),
    )


def remove_first_n_slides(prs: Presentation, n: int):
    # remove prototype slides from final export
    for _ in range(min(n, len(prs.slides))):
        slide_id = prs.slides._sldIdLst[0]
        r_id = slide_id.rId
        prs.part.drop_rel(r_id)
        del prs.slides._sldIdLst[0]


def build_pptx(slides_data, output_path: str):
    prs = Presentation(TEMPLATE_PATH)

    for slide_data in slides_data:
        spec = PLACEHOLDERS[slide_data["type"]]
        slide = duplicate_slide(prs, spec["prototype_index"])

        ok = replace_token_in_slide(slide, spec["token"], slide_data["title"])
        if not ok:
            raise ValueError(
                f"Repère introuvable dans le template pour {slide_data['type']} : {spec['token']}"
            )

        if slide_data["type"] == "Contenu":
            for item in slide_data.get("items", []):
                if item["kind"] == "text":
                    add_textbox(slide, item)
                elif item["kind"] == "image":
                    add_image(slide, item)

    remove_first_n_slides(prs, PROTOTYPE_COUNT)
    prs.save(output_path)
    return output_path


# ---------- Helpers preview ----------
def load_font(size: int, bold: bool = False):
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf" if bold else "/System/Library/Fonts/Supplemental/Arial.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            return ImageFont.truetype(path, size=size)
    return ImageFont.load_default()


def inch_to_px_x(v):
    return int(v / 10 * SLIDE_SIZE[0])


def inch_to_px_y(v):
    return int(v / 7.5 * SLIDE_SIZE[1])


def preview_image_for_slide(slide_data):
    img = Image.new("RGB", SLIDE_SIZE, "white")
    draw = ImageDraw.Draw(img)

    slide_type = slide_data["type"]

    if slide_type == "Titre général":
        draw.rectangle([0, 0, SLIDE_SIZE[0], SLIDE_SIZE[1]], fill=(59, 84, 135))
        font = load_font(34, bold=True)
        title = slide_data["title"] or "Titre général"
        bbox = draw.textbbox((0, 0), title, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text(((SLIDE_SIZE[0] - tw) / 2, (SLIDE_SIZE[1] - th) / 2), title, fill="white", font=font)

    elif slide_type == "Titre intermédiaire":
        draw.rectangle([0, 0, SLIDE_SIZE[0], SLIDE_SIZE[1]], fill=(103, 139, 196))
        font = load_font(30, bold=True)
        title = slide_data["title"] or "Titre intermédiaire"
        bbox = draw.textbbox((0, 0), title, font=font)
        tw = bbox[2] - bbox[0]
        th = bbox[3] - bbox[1]
        draw.text(((SLIDE_SIZE[0] - tw) / 2, (SLIDE_SIZE[1] - th) / 2), title, fill="white", font=font)

    else:
        header_h = 90
        draw.rectangle([0, 0, SLIDE_SIZE[0], header_h], fill=(103, 139, 196))
        draw.rectangle([40, header_h + 20, SLIDE_SIZE[0] - 40, SLIDE_SIZE[1] - 30], outline=(210, 210, 210), width=2)

        font = load_font(24, bold=True)
        title = slide_data["title"] or "Titre"
        bbox = draw.textbbox((0, 0), title, font=font)
        tw = bbox[2] - bbox[0]
        draw.text(((SLIDE_SIZE[0] - tw) / 2, 28), title, fill="white", font=font)

        for item in slide_data.get("items", []):
            x = inch_to_px_x(item["x"])
            y = inch_to_px_y(item["y"])
            w = inch_to_px_x(item["w"])
            h = inch_to_px_y(item["h"])

            if item["kind"] == "text":
                draw.rectangle([x, y, x + w, y + h], outline=(90, 90, 90), width=2)
                tfont = load_font(18, bold=item.get("bold", False))
                txt = item["text"][:180] if item["text"] else "Texte"
                draw.multiline_text((x + 10, y + 10), txt, fill=(40, 40, 40), font=tfont, spacing=4)
            elif item["kind"] == "image":
                draw.rectangle([x, y, x + w, y + h], outline=(90, 90, 90), width=2)
                draw.line([x, y, x + w, y + h], fill=(120, 120, 120), width=2)
                draw.line([x + w, y, x, y + h], fill=(120, 120, 120), width=2)
                label_font = load_font(18, bold=True)
                label = "Image"
                bbox = draw.textbbox((0, 0), label, font=label_font)
                lw = bbox[2] - bbox[0]
                lh = bbox[3] - bbox[1]
                draw.text((x + (w - lw) / 2, y + (h - lh) / 2), label, fill=(80, 80, 80), font=label_font)

    return img


# ---------- Streamlit UI ----------
st.set_page_config(page_title="Solimed PPT Builder", layout="wide")
st.title("Solimed — Générateur de présentation")
st.caption("Utilise les 3 slides modèles du fichier templates.pptx avec les repères {{TITLE_GENERAL}}, {{TITLE_SECTION}} et {{CONTENT_TITLE}}.")

if "slides" not in st.session_state:
    st.session_state.slides = []

if "draft_items" not in st.session_state:
    st.session_state.draft_items = []


def clamp_content_item(item):
    item["x"] = max(CONTENT_BOUNDS["x_min"], min(item["x"], CONTENT_BOUNDS["x_max"] - 0.3))
    item["y"] = max(CONTENT_BOUNDS["y_min"], min(item["y"], CONTENT_BOUNDS["y_max"] - 0.3))
    item["w"] = max(0.5, min(item["w"], CONTENT_BOUNDS["x_max"] - item["x"]))
    item["h"] = max(0.4, min(item["h"], CONTENT_BOUNDS["y_max"] - item["y"]))
    return item


with st.sidebar:
    st.subheader("Template")
    st.write(f"Fichier attendu : `{TEMPLATE_PATH}`")
    if os.path.exists(TEMPLATE_PATH):
        st.success("Template détecté")
    else:
        st.error("Template introuvable")

    st.subheader("Présentation")
    if st.button("Supprimer la dernière slide", width="stretch"):
        if st.session_state.slides:
            st.session_state.slides.pop()
    if st.button("Vider toute la présentation", width="stretch"):
        st.session_state.slides = []

left, right = st.columns([1.05, 0.95])

with left:
    st.subheader("Créer une slide")
    slide_type = st.selectbox("Type de slide", list(PLACEHOLDERS.keys()))
    title = st.text_input("Titre", placeholder="Saisis le titre de la slide")

    working_slide = {"type": slide_type, "title": title}

    if slide_type == "Contenu":
        st.markdown("#### Ajouter un élément dans la zone blanche")
        item_kind = st.selectbox("Type d’élément", ["text", "image"], format_func=lambda x: "Texte" if x == "text" else "Image")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            x = st.number_input("x", min_value=0.0, max_value=10.0, value=0.8, step=0.1)
        with c2:
            y = st.number_input("y", min_value=0.0, max_value=7.5, value=1.5, step=0.1)
        with c3:
            w = st.number_input("largeur", min_value=0.5, max_value=10.0, value=3.5, step=0.1)
        with c4:
            h = st.number_input("hauteur", min_value=0.4, max_value=7.0, value=1.5, step=0.1)

        if item_kind == "text":
            text = st.text_area("Texte", height=140)
            c5, c6, c7 = st.columns(3)
            with c5:
                font_size_value = st.number_input("Taille", min_value=8, max_value=36, value=14, step=1)
            with c6:
                bold = st.checkbox("Bold")
            with c7:
                font_name = st.text_input("Police", value="Aptos")

            if st.button("Ajouter cet élément", width="stretch"):
                from pptx.util import Pt
                item = {
                    "kind": "text",
                    "x": x,
                    "y": y,
                    "w": w,
                    "h": h,
                    "text": text,
                    "font_size_pt": Pt(font_size_value),
                    "bold": bold,
                    "font_name": font_name,
                }
                st.session_state.draft_items.append(clamp_content_item(item))

        else:
            uploaded = st.file_uploader("Image", type=["png", "jpg", "jpeg"])
            if st.button("Ajouter cet élément", width="stretch"):
                if uploaded is None:
                    st.warning("Ajoute une image avant de valider.")
                else:
                    ext = Path(uploaded.name).suffix.lower() or ".png"
                    tmp_path = f"tmp_{uuid.uuid4().hex}{ext}"
                    with open(tmp_path, "wb") as f:
                        f.write(uploaded.getbuffer())
                    item = {
                        "kind": "image",
                        "x": x,
                        "y": y,
                        "w": w,
                        "h": h,
                        "path": tmp_path,
                    }
                    st.session_state.draft_items.append(clamp_content_item(item))

        working_slide["items"] = st.session_state.draft_items

        if st.session_state.draft_items:
            st.markdown("#### Éléments déjà ajoutés")
            for idx, item in enumerate(st.session_state.draft_items, start=1):
                label = f"{idx}. {'Texte' if item['kind'] == 'text' else 'Image'} — x={item['x']}, y={item['y']}, w={item['w']}, h={item['h']}"
                st.write(label)
            if st.button("Effacer les éléments de cette slide", width="stretch"):
                st.session_state.draft_items = []
                st.rerun()

    if st.button("Ajouter la slide à la présentation", width="stretch"):
        if not title.strip():
            st.error("Le titre est obligatoire.")
        else:
            slide_to_store = copy.deepcopy(working_slide)
            if slide_type != "Contenu":
                slide_to_store["items"] = []
            st.session_state.slides.append(slide_to_store)
            st.session_state.draft_items = []
            st.success("Slide ajoutée.")

    st.markdown("---")
    st.subheader("Slides enregistrées")
    if not st.session_state.slides:
        st.info("Aucune slide pour le moment.")
    else:
        for idx, slide in enumerate(st.session_state.slides, start=1):
            with st.expander(f"Slide {idx} — {slide['type']} — {slide['title']}"):
                st.image(preview_image_for_slide(slide), width="stretch")

with right:
    st.subheader("Preview de la slide en cours")
    st.image(preview_image_for_slide(working_slide), width="stretch")

    st.markdown("---")
    st.subheader("Export")
    if st.button("Générer le PowerPoint", width="stretch"):
        if not os.path.exists(TEMPLATE_PATH):
            st.error("Le fichier templates.pptx doit être à la racine du repo.")
        elif len(st.session_state.slides) == 0:
            st.error("Ajoute au moins une slide avant de générer le fichier.")
        else:
            output_path = f"presentation_{uuid.uuid4().hex[:8]}.pptx"
            try:
                build_pptx(st.session_state.slides, output_path)
                with open(output_path, "rb") as f:
                    st.download_button(
                        "Télécharger le PPT",
                        data=f,
                        file_name="presentation_solimed.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        width="stretch",
                    )
                st.success("PowerPoint généré.")
            except Exception as e:
                st.error(f"Erreur pendant la génération : {e}")
