from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.oxml.ns import qn
from lxml import etree

# ── Colors ────────────────────────────────────────────────────────────────────
BLUE_DARK  = RGBColor(0x1A, 0x52, 0x76)
BLUE_MED   = RGBColor(0x2E, 0x86, 0xC1)
BLUE_LIGHT = RGBColor(0xD6, 0xEA, 0xF8)
GRAY_BG    = RGBColor(0xF2, 0xF2, 0xF2)
GRAY_TEXT  = RGBColor(0x88, 0x88, 0x88)
GREEN      = RGBColor(0x1E, 0x8A, 0x44)
RED        = RGBColor(0xC0, 0x39, 0x2B)
WHITE      = RGBColor(0xFF, 0xFF, 0xFF)
DARK       = RGBColor(0x1C, 0x1C, 0x1C)

FONT = "Aptos Display"

# ── Dimensions (cm) ───────────────────────────────────────────────────────────
SLIDE_W     = 33.87
SLIDE_H     = 19.05
MARGIN      = 2.0
CONTENT_W   = SLIDE_W - 2 * MARGIN        # 29.87
TITLE_LEFT  = MARGIN
TITLE_TOP   = 0.7
TITLE_H     = 2.0
LISERET_TOP = TITLE_TOP + TITLE_H + 0.8   # 3.5
LISERET_H   = 0.08
CONTENT_TOP = LISERET_TOP + LISERET_H + 1.0  # 4.58
CONTENT_H   = SLIDE_H - CONTENT_TOP - 1.0   # 13.47
LEFT_COL_W  = CONTENT_W * 3 / 5             # 17.922
GAP         = 0.4
RIGHT_COL_L = MARGIN + LEFT_COL_W + GAP     # 20.322
BOX_W       = CONTENT_W * 2 / 5 - 1.8      # 10.148


# ── Low-level helpers ─────────────────────────────────────────────────────────

def _blank_slide(prs):
    return prs.slides.add_slide(prs.slide_layouts[6])


def _no_border(shape):
    sp = shape._element
    spPr = sp.find(qn("p:spPr"))
    if spPr is None:
        return
    ln = spPr.find(qn("a:ln"))
    if ln is None:
        ln = etree.SubElement(spPr, qn("a:ln"))
    for child in list(ln):
        ln.remove(child)
    etree.SubElement(ln, qn("a:noFill"))


def _rect(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(1, Cm(left), Cm(top), Cm(width), Cm(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    _no_border(shape)
    return shape


def _oval(slide, left, top, width, height, fill_color):
    shape = slide.shapes.add_shape(9, Cm(left), Cm(top), Cm(width), Cm(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    _no_border(shape)
    return shape


def _set_run(run, text, size, bold, color):
    run.text = text
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color


def _textbox(slide, text, left, top, width, height,
             size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    txb = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    _set_run(p.add_run(), text, size, bold, color)
    return txb


def _multiline_textbox(slide, lines, left, top, width, height, size=14):
    txb = slide.shapes.add_textbox(Cm(left), Cm(top), Cm(width), Cm(height))
    tf = txb.text_frame
    tf.word_wrap = True
    for i, (text, bold, color) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        _set_run(p.add_run(), text, size, bold, color if color else DARK)
    return txb


def _add_oval_number(slide, number, left, top, size=1.2):
    circ = _oval(slide, left, top, size, size, BLUE_DARK)
    tf = circ.text_frame
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf.margin_top = Cm(0)
    tf.margin_bottom = Cm(0)
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    _set_run(p.add_run(), str(number), 16, True, WHITE)
    return circ


# ── Reusable components ───────────────────────────────────────────────────────

def add_title(slide, text):
    txb = slide.shapes.add_textbox(
        Cm(TITLE_LEFT), Cm(TITLE_TOP), Cm(CONTENT_W), Cm(TITLE_H)
    )
    tf = txb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    _set_run(p.add_run(), text, 36, True, DARK)
    _rect(slide, TITLE_LEFT, LISERET_TOP, CONTENT_W, LISERET_H, BLUE_DARK)


def add_page_number(slide, number, is_cover=False):
    color = WHITE if is_cover else DARK
    _textbox(slide, str(number),
             SLIDE_W - 2.2, SLIDE_H - 1.1, 1.8, 0.8,
             size=14, color=color, align=PP_ALIGN.RIGHT)


def add_kpi_box(slide, value, label, left, top, width, height, value_color=BLUE_MED):
    _rect(slide, left, top, width, height, GRAY_BG)
    pad = 0.4
    val_top = top + height * 0.12
    lbl_top = top + height * 0.58
    _textbox(slide, value,
             left + pad, val_top, width - 2*pad, height * 0.45,
             size=40, bold=True, color=value_color, align=PP_ALIGN.CENTER)
    _textbox(slide, label,
             left + pad, lbl_top, width - 2*pad, height * 0.35,
             size=12, color=GRAY_TEXT, align=PP_ALIGN.CENTER)


def add_card(slide, title, body_lines, left, top, width, height):
    _rect(slide, left, top, width, height, GRAY_BG)
    BAND_H = 1.8
    _rect(slide, left, top, width, BAND_H, BLUE_DARK)
    pad = 0.35
    _textbox(slide, title,
             left + pad, top + 0.35, width - 2*pad, BAND_H - 0.3,
             size=18, bold=True, color=WHITE, align=PP_ALIGN.LEFT)
    body_top = top + BAND_H + 0.3
    body_h = height - BAND_H - 0.5
    _multiline_textbox(slide, body_lines,
                       left + pad, body_top, width - 2*pad, body_h,
                       size=13)


def add_pipeline_step(slide, number, title, body, left, top, width, height):
    _rect(slide, left, top, width, height, BLUE_LIGHT)
    circ_size = 1.0
    circ_top = top + (height - circ_size) / 2
    _add_oval_number(slide, number, left + 0.35, circ_top, circ_size)
    text_left = left + 0.35 + circ_size + 0.35
    text_w = width - 0.35 - circ_size - 0.35 - 0.35
    _textbox(slide, title,
             text_left, top + 0.2, text_w, height * 0.42,
             size=14, bold=True, color=DARK)
    _textbox(slide, body,
             text_left, top + height * 0.44, text_w, height * 0.5,
             size=12, color=GRAY_TEXT)


# ── Slide builders ────────────────────────────────────────────────────────────

def build_cover(prs):
    slide = _blank_slide(prs)
    _rect(slide, 0, 0, SLIDE_W, SLIDE_H, BLUE_DARK)
    _textbox(slide,
             "Accès à l'eau potable dans le monde",
             3.0, 5.5, SLIDE_W - 6.0, 4.0,
             size=50, bold=True, color=WHITE, align=PP_ALIGN.CENTER)
    _textbox(slide,
             "Projet DWFA  ·  Analyse exploratoire et tableau de bord interactif",
             3.0, 10.2, SLIDE_W - 6.0, 2.5,
             size=36, color=WHITE, align=PP_ALIGN.CENTER)
    _rect(slide, MARGIN, 14.5, CONTENT_W, 0.06, RGBColor(0x5D, 0xAD, 0xE8))
    add_page_number(slide, 1, is_cover=True)


def build_contexte(prs):
    slide = _blank_slide(prs)
    add_title(slide,
              "387 millions de personnes sans accès à l'eau basique"
              " — l'Afrique concentre l'urgence humaine")

    body_lines = [
        ("DWFA (Drinking Water For All) est une organisation internationale dédiée à l'amélioration "
         "de l'accès à l'eau potable dans les pays en développement.",
         False, None),
        ("", False, None),
        ("Contexte mondial en 2017 :", True, BLUE_DARK),
        ("• 68 % de la population mondiale a accès à l'eau basique", False, None),
        ("• Seuls 24 % bénéficient d'un accès sécurisé (qualité OMS garantie)", False, None),
        ("• L'écart entre basique et sécurisé révèle un déficit de qualité massif", False, None),
        ("• L'Afrique et l'Asie du Sud-Est concentrent la majorité des personnes non couvertes", False, None),
        ("", False, None),
        ("Périmètre de l'étude :", True, BLUE_DARK),
        ("• 194 pays  ·  6 régions OMS  ·  2000 à 2017", False, None),
        ("", False, None),
        ("Objectif DWFA :", True, BLUE_DARK),
        ("Identifier les pays prioritaires pour une intervention adaptée :", False, None),
        ("D1 Construction   ·   D2 Modernisation   ·   D3 Conseil gouvernemental", False, None),
    ]
    _multiline_textbox(slide, body_lines, MARGIN, CONTENT_TOP,
                       LEFT_COL_W, CONTENT_H, size=14)

    kpi_h = (CONTENT_H - 2 * 0.35) / 3
    kpis = [
        ("387 M",  "Population sans accès à l'eau basique (2017)", RED),
        ("68 %",   "Taux d'accès basique mondial (2017)",          BLUE_MED),
        ("464 K",  "Décès liés au WASH (2016)",                    RED),
    ]
    for i, (val, lbl, col) in enumerate(kpis):
        add_kpi_box(slide, val, lbl,
                    RIGHT_COL_L, CONTENT_TOP + i * (kpi_h + 0.35),
                    BOX_W, kpi_h, col)

    add_page_number(slide, 2)


def build_sources(prs):
    slide = _blank_slide(prs)
    add_title(slide,
              "5 sources OMS / Banque Mondiale, 194 pays sur 18 ans — périmètre mondial validé")

    sources = [
        ("Accès à l'eau",
         "BasicAndSafelyManaged\nDrinkingWaterServices.csv",
         ["% accès basique & sécurisé",
          "Granularités : Total / Urban / Rural",
          "Période : 2000–2017"]),
        ("Mortalité WASH",
         "MortalityRateAttributed\nToWater.csv",
         ["Taux de mortalité WASH",
          "Année disponible : 2016",
          "183 pays couverts"]),
        ("Population",
         "Population.csv",
         ["Pop. totale, urbaine, rurale",
          "Période : 2000–2017",
          "194 pays — source ONU"]),
        ("Stabilité politique",
         "PoliticalStability.csv",
         ["Score WGI stabilité politique",
          "Échelle : −2,5 à +2,5",
          "Période : 2000–2017"]),
        ("Régions OMS",
         "RegionCountry.csv",
         ["Correspondance pays → région",
          "6 régions OMS",
          "194 pays — couverture 100 %"]),
    ]

    n = len(sources)
    card_gap = 0.3
    card_w = (CONTENT_W - (n - 1) * card_gap) / n
    card_h = CONTENT_H - 2.1

    for i, (title, src_file, body_parts) in enumerate(sources):
        left = MARGIN + i * (card_w + card_gap)
        body_lines = (
            [(line, True, BLUE_DARK) for line in src_file.split("\n")]
            + [("", False, None)]
            + [(bp, False, None) for bp in body_parts]
        )
        add_card(slide, title, body_lines, left, CONTENT_TOP, card_w, card_h)

    kpi_w = (CONTENT_W - 2 * 0.35) / 3
    bottom_top = CONTENT_TOP + card_h + 0.35
    bottom_kpis = [
        ("194 pays",  "couverture mondiale"),
        ("18 ans",    "2000 – 2017"),
        ("6 régions", "OMS"),
    ]
    for i, (val, lbl) in enumerate(bottom_kpis):
        add_kpi_box(slide, val, lbl,
                    MARGIN + i * (kpi_w + 0.35), bottom_top,
                    kpi_w, 1.6, BLUE_DARK)

    add_page_number(slide, 3)


def build_pretraitement(prs):
    slide = _blank_slide(prs)
    add_title(slide,
              "4 étapes de consolidation produisent 2 tables analytiques"
              " — zéro doublon, validation automatique")

    steps = [
        ("Renommage & types",
         "Colonnes standardisées en français · types numériques validés"
         " · pourcentages plafonnés à 100 (erreurs d'arrondi OMS)"),
        ("Filtrage granularité",
         "Extraction granularité «Total» pour les jointures principales"
         " · «Urban / Rural» conservés séparément pour les charts"),
        ("Indicateurs dérivés",
         "Population en personnes (× 1 000) · pct_pop_rurale"
         " · nb_deces_wash = taux_mortalite × pop / 100 000"),
        ("Jointure & export",
         "5 sources consolidées en 1 table · tests assert automatiques"
         " · 2 CSV exportés : dwfa_consolide.csv + dwfa_eau_granulaire.csv"),
    ]

    step_gap = 0.35
    step_h = (CONTENT_H - (len(steps) - 1) * step_gap - 1.3) / len(steps)

    for i, (title, body) in enumerate(steps):
        add_pipeline_step(slide, i + 1, title, body,
                          MARGIN, CONTENT_TOP + i * (step_h + step_gap),
                          LEFT_COL_W, step_h)

    kpi_h = (CONTENT_H - 2 * 0.35 - 1.3) / 3
    kpis = [
        ("3 492", "Lignes — dwfa_consolide.csv",    BLUE_MED),
        ("13",    "Colonnes analytiques",             BLUE_MED),
        ("0",     "Doublon — validation réussie",     GREEN),
    ]
    for i, (val, lbl, col) in enumerate(kpis):
        add_kpi_box(slide, val, lbl,
                    RIGHT_COL_L, CONTENT_TOP + i * (kpi_h + 0.35),
                    BOX_W, kpi_h, col)

    note = (
        "Valeurs manquantes documentées : 50 % pour l'accès sécurisé (non collecté par l'OMS) "
        "· 94,8 % pour la mortalité (une seule année disponible : 2016)"
    )
    _textbox(slide, note,
             MARGIN, SLIDE_H - 1.7, CONTENT_W, 0.9,
             size=10, color=GRAY_TEXT)

    add_page_number(slide, 4)


def build_justification(prs):
    slide = _blank_slide(prs)
    add_title(slide,
              "Tableau Public choisi pour son interactivité et sa publication sans infrastructure"
              " — adapté aux besoins DWFA")

    criteres = [
        ("Interactivité",
         [("Filtres dynamiques par région et par pays", False, None),
          ("Actions de navigation entre les 3 vues", False, None),
          ("Paramètres ajustables en temps réel", False, None),
          ("Highlight et sélection inter-graphiques", False, None)]),
        ("Publication",
         [("Gratuit — hébergé sur tableau.com", False, None),
          ("Lien partageable sans serveur", False, None),
          ("Aucune infrastructure requise", False, None),
          ("Mise à jour des données simplifiée", False, None)]),
        ("Multi-vues connectées",
         [("3 dashboards interconnectés", False, None),
          ("Navigation clic → zoom sur un pays", False, None),
          ("Contexte régional conservé entre vues", False, None),
          ("Paramètre pays partagé entre les vues", False, None)]),
        ("Accessibilité",
         [("Aucune installation côté lecteur", False, None),
          ("Consultation via navigateur standard", False, None),
          ("Compatible mobile et desktop", False, None),
          ("Exportable en PDF ou image", False, None)]),
    ]

    n = len(criteres)
    card_gap = 0.3
    card_w = (CONTENT_W - (n - 1) * card_gap) / n
    card_h = CONTENT_H

    for i, (title, body_lines) in enumerate(criteres):
        left = MARGIN + i * (card_w + card_gap)
        add_card(slide, title, body_lines, left, CONTENT_TOP, card_w, card_h)

    add_page_number(slide, 5)


def build_architecture(prs):
    slide = _blank_slide(prs)
    add_title(slide,
              "3 vues imbriquées pour répondre à une seule question"
              " : quel pays prioriser pour l'intervention DWFA ?")

    vues = [
        ("Vue MONDIALE",
         [("Question : Quelle est l'ampleur du problème ?", True, BLUE_MED),
          ("", False, None),
          ("KPIs : pop. sans accès · accès basique · sécurisé · décès · stabilité", False, None),
          ("", False, None),
          ("Graphiques :", True, DARK),
          ("• Map mondiale taux de mortalité WASH", False, None),
          ("• Bar instabilité politique par région", False, None),
          ("• Line chart évolution accès 2000–2017", False, None),
          ("• Bar décès WASH par région", False, None)]),
        ("Vue CONTINENTALE",
         [("Question : Quels pays prioriser dans la région ?", True, BLUE_MED),
          ("", False, None),
          ("KPIs filtrés par région sélectionnée", False, None),
          ("", False, None),
          ("Graphiques :", True, DARK),
          ("• Scatter mortalité vs accès (urgence humaine)", False, None),
          ("• Scatter D3 : stabilité vs accès sécurisé", False, None),
          ("• Bar classement pays par accès basique", False, None),
          ("• Line chart évolution par pays dans la région", False, None)]),
        ("Vue PAYS",
         [("Question : Comment intervenir ?", True, BLUE_MED),
          ("", False, None),
          ("KPIs filtrés par pays sélectionné", False, None),
          ("", False, None),
          ("Graphiques :", True, DARK),
          ("• Scatter D1 : urbanisation vs accès basique", False, None),
          ("• Scatter D2 : accès basique vs accès sécurisé", False, None),
          ("• Area chart évolution accès 2000–2017", False, None),
          ("• Panel diagnostic D1 / D2 / D3", False, None)]),
    ]

    n = len(vues)
    card_gap = 0.35
    card_w = (CONTENT_W - (n - 1) * card_gap) / n
    card_h = CONTENT_H

    for i, (title, body_lines) in enumerate(vues):
        left = MARGIN + i * (card_w + card_gap)
        add_card(slide, title, body_lines, left, CONTENT_TOP, card_w, card_h)

    add_page_number(slide, 6)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    prs = Presentation()
    prs.slide_width  = Cm(SLIDE_W)
    prs.slide_height = Cm(SLIDE_H)

    build_cover(prs)
    build_contexte(prs)
    build_sources(prs)
    build_pretraitement(prs)
    build_justification(prs)
    build_architecture(prs)

    output = "/home/elicesjo/Formation/PROJET-10-/presentation_DWFA.pptx"
    prs.save(output)
    print(f"presentation_DWFA.pptx créé — {len(prs.slides)} slides")


if __name__ == "__main__":
    main()
