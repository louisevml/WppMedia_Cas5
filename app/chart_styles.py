"""
Fonctions de mise en forme des graphiques PowerPoint.
Applique couleurs, légendes, axes, data-labels selon le type de graphique.
"""

from __future__ import annotations

import math

from pptx.chart.chart import Chart
from pptx.chart.plot import BarPlot
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Pt

from .slide_data import SlideData


# ── Layout helpers : titre, plotArea, légende ─────────────────────────────────

_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"


def _inject_manual_layout(parent_el, *, x: float, y: float,
                          w: float | None = None, h: float | None = None,
                          target_inner: bool = False) -> None:
    """Ajoute (ou remplace) un <c:layout><c:manualLayout> sur *parent_el*."""
    old = parent_el.find(f"{{{_NS_C}}}layout")
    if old is not None:
        parent_el.remove(old)

    parts = [
        f'<c:layout {nsdecls("c")}>',
        "<c:manualLayout>",
        *([ '<c:layoutTarget val="inner"/>'] if target_inner else []),
        '<c:xMode val="edge"/>',
        '<c:yMode val="edge"/>',
        f'<c:x val="{x}"/>',
        f'<c:y val="{y}"/>',
    ]
    if w is not None:
        parts.append(f'<c:w val="{w}"/>')
    if h is not None:
        parts.append(f'<c:h val="{h}"/>')
    parts += ["</c:manualLayout>", "</c:layout>"]

    el = parse_xml("".join(parts))
    parent_el.insert(0, el)


def _compute_layout(
    *,
    has_title: bool,
    is_horizontal: bool,
    n_categories: int,
    max_label_len: int,
    legend_n_lines: int = 1,
    show_legend: bool = True,
) -> dict[str, float]:
    """
    Répartition verticale : **zone de tracé maximale**, titre lisible mais
    compact, légende **en bas** (si ``show_legend``) à distance normale du graphique.

    Toutes les valeurs sont des fractions (0-1) de la hauteur du cadre chart.

    Retourne : plot_x, plot_y, plot_w, plot_h, legend_x, legend_y, legend_w, legend_h.
    """
    # ── Titre : bande minimale suffisante pour Pt(9), 1–2 lignes courtes ────
    title_h = 0.048 if has_title else 0.012

    # ── Sous le tracé : labels d'axe (vertical) puis léger interstice + légende
    if is_horizontal:
        axis_labels_h = 0.018
    else:
        crowding = max(n_categories / 10.0, max_label_len / 12.0)
        if crowding <= 1.0:
            axis_labels_h = 0.068
        elif crowding <= 1.5:
            axis_labels_h = 0.105
        elif crowding <= 2.2:
            axis_labels_h = 0.142
        else:
            axis_labels_h = 0.175

    if show_legend:
        legend_gap = 0.010
        if legend_n_lines <= 1:
            legend_h = 0.048
        else:
            legend_h = 0.082
    else:
        legend_gap = 0.0
        legend_h = 0.0

    bottom_pad = 0.006

    # plot = tout ce qui reste → barres le plus grandes possible
    plot_y = title_h
    plot_h = 1.0 - title_h - axis_labels_h - legend_gap - legend_h - bottom_pad
    if plot_h < 0.38:
        plot_h = 0.38

    legend_y = title_h + plot_h + axis_labels_h + legend_gap

    return {
        "plot_x":   0.0,
        "plot_y":   plot_y,
        "plot_w":   1.0,
        "plot_h":   plot_h,
        "legend_x": 0.05,
        "legend_y": legend_y,
        "legend_w": 0.90,
        "legend_h": legend_h,
    }


def _configure_legend(chart: Chart, data: SlideData, *, font_pt: int) -> None:
    """
    Une seule série → pas de légende (inutile, ex. barres horizontales « Ensemble » seul).
    Sinon légende en bas, police homogène.
    """
    if len(data.series_names) <= 1:
        chart.has_legend = False
        return
    chart.has_legend = True
    leg = chart.legend
    if leg is None:
        return
    leg.include_in_layout = False
    leg.position = XL_LEGEND_POSITION.BOTTOM
    leg.font.size = Pt(font_pt)
    leg.font.color.rgb = _LEGEND_COLOR


def _set_chart_layout(
    chart: Chart,
    data: SlideData,
    *,
    chart_type: str,
) -> None:
    """
    Positionne le plotArea et la légende en fonction du contenu du graphique
    pour garantir une mise en page harmonieuse (pas de chevauchement, pas de
    grand vide).
    """
    ct = chart_type.lower()
    is_pie_like = ct in ("pie", "doughnut")
    is_horizontal = "horizontal" in ct or is_pie_like
    n_cat = len(data.categories)
    max_label_len = max((len(c) for c in data.categories), default=0)
    if is_pie_like:
        legend_items = n_cat
    else:
        legend_items = len(data.series_names)
    legend_n_lines = 1 if legend_items <= 4 else 2
    show_legend = bool(chart.has_legend) and legend_items > 0

    layout = _compute_layout(
        has_title=chart.has_title,
        is_horizontal=is_horizontal,
        n_categories=n_cat,
        max_label_len=max_label_len,
        legend_n_lines=legend_n_lines,
        show_legend=show_legend,
    )

    chart_el = chart._chartSpace.find(f"{{{_NS_C}}}chart")

    plot_area = chart_el.find(f"{{{_NS_C}}}plotArea") if chart_el is not None else None
    if plot_area is not None:
        # Pie / donut : éviter layoutTarget="inner" sur le plotArea — PowerPoint
        # peut marquer le fichier comme corrompu puis supprimer les graphiques
        # à la réparation (slide multi-camemberts surtout).
        _inject_manual_layout(
            plot_area,
            x=layout["plot_x"], y=layout["plot_y"],
            w=layout["plot_w"], h=layout["plot_h"],
            target_inner=not is_pie_like,
        )

    legend_el = chart_el.find(f"{{{_NS_C}}}legend") if chart_el is not None else None
    if legend_el is not None:
        if show_legend:
            _inject_manual_layout(
                legend_el,
                x=layout["legend_x"], y=layout["legend_y"],
                w=layout["legend_w"], h=layout["legend_h"],
            )
        else:
            old = legend_el.find(f"{{{_NS_C}}}layout")
            if old is not None:
                legend_el.remove(old)


# ── Palettes historiques (référence visuelle ; les couleurs effectives sont
#    désormais générées par série via generate_distinct_colors) ─────────────

DEFAULT_SERIES_COLORS: tuple[RGBColor, ...] = (
    RGBColor(0, 176, 80),
    RGBColor(192,   0,   0),
    RGBColor(84, 101, 255),
    RGBColor(147,  75, 201),
)

_LABEL_COLOR      = RGBColor(64, 64, 120)
_LEGEND_COLOR     = RGBColor(64, 64, 120)
_SEPARATOR_COLOR  = RGBColor(217, 217, 217)

# Partie fractionnaire du nombre d'or : teintes maximalement espacées sur le cercle HSV
_PHI = 0.618033988749895


def _hsv_to_rgb(h: float, s: float, v: float) -> RGBColor:
    """h, s, v dans [0, 1] → RGBColor."""
    if s <= 0:
        x = int(round(v * 255))
        return RGBColor(x, x, x)
    h = (h % 1.0) * 6.0
    i = int(h)
    f = h - i
    p = v * (1.0 - s)
    q = v * (1.0 - f * s)
    t = v * (1.0 - (1.0 - f) * s)
    if i == 0:
        r, g, b = v, t, p
    elif i == 1:
        r, g, b = q, v, p
    elif i == 2:
        r, g, b = p, v, t
    elif i == 3:
        r, g, b = p, q, v
    elif i == 4:
        r, g, b = t, p, v
    else:
        r, g, b = v, p, q
    return RGBColor(
        int(max(0, min(255, round(r * 255)))),
        int(max(0, min(255, round(g * 255)))),
        int(max(0, min(255, round(b * 255)))),
    )


def _cap_fill_luminance(rgb: RGBColor, max_lum: float = 0.50) -> RGBColor:
    """
    Assombrit un remplissage trop clair (souvent verts / jaunes pâles illisibles
    sur fond blanc) en ramenant la luminance relative sous *max_lum*.
    """
    r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    for _ in range(14):
        if _relative_luminance(RGBColor(r, g, b)) <= max_lum:
            break
        r = max(0, int(r * 0.88))
        g = max(0, int(g * 0.88))
        b = max(0, int(b * 0.88))
    return RGBColor(r, g, b)


def generate_distinct_colors(n: int) -> tuple[RGBColor, ...]:
    """
    Génère *n* couleurs les plus distinctes possible sans dépendance externe :
    teintes réparties par angle d'or sur le cercle chromatique, avec légères
    variations de saturation et de luminance pour limiter les collisions visuelles.

    Les teintes jaune–vert sont volontairement un peu plus sombres (V réduit)
    car elles ressortent très claires en RVB à saturation modérée.
    """
    if n <= 0:
        return ()
    out: list[RGBColor] = []
    for i in range(n):
        hue = (i * _PHI) % 1.0
        # Saturation élevée pour éviter les pastels
        sat = 0.72 + 0.16 * ((i + 1) % 3) / 2.0
        # Valeur plus basse qu'avant : couleurs plus « pleines »
        val = 0.78 - 0.12 * ((i // 3) % 2)
        # Plage de teinte ~ jaune–vert–cyan (0,12–0,42 en 0–1) : trop clair en RVB
        if 0.12 <= hue <= 0.42:
            val *= 0.82
        c = _hsv_to_rgb(hue, sat, val)
        out.append(_cap_fill_luminance(c, max_lum=0.50))
    return tuple(out)


def _relative_luminance(rgb: RGBColor) -> float:
    """Luminance WCAG 2.1 relative (0 = noir, 1 = blanc)."""
    def lin(c: float) -> float:
        c /= 255.0
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r = lin(float(rgb[0]))
    g = lin(float(rgb[1]))
    b = lin(float(rgb[2]))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def _contrast_label_color(fill: RGBColor) -> RGBColor:
    """Texte lisible sur une zone de couleur *fill* (data labels empilés)."""
    lum = _relative_luminance(fill)
    # Texte blanc seulement sur fond assez sombre ; pastels / verts clairs → texte foncé
    return RGBColor(255, 255, 255) if lum < 0.38 else _LABEL_COLOR


# ═════════════════════════════════════════════════════════════════════════════
# Helpers internes
# ═════════════════════════════════════════════════════════════════════════════

def _hide_legend_entries(legend_elm, indices: tuple[int, ...]) -> None:
    for idx in indices:
        entry = parse_xml(
            f"<c:legendEntry {nsdecls('c')}>"
            f'<c:idx val="{idx}"/>'
            f'<c:delete val="1"/>'
            f"</c:legendEntry>"
        )
        legend_elm.append(entry)


def _apply_series_colors(chart: Chart) -> tuple[RGBColor, ...]:
    """
    Applique une couleur de remplissage à chaque série, palette générée
    selon le nombre de séries (couleurs distinctes).
    Retourne la palette utilisée (pour data labels, etc.).
    """
    n = len(chart.series)
    if n == 0:
        return ()
    palette = generate_distinct_colors(n)
    for i, ser in enumerate(chart.series):
        ser.invert_if_negative = False
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = palette[i]
    return palette


def _apply_pie_doughnut_slice_colors(chart: Chart) -> tuple[RGBColor, ...]:
    """
    Pour un camembert / anneau : une seule série dont les **points** représentent
    les catégories. Cette fonction colore chaque point individuellement avec
    une palette de couleurs distinctes.

    Retourne la palette utilisée (index par catégorie).
    """
    if not chart.series:
        return ()
    ser = chart.series[0]

    ser_el = ser._element
    cat_els = ser_el.findall(
        f"{{{_NS_C}}}cat/{{{_NS_C}}}strRef/{{{_NS_C}}}strCache/{{{_NS_C}}}pt"
    )
    if not cat_els:
        cat_els = ser_el.findall(
            f"{{{_NS_C}}}cat/{{{_NS_C}}}numRef/{{{_NS_C}}}numCache/{{{_NS_C}}}pt"
        )
    n = len(cat_els)
    if n == 0:
        val_els = ser_el.findall(
            f"{{{_NS_C}}}val/{{{_NS_C}}}numRef/{{{_NS_C}}}numCache/{{{_NS_C}}}pt"
        )
        n = len(val_els)
    if n == 0:
        return _apply_series_colors(chart)

    palette = generate_distinct_colors(n)

    for idx in range(n):
        dPt = ser_el.get_or_add_dPt_for_point(idx)
        existing = dPt.find(f"{{{_NS_C}}}spPr")
        if existing is not None:
            dPt.remove(existing)
        rgb = palette[idx]
        hex_color = "{:02X}{:02X}{:02X}".format(int(rgb[0]), int(rgb[1]), int(rgb[2]))
        spPr = parse_xml(
            f"<c:spPr {nsdecls('c', 'a')}>"
            f'  <a:solidFill><a:srgbClr val="{hex_color}"/></a:solidFill>'
            f'  <a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln>'
            f"</c:spPr>"
        )
        bubble3D = dPt.find(f"{{{_NS_C}}}bubble3D")
        if bubble3D is not None:
            bubble3D.addnext(spPr)
        else:
            dPt.append(spPr)

    return palette


def _force_legend(chart: Chart, *, font_pt: int) -> None:
    """Active et met en forme la légende (bas, police homogène), indépendamment
    du nombre de séries — utile pour pie / doughnut où les entrées de légende
    correspondent aux catégories."""
    chart.has_legend = True
    leg = chart.legend
    if leg is None:
        return
    leg.include_in_layout = False
    leg.position = XL_LEGEND_POSITION.BOTTOM
    leg.font.size = Pt(font_pt)
    leg.font.color.rgb = _LEGEND_COLOR


def _try_disable_leader_lines(dlbls_elm) -> None:
    el = dlbls_elm.find(qn("c:showLeaderLines"))
    if el is not None:
        el.set("val", "0")


def _disable_leader_lines_on_chart(chart: Chart) -> None:
    """Désactive les traits d'appel sur tous les ``c:dLbls`` du chartSpace."""
    for el in chart._chartSpace.iter():
        if el.tag == qn("c:dLbls"):
            _try_disable_leader_lines(el)


def _style_chart_title(chart: Chart) -> None:
    """Formate le titre du graphique s'il existe : taille, couleur, positionné en haut."""
    if not chart.has_title:
        return
    tf = chart.chart_title.text_frame
    for para in tf.paragraphs:
        para.font.size = Pt(9)
        para.font.bold = False
        para.font.color.rgb = _LABEL_COLOR
        para.alignment = None

    chart.chart_title.include_in_layout = False

    title_el = chart._chartSpace.find(f"{{{_NS_C}}}chart/{{{_NS_C}}}title")
    if title_el is not None:
        _inject_manual_layout(title_el, x=0.05, y=0.0)


# ═════════════════════════════════════════════════════════════════════════════
# Style : histogramme groupé (vertical ou horizontal)
# ═════════════════════════════════════════════════════════════════════════════

def _nice_axis_max(data: SlideData) -> float:
    """
    Calcule un maximum d'axe arrondi au palier supérieur « propre »
    juste au-dessus de la valeur max réelle.  Laisse ~15 % de marge
    pour les data-labels au-dessus des barres.

    Ex. max réel 26 → renvoie 35,  max réel 68 → renvoie 80.
    """
    flat = [v for row in data.values_matrix for v in row]
    if not flat:
        return 100.0
    peak = max(flat)
    target = peak * 1.15

    if target <= 10:
        step = 5
    elif target <= 50:
        step = 5
    else:
        step = 10
    return max(step, math.ceil(target / step) * step)


def _compute_gap_width(n_categories: int, n_series: int) -> int:
    """
    Calcule un gap_width adapté au contenu.
    L'original utilise 182 pour ≤14 catégories.
    On réduit progressivement au-delà pour garder les barres lisibles.
    """
    if n_categories <= 14:
        return 182
    if n_categories <= 20:
        return 150
    return 120


def style_bar_clustered(chart: Chart, _data: SlideData, *, chart_type: str = "bar_clustered_vertical") -> None:
    """Style pour bar_clustered_vertical ET bar_clustered_horizontal."""
    n_cat = len(_data.categories)
    n_ser = len(_data.series_names)

    plot = chart.plots[0]
    if isinstance(plot, BarPlot):
        plot.gap_width = _compute_gap_width(n_cat, n_ser)
        plot.overlap = 0
    plot.has_data_labels = False

    palette = _apply_series_colors(chart)

    for i, ser in enumerate(chart.series):
        dls = ser.data_labels
        dls.show_value = True
        dls.show_category_name = False
        dls.show_series_name = False
        dls.show_legend_key = False
        dls.number_format = '0"%"'
        dls.position = XL_LABEL_POSITION.OUTSIDE_END
        dls.font.size = Pt(7)
        if i < len(palette):
            dls.font.color.rgb = palette[i]
        _try_disable_leader_lines(dls._element)

    _configure_legend(chart, _data, font_pt=7)

    try:
        chart.value_axis.visible = False
        chart.value_axis.has_major_gridlines = False
        chart.value_axis.has_minor_gridlines = False
        chart.value_axis.maximum_scale = _nice_axis_max(_data)
        chart.value_axis.minimum_scale = 0.0
    except Exception:
        pass

    try:
        cat = chart.category_axis
        cat.has_major_gridlines = False
        cat.has_minor_gridlines = False
        cat.major_tick_mark = XL_TICK_MARK.NONE
        cat.minor_tick_mark = XL_TICK_MARK.NONE
        cat.format.line.fill.solid()
        cat.format.line.color.rgb = _SEPARATOR_COLOR
        cat.format.line.width = Pt(0.75)
        cat.tick_labels.font.size = Pt(8)
        cat.tick_labels.font.bold = True
        cat.tick_labels.font.color.rgb = _LABEL_COLOR
    except Exception:
        pass

    _set_chart_layout(chart, _data, chart_type=chart_type)


# ═════════════════════════════════════════════════════════════════════════════
# Style : barres / colonnes empilées à 100 %
# ═════════════════════════════════════════════════════════════════════════════

def style_bar_stacked_100(chart: Chart, _data: SlideData, *, chart_type: str = "bar_stacked_100_vertical") -> None:
    """Style pour bar_stacked_100_vertical ET bar_stacked_100_horizontal."""
    n_cat = len(_data.categories)
    n_ser = len(_data.series_names)

    plot = chart.plots[0]
    if isinstance(plot, BarPlot):
        plot.gap_width = _compute_gap_width(n_cat, n_ser)
        plot.overlap = 100
    plot.has_data_labels = False

    palette = _apply_series_colors(chart)

    for i, ser in enumerate(chart.series):
        dls = ser.data_labels
        dls.show_value = True
        dls.show_category_name = False
        dls.show_series_name = False
        dls.show_legend_key = False
        dls.number_format = '0"%"'
        dls.number_format_is_linked = False
        dls.position = XL_LABEL_POSITION.CENTER
        dls.font.size = Pt(7)
        dls.font.bold = True
        dls.font.color.rgb = _contrast_label_color(palette[i]) if i < len(palette) else _LABEL_COLOR
        _try_disable_leader_lines(dls._element)

    _configure_legend(chart, _data, font_pt=7)

    try:
        va = chart.value_axis
        va.visible = False
        va.has_major_gridlines = False
        va.has_minor_gridlines = False
    except Exception:
        pass

    try:
        cat = chart.category_axis
        cat.has_major_gridlines = False
        cat.has_minor_gridlines = False
        cat.major_tick_mark = XL_TICK_MARK.NONE
        cat.minor_tick_mark = XL_TICK_MARK.NONE
        cat.format.line.fill.solid()
        cat.format.line.color.rgb = _SEPARATOR_COLOR
        cat.format.line.width = Pt(0.75)
        cat.tick_labels.font.size = Pt(8)
        cat.tick_labels.font.bold = True
        cat.tick_labels.font.color.rgb = _LABEL_COLOR
    except Exception:
        pass

    _set_chart_layout(chart, _data, chart_type=chart_type)


# ═════════════════════════════════════════════════════════════════════════════
# Style : donut
# ═════════════════════════════════════════════════════════════════════════════

def pie_doughnut_palette(n: int) -> tuple[RGBColor, ...]:
    """Expose la palette appliquée aux parts pour la légende partagée."""
    return generate_distinct_colors(n)


def _style_pie_or_doughnut(
    chart: Chart,
    data: SlideData,
    *,
    chart_type: str,
) -> None:
    """Style commun camembert / anneau :
    - une couleur distincte par part (point = modalité)
    - data labels : **pourcentage natif** (``show_percentage``), format ``0"%"``
      pour des entiers — pas de ``c:tx``/rich par point : ces libellés
      « maison » cassent souvent l'OOXML et déclenchent réparation + graphiques
      supprimés (cf. python-pptx / schéma ``c:dLbl``).
    - légende cachée par défaut : l'orchestrateur décide d'ajouter une légende
      partagée unique en bas de diapo quand plusieurs donuts sont groupés.
    """
    _apply_pie_doughnut_slice_colors(chart)

    try:
        plot = chart.plots[0]
        plot.vary_by_categories = True
        plot.has_data_labels = True
    except Exception:
        pass

    if chart.series and data.values_matrix:
        ser = chart.series[0]
        dls = ser.data_labels
        # Même logique que les barres : les valeurs du tableau sont déjà des
        # pourcentages (0–100), on les affiche telles quelles avec le format
        # ``0"%"``. show_percentage = True renverrait la fraction 0–1 calculée
        # par PowerPoint, ce qui s'affiche « 0% / 1% » avec ce format.
        dls.show_value = True
        dls.show_percentage = False
        dls.show_category_name = False
        dls.show_series_name = False
        dls.show_legend_key = False
        dls.number_format = '0"%"'
        dls.number_format_is_linked = False
        try:
            dls.font.size = Pt(8)
            dls.font.bold = True
            dls.font.color.rgb = _LABEL_COLOR
            dls.position = XL_LABEL_POSITION.CENTER
        except Exception:
            pass
        _try_disable_leader_lines(dls._element)

    chart.has_legend = False

    # python-pptx met showLeaderLines=1 par défaut sur dLbls ; sur pie/donut
    # cela peut contribuer aux alertes « réparer » selon les versions Office.
    _disable_leader_lines_on_chart(chart)

    _set_chart_layout(chart, data, chart_type=chart_type)


def style_doughnut(chart: Chart, _data: SlideData, *, chart_type: str = "doughnut") -> None:
    """Style pour les graphiques en anneau (donut)."""
    _style_pie_or_doughnut(chart, _data, chart_type=chart_type)


# ═════════════════════════════════════════════════════════════════════════════
# Style : camembert (pie)
# ═════════════════════════════════════════════════════════════════════════════

def style_pie(chart: Chart, _data: SlideData, *, chart_type: str = "pie") -> None:
    """Style pour les graphiques en camembert."""
    _style_pie_or_doughnut(chart, _data, chart_type=chart_type)


# ═════════════════════════════════════════════════════════════════════════════
# Dispatcher principal
# ═════════════════════════════════════════════════════════════════════════════

_STYLE_DISPATCH: dict[str, callable] = {
    "bar_clustered_vertical":    style_bar_clustered,
    "bar_clustered_horizontal":  style_bar_clustered,
    "bar_stacked_100_vertical":  style_bar_stacked_100,
    "bar_stacked_100_horizontal": style_bar_stacked_100,
    "doughnut":                  style_doughnut,
    "pie":                       style_pie,
}


def apply_chart_style(chart: Chart, chart_type_str: str, data: SlideData) -> None:
    """
    Point d'entrée : applique le style correspondant au chart_type.
    Sert de callback pour build_presentation().
    """
    key = chart_type_str.strip().lower()
    style_fn = _STYLE_DISPATCH.get(key)
    if style_fn is not None:
        style_fn(chart, data, chart_type=key)
    else:
        print(f"[WARN] Pas de style défini pour chart_type « {chart_type_str} », rendu par défaut.")

    _style_chart_title(chart)
