"""
docx_chapter_extractor.py
=========================

Segmenta um arquivo .docx em capítulos (e seções especiais como prefácio,
anexos, glossário, bibliografia, etc.) de forma genérica, sem depender de
uma formatação específica.

Estratégia (em camadas, da mais confiável para a mais heurística):

  1. Resolver a cascata de estilos (styles.xml) e calcular o outline level
     efetivo de cada parágrafo.
  2. Se houver TOC (sumário) no documento, extrair as entradas como
     "dica" — mas NÃO confiar cegamente, pois pode estar desatualizado.
  3. Varrer todos os parágrafos coletando features (estilo, outline level,
     page-break-before, alinhamento, tamanho de fonte, negrito, caixa
     alta, regex de capítulo, etc.).
  4. Pontuar cada parágrafo como candidato a título de capítulo.
  5. Aplicar threshold adaptativo + validação (espaçamento regular,
     sequência numérica) para escolher os títulos finais.
  6. Classificar cada seção (capítulo, prefácio, anexo, glossário,
     bibliografia, índice, dedicatória, etc.) por padrão textual.
  7. Extrair o texto entre títulos consecutivos.

Uso:
    python docx_chapter_extractor.py caminho/para/livro.docx
    python docx_chapter_extractor.py livro.docx --json saida.json
    python docx_chapter_extractor.py livro.docx --outdir capitulos/

Dependências:
    pip install python-docx lxml
"""

from __future__ import annotations

import argparse
import json
import re
import statistics
import sys
import unicodedata
from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Optional

from docx import Document
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph
from lxml import etree


# ---------------------------------------------------------------------------
# Constantes / namespaces
# ---------------------------------------------------------------------------

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NSMAP = {"w": W_NS}


# Regex para detectar padrões de "capítulo" em PT/EN/ES.
# Cobre: "Capítulo 1", "Cap. I", "CAPÍTULO PRIMEIRO", "Chapter 3", numerais
# romanos isolados em linha curta, e numerais por extenso comuns.
CHAPTER_WORD_RE = re.compile(
    r"""
    ^\s*
    (
        cap[ií]tulo |
        cap\.? |
        chapter |
        cap[ií]tulo\s+ |
        livro |
        parte |
        se[cç][aã]o |
        unidade
    )
    \b
    """,
    re.IGNORECASE | re.VERBOSE,
)

ROMAN_LINE_RE = re.compile(r"^\s*[IVXLCDM]{1,6}\s*[\.\-–—:]?\s*.{0,80}$")
ARABIC_LINE_RE = re.compile(r"^\s*\d{1,3}\s*[\.\-–—:]\s*.{0,80}$")

# Padrões para classificar seções especiais (frontmatter / backmatter).
SECTION_PATTERNS: list[tuple[str, re.Pattern]] = [
    ("prefacio",      re.compile(r"^\s*pref[áa]cio\b", re.I)),
    ("apresentacao",  re.compile(r"^\s*apresenta[çc][ãa]o\b", re.I)),
    ("introducao",    re.compile(r"^\s*introdu[çc][ãa]o\b", re.I)),
    ("dedicatoria",   re.compile(r"^\s*dedicat[óo]ria\b", re.I)),
    ("agradecimentos",re.compile(r"^\s*agradecimentos?\b", re.I)),
    ("sumario",       re.compile(r"^\s*(sum[áa]rio|[íi]ndice)\s*$", re.I)),
    ("epigrafe",      re.compile(r"^\s*ep[íi]grafe\b", re.I)),
    ("prologo",       re.compile(r"^\s*pr[óo]logo\b", re.I)),
    ("epilogo",       re.compile(r"^\s*ep[íi]logo\b", re.I)),
    ("posfacio",      re.compile(r"^\s*pos?f[áa]cio\b", re.I)),
    ("conclusao",     re.compile(r"^\s*conclus[ãa]o\b", re.I)),
    ("anexo",         re.compile(r"^\s*anexo\b", re.I)),
    ("apendice",      re.compile(r"^\s*ap[êe]ndice\b", re.I)),
    ("glossario",     re.compile(r"^\s*gloss[áa]rio\b", re.I)),
    ("bibliografia",  re.compile(r"^\s*(bibliografia|refer[êe]ncias)\b", re.I)),
    ("notas",         re.compile(r"^\s*notas?\s*(de\s+rodap[ée])?\s*$", re.I)),
    ("indice_remissivo", re.compile(r"^\s*[íi]ndice\s+remissivo\b", re.I)),
    ("sobre_autor",   re.compile(r"^\s*sobre\s+o\s+autor\b", re.I)),
]


# ---------------------------------------------------------------------------
# Estruturas de dados
# ---------------------------------------------------------------------------

@dataclass
class ParaFeatures:
    """Features extraídas de um parágrafo para pontuação."""
    index: int                       # índice do parágrafo no documento
    text: str
    style_name: str = ""
    outline_level: Optional[int] = None   # 0..8 (None se não tem)
    page_break_before: bool = False
    alignment: str = ""              # left / center / right / justify
    font_size: Optional[float] = None     # em pontos
    is_bold: bool = False
    is_all_caps: bool = False
    char_count: int = 0
    word_count: int = 0
    matches_chapter_regex: bool = False
    matches_roman: bool = False
    matches_arabic: bool = False
    score: float = 0.0


@dataclass
class Section:
    """Uma seção detectada (capítulo ou seção especial)."""
    kind: str                        # 'capitulo', 'prefacio', 'anexo', ...
    title: str
    start_para: int
    end_para: int                    # exclusivo
    text: str = ""
    confidence: float = 0.0
    detection_score: float = 0.0


# ---------------------------------------------------------------------------
# Resolução de estilos
# ---------------------------------------------------------------------------

class StyleResolver:
    """
    Resolve a cascata de estilos do documento, expondo principalmente o
    outline level efetivo de cada estilo (herdado via basedOn).
    """

    def __init__(self, doc: DocxDocument):
        self.doc = doc
        self._cache: dict[str, Optional[int]] = {}
        # Map styleId -> element
        self._by_id: dict[str, etree._Element] = {}
        try:
            root = doc.styles.element
            for s in root.findall(f"{{{W_NS}}}style"):
                sid = s.get(f"{{{W_NS}}}styleId")
                if sid:
                    self._by_id[sid] = s
        except Exception:
            pass

    def outline_level(self, style_id: str) -> Optional[int]:
        """Retorna o outline level resolvido (com herança) ou None."""
        if not style_id:
            return None
        if style_id in self._cache:
            return self._cache[style_id]

        visited: set[str] = set()
        current = style_id
        result: Optional[int] = None

        while current and current not in visited:
            visited.add(current)
            el = self._by_id.get(current)
            if el is None:
                break

            # Procurar outlineLvl no pPr do estilo
            lvl_el = el.find(f"{{{W_NS}}}pPr/{{{W_NS}}}outlineLvl")
            if lvl_el is not None:
                val = lvl_el.get(f"{{{W_NS}}}val")
                if val is not None:
                    try:
                        result = int(val)
                        break
                    except ValueError:
                        pass

            # Subir para basedOn
            based = el.find(f"{{{W_NS}}}basedOn")
            if based is not None:
                current = based.get(f"{{{W_NS}}}val", "")
            else:
                current = ""

        # Heurística por nome: "Heading 1" / "Título 1" → 0
        if result is None:
            el = self._by_id.get(style_id)
            if el is not None:
                name_el = el.find(f"{{{W_NS}}}name")
                if name_el is not None:
                    name = (name_el.get(f"{{{W_NS}}}val") or "").lower()
                    m = re.search(r"(heading|t[íi]tulo|t[íi]tul)\s*(\d)", name)
                    if m:
                        result = int(m.group(2)) - 1

        self._cache[style_id] = result
        return result


# ---------------------------------------------------------------------------
# Extração de features dos parágrafos
# ---------------------------------------------------------------------------

def _is_all_caps(text: str) -> bool:
    letters = [c for c in text if c.isalpha()]
    if len(letters) < 3:
        return False
    return sum(1 for c in letters if c.isupper()) / len(letters) > 0.85


def _has_page_break_before(p: Paragraph) -> bool:
    el = p._p
    # pageBreakBefore no pPr
    if el.find(f"{{{W_NS}}}pPr/{{{W_NS}}}pageBreakBefore") is not None:
        return True
    # <w:br w:type="page"/> em qualquer run do parágrafo
    for br in el.iter(f"{{{W_NS}}}br"):
        if br.get(f"{{{W_NS}}}type") == "page":
            return True
    return False


def _font_size_pt(p: Paragraph) -> Optional[float]:
    """Pega o maior font size encontrado nos runs (em pontos)."""
    sizes = []
    for run in p.runs:
        sz = run.font.size
        if sz is not None:
            sizes.append(sz.pt)
    if sizes:
        return max(sizes)
    return None


def _is_bold(p: Paragraph) -> bool:
    if not p.runs:
        return False
    bolds = [bool(r.bold) for r in p.runs if r.text.strip()]
    if not bolds:
        return False
    return sum(bolds) / len(bolds) > 0.6


def extract_features(
    doc: DocxDocument, resolver: StyleResolver
) -> list[ParaFeatures]:
    feats: list[ParaFeatures] = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        style_name = p.style.name if p.style else ""
        style_id = p.style.style_id if p.style else ""
        outline = resolver.outline_level(style_id)

        f = ParaFeatures(
            index=i,
            text=text,
            style_name=style_name,
            outline_level=outline,
            page_break_before=_has_page_break_before(p),
            alignment=str(p.alignment).split(".")[-1].lower() if p.alignment else "",
            font_size=_font_size_pt(p),
            is_bold=_is_bold(p),
            is_all_caps=_is_all_caps(text),
            char_count=len(text),
            word_count=len(text.split()),
            matches_chapter_regex=bool(CHAPTER_WORD_RE.search(text)) if text else False,
            matches_roman=bool(ROMAN_LINE_RE.match(text)) if text else False,
            matches_arabic=bool(ARABIC_LINE_RE.match(text)) if text else False,
        )
        feats.append(f)
    return feats


# ---------------------------------------------------------------------------
# Extração do TOC (sumário) — usado como dica, não como verdade absoluta
# ---------------------------------------------------------------------------

def extract_toc_hints(doc: DocxDocument) -> list[str]:
    """
    Tenta extrair os títulos das entradas do TOC. Funciona quando o
    sumário foi gerado pelo Word com 'Inserir Sumário' (TOC field).
    Retorna apenas os textos dos títulos, normalizados.
    """
    hints: list[str] = []
    body = doc.element.body
    in_toc = False
    for p in body.iter(f"{{{W_NS}}}p"):
        # Estilo TOC?
        pstyle = p.find(f"{{{W_NS}}}pPr/{{{W_NS}}}pStyle")
        if pstyle is not None:
            sid = (pstyle.get(f"{{{W_NS}}}val") or "").lower()
            if sid.startswith("toc") or "sumario" in sid or "sum" in sid:
                in_toc = True
                text = "".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t"))
                # remove números de página soltos no fim
                text = re.sub(r"\s*\.{2,}\s*\d+\s*$", "", text).strip()
                text = re.sub(r"\s+\d+\s*$", "", text).strip()
                if text:
                    hints.append(text)
    return hints


def _normalize(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().lower()


def detect_toc_region(feats: list[ParaFeatures]) -> tuple[int, int]:
    """
    Detecta o intervalo [start, end) que corresponde ao sumário inline
    do documento (não o TOC field do Word, e sim a lista digitada
    manualmente que aparece logo após uma linha 'SUMÁRIO' ou 'ÍNDICE').

    Estratégia: procura a primeira linha que case com 'sumário'/'índice'
    e estende a região enquanto os parágrafos seguintes forem curtos
    (< 25 palavras). Para quando encontra texto longo (corpo) por
    pelo menos 3 parágrafos seguidos.
    """
    start = -1
    for f in feats[:300]:  # só procura no início
        n = _normalize(f.text)
        if n in ("sumario", "indice", "sumário", "índice") or \
           re.match(r"^(sum[áa]rio|[íi]ndice)\s*$", f.text, re.I):
            start = f.index
            break
    if start < 0:
        return (-1, -1)

    end = start + 1
    long_streak = 0
    for f in feats[start + 1:]:
        if f.word_count > 25:
            long_streak += 1
            if long_streak >= 3:
                break
        else:
            long_streak = 0
        end = f.index + 1
    return (start, end)


def merge_label_titles(feats: list[ParaFeatures]) -> None:
    """
    Quando um parágrafo é só o rótulo 'Capítulo N' (ou 'Cap. N'), e o
    parágrafo seguinte não-vazio é curto, fundimos visualmente o título
    no rótulo (in-place: o rótulo recebe ' — Título' anexado, e ganha
    bônus de score implícito por estar mais informativo).
    """
    label_only = re.compile(
        r"^\s*(cap[ií]tulo|cap\.?|chapter)\s+([IVXLCDM]+|\d+)\s*[\.\:\-—]?\s*$",
        re.I,
    )
    for i, f in enumerate(feats):
        if not f.text or not label_only.match(f.text):
            continue
        # Procurar próximo parágrafo não-vazio
        for j in range(i + 1, min(i + 5, len(feats))):
            nxt = feats[j]
            if not nxt.text:
                continue
            if nxt.word_count <= 25:
                f.text = f"{f.text.strip()} — {nxt.text.strip()}"
                f.char_count = len(f.text)
                f.word_count = len(f.text.split())
                # Marca o seguido para não ser também candidato
                nxt._merged_into_prev = True  # type: ignore[attr-defined]
            break


# ---------------------------------------------------------------------------
# Pontuação
# ---------------------------------------------------------------------------

def score_paragraphs(
    feats: list[ParaFeatures],
    toc_hints: list[str],
) -> None:
    # Estatísticas globais para comparação relativa
    sizes = [f.font_size for f in feats if f.font_size is not None]
    avg_size = statistics.mean(sizes) if sizes else 11.0
    std_size = statistics.pstdev(sizes) if len(sizes) > 1 else 1.0

    norm_hints = {_normalize(h) for h in toc_hints}

    for f in feats:
        if not f.text:
            continue
        if getattr(f, "_merged_into_prev", False):
            continue

        s = 0.0

        # Penalidade severa: parágrafos com pouquíssimo conteúdo alfabético
        # (letras isoladas em glossários, divisores, etc.)
        alpha_count = sum(1 for c in f.text if c.isalpha())
        if alpha_count <= 2:
            f.score = -100
            continue
        if f.word_count == 1 and f.char_count <= 4:
            f.score = -100
            continue

        # Outline level — sinal mais forte
        if f.outline_level is not None:
            if f.outline_level == 0:
                s += 45
            elif f.outline_level == 1:
                s += 18
            elif f.outline_level == 2:
                s += 5

        # Estilo cujo nome sugere capítulo
        sn = f.style_name.lower()
        if any(k in sn for k in ("heading 1", "título 1", "titulo 1", "chapter")):
            s += 25
        elif any(k in sn for k in ("heading 2", "título 2", "titulo 2")):
            s += 8

        # Regex textual
        if f.matches_chapter_regex:
            s += 30
        elif f.matches_roman and f.char_count <= 40:
            s += 12
        elif f.matches_arabic and f.char_count <= 60:
            s += 8

        # Bate com seção especial?
        for _, pat in SECTION_PATTERNS:
            if pat.match(f.text):
                s += 28
                break

        # Quebra de página antes — forte indicador de início de capítulo
        if f.page_break_before:
            s += 18

        # Tipografia: fonte maior que a média
        if f.font_size is not None and std_size > 0:
            z = (f.font_size - avg_size) / std_size
            if z > 1.5:
                s += 12
            elif z > 0.8:
                s += 6

        # Negrito + curto + centralizado
        if f.is_bold and f.char_count <= 100:
            s += 6
        if f.alignment == "center" and f.char_count <= 100:
            s += 6
        if f.is_all_caps and 3 <= f.word_count <= 12:
            s += 8

        # Penalidade: muito longo (parágrafo de corpo)
        if f.word_count > 25:
            s -= 30
        if f.word_count > 60:
            s -= 50

        # Bônus se aparece no TOC (dica, não decisão)
        if norm_hints and _normalize(f.text) in norm_hints:
            s += 20

        f.score = s


# ---------------------------------------------------------------------------
# Seleção dos títulos finais e validação
# ---------------------------------------------------------------------------

def select_titles(feats: list[ParaFeatures]) -> list[ParaFeatures]:
    """
    Escolhe os parágrafos que são títulos de seção. Usa threshold
    relativo ao melhor score (tudo dentro de uma janela do topo),
    com piso absoluto. Aplica supressão de proximidade.
    """
    candidates = [f for f in feats if f.score >= 40 and f.text]
    if not candidates:
        # fallback frouxo se nada bate
        candidates = [f for f in feats if f.score >= 25 and f.text]
    if not candidates:
        return []

    top = max(f.score for f in candidates)
    threshold = max(40.0, top - 20.0)
    chosen = [f for f in candidates if f.score >= threshold]

    while len(chosen) < 3 and threshold > 25:
        threshold -= 5
        chosen = [f for f in candidates if f.score >= threshold]

    # Supressão de proximidade
    chosen.sort(key=lambda x: x.index)
    suppressed: list[ParaFeatures] = []
    MIN_GAP = 3
    for f in chosen:
        if suppressed and (f.index - suppressed[-1].index) < MIN_GAP:
            if f.score > suppressed[-1].score:
                suppressed[-1] = f
            continue
        suppressed.append(f)

    return suppressed


# ---------------------------------------------------------------------------
# Classificação e montagem das seções
# ---------------------------------------------------------------------------

def classify(title: str) -> str:
    t = title.strip()
    for kind, pat in SECTION_PATTERNS:
        if pat.match(t):
            return kind
    if CHAPTER_WORD_RE.search(t):
        return "capitulo"
    return "capitulo"  # default


def build_sections(
    feats: list[ParaFeatures], titles: list[ParaFeatures]
) -> list[Section]:
    sections: list[Section] = []
    if not titles:
        return sections

    # Conteúdo antes do primeiro título = frontmatter genérico (ignorado
    # ou marcado como tal). Aqui, marcamos como "frontmatter" se houver
    # texto não-trivial.
    first = titles[0].index
    pre_text = "\n".join(
        f.text for f in feats[:first] if f.text
    ).strip()
    if pre_text and len(pre_text) > 200:
        sections.append(Section(
            kind="frontmatter",
            title="(material pré-textual)",
            start_para=0,
            end_para=first,
            text=pre_text,
            confidence=0.5,
        ))

    for i, t in enumerate(titles):
        start = t.index
        end = titles[i + 1].index if i + 1 < len(titles) else len(feats)
        body = "\n".join(
            f.text for f in feats[start + 1:end] if f.text
        ).strip()
        kind = classify(t.text)
        # Confiança proporcional ao score (saturando)
        conf = max(0.0, min(1.0, t.score / 80.0))
        sections.append(Section(
            kind=kind,
            title=t.text,
            start_para=start,
            end_para=end,
            text=body,
            confidence=conf,
            detection_score=t.score,
        ))
    return sections


# ---------------------------------------------------------------------------
# API principal
# ---------------------------------------------------------------------------

def extract_chapters(path: str | Path) -> list[Section]:
    doc = Document(str(path))
    resolver = StyleResolver(doc)
    feats = extract_features(doc, resolver)
    toc_hints = extract_toc_hints(doc)

    # Funde rótulos "Capítulo N" + título da linha seguinte
    merge_label_titles(feats)

    # Detecta região do sumário inline (texto digitado, não TOC field)
    toc_start, toc_end = detect_toc_region(feats)

    score_paragraphs(feats, toc_hints)

    # Zera scores dentro da região do sumário inline
    if toc_start >= 0:
        for f in feats[toc_start:toc_end]:
            f.score = -100

    titles = select_titles(feats)
    return build_sections(feats, titles)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def _safe_filename(s: str, maxlen: int = 80) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = re.sub(r"[^\w\s\-]", "", s).strip()
    s = re.sub(r"\s+", "_", s)
    return s[:maxlen] or "secao"


def main() -> int:
    ap = argparse.ArgumentParser(
        description="Extrai capítulos e seções de um arquivo .docx"
    )
    ap.add_argument("docx", help="caminho do arquivo .docx")
    ap.add_argument("--json", help="salvar resultado em JSON")
    ap.add_argument("--outdir", help="salvar cada seção em um .txt nesse diretório")
    ap.add_argument("--quiet", action="store_true", help="não imprimir resumo")
    args = ap.parse_args()

    sections = extract_chapters(args.docx)

    if not args.quiet:
        print(f"Detectadas {len(sections)} seções:\n")
        for i, s in enumerate(sections, 1):
            print(
                f"  {i:3d}. [{s.kind:14s}] "
                f"score={s.detection_score:5.1f}  "
                f"conf={s.confidence:.2f}  "
                f"chars={len(s.text):7d}  "
                f"{s.title[:70]}"
            )

    if args.json:
        data = [
            {
                "kind": s.kind,
                "title": s.title,
                "start_para": s.start_para,
                "end_para": s.end_para,
                "confidence": s.confidence,
                "detection_score": s.detection_score,
                "text": s.text,
            }
            for s in sections
        ]
        Path(args.json).write_text(
            json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
        )
        if not args.quiet:
            print(f"\nJSON salvo em {args.json}")

    if args.outdir:
        outdir = Path(args.outdir)
        outdir.mkdir(parents=True, exist_ok=True)
        for i, s in enumerate(sections, 1):
            fname = f"{i:03d}_{s.kind}_{_safe_filename(s.title)}.txt"
            (outdir / fname).write_text(
                f"# {s.title}\n\n{s.text}\n", encoding="utf-8"
            )
        if not args.quiet:
            print(f"Seções salvas em {outdir}/")

    return 0


if __name__ == "__main__":
    sys.exit(main())