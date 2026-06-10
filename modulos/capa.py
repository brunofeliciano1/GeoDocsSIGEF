import os
import shutil

from ..docx_utils import (
    ensure_docx_path,
    substituir_marcadores_docx,
    normalize_key,
    all_document_paragraphs,
    PLACEHOLDER_PATTERN,
    replace_match_across_runs,
    remover_paragrafos_placeholder_vazio,
    gerar_qrcode_imagem,
    substituir_qrcode_por_imagem,
)
from .pessoas_utils import (
    enriquecer_proprietarios,
    normalizar_nome_destaque,
    limpar_valor,
)


def find_capa_template(plugin_dir):
    path = os.path.join(plugin_dir, "models", "capa.docx")
    return path if os.path.exists(path) else None


def fill_capa_template(template_path, output_path, data, layer=None):
    try:
        from docx import Document
    except ImportError as exc:
        raise RuntimeError(
            "Biblioteca python-docx não encontrada. Instale python-docx no Python do QGIS."
        ) from exc

    template_path = os.path.abspath(template_path)
    output_path   = os.path.abspath(ensure_docx_path(output_path))

    if template_path == output_path:
        raise RuntimeError("Escolha um arquivo de saída diferente do modelo original.")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    shutil.copyfile(template_path, output_path)

    document = Document(output_path)

    # Enriquece owners para detectar cônjuge corretamente
    owners      = data.get("__owners", [])
    declarantes = enriquecer_proprietarios(owners, layer)

    # {{BLOCO_PROPRIETARIOS}}: substituído antes do loop normal para não ser
    # apagado como texto vazio por substituir_marcadores_docx().
    _replace_bloco_proprietarios_capa(document, declarantes)

    # {{QR_CODE}}: imagem gerada a partir do campo qrcode_pasta_drive
    qr_link = str(data.get("qr_code") or "").strip()
    if qr_link:
        import tempfile
        _tmp_fd, _tmp_path = tempfile.mkstemp(suffix=".png")
        os.close(_tmp_fd)
        try:
            gerar_qrcode_imagem(qr_link, _tmp_path)
            substituir_qrcode_por_imagem(document, _tmp_path, largura_cm=4.17, altura_cm=4.17)
        except Exception:
            # Se a geração falhar, garante que o placeholder seja removido
            remover_paragrafos_placeholder_vazio(document, {"qr_code": ""})
        finally:
            try:
                os.remove(_tmp_path)
            except Exception:
                pass
    else:
        # Sem link: remove o parágrafo contendo {{QR_CODE}}
        remover_paragrafos_placeholder_vazio(document, {"qr_code": ""})

    # Demais marcadores substituídos normalmente
    substituir_marcadores_docx(document, data)

    # Aplica Arial 12pt em todos os runs do documento
    _aplicar_fonte_capa(document)

    document.save(output_path)


def _replace_bloco_proprietarios_capa(document, declarantes):
    """Substitui APENAS o marcador {{BLOCO_PROPRIETARIOS}} dentro do parágrafo,
    preservando qualquer texto fixo anterior (ex: "PROPRIETÁRIO(A):").

    Regras:
    - Linha única, proprietários separados por vírgula.
    - PF com cônjuge: "NOME e CÔNJUGE" mantidos juntos antes da vírgula.
    - Sem CPF, CNPJ, qualificação ou rótulos.
    - Não limpa o parágrafo inteiro — troca somente o marcador via replace_match_across_runs.
    """
    entradas = []
    for decl in declarantes:
        nome = normalizar_nome_destaque(decl.get("nome", ""))
        if not nome:
            continue
        conjuge = decl.get("_conjuge")
        if conjuge and limpar_valor(conjuge.get("nome", "")):
            cnome = normalizar_nome_destaque(conjuge["nome"])
            entradas.append("%s e %s" % (nome, cnome))
        else:
            entradas.append(nome)

    nomes = ", ".join(entradas) if entradas else "[PROPRIETÁRIO NÃO INFORMADO]"

    for paragraph in all_document_paragraphs(document):
        runs = paragraph.runs
        if not runs:
            continue

        full_text = "".join(run.text for run in runs)

        # Encontra ocorrências do marcador bloco_proprietarios neste parágrafo
        matches = [
            m for m in PLACEHOLDER_PATTERN.finditer(full_text)
            if normalize_key(m.group(2) or m.group(3) or m.group(4) or "") == "bloco_proprietarios"
        ]
        if not matches:
            continue

        # Monta mapa de posições dos runs para replace_match_across_runs
        run_ranges = []
        position = 0
        for run in runs:
            start = position
            end   = start + len(run.text)
            run_ranges.append((run, start, end))
            position = end

        # Substitui cada ocorrência (em ordem reversa para não deslocar índices)
        for match in reversed(matches):
            replace_match_across_runs(runs, run_ranges, match.start(), match.end(), nomes)


def _aplicar_fonte_capa(document):
    """Aplica Arial 12pt em todos os runs do documento da capa."""
    try:
        from docx.shared import Pt
        from docx.oxml.ns import qn
    except ImportError:
        return

    for paragraph in all_document_paragraphs(document):
        for run in paragraph.runs:
            run.font.name = "Arial"
            run.font.size = Pt(12)
            try:
                r_pr = run._element.get_or_add_rPr()
                r_fonts = r_pr.rFonts
                if r_fonts is not None:
                    r_fonts.set(qn("w:eastAsia"), "Arial")
            except Exception:
                pass
