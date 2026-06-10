import os
from datetime import datetime

from .declaracao_confrontantes import (
    fill_declaracao_confrontantes_template,
    nome_confinante_para_arquivo,
    sanitizar_nome_arquivo,
)


def gerar_declaracoes_confrontantes_lote(
    template_path,
    output_dir,
    confinante_layer,
    servico_id,
    pdf_segments,
    pdf_data,
    denominacao=None,
):
    """Gera uma declaração de confrontante para cada confinante do serviço.

    Busca na camada todas as features onde int(feature["codigo"]) == servico_id
    e chama fill_declaracao_confrontantes_template para cada uma.

    Retorna (gerados, falhas) onde:
      gerados — lista de caminhos de arquivos criados com sucesso
      falhas  — lista de (identificador, mensagem_de_erro)
    """
    # ── Busca confinantes na camada ──────────────────────────────────────────
    confinante_features = []
    try:
        servico_id_int = int(servico_id)
        for feature in confinante_layer.getFeatures():
            try:
                codigo_raw = str(feature["codigo"] or "").strip()
                if not codigo_raw or codigo_raw.upper() in ("NULL", "NONE"):
                    continue
                if int(float(codigo_raw)) == servico_id_int:
                    confinante_features.append(feature)
            except (ValueError, TypeError):
                continue
    except Exception as exc:
        raise RuntimeError(
            "Erro ao buscar confinantes na camada '%s': %s"
            % (confinante_layer.name(), str(exc))
        )

    if not confinante_features:
        raise RuntimeError(
            "Nenhum confinante encontrado para o serviço id=%s." % servico_id
        )

    # ── Cria pasta de saída do lote ──────────────────────────────────────────
    if denominacao:
        sufixo = sanitizar_nome_arquivo(denominacao)[:40]
        nome_pasta = "DECLARACOES_CONFRONTANTES_%s" % sufixo if sufixo else "DECLARACOES_CONFRONTANTES"
    else:
        nome_pasta = "DECLARACOES_CONFRONTANTES"
    lote_dir = os.path.join(output_dir, nome_pasta)
    os.makedirs(lote_dir, exist_ok=True)

    # ── Gera declaração para cada confinante ─────────────────────────────────
    gerados = []
    falhas  = []

    for cf in confinante_features:
        nome_base = nome_confinante_para_arquivo(confinante_layer, cf)
        try:
            cid_raw = str(cf["id"] or "").strip()
        except Exception:
            cid_raw = "?"
        label_falha = (
            nome_base if nome_base != "CONFINANTE"
            else ("CONFINANTE_id_%s" % cid_raw)
        )

        nome_arquivo = "DECLARACAO_CONFRONTANTE_%s.docx" % nome_base
        output_path  = os.path.join(lote_dir, nome_arquivo)

        # Evita sobrescrever arquivo existente (mesmo comportamento de _safe_output_path)
        if os.path.exists(output_path):
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            base, ext = os.path.splitext(output_path)
            output_path = "%s_%s%s" % (base, ts, ext)

        try:
            fill_declaracao_confrontantes_template(
                template_path,
                output_path,
                confinante_layer,
                cf,
                pdf_segments,
                pdf_data,
            )
            gerados.append(output_path)
        except Exception as exc:
            falhas.append((label_falha, str(exc)))

    return gerados, falhas
