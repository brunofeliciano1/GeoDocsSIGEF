"""pessoas_utils.py — Módulo central de regras de pessoas para o plugin GeoDocsSIGEF.

Concentra funções de limpeza, formatação, busca em banco de dados, qualificação
jurídica e montagem de blocos textuais relacionados a pessoas físicas e jurídicas.
"""
import re

from ..docx_utils import normalize_key


# ══════════════════════════════════════════════════════════════════════════════
#  Limpeza e formatação básica
# ══════════════════════════════════════════════════════════════════════════════

def limpar_valor(v):
    """Remove None, NULL, 'NULL', 'None' e espaços extras."""
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.upper() in ("NULL", "NONE", "0", "FALSE") else s


def valor_valido(v):
    s = str(v).strip() if v is not None else ""
    return bool(s) and s.upper() not in ("NULL", "NONE", "0", "FALSE")


def formatar_cpf(cpf):
    d = re.sub(r"\D", "", str(cpf))
    return ("%s.%s.%s-%s" % (d[:3], d[3:6], d[6:9], d[9:])) if len(d) == 11 else cpf


def formatar_cnpj(cnpj):
    d = re.sub(r"\D", "", str(cnpj))
    return ("%s.%s.%s/%s-%s" % (d[:2], d[2:5], d[5:8], d[8:12], d[12:])) if len(d) == 14 else cnpj


def normalizar_nome_destaque(texto):
    """MAIÚSCULO para nomes em destaque (proprietário, empresa, cônjuge)."""
    v = limpar_valor(texto)
    return v.upper() if v else ""


def normalizar_nome_proprio(texto):
    """Inicial maiúscula em cada palavra (pai, mãe, logradouro)."""
    v = limpar_valor(texto)
    return " ".join(w.capitalize() for w in v.split()) if v else ""


def normalizar_texto_comum(texto):
    """Minúsculo para textos jurídicos comuns (profissão, estado civil)."""
    v = limpar_valor(texto)
    return v.lower() if v else ""


def _ec_com_regime(ec_str, regime):
    """Acrescenta 'sob o regime de ...' quando estado civil é casado/casada e regime estiver preenchido."""
    if not ec_str:
        return ec_str
    regime = limpar_valor(regime)
    if not regime:
        return ec_str
    if ec_str.lower().strip() in ("casado", "casada"):
        regime_fmt = normalizar_nome_proprio(regime)  # "comunhão parcial de bens" → "Comunhão Parcial De Bens"
        return "%s sob o regime de %s" % (ec_str, regime_fmt)
    return ec_str


# ══════════════════════════════════════════════════════════════════════════════
#  Sexo e flexão
# ══════════════════════════════════════════════════════════════════════════════

def obter_sexo(pessoa):
    sx = str(pessoa.get("sexo") or "").strip().upper()
    if sx in ("F", "FEM", "FEMIN", "FEMININO", "FEMININA", "MULHER"):
        return "FEMININO"
    if sx in ("M", "MAS", "MASC", "MASCULINO", "MASCULINA", "HOMEM"):
        return "MASCULINO"
    return sx


def flexionar(pessoa, masc, fem, neutro=None):
    """Retorna forma masculina, feminina ou neutra conforme pessoa.sexo."""
    sx = obter_sexo(pessoa)
    if sx == "MASCULINO":
        return masc
    if sx == "FEMININO":
        return fem
    return neutro if neutro is not None else masc


# ══════════════════════════════════════════════════════════════════════════════
#  Tipo de pessoa
# ══════════════════════════════════════════════════════════════════════════════

def detectar_tipo_pessoa(owner):
    """Retorna 'PF' ou 'PJ' baseado em tipo_pessoa/CPF/CNPJ."""
    tipo = normalize_key(str(owner.get("tipo_pessoa", "") or ""))
    if "juridica" in tipo:
        return "PJ"
    if "fisica" in tipo:
        return "PF"
    return "PJ" if (owner.get("cnpj") and not owner.get("cpf")) else "PF"


# ══════════════════════════════════════════════════════════════════════════════
#  Endereço
# ══════════════════════════════════════════════════════════════════════════════

def montar_endereco_pessoa(pessoa):
    """Monta endereço completo ignorando campos vazios/None/NULL."""
    if not pessoa:
        return ""
    rua = normalizar_nome_proprio(limpar_valor(pessoa.get("rua", "")))
    # Remove sufixo "Numero"/"Número" que alguns registros incluem no campo rua
    rua = re.sub(r"\s+N[uúUÚ]mero\s*$", "", rua, flags=re.IGNORECASE).strip()

    num_raw = limpar_valor(pessoa.get("numero", ""))  # limpar_valor já descarta "0"
    if num_raw.upper() == "S/N":
        num_fmt = "S/N"
    elif num_raw:
        num_fmt = "Nº %s" % num_raw
    else:
        num_fmt = ""

    comp   = normalizar_nome_proprio(limpar_valor(pessoa.get("complemento", "")))
    bairro = normalizar_nome_proprio(limpar_valor(pessoa.get("bairro", "")))
    cidade = normalizar_nome_proprio(limpar_valor(pessoa.get("cidade", "")))
    estado = normalizar_nome_proprio(limpar_valor(pessoa.get("estado", "")))
    cep    = limpar_valor(pessoa.get("cep", ""))

    logradouro = ", ".join(filter(None, [rua, num_fmt, comp]))
    partes = []
    if logradouro:
        partes.append(logradouro)
    if bairro:
        partes.append("Bairro %s" % bairro)
    local = " - ".join(filter(None, [cidade, estado]))
    if local:
        partes.append("município de %s" % local)
    if cep:
        partes.append("CEP %s" % cep)
    return ", ".join(partes)


# ══════════════════════════════════════════════════════════════════════════════
#  Data
# ══════════════════════════════════════════════════════════════════════════════

def _data_ddmmyyyy(valor):
    """Converte qualquer formato de data para dd/mm/aaaa."""
    if valor is None:
        return ""
    if hasattr(valor, "month") and callable(getattr(valor, "month", None)):
        try:
            return "%02d/%02d/%04d" % (int(valor.day()), int(valor.month()), int(valor.year()))
        except Exception:
            pass
    if hasattr(valor, "month") and not callable(getattr(valor, "month", None)):
        try:
            return "%02d/%02d/%04d" % (int(valor.day), int(valor.month), int(valor.year))
        except Exception:
            pass
    texto = str(valor).strip()
    if not texto or texto.upper() in ("NULL", "NONE", ""):
        return ""
    m = re.search(r"QDate\((\d{4}),\s*(\d{1,2}),\s*(\d{1,2})\)", texto)
    if m:
        return "%02d/%02d/%04d" % (int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})", texto)
    if m:
        return "%02d/%02d/%04d" % (int(m.group(3)), int(m.group(2)), int(m.group(1)))
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", texto)
    if m:
        dia, mes, ano = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if ano < 100:
            ano += 2000
        return "%02d/%02d/%04d" % (dia, mes, ano)
    return texto


# Linhas em branco para campos obrigatórios sem cadastro
LINHA_CURTA  = "____________________"            # nomes, CPF, RG, profissão, estado civil, nacionalidade
LINHA_MEDIA  = "________________________________________"   # nome de imóvel/propriedade
LINHA_LONGA  = "________________________________________________________________"  # endereço completo

# Alias interno mantido por compatibilidade com funções abaixo
_TRACO = LINHA_CURTA

# Valores de nacionalidade que indicam "não informado" (normalize_key já em minúsculo)
_NAT_NAO_INFORMADA = frozenset({"nao_informado", "nao informado", "n/a", "na", "-"})


def _nat_segs(s, pessoa):
    """Adiciona segmento de nacionalidade; usa traço se vazia ou 'não informado'."""
    nat_raw = normalizar_texto_comum(pessoa.get("nacionalidade", ""))
    if nat_raw and nat_raw.replace(" ", "_") not in _NAT_NAO_INFORMADA:
        s.append((", %s" % nat_raw, False))
    else:
        s.append((", nacionalidade %s" % _TRACO, False))


# ══════════════════════════════════════════════════════════════════════════════
#  Segmentos de qualificação (helpers internos)
# ══════════════════════════════════════════════════════════════════════════════

def _filiacao_segs(s, pessoa):
    pai = normalizar_nome_proprio(pessoa.get("nome_do_pai", ""))
    mae = normalizar_nome_proprio(pessoa.get("nome_da_mae", ""))
    if not (pai or mae):
        return
    filha = flexionar(pessoa, "filho", "filha", "filho(a)")
    if pai and mae:
        s.append((", %s de %s e %s" % (filha, pai, mae), False))
    elif pai:
        s.append((", %s de %s" % (filha, pai), False))
    else:
        s.append((", %s de %s" % (filha, mae), False))


def _rg_segs(s, pessoa):
    """Adiciona segmento de RG: bold quando preenchido, traço (não bold) quando vazio."""
    rg       = limpar_valor(pessoa.get("identidade", ""))
    portador = flexionar(pessoa, "portador", "portadora", "portador(a)")
    if rg:
        orgao   = limpar_valor(pessoa.get("orgao", ""))
        uf_rg   = limpar_valor(pessoa.get("uf", ""))
        emissor = "/".join(filter(None, [orgao, uf_rg]))
        rg_str  = "%s (%s)" % (rg, emissor) if emissor else rg
        s.append((", %s do RG nº %s" % (portador, rg_str), True))
    else:
        s.append((", %s do RG nº %s" % (portador, _TRACO), False))


def _cpf_segs(s, pessoa):
    """Adiciona segmento de CPF: bold quando preenchido, traço (não bold) quando vazio."""
    cpf      = limpar_valor(pessoa.get("cpf", ""))
    inscrito = flexionar(pessoa, "inscrito", "inscrita", "inscrito(a)")
    if cpf:
        s.append((", %s no CPF/MF sob o nº %s" % (inscrito, cpf), True))
    else:
        s.append((", %s no CPF/MF sob o nº %s" % (inscrito, _TRACO), False))


# ══════════════════════════════════════════════════════════════════════════════
#  Qualificação simples (helper para representantes e cônjuges)
# ══════════════════════════════════════════════════════════════════════════════

def _qualificar_simples_segs(pessoa):
    """Qualificação sem cônjuge nem sub-representantes — usada para representantes e procuradores.

    Campos obrigatórios (profissão, estado civil, RG, CPF, endereço) recebem traço
    quando não estão preenchidos no cadastro.
    """
    s = []
    nome = normalizar_nome_destaque(pessoa.get("nome", ""))
    if not nome:
        return s
    s.append((nome, True))

    # Nacionalidade — traço se vazia ou "NAO_INFORMADO"
    _nat_segs(s, pessoa)

    # Profissão — traço se vazia
    prof = normalizar_texto_comum(pessoa.get("profissao", ""))
    s.append((", %s" % (prof if prof else "profissão %s" % _TRACO), False))

    # Estado civil — traço se vazio
    ec = _ec_com_regime(
        normalizar_texto_comum(pessoa.get("estado_civil", "")),
        pessoa.get("regime_casamento", ""),
    )
    s.append((", %s" % (ec if ec else "estado civil %s" % _TRACO), False))

    # RG — traço se vazio (bold quando preenchido)
    _rg_segs(s, pessoa)

    # CPF — traço se vazio (bold quando preenchido)
    _cpf_segs(s, pessoa)

    # Endereço — linha longa se vazio (deve caber endereço completo)
    end = montar_endereco_pessoa(pessoa)
    dom = flexionar(
        pessoa,
        "residente e domiciliado",
        "residente e domiciliada",
        "residente e domiciliado(a)",
    )
    s.append((", %s em %s" % (dom, end if end else LINHA_LONGA), False))
    return s


# ══════════════════════════════════════════════════════════════════════════════
#  Qualificação jurídica — retorna [(texto, bold)]
# ══════════════════════════════════════════════════════════════════════════════

def qualificar_pessoa_fisica_segs(pessoa, conjuge=None, representantes=None, tipo_representacao=None):
    """Retorna ([(texto, bold)], avisos) para pessoa física.

    Campos obrigatórios de qualificação recebem traço quando não preenchidos no
    cadastro (profissão, estado civil, RG, CPF, endereço).
    Omitidos quando ausentes: data de nascimento, filiação, regime de casamento.
    """
    avisos = []
    s = []

    # Nome — bold; traço bold se vazio (o nome é sempre obrigatório)
    nome = normalizar_nome_destaque(pessoa.get("nome", ""))
    s.append((nome if nome else _TRACO, True))

    # Nacionalidade — traço se vazia ou "NAO_INFORMADO" (sem presumir brasileiro(a))
    _nat_segs(s, pessoa)

    # Profissão — traço se vazia
    prof = normalizar_texto_comum(pessoa.get("profissao", ""))
    s.append((", %s" % (prof if prof else "profissão %s" % _TRACO), False))

    # Estado civil — traço se vazio (cônjuge existente força "casado/casada")
    if conjuge and limpar_valor(conjuge.get("nome", "")):
        ec = _ec_com_regime(
            flexionar(pessoa, "casado", "casada", "casado(a)"),
            pessoa.get("regime_casamento", ""),
        )
    else:
        ec = _ec_com_regime(
            normalizar_texto_comum(pessoa.get("estado_civil", "")),
            pessoa.get("regime_casamento", ""),
        )
    s.append((", %s" % (ec if ec else "estado civil %s" % _TRACO), False))

    # Data de nascimento — omitir se vazia (sem traço)
    dn = _data_ddmmyyyy(pessoa.get("data_nascimento"))
    if dn:
        s.append((
            ", %s em %s" % (flexionar(pessoa, "nascido", "nascida", "nascido(a)"), dn),
            False,
        ))

    # Filiação — omitir se vazia (sem traço)
    _filiacao_segs(s, pessoa)

    # RG — traço se vazio (bold quando preenchido)
    _rg_segs(s, pessoa)

    # CPF — traço se vazio (bold quando preenchido)
    _cpf_segs(s, pessoa)

    # Cônjuge
    if conjuge and limpar_valor(conjuge.get("nome", "")):
        cnome      = normalizar_nome_destaque(conjuge["nome"])
        conj_label = flexionar(conjuge, "seu esposo", "sua esposa", "seu/sua cônjuge")
        s.append((" e %s " % conj_label, False))
        s.append((cnome, True))

        _nat_segs(s, conjuge)

        c_prof = normalizar_texto_comum(conjuge.get("profissao", ""))
        s.append((", %s" % (c_prof if c_prof else "profissão %s" % _TRACO), False))

        s.append((", %s" % _ec_com_regime(
            flexionar(conjuge, "casado", "casada", "casado(a)"),
            conjuge.get("regime_casamento", ""),
        ), False))

        c_dn = _data_ddmmyyyy(conjuge.get("data_nascimento"))
        if c_dn:
            s.append((
                ", %s em %s" % (flexionar(conjuge, "nascido", "nascida", "nascido(a)"), c_dn),
                False,
            ))
        _filiacao_segs(s, conjuge)
        _rg_segs(s, conjuge)
        _cpf_segs(s, conjuge)

    # Endereço (após cônjuge se existir) — traço se vazio
    end = montar_endereco_pessoa(conjuge or pessoa) or montar_endereco_pessoa(pessoa)
    if conjuge and limpar_valor(conjuge.get("nome", "")):
        dom = "residentes e domiciliados"
    else:
        dom = flexionar(
            pessoa,
            "residente e domiciliado",
            "residente e domiciliada",
            "residente e domiciliado(a)",
        )
    s.append((", %s em %s" % (dom, end if end else LINHA_LONGA), False))

    # Representante/procurador
    reps_validos = [r for r in (representantes or []) if r and limpar_valor(r.get("nome", ""))]
    if reps_validos:
        tv = str(tipo_representacao or "").upper()
        if tv == "PROCURADOR":
            pref_m = ", neste ato representado por seu procurador "
            pref_f = ", neste ato representada por seu procurador "
            pref_n = ", neste ato representado(a) por seu procurador "
        else:
            pref_m = ", neste ato representado por "
            pref_f = ", neste ato representada por "
            pref_n = ", neste ato representado(a) por "

        if conjuge and limpar_valor(conjuge.get("nome", "")):
            prefixo = ", neste ato representados por "
        else:
            prefixo = flexionar(pessoa, pref_m, pref_f, pref_n)

        for idx, rep in enumerate(reps_validos):
            s.append((prefixo if idx == 0 else ", ", False))
            s.extend(_qualificar_simples_segs(rep))

    return s, avisos


def qualificar_pessoa_juridica_segs(empresa, representantes):
    """Retorna ([(texto, bold)], avisos) para pessoa jurídica.

    Campos obrigatórios (CNPJ, endereço, representante) recebem traço quando ausentes.
    """
    avisos = []
    s = []

    # Nome — bold; traço bold se vazio
    nome_emp = normalizar_nome_destaque(empresa.get("nome", ""))
    s.append((nome_emp if nome_emp else _TRACO, True))
    s.append((", pessoa jurídica de direito privado", False))

    # CNPJ — traço se vazio
    cnpj = limpar_valor(empresa.get("cnpj", ""))
    if cnpj:
        s.append((", inscrita no CNPJ sob o nº ", False))
        s.append((cnpj, True))
    else:
        s.append((", inscrita no CNPJ sob o nº %s" % _TRACO, False))

    # Endereço/sede — linha longa se vazio
    end_emp = montar_endereco_pessoa(empresa)
    s.append((", com sede em %s" % (end_emp if end_emp else LINHA_LONGA), False))

    # Representantes — traço se nenhum cadastrado
    reps_validos = [r for r in representantes if r and limpar_valor(r.get("nome", ""))]
    if not reps_validos:
        s.append((", neste ato representada por %s, CPF nº %s" % (_TRACO, _TRACO), False))
        avisos.append("Aviso: '%s' não possui representante legal cadastrado." % nome_emp)
    elif len(reps_validos) == 1:
        s.append((", neste ato representada por ", False))
        s.extend(_qualificar_simples_segs(reps_validos[0]))
    else:
        s.append((", neste ato representada por seus representantes legais ", False))
        for idx, r in enumerate(reps_validos):
            if idx:
                s.append((", ", False))
            s.extend(_qualificar_simples_segs(r))

    return s, avisos


# ══════════════════════════════════════════════════════════════════════════════
#  Qualificação completa de um bloco de proprietários → [(texto, bold)]
# ══════════════════════════════════════════════════════════════════════════════

def qualificar_proprietarios_segs(proprietarios, avisos_out=None):
    """Retorna [(texto, bold)] com qualificação completa de todos os proprietários."""
    if avisos_out is None:
        avisos_out = []
    if not proprietarios:
        return [("[PROPRIETÁRIO DO IMÓVEL PRINCIPAL NÃO INFORMADO]", True)]

    todos_segs = []
    for i, owner in enumerate(proprietarios):
        if i:
            todos_segs.append(("; ", False))
        tipo = owner.get("_tipo") or detectar_tipo_pessoa(owner)
        if tipo == "PJ":
            segs, avs = qualificar_pessoa_juridica_segs(owner, owner.get("_representantes", []))
        else:
            segs, avs = qualificar_pessoa_fisica_segs(
                owner,
                conjuge=owner.get("_conjuge"),
                representantes=owner.get("_representantes"),
                tipo_representacao=owner.get("_tipo_representacao"),
            )
        avisos_out.extend(avs)
        todos_segs.extend(segs)
    return todos_segs


# ══════════════════════════════════════════════════════════════════════════════
#  Assinaturas
# ══════════════════════════════════════════════════════════════════════════════

def montar_assinaturas_pessoas(declarantes):
    """Retorna [(linha1, linha2)] para cada assinante.

    - PF sem representante: próprio assina (+ cônjuge se houver)
    - PF com representante: representante assina com rótulo 'Representante/Procurador de: NOME'
    - PJ: representante legal assina com rótulo 'Representante legal de: EMPRESA'
    """
    resultado = []
    for decl in declarantes:
        tipo     = decl.get("_tipo") or detectar_tipo_pessoa(decl)
        nome_dec = normalizar_nome_destaque(decl.get("nome", ""))
        reps     = [r for r in decl.get("_representantes", []) if r and r.get("nome")]

        if tipo == "PJ":
            if reps:
                for rep in reps:
                    rnome  = normalizar_nome_destaque(rep["nome"])
                    rcpf   = limpar_valor(rep.get("cpf", ""))
                    partes = ["Representante legal de: %s" % nome_dec]
                    if rcpf:
                        partes.append("CPF: %s" % rcpf)
                    resultado.append((rnome, "\n".join(partes)))
            else:
                resultado.append((nome_dec, "Representante Legal"))
        else:
            # PF
            if reps:
                # Representante/procurador assina pelo proprietário PF
                tv = str(decl.get("_tipo_representacao") or "REPRESENTANTE").upper()
                tv_label = "Procurador de" if tv == "PROCURADOR" else "Representante de"
                for rep in reps:
                    rnome  = normalizar_nome_destaque(rep.get("nome", ""))
                    rcpf   = limpar_valor(rep.get("cpf", ""))
                    partes = ["%s: %s" % (tv_label, nome_dec)]
                    if rcpf:
                        partes.append("CPF: %s" % rcpf)
                    resultado.append((rnome, "\n".join(partes)))
            else:
                # Próprio PF assina (+ cônjuge)
                cpf = limpar_valor(decl.get("cpf", ""))
                resultado.append((nome_dec, "CPF: %s" % cpf if cpf else ""))
                conjuge = decl.get("_conjuge")
                if conjuge and conjuge.get("nome"):
                    cnome = normalizar_nome_destaque(conjuge["nome"])
                    ccpf  = limpar_valor(conjuge.get("cpf", ""))
                    resultado.append((cnome, "CPF: %s" % ccpf if ccpf else ""))
    return resultado


# ══════════════════════════════════════════════════════════════════════════════
#  Execução de SQL
# ══════════════════════════════════════════════════════════════════════════════

def _exec_psycopg2(layer, sql, params):
    import psycopg2
    from qgis.core import QgsDataSourceUri
    uri = QgsDataSourceUri(layer.dataProvider().dataSourceUri())
    conn = psycopg2.connect(
        host=uri.host() or "localhost",
        port=int(uri.port()) if uri.port() else 5432,
        dbname=uri.database(),
        user=uri.username(),
        password=uri.password(),
    )
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    cols = [d[0] for d in cur.description]
    cur.close()
    conn.close()
    return rows, cols


def _exec_qgis(layer, sql):
    try:
        from qgis.core import QgsProviderRegistry
        meta = QgsProviderRegistry.instance().providerMetadata("postgres")
        if meta:
            conn = meta.createConnection(layer.dataProvider().dataSourceUri(), {})
            if conn:
                return conn.executeSql(sql) or []
    except Exception:
        pass
    return []


# ══════════════════════════════════════════════════════════════════════════════
#  Conversão de rows → dicionário de pessoa
# ══════════════════════════════════════════════════════════════════════════════

_COLS_PESSOA = [
    "ordem", "pessoa_id", "nome", "tipo_pessoa", "cpf", "cnpj",
    "identidade", "orgao", "uf", "data_nascimento", "profissao", "estado_civil",
    "nacionalidade", "regime_casamento", "nome_da_mae", "nome_do_pai",
    "rua", "numero", "complemento", "cep", "bairro", "cidade", "estado", "sexo",
]


def _row_to_pessoa(row, cols):
    record = dict(zip(cols, row))
    pessoa = {}
    for k, v in record.items():
        val = limpar_valor(v)
        if valor_valido(val):
            pessoa[k] = val
    if pessoa.get("cpf"):
        pessoa["cpf"] = formatar_cpf(pessoa["cpf"])
    if pessoa.get("cnpj"):
        pessoa["cnpj"] = formatar_cnpj(pessoa["cnpj"])
    return pessoa if pessoa.get("nome") else None


def linhas_para_pessoas(rows, cols):
    """Converte lista de tuplas SQL em lista de dicionários de pessoa."""
    pessoas = []
    for i, row in enumerate(rows):
        rec = dict(zip(cols, row))
        nome = limpar_valor(rec.get("nome", ""))
        if not nome:
            continue
        suffix = "" if i == 0 else "_%d" % i
        pessoa = {"nome": nome, "suffix": suffix}
        for k, v in rec.items():
            val = limpar_valor(v)
            if valor_valido(val) and k != "nome":
                pessoa[k] = val
        if pessoa.get("cpf"):
            pessoa["cpf"] = formatar_cpf(pessoa["cpf"])
        if pessoa.get("cnpj"):
            pessoa["cnpj"] = formatar_cnpj(pessoa["cnpj"])
        pessoas.append(pessoa)
    return pessoas


# ══════════════════════════════════════════════════════════════════════════════
#  Busca de vínculos (pessoas_vinculos)
# ══════════════════════════════════════════════════════════════════════════════

_SQL_VINCULO_BASE = """
    SELECT p.id AS pessoa_id, p.nome, p.tipo_pessoa, p.cpf, p.cnpj,
           p.identidade, p.orgao, p.uf,
           p.data_nascimento, p.profissao, p.estado_civil,
           p.nacionalidade, p.regime_casamento, p.nome_da_mae, p.nome_do_pai,
           p.rua, p.numero, p.complemento, p.cep, p.bairro, p.cidade, p.estado,
           p.sexo
    FROM servicos.pessoas_vinculos pv
    JOIN servicos.pessoas p ON p.id = pv.pessoa_vinculada_id
    WHERE pv.pessoa_id = %s AND pv.tipo_vinculo = %s
    ORDER BY pv.id
"""
_SQL_VINCULO_ATIVO = _SQL_VINCULO_BASE.replace(
    "ORDER BY", "AND pv.ativo = true ORDER BY"
)
_COLS_VINCULO = [
    "pessoa_id", "nome", "tipo_pessoa", "cpf", "cnpj",
    "identidade", "orgao", "uf", "data_nascimento", "profissao", "estado_civil",
    "nacionalidade", "regime_casamento", "nome_da_mae", "nome_do_pai",
    "rua", "numero", "complemento", "cep", "bairro", "cidade", "estado", "sexo",
]


def buscar_vinculados(layer, pessoa_id, tipo_vinculo, limite=None):
    """Busca todos os vínculos ativos de um tipo. Retorna lista de pessoas."""
    rows, desc = [], _COLS_VINCULO

    try:
        rows, desc = _exec_psycopg2(layer, _SQL_VINCULO_ATIVO, (int(pessoa_id), tipo_vinculo))
    except Exception:
        try:
            rows, desc = _exec_psycopg2(layer, _SQL_VINCULO_BASE, (int(pessoa_id), tipo_vinculo))
        except Exception:
            pass

    if not rows:
        try:
            from qgis.core import QgsProviderRegistry
            meta = QgsProviderRegistry.instance().providerMetadata("postgres")
            if meta:
                conn = meta.createConnection(layer.dataProvider().dataSourceUri(), {})
                if conn:
                    for sql in (_SQL_VINCULO_ATIVO, _SQL_VINCULO_BASE):
                        sql2 = sql.replace("%s", "{}").format(
                            int(pessoa_id), "'%s'" % tipo_vinculo
                        )
                        try:
                            rows = conn.executeSql(sql2) or []
                            if rows:
                                break
                        except Exception:
                            continue
        except Exception:
            pass

    pessoas = [_row_to_pessoa(r, desc) for r in rows]
    pessoas = [p for p in pessoas if p]
    return pessoas[:limite] if limite else pessoas


def buscar_vinculado(layer, pessoa_id, tipo_vinculo):
    """Busca um único vínculo. Retorna pessoa ou None."""
    resultado = buscar_vinculados(layer, pessoa_id, tipo_vinculo, limite=1)
    return resultado[0] if resultado else None


def buscar_conjuge(layer, pessoa_id):
    """Busca o cônjuge de uma pessoa. Retorna pessoa ou None."""
    return buscar_vinculado(layer, pessoa_id, "CONJUGE")


def buscar_representantes_legais(layer, pessoa_id):
    """Busca representantes legais de uma pessoa jurídica."""
    return buscar_vinculados(layer, pessoa_id, "REPRESENTANTE_LEGAL")


_TIPOS_REPRESENTACAO = ("REPRESENTANTE_LEGAL", "PROCURADOR", "REPRESENTANTE", "RESPONSAVEL")


def buscar_procuradores_representantes(layer, pessoa_id):
    """Busca representantes/procuradores em ordem de prioridade.

    Retorna (lista_representantes, tipo_vinculo_encontrado).
    """
    if not layer:
        return [], None
    for tv in _TIPOS_REPRESENTACAO:
        try:
            reps = buscar_vinculados(layer, pessoa_id, tv)
            if reps:
                return reps, tv
        except Exception:
            continue
    return [], None


# ══════════════════════════════════════════════════════════════════════════════
#  Busca de proprietários (servico_pessoas e confinante_pessoas)
# ══════════════════════════════════════════════════════════════════════════════

_SQL_PROPRIETARIOS = """
    SELECT sp.ordem, p.id AS pessoa_id, p.nome, p.tipo_pessoa, p.cpf, p.cnpj,
           p.identidade, p.orgao, p.uf,
           p.data_nascimento, p.profissao, p.estado_civil,
           p.nacionalidade, p.regime_casamento,
           p.nome_da_mae, p.nome_do_pai,
           p.rua, p.numero, p.complemento, p.cep, p.bairro, p.cidade, p.estado,
           p.sexo
    FROM servicos.servico_pessoas sp
    JOIN servicos.pessoas p ON p.id = sp.pessoa_id
    WHERE sp.servico_id = %s
      AND COALESCE(sp.tipo_vinculo, 'PROPRIETARIO') = 'PROPRIETARIO'
    ORDER BY sp.ordem, sp.id
"""

_SQL_CONFINANTE_PESSOAS = """
    SELECT cp.ordem, p.id AS pessoa_id, p.nome, p.tipo_pessoa, p.cpf, p.cnpj,
           p.identidade, p.orgao, p.uf,
           p.data_nascimento, p.profissao, p.estado_civil,
           p.nacionalidade, p.regime_casamento,
           p.nome_da_mae, p.nome_do_pai,
           p.rua, p.numero, p.complemento, p.cep, p.bairro, p.cidade, p.estado,
           p.sexo
    FROM servicos.confinante_pessoas cp
    JOIN servicos.pessoas p ON p.id = cp.pessoa_id
    WHERE cp.confinante_id = %s
    ORDER BY cp.ordem, cp.id
"""


def buscar_proprietarios_servico(layer, servico_id):
    """Busca proprietários do serviço principal (servico_pessoas)."""
    try:
        rows, cols = _exec_psycopg2(layer, _SQL_PROPRIETARIOS, (int(servico_id),))
        if rows:
            return linhas_para_pessoas(rows, cols)
    except Exception:
        pass
    try:
        sql = _SQL_PROPRIETARIOS.replace("%s", str(int(servico_id)))
        rows = _exec_qgis(layer, sql)
        if rows:
            return linhas_para_pessoas(rows, _COLS_PESSOA)
    except Exception:
        pass
    return []


def buscar_proprietarios_confinante(layer, confinante_id):
    """Busca proprietários do confinante (confinante_pessoas)."""
    try:
        rows, cols = _exec_psycopg2(layer, _SQL_CONFINANTE_PESSOAS, (int(confinante_id),))
        if rows:
            return linhas_para_pessoas(rows, cols)
    except Exception:
        pass
    try:
        sql = _SQL_CONFINANTE_PESSOAS.replace("%s", str(int(confinante_id)))
        rows = _exec_qgis(layer, sql)
        if rows:
            return linhas_para_pessoas(rows, _COLS_PESSOA)
    except Exception:
        pass
    return []


def buscar_dados_servico(layer, servico_id):
    """Busca dados do serviço principal via SELECT *."""
    try:
        rows, cols = _exec_psycopg2(
            layer,
            "SELECT * FROM servicos.servicos WHERE id = %s",
            (int(servico_id),),
        )
        if rows:
            rec = dict(zip(cols, rows[0]))
            if "nome_cartorio" not in rec:
                rec["nome_cartorio"] = limpar_valor(rec.get("cartorio", ""))
            return rec
    except Exception:
        pass
    return {}


# ══════════════════════════════════════════════════════════════════════════════
#  Enriquecimento com vínculos
# ══════════════════════════════════════════════════════════════════════════════

def _enriquecer_entry(entry, pid, tipo, layer):
    """Enriquece um dict de pessoa com cônjuge (PF) e/ou representantes."""
    if not pid or not layer:
        return entry
    if tipo == "PF":
        try:
            entry["_conjuge"] = buscar_conjuge(layer, pid)
        except Exception:
            pass
    # Busca representantes para PF e PJ
    for tv in _TIPOS_REPRESENTACAO:
        try:
            reps = buscar_vinculados(layer, pid, tv)
            if reps:
                entry["_representantes"] = reps
                entry["_tipo_representacao"] = tv
                break
        except Exception:
            continue
    return entry


def enriquecer_com_conjuge(owners, layer):
    """Busca cônjuge de cada proprietário PF em pessoas_vinculos (compat. legada).

    Para enriquecimento completo (cônjuge + representantes), use enriquecer_proprietarios.
    """
    if not layer:
        return owners
    enriched = []
    for owner in owners:
        entry = dict(owner)
        entry["_conjuge"] = None
        pid = owner.get("pessoa_id")
        if pid and detectar_tipo_pessoa(owner) == "PF":
            try:
                entry["_conjuge"] = buscar_conjuge(layer, pid)
            except Exception:
                pass
        enriched.append(entry)
    return enriched


def enriquecer_proprietarios(owners, layer):
    """Enriquece cada proprietário com cônjuge (PF) e representantes (PF e PJ).

    Diferente de enriquecer_com_conjuge: também busca REPRESENTANTE_LEGAL, PROCURADOR,
    REPRESENTANTE e RESPONSAVEL para proprietários PF e PJ.
    """
    enriched = []
    for owner in owners:
        entry = dict(owner)
        tipo = detectar_tipo_pessoa(owner)
        entry["_tipo"] = tipo
        entry["_conjuge"] = None
        entry["_representantes"] = []
        entry["_tipo_representacao"] = None
        pid = owner.get("pessoa_id")
        if pid and layer:
            _enriquecer_entry(entry, pid, tipo, layer)
        enriched.append(entry)
    return enriched


def expandir_com_conjuges(owners):
    """Lista plana de proprietários + cônjuges para blocos de assinatura legada."""
    flat = []
    for owner in owners:
        flat.append(owner)
        conjuge = owner.get("_conjuge")
        if conjuge and conjuge.get("nome"):
            flat.append(conjuge)
    return flat


def enriquecer_confinante(pessoas, layer):
    """Enriquece cada pessoa do confinante com cônjuge (PF) e representantes (PF e PJ)."""
    declarantes = []
    for p in pessoas:
        tipo = detectar_tipo_pessoa(p)
        entry = dict(p)
        entry["_tipo"] = tipo
        entry["_conjuge"] = None
        entry["_representantes"] = []
        entry["_tipo_representacao"] = None
        pid = p.get("pessoa_id")
        if pid and layer:
            _enriquecer_entry(entry, pid, tipo, layer)
        declarantes.append(entry)
    return declarantes


# ══════════════════════════════════════════════════════════════════════════════
#  Bloco de proprietários simples (texto inline — compatibilidade legada)
# ══════════════════════════════════════════════════════════════════════════════

def montar_bloco_proprietarios_simples(owners):
    """Texto inline resumido: 'NOME, CPF: xxx, e sua esposa NOME2, CPF: yyy'.

    Usado como fallback. Para qualificação completa, use qualificar_proprietarios_segs.
    """
    if not owners:
        return "[PROPRIETÁRIO DO IMÓVEL PRINCIPAL NÃO INFORMADO]"
    partes = []
    for owner in owners:
        nome = owner.get("nome", "").upper()
        cpf  = limpar_valor(owner.get("cpf", ""))
        cnpj = limpar_valor(owner.get("cnpj", ""))
        if cpf:
            parte = "%s, CPF: %s" % (nome, cpf)
        elif cnpj:
            parte = "%s, CNPJ: %s" % (nome, cnpj)
        else:
            parte = nome
        conjuge = owner.get("_conjuge")
        if conjuge and conjuge.get("nome"):
            conj_nome = conjuge.get("nome", "").upper()
            conj_cpf  = limpar_valor(conjuge.get("cpf", ""))
            try:
                label = flexionar(conjuge, "seu esposo", "sua esposa", "seu/sua cônjuge")
            except Exception:
                label = "seu/sua cônjuge"
            conj_parte = "%s %s" % (label, conj_nome)
            if conj_cpf:
                conj_parte += ", CPF: %s" % conj_cpf
            parte = "%s, e %s" % (parte, conj_parte)
        partes.append(parte)
    if len(partes) == 1:
        return partes[0]
    return "; ".join(partes)


# ══════════════════════════════════════════════════════════════════════════════
#  Bloco resumido de proprietários para Laudo Técnico — retorna [(texto, bold)]
# ══════════════════════════════════════════════════════════════════════════════

def _rep_resumido_segs(rep):
    """Qualificação resumida de representante/procurador para Laudo Técnico.

    Formato: NOME[, CPF nº XXX][, RG nº XXX (ORGÃO/UF)]
    """
    s = []
    nome = normalizar_nome_destaque(rep.get("nome", ""))
    if not nome:
        return s
    s.append((nome, True))
    cpf = limpar_valor(rep.get("cpf", ""))
    if cpf:
        s.append((", CPF nº %s" % cpf, True))
    else:
        s.append((", CPF nº %s" % _TRACO, False))
    rg = limpar_valor(rep.get("identidade", ""))
    if rg:
        orgao   = limpar_valor(rep.get("orgao", ""))
        uf_rg   = limpar_valor(rep.get("uf", ""))
        emissor = "/".join(filter(None, [orgao, uf_rg]))
        rg_str  = "%s (%s)" % (rg, emissor) if emissor else rg
        s.append((", RG nº %s" % rg_str, True))
    return s


def montar_bloco_proprietarios_resumido_segs(proprietarios):
    """Bloco resumido para Laudo Técnico: nome + CPF/CNPJ + cônjuge + representante.

    Retorna [(texto, bold)]. Não inclui: nacionalidade, profissão, estado civil,
    regime de casamento, nascimento, filiação, endereço do proprietário.
    Representante/procurador: nome + CPF nº + RG nº (resumido).
    """
    if not proprietarios:
        return [("[PROPRIETÁRIO DO IMÓVEL PRINCIPAL NÃO INFORMADO]", True)]

    todos_segs = []
    for i, owner in enumerate(proprietarios):
        if i:
            todos_segs.append(("; ", False))

        tipo = owner.get("_tipo") or detectar_tipo_pessoa(owner)
        nome = normalizar_nome_destaque(owner.get("nome", ""))
        todos_segs.append((nome if nome else _TRACO, True))

        if tipo == "PJ":
            cnpj = limpar_valor(owner.get("cnpj", ""))
            if cnpj:
                todos_segs.append((", CNPJ: %s" % cnpj, True))
            else:
                todos_segs.append((", CNPJ: %s" % _TRACO, False))
        else:
            cpf = limpar_valor(owner.get("cpf", ""))
            if cpf:
                todos_segs.append((", CPF: %s" % cpf, True))
            else:
                todos_segs.append((", CPF: %s" % _TRACO, False))

            conjuge = owner.get("_conjuge")
            if conjuge and limpar_valor(conjuge.get("nome", "")):
                conj_nome = normalizar_nome_destaque(conjuge["nome"])
                conj_cpf  = limpar_valor(conjuge.get("cpf", ""))
                label     = flexionar(conjuge, "seu esposo", "sua esposa", "seu/sua cônjuge")
                todos_segs.append((", e %s " % label, False))
                todos_segs.append((conj_nome, True))
                if conj_cpf:
                    todos_segs.append((", CPF: %s" % conj_cpf, True))

        # Representante/procurador — formato resumido: nome + CPF nº + RG nº
        reps = [r for r in (owner.get("_representantes") or []) if r and limpar_valor(r.get("nome", ""))]
        if reps:
            tv = str(owner.get("_tipo_representacao") or "").upper()
            if tv == "PROCURADOR":
                pref_m = ", neste ato representado por seu procurador "
                pref_f = ", neste ato representada por seu procurador "
                pref_n = ", neste ato representado(a) por seu procurador "
            else:
                pref_m = ", neste ato representado por "
                pref_f = ", neste ato representada por "
                pref_n = ", neste ato representado(a) por "

            conj = owner.get("_conjuge")
            if conj and limpar_valor(conj.get("nome", "")):
                prefixo = ", neste ato representados por "
            else:
                prefixo = flexionar(owner, pref_m, pref_f, pref_n)

            for idx, rep in enumerate(reps):
                todos_segs.append((prefixo if idx == 0 else ", ", False))
                todos_segs.extend(_rep_resumido_segs(rep))

    return todos_segs
