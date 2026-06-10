"""Classificação de confrontantes que dispensam anuência formal.

Função principal:
    classificar_confrontante_dispensa(nome_confrontante, municipio=None, uf=None)
    → {"categoria": str, "responsavel": str, "texto": str}

O texto retornado é o complemento do bullet da Declaração de Dispensa de Anuência
(o que vem após a vírgula, sem ponto-e-vírgula final).
"""

import re
import unicodedata


# ── Normalização ───────────────────────────────────────────────────────────────

def _norm(texto):
    """Remove acentos, maiúscula, colapsa espaços."""
    s = unicodedata.normalize("NFKD", str(texto or ""))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s.upper().strip())


def _substituir_locais(texto, municipio, uf):
    """Substitui {{MUNICIPIO}} e {{UF}} quando informados."""
    if municipio:
        texto = texto.replace("{{MUNICIPIO}}", municipio.strip())
    if uf:
        texto = texto.replace("{{UF}}", uf.strip().upper())
    return texto


# ── Tabela de órgãos rodoviários estaduais ────────────────────────────────────

_ORGAOS_ESTADUAIS = {
    "AC": "Departamento de Estradas de Rodagem do Acre – DERACRE",
    "AL": "Departamento de Estradas de Rodagem de Alagoas – DER/AL",
    "AM": "Departamento de Estradas de Rodagem do Amazonas – DEAM",
    "AP": "Departamento de Estradas e Rodagens do Amapá – DER/AP",
    "BA": "Departamento de Infraestrutura de Transportes da Bahia – DER/BA",
    "CE": "Superintendência de Obras Públicas do Estado do Ceará – SOP",
    "ES": "Departamento de Estradas de Rodagem do Espírito Santo – DER/ES",
    "GO": "Agência Goiana de Infraestrutura e Transportes – GOINFRA",
    "MA": "Departamento de Estradas de Rodagem do Maranhão – DEINFRA/MA",
    "MG": "Departamento de Estradas de Rodagem de Minas Gerais – DER/MG",
    "MS": "Agência Estadual de Gestão de Empreendimentos do Mato Grosso do Sul – AGESUL",
    "MT": "Secretaria de Estado de Infraestrutura e Logística do Mato Grosso",
    "PA": "Secretaria de Estado de Transportes do Pará – SETRAN/PA",
    "PB": "Departamento de Estradas de Rodagem da Paraíba – DER/PB",
    "PE": "Departamento de Infraestrutura e Transporte de Pernambuco – DIRCON/PE",
    "PI": "Departamento de Estradas de Rodagem do Piauí – DER/PI",
    "PR": "Departamento de Estradas de Rodagem do Paraná – DER/PR",
    "RJ": "Departamento de Estradas de Rodagem do Rio de Janeiro – DER/RJ",
    "RN": "Departamento de Estradas de Rodagem do Rio Grande do Norte – DER/RN",
    "RO": "Departamento de Estradas de Rodagem do Estado de Rondônia – DER/RO",
    "RR": "Departamento de Estradas e Rodagens de Roraima – DER/RR",
    "RS": "Departamento Autônomo de Estradas de Rodagem do Rio Grande do Sul – DAER/RS",
    "SC": "Departamento Estadual de Infraestrutura de Santa Catarina – DEINFRA/SC",
    "SE": "Departamento Estadual de Infraestrutura Rodoviária de Sergipe – DER/SE",
    "SP": "Departamento de Estradas de Rodagem de São Paulo – DER/SP",
    "TO": "Departamento de Estradas de Rodagem do Tocantins – DER/TO",
}

# Siglas de UF usadas como prefixo de número de rodovia (CE-085, GO-010 etc.)
_SIGLAS_UF = sorted(_ORGAOS_ESTADUAIS.keys(), key=len, reverse=True)


# ── Tabela de regras: (padrão_regex, categoria, responsavel, texto_template) ──
# Avaliada em ordem — primeira correspondência vence.
# Os textos usam {{MUNICIPIO}} e {{UF}} como placeholders substituídos em runtime.

_REGRAS = [

    # A) Recursos hídricos
    (
        r"\b(RIO|RIACHO|CORREGO|RIBEIRAO|IGARAPE|CANAL\s+NATURAL|ACUDE|"
        r"LAGOA|LAGO|RESERVATORIO|BARRAGEM)\b",
        "Recurso hídrico",
        "ente público competente (especialmente SPU)",
        (
            "por se tratar de recurso hídrico, sob responsabilidade do ente público "
            "competente, especialmente da Secretaria do Patrimônio da União – SPU, "
            "quando caracterizado como bem da União"
        ),
    ),

    # B) Terrenos de marinha e bens da União
    (
        r"\b(TERRENO\s+DE\s+MARINHA|PRAIA|MANGUE|ILHA\s+FEDERAL|"
        r"PATRIMONIO\s+DA\s+UNIAO)\b",
        "Bem da União",
        "Secretaria do Patrimônio da União – SPU",
        (
            "por se tratar de bem da União, sob responsabilidade da "
            "Secretaria do Patrimônio da União – SPU"
        ),
    ),

    # E) Rodovia federal / faixa de domínio federal  (BR antes das estaduais)
    (
        r"\bBR\s*[-–]?\s*\d+|\bRODOVIA\s+FEDERAL\b",
        "Rodovia federal/faixa de domínio",
        "Departamento Nacional de Infraestrutura de Transportes – DNIT",
        (
            "por se tratar de rodovia federal/faixa de domínio, sob responsabilidade do "
            "Departamento Nacional de Infraestrutura de Transportes – DNIT, "
            "salvo concessão ou órgão específico competente"
        ),
    ),

    # G) Faixa de domínio genérica
    (
        r"\bFAIXA\s+DE\s+DOMINIO\b",
        "Faixa de domínio",
        "órgão público ou concessionária competente",
        (
            "por se tratar de faixa de domínio, sob responsabilidade do "
            "órgão público ou concessionária competente"
        ),
    ),

    # H) Área / servidão de uso público  (mais específico antes de "servidão" genérico)
    (
        r"\b(AREA\s+DE\s+SERVIDAO|SERVIDAO\s+PUBLICA|AREA\s+DE\s+SERVIDAO\s+PUBLICA)\b",
        "Servidão pública",
        "Prefeitura Municipal ou órgão público específico competente",
        (
            "por se tratar de área/servidão de interesse público, sob responsabilidade da "
            "Prefeitura Municipal de {{MUNICIPIO}}-{{UF}}, "
            "salvo órgão público específico competente"
        ),
    ),

    # Servidão genérica
    (
        r"\bSERVIDAO\b",
        "Servidão",
        "órgão público ou concessionária competente",
        (
            "por se tratar de servidão, sob responsabilidade do "
            "órgão público ou concessionária competente"
        ),
    ),

    # C) Vias públicas urbanas
    (
        r"\b(RUA|AVENIDA|TRAVESSA|ALAMEDA|BECO|VIELA|PRACA)\b",
        "Via pública urbana",
        "Prefeitura Municipal",
        (
            "por se tratar de via pública urbana, sob responsabilidade da "
            "Prefeitura Municipal de {{MUNICIPIO}}-{{UF}}"
        ),
    ),

    # D) Vias públicas rurais / municipais
    (
        r"\b(ESTRADA|ESTRADA\s+CARROCA|ESTRADA\s+VICINAL|"
        r"CAMINHO\s+PUBLICO|PASSAGEM\s+PUBLICA|VICINAL)\b",
        "Via pública rural",
        "Prefeitura Municipal",
        (
            "por se tratar de via pública rural, sob responsabilidade da "
            "Prefeitura Municipal de {{MUNICIPIO}}-{{UF}}"
        ),
    ),

    # I) Infraestrutura elétrica
    (
        r"\b(LINHA\s+DE\s+TRANSMISSAO|LINHA\s+DE\s+DISTRIBUICAO|"
        r"REDE\s+ELETRICA|ENEL|EQUATORIAL|COELCE|CHESF|ELETROBRAS)\b"
        r"|\bLT\s*[-–]\s*\d+|\bLT\b(?!\w)",
        "Infraestrutura elétrica",
        "concessionária/permissionária de energia elétrica competente",
        (
            "por se tratar de infraestrutura elétrica, sob responsabilidade da "
            "concessionária, permissionária ou entidade responsável competente"
        ),
    ),

    # J) Infraestrutura de saneamento
    (
        r"\b(ADUTORA|REDE\s+DE\s+AGUA|REDE\s+DE\s+ESGOTO|"
        r"CAGECE|SAAE|COMPESA|SABESP|CAERN|CAEMA)\b"
        r"|\bETA\b|\bETE\b",
        "Infraestrutura de saneamento",
        "concessionária ou órgão de saneamento competente",
        (
            "por se tratar de infraestrutura de saneamento, sob responsabilidade da "
            "concessionária, autarquia ou órgão de saneamento competente"
        ),
    ),

    # K) Ferrovia
    (
        r"\b(FERROVIA|LINHA\s+FERREA|VIA\s+FERREA)\b",
        "Ferrovia",
        "órgão ou concessionária ferroviária competente",
        (
            "por se tratar de ferrovia/faixa ferroviária, sob responsabilidade do "
            "órgão ou concessionária ferroviária competente"
        ),
    ),

    # L) Dutos
    (
        r"\b(GASODUTO|OLEODUTO|DUTOVIA|MINERODUTO)\b",
        "Infraestrutura dutoviária",
        "concessionária, operadora ou entidade competente",
        (
            "por se tratar de infraestrutura dutoviária, sob responsabilidade da "
            "concessionária, operadora ou entidade competente"
        ),
    ),

    # Q) Terra indígena
    (
        r"\bTERRA\s+INDIGENA\b",
        "Área pública especial",
        "órgão indigenista federal competente e da União",
        (
            "por se tratar de terra indígena, sob responsabilidade do "
            "órgão indigenista federal competente e da União, quando aplicável"
        ),
    ),

    # R) Território quilombola
    (
        r"\b(QUILOMBOLA|TERRITORIO\s+QUILOMBOLA)\b",
        "Área pública/coletiva especial",
        "órgão público competente",
        (
            "por se tratar de território quilombola, sob responsabilidade do "
            "órgão público competente, quando aplicável"
        ),
    ),

    # P) Área ambiental / restrição ambiental
    (
        r"\b(APP|AREA\s+DE\s+PRESERVACAO\s+PERMANENTE|"
        r"RESERVA\s+LEGAL|UNIDADE\s+DE\s+CONSERVACAO|"
        r"PARQUE|RESERVA\s+AMBIENTAL)\b|\bAPA\b",
        "Área ambiental/restrição ambiental",
        "órgão ambiental competente",
        (
            "por se tratar de área ambiental ou área sujeita a restrição ambiental, "
            "sob responsabilidade do órgão ambiental competente, quando aplicável"
        ),
    ),

    # M) Área pública genérica
    (
        r"\b(AREA\s+PUBLICA|TERRENO\s+PUBLICO|PATRIMONIO\s+PUBLICO)\b",
        "Área pública",
        "ente público competente",
        (
            "por se tratar de área pública, sob responsabilidade do "
            "ente público competente"
        ),
    ),

    # N) Bem municipal (referência explícita à Prefeitura ou Município)
    (
        r"\b(PREFEITURA|MUNICIPIO\b|MUNICIPAL)\b",
        "Bem municipal",
        "Prefeitura Municipal",
        (
            "por se tratar de bem municipal, sob responsabilidade da "
            "Prefeitura Municipal de {{MUNICIPIO}}-{{UF}}"
        ),
    ),

    # O) Bem estadual (referência explícita ao Estado)
    (
        r"\b(GOVERNO\s+DO\s+ESTADO|SECRETARIA\s+ESTADUAL)\b",
        "Bem estadual",
        "Estado ou órgão estadual competente",
        (
            "por se tratar de bem estadual, sob responsabilidade do "
            "Estado de {{UF}} ou órgão estadual competente"
        ),
    ),
]


# ── Função principal ───────────────────────────────────────────────────────────

def classificar_confrontante_dispensa(nome_confrontante, municipio=None, uf=None):
    """Classifica o confrontante e retorna categoria, responsável e texto do bullet.

    Args:
        nome_confrontante: str — nome do confrontante (campo nome_propriedade).
        municipio:         str | None — município do imóvel principal.
        uf:                str | None — UF do imóvel principal.

    Returns:
        dict: {
            "categoria":   str,   # ex.: "Recurso hídrico"
            "responsavel": str,   # ex.: "ente público competente (especialmente SPU)"
            "texto":       str,   # complemento do bullet, sem ponto-e-vírgula final
        }

    Nunca retorna "[órgão responsável]". Fallback = "órgão público ou entidade competente".
    """
    n    = _norm(nome_confrontante)
    _uf  = (uf        or "").strip().upper()
    _mun = (municipio or "").strip()

    # ── F) Rodovias estaduais por sigla UF + número (CE-085, GO-010 …) ─────
    # Avaliado antes da tabela _REGRAS para capturar siglas estaduais específicas.
    for sigla in _SIGLAS_UF:
        if re.search(r"\b%s\s*[-–]?\s*\d+" % re.escape(sigla), n):
            orgao = _ORGAOS_ESTADUAIS.get(sigla, "DER/%s" % sigla)
            texto = (
                "por se tratar de rodovia estadual/faixa de domínio, sob responsabilidade do "
                "%s, salvo concessão ou órgão específico competente" % orgao
            )
            return {
                "categoria":   "Rodovia estadual/faixa de domínio",
                "responsavel": orgao,
                "texto":       _substituir_locais(texto, _mun, _uf),
            }

    # ── F) Rodovia estadual genérica (sem código numérico) ─────────────────
    if re.search(r"\bRODOVIA\s+ESTADUAL\b", n):
        if _uf in _ORGAOS_ESTADUAIS:
            orgao = _ORGAOS_ESTADUAIS[_uf]
        elif _uf:
            orgao = "DER/%s" % _uf
        else:
            orgao = "órgão rodoviário estadual competente"
        texto = (
            "por se tratar de rodovia estadual/faixa de domínio, sob responsabilidade do "
            "%s, salvo concessão ou órgão específico competente" % orgao
        )
        return {
            "categoria":   "Rodovia estadual/faixa de domínio",
            "responsavel": orgao,
            "texto":       _substituir_locais(texto, _mun, _uf),
        }

    # ── Tabela geral de regras ─────────────────────────────────────────────
    for padrao, categoria, responsavel, texto_tpl in _REGRAS:
        if re.search(padrao, n):
            return {
                "categoria":   categoria,
                "responsavel": responsavel,
                "texto":       _substituir_locais(texto_tpl, _mun, _uf),
            }

    # ── Fallback ───────────────────────────────────────────────────────────
    return {
        "categoria":   "Confrontação pública ou não individualizada",
        "responsavel": "órgão público ou entidade competente",
        "texto":       "sob responsabilidade do órgão público ou entidade competente",
    }
