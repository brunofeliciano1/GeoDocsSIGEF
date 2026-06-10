# Uso — GeoDocs SIGEF

## Fluxo básico

### 1. Carregar a camada Serviços 2

Conecte o QGIS ao banco PostgreSQL/PostGIS e carregue a camada **`Serviços 2`**. Essa camada contém os dados cadastrais do imóvel (proprietários, matrícula, CIB, CAR, CCIR, CNS, etc.).

### 2. Informar o ID do imóvel

No painel do plugin, informe o `id` do registro em `Serviços 2` correspondente ao processo que será documentado.

### 3. Usar o PDF SIGEF

Selecione o PDF exportado pelo SIGEF/INCRA para o imóvel. O plugin extrai automaticamente os vértices, azimutes, distâncias e confrontações para gerar o memorial descritivo e a planilha de cálculo.

### 4. Planilha Excel (opcional)

Uma planilha `.xlsx` pode ser fornecida para complementar ou substituir os dados extraídos do PDF.

### 5. Gerar os documentos

Selecione quais documentos deseja gerar e informe a pasta de saída. O plugin gera todos os arquivos `.docx` selecionados com base nos dados informados.

### 6. Confrontantes enriquecidos (camada Confinantes Principal)

Para enriquecer automaticamente os confrontantes do memorial e da planilha, carregue também a camada **`Confinantes Principal`**. O plugin vincula os segmentos do PDF ao nome do confrontante pela relação:

```
Serviços 2.id = Confinantes Principal.codigo
```

### 7. QR Code na Capa

O QR Code da capa é gerado automaticamente a partir do link armazenado no campo **`qrcode_pasta_drive`** da camada `Serviços 2`. Se o campo estiver vazio, o QR Code é omitido.

---

## Campos principais da camada Serviços 2

| Campo | Uso |
|---|---|
| `id` | Identificador do processo — usado para vincular às demais tabelas |
| `nome_propriedade` | Nome do imóvel rural (denominação) |
| `matricula` | Número da matrícula do imóvel |
| `cib` | Código CIB do imóvel (formatado automaticamente) |
| `car` | Código CAR do imóvel (formatado automaticamente) |
| `ccir` | Número do CCIR (formatado automaticamente) |
| `cns` | Código CNS do cartório |
| `nome_cartorio` | Nome do cartório de registro |
| `qrcode_pasta_drive` | Link da pasta do processo — usado para gerar o QR Code da capa |

## Relação com Confinantes Principal

| Campo | Uso |
|---|---|
| `Confinantes Principal.codigo` | Vincula ao `Serviços 2.id` — identifica os confrontantes do imóvel |

---

## Módulos disponíveis

Os documentos gerados estão organizados em módulos independentes na pasta `modulos/`:

- `memorial.py` — Memorial Descritivo
- `planilha_calculo.py` — Planilha de Cálculo
- `capa.py` — Capa com QR Code
- `laudo_tecnico.py` — Laudo Técnico
- `declaracao_respeito_limites.py` — Declaração de Respeito de Limites
- `declaracao_confrontantes.py` — Declaração de Confrontantes Individual
- `declaracao_confrontantes_lote.py` — Declaração de Confrontantes em Lote
- `declaracao_dispensa_anuencia.py` — Declaração de Dispensa de Anuência
- `requerimento_retificacao_area.py` — Requerimento de Retificação de Área
- `requerimento_desmembramento.py` — Requerimento de Desmembramento
- `ferramentas.py` — Conversor/Ferramentas PDF SIGEF
