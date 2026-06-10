# GeoDocs SIGEF

**EN:** QGIS plugin for automated generation of technical documents for rural property georreferencing processes (SIGEF/INCRA, Brazil).

**PT:** Plugin QGIS para geração automatizada de peças técnicas em processos de georreferenciamento de imóveis rurais (SIGEF/INCRA).

---

## Objetivo

O GeoDocs SIGEF integra dados de PDFs do SIGEF/INCRA, planilhas Excel, camadas QGIS e banco de dados PostgreSQL/PostGIS para gerar os documentos exigidos na certificação de imóveis rurais — com formatação padronizada, nomenclatura técnica correta e preenchimento automático a partir dos dados do processo.

---

## Funcionalidades

| Documento | Descrição |
|---|---|
| **Memorial Descritivo** | Gerado a partir do PDF SIGEF, com descrição de perímetro, confrontantes enriquecidos e assinaturas |
| **Planilha de Cálculo** | Planilha analítica com vértices, azimutes, distâncias e confrontações |
| **Capa** | Folha de rosto com dados do imóvel e QR Code gerado a partir do link da pasta do processo |
| **Laudo Técnico** | Laudo com qualificação completa de proprietários e representantes |
| **Declaração de Respeito de Limites** | Declaração com classificação automática do tipo de imóvel de cada confrontante |
| **Declaração de Confrontantes Individual** | Declaração individual por confrontante do processo |
| **Declaração de Confrontantes em Lote** | Geração em lote de declarações para todos os confrontantes |
| **Declaração de Dispensa de Anuência** | Declaração de dispensa para confrontantes enquadrados nos critérios legais |
| **Requerimento de Retificação de Área** | Requerimento com área por extenso e dados do imóvel |
| **Requerimento de Desmembramento** | Requerimento de desmembramento com parcelas e áreas |
| **Ferramentas PDF SIGEF** | Leitura direta de PDF SIGEF para extração de vértices, memorial narrativo e cartas de anuência |

---

## Requisitos de dados

O plugin trabalha com:

- **PDFs do SIGEF/INCRA** — fonte principal dos vértices e dados do processo
- **Camada QGIS `Serviços 2`** — vinculada a banco PostgreSQL/PostGIS, contém os dados cadastrais do imóvel
- **Camada QGIS `Confinantes Principal`** — vinculada ao mesmo banco, fornece os confrontantes enriquecidos pela relação `Serviços 2.id = Confinantes Principal.codigo`
- **Planilha Excel** (opcional) — pode complementar ou substituir dados do PDF

---

## Autor

Bruno Feliciano — brunofelicianodelima036@gmail.com
