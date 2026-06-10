# Dependências — GeoDocs SIGEF

## Dependências Python

| Biblioteca | Uso no plugin | Instalação necessária? |
|---|---|---|
| **python-docx** | Geração e manipulação de documentos `.docx` | Sim |
| **openpyxl** | Leitura e geração de planilhas `.xlsx` | Sim |
| **qrcode** | Geração do QR Code na capa | Sim |
| **Pillow (PIL)** | Processamento de imagem do QR Code | Sim (geralmente instalado com qrcode) |
| **pdfplumber** | Leitura e extração de texto de PDFs SIGEF | Sim |
| **psycopg2** | Conexão com banco PostgreSQL/PostGIS | Opcional (já disponível via QGIS na maioria dos casos) |
| **PyQGIS** | API do QGIS (camadas, feições, conexões) | Não — já incluso no QGIS |

---

## Como instalar no Python do QGIS (Windows)

Abra o **OSGeo4W Shell** (instalado junto com o QGIS) ou localize o executável Python do QGIS e execute:

```
python -m pip install python-docx openpyxl "qrcode[pil]" pdfplumber
```

No Windows com QGIS 3.40, o Python geralmente está em:
```
C:\Program Files\QGIS 3.40.x\apps\Python312\python.exe
```

Exemplo direto:
```
"C:\Program Files\QGIS 3.40.4\apps\Python312\python.exe" -m pip install python-docx openpyxl "qrcode[pil]" pdfplumber
```

---

## Observações

- `qrcode[pil]` instala o `qrcode` com suporte a imagem via Pillow — use sempre essa forma.
- `pdfplumber` depende de `pdfminer.six`, instalado automaticamente como dependência.
- `psycopg2` normalmente já está disponível no QGIS instalado com o pacote oficial Windows. Se não estiver: `pip install psycopg2-binary`.

PostgreSQL/PostGIS access

GeoDocs SIGEF primarily reads data from QGIS layers loaded in the current project. When additional related records are required, such as owners, representatives and relationship tables, the plugin may execute SQL queries using the PostgreSQL/PostGIS connection already configured in the QGIS layer.

The plugin does not store database credentials, passwords, hosts, users or connection strings in the source code. Connection parameters are obtained dynamically from the QGIS data source URI of the loaded layer.
