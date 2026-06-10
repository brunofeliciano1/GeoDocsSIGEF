# Instalação — GeoDocs SIGEF

## Requisitos mínimos

- **QGIS 3.40** ou superior
- Python do QGIS (já incluído na instalação do QGIS)
- Dependências Python listadas em [DEPENDENCIES.md](DEPENDENCIES.md)

---

## Instalar a partir do ZIP

1. No QGIS, acesse o menu **Complementos → Gerenciar e Instalar Complementos**
2. Clique na aba **Instalar a partir de ZIP**
3. Selecione o arquivo `GeoDocsSIGEF.zip`
4. Clique em **Instalar Plugin**

> **Atenção:** o ZIP deve manter a estrutura de pastas com o nome do plugin como pasta raiz:
> ```
> GeoDocsSIGEF/
>     metadata.txt
>     __init__.py
>     geodocssigef.py
>     docx_utils.py
>     modulos/
>     models/
>     ...
> ```
> O QGIS exige que `GeoDocsSIGEF/metadata.txt` e `GeoDocsSIGEF/__init__.py` existam dentro do ZIP para reconhecer o plugin.

---

## Após a instalação

1. Ative o plugin em **Complementos → Gerenciar e Instalar Complementos → Instalados**
2. O ícone do GeoDocs SIGEF aparecerá na barra de ferramentas do QGIS

---

## Instalar dependências Python

Algumas bibliotecas precisam ser instaladas manualmente no Python do QGIS.
Consulte [DEPENDENCIES.md](DEPENDENCIES.md) para instruções completas.
