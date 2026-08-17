# ConversorImagens

Conversor de imagens com interface gráfica (PyQt5 + Pillow).

Localiza imagens em uma pasta (qualquer formato), permite escolher o tipo de origem e o formato de destino, salva cópias no mesmo local das originais e exporta Excel com caminhos clicáveis.

## Nome do projeto

**ConversorImagens**

## Caminho

```
C:\Users\Tiago Holanda\Downloads\SCRIPT
```

## Arquivos principais

| Arquivo | Função |
|---------|--------|
| `conversor_imagens.py` | Programa com a tela |
| `rodar.bat` | Atalho para abrir o programa |
| `requirements.txt` | Dependências (Pillow, PyQt5, openpyxl) |
| `SISTEMA\` | Pasta de teste com imagens misturadas |

## Como ativar / abrir

### Opção 1 — Duplo clique

1. Abra: `C:\Users\Tiago Holanda\Downloads\SCRIPT`
2. Duplo clique em `rodar.bat`

### Opção 2 — Terminal

```powershell
cd "C:\Users\Tiago Holanda\Downloads\SCRIPT"
python conversor_imagens.py
```

### Se faltar biblioteca

```powershell
cd "C:\Users\Tiago Holanda\Downloads\SCRIPT"
pip install -r requirements.txt
```

## Uso

1. Clique em **Procurar...** e selecione a pasta
2. Em **Converter de**, escolha o tipo (ICO, PNG, JPEG, Todos…)
3. Em **Para**, escolha o formato de saída
4. Marque **Incluir subpastas** se quiser
5. **1. Buscar imagens**
6. **2. Converter e salvar cópias**
7. **3. Salvar Excel (caminhos)** — links clicáveis no Excel

## Pasta de teste

```
C:\Users\Tiago Holanda\Downloads\SCRIPT\SISTEMA
```

## Formatos

**Entrada:** PNG, JPG, JPEG, BMP, GIF, TIFF, WEBP, ICO e outros comuns.  
**Saída:** PNG, JPEG, WEBP, BMP, TIFF, GIF, ICO.

Sim: **ICO → PNG** funciona. O original não é apagado.
