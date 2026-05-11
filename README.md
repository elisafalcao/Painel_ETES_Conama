# Monitoramento de ETEs — CONAMA 430/2011

Este projeto lê uma planilha de resultados laboratoriais de ETEs, consolida os dados de **afluente** e **efluente**, calcula a **eficiência de remoção de DBO**, avalia a conformidade dos parâmetros selecionados e gera:

- `tabela_conformidade_etes.xlsx`: tabela final com formatação condicional;
- `tabela_conformidade_etes.html`: tabela em HTML;
- gráficos em PNG com visão geral, Top 10 conformidade, Top 10 não conformidade e menores eficiências de DBO.

## Colunas geradas

A tabela principal contém:

- Mês / Ano
- Estação
- pH Efluente
- Temperatura Efluente
- Materiais Sedimentares Efluente
- DBO Afluente
- DBO Efluente
- Eficiência DBO
- Óleos e Graxas Efluente
- Materiais Flutuantes Efluente

Além dessas, o script adiciona colunas de status para cada parâmetro, `Status Geral`, contadores de conformidade e `Índice de Conformidade (%)`.

## Regras adotadas

As regras estão configuradas no dicionário `LIMITES` dentro de `monitoramento_etes.py`:

| Parâmetro | Regra adotada |
|---|---:|
| pH | entre 5 e 9 |
| Temperatura do efluente | até 40 °C |
| Materiais sedimentáveis | até 1 mL/L |
| Eficiência de remoção de DBO | mínimo de 60% |
| Óleos e graxas | até 50 mg/L |
| Materiais flutuantes | ausência |

> Atenção: licenças ambientais, normas estaduais ou enquadramento do corpo receptor podem impor limites mais restritivos. Se necessário, altere os valores em `LIMITES`.

## Cores de status

- Verde: conforme
- Amarelo: atenção, valor próximo do limite ou informação incompleta
- Vermelho: não conforme
- Azul claro: sem dado

## Como usar no GitHub

### 1. Criar o repositório

No GitHub, clique em **New repository**, escolha um nome como `monitoramento-etes-conama430` e crie o repositório.

### 2. Enviar os arquivos

Inclua no repositório:

```text
monitoramento_etes.py
requirements.txt
README.md
ResultadoETEsCESAN_JAN2025python.xlsx
```

Você pode enviar pelo botão **Add file > Upload files** ou pelo terminal:

```bash
git clone https://github.com/SEU-USUARIO/monitoramento-etes-conama430.git
cd monitoramento-etes-conama430
cp /caminho/da/sua/planilha.xlsx ResultadoETEsCESAN_JAN2025python.xlsx
git add .
git commit -m "Adiciona script de monitoramento de ETEs"
git push
```

### 3. Instalar dependências

No terminal, dentro da pasta do projeto:

```bash
python -m venv .venv
source .venv/bin/activate  # Linux/Mac
# .venv\Scripts\activate   # Windows
pip install -r requirements.txt
```

### 4. Executar o script

```bash
python monitoramento_etes.py --entrada ResultadoETEsCESAN_JAN2025python.xlsx --saida resultados
```

### 5. Consultar os resultados

Após a execução, abra:

```text
resultados/tabela_conformidade_etes.xlsx
resultados/tabela_conformidade_etes.html
resultados/graficos/
```

## Publicar resultados no GitHub

Para versionar os resultados gerados:

```bash
git add resultados/
git commit -m "Gera tabela e gráficos de conformidade"
git push
```

## Personalização visual

As cores principais estão no dicionário `PALETA`, usando tons claros de azul, rosa e cinza inspirados na bandeira do Espírito Santo. Ajuste os códigos hexadecimais conforme a identidade visual desejada.
