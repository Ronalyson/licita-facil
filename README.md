# Licita Fácil

Esse projeto nasceu pra resolver um problema bem simples, mas que tomava tempo demais:

ficar ajustando planilha manualmente (principalmente pra licitação), mudando valores de coluna inteira, arredondando, revisando… toda vez.

Antes disso eu já tinha feito scripts em VBA, mas nem todo mundo sabe mexer com código no Excel.  
Então a ideia aqui foi: transformar isso num programa simples que qualquer pessoa consiga usar.

---

## O que o app faz

Você seleciona uma planilha `.xlsx`, escolhe uma coluna e aplica uma regra:

- aumentar valores
- diminuir valores

com porcentagem personalizada.

No caso de diminuição, dá pra escolher:
- arredondar pra cima
- arredondar pra baixo
- ou não arredondar

Ele só altera células que fazem sentido:
- números
- que não estejam vazias
- e que não sejam fórmulas

O resto ele ignora.

No final, salva em um novo arquivo (não sobrescreve o original).

---

## De onde veio a lógica

Esse projeto é basicamente a evolução de dois scripts VBA que eu já usava:

- `column-percentage-increase.bas`
- `column-percentage-decrease.bas`

A lógica continua a mesma:

- aumento → `valor * (1 + porcentagem/100)`
- diminuição → `valor * (1 - porcentagem/100)`

Só que agora com interface e sem precisar abrir editor VBA.

---

## Stack

Nada pesado, a ideia foi manter simples e rápido:

- Python
- openpyxl (pra mexer nas planilhas)
- customtkinter (pra interface)

Sem Electron, sem browser, sem frescura.

---

## Estrutura

```text
app/
  ui/          → interface
  services/    → processamento da planilha
  utils/       → validações
  models/      → estrutura de dados
vba/           → scripts originais
run.py         → ponto de entrada