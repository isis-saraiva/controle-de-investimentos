# controle-de-investimentos
## Ferramenta de controle de investimentos com Excel
### 1. Quais perguntas a ferramenta responde?
A ferramenta de controle de investimentos responde a várias perguntas estratégicas de planejamento financeiro, divididas em blocos na aba **APP**:
* **Quanto devo investir por mês?** Aparece na célula `D14`, calculada a partir de 30% do salário configurado (`D11` = `=salario*30%`).
* **Quanto de patrimônio acumulado?** (Projeção do montante futuro) Aparece na célula `D17`, calculada com base na taxa mensal, período em anos e aporte mensal.
* **Quantos dividendos mensais?** Aparece na célula `D18`, calculada aplicando o rendimento da carteira sobre o patrimônio acumulado (`=patrimonio*rendimento_carteira`).
* **Cenários de evolução (Quanto em 2, 5, 10, 20 e 30 anos?):** Aparecem na coluna **C** (para o montante acumulado em cada prazo de 2 a 30 anos) e na coluna **D** (para os dividendos mensais correspondentes a cada prazo nas linhas 21 a 25).
* **Qual o valor a ser investido por mês por tipo de FII (Alocação de ativos)?** Aparece na coluna **D** (linhas 33 a 38), detalhando os valores alocados em Papel, Tijolo, Híbridos, FOFs, Desenvolvimento e Hotelarias com base no perfil escolhido.
---

### 2. Como o `VF` e o `PROCV` entram nos cálculos
* **`VF` (Valor Futuro / `FV`):** É utilizado nas fórmulas de projeção de juros compostos para calcular o montante acumulado de patrimônio ao longo do tempo.
* Na célula `D17`: `=FV(taxa_mensal, qtd_anos*12, aporte*(-1))` calcula o valor futuro para o prazo total de anos definido.
* Nas linhas 21 a 25 (tabela de cenários): `=FV(taxa_mensal, A21*12, aporte*(-1))` calcula o valor acumulado para prazos específicos (2, 5, 10, 20 e 30 anos), utilizando o aporte mensal negativo (padrão da função financeira do Excel para saídas de caixa/investimentos).

* **`PROCV` (`VLOOKUP`):** É utilizado para buscar os percentuais de alocação de acordo com o perfil de investidor selecionado e o tipo de FII.
* Nas células da coluna `C` (linhas 33 a 38): `=VLOOKUP($C$29&"-"&B33, Chave!$A$2:$D$20, 4, FALSE)`. Ele concatena o perfil selecionado (ex: "Moderado") com o tipo de FII (ex: "PAPEL") para formar uma chave composta (`Moderado-PAPEL`), busca essa chave na aba **Chave** e retorna o percentual correspondente (4ª coluna da tabela de referência).
---

### 3. Quais os intervalos nomeados?
Na planilha, foram definidos os seguintes intervalos/células nomeadas (com seus respectivos locais na aba **APP**):

* `aporte` $\rightarrow$ `APP!$D$14` (Valor do aporte mensal)
* `patrimonio` $\rightarrow$ `APP!$D$17` (Patrimônio acumulado projetado)
* `qtd_anos` $\rightarrow$ `APP!$D$15` (Prazo em anos)
* `rendimento_carteira` $\rightarrow$ `APP!$D$10` (Taxa de rendimento dos dividendos/carteira, ex: 0,6% ao mês)
* `salario` $\rightarrow$ `APP!$D$9` (Salário base)
* `sugestao_investimento` $\rightarrow$ `APP!$D$11` (Sugestão de 30% do salário)
* `taxa_mensal` $\rightarrow$ `APP!$D$16` (Taxa de juros/crescimento mensal dos investimentos)
---

### 4. Os percentuais de cada perfil
Os percentuais de alocação por tipo de Fundo de Investimento Imobiliário (FII) variam conforme o perfil de risco selecionado na célula `C29` ( Conservador, Moderado ou Agressivo). Eles vieram diretamente da tabela de referência localizada na aba **Chave** com base nos dados enviados pelo instrutor do curso (`Chave!$A$2:$D$20`):

* **Perfil Conservador:**
* Papel: **30%**
* Tijolo: **50%**
* Híbridos: **10%**
* FOFs: **10%**
* Desenvolvimento: **0%**
* Hotelarias: **0%**


* **Perfil Moderado:**
* Papel: **32%**
* Tijolo: **40%**
* Híbridos: **8%**
* FOFs: **10%**
* Desenvolvimento: **10%**
* Hotelarias: **0%**


* **Perfil Agressivo:**
* Papel: **50%**
* Tijolo: **10%**
* Híbridos: **5%**
* FOFs: **5%**
* Desenvolvimento: **20%**
* Hotelarias: **10%**
Os valores são uma simulação e podem ser alterados para atender a realidade do usuário
---

### 5. O que mudou em relação à ferramenta do Instrutor
Ajuste na formatação e cores da aba APP.

<img width="832" height="738" alt="APP_teste02" src="https://github.com/user-attachments/assets/c89540ba-3c67-42e3-a529-3f5fcd31d041" />
<img width="837" height="770" alt="APP_teste01" src="https://github.com/user-attachments/assets/157e709e-9cc0-4794-b718-b72e232d10bb" />
