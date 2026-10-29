# 📊 WFM Performance Dashboard — Intradia & Aderência

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![SQL Server](https://img.shields.io/badge/T--SQL-CC292B?style=for-the-badge&logo=microsoft-sql-server&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-0073EC?style=for-the-badge&logo=visual-basic&logoColor=white)
![Status](https://img.shields.io/badge/Status-Conclu%C3%ADdo-brightgreen?style=for-the-badge)

> Solution de Workforce Management (WFM) projetada para auditoria operacional de Call Center em janelas intradia de 30 minutos, integrando análise de capacidade, curvas de oferta, desvios de TMA e métricas de Aderência em layout executivo.

---

## 📌 Visão Geral da Solução

O dashboard consolida dados operacionais de atendimento telefônico, mitigando distorções volumétricas e permitindo que a gestão tome decisões táticas com base em dados auditáveis. A arquitetura visual adota a paleta corporativa **Laranja (Alta Demanda/Alerta)**, **Grafite (Baixa Demanda/Fundo)** e **Branco (Contraste/Sinalização)**.

### 🧩 Módulos do Sistema

1. **OnePage (Executive Summary):** Visão holística de tráfego, nível de serviço, taxa de ocupação e entregas de escala em um único painel de alta densidade informativa.
2. **Volumetria:** Análise detalhada de chamadas Oferecidas, Atendidas, Abandonadas, Efetuadas e Transferidas, desdobradas por turno (Manhã, Tarde, Noite) e janelas de demanda.
3. **TMA & TME (Intradia e Aderência):**
   * **Intradia:** Evolução temporal do Tempo Médio de Atendimento (TMA) vs. Planejado por turno e por intervalo de 30 minutos.
   * **Aderência:** Análise de dispersão relativa indicando **Sobrecarga de TMA** nos $N$ maiores picos e **Alívio de TMA** nas janelas de menor pressão.

---

## 📐 Métricas e Regras de Negócio

Para garantir rigor matemático e evitar distorções de interpretação por usuários não técnicos, o painel aplica as seguintes convenções:

| Métrica | Conceito / Fórmula | Aplicação no Dashboard |
| :--- | :--- | :--- |
| **PCA** | $\frac{\text{Atendidas}}{\text{Oferecidas}}$ | Mede a eficiência de conversão da fila. |
| **Nível de Serviço (NS20)** | $\frac{\text{Atendidas até 20s}}{\text{Oferecidas}}$ | Métrica primária de experiência do cliente. |
| **% Atendidas em 20s** | $\frac{\text{Atendidas até 20s}}{\text{Atendidas Totais}}$ | Participação do atendimento rápido dentro do universo de chamadas concluídas. |
| **Sobrecarga de TMA** | $\left( \frac{\text{TMA Top 5}}{\text{TMA Consolidado}} \right) - 1$ | Quantifica o estresse operacional nos intervalos de pico. |
| **Alívio de TMA** | $1 - \left( \frac{\text{TMA Vales}}{\text{TMA Consolidado}} \right)$ | Quantifica o ganho de eficiência e velocidade nas janelas de baixa demanda. |

> **Nota de Tratamento de Dados (SQL):** As queries de extração utilizam CTEs para deduplicação estrita de `agentid` por intervalo de 30 minutos, impedindo a inflação artificial de tempo por multiplicadores de filas/skills.

---

## 🖼️ Mapeamento de Telas

### 1. Painel OnePage
- Visão executiva em grid estruturado com cards de KPI e velocímetro de Nível de Serviço.
- Gráfico de colunas indicando a proporção de **Intervalos Entregues vs. Não Entregues**.

### 2. Volumetria & Comparativo
- Segmentação de volumetria por turnos e botão seletor de indicadores.
- Gráfico de rosca dinâmico para cálculo de **Concentração de Volume em Alta Demanda**.

### 3. Módulo de TMA & Aderência
- Gráfico de barras combinadas indicando desvios acima/abaixo do planejamento.
- Painel lateral com anéis de **Sobrecarga** e **Alívio de TMA** contendo redação direta para leitura executiva.

---

## 💻 Estrutura de Código

### SQL (T-SQL Engine)
Os scripts localizados em `src/sql/` aplicam o padrão de agrupamento prévio em memória:

```sql
with CteAgenteIntervalo as (
    select
        c.Data
        , convert(time(0), dbo.fn_IntervaloPor(c.Intervalo, 30)) as Intervalo
        , c.agentid
        , max(c.TempoDisponivel) as TempoDisponivel
    from dbo.TB_AtendimentoRaw as c with(nolock)
    where 1=1
        and c.Data = @DataInicio
    group by 
        c.Data
        , convert(time(0), dbo.fn_IntervaloPor(c.Intervalo, 30))
        , c.agentid
)
select 
    Intervalo
    , dbo.fn_divisao(sum(TempoDisponivel), count(distinct agentid)) as TempoDisponivelMedio
from CteAgenteIntervalo
group by Intervalo;