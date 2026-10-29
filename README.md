# 📚 Além da Capa: Analytics de Leitura

![Power Bi](https://img.shields.io/badge/power_bi-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Status](https://img.shields.io/badge/Status-Em_Andamento-yellow?style=for-the-badge)

> **"Do Papel ao Pixel":** Transformando um diário de leitura manuscrito em um Dashboard de Business Intelligence.

## 📖 Sobre o Projeto

Este projeto nasceu de um case real e pessoal. Minha esposa mantinha um registro detalhado de suas leituras em um caderno físico, anotando dados como início e fim da leitura, avaliação, autores e observações.

Embora o registro fosse disciplinado, extrair *insights* era impossível. Perguntas como *"Qual meu ritmo de leitura anual?"* ou *"Qual gênero eu mais abandono?"* não tinham resposta imediata.

**O Objetivo:** Digitalizar esses dados, estruturar um Data Warehouse pessoal e criar uma visualização interativa para descobrir os padrões de hábito de leitura.

---

## 📊 Visão Geral do Dashboard

Aqui está o resultado final da análise:

| Capa (Visão Geral) | Detalhe por Autor |
| :---: | :---: |
| ![Dashboard Overview](assets/resumo.png) | ![Detalhes](assets/tesouros-lidos.png) |
*(Screenshots do projeto - Substitua pelas suas imagens reais)*

### Principais Insights (KPIs):
* **Volume:** Total de livros lidos e páginas viradas.
* **Velocidade:** Média de páginas lidas por dia e tempo médio de conclusão de um livro.
* **Qualidade:** Média de avaliações (1 a 5 ⭐) e taxa de livros "Favoritos".
* **Preferências:** Autores mais lidos e gêneros predominantes.

---

## 🛠️ Engenharia de Dados & Solução Técnica

Para sair do caderno e chegar no dashboard, o projeto seguiu um fluxo rigoroso de BI:

### 1. Arquitetura e Normalização (Data Design)
A grande sacada técnica deste projeto não foi o tratamento posterior, mas sim o **Design Relacional** definido na etapa de coleta.

Diferente de manter uma planilha única com textos repetidos ("Flat File"), optei por estruturar os dados seguindo princípios de **Normalização de Banco de Dados**:

* **Separação de Entidades:** Criação de tabelas dimensão dedicadas (`tb_autor`, `tb_genero`, `tb_editora`) para garantir integridade única dos dados.
* **Uso de Chaves Estrangeiras (FK):** A tabela principal (`tb_livros`) atua como a tabela fato centralizadora. Ela não armazena os nomes textuais, mas sim os IDs de referência (`ID_Autor`, `ID_Genero`), otimizando a performance e a organização.
* **Power Query:** Com a estrutura bem definida na origem, o Power Query atuou de forma leve, apenas validando tipos de dados e carregando o modelo relacional pronto.

### 2. Modelagem de Dados (Star Schema)
Diferente de uma simples "tabela dinâmica", optei por criar um modelo dimensional robusto (**Star Schema**) para garantir performance e escalabilidade:

* **f_Leituras (Fato):** Tabela central com as transações (cada livro lido), contendo as métricas (páginas, dias, nota).
* **d_Livros (Dimensão):** Atributos do livro (Título, ISBN, Gênero).
* **d_Autor (Dimensão):** Dados normalizados dos autores.
* **d_Calendario (Dimensão):** Tabela auxiliar para inteligência de tempo (DAX Time Intelligence).

### 3. Visualização & UX
Utilizei um tema personalizado **"Classic Library"**, focado em tons de papel antigo e azul tinta para evocar a sensação de leitura, evitando a fadiga visual.

---

## ✅ Checklist de Desenvolvimento & Roadmap

### 🏗️ Engenharia e Modelagem (Backend)
- [x] **Arquitetura de Dados:** Definição das tabelas e relacionamentos (Chaves Primárias e Estrangeiras).
- [x] **Normalização:** Criação das tabelas dimensão (`tb_autor`, `tb_genero`) para evitar redundância na tabela fato (`tb_livros`).
- [x] **Modelagem:** Implementação do modelo **Star Schema** no Power BI.
- [x] **DAX:** Criação das medidas calculadas para KPIs de leitura e inteligência de tempo.

### 🎨 Design e Experiência (Frontend)
- [x] **Identidade Visual:** Definição do tema "Classic Library" (Paleta de Cores e Tipografia).
- [x] **Assets:** Criação de backgrounds e ícones personalizados.
- [x] **Storytelling:** Construção dos painéis de Visão Geral e Detalhamento por Autor.

### 🚀 Melhorias Futuras (Backlog)
- [ ] **UX Avançada:** Criar *Tooltips* (Dicas de Ferramenta) personalizadas que exibem a capa do livro ao passar o mouse sobre o título.
- [ ] **Enriquecimento de Dados:** Investigar integração com API do Google Books para buscar capas e sinopses automaticamente.

### 🛠️ Próximos passos
- [ ] **visão por gênero:**
- [ ] **visão por editora:**
- [ ] **visão por personagens:**
- [ ] **visão por livro do mês e favoritos:**

- [ ] **%lido por autor:**
- [ ] **%lido por gênero:**
- [ ] **%lido por editora:**
- [ ] **%lido por formato:**

- [ ] **tempo de leitura x avaliação:**
- [ ] **top 10 livros mais longos:**
- [ ] **top 10 livros mais curtos:**

- [ ] **personagem de acordo com o tema da leitura ou com o autor/gênero mais lido:**

- [ ] **projeção de leitura:**


## 📂 Estrutura do Repositório

```text
📂 alem-da-capa-analytics
├── 📂 assets                 # Screenshots para documentação
├── 📂 docs                   # Design system e documentação de regras
├── 📂 src                    # Código Fonte
│   ├── 📂 data-source        # Bases normalizadas (CSV)
│   ├── 📂 images             # Ícones e Backgrounds usados no PBI
│   └── dashboard-habitos-leitura.pbix  # Arquivo principal
└── README.md