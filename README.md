# 🏪 Projeto Ranking de Lojas

Sistema de ranking de lojas com interface gráfica em Python que importa dados de Excel, armazena em SQLite e exibe análises de faturamento e lucro.

## 📋 Funcionalidades

- ✅ **Importação de Dados**: Lê arquivos Excel e importa para banco de dados SQLite
- 📊 **Ranking de Lojas**: Exibe ranking por faturamento e lucro
- 📈 **Gráfico de Barras**: Visualização visual do faturamento de cada loja
- 🔍 **Detalhes por Loja**: Análise de vendas por vendedor
- 📧 **Envio por E-mail**: Compartilha ranking via e-mail com HTML formatado

## 🚀 Como Usar

### Instalação de Dependências

```bash
pip install pandas openpyxl sqlalchemy yagmail matplotlib
```

### Executar a Aplicação

```bash
python RankingLojas_InterfaceGrafica.py
```

### Arquivos Necessários

Coloque os arquivos Excel na mesma pasta que o script:
- `Loja BH.xlsx`
- `Loja DF.xlsx`
- `Loja Manaus.xlsx`
- `Loja Rio.xlsx`
- `Loja Salvador.xlsx`
- `Loja SP.xlsx`

## 🛠️ Configuração

Edite o arquivo `chave.py` para adicionar suas credenciais do Gmail:

```python
senha = "sua_senha_de_app_gmail"
```

Ajuste também as configurações no início do `RankingLojas_InterfaceGrafica.py`:
- `GMAIL_USER`: Seu e-mail do Gmail
- `EMAIL_TO`: Lista de destinatários
- `EMAIL_CC`: Lista de cópia
- `EMAIL_BCC`: Lista de cópia oculta

## 📁 Estrutura do Projeto

```
Projeto_Ranking_de_Lojas/
├── RankingLojas_InterfaceGrafica.py  # Aplicação principal
├── chave.py                          # Configurações de senha
├── Loja *.xlsx                       # Dados de entrada
└── lojas.db                          # Banco de dados SQLite (gerado automaticamente)
```

## 🎨 Interface Gráfica

A aplicação possui 3 abas principais:

### 🏆 Ranking
- Tabela com ranking de lojas
- Cards com resumo (faturamento total, lucro total, melhor loja)
- Botões para ações rápidas

### 📈 Gráfico
- Gráfico de barras horizontal
- Visualização do faturamento de cada loja
- Valores exibidos nas barras

### 🔍 Detalhes por Loja
- Seletor de loja
- Tabela de vendedores
- Detalhes de vendas e lucro por vendedor

## 💡 Recursos Adicionais

- **Reimportação**: Botão para reimportar dados dos Excel
- **E-mail**: Enviar ranking formatado via e-mail
- **Cores Personalizadas**: Interface com paleta de cores moderna
- **Detecção Automática**: Script encontra automaticamente arquivos Excel na pasta

## 👨‍💼 Autor

EmersonLopes2589

## 📝 Licença

Projeto pessoal - Livre para uso e modificação
