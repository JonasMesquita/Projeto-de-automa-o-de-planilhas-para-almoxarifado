# 🧾 Sistema de Controle de Estoque – MechaMachines
**Versão 3.0 | Desenvolvido em Python (Tkinter + OpenPyXL + ReportLab)**

---

## 📘 Descrição

O **Sistema de Controle de Estoque MechaMachines** foi desenvolvido para auxiliar o gerenciamento de materiais em almoxarifados, obras e depósitos.  
Ele permite **registrar entradas e saídas**, **gerar relatórios em PDF** e **acompanhar o saldo de estoque em tempo real** através de uma interface simples e intuitiva.  

O programa utiliza **Excel (.xlsx)** como base de dados, garantindo portabilidade e facilidade de acesso às informações.  

---

## ⚙️ Funcionalidades Principais

✅ Registro de **entradas** e **saídas** de produtos  
✅ Cálculo automático de **saldo de estoque**  
✅ Geração de **relatórios completos em PDF**  
✅ **Painel visual de produtos** com alerta de estoque baixo  
✅ Botão para **excluir registros** de forma segura  
✅ Escolha do **local do arquivo Excel** no primeiro uso  
✅ Total de **entradas e saídas por período**  
✅ Interface 100% em **Tkinter**, leve e compatível com Windows  

---

## 🪟 Interface

A tela principal apresenta:

| Função | Descrição |
|--------|------------|
| **Registrar Entrada** | Adiciona novos produtos ou atualiza a quantidade de um produto existente. |
| **Registrar Saída** | Registra materiais que saíram do estoque, com data e destino. |
| **Excluir Registro** | Remove registros incorretos (com confirmação). |
| **Gerar Relatório PDF** | Gera relatório detalhado com entradas, saídas e totais. |
| **Atualizar Painel** | Atualiza os dados do painel e verifica alertas de estoque baixo. |

Produtos com estoque **abaixo do limite mínimo (10 unidades)** são destacados em **vermelho**.  

---

## 📦 Estrutura de Arquivos

```
📂 Sistema_Estoque/
│
├── almoxarifado_v3.0.py        # Código principal
├── estoque.xlsx                 # Planilha de dados
├── logo.ico                     # Ícone do executável
├── Instruções de Uso.txt        # Manual detalhado
├── README.md                    # Este arquivo
└── /relatorios/                 # Pasta onde os PDFs são salvos
```

---

## 🧠 Tecnologias Utilizadas

- **Python 3.10+**
- **Tkinter** → Interface gráfica  
- **OpenPyXL** → Manipulação de planilhas Excel  
- **ReportLab** → Geração de PDFs  
- **PyInstaller** → Criação do executável  

---

## 🧾 Relatórios PDF

Os relatórios incluem:

- Total de **entradas e saídas**
- **Saldo de estoque** atual
- **Produtos com estoque baixo**
- Data e hora da geração  
- Cabeçalho com logo e informações do sistema  

Exemplo de nome do arquivo:
```
relatorio_estoque_02112025_1612.pdf
```

---

## ⚠️ Requisitos

Se usar o **.exe** → Não é necessário Python instalado.  
Se rodar o código diretamente, instale os pacotes com:

```
pip install openpyxl reportlab
```

---

## 🧰 Compilação em Executável (opcional)

Para gerar o `.exe` com ícone e sem console:

```
pyinstaller --onefile --noconsole --icon=logo.ico almoxarifado_v3.0.py
```

O executável aparecerá dentro da pasta `dist/`.

---

## 💡 Dicas

- Faça **backup periódico** da planilha `estoque.xlsx`.  
- Gere **relatórios mensais** para histórico.  
- Padronize os nomes dos produtos (ex: “Tinta 18L Azul”).  
- Não altere manualmente as fórmulas do Excel.  

---

## 🧑‍💻 Autor

**Desenvolvido por:** Jonas Mesquita  
**Projeto:** Sistema de Controle de Estoque – MechaMachines  
**Linguagem:** Python  
**Versão:** 3.0 (2025)  

---

## 📜 Licença

Este projeto é distribuído para uso pessoal ou interno.  
Modificações são permitidas, desde que mantidos os créditos ao autor.
