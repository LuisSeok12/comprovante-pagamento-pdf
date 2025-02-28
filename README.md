# Comprovante Pagamento PDF

Este projeto contém um script Python que processa comprovantes de pagamento em PDF e atualiza um arquivo Excel com o status de cada transação.

## Funcionalidades

- **Seleção de arquivo Excel:** O usuário escolhe o arquivo com os dados das transações.
- **Seleção da pasta base dos PDFs:** O script permite escolher a pasta que contém os comprovantes organizados por ano e mês.
- **Busca e extração de páginas:** O script identifica o PDF correspondente à data da transação e busca uma página específica usando o número da fatura ou o valor do pagamento.
- **Exportação de página:** Se a transação for encontrada, a página é extraída e salva em uma subpasta chamada "Notas".
- **Atualização do Excel:** A coluna "Encontrado" do Excel é atualizada para indicar se o comprovante foi localizado.

## Requisitos

- Python 3.x
- [pandas](https://pandas.pydata.org/)
- [PyPDF2](https://pypi.org/project/PyPDF2/)
- [tkinter](https://docs.python.org/3/library/tkinter.html) (geralmente já incluído com o Python)
- [openpyxl](https://pypi.org/project/openpyxl/)

## Uso

1. Clone o repositório para sua máquina:
   ```bash
   git clone https://github.com/seu-usuario/comprovante-pagamento-pdf.git
