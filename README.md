# leitor-pdf-python
Leitor de PDF offline em Python com extração avançada de imagens e tabelas (CSV), conversão para Word/TXT e busca visual

---

# 👑 Leitor PDF By SrBozzo v1.0

O **Leitor PDF By SrBozzo** é um "canivete suíço" para manipulação de documentos PDF, desenvolvido inteiramente em Python. O software foi projetado para oferecer uma experiência de leitura fluida e profissional, garantindo 100% de privacidade ao processar todos os dados localmente.

---

## 🚀 Funcionalidades Principais

* **👀 Visualização Inteligente**: Navegação intuitiva por scroll do mouse com troca automática de páginas ao atingir as bordas.
* **🔍 Busca Visual (Ctrl+F)**: Sistema de pesquisa que destaca termos encontrados com marcações vermelhas em tempo real.
* **📘 Conversão Profissional**: Transforma PDFs em arquivos Word (.docx) preservando o layout original ou exporta apenas o texto puro (.txt).
* **📊 Extração de Dados**: Identifica tabelas estruturadas e as exporta para planilhas CSV utilizando o motor de análise do Pandas.
* **🖼️ Extrator de Mídia**: Varre o documento e salva todas as imagens internas em uma pasta separada.
* **🌓 Temas Adaptáveis**: Suporte a Modo Escuro (Dark) e Claro (Light) com ajuste dinâmico de cores.
* **📜 Monitor de Atividades (Logs)**: Registro em tempo real que mostra o tempo exato de execução de cada tarefa em milissegundos.

---

## 🛠️ Tecnologias e Motores

O projeto utiliza bibliotecas de alta performance para garantir rapidez e precisão:

* **CustomTkinter**: Interface gráfica moderna e responsiva.
* **PyMuPDF (fitz)**: Motor de renderização e extração de dados ultrarrápido.
* **Pandas**: Estruturação de dados para extração de tabelas.
* **pdf2docx**: Reconstrução de documentos PDF para o formato Word.
* **Pillow (PIL)**: Processamento de imagens para exibição em tela.

---

## 💻 Como Editar e Contribuir

Se você deseja modificar o código ou adicionar novas funções, siga os passos abaixo:

### 1. Pré-requisitos
Certifique-se de ter o **Python 3.x** instalado.

### 2. Instalação das Dependências
Abra o terminal na pasta do projeto e instale as bibliotecas necessárias:
```bash
pip install customtkinter Pillow PyMuPDF pandas pdf2docx
```

### 3. Executando o Projeto
Para rodar a interface de desenvolvimento, execute o arquivo principal:
```Bash
python app_desktop.py
```

### 4. Estrutura de Arquivos Essenciais
Para que o software funcione corretamente, mantenha estes arquivos na mesma pasta:

app_desktop.py: O "cérebro" da interface e lógica do app.
leitor_pdf.py: O motor de extração de texto.
logo.ico: O ícone oficial do sistema.

---

## 💾 Versão Portátil (.exe)
Para usuários que não possuem Python instalado, acesse a aba Releases deste repositório e baixe o executável portátil. Ele roda sem necessidade de instalação ou permissões de administrador.

---

# 👑 Leitor PDF By [SrBozzo](https://github.com/SrBozzo)
