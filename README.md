# Consolidador Inteligente de Planilhas de Algodão

## 📖 Sobre o Projeto

Este é um aplicativo de desktop completo, desenvolvido em Python, para automatizar e gerenciar a consolidação de relatórios de análise de algodão. A ferramenta foi projetada para otimizar um fluxo de trabalho empresarial real, substituindo processos manuais, demorados e suscetíveis a erros pela automação inteligente.

O sistema não apenas agrega dados, mas também os valida, limpa e oferece funcionalidades de atualização, tudo através de uma interface gráfica moderna, segura e intuitiva.

![Screenshot da Aplicação](https://i.imgur.com/JciaiVr.png)
![Screenshot da Aplicação](https://i.imgur.com/3xPU8yQ.png)
![Screenshot da Aplicação](https://i.imgur.com/btfTgZo.png)
![Screenshot da Aplicação](https://i.imgur.com/kpsteN2.png)

---

## ✨ Funcionalidades Avançadas

* **Interface Gráfica Profissional:** Desenvolvida com a biblioteca **CustomTkinter**, oferecendo um visual moderno e uma experiência de usuário aprimorada.
* **Sistema de Login:** Garante que apenas usuários autorizados tenham acesso à ferramenta.
* **Temas Customizáveis:** O usuário pode alternar entre os temas *Light* (claro) e *Dark* (escuro) diretamente na interface.
* **Adição e Atualização Inteligente:**
    * **Adicionar Dados:** O sistema verifica se um lançamento já existe (baseado em `Número` e `Vendedor`) e ignora duplicatas, prevenindo a inserção de dados repetidos.
    * **Atualizar Lançamento:** Permite ao usuário selecionar um lançamento existente e substituí-lo completamente com dados de uma nova planilha, ideal para correções e atualizações.
* **Seleção e Ordenação Interativa:** O usuário pode selecionar múltiplos arquivos de origem e reordená-los em uma lista interativa (arrastando para cima/baixo) para definir a ordem exata da consolidação.
* **Limpeza e Validação de Dados (Data Cleaning):**
    * O "motor" de processamento padroniza nomes de colunas (traduzindo de Português para Inglês).
    * Limpa colunas numéricas, removendo caracteres de texto (`R$`, letras) e tratando corretamente separadores decimais (`,`) e de milhar (`.`) para garantir a integridade dos dados para cálculos.
* **Segurança e Integridade:**
    * **Backups Automáticos:** A cada execução, um backup da planilha mestra é criado com data e hora.
    * **Preservação de Formato:** A escrita dos dados no Excel é feita de forma a preservar a formatação original da planilha, incluindo tabelas nomeadas, cores e outras características.
* **Aplicação Responsiva:** A lógica de processamento de dados roda em uma *thread* separada, garantindo que a interface gráfica nunca "congele", mesmo durante operações longas, e exibindo um log de progresso em tempo real.

---

## 🛠️ Tecnologias Utilizadas

* **Linguagem:** Python
* **Interface Gráfica (GUI):** CustomTkinter, Tkinter
* **Manipulação de Dados:** Pandas
* **Interação com Excel:** Openpyxl
* **Utilitários:** Pillow (imagens), python-dateutil (datas)
* **Concorrência:** Threading, Queue (para responsividade da UI)

---

## 🚀 Como Executar o Projeto

1.  **Clone o repositório e navegue até a pasta:**
    ```bash
    git clone https://github.com/dev-vra/consolidador_cucala.git
    cd consolidador_cucala
    ```
2.  **Crie e ative um ambiente virtual:**
    ```bash
    # Criar
    python -m venv .venv
    # Ativar (Windows)
    .\.venv\Scripts\Activate.ps1
    ```
3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```
4.  **Execute a aplicação:**
    ```bash
    python consolidador.py
    ```
5.  **Login e Senha para acesso:**
    ```bash
    Login: admin
    Senha: admin
    ```

---

## 👨‍💻 Desenvolvido por **Vinicios Reis**

* **LinkedIn** - [LinkdIn](https://www.linkedin.com/in/vinicios-reis-de-araújo-336b5430a)
* **GitHub** - [GitHub](https://github.com/dev-vra)
* **Email** - [Email](mailto:dev.vinnreis@gmail.com)

---

## 📄 Licença

Este projeto é distribuído sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para obter mais detalhes.

