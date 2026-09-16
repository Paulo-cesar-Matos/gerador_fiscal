# 🔁 Atualizações

### 🚀 Novidades da Versão 2.0.0

* **🔍 Painel Central de Busca:** Caixas dedicadas para busca individual por Cliente e Endereço.
* **💡 Grifado Visual:** Termos pesquisados são destacados em amarelo na prévia da mensagem.
* **✉️ Mensagens Individuais:** Navegação rápida (Anterior/Próximo) entre as ordens dos fiscais.
* **🧹 Formatação Limpa:** Remoção do campo de leitura atual e correção de números de ordens/UCs.
# 📋 Gerador de Mensagens de Fiscalização

Aplicativo desktop em Python/Tkinter para leitura de planilhas de ordens de serviço (`GERAL`) e geração de mensagens formatadas por fiscal para envio operacional via WhatsApp.

---

## 🚀 Como Baixar e Usar (Para Usuários)

Não é necessário ter o Python instalado para usar o programa:
1. Vá até a seção de **[Releases](../../releases)** do repositório.
2. Faça o download do arquivo executável **`gerador_fiscal.exe`**.
3. Execute o aplicativo no Windows.

---

## 🛠️ Como Rodar o Código Fonte (Para Desenvolvedores)

### Pré-requisitos
* Python 3.10+
* Git

### Instalação (Use o Prompt de Comando/CMD)
```bash
# Clocar o repositório
git clone https://github.com/Paulo-cesar-Matos/gerador_fiscal.git

# Entrar na pasta do projeto
cd C:\Users\[seuUsuario]\gerador_fiscal

# Instalar dependências
pip install numpy pandas openpyxl pyperclip

# Executar a aplicação
python gerador_fiscal.py
