# 📊 Script: Gerar Base e Enviar por E-mail

Este script em Python conecta-se a uma base de dados Oracle, executa uma consulta SQL, salva os resultados em uma planilha Excel, e envia esse arquivo por e-mail usando o Outlook com uma assinatura personalizada. Ele também limpa uma pasta de arquivos temporários após o envio.

---

## ⚙️ Requisitos

- Python 3.x
- Oracle Instant Client (ex: `instantclient_23_8`)
- Conexão com banco de dados Oracle
- Microsoft Outlook instalado
- Bibliotecas Python:

```bash
pip install oracledb pandas openpyxl xlwings pywin32 pillow
```

---

## 📂 Estrutura do Script

1. **Conexão com o banco Oracle**
2. **Execução de uma consulta SQL**
3. **Geração de uma planilha Excel**
4. **Criação de um e-mail no Outlook**
5. **Envio da planilha como anexo**
6. **Limpeza de uma pasta local após envio**

---

## ✏️ Configurações Necessárias

Antes de rodar o script, atualize os seguintes pontos:

- **Credenciais de banco de dados:**
  ```python
  username = 'xxxxxx'
  password = 'xxxxxxxx'
  host = 'xxxxxxxxx'
  port = xxxxxxx
  service_name = 'xxxxxx'
  ```

- **Caminhos:**
  - `lib_dir=r"instantclient_23_8"`
  - Caminho do arquivo Excel: `arquivo = r'endereço da pasta'`
  - Caminho do anexo: `base = r'caminho do arquivo'`
  - Pasta a ser limpa: `pasta_prints = r'endereço da pasta'`

- **E-mails:**
  ```python
  email_para = ['exemplo@empresa.com']
  email_cc = ['copia@empresa.com']
  assunto = 'Assunto do e-mail'
  ```

- **Assinatura personalizada (opcional):**
  ```python
  Assinatura = 'Nome'
  Cargo = 'Cargo'
  Setor = 'Setor'
  Gerencia = 'Gerência'
  contato = 'WhatsApp ou telefone'
  Link = 'https://wa.me/seunumero'
  email = 'seuemail@empresa.com'
  ```

---

## 🧪 Execução

Basta rodar o script com:

```bash
python "gerar base e enviar.py"
```

---

## ⚠️ Observações

- O Outlook precisa estar instalado e configurado para envio automático.
- A imagem de assinatura e o anexo devem existir nos caminhos definidos.
- A query SQL deve ser ajustada para suas necessidades.
