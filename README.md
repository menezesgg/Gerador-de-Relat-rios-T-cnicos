
```markdown
# 🛠️ Gerador de Relatórios Técnicos (RAT) para Escolas

### 📌 Descrição

Automatizador desenvolvido no **segundo dia de trabalho** para gerar documentos Word (.docx) personalizados com informações técnicas de **635 escolas** da rede estadual. O script substitui uma tarefa que antes era feita por **quatro colaboradores durante vários dias** — agora concluída em minutos.

---

### ⚙️ Funcionalidades

- 📥 Leitura de múltiplos arquivos CSV (acompanhamento, racks e endereçamento IP).
- 📝 Geração automatizada de documentos `.docx` para cada escola.
- 🗂️ Organização dos arquivos em pastas nomeadas com o CIE da escola.
- 📊 Barra de progresso para acompanhar o processamento.
- 🚀 Redução de tempo, esforço manual e chance de erro.

---

### 📁 Requisitos

- Python 3.x  
- [pandas](https://pandas.pydata.org/)
- [python-docx](https://python-docx.readthedocs.io/)
- [tqdm](https://github.com/tqdm/tqdm) (barra de progresso)

Instale com:

```bash
pip install pandas python-docx tqdm
```

---

### ▶️ Como usar

1. Coloque os arquivos CSV na mesma pasta do script:
   - `acompanhamento.csv`
   - `dados-redes.csv`
   - `quantidade-de-racks.csv`

2. Execute o script:

```bash
python gerar_documentos.py
```

3. Os documentos serão gerados dentro da pasta `saida_docs`, organizados por escola:

```
saida_docs/
├── 123456789012/
│   └── Termo_123456789012.docx
├── 987654321098/
│   └── Termo_987654321098.docx
...
```

---

### 🧠 Exemplo de conteúdo gerado

```
Escola: ESCOLA ESTADUAL VALENTIM GENTIL
Diretoria de Ensino: [DIRETORIA]
Município: [MUNICÍPIO]
Endereço: [ENDEREÇO]
Quantidade de Switches para instalar: 3

📷 3 fotos dos racks antes da substituição  
📷 3 fotos dos switches depois da substituição  
🧾 3 RATs com número de série  
...
```

---

### 📌 Autor

**Alexandre Menezes Gomes**  
📍 Goiânia, GO  

---

### 🤝 Contribuições

Pull requests são bem-vindos. Sinta-se livre para sugerir melhorias ou adaptar o projeto para novos contextos!

---

### 🏷️ Licença

Este projeto está licenciado sob a **MIT License**. Veja o arquivo `LICENSE` para mais detalhes.

```
