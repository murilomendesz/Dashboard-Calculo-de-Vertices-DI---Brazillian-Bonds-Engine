# Como Rodar o Projeto

> Pré-requisito: Python 3.10 ou superior instalado na máquina.

---

### 1. Baixe o projeto

Clique em **Code → Download ZIP** no GitHub e extraia a pasta, ou se tiver Git instalado:

```bash
git clone https://github.com/murilomendesz/curva-di.git
```

---

### 2. Abra o terminal na pasta do projeto

- Abra a pasta extraída
- Clique na barra de endereço do Windows Explorer, digite `cmd` e pressione Enter

---

### 3. Instale as dependências

```bash
pip install -r requirements.txt
```

Aguarde a instalação terminar.

---

### 4. Execute o projeto

```bash
python main.py
```

O Excel abrirá automaticamente com o dashboard completo.

---

### Observações

- É necessário ter o **Microsoft Excel** instalado na máquina
- Funciona apenas no **Windows**
- Os dados são buscados diretamente da ANBIMA — é necessário **conexão com a internet**
- Se rodar antes das ~14h, o projeto usa automaticamente os dados do dia anterior (a ANBIMA publica à tarde)
