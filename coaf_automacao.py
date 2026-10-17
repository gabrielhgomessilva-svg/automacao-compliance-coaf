# -*- coding: utf-8 -*-
"""
Automação COAF - gera a planilha diária de classificação de risco por operação.

Para cada operação (número do borderô):
  1. localiza o PDF do borderô de pagamento e lê cedente, CNPJ e pagamentos;
  2. abre o modelo, preenche operação + código do cedente (o risco vem da
     planilha de controle, como já funciona hoje);
  3. se o risco é ELEVADO, copia o MOTIVO da planilha de controle para a
     Fundamentação;
  4. se há pagamento a terceiro > R$ 10.000 (CNPJ do favorecido diferente do
     CNPJ do cedente), marca SIM no parâmetro correspondente e inclui a
     fundamentação;
  5. salva em  <pasta do dia> como  <operação>.xls

Requer Windows + Microsoft Excel instalado (usa o próprio Excel para abrir e
salvar o .xls, preservando layout, logos e fórmulas do modelo).
"""
import configparser
import datetime
import math
import os
import re
import sys
from dataclasses import dataclass, field

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO",
         "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"]

TEXTO_TERCEIRO = ("Pagamento a terceiro superior a R$ 10.000,00, autorizado pelo "
                  "cliente, desde que NÃO destinado, comprovadamente, a fornecedor "
                  "de bens ou serviços do cliente.")

# Células do formulário (aba "Plan1 (2)")
CEL_OPERACAO = "C10"
CEL_COD_CEDENTE = "C13"
CEL_RISCO = "C14"
CEL_NOME = "C11"
CEL_CNPJ = "C12"
CEL_TERCEIRO_SIM = "F23"
CEL_TERCEIRO_NAO = "G23"
CEL_FUNDAMENTACAO = "B46"      # B46:G46 mesclada (linha alta, com quebra de texto)
CEL_RESPONSAVEL = "A49"
CEL_DATA_COAF = "F58"          # F58:G58 mesclada - "Data da Comunicação ao COAF"
FUND_LARGURA_CHARS = 120       # estimativa conservadora de caracteres por linha
FUND_ALTURA_MIN = 65.25


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def so_digitos(s):
    return re.sub(r"\D", "", s or "")


def br_para_float(s):
    return float(s.replace(".", "").replace(",", "."))


def float_para_br(v):
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def carregar_config():
    cfg = configparser.RawConfigParser()
    caminho = os.path.join(base_dir(), "config.ini")
    if not os.path.exists(caminho):
        raise SystemExit(f"config.ini não encontrado em {caminho}")
    with open(caminho, encoding="utf-8-sig") as f:
        cfg.read_string("[geral]\n" + f.read())
    g = cfg["geral"]
    return {
        "pasta_arquivo": g.get("PASTA_ARQUIVO").strip(),
        "pasta_borderos": g.get("PASTA_BORDEROS").strip(),
        "controle": g.get("PLANILHA_CONTROLE").strip(),
        "modelo": g.get("PLANILHA_MODELO").strip(),
        "aba": g.get("ABA_FORMULARIO", "Plan1 (2)").strip(),
        "responsavel": g.get("RESPONSAVEL", "").strip(),
        "limite": float(g.get("LIMITE_TERCEIRO", "10000").strip().replace(",", ".")),
        "fonte": g.get("FONTE", "pdf").strip().lower(),                 # pdf | sql
        "sql_driver": g.get("SQL_DRIVER", "ODBC Driver 17 for SQL Server").strip(),
        "sql_servidor": g.get("SQL_SERVIDOR", "").strip(),
        "sql_banco": g.get("SQL_BANCO", "").strip(),
        "sql_usuario": g.get("SQL_USUARIO", "").strip(),                # vazio = autenticação do Windows
        "sql_senha": g.get("SQL_SENHA", "").strip(),
    }


# --------------------------------------------------------------------------
# Borderô de pagamento (PDF)
# --------------------------------------------------------------------------
@dataclass
class Pagamento:
    forma: str
    valor: float
    favorecido: str = ""
    documento: str = ""   # CNPJ/CPF do favorecido


@dataclass
class Bordero:
    numero: str
    cod_cedente: int
    nome_cedente: str
    pagamentos: list = field(default_factory=list)
    liquido: float = None      # valor líquido do borderô (None = desconhecido)


def texto_do_pdf(caminho):
    import pdfplumber
    partes = []
    with pdfplumber.open(caminho) as pdf:
        for p in pdf.pages:
            partes.append(p.extract_text() or "")
    return "\n".join(partes)


def interpretar_bordero(texto):
    m = re.search(r"BORDERO:\s*(\d+)\s+CEDENTE:\s*(\d+)\s*-\s*(.+)", texto)
    if not m:   # layout dos borderôs sem pagamento: "BORDERÔ: n DATA: d CLIENTE: cod - nome"
        m = re.search(r"BORDER[OÔ]:\s*(\d+)\s+DATA:\s*\S+\s+CLIENTE:\s*(\d+)\s*-\s*(.+)", texto)
    if not m:
        raise ValueError("Não encontrei 'BORDERO: ... CEDENTE: ...' no PDF.")
    b = Bordero(m.group(1), int(m.group(2)), m.group(3).strip())
    ml = re.search(r"Líquido\s+([\d.]+,\d{2})", texto)
    if ml:
        b.liquido = br_para_float(ml.group(1))

    m = re.search(r"FORMA DE PAGAMENTO(.*?)(TOTAL DOS PAGAMENTOS|$)", texto, re.S)
    bloco = m.group(1) if m else ""
    atual = None
    for linha in bloco.splitlines():
        linha = linha.strip()
        ini = re.match(r"^(\d{1,3})\s+([\d.]+,\d{2})$", linha)
        if ini:
            atual = Pagamento(ini.group(1), br_para_float(ini.group(2)))
            b.pagamentos.append(atual)
            continue
        if atual is None:
            continue
        f = re.match(r"(?i)^Favorecido:\s*(.*)$", linha)
        if f:
            atual.favorecido = f.group(1).strip()
        d = re.match(r"(?i)^CNPJ/CPF:\s*([\d.\-/]*)", linha)
        if d:
            atual.documento = d.group(1).strip()
    return b


def achar_pdf(pasta, operacao):
    """PDF salvo como '327101 - NOME.pdf' (aceita também '327101.pdf')."""
    candidatos = []
    for nome in os.listdir(pasta):
        if not nome.lower().endswith(".pdf"):
            continue
        m = re.match(r"^\s*(\d+)\b", nome)
        if m and m.group(1) == operacao:
            candidatos.append(nome)
    if not candidatos:
        return None
    candidatos.sort()
    return os.path.join(pasta, candidatos[0])


def avaliar_terceiros(pagamentos, cnpj_cedente, limite):
    """Retorna (terceiros_acima_do_limite, nao_verificaveis)."""
    terceiros, revisar = [], []
    alvo = so_digitos(cnpj_cedente)
    for p in pagamentos:
        if p.valor <= limite:
            continue
        if not so_digitos(p.documento):
            revisar.append(p)           # sem CNPJ/CPF no PDF -> não dá para concluir
        elif so_digitos(p.documento) != alvo:
            terceiros.append(p)
    return terceiros, revisar


def montar_fundamentacao(risco, motivo, terceiros):
    linhas, avisos = [], []
    if (risco or "").strip().lower() == "elevado":
        mot = (motivo or "").strip()
        if not mot or mot.upper() == "N/D":
            avisos.append("Risco ELEVADO sem MOTIVO na planilha de controle - fundamentação deixada em branco.")
        else:
            linhas.append(mot)
    for p in terceiros:
        linhas.append(f"{TEXTO_TERCEIRO} (Favorecido: {p.favorecido}, CNPJ/CPF {p.documento}, R$ {float_para_br(p.valor)})")
    return linhas, avisos


def altura_fundamentacao(linhas):
    n = sum(max(1, math.ceil(len(l) / FUND_LARGURA_CHARS)) for l in linhas)
    return max(FUND_ALTURA_MIN, 15.0 * n + 3)


# --------------------------------------------------------------------------
# Fonte SQL (ERP da empresa / SQL Server) - SOMENTE LEITURA
# --------------------------------------------------------------------------
# Tabelas (conferidas com borderôs reais):
#   SIGBORS        cabeçalho: BORDERO, CLIFOR (= código do cedente), EstadoBordero, DataConclusao
#   SIGCAD         cadastro: CODIGO, NOME
#   SIGBDCHQ       pagamentos FACTORING: BORDERO, VALOR, forma, favorecido (= cadfavor.ctrl_id)
#   SIGBDPAG_FIDC  pagamentos FIDC:      BORDERO, VALOR, CODFAVOR   (= cadfavor.ctrl_id)
#   cadfavor       favorecidos: ctrl_id, favorecido (nome), CGC (CNPJ/CPF), banco, agencia, conta
SQL_CABECALHO = ("SELECT b.BORDERO, b.CLIFOR, c.NOME, b.LIQUIDO FROM SIGBORS b "
                 "LEFT JOIN SIGCAD c ON c.CODIGO = b.CLIFOR WHERE b.BORDERO = ?")
SQL_PAG_FACTORING = ("SELECT q.VALOR, q.forma, f.favorecido, f.CGC FROM SIGBDCHQ q "
                     "LEFT JOIN cadfavor f ON f.ctrl_id = q.favorecido WHERE q.BORDERO = ?")
SQL_PAG_FIDC = ("SELECT p.VALOR, f.favorecido, f.CGC FROM SIGBDPAG_FIDC p "
                "LEFT JOIN cadfavor f ON f.ctrl_id = p.CODFAVOR WHERE p.BORDERO = ?")
SQL_OPS_DO_DIA = ("SELECT BORDERO FROM SIGBORS WHERE EstadoBordero LIKE 'Conclu%' "
                  "AND DataConclusao >= ? AND DataConclusao < ? ORDER BY BORDERO")


def conectar_sql(cfg):
    import pyodbc
    if not cfg["sql_servidor"] or not cfg["sql_banco"]:
        raise ValueError("Preencha SQL_SERVIDOR e SQL_BANCO no config.ini")
    partes = [f"DRIVER={{{cfg['sql_driver']}}}", f"SERVER={cfg['sql_servidor']}",
              f"DATABASE={cfg['sql_banco']}", "Encrypt=no", "TrustServerCertificate=yes",
              "ApplicationName=AutomacaoCOAF"]
    if cfg["sql_usuario"]:
        senha = cfg["sql_senha"].replace("}", "}}")      # permite ; { = na senha
        partes += [f"UID={cfg['sql_usuario']}", f"PWD={{{senha}}}"]
    else:
        partes.append("Trusted_Connection=yes")
    conn = pyodbc.connect(";".join(partes), timeout=10, autocommit=True)
    cur = conn.cursor()
    # leitura sem bloquear quem está gravando no sistema
    cur.execute("SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED")
    cur.execute("SET LOCK_TIMEOUT 5000")
    return conn


def sql_operacoes_do_dia(conn, data):
    ini = datetime.datetime(data.year, data.month, data.day)
    cur = conn.cursor()
    cur.execute(SQL_OPS_DO_DIA, (ini, ini + datetime.timedelta(days=1)))
    return [str(int(r[0])) for r in cur.fetchall()]


def sql_bordero(conn, operacao):
    """Retorna Bordero (com pagamentos de factoring e FIDC) ou None se o borderô não existe."""
    cur = conn.cursor()
    cur.execute(SQL_CABECALHO, (int(operacao),))
    row = cur.fetchone()
    if not row:
        return None
    b = Bordero(str(int(row[0])), int(row[1] or 0), (row[2] or "").strip(), liquido=float(row[3] or 0))
    cur.execute(SQL_PAG_FACTORING, (int(operacao),))
    for valor, forma, nome, cgc in cur.fetchall():
        b.pagamentos.append(Pagamento(str(forma or "").strip() or "FACTORING", float(valor or 0),
                                      (nome or "").strip(), (cgc or "").strip()))
    cur.execute(SQL_PAG_FIDC, (int(operacao),))
    for valor, nome, cgc in cur.fetchall():
        b.pagamentos.append(Pagamento("FIDC", float(valor or 0), (nome or "").strip(), (cgc or "").strip()))
    return b


def ler_bordero(cfg, sql_conn, operacao, r):
    """Lê o borderô da fonte configurada (SQL ou PDF). Em caso de falha preenche r e retorna None."""
    if cfg["fonte"] == "sql":
        try:
            b = sql_bordero(sql_conn, operacao)
        except Exception as e:
            r["status"] = "ERRO"
            r["msg"].append(f"Falha ao consultar o banco: {e}")
            return None
        if b is None:
            r["status"] = "ERRO"
            r["msg"].append(f"Borderô {operacao} não encontrado no banco '{cfg['sql_banco']}' "
                            f"(confira se a base é a atualizada).")
            return None
        return b
    pdf = achar_pdf(cfg["pasta_borderos"], operacao)
    if not pdf:
        r["status"] = "ERRO"
        r["msg"].append(f"PDF do borderô {operacao} não encontrado em {cfg['pasta_borderos']}")
        return None
    try:
        return interpretar_bordero(texto_do_pdf(pdf))
    except Exception as e:
        r["status"] = "ERRO"
        r["msg"].append(f"Falha ao ler o PDF ({os.path.basename(pdf)}): {e}")
        return None


# --------------------------------------------------------------------------
# Excel (COM)
# --------------------------------------------------------------------------
def iniciar_excel():
    import win32com.client as win32
    xl = win32.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    xl.AskToUpdateLinks = False
    return xl


def ler_controle(xl, caminho):
    """Lê a aba CEDENTES (somente leitura) -> {cod: dict}. Mantém a 1ª ocorrência (como o PROCV)."""
    wb = xl.Workbooks.Open(caminho, 0, True)    # UpdateLinks=0, ReadOnly=True
    try:
        ws = wb.Worksheets("CEDENTES")
        ultima = ws.Cells(ws.Rows.Count, 1).End(-4162).Row   # xlUp
        dados = ws.Range(ws.Cells(1, 1), ws.Cells(ultima, 14)).Value
    finally:
        wb.Close(False)
    tab = {}
    for linha in dados[1:]:
        cod = linha[0]
        if cod is None or str(cod).strip() == "":
            continue
        try:
            cod = int(float(cod))
        except (ValueError, TypeError):
            continue
        if cod in tab:
            continue
        tab[cod] = {
            "nome": str(linha[2] or "").strip(),
            "risco": str(linha[7] or "").strip(),
            "cnpj": str(linha[9] or "").strip(),
            "motivo": str(linha[11] or "").strip(),
        }
    return tab


def _limpar(ws, endereco):
    """Limpa a célula inteira, inclusive quando faz parte de uma área mesclada."""
    ws.Range(endereco).MergeArea.ClearContents()


def _escrever(ws, endereco, valor):
    """Escreve na célula (ou na célula superior-esquerda da área mesclada)."""
    ws.Range(endereco).MergeArea.Cells(1, 1).Value = valor


def gerar_planilha(xl, cfg, operacao, cedente, linhas_fund, terceiro, destino, data=None):
    """Abre o modelo, preenche e salva em 'destino' (.xls). Retorna (nome, cnpj, risco, avisos) lidos da planilha."""
    avisos = []
    controle_abs = os.path.abspath(cfg["controle"])
    # Abre a planilha de controle (somente leitura) ANTES: com ela aberta nesta instância do Excel,
    # as fórmulas leem os valores atuais dela, e não o "cache" antigo guardado dentro do modelo.
    ctl = None
    try:
        ctl = xl.Workbooks.Open(controle_abs, 0, True)
    except Exception as e:
        avisos.append(f"ATENÇÃO: não consegui abrir a planilha de controle ({e}) - os dados podem vir do cache do modelo.")
    wb = xl.Workbooks.Open(os.path.abspath(cfg["modelo"]), 0)   # sem atualizar vínculos ainda
    try:
        ws = wb.Worksheets(cfg["aba"])
        # religa o vínculo do modelo para a planilha de controle atual (PLANILHA_CONTROLE do config.ini)
        try:
            for src in (wb.LinkSources(1) or []):
                if os.path.normcase(os.path.abspath(src)) != os.path.normcase(controle_abs):
                    wb.ChangeLink(src, controle_abs, 1)
        except Exception as e:
            avisos.append(f"ATENÇÃO: não consegui religar o vínculo com a planilha de controle ({e}).")

        _escrever(ws, CEL_OPERACAO, int(operacao))
        _escrever(ws, CEL_COD_CEDENTE, int(cedente))
        if cfg["responsavel"]:
            _escrever(ws, CEL_RESPONSAVEL, int(cfg["responsavel"]))
        if data is not None:
            _escrever(ws, CEL_DATA_COAF, datetime.datetime(data.year, data.month, data.day))

        # parâmetro "pagamento a terceiro": volta ao padrão e marca se for o caso
        _limpar(ws, CEL_TERCEIRO_SIM)
        _escrever(ws, CEL_TERCEIRO_NAO, "X")
        if terceiro:
            _escrever(ws, CEL_TERCEIRO_SIM, "X")
            _limpar(ws, CEL_TERCEIRO_NAO)

        # fundamentação (B46:G46 é mesclada)
        _limpar(ws, CEL_FUNDAMENTACAO)
        if linhas_fund:
            _escrever(ws, CEL_FUNDAMENTACAO, "\n".join(linhas_fund))
            area = ws.Range(CEL_FUNDAMENTACAO).MergeArea
            try:        # o modelo já tem quebra de texto; isto é só um reforço e não pode travar
                area.WrapText = True
            except Exception:
                pass
            try:
                ws.Rows(area.Row).RowHeight = altura_fundamentacao(linhas_fund)
            except Exception:
                pass

        # Com a planilha de controle aberta nesta instância do Excel, as fórmulas já leem os valores atuais;
        # UpdateLink é só um reforço (precisa do nome EXATO do vínculo) e pode falhar sem prejuízo.
        try:
            nomes = wb.LinkSources(1) or []
            wb.UpdateLink(nomes[0] if nomes else controle_abs, 1)
        except Exception as e:
            if ctl is None:    # sem o controle aberto, a atualização é essencial
                avisos.append(f"ATENÇÃO: não consegui atualizar o vínculo ({e}).")
        xl.CalculateFull()

        # confere se o vínculo realmente aponta para a planilha de controle
        try:
            fontes = [os.path.normcase(os.path.abspath(x)) for x in (wb.LinkSources(1) or [])]
            if fontes and fontes != [os.path.normcase(controle_abs)]:
                avisos.append(f"ATENÇÃO: o vínculo da planilha aponta para {', '.join(wb.LinkSources(1))} "
                              f"e não para {controle_abs}.")
        except Exception:
            pass

        lido = (str(ws.Range(CEL_NOME).Value), str(ws.Range(CEL_CNPJ).Value), str(ws.Range(CEL_RISCO).Value), avisos)
        wb.SaveAs(destino, 56)    # xlExcel8 (.xls)
        return lido
    finally:
        wb.Close(False)
        if ctl is not None:
            try:
                ctl.Close(False)
            except Exception:
                pass


# --------------------------------------------------------------------------
# Fluxo principal
# --------------------------------------------------------------------------
def pasta_do_dia(cfg, data):
    base = cfg["pasta_arquivo"].replace("{ano}", str(data.year))
    return os.path.join(base, MESES[data.month - 1], data.strftime("%d.%m.%Y"))


def processar(xl, cfg, controle, operacao, pasta_dia, sobrescrever=None, data=None, sql_conn=None):
    r = {"op": operacao, "status": "", "msg": []}
    existentes = [os.path.join(pasta_dia, f"{operacao}{suf}.xls") for suf in ("", "S")]
    existentes = [e for e in existentes if os.path.exists(e)]
    if existentes:
        if sobrescrever is None or not sobrescrever(existentes[0]):
            r["status"] = "JÁ EXISTE"
            r["msg"].append(f"{existentes[0]} já existe - não foi sobrescrito.")
            return r
        for e in existentes:       # sobrescrever: remove a variante antiga (com ou sem S)
            try:
                os.remove(e)
            except OSError:
                pass

    b = ler_bordero(cfg, sql_conn, operacao, r)
    if b is None:
        return r
    if b.numero != operacao:
        r["msg"].append(f"ATENÇÃO: o PDF diz borderô {b.numero}, mas você pediu {operacao}.")

    ced = controle.get(b.cod_cedente)
    if not ced:
        r["status"] = "PENDENTE"
        r["msg"].append(f"Cedente {b.cod_cedente} ({b.nome_cedente}) NÃO está na planilha de controle de risco "
                        f"- classificar o cedente primeiro e rodar de novo.")
        return r

    terceiros, revisar = avaliar_terceiros(b.pagamentos, ced["cnpj"], cfg["limite"])
    if not b.pagamentos and (b.liquido is None or b.liquido > 0):
        r["msg"].append("ATENÇÃO: não encontrei nenhum pagamento (PDF/banco) mas o líquido do borderô não é zero - conferir pagamento a terceiro manualmente.")
    elif not b.pagamentos:
        r["msg"].append("Borderô sem pagamento (líquido R$ 0,00): planilha gerada normalmente, sem pagamento a terceiro.")
    for p in revisar:
        r["msg"].append(f"ATENÇÃO: pagamento de R$ {float_para_br(p.valor)} sem CNPJ/CPF no borderô "
                        f"({p.favorecido or p.forma}) - conferir pagamento a terceiro manualmente.")

    linhas, avisos = montar_fundamentacao(ced["risco"], ced["motivo"], terceiros)
    r["msg"].extend(avisos)
    # "S" no nome: risco Elevado OU pagamento a terceiro acima do limite
    com_s = ced["risco"].strip().lower() == "elevado" or bool(terceiros)
    destino = os.path.join(pasta_dia, f"{operacao}{'S' if com_s else ''}.xls")

    os.makedirs(pasta_dia, exist_ok=True)
    try:
        nome, cnpj, risco, avisos_xl = gerar_planilha(xl, cfg, operacao, b.cod_cedente, linhas, bool(terceiros), destino, data)
    except Exception as e:
        r["status"] = "ERRO"
        r["msg"].append(f"Falha no Excel: {e}")
        return r

    r["msg"].extend(avisos_xl)
    if risco.strip().lower() != ced["risco"].strip().lower():
        r["msg"].append(f"ATENÇÃO: a planilha mostra risco '{risco}' mas o controle diz '{ced['risco']}' "
                        f"- conferir os vínculos da planilha.")
    r["status"] = "OK"
    r["cedente"] = f"{b.cod_cedente} - {ced['nome']}"
    r["risco"] = ced["risco"]
    r["terceiro"] = bool(terceiros)
    r["fund"] = linhas
    r["destino"] = destino
    return r


def mostrar(r):
    ic = {"OK": "[OK]", "ERRO": "[ERRO]", "PENDENTE": "[PENDENTE]", "JÁ EXISTE": "[JÁ EXISTE]"}.get(r["status"], "[?]")
    print(f"{ic} Operação {r['op']}")
    if r["status"] == "OK":
        print(f"      Cedente: {r['cedente']}")
        print(f"      Risco:   {r['risco']}")
        print(f"      Pagamento a terceiro > limite: {'SIM' if r['terceiro'] else 'não'}")
        for l in r["fund"]:
            print(f"      Fundamentação: {l}")
        print(f"      Salvo em: {r['destino']}")
    for m in r["msg"]:
        print(f"      * {m}")


def registrar(resultados, pasta_dia):
    try:
        with open(os.path.join(base_dir(), "registro.log"), "a", encoding="utf-8") as f:
            for r in resultados:
                f.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}\t{pasta_dia}\t{r['op']}\t{r['status']}\t"
                        f"{r.get('risco','')}\t{'; '.join(r['msg'])}\n")
    except OSError:
        pass


def ler_operacoes(texto):
    return [t for t in re.split(r"[\s,;]+", texto.strip()) if t.isdigit()]


def listar_operacoes_do_dia(pasta, data):
    """Números dos borderôs (PDFs) salvos na pasta na data informada (data de modificação, a mesma que o Explorer mostra)."""
    ops = set()
    for e in os.scandir(pasta):
        if not e.is_file() or not e.name.lower().endswith(".pdf"):
            continue
        m = re.match(r"^\s*(\d+)\b", e.name)
        if not m:
            continue
        st = e.stat()
        if datetime.date.fromtimestamp(st.st_mtime) == data:
            ops.add(m.group(1))
    return sorted(ops, key=int)


def ler_data(texto, padrao):
    """Aceita: vazio (=padrão), 'ontem', dd.mm.aaaa, dd/mm/aaaa, dd/mm, dd.mm. Retorna date ou None se inválida."""
    t = texto.strip().lower()
    if not t:
        return padrao
    if t == "ontem":
        return datetime.date.today() - datetime.timedelta(days=1)
    t = t.replace("/", ".").replace("-", ".")
    for fmt in ("%d.%m.%Y", "%d.%m.%y"):
        try:
            return datetime.datetime.strptime(t, fmt).date()
        except ValueError:
            pass
    try:
        d = datetime.datetime.strptime(t, "%d.%m")
        return datetime.date(datetime.date.today().year, d.month, d.day)
    except ValueError:
        return None


def main(argv):
    """
    Sem argumentos      -> AUTOMÁTICO: pega os borderôs salvos HOJE e gera as planilhas que faltam.
    dd.mm.aaaa          -> AUTOMÁTICO para aquela data.
    números             -> gera só essas operações (data = hoje ou a informada).
    --manual            -> pergunta data e números (modo antigo).
    """
    try:
        sys.stdout.reconfigure(errors="replace")
    except Exception:
        pass
    cfg = carregar_config()
    cfg["modelo"] = cfg["modelo"] if os.path.isabs(cfg["modelo"]) else os.path.join(base_dir(), cfg["modelo"])

    data = datetime.date.today()
    ops = []
    manual = "--manual" in argv
    for a in argv:
        if re.match(r"^\d{2}\.\d{2}\.\d{4}$", a):
            data = datetime.datetime.strptime(a, "%d.%m.%Y").date()
        elif a.isdigit():
            ops.append(a)

    print("=== Automação COAF ===")
    tem_data_arg = any(re.match(r"^\d{2}\.\d{2}\.\d{4}$", a) for a in argv)
    if manual or (not ops and not tem_data_arg):
        # pergunta o dia; Enter = hoje
        while True:
            resp = input(f"Dia a processar [{data:%d.%m.%Y}] (Enter = hoje; ex.: 08.10.2026, 08/10 ou 'ontem'): ")
            d = ler_data(resp, data)
            if d:
                data = d
                break
            print("  Data inválida, tente de novo.")
    pasta_dia = pasta_do_dia(cfg, data)
    print(f"Data: {data:%d.%m.%Y}   Pasta do dia: {pasta_dia}")

    sql_conn = None
    if cfg["fonte"] == "sql":
        try:
            print(f"Conectando ao banco {cfg['sql_servidor']} / {cfg['sql_banco']} (somente leitura)...")
            sql_conn = conectar_sql(cfg)
        except Exception as e:
            print(f"Não consegui conectar ao banco: {e}")
            input("\nEnter para fechar...")
            return

    if not manual and not ops:
        try:
            if cfg["fonte"] == "sql":
                ops = sql_operacoes_do_dia(sql_conn, data)
                print(f"{len(ops)} borderô(s) concluído(s) em {data:%d.%m.%Y}: {', '.join(ops) if ops else '-'}")
            else:
                ops = listar_operacoes_do_dia(cfg["pasta_borderos"], data)
                print(f"{len(ops)} borderô(s) salvo(s) em {data:%d.%m.%Y}: {', '.join(ops) if ops else '-'}")
        except Exception as e:
            print(f"Não consegui listar os borderôs do dia: {e}")
            input("\nEnter para fechar...")
            return
        if not ops:
            input("\nNada a fazer. Enter para fechar...")
            return

    xl = iniciar_excel()
    resultados = []
    try:
        print("Lendo planilha de controle de risco...")
        controle = ler_controle(xl, cfg["controle"])
        print(f"  {len(controle)} cedentes carregados.\n")

        def perguntar(destino):
            return input(f"   {os.path.basename(destino)} já existe. Sobrescrever? (s/N) ").strip().lower() == "s"

        if manual:
            while True:
                lote = ler_operacoes(input("Número(s) da operação (vários separados por espaço; Enter para sair): "))
                if not lote:
                    break
                for op in lote:
                    r = processar(xl, cfg, controle, op, pasta_dia, perguntar, data, sql_conn)
                    mostrar(r)
                    resultados.append(r)
        else:
            for op in ops:
                r = processar(xl, cfg, controle, op, pasta_dia, None, data, sql_conn)   # já existentes são pulados
                mostrar(r)
                resultados.append(r)
        registrar(resultados, pasta_dia)
    finally:
        try:
            xl.Quit()
        except Exception:
            pass

    if resultados:
        cont = {}
        for r in resultados:
            cont[r["status"]] = cont.get(r["status"], 0) + 1
        print("\nResumo: " + " | ".join(f"{k}: {v}" for k, v in cont.items()))
    input("\nEnter para fechar...")


if __name__ == "__main__":
    main(sys.argv[1:])
