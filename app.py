import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from openpyxl import Workbook, load_workbook
from datetime import datetime
from pathlib import Path
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import sys


ESTOQUE_MINIMO_GLOBAL = 10
DATE_FORMAT = "%d/%m/%Y"

CAMINHO: Path | None = None

def criar_planilha_inicial(caminho: Path):
    """Cria planilha com abas 'Entradas' e 'Saidas' e cabeçalhos, se não existir."""
    wb = Workbook()
    wsE = wb.active
    wsE.title = "Entradas"
    wsE.append(["ID", "Código", "Produto", "Quantidade", "Data Entrada", "Fornecedor", "Observação"])
    wsS = wb.create_sheet("Saidas")
    wsS.append(["ID", "Código", "Produto", "Quantidade", "Data Saida", "Destino/Obra", "Observação"])
    wb.save(caminho)

def carregar_planilha():
    """Retorna objeto Workbook ou None com mensagem de erro."""
    global CAMINHO
    if not CAMINHO:
        messagebox.showerror("Erro", "Nenhuma planilha selecionada.")
        return None
    if not CAMINHO.exists():
        messagebox.showerror("Erro", f"Arquivo não encontrado:\n{CAMINHO}")
        return None
    try:
        wb = load_workbook(CAMINHO)
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir o arquivo:\n{e}")
        return None

    if "Entradas" not in wb.sheetnames:
        ws = wb.create_sheet("Entradas")
        ws.append(["ID", "Código", "Produto", "Quantidade", "Data Entrada", "Fornecedor", "Observação"])
    if "Saidas" not in wb.sheetnames:
        ws = wb.create_sheet("Saidas")
        ws.append(["ID", "Código", "Produto", "Quantidade", "Data Saida", "Destino/Obra", "Observação"])
    return wb

def next_id_for_sheet(ws):
    """Gera ID incremental para a próxima linha da sheet (coluna A)."""
    max_id = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        try:
            if row[0]:
                v = int(row[0])
                if v > max_id:
                    max_id = v
        except Exception:
            continue
    return max_id + 1

def safe_int(value):
    try:
        return int(value)
    except Exception:
        return 0

def escolher_arquivo():
    """Escolher arquivo .xlsx existente ou criar um novo."""
    global CAMINHO
    caminho = filedialog.askopenfilename(
        title="Escolha a planilha de estoque (.xlsx)",
        filetypes=[("Excel Files", "*.xlsx")],
    )
    if not caminho:
        # criar novo
        caminho = filedialog.asksaveasfilename(
            title="Criar nova planilha de estoque (.xlsx)",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
        )
        if caminho:
            CAMINHO = Path(caminho)
            criar_planilha_inicial(CAMINHO)
            messagebox.showinfo("Criado", f"Planilha criada:\n{CAMINHO}")
            atualizar_tudo()
            return
        else:
            return
    CAMINHO = Path(caminho)
    wb = carregar_planilha()
    if wb:
        wb.save(CAMINHO)
    messagebox.showinfo("Arquivo selecionado", f"{CAMINHO}")
    atualizar_tudo()

def adicionar_entrada():
    wb = carregar_planilha()
    if not wb:
        return
    ws = wb["Entradas"]
    codigo = ent_codigo.get().strip()
    produto = ent_produto.get().strip()
    qtd_text = ent_qtd_entrada.get().strip()
    obs = ent_obs_entrada.get().strip()
    fornecedor = ent_fornecedor.get().strip()
    if not codigo or not produto or not qtd_text:
        messagebox.showwarning("Aviso", "Preencha Código, Produto e Quantidade (Entrada).")
        return
    try:
        qtd = int(qtd_text)
        if qtd <= 0:
            raise ValueError
    except Exception:
        messagebox.showerror("Erro", "Quantidade de entrada inválida (deve ser inteiro > 0).")
        return
    nova_id = next_id_for_sheet(ws)
    data = datetime.now().strftime(DATE_FORMAT)
    ws.append([nova_id, codigo, produto, qtd, data, fornecedor, obs])
    wb.save(CAMINHO)
    messagebox.showinfo("Sucesso", f"Entrada registrada: {produto} — {qtd}")
    limpar_campos_entrada()
    atualizar_tudo()

def registrar_saida():
    wb = carregar_planilha()
    if not wb:
        return
    wsE = wb["Entradas"]
    wsS = wb["Saidas"]
    codigo = ent_codigo.get().strip()
    produto = ent_produto.get().strip()
    qtd_text = ent_qtd_saida.get().strip()
    destino = ent_destino.get().strip()
    obs = ent_obs_saida.get().strip()
    if not codigo or not produto or not qtd_text:
        messagebox.showwarning("Aviso", "Preencha Código, Produto e Quantidade (Saída).")
        return
    try:
        qtd = int(qtd_text)
        if qtd <= 0:
            raise ValueError
    except Exception:
        messagebox.showerror("Erro", "Quantidade de saída inválida (deve ser inteiro > 0).")
        return

    resumo = calcular_resumo_por_produto(wb)
    atual = resumo.get(codigo, {}).get("saldo", 0)
    if qtd > atual:
        messagebox.showerror("Erro", f"Estoque insuficiente para saída. Estoque atual: {atual}")
        return

    nova_id = next_id_for_sheet(wsS)
    data = datetime.now().strftime(DATE_FORMAT)
    wsS.append([nova_id, codigo, produto, qtd, data, destino, obs])
    wb.save(CAMINHO)
    messagebox.showinfo("Sucesso", f"Saída registrada: {produto} — {qtd}")
    limpar_campos_saida()
    atualizar_tudo()

def calcular_resumo_por_produto(wb=None):
    """Retorna dict {codigo: {produto: name, entradas: n, saidas: n, saldo: n}}"""
    if wb is None:
        wb = carregar_planilha()
        if not wb:
            return {}
    entradas = {}
    saidas = {}
    for row in wb["Entradas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        codigo = str(row[1])
        produto = str(row[2]) if row[2] is not None else ""
        qtd = safe_int(row[3])
        entradas.setdefault(codigo, {"produto": produto, "total": 0})
        entradas[codigo]["total"] += qtd
    for row in wb["Saidas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[1]:
            continue
        codigo = str(row[1])
        produto = str(row[2]) if row[2] is not None else ""
        qtd = safe_int(row[3])
        saidas.setdefault(codigo, {"produto": produto, "total": 0})
        saidas[codigo]["total"] += qtd
    resumo = {}
    codigos = set(list(entradas.keys()) + list(saidas.keys()))
    for c in codigos:
        nome = entradas.get(c, {}).get("produto") or saidas.get(c, {}).get("produto") or ""
        ent = entradas.get(c, {}).get("total", 0)
        sai = saidas.get(c, {}).get("total", 0)
        saldo = ent - sai
        resumo[c] = {"produto": nome, "entradas": ent, "saidas": sai, "saldo": saldo}
    return resumo

def atualizar_resumo_tree(filtro=""):
    for item in tree_resumo.get_children():
        tree_resumo.delete(item)
    wb = carregar_planilha()
    if not wb:
        return
    resumo = calcular_resumo_por_produto(wb)
    for codigo, info in resumo.items():
        if filtro and filtro.lower() not in codigo.lower() and filtro.lower() not in info["produto"].lower():
            continue
        ent = info["entradas"]
        sai = info["saidas"]
        saldo = info["saldo"]
        item_id = tree_resumo.insert("", "end", values=(codigo, info["produto"], ent, sai, saldo))
        if saldo <= ESTOQUE_MINIMO_GLOBAL:
            tree_resumo.item(item_id, tags=("baixo",))
    tree_resumo.tag_configure("baixo", background="salmon")

def atualizar_entradas_tree():
    for item in tree_entradas.get_children():
        tree_entradas.delete(item)
    wb = carregar_planilha()
    if not wb:
        return
    ws = wb["Entradas"]
    for row in ws.iter_rows(min_row=2):
        cells = [c.value for c in row]
        if not cells or not cells[0]:
            continue
        excel_row_index = row[0].row
        tree_entradas.insert("", "end", iid=str(excel_row_index), values=tuple(cells))

def atualizar_saidas_tree():
    for item in tree_saidas.get_children():
        tree_saidas.delete(item)
    wb = carregar_planilha()
    if not wb:
        return
    ws = wb["Saidas"]
    for row in ws.iter_rows(min_row=2):
        cells = [c.value for c in row]
        if not cells or not cells[0]:
            continue
        excel_row_index = row[0].row
        tree_saidas.insert("", "end", iid=str(excel_row_index), values=tuple(cells))

def atualizar_tudo(filtro=""):
    atualizar_resumo_tree(filtro=filtro)
    atualizar_entradas_tree()
    atualizar_saidas_tree()

def excluir_registro_entrada():
    wb = carregar_planilha()
    if not wb:
        return
    sel = tree_entradas.selection()
    if not sel:
        messagebox.showwarning("Aviso", "Selecione a entrada a excluir (clique na linha).")
        return
    confirm = messagebox.askyesno("Confirmar", "Deseja excluir o registro de entrada selecionado?")
    if not confirm:
        return
    ws = wb["Entradas"]
    rows_to_delete = sorted([int(i) for i in sel], reverse=True)
    for r in rows_to_delete:
        ws.delete_rows(r, 1)
    wb.save(CAMINHO)
    messagebox.showinfo("Sucesso", "Registro(s) de entrada excluído(s).")
    atualizar_tudo()

def excluir_registro_saida():
    wb = carregar_planilha()
    if not wb:
        return
    sel = tree_saidas.selection()
    if not sel:
        messagebox.showwarning("Aviso", "Selecione a saída a excluir (clique na linha).")
        return
    confirm = messagebox.askyesno("Confirmar", "Deseja excluir o registro de saída selecionado?")
    if not confirm:
        return
    ws = wb["Saidas"]
    rows_to_delete = sorted([int(i) for i in sel], reverse=True)
    for r in rows_to_delete:
        ws.delete_rows(r, 1)
    wb.save(CAMINHO)
    messagebox.showinfo("Sucesso", "Registro(s) de saída excluído(s).")
    atualizar_tudo()

def limpar_campos_entrada():
    ent_codigo.delete(0, tk.END)
    ent_produto.delete(0, tk.END)
    ent_qtd_entrada.delete(0, tk.END)
    ent_fornecedor.delete(0, tk.END)
    ent_obs_entrada.delete(0, tk.END)

def limpar_campos_saida():
    ent_codigo.delete(0, tk.END)
    ent_produto.delete(0, tk.END)
    ent_qtd_saida.delete(0, tk.END)
    ent_destino.delete(0, tk.END)
    ent_obs_saida.delete(0, tk.END)

def gerar_relatorio_pdf():
    wb = carregar_planilha()
    if not wb:
        return
    data_inicio_txt = ent_data_inicio.get().strip()
    data_fim_txt = ent_data_fim.get().strip()
    if not data_inicio_txt or not data_fim_txt:
        messagebox.showwarning("Aviso", "Preencha data início e fim (dd/mm/aaaa).")
        return
    try:
        data_inicio = datetime.strptime(data_inicio_txt, DATE_FORMAT)
        data_fim = datetime.strptime(data_fim_txt, DATE_FORMAT)
    except Exception:
        messagebox.showerror("Erro", "Formato de data inválido. Use dd/mm/aaaa.")
        return
    out_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], title="Salvar relatório como")
    if not out_path:
        return
    entradas_filtradas = []
    saidas_filtradas = []
    for row in wb["Entradas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        dt = None
        try:
            dt = datetime.strptime(str(row[4]), DATE_FORMAT)
        except Exception:
            continue
        if data_inicio <= dt <= data_fim:
            entradas_filtradas.append(row)
    for row in wb["Saidas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        dt = None
        try:
            dt = datetime.strptime(str(row[4]), DATE_FORMAT)
        except Exception:
            continue
        if data_inicio <= dt <= data_fim:
            saidas_filtradas.append(row)
    entradas_ate = {}
    saidas_ate = {}
    for row in wb["Entradas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        try:
            dt = datetime.strptime(str(row[4]), DATE_FORMAT)
        except Exception:
            continue
        if dt <= data_fim:
            codigo = str(row[1])
            entradas_ate.setdefault(codigo, {"produto": row[2], "total": 0})
            entradas_ate[codigo]["total"] += safe_int(row[3])
    for row in wb["Saidas"].iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        try:
            dt = datetime.strptime(str(row[4]), DATE_FORMAT)
        except Exception:
            continue
        if dt <= data_fim:
            codigo = str(row[1])
            saidas_ate.setdefault(codigo, {"produto": row[2], "total": 0})
            saidas_ate[codigo]["total"] += safe_int(row[3])
    codigos = set(list(entradas_ate.keys()) + list(saidas_ate.keys()))
    resumo_atual = []
    for c in sorted(codigos):
        nome = entradas_ate.get(c, {}).get("produto") or saidas_ate.get(c, {}).get("produto") or ""
        ent = entradas_ate.get(c, {}).get("total", 0)
        sai = saidas_ate.get(c, {}).get("total", 0)
        saldo = ent - sai
        resumo_atual.append((c, str(nome), ent, sai, saldo))
    try:
        c = canvas.Canvas(out_path, pagesize=A4)
        largura, altura = A4
        margem_x = 40
        y = altura - 40
        c.setFont("Helvetica-Bold", 16)
        titulo = f"Relatório de Estoque: {data_inicio.strftime(DATE_FORMAT)} até {data_fim.strftime(DATE_FORMAT)}"
        c.drawString(margem_x, y, titulo)
        y -= 30
        c.setFont("Helvetica", 10)
        c.drawString(margem_x, y, f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
        y -= 25
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margem_x, y, "ENTRADAS")
        y -= 18
        c.setFont("Helvetica", 10)
        c.drawString(margem_x, y, "ID")
        c.drawString(margem_x + 50, y, "Código")
        c.drawString(margem_x + 150, y, "Produto")
        c.drawString(margem_x + 350, y, "Qtd")
        c.drawString(margem_x + 400, y, "Data")
        y -= 14
        c.line(margem_x, y, largura - margem_x, y)
        y -= 8
        total_entradas = 0
        for row in entradas_filtradas:
            if y < 80:
                c.showPage()
                y = altura - 40
            idr, codigo, produto, qtd, data_e, fornecedor, obs = row
            c.drawString(margem_x, y, str(idr))
            c.drawString(margem_x + 50, y, str(codigo))
            c.drawString(margem_x + 150, y, str(produto)[:28])
            c.drawString(margem_x + 350, y, str(qtd))
            c.drawString(margem_x + 400, y, str(data_e))
            y -= 14
            total_entradas += safe_int(qtd)
        y -= 10
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margem_x, y, f"Total de Entradas no período: {total_entradas}")
        y -= 20
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margem_x, y, "SAÍDAS")
        y -= 18
        c.setFont("Helvetica", 10)
        c.drawString(margem_x, y, "ID")
        c.drawString(margem_x + 50, y, "Código")
        c.drawString(margem_x + 150, y, "Produto")
        c.drawString(margem_x + 350, y, "Qtd")
        c.drawString(margem_x + 400, y, "Data")
        y -= 14
        c.line(margem_x, y, largura - margem_x, y)
        y -= 8
        total_saidas = 0
        for row in saidas_filtradas:
            if y < 80:
                c.showPage()
                y = altura - 40
            idr, codigo, produto, qtd, data_s, destino, obs = row
            c.drawString(margem_x, y, str(idr))
            c.drawString(margem_x + 50, y, str(codigo))
            c.drawString(margem_x + 150, y, str(produto)[:28])
            c.drawString(margem_x + 350, y, str(qtd))
            c.drawString(margem_x + 400, y, str(data_s))
            y -= 14
            total_saidas += safe_int(qtd)
        y -= 10
        c.setFont("Helvetica-Bold", 10)
        c.drawString(margem_x, y, f"Total de Saídas no período: {total_saidas}")
        y -= 20
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margem_x, y, "ESTOQUE ATUAL (até data fim)")
        y -= 18
        c.setFont("Helvetica", 10)
        c.drawString(margem_x, y, "Código")
        c.drawString(margem_x + 120, y, "Produto")
        c.drawString(margem_x + 340, y, "Entradas")
        c.drawString(margem_x + 420, y, "Saídas")
        c.drawString(margem_x + 480, y, "Saldo")
        y -= 14
        c.line(margem_x, y, largura - margem_x, y)
        y -= 8
        for cinfo in resumo_atual:
            if y < 80:
                c.showPage()
                y = altura - 40
            codigo, produto, ent, sai, saldo = cinfo
            c.drawString(margem_x, y, str(codigo))
            c.drawString(margem_x + 120, y, str(produto)[:28])
            c.drawString(margem_x + 340, y, str(ent))
            c.drawString(margem_x + 420, y, str(sai))
            c.drawString(margem_x + 480, y, str(saldo))
            y -= 14

        c.save()
        messagebox.showinfo("Relatório gerado", f"PDF salvo em:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha ao gerar PDF:\n{e}")

root = tk.Tk()
root.title("Almoxarifado v3.0")
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = int(screen_width * 0.8)
window_height = int(screen_height * 0.8)
x = (screen_width // 2) - (window_width // 2)
y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{x}+{y}")
root.minsize(900, 600)
root.resizable(True, True)

top_frame = tk.Frame(root)
top_frame.pack(fill="x", padx=8, pady=6)

btn_arquivo = tk.Button(top_frame, text="📁 Selecionar / Criar Planilha", command=escolher_arquivo)
btn_arquivo.pack(side="left")

lbl_min = tk.Label(top_frame, text="Estoque mínimo (global):")
lbl_min.pack(side="left", padx=(20, 4))
ent_min_global = tk.Entry(top_frame, width=6)
ent_min_global.insert(0, str(ESTOQUE_MINIMO_GLOBAL))
ent_min_global.pack(side="left")

def set_min_global():
    global ESTOQUE_MINIMO_GLOBAL
    try:
        v = int(ent_min_global.get().strip())
        ESTOQUE_MINIMO_GLOBAL = v
        atualizar_tudo()
    except Exception:
        messagebox.showerror("Erro", "Valor de estoque mínimo inválido.")
tk.Button(top_frame, text="Salvar mínimo", command=set_min_global).pack(side="left", padx=6)

middle = tk.Frame(root)
middle.pack(fill="both", expand=True, padx=8, pady=6)

left = tk.Frame(middle)
left.pack(side="left", fill="both", expand=True)

frame_entrada = tk.LabelFrame(left, text="Registrar Entrada")
frame_entrada.pack(fill="x", padx=6, pady=6)

tk.Label(frame_entrada, text="Código:").grid(row=0, column=0, sticky="w")
ent_codigo = tk.Entry(frame_entrada, width=15)
ent_codigo.grid(row=0, column=1, padx=4, pady=2)

tk.Label(frame_entrada, text="Produto:").grid(row=0, column=2, sticky="w")
ent_produto = tk.Entry(frame_entrada, width=30)
ent_produto.grid(row=0, column=3, padx=4, pady=2)

tk.Label(frame_entrada, text="Qtd:").grid(row=1, column=0, sticky="w")
ent_qtd_entrada = tk.Entry(frame_entrada, width=10)
ent_qtd_entrada.grid(row=1, column=1, padx=4, pady=2)

tk.Label(frame_entrada, text="Fornecedor:").grid(row=1, column=2, sticky="w")
ent_fornecedor = tk.Entry(frame_entrada, width=20)
ent_fornecedor.grid(row=1, column=3, padx=4, pady=2)

tk.Label(frame_entrada, text="Observação:").grid(row=2, column=0, sticky="w")
ent_obs_entrada = tk.Entry(frame_entrada, width=70)
ent_obs_entrada.grid(row=2, column=1, columnspan=3, padx=4, pady=2)

tk.Button(frame_entrada, text="➕ Adicionar Entrada", command=adicionar_entrada).grid(row=3, column=0, pady=6)
tk.Button(frame_entrada, text="🧹 Limpar Entrada", command=limpar_campos_entrada).grid(row=3, column=1, pady=6)

frame_saida = tk.LabelFrame(left, text="Registrar Saída")
frame_saida.pack(fill="x", padx=6, pady=6)

tk.Label(frame_saida, text="Código:").grid(row=0, column=0, sticky="w")
ent_codigo_s = tk.Entry(frame_saida, width=15)
ent_codigo_s.grid(row=0, column=1, padx=4, pady=2)

def sync_codigo_to_main(event=None):
    ent_codigo.delete(0, tk.END)
    ent_codigo.insert(0, ent_codigo_s.get())
ent_codigo_s.bind("<FocusOut>", sync_codigo_to_main)

tk.Label(frame_saida, text="Produto:").grid(row=0, column=2, sticky="w")
ent_produto_s = tk.Entry(frame_saida, width=30)
ent_produto_s.grid(row=0, column=3, padx=4, pady=2)
def sync_produto_to_main(event=None):
    ent_produto.delete(0, tk.END)
    ent_produto.insert(0, ent_produto_s.get())
ent_produto_s.bind("<FocusOut>", sync_produto_to_main)

tk.Label(frame_saida, text="Qtd:").grid(row=1, column=0, sticky="w")
ent_qtd_saida = tk.Entry(frame_saida, width=10)
ent_qtd_saida.grid(row=1, column=1, padx=4, pady=2)

tk.Label(frame_saida, text="Destino/Obra:").grid(row=1, column=2, sticky="w")
ent_destino = tk.Entry(frame_saida, width=20)
ent_destino.grid(row=1, column=3, padx=4, pady=2)

tk.Label(frame_saida, text="Observação:").grid(row=2, column=0, sticky="w")
ent_obs_saida = tk.Entry(frame_saida, width=70)
ent_obs_saida.grid(row=2, column=1, columnspan=3, padx=4, pady=2)

tk.Button(frame_saida, text="📦 Registrar Saída", command=registrar_saida).grid(row=3, column=0, pady=6)
tk.Button(frame_saida, text="🧹 Limpar Saída", command=limpar_campos_saida).grid(row=3, column=1, pady=6)

btns_del = tk.Frame(left)
btns_del.pack(fill="x", padx=6, pady=6)
tk.Button(btns_del, text="❌ Excluir Entrada Selecionada", command=excluir_registro_entrada).pack(side="left", padx=6)
tk.Button(btns_del, text="❌ Excluir Saída Selecionada", command=excluir_registro_saida).pack(side="left", padx=6)
tk.Button(btns_del, text="🔄 Atualizar Tudo", command=atualizar_tudo).pack(side="left", padx=6)

tree_frame = tk.Frame(left)
tree_frame.pack(fill="both", expand=True, padx=6, pady=6)

sub_left_top = tk.LabelFrame(tree_frame, text="Entradas (histórico)")
sub_left_top.pack(fill="both", expand=True, padx=4, pady=4)
cols_e = ("ID", "Código", "Produto", "Quantidade", "Data", "Fornecedor", "Obs")
tree_entradas = ttk.Treeview(sub_left_top, columns=cols_e, show="headings", selectmode="extended")

for c in cols_e:
    tree_entradas.heading(c, text=c)
    tree_entradas.column(c, width=100 if c in ("ID", "Quantidade", "Data") else 180)
scroll_y_e = ttk.Scrollbar(sub_left_top, orient="vertical", command=tree_entradas.yview)
scroll_x_e = ttk.Scrollbar(sub_left_top, orient="horizontal", command=tree_entradas.xview)
tree_entradas.configure(yscrollcommand=scroll_y_e.set, xscrollcommand=scroll_x_e.set)

scroll_y_e.pack(side="right", fill="y")
scroll_x_e.pack(side="bottom", fill="x")
tree_entradas.pack(fill="both", expand=True)

sub_left_bot = tk.LabelFrame(tree_frame, text="Saídas (histórico)")
sub_left_bot.pack(fill="both", expand=True, padx=4, pady=4)
cols_s = ("ID", "Código", "Produto", "Quantidade", "Data", "Destino", "Obs")
tree_saidas = ttk.Treeview(sub_left_bot, columns=cols_s, show="headings", selectmode="extended")

for c in cols_s:
    tree_saidas.heading(c, text=c)
    tree_saidas.column(c, width=100 if c in ("ID", "Quantidade", "Data") else 180)
scroll_y_s = ttk.Scrollbar(sub_left_bot, orient="vertical", command=tree_saidas.yview)
scroll_x_s = ttk.Scrollbar(sub_left_bot, orient="horizontal", command=tree_saidas.xview)
tree_saidas.configure(yscrollcommand=scroll_y_s.set, xscrollcommand=scroll_x_s.set)

scroll_y_s.pack(side="right", fill="y")
scroll_x_s.pack(side="bottom", fill="x")
tree_saidas.pack(fill="both", expand=True)

right = tk.Frame(middle)
right.pack(side="left", fill="both", expand=True, padx=6)

search_frame = tk.Frame(right)
search_frame.pack(fill="x", pady=6)
tk.Label(search_frame, text="Buscar (código ou produto):").pack(side="left")
ent_busca = tk.Entry(search_frame, width=30)
ent_busca.pack(side="left", padx=6)

def on_busca_key(event):
    filtro = ent_busca.get().strip()
    atualizar_resumo_tree(filtro=filtro)
ent_busca.bind("<KeyRelease>", on_busca_key)

frame_resumo = tk.LabelFrame(right, text="Resumo de Produtos")
frame_resumo.pack(fill="both", expand=True, padx=4, pady=4)
cols_r = ("Código", "Produto", "Entradas", "Saídas", "Saldo")
tree_resumo = ttk.Treeview(frame_resumo, columns=cols_r, show="headings")

for c in cols_r:
    tree_resumo.heading(c, text=c)
    tree_resumo.column(c, width=120 if c != "Produto" else 240)

scroll_y_r = ttk.Scrollbar(frame_resumo, orient="vertical", command=tree_resumo.yview)
scroll_x_r = ttk.Scrollbar(frame_resumo, orient="horizontal", command=tree_resumo.xview)
tree_resumo.configure(yscrollcommand=scroll_y_r.set, xscrollcommand=scroll_x_r.set)

scroll_y_r.pack(side="right", fill="y")
scroll_x_r.pack(side="bottom", fill="x")
tree_resumo.pack(fill="both", expand=True)

frame_rel = tk.LabelFrame(right, text="Relatório (PDF)")
frame_rel.pack(fill="x", padx=4, pady=6)
tk.Label(frame_rel, text="Data Início (dd/mm/aaaa):").grid(row=0, column=0, padx=4, pady=4)
ent_data_inicio = tk.Entry(frame_rel, width=12)
ent_data_inicio.grid(row=0, column=1, padx=4, pady=4)
tk.Label(frame_rel, text="Data Fim (dd/mm/aaaa):").grid(row=0, column=2, padx=4, pady=4)
ent_data_fim = tk.Entry(frame_rel, width=12)
ent_data_fim.grid(row=0, column=3, padx=4, pady=4)
tk.Button(frame_rel, text="📄 Gerar Relatório PDF", command=gerar_relatorio_pdf).grid(row=1, column=0, columnspan=4, pady=8)

status_var = tk.StringVar()
status_var.set("Pronto.")
status_bar = tk.Label(root, textvariable=status_var, anchor="w")
status_bar.pack(fill="x", side="bottom")

def inicializar_ui():
    if getattr(sys, "frozen", False):
        status_var.set("Executável rodando. Selecione/Crie a planilha para iniciar.")
    else:
        status_var.set("Selecione ou crie uma planilha (.xlsx).")
    atualizar_tudo()

inicializar_ui()
root.mainloop()