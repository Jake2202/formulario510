import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
import os, sys, sqlite3, json, threading, urllib.request

try:
    import openpyxl
    HAS_XLSX = True
except ImportError:
    HAS_XLSX = False

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                 Paragraph, Spacer, HRFlowable)
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT

# ── Colors ───────────────────────────────────────────────────────────────────
BLUE   = colors.HexColor("#1d4ed8")
LBLUE  = colors.HexColor("#dbeafe")
GRAY   = colors.HexColor("#f1f5f9")
DGRAY  = colors.HexColor("#475569")
BLACK  = colors.HexColor("#0f172a")
WHITE  = colors.white
BORDER = colors.HexColor("#cbd5e1")

# ── DB path ───────────────────────────────────────────────────────────────────
def get_db_path():
    base = os.path.dirname(sys.executable) if getattr(sys,"frozen",False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "formulario510.db")

# ── Database ──────────────────────────────────────────────────────────────────
def init_db():
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS clientes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nit TEXT, dv TEXT, razon_social TEXT, direccion TEXT,
        telefono TEXT, cod_seccional TEXT, cod_dpto TEXT, cod_municipio TEXT,
        creado TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS declaraciones (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        cliente_id INTEGER, numero TEXT, fecha TEXT,
        datos TEXT, total_cop REAL, estado TEXT DEFAULT "Borrador",
        pdf_path TEXT,
        creado TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(cliente_id) REFERENCES clientes(id)
    )''')
    conn.commit(); conn.close()

def db_exec(sql, params=()):
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor(); c.execute(sql, params); conn.commit()
    last = c.lastrowid; conn.close(); return last

def db_fetch(sql, params=()):
    conn = sqlite3.connect(get_db_path())
    c = conn.cursor(); c.execute(sql, params)
    rows = c.fetchall(); conn.close(); return rows

# ── Excel map ─────────────────────────────────────────────────────────────────
EXCEL_MAP = {
    "nit":"B5","dv":"D5","razonSocial":"B6","direccion":"B7","telefono":"D7",
    "codSeccional":"B8","codDpto":"D8","codMunicipio":"B9",
    "nitDecl":"B12","dvDecl":"D12","razonDecl":"B13","tipoUsuario":"D13",
    "nombresDecl":"B14","numDocDecl":"D14","codUsuario":"B15",
    "tipoDecl":"B18","numFormAnterior":"D18",
    "manifestoCarga":"B19","fechaLlegada":"D19",
    "codLugarIngreso":"B20","codModo":"D20",
    "docTransporte":"B21","fechaDocTransporte":"D21",
    "codProcedencia":"B22","tasaCambio":"D22",
    "nombreExportador":"B25","formaPago":"D25",
    "codPaisCompra":"B26","codPaisOrigen":"D26",
    "subpartida":"B27","numBultos":"D27",
    "cantidad":"B28","pesoBruto":"D28",
    "pesoNeto":"B29","fob":"D29",
    "fletes":"B30","seguros":"D30",
    "otrosGastos":"B31","ajuste":"D31",
    "descripcion":"B32",
    "arancelPct":"B35","ivaPct":"D35","icPct":"B36",
}

SUBPARTIDAS = [
    ("Smartphones / Celulares",    "8517.13.00.00", "0",  "19"),
    ("Tablets",                    "8471.30.00.00", "0",  "19"),
    ("Laptops / Computadores",     "8471.30.00.00", "0",  "19"),
    ("Computadores escritorio",    "8471.41.00.00", "0",  "19"),
    ("Audífonos / Auriculares",    "8518.30.00.00", "15", "19"),
    ("Cámaras fotográficas",       "9006.53.00.00", "15", "19"),
    ("Relojes de pulsera",         "9102.11.00.00", "15", "19"),
    ("Consolas videojuegos",       "9504.50.00.00", "15", "19"),
    ("Accesorios electrónicos",    "8544.42.00.00", "15", "19"),
    ("Ropa exterior hombre",       "6203.42.00.00", "40", "19"),
    ("Ropa exterior mujer",        "6204.62.00.00", "40", "19"),
    ("Calzado deportivo",          "6404.11.00.00", "40", "19"),
    ("Bolsos / Maletines",         "4202.12.00.00", "15", "19"),
    ("Perfumes / Cosméticos",      "3303.00.10.00", "15", "19"),
    ("Suplementos vitamínicos",    "2106.90.72.00", "15", "19"),
    ("Libros impresos",            "4901.99.00.00", "0",  "0"),
    ("Juguetes",                   "9503.00.90.00", "15", "19"),
    ("Motos hasta 185cc",          "8711.20.00.00", "0",  "19"),
    ("Motos más de 185cc",         "8711.40.00.00", "35", "19"),
    ("Autos hasta 1400cc",         "8703.22.90.00", "35", "19"),
    ("Autos 1400-2000cc",          "8703.23.90.00", "35", "19"),
    ("Vehículos eléctricos",       "8703.80.10.00", "5",  "19"),
]

def fmt_cop(n):
    try: return f"${int(round(float(n))):,}".replace(",",".")
    except: return "$0"

def calcular_dv(nit):
    try:
        nit_str = str(nit).strip().replace(".","").replace("-","")
        if not nit_str.isdigit(): return ""
        factores = [71,67,59,53,47,43,41,37,29,23,19,17,13,7,3,2]
        nit_pad = nit_str.zfill(15)
        total = sum(int(d)*f for d,f in zip(nit_pad, factores))
        residuo = total % 11
        dv = 11 - residuo if residuo > 1 else residuo
        return str(dv)
    except: return ""

# ── PDF ───────────────────────────────────────────────────────────────────────
def make_pdf(data, path):
    doc = SimpleDocTemplate(path, pagesize=A4,
                            topMargin=14*mm, bottomMargin=14*mm,
                            leftMargin=14*mm, rightMargin=14*mm)
    story = []
    s_title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15, textColor=BLUE, spaceAfter=2)
    s_sub   = ParagraphStyle("s", fontName="Helvetica",      fontSize=8,  textColor=DGRAY, spaceAfter=8)
    s_sec   = ParagraphStyle("sc",fontName="Helvetica-Bold", fontSize=7,  textColor=BLUE, spaceBefore=10, spaceAfter=4, leading=10)
    s_lbl   = ParagraphStyle("l", fontName="Helvetica",      fontSize=7,  textColor=DGRAY)
    s_val   = ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=9,  textColor=BLACK)
    s_total = ParagraphStyle("to",fontName="Helvetica-Bold", fontSize=13, textColor=BLUE, alignment=TA_RIGHT)
    s_foot  = ParagraphStyle("f", fontName="Helvetica-Oblique", fontSize=7, textColor=DGRAY, alignment=TA_CENTER)

    hd = Table([[Paragraph("Formulario <font color='#1d4ed8'>510</font>", s_title),
                 Paragraph(f"Año: <b>{date.today().year}</b>", s_val)]],
               colWidths=[120*mm, 60*mm])
    hd.setStyle(TableStyle([("ALIGN",(1,0),(1,0),"RIGHT"),("VALIGN",(0,0),(-1,-1),"MIDDLE")]))
    story.append(hd)
    story.append(Paragraph("Declaración de Importación Simplificada Privada · DIAN Colombia", s_sub))
    story.append(HRFlowable(width="100%", thickness=1.5, color=BLUE, spaceAfter=10))

    def section(title, rows):
        story.append(Paragraph(title, s_sec))
        cw = [45*mm, 45*mm, 45*mm, 47*mm]
        tdata, row = [], []
        for lbl, val in rows:
            row.append([Paragraph(lbl, s_lbl), Paragraph(str(val) if val else "—", s_val)])
            if len(row) == 4:
                tdata.append(row); row = []
        if row:
            while len(row) < 4: row.append([Paragraph("", s_lbl), Paragraph("", s_val)])
            tdata.append(row)
        if not tdata: return
        t = Table(tdata, colWidths=cw, hAlign="LEFT")
        t.setStyle(TableStyle([
            ("ROWBACKGROUNDS",(0,0),(-1,-1),[WHITE,GRAY]),
            ("BOX",(0,0),(-1,-1),0.5,BORDER),("INNERGRID",(0,0),(-1,-1),0.3,BORDER),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
            ("LEFTPADDING",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"TOP"),
        ]))
        story.append(t); story.append(Spacer(1,4))

    section("01 — Importador", [
        ("NIT", data.get("nit","")+"-"+data.get("dv","")),
        ("Razón social", data.get("razonSocial","")),
        ("Dirección", data.get("direccion","")),
        ("Teléfono", data.get("telefono","")),
        ("Cód. Seccional", data.get("codSeccional","")),
        ("Cód. Departamento", data.get("codDpto","")),
        ("Cód. Municipio", data.get("codMunicipio","")),
    ])
    section("02 — Declarante", [
        ("NIT Declarante", data.get("nitDecl","")+"-"+data.get("dvDecl","")),
        ("Razón social", data.get("razonDecl","")),
        ("Tipo usuario", data.get("tipoUsuario","")),
        ("No. documento", data.get("numDocDecl","")),
        ("Nombres", data.get("nombresDecl","")),
    ])
    section("03 — Transporte", [
        ("Tipo declaración", data.get("tipoDecl","")),
        ("Manifiesto carga", data.get("manifestoCarga","")),
        ("Fecha llegada", data.get("fechaLlegada","")),
        ("Lugar ingreso", data.get("codLugarIngreso","")),
        ("Doc. transporte", data.get("docTransporte","")),
        ("Fecha doc.", data.get("fechaDocTransporte","")),
        ("Modo transporte", data.get("codModo","")),
        ("País procedencia", data.get("codProcedencia","")),
        ("Tasa de cambio", data.get("tasaCambio","")+" COP/USD"),
    ])
    section("04 — Mercancía", [
        ("Proveedor", data.get("nombreExportador","")),
        ("País compra", data.get("codPaisCompra","")),
        ("País origen", data.get("codPaisOrigen","")),
        ("Forma de pago", data.get("formaPago","")),
        ("Subpartida", data.get("subpartida","")),
        ("No. bultos", data.get("numBultos","")),
        ("Cantidad", data.get("cantidad","")),
        ("Peso bruto (kg)", data.get("pesoBruto","")),
        ("Peso neto (kg)", data.get("pesoNeto","")),
        ("FOB (USD)", "$"+data.get("fob","")),
        ("Fletes (USD)", "$"+data.get("fletes","")),
        ("Seguros (USD)", "$"+data.get("seguros","")),
        ("Otros gastos (USD)", "$"+data.get("otrosGastos","")),
        ("Valor Aduana CIF", "$"+data.get("valorAduana","")+" USD"),
    ])
    story.append(Paragraph("Descripción de las Mercancías (Cas. 68)", s_sec))
    dt = Table([[Paragraph(data.get("descripcion","—"), s_val)]], colWidths=[182*mm])
    dt.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BORDER),("BACKGROUND",(0,0),(-1,-1),WHITE),
                             ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
                             ("LEFTPADDING",(0,0),(-1,-1),8)]))
    story.append(dt); story.append(Spacer(1,6))

    story.append(Paragraph("05 — Liquidación Tributaria", s_sec))
    fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
    seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
    adj=float(data.get("ajuste","0") or 0); trm=float(data.get("tasaCambio","4150") or 4150)
    ap=float(data.get("arancelPct","0") or 0); ip=float(data.get("ivaPct","19") or 19)
    icp=float(data.get("icPct","0") or 0)
    cif=fob+flt+seg+otr+adj; cifC=cif*trm; araC=cifC*(ap/100)
    ivaC=(cifC+araC)*(ip/100); icC=cifC*(icp/100); total=araC+ivaC+icC

    def lr(label, val, bold=False, blue=False):
        ls = ParagraphStyle("lr", fontName="Helvetica-Bold" if bold else "Helvetica",
                             fontSize=9, textColor=BLUE if blue else (BLACK if bold else DGRAY))
        vs = ParagraphStyle("vr", fontName="Helvetica-Bold" if bold else "Helvetica",
                             fontSize=9 if not blue else 13,
                             textColor=BLUE if blue else (BLACK if bold else DGRAY), alignment=TA_RIGHT)
        return [Paragraph(label, ls), Paragraph(val, vs)]

    liq = [
        lr(f"FOB", f"${fob:.2f} USD"), lr(f"+ Fletes", f"${flt:.2f} USD"),
        lr(f"+ Seguros", f"${seg:.2f} USD"), lr(f"+ Otros gastos", f"${otr:.2f} USD"),
        lr(f"+/- Ajuste", f"${adj:.2f} USD"),
        lr(f"Valor Aduana CIF", f"${cif:.2f} USD = {fmt_cop(cifC)} COP", bold=True),
        lr(f"Arancel ({ap}%)", f"{fmt_cop(araC)} COP"),
        lr(f"IVA ({ip}%) sobre CIF + Arancel", f"{fmt_cop(ivaC)} COP"),
        lr(f"Impuesto al consumo ({icp}%)", f"{fmt_cop(icC)} COP"),
        lr("TOTAL LIQUIDADO (Cas. 93)", f"{fmt_cop(total)} COP", bold=True, blue=True),
    ]
    lt = Table(liq, colWidths=[130*mm, 52*mm])
    lt.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),0.5,BORDER),("INNERGRID",(0,0),(-1,-1),0.3,BORDER),
        ("ROWBACKGROUNDS",(0,0),(-1,-1),[WHITE,GRAY]),
        ("BACKGROUND",(0,9),(1,9),LBLUE),("BACKGROUND",(0,5),(1,5),GRAY),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5),
        ("LEFTPADDING",(0,0),(-1,-1),8),("RIGHTPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(lt); story.append(Spacer(1,8))

    ct_data = [
        [Paragraph("Casilla 72 — Total Arancel $", s_lbl),
         Paragraph("Casilla 76 — Total IVA $", s_lbl),
         Paragraph("Casilla 980 — Pago total $", s_lbl)],
        [Paragraph(fmt_cop(araC), s_total),
         Paragraph(fmt_cop(ivaC), s_total),
         Paragraph(fmt_cop(total), s_total)],
    ]
    ct = Table(ct_data, colWidths=[60*mm, 60*mm, 62*mm])
    ct.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.5,BLUE),("INNERGRID",(0,0),(-1,-1),0.5,BORDER),
        ("BACKGROUND",(0,0),(-1,-1),LBLUE),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),10),("VALIGN",(0,0),(-1,-1),"MIDDLE"),
    ]))
    story.append(ct); story.append(Spacer(1,16))
    firma_data = [
        [Paragraph("Firma del declarante", s_lbl), Paragraph("Nombre completo", s_lbl), Paragraph("C.C. No.", s_lbl)],
        [Paragraph("_______________________", s_val), Paragraph("_______________________", s_val), Paragraph("_______________________", s_val)],
    ]
    ft = Table(firma_data, colWidths=[60*mm, 60*mm, 62*mm])
    ft.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
        ("LEFTPADDING",(0,0),(-1,-1),6),("VALIGN",(0,0),(-1,-1),"BOTTOM"),
        ("BOX",(0,0),(-1,-1),0.5,BORDER),("INNERGRID",(0,0),(-1,-1),0.3,BORDER),
    ]))
    story.append(ft); story.append(Spacer(1,8))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER)); story.append(Spacer(1,4))
    story.append(Paragraph(
        f"Pre-diligenciamiento de referencia · No reemplaza declaración oficial ante la DIAN · "
        f"Generado el {date.today().strftime('%d/%m/%Y')}", s_foot))
    doc.build(story)


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Gestión de Clientes
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaClientes(tk.Toplevel):
    def __init__(self, parent, on_select=None):
        super().__init__(parent)
        self.title("Gestión de Clientes")
        self.geometry("900x560")
        self.configure(bg="white")
        self.on_select = on_select
        self._build()
        self._cargar()

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg="#1d4ed8", height=48)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="👥  Gestión de Clientes / Importadores",
                 font=("Arial",13,"bold"), bg="#1d4ed8", fg="white").pack(side="left", padx=16, pady=10)

        # Toolbar
        tb = tk.Frame(self, bg="#f8fafc", height=46)
        tb.pack(fill="x"); tb.pack_propagate(False)
        for text, cmd, bg in [
            ("➕ Nuevo cliente",   self._nuevo,    "#1d4ed8"),
            ("✏️ Editar",          self._editar,   "#0f766e"),
            ("🗑️ Eliminar",        self._eliminar, "#dc2626"),
        ]:
            tk.Button(tb, text=text, font=("Arial",10,"bold"), bg=bg, fg="white",
                      relief="flat", padx=12, pady=6, cursor="hand2",
                      command=cmd).pack(side="left", padx=(8,2), pady=6)

        # Search
        sf = tk.Frame(tb, bg="#f8fafc"); sf.pack(side="right", padx=12, pady=8)
        tk.Label(sf, text="🔍", bg="#f8fafc", font=("Arial",12)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace("w", lambda *a: self._cargar())
        tk.Entry(sf, textvariable=self.search_var, font=("Arial",11), width=22,
                 relief="flat", highlightbackground="#cbd5e1", highlightthickness=1
                 ).pack(side="left", ipady=4)

        # Table
        cols = ("ID","NIT","Razón Social","Dirección","Teléfono","Declaraciones")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=18)
        widths = {"ID":40,"NIT":110,"Razón Social":280,"Dirección":200,"Teléfono":110,"Declaraciones":90}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree.bind("<Double-1>", lambda e: self._seleccionar())

        if self.on_select:
            tk.Button(self, text="✅  Seleccionar cliente",
                      font=("Arial",11,"bold"), bg="#1d4ed8", fg="white",
                      relief="flat", pady=10, cursor="hand2",
                      command=self._seleccionar).pack(fill="x", padx=8, pady=(0,8))

    def _cargar(self):
        q = self.search_var.get().strip()
        if q:
            rows = db_fetch("SELECT id,nit,razon_social,direccion,telefono FROM clientes WHERE razon_social LIKE ? OR nit LIKE ? ORDER BY razon_social", (f"%{q}%",f"%{q}%"))
        else:
            rows = db_fetch("SELECT id,nit,razon_social,direccion,telefono FROM clientes ORDER BY razon_social")
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            cnt = db_fetch("SELECT COUNT(*) FROM declaraciones WHERE cliente_id=?", (r[0],))[0][0]
            self.tree.insert("", "end", values=(r[0],r[1],r[2],r[3],r[4],cnt))

    def _get_sel(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso","Seleccione un cliente primero."); return None
        return self.tree.item(sel[0])["values"]

    def _nuevo(self): FormCliente(self, callback=self._cargar)
    def _editar(self):
        v = self._get_sel()
        if v: FormCliente(self, cliente_id=v[0], callback=self._cargar)

    def _eliminar(self):
        v = self._get_sel()
        if not v: return
        cnt = db_fetch("SELECT COUNT(*) FROM declaraciones WHERE cliente_id=?", (v[0],))[0][0]
        msg = f"¿Eliminar a {v[2]}?"
        if cnt > 0: msg += f"\n\n⚠️ Tiene {cnt} declaración(es) registrada(s). También se eliminarán."
        if messagebox.askyesno("Confirmar", msg):
            db_exec("DELETE FROM declaraciones WHERE cliente_id=?", (v[0],))
            db_exec("DELETE FROM clientes WHERE id=?", (v[0],))
            self._cargar()

    def _seleccionar(self):
        v = self._get_sel()
        if v and self.on_select:
            row = db_fetch("SELECT * FROM clientes WHERE id=?", (v[0],))
            if row: self.on_select(row[0]); self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# FORMULARIO: Nuevo / Editar Cliente
# ═══════════════════════════════════════════════════════════════════════════════
class FormCliente(tk.Toplevel):
    def __init__(self, parent, cliente_id=None, callback=None):
        super().__init__(parent)
        self.cliente_id = cliente_id
        self.callback = callback
        self.title("Nuevo Cliente" if not cliente_id else "Editar Cliente")
        self.geometry("500x440")
        self.configure(bg="white")
        self.resizable(False, False)
        self.fields = {}
        self._build()
        if cliente_id: self._cargar()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="Datos del Importador / Cliente",
                 font=("Arial",12,"bold"), bg="white", fg="#1d4ed8").pack(pady=(16,12))
        form = tk.Frame(self, bg="white"); form.pack(fill="x", padx=24)
        campos = [
            ("NIT (sin DV)","nit",0,0), ("DV","dv",0,1),
            ("Razón social / Nombre completo","razonSocial",1,0,2),
            ("Dirección","direccion",2,0,2),
            ("Teléfono","telefono",3,0), ("Cód. Seccional","codSeccional",3,1),
            ("Cód. Departamento","codDpto",4,0), ("Cód. Municipio","codMunicipio",4,1),
        ]
        for i in range(2): form.columnconfigure(i, weight=1)
        for item in campos:
            label, key, row, col = item[0], item[1], item[2], item[3]
            span = item[4] if len(item)>4 else 1
            fc = tk.Frame(form, bg="white")
            fc.grid(row=row, column=col, columnspan=span, sticky="ew", padx=4, pady=4)
            tk.Label(fc, text=label, font=("Arial",9), bg="white", fg="#64748b").pack(anchor="w")
            w = tk.Entry(fc, font=("Arial",11), relief="flat", bg="#f8fafc",
                         highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x", ipady=5)
            if key == "nit":
                w.bind("<FocusOut>", lambda e: self._auto_dv())
            self.fields[key] = w

        btns = tk.Frame(self, bg="white"); btns.pack(fill="x", padx=24, pady=16)
        tk.Button(btns, text="💾  Guardar", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10, cursor="hand2",
                  command=self._guardar).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="Cancelar", font=("Arial",11), bg="#f1f5f9", fg="#64748b",
                  relief="flat", pady=10, cursor="hand2",
                  command=self.destroy).pack(side="left", fill="x", expand=True)

    def _auto_dv(self):
        nit = self.fields["nit"].get().strip()
        if len(nit) >= 6:
            dv = calcular_dv(nit)
            if dv:
                self.fields["dv"].delete(0,"end")
                self.fields["dv"].insert(0,dv)

    def _cargar(self):
        row = db_fetch("SELECT * FROM clientes WHERE id=?", (self.cliente_id,))
        if not row: return
        r = row[0]
        vals = {"nit":r[1],"dv":r[2],"razonSocial":r[3],"direccion":r[4],
                "telefono":r[5],"codSeccional":r[6],"codDpto":r[7],"codMunicipio":r[8]}
        for k,v in vals.items():
            w = self.fields.get(k)
            if w: w.delete(0,"end"); w.insert(0,v or "")

    def _guardar(self):
        nit = self.fields["nit"].get().strip()
        razon = self.fields["razonSocial"].get().strip()
        if not nit or not razon:
            messagebox.showwarning("Aviso","NIT y Razón social son obligatorios."); return
        vals = (nit, self.fields["dv"].get(), razon, self.fields["direccion"].get(),
                self.fields["telefono"].get(), self.fields["codSeccional"].get(),
                self.fields["codDpto"].get(), self.fields["codMunicipio"].get())
        if self.cliente_id:
            db_exec("UPDATE clientes SET nit=?,dv=?,razon_social=?,direccion=?,telefono=?,cod_seccional=?,cod_dpto=?,cod_municipio=? WHERE id=?",
                    vals+(self.cliente_id,))
        else:
            db_exec("INSERT INTO clientes(nit,dv,razon_social,direccion,telefono,cod_seccional,cod_dpto,cod_municipio) VALUES(?,?,?,?,?,?,?,?)", vals)
        if self.callback: self.callback()
        self.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Historial de Declaraciones
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaHistorial(tk.Toplevel):
    def __init__(self, parent, app_ref):
        super().__init__(parent)
        self.app = app_ref
        self.title("Historial de Declaraciones")
        self.geometry("1000x580")
        self.configure(bg="white")
        self._build()
        self._cargar()

    def _build(self):
        hdr = tk.Frame(self, bg="#1d4ed8", height=48)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="📋  Historial de Declaraciones",
                 font=("Arial",13,"bold"), bg="#1d4ed8", fg="white").pack(side="left", padx=16, pady=10)

        tb = tk.Frame(self, bg="#f8fafc", height=46); tb.pack(fill="x"); tb.pack_propagate(False)
        for text, cmd, bg in [
            ("📂 Abrir",    self._abrir,    "#0f766e"),
            ("🗑️ Eliminar", self._eliminar, "#dc2626"),
            ("📄 Ver PDF",  self._ver_pdf,  "#7c3aed"),
        ]:
            tk.Button(tb, text=text, font=("Arial",10,"bold"), bg=bg, fg="white",
                      relief="flat", padx=12, pady=6, cursor="hand2",
                      command=cmd).pack(side="left", padx=(8,2), pady=6)

        # Filtro cliente
        ff = tk.Frame(tb, bg="#f8fafc"); ff.pack(side="right", padx=12, pady=8)
        tk.Label(ff, text="Cliente:", bg="#f8fafc", font=("Arial",10)).pack(side="left", padx=(0,4))
        self.filtro_var = tk.StringVar(value="Todos")
        self.filtro_cb = ttk.Combobox(ff, textvariable=self.filtro_var, font=("Arial",10),
                                       state="readonly", width=24)
        self.filtro_cb.pack(side="left")
        self.filtro_cb.bind("<<ComboboxSelected>>", lambda e: self._cargar())
        self._actualizar_filtro()

        cols = ("ID","Fecha","Cliente","NIT","No. Doc. Transporte","Total COP","Estado")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=20)
        widths = {"ID":40,"Fecha":90,"Cliente":260,"NIT":110,"No. Doc. Transporte":160,"Total COP":110,"Estado":80}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w")
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree.bind("<Double-1>", lambda e: self._abrir())

    def _actualizar_filtro(self):
        clientes = db_fetch("SELECT razon_social FROM clientes ORDER BY razon_social")
        vals = ["Todos"] + [r[0] for r in clientes]
        self.filtro_cb["values"] = vals

    def _cargar(self):
        filtro = self.filtro_var.get()
        if filtro == "Todos":
            rows = db_fetch("""SELECT d.id, d.fecha, c.razon_social, c.nit, d.numero, d.total_cop, d.estado
                               FROM declaraciones d JOIN clientes c ON d.cliente_id=c.id
                               ORDER BY d.creado DESC""")
        else:
            rows = db_fetch("""SELECT d.id, d.fecha, c.razon_social, c.nit, d.numero, d.total_cop, d.estado
                               FROM declaraciones d JOIN clientes c ON d.cliente_id=c.id
                               WHERE c.razon_social=? ORDER BY d.creado DESC""", (filtro,))
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            total = fmt_cop(r[5]) if r[5] else "$0"
            self.tree.insert("", "end", values=(r[0],r[1],r[2],r[3],r[4] or "—",total,r[6]))

    def _get_sel_id(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso","Seleccione una declaración."); return None
        return self.tree.item(sel[0])["values"][0]

    def _abrir(self):
        did = self._get_sel_id()
        if not did: return
        row = db_fetch("SELECT datos FROM declaraciones WHERE id=?", (did,))
        if not row: return
        data = json.loads(row[0][0])
        self.app._cargar_datos(data)
        self.app._decl_id = did
        self.destroy()
        messagebox.showinfo("Cargado","✅ Declaración cargada en el formulario.")

    def _eliminar(self):
        did = self._get_sel_id()
        if not did: return
        if messagebox.askyesno("Confirmar","¿Eliminar esta declaración?"):
            db_exec("DELETE FROM declaraciones WHERE id=?", (did,))
            self._cargar()

    def _ver_pdf(self):
        did = self._get_sel_id()
        if not did: return
        row = db_fetch("SELECT pdf_path FROM declaraciones WHERE id=?", (did,))
        if row and row[0][0] and os.path.exists(row[0][0]):
            os.startfile(row[0][0])
        else:
            messagebox.showinfo("Sin PDF","No hay PDF guardado para esta declaración.")


# ═══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("DeclaraFácil 510 — Agencia de Aduanas")
        self.geometry("1150x800")
        self.minsize(980, 680)
        self.configure(bg="#f1f5f9")
        self.resizable(True, True)
        self.fields = {}
        self._sections = []
        self._nav_btns = []
        self._cliente_id = None
        self._decl_id = None
        init_db()
        self._build_ui()
        self._set_defaults()

    def _build_ui(self):
        # ── Top bar ──
        top = tk.Frame(self, bg="#1d4ed8", height=54)
        top.pack(fill="x"); top.pack_propagate(False)
        tk.Label(top, text="DeclaraFácil 510", font=("Arial",16,"bold"),
                 bg="#1d4ed8", fg="white").pack(side="left", padx=20, pady=10)
        tk.Label(top, text="Declaración de Importación Simplificada · DIAN Colombia",
                 font=("Arial",10), bg="#1d4ed8", fg="#bfdbfe").pack(side="left")

        # Cliente badge
        self.lbl_cliente_badge = tk.Label(top, text="Sin cliente seleccionado",
                                           font=("Arial",9), bg="#1e3a5f", fg="#93c5fd",
                                           padx=10, pady=4)
        self.lbl_cliente_badge.pack(side="right", padx=16, pady=10)

        body = tk.Frame(self, bg="#f1f5f9"); body.pack(fill="both", expand=True)

        # ── Sidebar ──
        sb = tk.Frame(body, bg="#0f1724", width=220)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)

        tk.Label(sb, text="SECCIONES", font=("Arial",9,"bold"),
                 bg="#0f1724", fg="#1e3a5f").pack(pady=(18,6), padx=16, anchor="w")

        for i,(sid,label) in enumerate([
            ("s0","01  Importador"),("s1","02  Declarante"),
            ("s2","03  Transporte"),("s3","04  Mercancía"),("s4","05  Liquidación")]):
            b = tk.Button(sb, text=label, font=("Arial",10), anchor="w",
                          bg="#0f1724", fg="#4a6a8a", relief="flat",
                          activebackground="#1a2535", activeforeground="#e0eaf4",
                          bd=0, padx=12, pady=5, cursor="hand2",
                          command=lambda idx=i: self._jump(idx))
            b.pack(fill="x", padx=8, pady=1)
            self._nav_btns.append(b)

        # Total box
        tf = tk.Frame(sb, bg="#070d15"); tf.pack(fill="x", padx=10, pady=(8,2))
        tk.Label(tf, text="TOTAL A PAGAR", font=("Arial",8,"bold"),
                 bg="#070d15", fg="#1e3a5f").pack(anchor="w", padx=10, pady=(4,0))
        self.lbl_total = tk.Label(tf, text="$0", font=("Arial",18,"bold"),
                                   bg="#070d15", fg="#3b82f6")
        self.lbl_total.pack(anchor="w", padx=10)
        tk.Label(tf, text="COP", font=("Arial",8), bg="#070d15", fg="#2a4060").pack(anchor="w", padx=10, pady=(0,2))
        for lbl, attr in [("Arancel","lbl_ara"),("IVA","lbl_iva"),("Imp. Consumo","lbl_ic")]:
            row = tk.Frame(tf, bg="#070d15"); row.pack(fill="x", padx=10, pady=1)
            tk.Label(row, text=lbl, font=("Arial",9), bg="#070d15", fg="#1e3a5f").pack(side="left")
            w = tk.Label(row, text="$0", font=("Arial",9), bg="#070d15", fg="#4a6a8a"); w.pack(side="right")
            setattr(self, attr, w)
        tk.Frame(tf, bg="#070d15", height=6).pack()

        # TRM label
        self.lbl_trm_status = tk.Label(sb, text="", font=("Arial",8),
                                        bg="#0f1724", fg="#4a6a8a", wraplength=190)
        self.lbl_trm_status.pack(padx=10, pady=(2,2))

        # Action buttons
        tk.Label(sb, text="ACCIONES", font=("Arial",8,"bold"),
                 bg="#0f1724", fg="#1e3a5f").pack(pady=(6,2), padx=16, anchor="w")

        for text, cmd, bg, fg in [
            ("👥  Clientes",          self._abrir_clientes,   "#1e3a5f", "white"),
            ("📋  Historial",          self._abrir_historial,  "#1e3a5f", "white"),
            ("🔄  Actualizar TRM",     self._update_trm,       "#7c3aed", "white"),
            ("📥  Cargar Excel",       self._load_excel,       "#0f766e", "white"),
            ("📋  Subpartidas",        self._subpartidas,      "#92400e", "white"),
            ("💾  Guardar decl.",      self._guardar_decl,     "#0f4c35", "white"),
            ("📄  Generar PDF",        self._generate,         "#3b82f6", "white"),
            ("🗑️  Limpiar",            self._clear,            "#1e2535", "#64748b"),
        ]:
            tk.Button(sb, text=text, font=("Arial",9,"bold"), bg=bg, fg=fg,
                      relief="flat", bd=0, padx=8, pady=5, cursor="hand2",
                      activebackground=bg, activeforeground=fg,
                      command=cmd).pack(fill="x", padx=8, pady=1)

        # ── Scrollable main ──
        cf = tk.Frame(body, bg="#f1f5f9"); cf.pack(side="left", fill="both", expand=True)
        self.canvas = tk.Canvas(cf, bg="#f1f5f9", highlightthickness=0)
        vbar = ttk.Scrollbar(cf, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.inner = tk.Frame(self.canvas, bg="#f1f5f9")
        self._win = self.canvas.create_window((0,0), window=self.inner, anchor="nw")
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfig(self._win, width=e.width))
        self.canvas.bind_all("<MouseWheel>", lambda e: self.canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
        self._build_form()

    def _card(self, title):
        outer = tk.Frame(self.inner, bg="#f1f5f9")
        outer.pack(fill="x", padx=20, pady=(10,0))
        self._sections.append(outer)
        fr = tk.Frame(outer, bg="white", highlightbackground="#e2e8f0", highlightthickness=1)
        fr.pack(fill="x")
        tk.Frame(fr, bg="#1d4ed8", height=3).pack(fill="x")
        hdr = tk.Frame(fr, bg="white"); hdr.pack(fill="x", padx=18, pady=(8,4))
        tk.Label(hdr, text=title, font=("Arial",11,"bold"), bg="white", fg="#1d4ed8").pack(side="left")
        c = tk.Frame(fr, bg="white"); c.pack(fill="x", padx=18, pady=(0,14))
        return c

    def _field(self, parent, label, key, row, col, colspan=1, widget="entry", opts=None, width=20):
        fc = tk.Frame(parent, bg="white")
        fc.grid(row=row, column=col, columnspan=colspan, sticky="ew", padx=4, pady=4)
        for c in range(col, col+colspan): parent.columnconfigure(c, weight=1)
        tk.Label(fc, text=label, font=("Arial",9), bg="white", fg="#64748b").pack(anchor="w")
        if widget == "entry":
            w = tk.Entry(fc, font=("Arial",12), relief="flat", bd=0, bg="#f8fafc",
                         width=width, highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x", ipady=6)
            w.bind("<KeyRelease>", lambda e: self._calc())
            if key == "nit":
                w.bind("<FocusOut>", self._auto_dv)
                w.bind("<KeyRelease>", lambda e: (self._calc(), self._auto_dv()))
        elif widget == "combo":
            w = ttk.Combobox(fc, values=[o[1] for o in opts], font=("Arial",12),
                              state="readonly", width=width-2)
            w.set(opts[0][1]); w._opts = opts
            w.bind("<<ComboboxSelected>>", lambda e: self._calc())
            w.pack(fill="x")
        elif widget == "text":
            w = tk.Text(fc, font=("Arial",12), relief="flat", bd=0, bg="#f8fafc",
                        height=3, width=width, highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x")
            w.bind("<KeyRelease>", lambda e: self._calc())
        self.fields[key] = w

    def _build_form(self):
        F = self._field

        c = self._card("01 — Importador")
        F(c,"NIT (sin DV)","nit",0,0); F(c,"DV","dv",0,1)
        F(c,"Razón social / Nombres y apellidos","razonSocial",0,2,colspan=2,width=40)
        F(c,"Dirección","direccion",1,0,colspan=2,width=40)
        F(c,"Teléfono","telefono",1,2)
        F(c,"Cód. Seccional","codSeccional",1,3,widget="combo",opts=[
            ("18","18 — San Andrés"),("11","11 — Bogotá"),("08","08 — Barranquilla"),
            ("13","13 — Cartagena"),("76","76 — Cali"),("05","05 — Medellín")])
        F(c,"Cód. Departamento DANE","codDpto",2,0); F(c,"Cód. Ciudad DANE","codMunicipio",2,1)

        c = self._card("02 — Declarante Autorizado")
        F(c,"NIT Declarante","nitDecl",0,0); F(c,"DV","dvDecl",0,1)
        F(c,"Razón social declarante","razonDecl",0,2,colspan=2,width=40)
        F(c,"Tipo usuario","tipoUsuario",1,0); F(c,"Cód. usuario DIAN","codUsuario",1,1)
        F(c,"No. Documento","numDocDecl",1,2); F(c,"Apellidos y nombres","nombresDecl",1,3)

        c = self._card("03 — Manifiesto y Transporte")
        F(c,"Tipo declaración","tipoDecl",0,0,widget="combo",opts=[
            ("1","1 — Inicial"),("2","2 — Legalización"),("3","3 — Anticipada"),
            ("4","4 — Corrección"),("5","5 — Modificación")])
        F(c,"No. Form. anterior","numFormAnterior",0,1)
        F(c,"Manifiesto de carga","manifestoCarga",0,2)
        F(c,"Fecha llegada (AAAA-MM-DD)","fechaLlegada",0,3)
        F(c,"Lugar ingreso","codLugarIngreso",1,0,widget="combo",opts=[
            ("ADZ","ADZ — San Andrés"),("BOG","BOG — Bogotá"),("CTG","CTG — Cartagena"),
            ("BAQ","BAQ — Barranquilla"),("CLO","CLO — Cali"),("MDE","MDE — Medellín"),
            ("BUN","BUN — Buenaventura"),("SMR","SMR — Santa Marta")])
        F(c,"Modo transporte","codModo",1,1,widget="combo",opts=[
            ("4","4 — Aéreo"),("5","5 — Postal"),("1","1 — Marítimo"),("7","7 — Terrestre")])
        F(c,"Doc. transporte (AWB/BL/Guía)","docTransporte",1,2)
        F(c,"Fecha doc. transporte","fechaDocTransporte",1,3)
        F(c,"País procedencia","codProcedencia",2,0,widget="combo",opts=[('AF', 'AF — Afganistán'), ('AL', 'AL — Albania'), ('DE', 'DE — Alemania'), ('AD', 'AD — Andorra'), ('AO', 'AO — Angola'), ('AG', 'AG — Antigua y Barbuda'), ('SA', 'SA — Arabia Saudita'), ('DZ', 'DZ — Argelia'), ('AR', 'AR — Argentina'), ('AM', 'AM — Armenia'), ('AU', 'AU — Australia'), ('AT', 'AT — Austria'), ('AZ', 'AZ — Azerbaiyán'), ('BS', 'BS — Bahamas'), ('BD', 'BD — Bangladesh'), ('BB', 'BB — Barbados'), ('BH', 'BH — Baréin'), ('BE', 'BE — Bélgica'), ('BZ', 'BZ — Belice'), ('BJ', 'BJ — Benín'), ('BY', 'BY — Bielorrusia'), ('BO', 'BO — Bolivia'), ('BA', 'BA — Bosnia y Herzegovina'), ('BW', 'BW — Botsuana'), ('BR', 'BR — Brasil'), ('BN', 'BN — Brunéi'), ('BG', 'BG — Bulgaria'), ('BF', 'BF — Burkina Faso'), ('BI', 'BI — Burundi'), ('BT', 'BT — Bután'), ('CV', 'CV — Cabo Verde'), ('KH', 'KH — Camboya'), ('CM', 'CM — Camerún'), ('CA', 'CA — Canadá'), ('QA', 'QA — Catar'), ('TD', 'TD — Chad'), ('CL', 'CL — Chile'), ('CN', 'CN — China'), ('CY', 'CY — Chipre'), ('CO', 'CO — Colombia'), ('KM', 'KM — Comoras'), ('CG', 'CG — Congo'), ('CD', 'CD — Congo (RDC)'), ('KP', 'KP — Corea del Norte'), ('KR', 'KR — Corea del Sur'), ('CI', 'CI — Costa de Marfil'), ('CR', 'CR — Costa Rica'), ('HR', 'HR — Croacia'), ('CU', 'CU — Cuba'), ('DK', 'DK — Dinamarca'), ('DJ', 'DJ — Djibouti'), ('DM', 'DM — Dominica'), ('EC', 'EC — Ecuador'), ('EG', 'EG — Egipto'), ('SV', 'SV — El Salvador'), ('AE', 'AE — Emiratos Árabes Unidos'), ('ER', 'ER — Eritrea'), ('SK', 'SK — Eslovaquia'), ('SI', 'SI — Eslovenia'), ('ES', 'ES — España'), ('US', 'US — Estados Unidos'), ('EE', 'EE — Estonia'), ('ET', 'ET — Etiopía'), ('PH', 'PH — Filipinas'), ('FI', 'FI — Finlandia'), ('FJ', 'FJ — Fiyi'), ('FR', 'FR — Francia'), ('GA', 'GA — Gabón'), ('GM', 'GM — Gambia'), ('GE', 'GE — Georgia'), ('GH', 'GH — Ghana'), ('GD', 'GD — Granada'), ('GR', 'GR — Grecia'), ('GT', 'GT — Guatemala'), ('GN', 'GN — Guinea'), ('GW', 'GW — Guinea-Bisáu'), ('GQ', 'GQ — Guinea Ecuatorial'), ('GY', 'GY — Guyana'), ('HT', 'HT — Haití'), ('HN', 'HN — Honduras'), ('HU', 'HU — Hungría'), ('IN', 'IN — India'), ('ID', 'ID — Indonesia'), ('IQ', 'IQ — Irak'), ('IR', 'IR — Irán'), ('IE', 'IE — Irlanda'), ('IS', 'IS — Islandia'), ('MH', 'MH — Islas Marshall'), ('SB', 'SB — Islas Salomón'), ('IL', 'IL — Israel'), ('IT', 'IT — Italia'), ('JM', 'JM — Jamaica'), ('JP', 'JP — Japón'), ('JO', 'JO — Jordania'), ('KZ', 'KZ — Kazajistán'), ('KE', 'KE — Kenia'), ('KG', 'KG — Kirguistán'), ('KI', 'KI — Kiribati'), ('KW', 'KW — Kuwait'), ('LA', 'LA — Laos'), ('LS', 'LS — Lesoto'), ('LV', 'LV — Letonia'), ('LB', 'LB — Líbano'), ('LR', 'LR — Liberia'), ('LY', 'LY — Libia'), ('LI', 'LI — Liechtenstein'), ('LT', 'LT — Lituania'), ('LU', 'LU — Luxemburgo'), ('MK', 'MK — Macedonia del Norte'), ('MG', 'MG — Madagascar'), ('MY', 'MY — Malasia'), ('MW', 'MW — Malaui'), ('MV', 'MV — Maldivas'), ('ML', 'ML — Malí'), ('MT', 'MT — Malta'), ('MA', 'MA — Marruecos'), ('MU', 'MU — Mauricio'), ('MR', 'MR — Mauritania'), ('MX', 'MX — México'), ('FM', 'FM — Micronesia'), ('MD', 'MD — Moldavia'), ('MC', 'MC — Mónaco'), ('MN', 'MN — Mongolia'), ('ME', 'ME — Montenegro'), ('MZ', 'MZ — Mozambique'), ('MM', 'MM — Myanmar'), ('NA', 'NA — Namibia'), ('NR', 'NR — Nauru'), ('NP', 'NP — Nepal'), ('NI', 'NI — Nicaragua'), ('NE', 'NE — Níger'), ('NG', 'NG — Nigeria'), ('NO', 'NO — Noruega'), ('NZ', 'NZ — Nueva Zelanda'), ('OM', 'OM — Omán'), ('NL', 'NL — Países Bajos'), ('PK', 'PK — Pakistán'), ('PW', 'PW — Palaos'), ('PA', 'PA — Panamá'), ('PG', 'PG — Papúa Nueva Guinea'), ('PY', 'PY — Paraguay'), ('PE', 'PE — Perú'), ('PL', 'PL — Polonia'), ('PT', 'PT — Portugal'), ('GB', 'GB — Reino Unido'), ('CF', 'CF — República Centroafricana'), ('DO', 'DO — República Dominicana'), ('RW', 'RW — Ruanda'), ('RO', 'RO — Rumania'), ('RU', 'RU — Rusia'), ('WS', 'WS — Samoa'), ('LC', 'LC — Santa Lucía'), ('VC', 'VC — San Vicente y las Granadinas'), ('KN', 'KN — San Cristóbal y Nieves'), ('SM', 'SM — San Marino'), ('ST', 'ST — Santo Tomé y Príncipe'), ('SN', 'SN — Senegal'), ('RS', 'RS — Serbia'), ('SC', 'SC — Seychelles'), ('SL', 'SL — Sierra Leona'), ('SG', 'SG — Singapur'), ('SY', 'SY — Siria'), ('SO', 'SO — Somalia'), ('LK', 'LK — Sri Lanka'), ('SZ', 'SZ — Suazilandia'), ('ZA', 'ZA — Sudáfrica'), ('SD', 'SD — Sudán'), ('SS', 'SS — Sudán del Sur'), ('SE', 'SE — Suecia'), ('CH', 'CH — Suiza'), ('SR', 'SR — Surinam'), ('TH', 'TH — Tailandia'), ('TZ', 'TZ — Tanzania'), ('TJ', 'TJ — Tayikistán'), ('TL', 'TL — Timor Oriental'), ('TG', 'TG — Togo'), ('TO', 'TO — Tonga'), ('TT', 'TT — Trinidad y Tobago'), ('TN', 'TN — Túnez'), ('TM', 'TM — Turkmenistán'), ('TR', 'TR — Turquía'), ('TV', 'TV — Tuvalu'), ('UA', 'UA — Ucrania'), ('UG', 'UG — Uganda'), ('UY', 'UY — Uruguay'), ('UZ', 'UZ — Uzbekistán'), ('VU', 'VU — Vanuatu'), ('VE', 'VE — Venezuela'), ('VN', 'VN — Vietnam'), ('YE', 'YE — Yemen'), ('ZM', 'ZM — Zambia'), ('ZW', 'ZW — Zimbabue')])
        F(c,"Tasa de cambio COP/USD","tasaCambio",2,1)

        c = self._card("04 — Mercancía")
        F(c,"Nombre exportador / proveedor","nombreExportador",0,0,colspan=2,width=40)
        F(c,"País compra","codPaisCompra",0,2,widget="combo",opts=[('AF', 'AF — Afganistán'), ('AL', 'AL — Albania'), ('DE', 'DE — Alemania'), ('AD', 'AD — Andorra'), ('AO', 'AO — Angola'), ('AG', 'AG — Antigua y Barbuda'), ('SA', 'SA — Arabia Saudita'), ('DZ', 'DZ — Argelia'), ('AR', 'AR — Argentina'), ('AM', 'AM — Armenia'), ('AU', 'AU — Australia'), ('AT', 'AT — Austria'), ('AZ', 'AZ — Azerbaiyán'), ('BS', 'BS — Bahamas'), ('BD', 'BD — Bangladesh'), ('BB', 'BB — Barbados'), ('BH', 'BH — Baréin'), ('BE', 'BE — Bélgica'), ('BZ', 'BZ — Belice'), ('BJ', 'BJ — Benín'), ('BY', 'BY — Bielorrusia'), ('BO', 'BO — Bolivia'), ('BA', 'BA — Bosnia y Herzegovina'), ('BW', 'BW — Botsuana'), ('BR', 'BR — Brasil'), ('BN', 'BN — Brunéi'), ('BG', 'BG — Bulgaria'), ('BF', 'BF — Burkina Faso'), ('BI', 'BI — Burundi'), ('BT', 'BT — Bután'), ('CV', 'CV — Cabo Verde'), ('KH', 'KH — Camboya'), ('CM', 'CM — Camerún'), ('CA', 'CA — Canadá'), ('QA', 'QA — Catar'), ('TD', 'TD — Chad'), ('CL', 'CL — Chile'), ('CN', 'CN — China'), ('CY', 'CY — Chipre'), ('CO', 'CO — Colombia'), ('KM', 'KM — Comoras'), ('CG', 'CG — Congo'), ('CD', 'CD — Congo (RDC)'), ('KP', 'KP — Corea del Norte'), ('KR', 'KR — Corea del Sur'), ('CI', 'CI — Costa de Marfil'), ('CR', 'CR — Costa Rica'), ('HR', 'HR — Croacia'), ('CU', 'CU — Cuba'), ('DK', 'DK — Dinamarca'), ('DJ', 'DJ — Djibouti'), ('DM', 'DM — Dominica'), ('EC', 'EC — Ecuador'), ('EG', 'EG — Egipto'), ('SV', 'SV — El Salvador'), ('AE', 'AE — Emiratos Árabes Unidos'), ('ER', 'ER — Eritrea'), ('SK', 'SK — Eslovaquia'), ('SI', 'SI — Eslovenia'), ('ES', 'ES — España'), ('US', 'US — Estados Unidos'), ('EE', 'EE — Estonia'), ('ET', 'ET — Etiopía'), ('PH', 'PH — Filipinas'), ('FI', 'FI — Finlandia'), ('FJ', 'FJ — Fiyi'), ('FR', 'FR — Francia'), ('GA', 'GA — Gabón'), ('GM', 'GM — Gambia'), ('GE', 'GE — Georgia'), ('GH', 'GH — Ghana'), ('GD', 'GD — Granada'), ('GR', 'GR — Grecia'), ('GT', 'GT — Guatemala'), ('GN', 'GN — Guinea'), ('GW', 'GW — Guinea-Bisáu'), ('GQ', 'GQ — Guinea Ecuatorial'), ('GY', 'GY — Guyana'), ('HT', 'HT — Haití'), ('HN', 'HN — Honduras'), ('HU', 'HU — Hungría'), ('IN', 'IN — India'), ('ID', 'ID — Indonesia'), ('IQ', 'IQ — Irak'), ('IR', 'IR — Irán'), ('IE', 'IE — Irlanda'), ('IS', 'IS — Islandia'), ('MH', 'MH — Islas Marshall'), ('SB', 'SB — Islas Salomón'), ('IL', 'IL — Israel'), ('IT', 'IT — Italia'), ('JM', 'JM — Jamaica'), ('JP', 'JP — Japón'), ('JO', 'JO — Jordania'), ('KZ', 'KZ — Kazajistán'), ('KE', 'KE — Kenia'), ('KG', 'KG — Kirguistán'), ('KI', 'KI — Kiribati'), ('KW', 'KW — Kuwait'), ('LA', 'LA — Laos'), ('LS', 'LS — Lesoto'), ('LV', 'LV — Letonia'), ('LB', 'LB — Líbano'), ('LR', 'LR — Liberia'), ('LY', 'LY — Libia'), ('LI', 'LI — Liechtenstein'), ('LT', 'LT — Lituania'), ('LU', 'LU — Luxemburgo'), ('MK', 'MK — Macedonia del Norte'), ('MG', 'MG — Madagascar'), ('MY', 'MY — Malasia'), ('MW', 'MW — Malaui'), ('MV', 'MV — Maldivas'), ('ML', 'ML — Malí'), ('MT', 'MT — Malta'), ('MA', 'MA — Marruecos'), ('MU', 'MU — Mauricio'), ('MR', 'MR — Mauritania'), ('MX', 'MX — México'), ('FM', 'FM — Micronesia'), ('MD', 'MD — Moldavia'), ('MC', 'MC — Mónaco'), ('MN', 'MN — Mongolia'), ('ME', 'ME — Montenegro'), ('MZ', 'MZ — Mozambique'), ('MM', 'MM — Myanmar'), ('NA', 'NA — Namibia'), ('NR', 'NR — Nauru'), ('NP', 'NP — Nepal'), ('NI', 'NI — Nicaragua'), ('NE', 'NE — Níger'), ('NG', 'NG — Nigeria'), ('NO', 'NO — Noruega'), ('NZ', 'NZ — Nueva Zelanda'), ('OM', 'OM — Omán'), ('NL', 'NL — Países Bajos'), ('PK', 'PK — Pakistán'), ('PW', 'PW — Palaos'), ('PA', 'PA — Panamá'), ('PG', 'PG — Papúa Nueva Guinea'), ('PY', 'PY — Paraguay'), ('PE', 'PE — Perú'), ('PL', 'PL — Polonia'), ('PT', 'PT — Portugal'), ('GB', 'GB — Reino Unido'), ('CF', 'CF — República Centroafricana'), ('DO', 'DO — República Dominicana'), ('RW', 'RW — Ruanda'), ('RO', 'RO — Rumania'), ('RU', 'RU — Rusia'), ('WS', 'WS — Samoa'), ('LC', 'LC — Santa Lucía'), ('VC', 'VC — San Vicente y las Granadinas'), ('KN', 'KN — San Cristóbal y Nieves'), ('SM', 'SM — San Marino'), ('ST', 'ST — Santo Tomé y Príncipe'), ('SN', 'SN — Senegal'), ('RS', 'RS — Serbia'), ('SC', 'SC — Seychelles'), ('SL', 'SL — Sierra Leona'), ('SG', 'SG — Singapur'), ('SY', 'SY — Siria'), ('SO', 'SO — Somalia'), ('LK', 'LK — Sri Lanka'), ('SZ', 'SZ — Suazilandia'), ('ZA', 'ZA — Sudáfrica'), ('SD', 'SD — Sudán'), ('SS', 'SS — Sudán del Sur'), ('SE', 'SE — Suecia'), ('CH', 'CH — Suiza'), ('SR', 'SR — Surinam'), ('TH', 'TH — Tailandia'), ('TZ', 'TZ — Tanzania'), ('TJ', 'TJ — Tayikistán'), ('TL', 'TL — Timor Oriental'), ('TG', 'TG — Togo'), ('TO', 'TO — Tonga'), ('TT', 'TT — Trinidad y Tobago'), ('TN', 'TN — Túnez'), ('TM', 'TM — Turkmenistán'), ('TR', 'TR — Turquía'), ('TV', 'TV — Tuvalu'), ('UA', 'UA — Ucrania'), ('UG', 'UG — Uganda'), ('UY', 'UY — Uruguay'), ('UZ', 'UZ — Uzbekistán'), ('VU', 'VU — Vanuatu'), ('VE', 'VE — Venezuela'), ('VN', 'VN — Vietnam'), ('YE', 'YE — Yemen'), ('ZM', 'ZM — Zambia'), ('ZW', 'ZW — Zimbabue')])
        F(c,"País origen","codPaisOrigen",0,3,widget="combo",opts=[('AF', 'AF — Afganistán'), ('AL', 'AL — Albania'), ('DE', 'DE — Alemania'), ('AD', 'AD — Andorra'), ('AO', 'AO — Angola'), ('AG', 'AG — Antigua y Barbuda'), ('SA', 'SA — Arabia Saudita'), ('DZ', 'DZ — Argelia'), ('AR', 'AR — Argentina'), ('AM', 'AM — Armenia'), ('AU', 'AU — Australia'), ('AT', 'AT — Austria'), ('AZ', 'AZ — Azerbaiyán'), ('BS', 'BS — Bahamas'), ('BD', 'BD — Bangladesh'), ('BB', 'BB — Barbados'), ('BH', 'BH — Baréin'), ('BE', 'BE — Bélgica'), ('BZ', 'BZ — Belice'), ('BJ', 'BJ — Benín'), ('BY', 'BY — Bielorrusia'), ('BO', 'BO — Bolivia'), ('BA', 'BA — Bosnia y Herzegovina'), ('BW', 'BW — Botsuana'), ('BR', 'BR — Brasil'), ('BN', 'BN — Brunéi'), ('BG', 'BG — Bulgaria'), ('BF', 'BF — Burkina Faso'), ('BI', 'BI — Burundi'), ('BT', 'BT — Bután'), ('CV', 'CV — Cabo Verde'), ('KH', 'KH — Camboya'), ('CM', 'CM — Camerún'), ('CA', 'CA — Canadá'), ('QA', 'QA — Catar'), ('TD', 'TD — Chad'), ('CL', 'CL — Chile'), ('CN', 'CN — China'), ('CY', 'CY — Chipre'), ('CO', 'CO — Colombia'), ('KM', 'KM — Comoras'), ('CG', 'CG — Congo'), ('CD', 'CD — Congo (RDC)'), ('KP', 'KP — Corea del Norte'), ('KR', 'KR — Corea del Sur'), ('CI', 'CI — Costa de Marfil'), ('CR', 'CR — Costa Rica'), ('HR', 'HR — Croacia'), ('CU', 'CU — Cuba'), ('DK', 'DK — Dinamarca'), ('DJ', 'DJ — Djibouti'), ('DM', 'DM — Dominica'), ('EC', 'EC — Ecuador'), ('EG', 'EG — Egipto'), ('SV', 'SV — El Salvador'), ('AE', 'AE — Emiratos Árabes Unidos'), ('ER', 'ER — Eritrea'), ('SK', 'SK — Eslovaquia'), ('SI', 'SI — Eslovenia'), ('ES', 'ES — España'), ('US', 'US — Estados Unidos'), ('EE', 'EE — Estonia'), ('ET', 'ET — Etiopía'), ('PH', 'PH — Filipinas'), ('FI', 'FI — Finlandia'), ('FJ', 'FJ — Fiyi'), ('FR', 'FR — Francia'), ('GA', 'GA — Gabón'), ('GM', 'GM — Gambia'), ('GE', 'GE — Georgia'), ('GH', 'GH — Ghana'), ('GD', 'GD — Granada'), ('GR', 'GR — Grecia'), ('GT', 'GT — Guatemala'), ('GN', 'GN — Guinea'), ('GW', 'GW — Guinea-Bisáu'), ('GQ', 'GQ — Guinea Ecuatorial'), ('GY', 'GY — Guyana'), ('HT', 'HT — Haití'), ('HN', 'HN — Honduras'), ('HU', 'HU — Hungría'), ('IN', 'IN — India'), ('ID', 'ID — Indonesia'), ('IQ', 'IQ — Irak'), ('IR', 'IR — Irán'), ('IE', 'IE — Irlanda'), ('IS', 'IS — Islandia'), ('MH', 'MH — Islas Marshall'), ('SB', 'SB — Islas Salomón'), ('IL', 'IL — Israel'), ('IT', 'IT — Italia'), ('JM', 'JM — Jamaica'), ('JP', 'JP — Japón'), ('JO', 'JO — Jordania'), ('KZ', 'KZ — Kazajistán'), ('KE', 'KE — Kenia'), ('KG', 'KG — Kirguistán'), ('KI', 'KI — Kiribati'), ('KW', 'KW — Kuwait'), ('LA', 'LA — Laos'), ('LS', 'LS — Lesoto'), ('LV', 'LV — Letonia'), ('LB', 'LB — Líbano'), ('LR', 'LR — Liberia'), ('LY', 'LY — Libia'), ('LI', 'LI — Liechtenstein'), ('LT', 'LT — Lituania'), ('LU', 'LU — Luxemburgo'), ('MK', 'MK — Macedonia del Norte'), ('MG', 'MG — Madagascar'), ('MY', 'MY — Malasia'), ('MW', 'MW — Malaui'), ('MV', 'MV — Maldivas'), ('ML', 'ML — Malí'), ('MT', 'MT — Malta'), ('MA', 'MA — Marruecos'), ('MU', 'MU — Mauricio'), ('MR', 'MR — Mauritania'), ('MX', 'MX — México'), ('FM', 'FM — Micronesia'), ('MD', 'MD — Moldavia'), ('MC', 'MC — Mónaco'), ('MN', 'MN — Mongolia'), ('ME', 'ME — Montenegro'), ('MZ', 'MZ — Mozambique'), ('MM', 'MM — Myanmar'), ('NA', 'NA — Namibia'), ('NR', 'NR — Nauru'), ('NP', 'NP — Nepal'), ('NI', 'NI — Nicaragua'), ('NE', 'NE — Níger'), ('NG', 'NG — Nigeria'), ('NO', 'NO — Noruega'), ('NZ', 'NZ — Nueva Zelanda'), ('OM', 'OM — Omán'), ('NL', 'NL — Países Bajos'), ('PK', 'PK — Pakistán'), ('PW', 'PW — Palaos'), ('PA', 'PA — Panamá'), ('PG', 'PG — Papúa Nueva Guinea'), ('PY', 'PY — Paraguay'), ('PE', 'PE — Perú'), ('PL', 'PL — Polonia'), ('PT', 'PT — Portugal'), ('GB', 'GB — Reino Unido'), ('CF', 'CF — República Centroafricana'), ('DO', 'DO — República Dominicana'), ('RW', 'RW — Ruanda'), ('RO', 'RO — Rumania'), ('RU', 'RU — Rusia'), ('WS', 'WS — Samoa'), ('LC', 'LC — Santa Lucía'), ('VC', 'VC — San Vicente y las Granadinas'), ('KN', 'KN — San Cristóbal y Nieves'), ('SM', 'SM — San Marino'), ('ST', 'ST — Santo Tomé y Príncipe'), ('SN', 'SN — Senegal'), ('RS', 'RS — Serbia'), ('SC', 'SC — Seychelles'), ('SL', 'SL — Sierra Leona'), ('SG', 'SG — Singapur'), ('SY', 'SY — Siria'), ('SO', 'SO — Somalia'), ('LK', 'LK — Sri Lanka'), ('SZ', 'SZ — Suazilandia'), ('ZA', 'ZA — Sudáfrica'), ('SD', 'SD — Sudán'), ('SS', 'SS — Sudán del Sur'), ('SE', 'SE — Suecia'), ('CH', 'CH — Suiza'), ('SR', 'SR — Surinam'), ('TH', 'TH — Tailandia'), ('TZ', 'TZ — Tanzania'), ('TJ', 'TJ — Tayikistán'), ('TL', 'TL — Timor Oriental'), ('TG', 'TG — Togo'), ('TO', 'TO — Tonga'), ('TT', 'TT — Trinidad y Tobago'), ('TN', 'TN — Túnez'), ('TM', 'TM — Turkmenistán'), ('TR', 'TR — Turquía'), ('TV', 'TV — Tuvalu'), ('UA', 'UA — Ucrania'), ('UG', 'UG — Uganda'), ('UY', 'UY — Uruguay'), ('UZ', 'UZ — Uzbekistán'), ('VU', 'VU — Vanuatu'), ('VE', 'VE — Venezuela'), ('VN', 'VN — Vietnam'), ('YE', 'YE — Yemen'), ('ZM', 'ZM — Zambia'), ('ZW', 'ZW — Zimbabue')])
        F(c,"Forma pago","formaPago",1,0,widget="combo",opts=[
            ("99","99 — Sin pago exterior"),("01","01 — Giro directo"),("02","02 — Carta crédito")])
        F(c,"Subpartida arancelaria (10 dígitos)","subpartida",1,1)
        F(c,"No. bultos","numBultos",1,2); F(c,"Cantidad","cantidad",1,3)
        F(c,"Peso bruto (kg)","pesoBruto",2,0); F(c,"Peso neto (kg)","pesoNeto",2,1)
        F(c,"Valor FOB (USD)","fob",2,2); F(c,"Valor fletes (USD)","fletes",2,3)
        F(c,"Valor seguros (USD)","seguros",3,0); F(c,"Otros gastos (USD)","otrosGastos",3,1)
        F(c,"Ajuste valor (USD)","ajuste",3,2)
        F(c,"Valor Aduana CIF (auto)","valorAduana",3,3)
        F(c,"Descripción — marca, modelo, serial","descripcion",4,0,colspan=4,widget="text",width=80)

        c = self._card("05 — Liquidación Tributaria")
        F(c,"% Arancel","arancelPct",0,0); F(c,"% IVA","ivaPct",0,1); F(c,"% Imp. consumo","icPct",0,2)

        lf = tk.Frame(c, bg="#eff6ff", highlightbackground="#bfdbfe", highlightthickness=1)
        lf.grid(row=1, column=0, columnspan=4, sticky="ew", padx=4, pady=6)
        liq_rows = [("FOB","l_fob"),("+ Fletes","l_flt"),("+ Seguros","l_seg"),
                    ("+ Otros","l_otr"),("± Ajuste","l_adj"),
                    ("Valor Aduana CIF","l_cif"),("Arancel","l_ara"),("IVA","l_iva"),("Imp. Consumo","l_ic")]
        lf.columnconfigure(0, weight=1)
        for i,(lbl,attr) in enumerate(liq_rows):
            tk.Label(lf,text=lbl,font=("Arial",11),bg="#eff6ff",fg="#475569",anchor="w"
                     ).grid(row=i,column=0,sticky="w",padx=12,pady=2)
            v = tk.Label(lf,text="$0",font=("Arial",11),bg="#eff6ff",fg="#475569",anchor="e")
            v.grid(row=i,column=1,sticky="e",padx=12,pady=2); setattr(self,attr,v)
        tk.Frame(lf,bg="#bfdbfe",height=2).grid(row=len(liq_rows),column=0,columnspan=2,sticky="ew",padx=8,pady=4)
        tk.Label(lf,text="TOTAL LIQUIDADO (Cas. 93)",font=("Arial",12,"bold"),
                 bg="#eff6ff",fg="#1e3a5f",anchor="w").grid(row=len(liq_rows)+1,column=0,sticky="w",padx=12,pady=4)
        self.l_total_big = tk.Label(lf,text="$0 COP",font=("Arial",15,"bold"),bg="#eff6ff",fg="#1d4ed8",anchor="e")
        self.l_total_big.grid(row=len(liq_rows)+1,column=1,sticky="e",padx=12,pady=4)

        cas = tk.Frame(c, bg="white")
        cas.grid(row=2,column=0,columnspan=4,sticky="ew",padx=4,pady=(4,0))
        for i,(cl,ca) in enumerate([("Cas. 72 — Total Arancel $","cas72"),
                                     ("Cas. 76 — Total IVA $","cas76"),
                                     ("Cas. 980 — Pago total $","cas980")]):
            cf2 = tk.Frame(cas,bg="#dbeafe",highlightbackground="#93c5fd",highlightthickness=1)
            cf2.grid(row=0,column=i,sticky="ew",padx=5,pady=4); cas.columnconfigure(i,weight=1)
            tk.Label(cf2,text=cl,font=("Arial",9),bg="#dbeafe",fg="#1e40af").pack(anchor="w",padx=8,pady=(6,0))
            w = tk.Label(cf2,text="$0",font=("Arial",14,"bold"),bg="#dbeafe",fg="#1d4ed8",anchor="e")
            w.pack(fill="x",padx=8,pady=(0,6)); setattr(self,ca,w)

        tk.Frame(self.inner,bg="#f1f5f9",height=30).pack()

    # ── Logic ────────────────────────────────────────────────────────────────
    def _set_defaults(self):
        today = date.today().isoformat()
        defaults = {
            "nit":"","dv":"","razonSocial":"","direccion":"","telefono":"",
            "codDpto":"","codMunicipio":"",
            "nitDecl":"","dvDecl":"","razonDecl":"","tipoUsuario":"",
            "codUsuario":"","numDocDecl":"","nombresDecl":"",
            "numFormAnterior":"","manifestoCarga":"",
            "fechaLlegada":today,"docTransporte":"","fechaDocTransporte":today,
            "tasaCambio":"","nombreExportador":"","subpartida":"",
            "numBultos":"","cantidad":"","pesoBruto":"","pesoNeto":"",
            "fob":"","fletes":"","seguros":"","otrosGastos":"","ajuste":"",
            "descripcion":"","arancelPct":"","ivaPct":"19","icPct":"",
        }
        for k,v in defaults.items(): self._set_field(k,v)
        self._calc()

    def _get_field(self, key):
        w = self.fields.get(key)
        if w is None: return ""
        if isinstance(w, ttk.Combobox):
            sel = w.get()
            for cod,lbl in w._opts:
                if lbl == sel: return cod
            return sel.split("—")[0].strip() if "—" in sel else sel
        if isinstance(w, tk.Text): return w.get("1.0","end-1c")
        return w.get()

    def _set_field(self, key, value):
        w = self.fields.get(key)
        if w is None: return
        val = str(value) if value is not None else ""
        if isinstance(w, ttk.Combobox):
            for cod,lbl in w._opts:
                if str(cod).strip() == val.strip() or str(lbl).strip() == val.strip():
                    w.set(lbl); return
            for cod,lbl in w._opts:
                if val.strip().upper() in lbl.upper():
                    w.set(lbl); return
        elif isinstance(w, tk.Text):
            w.delete("1.0","end"); w.insert("1.0",val)
        elif isinstance(w, tk.Entry):
            state = w.cget("state")
            if state == "readonly": w.config(state="normal")
            w.delete(0,"end"); w.insert(0,val)
            if state == "readonly": w.config(state="readonly")

    def _cargar_datos(self, data):
        for k,v in data.items(): self._set_field(k, str(v))
        self._calc()

    def _calc(self):
        def fv(k):
            try: return float(self._get_field(k) or 0)
            except: return 0.0
        fob=fv("fob"); flt=fv("fletes"); seg=fv("seguros")
        otr=fv("otrosGastos"); adj=fv("ajuste")
        trm=fv("tasaCambio") or 4150
        ap=fv("arancelPct"); ip=fv("ivaPct"); icp=fv("icPct")
        cif=fob+flt+seg+otr+adj
        cifC=cif*trm; araC=cifC*(ap/100); ivaC=(cifC+araC)*(ip/100); icC=cifC*(icp/100)
        total=araC+ivaC+icC

        w=self.fields.get("valorAduana")
        if isinstance(w,tk.Entry):
            w.config(state="normal"); w.delete(0,"end"); w.insert(0,f"{cif:.2f}")
            w.config(state="readonly",readonlybackground="#eef4ff",fg="#1d4ed8")

        def f(n): return f"${int(round(n)):,}".replace(",",".")
        self.l_fob.config(text=f"${fob:.2f} USD"); self.l_flt.config(text=f"${flt:.2f} USD")
        self.l_seg.config(text=f"${seg:.2f} USD"); self.l_otr.config(text=f"${otr:.2f} USD")
        self.l_adj.config(text=f"${adj:.2f} USD")
        self.l_cif.config(text=f"${cif:.2f} USD = {f(cifC)} COP", font=("Arial",11,"bold"), fg="#0f172a")
        self.l_ara.config(text=f"{f(araC)} COP ({ap}%)")
        self.l_iva.config(text=f"{f(ivaC)} COP ({ip}%)")
        self.l_ic.config(text=f"{f(icC)} COP ({icp}%)")
        self.l_total_big.config(text=f"{f(total)} COP")
        self.cas72.config(text=f(araC)); self.cas76.config(text=f(ivaC)); self.cas980.config(text=f(total))
        self.lbl_total.config(text=f(total))
        self.lbl_ara.config(text=f(araC)); self.lbl_iva.config(text=f(ivaC)); self.lbl_ic.config(text=f(icC))
        self._total_cop = total

    def _jump(self, idx):
        self.canvas.update_idletasks()
        if idx >= len(self._sections): return
        sec = self._sections[idx]
        # Walk up to get y relative to inner frame
        y = 0
        w = sec
        while w is not None and w != self.inner:
            y += w.winfo_y()
            w = w.master
        total_h = max(self.inner.winfo_height(), 1)
        canvas_h = max(self.canvas.winfo_height(), 1)
        frac = y / max(total_h - canvas_h, 1)
        self.canvas.yview_moveto(max(0.0, min(1.0, frac)))
        for i, b in enumerate(self._nav_btns):
            b.config(
                bg="#0d1e30" if i == idx else "#0f1724",
                fg="#e0eaf4" if i == idx else "#4a6a8a",
                font=("Arial", 10, "bold") if i == idx else ("Arial", 10))

    def _auto_dv(self, *args):
        nit = self._get_field("nit").strip()
        if len(nit) >= 6:
            dv = calcular_dv(nit)
            if dv: self._set_field("dv", dv)

    # ── Clientes ─────────────────────────────────────────────────────────────
    def _abrir_clientes(self):
        VentanaClientes(self, on_select=self._aplicar_cliente)

    def _aplicar_cliente(self, row):
        # row: (id, nit, dv, razon_social, direccion, telefono, cod_seccional, cod_dpto, cod_municipio, creado)
        self._cliente_id = row[0]
        self._set_field("nit", row[1] or "")
        self._set_field("dv",  row[2] or "")
        self._set_field("razonSocial", row[3] or "")
        self._set_field("direccion",   row[4] or "")
        self._set_field("telefono",    row[5] or "")
        self._set_field("codSeccional",row[6] or "")
        self._set_field("codDpto",     row[7] or "")
        self._set_field("codMunicipio",row[8] or "")
        self.lbl_cliente_badge.config(text=f"👤 {row[3]}")
        self._calc()

    # ── Historial ─────────────────────────────────────────────────────────────
    def _abrir_historial(self):
        VentanaHistorial(self, self)

    def _guardar_decl(self):
        if not self._cliente_id:
            # Try to find client by NIT already in form
            nit = self._get_field("nit").strip()
            if nit:
                rows = db_fetch("SELECT id FROM clientes WHERE nit=?", (nit,))
                if rows:
                    self._cliente_id = rows[0][0]
            if not self._cliente_id:
                # Auto-create client from form data
                nit = self._get_field("nit").strip()
                razon = self._get_field("razonSocial").strip()
                if nit and razon:
                    self._cliente_id = db_exec(
                        "INSERT INTO clientes(nit,dv,razon_social,direccion,telefono,cod_seccional,cod_dpto,cod_municipio) VALUES(?,?,?,?,?,?,?,?)",
                        (nit, self._get_field("dv"), razon, self._get_field("direccion"),
                         self._get_field("telefono"), self._get_field("codSeccional"),
                         self._get_field("codDpto"), self._get_field("codMunicipio")))
                    self.lbl_cliente_badge.config(text=f"👤 {razon}")
                else:
                    if not messagebox.askyesno("Sin cliente",
                        "No hay cliente seleccionado ni NIT ingresado.\n¿Guardar sin cliente?"):
                        return
        self._calc()
        data = {k: self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        total = getattr(self, "_total_cop", 0)
        doc_transp = self._get_field("docTransporte")
        fecha = date.today().isoformat()
        if self._decl_id:
            db_exec("UPDATE declaraciones SET datos=?,total_cop=?,numero=?,fecha=?,estado='Borrador' WHERE id=?",
                    (json.dumps(data, ensure_ascii=False), total, doc_transp, fecha, self._decl_id))
        else:
            self._decl_id = db_exec(
                "INSERT INTO declaraciones(cliente_id,numero,fecha,datos,total_cop,estado) VALUES(?,?,?,?,?,?)",
                (self._cliente_id, doc_transp, fecha, json.dumps(data, ensure_ascii=False), total, "Borrador"))
        messagebox.showinfo("Guardado","✅ Declaración guardada en el historial.")

    # ── TRM ───────────────────────────────────────────────────────────────────
    def _update_trm(self):
        self.lbl_trm_status.config(text="Consultando TRM...", fg="#f59e0b")
        def fetch():
            trm = None
            today = date.today().strftime("%Y-%m-%d")
            # Intento 1: datos.gov.co
            try:
                url = (f"https://www.datos.gov.co/resource/mcec-87by.json"
                       f"?$where=vigenciadesde>='{today}T00:00:00.000'&$limit=1&$order=vigenciadesde DESC")
                req = urllib.request.Request(url, headers={"User-Agent":"Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=6) as r:
                    data = json.loads(r.read())
                if data and "valor" in data[0]:
                    trm = float(str(data[0]["valor"]).replace(",","."))
            except: pass
            # Intento 2: si no hay TRM para hoy (fin de semana/festivo), buscar la más reciente
            if not trm:
                try:
                    url2 = "https://www.datos.gov.co/resource/mcec-87by.json?$limit=1&$order=vigenciadesde DESC"
                    req2 = urllib.request.Request(url2, headers={"User-Agent":"Mozilla/5.0"})
                    with urllib.request.urlopen(req2, timeout=6) as r2:
                        data2 = json.loads(r2.read())
                    if data2 and "valor" in data2[0]:
                        trm = float(str(data2[0]["valor"]).replace(",","."))
                        today = str(data2[0].get("vigenciadesde",""))[:10]
                except: pass
            if trm and trm > 0:
                self.after(0, lambda t=trm, d=today: self._apply_trm(t, d))
            else:
                self.after(0, lambda: self.lbl_trm_status.config(
                    text="Sin conexión.\nIngrese TRM manualmente.", fg="#ef4444"))
        threading.Thread(target=fetch, daemon=True).start()

    def _apply_trm(self, trm, fecha):
        self._set_field("tasaCambio", f"{trm:.2f}")
        self._calc()
        self.lbl_trm_status.config(text=f"TRM {fecha}: ${trm:,.2f}", fg="#22c55e")

    # ── Excel ─────────────────────────────────────────────────────────────────
    def _load_excel(self):
        if not HAS_XLSX:
            messagebox.showerror("Error","openpyxl no está instalado."); return
        path = filedialog.askopenfilename(
            title="Seleccionar plantilla Excel",
            filetypes=[("Excel","*.xlsx *.xls")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["Formulario510"]
            loaded = 0
            for key, cell_addr in EXCEL_MAP.items():
                val = ws[cell_addr].value
                if val is not None:
                    self._set_field(key, str(val).strip()); loaded += 1
            self._calc()
            messagebox.showinfo("Listo", f"✅ {loaded} campos importados desde Excel.")
        except KeyError:
            messagebox.showerror("Error","No se encontró la hoja 'Formulario510'.")
        except Exception as e:
            messagebox.showerror("Error al cargar Excel", str(e))

    # ── Subpartidas ───────────────────────────────────────────────────────────
    def _subpartidas(self):
        win = tk.Toplevel(self)
        win.title("Subpartidas Arancelarias Comunes")
        win.geometry("540x440")
        win.configure(bg="white"); win.resizable(False,False)
        tk.Label(win, text="Subpartidas Arancelarias Comunes", font=("Arial",12,"bold"),
                 bg="white", fg="#1d4ed8").pack(pady=(14,2))
        tk.Label(win, text="Doble clic para seleccionar",
                 font=("Arial",9), bg="white", fg="#94a3b8").pack(pady=(0,8))
        frame = tk.Frame(win, bg="white"); frame.pack(fill="both", expand=True, padx=14, pady=(0,14))
        cols = ("Producto","Subpartida","Arancel","IVA")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width={"Producto":210,"Subpartida":120,"Arancel":80,"IVA":70}[c], anchor="w")
        for row in SUBPARTIDAS:
            tree.insert("", "end", values=row)
        sb2 = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side="right", fill="y"); tree.pack(side="left", fill="both", expand=True)
        def sel(e=None):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])["values"]
            self._set_field("subpartida", v[1])
            self._set_field("arancelPct", str(v[2]).replace("%",""))
            self._calc(); win.destroy()
        tree.bind("<Double-1>", sel)
        tk.Button(win, text="Seleccionar", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=sel).pack(fill="x", padx=14, pady=(0,8))

    # ── Validar ───────────────────────────────────────────────────────────────
    def _validar(self):
        errores = []
        for key, label in [("nit","NIT del importador"),("razonSocial","Razón social"),
                            ("fechaLlegada","Fecha de llegada"),("docTransporte","Doc. transporte"),
                            ("subpartida","Subpartida arancelaria"),("fob","Valor FOB"),
                            ("descripcion","Descripción de la mercancía")]:
            if not self._get_field(key).strip():
                errores.append(f"  • {label}")
        return errores

    # ── PDF ───────────────────────────────────────────────────────────────────
    def _generate(self):
        self._calc()
        errores = self._validar()
        if errores:
            msg = "Faltan campos obligatorios:\n\n" + "\n".join(errores) + "\n\n¿Generar de todas formas?"
            if not messagebox.askyesno("Campos incompletos", msg): return
        try:
            fob=float(self._get_field("fob") or 0); flt=float(self._get_field("fletes") or 0)
            seg=float(self._get_field("seguros") or 0); otr=float(self._get_field("otrosGastos") or 0)
            adj=float(self._get_field("ajuste") or 0); cif=fob+flt+seg+otr+adj
            if cif > 2000:
                messagebox.showwarning("⚠️ Límite",
                    f"CIF ${cif:,.2f} USD supera $2.000 USD.\nUse Formulario 500 (declaración formal).")
            elif cif > 200:
                messagebox.showinfo("ℹ️ Aviso", f"CIF ${cif:,.2f} USD supera $200 USD.\nAplican impuestos.")
        except: pass
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF","*.pdf")],
            initialfile=f"Formulario510_{date.today().isoformat()}.pdf",
            title="Guardar PDF")
        if not path: return
        data = {k: self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        try:
            make_pdf(data, path)
            # Actualizar PDF en historial si existe
            if self._decl_id:
                db_exec("UPDATE declaraciones SET pdf_path=?,estado='Generado' WHERE id=?",
                        (path, self._decl_id))
            messagebox.showinfo("PDF Generado", f"✅ PDF generado:\n{path}")
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e))

    # ── Clear ─────────────────────────────────────────────────────────────────
    def _clear(self):
        if messagebox.askyesno("Confirmar","¿Limpiar todos los campos?"):
            self._cliente_id = None; self._decl_id = None
            self.lbl_cliente_badge.config(text="Sin cliente seleccionado")
            self._set_defaults()


if __name__ == "__main__":
    app = App()
    app.mainloop()
