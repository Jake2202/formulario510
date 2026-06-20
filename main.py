import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date, datetime, timedelta
import os, sys, sqlite3, json, threading, urllib.request, urllib.parse, webbrowser, hashlib, random, string, shutil

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
        pdf_path TEXT, fecha_levante TEXT, fecha_vencimiento TEXT,
        creado TEXT DEFAULT CURRENT_TIMESTAMP,
        FOREIGN KEY(cliente_id) REFERENCES clientes(id)
    )''')
    # Add columns if upgrading from old DB
    try:
        c.execute("ALTER TABLE declaraciones ADD COLUMN fecha_levante TEXT")
        c.execute("ALTER TABLE declaraciones ADD COLUMN fecha_vencimiento TEXT")
        conn.commit()
    except: pass
    c.execute('''CREATE TABLE IF NOT EXISTS usuarios (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE, password_hash TEXT,
        rol TEXT DEFAULT "operador", activo INTEGER DEFAULT 1,
        creado TEXT DEFAULT CURRENT_TIMESTAMP
    )''')
    c.execute('''CREATE TABLE IF NOT EXISTS config (
        clave TEXT PRIMARY KEY, valor TEXT
    )''')
    # Default admin user if none exists
    c.execute("SELECT COUNT(*) FROM usuarios")
    if c.fetchone()[0] == 0:
        pwd = hashlib.sha256("admin123".encode()).hexdigest()
        c.execute("INSERT INTO usuarios(username,password_hash,rol) VALUES(?,?,?)",
                  ("admin", pwd, "admin"))
    # Default config
    for k,v in [("agencia_nombre","Mi Agencia de Aduanas"),
                ("agencia_nit",""),("agencia_tel",""),
                ("agencia_dir",""),("licencia_key",""),
                ("licencia_activa","0"),("max_decl_trial","5")]:
        c.execute("INSERT OR IGNORE INTO config(clave,valor) VALUES(?,?)",(k,v))
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
    """Calcula el dígito de verificación del NIT colombiano (algoritmo oficial DIAN,
    módulo 11). Los pesos se aplican de derecha a izquierda: 3,7,13,17,19,23,29,37,41,43,47,53,59,67,71."""
    try:
        nit_str = str(nit).strip().replace(".","").replace("-","")
        if not nit_str.isdigit() or not nit_str: return ""
        factores = [3,7,13,17,19,23,29,37,41,43,47,53,59,67,71]
        digitos_invertidos = nit_str[::-1]
        total = sum(int(d) * factores[i] for i, d in enumerate(digitos_invertidos) if i < len(factores))
        residuo = total % 11
        if residuo == 0: return "0"
        if residuo == 1: return "1"
        return str(11 - residuo)
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
# VENTANA: Checklist de Documentos Soporte
# ═══════════════════════════════════════════════════════════════════════════════
DOCUMENTOS_BASE = [
    ("Factura comercial del proveedor",              True),
    ("Lista de empaque (Packing list)",              True),
    ("Documento de transporte (AWB / BL / Guía)",   True),
    ("Declaración de importación Formulario 510",    True),
    ("Recibo Oficial de Pago (ROP)",                 True),
    ("Poder o autorización del importador",          False),
    ("Registro de importación (si aplica)",          False),
    ("Visto bueno INVIMA (medicamentos/alimentos)",  False),
    ("Visto bueno ICA (productos agropecuarios)",    False),
    ("Visto bueno Min. Transporte (vehículos)",      False),
    ("Certificado de origen (si aplica TLC)",        False),
    ("Garantía / Fianza (mercancía restringida)",    False),
    ("Foto de la mercancía (si lo exige el aforo)",  False),
]

class VentanaChecklist(tk.Toplevel):
    def __init__(self, parent, decl_info=""):
        super().__init__(parent)
        self.title("Checklist — Documentos Soporte")
        self.geometry("560x580")
        self.configure(bg="white")
        self.vars = []
        self._build(decl_info)

    def _build(self, decl_info):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="✅  Checklist de Documentos Soporte",
                 font=("Arial",13,"bold"), bg="white", fg="#1d4ed8").pack(pady=(14,2))
        if decl_info:
            tk.Label(self, text=decl_info, font=("Arial",9),
                     bg="white", fg="#64748b").pack(pady=(0,8))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="both", expand=True, padx=20)

        tk.Label(frame, text="Marque los documentos que ya tiene listos:",
                 font=("Arial",10), bg="white", fg="#475569").pack(anchor="w", pady=(0,10))

        for doc, obligatorio in DOCUMENTOS_BASE:
            row = tk.Frame(frame, bg="white"); row.pack(fill="x", pady=3)
            var = tk.BooleanVar(value=obligatorio)
            self.vars.append((var, doc, obligatorio))
            color = "#dc2626" if obligatorio else "#64748b"
            tag = " ★" if obligatorio else ""
            cb = tk.Checkbutton(row, text=doc+tag, variable=var,
                                font=("Arial",10), bg="white", fg=color,
                                activebackground="white", cursor="hand2",
                                selectcolor="#dbeafe")
            cb.pack(side="left")

        tk.Label(frame, text="★ = Documento obligatorio siempre",
                 font=("Arial",8), bg="white", fg="#94a3b8").pack(anchor="w", pady=(8,0))

        # Progress bar
        self.lbl_progress = tk.Label(self, text="", font=("Arial",10,"bold"),
                                      bg="white", fg="#1d4ed8")
        self.lbl_progress.pack(pady=4)
        self._update_progress()
        for var,_,_ in self.vars:
            var.trace("w", lambda *a: self._update_progress())

        btns = tk.Frame(self, bg="white"); btns.pack(fill="x", padx=20, pady=(4,16))
        tk.Button(btns, text="🖨️  Imprimir checklist", font=("Arial",10,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=8, cursor="hand2",
                  command=self._imprimir).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="Cerrar", font=("Arial",10), bg="#f1f5f9", fg="#64748b",
                  relief="flat", pady=8, cursor="hand2",
                  command=self.destroy).pack(side="left", fill="x", expand=True)

    def _update_progress(self):
        total = len(self.vars); marcados = sum(1 for v,_,_ in self.vars if v.get())
        oblig_total = sum(1 for _,_,o in self.vars if o)
        oblig_ok = sum(1 for v,_,o in self.vars if o and v.get())
        color = "#22c55e" if oblig_ok == oblig_total else "#f59e0b"
        self.lbl_progress.config(
            text=f"{marcados}/{total} documentos listos  |  Obligatorios: {oblig_ok}/{oblig_total}",
            fg=color)

    def _imprimir(self):
        lines = ["CHECKLIST DOCUMENTOS SOPORTE — FORMULARIO 510", "="*50, ""]
        for var, doc, oblig in self.vars:
            estado = "[X]" if var.get() else "[ ]"
            tag = " (OBLIGATORIO)" if oblig else ""
            lines.append(f"{estado} {doc}{tag}")
        lines += ["","Fecha: "+date.today().strftime("%d/%m/%Y")]
        txt = "\n".join(lines)
        import tempfile, os
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt",
                                          delete=False, encoding="utf-8")
        tmp.write(txt); tmp.close()
        os.startfile(tmp.name)


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Control de Plazos
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaPlazos(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Control de Plazos y Alertas")
        self.geometry("860x520")
        self.configure(bg="white")
        self._build()
        self._cargar()

    def _build(self):
        hdr = tk.Frame(self, bg="#1d4ed8", height=48); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="⏰  Control de Plazos — Declaraciones",
                 font=("Arial",13,"bold"), bg="#1d4ed8", fg="white").pack(side="left", padx=16, pady=10)

        # Info
        info = tk.Frame(self, bg="#fef9c3"); info.pack(fill="x")
        tk.Label(info, text="⚠️  Plazo legal: máximo 2 meses desde el levante para presentar la declaración formal.",
                 font=("Arial",9), bg="#fef9c3", fg="#92400e").pack(padx=16, pady=6)

        # Toolbar
        tb = tk.Frame(self, bg="#f8fafc", height=46); tb.pack(fill="x"); tb.pack_propagate(False)
        tk.Button(tb, text="📅 Registrar levante", font=("Arial",10,"bold"),
                  bg="#0f766e", fg="white", relief="flat", padx=12, pady=6,
                  cursor="hand2", command=self._registrar_levante).pack(side="left", padx=8, pady=6)
        tk.Button(tb, text="🔄 Actualizar", font=("Arial",10,"bold"),
                  bg="#1e3a5f", fg="white", relief="flat", padx=12, pady=6,
                  cursor="hand2", command=self._cargar).pack(side="left", padx=2, pady=6)

        cols = ("ID","Cliente","Doc. Transporte","Fecha Levante","Vencimiento","Días restantes","Estado")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=16)
        widths = {"ID":40,"Cliente":220,"Doc. Transporte":140,"Fecha Levante":110,
                  "Vencimiento":110,"Días restantes":100,"Estado":90}
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], anchor="w")

        # Color tags
        self.tree.tag_configure("vencido",   background="#fee2e2", foreground="#dc2626")
        self.tree.tag_configure("urgente",   background="#fef9c3", foreground="#92400e")
        self.tree.tag_configure("ok",        background="#f0fdf4", foreground="#15803d")
        self.tree.tag_configure("sin_fecha", background="#f8fafc", foreground="#94a3b8")

        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True, padx=8, pady=8)

    def _cargar(self):
        rows = db_fetch("""SELECT d.id, c.razon_social, d.numero, d.fecha_levante, d.fecha_vencimiento
                           FROM declaraciones d
                           LEFT JOIN clientes c ON d.cliente_id=c.id
                           ORDER BY d.fecha_vencimiento ASC""")
        self.tree.delete(*self.tree.get_children())
        today = date.today()
        alertas = []
        for r in rows:
            did, cliente, doc, f_lev, f_venc = r
            if not f_lev:
                self.tree.insert("","end",values=(did,cliente or "—",doc or "—","Sin registrar","—","—","Pendiente"),tags=("sin_fecha",))
                continue
            try:
                lev_date  = datetime.strptime(f_lev, "%Y-%m-%d").date()
                venc_date = lev_date + timedelta(days=60)
                dias = (venc_date - today).days
                if dias < 0:
                    tag="vencido"; estado="VENCIDO"
                elif dias <= 10:
                    tag="urgente"; estado=f"URGENTE"
                    alertas.append((cliente, doc, dias))
                else:
                    tag="ok"; estado="Al día"
                self.tree.insert("","end",values=(did,cliente or "—",doc or "—",
                    f_lev, venc_date.strftime("%Y-%m-%d"), f"{dias} días", estado),tags=(tag,))
            except:
                self.tree.insert("","end",values=(did,cliente or "—",doc or "—",f_lev,"Error","—","Error"),tags=("sin_fecha",))
        if alertas:
            msg = "⚠️ DECLARACIONES PRÓXIMAS A VENCER:\n\n"
            for c,d,dias in alertas:
                msg += f"• {c} — Guía {d}: {dias} días restantes\n"
            messagebox.showwarning("Alertas de vencimiento", msg)

    def _registrar_levante(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso","Seleccione una declaración."); return
        did = self.tree.item(sel[0])["values"][0]
        win = tk.Toplevel(self); win.title("Registrar Levante"); win.geometry("320x180")
        win.configure(bg="white"); win.resizable(False,False)
        tk.Label(win, text="Fecha de levante (AAAA-MM-DD):", font=("Arial",10),
                 bg="white").pack(pady=(20,4), padx=20, anchor="w")
        entry = tk.Entry(win, font=("Arial",12), relief="flat",
                         highlightbackground="#cbd5e1", highlightthickness=1)
        entry.insert(0, date.today().isoformat())
        entry.pack(fill="x", padx=20, ipady=5)
        def guardar():
            fecha = entry.get().strip()
            try:
                datetime.strptime(fecha, "%Y-%m-%d")
                venc = (datetime.strptime(fecha,"%Y-%m-%d") + timedelta(days=60)).strftime("%Y-%m-%d")
                db_exec("UPDATE declaraciones SET fecha_levante=?,fecha_vencimiento=?,estado='Con levante' WHERE id=?",
                        (fecha, venc, did))
                win.destroy(); self._cargar()
            except:
                messagebox.showerror("Error","Fecha inválida. Use formato AAAA-MM-DD.")
        tk.Button(win, text="Guardar", font=("Arial",11,"bold"), bg="#1d4ed8", fg="white",
                  relief="flat", pady=8, cursor="hand2", command=guardar).pack(fill="x", padx=20, pady=12)


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Recibo Oficial de Pago (ROP)
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaROP(tk.Toplevel):
    def __init__(self, parent, data, total_cop, trm):
        super().__init__(parent)
        self.title("Recibo Oficial de Pago — ROP")
        self.geometry("540x620")
        self.configure(bg="white")
        self.data = data; self.total_cop = total_cop; self.trm = trm
        self._build()

    def _build(self):
        # Header
        hdr = tk.Frame(self, bg="#1d4ed8", height=54); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="Recibo Oficial de Pago", font=("Arial",14,"bold"),
                 bg="#1d4ed8", fg="white").pack(side="left", padx=16, pady=8)
        tk.Label(hdr, text="Formulario 510 — DIAN Colombia",
                 font=("Arial",9), bg="#1d4ed8", fg="#bfdbfe").pack(side="left")

        frame = tk.Frame(self, bg="white"); frame.pack(fill="both", expand=True, padx=24, pady=16)

        def row(label, value, bold=False, big=False):
            r = tk.Frame(frame, bg="white"); r.pack(fill="x", pady=3)
            tk.Label(r, text=label, font=("Arial",9), bg="white",
                     fg="#64748b", width=28, anchor="w").pack(side="left")
            tk.Label(r, text=value, font=("Arial",12 if big else 10, "bold" if bold else "normal"),
                     bg="white", fg="#1d4ed8" if big else "#0f172a", anchor="w").pack(side="left")

        # Cálculos
        fob  = float(self.data.get("fob","0") or 0)
        flt  = float(self.data.get("fletes","0") or 0)
        seg  = float(self.data.get("seguros","0") or 0)
        otr  = float(self.data.get("otrosGastos","0") or 0)
        adj  = float(self.data.get("ajuste","0") or 0)
        ap   = float(self.data.get("arancelPct","0") or 0)
        ip   = float(self.data.get("ivaPct","19") or 19)
        icp  = float(self.data.get("icPct","0") or 0)
        cif  = fob+flt+seg+otr+adj
        cifC = cif*self.trm
        araC = cifC*(ap/100)
        ivaC = (cifC+araC)*(ip/100)
        icC  = cifC*(icp/100)
        total= araC+ivaC+icC

        def fc(n): return f"${int(round(n)):,} COP".replace(",",".")
        def fu(n): return f"${n:,.2f} USD"

        tk.Frame(frame, bg="#e2e8f0", height=1).pack(fill="x", pady=6)
        tk.Label(frame, text="DATOS DEL IMPORTADOR", font=("Arial",9,"bold"),
                 bg="white", fg="#1d4ed8").pack(anchor="w", pady=(0,4))
        row("NIT / Razón social:", f"{self.data.get('nit','')} - {self.data.get('razonSocial','')}")
        row("Doc. transporte:", self.data.get("docTransporte",""))
        row("Fecha declaración:", date.today().strftime("%d/%m/%Y"))
        row("Seccional DIAN:", self.data.get("codSeccional",""))

        tk.Frame(frame, bg="#e2e8f0", height=1).pack(fill="x", pady=8)
        tk.Label(frame, text="LIQUIDACIÓN", font=("Arial",9,"bold"),
                 bg="white", fg="#1d4ed8").pack(anchor="w", pady=(0,4))
        row("Valor CIF:", f"{fu(cif)} = {fc(cifC)}")
        row("Tasa de cambio:", f"${self.trm:,.2f} COP/USD")
        row(f"Casilla 72 — Arancel ({ap}%):", fc(araC))
        row(f"Casilla 76 — IVA ({ip}%):", fc(ivaC))
        if icp > 0: row(f"Imp. al consumo ({icp}%):", fc(icC))

        tk.Frame(frame, bg="#1d4ed8", height=2).pack(fill="x", pady=8)
        row("CASILLA 980 — TOTAL A PAGAR:", fc(total), bold=True, big=True)
        tk.Frame(frame, bg="#1d4ed8", height=2).pack(fill="x", pady=(0,8))

        # Código de referencia
        import hashlib
        ref = hashlib.md5(f"{self.data.get('nit','')}{self.data.get('docTransporte','')}{date.today()}".encode()).hexdigest()[:12].upper()
        row("Referencia de pago:", ref)
        row("Banco autorizado:", "Cualquier banco habilitado DIAN")
        row("Vigencia:", date.today().strftime("%d/%m/%Y")+" (mismo día hábil)")

        tk.Label(frame, text="⚠️  Este ROP es de referencia. El oficial se genera en el SYGA.",
                 font=("Arial",8), bg="white", fg="#94a3b8", wraplength=460).pack(pady=(8,0))

        btns = tk.Frame(self, bg="white"); btns.pack(fill="x", padx=24, pady=(0,16))
        tk.Button(btns, text="🖨️  Imprimir ROP", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10, cursor="hand2",
                  command=self._imprimir).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="🌐  Ir al SYGA", font=("Arial",11,"bold"),
                  bg="#0f766e", fg="white", relief="flat", pady=10, cursor="hand2",
                  command=lambda: webbrowser.open("https://importaciones.dian.gov.co")
                  ).pack(side="left", fill="x", expand=True)

    def _imprimir(self):
        fob=float(self.data.get("fob","0") or 0); flt=float(self.data.get("fletes","0") or 0)
        seg=float(self.data.get("seguros","0") or 0); otr=float(self.data.get("otrosGastos","0") or 0)
        adj=float(self.data.get("ajuste","0") or 0); ap=float(self.data.get("arancelPct","0") or 0)
        ip=float(self.data.get("ivaPct","19") or 19); icp=float(self.data.get("icPct","0") or 0)
        cif=fob+flt+seg+otr+adj; cifC=cif*self.trm
        araC=cifC*(ap/100); ivaC=(cifC+araC)*(ip/100); icC=cifC*(icp/100); total=araC+ivaC+icC
        def fc(n): return f"${int(round(n)):,} COP".replace(",",".")
        import hashlib, tempfile, os
        ref = hashlib.md5(f"{self.data.get('nit','')}{self.data.get('docTransporte','')}{date.today()}".encode()).hexdigest()[:12].upper()
        txt = f"""
RECIBO OFICIAL DE PAGO — REFERENCIA
DIAN Colombia — Formulario 510
{'='*50}
NIT:             {self.data.get('nit','')}
Importador:      {self.data.get('razonSocial','')}
Doc. Transporte: {self.data.get('docTransporte','')}
Fecha:           {date.today().strftime('%d/%m/%Y')}
Seccional:       {self.data.get('codSeccional','')}
{'='*50}
Valor CIF:       ${cif:.2f} USD = {fc(cifC)}
TRM:             ${self.trm:,.2f} COP/USD
Arancel ({ap}%): {fc(araC)}
IVA ({ip}%):     {fc(ivaC)}
Imp. Consumo:    {fc(icC)}
{'='*50}
TOTAL A PAGAR:   {fc(total)}
Referencia:      {ref}
{'='*50}
Pague en cualquier banco habilitado DIAN
Vigencia: {date.today().strftime('%d/%m/%Y')}

NOTA: Recibo de referencia. El oficial se genera en el SYGA DIAN.
"""
        tmp = tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False, encoding="utf-8")
        tmp.write(txt); tmp.close()
        os.startfile(tmp.name)


# ═══════════════════════════════════════════════════════════════════════════════
# FASE 3 — EDI / XML, Consulta Levante, Entrega Urgente
# ═══════════════════════════════════════════════════════════════════════════════

# ── Generador de archivo EDI (XML SYGA) ──────────────────────────────────────
class VentanaEDI(tk.Toplevel):
    def __init__(self, parent, data):
        super().__init__(parent)
        self.title("Generar Archivo EDI — SYGA")
        self.geometry("700x600")
        self.configure(bg="white")
        self.data = data
        self._build()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="📤  Archivo EDI — Transmisión SYGA",
                 font=("Arial",13,"bold"), bg="white", fg="#1d4ed8").pack(pady=(14,2))
        tk.Label(self, text="Genera el archivo XML/EDI para subir directamente al SYGA DIAN",
                 font=("Arial",9), bg="white", fg="#64748b").pack(pady=(0,10))

        # Preview area
        frame = tk.Frame(self, bg="white"); frame.pack(fill="both", expand=True, padx=16)
        self.txt = tk.Text(frame, font=("Courier New",9), bg="#0f172a", fg="#22d3ee",
                           relief="flat", wrap="none",
                           highlightbackground="#1e3a5f", highlightthickness=1)
        sbx = ttk.Scrollbar(frame, orient="horizontal", command=self.txt.xview)
        sby = ttk.Scrollbar(frame, orient="vertical",   command=self.txt.yview)
        self.txt.configure(xscrollcommand=sbx.set, yscrollcommand=sby.set)
        sby.pack(side="right", fill="y"); sbx.pack(side="bottom", fill="x")
        self.txt.pack(fill="both", expand=True)
        self._generar_xml()

        btns = tk.Frame(self, bg="white"); btns.pack(fill="x", padx=16, pady=12)
        tk.Button(btns, text="💾  Guardar .xml", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10, cursor="hand2",
                  command=self._guardar).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="🌐  Ir al SYGA", font=("Arial",11,"bold"),
                  bg="#0f766e", fg="white", relief="flat", pady=10, cursor="hand2",
                  command=lambda: webbrowser.open("https://importaciones.dian.gov.co")
                  ).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="Cerrar", font=("Arial",11),
                  bg="#f1f5f9", fg="#64748b", relief="flat", pady=10,
                  cursor="hand2", command=self.destroy).pack(side="left", fill="x", expand=True)

    def _generar_xml(self):
        d = self.data
        import xml.sax.saxutils as saxutils
        def g(k): return saxutils.escape(str(d.get(k,"") or "").strip())
        fob=float(d.get("fob","0") or 0); flt=float(d.get("fletes","0") or 0)
        seg=float(d.get("seguros","0") or 0); otr=float(d.get("otrosGastos","0") or 0)
        adj=float(d.get("ajuste","0") or 0); trm=float(d.get("tasaCambio","4150") or 4150)
        ap=float(d.get("arancelPct","0") or 0); ip=float(d.get("ivaPct","19") or 19); icp=float(d.get("icPct","0") or 0)
        cif=fob+flt+seg+otr+adj; cifC=cif*trm
        araC=cifC*(ap/100); ivaC=(cifC+araC)*(ip/100); icC=cifC*(icp/100); total=araC+ivaC+icC

        xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<!-- Archivo EDI generado por DeclaraFácil 510 — {date.today().isoformat()} -->
<!-- Para uso con SYGA DIAN — Declaración de Importación Simplificada -->
<DeclaracionImportacion xmlns="urn:dian:gov:co:syga:1.0"
    tipo="510" version="1.0" fecha="{date.today().isoformat()}">

  <!-- SECCIÓN 01: IMPORTADOR -->
  <Importador>
    <NIT>{g("nit")}</NIT>
    <DigitoVerificacion>{g("dv")}</DigitoVerificacion>
    <RazonSocial>{g("razonSocial")}</RazonSocial>
    <Direccion>{g("direccion")}</Direccion>
    <Telefono>{g("telefono")}</Telefono>
    <CodigoSeccional>{g("codSeccional")}</CodigoSeccional>
    <CodigoDepartamento>{g("codDpto")}</CodigoDepartamento>
    <CodigoMunicipio>{g("codMunicipio")}</CodigoMunicipio>
  </Importador>

  <!-- SECCIÓN 02: DECLARANTE -->
  <Declarante>
    <NIT>{g("nitDecl")}</NIT>
    <DigitoVerificacion>{g("dvDecl")}</DigitoVerificacion>
    <RazonSocial>{g("razonDecl")}</RazonSocial>
    <TipoUsuario>{g("tipoUsuario")}</TipoUsuario>
    <CodigoUsuario>{g("codUsuario")}</CodigoUsuario>
    <NumeroDocumento>{g("numDocDecl")}</NumeroDocumento>
    <NombresApellidos>{g("nombresDecl")}</NombresApellidos>
  </Declarante>

  <!-- SECCIÓN 03: MANIFIESTO Y TRANSPORTE -->
  <Transporte>
    <TipoDeclaracion>{g("tipoDecl")}</TipoDeclaracion>
    <ManifiestoCarga>{g("manifestoCarga")}</ManifiestoCarga>
    <FechaLlegada>{g("fechaLlegada")}</FechaLlegada>
    <LugarIngreso>{g("codLugarIngreso")}</LugarIngreso>
    <ModoTransporte>{g("codModo")}</ModoTransporte>
    <DocumentoTransporte>{g("docTransporte")}</DocumentoTransporte>
    <FechaDocTransporte>{g("fechaDocTransporte")}</FechaDocTransporte>
    <PaisProcedencia>{g("codProcedencia")}</PaisProcedencia>
    <TasaCambio>{g("tasaCambio")}</TasaCambio>
  </Transporte>

  <!-- SECCIÓN 04: MERCANCÍA -->
  <Mercancia>
    <Exportador>{g("nombreExportador")}</Exportador>
    <PaisCompra>{g("codPaisCompra")}</PaisCompra>
    <PaisOrigen>{g("codPaisOrigen")}</PaisOrigen>
    <FormaPago>{g("formaPago")}</FormaPago>
    <SubpartidaArancelaria>{g("subpartida")}</SubpartidaArancelaria>
    <NumeroBultos>{g("numBultos")}</NumeroBultos>
    <Cantidad>{g("cantidad")}</Cantidad>
    <PesoBrutoKg>{g("pesoBruto")}</PesoBrutoKg>
    <PesoNetoKg>{g("pesoNeto")}</PesoNetoKg>
    <ValorFOB moneda="USD">{fob:.2f}</ValorFOB>
    <ValorFletes moneda="USD">{flt:.2f}</ValorFletes>
    <ValorSeguros moneda="USD">{seg:.2f}</ValorSeguros>
    <OtrosGastos moneda="USD">{otr:.2f}</OtrosGastos>
    <AjusteValor moneda="USD">{adj:.2f}</AjusteValor>
    <ValorAduanaCIF moneda="USD">{cif:.2f}</ValorAduanaCIF>
    <Descripcion>{g("descripcion")}</Descripcion>
  </Mercancia>

  <!-- SECCIÓN 05: LIQUIDACIÓN TRIBUTARIA -->
  <Liquidacion>
    <TasaCambioCOPUSD>{trm:.2f}</TasaCambioCOPUSD>
    <ValorAduanaCOP>{cifC:.2f}</ValorAduanaCOP>
    <PorcentajeArancel>{ap}</PorcentajeArancel>
    <Casilla72ArancelCOP>{araC:.2f}</Casilla72ArancelCOP>
    <PorcentajeIVA>{ip}</PorcentajeIVA>
    <Casilla76IVACOP>{ivaC:.2f}</Casilla76IVACOP>
    <PorcentajeImpConsumo>{icp}</PorcentajeImpConsumo>
    <ImpuestoConsumoCOP>{icC:.2f}</ImpuestoConsumoCOP>
    <Casilla980TotalPagarCOP>{total:.2f}</Casilla980TotalPagarCOP>
  </Liquidacion>

</DeclaracionImportacion>'''
        self.xml_content = xml
        self.txt.delete("1.0","end")
        self.txt.insert("1.0", xml)

    def _guardar(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xml", filetypes=[("XML EDI","*.xml")],
            initialfile=f"EDI_510_{date.today().isoformat()}.xml",
            title="Guardar archivo EDI")
        if not path: return
        with open(path,"w",encoding="utf-8") as f:
            f.write(self.xml_content)
        messagebox.showinfo("Guardado",
            f"✅ Archivo EDI guardado:\n{path}\n\n"
            f"Súbalo al SYGA DIAN en:\nhttps://importaciones.dian.gov.co")


# ── Consulta Estado de Levante ────────────────────────────────────────────────
class VentanaConsultaLevante(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Consulta Estado de Levante — SYGA")
        self.geometry("540x420")
        self.configure(bg="white")
        self._build()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="🔍  Consulta Estado de Levante",
                 font=("Arial",13,"bold"), bg="white", fg="#1d4ed8").pack(pady=(16,4))
        tk.Label(self, text="Consulte el estado de su declaración en el SYGA DIAN",
                 font=("Arial",9), bg="white", fg="#64748b").pack(pady=(0,16))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="x", padx=24)

        # Campos de búsqueda
        for label, attr in [
            ("NIT del importador:", "ent_nit"),
            ("No. doc. transporte (Guía/AWB/BL):", "ent_doc"),
            ("No. declaración (si tiene):", "ent_decl"),
        ]:
            tk.Label(frame, text=label, font=("Arial",10), bg="white",
                     fg="#475569").pack(anchor="w", pady=(8,2))
            ent = tk.Entry(frame, font=("Arial",12), relief="flat", bg="#f8fafc",
                           highlightbackground="#cbd5e1", highlightthickness=1)
            ent.pack(fill="x", ipady=5)
            setattr(self, attr, ent)

        tk.Button(frame, text="🔍  Consultar en SYGA DIAN",
                  font=("Arial",11,"bold"), bg="#1d4ed8", fg="white",
                  relief="flat", pady=10, cursor="hand2",
                  command=self._consultar).pack(fill="x", pady=16)

        # Resultado
        self.resultado = tk.Frame(self, bg="#f8fafc",
                                   highlightbackground="#e2e8f0", highlightthickness=1)
        self.resultado.pack(fill="x", padx=24, pady=(0,8))
        self.lbl_resultado = tk.Label(self.resultado,
            text="Ingrese los datos y presione Consultar",
            font=("Arial",10), bg="#f8fafc", fg="#94a3b8",
            wraplength=440, justify="left", pady=12, padx=12)
        self.lbl_resultado.pack(fill="x")

        tk.Label(self, text="⚠️  La consulta abre el portal oficial del SYGA. La información exacta solo está disponible con credenciales DIAN.",
                 font=("Arial",8), bg="white", fg="#94a3b8",
                 justify="center").pack(pady=8)

    def _consultar(self):
        nit  = self.ent_nit.get().strip()
        doc  = self.ent_doc.get().strip()
        decl = self.ent_decl.get().strip()
        if not nit and not doc:
            messagebox.showwarning("Aviso","Ingrese al menos el NIT o el doc. de transporte.")
            return
        # Build SYGA URL with params
        params = []
        if nit:  params.append(f"nit={nit}")
        if doc:  params.append(f"docTransporte={doc}")
        if decl: params.append(f"numDecl={decl}")
        url = "https://importaciones.dian.gov.co/Formulario510/ConsultaEstado?" + "&".join(params)
        self.lbl_resultado.config(
            text=f"✅ Abriendo SYGA DIAN con los datos ingresados...\n\nNIT: {nit}\nDoc. Transporte: {doc}\n\nSi la página no carga automáticamente, ingrese manualmente al portal.",
            fg="#0f172a")
        webbrowser.open("https://importaciones.dian.gov.co")


# ── Solicitud Entrega Urgente ─────────────────────────────────────────────────
class VentanaEntregaUrgente(tk.Toplevel):
    def __init__(self, parent, data):
        super().__init__(parent)
        self.title("Solicitud de Entrega Urgente")
        self.geometry("580x620")
        self.configure(bg="white")
        self.data = data
        self._build()

    def _build(self):
        tk.Frame(self, bg="#dc2626", height=4).pack(fill="x")
        tk.Label(self, text="🚨  Solicitud de Entrega Urgente / Levante Especial",
                 font=("Arial",12,"bold"), bg="white", fg="#dc2626").pack(pady=(14,2))
        tk.Label(self, text="Para mercancía perecedera, medicamentos urgentes o casos especiales",
                 font=("Arial",9), bg="white", fg="#64748b").pack(pady=(0,12))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="both", expand=True, padx=20)

        def field(label, attr, default="", height=1):
            tk.Label(frame, text=label, font=("Arial",10), bg="white",
                     fg="#475569").pack(anchor="w", pady=(8,2))
            if height == 1:
                w = tk.Entry(frame, font=("Arial",11), relief="flat", bg="#f8fafc",
                             highlightbackground="#cbd5e1", highlightthickness=1)
                w.insert(0, default); w.pack(fill="x", ipady=5)
            else:
                w = tk.Text(frame, font=("Arial",11), relief="flat", bg="#f8fafc",
                            height=height, highlightbackground="#cbd5e1", highlightthickness=1)
                w.insert("1.0", default); w.pack(fill="x")
            setattr(self, attr, w)

        d = self.data
        field("Importador (Razón social):", "ent_importador",
              d.get("razonSocial",""))
        field("NIT:", "ent_nit", d.get("nit",""))
        field("No. doc. transporte (Guía/AWB):", "ent_doc",
              d.get("docTransporte",""))
        field("Descripción de la mercancía:", "ent_desc",
              d.get("descripcion",""))

        tk.Label(frame, text="Tipo de urgencia:", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w", pady=(8,2))
        self.tipo_var = tk.StringVar(value="Mercancía perecedera")
        tipos = ["Mercancía perecedera","Medicamentos / insumos médicos urgentes",
                 "Materia prima para producción urgente","Animales vivos",
                 "Órganos / tejidos para trasplante","Otro"]
        ttk.Combobox(frame, values=tipos, textvariable=self.tipo_var,
                     font=("Arial",11), state="readonly").pack(fill="x")

        field("Justificación detallada de la urgencia:", "ent_justif",
              "Indique el motivo por el cual se requiere levante urgente...", height=4)

        field("Nombre del solicitante:", "ent_solicitante", d.get("nombresDecl",""))
        field("Teléfono de contacto:", "ent_telefono", d.get("telefono",""))

        btns = tk.Frame(self, bg="white"); btns.pack(fill="x", padx=20, pady=12)
        tk.Button(btns, text="🖨️  Generar solicitud",
                  font=("Arial",11,"bold"), bg="#dc2626", fg="white",
                  relief="flat", pady=10, cursor="hand2",
                  command=self._generar).pack(side="left", fill="x", expand=True, padx=(0,4))
        tk.Button(btns, text="🌐  Ir al SYGA",
                  font=("Arial",11,"bold"), bg="#0369a1", fg="white",
                  relief="flat", pady=10, cursor="hand2",
                  command=lambda: webbrowser.open("https://importaciones.dian.gov.co")
                  ).pack(side="left", fill="x", expand=True)

    def _get(self, attr):
        w = getattr(self, attr, None)
        if w is None: return ""
        if isinstance(w, tk.Text): return w.get("1.0","end-1c").strip()
        return w.get().strip()

    def _generar(self):
        txt = f"""
SOLICITUD DE ENTREGA URGENTE / LEVANTE ESPECIAL
DIAN Colombia — Seccional {self.data.get("codSeccional","")}
{"="*56}
Fecha de solicitud: {date.today().strftime("%d/%m/%Y")}
Hora: {datetime.now().strftime("%H:%M")}

DATOS DEL IMPORTADOR
{"─"*40}
Importador:       {self._get("ent_importador")}
NIT:              {self._get("ent_nit")}
Doc. transporte:  {self._get("ent_doc")}
Descripción:      {self._get("ent_desc")}

TIPO DE URGENCIA
{"─"*40}
{self.tipo_var.get()}

JUSTIFICACIÓN
{"─"*40}
{self._get("ent_justif")}

DATOS DEL SOLICITANTE
{"─"*40}
Nombre:    {self._get("ent_solicitante")}
Teléfono:  {self._get("ent_telefono")}

{"="*56}
Presente este documento en la Seccional DIAN correspondiente
junto con los documentos soporte de la declaración.
Teléfono DIAN: 57 (1) 307 1111
Portal SYGA: https://importaciones.dian.gov.co
"""
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Texto","*.txt")],
            initialfile=f"EntregaUrgente_{date.today().isoformat()}.txt",
            title="Guardar solicitud")
        if not path: return
        with open(path,"w",encoding="utf-8") as f:
            f.write(txt)
        import os; os.startfile(path)
        messagebox.showinfo("Generado",
            f"✅ Solicitud generada:\n{path}\n\n"
            "Preséntela en la seccional DIAN junto con los documentos soporte.")


# ═══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════
class App(tk.Toplevel):
    def __init__(self, master, user="admin", rol="admin"):
        super().__init__(master)
        self.master = master
        self.title(f"DeclaraFácil 510 — {user} ({rol})")
        self.geometry("1150x800")
        self.minsize(980, 680)
        self.configure(bg="#f1f5f9")
        self.resizable(True, True)
        self.user = user; self.rol = rol
        self.fields = {}
        self._sections = []
        self._nav_btns = []
        self._cliente_id = None
        self._decl_id = None
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self._set_defaults()
        # Auto-backup on start
        self._auto_backup()

    def _on_close(self):
        import sys
        self.master.destroy()
        sys.exit(0)

    def _auto_backup(self):
        try:
            db = get_db_path()
            bak = db.replace(".db", f"_bak_{date.today().isoformat()}.db")
            if not os.path.exists(bak):
                shutil.copy2(db, bak)
        except: pass

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
        sb_outer = tk.Frame(body, bg="#0f1724", width=220)
        sb_outer.pack(side="left", fill="y"); sb_outer.pack_propagate(False)

        # Scrollable sidebar
        sb_canvas = tk.Canvas(sb_outer, bg="#0f1724", highlightthickness=0, width=220)
        sb_scroll = ttk.Scrollbar(sb_outer, orient="vertical", command=sb_canvas.yview)
        sb_canvas.configure(yscrollcommand=sb_scroll.set)
        sb_scroll.pack(side="right", fill="y")
        sb_canvas.pack(side="left", fill="both", expand=True)
        sb = tk.Frame(sb_canvas, bg="#0f1724")
        sb_win = sb_canvas.create_window((0,0), window=sb, anchor="nw")
        sb.bind("<Configure>", lambda e: sb_canvas.configure(scrollregion=sb_canvas.bbox("all")))
        sb_canvas.bind("<Configure>", lambda e: sb_canvas.itemconfig(sb_win, width=e.width))
        sb_canvas.bind("<MouseWheel>", lambda e: sb_canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
        sb.bind("<MouseWheel>", lambda e: sb_canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

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
            b.bind("<MouseWheel>", lambda e: sb_canvas.yview_scroll(int(-1*(e.delta/120)),"units"))
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
            ("👥  Clientes",          self._abrir_clientes,      "#1e3a5f", "white"),
            ("📋  Historial",          self._abrir_historial,     "#1e3a5f", "white"),
            ("📊  Estadísticas",       self._abrir_stats,         "#1e3a5f", "white"),
            ("⏰  Plazos y alertas",   self._abrir_plazos,        "#7c2d12", "white"),
            ("✅  Checklist docs",     self._abrir_checklist,     "#14532d", "white"),
            ("🧾  Recibo de pago",     self._abrir_rop,           "#1e3a5f", "white"),
            ("📤  Generar EDI/XML",    self._abrir_edi,           "#0c4a6e", "white"),
            ("🔍  Consultar levante",  self._abrir_consulta,      "#1e3a5f", "white"),
            ("🚨  Entrega urgente",    self._abrir_urgente,       "#7f1d1d", "white"),
            ("⚖️  Calculadora multas", self._abrir_multas,        "#7c2d12", "white"),
            ("📝  Poder/Autorización", self._abrir_poder,         "#1e3a5f", "white"),
            ("🌐  Ir al SYGA",         self._abrir_syga,          "#0369a1", "white"),
            ("🔄  Actualizar TRM",     self._update_trm,          "#7c3aed", "white"),
            ("📥  Cargar Excel",       self._load_excel,          "#0f766e", "white"),
            ("📋  Subpartidas",        self._subpartidas,         "#92400e", "white"),
            ("💾  Guardar decl.",      self._guardar_decl,        "#0f4c35", "white"),
            ("📄  Generar PDF",        self._generate,            "#3b82f6", "white"),
            ("⚙️  Configuración",      self._abrir_config,        "#1e2535", "#94a3b8"),
            ("🗑️  Limpiar",            self._clear,               "#1e2535", "#64748b"),
        ]:
            btn = tk.Button(sb, text=text, font=("Arial",9,"bold"), bg=bg, fg=fg,
                      relief="flat", bd=0, padx=8, pady=5, cursor="hand2",
                      activebackground=bg, activeforeground=fg,
                      command=cmd)
            btn.pack(fill="x", padx=8, pady=1)
            btn.bind("<MouseWheel>", lambda e: sb_canvas.yview_scroll(int(-1*(e.delta/120)),"units"))

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

        tk.Frame(self.inner,bg="#f1f5f9",height=60).pack()

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
            trm, fecha_trm = None, None
            HEADERS = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                              "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                "Accept": "application/json",
            }

            # Intento 1: dataset oficial datos.gov.co (32sa-8pi3), retrocediendo hasta 10 días
            # por si hoy es festivo/fin de semana y no hay TRM publicada.
            try:
                from datetime import timedelta
                for offset in range(0, 10):
                    d = (date.today() - timedelta(days=offset)).strftime("%Y-%m-%d")
                    params = urllib.parse.urlencode({"vigenciadesde": d})
                    url = f"https://www.datos.gov.co/resource/32sa-8pi3.json?{params}"
                    req = urllib.request.Request(url, headers=HEADERS)
                    with urllib.request.urlopen(req, timeout=6) as r:
                        data = json.loads(r.read())
                    if data and "valor" in data[0]:
                        trm = float(str(data[0]["valor"]).replace(",", "."))
                        fecha_trm = d
                        break
            except Exception:
                pass

            # Intento 2: dataset alterno mcec-87by con orden correcto (sin espacios crudos)
            if not trm:
                try:
                    params2 = urllib.parse.urlencode({"$limit": "1", "$order": "vigenciadesde DESC"})
                    url2 = f"https://www.datos.gov.co/resource/mcec-87by.json?{params2}"
                    req2 = urllib.request.Request(url2, headers=HEADERS)
                    with urllib.request.urlopen(req2, timeout=6) as r2:
                        data2 = json.loads(r2.read())
                    if data2 and "valor" in data2[0]:
                        trm = float(str(data2[0]["valor"]).replace(",", "."))
                        fecha_trm = str(data2[0].get("vigenciadesde",""))[:10]
                except Exception:
                    pass

            # Intento 3: API pública de tipo de cambio (respaldo, no oficial DIAN pero útil)
            if not trm:
                try:
                    req3 = urllib.request.Request("https://open.er-api.com/v6/latest/USD", headers=HEADERS)
                    with urllib.request.urlopen(req3, timeout=6) as r3:
                        data3 = json.loads(r3.read())
                    cop = data3.get("rates", {}).get("COP")
                    if cop:
                        trm = float(cop)
                        fecha_trm = date.today().strftime("%Y-%m-%d") + " (referencial)"
                except Exception:
                    pass

            if trm and trm > 0:
                self.after(0, lambda t=trm, d=fecha_trm: self._apply_trm(t, d))
            else:
                self.after(0, lambda: self.lbl_trm_status.config(
                    text="Sin conexión a internet.\nIngrese la TRM manualmente.", fg="#ef4444"))
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

    # ── EDI ──────────────────────────────────────────────────────────────────
    def _abrir_edi(self):
        data = {k: self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        if not self._get_field("nit"):
            messagebox.showwarning("Aviso","Ingrese los datos del formulario antes de generar el EDI.")
            return
        VentanaEDI(self, data)

    # ── Consulta levante ──────────────────────────────────────────────────────
    def _abrir_consulta(self):
        win = VentanaConsultaLevante(self)
        win.ent_nit.insert(0, self._get_field("nit"))
        win.ent_doc.insert(0, self._get_field("docTransporte"))

    # ── Entrega urgente ───────────────────────────────────────────────────────
    def _abrir_urgente(self):
        data = {k: self._get_field(k) for k in self.fields}
        VentanaEntregaUrgente(self, data)

    # ── Plazos ───────────────────────────────────────────────────────────────
    def _abrir_plazos(self):
        VentanaPlazos(self)

    # ── Checklist ─────────────────────────────────────────────────────────────
    def _abrir_checklist(self):
        cliente = self._get_field("razonSocial") or "Sin cliente"
        doc     = self._get_field("docTransporte") or "Sin doc."
        VentanaChecklist(self, decl_info=f"Cliente: {cliente}  |  Doc: {doc}")

    # ── ROP ───────────────────────────────────────────────────────────────────
    def _abrir_rop(self):
        data = {k: self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        trm = float(self._get_field("tasaCambio") or 4150)
        total = getattr(self,"_total_cop",0)
        if not self._get_field("nit"):
            messagebox.showwarning("Aviso","Ingrese al menos el NIT y los valores antes de generar el ROP.")
            return
        VentanaROP(self, data, total, trm)

    # ── SYGA ──────────────────────────────────────────────────────────────────
    def _abrir_syga(self):
        webbrowser.open("https://importaciones.dian.gov.co")

    # ── Estadísticas ──────────────────────────────────────────────────────────
    def _abrir_stats(self):
        VentanaEstadisticas(self)

    # ── Multas ────────────────────────────────────────────────────────────────
    def _abrir_multas(self):
        VentanaMultas(self)

    # ── Poder ─────────────────────────────────────────────────────────────────
    def _abrir_poder(self):
        data = {k: self._get_field(k) for k in self.fields}
        VentanaPoder(self, data)

    # ── Configuración ─────────────────────────────────────────────────────────
    def _abrir_config(self):
        VentanaConfigAgencia(self)

    # ── Clear ─────────────────────────────────────────────────────────────────
    def _clear(self):
        if messagebox.askyesno("Confirmar","¿Limpiar todos los campos?"):
            self._cliente_id = None; self._decl_id = None
            self.lbl_cliente_badge.config(text="Sin cliente seleccionado")
            self._set_defaults()


# __main__ moved to end of file


# ═══════════════════════════════════════════════════════════════════════════════
# PANTALLA DE INICIO / SPLASH
# ═══════════════════════════════════════════════════════════════════════════════
class SplashScreen(tk.Toplevel):
    def __init__(self, parent, on_ready):
        super().__init__(parent)
        self.on_ready = on_ready
        self.overrideredirect(True)
        w, h = 480, 300
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(bg="#0f172a")
        self._build()
        self.after(2800, self._cerrar)

    def _build(self):
        # Logo area
        logo = tk.Frame(self, bg="#1d4ed8", width=80, height=80)
        logo.place(x=200, y=40)
        tk.Label(logo, text="510", font=("Arial",22,"bold"),
                 bg="#1d4ed8", fg="white").place(relx=0.5, rely=0.5, anchor="center")

        agencia = db_fetch("SELECT valor FROM config WHERE clave='agencia_nombre'")
        nombre = agencia[0][0] if agencia else "Mi Agencia de Aduanas"

        tk.Label(self, text="DeclaraFácil", font=("Arial",22,"bold"),
                 bg="#0f172a", fg="white").place(x=240, y=140, anchor="center")
        tk.Label(self, text=nombre, font=("Arial",11),
                 bg="#0f172a", fg="#60a5fa").place(x=240, y=172, anchor="center")
        tk.Label(self, text="Sistema de Gestión Aduanera · DIAN Colombia",
                 font=("Arial",8), bg="#0f172a", fg="#475569").place(x=240, y=196, anchor="center")

        # Progress bar
        self.bar_bg = tk.Frame(self, bg="#1e293b", width=360, height=4)
        self.bar_bg.place(x=60, y=240)
        self.bar = tk.Frame(self, bg="#1d4ed8", width=0, height=4)
        self.bar.place(x=60, y=240)

        self.lbl_status = tk.Label(self, text="Iniciando...", font=("Arial",8),
                                    bg="#0f172a", fg="#475569")
        self.lbl_status.place(x=240, y=256, anchor="center")

        tk.Label(self, text="v2.0", font=("Arial",7),
                 bg="#0f172a", fg="#1e293b").place(x=460, y=286, anchor="se")

        self._animar(0)

    def _animar(self, step):
        msgs = ["Cargando base de datos...","Verificando licencia...","Preparando formularios...","Listo ✓"]
        if step <= 100:
            self.bar.config(width=int(3.6*step))
            self.lbl_status.config(text=msgs[min(step//30, 3)])
            self.after(22, lambda: self._animar(step+2))

    def _cerrar(self):
        self.destroy()
        self.on_ready()


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Login
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaLogin(tk.Toplevel):
    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.on_success = on_success
        self.title("Iniciar sesión — DeclaraFácil 510")
        self.geometry("380x320")
        self.resizable(False, False)
        self.configure(bg="white")
        self.protocol("WM_DELETE_WINDOW", self._cancelar)
        self._build()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        agencia = db_fetch("SELECT valor FROM config WHERE clave='agencia_nombre'")
        nombre = agencia[0][0] if agencia else "DeclaraFácil 510"
        tk.Label(self, text="🔐  Iniciar Sesión", font=("Arial",14,"bold"),
                 bg="white", fg="#1d4ed8").pack(pady=(20,4))
        tk.Label(self, text=nombre, font=("Arial",9),
                 bg="white", fg="#64748b").pack(pady=(0,20))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="x", padx=40)
        tk.Label(frame, text="Usuario:", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w")
        self.ent_user = tk.Entry(frame, font=("Arial",12), relief="flat",
                                  highlightbackground="#cbd5e1", highlightthickness=1)
        self.ent_user.pack(fill="x", ipady=6, pady=(2,12))
        self.ent_user.insert(0,"admin")

        tk.Label(frame, text="Contraseña:", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w")
        self.ent_pwd = tk.Entry(frame, font=("Arial",12), relief="flat", show="•",
                                 highlightbackground="#cbd5e1", highlightthickness=1)
        self.ent_pwd.pack(fill="x", ipady=6, pady=(2,0))
        self.ent_pwd.bind("<Return>", lambda e: self._login())

        self.lbl_err = tk.Label(frame, text="", font=("Arial",9),
                                 bg="white", fg="#dc2626")
        self.lbl_err.pack(pady=4)

        tk.Button(frame, text="Iniciar sesión", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10,
                  cursor="hand2", command=self._login).pack(fill="x", pady=(4,0))

    def _login(self):
        user = self.ent_user.get().strip()
        pwd  = hashlib.sha256(self.ent_pwd.get().encode()).hexdigest()
        rows = db_fetch("SELECT id,rol FROM usuarios WHERE username=? AND password_hash=? AND activo=1",
                        (user, pwd))
        if rows:
            self.destroy()
            self.on_success(user, rows[0][1])
        else:
            self.lbl_err.config(text="Usuario o contraseña incorrectos")
            self.ent_pwd.delete(0,"end")

    def _cancelar(self):
        import sys; sys.exit(0)


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Configuración de la Agencia
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaConfigAgencia(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Configuración de la Agencia")
        self.geometry("520x580")
        self.configure(bg="white")
        self.resizable(False, False)
        self.fields = {}
        self._build()
        self._cargar()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="⚙️  Configuración de la Agencia",
                 font=("Arial",13,"bold"), bg="white", fg="#1d4ed8").pack(pady=(16,4))

        nb = ttk.Notebook(self); nb.pack(fill="both", expand=True, padx=16, pady=8)

        # Tab 1: Datos agencia
        t1 = tk.Frame(nb, bg="white"); nb.add(t1, text="  Agencia  ")
        for label, key in [
            ("Nombre de la agencia:", "agencia_nombre"),
            ("NIT de la agencia:", "agencia_nit"),
            ("Teléfono:", "agencia_tel"),
            ("Dirección:", "agencia_dir"),
        ]:
            tk.Label(t1, text=label, font=("Arial",10), bg="white",
                     fg="#475569").pack(anchor="w", padx=20, pady=(12,2))
            w = tk.Entry(t1, font=("Arial",11), relief="flat",
                         highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x", padx=20, ipady=6)
            self.fields[key] = w

        # Tab 2: Usuarios
        t2 = tk.Frame(nb, bg="white"); nb.add(t2, text="  Usuarios  ")
        tk.Label(t2, text="Usuarios del sistema:", font=("Arial",10,"bold"),
                 bg="white", fg="#1d4ed8").pack(anchor="w", padx=20, pady=(16,8))
        cols = ("Usuario","Rol","Estado")
        self.tree_users = ttk.Treeview(t2, columns=cols, show="headings", height=6)
        for c in cols:
            self.tree_users.heading(c, text=c)
            self.tree_users.column(c, width={"Usuario":160,"Rol":100,"Estado":80}[c])
        self.tree_users.pack(fill="x", padx=20)
        self._cargar_usuarios()

        btns_u = tk.Frame(t2, bg="white"); btns_u.pack(fill="x", padx=20, pady=8)
        tk.Button(btns_u, text="➕ Nuevo usuario", font=("Arial",9,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", padx=10, pady=6,
                  cursor="hand2", command=self._nuevo_usuario).pack(side="left", padx=(0,4))
        tk.Button(btns_u, text="🔑 Cambiar contraseña", font=("Arial",9,"bold"),
                  bg="#0f766e", fg="white", relief="flat", padx=10, pady=6,
                  cursor="hand2", command=self._cambiar_pwd).pack(side="left")

        # Tab 3: Licencia
        t3 = tk.Frame(nb, bg="white"); nb.add(t3, text="  Licencia  ")
        self._build_licencia(t3)

        # Tab 4: Respaldo
        t4 = tk.Frame(nb, bg="white"); nb.add(t4, text="  Respaldo  ")
        self._build_respaldo(t4)

        # Save button
        tk.Button(self, text="💾  Guardar configuración", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10,
                  cursor="hand2", command=self._guardar).pack(fill="x", padx=16, pady=(0,16))

    def _build_licencia(self, parent):
        lic = db_fetch("SELECT valor FROM config WHERE clave='licencia_activa'")
        activa = lic[0][0] == "1" if lic else False
        cnt = db_fetch("SELECT COUNT(*) FROM declaraciones")[0][0]
        max_t = db_fetch("SELECT valor FROM config WHERE clave='max_decl_trial'")
        max_trial = int(max_t[0][0]) if max_t else 5

        estado_color = "#22c55e" if activa else "#f59e0b"
        estado_text  = "✅ LICENCIA ACTIVA" if activa else f"⚠️ MODO PRUEBA ({cnt}/{max_trial} declaraciones)"

        tk.Label(parent, text=estado_text, font=("Arial",12,"bold"),
                 bg="white", fg=estado_color).pack(pady=(20,8))

        if not activa:
            tk.Label(parent, text=f"En modo prueba puede crear hasta {max_trial} declaraciones.\nIngrese su clave de activación para desbloquear.",
                     font=("Arial",9), bg="white", fg="#64748b",
                     justify="center").pack(pady=(0,16))

        tk.Label(parent, text="Clave de activación:", font=("Arial",10),
                 bg="white", fg="#475569").pack(padx=20, anchor="w")
        self.ent_lic = tk.Entry(parent, font=("Arial",12), relief="flat",
                                 highlightbackground="#cbd5e1", highlightthickness=1)
        lic_key = db_fetch("SELECT valor FROM config WHERE clave='licencia_key'")
        if lic_key and lic_key[0][0]:
            self.ent_lic.insert(0, lic_key[0][0])
        self.ent_lic.pack(fill="x", padx=20, ipady=6, pady=(2,12))
        tk.Button(parent, text="🔑  Activar licencia", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=self._activar).pack(fill="x", padx=20)

        tk.Label(parent,
                 text="Para obtener su clave de activación contáctenos:\nbentjake15@gmail.com",
                 font=("Arial",9), bg="white", fg="#94a3b8",
                 justify="center").pack(pady=16)

    def _build_respaldo(self, parent):
        tk.Label(parent, text="💾  Respaldo de Base de Datos",
                 font=("Arial",11,"bold"), bg="white", fg="#1d4ed8").pack(pady=(20,8))
        tk.Label(parent, text="Guarde una copia de seguridad de todos sus clientes,\ndeclaraciones e historial.",
                 font=("Arial",9), bg="white", fg="#64748b", justify="center").pack(pady=(0,16))

        db_path = get_db_path()
        size = os.path.getsize(db_path) if os.path.exists(db_path) else 0
        cnt_c = db_fetch("SELECT COUNT(*) FROM clientes")[0][0]
        cnt_d = db_fetch("SELECT COUNT(*) FROM declaraciones")[0][0]

        info = tk.Frame(parent, bg="#f8fafc", highlightbackground="#e2e8f0", highlightthickness=1)
        info.pack(fill="x", padx=20, pady=(0,16))
        for label, val in [("Clientes registrados:", str(cnt_c)),
                            ("Declaraciones:", str(cnt_d)),
                            ("Tamaño DB:", f"{size/1024:.1f} KB")]:
            r = tk.Frame(info, bg="#f8fafc"); r.pack(fill="x", padx=12, pady=3)
            tk.Label(r, text=label, font=("Arial",9), bg="#f8fafc", fg="#64748b").pack(side="left")
            tk.Label(r, text=val, font=("Arial",9,"bold"), bg="#f8fafc", fg="#0f172a").pack(side="right")

        tk.Button(parent, text="📤  Guardar respaldo ahora", font=("Arial",11,"bold"),
                  bg="#0f766e", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=self._respaldar).pack(fill="x", padx=20, pady=(0,8))
        tk.Button(parent, text="📥  Restaurar respaldo", font=("Arial",11,"bold"),
                  bg="#7c3aed", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=self._restaurar).pack(fill="x", padx=20)

    def _cargar(self):
        for key, w in self.fields.items():
            rows = db_fetch("SELECT valor FROM config WHERE clave=?", (key,))
            if rows: w.delete(0,"end"); w.insert(0, rows[0][0])

    def _cargar_usuarios(self):
        self.tree_users.delete(*self.tree_users.get_children())
        for r in db_fetch("SELECT username,rol,activo FROM usuarios"):
            self.tree_users.insert("","end", values=(r[0],r[1],"Activo" if r[2] else "Inactivo"))

    def _guardar(self):
        for key, w in self.fields.items():
            db_exec("UPDATE config SET valor=? WHERE clave=?", (w.get().strip(), key))
        messagebox.showinfo("Guardado","✅ Configuración guardada.")

    def _activar(self):
        key = self.ent_lic.get().strip()
        # Validación simple de formato: prefijo DF510- y longitud mínima.
        if key.upper().startswith("DF510-") and len(key) >= 20:
            db_exec("UPDATE config SET valor='1' WHERE clave='licencia_activa'")
            db_exec("UPDATE config SET valor=? WHERE clave='licencia_key'", (key,))
            messagebox.showinfo("Activado","✅ Licencia activada correctamente.\nReinicie la aplicación.")
        else:
            messagebox.showerror("Clave inválida",
                "La clave ingresada no es válida.\nContacte a bentjake15@gmail.com para obtener su clave.")

    def _nuevo_usuario(self):
        win = tk.Toplevel(self); win.title("Nuevo usuario")
        win.geometry("340x280"); win.configure(bg="white"); win.resizable(False,False)
        tk.Frame(win, bg="#1d4ed8", height=3).pack(fill="x")
        tk.Label(win, text="Nuevo Usuario", font=("Arial",12,"bold"),
                 bg="white", fg="#1d4ed8").pack(pady=(14,12))
        frame = tk.Frame(win, bg="white"); frame.pack(fill="x", padx=24)
        entries = {}
        for label, key, show in [("Usuario:","user",""),("Contraseña:","pwd","•"),("Repetir contraseña:","pwd2","•")]:
            tk.Label(frame, text=label, font=("Arial",10), bg="white", fg="#475569").pack(anchor="w")
            e = tk.Entry(frame, font=("Arial",11), show=show, relief="flat",
                         highlightbackground="#cbd5e1", highlightthickness=1)
            e.pack(fill="x", ipady=5, pady=(2,8)); entries[key] = e
        tk.Label(frame, text="Rol:", font=("Arial",10), bg="white", fg="#475569").pack(anchor="w")
        rol_var = tk.StringVar(value="operador")
        ttk.Combobox(frame, values=["operador","admin"], textvariable=rol_var,
                     state="readonly", font=("Arial",11)).pack(fill="x")
        def crear():
            u=entries["user"].get().strip(); p=entries["pwd"].get(); p2=entries["pwd2"].get()
            if not u or not p: messagebox.showwarning("Aviso","Complete todos los campos."); return
            if p != p2: messagebox.showwarning("Aviso","Las contraseñas no coinciden."); return
            try:
                db_exec("INSERT INTO usuarios(username,password_hash,rol) VALUES(?,?,?)",
                        (u, hashlib.sha256(p.encode()).hexdigest(), rol_var.get()))
                self._cargar_usuarios(); win.destroy()
                messagebox.showinfo("Creado",f"✅ Usuario '{u}' creado.")
            except: messagebox.showerror("Error","El usuario ya existe.")
        tk.Button(win, text="Crear usuario", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=crear).pack(fill="x", padx=24, pady=12)

    def _cambiar_pwd(self):
        sel = self.tree_users.selection()
        if not sel: messagebox.showwarning("Aviso","Seleccione un usuario."); return
        user = self.tree_users.item(sel[0])["values"][0]
        win = tk.Toplevel(self); win.title("Cambiar contraseña")
        win.geometry("320x220"); win.configure(bg="white"); win.resizable(False,False)
        tk.Label(win, text=f"Cambiar contraseña: {user}", font=("Arial",11,"bold"),
                 bg="white", fg="#1d4ed8").pack(pady=(16,12))
        frame = tk.Frame(win, bg="white"); frame.pack(fill="x", padx=24)
        entries = {}
        for label, key in [("Nueva contraseña:","pwd"),("Repetir:","pwd2")]:
            tk.Label(frame, text=label, font=("Arial",10), bg="white", fg="#475569").pack(anchor="w")
            e = tk.Entry(frame, font=("Arial",11), show="•", relief="flat",
                         highlightbackground="#cbd5e1", highlightthickness=1)
            e.pack(fill="x", ipady=5, pady=(2,8)); entries[key] = e
        def cambiar():
            p=entries["pwd"].get(); p2=entries["pwd2"].get()
            if not p: return
            if p != p2: messagebox.showwarning("Aviso","Las contraseñas no coinciden."); return
            db_exec("UPDATE usuarios SET password_hash=? WHERE username=?",
                    (hashlib.sha256(p.encode()).hexdigest(), user))
            win.destroy(); messagebox.showinfo("Cambiado","✅ Contraseña actualizada.")
        tk.Button(win, text="Cambiar", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=8,
                  cursor="hand2", command=cambiar).pack(fill="x", padx=24)

    def _respaldar(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".db",
            filetypes=[("Base de datos","*.db")],
            initialfile=f"Respaldo_DeclaraFacil_{date.today().isoformat()}.db",
            title="Guardar respaldo")
        if not path: return
        shutil.copy2(get_db_path(), path)
        messagebox.showinfo("Respaldo","✅ Respaldo guardado correctamente.")

    def _restaurar(self):
        path = filedialog.askopenfilename(
            title="Seleccionar respaldo",
            filetypes=[("Base de datos","*.db")])
        if not path: return
        if messagebox.askyesno("Confirmar",
            "⚠️ Esto reemplazará TODOS los datos actuales.\n¿Está seguro?"):
            shutil.copy2(path, get_db_path())
            messagebox.showinfo("Restaurado","✅ Datos restaurados. Reinicie la aplicación.")


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Panel de Estadísticas
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaEstadisticas(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Panel de Estadísticas")
        self.geometry("720x560")
        self.configure(bg="white")
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg="#1d4ed8", height=48); hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="📊  Panel de Estadísticas",
                 font=("Arial",13,"bold"), bg="#1d4ed8", fg="white").pack(side="left", padx=16, pady=10)

        # KPI cards
        kpi_frame = tk.Frame(self, bg="#f8fafc"); kpi_frame.pack(fill="x", padx=16, pady=12)

        total_decl  = db_fetch("SELECT COUNT(*) FROM declaraciones")[0][0]
        total_cli   = db_fetch("SELECT COUNT(*) FROM clientes")[0][0]
        total_cop   = db_fetch("SELECT SUM(total_cop) FROM declaraciones")[0][0] or 0
        mes_actual  = date.today().strftime("%Y-%m")
        decl_mes    = db_fetch("SELECT COUNT(*) FROM declaraciones WHERE fecha LIKE ?", (f"{mes_actual}%",))[0][0]

        for i,(label,val,color) in enumerate([
            ("Total declaraciones", str(total_decl), "#1d4ed8"),
            ("Clientes registrados", str(total_cli),  "#0f766e"),
            (f"Declaraciones {date.today().strftime('%b %Y')}", str(decl_mes), "#7c3aed"),
            ("Total liquidado", fmt_cop(total_cop)+" COP", "#dc2626"),
        ]):
            card = tk.Frame(kpi_frame, bg="white",
                            highlightbackground="#e2e8f0", highlightthickness=1)
            card.grid(row=0, column=i, padx=6, pady=4, sticky="ew")
            kpi_frame.columnconfigure(i, weight=1)
            tk.Label(card, text=val, font=("Arial",18,"bold"),
                     bg="white", fg=color).pack(pady=(12,2))
            tk.Label(card, text=label, font=("Arial",8),
                     bg="white", fg="#94a3b8").pack(pady=(0,12))

        # Top clientes
        tk.Label(self, text="Top 5 Clientes por declaraciones",
                 font=("Arial",10,"bold"), bg="white", fg="#1d4ed8").pack(anchor="w", padx=16, pady=(8,4))
        cols = ("Cliente","NIT","Declaraciones","Total COP")
        tree = ttk.Treeview(self, columns=cols, show="headings", height=5)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width={"Cliente":260,"NIT":120,"Declaraciones":110,"Total COP":160}[c])
        top = db_fetch("""SELECT c.razon_social, c.nit, COUNT(d.id), SUM(d.total_cop)
                          FROM clientes c LEFT JOIN declaraciones d ON c.id=d.cliente_id
                          GROUP BY c.id ORDER BY COUNT(d.id) DESC LIMIT 5""")
        for r in top:
            tree.insert("","end", values=(r[0],r[1],r[2],fmt_cop(r[3] or 0)+" COP"))
        tree.pack(fill="x", padx=16)

        # Declaraciones por mes
        tk.Label(self, text="Declaraciones por mes (últimos 6 meses)",
                 font=("Arial",10,"bold"), bg="white", fg="#1d4ed8").pack(anchor="w", padx=16, pady=(16,4))
        cols2 = ("Mes","Declaraciones","Total COP","Promedio COP")
        tree2 = ttk.Treeview(self, columns=cols2, show="headings", height=6)
        for c in cols2:
            tree2.heading(c, text=c)
            tree2.column(c, width={"Mes":120,"Declaraciones":120,"Total COP":180,"Promedio COP":180}[c])
        meses = db_fetch("""SELECT strftime('%Y-%m',fecha) as mes, COUNT(*), SUM(total_cop), AVG(total_cop)
                            FROM declaraciones WHERE fecha IS NOT NULL
                            GROUP BY mes ORDER BY mes DESC LIMIT 6""")
        for r in meses:
            tree2.insert("","end", values=(r[0],r[1],fmt_cop(r[2] or 0)+" COP",fmt_cop(r[3] or 0)+" COP"))
        tree2.pack(fill="x", padx=16, pady=(0,16))


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Calculadora de Multas
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaMultas(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Calculadora de Multas y Sanciones DIAN")
        self.geometry("560x520")
        self.configure(bg="white")
        self._build()

    def _build(self):
        tk.Frame(self, bg="#dc2626", height=4).pack(fill="x")
        tk.Label(self, text="⚖️  Calculadora de Multas y Sanciones",
                 font=("Arial",13,"bold"), bg="white", fg="#dc2626").pack(pady=(14,2))
        tk.Label(self, text="Estimación basada en el Estatuto Aduanero (Decreto 1165/2019)",
                 font=("Arial",9), bg="white", fg="#64748b").pack(pady=(0,12))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="x", padx=24)

        tk.Label(frame, text="Tipo de infracción:", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w", pady=(8,2))
        self.tipo_var = tk.StringVar()
        tipos = [
            "Presentación extemporánea de declaración",
            "Corrección de declaración",
            "Inexactitud en valores declarados",
            "No presentación de declaración",
            "Infracción de importación (Art. 596)",
        ]
        cb = ttk.Combobox(frame, values=tipos, textvariable=self.tipo_var,
                          font=("Arial",11), state="readonly")
        cb.pack(fill="x"); cb.set(tipos[0])
        cb.bind("<<ComboboxSelected>>", lambda e: self._calcular())

        tk.Label(frame, text="Valor de los tributos (COP):", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w", pady=(12,2))
        self.ent_tributos = tk.Entry(frame, font=("Arial",12), relief="flat",
                                      highlightbackground="#cbd5e1", highlightthickness=1)
        self.ent_tributos.pack(fill="x", ipady=6)
        self.ent_tributos.bind("<KeyRelease>", lambda e: self._calcular())

        tk.Label(frame, text="Días de retraso (si aplica):", font=("Arial",10),
                 bg="white", fg="#475569").pack(anchor="w", pady=(12,2))
        self.ent_dias = tk.Entry(frame, font=("Arial",12), relief="flat",
                                  highlightbackground="#cbd5e1", highlightthickness=1)
        self.ent_dias.insert(0,"0"); self.ent_dias.pack(fill="x", ipady=6)
        self.ent_dias.bind("<KeyRelease>", lambda e: self._calcular())

        # Resultado
        self.res_frame = tk.Frame(self, bg="#fef2f2",
                                   highlightbackground="#fecaca", highlightthickness=1)
        self.res_frame.pack(fill="x", padx=24, pady=16)
        self.lbl_tipo_mul = tk.Label(self.res_frame, text="",
                                      font=("Arial",9), bg="#fef2f2", fg="#7f1d1d")
        self.lbl_tipo_mul.pack(anchor="w", padx=12, pady=(10,2))
        self.lbl_multa = tk.Label(self.res_frame, text="$0 COP",
                                   font=("Arial",18,"bold"), bg="#fef2f2", fg="#dc2626")
        self.lbl_multa.pack(pady=(0,4))
        self.lbl_base = tk.Label(self.res_frame, text="",
                                  font=("Arial",8), bg="#fef2f2", fg="#94a3b8")
        self.lbl_base.pack(pady=(0,10))

        tk.Label(self, text="⚠️  Valores de referencia. Consulte con un abogado aduanero para casos específicos.",
                 font=("Arial",8), bg="white", fg="#94a3b8",
                 wraplength=500, justify="center").pack(pady=8)

        self._calcular()

    def _calcular(self):
        try:
            tributos = float(self.ent_tributos.get().replace(".","").replace(",",".") or 0)
            dias     = int(self.ent_dias.get() or 0)
        except: return

        tipo  = self.tipo_var.get()
        uvt   = 47065  # UVT 2025
        multa = 0
        base  = ""

        if "extemporánea" in tipo:
            # 1.5% por mes o fracción, mínimo 10 UVT
            meses  = max(1, (dias + 29) // 30)
            multa  = max(tributos * 0.015 * meses, 10 * uvt)
            base   = f"1.5% × {meses} mes(es) sobre tributos. Mín. 10 UVT ({fmt_cop(10*uvt)} COP)"
        elif "Corrección" in tipo:
            multa = max(tributos * 0.10, 10 * uvt)
            base  = f"10% sobre tributos. Mínimo 10 UVT ({fmt_cop(10*uvt)} COP)"
        elif "Inexactitud" in tipo:
            multa = max(tributos * 1.60, 20 * uvt)
            base  = f"160% sobre tributos. Mínimo 20 UVT ({fmt_cop(20*uvt)} COP)"
        elif "No presentación" in tipo:
            multa = max(tributos * 0.20, 20 * uvt)
            base  = f"20% sobre tributos. Mínimo 20 UVT ({fmt_cop(20*uvt)} COP)"
        elif "Art. 596" in tipo:
            multa = max(200 * uvt, tributos * 0.20)
            base  = f"Mínimo 200 UVT ({fmt_cop(200*uvt)} COP)"

        self.lbl_multa.config(text=fmt_cop(multa)+" COP")
        self.lbl_base.config(text=base)
        self.lbl_tipo_mul.config(text=f"Base de cálculo: {tipo}")


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANA: Generador de Poder / Autorización
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaPoder(tk.Toplevel):
    def __init__(self, parent, data):
        super().__init__(parent)
        self.title("Generador de Poder / Autorización")
        self.geometry("600x560")
        self.configure(bg="white")
        self.data = data
        self._build()

    def _build(self):
        tk.Frame(self, bg="#1d4ed8", height=4).pack(fill="x")
        tk.Label(self, text="📝  Poder / Autorización para Agencia de Aduanas",
                 font=("Arial",12,"bold"), bg="white", fg="#1d4ed8").pack(pady=(14,4))

        frame = tk.Frame(self, bg="white"); frame.pack(fill="both", expand=True, padx=20)
        agencia = db_fetch("SELECT valor FROM config WHERE clave='agencia_nombre'")
        ag_nombre = agencia[0][0] if agencia else "LA AGENCIA DE ADUANAS"
        agencia_nit = db_fetch("SELECT valor FROM config WHERE clave='agencia_nit'")
        ag_nit = agencia_nit[0][0] if agencia_nit else "NIT AGENCIA"

        fields = {}
        for label, key, default in [
            ("Nombre del poderdante (importador):", "poderdante", self.data.get("razonSocial","")),
            ("NIT / Cédula del poderdante:", "pod_nit", self.data.get("nit","")),
            ("Descripción de la mercancía:", "mercancias", self.data.get("descripcion","")),
            ("No. doc. transporte:", "doc_transp", self.data.get("docTransporte","")),
        ]:
            tk.Label(frame, text=label, font=("Arial",10), bg="white", fg="#475569").pack(anchor="w", pady=(8,2))
            w = tk.Entry(frame, font=("Arial",11), relief="flat",
                         highlightbackground="#cbd5e1", highlightthickness=1)
            w.insert(0, default); w.pack(fill="x", ipady=5); fields[key] = w

        def generar():
            pod   = fields["poderdante"].get().strip()
            pnit  = fields["pod_nit"].get().strip()
            merc  = fields["mercancias"].get().strip()
            doc   = fields["doc_transp"].get().strip()
            fecha = date.today().strftime("%d de %B de %Y")
            txt = f"""
                        PODER ESPECIAL

Yo, {pod}, identificado(a) con NIT/CC No. {pnit}, actuando en mi propio nombre
y representación, por medio del presente documento OTORGO PODER ESPECIAL,
amplio y suficiente a {ag_nombre} — NIT {ag_nit}, para que en mi nombre
y representación adelante todos los trámites necesarios ante la DIAN y demás
autoridades aduaneras, para la importación y nacionalización de la siguiente mercancía:

Descripción: {merc}
Documento de transporte: {doc}

Este poder incluye facultades para:
- Presentar y firmar la Declaración de Importación Simplificada (Formulario 510)
- Pagar los tributos aduaneros correspondientes
- Recibir la mercancía una vez otorgado el levante
- Realizar correcciones o modificaciones que sean necesarias

El presente poder se otorga en la ciudad de San Andrés Isla,
a los {fecha}.


_________________________________          _________________________________
Firma del Poderdante                        Firma del Apoderado
{pod}                                       {ag_nombre}
NIT/CC: {pnit}                              NIT: {ag_nit}

Nota: Este documento puede requerir autenticación notarial según el caso.
"""
            path = filedialog.asksaveasfilename(
                defaultextension=".txt", filetypes=[("Texto","*.txt")],
                initialfile=f"Poder_{pnit}_{date.today().isoformat()}.txt")
            if not path: return
            with open(path,"w",encoding="utf-8") as f: f.write(txt)
            import os; os.startfile(path)

        tk.Button(frame, text="📄  Generar Poder", font=("Arial",11,"bold"),
                  bg="#1d4ed8", fg="white", relief="flat", pady=10,
                  cursor="hand2", command=generar).pack(fill="x", pady=16)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # root stays hidden forever; App is its own Toplevel window
    init_db()

    def on_login(user, rol):
        App(root, user, rol)

    def on_splash_ready():
        VentanaLogin(root, on_login)

    SplashScreen(root, on_splash_ready)
    root.mainloop()
