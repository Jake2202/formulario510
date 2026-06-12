import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import date
import os, sys

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

BLUE   = colors.HexColor("#1d4ed8")
LBLUE  = colors.HexColor("#dbeafe")
GRAY   = colors.HexColor("#f1f5f9")
DGRAY  = colors.HexColor("#475569")
BLACK  = colors.HexColor("#0f172a")
WHITE  = colors.white
BORDER = colors.HexColor("#cbd5e1")

# ── Mapping: field key -> Excel cell in PlantillaFormulario510.xlsx ─────────
EXCEL_MAP = {
    "nit":"B6","dv":"D6","razonSocial":"B7","direccion":"B8","telefono":"D8",
    "codSeccional":"D9","codDpto":"B9","codMunicipio":"B10",
    "nitDecl":"B13","dvDecl":"D13","razonDecl":"B14","tipoUsuario":"D14",
    "codUsuario":"B15","numDocDecl":"D15","nombresDecl":"B16",
    "tipoDecl":"B19","numFormAnterior":"D19","manifestoCarga":"B20","fechaLlegada":"D20",
    "codLugarIngreso":"B21","codModo":"D21","docTransporte":"B22","fechaDocTransporte":"D22",
    "codProcedencia":"B23","tasaCambio":"D23",
    "nombreExportador":"B26","formaPago":"D26","codPaisCompra":"B27","codPaisOrigen":"D27",
    "subpartida":"B28","numBultos":"D28","cantidad":"B29","pesoBruto":"D29",
    "pesoNeto":"B30","fob":"D30","fletes":"B31","seguros":"D31",
    "otrosGastos":"B32","ajuste":"D32","descripcion":"B33",
    "arancelPct":"B36","ivaPct":"D36","icPct":"B37",
}

def fmt_cop(n):
    try: return f"${int(round(float(n))):,}".replace(",",".")
    except: return "$0"

# ── PDF generation ───────────────────────────────────────────────────────────
def make_pdf(data, path):
    doc = SimpleDocTemplate(path, pagesize=A4,
                            topMargin=14*mm, bottomMargin=14*mm,
                            leftMargin=14*mm, rightMargin=14*mm)
    story = []

    s_title = ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=15, textColor=BLUE, spaceAfter=2)
    s_sub   = ParagraphStyle("s", fontName="Helvetica",      fontSize=8,  textColor=DGRAY, spaceAfter=8)
    s_sec   = ParagraphStyle("sc",fontName="Helvetica-Bold", fontSize=7,  textColor=BLUE,
                              spaceBefore=10, spaceAfter=4, leading=10)
    s_lbl   = ParagraphStyle("l", fontName="Helvetica",      fontSize=7,  textColor=DGRAY)
    s_val   = ParagraphStyle("v", fontName="Helvetica-Bold", fontSize=9,  textColor=BLACK)
    s_total = ParagraphStyle("to",fontName="Helvetica-Bold", fontSize=13, textColor=BLUE, alignment=TA_RIGHT)
    s_foot  = ParagraphStyle("f", fontName="Helvetica-Oblique", fontSize=7, textColor=DGRAY, alignment=TA_CENTER)

    # Header
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
            while len(row) < 4: row.append(["",""])
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
        ("Cód. usuario", data.get("codUsuario","")),
        ("No. documento", data.get("numDocDecl","")),
        ("Nombres", data.get("nombresDecl","")),
    ])
    section("03 — Transporte", [
        ("Tipo declaración", data.get("tipoDecl","")),
        ("Manifiesto carga", data.get("manifestoCarga","")),
        ("Fecha llegada", data.get("fechaLlegada","")),
        ("Lugar ingreso", data.get("codLugarIngreso","")),
        ("Doc. transporte", data.get("docTransporte","")),
        ("Fecha doc. transp.", data.get("fechaDocTransporte","")),
        ("Modo transporte", data.get("codModo","")),
        ("País procedencia", data.get("codProcedencia","")),
        ("Tasa de cambio", data.get("tasaCambio","")+" COP/USD"),
    ])
    section("04 — Mercancía", [
        ("Proveedor/Exportador", data.get("nombreExportador","")),
        ("País compra", data.get("codPaisCompra","")),
        ("País origen", data.get("codPaisOrigen","")),
        ("Forma de pago", data.get("formaPago","")),
        ("Subpartida arancelaria", data.get("subpartida","")),
        ("No. bultos", data.get("numBultos","")),
        ("Cantidad", data.get("cantidad","")),
        ("Peso bruto (kg)", data.get("pesoBruto","")),
        ("Peso neto (kg)", data.get("pesoNeto","")),
        ("Valor FOB (USD)", "$"+data.get("fob","")),
        ("Valor fletes (USD)", "$"+data.get("fletes","")),
        ("Valor seguros (USD)", "$"+data.get("seguros","")),
        ("Otros gastos (USD)", "$"+data.get("otrosGastos","")),
        ("Valor Aduana CIF", "$"+data.get("valorAduana","")+" USD"),
    ])

    story.append(Paragraph("Descripción de las Mercancías (Cas. 68)", s_sec))
    dt = Table([[Paragraph(data.get("descripcion","—"), s_val)]], colWidths=[182*mm])
    dt.setStyle(TableStyle([("BOX",(0,0),(-1,-1),0.5,BORDER),
                             ("BACKGROUND",(0,0),(-1,-1),WHITE),
                             ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
                             ("LEFTPADDING",(0,0),(-1,-1),8)]))
    story.append(dt); story.append(Spacer(1,6))

    story.append(Paragraph("05 — Liquidación Tributaria", s_sec))
    fob   = float(data.get("fob","0") or 0)
    flt   = float(data.get("fletes","0") or 0)
    seg   = float(data.get("seguros","0") or 0)
    otr   = float(data.get("otrosGastos","0") or 0)
    adj   = float(data.get("ajuste","0") or 0)
    trm   = float(data.get("tasaCambio","4150") or 4150)
    ap    = float(data.get("arancelPct","0") or 0)
    ip    = float(data.get("ivaPct","19") or 19)
    icp   = float(data.get("icPct","0") or 0)
    cif   = fob+flt+seg+otr+adj
    cifC  = cif*trm
    araC  = cifC*(ap/100)
    ivaC  = (cifC+araC)*(ip/100)
    icC   = cifC*(icp/100)
    total = araC+ivaC+icC

    def lr(label, val, bold=False, blue=False):
        ls = ParagraphStyle("lr", fontName="Helvetica-Bold" if bold else "Helvetica",
                             fontSize=9, textColor=BLUE if blue else (BLACK if bold else DGRAY))
        vs = ParagraphStyle("vr", fontName="Helvetica-Bold" if bold else "Helvetica",
                             fontSize=9 if not blue else 13,
                             textColor=BLUE if blue else (BLACK if bold else DGRAY), alignment=TA_RIGHT)
        return [Paragraph(label, ls), Paragraph(val, vs)]

    liq = [
        lr(f"FOB", f"${fob:.2f} USD"),
        lr(f"+ Fletes", f"${flt:.2f} USD"),
        lr(f"+ Seguros", f"${seg:.2f} USD"),
        lr(f"+ Otros gastos", f"${otr:.2f} USD"),
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

    ct = Table([[
        [Paragraph("Casilla 72 — Total Arancel $", s_lbl), Paragraph(fmt_cop(araC), s_total)],
        [Paragraph("Casilla 76 — Total IVA $",     s_lbl), Paragraph(fmt_cop(ivaC), s_total)],
        [Paragraph("Casilla 980 — Pago total $",   s_lbl), Paragraph(fmt_cop(total),s_total)],
    ]], colWidths=[60*mm, 60*mm, 62*mm])
    ct.setStyle(TableStyle([
        ("BOX",(0,0),(-1,-1),1.5,BLUE),("INNERGRID",(0,0),(-1,-1),0.5,BORDER),
        ("BACKGROUND",(0,0),(-1,-1),LBLUE),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),10),
    ]))
    story.append(ct); story.append(Spacer(1,16))

    ft = Table([[
        [Paragraph("Firma del declarante", s_lbl), Spacer(1,20), HRFlowable(width="80%",thickness=0.5,color=BORDER)],
        [Paragraph("Nombre completo",       s_lbl), Spacer(1,20), HRFlowable(width="80%",thickness=0.5,color=BORDER)],
        [Paragraph("C.C. No.",              s_lbl), Spacer(1,20), HRFlowable(width="80%",thickness=0.5,color=BORDER)],
    ]], colWidths=[60*mm, 60*mm, 62*mm])
    ft.setStyle(TableStyle([
        ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(ft)
    story.append(Spacer(1,8))
    story.append(HRFlowable(width="100%", thickness=0.5, color=BORDER))
    story.append(Spacer(1,4))
    story.append(Paragraph(
        f"Pre-diligenciamiento de referencia · No reemplaza declaración oficial ante la DIAN · "
        f"Generado el {date.today().strftime('%d/%m/%Y')}",
        s_foot))
    doc.build(story)


# ── GUI ──────────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Formulario 510 — Declaración de Importación Simplificada")
        self.geometry("1020x720")
        self.minsize(880, 600)
        self.configure(bg="#f1f5f9")
        self.resizable(True, True)
        self.fields = {}
        self._sections = []
        self._nav_btns = []
        self._build_ui()
        self._set_defaults()

    # ─── UI ──────────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top bar
        top = tk.Frame(self, bg="#1d4ed8", height=54)
        top.pack(fill="x"); top.pack_propagate(False)
        tk.Label(top, text="Formulario 510", font=("Arial",16,"bold"),
                 bg="#1d4ed8", fg="white").pack(side="left", padx=20, pady=10)
        tk.Label(top, text="Declaración de Importación Simplificada · DIAN Colombia",
                 font=("Arial",9), bg="#1d4ed8", fg="#bfdbfe").pack(side="left")

        body = tk.Frame(self, bg="#f1f5f9")
        body.pack(fill="both", expand=True)

        # ── Sidebar ──
        sb = tk.Frame(body, bg="#0f1724", width=200)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)

        tk.Label(sb, text="SECCIONES", font=("Arial",7,"bold"),
                 bg="#0f1724", fg="#1e3a5f").pack(pady=(18,6), padx=16, anchor="w")

        for i,(sid,label) in enumerate([
            ("s0","01  Importador"),("s1","02  Declarante"),
            ("s2","03  Transporte"),("s3","04  Mercancía"),("s4","05  Liquidación")]):
            b = tk.Button(sb, text=label, font=("Arial",10), anchor="w",
                          bg="#0f1724", fg="#4a6a8a", relief="flat",
                          activebackground="#1a2535", activeforeground="#e0eaf4",
                          bd=0, padx=14, pady=7, cursor="hand2",
                          command=lambda idx=i: self._jump(idx))
            b.pack(fill="x", padx=8, pady=1)
            self._nav_btns.append(b)

        # Total box
        tf = tk.Frame(sb, bg="#070d15"); tf.pack(fill="x", padx=10, pady=(20,4))
        tk.Label(tf, text="TOTAL A PAGAR", font=("Arial",7,"bold"),
                 bg="#070d15", fg="#1e3a5f").pack(anchor="w", padx=10, pady=(8,0))
        self.lbl_total = tk.Label(tf, text="$0", font=("Arial",20,"bold"),
                                   bg="#070d15", fg="#3b82f6")
        self.lbl_total.pack(anchor="w", padx=10)
        tk.Label(tf, text="COP", font=("Arial",8), bg="#070d15", fg="#2a4060").pack(anchor="w", padx=10, pady=(0,4))
        for lbl, attr in [("Arancel","lbl_ara"),("IVA","lbl_iva"),("Imp. Consumo","lbl_ic")]:
            row = tk.Frame(tf, bg="#070d15"); row.pack(fill="x", padx=10, pady=1)
            tk.Label(row, text=lbl, font=("Arial",8), bg="#070d15", fg="#1e3a5f").pack(side="left")
            w = tk.Label(row, text="$0", font=("Arial",8), bg="#070d15", fg="#4a6a8a"); w.pack(side="right")
            setattr(self, attr, w)
        tk.Frame(tf, bg="#070d15", height=8).pack()

        # Action buttons
        for text, cmd, bg, fg in [
            ("🔄  Actualizar TRM",     self._update_trm, "#7c3aed", "white"),
            ("📥  Cargar desde Excel", self._load_excel, "#0f766e", "white"),
            ("📄  Generar PDF",        self._generate,   "#3b82f6", "white"),
            ("🗑️  Limpiar formulario", self._clear,      "#1e2535", "#64748b"),
        ]:
            tk.Button(sb, text=text, font=("Arial",9,"bold"), bg=bg, fg=fg,
                      relief="flat", bd=0, padx=10, pady=8, cursor="hand2",
                      activebackground=bg, activeforeground=fg,
                      command=cmd).pack(fill="x", padx=10, pady=2)

        # TRM status label
        self.lbl_trm_status = tk.Label(sb, text="", font=("Arial",7),
                                        bg="#0f1724", fg="#4a6a8a", wraplength=170)
        self.lbl_trm_status.pack(padx=10, pady=(0,4))

        # ── Scrollable main area ──
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
        tk.Label(hdr, text=title, font=("Arial",9,"bold"), bg="white", fg="#1d4ed8").pack(side="left")
        c = tk.Frame(fr, bg="white"); c.pack(fill="x", padx=18, pady=(0,14))
        return c

    def _field(self, parent, label, key, row, col, colspan=1, widget="entry", opts=None, width=20):
        fc = tk.Frame(parent, bg="white")
        fc.grid(row=row, column=col, columnspan=colspan, sticky="ew", padx=4, pady=3)
        for c in range(col, col+colspan): parent.columnconfigure(c, weight=1)
        tk.Label(fc, text=label, font=("Arial",7), bg="white", fg="#64748b").pack(anchor="w")
        if widget == "entry":
            w = tk.Entry(fc, font=("Arial",10), relief="flat", bd=0, bg="#f8fafc",
                         width=width, highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x", ipady=4)
            w.bind("<KeyRelease>", lambda e: self._calc())
        elif widget == "combo":
            w = ttk.Combobox(fc, values=[o[1] for o in opts], font=("Arial",10),
                              state="readonly", width=width-2)
            w.set(opts[0][1]); w._opts = opts
            w.bind("<<ComboboxSelected>>", lambda e: self._calc())
            w.pack(fill="x")
        elif widget == "text":
            w = tk.Text(fc, font=("Arial",10), relief="flat", bd=0, bg="#f8fafc",
                        height=3, width=width, highlightbackground="#cbd5e1", highlightthickness=1)
            w.pack(fill="x")
            w.bind("<KeyRelease>", lambda e: self._calc())
        self.fields[key] = w

    def _build_form(self):
        F = self._field

        # 01 Importador
        c = self._card("01 — Importador")
        F(c,"NIT (sin DV)","nit",0,0); F(c,"DV","dv",0,1)
        F(c,"Razón social / Nombres y apellidos","razonSocial",0,2,colspan=2,width=40)
        F(c,"Dirección","direccion",1,0,colspan=2,width=40)
        F(c,"Teléfono","telefono",1,2)
        F(c,"Cód. Seccional","codSeccional",1,3,widget="combo",opts=[
            ("18","18 — San Andrés"),("11","11 — Bogotá"),("08","08 — Barranquilla"),
            ("13","13 — Cartagena"),("76","76 — Cali"),("05","05 — Medellín")])
        F(c,"Cód. Departamento DANE","codDpto",2,0); F(c,"Cód. Ciudad DANE","codMunicipio",2,1)

        # 02 Declarante
        c = self._card("02 — Declarante Autorizado")
        F(c,"NIT Declarante","nitDecl",0,0); F(c,"DV","dvDecl",0,1)
        F(c,"Razón social declarante","razonDecl",0,2,colspan=2,width=40)
        F(c,"Tipo usuario","tipoUsuario",1,0); F(c,"Cód. usuario DIAN","codUsuario",1,1)
        F(c,"No. Documento","numDocDecl",1,2); F(c,"Apellidos y nombres","nombresDecl",1,3)

        # 03 Transporte
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
        F(c,"País procedencia","codProcedencia",2,0,widget="combo",opts=[
            ("US","US — Estados Unidos"),("CN","CN — China"),("DE","DE — Alemania"),
            ("JP","JP — Japón"),("KR","KR — Corea del Sur"),("GB","GB — Reino Unido"),
            ("ES","ES — España"),("MX","MX — México"),("BR","BR — Brasil")])
        F(c,"Tasa de cambio COP/USD","tasaCambio",2,1)

        # 04 Mercancía
        c = self._card("04 — Mercancía")
        F(c,"Nombre exportador / proveedor","nombreExportador",0,0,colspan=2,width=40)
        F(c,"País compra","codPaisCompra",0,2,widget="combo",opts=[
            ("US","US — Estados Unidos"),("CN","CN — China"),("DE","DE — Alemania"),
            ("JP","JP — Japón"),("KR","KR — Corea del Sur"),("GB","GB — Reino Unido"),("ES","ES — España")])
        F(c,"País origen","codPaisOrigen",0,3,widget="combo",opts=[
            ("US","US — Estados Unidos"),("CN","CN — China"),("DE","DE — Alemania"),
            ("JP","JP — Japón"),("KR","KR — Corea del Sur"),("GB","GB — Reino Unido"),("ES","ES — España")])
        F(c,"Forma pago","formaPago",1,0,widget="combo",opts=[
            ("99","99 — Sin pago exterior"),("01","01 — Giro directo"),("02","02 — Carta crédito")])
        F(c,"Subpartida arancelaria (10 dígitos)","subpartida",1,1)
        F(c,"No. bultos","numBultos",1,2); F(c,"Cantidad","cantidad",1,3)
        F(c,"Peso bruto (kg)","pesoBruto",2,0); F(c,"Peso neto (kg)","pesoNeto",2,1)
        F(c,"Valor FOB (USD)","fob",2,2); F(c,"Valor fletes (USD)","fletes",2,3)
        F(c,"Valor seguros (USD)","seguros",3,0); F(c,"Otros gastos (USD)","otrosGastos",3,1)
        F(c,"Ajuste valor (USD)","ajuste",3,2)
        F(c,"Valor Aduana CIF (auto)","valorAduana",3,3)
        F(c,"Descripción — marca, modelo, serial, otros","descripcion",4,0,colspan=4,widget="text",width=80)

        # 05 Liquidación
        c = self._card("05 — Liquidación Tributaria")
        F(c,"% Arancel","arancelPct",0,0); F(c,"% IVA","ivaPct",0,1); F(c,"% Imp. consumo","icPct",0,2)

        lf = tk.Frame(c, bg="#eff6ff", highlightbackground="#bfdbfe", highlightthickness=1)
        lf.grid(row=1, column=0, columnspan=4, sticky="ew", padx=4, pady=6)
        liq_rows = [
            ("FOB","l_fob"),("+ Fletes","l_flt"),("+ Seguros","l_seg"),
            ("+ Otros","l_otr"),("± Ajuste","l_adj"),
            ("Valor Aduana CIF","l_cif"),
            ("Arancel","l_ara"),("IVA","l_iva"),("Imp. Consumo","l_ic"),
        ]
        lf.columnconfigure(0,weight=1)
        for i,(lbl,attr) in enumerate(liq_rows):
            tk.Label(lf,text=lbl,font=("Arial",9),bg="#eff6ff",fg="#475569",anchor="w"
                     ).grid(row=i,column=0,sticky="w",padx=12,pady=2)
            v = tk.Label(lf,text="$0",font=("Arial",9),bg="#eff6ff",fg="#475569",anchor="e")
            v.grid(row=i,column=1,sticky="e",padx=12,pady=2); setattr(self,attr,v)
        tk.Frame(lf,bg="#bfdbfe",height=2).grid(row=len(liq_rows),column=0,columnspan=2,sticky="ew",padx=8,pady=4)
        tk.Label(lf,text="TOTAL LIQUIDADO (Cas. 93)",font=("Arial",10,"bold"),
                 bg="#eff6ff",fg="#1e3a5f",anchor="w").grid(row=len(liq_rows)+1,column=0,sticky="w",padx=12,pady=4)
        self.l_total_big = tk.Label(lf,text="$0 COP",font=("Arial",14,"bold"),bg="#eff6ff",fg="#1d4ed8",anchor="e")
        self.l_total_big.grid(row=len(liq_rows)+1,column=1,sticky="e",padx=12,pady=4)

        cas = tk.Frame(c, bg="white")
        cas.grid(row=2,column=0,columnspan=4,sticky="ew",padx=4,pady=(4,0))
        for i,(cl,ca) in enumerate([("Cas. 72 — Total Arancel $","cas72"),
                                     ("Cas. 76 — Total IVA $","cas76"),
                                     ("Cas. 980 — Pago total $","cas980")]):
            cf = tk.Frame(cas,bg="#dbeafe",highlightbackground="#93c5fd",highlightthickness=1)
            cf.grid(row=0,column=i,sticky="ew",padx=5,pady=4); cas.columnconfigure(i,weight=1)
            tk.Label(cf,text=cl,font=("Arial",7),bg="#dbeafe",fg="#1e40af").pack(anchor="w",padx=8,pady=(6,0))
            w = tk.Label(cf,text="$0",font=("Arial",13,"bold"),bg="#dbeafe",fg="#1d4ed8",anchor="e")
            w.pack(fill="x",padx=8,pady=(0,6)); setattr(self,ca,w)

        tk.Frame(self.inner,bg="#f1f5f9",height=30).pack()

    # ─── Logic ───────────────────────────────────────────────────────────────
    def _set_defaults(self):
        today = date.today().isoformat()
        defaults = {
            "nit":"","dv":"","razonSocial":"","direccion":"","telefono":"",
            "codDpto":"","codMunicipio":"",
            "nitDecl":"","dvDecl":"","razonDecl":"","tipoUsuario":"",
            "codUsuario":"","numDocDecl":"","nombresDecl":"",
            "numFormAnterior":"","manifestoCarga":"",
            "fechaLlegada":today,"docTransporte":"","fechaDocTransporte":today,
            "tasaCambio":"",
            "nombreExportador":"","subpartida":"","numBultos":"","cantidad":"",
            "pesoBruto":"","pesoNeto":"",
            "fob":"","fletes":"","seguros":"","otrosGastos":"","ajuste":"",
            "descripcion":"","arancelPct":"","ivaPct":"19","icPct":"",
        }
        for k,v in defaults.items():
            self._set_field(k,v)
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
            # Try to match by code
            for cod,lbl in w._opts:
                if str(cod).strip() == val.strip() or str(lbl).strip() == val.strip():
                    w.set(lbl); return
            # Try partial
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

        # Update CIF display field
        w=self.fields.get("valorAduana")
        if isinstance(w,tk.Entry):
            w.config(state="normal"); w.delete(0,"end"); w.insert(0,f"{cif:.2f}")
            w.config(state="readonly",readonlybackground="#eef4ff",fg="#1d4ed8")

        def f(n): return f"${int(round(n)):,}".replace(",",".")
        self.l_fob.config(text=f"${fob:.2f} USD")
        self.l_flt.config(text=f"${flt:.2f} USD")
        self.l_seg.config(text=f"${seg:.2f} USD")
        self.l_otr.config(text=f"${otr:.2f} USD")
        self.l_adj.config(text=f"${adj:.2f} USD")
        self.l_cif.config(text=f"${cif:.2f} USD = {f(cifC)} COP", font=("Arial",9,"bold"), fg="#0f172a")
        self.l_ara.config(text=f"{f(araC)} COP ({ap}%)")
        self.l_iva.config(text=f"{f(ivaC)} COP ({ip}%)")
        self.l_ic.config(text=f"{f(icC)} COP ({icp}%)")
        self.l_total_big.config(text=f"{f(total)} COP")
        self.cas72.config(text=f(araC)); self.cas76.config(text=f(ivaC)); self.cas980.config(text=f(total))
        self.lbl_total.config(text=f(total))
        self.lbl_ara.config(text=f(araC)); self.lbl_iva.config(text=f(ivaC)); self.lbl_ic.config(text=f(icC))

    def _jump(self, idx):
        sec = self._sections[idx]
        self.canvas.update_idletasks()
        y = sec.winfo_y()
        total_h = self.inner.winfo_height()
        canvas_h = self.canvas.winfo_height()
        frac = y / max(total_h - canvas_h, 1)
        self.canvas.yview_moveto(max(0, min(1, frac)))
        for i,b in enumerate(self._nav_btns):
            b.config(bg="#0d1e30" if i==idx else "#0f1724",
                     fg="#e0eaf4" if i==idx else "#4a6a8a",
                     font=("Arial",10,"bold") if i==idx else ("Arial",10))

    # ─── TRM automática ──────────────────────────────────────────────────────
    def _update_trm(self):
        import urllib.request, json, threading
        self.lbl_trm_status.config(text="Consultando TRM...", fg="#f59e0b")
        def fetch():
            try:
                today = date.today().strftime("%Y-%m-%d")
                url = (f"https://www.datos.gov.co/resource/mcec-87by.json"
                       f"?vigenciadesde={today}&vigenciahasta={today}")
                req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
                with urllib.request.urlopen(req, timeout=8) as r:
                    data = json.loads(r.read())
                if data and "valor" in data[0]:
                    trm = float(data[0]["valor"])
                    self.after(0, lambda: self._apply_trm(trm, today))
                else:
                    # Fallback: API Banco de la República
                    url2 = (f"https://www.banrep.gov.co/es/trm-api?"
                            f"startDate={today}&endDate={today}")
                    req2 = urllib.request.Request(url2, headers={"User-Agent": "Mozilla/5.0"})
                    with urllib.request.urlopen(req2, timeout=8) as r2:
                        data2 = json.loads(r2.read())
                    trm = float(data2[0]["value"])
                    self.after(0, lambda: self._apply_trm(trm, today))
            except Exception as e:
                self.after(0, lambda: self.lbl_trm_status.config(
                    text=f"Sin conexión. Ingrese TRM manualmente.", fg="#ef4444"))
        threading.Thread(target=fetch, daemon=True).start()

    def _apply_trm(self, trm, fecha):
        self._set_field("tasaCambio", f"{trm:.2f}")
        self._calc()
        self.lbl_trm_status.config(
            text=f"TRM {fecha}\n${trm:,.2f} COP/USD", fg="#22c55e")

    # ─── Load from Excel ─────────────────────────────────────────────────────
    def _load_excel(self):
        if not HAS_XLSX:
            messagebox.showerror("Error","openpyxl no está instalado.")
            return
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel (PlantillaFormulario510.xlsx)",
            filetypes=[("Excel","*.xlsx *.xls")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["Formulario510"]
            loaded = 0
            for key, cell_addr in EXCEL_MAP.items():
                val = ws[cell_addr].value
                if val is not None:
                    self._set_field(key, str(val).strip())
                    loaded += 1
            self._calc()
            messagebox.showinfo("Listo",
                f"✅ Datos cargados desde Excel exitosamente.\n{loaded} campos importados.\n\nRevise el formulario y genere el PDF.")
        except KeyError:
            messagebox.showerror("Error",
                "No se encontró la hoja 'Formulario510'.\nAsegúrese de usar la plantilla oficial PlantillaFormulario510.xlsx")
        except Exception as e:
            messagebox.showerror("Error al cargar Excel", str(e))

    # ─── Generate PDF ─────────────────────────────────────────────────────────
    def _generate(self):
        self._calc()
        path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")],
            initialfile=f"Formulario510_{date.today().isoformat()}.pdf",
            title="Guardar declaración como PDF")
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
            messagebox.showinfo("PDF Generado",
                f"✅ Declaración generada correctamente:\n\n{path}\n\nAbra el archivo con cualquier visor de PDF (Adobe, Edge, Chrome).")
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e))

    # ─── Clear ────────────────────────────────────────────────────────────────
    def _clear(self):
        if messagebox.askyesno("Confirmar", "¿Desea limpiar todos los campos del formulario?"):
            self._set_defaults()


if __name__ == "__main__":
    app = App()
    app.mainloop()
