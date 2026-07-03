# ── BACKEND (sin cambios) ──────────────────────────────────────────────────────
import sys
exec(open("/tmp/backend.py").read())
# ── FIN BACKEND ───────────────────────────────────────────────────────────────

import customtkinter as ctk
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

# ── Tema global ───────────────────────────────────────────────────────────────
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

BLUE    = "#3b82f6"
BLUE_DK = "#1d4ed8"
BG      = "#0f1117"
BG2     = "#161b22"
BG3     = "#1c2128"
BORDER  = "#30363d"
TEXT    = "#e6edf3"
TEXT2   = "#8b949e"
GREEN   = "#3fb950"
RED     = "#f85149"
YELLOW  = "#d29922"
PURPLE  = "#8b5cf6"

FONT_TITLE  = ("Inter", 22, "bold")
FONT_H2     = ("Inter", 13, "bold")
FONT_LABEL  = ("Inter", 10)
FONT_SMALL  = ("Inter", 9)
FONT_MONO   = ("JetBrains Mono", 10)

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

PAISES = [
    ("AF","Afganistán"),("AL","Albania"),("DE","Alemania"),("AD","Andorra"),
    ("AO","Angola"),("AG","Antigua y Barbuda"),("SA","Arabia Saudita"),("DZ","Argelia"),
    ("AR","Argentina"),("AM","Armenia"),("AU","Australia"),("AT","Austria"),
    ("AZ","Azerbaiyán"),("BS","Bahamas"),("BD","Bangladesh"),("BE","Bélgica"),
    ("BZ","Belice"),("BO","Bolivia"),("BR","Brasil"),("BN","Brunéi"),
    ("CA","Canadá"),("QA","Catar"),("CL","Chile"),("CN","China"),
    ("CO","Colombia"),("KR","Corea del Sur"),("CR","Costa Rica"),("HR","Croacia"),
    ("CU","Cuba"),("DK","Dinamarca"),("EC","Ecuador"),("EG","Egipto"),
    ("SV","El Salvador"),("AE","Emiratos Árabes"),("ES","España"),("US","Estados Unidos"),
    ("FR","Francia"),("GH","Ghana"),("GR","Grecia"),("GT","Guatemala"),
    ("HN","Honduras"),("HU","Hungría"),("IN","India"),("ID","Indonesia"),
    ("IE","Irlanda"),("IL","Israel"),("IT","Italia"),("JM","Jamaica"),
    ("JP","Japón"),("KE","Kenia"),("MY","Malasia"),("MX","México"),
    ("NI","Nicaragua"),("NG","Nigeria"),("NO","Noruega"),("NZ","Nueva Zelanda"),
    ("PA","Panamá"),("PY","Paraguay"),("PE","Perú"),("PL","Polonia"),
    ("PT","Portugal"),("GB","Reino Unido"),("DO","República Dominicana"),
    ("RU","Rusia"),("SG","Singapur"),("ZA","Sudáfrica"),("SE","Suecia"),
    ("CH","Suiza"),("TH","Tailandia"),("TZ","Tanzania"),("TR","Turquía"),
    ("UA","Ucrania"),("UY","Uruguay"),("VE","Venezuela"),("VN","Vietnam"),
]

def pais_opts():
    return [f"{cod} — {nombre}" for cod, nombre in PAISES]

def pais_cod(display):
    return display.split(" — ")[0] if display else ""

# ═══════════════════════════════════════════════════════════════════════════════
# HELPERS UI
# ═══════════════════════════════════════════════════════════════════════════════
def make_entry(parent, placeholder="", width=200):
    e = ctk.CTkEntry(parent, placeholder_text=placeholder,
                     font=FONT_LABEL, width=width,
                     fg_color=BG3, border_color=BORDER,
                     text_color=TEXT, placeholder_text_color=TEXT2)
    return e

def make_label(parent, text, size=10, bold=False, color=None):
    f = ("Inter", size, "bold") if bold else ("Inter", size)
    return ctk.CTkLabel(parent, text=text, font=f,
                        text_color=color or TEXT2)

def make_button(parent, text, cmd, color=BLUE, width=160, height=34):
    return ctk.CTkButton(parent, text=text, command=cmd,
                         font=("Inter",11,"bold"),
                         fg_color=color, hover_color=BLUE_DK,
                         width=width, height=height,
                         corner_radius=6)

def make_combo(parent, values, width=200):
    return ctk.CTkComboBox(parent, values=values,
                           font=FONT_LABEL, width=width,
                           fg_color=BG3, border_color=BORDER,
                           text_color=TEXT, button_color=BLUE,
                           dropdown_fg_color=BG2,
                           dropdown_text_color=TEXT)

def section_card(parent):
    return ctk.CTkFrame(parent, fg_color=BG2,
                        corner_radius=10,
                        border_width=1, border_color=BORDER)

def section_title(parent, num, text):
    f = ctk.CTkFrame(parent, fg_color="transparent")
    f.pack(fill="x", padx=20, pady=(18,8))
    ctk.CTkLabel(f, text=num, font=("Inter",9,"bold"),
                 text_color=BLUE,
                 width=28, height=22,
                 fg_color=BG3, corner_radius=4).pack(side="left", padx=(0,8))
    ctk.CTkLabel(f, text=text, font=("Inter",13,"bold"),
                 text_color=TEXT).pack(side="left")

def field_group(parent, label, widget_fn, row, col, colspan=1, padx=(4,4), pady=(4,4)):
    fr = ctk.CTkFrame(parent, fg_color="transparent")
    fr.grid(row=row, column=col, columnspan=colspan,
            sticky="ew", padx=padx, pady=pady)
    parent.grid_columnconfigure(col, weight=1)
    if colspan > 1:
        for c in range(col, col+colspan):
            parent.grid_columnconfigure(c, weight=1)
    ctk.CTkLabel(fr, text=label, font=("Inter",9),
                 text_color=TEXT2).pack(anchor="w", pady=(0,2))
    w = widget_fn(fr)
    w.pack(fill="x")
    return w

# ═══════════════════════════════════════════════════════════════════════════════
# SPLASH SCREEN
# ═══════════════════════════════════════════════════════════════════════════════
class SplashScreen(ctk.CTkToplevel):
    def __init__(self, parent, on_ready):
        super().__init__(parent)
        self.on_ready = on_ready
        self.overrideredirect(True)
        w, h = 460, 280
        sw = self.winfo_screenwidth(); sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")
        self.configure(fg_color=BG)
        self._build()
        self.after(2800, self._cerrar)

    def _build(self):
        # Logo badge
        logo = ctk.CTkFrame(self, width=72, height=72, corner_radius=16,
                             fg_color=BLUE_DK)
        logo.place(x=194, y=36)
        logo.pack_propagate(False)
        ctk.CTkLabel(logo, text="510", font=("Inter",22,"bold"),
                     text_color="white").place(relx=0.5, rely=0.5, anchor="center")

        # Texts
        agencia = db_fetch("SELECT valor FROM config WHERE clave='agencia_nombre'")
        nombre = agencia[0][0] if agencia else "Mi Agencia de Aduanas"

        ctk.CTkLabel(self, text="DeclaraFácil", font=("Inter",26,"bold"),
                     text_color=TEXT).place(x=230, y=124, anchor="center")
        ctk.CTkLabel(self, text=nombre, font=("Inter",11),
                     text_color=BLUE).place(x=230, y=152, anchor="center")
        ctk.CTkLabel(self, text="Sistema de Gestión Aduanera · DIAN Colombia",
                     font=("Inter",8), text_color=TEXT2).place(x=230, y=172, anchor="center")

        # Progress bar
        self.bar = ctk.CTkProgressBar(self, width=360, height=4,
                                       fg_color=BG3, progress_color=BLUE,
                                       corner_radius=2)
        self.bar.place(x=50, y=220)
        self.bar.set(0)

        self.lbl_status = ctk.CTkLabel(self, text="Iniciando...",
                                        font=("Inter",8), text_color=TEXT2)
        self.lbl_status.place(x=230, y=238, anchor="center")

        ctk.CTkLabel(self, text="v2.0", font=("Inter",7),
                     text_color=BORDER).place(x=448, y=270, anchor="se")
        self._animar(0)

    def _animar(self, step):
        msgs = ["Cargando base de datos...","Verificando licencia...","Preparando formularios...","Listo ✓"]
        if step <= 100:
            self.bar.set(step/100)
            self.lbl_status.configure(text=msgs[min(step//30, 3)])
            self.after(22, lambda: self._animar(step+2))

    def _cerrar(self):
        self.destroy()
        self.on_ready()


# ═══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ═══════════════════════════════════════════════════════════════════════════════
class VentanaLogin(ctk.CTkToplevel):
    def __init__(self, parent, on_success):
        super().__init__(parent)
        self.on_success = on_success
        self.title("DeclaraFácil 510 — Iniciar sesión")
        self.geometry("400x420")
        self.resizable(False, False)
        self.configure(fg_color=BG)
        self.protocol("WM_DELETE_WINDOW", lambda: sys.exit(0))
        self._build()

    def _build(self):
        # Card central
        card = ctk.CTkFrame(self, fg_color=BG2, corner_radius=12,
                             border_width=1, border_color=BORDER)
        card.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.88)

        # Badge logo
        badge = ctk.CTkFrame(card, width=56, height=56, corner_radius=12,
                              fg_color=BLUE_DK)
        badge.pack(pady=(28,0))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text="510", font=("Inter",16,"bold"),
                     text_color="white").place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(card, text="DeclaraFácil 510",
                     font=("Inter",16,"bold"), text_color=TEXT).pack(pady=(10,2))

        agencia = db_fetch("SELECT valor FROM config WHERE clave='agencia_nombre'")
        nombre = agencia[0][0] if agencia else ""
        ctk.CTkLabel(card, text=nombre, font=("Inter",9),
                     text_color=TEXT2).pack(pady=(0,20))

        # Fields
        fr = ctk.CTkFrame(card, fg_color="transparent")
        fr.pack(fill="x", padx=28)

        ctk.CTkLabel(fr, text="Usuario", font=("Inter",10),
                     text_color=TEXT2).pack(anchor="w")
        self.ent_user = ctk.CTkEntry(fr, placeholder_text="admin",
                                      font=FONT_LABEL, height=38,
                                      fg_color=BG3, border_color=BORDER, text_color=TEXT)
        self.ent_user.pack(fill="x", pady=(2,12))
        self.ent_user.insert(0,"admin")

        ctk.CTkLabel(fr, text="Contraseña", font=("Inter",10),
                     text_color=TEXT2).pack(anchor="w")
        self.ent_pwd = ctk.CTkEntry(fr, placeholder_text="••••••••",
                                     font=FONT_LABEL, height=38, show="•",
                                     fg_color=BG3, border_color=BORDER, text_color=TEXT)
        self.ent_pwd.pack(fill="x", pady=(2,0))
        self.ent_pwd.bind("<Return>", lambda e: self._login())

        self.lbl_err = ctk.CTkLabel(fr, text="", font=("Inter",9),
                                     text_color=RED)
        self.lbl_err.pack(pady=(6,0))

        ctk.CTkButton(fr, text="Iniciar sesión", command=self._login,
                      font=("Inter",12,"bold"), height=40,
                      fg_color=BLUE, hover_color=BLUE_DK,
                      corner_radius=8).pack(fill="x", pady=(8,24))

    def _login(self):
        user = self.ent_user.get().strip()
        pwd  = hashlib.sha256(self.ent_pwd.get().encode()).hexdigest()
        rows = db_fetch("SELECT id,rol FROM usuarios WHERE username=? AND password_hash=? AND activo=1", (user,pwd))
        if rows:
            self.destroy()
            self.on_success(user, rows[0][1])
        else:
            self.lbl_err.configure(text="Usuario o contraseña incorrectos")
            self.ent_pwd.delete(0,"end")


# ═══════════════════════════════════════════════════════════════════════════════
# APP PRINCIPAL
# ═══════════════════════════════════════════════════════════════════════════════
class App(ctk.CTkToplevel):
    def __init__(self, master, user="admin", rol="admin"):
        super().__init__(master)
        self.master = master
        self.title(f"DeclaraFácil 510")
        self.geometry("1280x820")
        self.minsize(1000, 700)
        self.configure(fg_color=BG)
        self.user = user; self.rol = rol
        self.fields = {}
        self._sections = []
        self._nav_btns = []
        self._cliente_id = None
        self._decl_id = None
        self._total_cop = 0
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._build_ui()
        self._set_defaults()
        self._auto_backup()

    def _on_close(self):
        self.master.destroy(); sys.exit(0)

    def _auto_backup(self):
        try:
            db = get_db_path()
            bak = db.replace(".db", f"_bak_{date.today().isoformat()}.db")
            if not os.path.exists(bak):
                shutil.copy2(db, bak)
        except: pass

    def _build_ui(self):
        # ── Top bar ──
        top = ctk.CTkFrame(self, height=56, fg_color=BG2,
                            corner_radius=0, border_width=0)
        top.pack(fill="x"); top.pack_propagate(False)

        # Logo
        logo_fr = ctk.CTkFrame(top, width=40, height=32, corner_radius=6,
                                 fg_color=BLUE_DK)
        logo_fr.pack(side="left", padx=(16,8), pady=12)
        logo_fr.pack_propagate(False)
        ctk.CTkLabel(logo_fr, text="510", font=("Inter",11,"bold"),
                     text_color="white").place(relx=0.5,rely=0.5,anchor="center")

        ctk.CTkLabel(top, text="DeclaraFácil", font=("Inter",15,"bold"),
                     text_color=TEXT).pack(side="left")
        ctk.CTkLabel(top, text="Declaración de Importación Simplificada · DIAN Colombia",
                     font=("Inter",9), text_color=TEXT2).pack(side="left", padx=(8,0))

        # Right side of topbar
        self.lbl_cliente = ctk.CTkLabel(top, text="Sin cliente",
                                         font=("Inter",9), text_color=TEXT2,
                                         fg_color=BG3, corner_radius=6,
                                         padx=12, pady=4)
        self.lbl_cliente.pack(side="right", padx=16, pady=14)
        ctk.CTkLabel(top, text=f"👤 {self.user}",
                     font=("Inter",9), text_color=TEXT2).pack(side="right", padx=8)

        # ── Body ──
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)

        # ── Sidebar ──
        sb = ctk.CTkFrame(body, width=200, fg_color=BG2,
                           corner_radius=0, border_width=0)
        sb.pack(side="left", fill="y"); sb.pack_propagate(False)

        # Section nav
        ctk.CTkLabel(sb, text="SECCIONES", font=("Inter",9,"bold"),
                     text_color=TEXT2).pack(anchor="w", padx=16, pady=(20,6))

        for i, (num, label) in enumerate([
            ("01","Importador"),("02","Declarante"),
            ("03","Transporte"),("04","Mercancía"),("05","Liquidación")]):
            btn = ctk.CTkButton(sb, text=f"  {num}  {label}",
                                font=("Inter",11), anchor="w",
                                fg_color="transparent", hover_color=BG3,
                                text_color=TEXT2, corner_radius=6,
                                height=34, width=180,
                                command=lambda idx=i: self._jump(idx))
            btn.pack(padx=8, pady=1)
            self._nav_btns.append(btn)

        # Divider
        ctk.CTkFrame(sb, height=1, fg_color=BORDER).pack(fill="x", padx=12, pady=12)

        # Total box
        total_card = ctk.CTkFrame(sb, fg_color=BG3, corner_radius=10)
        total_card.pack(fill="x", padx=10, pady=(0,8))
        ctk.CTkLabel(total_card, text="TOTAL A PAGAR",
                     font=("Inter",8,"bold"), text_color=TEXT2).pack(anchor="w", padx=12, pady=(10,0))
        self.lbl_total = ctk.CTkLabel(total_card, text="$0",
                                       font=("Inter",22,"bold"), text_color=BLUE)
        self.lbl_total.pack(anchor="w", padx=12)
        ctk.CTkLabel(total_card, text="COP", font=("Inter",8),
                     text_color=TEXT2).pack(anchor="w", padx=12, pady=(0,2))
        for lbl, attr in [("Arancel","lbl_ara"),("IVA","lbl_iva"),("Imp. Consumo","lbl_ic")]:
            r = ctk.CTkFrame(total_card, fg_color="transparent")
            r.pack(fill="x", padx=12, pady=1)
            ctk.CTkLabel(r, text=lbl, font=("Inter",8), text_color=TEXT2).pack(side="left")
            w = ctk.CTkLabel(r, text="$0", font=("Inter",8), text_color=TEXT2)
            w.pack(side="right"); setattr(self, attr, w)
        self.lbl_trm = ctk.CTkLabel(total_card, text="", font=("Inter",7),
                                     text_color=GREEN)
        self.lbl_trm.pack(pady=(4,10))

        # Divider
        ctk.CTkFrame(sb, height=1, fg_color=BORDER).pack(fill="x", padx=12, pady=(0,8))

        # Action buttons
        ctk.CTkLabel(sb, text="ACCIONES", font=("Inter",9,"bold"),
                     text_color=TEXT2).pack(anchor="w", padx=16, pady=(0,6))

        sb_scroll = ctk.CTkScrollableFrame(sb, fg_color="transparent",
                                            scrollbar_button_color=BORDER,
                                            height=280)
        sb_scroll.pack(fill="x", padx=8)

        actions = [
            ("👥  Clientes",         self._abrir_clientes,  BG3),
            ("📋  Historial",         self._abrir_historial, BG3),
            ("📊  Estadísticas",      self._abrir_stats,     BG3),
            ("⏰  Plazos",            self._abrir_plazos,    BG3),
            ("✅  Checklist docs",    self._abrir_checklist, BG3),
            ("🧾  Recibo de pago",    self._abrir_rop,       BG3),
            ("📤  Generar EDI/XML",   self._abrir_edi,       BG3),
            ("🔍  Consultar levante", self._abrir_consulta,  BG3),
            ("🚨  Entrega urgente",   self._abrir_urgente,   BG3),
            ("⚖️  Multas",            self._abrir_multas,    BG3),
            ("📝  Poder",             self._abrir_poder,     BG3),
            ("🌐  Ir al SYGA",        self._abrir_syga,      BG3),
            ("🔄  Actualizar TRM",    self._update_trm,      PURPLE),
            ("📥  Cargar Excel",      self._load_excel,      "#0f4c35"),
            ("📋  Subpartidas",       self._subpartidas,     "#7c2d12"),
            ("💾  Guardar",           self._guardar_decl,    "#0c4a6e"),
            ("📄  Generar PDF",       self._generate,        BLUE_DK),
            ("⚙️  Configuración",     self._abrir_config,    BG3),
            ("🗑️  Limpiar",           self._clear,           BG3),
        ]
        for text, cmd, color in actions:
            ctk.CTkButton(sb_scroll, text=text, command=cmd,
                          font=("Inter",10), anchor="w",
                          fg_color=color, hover_color=BLUE_DK,
                          text_color=TEXT, corner_radius=6,
                          height=30, width=180).pack(fill="x", pady=1)

        # ── Main scrollable area ──
        main_fr = ctk.CTkFrame(body, fg_color=BG, corner_radius=0)
        main_fr.pack(side="left", fill="both", expand=True)

        self.scroll_main = ctk.CTkScrollableFrame(main_fr, fg_color=BG,
                                                   scrollbar_button_color=BORDER)
        self.scroll_main.pack(fill="both", expand=True, padx=0, pady=0)

        self._build_form()

    # ── Form builder ──────────────────────────────────────────────────────────
    def _card(self, num, title):
        outer = ctk.CTkFrame(self.scroll_main, fg_color="transparent")
        outer.pack(fill="x", padx=24, pady=(12,0))
        self._sections.append(outer)

        card = ctk.CTkFrame(outer, fg_color=BG2, corner_radius=10,
                             border_width=1, border_color=BORDER)
        card.pack(fill="x")

        # Title bar
        title_bar = ctk.CTkFrame(card, fg_color="transparent", height=44)
        title_bar.pack(fill="x", padx=20, pady=(14,0))
        title_bar.pack_propagate(False)

        badge = ctk.CTkFrame(title_bar, width=28, height=20, corner_radius=4,
                              fg_color=BLUE_DK)
        badge.pack(side="left", padx=(0,8))
        badge.pack_propagate(False)
        ctk.CTkLabel(badge, text=num, font=("Inter",9,"bold"),
                     text_color="white").place(relx=0.5,rely=0.5,anchor="center")
        ctk.CTkLabel(title_bar, text=title, font=("Inter",13,"bold"),
                     text_color=TEXT).pack(side="left")

        content = ctk.CTkFrame(card, fg_color="transparent")
        content.pack(fill="x", padx=20, pady=(8,18))
        return content

    def _field(self, parent, label, key, row, col, colspan=1,
               widget="entry", opts=None, width=None):
        fr = ctk.CTkFrame(parent, fg_color="transparent")
        fr.grid(row=row, column=col, columnspan=colspan,
                sticky="ew", padx=6, pady=4)
        for c in range(col, col+colspan):
            parent.grid_columnconfigure(c, weight=1)

        ctk.CTkLabel(fr, text=label, font=("Inter",9),
                     text_color=TEXT2).pack(anchor="w", pady=(0,3))

        if widget == "entry":
            w = ctk.CTkEntry(fr, font=FONT_LABEL, height=36,
                             fg_color=BG3, border_color=BORDER,
                             text_color=TEXT, placeholder_text_color=TEXT2)
            w.pack(fill="x")
            w.bind("<KeyRelease>", lambda e: self._calc())
            if key == "nit":
                w.bind("<FocusOut>", self._auto_dv)
                w.bind("<KeyRelease>", lambda e: (self._calc(), self._auto_dv()))
        elif widget == "combo":
            values = [o[1] for o in opts]
            w = ctk.CTkComboBox(fr, values=values, font=FONT_LABEL,
                                 width=width or 200,
                                 fg_color=BG3, border_color=BORDER,
                                 text_color=TEXT, button_color=BLUE,
                                 dropdown_fg_color=BG2,
                                 dropdown_text_color=TEXT,
                                 command=lambda v: self._calc())
            w.pack(fill="x")
            w.set(values[0])
            w._opts = opts
        elif widget == "text":
            w = ctk.CTkTextbox(fr, font=FONT_LABEL, height=72,
                                fg_color=BG3, border_color=BORDER,
                                text_color=TEXT, border_width=1)
            w.pack(fill="x")
            w.bind("<KeyRelease>", lambda e: self._calc())

        self.fields[key] = w

    def _build_form(self):
        F = self._field

        # 01 Importador
        c = self._card("01","Importador")
        for i in range(4): c.grid_columnconfigure(i, weight=1)
        F(c,"NIT (sin DV)","nit",0,0)
        F(c,"DV","dv",0,1)
        F(c,"Razón social / Nombres y apellidos","razonSocial",0,2,colspan=2)
        F(c,"Dirección","direccion",1,0,colspan=2)
        F(c,"Teléfono","telefono",1,2)
        F(c,"Cód. Seccional","codSeccional",1,3,widget="combo",opts=[
            ("18","18 — San Andrés"),("11","11 — Bogotá"),("08","08 — Barranquilla"),
            ("13","13 — Cartagena"),("76","76 — Cali"),("05","05 — Medellín")])
        F(c,"Cód. Departamento DANE","codDpto",2,0)
        F(c,"Cód. Ciudad DANE","codMunicipio",2,1)

        # 02 Declarante
        c = self._card("02","Declarante Autorizado")
        for i in range(4): c.grid_columnconfigure(i, weight=1)
        F(c,"NIT Declarante","nitDecl",0,0); F(c,"DV","dvDecl",0,1)
        F(c,"Razón social declarante","razonDecl",0,2,colspan=2)
        F(c,"Tipo usuario","tipoUsuario",1,0)
        F(c,"Cód. usuario DIAN","codUsuario",1,1)
        F(c,"No. Documento","numDocDecl",1,2)
        F(c,"Apellidos y nombres","nombresDecl",1,3)

        # 03 Transporte
        c = self._card("03","Manifiesto y Transporte")
        for i in range(4): c.grid_columnconfigure(i, weight=1)
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
        F(c,"País procedencia","codProcedencia",2,0,widget="combo",
          opts=[(cod, f"{cod} — {nom}") for cod,nom in PAISES])
        F(c,"Tasa de cambio COP/USD","tasaCambio",2,1)

        # 04 Mercancía
        c = self._card("04","Mercancía")
        for i in range(4): c.grid_columnconfigure(i, weight=1)
        F(c,"Nombre exportador / proveedor","nombreExportador",0,0,colspan=2)
        F(c,"País compra","codPaisCompra",0,2,widget="combo",
          opts=[(cod, f"{cod} — {nom}") for cod,nom in PAISES])
        F(c,"País origen","codPaisOrigen",0,3,widget="combo",
          opts=[(cod, f"{cod} — {nom}") for cod,nom in PAISES])
        F(c,"Forma pago","formaPago",1,0,widget="combo",opts=[
            ("99","99 — Sin pago exterior"),("01","01 — Giro directo"),("02","02 — Carta crédito")])
        F(c,"Subpartida arancelaria (10 dígitos)","subpartida",1,1)
        F(c,"No. bultos","numBultos",1,2); F(c,"Cantidad","cantidad",1,3)
        F(c,"Peso bruto (kg)","pesoBruto",2,0); F(c,"Peso neto (kg)","pesoNeto",2,1)
        F(c,"Valor FOB (USD)","fob",2,2); F(c,"Valor fletes (USD)","fletes",2,3)
        F(c,"Valor seguros (USD)","seguros",3,0); F(c,"Otros gastos (USD)","otrosGastos",3,1)
        F(c,"Ajuste valor (USD)","ajuste",3,2)
        F(c,"Valor Aduana CIF (auto)","valorAduana",3,3)
        F(c,"Descripción — marca, modelo, serial","descripcion",4,0,colspan=4,widget="text")

        # 05 Liquidación
        c = self._card("05","Liquidación Tributaria")
        for i in range(4): c.grid_columnconfigure(i, weight=1)
        F(c,"% Arancel","arancelPct",0,0)
        F(c,"% IVA","ivaPct",0,1)
        F(c,"% Imp. consumo","icPct",0,2)

        # Liquidación result card
        liq_outer = ctk.CTkFrame(c, fg_color="transparent")
        liq_outer.grid(row=1, column=0, columnspan=4, sticky="ew", padx=6, pady=(8,0))

        liq = ctk.CTkFrame(liq_outer, fg_color=BG3, corner_radius=8,
                            border_width=1, border_color=BORDER)
        liq.pack(fill="x")

        liq_rows_cfg = [
            ("FOB","l_fob"),("+ Fletes","l_flt"),("+ Seguros","l_seg"),
            ("+ Otros","l_otr"),("± Ajuste","l_adj"),
            ("Valor Aduana CIF","l_cif"),
            ("Arancel","l_ara"),("IVA","l_iva"),("Imp. Consumo","l_ic"),
        ]
        for i,(lbl,attr) in enumerate(liq_rows_cfg):
            r = ctk.CTkFrame(liq, fg_color="transparent")
            r.pack(fill="x", padx=16, pady=2)
            ctk.CTkLabel(r, text=lbl, font=("Inter",10),
                         text_color=TEXT2).pack(side="left")
            v = ctk.CTkLabel(r, text="$0.00 USD" if i < 5 else "$0 COP",
                              font=("Inter",10), text_color=TEXT2)
            v.pack(side="right"); setattr(self, attr, v)

        # Divider
        ctk.CTkFrame(liq, height=1, fg_color=BORDER).pack(fill="x", padx=12, pady=4)

        # Total row
        total_row = ctk.CTkFrame(liq, fg_color="transparent")
        total_row.pack(fill="x", padx=16, pady=(0,16))
        ctk.CTkLabel(total_row, text="TOTAL LIQUIDADO (Cas. 93)",
                     font=("Inter",12,"bold"), text_color=TEXT).pack(side="left")
        self.l_total_big = ctk.CTkLabel(total_row, text="$0 COP",
                                         font=("Inter",16,"bold"), text_color=BLUE)
        self.l_total_big.pack(side="right")

        # Casillas oficiales
        cas_fr = ctk.CTkFrame(c, fg_color="transparent")
        cas_fr.grid(row=2, column=0, columnspan=4, sticky="ew", padx=6, pady=(8,0))
        for i,(cl,ca) in enumerate([
            ("Cas. 72 — Total Arancel","cas72"),
            ("Cas. 76 — Total IVA","cas76"),
            ("Cas. 980 — Pago total","cas980")]):
            cf = ctk.CTkFrame(cas_fr, fg_color=BG3, corner_radius=8,
                               border_width=1, border_color=BLUE_DK)
            cf.grid(row=0, column=i, sticky="ew", padx=4)
            cas_fr.grid_columnconfigure(i, weight=1)
            ctk.CTkLabel(cf, text=cl, font=("Inter",8),
                         text_color=TEXT2).pack(anchor="w", padx=12, pady=(8,0))
            w = ctk.CTkLabel(cf, text="$0", font=("Inter",16,"bold"), text_color=BLUE)
            w.pack(anchor="e", padx=12, pady=(0,8)); setattr(self, ca, w)

        # Bottom padding
        ctk.CTkFrame(self.scroll_main, fg_color="transparent", height=40).pack()

    # ── Field get/set ──────────────────────────────────────────────────────────
    def _get_field(self, key):
        w = self.fields.get(key)
        if w is None: return ""
        if isinstance(w, ctk.CTkTextbox): return w.get("1.0","end-1c")
        if isinstance(w, ctk.CTkComboBox):
            sel = w.get()
            opts = getattr(w, "_opts", None)
            if opts:
                for cod,lbl in opts:
                    if lbl == sel: return cod
            return sel.split("—")[0].strip() if "—" in sel else sel
        return w.get()

    def _set_field(self, key, value):
        w = self.fields.get(key)
        if w is None: return
        val = str(value) if value is not None else ""
        if isinstance(w, ctk.CTkTextbox):
            w.delete("1.0","end"); w.insert("1.0",val)
        elif isinstance(w, ctk.CTkComboBox):
            opts = getattr(w, "_opts", None)
            if opts:
                for cod,lbl in opts:
                    if str(cod).strip() == val.strip() or lbl == val:
                        w.set(lbl); return
                for cod,lbl in opts:
                    if val.strip().upper() in lbl.upper():
                        w.set(lbl); return
        elif isinstance(w, ctk.CTkEntry):
            state = w.cget("state")
            if state == "disabled": w.configure(state="normal")
            w.delete(0,"end"); w.insert(0,val)
            if state == "disabled": w.configure(state="disabled")

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
        cifC=cif*trm; araC=cifC*(ap/100)
        ivaC=(cifC+araC)*(ip/100); icC=cifC*(icp/100)
        total=araC+ivaC+icC

        # Update CIF display field
        w=self.fields.get("valorAduana")
        if isinstance(w, ctk.CTkEntry):
            w.configure(state="normal"); w.delete(0,"end")
            w.insert(0,f"{cif:.2f}"); w.configure(state="disabled",
            text_color=BLUE)

        def f(n): return f"${int(round(n)):,}".replace(",",".")
        self.l_fob.configure(text=f"${fob:.2f} USD")
        self.l_flt.configure(text=f"${flt:.2f} USD")
        self.l_seg.configure(text=f"${seg:.2f} USD")
        self.l_otr.configure(text=f"${otr:.2f} USD")
        self.l_adj.configure(text=f"${adj:.2f} USD")
        self.l_cif.configure(text=f"${cif:.2f} USD = {f(cifC)} COP",
                              text_color=TEXT, font=("Inter",10,"bold"))
        self.l_ara.configure(text=f"{f(araC)} COP ({ap}%)")
        self.l_iva.configure(text=f"{f(ivaC)} COP ({ip}%)")
        self.l_ic.configure(text=f"{f(icC)} COP ({icp}%)")
        self.l_total_big.configure(text=f"{f(total)} COP")
        self.cas72.configure(text=f(araC))
        self.cas76.configure(text=f(ivaC))
        self.cas980.configure(text=f(total))
        self.lbl_total.configure(text=f(total))
        self.lbl_ara.configure(text=f(araC))
        self.lbl_iva.configure(text=f(ivaC))
        self.lbl_ic.configure(text=f(icC))
        self._total_cop = total

    def _jump(self, idx):
        if idx >= len(self._sections): return
        sec = self._sections[idx]
        self.scroll_main.update_idletasks()
        # Scroll to section
        canvas = self.scroll_main._parent_canvas
        y = sec.winfo_y()
        total_h = self.scroll_main._parent_frame.winfo_height()
        canvas_h = canvas.winfo_height()
        frac = y / max(total_h - canvas_h, 1)
        canvas.yview_moveto(max(0.0, min(1.0, frac)))
        # Highlight active nav button
        for i,b in enumerate(self._nav_btns):
            b.configure(fg_color=BG3 if i==idx else "transparent",
                        text_color=TEXT if i==idx else TEXT2,
                        font=("Inter",11,"bold") if i==idx else ("Inter",11))

    def _auto_dv(self, *args):
        nit = self._get_field("nit").strip()
        if len(nit) >= 6:
            dv = calcular_dv(nit)
            if dv: self._set_field("dv", dv)

    # ── Clientes ──────────────────────────────────────────────────────────────
    def _abrir_clientes(self):
        VentanaClientes(self, on_select=self._aplicar_cliente)

    def _aplicar_cliente(self, row):
        self._cliente_id = row[0]
        for k,v in zip(["nit","dv","razonSocial","direccion","telefono",
                         "codSeccional","codDpto","codMunicipio"],
                        row[1:9]):
            self._set_field(k, v or "")
        self.lbl_cliente.configure(text=f"👤 {row[3]}", text_color=GREEN)
        self._calc()

    def _abrir_historial(self): VentanaHistorial(self, self)
    def _abrir_stats(self):     VentanaEstadisticas(self)
    def _abrir_plazos(self):    VentanaPlazos(self)
    def _abrir_checklist(self):
        VentanaChecklist(self, decl_info=f"Cliente: {self._get_field('razonSocial') or '—'}  |  Doc: {self._get_field('docTransporte') or '—'}")
    def _abrir_syga(self):      webbrowser.open("https://importaciones.dian.gov.co")
    def _abrir_multas(self):    VentanaMultas(self)
    def _abrir_poder(self):     VentanaPoder(self, {k:self._get_field(k) for k in self.fields})
    def _abrir_consulta(self):
        w = VentanaConsultaLevante(self)
        w.ent_nit.insert(0, self._get_field("nit"))
        w.ent_doc.insert(0, self._get_field("docTransporte"))

    def _abrir_urgente(self):
        VentanaEntregaUrgente(self, {k:self._get_field(k) for k in self.fields})

    def _abrir_rop(self):
        data = {k:self._get_field(k) for k in self.fields}
        if not self._get_field("nit"):
            messagebox.showwarning("Aviso","Ingrese al menos el NIT antes de generar el ROP.")
            return
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        VentanaROP(self, data, self._total_cop, float(self._get_field("tasaCambio") or 4150))

    def _abrir_edi(self):
        if not self._check_licencia(): return
        data = {k:self._get_field(k) for k in self.fields}
        if not self._get_field("nit"):
            messagebox.showwarning("Aviso","Ingrese los datos del formulario antes de generar el EDI.")
            return
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        VentanaEDI(self, data)

    def _abrir_config(self): VentanaConfigAgencia(self)
    def _subpartidas(self):  VentanaSubpartidas(self)

    def _update_trm(self):
        self.lbl_trm.configure(text="Consultando...", text_color=YELLOW)
        def fetch():
            trm, fecha_trm = None, None
            HEADERS = {"User-Agent":"Mozilla/5.0 (Windows NT 10.0; Win64; x64)","Accept":"application/json"}
            try:
                from datetime import timedelta
                for offset in range(0,10):
                    d = (date.today()-timedelta(days=offset)).strftime("%Y-%m-%d")
                    params = urllib.parse.urlencode({"vigenciadesde":d})
                    url = f"https://www.datos.gov.co/resource/32sa-8pi3.json?{params}"
                    req = urllib.request.Request(url, headers=HEADERS)
                    with urllib.request.urlopen(req, timeout=6) as r:
                        data = json.loads(r.read())
                    if data and "valor" in data[0]:
                        trm = float(str(data[0]["valor"]).replace(",","."))
                        fecha_trm = d; break
            except: pass
            if not trm:
                try:
                    params2 = urllib.parse.urlencode({"$limit":"1","$order":"vigenciadesde DESC"})
                    url2 = f"https://www.datos.gov.co/resource/mcec-87by.json?{params2}"
                    req2 = urllib.request.Request(url2, headers=HEADERS)
                    with urllib.request.urlopen(req2, timeout=6) as r2:
                        data2 = json.loads(r2.read())
                    if data2 and "valor" in data2[0]:
                        trm = float(str(data2[0]["valor"]).replace(",","."))
                        fecha_trm = str(data2[0].get("vigenciadesde",""))[:10]
                except: pass
            if not trm:
                try:
                    req3 = urllib.request.Request("https://open.er-api.com/v6/latest/USD", headers=HEADERS)
                    with urllib.request.urlopen(req3, timeout=6) as r3:
                        data3 = json.loads(r3.read())
                    cop = data3.get("rates",{}).get("COP")
                    if cop: trm=float(cop); fecha_trm=date.today().strftime("%Y-%m-%d")+" (ref.)"
                except: pass
            if trm and trm > 0:
                self.after(0, lambda t=trm, d=fecha_trm: self._apply_trm(t,d))
            else:
                self.after(0, lambda: self.lbl_trm.configure(
                    text="Sin conexión — ingrese TRM manualmente", text_color=RED))
        threading.Thread(target=fetch, daemon=True).start()

    def _apply_trm(self, trm, fecha):
        self._set_field("tasaCambio", f"{trm:.2f}")
        self._calc()
        self.lbl_trm.configure(text=f"TRM {fecha}: ${trm:,.2f}", text_color=GREEN)

    def _load_excel(self):
        if not HAS_XLSX:
            messagebox.showerror("Error","openpyxl no está instalado."); return
        path = filedialog.askopenfilename(title="Seleccionar plantilla Excel",
                                           filetypes=[("Excel","*.xlsx *.xls")])
        if not path: return
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
            ws = wb["Formulario510"]
            EXCEL_MAP = {
                "nit":"B5","dv":"D5","razonSocial":"B6","direccion":"B7","telefono":"D7",
                "codSeccional":"B8","codDpto":"D8","codMunicipio":"B9",
                "nitDecl":"B12","dvDecl":"D12","razonDecl":"B13","tipoUsuario":"D13",
                "nombresDecl":"B14","numDocDecl":"D14","codUsuario":"B15",
                "tipoDecl":"B18","numFormAnterior":"D18","manifestoCarga":"B19",
                "fechaLlegada":"D19","codLugarIngreso":"B20","codModo":"D20",
                "docTransporte":"B21","fechaDocTransporte":"D21",
                "codProcedencia":"B22","tasaCambio":"D22",
                "nombreExportador":"B25","formaPago":"D25","codPaisCompra":"B26",
                "codPaisOrigen":"D26","subpartida":"B27","numBultos":"D27",
                "cantidad":"B28","pesoBruto":"D28","pesoNeto":"B29","fob":"D29",
                "fletes":"B30","seguros":"D30","otrosGastos":"B31","ajuste":"D31",
                "descripcion":"B32","arancelPct":"B35","ivaPct":"D35","icPct":"B36",
            }
            loaded = 0
            for key, cell_addr in EXCEL_MAP.items():
                val = ws[cell_addr].value
                if val is not None:
                    self._set_field(key, str(val).strip()); loaded += 1
            self._calc()
            messagebox.showinfo("Listo",f"✅ {loaded} campos importados desde Excel.")
        except KeyError:
            messagebox.showerror("Error","No se encontró la hoja 'Formulario510'.")
        except Exception as e:
            messagebox.showerror("Error al cargar Excel", str(e))

    def _check_licencia(self):
        lic = db_fetch("SELECT valor FROM config WHERE clave='licencia_activa'")
        if lic and lic[0][0] == "1": return True
        max_t = db_fetch("SELECT valor FROM config WHERE clave='max_decl_trial'")
        max_trial = int(max_t[0][0]) if max_t else 5
        cnt = obtener_contador_uso()
        if cnt >= max_trial:
            messagebox.showwarning("Modo prueba agotado",
                f"Límite de {max_trial} declaraciones alcanzado.\n"
                "Active su licencia en ⚙️ Configuración → Licencia\n\n"
                "Contacto: bentjake15@gmail.com")
            return False
        return True

    def _validar(self):
        errores = []
        for key, label in [("nit","NIT del importador"),("razonSocial","Razón social"),
                            ("fechaLlegada","Fecha de llegada"),("docTransporte","Doc. transporte"),
                            ("subpartida","Subpartida arancelaria"),("fob","Valor FOB"),
                            ("descripcion","Descripción de la mercancía")]:
            if not self._get_field(key).strip():
                errores.append(f"  • {label}")
        return errores

    def _generate(self):
        if not self._check_licencia(): return
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
                messagebox.showwarning("⚠️ Límite",f"CIF ${cif:,.2f} USD supera $2.000 USD.\nUse Formulario 500.")
            elif cif > 200:
                messagebox.showinfo("ℹ️ Aviso",f"CIF ${cif:,.2f} USD supera $200 USD.\nAplican impuestos.")
        except: pass
        path = filedialog.asksaveasfilename(defaultextension=".pdf",
            filetypes=[("PDF","*.pdf")],
            initialfile=f"Formulario510_{date.today().isoformat()}.pdf")
        if not path: return
        data = {k:self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        try:
            make_pdf(data, path)
            if self._decl_id:
                db_exec("UPDATE declaraciones SET pdf_path=?,estado='Generado' WHERE id=?", (path,self._decl_id))
            messagebox.showinfo("PDF Generado",f"✅ PDF generado:\n{path}")
        except Exception as e:
            messagebox.showerror("Error al generar PDF", str(e))

    def _guardar_decl(self):
        if not self._check_licencia(): return
        if not self._cliente_id:
            nit = self._get_field("nit").strip()
            if nit:
                rows = db_fetch("SELECT id FROM clientes WHERE nit=?", (nit,))
                if rows: self._cliente_id = rows[0][0]
            if not self._cliente_id:
                nit = self._get_field("nit").strip()
                razon = self._get_field("razonSocial").strip()
                if nit and razon:
                    self._cliente_id = db_exec(
                        "INSERT INTO clientes(nit,dv,razon_social,direccion,telefono,cod_seccional,cod_dpto,cod_municipio) VALUES(?,?,?,?,?,?,?,?)",
                        (nit,self._get_field("dv"),razon,self._get_field("direccion"),
                         self._get_field("telefono"),self._get_field("codSeccional"),
                         self._get_field("codDpto"),self._get_field("codMunicipio")))
                    self.lbl_cliente.configure(text=f"👤 {razon}", text_color=GREEN)
                else:
                    if not messagebox.askyesno("Sin cliente","¿Guardar sin cliente asociado?"): return
        self._calc()
        data = {k:self._get_field(k) for k in self.fields}
        try:
            fob=float(data.get("fob","0") or 0); flt=float(data.get("fletes","0") or 0)
            seg=float(data.get("seguros","0") or 0); otr=float(data.get("otrosGastos","0") or 0)
            adj=float(data.get("ajuste","0") or 0)
            data["valorAduana"] = f"{fob+flt+seg+otr+adj:.2f}"
        except: pass
        doc_transp = self._get_field("docTransporte")
        fecha = date.today().isoformat()
        if self._decl_id:
            db_exec("UPDATE declaraciones SET datos=?,total_cop=?,numero=?,fecha=?,estado='Borrador' WHERE id=?",
                    (json.dumps(data,ensure_ascii=False), self._total_cop, doc_transp, fecha, self._decl_id))
        else:
            self._decl_id = db_exec(
                "INSERT INTO declaraciones(cliente_id,numero,fecha,datos,total_cop,estado) VALUES(?,?,?,?,?,?)",
                (self._cliente_id, doc_transp, fecha, json.dumps(data,ensure_ascii=False), self._total_cop, "Borrador"))
            incrementar_contador_uso()
        messagebox.showinfo("Guardado","✅ Declaración guardada en el historial.")

    def _clear(self):
        if messagebox.askyesno("Confirmar","¿Limpiar todos los campos?"):
            self._cliente_id = None; self._decl_id = None
            self.lbl_cliente.configure(text="Sin cliente", text_color=TEXT2)
            self._set_defaults()


# ═══════════════════════════════════════════════════════════════════════════════
# VENTANAS SECUNDARIAS — mantienen la lógica de la versión anterior
# con los widgets de customtkinter solo para ventanas nuevas simples
# ═══════════════════════════════════════════════════════════════════════════════

class VentanaSubpartidas(ctk.CTkToplevel):
    def __init__(self, app_ref):
        super().__init__(app_ref)
        self.app = app_ref
        self.title("Subpartidas Arancelarias Comunes")
        self.geometry("560x460")
        self.configure(fg_color=BG)
        self._build()

    def _build(self):
        ctk.CTkLabel(self, text="Subpartidas Arancelarias Comunes",
                     font=("Inter",14,"bold"), text_color=TEXT).pack(pady=(20,4))
        ctk.CTkLabel(self, text="Doble clic para seleccionar",
                     font=("Inter",9), text_color=TEXT2).pack(pady=(0,12))

        frame = ctk.CTkFrame(self, fg_color=BG2, corner_radius=10,
                              border_width=1, border_color=BORDER)
        frame.pack(fill="both", expand=True, padx=20, pady=(0,16))

        import tkinter.ttk as ttk2
        style = ttk2.Style()
        style.theme_use("default")
        style.configure("Dark.Treeview", background=BG2, foreground=TEXT,
                        fieldbackground=BG2, borderwidth=0, rowheight=28)
        style.configure("Dark.Treeview.Heading", background=BG3,
                        foreground=TEXT2, borderwidth=0)
        style.map("Dark.Treeview", background=[("selected", BLUE_DK)])

        cols = ("Producto","Subpartida","Arancel","IVA")
        tree = ttk2.Treeview(frame, columns=cols, show="headings",
                              height=14, style="Dark.Treeview")
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width={"Producto":220,"Subpartida":120,"Arancel":80,"IVA":70}[c])
        for row in SUBPARTIDAS:
            tree.insert("","end",values=row)
        sb = ttk2.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); tree.pack(fill="both", expand=True, padx=2, pady=2)

        def sel(e=None):
            s = tree.selection()
            if not s: return
            v = tree.item(s[0])["values"]
            self.app._set_field("subpartida", v[1])
            self.app._set_field("arancelPct", str(v[2]).replace("%",""))
            self.app._calc(); self.destroy()
        tree.bind("<Double-1>", sel)
        ctk.CTkButton(self, text="Seleccionar", command=sel,
                      font=("Inter",11,"bold"), height=38,
                      fg_color=BLUE, hover_color=BLUE_DK,
                      corner_radius=8).pack(fill="x", padx=20, pady=(0,16))


# ── Reusar clases de ventanas de la versión anterior (tk.Toplevel) ────────────
# Estas ventanas funcionan correctamente con tk.Toplevel — no es necesario migrar
# todas a customtkinter para que la app se vea bien, ya que son ventanas secundarias
# que el usuario abre ocasionalmente.

import tkinter.ttk as ttk

class VentanaClientes(tk.Toplevel):
    def __init__(self, parent, on_select=None):
        super().__init__(parent)
        self.title("Gestión de Clientes"); self.geometry("900x540")
        self.configure(bg=BG2); self.on_select = on_select
        self._build(); self._cargar()

    def _build(self):
        hdr = tk.Frame(self, bg=BLUE_DK, height=48)
        hdr.pack(fill="x"); hdr.pack_propagate(False)
        tk.Label(hdr, text="👥  Gestión de Clientes / Importadores",
                 font=("Inter",12,"bold"), bg=BLUE_DK, fg="white").pack(side="left", padx=16, pady=10)
        tb = tk.Frame(self, bg=BG3, height=46); tb.pack(fill="x"); tb.pack_propagate(False)
        for text, cmd, bg in [
            ("➕ Nuevo", self._nuevo, BLUE),
            ("✏️ Editar", self._editar, "#0f766e"),
            ("🗑️ Eliminar", self._eliminar, "#dc2626")]:
            tk.Button(tb, text=text, font=("Inter",10,"bold"), bg=bg, fg="white",
                      relief="flat", padx=10, pady=5, cursor="hand2",
                      command=cmd).pack(side="left", padx=(8,2), pady=6)
        sf = tk.Frame(tb, bg=BG3); sf.pack(side="right", padx=12, pady=8)
        self.search_var = tk.StringVar(); self.search_var.trace("w", lambda *a: self._cargar())
        tk.Entry(sf, textvariable=self.search_var, font=("Inter",11), width=22,
                 bg=BG, fg=TEXT, relief="flat", insertbackground=TEXT).pack(side="left", ipady=4)
        cols = ("ID","NIT","Razón Social","Dirección","Teléfono","Declaraciones")
        style = ttk.Style(); style.theme_use("default")
        style.configure("D.Treeview", background=BG2, foreground=TEXT,
                        fieldbackground=BG2, rowheight=26)
        style.configure("D.Treeview.Heading", background=BG3, foreground=TEXT2)
        style.map("D.Treeview", background=[("selected",BLUE_DK)])
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=16, style="D.Treeview")
        widths = {"ID":40,"NIT":110,"Razón Social":280,"Dirección":200,"Teléfono":110,"Declaraciones":90}
        for c in cols:
            self.tree.heading(c, text=c); self.tree.column(c, width=widths[c])
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y"); self.tree.pack(fill="both", expand=True, padx=8, pady=8)
        self.tree.bind("<Double-1>", lambda e: self._seleccionar())
        if self.on_select:
            tk.Button(self, text="✅  Seleccionar cliente",
                      font=("Inter",11,"bold"), bg=BLUE, fg="white",
                      relief="flat", pady=10, cursor="hand2",
                      command=self._seleccionar).pack(fill="x", padx=8, pady=(0,8))

    def _cargar(self):
        q = self.search_var.get().strip()
        rows = db_fetch("SELECT id,nit,razon_social,direccion,telefono FROM clientes WHERE razon_social LIKE ? OR nit LIKE ? ORDER BY razon_social",
                        (f"%{q}%",f"%{q}%")) if q else db_fetch("SELECT id,nit,razon_social,direccion,telefono FROM clientes ORDER BY razon_social")
        self.tree.delete(*self.tree.get_children())
        for r in rows:
            cnt = db_fetch("SELECT COUNT(*) FROM declaraciones WHERE cliente_id=?", (r[0],))[0][0]
            self.tree.insert("","end",values=(r[0],r[1],r[2],r[3],r[4],cnt))

    def _get_sel(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("Aviso","Seleccione un cliente."); return None
        return self.tree.item(sel[0])["values"]

    def _nuevo(self): FormCliente(self, callback=self._cargar)
    def _editar(self):
        v = self._get_sel()
        if v: FormCliente(self, cliente_id=v[0], callback=self._cargar)
    def _eliminar(self):
        v = self._get_sel()
        if not v: return
        if messagebox.askyesno("Confirmar", f"¿Eliminar a {v[2]}?"):
            db_exec("DELETE FROM declaraciones WHERE cliente_id=?", (v[0],))
            db_exec("DELETE FROM clientes WHERE id=?", (v[0],))
            self._cargar()
    def _seleccionar(self):
        v = self._get_sel()
        if v and self.on_select:
            row = db_fetch("SELECT * FROM clientes WHERE id=?", (v[0],))
            if row: self.on_select(row[0]); self.destroy()


class FormCliente(tk.Toplevel):
    def __init__(self, parent, cliente_id=None, callback=None):
        super().__init__(parent)
        self.cliente_id = cliente_id; self.callback = callback
        self.title("Nuevo Cliente" if not cliente_id else "Editar Cliente")
        self.geometry("500x420"); self.configure(bg=BG2); self.resizable(False,False)
        self.fields = {}; self._build()
        if cliente_id: self._cargar()

    def _build(self):
        tk.Frame(self, bg=BLUE_DK, height=4).pack(fill="x")
        tk.Label(self, text="Datos del Importador", font=("Inter",12,"bold"),
                 bg=BG2, fg=TEXT).pack(pady=(16,12))
        form = tk.Frame(self, bg=BG2); form.pack(fill="x", padx=24)
        for i in range(2): form.columnconfigure(i, weight=1)
        campos = [
            ("NIT (sin DV)","nit",0,0),("DV","dv",0,1),
            ("Razón social","razonSocial",1,0,2),("Dirección","direccion",2,0,2),
            ("Teléfono","telefono",3,0),("Cód. Seccional","codSeccional",3,1),
            ("Cód. Departamento","codDpto",4,0),("Cód. Municipio","codMunicipio",4,1),
        ]
        for item in campos:
            label,key,row,col = item[0],item[1],item[2],item[3]
            span = item[4] if len(item)>4 else 1
            fc = tk.Frame(form, bg=BG2)
            fc.grid(row=row,column=col,columnspan=span,sticky="ew",padx=4,pady=4)
            tk.Label(fc,text=label,font=("Inter",9),bg=BG2,fg=TEXT2).pack(anchor="w")
            w = tk.Entry(fc,font=("Inter",11),bg=BG3,fg=TEXT,relief="flat",
                         insertbackground=TEXT,highlightbackground=BORDER,highlightthickness=1)
            w.pack(fill="x",ipady=5)
            if key=="nit": w.bind("<FocusOut>", lambda e: self._auto_dv())
            self.fields[key] = w
        btns = tk.Frame(self,bg=BG2); btns.pack(fill="x",padx=24,pady=16)
        tk.Button(btns,text="💾  Guardar",font=("Inter",11,"bold"),bg=BLUE,fg="white",
                  relief="flat",pady=10,cursor="hand2",command=self._guardar
                  ).pack(side="left",fill="x",expand=True,padx=(0,4))
        tk.Button(btns,text="Cancelar",font=("Inter",11),bg=BG3,fg=TEXT2,
                  relief="flat",pady=10,cursor="hand2",command=self.destroy
                  ).pack(side="left",fill="x",expand=True)

    def _auto_dv(self):
        nit = self.fields["nit"].get().strip()
        if len(nit)>=6:
            dv=calcular_dv(nit)
            if dv: self.fields["dv"].delete(0,"end"); self.fields["dv"].insert(0,dv)

    def _cargar(self):
        row = db_fetch("SELECT * FROM clientes WHERE id=?", (self.cliente_id,))
        if not row: return
        r=row[0]
        for k,v in zip(["nit","dv","razonSocial","direccion","telefono",
                         "codSeccional","codDpto","codMunicipio"], r[1:9]):
            w=self.fields.get(k)
            if w: w.delete(0,"end"); w.insert(0,v or "")

    def _guardar(self):
        nit=self.fields["nit"].get().strip(); razon=self.fields["razonSocial"].get().strip()
        if not nit or not razon: messagebox.showwarning("Aviso","NIT y Razón social son obligatorios."); return
        vals=(nit,self.fields["dv"].get(),razon,self.fields["direccion"].get(),
              self.fields["telefono"].get(),self.fields["codSeccional"].get(),
              self.fields["codDpto"].get(),self.fields["codMunicipio"].get())
        if self.cliente_id:
            db_exec("UPDATE clientes SET nit=?,dv=?,razon_social=?,direccion=?,telefono=?,cod_seccional=?,cod_dpto=?,cod_municipio=? WHERE id=?", vals+(self.cliente_id,))
        else:
            db_exec("INSERT INTO clientes(nit,dv,razon_social,direccion,telefono,cod_seccional,cod_dpto,cod_municipio) VALUES(?,?,?,?,?,?,?,?)", vals)
        if self.callback: self.callback()
        self.destroy()


# ── Las demás ventanas secundarias (Historial, ROP, EDI, etc.) se importan
# del archivo original que ya funciona correctamente ──────────────────────────
_orig = open("/home/claude/formulario510_exe/main.py").read()

# Extract each secondary window class from original
import types
_ns = {
    "tk":tk,"ttk":ttk,"messagebox":messagebox,"filedialog":filedialog,
    "date":date,"datetime":datetime,"timedelta":timedelta,
    "os":os,"sys":sys,"json":json,"threading":threading,
    "hashlib":hashlib,"shutil":shutil,"webbrowser":webbrowser,
    "urllib":urllib,"db_fetch":db_fetch,"db_exec":db_exec,
    "fmt_cop":fmt_cop,"get_machine_id":get_machine_id,
    "get_app_data_dir":get_app_data_dir,"get_db_path":get_db_path,
    "calcular_dv":calcular_dv,"generar_clave_licencia":generar_clave_licencia,
    "validar_clave_licencia":validar_clave_licencia,
    "incrementar_contador_uso":incrementar_contador_uso,
    "obtener_contador_uso":obtener_contador_uso,
    "make_pdf":make_pdf,"init_db":init_db,
    "HAS_WINREG":HAS_WINREG,"HAS_XLSX":HAS_XLSX,
    "SUBPARTIDAS":SUBPARTIDAS,"PAISES":PAISES,
    "BG":BG,"BG2":BG2,"BG3":BG3,"BORDER":BORDER,"TEXT":TEXT,"TEXT2":TEXT2,
    "BLUE":BLUE,"BLUE_DK":BLUE_DK,"GREEN":GREEN,"RED":RED,"YELLOW":YELLOW,
    "LICENSE_SECRET":LICENSE_SECRET,
}

for cls_name in ["VentanaHistorial","VentanaChecklist","VentanaPlazos",
                  "VentanaROP","VentanaEDI","VentanaConsultaLevante",
                  "VentanaEntregaUrgente","VentanaConfigAgencia",
                  "VentanaEstadisticas","VentanaMultas","VentanaPoder"]:
    start = _orig.find(f"\nclass {cls_name}")
    nxt = _orig.find("\nclass ", start+1)
    code = _orig[start:nxt if nxt>0 else len(_orig)]
    exec(code, _ns)
    globals()[cls_name] = _ns[cls_name]


# ═══════════════════════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    init_db()
    root = ctk.CTk()
    root.withdraw()

    def on_login(user, rol):
        App(root, user, rol)

    def on_splash_ready():
        VentanaLogin(root, on_login)

    SplashScreen(root, on_splash_ready)
    root.mainloop()
