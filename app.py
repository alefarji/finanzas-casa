from datetime import date, datetime
import json
import os
import pandas as pd
import plotly.express as px
import streamlit as st

# 1. CONFIGURACIÓN DE PÁGINA Y ESTILOS VISUALES
st.set_page_config(
    page_title="Finanzas Casa | Ale, Denu y Mila",
    page_icon="🏠",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
        .main { background-color: #f8fafc; }
        div[data-testid="stMetric"] {
            background-color: #ffffff;
            padding: 16px 20px;
            border-radius: 14px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 2px 6px rgba(15, 23, 42, 0.04);
        }
        .progress-card {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            color: white;
            padding: 22px;
            border-radius: 16px;
            margin-bottom: 20px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        }
        .bank-card {
            background: white;
            padding: 16px;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            margin-bottom: 12px;
        }
    </style>
""",
    unsafe_allow_html=True,
)

ARCHIVO_DATOS = "datos_finanzas_hogar.json"


# 2. DATOS BASE DEL EXCEL (Los 24 gastos exactos de Noviembre 2026)
def obtener_gastos_base(mes_num=11, anio_num=2026):
  m_str = f"{mes_num:02d}"
  base = [
      (
          "GIMNASIO DENI",
          50.0,
          f"05/{m_str}/{anio_num}",
          "Mensual",
          "Denu",
          "kutx 5377",
          "Denu",
          "Personal / Salud",
          "Fijo",
      ),
      (
          "LUZ",
          85.0,
          f"26/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 5358",
          "Ale",
          "Suministros",
          "Variable",
      ),
      (
          "AGUA",
          84.0,
          f"06/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Ale",
          "Suministros",
          "Variable",
      ),
      (
          "MOVIL E INTERNET",
          72.0,
          f"06/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 5358",
          "Ale",
          "Suministros",
          "Fijo",
      ),
      (
          "Santander Tarjeta",
          0.0,
          f"06/{m_str}/{anio_num}",
          "Mensual",
          "Denu",
          "Santander",
          "Denu",
          "Tarjetas y Préstamos",
          "Variable",
      ),
      (
          "Sanitas BBVA Mila",
          92.0,
          f"06/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "BBVA",
          "Ale",
          "Mila (Colegio y Salud)",
          "Fijo",
      ),
      (
          "Sanitas Dental",
          14.0,
          f"07/{m_str}/{anio_num}",
          "Mensual",
          "Denu",
          "Santander",
          "Denu",
          "Personal / Salud",
          "Fijo",
      ),
      (
          "Sanitas BBVA Ale",
          92.0,
          f"11/{m_str}/{anio_num}",
          "Mensual",
          "Ale",
          "BBVA",
          "Ale",
          "Personal / Salud",
          "Fijo",
      ),
      (
          "SEGURO MOTO",
          74.0,
          f"11/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 5358",
          "Ale",
          "Vehículos / Seguros",
          "Fijo",
      ),
      (
          "Kutxabank Tarjeta A",
          100.0,
          f"11/{m_str}/{anio_num}",
          "Mensual",
          "Ale",
          "kutx 5358",
          "Ale",
          "Tarjetas y Préstamos",
          "Variable",
      ),
      (
          "Kutxabank Tarjeta D",
          0.0,
          f"01/{m_str}/{anio_num}",
          "Mensual",
          "Denu",
          "kutx",
          "Denu",
          "Tarjetas y Préstamos",
          "Variable",
      ),
      (
          "BBVA Tarjeta",
          100.0,
          f"11/{m_str}/{anio_num}",
          "Mensual",
          "Ale",
          "BBVA",
          "Ale",
          "Tarjetas y Préstamos",
          "Variable",
      ),
      (
          "Movil Ale",
          84.0,
          f"12/{m_str}/{anio_num}",
          "Mensual",
          "Ale",
          "BBVA",
          "Ale",
          "Personal / Salud",
          "Fijo",
      ),
      (
          "Comida",
          600.0,
          f"13/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "N26",
          "Denu",
          "Comida y Supermercado",
          "Variable",
      ),
      (
          "HIPOTECA",
          898.0,
          f"14/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Ale",
          "Hipoteca y Vivienda",
          "Fijo",
      ),
      (
          "Prestamo Gravity",
          400.0,
          f"15/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 5358",
          "Ale",
          "Tarjetas y Préstamos",
          "Fijo",
      ),
      (
          "Prestamo BBVA",
          154.0,
          f"16/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "BBVA",
          "Ale",
          "Tarjetas y Préstamos",
          "Fijo",
      ),
      (
          "Ahorro Kutxabanc",
          125.0,
          f"17/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Denu",
          "Ahorro Familiar",
          "Fijo",
      ),
      (
          "Comunidad",
          48.0,
          f"18/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Denu",
          "Hipoteca y Vivienda",
          "Fijo",
      ),
      (
          "ChatGPT/YouTube/AMAZON",
          77.0,
          f"19/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx",
          "Ale",
          "Suscripciones y Hogar",
          "Fijo",
      ),
      (
          "Lina",
          70.0,
          f"15/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "BBVA",
          "Ale",
          "Suscripciones y Hogar",
          "Fijo",
      ),
      (
          "Psico",
          80.0,
          f"15/{m_str}/{anio_num}",
          "Mensual",
          "Denu",
          "Santander",
          "Denu",
          "Personal / Salud",
          "Fijo",
      ),
      (
          "Maternal Extra",
          100.0,
          f"01/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Denu",
          "Mila (Colegio y Salud)",
          "Variable",
      ),
      (
          "Maternal",
          490.0,
          f"01/{m_str}/{anio_num}",
          "Mensual",
          "Ambos",
          "kutx 4844",
          "Denu",
          "Mila (Colegio y Salud)",
          "Fijo",
      ),
  ]
  lista = []
  for i, g in enumerate(base):
    lista.append({
        "id": i + 1,
        "concepto": g[0],
        "valor": g[1],
        "fecha": g[2],
        "estado": g[3],
        "de_quien": g[4],
        "banco": g[5],
        "quien_paga": g[6],
        "categoria": g[7],
        "tipo": g[8],
        "forzar_manual": (  # None = sigue la fecha automática; True/False = forzado manual
            None
        ),
    })
  return lista


# 3. SISTEMA DE GUARDADO PERMANENTE (JSON EN DISCO)
def cargar_bd():
  if os.path.exists(ARCHIVO_DATOS):
    try:
      with open(ARCHIVO_DATOS, "r", encoding="utf-8") as f:
        return json.load(f)
    except Exception:
      pass
  return {
      "meses": {
          "Noviembre 2026": {
              "salarios": {"Ale": 3000.0, "Denu": 2200.0},
              "gastos": obtener_gastos_base(11, 2026),
          }
      }
  }


def guardar_bd(datos):
  with open(ARCHIVO_DATOS, "w", encoding="utf-8") as f:
    json.dump(datos, f, ensure_ascii=False, indent=2)


if "bd" not in st.session_state:
  st.session_state.bd = cargar_bd()

# 4. BARRA LATERAL: MES, SALARIOS Y SIMULADOR DE FECHA AUTOMÁTICA
st.sidebar.title("🏠 Panel Familiar")

MAPA_MESES = {
    "Noviembre 2026": (11, 2026),
    "Diciembre 2026": (12, 2026),
    "Enero 2027": (1, 2027),
    "Febrero 2027": (2, 2027),
    "Marzo 2027": (3, 2027),
    "Abril 2027": (4, 2027),
}

mes_actual = st.sidebar.selectbox(
    "📅 Seleccionar Mes", list(MAPA_MESES.keys()), index=0
)
mes_num, anio_num = MAPA_MESES[mes_actual]

# Si el mes no existe aún en la base de datos, lo creamos adaptando las fechas al nuevo mes
if mes_actual not in st.session_state.bd["meses"]:
  ultimo_mes = list(st.session_state.bd["meses"].keys())[-1]
  datos_previos = st.session_state.bd["meses"][ultimo_mes]
  nuevos_gastos = []
  for g in datos_previos["gastos"]:
    copia = dict(g)
    try:
      dia = int(str(g["fecha"]).split("/")[0])
    except Exception:
      dia = 1
    copia["fecha"] = f"{dia:02d}/{mes_num:02d}/{anio_num}"
    copia["forzar_manual"] = None  # Reiniciamos pagos para el nuevo mes
    nuevos_gastos.append(copia)

  st.session_state.bd["meses"][mes_actual] = {
      "salarios": dict(datos_previos["salarios"]),
      "gastos": nuevos_gastos,
  }
  guardar_bd(st.session_state.bd)

datos_mes = st.session_state.bd["meses"][mes_actual]

# Control de Auto-pago por Fecha
st.sidebar.markdown("---")
st.sidebar.subheader("⚡ Auto-Pago por Fecha")
auto_pago_activo = st.sidebar.toggle(
    "Marcar Pagado automáticamente al llegar la fecha", value=True
)

fecha_referencia = st.sidebar.date_input(
    "Fecha de referencia (para comprobar cobros)",
    value=date(anio_num, mes_num, 10),
    format="DD/MM/YYYY",
)

# Gestión de Salarios / Ingresos del Mes
st.sidebar.markdown("---")
st.sidebar.subheader("💰 Salarios / Ingresos del Mes")
sal_ale = st.sidebar.number_input(
    "Ingresos netos Ale (€)",
    min_value=0.0,
    value=float(datos_mes["salarios"].get("Ale", 3000.0)),
    step=50.0,
)
sal_denu = st.sidebar.number_input(
    "Ingresos netos Denu (€)",
    min_value=0.0,
    value=float(datos_mes["salarios"].get("Denu", 2200.0)),
    step=50.0,
)

if (
    sal_ale != datos_mes["salarios"].get("Ale")
    or sal_denu != datos_mes["salarios"].get("Denu")
):
  datos_mes["salarios"]["Ale"] = sal_ale
  datos_mes["salarios"]["Denu"] = sal_denu
  guardar_bd(st.session_state.bd)


# 5. CÁLCULO DE ESTADO PAGADO / NO PAGADO (AUTOMÁTICO + MANUAL)
def parsear_fecha(fecha_str):
  try:
    return datetime.strptime(str(fecha_str).strip(), "%d/%m/%Y").date()
  except Exception:
    return date(anio_num, mes_num, 1)


filas_procesadas = []
for g in datos_mes["gastos"]:
  f_cobro = parsear_fecha(g["fecha"])
  vencido_por_fecha = f_cobro <= fecha_referencia

  if g.get("forzar_manual") is not None:
    esta_pagado = bool(g["forzar_manual"])
    origen_estado = "Manual"
  else:
    esta_pagado = vencido_por_fecha if auto_pago_activo else False
    origen_estado = (
        "⚡ Auto por fecha"
        if (esta_pagado and auto_pago_activo)
        else "Pendiente"
    )

  item = dict(g)
  item["pagado_efectivo"] = esta_pagado
  item["origen_estado"] = origen_estado
  item["fecha_dt"] = f_cobro
  filas_procesadas.append(item)

df = pd.DataFrame(filas_procesadas)

# 6. CABECERA Y MÉTRICAS GLOBALES (INCLUYENDO SALARIOS Y AHORRO)
st.title(f"🏠 Finanzas Casa — {mes_actual}")

ingresos_totales = sal_ale + sal_denu
total_gastos = df["valor"].sum() if not df.empty else 0.0
total_pagado = (
    df[df["pagado_efectivo"]]["valor"].sum() if not df.empty else 0.0
)
total_pendiente = total_gastos - total_pagado
pct_pagado = (total_pagado / total_gastos * 100) if total_gastos > 0 else 0.0
pct_pendiente = 100.0 - pct_pagado
disponible_mes = ingresos_totales - total_gastos
pct_sueldo_gastado = (
    (total_gastos / ingresos_totales * 100) if ingresos_totales > 0 else 0.0
)

gasto_comida = (
    df[df["categoria"] == "Comida y Supermercado"]["valor"].sum()
    if not df.empty
    else 0.0
)
gasto_hipoteca = (
    df[df["categoria"] == "Hipoteca y Vivienda"]["valor"].sum()
    if not df.empty
    else 0.0
)
gasto_mila = (
    df[df["categoria"] == "Mila (Colegio y Salud)"]["valor"].sum()
    if not df.empty
    else 0.0
)

# Panel Oscuro: ¿En qué parte del mes estamos?
pendiente_ale = (
    df[(~df["pagado_efectivo"]) & (df["quien_paga"] == "Ale")]["valor"].sum()
    if not df.empty
    else 0.0
)
pendiente_denu = (
    df[(~df["pagado_efectivo"]) & (df["quien_paga"] == "Denu")]["valor"].sum()
    if not df.empty
    else 0.0
)

st.markdown(
    f"""
<div class="progress-card">
    <div style="display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:10px;">
        <div>
            <span style="background:#10b981; color:#052e16; padding:4px 10px; border-radius:99px; font-size:12px; font-weight:700;">
                CONTROL DE COBROS DEL MES ({fecha_referencia.strftime('%d/%m/%Y')})
            </span>
            <h3 style="margin:8px 0 2px 0; color:white;">¿En qué parte del mes estamos?</h3>
        </div>
        <div style="text-align:right;">
            <span style="font-size:14px; color:#94a3b8;">Ya cobrado: <b style="color:#34d399;">{total_pagado:,.0f} € ({pct_pagado:.1f}%)</b></span> &nbsp;|&nbsp;
            <span style="font-size:14px; color:#94a3b8;">Falta por salir: <b style="color:#fbbf24;">{total_pendiente:,.0f} € ({pct_pendiente:.1f}%)</b></span>
        </div>
    </div>
    <div style="width:100%; background:#334155; height:16px; border-radius:99px; margin:14px 0; overflow:hidden;">
        <div style="width:{min(100.0, pct_pagado):.1f}%; background:linear-gradient(90deg, #10b981, #34d399); height:100%;"></div>
    </div>
    <div style="display:flex; justify-content:space-between; font-size:13px; color:#cbd5e1; flex-wrap:wrap; gap:8px;">
        <span>👨‍💻 Falta salir de cuentas de <b>Ale</b>: <b style="color:#fbbf24;">{pendiente_ale:,.0f} €</b></span>
        <span>👩 Falta salir de cuentas de <b>Denu</b>: <b style="color:#fbbf24;">{pendiente_denu:,.0f} €</b></span>
        <span>💰 Saldo libre tras todos los gastos: <b style="color:#38bdf8;">{disponible_mes:,.0f} €</b> ({100 - pct_sueldo_gastado:.1f}% del ingreso familiar)</span>
    </div>
</div>
""".replace(",", "."),
    unsafe_allow_html=True,
)

# 5 Tarjetas de Resumen Rápido con Porcentajes Clave
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric(
    "Total Gastos Mes",
    f"{total_gastos:,.0f} €".replace(",", "."),
    f"{pct_sueldo_gastado:.1f}% de vuestros sueldos",
)
c2.metric(
    "🏠 Hipoteca + Comunidad",
    f"{gasto_hipoteca:,.0f} €".replace(",", "."),
    f"{(gasto_hipoteca/total_gastos*100 if total_gastos else 0):.1f}% del gasto",
)
c3.metric(
    "🛒 Comida / Súper",
    f"{gasto_comida:,.0f} €".replace(",", "."),
    f"{(gasto_comida/total_gastos*100 if total_gastos else 0):.1f}% del gasto",
)
c4.metric(
    "👧🏻 Mila (Maternal + Sanitas)",
    f"{gasto_mila:,.0f} €".replace(",", "."),
    f"{(gasto_mila/total_gastos*100 if total_gastos else 0):.1f}% del gasto",
)
c5.metric(
    "💵 Libre / Ahorro Neto",
    f"{disponible_mes:,.0f} €".replace(",", "."),
    f"De {ingresos_totales:,.0f} € ingresados".replace(",", "."),
)

# 7. PESTAÑAS PRINCIPALES
tab_tabla, tab_bancos, tab_graficos, tab_nuevo = st.tabs([
    "📋 Tabla de Gastos (Editar Fechas, Importes y Pagos)",
    "🏦 Dinero por Banco y Quién Paga",
    "📊 Porcentajes (Hipoteca, Comida, Mila, Fijos vs Variables)",
    "➕ Añadir Nuevo Gasto / Backup",
])

# --- PESTAÑA 1: TABLA INTERACTIVA COMPLETA ---
with tab_tabla:
  col_f1, col_f2, col_f3 = st.columns([2, 2, 4])
  with col_f1:
    if st.button("⚡ Dejar que todas las fechas decidan (Modo Auto)"):
      for g in datos_mes["gastos"]:
        g["forzar_manual"] = None
      guardar_bd(st.session_state.bd)
      st.rerun()
  with col_f2:
    if st.button("✅ Marcar TODOS como Pagados"):
      for g in datos_mes["gastos"]:
        g["forzar_manual"] = True
      guardar_bd(st.session_state.bd)
      st.rerun()

  st.caption(
      "💡 **Cómo editar:** Haz doble clic en cualquier celda para cambiar la"
      " **Fecha de cobro**, el **Valor (€)**, el **Banco** o marca/desmarca la"
      " casilla **Pagado**. También puedes borrar una fila seleccionándola a"
      " la izquierda y pulsando `Supr`, o añadir una al final."
  )

  if not df.empty:
    df_para_editar = pd.DataFrame({
        "Pagado": df["pagado_efectivo"],
        "Modo": df["origen_estado"],
        "Concepto": df["concepto"],
        "Valor (€)": df["valor"],
        "Fecha Cobro": df["fecha_dt"],
        "De quién es": df["de_quien"],
        "Banco": df["banco"],
        "Quién paga": df["quien_paga"],
        "Categoría": df["categoria"],
        "Tipo": df["tipo"],
    })

    bancos_disponibles = sorted(
        list(
            set(
                df["banco"].tolist()
                + [
                    "kutx 4844",
                    "kutx 5358",
                    "kutx 5377",
                    "kutx",
                    "BBVA",
                    "N26",
                    "Santander",
                    "Efectivo",
                ]
            )
        )
    )
    categorias_disponibles = [
        "Hipoteca y Vivienda",
        "Comida y Supermercado",
        "Mila (Colegio y Salud)",
        "Suministros",
        "Tarjetas y Préstamos",
        "Personal / Salud",
        "Vehículos / Seguros",
        "Ahorro Familiar",
        "Suscripciones y Hogar",
    ]

    tabla_editada = st.data_editor(
        df_para_editar,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        disabled=["Modo"],
        column_config={
            "Pagado": st.column_config.CheckboxColumn(
                "✅ Pagado",
                help="Marca o desmarca para cambiar a Pagado / No Pagado",
            ),
            "Modo": st.column_config.TextColumn("Estado Auto"),
            "Concepto": st.column_config.TextColumn(
                "GASTOS MES", required=True
            ),
            "Valor (€)": st.column_config.NumberColumn(
                "Valor (€)", format="%.2f €", min_value=0.0, required=True
            ),
            "Fecha Cobro": st.column_config.DateColumn(
                "Fecha Cobro", format="DD/MM/YYYY", required=True
            ),
            "De quién es": st.column_config.SelectboxColumn(
                "DE QUIÉN ES", options=["Ambos", "Ale", "Denu"], required=True
            ),
            "Banco": st.column_config.SelectboxColumn(
                "BANCO", options=bancos_disponibles, required=True
            ),
            "Quién paga": st.column_config.SelectboxColumn(
                "Quién paga", options=["Ale", "Denu"], required=True
            ),
            "Categoría": st.column_config.SelectboxColumn(
                "Categoría (%)", options=categorias_disponibles, required=True
            ),
            "Tipo": st.column_config.SelectboxColumn(
                "Fijo / Variable", options=["Fijo", "Variable"], required=True
            ),
        },
        key=f"tabla_gastos_{mes_actual}",
    )

    if st.button(
        "💾 Guardar cambios hechos en la tabla",
        type="primary",
        use_container_width=True,
    ):
      lista_actualizada = []
      for i, fila in tabla_editada.iterrows():
        if pd.isna(fila["Concepto"]) or str(fila["Concepto"]).strip() == "":
          continue
        f_dt = fila["Fecha Cobro"]
        if isinstance(f_dt, (date, datetime)):
          f_str = f_dt.strftime("%d/%m/%Y")
          vencido = f_dt <= fecha_referencia
        else:
          f_str = f"01/{mes_num:02d}/{anio_num}"
          vencido = False

        pagado_usuario = bool(fila["Pagado"])
        auto_esperado = vencido if auto_pago_activo else False

        # Si el usuario puso un estado distinto al automático, lo guardamos como manual
        forzar = None if (pagado_usuario == auto_esperado) else pagado_usuario

        lista_actualizada.append({
            "id": i + 1,
            "concepto": str(fila["Concepto"]),
            "valor": float(fila["Valor (€)"]),
            "fecha": f_str,
            "estado": "Mensual",
            "de_quien": str(fila["De quién es"]),
            "banco": str(fila["Banco"]),
            "quien_paga": str(fila["Quién paga"]),
            "categoria": str(fila["Categoría"]),
            "tipo": str(fila["Tipo"]),
            "forzar_manual": forzar,
        })
      datos_mes["gastos"] = lista_actualizada
      guardar_bd(st.session_state.bd)
      st.success(
          "✅ ¡Todos los cambios de fechas, importes y estados se han guardado"
          " permanentemente!"
      )
      st.rerun()

# --- PESTAÑA 2: DINERO POR BANCO Y QUIÉN PAGA ---
with tab_bancos:
  col_izq, col_der = st.columns([1.2, 1])

  with col_izq:
    st.subheader("🏦 Saldo por Banco (Total vs. Ya Cobrado vs. Falta Salir)")
    resumen_bancos = []
    for banco, grupo in df.groupby("banco"):
      tot_b = grupo["valor"].sum()
      pag_b = grupo[grupo["pagado_efectivo"]]["valor"].sum()
      pen_b = tot_b - pag_b
      resumen_bancos.append({
          "Banco / Cuenta": banco,
          "Total Mes (€)": tot_b,
          "Ya Cobrado (€)": pag_b,
          "Falta Salir (€)": pen_b,
          "Recibos Pendientes": len(grupo[~grupo["pagado_efectivo"]]),
      })
    df_bancos = pd.DataFrame(resumen_bancos).sort_values(
        by="Total Mes (€)", ascending=False
    )
    st.dataframe(df_bancos, use_container_width=True, hide_index=True)

  with col_der:
    st.subheader("⚖️ Balance entre Ale y Denu")
    comunes_total = df[df["de_quien"] == "Ambos"]["valor"].sum()
    propios_ale = df[df["de_quien"] == "Ale"]["valor"].sum()
    propios_denu = df[df["de_quien"] == "Denu"]["valor"].sum()

    paga_ale_total = df[df["quien_paga"] == "Ale"]["valor"].sum()
    paga_denu_total = df[df["quien_paga"] == "Denu"]["valor"].sum()

    st.info(
        "👨‍💻 **Ale paga desde sus cuentas:**"
        f" **{paga_ale_total:,.0f} €**\n\n• Ya cobrado:"
        f" {df[(df['quien_paga']=='Ale') & (df['pagado_efectivo'])]['valor'].sum():,.0f} €"
        f" | **Falta salir: {pendiente_ale:,.0f} €**\n• Gastos propios de Ale:"
        f" {propios_ale:,.0f} € | De Ambos que paga Ale:"
        f" {paga_ale_total - propios_ale:,.0f} €"
    )
    st.success(
        "👩 **Denu paga desde sus cuentas:**"
        f" **{paga_denu_total:,.0f} €**\n\n• Ya cobrado:"
        f" {df[(df['quien_paga']=='Denu') & (df['pagado_efectivo'])]['valor'].sum():,.0f} €"
        f" | **Falta salir: {pendiente_denu:,.0f} €**\n• Gastos propios de"
        f" Denu: {propios_denu:,.0f} € | De Ambos que paga Denu:"
        f" {paga_denu_total - propios_denu:,.0f} €"
    )
    st.caption(
        f"📌 **Total gastos de 'Ambos':** {comunes_total:,.0f} € (50% ="
        f" {comunes_total/2:,.1f} € cada uno)."
    )

# --- PESTAÑA 3: GRÁFICOS DE PORCENTAJES ---
with tab_graficos:
  g_col1, g_col2 = st.columns(2)
  with g_col1:
    st.subheader("📊 % de Gasto por Categoría (Hipoteca, Comida, Mila...)")
    df_cat = df.groupby("categoria", as_index=False)["valor"].sum()
    df_cat = df_cat[df_cat["valor"] > 0].sort_values(
        by="valor", ascending=False
    )
    fig_pie = px.pie(
        df_cat,
        values="valor",
        names="categoria",
        hole=0.45,
        color_discrete_sequence=px.colors.qualitative.Safe,
    )
    fig_pie.update_traces(
        textposition="inside", textinfo="percent+label+value"
    )
    st.plotly_chart(fig_pie, use_container_width=True)

  with g_col2:
    st.subheader("📌 Gastos Fijos vs. Gastos Variables")
    df_tipo = df.groupby("tipo", as_index=False)["valor"].sum()
    fig_bar = px.bar(
        df_tipo,
        x="tipo",
        y="valor",
        color="tipo",
        text_auto=".0f",
        color_discrete_map={"Fijo": "#3b82f6", "Variable": "#f59e0b"},
    )
    st.plotly_chart(fig_bar, use_container_width=True)

# --- PESTAÑA 4: AÑADIR NUEVO GASTO Y COPIA DE SEGURIDAD ---
with tab_nuevo:
  col_n1, col_n2 = st.columns([1.4, 1])
  with col_n1:
    st.subheader("➕ Añadir un gasto extra a este mes")
    with st.form("form_alta_rapida"):
      f1, f2 = st.columns(2)
      with f1:
        n_conc = st.text_input(
            "Concepto (ej. Ropa Mila, Taller moto, Cena...)"
        )
        n_val = st.number_input("Importe (€)", min_value=0.0, step=5.0)
        n_fec = st.date_input(
            "Fecha de cobro", value=date(anio_num, mes_num, 15)
        )
        n_cat = st.selectbox("Categoría", categorias_disponibles)
      with f2:
        n_deq = st.selectbox("¿De quién es?", ["Ambos", "Ale", "Denu"])
        n_ban = st.selectbox("Banco por donde sale", bancos_disponibles)
        n_qp = st.selectbox("¿Quién lo paga?", ["Ale", "Denu"])
        n_tip = st.selectbox("Tipo de gasto", ["Variable", "Fijo"])

      if st.form_submit_button("➕ Guardar nuevo gasto", type="primary"):
        if n_conc.strip():
          datos_mes["gastos"].append({
              "id": len(datos_mes["gastos"]) + 1,
              "concepto": n_conc.strip(),
              "valor": float(n_val),
              "fecha": n_fec.strftime("%d/%m/%Y"),
              "estado": "Mensual",
              "de_quien": n_deq,
              "banco": n_ban,
              "quien_paga": n_qp,
              "categoria": n_cat,
              "tipo": n_tip,
              "forzar_manual": None,
          })
          guardar_bd(st.session_state.bd)
          st.success(f"✅ Gasto '{n_conc}' añadido a {mes_actual}.")
          st.rerun()

  with col_n2:
    st.subheader("💾 Copia de Seguridad (Backup)")
    st.write(
        "Descarga todos vuestros meses o restaura una copia cuando quieras:"
    )
    json_bytes = json.dumps(
        st.session_state.bd, ensure_ascii=False, indent=2
    ).encode("utf-8")
    st.download_button(
        "⬇️ Descargar copia de seguridad (.json)",
        data=json_bytes,
        file_name="finanzas_ale_denu_backup.json",
        mime="application/json",
    )
