import sys, json, os, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import numpy as np

datos = json.loads(sys.stdin.read())
out_dir = sys.argv[1] if len(sys.argv) > 1 else "/tmp/graficos_informe"
os.makedirs(out_dir, exist_ok=True)
paths = {}

meses  = datos.get("historial_meses", ["Ene","Feb","Mar","Abr","May"])
hist_v = datos.get("historial_tension", [])
hist_c = datos.get("historial_corriente", [])
hist_t = datos.get("historial_temperatura", [])

def base_fig(title, ylabel, ct="#1F4E79"):
    fig, ax = plt.subplots(figsize=(7.2, 3.0))
    fig.patch.set_facecolor('white')
    ax.set_facecolor("#F8F8F8")
    ax.set_title(title, fontsize=11, fontweight='bold', color=ct, pad=8, loc='left')
    ax.set_ylabel(ylabel, fontsize=8, color='#666')
    ax.tick_params(axis='both', labelsize=8, colors='#555')
    for s in ['top','right']: ax.spines[s].set_visible(False)
    ax.spines['left'].set_color('#CCC')
    ax.spines['bottom'].set_color('#CCC')
    ax.grid(axis='y', color='#E0E0E0', linewidth=0.6, linestyle='--')
    ax.set_axisbelow(True)
    return fig, ax

# Tensiones históricas
if hist_v:
    fig, ax = base_fig("Tensiones históricas (V)", "Voltios")
    x = range(len(meses))
    rs = [m.get('rs') for m in hist_v]
    rt = [m.get('rt') for m in hist_v]
    st = [m.get('st') for m in hist_v]
    if any(v is not None for v in rs):
        ax.plot(x, rs, 'o-', color='#378ADD', linewidth=2, markersize=5, label='R/S')
        ax.plot(x, rt, 's--', color='#85B7EB', linewidth=1.5, markersize=4, label='R/T')
        ax.plot(x, st, '^:', color='#B5D4F4', linewidth=1.2, markersize=4, label='S/T')
        ax.set_xticks(list(x)); ax.set_xticklabels(meses)
        ax.legend(fontsize=8, framealpha=0.7, loc='lower right')
        for i, v in enumerate(rs):
            if v: ax.annotate(str(v), (i,v), textcoords="offset points", xytext=(0,6), ha='center', fontsize=7, color='#378ADD')
    p = f"{out_dir}/graf_tension.png"
    fig.tight_layout(); fig.savefig(p, dpi=130, bbox_inches='tight', facecolor='white'); plt.close(fig)
    paths['tension'] = p

# Corrientes históricas
if hist_c:
    fig, ax = base_fig("Corrientes históricas (A)", "Amperes", "#1F6B3E")
    x = range(len(meses))
    cr = [m.get('r') for m in hist_c]
    cs = [m.get('s') for m in hist_c]
    ct2 = [m.get('t') for m in hist_c]
    if any(v is not None for v in cr):
        ax.plot(x, cr, 'o-', color='#3B6D11', linewidth=2, markersize=5, label='R')
        ax.plot(x, cs, 's--', color='#97C459', linewidth=1.5, markersize=4, label='S')
        ax.plot(x, ct2, '^:', color='#C0DD97', linewidth=1.2, markersize=4, label='T')
        ax.set_xticks(list(x)); ax.set_xticklabels(meses)
        ax.legend(fontsize=8, framealpha=0.7, loc='lower right')
        for i, v in enumerate(cr):
            if v: ax.annotate(str(v), (i,v), textcoords="offset points", xytext=(0,6), ha='center', fontsize=7, color='#3B6D11')
    p = f"{out_dir}/graf_corriente.png"
    fig.tight_layout(); fig.savefig(p, dpi=130, bbox_inches='tight', facecolor='white'); plt.close(fig)
    paths['corriente'] = p

# Temperaturas históricas
if hist_t:
    fig, ax = base_fig("Temperaturas históricas (°C)", "Grados °C", "#7F6000")
    x = range(len(meses))
    tr = [m.get('r') for m in hist_t]
    ts2 = [m.get('s') for m in hist_t]
    tt = [m.get('t') for m in hist_t]
    if any(v is not None for v in tr):
        ax.plot(x, tr, 'o-', color='#EF9F27', linewidth=2, markersize=5, label='R')
        ax.plot(x, ts2, 's--', color='#FAC775', linewidth=1.5, markersize=4, label='S')
        ax.plot(x, tt, '^:', color='#FAEEDA', linewidth=1.2, markersize=4, label='T')
        ax.axhline(y=35, color='#E24B4A', linewidth=0.8, linestyle='--', alpha=0.6, label='Límite 35°C')
        ax.set_xticks(list(x)); ax.set_xticklabels(meses)
        ax.legend(fontsize=8, framealpha=0.7, loc='lower right')
        for i, v in enumerate(tr):
            if v: ax.annotate(str(v), (i,v), textcoords="offset points", xytext=(0,6), ha='center', fontsize=7, color='#EF9F27')
    p = f"{out_dir}/graf_temperatura.png"
    fig.tight_layout(); fig.savefig(p, dpi=130, bbox_inches='tight', facecolor='white'); plt.close(fig)
    paths['temperatura'] = p

# Barras tensión mes actual
tableros = datos.get("tableros", [])
nombres = [t['nombre'][:14] for t in tableros if t.get('v_rs') is not None]
v_rs_mes = [t['v_rs'] for t in tableros if t.get('v_rs') is not None]
if nombres:
    fig, ax = base_fig("Tensión R/S por tablero — mes actual (V)", "Voltios")
    x2 = np.arange(len(nombres))
    ax.bar(x2, v_rs_mes, color='#378ADD', width=0.6, edgecolor='white', linewidth=0.5)
    ax.set_xticks(x2); ax.set_xticklabels(nombres, rotation=35, ha='right', fontsize=7)
    avg = float(np.mean(v_rs_mes))
    ax.axhline(y=avg, color='#E24B4A', linewidth=1, linestyle='--', alpha=0.7, label=f'Promedio: {avg:.0f}V')
    ax.legend(fontsize=8)
    p = f"{out_dir}/graf_barras_tension.png"
    fig.tight_layout(); fig.savefig(p, dpi=130, bbox_inches='tight', facecolor='white'); plt.close(fig)
    paths['barras_tension'] = p

print(json.dumps(paths))
