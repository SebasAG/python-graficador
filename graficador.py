import numpy as np
import tkinter as tk
# import matplotlib as mpl
# mpl.use('TkAgg')
import matplotlib.figure as fg
from matplotlib.backends.tkagg import blit
from matplotlib.backends.backend_agg import FigureCanvasAgg
# from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib import rcParams
import openpyxl as xl

rcParams.update({
    'figure.autolayout': True,
    'font.size': 16,
    'font.family': 'serif',
    'font.serif': ['Computer Modern Roman'],
    'text.usetex': True,
    'text.latex.unicode': True,
    'text.latex.preamble': [r'\usepackage{siunitx}'
                            r'\usepackage{physics}']
})


def draw_figure(canvas, figure, loc=(0, 0)):
    """
    Draw a matplotlib figure onto a Tk grafica

    loc: location of top-left corner of figure on grafica in pixels.

    Inspired by matplotlib source: lib/matplotlib/backends/backend_tkagg.py
    """
    figure_canvas_agg = FigureCanvasAgg(figure)
    figure_canvas_agg.draw()
    figure_x, figure_y, figure_w, figure_h = figure.bbox.bounds
    figure_w, figure_h = int(figure_w+1), int(figure_h+1)
    photo = tk.PhotoImage(master=canvas, width=figure_w, height=figure_h)

    # Position: convert from top-left anchor to center anchor
    canvas.create_image(loc[0] + figure_w/2, loc[1] + figure_h/2, image=photo)
    # Unfortunatly, there's no accessor for the pointer to the native renderer
    blit(photo, figure_canvas_agg.get_renderer()._renderer, colormode=2)

    # Return a handle which contains a reference to the photo object
    # which must be kept live or else the picture disappears
    return photo


def ajuste_lineal(datos_x, datos_y):
    datos_xy = datos_y * datos_x
    datos_xx = datos_x ** 2
    num = len(datos_x)
    y_prom = np.mean(datos_y)

    suma_x = np.sum(datos_x)
    suma_y = np.sum(datos_y)
    suma_xx = np.sum(datos_xx)
    suma_xy = np.sum(datos_xy)

    delta = num * suma_xx - suma_x ** 2

    m = (num * suma_xy - suma_x * suma_y) / delta
    b = (suma_xx * suma_y - suma_x * suma_xy) / delta

    tot = (datos_y - y_prom) ** 2
    recta = m * datos_x + b
    res = (datos_y - recta) ** 2

    suma_tot = np.sum(tot)
    suma_res = np.sum(res)

    rcuad = 1 - suma_res / suma_tot

    d_est = np.sqrt(suma_res / (num - 2))
    d_m = d_est * np.sqrt(num / delta)
    d_b = d_est * np.sqrt(suma_xx / delta)

    return m, d_m, b, d_b, d_est, rcuad, recta


def importar_datos(archivo):
    global datosxy, x_label, y_label

    wb = xl.load_workbook(archivo, data_only=True, read_only=True)
    sheet = wb.active

    x_vals = []
    for celda in sheet.columns[int(colx.get())-1]:
        x_vals.append(celda.value)

    x_label = x_vals.pop(0)
    etiqueta_x.delete(0, tk.END)
    etiqueta_x.insert(tk.END, x_label)
    etiqueta_x.grid_forget()
    etiqueta_x.grid(row=4, column=0)

    x_inc = []
    for celda in sheet.columns[int(coldx.get())-1]:
        x_inc.append(celda.value)
    del x_inc[0]

    y_vals = []
    for celda in sheet.columns[int(coly.get())-1]:
        y_vals.append(celda.value)

    y_label = y_vals.pop(0)
    etiqueta_y.delete(0, tk.END)
    etiqueta_y.insert(tk.END, y_label)
    etiqueta_y.grid_forget()
    etiqueta_y.grid(row=5, column=0)

    y_inc = []
    for celda in sheet.columns[int(coldy.get())-1]:
        y_inc.append(celda.value)
    del y_inc[0]

    datosxy = np.array(x_vals), np.array(x_inc), np.array(y_vals), np.array(y_inc)

    wb._archive.close()
    print("Datos cargados.")
    fig_nombre.delete(0, tk.END)
    fig_nombre.insert(tk.END, archivo[:-5]+".png")


def graficar():
    global datosxy, fig_photo, fig, ax, x_label, y_label

    x_label = etiqueta_x.get()  # .replace("\\", "\\\\")
    y_label = etiqueta_y.get()  # .replace("\\", "\\\\")

    fig = fg.Figure()
    ax = fig.add_subplot(111)

    ax.errorbar(datosxy[0], datosxy[2], xerr=datosxy[1], yerr=datosxy[3], fmt='.', zorder=3)

    ax.set_xlabel(x_label)
    ax.set_ylabel(y_label)
    ax.minorticks_on()
    ax.grid(which='major', zorder=0, linestyle='-', color='0.75')
    ax.grid(which='minor', zorder=0, color='0.75')

    xmin.set(ax.get_xlim()[0])
    xmax.set(ax.get_xlim()[1])
    ymin.set(ax.get_ylim()[0])
    ymax.set(ax.get_ylim()[1])

    fig_photo = draw_figure(grafica, fig)
    grafica.config(width=fig.bbox.bounds[2], height=fig.bbox.bounds[3])


def aplicar_ajuste():
    global m, d_m, b, d_b, d_est, rcuad, recta, fig_photo, ecu_label, recta_ecu, ajuste, ecuacion

    m, d_m, b, d_b, d_est, rcuad, recta = ajuste_lineal(datosxy[0], datosxy[2])

    ecu = r"$y = ({:.2f} \pm {:.2f}) x + ({:.2f} \pm {:.2f})$".format(m, d_m, b, d_b)

    ecu_label.grid_forget()
    ecu_label.grid(row=7, column=0, sticky=tk.SW)

    recta_ecu.grid_forget()
    recta_ecu.grid(row=8, column=0)

    if recta_ecu.get() == '':
        recta_ecu.insert(tk.END, ecu)

    ajuste = ax.plot(datosxy[0], recta, 'g', zorder=2)
    ecuacion = ax.annotate(
        recta_ecu.get() + '\n' + r"$R^2 = {:6.4f}$".format(rcuad),
        xy=(0, 1), xycoords='axes fraction', xytext=(10, -45), textcoords='offset points',
        bbox=dict(boxstyle='round', fc='w', alpha=0.5), size=16
    )

    fig_photo = draw_figure(grafica, fig)
    grafica.config(width=fig.bbox.bounds[2], height=fig.bbox.bounds[3])


def ajustar_graf():
    global fig_photo
    
    ax.set_xlim(float(xmin.get()), float(xmax.get()))
    ax.set_ylim(float(ymin.get()), float(ymax.get()))

    fig_photo = draw_figure(grafica, fig)
    grafica.config(width=fig.bbox.bounds[2], height=fig.bbox.bounds[3])


def logx_switch():
    global isxlog, fig_photo

    if isxlog:
        ax.set_xscale('linear')
    else:
        ax.set_xscale('log')
    fig_photo = draw_figure(grafica, fig)
    grafica.config(width=fig.bbox.bounds[2], height=fig.bbox.bounds[3])


def logy_switch():
    global isylog, fig_photo

    if isylog:
        ax.set_yscale('linear')
    else:
        ax.set_yscale('log')
    fig_photo = draw_figure(grafica, fig)
    grafica.config(width=fig.bbox.bounds[2], height=fig.bbox.bounds[3])



def num_cifras(datos):
    diglist = []
    for num in datos:
        string = '{:e}'.format(num)
        digit = int(string.partition('e')[2])
        diglist.append(max(0, -digit))
    return diglist


def gen_tab1(tabla, datos, datos_inc):
    cifras = num_cifras(datos_inc)

    tabla.delete("1.0", tk.END)

    tabla.insert("1.0", "\\begin{center}\n")
    tabla.insert("2.0", "\t\\begin{small}\n")
    tabla.insert("3.0", "\t\\tablefirsthead{%")
    tabla.insert("4.0", "\t\t$n$ & {} \\\\ \\hline}}\n".format(
        etiqueta_x.get()
    ))
    tabla.insert("5.0", "\t\\tablehead{%")
    tabla.insert("6.0", "\t\t\\multicolumn{2}{c}{Continuación de la tabla anterior.} \\\\")
    tabla.insert("7.0", "\t\t$n$ & {} \\\\ \\hline}}\n".format(
        etiqueta_x.get()
    ))
    tabla.insert("8.0", "\t\\tabletail{%")
    tabla.insert("9.0", "\t\t\\multicolumn{2}{c}{...} \\\\}")
    tabla.insert("10.0", "\t\\tablelasttail{}")
    tabla.insert("11.0", "\t\\bottomcaption{}")
    tabla.insert("12.0", "\t\\begin{supertabular}{c|c}\n")

    for i in range(len(datos)):
        tabla.insert(
            "{}.0".format(i+13),
            "\t\t{} & ${: .{digs}f} \pm {: .{digs}f}$ \\\\\n".format(
                i+1, datos[i], datos_inc[i], digs=cifras[i]
            )
        )

    tabla.insert("{}.0".format(len(datosxy[0]) + 13), "\t\\end{supertabular}\n")
    tabla.insert("{}.0".format(len(datosxy[0]) + 14), "\t\\label{tab:}\n")
    tabla.insert("{}.0".format(len(datosxy[0]) + 15), "\t\\end{small}\n")
    tabla.insert("{}.0".format(len(datosxy[0]) + 16), "\\end{center}\n")


def gen_tab2():
    global datosxy

    x_digs, y_digs = num_cifras(datosxy[1]), num_cifras(datosxy[3])

    tabla_una.delete("1.0", tk.END)

    tabla_una.insert("1.0", "\\begin{center}\n")
    tabla_una.insert("2.0", "\t\\begin{small}\n")
    tabla_una.insert("3.0", "\t\\tablefirsthead{%")
    tabla_una.insert("4.0", "\t\t$n$ & {} & {} \\\\ \\hline}}\n".format(
        etiqueta_x.get(), etiqueta_y.get()
    ))
    tabla_una.insert("5.0", "\t\\tablehead{%")
    tabla_una.insert("6.0", "\t\t\\multicolumn{3}{c}{Continuación de la tabla anterior.} \\\\")
    tabla_una.insert("7.0", "\t\t$n$ & {} & {} \\\\ \\hline}}\n".format(
        etiqueta_x.get(), etiqueta_y.get()
    ))
    tabla_una.insert("8.0", "\t\\tabletail{%")
    tabla_una.insert("9.0", "\t\t\\multicolumn{3}{c}{...} \\\\}")
    tabla_una.insert("10.0", "\t\\tablelasttail{}")
    tabla_una.insert("11.0", "\t\\bottomcaption{}")
    tabla_una.insert("12.0", "\t\\begin{supertabular}{c|cc}\n")

    for i in range(len(datosxy[0])):
        tabla_una.insert(
            "{}.0".format(i+13),
            "\t\t{} & ${: .{digsx}f} \\pm {: .{digsx}f}$ & ${: .{digsy}f} \\pm {: .{digsy}f}$ \\\\\n".format(
                i+1, datosxy[0][i], datosxy[1][i], datosxy[2][i], datosxy[3][i], digsx=x_digs[i], digsy=y_digs[i]
            )
        )

    tabla_una.insert("{}.0".format(len(datosxy[0]) + 13), "\t\\end{supertabular}\n")
    tabla_una.insert("{}.0".format(len(datosxy[0]) + 14), "\t\\label{tab:}\n")
    tabla_una.insert("{}.0".format(len(datosxy[0]) + 15), "\t\\end{small}\n")
    tabla_una.insert("{}.0".format(len(datosxy[0]) + 16), "\\end{center}\n")


def generar_tablas():
    global datosxy

    tabla_grupo.grid(row=16, column=0, columnspan=2)

    tabla_una.grid_forget()
    tabla_x.grid_forget()
    tabla_y.grid_forget()

    if opt_var.get() == opciones[0]:
        gen_tab2()
        tabla_una.grid(row=0, column=0)
    else:
        gen_tab1(tabla_x, datosxy[0], datosxy[1])
        gen_tab1(tabla_y, datosxy[2], datosxy[3])
        tabla_x.grid(row=0, column=0)
        tabla_y.grid(row=0, column=1)


# Crear ventana
window = tk.Tk()
window.title("Graficador de datosxy de Excel")


# Gráfica en la ventana (canvas)
grafica = tk.Canvas(window)
grafica.grid(row=0, column=1, rowspan=15)


# Importar datosxy
x_label = 'x <LaTeX>'
y_label = 'y <LaTeX>'

import_grupo = tk.LabelFrame(window)
import_grupo.grid(row=0, column=0, rowspan=3)

tk.Label(import_grupo, text="Cargar archivo de datosxy:").grid(row=0, column=0, sticky=tk.SW, columnspan=8)

datos_archivo = tk.Entry(import_grupo, width=24)
datos_archivo.insert(tk.END, "datos.xlsx")
datos_archivo.grid(row=1, column=0, columnspan=8)


tk.Label(import_grupo, text="x").grid(row=2, column=0, sticky=tk.W)
colx = tk.StringVar()
colx_in = tk.Entry(import_grupo, textvariable=colx, width=3)
colx_in.grid(row=2, column=1)
colx_in.insert(tk.END, 1)

tk.Label(import_grupo, text="dx").grid(row=2, column=2, sticky=tk.W)
coldx = tk.StringVar()
coldx_in = tk.Entry(import_grupo, textvariable=coldx, width=3)
coldx_in.grid(row=2, column=3)
coldx_in.insert(tk.END, 2)

tk.Label(import_grupo, text="y").grid(row=2, column=4, sticky=tk.W)
coly = tk.StringVar()
coly_in = tk.Entry(import_grupo, textvariable=coly, width=3)
coly_in.grid(row=2, column=5)
coly_in.insert(tk.END, 3)

tk.Label(import_grupo, text="dy").grid(row=2, column=6, sticky=tk.W)
coldy = tk.StringVar()
coldy_in = tk.Entry(import_grupo, textvariable=coldy, width=3)
coldy_in.grid(row=2, column=7)
coldy_in.insert(tk.END, 4)


tk.Label(window, text="Nombres de los ejes:").grid(row=3, column=0, sticky=tk.SW)

etiqueta_x = tk.Entry(window)
etiqueta_x.insert(tk.END, x_label)
etiqueta_x.grid(row=4, column=0)

etiqueta_y = tk.Entry(window)
etiqueta_y.insert(tk.END, y_label)
etiqueta_y.grid(row=5, column=0)

datos_cargar = tk.Button(import_grupo, text="Cargar",
                         command=lambda: importar_datos(datos_archivo.get()))
datos_cargar.grid(row=3, column=0, sticky=tk.N, columnspan=8)


# Graficar datosxy
datos_graficar = tk.Button(window, text="Graficar / Resetear gráfica", command=graficar)
datos_graficar.grid(row=6, column=0)


# Ajustar graficas
grafajus_grupo = tk.LabelFrame(window)
grafajus_grupo.grid(row=0, column=2)

tk.Label(grafajus_grupo, text="x:").grid(row=0, column=0)
tk.Label(grafajus_grupo, text="y:").grid(row=1, column=0)
tk.Label(grafajus_grupo, text="-").grid(row=0, column=2)
tk.Label(grafajus_grupo, text="-").grid(row=1, column=2)

xmin = tk.StringVar()
xmin_in = tk.Entry(grafajus_grupo, textvariable=xmin, width=3)
xmin_in.grid(row=0, column=1)

xmax = tk.StringVar()
xmax_in = tk.Entry(grafajus_grupo, textvariable=xmax, width=3)
xmax_in.grid(row=0, column=3)

ymin = tk.StringVar()
ymin_in = tk.Entry(grafajus_grupo, textvariable=ymin, width=3)
ymin_in.grid(row=1, column=1)

ymax = tk.StringVar()
ymax_in = tk.Entry(grafajus_grupo, textvariable=ymax, width=3)
ymax_in.grid(row=1, column=3)

grafajus_bot = tk.Button(grafajus_grupo, text="Ajustar Gráfica", command=ajustar_graf)
grafajus_bot.grid(row=2, column=0, columnspan=4)

isxlog, isylog = False, False
logx_bot = tk.Button(grafajus_grupo, text="logX", command=logx_switch)
logx_bot.grid(row=3, column=0, columnspan=2)
logy_bot = tk.Button(grafajus_grupo, text="logY", command=logy_switch)
logy_bot.grid(row=3, column=2, columnspan=2)


# Ajuste lineal
datos_ajuste = tk.Button(window, text="Ajuste lineal", command=aplicar_ajuste)
datos_ajuste.grid(row=9, column=0)

ecu_label = tk.Label(window, text='Ecuación:')
recta_ecu = tk.Entry(window)


# Guardar imagen
tk.Label(window, text="Nombre:").grid(row=10, column=0, sticky=tk.SW)

fig_nombre = tk.Entry(window)
fig_nombre.insert(tk.END, "fig.png")
fig_nombre.grid(row=11, column=0)

fig_guardar = tk.Button(window, text="Guardar",
                        command=lambda: fig.savefig(fig_nombre.get()))
fig_guardar.grid(row=12, column=0, sticky=tk.N)


# Generar tabla de LaTeX
tk.Label(window, text="Generar tabla en LaTeX:").grid(row=13, column=0, sticky=tk.SW)

opciones = [
    'Una sola tabla',
    'Tablas diferentes'
]

opt_var = tk.StringVar(window)
opt_var.set(opciones[0])

tabla_opt = tk.OptionMenu(window, opt_var, *opciones)
tabla_opt.grid(row=14, column=0)

tabla_boton = tk.Button(window, text="Generar", command=generar_tablas)
tabla_boton.grid(row=15, column=0, sticky=tk.N)

tabla_grupo = tk.LabelFrame(window)

tabla_una = tk.Text(tabla_grupo)
tabla_x = tk.Text(tabla_grupo)
tabla_y = tk.Text(tabla_grupo)


# Let Tk take over
tk.mainloop()
