import tkinter as tk
from tkinter import messagebox
import matplotlib.pyplot as plt

from openpyxl import Workbook
from datetime import datetime

def calcular():
    try:
        capital = float(entrada_capital.get())
        taxa_ano= float(entrada_taxa.get()) / 100
        tempo_ano = float(entrada_tempo.get())
        aporte = float(entrada_aporte.get())
        tipo = tipo_juros.get()
        freq = frequencia.get()

        if freq == "Anual":
            n = 1
        elif freq == "Mensal":
            n = 12
        elif freq == "Diária":
            n = 365

        taxa_periodo = taxa_ano / n
        tempo_total = int(tempo_ano * n)

        montantes = []
        valor = capital

        for t in range(tempo_total + 1):
            if tipo == "composto":
                if t > 0:
                    valor = valor * (1 + taxa_periodo) + aporte
            else:
                if t > 0:
                    valor = capital * (1 + taxa_periodo * t) + aporte * t
                else:
                    valor = capital
            montantes.append(valor)

        montante_final = montantes[-1]

        valor_aportado = aporte * tempo_total
        juros_total = montante_final - capital - valor_aportado

        resultado_var.set(
            f"Montate: R$ {montante_final:.2f}\n"
            f"Aportado: R$ {valor_aportado:.2f}\n"
            f"Juros: R$ {juros_total:.2f}")

        plt.figure(figsize=(8,4))
        plt.plot(range(tempo_total + 1), montantes, marker='o', color='purple')
        plt.title(f'Crecimento do Juros ({tipo.capitalize()}, {freq.lower()})')
        plt.xlabel(f'Período ({freq.lower()})')
        plt.ylabel('Valor (R$)')
        plt.grid(True)
        plt.tight_layout()
        plt.show()
        exportar_excel(montantes, taxa_periodo, aporte, capital, tipo, freq)


    except ValueError:
        messagebox.showerror("Error", "Por Favor, preencha todos os campos corretamente")


def limpar():
    entrada_capital.delete(0, tk.END)
    entrada_taxa.delete(0, tk.END)
    entrada_tempo.delete(0, tk.END)
    tipo_juros.set("composto")
    frequencia.set("Mensal")
    resultado_var.set("")

def exportar_excel(montantes, taxa_periodo, aporte, capital, tipo, freq):
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "Histórico de Juros"

        ws.append(["Período", "Juros Acumulados (R$)", "Montante Total (R$)"])

        for t, montante in enumerate(montantes):
            if tipo == "composto":
                if t == 0:
                    juros == 0
                else:
                    juros = montante - (capital * (1 + taxa_periodo) ** t) if aporte == 0 else montante - capital - aporte * t
            else:
                juros = montante - capital - aporte * t
        
            ws.append([t, round(juros, 2), round(montante, 2)])
        
        from datetime import datetime
        import os

        nome_arquivo = f"historico_juros_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        caminho = os.path.abspath(nome_arquivo)
        wb.save(nome_arquivo)
        print(f"Arquivo salvo em: {caminho}")
        messagebox.showinfo("Exportado", f"Tabela salva em:\n{caminho}")

    except Exception as e:
        messagebox.showerror("Erro ao salvar", f"Ocorreu um erro:\n{str(e)}")
        print(f"Erro ao salvar arquivo: {e}")

janela = tk.Tk()
janela.title("Calculadora de Juros")
janela.geometry("380x420")

tk.Label(janela, text="Capital Inicial (R$):").pack()
entrada_capital = tk.Entry(janela)
entrada_capital.pack()

tk.Label(janela, text="Taxa de Juros (%):").pack()
entrada_taxa = tk.Entry(janela)
entrada_taxa.pack()

tk.Label(janela, text="Tempo (em Períodos):").pack()
entrada_tempo = tk.Entry(janela)
entrada_tempo.pack()

tk.Label(janela, text="Aporte por período (R$):").pack()
entrada_aporte = tk.Entry(janela)
entrada_aporte.pack()

tipo_juros = tk.StringVar(value="composto")
frame_opcoes = tk.Frame(janela)
tk.Radiobutton(frame_opcoes, text="Juros Compostos", variable=tipo_juros, value="composto").pack(side=tk.LEFT, padx=10)
tk.Radiobutton(frame_opcoes, text="Juros Simples", variable=tipo_juros, value="simples").pack(side=tk.LEFT, padx=10)
frame_opcoes.pack(pady=10)

tk.Label(janela, text="Frequência de capitalização:").pack()
frequencia = tk.StringVar(value="Mensal")
opcoes_freq = ["Anual", "Mensal", "Diária"]
tk.OptionMenu(janela, frequencia, *opcoes_freq).pack()

tk.Button(janela, text="Calcular", command=calcular).pack(pady=10)
tk.Button(janela, text="Limpar Tudo", command=limpar).pack(pady=5)

resultado_var = tk.StringVar()

tk.Label(janela, textvariable=resultado_var, font=("Arial", 12), fg="blue").pack(pady=10)

janela.mainloop()