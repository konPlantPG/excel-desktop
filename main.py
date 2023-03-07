import tkinter as tk
from tkinter import filedialog

import matplotlib.pyplot as plt
import japanize_matplotlib
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import pandas as pd



class GraphPlotter:
    def __init__(self, master):
        self.master = master
        master.title("Graph Plotter")

        self.file_path = ""
        self.data = pd.DataFrame()

        #ラベルとリストボックスの作成
        self.x_label = tk.Label(master, text="x軸")
        self.y_label = tk.Label(master, text="y軸")
        self.x_listbox = tk.Listbox(master, selectmode=tk.SINGLE, selectforeground="red", exportselection=False)
        self.y_listbox = tk.Listbox(master, selectmode=tk.SINGLE, selectforeground="red", exportselection=False)

        #ラベルとリストボックスの配置
        self.x_label.grid(row=1, column=0)
        self.y_label.grid(row=1, column=1)
        self.x_listbox.grid(row=2, column=0)
        self.y_listbox.grid(row=2, column=1)

        
        #ファイル読み込みとグラフ描画のボタン
        self.file_button = tk.Button(master, text="Open File", command=self.load_file)
        self.plot_button = tk.Button(master, text="Plot Graph", command=self.plot_graph)

        #グラフ種類を選択するラジオボタン
        self.graph_type = tk.StringVar(value="line")
        self.line_rb = tk.Radiobutton(master, text="Line", variable=self.graph_type, value="line")
        self.scatter_rb = tk.Radiobutton(master, text="Scatter", variable=self.graph_type, value="scatter")

        #グラフの描画
        self.fig = plt.Figure(figsize=(5, 4), dpi=100)
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=master)
        self.canvas.get_tk_widget().grid(row=3, column=0, columnspan=2)

        #ボタンや文字の配置
        self.file_button.grid(row=4, column=0, pady=10)
        self.plot_button.grid(row=4, column=1, pady=10)
        self.line_rb.grid(row=5, column=0)
        self.scatter_rb.grid(row=5, column=1)

    #ファイルを読み込む
    def load_file(self):
        self.file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if self.file_path:
            #pandasで選んだデータを読み込む
            self.data = pd.read_excel(self.file_path)
            columns = self.data.columns.tolist()
            self.x_listbox.delete(0, tk.END)
            self.y_listbox.delete(0, tk.END)

            for column in columns:
                self.x_listbox.insert(tk.END, column)
                self.y_listbox.insert(tk.END, column)

    #グラフの描画
    def plot_graph(self):
        self.ax.clear()

        x_column = self.x_listbox.get(tk.ACTIVE)
        y_column = self.y_listbox.get(tk.ACTIVE)
        x_data = self.data[x_column]
        y_data = self.data[y_column]

        self.ax.set_xlabel(x_column)
        self.ax.set_ylabel(y_column)

        if self.graph_type.get() == "line":
            self.ax.plot(x_data, y_data)
        elif self.graph_type.get() == "scatter":
            self.ax.scatter(x_data, y_data)

        self.canvas.draw()

def main():
    root = tk.Tk()
    GraphPlotter(root)
    root.mainloop()

if __name__=='__main__':
    main()