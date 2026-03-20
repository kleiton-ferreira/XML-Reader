import customtkinter as ctk
from tkinter import filedialog, messagebox
import xml.etree.ElementTree as ET
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
from reportlab.lib import colors
from datetime import datetime

ctk.set_appearance_mode("light")


class XMLIntelligenceUltra(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("XML Reader")
        self.geometry("1400x900")
        self.configure(fg_color="#F4F7FA")

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        # Variáveis de Controle
        self.parsed_rows = []
        self.total_geral = 0.0
        self.file_count = 0
        self.row_index = 0

        # Definição de Colunas e Larguras
        self.headers = ["TIPO", "NÚMERO", "EMISSÃO", "VALOR", "EMITENTE", "DESTINATÁRIO"]
        self.col_widths = [120, 150, 170, 200, 450, 450]
        self.total_table_width = sum(self.col_widths) + 120

        # --- SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color="#FFFFFF", border_width=1,
                                    border_color="#DDE1E7")
        self.sidebar.grid(row=0, column=0, sticky="nsew")

        ctk.CTkLabel(self.sidebar, text="XML Reader", font=("Inter", 32, "bold"), text_color="#1A73E8").pack(pady=(60, 40))

        self.btn_import = ctk.CTkButton(self.sidebar, text="📂 Importar XMLs", command=self.import_files,
                                        font=("Segoe UI", 16, "bold"), height=55, corner_radius=12, fg_color="#1A73E8")
        self.btn_import.pack(padx=30, pady=10)

        self.btn_excel = ctk.CTkButton(self.sidebar, text="📊 Exportar Excel", command=self.export_to_excel,
                                       font=("Segoe UI", 14), height=45, fg_color="transparent", border_width=1.5,
                                       border_color="#2E7D32", text_color="#2E7D32", hover_color="#E8F5E9")
        self.btn_excel.pack(padx=30, pady=5)

        self.btn_pdf = ctk.CTkButton(self.sidebar, text="📕 Exportar PDF", command=self.export_to_pdf,
                                     font=("Segoe UI", 14), height=45, fg_color="transparent", border_width=1.5,
                                     border_color="#C62828", text_color="#C62828", hover_color="#FFEBEE")
        self.btn_pdf.pack(padx=30, pady=5)

        self.btn_clear = ctk.CTkButton(self.sidebar, text="Limpar Dashboard", command=self.clear_data,
                                       font=("Segoe UI", 13), height=40, fg_color="transparent", border_width=1.2,
                                       border_color="#B0B8C1", text_color="#78909C", hover_color="#ECEFF1")
        self.btn_clear.pack(padx=30, pady=(40, 10))

        # --- CONTEÚDO ---
        self.main_container = ctk.CTkFrame(self, fg_color="transparent")
        self.main_container.grid(row=0, column=1, padx=40, pady=40, sticky="nsew")

        self.stats_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.stats_frame.pack(fill="x", pady=(0, 30))
        self.card_total = self.create_stat_card(self.stats_frame, "VALOR TOTAL", "R$ 0,00", "#1A73E8")
        self.card_files = self.create_stat_card(self.stats_frame, "DOCUMENTOS", "0", "#2E7D32")

        self.view_container = ctk.CTkFrame(self.main_container, fg_color="#FFFFFF", corner_radius=20, border_width=1,
                                           border_color="#DDE1E7")
        self.view_container.pack(fill="both", expand=True)

        self.header_canvas = ctk.CTkCanvas(self.view_container, height=70, bg="#FFFFFF", highlightthickness=0)
        self.header_canvas.pack(fill="x", padx=30, pady=(20, 0))
        self.header_frame = ctk.CTkFrame(self.header_canvas, fg_color="transparent")
        self.header_canvas.create_window((0, 0), window=self.header_frame, anchor="nw")

        for i, text in enumerate(self.headers):
            lbl = ctk.CTkLabel(self.header_frame, text=text, font=("Segoe UI", 16, "bold"),
                               text_color="#455A64", width=self.col_widths[i], anchor="w")
            lbl.grid(row=0, column=i, padx=10, pady=15)

        self.data_container = ctk.CTkFrame(self.view_container, fg_color="transparent")
        self.data_container.pack(fill="both", expand=True, padx=30, pady=(0, 10))
        self.v_scroll = ctk.CTkScrollbar(self.data_container, orientation="vertical")
        self.v_scroll.pack(side="right", fill="y")
        self.h_scroll = ctk.CTkScrollbar(self.view_container, orientation="horizontal", height=18)
        self.h_scroll.pack(side="bottom", fill="x", padx=30, pady=10)

        self.data_canvas = ctk.CTkCanvas(self.data_container, bg="#FFFFFF", highlightthickness=0,
                                         yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set)
        self.data_canvas.pack(side="left", fill="both", expand=True)
        self.v_scroll.configure(command=self.data_canvas.yview)
        self.h_scroll.configure(command=self._sync_scrolls)

        self.data_inner_frame = ctk.CTkFrame(self.data_canvas, fg_color="white")
        self.data_canvas.create_window((0, 0), window=self.data_inner_frame, anchor="nw")

        self.update_scroll_regions()
        self.data_inner_frame.bind("<Configure>", lambda e: self.update_scroll_regions())
        self.data_canvas.bind_all("<MouseWheel>",
                                  lambda e: self.data_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _sync_scrolls(self, *args):
        self.data_canvas.xview(*args)
        self.header_canvas.xview(*args)

    def update_scroll_regions(self):
        self.data_canvas.configure(scrollregion=(0, 0, self.total_table_width, self.data_inner_frame.winfo_height()))
        self.header_canvas.configure(scrollregion=(0, 0, self.total_table_width, 70))

    def create_stat_card(self, parent, title, value, color):
        card = ctk.CTkFrame(parent, fg_color="#FFFFFF", height=130, width=350, corner_radius=15, border_width=1,
                            border_color="#DDE1E7")
        card.pack(side="left", padx=10)
        card.pack_propagate(False)
        ctk.CTkLabel(card, text=title, font=("Segoe UI", 12, "bold"), text_color="#94A3B8").pack(pady=(20, 0), padx=25,
                                                                                                 anchor="w")
        lbl = ctk.CTkLabel(card, text=value, font=("Segoe UI", 38, "bold"), text_color=color)
        lbl.pack(padx=25, anchor="w")
        return lbl

    def show_success_popup(self, msg):
        popup = ctk.CTkToplevel(self)
        popup.title("Sucesso")
        popup.geometry("380x200")
        popup.attributes("-topmost", True)
        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - 190
        y = self.winfo_y() + (self.winfo_height() // 2) - 100
        popup.geometry(f"+{x}+{y}")
        ctk.CTkLabel(popup, text="✔️", font=("Segoe UI", 45)).pack(pady=(20, 5))
        ctk.CTkLabel(popup, text=msg, font=("Segoe UI", 15)).pack(pady=5)
        ctk.CTkButton(popup, text="OK", command=popup.destroy, width=120, corner_radius=8).pack(pady=10)

    def import_files(self):
        files = filedialog.askopenfilenames(filetypes=[("XML", "*.xml")])
        if not files: return
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe', 'cte': 'http://www.portalfiscal.inf.br/cte'}

        for file in files:
            try:
                tree = ET.parse(file)
                root = tree.getroot()
                is_nfe = root.find('.//nfe:infNFe', ns) is not None
                tipo = "💰 NFe" if is_nfe else "📄 CTe"

                if is_nfe:
                    num = root.find('.//nfe:ide/nfe:nNF', ns).text
                    data_raw = root.find('.//nfe:ide/nfe:dhEmi', ns).text[:10]
                    v = float(root.find('.//nfe:total/nfe:ICMSTot/nfe:vNF', ns).text)
                    emit = root.find('.//nfe:emit/nfe:xNome', ns).text
                    dest = root.find('.//nfe:dest/nfe:xNome', ns).text
                else:
                    num = root.find('.//cte:ide/cte:nCT', ns).text
                    data_raw = root.find('.//cte:ide/cte:dhEmi', ns).text[:10]
                    v = float(root.find('.//cte:vPrest/cte:vTPrest', ns).text)
                    emit = root.find('.//cte:emit/cte:xNome', ns).text
                    dest = root.find('.//cte:dest/cte:xNome', ns).text

                # --- Formatação da Data para PT-BR (DD/MM/AAAA) ---
                data_obj = datetime.strptime(data_raw, "%Y-%m-%d")
                data_br = data_obj.strftime("%d/%m/%Y")

                self.parsed_rows.append([tipo, num, data_br, v, emit, dest])
                v_f = f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

                for col, text in enumerate([tipo, num, data_br, v_f, emit, dest]):
                    e = ctk.CTkEntry(self.data_inner_frame, width=self.col_widths[col],
                                     font=("Segoe UI", 15), fg_color="transparent", border_width=0)
                    e.insert(0, text)
                    e.configure(state="readonly")
                    e.grid(row=self.row_index, column=col, padx=10, pady=4, sticky="w")

                self.row_index += 1
                self.total_geral += v
                self.file_count += 1
            except Exception as e:
                print(f"Erro ao processar arquivo: {e}")
                continue

        self.card_total.configure(
            text=f"R$ {self.total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        self.card_files.configure(text=str(self.file_count))
        self.update_scroll_regions()

    def export_to_excel(self):
        if not self.parsed_rows: return
        name = f"Relatorio_XML_{datetime.now().strftime('%d-%m-%Y')}.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=name)
        if path:
            pd.DataFrame(self.parsed_rows, columns=self.headers).to_excel(path, index=False)
            self.show_success_popup("Excel exportado com sucesso!")

    def export_to_pdf(self):
        if not self.parsed_rows: return
        name = f"Relatorio_XML_{datetime.now().strftime('%d-%m-%Y')}.pdf"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", initialfile=name)
        if path:
            doc = SimpleDocTemplate(path, pagesize=landscape(A4))
            pdf_data = [self.headers]
            for r in self.parsed_rows:
                # r[2] já está no formato brasileiro devido à alteração no import_files
                row = [r[0].replace("💰 ", "").replace("📄 ", ""), r[1], r[2], f"R$ {r[3]:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), r[4][:40], r[5][:40]]
                pdf_data.append(row)

            t = Table(pdf_data)
            t.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkslategray),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('FONTSIZE', (0, 0), (-1, -1), 10)
            ]))
            doc.build([t])
            self.show_success_popup("PDF exportado com sucesso!")

    def clear_data(self):
        for w in self.data_inner_frame.winfo_children(): w.destroy()
        self.parsed_rows = []
        self.row_index = 0
        self.total_geral = 0.0
        self.file_count = 0
        self.card_total.configure(text="R$ 0,00")
        self.card_files.configure(text="0")
        self.update_scroll_regions()


if __name__ == "__main__":
    app = XMLIntelligenceUltra()
    app.mainloop()