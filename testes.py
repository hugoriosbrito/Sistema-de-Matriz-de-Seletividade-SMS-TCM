import folium
import pandas as pd
import tkinter as tk
from tkinterweb import HtmlFrame
import tkhtmlview
from tkhtmlview import HTMLText, RenderHTML, HTMLLabel
import webview
import os

# Carregar dados do Excel
df = pd.read_excel("dados\\Matriz Modelo - VERSÃO SISTEMA.xlsx", sheet_name='MATRIZ CONTRATOS')

dfIDs = df.iloc[6:,0]
dfMunicipio= df.iloc[6:,1]
dfNota = df.iloc[6:,34]
dfIRCE = df.iloc[6:,2]
dfDCE=df.iloc[6:,3]

novo_df = {
    'id':dfIDs.values,
    'municipio':dfMunicipio.values,
    'irce':dfIRCE.values,
    'dce':dfDCE.values,
    'nota':dfNota.values
}
dfPlot = pd.DataFrame(novo_df)

# URL do GeoJSON
geojson_url = 'https://raw.githubusercontent.com/tbrugz/geodata-br/refs/heads/master/geojson/geojs-29-mun.json'

# Criar o mapa
mapa_mun_bahia = folium.Map(location=[-12.9704, -38.5124], zoom_start=6, tiles='cartodbpositron')

# Criar o choropleth
folium.Choropleth(
    geo_data=geojson_url,
    data=dfPlot,
    columns=['id', 'nota'],
    key_on='feature.properties.id',
    fill_color='YlOrRd',
    fill_opacity=0.9,
    line_opacity=0.5,
    legend_name="Notas"
).add_to(mapa_mun_bahia)

# Salvar o mapa em um arquivo HTML
map_file = 'mapa_cloropleto_bahia.html'
mapa_mun_bahia.save(map_file)

# Criar a interface Tkinter
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Mapa Cloroplético da Bahia")
        self.geometry("800x600")

        # Frame para o mapa
        self.frame_map = tk.Frame(self)
        self.frame_map.pack(fill='both', expand=True)

        # Carregar e exibir o mapa
        self.show_map()

    def show_map(self):
        html_label = HTMLLabel(self.frame_map, html=RenderHTML(map_file))
        html_label.pack(fill="both", expand=True)
        html_label.fit_height()

if __name__ == "__main__":
    app = App()
    app.mainloop()

    # Limpar o arquivo do mapa após o fechamento
    if os.path.exists(map_file):
        os.remove(map_file)
