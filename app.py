#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from shiny import App, ui, render, reactive
from shinywidgets import output_widget, render_plotly

# ============================================================
# CONFIGURACIÓN GLOBAL
# ============================================================

ARCHIVO_EXCEL = 'Base_PIB (1).xlsx'
años_energia  = ['2014','2015','2016','2017','2018','2019','2020','2021','2022','2023','2024']
años_pib      = [str(y) for y in range(2010, 2025)]

# ── Mapeos de continentes (Dashboard 2) ─────────────────────
CONTINENTE_MAP_D2 = {
    'Canada':'América','Mexico':'América','US':'América','Argentina':'América',
    'Brazil':'América','Chile':'América','Colombia':'América','Ecuador':'América',
    'Peru':'América','Venezuela':'América','Other South America':'América',
    'Trinidad & Tobago':'América','Central America':'América','Other Caribbean':'América',
    'Honduras':'América','The Dominican Republic':'América','Cuba':'América',
    'Puerto Rico & USVI':'América','Other S. & Cent. America':'América',
    'Austria':'Europa','Belgium':'Europa','Bulgaria':'Europa','Croatia':'Europa',
    'Cyprus':'Europa','Czech Republic':'Europa','Denmark':'Europa','Estonia':'Europa',
    'Finland':'Europa','France':'Europa','Germany':'Europa','Greece':'Europa',
    'Hungary':'Europa','Iceland':'Europa','Ireland':'Europa','Italy':'Europa',
    'Latvia':'Europa','Lithuania':'Europa','Luxembourg':'Europa','Malta':'Europa',
    'Netherlands':'Europa','North Macedonia':'Europa','Norway':'Europa','Poland':'Europa',
    'Portugal':'Europa','Romania':'Europa','Slovakia':'Europa','Slovenia':'Europa',
    'Spain':'Europa','Sweden':'Europa','Switzerland':'Europa','Türkiye':'Europa',
    'Ukraine':'Europa','United Kingdom':'Europa','Other Europe':'Europa',
    'Bosnia Herzegovina':'Europa','Albania':'Europa','Montenegro':'Europa',
    'Serbia':'Europa','Russian Federation':'Europa','Belarus':'Europa',
    'Azerbaijan':'Asia','Kazakhstan':'Asia','Turkmenistan':'Asia','Uzbekistan':'Asia',
    'Other CIS':'Asia','Iran':'Asia','Iraq':'Asia','Israel':'Asia','Jordan':'Asia',
    'Kuwait':'Asia','Oman':'Asia','Qatar':'Asia','Saudi Arabia':'Asia','Syria':'Asia',
    'United Arab Emirates':'Asia','Yemen':'Asia','Other Middle East':'Asia','Bahrain':'Asia',
    'Bangladesh':'Asia','China':'Asia','China Hong Kong SAR':'Asia','India':'Asia',
    'Indonesia':'Asia','Japan':'Asia','Malaysia':'Asia','Myanmar':'Asia','Pakistan':'Asia',
    'Philippines':'Asia','Singapore':'Asia','South Korea':'Asia','Sri Lanka':'Asia',
    'Taiwan':'Asia','Thailand':'Asia','Viet Nam':'Asia','Vietnam':'Asia',
    'Other Asia Pacific':'Asia','Cambodia':'Asia','Mongolia':'Asia',
    'Algeria':'África','Egypt':'África','Morocco':'África','South Africa':'África',
    'Tunisia':'África','Eastern Africa':'África','Middle Africa':'África',
    'Western Africa':'África','Other Northern Africa':'África','Other Southern Africa':'África',
    'Other Africa':'África','Nigeria':'África','Kenya':'África','Libya':'África',
    'Australia':'Oceanía','New Zealand':'Oceanía','Papua New Guinea':'Oceanía',
    'Fiji':'Oceanía','New Caledonia':'Oceanía',
}

ISO_CONTINENTE_D2 = {
    'USA':'América','CAN':'América','MEX':'América','GRL':'América','GTM':'América',
    'BLZ':'América','HND':'América','SLV':'América','NIC':'América','CRI':'América',
    'PAN':'América','CUB':'América','DOM':'América','HTI':'América','JAM':'América',
    'TTO':'América','BRB':'América','LCA':'América','VCT':'América','GRD':'América',
    'ATG':'América','KNA':'América','PRI':'América','BRA':'América','ARG':'América',
    'CHL':'América','COL':'América','PER':'América','VEN':'América','ECU':'América',
    'BOL':'América','PRY':'América','URY':'América','GUY':'América','SUR':'América',
    'DEU':'Europa','FRA':'Europa','GBR':'Europa','ITA':'Europa','ESP':'Europa','PRT':'Europa',
    'NLD':'Europa','BEL':'Europa','LUX':'Europa','CHE':'Europa','AUT':'Europa','SWE':'Europa',
    'NOR':'Europa','DNK':'Europa','FIN':'Europa','ISL':'Europa','IRL':'Europa','POL':'Europa',
    'CZE':'Europa','SVK':'Europa','HUN':'Europa','ROU':'Europa','BGR':'Europa','HRV':'Europa',
    'SVN':'Europa','EST':'Europa','LVA':'Europa','LTU':'Europa','GRC':'Europa','CYP':'Europa',
    'MLT':'Europa','ALB':'Europa','MNE':'Europa','SRB':'Europa','BIH':'Europa','MKD':'Europa',
    'UKR':'Europa','BLR':'Europa','RUS':'Europa','TUR':'Europa','GEO':'Europa','ARM':'Europa',
    'MDA':'Europa','XKX':'Europa',
    'KAZ':'Asia','UZB':'Asia','TKM':'Asia','KGZ':'Asia','TJK':'Asia','AZE':'Asia',
    'SAU':'Asia','IRN':'Asia','IRQ':'Asia','ARE':'Asia','KWT':'Asia','QAT':'Asia',
    'BHR':'Asia','OMN':'Asia','YEM':'Asia','JOR':'Asia','SYR':'Asia','LBN':'Asia',
    'ISR':'Asia','PSE':'Asia','CHN':'Asia','IND':'Asia','JPN':'Asia','KOR':'Asia',
    'IDN':'Asia','MYS':'Asia','THA':'Asia','VNM':'Asia','PHL':'Asia','SGP':'Asia',
    'MMR':'Asia','BGD':'Asia','PAK':'Asia','LKA':'Asia','NPL':'Asia','KHM':'Asia',
    'LAO':'Asia','MNG':'Asia','HKG':'Asia','TLS':'Asia','BRN':'Asia','AFG':'Asia',
    'BTN':'Asia','MDV':'Asia',
    'DZA':'África','EGY':'África','MAR':'África','TUN':'África','LBY':'África','SDN':'África',
    'ETH':'África','KEN':'África','TZA':'África','UGA':'África','ZAF':'África','NGA':'África',
    'GHA':'África','CIV':'África','SEN':'África','CMR':'África','AGO':'África','MOZ':'África',
    'ZMB':'África','ZWE':'África','MDG':'África','MLI':'África','NER':'África','TCD':'África',
    'SOM':'África','COD':'África','COG':'África','GAB':'África','BFA':'África','GIN':'África',
    'RWA':'África','BDI':'África','BWA':'África','NAM':'África','LSO':'África','SWZ':'África',
    'MRT':'África','GMB':'África','GNB':'África','SLE':'África','LBR':'África','TGO':'África',
    'BEN':'África','ERI':'África','DJI':'África','COM':'África','CPV':'África','STP':'África',
    'MUS':'África','SYC':'África','GNQ':'África','CAF':'África','MWI':'África',
    'AUS':'Oceanía','NZL':'Oceanía','FJI':'Oceanía','PNG':'Oceanía','WSM':'Oceanía',
    'TON':'Oceanía','VUT':'Oceanía','SLB':'Oceanía','FSM':'Oceanía','PLW':'Oceanía',
    'MHL':'Oceanía','KIR':'Oceanía','NRU':'Oceanía','TUV':'Oceanía',
}

# ── Mapeos Dashboard 3 ───────────────────────────────────────
CONTINENTE_MAP_D3 = {
    'Canada':'América del Norte','Mexico':'América del Norte','US':'América del Norte',
    'Argentina':'América del Sur','Brazil':'América del Sur','Chile':'América del Sur',
    'Colombia':'América del Sur','Ecuador':'América del Sur','Peru':'América del Sur',
    'Venezuela':'América del Sur','Other South America':'América del Sur',
    'Trinidad & Tobago':'América del Sur',
    'Central America':'América Central','Other Caribbean':'América Central',
    'Austria':'Europa','Belgium':'Europa','Bulgaria':'Europa','Croatia':'Europa',
    'Cyprus':'Europa','Czech Republic':'Europa','Denmark':'Europa','Estonia':'Europa',
    'Finland':'Europa','France':'Europa','Germany':'Europa','Greece':'Europa',
    'Hungary':'Europa','Iceland':'Europa','Ireland':'Europa','Italy':'Europa',
    'Latvia':'Europa','Lithuania':'Europa','Luxembourg':'Europa','Malta':'Europa',
    'Netherlands':'Europa','North Macedonia':'Europa','Norway':'Europa','Poland':'Europa',
    'Portugal':'Europa','Romania':'Europa','Slovakia':'Europa','Slovenia':'Europa',
    'Spain':'Europa','Sweden':'Europa','Switzerland':'Europa','Türkiye':'Europa',
    'Ukraine':'Europa','United Kingdom':'Europa','Other Europe':'Europa',
    'Bosnia Herzegovina':'Europa','Albania':'Europa','Montenegro':'Europa',
    'Serbia':'Europa','Russian Federation':'Europa','Belarus':'Europa',
    'Azerbaijan':'Asia Central','Kazakhstan':'Asia Central','Turkmenistan':'Asia Central',
    'Uzbekistan':'Asia Central','Other CIS':'Asia Central',
    'Iran':'Medio Oriente','Iraq':'Medio Oriente','Israel':'Medio Oriente',
    'Jordan':'Medio Oriente','Kuwait':'Medio Oriente','Oman':'Medio Oriente',
    'Qatar':'Medio Oriente','Saudi Arabia':'Medio Oriente','Syria':'Medio Oriente',
    'United Arab Emirates':'Medio Oriente','Yemen':'Medio Oriente',
    'Other Middle East':'Medio Oriente','Bahrain':'Medio Oriente',
    'Algeria':'África','Egypt':'África','Morocco':'África','South Africa':'África',
    'Tunisia':'África','Eastern Africa':'África','Middle Africa':'África',
    'Western Africa':'África','Other Northern Africa':'África',
    'Other Southern Africa':'África','Other Africa':'África',
    'Nigeria':'África','Kenya':'África','Libya':'África',
    'Australia':'Asia-Pacífico','Bangladesh':'Asia-Pacífico','China':'Asia-Pacífico',
    'China Hong Kong SAR':'Asia-Pacífico','India':'Asia-Pacífico','Indonesia':'Asia-Pacífico',
    'Japan':'Asia-Pacífico','Malaysia':'Asia-Pacífico','Myanmar':'Asia-Pacífico',
    'New Zealand':'Asia-Pacífico','Pakistan':'Asia-Pacífico','Philippines':'Asia-Pacífico',
    'Singapore':'Asia-Pacífico','South Korea':'Asia-Pacífico','Sri Lanka':'Asia-Pacífico',
    'Taiwan':'Asia-Pacífico','Thailand':'Asia-Pacífico','Viet Nam':'Asia-Pacífico',
    'Vietnam':'Asia-Pacífico','Other Asia Pacific':'Asia-Pacífico',
}

ISO_CONTINENTE_D3 = {
    'USA':'América del Norte','CAN':'América del Norte','MEX':'América del Norte','GRL':'América del Norte',
    'GTM':'América Central','BLZ':'América Central','HND':'América Central','SLV':'América Central',
    'NIC':'América Central','CRI':'América Central','PAN':'América Central','CUB':'América Central',
    'DOM':'América Central','HTI':'América Central','JAM':'América Central','TTO':'América Central',
    'BRB':'América Central','LCA':'América Central','VCT':'América Central','GRD':'América Central',
    'ATG':'América Central','KNA':'América Central','PRI':'América Central',
    'BRA':'América del Sur','ARG':'América del Sur','CHL':'América del Sur','COL':'América del Sur',
    'PER':'América del Sur','VEN':'América del Sur','ECU':'América del Sur','BOL':'América del Sur',
    'PRY':'América del Sur','URY':'América del Sur','GUY':'América del Sur','SUR':'América del Sur',
    'DEU':'Europa','FRA':'Europa','GBR':'Europa','ITA':'Europa','ESP':'Europa','PRT':'Europa',
    'NLD':'Europa','BEL':'Europa','LUX':'Europa','CHE':'Europa','AUT':'Europa','SWE':'Europa',
    'NOR':'Europa','DNK':'Europa','FIN':'Europa','ISL':'Europa','IRL':'Europa','POL':'Europa',
    'CZE':'Europa','SVK':'Europa','HUN':'Europa','ROU':'Europa','BGR':'Europa','HRV':'Europa',
    'SVN':'Europa','EST':'Europa','LVA':'Europa','LTU':'Europa','GRC':'Europa','CYP':'Europa',
    'MLT':'Europa','ALB':'Europa','MNE':'Europa','SRB':'Europa','BIH':'Europa','MKD':'Europa',
    'UKR':'Europa','BLR':'Europa','RUS':'Europa','TUR':'Europa','GEO':'Europa','ARM':'Europa',
    'MDA':'Europa','XKX':'Europa',
    'KAZ':'Asia Central','UZB':'Asia Central','TKM':'Asia Central','KGZ':'Asia Central',
    'TJK':'Asia Central','AZE':'Asia Central',
    'SAU':'Medio Oriente','IRN':'Medio Oriente','IRQ':'Medio Oriente','ARE':'Medio Oriente',
    'KWT':'Medio Oriente','QAT':'Medio Oriente','BHR':'Medio Oriente','OMN':'Medio Oriente',
    'YEM':'Medio Oriente','JOR':'Medio Oriente','SYR':'Medio Oriente','LBN':'Medio Oriente',
    'ISR':'Medio Oriente','PSE':'Medio Oriente',
    'DZA':'África','EGY':'África','MAR':'África','TUN':'África','LBY':'África','SDN':'África',
    'ETH':'África','KEN':'África','TZA':'África','UGA':'África','ZAF':'África','NGA':'África',
    'GHA':'África','CIV':'África','SEN':'África','CMR':'África','AGO':'África','MOZ':'África',
    'ZMB':'África','ZWE':'África','MDG':'África','MLI':'África','NER':'África','TCD':'África',
    'SOM':'África','COD':'África','COG':'África','GAB':'África','BFA':'África','GIN':'África',
    'RWA':'África','BDI':'África','BWA':'África','NAM':'África','LSO':'África','SWZ':'África',
    'MRT':'África','GMB':'África','GNB':'África','SLE':'África','LBR':'África','TGO':'África',
    'BEN':'África','ERI':'África','DJI':'África','COM':'África','CPV':'África','STP':'África',
    'MUS':'África','SYC':'África','GNQ':'África','CAF':'África','MWI':'África',
    'CHN':'Asia-Pacífico','IND':'Asia-Pacífico','JPN':'Asia-Pacífico','KOR':'Asia-Pacífico',
    'AUS':'Asia-Pacífico','NZL':'Asia-Pacífico','IDN':'Asia-Pacífico','MYS':'Asia-Pacífico',
    'THA':'Asia-Pacífico','VNM':'Asia-Pacífico','PHL':'Asia-Pacífico','SGP':'Asia-Pacífico',
    'MMR':'Asia-Pacífico','BGD':'Asia-Pacífico','PAK':'Asia-Pacífico','LKA':'Asia-Pacífico',
    'NPL':'Asia-Pacífico','KHM':'Asia-Pacífico','LAO':'Asia-Pacífico','MNG':'Asia-Pacífico',
    'HKG':'Asia-Pacífico','TLS':'Asia-Pacífico','BRN':'Asia-Pacífico','FJI':'Asia-Pacífico',
    'PNG':'Asia-Pacífico','WSM':'Asia-Pacífico','TON':'Asia-Pacífico','VUT':'Asia-Pacífico',
    'SLB':'Asia-Pacífico','FSM':'Asia-Pacífico','PLW':'Asia-Pacífico','MHL':'Asia-Pacífico',
    'KIR':'Asia-Pacífico','NRU':'Asia-Pacífico','TUV':'Asia-Pacífico','AFG':'Asia-Pacífico',
    'BTN':'Asia-Pacífico','MDV':'Asia-Pacífico',
}

BLOQUES_D3 = {
    'Oceanía': {
        'iso': ['AUS','NZL','PNG','FJI','WSM','TON','VUT','SLB','FSM','PLW','MHL','KIR','NRU','TUV'],
        'continentes': ['Asia-Pacífico'],
        'lataxis': [-55, 10], 'lonaxis': [110, 185],
    },
    'Asia': {
        'iso': ['CHN','IND','JPN','KOR','IDN','MYS','THA','VNM','PHL','SGP','MMR','BGD',
                'PAK','LKA','NPL','KHM','LAO','MNG','TWN','HKG','PRK','TLS','BRN',
                'KAZ','UZB','TKM','KGZ','TJK','AZE','GEO','ARM',
                'IRN','IRQ','ISR','SAU','ARE','KWT','QAT','BHR','OMN','YEM','JOR','SYR','LBN','PSE'],
        'continentes': ['Asia-Pacífico','Asia Central','Medio Oriente'],
        'lataxis': [-15, 60], 'lonaxis': [25, 150],
    },
    'América': {
        'iso': ['USA','CAN','MEX','GRL','GTM','BLZ','HND','SLV','NIC','CRI','PAN','CUB',
                'DOM','HTI','JAM','TTO','BRB','LCA','VCT','GRD','ATG','KNA',
                'BRA','ARG','CHL','COL','PER','VEN','ECU','BOL','PRY','URY','GUY','SUR'],
        'continentes': ['América del Norte','América Central','América del Sur'],
        'lataxis': [-60, 85], 'lonaxis': [-170, -30],
    },
    'Europa': {
        'iso': ['DEU','FRA','GBR','ITA','ESP','PRT','NLD','BEL','LUX','CHE','AUT',
                'SWE','NOR','DNK','FIN','ISL','IRL','POL','CZE','SVK','HUN','ROU',
                'BGR','HRV','SVN','EST','LVA','LTU','GRC','CYP','MLT','ALB','MNE',
                'SRB','BIH','MKD','UKR','BLR','RUS','TUR'],
        'continentes': ['Europa'],
        'lataxis': [35, 72], 'lonaxis': [-25, 65],
    },
    'África': {
        'iso': ['DZA','EGY','MAR','TUN','LBY','SDN','ETH','KEN','TZA','UGA','ZAF',
                'NGA','GHA','CIV','SEN','CMR','AGO','MOZ','ZMB','ZWE','MDG','MLI',
                'NER','TCD','SOM','COD','COG','GAB','BFA','GIN','RWA','BDI','BWA',
                'NAM','LSO','SWZ','MRT','GMB','GNB','SLE','LBR','TGO','BEN','ERI',
                'DJI','COM','CPV','STP','MUS'],
        'continentes': ['África'],
        'lataxis': [-40, 38], 'lonaxis': [-20, 55],
    },
}

ESCALAS_D3 = {
    'PIB':        {'color': [[0,'#EBF5FB'],[0.5,'#5DADE2'],[1,'#1A5276']], 'unidad':'%'},
    'Petróleo':   {'color': [[0,'#FDEDEC'],[0.5,'#E74C3C'],[1,'#7B241C']], 'unidad':'Exajoules'},
    'Renovables': {'color': [[0,'#E9F7EF'],[0.5,'#52BE80'],[1,'#145A32']], 'unidad':'MW/TWh'},
}
COL_MAP_D3 = {'PIB':'PIB','Petróleo':'Petroleo','Renovables':'Renovables'}
MEDALLAS   = ['🥇','🥈','🥉','4️⃣','5️⃣','6️⃣','7️⃣','8️⃣']

# ============================================================
# CARGA DE DATOS
# ============================================================
print("Cargando datos...")

with pd.ExcelFile(ARCHIVO_EXCEL) as xls:
    df_solar_raw    = pd.read_excel(xls, sheet_name='Dataset_solar')
    df_hidro_raw    = pd.read_excel(xls, sheet_name='Dataset_hidroelectrica')
    df_eolica_raw   = pd.read_excel(xls, sheet_name='Dataset_eolica')
    df_petroleo_raw = pd.read_excel(xls, sheet_name='Dataset_petroleo')
    df_pib_raw      = pd.read_excel(xls, sheet_name='Dataset_Pib', skiprows=4)

# ── Dashboard 1: datos por país ──────────────────────────────
def prep_d1(df, col_pais, tipo):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns={col_pais: 'Country'}, inplace=True)
    df[años_energia] = df[años_energia].apply(pd.to_numeric, errors='coerce')
    df['Promedio'] = df[años_energia].mean(axis=1)
    df['Tipo'] = tipo
    return df[['Country','Promedio','Tipo']].dropna(subset=['Promedio'])

d1_solar    = prep_d1(df_solar_raw,    'Megawatts',      'Solar')
d1_hidro    = prep_d1(df_hidro_raw,    'Terawatt-hours', 'Hidroeléctrica')
d1_eolica   = prep_d1(df_eolica_raw,   'Megawatts',      'Eólica')
d1_petroleo = prep_d1(df_petroleo_raw, 'Exajoules',      'Petróleo')
df_energia_d1 = pd.concat([d1_solar, d1_hidro, d1_eolica, d1_petroleo], ignore_index=True)
df_energia_d1 = df_energia_d1.sort_values('Country')

df_pib_raw.columns = [str(c).strip() for c in df_pib_raw.columns]
cols_pib = [c for c in años_pib if c in df_pib_raw.columns]
df_pib_raw[cols_pib] = df_pib_raw[cols_pib].apply(pd.to_numeric, errors='coerce')
df_pib_long = df_pib_raw.melt(
    id_vars=['Country Name'], value_vars=cols_pib, var_name='Año', value_name='PIB'
).rename(columns={'Country Name':'Country'})
df_pib_long['Año'] = pd.to_numeric(df_pib_long['Año'])
lista_paises = sorted(df_pib_long['Country'].dropna().unique().tolist())

# ── Dashboard 2: series temporales por continente ───────────
def proc_energia_d2(df, col_pais, tipo):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns={col_pais: 'Country'}, inplace=True)
    cols = [c for c in años_energia if c in df.columns]
    df[cols] = df[cols].apply(pd.to_numeric, errors='coerce')
    df_l = df.melt(id_vars=['Country'], value_vars=cols, var_name='Año', value_name='Valor')
    df_l['Año']        = pd.to_numeric(df_l['Año'])
    df_l['Tipo']       = tipo
    df_l['Continente'] = df_l['Country'].map(CONTINENTE_MAP_D2).fillna('Otros')
    return df_l[df_l['Continente'] != 'Otros'].dropna(subset=['Valor'])

d2_solar    = proc_energia_d2(df_solar_raw,    'Megawatts',      'Solar')
d2_hidro    = proc_energia_d2(df_hidro_raw,    'Terawatt-hours', 'Hidroeléctrica')
d2_eolica   = proc_energia_d2(df_eolica_raw,   'Megawatts',      'Eólica')
d2_petroleo = proc_energia_d2(df_petroleo_raw, 'Exajoules',      'Petróleo')
d2_renovables = pd.concat([d2_solar, d2_hidro, d2_eolica], ignore_index=True)

df_pib_d2 = df_pib_raw.copy()
df_pib_d2['Continente'] = df_pib_d2['Country Code'].map(ISO_CONTINENTE_D2).fillna('Otros')
df_pib_d2_long = df_pib_d2.melt(
    id_vars=['Country Name','Country Code','Continente'],
    value_vars=[c for c in cols_pib if c in df_pib_d2.columns],
    var_name='Año', value_name='PIB'
)
df_pib_d2_long.rename(columns={'Country Name':'Country'}, inplace=True)
df_pib_d2_long['Año'] = pd.to_numeric(df_pib_d2_long['Año'])
df_pib_d2_long = df_pib_d2_long[df_pib_d2_long['Continente'] != 'Otros'].dropna(subset=['PIB'])

lista_continentes_d2 = ['Todos'] + sorted(
    set(d2_petroleo['Continente'].unique()) | set(d2_renovables['Continente'].unique())
)

# ── Dashboard 3: mapas de coropletas por continente ─────────
def agg_cont_d3(df, col_pais, años):
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df.rename(columns={col_pais:'Country'}, inplace=True)
    df[años] = df[años].apply(pd.to_numeric, errors='coerce')
    df['Continente'] = df['Country'].map(CONTINENTE_MAP_D3).fillna('Otros')
    df['Promedio']   = df[años].mean(axis=1)
    return df.groupby('Continente')['Promedio'].sum().reset_index()

r_pet_d3 = agg_cont_d3(df_petroleo_raw,'Exajoules',      años_energia).rename(columns={'Promedio':'Petroleo'})
r_sol_d3 = agg_cont_d3(df_solar_raw,   'Megawatts',      años_energia)
r_hid_d3 = agg_cont_d3(df_hidro_raw,   'Terawatt-hours', años_energia)
r_eol_d3 = agg_cont_d3(df_eolica_raw,  'Megawatts',      años_energia)
r_ren_d3 = (pd.concat([r_sol_d3, r_hid_d3, r_eol_d3])
              .groupby('Continente')['Promedio'].sum()
              .reset_index().rename(columns={'Promedio':'Renovables'}))

df_pib_d3 = df_pib_raw.copy()
df_pib_d3['Continente'] = df_pib_d3['Country Code'].map(ISO_CONTINENTE_D3)
df_pib_d3['PIB_prom']   = df_pib_d3[cols_pib].mean(axis=1)
r_pib_d3 = (df_pib_d3.dropna(subset=['Continente','PIB_prom'])
                      .groupby('Continente')['PIB_prom'].mean()
                      .reset_index().rename(columns={'PIB_prom':'PIB'}))

df_cont_d3 = (r_pet_d3.merge(r_ren_d3, on='Continente', how='outer')
                       .merge(r_pib_d3, on='Continente', how='outer'))
df_cont_d3 = df_cont_d3[df_cont_d3['Continente'] != 'Otros'].reset_index(drop=True)

print("¡Datos listos!")

# ============================================================
# FUNCIONES DE AGREGACIÓN D2
# ============================================================

def agg_energia_d2(df, continente):
    if continente != 'Todos':
        df = df[df['Continente'] == continente]
    return df.groupby('Año')['Valor'].sum().reset_index()

def agg_pib_d2(continente):
    df = df_pib_d2_long.copy()
    if continente != 'Todos':
        df = df[df['Continente'] == continente]
    return df.groupby('Año')['PIB'].mean().reset_index()

def base_layout_d2(titulo, x_title):
    return dict(
        title=titulo, plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF',
        font=dict(family='Arial', size=12),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='right', x=1),
        margin=dict(l=70, r=70, t=60, b=60),
        xaxis=dict(title=x_title, showgrid=True, gridcolor='#E0E0E0',
                   zerolinecolor='#E0E0E0', dtick=1),
    )

COLOR_PET = '#d35400'
COLOR_PIB = '#2980b9'
COLOR_REN = '#27ae60'

# ============================================================
# FUNCIÓN MAPA D3
# ============================================================

GEO_DOMAINS_D3 = {
    'Oceanía': dict(x=[0.00, 0.30], y=[0.52, 1.00]),
    'Asia':    dict(x=[0.00, 0.30], y=[0.00, 0.50]),
    'América': dict(x=[0.32, 0.65], y=[0.00, 1.00]),
    'Europa':  dict(x=[0.67, 1.00], y=[0.52, 1.00]),
    'África':  dict(x=[0.67, 1.00], y=[0.00, 0.50]),
}

def construir_mapa_d3(variable):
    col    = COL_MAP_D3[variable]
    escala = ESCALAS_D3[variable]
    unidad = escala['unidad']
    val_dict = dict(zip(df_cont_d3['Continente'], df_cont_d3[col]))
    vmin = df_cont_d3[col].min()
    vmax = df_cont_d3[col].max()

    fig = go.Figure()
    geo_configs   = {}
    showscale_done = False

    for i, (nombre, info) in enumerate(BLOQUES_D3.items()):
        geo_id = f'geo{i+1}' if i > 0 else 'geo'
        vals   = [val_dict.get(c) for c in info['continentes'] if val_dict.get(c) is not None]
        valor  = sum(vals) if vals else 0
        dom    = GEO_DOMAINS_D3[nombre]

        fig.add_trace(go.Choropleth(
            locations=info['iso'], z=[valor]*len(info['iso']),
            geo=geo_id, colorscale=escala['color'], zmin=vmin, zmax=vmax,
            showscale=not showscale_done,
            colorbar=dict(title=unidad, x=0.5, y=-0.08, orientation='h',
                          len=0.6, thickness=12) if not showscale_done else {},
            marker_line_color='white', marker_line_width=0.8,
            hovertemplate=f'<b>{nombre}</b><br>{variable}: {valor:,.2f} {unidad}<extra></extra>',
        ))
        showscale_done = True

        geo_configs[geo_id] = dict(
            domain=dom, showframe=True, framecolor='#BDC3C7',
            showcoastlines=True, coastlinecolor='#BDC3C7',
            showland=True, landcolor='#F2F3F4',
            showocean=True, oceancolor='#D6EAF8',
            showlakes=False, projection_type='natural earth',
            lataxis_range=info['lataxis'], lonaxis_range=info['lonaxis'],
            bgcolor='white',
        )

    anotaciones = [
        dict(x=0.15,  y=1.01, text='<b>Oceanía</b>',  showarrow=False, xref='paper', yref='paper', font=dict(size=11, color='#2c3e50')),
        dict(x=0.15,  y=0.49, text='<b>Asia</b>',      showarrow=False, xref='paper', yref='paper', font=dict(size=11, color='#2c3e50')),
        dict(x=0.485, y=1.01, text='<b>América</b>',   showarrow=False, xref='paper', yref='paper', font=dict(size=11, color='#2c3e50')),
        dict(x=0.835, y=1.01, text='<b>Europa</b>',    showarrow=False, xref='paper', yref='paper', font=dict(size=11, color='#2c3e50')),
        dict(x=0.835, y=0.49, text='<b>África</b>',    showarrow=False, xref='paper', yref='paper', font=dict(size=11, color='#2c3e50')),
    ]

    layout_update = dict(
        paper_bgcolor='white', plot_bgcolor='white', height=620,
        margin=dict(l=5, r=5, t=55, b=60),
        title=dict(
            text=f'Desarrollo por Continente — {variable} ({unidad})',
            x=0.5, xanchor='center', font=dict(size=15, color='#2c3e50')
        ),
        annotations=anotaciones,
    )
    layout_update.update(geo_configs)
    fig.update_layout(**layout_update)
    return fig

# ============================================================
# UI UNIFICADA
# ============================================================

CSS_GLOBAL = """
    @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;700;800&family=DM+Sans:wght@300;400;500&display=swap');

    * { box-sizing: border-box; }

    body {
        background-color: #f5f7fa !important;
        font-family: 'DM Sans', sans-serif;
        color: #1e2a3a;
        margin: 0;
        padding: 0;
    }

    /* ── HERO ── */
    .hero {
        background: linear-gradient(135deg, #1565C0 0%, #00897B 60%, #1976D2 100%);
        padding: 72px 40px 56px;
        text-align: center;
        position: relative;
        overflow: hidden;
        border-bottom: 1px solid rgba(0,0,0,0.08);
    }
    .hero::before {
        content: '';
        position: absolute; inset: 0;
        background: radial-gradient(ellipse 80% 60% at 50% 0%, rgba(255,255,255,0.12) 0%, transparent 70%);
        pointer-events: none;
    }
    .hero-label {
        font-family: 'DM Sans', sans-serif;
        font-size: 0.72rem;
        font-weight: 500;
        letter-spacing: 0.22em;
        text-transform: uppercase;
        color: #b2dfdb;
        margin-bottom: 16px;
    }
    .hero-title {
        font-family: 'Syne', sans-serif;
        font-size: clamp(1.6rem, 3.5vw, 2.8rem);
        font-weight: 800;
        line-height: 1.15;
        color: #ffffff;
        margin: 0 auto 20px;
        max-width: 780px;
    }
    .hero-title span { color: #b2ebf2; }
    .hero-subtitle {
        font-size: 0.97rem;
        color: rgba(255,255,255,0.78);
        max-width: 560px;
        margin: 0 auto;
        line-height: 1.6;
    }

    /* ── NAV TABS ── */
    .nav-section {
        background: #ffffff;
        padding: 28px 40px;
        border-bottom: 1px solid #e0e6ef;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    }
    .nav-label {
        font-size: 0.68rem;
        letter-spacing: 0.18em;
        text-transform: uppercase;
        color: #90a0b7;
        margin-bottom: 14px;
        font-weight: 500;
    }
    .nav-cards {
        display: flex;
        gap: 14px;
        flex-wrap: wrap;
    }
    .nav-card {
        flex: 1;
        min-width: 200px;
        background: #f8fafc;
        border: 1px solid #dde4ef;
        border-radius: 12px;
        padding: 20px 22px;
        cursor: pointer;
        transition: all 0.25s ease;
        position: relative;
        overflow: hidden;
    }
    .nav-card::before {
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0;
        height: 3px;
        border-radius: 12px 12px 0 0;
        transition: opacity 0.25s;
    }
    .nav-card-1::before { background: #2196F3; }
    .nav-card-2::before { background: #00897B; }
    .nav-card-3::before { background: #f59e0b; }
    .nav-card.active { background: #ffffff; border-color: #b0c4de; box-shadow: 0 4px 16px rgba(0,0,0,0.08); }
    .nav-card:hover  { background: #ffffff; transform: translateY(-2px); box-shadow: 0 4px 16px rgba(0,0,0,0.08); }
    .nav-num {
        font-family: 'Syne', sans-serif;
        font-size: 0.75rem;
        font-weight: 700;
        letter-spacing: 0.1em;
        color: #90a0b7;
        margin-bottom: 6px;
    }
    .nav-card-title {
        font-family: 'Syne', sans-serif;
        font-size: 0.95rem;
        font-weight: 700;
        color: #1e2a3a;
        margin-bottom: 4px;
    }
    .nav-card-desc {
        font-size: 0.78rem;
        color: #6b7c93;
        line-height: 1.4;
    }

    /* ── CONTENT PANELS ── */
    .dashboard-panel {
        padding: 32px 36px;
        min-height: 70vh;
    }
    .panel-header {
        display: flex;
        align-items: center;
        gap: 12px;
        margin-bottom: 24px;
        padding-bottom: 16px;
        border-bottom: 1px solid #e0e6ef;
    }
    .panel-dot {
        width: 10px; height: 10px;
        border-radius: 50%;
        flex-shrink: 0;
    }
    .panel-title {
        font-family: 'Syne', sans-serif;
        font-size: 1.15rem;
        font-weight: 700;
        color: #1e2a3a;
    }
    .panel-subtitle {
        font-size: 0.82rem;
        color: #90a0b7;
        margin-left: auto;
    }

    /* ── CARDS ── */
    .card {
        background: #ffffff !important;
        border: 1px solid #e0e6ef !important;
        border-radius: 12px !important;
        margin-bottom: 20px !important;
        box-shadow: 0 2px 12px rgba(0,0,0,0.06) !important;
    }
    .card-header {
        background: #f5f7fa !important;
        border-bottom: 1px solid #e0e6ef !important;
        color: #3d5068 !important;
        font-size: 0.88rem;
        font-weight: 500;
        padding: 14px 20px !important;
    }

    /* ── SELECTS ── */
    .selectize-control .selectize-input {
        background: #ffffff !important;
        border: 1px solid #cdd6e0 !important;
        color: #1e2a3a !important;
        border-radius: 8px !important;
        font-size: 0.88rem;
    }
    .selectize-control .selectize-dropdown {
        background: #ffffff !important;
        border: 1px solid #cdd6e0 !important;
        color: #1e2a3a !important;
    }
    select, .form-select, .form-control {
        background: #ffffff !important;
        border: 1px solid #cdd6e0 !important;
        color: #1e2a3a !important;
        border-radius: 8px !important;
    }
    label { color: #4a5e72 !important; font-size: 0.84rem !important; }

    /* ── RANKING D3 ── */
    .rank-item { padding:7px 0; border-bottom:1px solid #eaeff5; font-size:.9rem; color:#3d5068; }
    .rank-1  { color:#e74c3c; font-weight:700; }
    .rank-2  { color:#e67e22; font-weight:600; }
    .rank-3  { color:#f39c12; font-weight:600; }
    .ranking-box { background: #f8fafc; border-radius:10px; padding:1rem 1.5rem; }

    /* ── MISC ── */
    h2 { font-family: 'Syne', sans-serif; color: #1e2a3a !important; }
    .text-muted { color: #90a0b7 !important; }
    p { color: #4a5e72; }
"""

app_ui = ui.page_fluid(
    ui.tags.style(CSS_GLOBAL),

    # ── HERO ──────────────────────────────────────────────────
    ui.div(
        ui.div("ANÁLISIS GLOBAL · ENERGÍA & ECONOMÍA", class_="hero-label"),
        ui.HTML('<h1 class="hero-title">El crecimiento económico está acompañado de una <span>transición hacia energías renovables</span></h1>'),
        ui.div(
            "Explora dashboards interactivos para analizar la relación entre desarrollo económico y sostenibilidad energética.",
            class_="hero-subtitle"
        ),
        class_="hero"
    ),

    # ── NAVEGACIÓN ────────────────────────────────────────────
    ui.div(
        ui.div("SELECCIONA UN DASHBOARD", class_="nav-label"),
        ui.div(
            ui.input_action_button(
                "btn1", "",
                onclick="selectTab('tab1')",
            ),
            ui.input_action_button(
                "btn2", "",
                onclick="selectTab('tab2')",
            ),
            ui.input_action_button(
                "btn3", "",
                onclick="selectTab('tab3')",
            ),
            style="display:none"
        ),
        ui.tags.div(
            ui.tags.div(
                ui.tags.div("01", class_="nav-num"),
                ui.tags.div("Crecimiento económico y energético", class_="nav-card-title"),
                ui.tags.div("PIB, consumo energético y producción renovable por país y período.", class_="nav-card-desc"),
                id="navcard1",
                class_="nav-card nav-card-1 active",
                onclick="Shiny.setInputValue('tab_activo', '1', {priority: 'event'})",
            ),
            ui.tags.div(
                ui.tags.div("02", class_="nav-num"),
                ui.tags.div("Comparativa de desarrollo", class_="nav-card-title"),
                ui.tags.div("Análisis comparativo entre continentes en indicadores clave.", class_="nav-card-desc"),
                id="navcard2",
                class_="nav-card nav-card-2",
                onclick="Shiny.setInputValue('tab_activo', '2', {priority: 'event'})",
            ),
            ui.tags.div(
                ui.tags.div("03", class_="nav-num"),
                ui.tags.div("Mapeo por desarrollo de continente", class_="nav-card-title"),
                ui.tags.div("Visualización geográfica del avance energético por continente.", class_="nav-card-desc"),
                id="navcard3",
                class_="nav-card nav-card-3",
                onclick="Shiny.setInputValue('tab_activo', '3', {priority: 'event'})",
            ),
            class_="nav-cards"
        ),
        class_="nav-section"
    ),

    # JS para resaltar tarjeta activa
    ui.tags.script("""
        $(document).on('shiny:inputchanged', function(e){
            if(e.name === 'tab_activo'){
                ['navcard1','navcard2','navcard3'].forEach(function(id){
                    document.getElementById(id).classList.remove('active');
                });
                document.getElementById('navcard'+e.value).classList.add('active');
            }
        });
    """),

    # ── PANEL CONTENIDO ───────────────────────────────────────
    ui.div(
        ui.output_ui("contenido_dashboard"),
        class_="dashboard-panel"
    ),
)

# ============================================================
# SERVIDOR
# ============================================================

def server(input, output, session):

    tab_actual = reactive.Value("1")

    @reactive.Effect
    @reactive.event(input.tab_activo)
    def _cambiar_tab():
        tab_actual.set(str(input.tab_activo()))

    # ── Render del panel principal ───────────────────────────
    @render.ui
    def contenido_dashboard():
        t = tab_actual()

        if t == "1":
            return ui.div(
                ui.div(
                    ui.tags.div(class_="panel-dot", style="background:#2196F3"),
                    ui.span("Crecimiento Económico y Energético", class_="panel-title"),
                    ui.span("Promedio 2014-2024 · Por País", class_="panel-subtitle"),
                    class_="panel-header"
                ),
                ui.card(
                    ui.card_header("Promedio de Consumo/Capacidad de Energía (2014-2024)"),
                    ui.row(
                        ui.column(4,
                            ui.input_select("energy_filter","Tipo de Energía:",
                                choices=["Petróleo","Eólica","Solar","Hidroeléctrica"],
                                selected="Petróleo")
                        )
                    ),
                    output_widget("grafico_energia")
                ),
                ui.card(
                    ui.card_header("Evolución del PIB"),
                    ui.row(
                        ui.column(4,
                            ui.input_select("country_filter","País:",
                                choices=lista_paises,
                                selected="Mexico" if "Mexico" in lista_paises else lista_paises[0])
                        )
                    ),
                    output_widget("grafico_pib")
                ),
            )

        elif t == "2":
            return ui.div(
                ui.div(
                    ui.tags.div(class_="panel-dot", style="background:#20b2aa"),
                    ui.span("Comparativa de Desarrollo", class_="panel-title"),
                    ui.span("Series temporales 2014-2024 · Por Continente", class_="panel-subtitle"),
                    class_="panel-header"
                ),
                ui.card(
                    ui.row(
                        ui.column(4,
                            ui.input_select("continente_d2","🌍 Filtrar por Continente:",
                                choices=lista_continentes_d2, selected="Todos")
                        ),
                        ui.column(8, ui.p("El filtro aplica simultáneamente a las tres gráficas.",
                                          class_="text-muted mt-2"))
                    )
                ),
                ui.card(ui.card_header("① Petróleo vs PIB — Crecimiento Anual"),
                        output_widget("d2_grafico1")),
                ui.card(ui.card_header("② PIB vs Energías Renovables — Crecimiento Anual"),
                        output_widget("d2_grafico2")),
                ui.card(ui.card_header("③ Petróleo vs Energías Renovables — Cantidad de Energía"),
                        output_widget("d2_grafico3")),
            )

        else:  # t == "3"
            return ui.div(
                ui.div(
                    ui.tags.div(class_="panel-dot", style="background:#f59e0b"),
                    ui.span("Mapeo por Desarrollo de Continente", class_="panel-title"),
                    ui.span("Coropletas · Promedio 2010-2024", class_="panel-subtitle"),
                    class_="panel-header"
                ),
                ui.card(
                    ui.row(
                        ui.column(4,
                            ui.input_select("variable_d3","🔍 Variable a visualizar:",
                                choices=["PIB","Petróleo","Renovables"], selected="Renovables")
                        ),
                        ui.column(8, ui.p("Tonos más oscuros = mayor valor.",
                                          class_="text-muted mt-2"))
                    )
                ),
                ui.card(output_widget("d3_mapa")),
                ui.card(
                    ui.card_header("🏆 Ranking de Continentes"),
                    ui.output_ui("d3_ranking")
                ),
            )

    # ── Dashboard 1 ─────────────────────────────────────────
    @render_plotly
    def grafico_energia():
        df_f = df_energia_d1[df_energia_d1['Tipo'] == input.energy_filter()]
        df_p = df_f.nlargest(30, 'Promedio')
        fig  = px.bar(
            df_p, x='Country', y='Promedio',
            title=f"Top 30 Países - Promedio de Energía {input.energy_filter()}",
            labels={'Country':'País','Promedio':'Energía Promedio'},
            color='Promedio', color_continuous_scale='Viridis'
        )
        fig.update_layout(
            xaxis_tickangle=-45, margin=dict(b=100),
            plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF'
        )
        return fig

    @render_plotly
    def grafico_pib():
        df_f = df_pib_long[df_pib_long['Country'] == input.country_filter()]
        fig  = px.line(
            df_f, x='Año', y='PIB',
            title=f"Crecimiento del PIB — {input.country_filter()} (2010-2024)",
            labels={'Año':'Año','PIB':'Crecimiento Anual del PIB (%)'},
            markers=True
        )
        fig.update_traces(line=dict(width=3, color='#1f77b4'), marker=dict(size=8))
        fig.update_layout(
            xaxis=dict(tickmode='linear', dtick=1),
            plot_bgcolor='#FFFFFF', paper_bgcolor='#FFFFFF'
        )
        return fig

    # ── Dashboard 2 ─────────────────────────────────────────
    @render_plotly
    def d2_grafico1():
        cont   = input.continente_d2()
        df_pet = agg_energia_d2(d2_petroleo, cont)
        df_pib = agg_pib_d2(cont)
        años_ok = sorted(set(df_pet['Año']) & set(df_pib['Año']))
        df_pet  = df_pet[df_pet['Año'].isin(años_ok)]
        df_pib  = df_pib[df_pib['Año'].isin(años_ok)]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_pet['Año'], y=df_pet['Valor'],
            name='Petróleo (Exajoules)', line=dict(color=COLOR_PET, width=3),
            marker=dict(size=7), mode='lines+markers'))
        fig.add_trace(go.Scatter(x=df_pib['Año'], y=df_pib['PIB'],
            name='PIB (% crecim. prom.)', line=dict(color=COLOR_PIB, width=3, dash='dot'),
            marker=dict(size=7), mode='lines+markers', yaxis='y2'))
        layout = base_layout_d2(f'Petróleo vs PIB — {cont}', 'Año')
        layout['yaxis']  = dict(title='Consumo Petróleo (Exajoules)',
                                showgrid=True, gridcolor='#E0E0E0', color=COLOR_PET)
        layout['yaxis2'] = dict(title='Crecimiento PIB (%)',
                                overlaying='y', side='right', color=COLOR_PIB, showgrid=False)
        fig.update_layout(**layout)
        return fig

    @render_plotly
    def d2_grafico2():
        cont   = input.continente_d2()
        df_ren = agg_energia_d2(d2_renovables, cont)
        df_pib = agg_pib_d2(cont)
        años_ok = sorted(set(df_ren['Año']) & set(df_pib['Año']))
        df_ren  = df_ren[df_ren['Año'].isin(años_ok)]
        df_pib  = df_pib[df_pib['Año'].isin(años_ok)]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_pib['Año'], y=df_pib['PIB'],
            name='PIB (% crecim. prom.)', line=dict(color=COLOR_PIB, width=3),
            marker=dict(size=7), mode='lines+markers'))
        fig.add_trace(go.Scatter(x=df_ren['Año'], y=df_ren['Valor'],
            name='Renovables (Solar+Eólica+Hidro)', line=dict(color=COLOR_REN, width=3, dash='dot'),
            marker=dict(size=7), mode='lines+markers', yaxis='y2'))
        layout = base_layout_d2(f'PIB vs Energías Renovables — {cont}', 'Año')
        layout['yaxis']  = dict(title='Crecimiento PIB (%)',
                                showgrid=True, gridcolor='#E0E0E0', color=COLOR_PIB)
        layout['yaxis2'] = dict(title='Capacidad Renovable (MW/TWh)',
                                overlaying='y', side='right', color=COLOR_REN, showgrid=False)
        fig.update_layout(**layout)
        return fig

    @render_plotly
    def d2_grafico3():
        cont   = input.continente_d2()
        df_pet = agg_energia_d2(d2_petroleo, cont)
        df_ren = agg_energia_d2(d2_renovables, cont)
        años_ok = sorted(set(df_pet['Año']) & set(df_ren['Año']))
        df_pet  = df_pet[df_pet['Año'].isin(años_ok)]
        df_ren  = df_ren[df_ren['Año'].isin(años_ok)]

        fig = go.Figure()
        fig.add_trace(go.Scatter(x=df_pet['Año'], y=df_pet['Valor'],
            name='Petróleo (Exajoules)', line=dict(color=COLOR_PET, width=3),
            marker=dict(size=7), mode='lines+markers'))
        fig.add_trace(go.Scatter(x=df_ren['Año'], y=df_ren['Valor'],
            name='Renovables (Solar+Eólica+Hidro)', line=dict(color=COLOR_REN, width=3, dash='dot'),
            marker=dict(size=7), mode='lines+markers', yaxis='y2'))
        layout = base_layout_d2(f'Petróleo vs Renovables — {cont}', 'Año')
        layout['yaxis']  = dict(title='Consumo Petróleo (Exajoules)',
                                showgrid=True, gridcolor='#E0E0E0', color=COLOR_PET)
        layout['yaxis2'] = dict(title='Capacidad Renovable (MW/TWh)',
                                overlaying='y', side='right', color=COLOR_REN, showgrid=False)
        fig.update_layout(**layout)
        return fig

    # ── Dashboard 3 ─────────────────────────────────────────
    @render_plotly
    def d3_mapa():
        return construir_mapa_d3(input.variable_d3())

    @render.ui
    def d3_ranking():
        col    = COL_MAP_D3[input.variable_d3()]
        unidad = ESCALAS_D3[input.variable_d3()]['unidad']
        df_r   = df_cont_d3[['Continente', col]].dropna().sort_values(col, ascending=False).reset_index(drop=True)
        items  = []
        for i, row in df_r.iterrows():
            medalla = MEDALLAS[i] if i < len(MEDALLAS) else ''
            clase   = f"rank-{i+1}" if i < 3 else ""
            items.append(ui.div(
                ui.span(f"{medalla}  {row['Continente']}  —  "),
                ui.span(f"{row[col]:,.2f} {unidad}"),
                class_=f"rank-item {clase}"
            ))
        return ui.div(*items, class_="ranking-box")


app = App(app_ui, server)



