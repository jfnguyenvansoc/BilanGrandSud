import datetime as dt
from collections import OrderedDict
from statistics import mean

import arcpy
import pandas as pd
import numpy as np
import scipy.stats as sp
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from numpy import polyfit
from pandas.plotting import register_matplotlib_converters

register_matplotlib_converters()

def get_del_field(fcIn, liste_champ):
    fields = arcpy.ListFields(fcIn)
    fields_del = []
    for f in fields:
        if f.name not in liste_champ:
            fields_del.append(f.name)
    return fields_del

gdb = r"C:\Users\come.daval\Documents\ArcGIS\Projects\Lea\ExportTHIO.gdb"
annee_fin_etude = 2019
fichier_excel_gamme_ref = "BilanThio2019.xlsx"

tb_parametre = "TB_PARAMETRES_THIO"
couche_stations = "FC_STATIONS_THIO"
tb_donnees = "TB_DONNEES_THIO"

annee_dbt_etude = annee_fin_etude - 4
dt_dbt_etude = dt.datetime(annee_dbt_etude, 1, 1)
dt_fin_etude = dt.datetime(annee_fin_etude, 12, 31)

tb_parametre_sorted = tb_parametre + "_sorted"
join_data = "join_data"

where_stations = "BilanENVRefSuiviTxt = 'SuiviTHIO'"
where_params = "BilanTHIO = 'Oui'"

field_data = ["OBJECTID", "id_parametre", "id_geometry", "date_", "valeur", "BilanEnv_id_parametre"]
field_parametres = ["OBJECTID", "id_parametre", "nom", "unite", "BilanTHIO", "OrdreThio"]
field_station = ["OBJECTID", "SHAPE", "nom", "id_geometry", "nom_simple", "BilanENVRefSuiviTxt", "BilanENVZone", "BilanENVOrdre", "BilanENVGraphique"]
field_del_data = get_del_field(gdb + "/" + tb_donnees, field_data)
field_del_param = get_del_field(gdb + "/" + tb_parametre, field_parametres)
field_del_station = get_del_field(gdb + "/" + couche_stations, field_station)
del_fields_l = [field_del_data, field_del_param, field_del_station]
del_fields = []
for table_name, liste_field in zip([tb_donnees, tb_parametre, couche_stations], del_fields_l):
    for f in liste_field:
        del_fields.append(table_name + "." + f)

arcpy.Delete_management(tb_parametre)
arcpy.Delete_management(tb_donnees)
arcpy.Delete_management(couche_stations)

arcpy.MakeTableView_management(gdb + "/" + couche_stations, couche_stations, where_stations)
arcpy.MakeTableView_management(gdb + "/" + tb_parametre, tb_parametre, where_params)
arcpy.MakeTableView_management(gdb + "/" + tb_donnees, tb_donnees)
# A adapter selon source
key_field_1 = tb_donnees + ".OBJECTID"
join_clause_1 = couche_stations + ".id_geometry = " + tb_donnees + ".id_geometry" +\
    " AND " + couche_stations + ".BilanENVRefSuiviTxt = 'SuiviTHIO'" +\
    " AND " + couche_stations + ".BilanENVGraphique = 'Oui'" +\
    " AND " + tb_parametre + ".id_parametre = " + tb_donnees + ".id_parametre" +\
    " AND " + tb_parametre + ".BilanTHIO = 'Oui'" +\
    " AND " + tb_donnees + ".date_ >= timestamp '" + str(annee_dbt_etude) + "-01-01'" +\
    " AND " + tb_donnees + ".date_ <= timestamp '" + str(annee_fin_etude) + "-12-31'"

# Jointure des données en conservants les data concernants les stations suivies
arcpy.Delete_management(join_data)
# Selection des données des 5 années d'études
arcpy.MakeQueryTable_management([tb_donnees, tb_parametre, couche_stations], join_data, "USE_KEY_FIELDS",\
    key_field_1, where_clause=join_clause_1)
# Tri des données de la table selon l'ordre des paramètre, puis des stations, puis des dates
fields_sort = [[tb_parametre+".OrdreThio", "ASCENDING"],\
    [couche_stations+".BilanENVOrdre", "ASCENDING"],\
    [tb_donnees+".date_", "ASCENDING"]]
arcpy.Delete_management(gdb + "/" + tb_parametre_sorted)
arcpy.Sort_management(join_data, gdb + "/" + tb_parametre_sorted, fields_sort)
try:
    arcpy.DeleteField_management(join_data, del_fields)
except:
    print("Pas de champs à supprimer")
arcpy.MakeTableView_management(gdb + "/" + tb_parametre_sorted, tb_parametre_sorted)

field_join = [f.name for f in arcpy.ListFields(tb_parametre_sorted)]

""" Ordre des champs : 

0:'OBJECTID', 1:'SHAPE',

2:'TB_DONNEES_THIO_id_parametre', 3:'TB_DONNEES_THIO_id_geometry', 4:'TB_DONNEES_THIO_date_',
5:'TB_DONNEES_THIO_valeur', 6:'TB_DONNEES_THIO_BilanEnv_id_parametre',

7:'TB_PARAMETRES_THIO_OBJECTID', 8:'TB_PARAMETRES_THIO_id_parametre', 9:'TB_PARAMETRES_THIO_nom',
10:'TB_PARAMETRES_THIO_unite', 11:'TB_PARAMETRES_THIO_BilanTHIO', 12:'TB_PARAMETRES_THIO_OrdreThio',

13:'FC_STATIONS_THIO_OBJECTID', 14:'FC_STATIONS_THIO_nom', 15:'FC_STATIONS_THIO_id_geometry',
16:'FC_STATIONS_THIO_nom_simple', 17:'FC_STATIONS_THIO_BilanENVRefSuiviTxt',
18:'FC_STATIONS_THIO_BilanENVZone', 19:'FC_STATIONS_THIO_BilanENVOrdre',
20:'FC_STATIONS_THIO_BilanENVGraphique'
"""

sCurs = [list(sc) for sc in arcpy.da.SearchCursor(tb_parametre_sorted, field_join)]
plot_data = OrderedDict()
unite_parametre = dict()
extrem_parametres = dict()
donnee_station = dict()
type_station = dict()
dico_percentile_ref = dict()

for row in sCurs:
    date = row[4]
    nom_param = row[9]
    nom_station = row[14]
    type_st = row[18]
    valeur = row[5]  
    unite = row[10]
    # Si nouveau paramètre
    if nom_param not in plot_data.keys():
        unite_parametre[nom_param] = unite
        plot_data[nom_param] = OrderedDict()
        plot_data[nom_param][nom_station] = [[], []]
        type_station[nom_station] = type_st
        dico_percentile_ref[nom_param] = {"Ultramafique": None, "Mixte": None}
    # Sinon si nouvelle station
    elif nom_station not in plot_data[nom_param].keys():
        plot_data[nom_param][nom_station] = [[], []]
        type_station[nom_station] = type_st
    # dans tout les cas
    plot_data[nom_param][nom_station][0].append(date)
    plot_data[nom_param][nom_station][1].append(valeur)

# Récupération des percentiles 75 ds gammes de référence
gamme_ref = pd.read_excel(fichier_excel_gamme_ref)
for i in range(len(gamme_ref)):
    if gamme_ref.loc[i, "TypeData"] == "Gamme de reference - Percentile 75":
        for param in dico_percentile_ref.keys():
            try:
                dico_percentile_ref[param][gamme_ref.loc[i, "Zone"]] = gamme_ref.loc[i, param]
            except:
                print("Veuillez vérifier les données du fichier excel")

# Recherche des valeurs min et max pour chaque paramètre (bug si min > 10 000 000 u)
for parametre, stations in plot_data.items():
    extrem_parametres[parametre] = [10000000, 0]
    for station, donnees in stations.items():
        for mesure in donnees[1]:
            if mesure < extrem_parametres[parametre][0]:
                extrem_parametres[parametre][0] = mesure
            if mesure > extrem_parametres[parametre][1]:
                extrem_parametres[parametre][1] = mesure

count = 0
with PdfPages("rapportThio.pdf") as pdf:
    for parametre, station_data in plot_data.items():
        count = 0
        for station, donnees in station_data.items():
            # Si le compteur est un multiple de 3
            if count%3 == 0:
                if count > 0:
                    pdf.savefig()
                    plt.close()
                # Remise à zéro et création d'une nouvelle figure
                count = 0
                fig, axes = plt.subplots(3, 1)
                fig.suptitle(parametre, fontsize=15)
                plt.subplots_adjust(hspace=0.5)
                for ax in axes:
                    ax.set_visible(False)
            # Mise en forme des données
            ax = axes[count]
            data = pd.DataFrame({'Date': donnees[0], 'Valeurs': donnees[1]})
            data.set_index('Date', inplace=True)
            p75 = dico_percentile_ref[parametre][type_station[station]]
            Y = np.array(data['Valeurs'].dropna().values, dtype=float)
            X = np.array(pd.to_datetime(data['Valeurs'].dropna()).index.values, dtype=float)
            Xp = [dt_dbt_etude, dt_fin_etude]
            Yp = [p75]*2
            # Calcul de la régression linéaire
            lin_fit = sp.linregress(X, Y)
            slope, intercept = lin_fit[0], lin_fit[1] 
            xf = np.linspace(min(X), max(X), 100)
            yf = (slope*xf) + intercept
            xf = pd.to_datetime(xf)
            ax.set_visible(True)
            # Tracé des courbes
            lin_fit_label = "{:.6f}x".format(slope*10**9*3600*24*365.25)
            ax.plot(data['Valeurs'].index, Y, marker='x', ls='', label='', color='orange')
            ax.plot(xf, yf, label=lin_fit_label, linewidth=0.7, color='blue')
            ax.plot(Xp, Yp, label="P75 : {:e}".format(p75), linewidth=0.7, color="red")
            # Décoration
            ax.legend()
            ax.set_ylabel(unite_parametre[parametre])
            ax.set_title(station + ", " + type_station[station], fontsize=11)
            trim_tick = mdates.MonthLocator([4, 7, 10])
            year_tick = mdates.YearLocator(1)
            ax.xaxis.set_minor_locator(trim_tick)
            ax.xaxis.set_major_locator(year_tick)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%Y"))
            ax.set_xlim(dt_dbt_etude, dt_fin_etude)
            ax.set_ylim(bottom=extrem_parametres[parametre][0]*1.1-extrem_parametres[parametre][1]*0.1, top=extrem_parametres[parametre][1] * 1.05)
            count += 1
        # Sauvegarde à la fin de chaque paramètre
        try:
            pdf.savefig()
            plt.close()
        except:
            print("figure introuvable")
