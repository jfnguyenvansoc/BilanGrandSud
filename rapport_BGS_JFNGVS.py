############################################MODULES###############################
#################################################################################
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
###################################################################################

########################FONCTIONS##################################################
###################################################################################
def get_del_field(fcIn, liste_champ):
    fields = arcpy.ListFields(fcIn)
    fields_del = []
    for f in fields:
        if f.name not in liste_champ:
            fields_del.append(f.name)
    return fields_del

def TrouveReference(GR,DicoPerc,ChpExcRef,inds,prm):
    dicoRefs = dict()
    for ind in inds:
        if np.isnan(GR.loc[ind,prm]):
            dicoRefs[GR.loc[ind, ChpExcRef]] = 0
        else:
            dicoRefs[GR.loc[ind, ChpExcRef]] = GR.loc[ind, prm]
    DicoPerc[prm] = dicoRefs
    return DicoPerc

def OrdreKeys(dico):
    Items = sorted(dico.items(), key=lambda x: x[1], reverse=False)
    cles = [item[0] for item in Items]
    return cles

###################################################################################

################################################PARAMETRES############################################
######################################################################################################
rep = r"C:\Users\jeanf\OneDrive - Association OEIL\DONNEES\Traitement_SIG\GALAXIA\Tous_Types\2020\200825_IntegrationSuiteGDSud2018_2019_THIO"
gdb = rep + r"\GALAXIA_TOPOMAT_FORMES.gdb"
FichierPDFSortie = "ExportBGSSout_Graphique_2020.pdf"
annee_fin_etude = 2020
fichier_excel_gamme_ref = rep + r"\ExportBilanGrandSud\ExportBGSSout2020.xlsx"
tb_parametre = "TB_PARAMETRES"
couche_stations = "FC_STATIONS"
tb_donnees = "TB_DATA_SOUT"
#Champ de reference sur la couche station FC_STATIONS
ChampStationRef = "BilanENVTypeStation"
#Sur Export Excel
ChampExcelRef = "Localisation"
#A changer selon les differents types de reference
#DicoTypeRef = {"Riviere": None}

annee_dbt_etude = annee_fin_etude - 4
dt_dbt_etude = dt.datetime(annee_dbt_etude, 1, 1)
dt_fin_etude = dt.datetime(annee_fin_etude, 12, 31)

tb_parametre_sorted = tb_parametre + "_sorted"
join_data = "join_data"
#Type de suivi ID
IdSuivi = "SOUT"
#Si station de reference en sortie
Reference = "Non"
#Selection particuliere sur la table de donnees
SelectionPart = "Non"

if SelectionPart == "Oui" and IdSuivi == "SOUT":
    SelectData = "BilanEnv_id_parametre LIKE '%/" + IdSuivi + "/%' AND BilanEnv_id_parametre NOT LIKE '%/PHYS_CHIMI_CONT/%'"
elif SelectionPart == "Oui":
    SelectData = "diffusion = 'Oui'"
if SelectionPart == "Oui":
    print(SelectData)
#Champ de station qui donne l'ordre des stations a mettre Ref ou non si Station de suivi ou Ref
if Reference == "Oui":
    ChampOrdreRefStat = "BilanENVOrdreRef"
else:
    ChampOrdreRefStat = "BilanENVOrdre"
#ChampOrdreRefStat = "BilanENVOrdre"
####REQUETE SUIVI#############
if Reference == "Oui" and IdSuivi == "SURF":
    where_stations = "id_geometry ='3_C' OR id_geometry ='WJ_01'"
elif Reference == "Oui" and IdSuivi == "SOUT":
    where_stations = "BilanENVRefSuiviTxt = 'Reference' AND BilanEnvCompartiment LIKE '%" + IdSuivi + "%' AND " + ChampOrdreRefStat + " IS NOT NULL"\
                     + "OR BilanENVRefSuiviTxt = 'Controle' AND BilanEnvCompartiment LIKE '%" + IdSuivi + "%' AND " + ChampOrdreRefStat + " IS NOT NULL"
elif Reference == "Oui":
    where_stations = "BilanENVRefSuiviTxt = 'Reference' AND BilanEnvCompartiment LIKE '%" + IdSuivi + "%' AND " + ChampOrdreRefStat + " IS NOT NULL"
else:
    where_stations = "BilanENVRefSuiviTxt = 'Suivi' AND BilanEnvCompartiment LIKE '%" + IdSuivi + "%' AND " + ChampOrdreRefStat + " IS NOT NULL"

#Requete parametres#
where_params = "BilanENV = 'Oui' AND id_cat2 = '" + IdSuivi + "'"
field_data = {"OBJECTID":0, "id_parametre":1, "id_geometry":2, "date_":3, "valeur":4, "BilanEnv_id_parametre":5}
field_parametres = {"OBJECTID":0, "id_parametre":1, "nom":2, "unite":3, "BilanENV":4, "OrdreBilanENV":5, "id_cat2":6, "BilanENVLimiteMax":7}
#changmeent BilanENVTypeStation
field_station = {"OBJECTID":0, "SHAPE":1, "nom":2, "id_geometry":3, "nom_simple":4, "BilanENVRefSuiviTxt":5, ChampStationRef:6, ChampOrdreRefStat:7, "BilanEnvCompartiment":8}
#Condition de prise en compte des limites de valeurs max par parametre mettre "Oui" ou "Non"
LimiteMax = "Non"
##############################################FIN DES PARAMETRES ########################################################
####MODIFICATION POUR LA RECUPERATION DES DONNEES############################
if LimiteMax == "Oui":
    print("Limite de valeur est un parametre a considerer")
else:
    print("Pas de limite de valeur a considerer")
arcpy.MakeTableView_management(gdb + "/" + tb_parametre, tb_parametre)
arcpy.MakeFeatureLayer_management(gdb + "/" + couche_stations, couche_stations)
if SelectionPart == "Oui":
    print("Filtre sur les donnes a ete realise avec la selection: " + SelectData)
    arcpy.MakeTableView_management(gdb + "/" + tb_donnees, tb_donnees, SelectData)
else:
    arcpy.MakeTableView_management(gdb + "/" + tb_donnees, tb_donnees)
FParams = OrdreKeys(field_parametres)
print(FParams)
FStations = OrdreKeys(field_station)
print(FStations)
print("Recuperation des parametres")
Parametres = [list(p) for p in arcpy.da.SearchCursor(tb_parametre, FParams , where_params, sql_clause=(None,"ORDER BY OrdreBilanENV"))]
print(str(len(Parametres)) + " parametres recuperes")
print("Recuperation des stations")
Stations = [list(s) for s in arcpy.da.SearchCursor(couche_stations, FStations, where_stations, sql_clause=(None, "ORDER BY " + ChampOrdreRefStat))]
print(str(len(Stations)) + " stations recuperees")
Datas = list()
DatasField = {"OBJECTID":0, "id_parametre":1, "nom_station":2, "date_":3, "valeur":4, "nom_parametre":5, "unite":6, "typestation":7}
FDatas = OrdreKeys(field_data)
arcpy.Delete_management(couche_stations)
arcpy.Delete_management(tb_parametre)
ParamLimites = dict()
if LimiteMax == "Oui":
    for p in Parametres:
        ParamLimites[p[field_parametres["id_parametre"]]] = p[field_parametres["BilanENVLimiteMax"]]

print("Recuperation des donnees")
for p in Parametres:
    for s in Stations:
        #Filtre SQL si on limite les valeurs max ou non
        if LimiteMax == "Oui":
            if p[field_parametres["nom"]] == "Cadmium dissous":
                SQLDataSel = "date_ >= date '" + str(annee_dbt_etude) + "-01-01' AND date_ < date '" + str(
                    annee_fin_etude + 1) + "-01-01' AND id_parametre = '" + p[
                                 field_parametres["id_parametre"]] + "' AND id_geometry = '" + s[
                                 field_station["id_geometry"]] + "' AND valeur < 0.01001"
            elif not ParamLimites[p[field_parametres["id_parametre"]]] is None:
                SQLDataSel = "date_ >= date '" + str(annee_dbt_etude) + "-01-01' AND date_ < date '" + str(
                    annee_fin_etude + 1) + "-01-01' AND id_parametre = '" + p[
                                 field_parametres["id_parametre"]] + "' AND id_geometry = '" + s[
                                 field_station["id_geometry"]] + "' AND valeur < " + str(ParamLimites[p[field_parametres["id_parametre"]]])
            else:
                SQLDataSel = "date_ >= date '" + str(annee_dbt_etude) + "-01-01' AND date_ < date '" + str(
                    annee_fin_etude + 1) + "-01-01' AND id_parametre = '" + p[
                                 field_parametres["id_parametre"]] + "' AND id_geometry = '" + s[
                                 field_station["id_geometry"]] + "'"
        else:
            SQLDataSel = "date_ >= date '" + str(annee_dbt_etude) + "-01-01' AND date_ < date '" + str(
                annee_fin_etude + 1) + "-01-01' AND id_parametre = '" + p[
                             field_parametres["id_parametre"]] + "' AND id_geometry = '" + s[
                             field_station["id_geometry"]] + "'"
        DataSel = [list(dt) for dt in arcpy.da.SearchCursor(tb_donnees, FDatas,SQLDataSel,sql_clause=(None,"ORDER BY date_"))]
        for dt in DataSel:
            idt = DataSel.index(dt)
            DataSel[idt][field_data["id_geometry"]] = s[field_station["nom_simple"]]
            DataSel[idt][field_data["BilanEnv_id_parametre"]] = p[field_parametres["nom"]]
            DataSel[idt].append(p[field_parametres["unite"]])
            DataSel[idt].append(s[field_station[ChampStationRef]])
        Datas.extend(DataSel)
print(str(len(Datas)) + " donnees recuperees")
arcpy.Delete_management(tb_donnees)
print("integration des donnees sur dictionnaires")
#####################FIN MODIFICATION JFNGVS###########################
plot_data = OrderedDict()
unite_parametre = dict()
extrem_parametres = dict()
donnee_station = dict()
type_station = dict()
dico_percentile_ref = dict()
Params = list()
for donnee in Datas:
    date = donnee[DatasField["date_"]]
    nom_param = donnee[DatasField["nom_parametre"]]
    nom_station = donnee[DatasField["nom_station"]]
    type_st = donnee[DatasField["typestation"]]
    valeur = donnee[DatasField["valeur"]]
    unite = donnee[DatasField["unite"]]
    Params.append(nom_param)
    # Si nouveau paramètre
    if nom_param not in plot_data.keys():
        unite_parametre[nom_param] = unite
        plot_data[nom_param] = OrderedDict()
        plot_data[nom_param][nom_station] = [[], []]
        type_station[nom_station] = type_st
        #dico_percentile_ref[nom_param] = DicoTypeRef
    # Sinon si nouvelle station
    elif nom_station not in plot_data[nom_param].keys():
        plot_data[nom_param][nom_station] = [[], []]
        type_station[nom_station] = type_st
    # dans tout les cas
    plot_data[nom_param][nom_station][0].append(date)
    plot_data[nom_param][nom_station][1].append(valeur)
print(plot_data.keys())
print("Recuperation des percentiles 75")
# Récupération des percentiles 75 ds gammes de référence
gamme_ref = pd.read_excel(fichier_excel_gamme_ref)
iGRs = list()
for i in range(len(gamme_ref)):
    if gamme_ref.loc[i, "TypeData"] == "Gamme de reference - Percentile 75":
        iGRs.append(i)

for param in Parametres:
    dico_percentile_ref = TrouveReference(gamme_ref, dico_percentile_ref, ChampExcelRef, iGRs, param[field_parametres["nom"]])
print(dico_percentile_ref)
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
print("Creation du PDF avec les graphiques")
with PdfPages(FichierPDFSortie) as pdf:
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
