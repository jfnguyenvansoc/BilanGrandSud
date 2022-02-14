# -*- coding: utf8 -*-

import arcpy, datetime, xlsxwriter
from unidecode import unidecode
from ModuleBilanEnv import SQLParam, LignesSup, PeriodeAnnee, PeriodeSsFLUX, FormatExcel, LigneExcelValeur
from ModuleBilanEnvStatistiquesGammeRef import StatistiquesGammeReference
from ModuleBilanEnvStatistiquesPeriode import Statistiques3ans, StatistiquesPeriode
from ModuleBilanEnvExportMetrique import ExportMetrique

TB_DONNEES = arcpy.GetParameterAsText(0)
ChampDataParam = arcpy.GetParameterAsText(1)
FC_STATIONS = arcpy.GetParameterAsText(2)
TB_PARAMETRES = arcpy.GetParameterAsText(3)
TB_SOURCE = arcpy.GetParameterAsText(4)
TB_CATEGORIE_NAME = arcpy.GetParameterAsText(5)
Suivi = arcpy.GetParameterAsText(6) # valeur texte du suivi a recuperer ensuite l'id
RepSortie= arcpy.GetParameterAsText(7) #Repertoire pour les sorties en sortie
ExcelSortie = arcpy.GetParameterAsText(8) #Nom du fichier Excel en sortie
TypePeriode = arcpy.GetParameterAsText(9) # type de periode en fonction du suivi
TB_BilanENVXMLParam = arcpy.GetParameterAsText(10) #Table de correspondance des parametres pour XML
TB_BilanENVXMLMetr = arcpy.GetParameterAsText(11) #Table de correspondance des metriques pour XML
TB_BilanENVXMLGam = arcpy.GetParameterAsText(12)
NomXMLSortie = arcpy.GetParameterAsText(13) #Nom du fichier XML en sortie
MetriqueXMLVoulues = arcpy.GetParameterAsText(14) #Liste des metriques voulues
MetriqueXMLVoulues = MetriqueXMLVoulues.split(";")
MetriqueXMLVoulues = [unidecode(mxml).replace("'","") for mxml in MetriqueXMLVoulues]
NomXMLGRSortie = arcpy.GetParameterAsText(15)
GammeXMLVoulues = arcpy.GetParameterAsText(16)
GammeXMLVoulues = GammeXMLVoulues.split(";")
GammeXMLVoulues = [unidecode(gxml).replace("'","") for gxml in GammeXMLVoulues]
AnneeEtude = int(arcpy.GetParameterAsText(17))
ExcelResultatPrec = arcpy.GetParameterAsText(18)
if NomXMLSortie[-4:] != ".xml":
    NomXMLSortie = NomXMLSortie + ".xml"
if NomXMLGRSortie[-4:] != ".xml":
    NomXMLGRSortie = NomXMLGRSortie + ".xml"
milieu = "Marin"

#organisation des premieres colonnes
indiceColMar = {"TypeData":0, "TypeStation":1, "Zone":2, "TypologieStation":3, "TypeRef":4, "Station":5, "Date":6, "Periode":7,"PremierParam":8}


########TRAITEMENT DES INFORMATIONS##########
arcpy.AddMessage("le gros matou du Milieu marin se reveille et vous dire BONJOUR, il commence a bosser. Alors il recupere l'identifiant de la categorie correspondant au Suivi")
IdSuivi = [str(c[1]) for c in arcpy.da.SearchCursor(TB_CATEGORIE_NAME,["nom_complet","id_cat"]) if c[0] == Suivi]
IdSuivi = IdSuivi[0]
arcpy.AddMessage("Le gros matou recupere les informations des stations et parametres du Bilan Grand Sud Identifiant Suivi = " + IdSuivi)
if IdSuivi == "RECOUVR32":
    Params = [list(p) for p in arcpy.da.SearchCursor(TB_PARAMETRES,[ChampDataParam,"BilanENV_nom","BilanENV","unite","BilanEnvSeuil", "OrdreBilanENV"],sql_clause=(None,'ORDER BY OrdreBilanENV'))
          if p[0].find("SUBSTR/RECOUVR/") >= 0 and p[2] == "Oui"]
    Params = [p[0] + ";" + p[1] + ";" + p[2] + ";" + p[3] + ";" + str(p[4]) for p in Params]
    Params = list(set(Params))
    Params = [p.split(";") for p in Params]
elif IdSuivi == "RECOUVR12":
    Params = [list(p) for p in arcpy.da.SearchCursor(TB_PARAMETRES,[ChampDataParam,"nom","BilanENV","unite","BilanEnvSeuil", "OrdreBilanENV"],sql_clause=(None,'ORDER BY OrdreBilanENV'))
          if p[0].find("ACROPORA/RECOUVR/") >= 0 and p[2] == "Oui"]
    Params = [p[0] + ";" + p[1] + ";" + p[2] + ";" + p[3] + ";" + str(p[4]) for p in Params]
    Params = list(set(Params))
    Params = [p.split(";") for p in Params]
elif IdSuivi == "EAU":
    Params = [list(p) for p in arcpy.da.SearchCursor(TB_PARAMETRES,[ChampDataParam,"nom","BilanENV","unite","BilanEnvSeuil", "OrdreBilanENV"],sql_clause=(None,'ORDER BY OrdreBilanENV'))]
    Ordres = list(set([ordr[5] for ordr in Params]))
    ParamsFilt = list()
    for o in Ordres:
        ParamSel = [p for p in Params if p[5] == o]
        ParamsFilt.append(ParamSel[0])
    Params = list()
    Params.extend(ParamsFilt)
    del ParamsFilt
    Params = [p[0] + ";" + p[1] + ";" + p[2] + ";" + p[3] + ";" + str(p[4]) for p in Params]
    Params = list(set(Params))
    Params = [p.split(";") for p in Params]
else:
    Params = [list(p) for p in arcpy.da.SearchCursor(TB_PARAMETRES,[ChampDataParam,"nom","BilanENV","unite","BilanEnvSeuil", "OrdreBilanENV"],"BilanENV = 'Oui' AND "+ ChampDataParam + " LIKE '%" + IdSuivi + "%'",
                                                     sql_clause=(None,'ORDER BY OrdreBilanENV'))]

if len(Params) == 0:
    arcpy.AddMessage("Dommage .... Aucun parametre pour le Bilan Environnement sur ce type de suivi ou categorie")
    del Params
    arcpy.AddMessage("Fin du travail pour le gros matou!!! Oueeee")
else:
    arcpy.AddMessage(str(len(Params)) + " parametres pour le suivi " + Suivi)
    SQLStation = "BilanENVRefSuiviTxt IS NOT NULL AND BilanENVCompartiment LIKE '%" + IdSuivi + "%'"
    Stations = [list(s) for s in arcpy.da.SearchCursor(FC_STATIONS,
                                                       ["id_geometry", "nom_simple", "BilanENVRefSuiviTxt", "BilanENVTypologie","BilanENVZone","BilanEnvType","BilanEnvCompartiment", "BilanENVOrdre", "BilanENVOrdreRef"],
                                                                                                          SQLStation, sql_clause=(None,'ORDER BY BilanENVOrdre'))]
    arcpy.AddMessage("Alors le gros matou constitue un Filtre SQL a constituer pour appliquer sur la table de donnees")
    PmsIds = [p[0] for p in Params]
    StatIds = [s[0] for s in Stations]
    SQL = SQLParam(PmsIds,StatIds,TB_DONNEES)
    arcpy.AddMessage("Nombre de stations: "+ str(len(Stations)))

    arcpy.AddMessage("Maintenant le gros matou recupere les donnees ordonnees par date de mesure, station, parametre")
    arcpy.SelectLayerByAttribute_management(TB_DONNEES, "CLEAR_SELECTION")
    #Periode specifique pour substrat Vale NC d'aout a Novembre

    if IdSuivi == "RECOUVR32":
        Datas = [list(d) for d in arcpy.da.SearchCursor(TB_DONNEES,["id_geometry",ChampDataParam,"date_","valeur","methode_analyse","valeur_texte","signe","limite_quantification"],sql_clause=(None,'ORDER BY date_, id_geometry,' + ChampDataParam))
                 if d[2].month >= 8 and d[2].month <= 11]
    else:
        Datas = [list(d) for d in arcpy.da.SearchCursor(TB_DONNEES,["id_geometry", ChampDataParam, "date_", "valeur", "methode_analyse", "valeur_texte", "signe",
                                                                    "limite_quantification"], ChampDataParam + " LIKE '%" + IdSuivi +"%' AND date_ < date '" + str(AnneeEtude+1) + "-01-01'",
                                                        sql_clause=(None,'ORDER BY date_, id_geometry,' + ChampDataParam))]
    arcpy.AddMessage(str(len(Datas)) + " donnees recuperees pour le Bilan Environnement sur " + Suivi + " pour le gros matou. Y'a du taff")
    arcpy.AddMessage("Faut bien stocker quelque part le resultat... le gros matou cree la fichier Excel en sortie. Vous lui avez dit qu'il s'appelle " + ExcelSortie)
    XLSX = xlsxwriter.Workbook(RepSortie + "/" + ExcelSortie)
    Feuille = XLSX.add_worksheet("Donnees")
    FormTitre = FormatExcel(XLSX, "Oui", 11, "Calibri Light", "#ffffff", "#5b9ad5", "")
    FormGamme = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#c00000", "", "")
    FormMet = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#c05a11", "", "")
    FormVal = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#0070c0", "", "")
    FormInfo = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "000000", "", "")
    FormNorm = FormatExcel(XLSX, "Non", 10, "Calibri Light", "#000000", "", "")
    FormDate = FormatExcel(XLSX, "Non", 10, "Calibri Light","#000000","", "dd/mm/yyyy HH:MM:SS")
    
    ###REPRISE PENSER A INSERER INFORMATION AMONT/AVAL STATION, GRANDBV/PETITBV, ZONEBILAN
    Feuille.write(0,0,"TypeData",FormTitre)
    Feuille.write(0, indiceColMar["TypeStation"],"TypeStation",FormTitre)
    Feuille.write(0, indiceColMar["Zone"],"Zone",FormTitre)
    Feuille.write(0, indiceColMar["TypologieStation"],"TypologieStation", FormTitre)
    Feuille.write(0, indiceColMar["TypeRef"],"TypeRef",FormTitre)
    Feuille.write(0, indiceColMar["Station"],"Station",FormTitre)
    Feuille.write(0, indiceColMar["Date"],"Date_",FormTitre)
    Feuille.write(0, indiceColMar["Periode"],"Periode",FormTitre)
    Line0 = ["Valeur Seuil", "", "", "", "", "", None, ""]
    Line1 = ["Filtre Bilan Env","","","","","", None,""]
    Line2 = ["Unite","","","","","", None,""]
    #on ajoute en champ chaque parametre avec id parametre en nom de champ et en alias le nom du parametre
    #on met sur la premiere ligne les valeurs seuils pour chaque parametre Line0
    #on met sur la deuxieme ligne le type de parametre si cest cle ou simple parametre de Bilan ENV avec code Cle ou Env Line1
    for p in Params:
        Feuille.write(0,indiceColMar["PremierParam"] + Params.index(p), p[1], FormTitre)
        Line0.append(p[4])
        Line1.append(p[2])
        Line2.append(p[3])
    lx = 0
    for l in Line0:
        Feuille.write(1,lx, l, FormInfo)
        lx = lx + 1
    lx = 0
    for l in Line1:
        Feuille.write(2,lx, l, FormInfo)
        lx = lx + 1
    lx = 0
    for l in Line2:
        Feuille.write(3,lx, l, FormInfo)
        lx = lx + 1
    #indice pour les lignes Excel dans le fichier et figer les volets sur A5
    Feuille.freeze_panes(4,0)
    Feuille.set_column(0, Feuille.dim_colmax,15)
    Feuille.autofilter(3, 0, 3, Feuille.dim_colmax)
    indxl = 4
    #on trouve les champs pour declarer linsertion de valeur dans la table en sortie
    Champs = ["TypeData","TypeStation","Zone","TypologieStation","TypeRef","Station","Date_","Periode"]
    PmsIdsNew = [p.replace("/","_") for p in PmsIds]
    Champs.extend(PmsIdsNew)
    arcpy.AddMessage(str(len(Champs)) + " Champs")
    ##Creation de liste pour recuperee methode signe a la place des valeurs
    ##mettre des cas en fonction annee, saison... ajouter un parametre a l'outil
    ##premier essai par annee
    DatasMeth = list()
    DatasSigne = list()
    DatasPeriode = list()
    DatasExp = list()
    arcpy.AddMessage("Premiere partie du travail, le gros matou traite les donnees brutes, les valeurs reellement mesurees... c'est parti")
    for s in Stations:
        nbS = len(Stations)
        sI = Stations.index(s)
        pct25 = round(nbS/4)
        pct50 = round(nbS/2)
        pct75 = pct25 + pct50
        pct100 = nbS - 1
        if sI == pct25:
            arcpy.AddMessage("C'est deja ca... Le gros matou a traite deja 25% des valeurs brutes traitees")
        if sI == pct50:
            arcpy.AddMessage("Aller la moitie... Le gros matou a traite deja 50% des valeurs brutes traitees")
        if sI == pct75:
            arcpy.AddMessage("C'est bien .... Le gros matou a traite deja 75% des valeurs brutes traitees")
        if sI == pct100:
            arcpy.AddMessage("Waouw le gros matou a traite 100% des valeurs brutes traitees")

        #on selectionne par station les donnees
        DatasSelect = [d for d in Datas if d[0] == s[0]]
        if len(DatasSelect) > 0:
            #on recupere les valeurs uniques de date
            DatesUniqueStat = [d[2] for d in DatasSelect]
            DatesUniqueStat = list(set(DatesUniqueStat))
            DatesUniqueStat.sort()
            # par date on cree une entree
            for dt in DatesUniqueStat:
                DatasSelDate = [d for d in DatasSelect if d[2] == dt]
                #mettre condition ici si differentes periodes
                ###MODIFICATION#################
                if TypePeriode == "Saisons - Flux a particule":
                    Line = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt, PeriodeSsFLUX(dt)]
                    LineM = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt, PeriodeSsFLUX(dt)]
                    LineS = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt, PeriodeSsFLUX(dt)]
                else:
                    Line = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt,PeriodeAnnee(dt)]
                    LineM = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt,PeriodeAnnee(dt)]
                    LineS = ["Valeur reelle",s[2],s[4],s[3],s[5],s[1],dt,PeriodeAnnee(dt)]
                #Les listes de type OtherDatas (M pour Methodes, S pour Signe) correspondent aux lignes supplementaires si plusieurs valeurs sont ...
                #rencontrees pour un meme trio parametre-station-date,on conserve ces nouvelles infos pour creer de nouvelles lignes avec la fonction ...
                #LignesSup les lignes sont constituees dun index unique par ligne, index parametre, index sur lemplacement de la ligne, puis valeur du parametre
                OtherDatas = list()
                OtherDatasM = list()
                OtherDatasS = list()
                for p in Params:
                    y = Params.index(p)
                    DatasSelDateP = [d for d in DatasSelDate if d[1] == p[0]]
                    if len(DatasSelDateP) == 0:
                        #on laisse vide si aucune valeur pour le parametre
                        Line.append("")
                        LineM.append("")
                        LineS.append("")
                    elif len(DatasSelDateP) > 1:
                        #cas si plusieurs valeurs pour une date sur une station pour un parametre
                        nbdt = len(DatasSelDateP)
                        a = 1
                        while a < nbdt:
                            #information servant a recreer ligne supplementaire grace a fonction LignesSup declaree precedemment
                            OtherDatas.append([a,y,y+indiceColMar["PremierParam"],DatasSelDateP[a][5].replace(".",",")])
                            OtherDatasM.append([a,y,y+indiceColMar["PremierParam"],DatasSelDateP[a][4]])
                            OtherDatasS.append([a, y,y+indiceColMar["PremierParam"],DatasSelDateP[a][6]])
                            a = a + 1
                        Line.append(DatasSelDateP[0][5].replace(".",","))
                        LineM.append(DatasSelDateP[0][4])
                        LineS.append(DatasSelDateP[0][6])
                    else:
                        #si une valeur on la recupere
                        DatasSelDateP = DatasSelDateP[0]
                        Line.append(DatasSelDateP[5].replace(".",","))
                        LineM.append(DatasSelDateP[4])
                        LineS.append(DatasSelDateP[6])
                LigneExcelValeur(FormVal, FormNorm, FormDate, Feuille, Line, indxl, milieu)
                indxl = indxl + 1
                DatasExp.append(Line)
                DatasMeth.append(LineM)
                DatasSigne.append(LineS)
                if len(OtherDatas) > 0:
                    indlines = [od[0] for od in OtherDatas]
                    indlines = list(set(indlines))
                    OtherLines = LignesSup(OtherDatas,Line,Params,indiceColMar["PremierParam"])
                    OtherLinesM = LignesSup(OtherDatasM,LineM,Params,indiceColMar["PremierParam"])
                    OtherLinesS = LignesSup(OtherDatasS,LineS,Params,indiceColMar["PremierParam"])
                    del OtherDatas,OtherDatasM, OtherDatasS
                    for ol in OtherLines:
                        LigneExcelValeur(FormVal, FormNorm, FormDate, Feuille, ol, indxl, milieu)
                        indxl = indxl + 1
                        DatasExp.append(ol)
                    del OtherLines
                    for olm in OtherLinesM:
                        DatasMeth.append(olm)
                    del OtherLinesM
                    for ols in OtherLinesS:
                        DatasSigne.append(ols)
                    del OtherLinesS
                    
            #arcpy.AddMessage("Donnees traitees pour la station " + s[0])
        else:
            VideComplet = ""

    arcpy.AddMessage("Y'a encore du taff!!! ppfff Deuxieme partie du travail: le gros matou va traiter les metriques (PS: le gros matou commence a fatiguer")
    del Datas
    if TypePeriode == "Saisons - Flux a particule":
        arcpy.AddMessage("Le gros matou vous parle: La periode est: " + TypePeriode)
        DatasMethSC = [d for d in DatasMeth if d[indiceColMar["Periode"]][0:13] == "Saison chaude"]
        DatasSigneSC = [d for d in DatasSigne if d[indiceColMar["Periode"]][0:13] == "Saison chaude"]
        DatasExpSC = [d for d in DatasExp if d[indiceColMar["Periode"]][0:13] == "Saison chaude"]

        DatasMethSF = [d for d in DatasMeth if d[indiceColMar["Periode"]][0:14] == "Saison fraiche"]
        DatasSigneSF = [d for d in DatasSigne if d[indiceColMar["Periode"]][0:14] == "Saison fraiche"]
        DatasExpSF = [d for d in DatasExp if d[indiceColMar["Periode"]][0:14] == "Saison fraiche"]
        #Gamme de reference avant Statistiques de Periode pour le Percentile 75
        arcpy.AddMessage("le gros matou est content pas de gamme de reference pour les flux a particule :-(")
        #suppression StatsGammP75 =[stg for stg in StatsGamm if stg[0] == "Gamme de reference - Percentile 75"]
        arcpy.AddMessage("le gros matou va donc traiter les metriques par Periode :-(")
        StatsGamm = list()
        StatsGammP = list()

        StatsPeriode = StatistiquesPeriode(milieu, DatasMeth, DatasSigne, DatasExp, Params, StatsGammP, IdSuivi)
        #insertion Statistiques par Periode avant les gammes de reference
        for stm in StatsPeriode:
            LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, stm, indxl, milieu)
            indxl = indxl + 1
    else:
        arcpy.AddMessage("Le gros matou vous parle: La periode est: " + TypePeriode)
        ##Mettre ici Gamme reference
        arcpy.AddMessage("Pour le milieu marin le gros matou va traiter d'abord les metriques par periode :-(")
        arcpy.AddMessage(str(len(DatasExp)) + " Donnees brutes")
        StatsGammP = list()
        StatsPeriode = StatistiquesPeriode(milieu, DatasMeth, DatasSigne, DatasExp, Params, StatsGammP, IdSuivi)
        arcpy.AddMessage("le gros matou va traiter ensuite les gammes de reference :-(")
        if IdSuivi == "EAU":
            Stat3ans = Statistiques3ans(DatasMeth,DatasSigne,DatasExp,Params,IdSuivi,AnneeEtude)
            arcpy.AddMessage("Nombre de Statistique pour les 3 ans: " + str(len(Stat3ans)))
            StatsPeriodeSel = [stp for stp in StatsPeriode if stp[indiceColMar["TypeData"]].find("Metrique - Ecart Moyenne") >= 0 and int(stp[indiceColMar["Periode"]]) < AnneeEtude
                                or stp[indiceColMar["TypeData"]].find("Metrique - Moyenne") >= 0 and int(stp[indiceColMar["Periode"]]) < AnneeEtude]
            arcpy.AddMessage(str(len(StatsPeriodeSel)) + " Donnees Periodes selectionnees")
            StatsGamm = StatistiquesGammeReference(milieu, DatasMeth, DatasSigne, DatasExp, Params,IdSuivi, AnneeEtude, DonneesPerRefEau=StatsPeriodeSel)
            arcpy.AddMessage("Nombre de Statistique pour les gammes de reference " + str(len(StatsGamm)))
        else:
            StatsGamm = StatistiquesGammeReference(milieu, DatasMeth, DatasSigne, DatasExp, Params, IdSuivi, AnneeEtude)
            arcpy.AddMessage("Nombre de Statistique pour les gammes de reference " + str(len(StatsGamm)))
            
        for stm in StatsPeriode:
            LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, stm, indxl, milieu)
            indxl = indxl + 1
        if IdSuivi == "EAU":
            for st3 in Stat3ans:
                LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, st3, indxl, milieu)
                indxl = indxl + 1
        if len(StatsGamm) > 0:
            for stg in StatsGamm:
                LigneExcelValeur(FormGamme, FormNorm, FormDate, Feuille, stg, indxl, milieu)
                indxl = indxl + 1
     #Creation des fichiers XML pour les metriques et gammes de reference
    XLSX.close()
    StationsOrdre = [stat for stat in Stations if stat[7] > 0]
    StationsOrdreRef = [stat for stat in Stations if stat[8] > 0]
    if IdSuivi == "EAU":
        StatsPer3a = list()
        StatsPer3a.extend(Stat3ans)
        StatsPer3a.extend(StatsPeriode)
        ExportMetrique(milieu,StatsPer3a, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi, MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLSortie, StationsOrdre,"Suivi", AnneeEtude, ExcelResultatPrec)
        #ExportMetrique(milieu,StatsPer3a, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi, MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLSortie, StationsOrdreRef,"Reference", AnneeEtude)

    else:
        ExportMetrique(milieu,StatsPeriode, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi, MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLSortie, StationsOrdre,"Suivi", AnneeEtude, ExcelResultatPrec)
        #ExportMetrique(milieu,StatsPeriode, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi, MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLSortie, StationsOrdreRef,"Reference", AnneeEtude)
    #partie metrique par annee
    del DatasMeth, DatasSigne, Stations, Params, StatistiquesPeriode,DatasExp
    arcpy.AddMessage("Le gros matou est fatigue, il va dormir. A plous!!!")
