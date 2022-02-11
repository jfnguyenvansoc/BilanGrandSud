# -*- coding: utf8 -*-
import arcpy, numpy, xlsxwriter, datetime
from unidecode import unidecode

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
AnneeEtude = int(arcpy.GetParameterAsText(17))

GammeXMLVoulues = GammeXMLVoulues.split(";")
GammeXMLVoulues = [unidecode(gxml).replace("'","") for gxml in GammeXMLVoulues]
if NomXMLSortie[-4:] != ".xml":
    NomXMLSortie = NomXMLSortie + ".xml"
if NomXMLGRSortie[-4:] != ".xml":
    NomXMLGRSortie = NomXMLGRSortie + ".xml"

#organisation des premieres colonnes
indiceCol = {"TypeData":0, "TypeStation":1, "Localisation":2, "TypeBV":3, "Zone":4, "ZoneReference":5, "Station":6, "Date":7, "Periode":8,"PremierParam":9}

def SQLParam(ParamList,StatList,TableDonnees):
    ParamList = [unidecode(p) for p in ParamList]
    StatList = [unidecode(s) for s in StatList]
    for pm in ParamList:
        i = ParamList.index(pm)
        if i == 0:
            SQLParam = "id_parametre = '" + pm + "'"
        else:
            SQLParam = SQLParam + " OR id_parametre = '" + pm + "'"
    arcpy.SelectLayerByAttribute_management(TableDonnees, "NEW_SELECTION", SQLParam)
    for s in StatList:
        y = StatList.index(s)
        if y == 0:
            SQLStat = "id_geometry = '" + s + "'"
        else:
            SQLStat = SQLStat + " OR id_geometry = '" + s + "'"
    arcpy.SelectLayerByAttribute_management(TableDonnees,"SUBSET_SELECTION",SQLStat)

#creation dautres lignes supplementaires car si un ou plusieurs parametres a la meme date et meme station plusieurs valeurs sont rencontrees
def LignesSup(AutresDonnees,PremiereLigne,ListParam):
    DebutLigne = PremiereLigne[0:9]
    IndexLigne = [ad[0] for ad in AutresDonnees]
    IndexLigne = list(set(IndexLigne))
    IndexLigne.sort()
    AutresLignes = list()
    for il in IndexLigne:
        DonneeFiltre = [ad for ad in AutresDonnees if ad[0] == il]
        Line = list()
        Line.extend(DebutLigne)
        indligneParam = len(ListParam) + 9
        x = 9
        #tous les index de lignes doivent etre parcours
        while x < indligneParam:
            DonneeFiltParam = [df for df in DonneeFiltre if df[2] == x]
            nbpl = len(DonneeFiltParam)
            #cas ou lindex correspond on va chercher la valeur
            if nbpl == 1:
                Line.append(DonneeFiltParam[0][3])
            #erreur si plusieurs valeurs
            elif nbpl > 1:
                arcpy.AddMessage("Le gros matou vous informe d'une erreur: plusieurs valeurs sont encore rencontrees, concatenation des valeurs")
                b = 0
                concat = ""
                while b < nbpl:
                    if b == 0:
                        concat = DonneeFiltParam[b][3]
                    else:
                        concat = concat + "|" + DonneeFiltParam[b][3]
                    b = b + 1
            #si ne correspond pas au parametre
            else:
                Line.append("")
            x = x + 1
            del DonneeFiltParam
        del DonneeFiltre
        AutresLignes.append(Line)
        del Line
    del DebutLigne, IndexLigne
    return AutresLignes


#Si periode correspond a Annee fonction a utiliser
def PeriodeAnnee(DateEntree):
    return DateEntree.strftime("%Y")


#Si periode correspond a Saison Chaude Macro Inverterbres fonction a utiliser
def PeriodeSsChaudeMIB(DateEntree):
    #saison chaude de mai a juillet
    if DateEntree.month > 9 and DateEntree.month < 13:
        periode = "Saison chaude " + DateEntree.strftime("%Y")
    else:
        periode = "Autre saison " + DateEntree.strftime("%Y")
    return periode


#Si periode correspond a Saison fraiche Poissons Crustaces fonction a utiliser
def PeriodeSsPoiss(DateEntree):
    #saison chaude de mai a juillet
    if DateEntree.month > 4 and DateEntree.month < 8:
        periode = "Saison fraiche " + DateEntree.strftime("%Y")
    else:
        periode = "Autre saison " + DateEntree.strftime("%Y")
    return periode

##Fonction pour la moyenne par jour puis par mois et puis par annee
##Regle abandonnee mais conserve dans le cas ou la regle sera a nouveau appliquee
#en parametres: une liste de valeur non vide souvent sur une station pour une periode sur tous les parametres necessaires, nombre de parametres, indice de colonne du parametre
def RegleMoyenne(DonneesExpPeriodeNnVide,nbparam,IndParam):
    for d in DonneesExpPeriodeNnVide:
        #ajout d'un attribut jour+mois sur chaque ligne
        d.append(d[6].strftime("%d%m"))
        #ajout d'un attribut mois sur chaque ligne
        d.append(d[6].strftime("%m"))
    #recuperation unique des informations jour et mois
    Jours = list()
    Mois = list()
    for d in DonneesExpPeriodeNnVide:
        Jours.append(d[nbparam])
        Mois.append(d[nbparam+1])
    Jours = list(set(Jours))
    Jours.sort()
    Mois = list(set(Mois))
    Mois.sort()
    #traitement des moyennes par jour puis integration dans DonneesJourMoy
    DonneesJourMoy = list()
    for j in Jours:
        DonneeJMoy = [d for d in DonneesExpPeriodeNnVide if d[nbparam] == j]
        LineJMoy  = list()
        LineJMoy.extend(DonneeJMoy[0][0:6])
        LineJMoy.append(None)
        LineJMoy.extend(DonneeJMoy[0][7:9])
        ValsJParam = [float(d[IndParam].replace(",",".")) for d in DonneeJMoy]
        MoyJ = numpy.mean(ValsJParam)
        #on conserve que la colonne correspondant au parametre avec IndParam
        LineJMoy.append(MoyJ)
        #seulement colonnes jour et mois integrees a la fin
        LineJMoy.append(DonneeJMoy[0][nbparam])
        LineJMoy.append(DonneeJMoy[0][nbparam+1])
        DonneesJourMoy.append(LineJMoy)
        del DonneeJMoy, LineJMoy, ValsJParam, MoyJ

    nm = len(DonneesJourMoy[0]) - 1 #longueur dune ligne sur stat moy jour liste car tous les autres parametres inutiles sont vires
    #traitement des moyennes par mois puis integration dans DonneesMoisMoy a partir de DonneesJourMoy
    DonneesMoisMoy = list()
    for m in Mois:
        DonneeMMoy = [d for d in DonneesJourMoy if d[nm] == m]
        LineMMoy = list()
        LineMMoy.extend(DonneeMMoy[0][0:9])
        ValsMParam = [d[9] for d in DonneeMMoy]
        MoyM = numpy.mean(ValsMParam)
        LineMMoy.append(MoyM)
        #on conserve seulement info Mois
        LineMMoy.append(DonneeMMoy[0][nm])
        DonneesMoisMoy.append(LineMMoy)
        del DonneeMMoy, MoyM, ValsMParam, LineMMoy
    del nm, DonneesJourMoy
    #moyenne sur toute la periode a partir de DonneesMoisMoy
    ValsMoisPeriode = [d[8] for d in DonneesMoisMoy]
    del DonneesExpPeriodeNnVide, nbparam, IndParam, Jours, Mois

    #retourne de la moyenne annuelle en texte
    return str(numpy.mean(ValsMoisPeriode)).replace(".",",")

#simple concatenation d une liste de valeur textuelle
def ConcatenationInformation(InfosData):
    InfoTexte = ""
    for i in InfosData:
        ii = InfosData.index(i)
        if ii == 0:
            InfoTexte = i
        else:
            InfoTexte = InfoTexte + "|" + i
    return InfoTexte


def AnneeMinMax(DatesListe):
    MinDate = min(DatesListe)
    MaxDate = max(DatesListe)
    return MinDate.strftime("%Y") + "-" + MaxDate.strftime("%Y")

#trouver les methodes Gamme de ref
def FindMethodes(DonneesM, ind):
    Meths = [d[ind] for d in DonneesM if d[ind] is not None and d[ind] != ""]
    Meths = list(set(Meths))
    MethsTxt = ""
    if len(Meths) > 0:
        MethsTxt = ConcatenationInformation(Meths)
    return MethsTxt
###Fonction pour trouver les LQ#########
def FindLQ(DonneesS,ind,DonneesEx):
    indexLQ = [a for a, e in enumerate(DonneesS) if e[ind] == "<"]
    RechLQ = ""
    if len(indexLQ) == 0:
        RechLQ = ""
    else:
        LQs = list()
        for ilq in indexLQ:
            lq = DonneesEx[ilq][ind]
            LQs.append(lq)
        LQs = list(set(LQs))
        RechLQ= ConcatenationInformation(LQs)
    del indexLQ
    return RechLQ
###Fonction pour trouver les Signes#########
def FindSigne(DonneesS,ind):
    Signs = [str(d[ind]) for d in DonneesS if d[ind] is not None]
    nblq = Signs.count("<")
    return str(nblq)
###Fonction pour trouver les valeurs non vides#########
def FindValsNonVides(DonneesEx, ind):
    DonneesExNnVide = [d for d in DonneesEx if d[ind] is not None and d[ind] != ""]
    valN = len(DonneesExNnVide)
    ValsEx = [float(d[ind].replace(",", ".")) for d in DonneesExNnVide]
    return [valN,ValsEx]  
###Fonction pour calculer les statistiques percentiles, min, moy, ....#########
def FindStatsCal(ValsC):
    # Moyenne en utilisant la fonction RegleMoyenne pour effectuer moyenne par jour puis par mois et enfin sur toute la periode
    # abandon remplace par moyenne numpy.mean sur liste de valeurs
    moy = str(numpy.mean(ValsC)).replace(".", ",")
    # Ecart-Type
    ect = str(numpy.std(ValsC)).replace(".", ",")
    med = str(numpy.median(ValsC)).replace(".", ",")
    # Percentile 10, 25, 75, 9O
    perc10 = str(numpy.percentile(ValsC, 10)).replace(".", ",")
    perc25 = str(numpy.percentile(ValsC, 25)).replace(".", ",")
    perc75 = str(numpy.percentile(ValsC, 75)).replace(".", ",")
    perc80 = str(numpy.percentile(ValsC, 80)).replace(".", ",")
    perc85 = str(numpy.percentile(ValsC, 85)).replace(".", ",")
    perc90 = str(numpy.percentile(ValsC, 90)).replace(".", ",")
    perc95 = str(numpy.percentile(ValsC, 95)).replace(".", ",")
    Min = str(min(ValsC)).replace(".", ",")
    Max = str(max(ValsC)).replace(".", ",")
    return  {"min" : Min, "max" : Max, "moy": moy, "ect" : ect, "med": med, "perc10" : perc10, "perc25" : perc25, "perc75" : perc75, "perc80" : perc80, "perc85" : perc85, "perc90" : perc90, "perc95" : perc95}

################################FONCTION PRINCIPALE POUR CALCULS DE GAMME DE REFERENCE##############################################
def StatistiquesGammeReference(DonneesM,DonneesS,DonneesEx,Parametres,IdentSuiv):
    DonneesStatGammeRef = list()
    #CAS EAUX DOUCES DE SURFACE SEDIMENT AMONT/AVAL GRAND BV/PETIT BV ET DOLINES A PART
    if IdentSuiv == "SURF":
        arcpy.AddMessage("Oh non Gammes de reference pour les eaux de surface")
        #Ultramafique
        DonneesM_UM = [d for d in DonneesM if d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Ultramafique"]
        StatTxt_UM = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesM_UM])))
        Periode_UM = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesM_UM])))
        DonneesS_UM = [d for d in DonneesS if
                         d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Ultramafique"]
        DonneesEx_UM = [d for d in DonneesEx if
                          d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Ultramafique"]

        #Mixte
        DonneesM_Mixte = [d for d in DonneesM if d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Mixte"]
        StatTxt_Mixte = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesM_Mixte])))
        Periode_Mixte = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesM_Mixte])))
        DonneesS_Mixte = [d for d in DonneesS if
                         d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Mixte"]
        DonneesEx_Mixte = [d for d in DonneesEx if
                          d[indiceCol["TypeStation"]] == "RefTHIO" and d[indiceCol["ZoneReference"]] == "Mixte"]

        # ligne pour les statistiques de reference Ultramafique
        Line_UMMeth = ["Gamme de reference - methodes analytiques", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM,
                         None, Periode_UM]
        Line_UMLQ = ["Gamme de reference - Limites quantitatives", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM,
                       None,
                       Periode_UM]
        Line_UMSigne = ["Gamme de reference - Nb LQ", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMN = ["Gamme de reference - N", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                      Periode_UM]
        Line_UMMoy = ["Gamme de reference - Moyenne", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                        Periode_UM]
        Line_UMEct = ["Gamme de reference - Ecart-Type", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                        Periode_UM]
        Line_UMMed = ["Gamme de reference - Mediane", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                        Periode_UM]
        Line_UMPct10 = ["Gamme de reference - Percentile 10", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct25 = ["Gamme de reference - Percentile 25", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct75 = ["Gamme de reference - Percentile 75", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct80 = ["Gamme de reference - Percentile 80", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct85 = ["Gamme de reference - Percentile 85", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct90 = ["Gamme de reference - Percentile 90", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMPct95 = ["Gamme de reference - Percentile 95", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                          Periode_UM]
        Line_UMMin = ["Gamme de reference - Min", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                        Periode_UM]
        Line_UMMax = ["Gamme de reference - Max", "Reference", "", "", "Ultramafique", "Ultramafique", StatTxt_UM, None,
                        Periode_UM]
        
        # ligne pour les statistiques de reference Mixte
        Line_MixteMeth = ["Gamme de reference - methodes analytiques", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte,
                         None, Periode_Mixte]
        Line_MixteLQ = ["Gamme de reference - Limites quantitatives", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte,
                       None,
                       Periode_Mixte]
        Line_MixteSigne = ["Gamme de reference - Nb LQ", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixteN = ["Gamme de reference - N", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                      Periode_Mixte]
        Line_MixteMoy = ["Gamme de reference - Moyenne", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                        Periode_Mixte]
        Line_MixteEct = ["Gamme de reference - Ecart-Type", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                        Periode_Mixte]
        Line_MixteMed = ["Gamme de reference - Mediane", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                        Periode_Mixte]
        Line_MixtePct10 = ["Gamme de reference - Percentile 10", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct25 = ["Gamme de reference - Percentile 25", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct75 = ["Gamme de reference - Percentile 75", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct80 = ["Gamme de reference - Percentile 80", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct85 = ["Gamme de reference - Percentile 85", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct90 = ["Gamme de reference - Percentile 90", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixtePct95 = ["Gamme de reference - Percentile 95", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                          Periode_Mixte]
        Line_MixteMin = ["Gamme de reference - Min", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                        Periode_Mixte]
        Line_MixteMax = ["Gamme de reference - Max", "Reference", "", "", "Mixte", "Mixte", StatTxt_Mixte, None,
                        Periode_Mixte]
        
        for p in Parametres:
            i = Parametres.index(p) + indiceCol["PremierParam"]
            # methodes analytiques _UM
            Line_UMMeth.append(FindMethodes(DonneesM_UM,i))
            ## VALEUR LQ
            # _UM
            Line_UMLQ.append(FindLQ(DonneesS_UM,i,DonneesEx_UM))
            # NOMBRE DE VALEURS = A LA LQ
            #_UM
            Line_UMSigne.append(FindSigne(DonneesS_UM,i))

            # DONNEES NON NULLES OU NON VIDES RECUPEREES + VALEURS MESUREES N
            # _UM
            VN_UM = FindValsNonVides(DonneesEx_UM, i)
            valN_UM = VN_UM[0]
            ValsEx_UM = VN_UM[1]
            Line_UMN.append(valN_UM)
            # _UM
            if valN_UM == 0:
                Line_UMMoy.append("")
                Line_UMEct.append("")
                Line_UMMed.append("")
                Line_UMPct10.append("")
                Line_UMPct25.append("")
                Line_UMPct75.append("")
                Line_UMPct80.append("")
                Line_UMPct85.append("")
                Line_UMPct90.append("")
                Line_UMPct95.append("")
                Line_UMMin.append("")
                Line_UMMax.append("")
            else:
                Calc_UM = FindStatsCal(ValsEx_UM)
                Line_UMMoy.append(Calc_UM["moy"])
                Line_UMEct.append(Calc_UM["ect"])
                Line_UMMed.append(Calc_UM["med"])
                Line_UMPct10.append(Calc_UM["perc10"])
                Line_UMPct25.append(Calc_UM["perc25"])
                Line_UMPct75.append(Calc_UM["perc75"])
                Line_UMPct80.append(Calc_UM["perc80"])
                Line_UMPct85.append(Calc_UM["perc85"])
                Line_UMPct90.append(Calc_UM["perc90"])
                Line_UMPct95.append(Calc_UM["perc95"])
                # Min et Max
                Line_UMMin.append(Calc_UM["min"])
                Line_UMMax.append(Calc_UM["max"])
            ###MIXTE
            # methodes analytiques _Mixte
            Line_MixteMeth.append(FindMethodes(DonneesM_Mixte,i))
            ## VALEUR LQ
            # _Mixte
            Line_MixteLQ.append(FindLQ(DonneesS_Mixte,i,DonneesEx_Mixte))
            # NOMBRE DE VALEURS = A LA LQ
            #_Mixte
            Line_MixteSigne.append(FindSigne(DonneesS_Mixte,i))

            # DONNEES NON NULLES OU NON VIDES RECUPEREES + VALEURS MESUREES N
            # _Mixte
            VN_Mixte = FindValsNonVides(DonneesEx_Mixte, i)
            valN_Mixte = VN_Mixte[0]
            ValsEx_Mixte = VN_Mixte[1]
            Line_MixteN.append(valN_Mixte)
            # _Mixte
            if valN_Mixte == 0:
                Line_MixteMoy.append("")
                Line_MixteEct.append("")
                Line_MixteMed.append("")
                Line_MixtePct10.append("")
                Line_MixtePct25.append("")
                Line_MixtePct75.append("")
                Line_MixtePct80.append("")
                Line_MixtePct85.append("")
                Line_MixtePct90.append("")
                Line_MixtePct95.append("")
                Line_MixteMin.append("")
                Line_MixteMax.append("")
            else:
                Calc_Mixte = FindStatsCal(ValsEx_Mixte)
                Line_MixteMoy.append(Calc_Mixte["moy"])
                Line_MixteEct.append(Calc_Mixte["ect"])
                Line_MixteMed.append(Calc_Mixte["med"])
                Line_MixtePct10.append(Calc_Mixte["perc10"])
                Line_MixtePct25.append(Calc_Mixte["perc25"])
                Line_MixtePct75.append(Calc_Mixte["perc75"])
                Line_MixtePct80.append(Calc_Mixte["perc80"])
                Line_MixtePct85.append(Calc_Mixte["perc85"])
                Line_MixtePct90.append(Calc_Mixte["perc90"])
                Line_MixtePct95.append(Calc_Mixte["perc95"])
                # Min et Max
                Line_MixteMin.append(Calc_Mixte["min"])
                Line_MixteMax.append(Calc_Mixte["max"])
        #######################INSERTION DES DONNEES###################################
        # Insertion Donnees _UM
        DonneesStatGammeRef.append(Line_UMMeth)
        DonneesStatGammeRef.append(Line_UMLQ)
        DonneesStatGammeRef.append(Line_UMSigne)
        DonneesStatGammeRef.append(Line_UMN)
        DonneesStatGammeRef.append(Line_UMMoy)
        DonneesStatGammeRef.append(Line_UMEct)
        DonneesStatGammeRef.append(Line_UMMed)
        DonneesStatGammeRef.append(Line_UMPct10)
        DonneesStatGammeRef.append(Line_UMPct25)
        DonneesStatGammeRef.append(Line_UMPct75)
        DonneesStatGammeRef.append(Line_UMPct80)
        DonneesStatGammeRef.append(Line_UMPct85)
        DonneesStatGammeRef.append(Line_UMPct90)
        DonneesStatGammeRef.append(Line_UMPct95)
        DonneesStatGammeRef.append(Line_UMMin)
        DonneesStatGammeRef.append(Line_UMMax)
        # Insertion Donnees _Mixte
        DonneesStatGammeRef.append(Line_MixteMeth)
        DonneesStatGammeRef.append(Line_MixteLQ)
        DonneesStatGammeRef.append(Line_MixteSigne)
        DonneesStatGammeRef.append(Line_MixteN)
        DonneesStatGammeRef.append(Line_MixteMoy)
        DonneesStatGammeRef.append(Line_MixteEct)
        DonneesStatGammeRef.append(Line_MixteMed)
        DonneesStatGammeRef.append(Line_MixtePct10)
        DonneesStatGammeRef.append(Line_MixtePct25)
        DonneesStatGammeRef.append(Line_MixtePct75)
        DonneesStatGammeRef.append(Line_MixtePct80)
        DonneesStatGammeRef.append(Line_MixtePct85)
        DonneesStatGammeRef.append(Line_MixtePct90)
        DonneesStatGammeRef.append(Line_MixtePct95)
        DonneesStatGammeRef.append(Line_MixteMin)
        DonneesStatGammeRef.append(Line_MixteMax)
    else:
        arcpy.AddMessage("Ouuueeeee  Pas de gamme de reference demandee pour le gros matou :-)")
        LineVideGamme = ["Pas de gamme de reference attendue","","","","","","",None,""]
        for p in Params:
            LineVideGamme.append("")
        DonneesStatGammeRef.append(LineVideGamme)

    return DonneesStatGammeRef

###Fonction pour trouver les valeurs superieures a des percentiles#########
###la ou les valeurs sur valcol sont celles qui servent de reference ou il y a les valeurs et l'indcol correspond au nom de la colonne Ex: ZoneReference
def NbSupPerc(StGP,indcol, valcol, percRef, ind, ValsEx, ValN):
    #indcol et valcol sont des listes
    GP = list()
    GP.extend(StGP)
    for indc in indcol:
        ii = indcol.index(indc)
        GP = [gp for gp in GP if gp[indiceCol[indc]] == valcol[ii]]
    if len(GP) > 0:
        P = [gp[ind] for gp in GP if gp[indiceCol["TypeData"]] == percRef]
        if len(P) > 0:
            P = P[0]
        else:
            P = ""
    else:
        P = ""
    if P == "":
        NbSupP = ""
        PctSupP = ""
    else:
        P = float(P.replace(",", "."))
        NbSupP = len([val for val in ValsEx if val > P])
        PctSupP = NbSupP * 100 / ValN
    return {"NbS" : str(NbSupP), "PctS" : str(PctSupP) }

################################FONCTION PRINCIPALE POUR METRIQUES (STATISTIQUES SELON LA PERIODE (ANNEE, SAISON))##############################################
def StatistiquesPeriode(DonneesMethodes,DonneesSignes,DonneesExp,Parametres,StatistGP,IdentSuiv):
    #on conserve toutes les infos de station pour conserver ensuite l'ordre
    StatsDts = sorted(list(set([d[indiceCol["TypeStation"]] + ";" + d[indiceCol["Localisation"]]+ ";" + d[indiceCol["TypeBV"]] + ";" + d[indiceCol["Zone"]]+ ";" + d[indiceCol["ZoneReference"]] + ";" + d[indiceCol["Station"]] for d in DonneesMethodes])))
    StatsDts = [s.split(";") for s in StatsDts]
    Stats = [s[indiceCol["Station"]-1] for s in StatsDts]
    SortieDonnees = list()
    for s in Stats:
        StZoneRef = [std[indiceCol["ZoneReference"]-1] for std in StatsDts if s == std[indiceCol["Station"]-1]][0]
        StLoc = [std[indiceCol["Localisation"]-1] for std in StatsDts if s == std[indiceCol["Station"]-1]][0]
        StBV = [std[indiceCol["TypeBV"]-1] for std in StatsDts if s == std[indiceCol["Station"]-1]][0]
        nbS = len(Stats)
        sti = Stats.index(s)
        pct25 = round(nbS/4)
        pct50 = round(nbS/2)
        pct75 = pct25 + pct50
        pct100 = nbS - 1
        if sti == pct25:
            arcpy.AddMessage("Ca avance Le gros matou a traite deja 25% des metriques traitees")
        if sti == pct50:
            arcpy.AddMessage("Deja au milieu... Le gros matou a traite 50% des metriques traitees")
        if sti == pct75:
            arcpy.AddMessage("Ca avance bien Le gros matou a traite deja 75% des metriques traitees")
        if sti == pct100:
            arcpy.AddMessage("Waouw Le gros matou a traite 100% des metriques traitees")

        DonneesMethStat = [d for d in DonneesMethodes if d[indiceCol["Station"]] == s]
        DonneesSignStat = [d for d in DonneesSignes if d[indiceCol["Station"]] == s]
        DonneesExpStat = [d for d in DonneesExp if d[indiceCol["Station"]] == s]
        Periodes = [d[8] for d in DonneesMethStat]
        Periodes = list(set(Periodes))
        Periodes.sort()
        for per in Periodes:
            DonneesMthStPer = [d for d in DonneesMethStat if d[indiceCol["Periode"]] == per]
            DonneesSiStPer = [d for d in DonneesSignStat if d[indiceCol["Periode"]] == per]
            DonneesExpStPer = [d for d in DonneesExpStat if d[indiceCol["Periode"]] == per]
            d = DonneesExpStPer[0]
            LineMeth = ["Metrique - methodes analytiques",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSign = ["Metrique - Nb LQ",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineLQ = ["Metrique - Limites quantitatives",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineN = ["Metrique - N",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineMoy = ["Metrique - Moyenne",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineEct = ["Metrique - Ecart-type",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineMed = ["Metrique - Mediane", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc10 = ["Metrique - Percentile 10",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePerc25 = ["Metrique - Percentile 25", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc75 = ["Metrique - Percentile 75", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc80 = ["Metrique - Percentile 80", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc85 = ["Metrique - Percentile 85", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc90 = ["Metrique - Percentile 90", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LinePerc95 = ["Metrique - Percentile 95", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LineMin = ["Metrique - Min", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LineMax = ["Metrique - Max", d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]], None, per]
            LineSupPerc10 = ["Metrique - Superieur Perc. 10",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc10 = ["Metrique - Superieur Perc. 10 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc25 = ["Metrique - Superieur Perc. 25",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc25 = ["Metrique - Superieur Perc. 25 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc75 = ["Metrique - Superieur Perc. 75",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc75 = ["Metrique - Superieur Perc. 75 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc80 = ["Metrique - Superieur Perc. 80",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc80 = ["Metrique - Superieur Perc. 80 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc85 = ["Metrique - Superieur Perc. 85",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc85 = ["Metrique - Superieur Perc. 85 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc90 = ["Metrique - Superieur Perc. 90",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc90 = ["Metrique - Superieur Perc. 90 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LineSupPerc95 = ["Metrique - Superieur Perc. 95",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            LinePctSupPerc95 = ["Metrique - Superieur Perc. 95 (%)",d[indiceCol["TypeStation"]],d[indiceCol["Localisation"]],d[indiceCol["TypeBV"]],d[indiceCol["Zone"]],d[indiceCol["ZoneReference"]],d[indiceCol["Station"]],None,per]
            for p in Parametres:
                i = Parametres.index(p) + indiceCol["PremierParam"]
                #methodes analytiques
                LineMeth.append(FindMethodes(DonneesMthStPer,i))
                #nombre de valeurs a la LQ
                LineSign.append(FindSigne(DonneesSiStPer,i))
                #valeurs LQ
                LineLQ.append(FindLQ(DonneesSiStPer,i,DonneesExpStPer))
                #recuperation des valeurs non vides ou nulles pour les prochaines statistiques
                VN = FindValsNonVides(DonneesExpStPer, i)
                valN = VN[0]
                ValsExpStPer = VN[1]
                LineN.append(valN)
                if valN == 0:
                    LineMoy.append("")
                    LineEct.append("")
                    LineMed.append("")
                    LinePerc10.append("")
                    LinePerc25.append("")
                    LinePerc75.append("")
                    LinePerc80.append("")
                    LinePerc85.append("")
                    LinePerc90.append("")
                    LinePerc95.append("")
                    LineMin.append("")
                    LineMax.append("")
                    LineSupPerc10.append("")
                    LinePctSupPerc10.append("")
                    LineSupPerc25.append("")
                    LinePctSupPerc25.append("")
                    LineSupPerc75.append("")
                    LinePctSupPerc75.append("")
                    LineSupPerc80.append("")
                    LinePctSupPerc80.append("")
                    LineSupPerc85.append("")
                    LinePctSupPerc85.append("")
                    LineSupPerc90.append("")
                    LinePctSupPerc90.append("")
                    LineSupPerc95.append("")
                    LinePctSupPerc95.append("")
                else:
                    Calc = FindStatsCal(ValsExpStPer)
                    LineMoy.append(Calc["moy"])
                    LineEct.append(Calc["ect"])
                    LineMed.append(Calc["med"])
                    LinePerc10.append(Calc["perc10"])
                    LinePerc25.append(Calc["perc25"])
                    LinePerc75.append(Calc["perc75"])
                    LinePerc80.append(Calc["perc80"])
                    LinePerc85.append(Calc["perc85"])
                    LinePerc90.append(Calc["perc90"])
                    LinePerc95.append(Calc["perc95"])
                    # Min et Max
                    LineMin.append(Calc["min"])
                    LineMax.append(Calc["max"])
                    #valeur superieure au percentile75
                    #Eaux de surface ou eaux souterraines
                    if IdentSuiv == "SURF":
                        Perc10 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 10" , i, ValsExpStPer, valN)
                        Perc25 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 25" , i, ValsExpStPer, valN)
                        Perc75 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 75" , i, ValsExpStPer, valN)
                        Perc80 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 80" , i, ValsExpStPer, valN)
                        Perc85 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 85" , i, ValsExpStPer, valN)
                        Perc90 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 90" , i, ValsExpStPer, valN)
                        Perc95 = NbSupPerc(StatistGP, ["ZoneReference"], [StZoneRef], "Gamme de reference - Percentile 95" , i, ValsExpStPer, valN)                  
                        LineSupPerc10.append(Perc10["NbS"])
                        LinePctSupPerc10.append(Perc10["PctS"])
                        LineSupPerc25.append(Perc25["NbS"])
                        LinePctSupPerc25.append(Perc25["PctS"])
                        LineSupPerc75.append(Perc75["NbS"])
                        LinePctSupPerc75.append(Perc75["PctS"])
                        LineSupPerc80.append(Perc80["NbS"])
                        LinePctSupPerc80.append(Perc80["PctS"])
                        LineSupPerc85.append(Perc85["NbS"])
                        LinePctSupPerc85.append(Perc85["PctS"])
                        LineSupPerc90.append(Perc90["NbS"])
                        LinePctSupPerc90.append(Perc90["PctS"])
                        LineSupPerc95.append(Perc95["NbS"])
                        LinePctSupPerc95.append(Perc95["PctS"])
                    else:
                        LineSupPerc10.append("")
                        LinePctSupPerc10.append("")
                        LineSupPerc25.append("")
                        LinePctSupPerc25.append("")
                        LineSupPerc75.append("")
                        LinePctSupPerc75.append("")
                        LineSupPerc80.append("")
                        LinePctSupPerc80.append("")
                        LineSupPerc85.append("")
                        LinePctSupPerc85.append("")
                        LineSupPerc90.append("")
                        LinePctSupPerc90.append("")
                        LineSupPerc95.append("")
                        LinePctSupPerc95.append("")

            SortieDonnees.append(LineMeth)
            SortieDonnees.append(LineLQ)
            SortieDonnees.append(LineSign)
            SortieDonnees.append(LineN)
            SortieDonnees.append(LineMoy)
            SortieDonnees.append(LineEct)
            SortieDonnees.append(LineMed)
            SortieDonnees.append(LinePerc10)
            SortieDonnees.append(LinePerc25)
            SortieDonnees.append(LinePerc75)
            SortieDonnees.append(LinePerc80)
            SortieDonnees.append(LinePerc85)
            SortieDonnees.append(LinePerc90)
            SortieDonnees.append(LinePerc95)
            SortieDonnees.append(LineMin)
            SortieDonnees.append(LineMax)
            SortieDonnees.append(LineSupPerc10)
            SortieDonnees.append(LinePctSupPerc10)
            SortieDonnees.append(LineSupPerc25)
            SortieDonnees.append(LinePctSupPerc25)
            SortieDonnees.append(LineSupPerc75)
            SortieDonnees.append(LinePctSupPerc75)
            SortieDonnees.append(LineSupPerc80)
            SortieDonnees.append(LinePctSupPerc80)
            SortieDonnees.append(LineSupPerc85)
            SortieDonnees.append(LinePctSupPerc85)
            SortieDonnees.append(LineSupPerc90)
            SortieDonnees.append(LinePctSupPerc90)
            SortieDonnees.append(LineSupPerc95)
            SortieDonnees.append(LinePctSupPerc95)
            
            del LineMeth, LineSign,LineLQ, LineN, LineEct, LineMed, LinePerc10, LinePerc25, LinePerc90, LinePerc75, LinePerc80, LinePerc85, LinePerc95
            del DonneesSiStPer, DonneesMthStPer, DonneesExpStPer
        del DonneesMethStat, DonneesSignStat, Periodes, DonneesExpStat
    del DonneesMethodes, DonneesSignes, Parametres,StatsDts, Stats
    return SortieDonnees

#FONCTION FORMATAGE CELLULE
#Fichier Excel, Oui ou Non, numerique pour la taille, police texte, couleur police, couleur de fond
def FormatExcel(FExcel,bold,taille,police,couleur,couleurfd,formatcell):
    FormF = FExcel.add_format()
    if bold == "Oui":
        FormF.set_bold()
    if couleurfd != "":
        FormF.set_bg_color(couleurfd)
    FormF.set_font_color(couleur)
    FormF.set_font(police)
    FormF.set_font_size(taille)
    if formatcell != "":
        FormF.set_num_format(formatcell)
    return FormF

#insertion de ligne de donnees sur Excel
def LigneExcelValeur(formDt, formN, formDate, feuilleX, ligne, Ind):
    colX = 0
    for L in ligne:
        if colX == 0:
            feuilleX.write(Ind, colX, L, formDt)
        elif colX == indiceCol["Date"]:
            feuilleX.write(Ind, colX, L, formDate)
        else:
            if str(type(L)) == "<type 'int'>" or str(type(L)) == "<type 'float'>":
                feuilleX.write_number(Ind, colX, L, formN)
            elif L != "" and L.find("|") == -1 and L.find("<") == -1:
                L = L.replace(",", ".")
                if L.find(".") > 0:
                    L = float(L)
                    feuilleX.write_number(Ind, colX, L, formN)
                else:
                    feuilleX.write(Ind, colX, L, formN)
            else:
                feuilleX.write(Ind, colX, L, formN)
        colX = colX + 1

        
##FONCTION  pour la sortie XLSX Metriques
def ExportMetrique(DonneesPeriodes, DonneesGamme,Parametres,XMLParamEquiv, XMLMetriqueEquiv, XMLGammeEquiv, idsuivi, metrAttendu, GRAttendu, FichierXML, OrdreStation, TypeSuivi):
    if idsuivi == "SURF" or idsuivi == "SOUT" or idsuivi == "SEDI":
        arcpy.AddMessage("Creation des informations pour le fichier de metrique XLSX")
        #on va chercher les 5 dernieres annees par l'annee en cours - 1
        #Exemple Annee en cours = 2019 on recupere 2018 à 2014
        MetXLSX = xlsxwriter.Workbook(FichierXML.replace(".xml",".xlsx"))
        ##FORMES des cellules
        #Titre
        FormTitle = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "", "")
        FormTitle.set_border(5)
        FormTitle.set_align("center")
        FormTitle.set_align("vcenter")
        FormTitleNote = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#bdd7ee", "")
        FormTitleNote.set_border(5)
        FormTitleNote.set_align("center")
        FormTitleNote.set_align("vcenter")
        FormTitleNote.set_text_wrap()
        #Normal
        FormN = FormatExcel(MetXLSX, "Non", 9, "Calibri Light", "#000000", "", "")
        FormN.set_align("center")
        FormN.set_align("vcenter")
        FormN.set_border(1)
        #Station
        FormSt = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#fce4d6", "")
        FormSt.set_align("center")
        FormSt.set_align("vcenter")
        FormSt.set_border(5)
        #Station Reference
        FormStR = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#E2A2F9", "")
        FormStR.set_align("center")
        FormStR.set_align("vcenter")
        FormStR.set_border(5)
        #cellule numerique
        FormNum = FormatExcel(MetXLSX, "Non", 9, "Calibri Light", "#000000", "", "")
        FormNum.set_align("center")
        FormNum.set_align("vcenter")
        #FormNum.set_num_format("0.000")
        FormNum.set_border(1)
        # cellule Fortement perturbe
        FormFortPerturb = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#e6b8b7", "")
        FormFortPerturb.set_align("center")
        FormFortPerturb.set_align("vcenter")
        # cellule Non perturbe
        FormNonPerturb = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#d8e4bc", "")
        FormNonPerturb.set_align("center")
        FormNonPerturb.set_align("vcenter")
        # cellule Pas de seuil
        FormNonSeuil = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#FFD966", "")
        FormNonSeuil.set_align("center")
        FormNonSeuil.set_align("vcenter")
        #An = int(datetime.datetime.now().strftime("%Y"))
        An = AnneeEtude
        Annees = [An - 4, An - 3, An - 2, An - 1, An]
        arcpy.AddMessage("Metriques attendues:")
        arcpy.AddMessage(metrAttendu)
        arcpy.AddMessage("Periodes attendues:")
        arcpy.AddMessage(Annees)
        arcpy.AddMessage("Gammes attendues:")
        arcpy.AddMessage(GRAttendu)
        #on filtre les donnees de metrique seulement sur les 5 dernieres annees et metriques voulues metrAttendu
        arcpy.AddMessage(str(len(DonneesPeriodes)) + " donnees de periodes entrees")
        DonneesPeriodesVoulues = [d for d in DonneesPeriodes if int(str(d[indiceCol["Periode"]])[-4:]) in Annees and d[indiceCol["TypeData"]] in metrAttendu]
        arcpy.AddMessage(str(len(DonneesPeriodesVoulues)) + " donnees sur les periodes et metriques attendues")
        #Liste unique de station - periode
        StationPer = list(set([dpa[indiceCol["TypeStation"]] + ";" + dpa[indiceCol["Localisation"]] + ";" + dpa[indiceCol["TypeBV"]] + ";" + dpa[indiceCol["Zone"]] + ";" +
                               dpa[indiceCol["Station"]] + ";" + dpa[indiceCol["Periode"]] for dpa in DonneesPeriodesVoulues]))
        StationPer = [stp.split(";") for stp in StationPer]
        arcpy.AddMessage(str(len(StationPer)) + " combinaison unique station-periode")
        #Liste unique de station
        Station = list(set([st[indiceCol["TypeStation"]-1] + ";" + st[indiceCol["Localisation"]-1] + ";" + st[indiceCol["TypeBV"]-1] + ";" + st[indiceCol["Zone"]-1] + ";" +
                                st[indiceCol["Station"]-2] for st in StationPer]))
        #On reorganise les stations selon l'ordre voulu grace OrdreStation avec seulement les stations de suivi (pas de reference)...
        #... on implemente egalement les stations de suivi n'ayant pas de donnee
        NbOrdreStat = len(OrdreStation)
        iord = 1
        StationOrdonnee = list()
        Station = [st.split(";") for st in Station]
        while iord < NbOrdreStat + 1:
            if TypeSuivi == "Reference":
                StatO = [os[1] for os in OrdreStation if os[9] == iord]
            else:
                StatO = [os[1] for os in OrdreStation if os[8] == iord]
            #si seulement on trouve une correspondance avec l'OrdreStation
            if len(StatO) > 0:
                StatO = StatO[0]
                StationSel = [s for s in Station if s[4] == StatO]
                if len(StationSel) > 0:
                    StationSel = StationSel[0]
                    StationOrdonnee.append(StationSel)
                else:
                    StatO = [os for os in OrdreStation if os[8] == iord]
                    StationSel = [[st[2],st[3],st[4],st[5],st[1]] for st in StatO]
                    if len(StationSel) > 0:
                        StationSel = StationSel[0]
                        StationOrdonnee.append(StationSel)
            iord += 1
        Station = [sto for sto in StationOrdonnee]
        del StationOrdonnee
        #Fin Ordre des stations et virer stations de reference
        arcpy.AddMessage(str(len(Station)) + " Stations concernees")
        #Equivalence Parametre
        EquivParam = [[ep[0],ep[1],ep[2],ep[3]] for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite","SeuilTHIO"],"id_cat2 = '" + idsuivi + "'")]
        #Equivalence Metrique
        EquivMetr = [[em[0],em[1],em[2]] for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml","metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
        # Equivalence Gamme de reference
        EquivGamme = [list(em) for em in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
        nbmet = len(metrAttendu)
        EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu]
        #dictionnaire Annee indice colonne fusion pour la ligne Annees sous Excel par feuille Excel de parametre
        AnInd = dict()
        xa = 0
        for an in Annees:
            AnInd[str(an)] = [xa, xa + nbmet - 1]
            xa = xa + nbmet
        arcpy.AddMessage(AnInd)
        AnMet = dict()
        #dictionnaire AnMet pour placer le titre des metrique par annee pour les annees normales et AnMet3 pour les metriques sur les 3 dernieres annees avec mise en forme
        met_init = 3
        for an in Annees:
            for met2 in EMetr:
                im = EMetr.index(met2)
                AnMet[an,met2[0]] = [2,met_init+im+AnInd[str(an)][0],met2[0],FormTitle]
        Metr = dict()
        #dictionnaire avec comme valeurs metrique pour XML balise (metr) et metrique pour le nom en Excel (metrXL)
        for ma in metrAttendu:
            metrXL =[met[2] for met in EquivMetr if met[0] == ma]
            if len(metrXL) > 0:
                Metr[ma] = metrXL[0]
            else:
                arcpy.AddMessage("Correspondance non trouvee pour la metrique attendue: " + ma)
                Metr[ma] = ma
        #dictionnaire avec parametre en cle nom parametre et valeurs (nom et unite)
        Paramdict = dict()
        for p in Parametres:
            nameparam = [np[1] for np in EquivParam if np[0] == p[1]]
            if len(nameparam) > 0:
                nameparam = nameparam[0]
                unitparam = [np[2] for np in EquivParam if np[0] == p[1]]
                seuilparam = [np[3] for np in EquivParam if np[0] == p[1]]
                if len(unitparam) > 0:
                    unitparam = unitparam[0]
                    seuilparam = seuilparam[0]
                else:
                    arcpy.AddMessage("Correspondance non trouvee pour l'unite de parametre: " + p[1])
                    unitparam = ""
            else:
                arcpy.AddMessage("Correspondance non trouvee pour le parametre: " + p[1])
                nameparam = p[1]
                unitparam = ""
                seuilparam = ""
            Paramdict[p[1]] = [nameparam,unitparam,seuilparam]
        #Donnees Gamme de reference#
        DonneesGammeVoulues = [d for d in DonneesGamme if d[0] in GRAttendu]
        TypeRef = list(set([dtgr[indiceCol["Zone"]] for dtgr in DonneesGammeVoulues]))
        #dictionnaire de periode par station avec comme valeurs periodes et stations en cle
        PerStdict = dict()
        for s in Station:
                PerStdict[s[indiceCol["Station"]-2]] = sorted(list(set([ast[indiceCol["Periode"]-3] for ast in StationPer if ast[indiceCol["Station"]-2] == s[indiceCol["Station"]-2]])))
        for p in Parametres:
            #seuil du parametre pour note reference
            seuilP = Paramdict[p[1]][2]
            #seuil#
            #colonne Note comparaison initialise a 0
            notex = 0
            ip = Parametres.index(p)
            arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            ipdt = ip+ indiceCol["PremierParam"]
            Fparam = MetXLSX.add_worksheet(Paramdict[p[1]][0])
            Fparam.merge_range(0,0,0,2, Paramdict[p[1]][0] + " (" + Paramdict[p[1]][1] + ")", FormTitle)
            if TypeSuivi == "Suivi":
                Fparam.merge_range(1,0,1,2, "Station de suivi", FormSt)
            else:
                Fparam.merge_range(1, 0, 1, 2, "Station de reference", FormStR)
            Fparam.write(2,0,"Zone", FormTitle)
            Fparam.write(2,1,"Geologie", FormTitle)
            Fparam.write(2,2,"Station", FormTitle)
            # Longueur des colonnes
            Fparam.set_column(0, 2, 20)
            Fparam.set_column(3, 32, 10)
            Fparam.set_column(33, 34, 18)
            Fparam.set_column(35, 38, 10)
            Fparam.set_column(39, 39, 75)
            Fparam.set_column(40, 45, 10)
            #reference a mettre en forme
            # On met en forme les cellules de valeur avant d'integrer
            LigneFin = len(Station) + 2
            ColonneFin = AnInd[str(AnneeEtude)][1] + 3
            ldeb = 3
            while ldeb <= LigneFin:
                cdeb = 3
                while cdeb <= ColonneFin:
                    Fparam.write(ldeb, cdeb, "", FormN)
                    cdeb += 1
                ldeb += 1
            sx = 3
            # rappel met_init = 3
            for an in Annees:
                if AnInd[str(an)][0] == AnInd[str(an)][1]:
                    Fparam.write(1,AnInd[str(an)][0]+met_init,str(an),FormTitle)
                else:
                    Fparam.merge_range(1,AnInd[str(an)][0] + met_init, 1, AnInd[str(an)][1] + met_init, str(an), FormTitle)
                for met2 in EMetr:
                    im = EMetr.index(met2)
                    Fparam.write(AnMet[(an, met2[0])][0],AnMet[(an, met2[0])][1], AnMet[(an, met2[0])][2],AnMet[(an, met2[0])][3])
            for s in Station:
                Fparam.write(sx,0, s[indiceCol["TypeBV"]-1],FormN)
                Fparam.write(sx,1, s[indiceCol["Zone"]-1],FormN)
                Fparam.write(sx,2, s[indiceCol["Station"]-2],FormN)
                for perst in PerStdict[s[indiceCol["Station"]-2]]:
                    if len(str(perst)) > 4:
                        if perst.find("Saison chaude") >= 0:
                            perstxml = "sc_" + perst[-4:]
                        elif perst.find("Saison fraiche") >= 0:
                            perstxml = "sf_" + perst[-4:]
                        elif perst[0] == "2":
                            perstxml = perst
                        else:
                            perstxml = perst[-4:]
                    else:
                        perstxml = perst
                    for ma in EMetr:
                        im = EMetr.index(ma)
                        valeur = [dpv[ipdt] for dpv in DonneesPeriodesVoulues if dpv[indiceCol["TypeData"]] == ma[1] and dpv[indiceCol["TypeStation"]] == s[indiceCol["TypeStation"]-1]
                                  and dpv[indiceCol["Localisation"]] == s[indiceCol["Localisation"]-1] and dpv[indiceCol["TypeBV"]] == s[indiceCol["TypeBV"]-1]
                                  and dpv[indiceCol["Zone"]] == s[indiceCol["Zone"]-1] and dpv[indiceCol["Station"]] == s[indiceCol["Station"]-2]
                                  and dpv[indiceCol["Periode"]] == perst]
                        if len(valeur) > 0:
                            valeur = valeur[0]
                            if valeur == "" or valeur == None:
                                valeur = ""
                        else:
                            valeur = ""
                        valeur = str(valeur).replace(".",",")
                        metx = im + met_init + AnInd[str(perstxml)][0]
                        if valeur != "" and valeur.find("|") == -1:
                            valeur = valeur.replace(",", ".")
                            valeur = float(valeur)
                            Fparam.write_number(sx,metx,valeur,FormNum)
                            ###NOTE COMPARAISON REFERENCE
                            if ma[0] == "% > Perc 75" and int(perst) == AnneeEtude:
                                typeseuilP = str(type(seuilP))
                                if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                    if valeur >= seuilP:
                                        Fparam.write(sx, metx+1, "Fortement perturbe", FormFortPerturb)
                                        notex = metx+1
                                    else:
                                        Fparam.write(sx, metx+1, "Non perturbe", FormNonPerturb)
                                        notex = metx + 1
                                else:
                                    Fparam.write(sx, metx+1, "Pas de seuil", FormNonSeuil)
                                    notex = metx + 1
                        else:
                            Fparam.write(sx,metx,valeur,FormNum)
                sx += 1
                #arcpy.AddMessage("Station traitee " + s[indiceCol["Station"]-2]  + " pour le fichier XML")
            arcpy.AddMessage("Parametre " + p[1] + " traite pour le fichier de metrique XML: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            if notex > 0:
                Fparam.merge_range(1, notex, 2, notex, "Note comparaison reference", FormTitleNote)
            else:
                Fparam.merge_range(1,Fparam.dim_colmax+1, 2, Fparam.dim_colmax+1, "Note comparaison reference", FormTitleNote)
            Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Evolution temporelle", FormTitle)
            Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2018", FormTitle)
            Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2019", FormTitle)
            Fparam.set_column(0,Fparam.dim_colmax,18)
            ##############################GAMME DE REFERENCE#################################################
            debColGamme = Fparam.dim_colmax + 2
            Fparam.merge_range(0, debColGamme, 0, debColGamme + 4, "Gamme de reference", FormTitleNote)
            Fparam.write(1, debColGamme, "Geologie", FormTitle)
            Fparam.write(1, debColGamme + 1, "Stations", FormTitle)
            Fparam.write(1, debColGamme + 2, "Periode", FormTitle)
            refx = 2
            for ref in TypeRef:
                DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceCol["Zone"]] == ref]
                Fparam.write(refx, debColGamme, ref, FormN)
                Fparam.write(refx, debColGamme + 1, DonneesGVSelect[0][indiceCol["Station"]], FormN)
                Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceCol["Periode"]], FormN)
                grx = debColGamme + 3
                for gr in GRAttendu:
                    DonneesGVSel = [dgv for dgv in DonneesGVSelect if dgv[indiceCol["TypeData"]] == gr]
                    gamXL = [gm[2] for gm in EquivGamme if gm[0] == gr]
                    if len(gamXL) > 0:
                        gamXL = gamXL[0]
                    else:
                        arcpy.AddMessage("Correspondance non trouvee pour la gamme attendue: " + gr)
                        gamXL = gr
                    valeur = [dga[ipdt] for dga in DonneesGVSel]
                    if len(valeur) > 0:
                        valeur = valeur[0]
                        if valeur == "" or valeur == None:
                            valeur = ""
                    else:
                        valeur = ""
                    valeur = str(valeur).replace(".", ",")
                    Fparam.write(1, grx, gamXL, FormTitle)
                    if valeur != "" and valeur.find("|") == -1:
                        valeur = valeur.replace(",",".")
                        valeur = float(valeur)
                        Fparam.write_number(refx, grx, valeur, FormNum)
                    else:
                        Fparam.write(refx, grx, valeur, FormNum)
                    grx += 1
                refx += 1
            #On ajoute a la fin seuil parametre
            Fparam.write(1,grx,"Seuil",FormTitle)
            typeseuilP = str(type(seuilP))
            if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                Fparam.write_number(2,grx,seuilP,FormNum)
            else:
                Fparam.write(2,grx,seuilP,FormNum)
        #POUR TOUS LES CAS DE SUIVIS
        MetXLSX.close()
    else:
        arcpy.AddMessage("Aucune donnee pour le fichier Excel de metrique")
    ###METTRE elif ICI pour les autres type de suivi SOUT ...


################################EXECUTION / TRAITEMENT##############################################
arcpy.AddMessage("le gros matou des Eaux Douces se reveille et vous dire BONJOUR, il commence a bosser. Alors il recupere l'identifiant de la categorie correspondant au Suivi")
IdSuivi = [str(c[1]) for c in arcpy.da.SearchCursor(TB_CATEGORIE_NAME,["nom_complet","id_cat"]) if c[0] == Suivi]
IdSuivi = IdSuivi[0]

arcpy.AddMessage("Le gros matou recupere les informations des stations et parametres du Bilan de Thio")
Params = [[p[0],p[1],p[2],p[3],p[4]] for p in arcpy.da.SearchCursor(TB_PARAMETRES,["id_parametre","nom","BilanTHIO","unite","BilanEnvSeuil","OrdreThio"],sql_clause=(None,'ORDER BY OrdreThio')) if p[0].find("/" + IdSuivi + "/") >= 0 and p[2] == "Oui"]

if len(Params) == 0:
    arcpy.AddMessage("Dommage .... Aucun parametre pour le Bilan Environnement sur ce type de suivi ou categorie")
    del Params
    arcpy.AddMessage("Fin du travail pour le gros matou!!! Oueeee")
else:
    #Attention BilanENVTypeStation correspond a Localisation dans le fichier en sortie
    SQLStation = "BilanENVRefSuiviTxt = 'RefTHIO' AND BilanENVCompartiment LIKE '%" + IdSuivi + "%' OR BilanENVRefSuiviTxt = 'SuiviTHIO' AND BilanENVCompartiment LIKE '%" + IdSuivi + "%'"
    Stations = [list(s) for s in arcpy.da.SearchCursor(FC_STATIONS,["id_geometry", "nom_simple", "BilanENVRefSuiviTxt", "BilanENVTypeStation", "BilanENVTypeBV","BilanENVZone","BilanEnvZoneRef","BilanEnvCompartiment","BilanENVOrdre","BilanENVOrdreRef"], SQLStation ,
                                                       sql_clause=(None,'ORDER BY BilanENVRefSuiviTxt, BilanENVTypeStation, BilanENVTypeBV, id_geometry'))]
    arcpy.AddMessage("Alors le gros matou constitue un Filtre SQL a constituer pour appliquer sur la table de donnees")
    PmsIds = [p[0] for p in Params]
    StatIds = [s[0] for s in Stations]
    SQL = SQLParam(PmsIds,StatIds,TB_DONNEES)
    arcpy.AddMessage("Nombre de stations: "+ str(len(Stations)))
    arcpy.AddMessage("Maintenant le gros matou recupere les donnees ordonnees par date de mesure, station, parametre")
    Datas = [list(d) for d in arcpy.da.SearchCursor(TB_DONNEES,["id_geometry",ChampDataParam,"date_","valeur","methode_analyse","valeur_texte","signe","limite_quantification"],sql_clause=(None,'ORDER BY date_, id_geometry, ' + ChampDataParam))]
    arcpy.SelectLayerByAttribute_management(TB_DONNEES, "CLEAR_SELECTION")
    arcpy.AddMessage(str(len(Datas)) + " donnees recuperees pour le Bilan Environnement sur " + Suivi + " pour le gros matou. Y'a du taff")
    arcpy.AddMessage("Faut bien stocker quelque part le resultat... le gros matou cree le fichier Excel en sortie. Vous lui avez dit qu'il se nomme " + ExcelSortie)
    XLSX = xlsxwriter.Workbook(RepSortie + "/" + ExcelSortie)
    Feuille = XLSX.add_worksheet("Donnees")
    FormTitre = FormatExcel(XLSX, "Oui", 11, "Calibri Light", "#ffffff", "#5b9ad5", "")
    FormGamme = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#c00000", "", "")
    FormMet = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#c05a11", "", "")
    FormVal = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "#0070c0", "", "")
    FormInfo = FormatExcel(XLSX, "Oui", 10, "Calibri Light", "000000", "", "")
    FormNorm = FormatExcel(XLSX, "Non", 10, "Calibri Light", "#000000", "", "")
    #FormNorm.set_num_format("0.000")
    FormDate = FormatExcel(XLSX, "Non", 10, "Calibri Light","#000000","", "dd/mm/yyyy HH:MM:SS")
    Feuille.write(0,0,"TypeData",FormTitre)
    Feuille.write(0, indiceCol["TypeStation"],"TypeStation",FormTitre)
    Feuille.write(0, indiceCol["Localisation"],"Localisation",FormTitre)
    Feuille.write(0, indiceCol["TypeBV"],"TypeBV", FormTitre)
    Feuille.write(0, indiceCol["Zone"],"Zone",FormTitre)
    Feuille.write(0, indiceCol["ZoneReference"],"ZoneReference",FormTitre)
    Feuille.write(0, indiceCol["Station"],"Station",FormTitre)
    Feuille.write(0, indiceCol["Date"],"Date_",FormTitre)
    Feuille.write(0, indiceCol["Periode"],"Periode",FormTitre)
    Line0 = ["Valeur Seuil", "", "", "", "", "","", None, ""]
    Line1 = ["TypeParametre","","","","","","",None,""]
    Line2 = ["Unite","","","","","","",None,""]
    #on ajoute en champ chaque parametre avec id parametre en nom de champ et en alias le nom du parametre
    #on met sur la premiere ligne les valeurs seuils pour chaque parametre Line0
    #on met sur la deuxieme ligne le type de parametre si cest cle ou simple parametre de Bilan ENV avec code Cle ou Env Line1
    for p in Params:
        Feuille.write(0,indiceCol["PremierParam"] + Params.index(p), p[1], FormTitre)
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
    Champs = ["TypeData","TypeStation","Localisation","TypeBV","Zone", "ZoneReference","Station","Date_","Periode"]
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
                if TypePeriode == "Saison chaude - MacroInvertebres":
                    Line = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeSsChaudeMIB(dt)]
                    LineM = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeSsChaudeMIB(dt)]
                    LineS = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeSsChaudeMIB(dt)]
                elif TypePeriode == "Saison fraiche - Poissons Crustaces":
                    Line = ["Valeur reelle", s[2], s[3], s[4], s[5],s[6], s[1], dt, PeriodeSsPoiss(dt)]
                    LineM = ["Valeur reelle", s[2], s[3], s[4], s[5],s[6], s[1], dt, PeriodeSsPoiss(dt)]
                    LineS = ["Valeur reelle", s[2], s[3], s[4], s[5],s[6], s[1], dt, PeriodeSsPoiss(dt)]
                else:
                    Line = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeAnnee(dt)]
                    LineM = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeAnnee(dt)]
                    LineS = ["Valeur reelle",s[2],s[3],s[4],s[5],s[6],s[1],dt,PeriodeAnnee(dt)]
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
                            OtherDatas.append([a,y,y+9,DatasSelDateP[a][5].replace(".",",")])
                            OtherDatasM.append([a,y,y+9,DatasSelDateP[a][4]])
                            OtherDatasS.append([a, y,y+9,DatasSelDateP[a][6]])
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
                LigneExcelValeur(FormVal, FormNorm, FormDate, Feuille, Line, indxl)
                indxl = indxl + 1
                DatasExp.append(Line)
                DatasMeth.append(LineM)
                DatasSigne.append(LineS)
                if len(OtherDatas) > 0:
                    indlines = [od[0] for od in OtherDatas]
                    indlines = list(set(indlines))
                    OtherLines = LignesSup(OtherDatas,Line,Params)
                    OtherLinesM = LignesSup(OtherDatasM,LineM,Params)
                    OtherLinesS = LignesSup(OtherDatasS,LineS,Params)
                    del OtherDatas,OtherDatasM, OtherDatasS
                    for ol in OtherLines:
                        LigneExcelValeur(FormVal, FormNorm, FormDate, Feuille, ol, indxl)
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
    if TypePeriode == "Saison chaude - MacroInvertebres":
        arcpy.AddMessage("Le gros matou vous parle: La periode est: " + TypePeriode)
        DatasMeth = [d for d in DatasMeth if d[8][0:13] == "Saison chaude"]
        DatasSigne = [d for d in DatasSigne if d[8][0:13] == "Saison chaude"]
        DatasExp = [d for d in DatasExp if d[8][0:13] == "Saison chaude"]
        #Gamme de reference avant Statistiques de Periode pour le Percentile 75
        arcpy.AddMessage("le gros matou va traiter d'abord les metriques pour les Gammes de reference :-(")
        StatsGamm = StatistiquesGammeReference(DatasMeth, DatasSigne, DatasExp, Params, IdSuivi)
        StatsGammP = [stg for stg in StatsGamm if stg[0].find("Gamme de reference - Percentile") >= 0]
        arcpy.AddMessage("le gros matou va traiter ensuite les metriques par Periode :-(")
        StatsPeriode = StatistiquesPeriode(DatasMeth,DatasSigne,DatasExp,Params, StatsGammP, IdSuivi)
        #insertion Statistiques par Periode avant les gammes de reference
        for stm in StatsPeriode:
            LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, stm, indxl)
            indxl = indxl + 1
        for stg in StatsGamm:
            LigneExcelValeur(FormGamme, FormNorm, FormDate, Feuille, stg, indxl)
            indxl = indxl + 1
        ##Mettre ici Gamme reference
    elif TypePeriode == "Saison fraiche - Poissons Crustaces":
        arcpy.AddMessage("Le gros matou vous parle: La periode est: " + TypePeriode)
        DatasMeth = [d for d in DatasMeth if d[8][0:14] == "Saison fraiche"]
        DatasSigne = [d for d in DatasSigne if d[8][0:14] == "Saison fraiche"]
        DatasExp = [d for d in DatasExp if d[8][0:14] == "Saison fraiche"]
        arcpy.AddMessage("le gros matou va traiter les metriques pour les Gammes de reference :-(")
        StatsGamm = StatistiquesGammeReference(DatasMeth, DatasSigne, DatasExp, Params, IdSuivi)
        StatsGammP = [stg for stg in StatsGamm if stg[0].find("Gamme de reference - Percentile") >= 0]
        arcpy.AddMessage("le gros matou va traiter ensuite les metriques par Periode :-(")
        StatsPeriode = StatistiquesPeriode(DatasMeth, DatasSigne, DatasExp, Params, StatsGammP, IdSuivi)
        ##Mettre ici Gamme reference
        for stm in StatsPeriode:
            LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, stm, indxl)
            indxl = indxl + 1
        for stg in StatsGamm:
            LigneExcelValeur(FormGamme, FormNorm, FormDate, Feuille, stg, indxl)
            indxl = indxl + 1
    else:
        arcpy.AddMessage("Le gros matou vous parle: La periode est: " + TypePeriode)
        ##Mettre ici Gamme reference
        arcpy.AddMessage("le gros matou va traiter d'abord les metriques pour les Gammes de reference :-(")
        StatsGamm = StatistiquesGammeReference(DatasMeth, DatasSigne, DatasExp, Params, IdSuivi)
        StatsGammP = [stg for stg in StatsGamm if stg[0].find("Gamme de reference - Percentile") >= 0]
        arcpy.AddMessage("le gros matou va traiter ensuite les metriques par Periode :-(")
        StatsPeriode = StatistiquesPeriode(DatasMeth, DatasSigne, DatasExp, Params, StatsGammP, IdSuivi)
        arcpy.AddMessage("Nombre de Statistique par annee " + str(len(StatsPeriode)))
        arcpy.AddMessage("Nombre de Statistique pour les gammes de reference " + str(len(StatsGamm)))
        for stm in StatsPeriode:
            LigneExcelValeur(FormMet, FormNorm, FormDate, Feuille, stm, indxl)
            indxl = indxl + 1
        for stg in StatsGamm:
            LigneExcelValeur(FormGamme, FormNorm, FormDate, Feuille, stg, indxl)
            indxl = indxl + 1
            
    #Creation des fichiers XML pour les metriques et gammes de reference
    XLSX.close()
    #Filtre sur l'Ordre a partir de 1 et 0 + NULL ignore
    StationsOrdre = [stat for stat in Stations if stat[8] != None and stat[8] > 0]
    ExportMetrique(StatsPeriode, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi, MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLSortie, StationsOrdre, "Suivi")
    #Export Stations de reference metrique
    StationsOrdreRef = [stat for stat in Stations if stat[9] != None and stat[9] > 0]
    ExportMetrique(StatsPeriode, StatsGamm, Params, TB_BilanENVXMLParam, TB_BilanENVXMLMetr, TB_BilanENVXMLGam, IdSuivi,
                   MetriqueXMLVoulues, GammeXMLVoulues, RepSortie + "/" + NomXMLGRSortie, StationsOrdreRef, "Reference")
    ##ExportGammeXML(StatsGamm,Params, TB_BilanENVXMLParam, TB_BilanENVXMLGam, IdSuivi, GammeXMLVoulues, RepSortie + "/" + NomXMLGRSortie)
    del DatasMeth, DatasSigne, Stations, Params, StatistiquesPeriode,DatasExp
    arcpy.AddMessage("Le gros matou est fatigue, il va dormir. A plous!!!")
