# -*- coding: utf8 -*-
import numpy, arcpy
from unidecode import unidecode

#organisation des premieres colonnes
indiceCol = {"TypeData":0, "TypeStation":1, "Localisation":2, "TypeBV":3, "Zone":4, "ZoneReference":5, "Station":6, "Date":7,"Periode":8,"PremierParam":9}
#organisation des premieres colonnes milieu marin
indiceColMar = {"TypeData":0, "TypeStation":1, "Zone":2, "TypologieStation":3, "TypeRef":4, "Station":5, "Date":6, "Periode":7,"PremierParam":8}
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
def LignesSup(AutresDonnees,PremiereLigne,ListParam,iPremParam):
    DebutLigne = PremiereLigne[0:iPremParam]
    IndexLigne = [ad[0] for ad in AutresDonnees]
    IndexLigne = list(set(IndexLigne))
    IndexLigne.sort()
    AutresLignes = list()
    for il in IndexLigne:
        DonneeFiltre = [ad for ad in AutresDonnees if ad[0] == il]
        Line = list()
        Line.extend(DebutLigne)
        indligneParam = len(ListParam) + iPremParam
        x = iPremParam
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
                    if DonneeFiltParam[b][3] is None:
                        dfp = ""
                    else:
                        dfp = DonneeFiltParam[b][3]
                    if b == 0:
                        concat = dfp
                    else:
                        concat = concat + "|" + dfp
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

#Si periode correspond a Saisons pour les Flux a particule fonction a utiliser Milieu MARIN
def PeriodeSsFLUX(DateEntree):
    #saison chaude de mai a juillet
    if DateEntree.day >= 16 and DateEntree.month == 5 or DateEntree.month > 5 and DateEntree.month < 11 or DateEntree.day <= 15 and DateEntree.month == 11:
        periode = "Saison fraiche " + DateEntree.strftime("%Y")
    else:
        periode = "Saison chaude " + DateEntree.strftime("%Y")
    return periode

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

def Annee3ans(DatesListe):
    MaxDate = max(DatesListe)
    MaxYear = MaxDate.strftime("%Y")
    MinYear = str(int(MaxYear)-2)
    return MinYear + "-" + MaxYear

#trouver les methodes Gamme de ref
def FindMethodes(DonneesM, ind):
    Meths = [d[ind] for d in DonneesM if d[ind] is not None and d[ind] != ""]
    Meths = list(set(Meths))
    MethsTxt = ""
    if len(Meths) > 0:
        MethsTxt = ConcatenationInformation(Meths)
    return MethsTxt

def FindLQ(DonneesS,ind,DonneesEx, idtifSuivi):
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
        #pour les cas SURF on reaffiche la LQ d'origine qui a ete divisee par deux
        if idtifSuivi == "SURF":
            LQs = [str(float(lq.replace(",",".")) * 2) for lq in LQs]
            RechLQ = ConcatenationInformation(LQs)
        else:
            RechLQ= ConcatenationInformation(LQs)
    del indexLQ
    return RechLQ

def FindSigne(DonneesS,ind):
    Signs = [str(d[ind]) for d in DonneesS if d[ind] is not None]
    nblq = Signs.count("<")
    return str(nblq)

def FindValsNonVides(DonneesEx, ind):
    DonneesExNnVide = [d for d in DonneesEx if d[ind] is not None and d[ind] != ""]
    valN = len(DonneesExNnVide)
    ValsEx = [float(d[ind].replace(",", ".")) for d in DonneesExNnVide if str(type(d[ind])) == "<type 'str'>" or str(type(d[ind])) == "<type 'unicode'>"]
    ValsEx.extend([d[ind] for d in DonneesExNnVide if str(type(d[ind])) == "<type 'float'>" or str(type(d[ind])) == "<type 'int'>"])
    return [valN,ValsEx]

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
    return {"min" : Min, "max" : Max, "moy": moy, "ect" : ect, "med": med, "perc10" : perc10, "perc25" : perc25, "perc75" : perc75, "perc80" : perc80, "perc85" : perc85, "perc90" : perc90, "perc95" : perc95}


def NbSupPerc(StGP, indcol, valcol, percRef, ind, ValsEx, ValN):
    # cas particulier pour nappe_base_littorale des eaux souterraines a comparer avec aquitard_lateritique_aquifere_profond
    # indcol et valcol sont des listes
    if "nappe_base_littorale" in valcol:
        iv = valcol.index("nappe_base_littorale")
        valcol[iv] = "aquitard_lateritique_aquifere_principal"
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
        P = str(P)
        P = float(P.replace(",", "."))
        NbSupP = len([val for val in ValsEx if val > P])
        PctSupP = NbSupP * 100 / ValN
    return {"NbS": str(NbSupP), "PctS": str(PctSupP)}

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
    FormF.set_num_format(formatcell)
    return FormF
#fonction pour savoir si le texte peut etre un numerique entier 0 = Oui 1=Non
def NumEnt(txt):
    val = 0
    for t in txt:
        if not t.isdigit():
            val = 1
    return val

#insertion de ligne de donnees sur Excel
def LigneExcelValeur(formDt, formN, formDate, feuilleX, ligne, Ind, mil):
    colX = 0
    if mil == "Eaux douces":
        icol = indiceCol
    else:
        icol = indiceColMar
    for L in ligne:
        if colX == 0:
            feuilleX.write(Ind, colX, L, formDt)
        #cellule date
        elif colX == icol["Date"]:
            feuilleX.write(Ind, colX, L, formDate)
        else:
            #valeur vide
            if str(type(L)) == "<type 'NoneType'>":
                feuilleX.write(Ind, colX,"",formN)
            #valeur numerique
            elif str(type(L)) == "<type 'int'>" or str(type(L)) == "<type 'float'>":
                feuilleX.write_number(Ind, colX, L, formN)
            #si valeur texte peut être ecrite en numerique decimale
            elif L != "" and L.find("|") == -1 and L.find("<") == -1 and L[0].isdigit() and L.find(" ") == -1:
                L = L.replace(",", ".")
                if L.find(".") > 0:
                    L = float(L)
                    feuilleX.write_number(Ind, colX, L, formN)
                #texte en valeur numerique entier
                elif str(type(L)) == "<type 'str'>" and NumEnt(L) == 0 or str(type(L)) == "<type 'unicode'>" and NumEnt(L) == 0 :
                    L = float(L)
                    feuilleX.write_number(Ind,colX, L, formN)
                else:
                    feuilleX.write(Ind, colX, L, formN)
            else:
                feuilleX.write(Ind, colX, L, formN)
        colX = colX + 1

