# -*- coding: utf8 -*-
import arcpy, xlsxwriter, xlrd, datetime
from unidecode import unidecode
from ModuleBilanEnv import FormatExcel

#organisation des premieres colonnes
indiceCol = {"TypeData":0, "TypeStation":1, "Localisation":2, "TypeBV":3, "Zone":4, "ZoneReference":5, "Station":6, "Date":7,"Periode":8,"PremierParam":9}
#organisation des premieres colonnes Milieu Marin
indiceColMar = {"TypeData":0, "TypeStation":1, "Zone":2, "TypologieStation":3, "TypeRef":4, "Station":5, "Date":6, "Periode":7,"PremierParam":8}

#fonction pour definir les indices de colonnes pour l'emplacement des annees sur la feuille Excel
def AnIndiceExcel(ListeAnnee, NbMetrique, Suivi="", Periode3ans="",NbMetrique3a=0):
    AnneeInd = dict()
    xa = 0
    if Suivi == "EAU":
        for an in ListeAnnee:
            if an == Periode3ans:
                AnneeInd[str(an)] = [xa, xa + NbMetrique3a - 1]
                xa = xa + NbMetrique
            else:
                AnneeInd[str(an)] = [xa, xa + NbMetrique - 1]
                xa = xa + NbMetrique
    elif Suivi == "FLUX":
        for an in ListeAnnee:
            AnneeInd["sf_" + str(an)] = [xa, xa + NbMetrique - 1]
            xa = xa + NbMetrique
            AnneeInd["sc_" + str(an)] = [xa, xa + NbMetrique - 1]
            xa = xa + NbMetrique
    else:
        for an in ListeAnnee:
            AnneeInd[str(an)] = [xa, xa + NbMetrique - 1]
            xa = xa + NbMetrique
    return AnneeInd

#Fonction pour definir sous les annees les differentes metriques en sous titre des colonnes de donnees pour Excel
def AnMetriqueExcel(ListeAnnee, EquivalMetrique, AnIndiceDict, FormTitreExcel, Periode3ans="", IndicePremColMet=3, Suivi="", EquivalMetrique3ans=list(),FormTitreExcelRS=""):
    AnMetr = dict()
    if Suivi == "EAU":
        AnMetr3 = dict()
        for an in ListeAnnee:
            if an == Periode3ans:
                for met3 in EquivalMetrique3ans:
                    im = EquivalMetrique3ans.index(met3)
                    if FormTitreExcelRS != "":
                        AnMetr3[an, met3[0]] = [2, IndicePremColMet + im + AnIndiceDict[str(an)][0], met3[0], FormTitreExcelRS]
                    else:
                        AnMetr3[an, met3[0]] = [2, IndicePremColMet + im + AnIndiceDict[str(an)][0], met3[0], FormTitreExcel]
            else:
                for met2 in EquivalMetrique:
                    im = EquivalMetrique.index(met2)
                    AnMetr[an, met2[0]] = [2, IndicePremColMet + im + AnIndiceDict[str(an)][0], met2[0], FormTitreExcel]
        SortieAnMetr = [AnMetr, AnMetr3]
    elif Suivi == "FLUX":
        for an in ListeAnnee:
            for met2 in EquivalMetrique:
                im = EquivalMetrique.index(met2)
                AnMetr["sc_" + an, met2[0]] = [2, IndicePremColMet + im + AnIndiceDict["sc_" + str(an)][0], met2[0], FormTitreExcel]
                AnMetr["sf_" + an, met2[0]] = [2, IndicePremColMet + im + AnIndiceDict["sf_" + str(an)][0], met2[0], FormTitreExcel]
            SortieAnMetr = AnMetr
    else:
        AnMetr = dict()
        for an in ListeAnnee:
            for met2 in EquivalMetrique:
                im = EquivalMetrique.index(met2)
                AnMetr[an, met2[0]] = [2, IndicePremColMet + im + AnIndiceDict[str(an)][0], met2[0], FormTitreExcel]
        SortieAnMetr = AnMetr
    return SortieAnMetr

#Fonction pour un dictionnaire avec l'equivalence de metrique pour XML et pour Excel
def EquivalenceMetriqueXMLExcel(MetriquesAttendues, EquivalenceMetrique):
    Met = dict()
    for ma in MetriquesAttendues:
        metr = [met[1] for met in EquivalenceMetrique if met[0] == ma]
        metrXL = [met[2] for met in EquivalenceMetrique if met[0] == ma]
        if len(metr) > 0:
            Met[ma] = [metr[0], metrXL[0]]
        else:
            arcpy.AddMessage("Correspondance non trouvee pour la metrique attendue: " + ma)
            Met[ma] = [ma, ma]
    return Met

#Fonction pour retourner un dictionnaire sur les equivalences de parametres nom et unite
def ParametreDico(Parametres, EquivalenceParametres,identSuivi=""):
    Prmdict = dict()
    if identSuivi == "EAU":
        for p in Parametres:
            nameparam = [np[1] for np in EquivalenceParametres if np[0] == p[1]]
            if len(nameparam) > 0:
                nameparam = nameparam[0]
                unitparam = [np[2] for np in EquivalenceParametres if np[0] == p[1]]
                seuillittoral = [np[3] for np in EquivalenceParametres if np[0] == p[1]]
                seuilcotier = [np[4] for np in EquivalenceParametres if np[0] == p[1]]
                seuilocean = [np[5] for np in EquivalenceParametres if np[0] == p[1]]
                metrique = [np[6] for np in EquivalenceParametres if np[0] == p[1]]
                if len(unitparam) > 0:
                    unitparam = unitparam[0]
                    seuillittoral = seuillittoral[0]
                    seuilcotier = seuilcotier[0]
                    seuilocean = seuilocean[0]
                    metrique = metrique[0]
                else:
                    arcpy.AddMessage("Correspondance non trouvee pour l'unite de parametre: " + p[1])
                    unitparam = ""
                    seuillittoral = ""
                    seuilcotier = ""
                    seuilocean = ""
                    metrique = ""
            else:
                arcpy.AddMessage("Correspondance non trouvee pour le parametre: " + p[1])
                nameparam = p[1]
                unitparam = ""
                seuillittoral = ""
                seuilcotier = ""
                seuilocean = ""
                metrique = ""
            Prmdict[p[1]] = dict()
            Prmdict[p[1]]["name"] = nameparam
            Prmdict[p[1]]["unite"] = unitparam
            Prmdict[p[1]]["littoral"] = seuillittoral
            Prmdict[p[1]]["cotier"] = seuilcotier
            Prmdict[p[1]]["oceanique"] = seuilocean
            Prmdict[p[1]]["metrique"] = metrique
    else:
        for p in Parametres:
            nameparam = [np[1] for np in EquivalenceParametres if np[0] == p[1]]
            if len(nameparam) > 0:
                nameparam = nameparam[0]
                unitparam = [np[2] for np in EquivalenceParametres if np[0] == p[1]]
                seuilparam = [np[3] for np in EquivalenceParametres if np[0] == p[1]]
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
            Prmdict[p[1]] = [nameparam, unitparam,seuilparam]
    return Prmdict

#Fonction pour retourner un dictionnaire avec comme cle les stations comme valeurs les periodes correspondantes
def PeriodeStat(LStations, LStationPeriode):
    PeriodeStdict = dict()
    for s in LStations:
        PeriodeStdict[s[indiceColMar["Station"] - 1]] = sorted(list(set([ast[indiceColMar["Periode"] - 2] for ast in LStationPeriode if
                                                                      ast[indiceColMar["Station"] - 1] == s[
                                                                          indiceColMar["Station"] - 1]])))
    return PeriodeStdict

#Fonction pour ecrire la valeur dans le tableau et voir si note comparaison
#arguments = valeur, Feuille excel, Types (texte ou numeriques), Types numériques, Types Texte, numero ligne, numero colonne, Format, Format decimal, Format entier
def TraitementValeur(val, Feuil, typesTT, typesNUM, typesSTR, irow, icol, Format, FormatDec, FormatEnt, MetriqueRef):
    #si au moins une Valeur a ete recuperee
    if len(val) > 0:
        Valeur = val[0]
        typeval = str(type(Valeur))
        typeval = typeval[typeval.find("'")+1:-2]
        #si la Valeur est vide ou nulle
        if Valeur == "" or Valeur is None:
            Valeur = ""
            Feuil.write(irow,icol,Valeur,Format)
            if MetriqueRef == "Metrique":
                NoteCompar = "Non"
        #si le type de Valeur est au moins du numerique ou textuel
        elif typeval in typesTT:
            if typeval in typesSTR:
                if Valeur != "" and Valeur.find("|") == -1:
                    Valeur = Valeur.replace(",", ".")
                    Valeur = float(Valeur)
                    if Valeur.is_integer():
                        Feuil.write_number(irow,icol,Valeur,FormatEnt)
                    else:
                        Feuil.write_number(irow, icol, Valeur, FormatDec)

            elif typeval == "float":
                Feuil.write_number(irow,icol,Valeur,FormatDec)
            elif typeval == "int":
                Feuil.write_number(irow,icol,Valeur,FormatEnt)
            if MetriqueRef == "Metrique":
                NoteCompar = "Oui"
        ####COMPARTIMENT NOTE DE COMPARAISON############
        else:
            Feuil.write(irow,icol,Valeur,Format)
            NoteCompar = "Non"
    else:
        Valeur = ""
        Feuil.write(irow,icol,Valeur,Format)
        if MetriqueRef == "Metrique":
            NoteCompar = "Non"
    if MetriqueRef == "Metrique":
        return NoteCompar, Valeur
    else:
        return Valeur

#Definir les notes par rapports aux seuils Zoneco
#arguments: valeur, seuils (min,max), feuille Excel, numero ligne, numero colonne, format non perturbe, format moderement perturbe, format fortement perturbe
def SeuilsZoneco(Valeur, Seuils, Feuille, irow, icol, FormNonPert, FormModerPert, FormFortPert):
    if Valeur < Seuils[0]:
        Feuille.write(irow, icol, "Non perturbe", FormNonPert)
    elif Valeur >= Seuils[0] and Valeur < Seuils[1]:
        Feuille.write(irow, icol, "Moderement perturbe", FormModerPert)
    else:
        Feuille.write(irow, icol, "Fortement perturbe", FormFortPert)


##FONCTION  pour la sortie XLSX Metriques
def ExportMetrique(Milieu, DonneesPeriodes, DonneesGamme,Parametres,XMLParamEquiv, XMLMetriqueEquiv, XMLGammeEquiv, idsuivi, metrAttendu, GRAttendu, FichierXML, OrdreStation, TypeSuivi, AnEtude, ExcelPrec=0):
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
    #Titre Fond Bleu
    FormTitleNote = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#bdd7ee", "")
    FormTitleNote.set_border(5)
    FormTitleNote.set_align("center")
    FormTitleNote.set_align("vcenter")
    FormTitleNote.set_text_wrap()
    #Titre Fond Rose saumon ZONECO
    FormTitleNoteRS = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#FFCCCC", "")
    FormTitleNoteRS.set_border(5)
    FormTitleNoteRS.set_align("center")
    FormTitleNoteRS.set_align("vcenter")
    FormTitleNoteRS.set_text_wrap()
    #Titre Fond Lavande
    FormTitleNoteLV = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#FFCCFF", "")
    FormTitleNoteLV.set_border(5)
    FormTitleNoteLV.set_align("center")
    FormTitleNoteLV.set_align("vcenter")
    FormTitleNoteLV.set_text_wrap()
    #Titre Fond Vert Clair PARAMETRE
    FormTitleNoteVC = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#CCFFCC", "")
    FormTitleNoteVC.set_border(5)
    FormTitleNoteVC.set_align("center")
    FormTitleNoteVC.set_align("vcenter")
    FormTitleNoteVC.set_text_wrap()
    #Titre Fond Ancien Score
    FormTitleScoreOld = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#DAD5FD", "")
    FormTitleScoreOld.set_border(5)
    FormTitleScoreOld.set_align("center")
    FormTitleScoreOld.set_align("vcenter")
    FormTitleScoreOld.set_text_wrap()
    #Titre Fond Score
    FormTitleScore = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#ABA0FA", "")
    FormTitleScore.set_border(5)
    FormTitleScore.set_align("center")
    FormTitleScore.set_align("vcenter")
    FormTitleScore.set_text_wrap()
    #Titre Evolution temporelle
    FormTitleEvoTemp = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light","#000000",  "#FFFFCC", "")
    FormTitleEvoTemp.set_border(5)
    FormTitleEvoTemp.set_align("center")
    FormTitleEvoTemp.set_align("vcenter")
    FormTitleEvoTemp.set_text_wrap()
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
    #cellule numerique decimale
    FormNum = FormatExcel(MetXLSX, "Non", 9, "Calibri Light", "#000000", "", "0.000")
    FormNum.set_align("center")
    FormNum.set_align("vcenter")
    #FormNum.set_num_format("0.000")
    FormNum.set_border(1)
    #cellule numerique decimal
    FormNumEnt = FormatExcel(MetXLSX, "Non", 9, "Calibri Light", "#000000", "", "#")
    FormNumEnt.set_align("center")
    FormNumEnt.set_align("vcenter")
    FormNumEnt.set_border(1)
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
    # cellule Inconnu
    FormInco = FormatExcel(MetXLSX, "Oui", 9, "Calibri Light", "#000000", "#C0C0C0", "")
    FormInco.set_align("center")
    FormInco.set_align("vcenter")
    #An = int(datetime.datetime.now().strftime("%Y"))
    An = AnEtude
    Annees = [An - 4, An - 3, An - 2, An - 1, An]
    arcpy.AddMessage("Metriques attendues:")
    arcpy.AddMessage(metrAttendu)
    arcpy.AddMessage("Periodes attendues:")
    arcpy.AddMessage(Annees)
    arcpy.AddMessage("Gammes attendues:")
    arcpy.AddMessage(GRAttendu)
    if Milieu == "Eaux douces":
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
        if idsuivi == "SURF" or idsuivi == "SEDI":
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
            EquivParam = [list(ep) for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite","SeuilBilanENV", "SeuilBilanENV_Etude", "PeriodeBilanENV_Etude", "LQ_BilanENV_Etude", "N_LQ_BilanENV_Etude", "N_BilanENV_Etude","Med_BilanENV_Etude", "Max_BilanENV_Etude"],"id_cat2 = '" + idsuivi + "'")]
            #Equivalence Metrique
            EquivMetr = [list(em) for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml","metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
            # Equivalence Gamme de reference
            EquivGamme = [list(eg) for eg in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
            nbmet = len(metrAttendu)
            EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu]
            #dictionnaire Annee indice colonne fusion pour la ligne Annees sous Excel par feuille Excel de parametre
            AnInd = dict()
            xa = 0
            for an in Annees:
                AnInd[str(an)] = [xa, xa + nbmet - 1]
                xa = xa + nbmet
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
            #derniere colonne de la derniere annee après la derniere metrique demandee
            LastColAnnee = AnMet[An, Metr[metrAttendu[len(metrAttendu)-1]]][1]
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
            TypeRef = list(set([dtgr[indiceCol["Localisation"]] for dtgr in DonneesGammeVoulues]))
            arcpy.AddMessage(TypeRef)
            #dictionnaire de periode par station avec comme valeurs periodes et stations en cle
            PerStdict = dict()
            for s in Station:
                PerStdict[s[indiceCol["Station"]-2]] = sorted(list(set([ast[indiceCol["Periode"]-3] for ast in StationPer if ast[indiceCol["Station"]-2] == s[indiceCol["Station"]-2]])))
            #dictionnaire pour les scores de l'annee precedente avec en cle les noms des feuilles (parametre nom reduit) et en valeurs dico clé station = liste station et clé scores liste des scores
            Scoredict = dict()
            if ExcelPrec != 0:
                FExcP = xlrd.open_workbook(ExcelPrec)
                NamesFeuilles = FExcP.sheet_names()
                for nf in NamesFeuilles:
                    if nf != "grilles":
                        ValsScore = list()
                        StatsPrec = list()
                        FeuilleP = FExcP.sheet_by_name(nf)
                        NamesCol = [c.value for c in FeuilleP.row(1)]
                        ColScore = "Score " + str(An - 1)
                        if ColScore in NamesCol:
                            iScore = NamesCol.index(ColScore)
                            ValsScore.extend([c.value for c in FeuilleP.col(iScore)][3:])
                            NamesCol = [c.value for c in FeuilleP.row(2)]
                            if "Nom" in NamesCol:
                                iStat = NamesCol.index("Nom")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                            elif "Station" in NamesCol:
                                iStat = NamesCol.index("Station")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                        Scoredict[nf] = {"Stations": StatsPrec, "Scores":ValsScore}
                del FExcP
            FeuillesNA = list()
            for p in Parametres:
                #seuil du parametre pour note reference
                seuilP = Paramdict[p[1]][2]
                #seuil#
                ip = Parametres.index(p)
                arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ipdt = ip+ indiceCol["PremierParam"]
                Fparam = MetXLSX.add_worksheet(Paramdict[p[1]][0])
                Fparam.merge_range(0,0,0,2, Paramdict[p[1]][0] + " (" + Paramdict[p[1]][1] + ")", FormTitle)
                if TypeSuivi == "Suivi":
                    Fparam.merge_range(1,0,1,2, "Station de suivi", FormSt)
                else:
                    Fparam.merge_range(1, 0, 1, 2, "Station de reference", FormStR)
                Fparam.write(2,0,"Type", FormTitle)
                Fparam.write(2,1,"Zone", FormTitle)
                Fparam.write(2,2,"Station", FormTitle)
                Fparam.freeze_panes(3,3)
                # Longueur des colonnes
                Fparam.set_column(0, 2, 20)
                Fparam.set_column(3, LastColAnnee, 10)
                Fparam.set_column(LastColAnnee + 1, LastColAnnee + 5, 18)
                Fparam.set_column(LastColAnnee + 7, LastColAnnee + 17, 15)

                #reference a mettre en forme
                # On met en forme les cellules de valeur avant d'integrer
                LigneFin = len(Station) + 2
                ColonneFin = AnInd[str(An)][1] + 3
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
                    Fparam.write(sx,0, s[indiceCol["Localisation"]-1],FormN)
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
                                if ma[0] == "% > Perc 75" and int(perst) == An:
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                                elif ma[0] == "% > Perc 75 - OEIL" and int(perst) == An and s[indiceCol["Localisation"]-1] == "Riviere":
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 2, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 2, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 2, "Pas de seuil", FormNonSeuil)
                            else:
                                Fparam.write(sx,metx,valeur,FormNum)
                    ###RECUPERATION DES ANCIENS SCORES
                    if ExcelPrec != 0:
                        if Paramdict[p[1]][0] in Scoredict.keys():
                            #si le nom de la station est dans les resultats precedents
                            if s[indiceCol["Station"]-2] in Scoredict[Paramdict[p[1]][0]]["Stations"]:
                                istp = Scoredict[Paramdict[p[1]][0]]["Stations"].index(s[indiceCol["Station"]-2])
                                if Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Bon" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Non perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Non perturbe", FormNonPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Mauvais" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Fortement perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Fortement perturbe", FormFortPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Inconnu":
                                    Fparam.write(sx, LastColAnnee + 4, "Inconnu", FormInco)
                            else:
                                arcpy.AddMessage("Pas de correspndance avec " + s[indiceCol["Station"]-2] + " pour " + Paramdict[p[1]][0])
                        else:
                            if Paramdict[p[1]][0] not in FeuillesNA:
                                arcpy.AddMessage("La Feuille " + Paramdict[p[1]][0] + " n'existe pas dans " + ExcelPrec)
                                FeuillesNA.append(Paramdict[p[1]][0])
                    sx += 1
                    #arcpy.AddMessage("Station traitee " + s[indiceCol["Station"]-2]  + " pour le fichier XML")
                arcpy.AddMessage("Parametre " + p[1] + " traite pour les metriques: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                # colonne Note comparaison initialise a 0
                notex = LastColAnnee + 1
                if notex > 0:
                    Fparam.merge_range(1, notex, 2, notex, "Note comparaison reference", FormTitleNote)
                else:
                    Fparam.merge_range(1,notex+1, 2, Fparam.dim_colmax+1, "Note comparaison reference", FormTitleNote)
                Fparam.merge_range(1, notex + 1, 2, notex + 1, "Note etude seuil", FormTitleNote)
                Fparam.merge_range(1,notex + 2, 2, notex + 2, "Evolution temporelle", FormTitle)
                Fparam.merge_range(1,notex + 3, 2, notex + 3, "Score " + str(An - 1), FormTitle)
                Fparam.merge_range(1,notex + 4, 2, notex + 4, "Score " + str(An), FormTitle)
                Fparam.set_column(0,notex,18)
                ##############################GAMME DE REFERENCE#################################################
                debColGamme = Fparam.dim_colmax + 2
                Fparam.merge_range(0, debColGamme, 0, debColGamme + 4, "Gamme de reference", FormTitleNote)
                Fparam.write(1, debColGamme, "Type", FormTitle)
                Fparam.write(1, debColGamme + 1, "Stations", FormTitle)
                Fparam.write(1, debColGamme + 2, "Periode", FormTitle)
                refx = 2
                for ref in TypeRef:
                    DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceCol["Localisation"]] == ref]
                    #recherche de la periode particuliere si Seuil Riviere OEIL
                    if ref == "Seuil Riviere OEIL":
                        EquivParamSel = [ep for ep in EquivParam if ep[0] == p[1]]
                        if len(EquivParamSel) > 0:
                            EquivParamSel = EquivParamSel[0]
                            PerG = EquivParamSel[5]
                            Fparam.write(refx, debColGamme + 2, PerG, FormN)
                        else:
                            Fparam.write(refx, debColGamme + 2, "", FormN)
                    else:
                        Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceCol["Periode"]], FormN)

                    Fparam.write(refx, debColGamme, ref, FormN)
                    Fparam.write(refx, debColGamme + 1, DonneesGVSelect[0][indiceCol["Station"]], FormN)

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
                        if str(type(valeur)) == "<type 'int'>" or str(type(valeur)) == "<type 'float'>":
                            Fparam.write(refx, grx, valeur, FormNum)
                            Fparam.write(1, grx, gamXL, FormTitle)
                        else:
                            valeur = unidecode(valeur).replace(".", ",")
                            Fparam.write(1, grx, gamXL, FormTitle)
                            if valeur != "" and valeur.find("|") == -1 and valeur.find(" ") == -1:
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
        elif idsuivi == "SOUT":
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
            arcpy.AddMessage("id_cat2 = '"+ idsuivi + "'")
            EquivParam = [list(ep) for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite","SeuilBilanENV", "SeuilBilanENV_Etude", "PeriodeBilanENV_Etude", "LQ_BilanENV_Etude", "N_LQ_BilanENV_Etude", "N_BilanENV_Etude","Med_BilanENV_Etude", "Max_BilanENV_Etude"],"id_cat2 = '" + idsuivi + "'")]
            #Equivalence Metrique
            EquivMetr = [list(em) for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml","metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
            # Equivalence Gamme de reference
            EquivGamme = [list(eg) for eg in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
            nbmet = len(metrAttendu)
            EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu]
            #dictionnaire Annee indice colonne fusion pour la ligne Annees sous Excel par feuille Excel de parametre
            AnInd = dict()
            xa = 0
            for an in Annees:
                AnInd[str(an)] = [xa, xa + nbmet - 1]
                xa = xa + nbmet
            AnMet = dict()
            #dictionnaire AnMet pour placer le titre des metrique par annee pour les annees normales et AnMet3 pour les metriques sur les 3 dernieres annees avec mise en forme
            met_init = 4
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
            #derniere colonne de la derniere annee après la derniere metrique demandee
            LastColAnnee = AnMet[AnEtude, Metr[metrAttendu[len(metrAttendu)-1]]][1]
            #dictionnaire avec parametre en cle nom parametre et valeurs (nom et unite)
            Paramdict = dict()
            for p in Parametres:
                nameparam = [np[1] for np in EquivParam if np[0] == p[1]]
                if len(nameparam) > 0:
                    nameparam = nameparam[0]
                    unitparam = [np[2] for np in EquivParam if np[0] == p[1]]
                    seuilparam = [np[3] for np in EquivParam if np[0] == p[1]]
                    seuilparam2 = [np[4] for np in EquivParam if np[0] == p[1]]
                    if len(unitparam) > 0:
                        unitparam = unitparam[0]
                        seuilparam = seuilparam[0]
                        seuilparam2 = seuilparam2[0]
                    else:
                        arcpy.AddMessage("Correspondance non trouvee pour l'unite de parametre: " + p[1])
                        unitparam = ""
                else:
                    arcpy.AddMessage("Correspondance non trouvee pour le parametre: " + p[1])
                    nameparam = p[1]
                    unitparam = ""
                    seuilparam = ""
                    seuilparam2 = ""
                Paramdict[p[1]] = [nameparam,unitparam,seuilparam,seuilparam2]
            #Donnees Gamme de reference#
            DonneesGammeVoulues = [d for d in DonneesGamme if d[0] in GRAttendu]
            TypeRef = list(set([dtgr[indiceCol["Localisation"]] for dtgr in DonneesGammeVoulues]))
            #demande d'organiser dans l'ordre lateritique, principal, mixte
            TypeRef.sort()
            TypeRef2 = list(set([dtgr[indiceCol["TypeStation"]] for dtgr in DonneesGammeVoulues if dtgr[indiceCol["TypeStation"]] != "Suivi"]))
            arcpy.AddMessage(TypeRef)
            #dictionnaire de periode par station avec comme valeurs periodes et stations en cle
            PerStdict = dict()
            for s in Station:
                PerStdict[s[indiceCol["Station"]-2]] = sorted(list(set([ast[indiceCol["Periode"]-3] for ast in StationPer if ast[indiceCol["Station"]-2] == s[indiceCol["Station"]-2]])))
            #dictionnaire pour les scores de l'annee precedente avec en cle les noms des feuilles (parametre nom reduit) et en valeurs dico clé station = liste station et clé scores liste des scores
            Scoredict = dict()
            if ExcelPrec != 0:
                FExcP = xlrd.open_workbook(ExcelPrec)
                NamesFeuilles = FExcP.sheet_names()
                for nf in NamesFeuilles:
                    if nf != "grilles":
                        ValsScore = list()
                        StatsPrec = list()
                        FeuilleP = FExcP.sheet_by_name(nf)
                        NamesCol = [c.value for c in FeuilleP.row(1)]
                        ColScore = "Score " + str(AnEtude - 1)
                        if ColScore in NamesCol:
                            iScore = NamesCol.index(ColScore)
                            ValsScore.extend([c.value for c in FeuilleP.col(iScore)][3:])
                            NamesCol = [c.value for c in FeuilleP.row(2)]
                            if "Nom" in NamesCol:
                                iStat = NamesCol.index("Nom")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                            elif "Station" in NamesCol:
                                iStat = NamesCol.index("Station")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                        Scoredict[nf] = {"Stations": StatsPrec, "Scores":ValsScore}
                del FExcP
            FeuillesNA = list()
            for p in Parametres:
                #seuil du parametre pour note reference
                seuilP = Paramdict[p[1]][2]
                seuilP2 = Paramdict[p[1]][3]
                #seuil#
                ip = Parametres.index(p)
                arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ipdt = ip+ indiceCol["PremierParam"]
                Fparam = MetXLSX.add_worksheet(Paramdict[p[1]][0])
                Fparam.merge_range(0,0,0,3, Paramdict[p[1]][0] + " (" + Paramdict[p[1]][1] + ")", FormTitle)
                if TypeSuivi == "Suivi":
                    Fparam.merge_range(1,0,1,3, "Piezometre de suivi", FormSt)
                else:
                    Fparam.merge_range(1, 0, 1, 3, "Piezometre de reference", FormStR)
                Fparam.write(2, 0, "Source d'influence", FormTitle)
                Fparam.write(2, 1, "Aquifere", FormTitle)
                Fparam.write(2, 2, "Zone", FormTitle)
                Fparam.write(2, 3, "Piezometre", FormTitle)
                Fparam.freeze_panes(3,4)
                # Longueur des colonnes
                Fparam.set_column(0, 2, 20)
                Fparam.set_column(3, LastColAnnee, 10)
                Fparam.set_column(LastColAnnee + 1, LastColAnnee + 5, 18)
                Fparam.set_column(LastColAnnee + 7, LastColAnnee + 17, 15)

                #reference a mettre en forme
                # On met en forme les cellules de valeur avant d'integrer
                LigneFin = len(Station) + 2
                ColonneFin = AnInd[str(AnEtude)][1] + 3
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
                    Fparam.write(sx,0, s[indiceCol["Zone"]-1],FormN)
                    Fparam.write(sx,1, s[indiceCol["Localisation"]-1],FormN)
                    Fparam.write(sx,2, s[indiceCol["TypeBV"]-1],FormN)
                    Fparam.write(sx, 3, s[indiceCol["Station"] - 2], FormN)
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
                                if ma[0] == "% > Perc 75" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "principal" or ma[0] == "% > Perc 75" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "principal lateritique":
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                                elif ma[0] == "% > Perc 75" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "lateritique":
                                    typeseuilP2 = str(type(seuilP2))
                                    if typeseuilP2 == "<type 'int'>" or typeseuilP2 == "<type 'float'>":
                                        if valeur >= seuilP2:
                                            Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                                ##NOTE COMPARAISON CONTROLE
                                if ma[0] == "% > Perc 75 - Controle" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "principal" or ma[0] == "% > Perc 75 - Controle" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "principal lateritique":
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 2, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 2, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 2, "Pas de seuil", FormNonSeuil)
                                elif ma[0] == "% > Perc 75 - Controle" and int(perst) == AnEtude and s[indiceCol["Localisation"]-1] == "lateritique":
                                    typeseuilP2 = str(type(seuilP2))
                                    if typeseuilP2 == "<type 'int'>" or typeseuilP2 == "<type 'float'>":
                                        if valeur >= seuilP2:
                                            Fparam.write(sx, LastColAnnee + 2, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 2, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 2, "Pas de seuil", FormNonSeuil)
                            else:
                                Fparam.write(sx,metx,valeur,FormNum)
                    ###RECUPERATION DES ANCIENS SCORES
                    if ExcelPrec != 0:
                        if Paramdict[p[1]][0] in Scoredict.keys():
                            #si le nom de la station est dans les resultats precedents
                            if s[indiceCol["Station"]-2] in Scoredict[Paramdict[p[1]][0]]["Stations"]:
                                istp = Scoredict[Paramdict[p[1]][0]]["Stations"].index(s[indiceCol["Station"]-2])
                                if Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Bon" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Non perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Non perturbe", FormNonPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Mauvais" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Fortement perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Fortement perturbe", FormFortPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Inconnu":
                                    Fparam.write(sx, LastColAnnee + 4, "Inconnu", FormInco)
                            else:
                                arcpy.AddMessage("Pas de correspndance avec " + s[indiceCol["Station"]-2] + " pour " + Paramdict[p[1]][0])
                        else:
                            if Paramdict[p[1]][0] not in FeuillesNA:
                                arcpy.AddMessage("La Feuille " + Paramdict[p[1]][0] + " n'existe pas dans " + ExcelPrec)
                                FeuillesNA.append(Paramdict[p[1]][0])
                    sx += 1
                    #arcpy.AddMessage("Station traitee " + s[indiceCol["Station"]-2]  + " pour le fichier XML")
                arcpy.AddMessage("Parametre " + p[1] + " traite pour les metriques: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                # colonne Note comparaison initialise a 0
                notex = LastColAnnee + 1
                if notex > 0:
                    Fparam.merge_range(1, notex, 2, notex, "Note comparaison reference - volontaire", FormTitleNote)
                else:
                    Fparam.merge_range(1,notex+1, 2, Fparam.dim_colmax+1, "Note comparaison reference - volontaire", FormTitleNote)
                Fparam.merge_range(1, notex + 1, 2, notex + 1, "Note comparaison reference - controle", FormTitleNote)
                Fparam.merge_range(1,notex + 2, 2, notex + 2, "Evolution temporelle", FormTitle)
                Fparam.merge_range(1,notex + 3, 2, notex + 3, "Score " + str(AnEtude - 1), FormTitle)
                Fparam.merge_range(1,notex + 4, 2, notex + 4, "Score " + str(AnEtude), FormTitle)
                Fparam.set_column(0,notex,18)
                ##############################GAMME DE REFERENCE#################################################
                debColGamme = Fparam.dim_colmax + 2
                Fparam.merge_range(0, debColGamme, 0, debColGamme + 4, "Gamme de reference", FormTitleNote)
                Fparam.write(1, debColGamme, "Type", FormTitle)
                Fparam.write(1, debColGamme + 1, "Aquifere", FormTitle)
                Fparam.write(1, debColGamme + 2, "Piezometre", FormTitle)
                Fparam.write(1, debColGamme + 3, "Periode", FormTitle)
                refx = 2
                for ref2 in TypeRef2:
                    for ref in TypeRef:
                        DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceCol["Localisation"]] == ref and dgv[indiceCol["TypeStation"]] == ref2]
                        Fparam.write(refx, debColGamme + 3, DonneesGVSelect[0][indiceCol["Periode"]], FormN)
                        if ref2 == "Reference":
                            Fparam.write(refx, debColGamme, "Volontaire", FormN)
                        else:
                            Fparam.write(refx, debColGamme, ref2, FormN)
                        Fparam.write(refx, debColGamme + 1, ref, FormN)
                        Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceCol["Station"]], FormN)

                        grx = debColGamme + 4
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
                            if str(type(valeur)) == "<type 'int'>" or str(type(valeur)) == "<type 'float'>":
                                Fparam.write(refx, grx, valeur, FormNum)
                                Fparam.write(1, grx, gamXL, FormTitle)
                            else:
                                valeur = unidecode(valeur).replace(".", ",")
                                Fparam.write(1, grx, gamXL, FormTitle)
                                if valeur != "" and valeur.find("|") == -1 and valeur.find(" ") == -1:
                                    valeur = valeur.replace(",",".")
                                    valeur = float(valeur)
                                    Fparam.write_number(refx, grx, valeur, FormNum)
                                else:
                                    Fparam.write(refx, grx, valeur, FormNum)
                            grx += 1
                        refx += 1
                #On ajoute a la fin seuil parametre
                Fparam.write(1,grx,"Seuil Aquifere principal",FormTitle)
                Fparam.write(1,grx + 1, "Seuil Aquitard lateritique", FormTitle)
                typeseuilP = str(type(seuilP))
                typeseuilP2 = str(type(seuilP2))
                if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                    Fparam.write_number(2,grx,seuilP,FormNum)
                else:
                    Fparam.write(2,grx,seuilP,FormNum)
                if typeseuilP2 == "<type 'int'>" or typeseuilP2 == "<type 'float'>":
                    Fparam.write_number(2,grx + 1,seuilP2,FormNum)
                else:
                    Fparam.write(2,grx + 1,seuilP2,FormNum)
            #POUR TOUS LES CAS DE SUIVIS
            MetXLSX.close()
        else:
            arcpy.AddMessage("Aucune donnee pour le fichier Excel de metrique")
    elif Milieu == "Marin":
        if idsuivi == "EAU":
            Annee3a = str(An - 2) + "-" + str(An)
            DonneesPeriodesVoulues = [d for d in DonneesPeriodes if int(str(d[indiceColMar["Periode"]])[-4:]) in Annees and d[indiceColMar["TypeData"]] in metrAttendu or d[indiceColMar["Periode"]] == Annee3a and d[indiceColMar["TypeData"]] in metrAttendu]
            arcpy.AddMessage(str(len(DonneesPeriodesVoulues)) + " donnees sur les periodes et metriques attendues")
            #Liste unique de station - periode
            StationPer = list(set([dpa[indiceColMar["TypeStation"]] + ";" + dpa[indiceColMar["Zone"]] + ";" + dpa[indiceColMar["TypologieStation"]] + ";" + dpa[indiceColMar["TypeRef"]] + ";" +
                                   dpa[indiceColMar["Station"]] + ";" + dpa[indiceColMar["Periode"]] for dpa in DonneesPeriodesVoulues]))
            StationPer = [stp.split(";") for stp in StationPer]
            arcpy.AddMessage(str(len(StationPer)) + " combinaison unique station-periode")
            #Liste unique de station
            Station = list(set([st[indiceColMar["TypeStation"]-1] + ";" + st[indiceColMar["Zone"]-1] + ";" + st[indiceColMar["TypologieStation"]-1] + ";" + st[indiceColMar["TypeRef"]-1] + ";" +
                                st[indiceColMar["Station"]-1] for st in StationPer]))
            NbOrdreStat = len(OrdreStation)
            iord = 1
            StationOrdonnee = list()
            Station = [st.split(";") for st in Station]
            while iord < NbOrdreStat + 1:
                if TypeSuivi == "Reference":
                    StatO = [os[1] for os in OrdreStation if os[8] == iord]
                else:
                    StatO = [os[1] for os in OrdreStation if os[7] == iord]
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
            arcpy.AddMessage(str(len(Station)) + " Stations concernees")
            EquivParam = [list(ep) for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite", "SeuilBilanEnv_Littoral", "SeuilBilanEnv_Cotier", "SeuilBilanEnv_Oceanique", "BilanEnv_ZonEco_Metrique"],"id_cat1 = '" + idsuivi + "' OR id_cat2 = '" + idsuivi + "'")]
            #Equivalence Metrique
            EquivMetr = [[em[0],em[1], em[2]] for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml", "metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
            # Equivalence Gamme de reference
            EquivGamme = [list(eg) for eg in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
            metrAttendu3a = [met for met in metrAttendu if met.find("3 ans") >= 0]
            metrAttendu1a = [met for met in metrAttendu if met.find("3 ans") < 0]
            nbmet = len(metrAttendu1a)
            EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu1a]
            arcpy.AddMessage(EMetr)
            nbmet3 = len(metrAttendu3a)
            EMetr3a = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu3a]
            arcpy.AddMessage(EMetr3a)
            #dictionnaire de AnInd annee et indice des colonnes en fonctions du nombre de metrique voulue ensuite pour fusionner sur la ligne Annee sur Excel
            AnneesEAU = list()
            AnneesEAU.extend(Annees)
            AnneesEAU.append(Annee3a)
            AnInd = AnIndiceExcel(AnneesEAU, nbmet,idsuivi, Annee3a,nbmet3)
            #dictionnaire de AnMet pour les annees normales et AnMet3 pour les metriques sur les 3 dernieres annees avec mise en forme
            met_init = 3
            AnMetAll = AnMetriqueExcel(AnneesEAU,EMetr,AnInd,FormTitle,Annee3a,met_init,idsuivi,EMetr3a,FormTitreExcelRS=FormTitleNoteRS)
            AnMet = AnMetAll[0]
            AnMet3 = AnMetAll[1]
            del AnMetAll
            arcpy.AddMessage("dictionnaire Metrique XML et Excel " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            #dictionnaire avec comme valeurs metrique pour XML balise (metr) et metrique pour le nom en Excel (metrXL)
            Metr = EquivalenceMetriqueXMLExcel(metrAttendu, EquivMetr)
            # dictionnaire avec parametre en cle nom parametrre et valeurs (nom et unite)
            arcpy.AddMessage("dictionnaire Metrique Parametre Nom et unite " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            #derniere colonne de la derniere annee après la derniere metrique demandee
            CoupleFin = (An, Metr[metrAttendu[len(metrAttendu)-1]][1])
            if not CoupleFin in AnMet.keys():
                CoupleFin = (Annee3a,Metr[metrAttendu[len(metrAttendu)-1]][1])
                LastColAnnee = AnMet3[CoupleFin][1]
            else:
                LastColAnnee = AnMet[CoupleFin][1]
            Paramdict = ParametreDico(Parametres,EquivParam,idsuivi)
            #Donnees Gamme de reference#
            DonneesGammeVoulues = [d for d in DonneesGamme if d[0] in GRAttendu]
            #Type de reference non utilisee pour la colonne d'eau marine car les ref se font par station
            TypeRef = list(set([dtgr[indiceColMar["TypologieStation"]] for dtgr in DonneesGammeVoulues]))
            # dictionnaire periode station avec station en cle periodes en valeurs
            arcpy.AddMessage("dictionnaire Periodes par station " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            PerStdict = PeriodeStat(Station, StationPer)
            #dictionnaire pour les scores de l'annee precedente avec en cle les noms des feuilles (parametre nom reduit) et en valeurs dico clé station = liste station et clé scores liste des scores
            Scoredict = dict()
            if ExcelPrec != 0:
                FExcP = xlrd.open_workbook(ExcelPrec)
                NamesFeuilles = FExcP.sheet_names()
                for nf in NamesFeuilles:
                    if nf != "grilles":
                        ValsScore = list()
                        StatsPrec = list()
                        FeuilleP = FExcP.sheet_by_name(nf)
                        NamesCol = [c.value for c in FeuilleP.row(1)]
                        ColScore = "Score " + str(An - 1)
                        if ColScore in NamesCol:
                            iScore = NamesCol.index(ColScore)
                            ValsScore.extend([c.value for c in FeuilleP.col(iScore)][3:])
                            NamesCol = [c.value for c in FeuilleP.row(2)]
                            if "Nom" in NamesCol:
                                iStat = NamesCol.index("Nom")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                            elif "Station" in NamesCol:
                                iStat = NamesCol.index("Station")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                        Scoredict[nf] = {"Stations": StatsPrec, "Scores":ValsScore}
                del FExcP
            FeuillesNA = list()
            #pour les types de valeurs attendues
            typeStr = ["str","unicode"]
            typeNum = ["int","float"]
            typeAll = list()
            typeAll.extend(typeNum)
            typeAll.extend(typeStr)
            for p in Parametres:
                #seuil du parametre pour note reference
                #seuilP = Paramdict[p[1]][2]
                ip = Parametres.index(p)
                arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ipdt = ip + indiceColMar["PremierParam"]
                Fparam = MetXLSX.add_worksheet(Paramdict[p[1]]["name"])
                Fparam.merge_range(0,0,0,2, Paramdict[p[1]]["name"] + " (" + Paramdict[p[1]]["unite"] + ")", FormTitleNoteVC)
                if TypeSuivi == "Suivi":
                    Fparam.merge_range(1,0,1,2, "Stations de la colonne d'eau marine", FormSt)
                else:
                    Fparam.merge_range(1, 0, 1, 2, "Station de reference", FormStR)
                Fparam.write(2,0,"Typologie", FormTitle)
                Fparam.write(2,1,"Zone", FormTitle)
                Fparam.write(2,2,"Station", FormTitle)
                Fparam.freeze_panes(3,3)
                # Longueur des colonnes
                Fparam.set_column(0, 2, 20)
                Fparam.set_column(3, LastColAnnee, 10)
                Fparam.set_column(LastColAnnee + 1, LastColAnnee + 6, 18)
                Fparam.set_column(LastColAnnee + 8, LastColAnnee + 17, 18)
                #reference a mettre en forme
                # On met en forme les cellules de valeur avant d'integrer
                LigneFin = len(Station) + 2
                ColonneFin = AnInd[str(An)][1] + 3
                ldeb = 3
                while ldeb <= LigneFin:
                    cdeb = 3
                    while cdeb <= ColonneFin:
                        Fparam.write(ldeb, cdeb, "", FormN)
                        cdeb += 1
                    ldeb += 1
                sx = 3
                met_init = 3
                for an in AnneesEAU:
                    #Fparam.merge_range(1,AnInd[str(an)][0] + met_init, 1, AnInd[str(an)][1] + met_init, str(an), FormTitle)
                    if an == Annee3a:
                        Fparam.merge_range(1,AnInd[str(an)][0] + met_init, 1, AnInd[str(an)][1] + met_init, str(an), FormTitleNoteRS)
                        for met3 in EMetr3a:
                            Fparam.write(AnMet3[(an,met3[0])][0],AnMet3[(an,met3[0])][1], AnMet3[(an,met3[0])][2],AnMet3[(an,met3[0])][3])
                    else:
                        Fparam.merge_range(1,AnInd[str(an)][0] + met_init, 1, AnInd[str(an)][1] + met_init, str(an), FormTitle)
                        for met2 in EMetr:
                            Fparam.write(AnMet[(an,met2[0])][0],AnMet[(an,met2[0])][1], AnMet[(an,met2[0])][2],AnMet[(an,met2[0])][3])
                for s in Station:
                    Fparam.write(sx,0, s[indiceColMar["TypologieStation"]-1],FormN)
                    Fparam.write(sx,1, s[indiceColMar["Zone"]-1],FormN)
                    Fparam.write(sx,2, s[indiceColMar["Station"]-1],FormN)
                    for perst in PerStdict[s[indiceColMar["Station"]-1]]:
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
                        if str(perst) == Annee3a:
                            Zoneco="Non"
                            for ma in EMetr3a:
                                ima = EMetr3a.index(ma)
                                #position de la valeur dans le tableau par rapport aux colonnes metrique et periode (anneee)
                                metx = ima + met_init + AnInd[str(perstxml)][0]
                                valeur = [dpv[ipdt] for dpv in DonneesPeriodesVoulues if dpv[indiceColMar["TypeData"]] == ma[1] and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1]
                                          and dpv[indiceColMar["TypologieStation"]] == s[indiceColMar["TypologieStation"]-1] and dpv[indiceColMar["TypeRef"]] == s[indiceColMar["TypeRef"]-1]
                                          and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1] and dpv[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]
                                          and dpv[indiceColMar["Periode"]] == perst]
                                #TraitementValeur(val, Feuil, typesTT, typesNUM, typesSTR, irow, icol, Format, FormatDec, FormatEnt):
                                NoteComparaison, valeur = TraitementValeur(valeur, Fparam, typeAll, typeNum, typeStr, sx, metx, FormN, FormNum, FormNumEnt,"Metrique")
                                #GRILLE ZONECO
                                if NoteComparaison == "Oui":
                                    ColNoteZoneco = LastColAnnee + 1
                                    if ma[0] == "Med 3 ans" and Paramdict[p[1]]["metrique"] == "Med 3 ans":
                                        if s[indiceColMar["TypologieStation"]-1] == "Fond de baie, littoral":
                                            seuils = Paramdict[p[1]]["littoral"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                        elif s[indiceColMar["TypologieStation"]-1] == "Lagon en milieu cotier":
                                            seuils = Paramdict[p[1]]["cotier"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                    elif ma[0] == "Per 90 3 ans" and Paramdict[p[1]]["metrique"] == "Per 90 3 ans":
                                        if s[indiceColMar["TypologieStation"]-1] == "Fond de baie, littoral":
                                            seuils = Paramdict[p[1]]["littoral"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                        elif s[indiceColMar["TypologieStation"]-1] == "Lagon en milieu cotier":
                                            seuils = Paramdict[p[1]]["cotier"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                    elif ma[0] == "Moy 3 ans" and Paramdict[p[1]]["metrique"] == "Moy 3 ans":
                                        if s[indiceColMar["TypologieStation"]-1] == "Fond de baie, littoral":
                                            seuils = Paramdict[p[1]]["littoral"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                        elif s[indiceColMar["TypologieStation"]-1] == "Lagon en milieu cotier":
                                            seuils = Paramdict[p[1]]["cotier"].split("|")
                                            seuils = [float(sl) for sl in seuils]
                                            SeuilsZoneco(valeur, seuils, Fparam, sx, ColNoteZoneco, FormNonPerturb, FormNonSeuil, FormFortPerturb)
                                            Zoneco = "Oui"
                                    elif Zoneco == "Non":
                                        Fparam.write(sx, ColNoteZoneco, "Non concerne",FormInco)
                        else:
                            for ma in EMetr:
                                im = EMetr.index(ma)
                                #position de la valeur dans le tableau par rapport aux colonnes metrique et periode (anneee)
                                metx = im + met_init + AnInd[str(perstxml)][0]
                                #valeur recuperee
                                valeur = [dpv[ipdt] for dpv in DonneesPeriodesVoulues if dpv[indiceColMar["TypeData"]] == ma[1] and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1]
                                          and dpv[indiceColMar["TypologieStation"]] == s[indiceColMar["TypologieStation"]-1] and dpv[indiceColMar["TypeRef"]] == s[indiceColMar["TypeRef"]-1]
                                          and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1] and dpv[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]
                                          and dpv[indiceColMar["Periode"]] == perst]
                                #Traitement et integration de la valeur au tableau + Definition si note de comparaison
                                NoteComparaison, valeur = TraitementValeur(valeur, Fparam, typeAll, typeNum, typeStr, sx, metx, FormN, FormNum, FormNumEnt,"Metrique")
                                if NoteComparaison == "Oui":
                                    ####COMPARTIMENT NOTE DE COMPARAISON############
                                    ###NOTE COMPARAISON REFERENCE TEMPORELLE
                                    ColNoteRefTemp = LastColAnnee + 2
                                    ColNoteRefSpat= LastColAnnee + 3
                                    if ma[0] == "Moy" and int(perst) == An:
                                        gr_p90 = "Gamme de reference - Percentile 90 Moyenne"
                                        Perc90GR = [dgr[ipdt] for dgr in DonneesGammeVoulues if dgr[indiceColMar["TypeData"]] == gr_p90 and dgr[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]]
                                        if len(Perc90GR) > 0:
                                            Perc90GR = Perc90GR[0]
                                            tPerc90GR = str(type(Perc90GR))
                                            if tPerc90GR == "<type 'str'>" or tPerc90GR == "<type 'unicode'>":
                                                seuilP = float(Perc90GR.replace(",","."))
                                            else:
                                                seuilP = Perc90GR
                                        else:
                                            seuilP = ""
                                        typeseuilP = str(type(seuilP))
                                        if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                            if valeur >= seuilP:
                                                Fparam.write(sx, ColNoteRefTemp, "Fortement perturbe", FormFortPerturb)
                                            else:
                                                Fparam.write(sx, ColNoteRefTemp, "Non perturbe", FormNonPerturb)
                                        else:
                                            Fparam.write(sx, ColNoteRefTemp, "Pas de seuil", FormNonSeuil)
                                    ###NOTE COMPARAISON REFERENCE SPATIAlE donnees seulement ST06 et ST09 donc pas besoind de filtre sur station
                                    elif ma[0] == "Ecart Moy" and int(perst) == An:
                                        gr_p90sp = "Gamme de reference - Percentile 90 Ecart Moyenne"
                                        Perc90SP_GR = [dgr[ipdt] for dgr in DonneesGammeVoulues if dgr[indiceColMar["TypeData"]] == gr_p90sp and dgr[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]]
                                        if len(Perc90SP_GR) > 0:
                                            Perc90SP_GR = Perc90SP_GR[0]
                                            tPerc90SP_GR = str(type(Perc90SP_GR))
                                            if tPerc90SP_GR == "<type 'str'>" or tPerc90SP_GR == "<type 'unicode'>":
                                                seuilP_SP = float(Perc90SP_GR.replace(",","."))
                                            else:
                                                seuilP_SP = Perc90SP_GR
                                        else:
                                            seuilP_SP = ""
                                        typeseuilP_SP = str(type(seuilP_SP))
                                        if typeseuilP_SP == "<type 'int'>" or typeseuilP_SP == "<type 'float'>":
                                            if valeur >= seuilP_SP:
                                                Fparam.write(sx, ColNoteRefSpat, "Fortement perturbe", FormFortPerturb)
                                            else:
                                                Fparam.write(sx, ColNoteRefSpat, "Non perturbe", FormNonPerturb)
                                        else:
                                            Fparam.write(sx, ColNoteRefSpat, "Pas de seuil", FormNonSeuil)
                    ###RECUPERATION DES ANCIENS SCORES
                    ColScoreOld = LastColAnnee + 5
                    if ExcelPrec != 0:
                        if Paramdict[p[1]]["name"] in Scoredict.keys():
                            #si le nom de la station est dans les resultats precedents
                            if s[indiceColMar["Station"]-1] in Scoredict[Paramdict[p[1]]["name"]]["Stations"]:
                                istp = Scoredict[Paramdict[p[1]]["name"]]["Stations"].index(s[indiceColMar["Station"]-1])
                                if Scoredict[Paramdict[p[1]]["name"]]["Scores"][istp] == "Bon" or Scoredict[Paramdict[p[1]]["name"]]["Scores"][istp] == "Non perturbe":
                                    Fparam.write(sx, ColScoreOld, "Non perturbe", FormNonPerturb)
                                elif Scoredict[Paramdict[p[1]]["name"]]["Scores"][istp] == "Mauvais" or Scoredict[Paramdict[p[1]]["name"]]["Scores"][istp] == "Fortement perturbe":
                                    Fparam.write(sx, ColScoreOld, "Fortement perturbe", FormFortPerturb)
                                elif Scoredict[Paramdict[p[1]]["name"]]["Scores"][istp] == "Inconnu":
                                    Fparam.write(sx, ColScoreOld, "Inconnu", FormInco)
                            else:
                                arcpy.AddMessage("Pas de correspondance avec " + s[indiceColMar["Station"]-1] + " pour " + Paramdict[p[1]]["name"])
                        else:
                            if Paramdict[p[1]]["name"] not in FeuillesNA:
                                arcpy.AddMessage("La feuille " + Paramdict[p[1]]["name"] + " n'existe pas dans " + ExcelPrec)
                                FeuillesNA.append(Paramdict[p[1]]["name"])
                    sx += 1
                #arcpy.AddMessage("Station traitee " + s[indiceColMar["Station"]-2]  + " pour le fichier XML")
                # colonne Note comparaison initialise a 0
                notex = LastColAnnee + 1
                if notex > 0:
                    Fparam.merge_range(1,notex, 2, notex, "Note grille ZONECO", FormTitleNoteRS)
                else:
                    Fparam.merge_range(1,notex+1, 2, Fparam.dim_colmax+1, "Note grille ZONECO", FormTitleNoteRS)
                Fparam.merge_range(1,notex + 1, 2, notex + 1, "Note comparaison reference temporelle", FormTitleNoteLV)
                Fparam.merge_range(1,notex + 2, 2, notex + 2, "Note comparaison reference spatiale", FormTitleNote)
                Fparam.merge_range(1,notex + 3, 2, notex + 3, "Evolution temporelle", FormTitleEvoTemp)
                Fparam.merge_range(1,notex + 4, 2, notex + 4, "Score " + str(An - 1), FormTitleScoreOld)
                Fparam.merge_range(1,notex + 5, 2, notex + 5, "Score " + str(An), FormTitleScore)
                arcpy.AddMessage("Parametre " + p[1] + " traite pour les metriques " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ''' REVOIR GAMME DE REFERENCE APRES AVOIR VU LE FICHIER BILANENV'''
                ##############################GAMME DE REFERENCE#################################################
                debColGamme = Fparam.dim_colmax + 2
                Fparam.merge_range(1, debColGamme, 1, debColGamme + 4, "Gamme de reference", FormTitleNote)
                Fparam.write(2, debColGamme, "Typologie", FormTitle)
                Fparam.write(2, debColGamme + 1, "Zone", FormTitle)
                Fparam.write(2, debColGamme + 2, "Stations", FormTitle)
                Fparam.write(2, debColGamme + 3, "Reference spatiale", FormTitle)
                Fparam.write(2, debColGamme + 4, "Periode", FormTitle)
                refx = 3
                #fonctionnement different du milieu Eaux douces par station des references
                for s in Station:
                    DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]]
                    Fparam.write(refx, debColGamme + 4, DonneesGVSelect[0][indiceColMar["Periode"]], FormN)
                    Fparam.write(refx, debColGamme, DonneesGVSelect[0][indiceColMar["TypologieStation"]], FormN)
                    Fparam.write(refx, debColGamme + 1, DonneesGVSelect[0][indiceColMar["Zone"]], FormN)
                    ref = DonneesGVSelect[0][indiceColMar["TypeRef"]]
                    if not ref is None and ref.find("Suivi EAU") >= 0:
                        Fparam.write(refx, debColGamme + 3, DonneesGVSelect[0][indiceColMar["TypeRef"]].replace("Suivi EAU|",""), FormN)
                    else:
                        Fparam.write(refx, debColGamme + 3, "", FormN)
                    Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceColMar["Station"]], FormN)
                    grx = debColGamme + 5
                    for gr in GRAttendu:
                        DonneesGVSel = [dgv for dgv in DonneesGVSelect if dgv[indiceColMar["TypeData"]] == gr]
                        gamXL = [gm[2] for gm in EquivGamme if gm[0] == gr]
                        if len(gamXL) > 0:
                            gamXL = gamXL[0]
                        else:
                            arcpy.AddMessage("Correspondance non trouvee pour la gamme attendue: " + gr)
                            gamXL = gr
                        Fparam.write(2, grx, gamXL, FormTitle)
                        if len(DonneesGVSel) > 0:
                            valeur = [dga[ipdt] for dga in DonneesGVSel]
                            valeur = TraitementValeur(valeur,Fparam,typeAll,typeNum,typeStr, refx, grx, FormN, FormNum, FormNumEnt, "Reference")
                        grx += 1
                    refx += 1
                #on verifie et ajoute si besoin grille ZONECO
                if Paramdict[p[1]]["metrique"] != "" and Paramdict[p[1]]["metrique"] != None:
                    debL = refx+2
                    Fparam.merge_range(debL, debColGamme, debL, debColGamme + 3, "Reference Grille ZONECO", FormTitleNoteRS)
                    Fparam.write(debL + 1, debColGamme,"Typologie",FormTitle)
                    Fparam.write(debL + 1, debColGamme + 1,"Seuil",FormTitle)
                    Fparam.write(debL + 1, debColGamme + 2,"Qualite",FormTitle)
                    Fparam.write(debL + 1, debColGamme + 3,"Metrique",FormTitle)
                    Typologies = {0:"littoral",1:"cotier",2:"oceanique"}
                    lze = debL + 2
                    for typo in Typologies.keys():
                        Fparam.merge_range(lze,debColGamme,lze+2,debColGamme,Typologies[typo],FormTitle)
                        seuilsL = Paramdict[p[1]][Typologies[typo]].split("|")
                        Fparam.write(lze,debColGamme + 1,seuilsL[0] + " <",FormN)
                        Fparam.write(lze,debColGamme + 2, "Non perturbe",FormNonPerturb)
                        Fparam.write(lze + 1,debColGamme + 1,"[" + seuilsL[0] + " - " + seuilsL[1] + "[",FormN)
                        Fparam.write(lze + 1,debColGamme + 2, "Moderement perturbe",FormNonSeuil)
                        Fparam.write(lze + 2,debColGamme + 1,">= " + seuilsL[1],FormN)
                        Fparam.write(lze + 2,debColGamme + 2, "Fortement perturbe",FormFortPerturb)
                        lze+=3
                    Fparam.merge_range(debL + 2, debColGamme + 3, lze-1, debColGamme + 3, Paramdict[p[1]]["metrique"],FormN)
                #On ajoute a la fin seuil parametre
                # Fparam.write(1, grx,"Seuil",FormTitle)
                # typeseuilP = str(type(seuilP))
                # if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                #     Fparam.write_number(2,grx,seuilP,FormNum)
                # else:
                #     Fparam.write(2,grx,seuilP,FormNum)
            #POUR TOUS LES CAS DE SUIVIS
            MetXLSX.close()
        elif idsuivi == "FLUX":
            DonneesPeriodesVoulues = [d for d in DonneesPeriodes if int(str(d[indiceColMar["Periode"]])[-4:]) in Annees and d[indiceColMar["TypeData"]] in metrAttendu]
            arcpy.AddMessage(str(len(DonneesPeriodesVoulues)) + " donnees sur les periodes et metriques attendues")
            #Liste unique de station - periode
            StationPer = list(set([dpa[indiceColMar["TypeStation"]] + ";" + dpa[indiceColMar["Zone"]] + ";" + dpa[indiceColMar["TypologieStation"]] + ";" + dpa[indiceColMar["TypeRef"]] + ";" +
                                   dpa[indiceColMar["Station"]] + ";" + dpa[indiceColMar["Periode"]] for dpa in DonneesPeriodesVoulues]))
            StationPer = [stp.split(";") for stp in StationPer]
            arcpy.AddMessage(str(len(StationPer)) + " combinaison unique station-periode")
            #Liste unique de station
            Station = list(set([st[indiceColMar["TypeStation"]-1] + ";" + st[indiceColMar["Zone"]-1] + ";" + st[indiceColMar["TypologieStation"]-1] + ";" + st[indiceColMar["TypeRef"]-1] + ";" +
                                st[indiceColMar["Station"]-1] for st in StationPer]))
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
            arcpy.AddMessage(str(len(Station)) + " Stations concernees")
            #Equivalence Parametre
            EquivParam = [[ep[0],ep[1],ep[2]] for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite"],"id_cat1 = '" + idsuivi + "' OR id_cat2 = '" + idsuivi + "'")]
            EquivMetr = [[em[0],em[1], em[2]] for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml", "metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
            # Equivalence Gamme de reference
            EquivGamme = [list(eg) for eg in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
            nbmet = len(metrAttendu)
            EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu]
            #dictionnaire Annee indice colonne fusion pour la ligne Annees sous Excel par feuille Excel de parametre
            AnInd = AnIndiceExcel(Annees,nbmet,idsuivi)
            #dictionnaire AnMet pour placer le titre des metrique par annee pour les annees normales et AnMet3 pour les metriques sur les 3 dernieres annees avec mise en forme
            met_init = 3
            AnMet = AnMetriqueExcel(Annees, EMetr, AnInd, FormTitle,IndicePremColMet=met_init)
            # dictionnaire avec comme valeurs metrique pour XML balise (metr) et metrique pour le nom en Excel (metrXL)
            Metr = EquivalenceMetriqueXMLExcel(metrAttendu, EquivMetr)
            #derniere colonne de la derniere annee après la derniere metrique demandee
            LastColAnnee = AnMet[An, Metr[metrAttendu[len(metrAttendu)-1]]][1]
            #dictionnaire avec parametre en cle nom parametrre et valeurs (nom et unite)
            Paramdict = ParametreDico(Parametres, EquivParam)
            #Donnees Gamme de reference#
            DonneesGammeVoulues = [d for d in DonneesGamme if d[0] in GRAttendu]
            '''A changer ici selon les gammes attendues pour les references'''
            TypeRef = list(set([dtgr[indiceColMar["Localisation"]] for dtgr in DonneesGammeVoulues]))
            arcpy.AddMessage(TypeRef)
            #dictionnaire  de periode par station avec comme valeurs periodes et stations en cle
            PerStdict = PeriodeStat(Station, StationPer)
            #dictionnaire pour les scores de l'annee precedente avec en cle les noms des feuilles (parametre nom reduit) et en valeurs dico clé station = liste station et clé scores liste des scores
            Scoredict = dict()
            if ExcelPrec != 0:
                FExcP = xlrd.open_workbook(ExcelPrec)
                NamesFeuilles = FExcP.sheet_names()
                for nf in NamesFeuilles:
                    if nf != "grilles":
                        ValsScore = list()
                        StatsPrec = list()
                        FeuilleP = FExcP.sheet_by_name(nf)
                        NamesCol = [c.value for c in FeuilleP.row(1)]
                        ColScore = "Score " + str(An - 1)
                        if ColScore in NamesCol:
                            iScore = NamesCol.index(ColScore)
                            ValsScore.extend([c.value for c in FeuilleP.col(iScore)][3:])
                            NamesCol = [c.value for c in FeuilleP.row(2)]
                            if "Nom" in NamesCol:
                                iStat = NamesCol.index("Nom")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                            elif "Station" in NamesCol:
                                iStat = NamesCol.index("Station")
                                StatsPrec.extend([c.value for c in FeuilleP.col(iStat)][3:])
                        Scoredict[nf] = {"Stations": StatsPrec, "Scores":ValsScore}
                del FExcP
            FeuillesNA = list()
            for p in Parametres:
                #seuil du parametre pour note reference
                seuilP = Paramdict[p[1]][2]
                ip = Parametres.index(p)
                arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ipdt = ip + indiceColMar["PremierParam"]
                Fparam = MetXLSX.add_worksheet(Paramdict[p[1]][0])
                Fparam.merge_range(0,0,0,2, Paramdict[p[1]][0] + " (" + Paramdict[p[1]][1] + ")", FormTitle)
                if TypeSuivi == "Suivi":
                    Fparam.merge_range(1,0,1,2, "Station de suivi", FormSt)
                else:
                    Fparam.merge_range(1, 0, 1, 2, "Station de reference", FormStR)
                Fparam.write(2,0,"position", FormTitle)
                Fparam.write(2,1,"zone", FormTitle)
                Fparam.write(2,2,"nom", FormTitle)
                Fparam.freeze_panes(3,3)
                # Longueur des colonnes
                Fparam.set_column(0, 2, 20)
                Fparam.set_column(3, LastColAnnee, 10)
                Fparam.set_column(LastColAnnee + 1, LastColAnnee + 5, 18)
                Fparam.set_column(LastColAnnee + 7, LastColAnnee + 17, 1)
                #reference a mettre en forme
                # On met en forme les cellules de valeur avant d'integrer
                LigneFin = len(Station) + 2
                ColonneFin = AnInd[str(An)][1] + 3
                ldeb = 3
                while ldeb <= LigneFin:
                    cdeb = 3
                    while cdeb <= ColonneFin:
                        Fparam.write(ldeb, cdeb, "", FormN)
                        cdeb += 1
                    ldeb += 1
                sx = 3
                met_init = 3
                for an in Annees:
                    Fparam.merge_range(1,AnInd["sf_" + str(an)][0] + met_init, 1, AnInd["sf_" + str(an)][1] + met_init, "Saison fraiche " + str(an), FormTitle)
                    Fparam.merge_range(1,AnInd["sc_" + str(an)][0] + met_init, 1, AnInd["sc_" + str(an)][1] + met_init, "Saison chaude " + str(an), FormTitle)
                    for met2 in EMetr:
                        im = EMetr.index(met2)
                        Fparam.write(AnMet[("sf_" + str(an), met2[0])][0],AnMet[("sf_" + str(an), met2[0])][1], AnMet[("sf_" + str(an), met2[0])][2],AnMet[("sf_" + str(an), met2[0])][3])
                        Fparam.write(AnMet[("sc_" + str(an), met2[0])][0],AnMet[("sc_" + str(an), met2[0])][1], AnMet[("sc_" + str(an), met2[0])][2],AnMet[("sc_" + str(an), met2[0])][3])
                for s in Station:
                    Fparam.write(sx,0, s[indiceColMar["TypologieStation"]-1],FormN)
                    Fparam.write(sx,1, s[indiceColMar["Zone"]-1],FormN)
                    Fparam.write(sx,2, s[indiceColMar["Station"]-1],FormN)
                    for perst in PerStdict[s[indiceColMar["Station"]-1]]:
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
                            ima = EMetr.index(ma)
                            valeur = [dpv[ipdt] for dpv in DonneesPeriodesVoulues if dpv[indiceColMar["TypeData"]] == ma[1] and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1]
                                      and dpv[indiceColMar["TypologieStation"]] == s[indiceColMar["TypologieStation"]-1] and dpv[indiceColMar["TypeRef"]] == s[indiceColMar["TypeRef"]-1]
                                      and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1] and dpv[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]
                                      and dpv[indiceColMar["Periode"]] == perst]
                            if len(valeur) > 0:
                                valeur = valeur[0]
                                if valeur == "" or valeur == None:
                                    valeur = ""
                            else:
                                valeur = "NA"
                            valeur = str(valeur).replace(".",",")
                            metx = ima + met_init + AnInd[str(perstxml)][0]
                            if valeur != "" and valeur.find("|") == -1:
                                valeur = valeur.replace(",", ".")
                                valeur = float(valeur)
                                Fparam.write_number(sx,metx,valeur,FormNum)
                                ###NOTE COMPARAISON REFERENCE
                                if ma[0] == "% > Perc 75" and int(perst) == An:
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                            else:
                                Fparam.write(sx,metx,valeur,FormN)
                    ###RECUPERATION DES ANCIENS SCORES
                    if ExcelPrec != 0:
                        if Paramdict[p[1]][0] in Scoredict.keys():
                            #si le nom de la station est dans les resultats precedents
                            if s[indiceColMar["Station"]-1] in Scoredict[Paramdict[p[1]][0]]["Stations"]:
                                istp = Scoredict[Paramdict[p[1]][0]]["Stations"].index(s[indiceColMar["Station"]-1])
                                if Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Bon" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Non perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Non perturbe", FormNonPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Mauvais" or Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Fortement perturbe":
                                    Fparam.write(sx, LastColAnnee + 4, "Fortement perturbe", FormFortPerturb)
                                elif Scoredict[Paramdict[p[1]][0]]["Scores"][istp] == "Inconnu":
                                    Fparam.write(sx, LastColAnnee + 4, "Inconnu", FormInco)
                            else:
                                arcpy.AddMessage("Pas de correspndance avec " + s[indiceColMar["Station"]-1] + " pour " + Paramdict[p[1]][0])
                        else:
                            if Paramdict[p[1]][0] not in FeuillesNA:
                                arcpy.AddMessage("La Feuille " + Paramdict[p[1]][0] + " n'existe pas dans " + ExcelPrec)
                                FeuillesNA.append(Paramdict[p[1]][0])
                    sx += 1

                    #arcpy.AddMessage("Station traitee " + s[indiceColMar["Station"]-2]  + " pour le fichier XML")
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Note comparaison reference", FormTitleNote)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Evolution temporelle", FormTitle)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2017", FormTitle)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2018", FormTitle)
                Fparam.set_column(0,Fparam.dim_colmax,18)
                arcpy.AddMessage("Parametre " + p[1] + " traite pour les metriques " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ##############################GAMME DE REFERENCE#################################################
                debColGamme = Fparam.dim_colmax + 2
                Fparam.merge_range(0, debColGamme, 0, debColGamme + 4, "Gamme de reference", FormTitleNote)
                Fparam.write(1, debColGamme, "Type", FormTitle)
                Fparam.write(1, debColGamme + 1, "Stations", FormTitle)
                Fparam.write(1, debColGamme + 2, "Periode", FormTitle)
                refx = 2
                for ref in TypeRef:
                    DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceColMar["Localisation"]] == ref]
                    #recherche de la periode particuliere si Seuil Riviere OEIL
                    if ref == "Seuil Riviere OEIL":
                        EquivParamSel = [ep for ep in EquivParam if ep[0] == p[1]]
                        if len(EquivParamSel) > 0:
                            EquivParamSel = EquivParamSel[0]
                            PerG = EquivParamSel[5]
                            Fparam.write(refx, debColGamme + 2, PerG, FormN)
                        else:
                            Fparam.write(refx, debColGamme + 2, "", FormN)
                    else:
                        Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceColMar["Periode"]], FormN)

                    Fparam.write(refx, debColGamme, ref, FormN)
                    Fparam.write(refx, debColGamme + 1, DonneesGVSelect[0][indiceColMar["Station"]], FormN)

                    grx = debColGamme + 3
                    for gr in GRAttendu:
                        DonneesGVSel = [dgv for dgv in DonneesGVSelect if dgv[indiceColMar["TypeData"]] == gr]
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
                        if str(type(valeur)) == "<type 'int'>" or str(type(valeur)) == "<type 'float'>":
                            Fparam.write(refx, grx, valeur, FormNum)
                            Fparam.write(1, grx, gamXL, FormTitle)
                        else:
                            valeur = unidecode(valeur).replace(".", ",")
                            Fparam.write(1, grx, gamXL, FormTitle)
                            if valeur != "" and valeur.find("|") == -1 and valeur.find(" ") == -1:
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
            MetXLSX.close()
        else:
            DonneesPeriodesVoulues = [d for d in DonneesPeriodes if int(str(d[indiceColMar["Periode"]])[-4:]) in Annees and d[indiceColMar["TypeData"]] in metrAttendu]
            arcpy.AddMessage(str(len(DonneesPeriodesVoulues)) + " donnees sur les periodes et metriques attendues")
            #Liste unique de station - periode
            StationPer = list(set([dpa[indiceColMar["TypeStation"]] + ";" + dpa[indiceColMar["Zone"]] + ";" + dpa[indiceColMar["TypologieStation"]] + ";" + dpa[indiceColMar["TypeRef"]] + ";" +
                                   dpa[indiceColMar["Station"]] + ";" + dpa[indiceColMar["Periode"]] for dpa in DonneesPeriodesVoulues]))
            StationPer = [stp.split(";") for stp in StationPer]
            arcpy.AddMessage(str(len(StationPer)) + " combinaison unique station-periode")
            #Liste unique de station
            Station = list(set([st[indiceColMar["TypeStation"]-1] + ";" + st[indiceColMar["Zone"]-1] + ";" + st[indiceColMar["TypologieStation"]-1] + ";" + st[indiceColMar["TypeRef"]-1] + ";" +
                                st[indiceColMar["Station"]-1] for st in StationPer]))
            Station = [st.split(";") for st in Station]
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
            arcpy.AddMessage(str(len(Station)) + " Stations concernees")
            #Equivalence Parametre
            if idsuivi == "RECOUVR32":
                EquivParam = [[ep[0],ep[1],ep[2]] for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite"],"id_parametre LIKE 'SUBSTR/RECOUVR/%'")]
            elif idsuivi == "RECOUVR12":
                EquivParam = [[ep[0],ep[1],ep[2]] for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite"],"id_parametre LIKE 'SUBSTR/ACROPORA/RECOUVR/%'")]
            else:
                EquivParam = [[ep[0],ep[1],ep[2]] for ep in arcpy.da.SearchCursor(XMLParamEquiv,["parametre","xml_param","unite"],"id_cat1 = '" + idsuivi + "' OR id_cat2 = '" + idsuivi + "'")]
            #Equivalence Metrique
            EquivMetr = [[em[0],em[1], em[2]] for em in arcpy.da.SearchCursor(XMLMetriqueEquiv,["metrique","metrique_xml", "metriqueExcel"],"id_suivi LIKE '%" + idsuivi + "%'")]
            # Equivalence Gamme de reference
            EquivGamme = [list(eg) for eg in arcpy.da.SearchCursor(XMLGammeEquiv, ["gamme_reference", "gamme_xml", "gammeExcel"], "id_suivi LIKE '%" + idsuivi + "%'")]
            nbmet = len(metrAttendu)
            EMetr = [[eme[2],eme[0]] for eme in EquivMetr if eme[0] in metrAttendu]
            #dictionnaire Annee indice colonne fusion pour la ligne Annees sous Excel par feuille Excel de parametre
            AnInd = AnIndiceExcel(Annees, nbmet, idsuivi)
            #dictionnaire AnMet pour placer le titre des metrique par annee pour les annees normales et AnMet3 pour les metriques sur les 3 dernieres annees avec mise en forme
            met_init = 3
            AnMet = AnMetriqueExcel(Annees, AnInd, FormTitle, IndicePremColMet=met_init)
            #dictionnaire avec comme valeurs metrique pour XML balise (metr) et metrique pour le nom en Excel (metrXL)
            Metr = EquivalenceMetriqueXMLExcel(metrAttendu, EquivMetr)
            #derniere colonne de la derniere annee après la derniere metrique demandee
            LastColAnnee = AnMet[An, Metr[metrAttendu[len(metrAttendu)-1]]][1]
            #dictionnaire avec parametre en cle nom parametrre et valeurs (nom et unite)
            Paramdict = ParametreDico(Parametres, EquivParam)
            #dictionnaire  de periode par station avec comme valeurs periodes et stations en cle
            PerStdict = PeriodeStat(Station, StationPer)
            #Donnees Gamme de reference#
            DonneesGammeVoulues = [d for d in DonneesGamme if d[0] in GRAttendu]
            '''A changer ici selon les gammes attendues pour les references'''
            TypeRef = list(set([dtgr[indiceColMar["Localisation"]] for dtgr in DonneesGammeVoulues]))
            arcpy.AddMessage(TypeRef)
            #TRAITEMENT PAR PARAMETRE
            for p in Parametres:
                #seuil du parametre pour note reference
                seuilP = Paramdict[p[1]][2]
                ip = Parametres.index(p)
                arcpy.AddMessage("Debut de traitement pour le parametre " + p[1] + ": " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ipdt = ip + indiceColMar["PremierParam"]
                Fparam = MetXLSX.add_worksheet(Paramdict[p[1]][0])
                Fparam.merge_range(0,0,0,2, Paramdict[p[1]][0] + " (" + Paramdict[p[1]][1] + ")", FormTitle)
                Fparam.merge_range(1,0,1,2, "Station de suivi", FormSt)
                Fparam.write(2,0,"position", FormTitle)
                Fparam.write(2,1,"zone", FormTitle)
                Fparam.write(2,2,"nom", FormTitle)
                Fparam.freeze_panes(3,3)
                # Longueur des colonnes
                Fparam.set_column(0, 2, 20)
                Fparam.set_column(3, LastColAnnee, 10)
                Fparam.set_column(LastColAnnee + 1, LastColAnnee + 5, 18)
                Fparam.set_column(LastColAnnee + 7, LastColAnnee + 17, 1)
                #reference a mettre en forme
                # On met en forme les cellules de valeur avant d'integrer
                LigneFin = len(Station) + 2
                ColonneFin = AnInd[str(An)][1] + 3
                ldeb = 3
                while ldeb <= LigneFin:
                    cdeb = 3
                    while cdeb <= ColonneFin:
                        Fparam.write(ldeb, cdeb, "", FormN)
                        cdeb += 1
                    ldeb += 1
                sx = 3
                met_init = 3
                for an in Annees:
                    Fparam.merge_range(1,AnInd[str(an)][0] + met_init, 1, AnInd[str(an)][1] + met_init, str(an), FormTitle)
                    for met2 in EMetr:
                        im = EMetr.index(met2)
                        Fparam.write(AnMet[(an, met2[0])][0],AnMet[(an, met2[0])][1], AnMet[(an, met2[0])][2],AnMet[(an, met2[0])][3])
                for s in Station:
                    Fparam.write(sx,0, s[indiceColMar["TypologieStation"]-1],FormN)
                    Fparam.write(sx,1, s[indiceColMar["Zone"]-1],FormN)
                    Fparam.write(sx,2, s[indiceColMar["Station"]-1],FormN)
                    for perst in PerStdict[s[indiceColMar["Station"]-1]]:
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
                            ima = EMetr.index(ma)
                            valeur = [dpv[ipdt] for dpv in DonneesPeriodesVoulues if dpv[indiceColMar["TypeData"]] == ma[1] and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1]
                                      and dpv[indiceColMar["TypologieStation"]] == s[indiceColMar["TypologieStation"]-1] and dpv[indiceColMar["TypeRef"]] == s[indiceColMar["TypeRef"]-1]
                                      and dpv[indiceColMar["Zone"]] == s[indiceColMar["Zone"]-1] and dpv[indiceColMar["Station"]] == s[indiceColMar["Station"]-1]
                                      and dpv[indiceColMar["Periode"]] == perst]
                            if len(valeur) > 0:
                                valeur = valeur[0]
                                if valeur == "" or valeur == None:
                                    valeur = ""
                            else:
                                valeur = "NA"
                            valeur = str(valeur).replace(".",",")
                            metx = ima + met_init + AnInd[str(perstxml)][0]
                            if valeur != "" and valeur.find("|") == -1:
                                valeur = valeur.replace(",", ".")
                                valeur = float(valeur)
                                Fparam.write_number(sx,metx,valeur,FormNum)
                                ###NOTE COMPARAISON REFERENCE
                                if ma[0] == "% > Perc 75" and int(perst) == An:
                                    typeseuilP = str(type(seuilP))
                                    if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                        if valeur >= seuilP:
                                            Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                        else:
                                            Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                            else:
                                Fparam.write(sx,metx,valeur,FormN)
                            #metx = EMetr.index(metrXL) + met_init + AnInd[str(perstxml)][0]
                    sx += 1
                    #arcpy.AddMessage("Station traitee " + s[indiceColMar["Station"]-2]  + " pour le fichier XML")
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Note comparaison reference", FormTitleNote)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Evolution temporelle", FormTitle)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2017", FormTitle)
                Fparam.merge_range(1,Fparam.dim_colmax + 1, 2, Fparam.dim_colmax + 1, "Score 2018", FormTitle)
                Fparam.set_column(0,Fparam.dim_colmax,18)
                arcpy.AddMessage("Parametre " + p[1]  + " traite pour le fichier de metrique XML" + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
                ##############################GAMME DE REFERENCE#################################################
                debColGamme = Fparam.dim_colmax + 2
                Fparam.merge_range(0, debColGamme, 0, debColGamme + 4, "Gamme de reference", FormTitleNote)
                Fparam.write(1, debColGamme, "Type", FormTitle)
                Fparam.write(1, debColGamme + 1, "Stations", FormTitle)
                Fparam.write(1, debColGamme + 2, "Periode", FormTitle)
                refx = 2
                for ref in TypeRef:
                    DonneesGVSelect = [dgv for dgv in DonneesGammeVoulues if dgv[indiceColMar["Localisation"]] == ref]
                    #recherche de la periode particuliere si Seuil Riviere OEIL
                    if ref == "Seuil Riviere OEIL":
                        EquivParamSel = [ep for ep in EquivParam if ep[0] == p[1]]
                        if len(EquivParamSel) > 0:
                            EquivParamSel = EquivParamSel[0]
                            PerG = EquivParamSel[5]
                            Fparam.write(refx, debColGamme + 2, PerG, FormN)
                        else:
                            Fparam.write(refx, debColGamme + 2, "", FormN)
                    else:
                        Fparam.write(refx, debColGamme + 2, DonneesGVSelect[0][indiceColMar["Periode"]], FormN)

                    Fparam.write(refx, debColGamme, ref, FormN)
                    Fparam.write(refx, debColGamme + 1, DonneesGVSelect[0][indiceColMar["Station"]], FormN)

                    grx = debColGamme + 3
                    for gr in GRAttendu:
                        DonneesGVSel = [dgv for dgv in DonneesGVSelect if dgv[indiceColMar["TypeData"]] == gr]
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
                        if str(type(valeur)) == "<type 'int'>" or str(type(valeur)) == "<type 'float'>":
                            Fparam.write(refx, grx, valeur, FormNum)
                            Fparam.write(1, grx, gamXL, FormTitle)
                        else:
                            valeur = unidecode(valeur).replace(".", ",")
                            Fparam.write(1, grx, gamXL, FormTitle)
                            if valeur != "" and valeur.find("|") == -1 and valeur.find(" ") == -1:
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

