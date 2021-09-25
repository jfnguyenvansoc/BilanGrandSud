# -*- coding: utf8 -*-
import arcpy, xlsxwriter, xlrd, datetime
from unidecode import unidecode
from ModuleBilanEnv import FormatExcel

#organisation des premieres colonnes
indiceCol = {"TypeData":0, "TypeStation":1, "Localisation":2, "TypeBV":3, "Zone":4, "ZoneReference":5, "Station":6, "Date":7,"Periode":8,"PremierParam":9}

##FONCTION  pour la sortie XLSX Metriques
def ExportMetrique(DonneesPeriodes, DonneesGamme,Parametres,XMLParamEquiv, XMLMetriqueEquiv, XMLGammeEquiv, idsuivi, metrAttendu, GRAttendu, FichierXML, OrdreStation, TypeSuivi, AnEtude, ExcelPrec=0):
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
        LastColAnnee = AnMet[AnneeEtude, Metr[metrAttendu[len(metrAttendu)-1]]][1]
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
                    ColScore = "Score " + str(AnneeEtude - 1)
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
                            if ma[0] == "% > Perc 75" and int(perst) == AnneeEtude:
                                typeseuilP = str(type(seuilP))
                                if typeseuilP == "<type 'int'>" or typeseuilP == "<type 'float'>":
                                    if valeur >= seuilP:
                                        Fparam.write(sx, LastColAnnee + 1, "Fortement perturbe", FormFortPerturb)
                                    else:
                                        Fparam.write(sx, LastColAnnee + 1, "Non perturbe", FormNonPerturb)
                                else:
                                    Fparam.write(sx, LastColAnnee + 1, "Pas de seuil", FormNonSeuil)
                            elif ma[0] == "% > Perc 75 - OEIL" and int(perst) == AnneeEtude and s[indiceCol["Localisation"]-1] == "Riviere":
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
            arcpy.AddMessage("Parametre " + p[1] + " traite pour le fichier de metrique XML: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
            # colonne Note comparaison initialise a 0
            notex = LastColAnnee + 1
            if notex > 0:
                Fparam.merge_range(1, notex, 2, notex, "Note comparaison reference", FormTitleNote)
            else:
                Fparam.merge_range(1,notex+1, 2, Fparam.dim_colmax+1, "Note comparaison reference", FormTitleNote)
            Fparam.merge_range(1, notex + 1, 2, notex + 1, "Note etude seuil", FormTitleNote)
            Fparam.merge_range(1,notex + 2, 2, notex + 2, "Evolution temporelle", FormTitle)
            Fparam.merge_range(1,notex + 3, 2, notex + 3, "Score " + str(AnneeEtude - 1), FormTitle)
            Fparam.merge_range(1,notex + 4, 2, notex + 4, "Score " + str(AnneeEtude), FormTitle)
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
            arcpy.AddMessage("Parametre " + p[1] + " traite pour le fichier de metrique XML: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S"))
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
    ###METTRE elif ICI pour les autres type de suivi SOUT ...

