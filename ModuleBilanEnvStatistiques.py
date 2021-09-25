# -*- coding: utf8 -*-
from ModuleBilanEnv import ConcatenationInformation, AnneeMinMax, FindSigne, FindLQ, FindMethodes, FindStatsCal, FindValsNonVides
from ModuleBilanEnv import NbSupPerc
import arcpy, datetime

# organisation des premieres colonnes
indiceCol = {"TypeData": 0, "TypeStation": 1, "Localisation": 2, "TypeBV": 3, "Zone": 4, "ZoneReference": 5,
             "Station": 6, "Date": 7, "Periode": 8, "PremierParam": 9}


##Calcul des Statistiques de Gamme de Reference
def StatistiquesGammeReference(DonneesM, DonneesS, DonneesEx, Parametres, IdentSuiv, XMLParam=0):
    DonneesStatGammeRef = list()
    # CAS EAUX DOUCES DE SURFACE SEDIMENT Riviere/AVAL GRAND BV/PETIT BV ET DOLINES A PART
    if IdentSuiv == "SURF":
        arcpy.AddMessage("Oh non Gammes de reference pour les eaux de surface")
        # Riviere
        DonneesMRiviere = [d for d in DonneesM if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]
        StatTxtRiviere = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesMRiviere])))
        PeriodeRiviere = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesMRiviere])))
        DonneesSRiviere = [d for d in DonneesS if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]
        DonneesExRiviere = [d for d in DonneesEx if
                            d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]

        # Dolines
        DonneesMDol = [d for d in DonneesM if
                       d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
        if len(DonneesMDol) > 0:
            StatTxtDol = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesMDol])))
            PeriodeDol = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesMDol])))
            DonneesSDol = [d for d in DonneesS if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
            DonneesExDol = [d for d in DonneesEx if
                            d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
        # CAS PARTICULIER Gamme Ref Seuil OEIL seulement pour certains parametres
        if XMLParam != 0:
            SQLPrmX = "id_cat2 = '" + IdentSuiv + "' AND SeuilBilanENV_Etude IS NOT NULL"
            ParamsEtudeOEIL = [list(px) for px in arcpy.da.SearchCursor(XMLParam, ["parametre", "id_parametre",
                                                                                   "SeuilBilanENV_Etude",
                                                                                   "PeriodeBilanENV_Etude",
                                                                                   "LQ_BilanENV_Etude",
                                                                                   "N_LQ_BilanENV_Etude",
                                                                                   "N_BilanENV_Etude",
                                                                                   "Med_BilanENV_Etude",
                                                                                   "Max_BilanENV_Etude"], SQLPrmX)]

        # ligne pour les statistiques de reference Riviere
        LineRiviereMeth = ["Gamme de reference - methodes analytiques", "Reference", "Riviere", "", "", "",
                           StatTxtRiviere,
                           None, PeriodeRiviere]
        LineRiviereLQ = ["Gamme de reference - Limites quantitatives", "Reference", "Riviere", "", "", "",
                         StatTxtRiviere,
                         None,
                         PeriodeRiviere]
        LineRiviereSigne = ["Gamme de reference - Nb LQ", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                            PeriodeRiviere]
        LineRiviereN = ["Gamme de reference - N", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                        PeriodeRiviere]
        LineRiviereMoy = ["Gamme de reference - Moyenne", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereEct = ["Gamme de reference - Ecart-Type", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereMed = ["Gamme de reference - Mediane", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRivierePct10 = ["Gamme de reference - Percentile 10", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct25 = ["Gamme de reference - Percentile 25", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct75 = ["Gamme de reference - Percentile 75", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct80 = ["Gamme de reference - Percentile 80", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct85 = ["Gamme de reference - Percentile 85", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct90 = ["Gamme de reference - Percentile 90", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct95 = ["Gamme de reference - Percentile 95", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRiviereMin = ["Gamme de reference - Min", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereMax = ["Gamme de reference - Max", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]

        # ligne pour les statistiques de reference Doline
        if len(DonneesMDol) > 0:
            LineDolMeth = ["Gamme de reference - methodes analytiques", "Reference", "Doline", "", "", "", StatTxtDol,
                           None,
                           PeriodeDol]
            LineDolLQ = ["Gamme de reference - Limites quantitatives", "Reference", "Doline", "", "", "", StatTxtDol,
                         None,
                         PeriodeDol]
            LineDolSigne = ["Gamme de reference - Nb LQ", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolN = ["Gamme de reference - N", "Reference", "Doline", "", "", "", StatTxtDol, None,
                        PeriodeDol]
            LineDolMoy = ["Gamme de reference - Moyenne", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolEct = ["Gamme de reference - Ecart-type", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolMed = ["Gamme de reference - Mediane", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolPct10 = ["Gamme de reference - Percentile 10", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct25 = ["Gamme de reference - Percentile 25", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct75 = ["Gamme de reference - Percentile 75", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct80 = ["Gamme de reference - Percentile 80", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct85 = ["Gamme de reference - Percentile 85", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct90 = ["Gamme de reference - Percentile 90", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct95 = ["Gamme de reference - Percentile 95", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolMin = ["Gamme de reference - Min", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolMax = ["Gamme de reference - Max", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
        # ligne pour les gammes de reference OEIL
        # "SeuilBilanENV_Etude" perc75, "PeriodeBilanENV_Etude" periode, "LQ_BilanENV_Etude" lq, "N_LQ_BilanENV_Etude" n < lq, "N_BilanENV_Etude" n]
        # "parametre", "id_parametre", "SeuilBilanENV_Etude", "PeriodeBilanENV_Etude", "LQ_BilanENV_Etude", "N_LQ_BilanENV_Etude", "N_BilanENV_Etude"
        # Attention la periode est completee selon le parametre via BilanENVXMLParametres
        LineOEILLQ = ["Gamme de reference - Limites quantitatives", "", "Seuil Riviere OEIL", "", "", "", "", None]
        LineOEILSigne = ["Gamme de reference - Nb LQ", "", "Seuil Riviere OEIL", "", "", "", "", None]
        LineOEILN = ["Gamme de reference - N", "", "Seuil Riviere OEIL", "", "", "", "", None]
        LineOEILMed = ["Gamme de reference - Mediane", "", "Seuil Riviere OEIL", "", "", "", "", None]
        LineOEILPct75 = ["Gamme de reference - Percentile 75", "", "Seuil Riviere OEIL", "", "", "", "", None]
        LineOEILMax = ["Gamme de reference - Max", "", "Seuil Riviere OEIL", "", "", "", "", None]

        for p in Parametres:
            i = Parametres.index(p) + indiceCol["PremierParam"]
            # methodes analytiques Riviere
            LineRiviereMeth.append(FindMethodes(DonneesMRiviere, i))
            # methodes analytiques doline
            if len(DonneesMDol) > 0:
                LineDolMeth.append(FindMethodes(DonneesMDol, i))
            ## VALEUR LQ
            # Riviere
            LineRiviereLQ.append(FindLQ(DonneesSRiviere, i, DonneesExRiviere, IdentSuiv))
            # Doline
            if len(DonneesMDol) > 0:
                LineDolLQ.append(FindLQ(DonneesSDol, i, DonneesExDol, IdentSuiv))
            # cas etude OEIL
            PrmEtudeOEIL = [xp for xp in ParamsEtudeOEIL if p[0] == xp[1]]
            if len(PrmEtudeOEIL) > 0:
                PrmEtudeOEIL = PrmEtudeOEIL[0]
                # si premier parametre ajouter d'abord la periode correspondante
                if Parametres.index(p) == 0:
                    periodeOEIL = PrmEtudeOEIL[3]
                    LineOEILLQ.append(periodeOEIL)
                    LineOEILLQ.append(PrmEtudeOEIL[4])
                    LineOEILSigne.append(PrmEtudeOEIL[3])
                    LineOEILSigne.append(PrmEtudeOEIL[5])
                    LineOEILN.append(PrmEtudeOEIL[3])
                    LineOEILN.append(PrmEtudeOEIL[6])
                    LineOEILMed.append(PrmEtudeOEIL[3])
                    LineOEILMed.append(PrmEtudeOEIL[7])
                    LineOEILPct75.append(PrmEtudeOEIL[3])
                    LineOEILPct75.append(PrmEtudeOEIL[2])
                    LineOEILMax.append(PrmEtudeOEIL[3])
                    LineOEILMax.append(PrmEtudeOEIL[8])

                else:
                    periodeOEIL = periodeOEIL + "|" + PrmEtudeOEIL[3]
                    LineOEILLQ[indiceCol["Periode"]] = periodeOEIL
                    LineOEILLQ.append(PrmEtudeOEIL[4])
                    LineOEILSigne[indiceCol["Periode"]] = periodeOEIL
                    LineOEILSigne.append(PrmEtudeOEIL[5])
                    LineOEILN[indiceCol["Periode"]] = periodeOEIL
                    LineOEILN.append(PrmEtudeOEIL[6])
                    LineOEILMed[indiceCol["Periode"]] = periodeOEIL
                    LineOEILMed.append(PrmEtudeOEIL[7])
                    LineOEILPct75[indiceCol["Periode"]] = periodeOEIL
                    LineOEILPct75.append(PrmEtudeOEIL[2])
                    LineOEILMax[indiceCol["Periode"]] = periodeOEIL
                    LineOEILMax.append(PrmEtudeOEIL[8])
            else:
                LineOEILLQ.append("")
                LineOEILSigne.append("")
                LineOEILN.append("")
                LineOEILMed.append("")
                LineOEILPct75.append("")
                LineOEILMax.append("")

            # NOMBRE DE VALEURS = A LA LQ
            # Riviere
            LineRiviereSigne.append(FindSigne(DonneesSRiviere, i))
            # Doline
            if len(DonneesMDol) > 0:
                LineDolSigne.append(FindSigne(DonneesSDol, i))

            # DONNEES NON NULLES OU NON VIDES RECUPEREES + VALEURS MESUREES N
            # Riviere
            VNRiviere = FindValsNonVides(DonneesExRiviere, i)
            valNRiviere = VNRiviere[0]
            ValsExRiviere = VNRiviere[1]
            LineRiviereN.append(valNRiviere)
            # Doline
            if len(DonneesMDol) > 0:
                VNDol = FindValsNonVides(DonneesExDol, i)
                valNDol = VNDol[0]
                ValsExDol = VNDol[1]
                LineDolN.append(valNDol)
            # Riviere
            if valNRiviere == 0:
                LineRiviereMoy.append("")
                LineRiviereEct.append("")
                LineRiviereMed.append("")
                LineRivierePct10.append("")
                LineRivierePct25.append("")
                LineRivierePct75.append("")
                LineRivierePct80.append("")
                LineRivierePct85.append("")
                LineRivierePct90.append("")
                LineRivierePct95.append("")
                LineRiviereMin.append("")
                LineRiviereMax.append("")
            else:
                CalcRiviere = FindStatsCal(ValsExRiviere)
                LineRiviereMoy.append(CalcRiviere["moy"])
                LineRiviereEct.append(CalcRiviere["ect"])
                LineRiviereMed.append(CalcRiviere["med"])
                LineRivierePct10.append(CalcRiviere["perc10"])
                LineRivierePct25.append(CalcRiviere["perc25"])
                LineRivierePct75.append(CalcRiviere["perc75"])
                LineRivierePct80.append(CalcRiviere["perc80"])
                LineRivierePct85.append(CalcRiviere["perc85"])
                LineRivierePct90.append(CalcRiviere["perc90"])
                LineRivierePct95.append(CalcRiviere["perc95"])
                # Min et Max
                LineRiviereMin.append(CalcRiviere["min"])
                LineRiviereMax.append(CalcRiviere["max"])
            # Doline
            if len(DonneesMDol) > 0:
                if valNDol == 0:
                    LineDolMoy.append("")
                    LineDolEct.append("")
                    LineDolMed.append("")
                    LineDolPct10.append("")
                    LineDolPct25.append("")
                    LineDolPct75.append("")
                    LineDolPct80.append("")
                    LineDolPct85.append("")
                    LineDolPct90.append("")
                    LineDolPct95.append("")
                    LineDolMin.append("")
                    LineDolMax.append("")
                else:
                    CalcDol = FindStatsCal(ValsExDol)
                    LineDolMoy.append(CalcDol["moy"])
                    LineDolEct.append(CalcDol["ect"])
                    LineDolMed.append(CalcDol["med"])
                    LineDolPct10.append(CalcDol["perc10"])
                    LineDolPct25.append(CalcDol["perc25"])
                    LineDolPct75.append(CalcDol["perc75"])
                    LineDolPct80.append(CalcDol["perc80"])
                    LineDolPct85.append(CalcDol["perc85"])
                    LineDolPct90.append(CalcDol["perc90"])
                    LineDolPct95.append(CalcDol["perc95"])
                    # Min et Max
                    LineDolMin.append(CalcDol["min"])
                    LineDolMax.append(CalcDol["max"])

        # Insertion Donnees Riviere
        DonneesStatGammeRef.append(LineRiviereMeth)
        DonneesStatGammeRef.append(LineRiviereLQ)
        DonneesStatGammeRef.append(LineRiviereSigne)
        DonneesStatGammeRef.append(LineRiviereN)
        DonneesStatGammeRef.append(LineRiviereMoy)
        DonneesStatGammeRef.append(LineRiviereEct)
        DonneesStatGammeRef.append(LineRiviereMed)
        DonneesStatGammeRef.append(LineRivierePct10)
        DonneesStatGammeRef.append(LineRivierePct25)
        DonneesStatGammeRef.append(LineRivierePct75)
        DonneesStatGammeRef.append(LineRivierePct80)
        DonneesStatGammeRef.append(LineRivierePct85)
        DonneesStatGammeRef.append(LineRivierePct90)
        DonneesStatGammeRef.append(LineRivierePct95)
        DonneesStatGammeRef.append(LineRiviereMin)
        DonneesStatGammeRef.append(LineRiviereMax)

        # Insertion Donnees Doline s'il y en a
        if len(DonneesMDol) > 0:
            DonneesStatGammeRef.append(LineDolMeth)
            DonneesStatGammeRef.append(LineDolLQ)
            DonneesStatGammeRef.append(LineDolSigne)
            DonneesStatGammeRef.append(LineDolN)
            DonneesStatGammeRef.append(LineDolMoy)
            DonneesStatGammeRef.append(LineDolEct)
            DonneesStatGammeRef.append(LineDolMed)
            DonneesStatGammeRef.append(LineDolPct10)
            DonneesStatGammeRef.append(LineDolPct25)
            DonneesStatGammeRef.append(LineDolPct75)
            DonneesStatGammeRef.append(LineDolPct80)
            DonneesStatGammeRef.append(LineDolPct85)
            DonneesStatGammeRef.append(LineDolPct90)
            DonneesStatGammeRef.append(LineDolPct95)
            DonneesStatGammeRef.append(LineDolMin)
            DonneesStatGammeRef.append(LineDolMax)
        DonneesStatGammeRef.append(LineOEILLQ)
        DonneesStatGammeRef.append(LineOEILSigne)
        DonneesStatGammeRef.append(LineOEILN)
        DonneesStatGammeRef.append(LineOEILMed)
        DonneesStatGammeRef.append(LineOEILPct75)
        DonneesStatGammeRef.append(LineOEILMax)


    # CAS MACROINVERTEBRES Riviere ET AVAL et SEDIMENTS EGALEMENT CAR SEULEMENT GRAND BASSIN VERSANT
    elif IdentSuiv == "MACROBEN" or IdentSuiv == "SEDI":
        arcpy.AddMessage("Oulalalaaaa Gammes de reference pour les macroinvertebres")
        DonneesMRiviere = [d for d in DonneesM if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]
        StatTxtRiviere = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesMRiviere])))
        PeriodeRiviere = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesMRiviere])))
        if IdentSuiv == "MACROBEN":
            PeriodeRiviere = "Saison chaude " + PeriodeRiviere
        DonneesSRiviere = [d for d in DonneesS if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]
        DonneesExRiviere = [d for d in DonneesEx if
                            d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Riviere"]
        # Aval Donnees
        DonneesMAval = [d for d in DonneesM if
                        d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Aval"]
        StatTxtAval = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesMAval])))
        PeriodeAval = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesMAval])))
        if IdentSuiv == "MACROBEN":
            PeriodeAval = "Saison chaude " + PeriodeAval
        DonneesSAval = [d for d in DonneesS if
                        d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Aval"]
        DonneesExAval = [d for d in DonneesEx if
                         d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Aval"]
        # Dolines Donnees
        DonneesMDol = [d for d in DonneesM if
                       d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
        if len(DonneesMDol) > 0:
            StatTxtDol = ConcatenationInformation(list(set([d[indiceCol["Station"]] for d in DonneesMDol])))
            PeriodeDol = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesMDol])))
            if IdentSuiv == "MACROBEN":
                PeriodeDol = "Saison chaude " + PeriodeDol
            DonneesSDol = [d for d in DonneesS if
                           d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
            DonneesExDol = [d for d in DonneesEx if
                            d[indiceCol["TypeStation"]] == "Reference" and d[indiceCol["Localisation"]] == "Doline"]
        # ligne pour les statistiques de reference Riviere
        LineRiviereMeth = ["Gamme de reference - methodes analytiques", "Reference", "Riviere", "", "", "",
                           StatTxtRiviere, None, PeriodeRiviere]
        LineRiviereLQ = ["Gamme de reference - Limites quantitatives", "Reference", "Riviere", "", "", "",
                         StatTxtRiviere, None,
                         PeriodeRiviere]
        LineRiviereSigne = ["Gamme de reference - Nb LQ", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                            PeriodeRiviere]
        LineRiviereN = ["Gamme de reference - N", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                        PeriodeRiviere]
        LineRiviereMoy = ["Gamme de reference - Moyenne", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereEct = ["Gamme de reference - Ecart-Type", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereMed = ["Gamme de reference - Mediane", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRivierePct10 = ["Gamme de reference - Percentile 10", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct25 = ["Gamme de reference - Percentile 25", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct75 = ["Gamme de reference - Percentile 75", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct80 = ["Gamme de reference - Percentile 80", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct85 = ["Gamme de reference - Percentile 85", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct90 = ["Gamme de reference - Percentile 90", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRivierePct95 = ["Gamme de reference - Percentile 95", "Reference", "Riviere", "", "", "", StatTxtRiviere,
                            None,
                            PeriodeRiviere]
        LineRiviereMin = ["Gamme de reference - Min", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        LineRiviereMax = ["Gamme de reference - Max", "Reference", "Riviere", "", "", "", StatTxtRiviere, None,
                          PeriodeRiviere]
        # ligne pour les statistiques de reference Aval
        LineAvalMeth = ["Gamme de reference - methodes analytiques", "Reference", "Aval", "", "", "", StatTxtAval, None,
                        PeriodeAval]
        LineAvalLQ = ["Gamme de reference - Limites quantitatives", "Reference", "Aval", "", "", "", StatTxtAval, None,
                      PeriodeAval]
        LineAvalSigne = ["Gamme de reference - Nb LQ", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalN = ["Gamme de reference - N", "Reference", "Aval", "", "", "", StatTxtAval, None,
                     PeriodeAval]
        LineAvalMoy = ["Gamme de reference - Moyenne", "Reference", "Aval", "", "", "", StatTxtAval, None,
                       PeriodeAval]
        LineAvalEct = ["Gamme de reference - Ecart-type", "Reference", "Aval", "", "", "", StatTxtAval, None,
                       PeriodeAval]
        LineAvalMed = ["Gamme de reference - Mediane", "Reference", "Aval", "", "", "", StatTxtAval, None,
                       PeriodeAval]
        LineAvalPct10 = ["Gamme de reference - Percentile 10", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct25 = ["Gamme de reference - Percentile 25", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct75 = ["Gamme de reference - Percentile 75", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct80 = ["Gamme de reference - Percentile 80", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct85 = ["Gamme de reference - Percentile 85", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct90 = ["Gamme de reference - Percentile 90", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalPct95 = ["Gamme de reference - Percentile 95", "Reference", "Aval", "", "", "", StatTxtAval, None,
                         PeriodeAval]
        LineAvalMin = ["Gamme de reference - Min", "Reference", "Aval", "", "", "", StatTxtAval, None,
                       PeriodeAval]
        LineAvalMax = ["Gamme de reference - Max", "Reference", "Aval", "", "", "", StatTxtAval, None,
                       PeriodeAval]
        # ligne pour les statistiques de reference Doline
        if len(DonneesMDol) > 0:
            LineDolMeth = ["Gamme de reference - methodes analytiques", "Reference", "Doline", "", "", "", StatTxtDol,
                           None,
                           PeriodeDol]
            LineDolLQ = ["Gamme de reference - Limites quantitatives", "Reference", "Doline", "", "", "", StatTxtDol,
                         None,
                         PeriodeDol]
            LineDolSigne = ["Gamme de reference - Nb LQ", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolN = ["Gamme de reference - N", "Reference", "Doline", "", "", "", StatTxtDol, None,
                        PeriodeDol]
            LineDolMoy = ["Gamme de reference - Moyenne", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolEct = ["Gamme de reference - Ecart-type", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolMed = ["Gamme de reference - Mediane", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolPct10 = ["Gamme de reference - Percentile 10", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct25 = ["Gamme de reference - Percentile 25", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct75 = ["Gamme de reference - Percentile 75", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct80 = ["Gamme de reference - Percentile 80", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct85 = ["Gamme de reference - Percentile 85", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct90 = ["Gamme de reference - Percentile 90", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolPct95 = ["Gamme de reference - Percentile 95", "Reference", "Doline", "", "", "", StatTxtDol, None,
                            PeriodeDol]
            LineDolMin = ["Gamme de reference - Min", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]
            LineDolMax = ["Gamme de reference - Max", "Reference", "Doline", "", "", "", StatTxtDol, None,
                          PeriodeDol]

        for p in Parametres:
            i = Parametres.index(p) + indiceCol["PremierParam"]
            # methodes analytiques Riviere et aval
            LineRiviereMeth.append(FindMethodes(DonneesMRiviere, i))
            LineAvalMeth.append(FindMethodes(DonneesMAval, i))
            if len(DonneesMDol) > 0:
                LineDolMeth.append(FindMethodes(DonneesMDol, i))
            # valeurs LQ
            # Riviere
            LineRiviereLQ.append(FindLQ(DonneesSRiviere, i, DonneesExRiviere, IdentSuiv))
            # Aval
            LineAvalLQ.append(FindLQ(DonneesSAval, i, DonneesExAval, IdentSuiv))
            # Doline
            if len(DonneesMDol) > 0:
                LineDolLQ.append(FindLQ(DonneesSDol, i, DonneesExDol, IdentSuiv))
            # nombre de valeurs a la LQ
            LineRiviereSigne.append(FindSigne(DonneesSRiviere, i))
            LineAvalSigne.append(FindSigne(DonneesSAval, i))
            if len(DonneesMDol) > 0:
                LineDolSigne.append(FindSigne(DonneesSDol, i))
            # Donnees non nulles ou non vides a recuperer + valeurs reellement mesurees pour N
            # Riviere
            VNRiviere = FindValsNonVides(DonneesExRiviere, i)
            valNRiviere = VNRiviere[0]
            ValsExRiviere = VNRiviere[1]
            LineRiviereN.append(valNRiviere)
            # Aval
            VNAval = FindValsNonVides(DonneesExAval, i)
            valNAval = VNAval[0]
            ValsExAval = VNAval[1]
            LineAvalN.append(valNAval)
            # Doline
            if len(DonneesMDol) > 0:
                VNDol = FindValsNonVides(DonneesExDol, i)
                valNDol = VNDol[0]
                ValsExDol = VNDol[1]
                LineDolN.append(valNDol)
            # Riviere
            if valNRiviere == 0:
                LineRiviereMoy.append("")
                LineRiviereEct.append("")
                LineRiviereMed.append("")
                LineRivierePct10.append("")
                LineRivierePct25.append("")
                LineRivierePct75.append("")
                LineRivierePct80.append("")
                LineRivierePct85.append("")
                LineRivierePct90.append("")
                LineRivierePct95.append("")
                LineRiviereMin.append("")
                LineRiviereMax.append("")
            else:
                CalcRiviere = FindStatsCal(ValsExRiviere)
                LineRiviereMoy.append(CalcRiviere["moy"])
                LineRiviereEct.append(CalcRiviere["ect"])
                LineRiviereMed.append(CalcRiviere["med"])
                LineRivierePct10.append(CalcRiviere["perc10"])
                LineRivierePct25.append(CalcRiviere["perc25"])
                LineRivierePct75.append(CalcRiviere["perc75"])
                LineRivierePct80.append(CalcRiviere["perc80"])
                LineRivierePct85.append(CalcRiviere["perc85"])
                LineRivierePct90.append(CalcRiviere["perc90"])
                LineRivierePct95.append(CalcRiviere["perc95"])
                # Min et Max
                LineRiviereMin.append(CalcRiviere["min"])
                LineRiviereMax.append(CalcRiviere["max"])
            # Aval
            if valNAval == 0:
                LineAvalMoy.append("")
                LineAvalEct.append("")
                LineAvalMed.append("")
                LineAvalPct10.append("")
                LineAvalPct25.append("")
                LineAvalPct75.append("")
                LineAvalPct80.append("")
                LineAvalPct85.append("")
                LineAvalPct90.append("")
                LineAvalPct95.append("")
                LineAvalMin.append("")
                LineAvalMax.append("")
            else:
                CalcAval = FindStatsCal(ValsExAval)
                LineAvalMoy.append(CalcAval["moy"])
                LineAvalEct.append(CalcAval["ect"])
                LineAvalMed.append(CalcAval["med"])
                LineAvalPct10.append(CalcAval["perc10"])
                LineAvalPct25.append(CalcAval["perc25"])
                LineAvalPct75.append(CalcAval["perc75"])
                LineAvalPct80.append(CalcAval["perc80"])
                LineAvalPct85.append(CalcAval["perc85"])
                LineAvalPct90.append(CalcAval["perc90"])
                LineAvalPct95.append(CalcAval["perc95"])
                # Min et Max
                LineAvalMin.append(CalcAval["min"])
                LineAvalMax.append(CalcAval["max"])
            # Doline
            if len(DonneesMDol) > 0:
                if valNDol == 0:
                    LineDolMoy.append("")
                    LineDolEct.append("")
                    LineDolMed.append("")
                    LineDolPct10.append("")
                    LineDolPct25.append("")
                    LineDolPct75.append("")
                    LineDolPct80.append("")
                    LineDolPct85.append("")
                    LineDolPct90.append("")
                    LineDolPct95.append("")
                    LineDolMin.append("")
                    LineDolMax.append("")
                else:
                    CalcDol = FindStatsCal(ValsExDol)
                    LineDolMoy.append(CalcDol["moy"])
                    LineDolEct.append(CalcDol["ect"])
                    LineDolMed.append(CalcDol["med"])
                    LineDolPct10.append(CalcDol["perc10"])
                    LineDolPct25.append(CalcDol["perc25"])
                    LineDolPct75.append(CalcDol["perc75"])
                    LineDolPct80.append(CalcDol["perc80"])
                    LineDolPct85.append(CalcDol["perc85"])
                    LineDolPct90.append(CalcDol["perc90"])
                    LineDolPct95.append(CalcDol["perc95"])
                    # Min et Max
                    LineDolMin.append(CalcDol["min"])
                    LineDolMax.append(CalcDol["max"])
        # Insertion Donnees Riviere
        DonneesStatGammeRef.append(LineRiviereMeth)
        DonneesStatGammeRef.append(LineRiviereLQ)
        DonneesStatGammeRef.append(LineRiviereSigne)
        DonneesStatGammeRef.append(LineRiviereN)
        DonneesStatGammeRef.append(LineRiviereMoy)
        DonneesStatGammeRef.append(LineRiviereEct)
        DonneesStatGammeRef.append(LineRiviereMed)
        DonneesStatGammeRef.append(LineRivierePct10)
        DonneesStatGammeRef.append(LineRivierePct25)
        DonneesStatGammeRef.append(LineRivierePct75)
        DonneesStatGammeRef.append(LineRivierePct80)
        DonneesStatGammeRef.append(LineRivierePct85)
        DonneesStatGammeRef.append(LineRivierePct90)
        DonneesStatGammeRef.append(LineRivierePct95)
        DonneesStatGammeRef.append(LineRiviereMin)
        DonneesStatGammeRef.append(LineRiviereMax)
        # insertion Donnees Aval
        DonneesStatGammeRef.append(LineAvalMeth)
        DonneesStatGammeRef.append(LineAvalLQ)
        DonneesStatGammeRef.append(LineAvalSigne)
        DonneesStatGammeRef.append(LineAvalN)
        DonneesStatGammeRef.append(LineAvalMoy)
        DonneesStatGammeRef.append(LineAvalEct)
        DonneesStatGammeRef.append(LineAvalMed)
        DonneesStatGammeRef.append(LineAvalPct10)
        DonneesStatGammeRef.append(LineAvalPct25)
        DonneesStatGammeRef.append(LineAvalPct75)
        DonneesStatGammeRef.append(LineAvalPct80)
        DonneesStatGammeRef.append(LineAvalPct85)
        DonneesStatGammeRef.append(LineAvalPct90)
        DonneesStatGammeRef.append(LineAvalPct95)
        DonneesStatGammeRef.append(LineAvalMin)
        DonneesStatGammeRef.append(LineAvalMax)
        # Insertion Donnees Doline s'il y en a
        if len(DonneesMDol) > 0:
            DonneesStatGammeRef.append(LineDolMeth)
            DonneesStatGammeRef.append(LineDolLQ)
            DonneesStatGammeRef.append(LineDolSigne)
            DonneesStatGammeRef.append(LineDolN)
            DonneesStatGammeRef.append(LineDolMoy)
            DonneesStatGammeRef.append(LineDolEct)
            DonneesStatGammeRef.append(LineDolMed)
            DonneesStatGammeRef.append(LineDolPct10)
            DonneesStatGammeRef.append(LineDolPct25)
            DonneesStatGammeRef.append(LineDolPct75)
            DonneesStatGammeRef.append(LineDolPct80)
            DonneesStatGammeRef.append(LineDolPct85)
            DonneesStatGammeRef.append(LineDolPct90)
            DonneesStatGammeRef.append(LineDolPct95)
            DonneesStatGammeRef.append(LineDolMin)
            DonneesStatGammeRef.append(LineDolMax)
    # CAS EAUX SOUTERRAINES
    elif IdentSuiv == "SOUT":
        arcpy.AddMessage(
            "Encore du taff pour le gros matou, il traite les metriques pour les gammes de reference des eaux souterraines")
        DonneesMOrig = list()
        DonneesMOrig.extend(DonneesM)
        DonneesSOrig = list()
        DonneesSOrig.extend(DonneesS)
        DonneesExOrig = list()
        DonneesExOrig.extend(DonneesEx)
        for Rf in ["Reference", "Controle"]:
            DonneesM = [d for d in DonneesMOrig if d[indiceCol["TypeStation"]] == Rf]
            DonneesS = [d for d in DonneesSOrig if d[indiceCol["TypeStation"]] == Rf]
            DonneesEx = [d for d in DonneesExOrig if d[indiceCol["TypeStation"]] == Rf]
            Localisations = [d[indiceCol["Localisation"]] for d in DonneesM]
            Localisations = list(set(Localisations))
            AnCours = max([int(d[indiceCol["Date"]].strftime("%Y")) for d in DonneesM])
            Ans3 = [AnCours - 3, AnCours - 2, AnCours - 1]
            if "principal lateritique" not in Localisations:
                Localisations.append("principal lateritique")
            arcpy.AddMessage(Localisations)
            for l in Localisations:
                # analyse par localisation type aquifere_principal aquifere_p
                if l == "principal lateritique":
                    DonneesML = [d for d in DonneesM if d[indiceCol["Localisation"]] == l or d[
                        indiceCol["Localisation"]] == "lateritique" or d[
                                     indiceCol["Localisation"]] == "principal"]
                    DonneesSL = [d for d in DonneesS if d[indiceCol["Localisation"]] == l or d[
                        indiceCol["Localisation"]] == "lateritique" or d[
                                     indiceCol["Localisation"]] == "principal"]
                    DonneesExL = [d for d in DonneesEx if d[indiceCol["Localisation"]] == l or d[
                        indiceCol["Localisation"]] == "lateritique" or d[
                                      indiceCol["Localisation"]] == "principal"]
                else:
                    DonneesML = [d for d in DonneesM if d[indiceCol["Localisation"]] == l]
                    DonneesSL = [d for d in DonneesS if d[indiceCol["Localisation"]] == l]
                    DonneesExL = [d for d in DonneesEx if d[indiceCol["Localisation"]] == l]
                StatsL = [d[indiceCol["Station"]] for d in DonneesML]
                StatsL = list(set(StatsL))
                StatTxtRef = ConcatenationInformation(StatsL)
                PeriodeRef = AnneeMinMax(list(set([d[indiceCol["Date"]] for d in DonneesML])))

                LineRefMeth = ["Gamme de reference - methodes analytiques", Rf, l, "", "", "", StatTxtRef, None,
                               PeriodeRef]
                LineRefLQ = ["Gamme de reference - Limites quantitatives", Rf, l, "", "", "", StatTxtRef, None,
                             PeriodeRef]
                LineRefSigne = ["Gamme de reference - Nb LQ", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefN = ["Gamme de reference - N", Rf, l, "", "", "", StatTxtRef, None,
                            PeriodeRef]
                LineRefMoy = ["Gamme de reference - Moyenne", Rf, l, "", "", "", StatTxtRef, None,
                              PeriodeRef]
                LineRefEct = ["Gamme de reference - Ecart-type", Rf, l, "", "", "", StatTxtRef, None,
                              PeriodeRef]
                LineRefMed = ["Gamme de reference - Mediane", Rf, l, "", "", "", StatTxtRef, None,
                              PeriodeRef]
                LineRefPct10 = ["Gamme de reference - Percentile 10", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct25 = ["Gamme de reference - Percentile 25", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct75 = ["Gamme de reference - Percentile 75", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct80 = ["Gamme de reference - Percentile 80", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct85 = ["Gamme de reference - Percentile 85", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct90 = ["Gamme de reference - Percentile 90", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefPct95 = ["Gamme de reference - Percentile 95", Rf, l, "", "", "", StatTxtRef, None,
                                PeriodeRef]
                LineRefMin = ["Gamme de reference - Min", Rf, l, "", "", "", StatTxtRef, None,
                              PeriodeRef]
                LineRefMax = ["Gamme de reference - Max", Rf, l, "", "", "", StatTxtRef, None,
                              PeriodeRef]

                for p in Parametres:
                    i = Parametres.index(p) + indiceCol["PremierParam"]
                    # methodes analytiques de reference par zone
                    LineRefMeth.append(FindMethodes(DonneesML, i))
                    # valeurs LQ
                    LineRefLQ.append(FindLQ(DonneesSL, i, DonneesExL, IdentSuiv))
                    # nombre de valeurs a la LQ
                    LineRefSigne.append(FindSigne(DonneesSL, i))
                    # Donnees non nulles ou non vides a recuperer + valeurs reellement mesurees pour N
                    VNRef = FindValsNonVides(DonneesExL, i)
                    valNRef = VNRef[0]
                    ValsExRef = VNRef[1]
                    LineRefN.append(valNRef)
                    # Valeurs reellement recuperees pour le calcul des differentes statistiques
                    if valNRef == 0:
                        LineRefMoy.append("")
                        LineRefEct.append("")
                        LineRefMed.append("")
                        LineRefPct10.append("")
                        LineRefPct25.append("")
                        LineRefPct75.append("")
                        LineRefPct80.append("")
                        LineRefPct85.append("")
                        LineRefPct90.append("")
                        LineRefPct95.append("")
                        LineRefMin.append("")
                        LineRefMax.append("")
                    else:
                        CalcRef = FindStatsCal(ValsExRef)
                        LineRefMoy.append(CalcRef["moy"])
                        LineRefEct.append(CalcRef["ect"])
                        LineRefMed.append(CalcRef["med"])
                        LineRefPct10.append(CalcRef["perc10"])
                        LineRefPct25.append(CalcRef["perc25"])
                        LineRefPct75.append(CalcRef["perc75"])
                        LineRefPct80.append(CalcRef["perc80"])
                        LineRefPct85.append(CalcRef["perc85"])
                        LineRefPct90.append(CalcRef["perc90"])
                        LineRefPct95.append(CalcRef["perc95"])
                        # Min et Max
                        LineRefMin.append(CalcRef["min"])
                        LineRefMax.append(CalcRef["max"])

                DonneesStatGammeRef.append(LineRefMeth)
                DonneesStatGammeRef.append(LineRefLQ)
                DonneesStatGammeRef.append(LineRefSigne)
                DonneesStatGammeRef.append(LineRefN)
                DonneesStatGammeRef.append(LineRefMoy)
                DonneesStatGammeRef.append(LineRefEct)
                DonneesStatGammeRef.append(LineRefMed)
                DonneesStatGammeRef.append(LineRefPct10)
                DonneesStatGammeRef.append(LineRefPct25)
                DonneesStatGammeRef.append(LineRefPct75)
                DonneesStatGammeRef.append(LineRefPct80)
                DonneesStatGammeRef.append(LineRefPct85)
                DonneesStatGammeRef.append(LineRefPct90)
                DonneesStatGammeRef.append(LineRefPct95)
                DonneesStatGammeRef.append(LineRefMin)
                DonneesStatGammeRef.append(LineRefMax)

    else:
        arcpy.AddMessage("Ouuueeeee  Pas de gamme de reference demandee pour le gros matou :-)")
        LineVideGamme = ["Pas de gamme de reference attendue", "", "", "", "", "", "", None, ""]
        for p in Params:
            LineVideGamme.append("")
        DonneesStatGammeRef.append(LineVideGamme)

    return DonneesStatGammeRef


# Fonction des Statistiques par Periode (Annee ou Saison selon le suivi)
def StatistiquesPeriode(DonneesMethodes, DonneesSignes, DonneesExp, Parametres, StatistGP, IdentSuiv):
    # on conserve toutes les infos de station pour conserver ensuite l'ordre
    StatsDts = sorted(list(set([d[indiceCol["TypeStation"]] + ";" + d[indiceCol["Localisation"]] + ";" + d[
        indiceCol["TypeBV"]] + ";" + d[indiceCol["Zone"]] + ";" + d[indiceCol["ZoneReference"]] + ";" + d[
                                    indiceCol["Station"]] for d in DonneesMethodes])))
    StatsDts = [s.split(";") for s in StatsDts]
    Stats = [s[indiceCol["Station"] - 1] for s in StatsDts]
    SortieDonnees = list()
    for s in Stats:
        StZoneRef = [std[indiceCol["ZoneReference"] - 1] for std in StatsDts if s == std[indiceCol["Station"] - 1]][0]
        StLoc = [std[indiceCol["Localisation"] - 1] for std in StatsDts if s == std[indiceCol["Station"] - 1]][0]
        StBV = [std[indiceCol["TypeBV"] - 1] for std in StatsDts if s == std[indiceCol["Station"] - 1]][0]
        nbS = len(Stats)
        sti = Stats.index(s)
        pct25 = round(nbS / 4)
        pct50 = round(nbS / 2)
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
            LineMeth = ["Metrique - methodes analytiques", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                        d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                        d[indiceCol["Station"]], None, per]
            LineSign = ["Metrique - Nb LQ", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                        d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                        d[indiceCol["Station"]], None, per]
            LineLQ = ["Metrique - Limites quantitatives", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                      d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                      d[indiceCol["Station"]], None, per]
            LineN = ["Metrique - N", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]],
                     d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineMoy = ["Metrique - Moyenne", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                       d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                       d[indiceCol["Station"]], None, per]
            LineEct = ["Metrique - Ecart-type", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                       d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                       d[indiceCol["Station"]], None, per]
            LineMed = ["Metrique - Mediane", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                       d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                       d[indiceCol["Station"]], None, per]
            LinePerc10 = ["Metrique - Percentile 10", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc25 = ["Metrique - Percentile 25", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc75 = ["Metrique - Percentile 75", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc80 = ["Metrique - Percentile 80", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc85 = ["Metrique - Percentile 85", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc90 = ["Metrique - Percentile 90", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LinePerc95 = ["Metrique - Percentile 95", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                          d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                          d[indiceCol["Station"]], None, per]
            LineMin = ["Metrique - Min", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                       d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                       d[indiceCol["Station"]], None, per]
            LineMax = ["Metrique - Max", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                       d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                       d[indiceCol["Station"]], None, per]
            LineSupPerc10 = ["Metrique - Superieur Perc. 10", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc10 = ["Metrique - Superieur Perc. 10 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc25 = ["Metrique - Superieur Perc. 25", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc25 = ["Metrique - Superieur Perc. 25 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc75 = ["Metrique - Superieur Perc. 75", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc75 = ["Metrique - Superieur Perc. 75 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc80 = ["Metrique - Superieur Perc. 80", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc80 = ["Metrique - Superieur Perc. 80 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc85 = ["Metrique - Superieur Perc. 85", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc85 = ["Metrique - Superieur Perc. 85 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc90 = ["Metrique - Superieur Perc. 90", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc90 = ["Metrique - Superieur Perc. 90 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            LineSupPerc95 = ["Metrique - Superieur Perc. 95", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                             d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                             d[indiceCol["Station"]], None, per]
            LinePctSupPerc95 = ["Metrique - Superieur Perc. 95 (%)", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            # Donnees Controle pour les eaux souterraines en plus de Reference
            if IdentSuiv == "SOUT":
                LineContSupPerc10 = ["Metrique - Controle - Superieur Perc. 10", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc10 = ["Metrique - Controle - Superieur Perc. 10 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc25 = ["Metrique - Controle - Superieur Perc. 25", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc25 = ["Metrique - Controle - Superieur Perc. 25 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc75 = ["Metrique - Controle - Superieur Perc. 75", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc75 = ["Metrique - Controle - Superieur Perc. 75 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc80 = ["Metrique - Controle - Superieur Perc. 80", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc80 = ["Metrique - Controle - Superieur Perc. 80 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc85 = ["Metrique - Controle - Superieur Perc. 85", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc85 = ["Metrique - Controle - Superieur Perc. 85 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc90 = ["Metrique - Controle - Superieur Perc. 90", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc90 = ["Metrique - Controle - Superieur Perc. 90 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContSupPerc95 = ["Metrique - Controle - Superieur Perc. 95", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LineContPctSupPerc95 = ["Metrique - Controle - Superieur Perc. 95 (%)", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                # 3 dernieres annees
                if Periodes.index(per) == len(Periodes) - 1:
                    YearNow = int(datetime.datetime.now().strftime("%Y"))
                    dernAn3 = [str(YearNow - 3), str(YearNow - 2), str(YearNow - 1)]
                    DonneesMthStPer3 = [d for d in DonneesMethStat if
                                        d[indiceCol["Periode"]] == dernAn3[0] or d[indiceCol["Periode"]] == dernAn3[
                                            1] or d[indiceCol["Periode"]] == dernAn3[2]]
                    DonneesSiStPer3 = [d for d in DonneesSignStat if
                                       d[indiceCol["Periode"]] == dernAn3[0] or d[indiceCol["Periode"]] == dernAn3[1] or
                                       d[indiceCol["Periode"]] == dernAn3[2]]
                    DonneesExpStPer3 = [d for d in DonneesExpStat if
                                        d[indiceCol["Periode"]] == dernAn3[0] or d[indiceCol["Periode"]] == dernAn3[
                                            1] or d[indiceCol["Periode"]] == dernAn3[2]]
                    Perd3 = dernAn3[0] + "-" + dernAn3[2]
                    Line3aSigne = ["Metrique - Nb LQ - 3 ans", d[indiceCol["TypeStation"]],
                                   d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                   d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aN = ["Metrique - N - 3 ans", d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                               d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                               d[indiceCol["Station"]], None, Perd3]
                    Line3aLQ = ["Metrique - Limites quantitatives - 3 ans", d[indiceCol["TypeStation"]],
                                d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc10 = ["Metrique - Superieur Perc. 10 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc10 = ["Metrique - Superieur Perc. 10 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc25 = ["Metrique - Superieur Perc. 25 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc25 = ["Metrique - Superieur Perc. 25 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc75 = ["Metrique - Superieur Perc. 75 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc75 = ["Metrique - Superieur Perc. 75 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc80 = ["Metrique - Superieur Perc. 80 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc80 = ["Metrique - Superieur Perc. 80 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc85 = ["Metrique - Superieur Perc. 85 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc85 = ["Metrique - Superieur Perc. 85 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc90 = ["Metrique - Superieur Perc. 90 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc90 = ["Metrique - Superieur Perc. 90 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aSupPerc95 = ["Metrique - Superieur Perc. 95 - 3 ans", d[indiceCol["TypeStation"]],
                                       d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                       d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    Line3aPctSupPerc95 = ["Metrique - Superieur Perc. 95 (%) - 3 ans", d[indiceCol["TypeStation"]],
                                          d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                          d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    # Controle
                    LineCont3aSupPerc10 = ["Metrique - Controle - Superieur Perc. 10 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc10 = ["Metrique - Controle - Superieur Perc. 10 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc25 = ["Metrique - Controle - Superieur Perc. 25 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc25 = ["Metrique - Controle - Superieur Perc. 25 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc75 = ["Metrique - Controle - Superieur Perc. 75 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc75 = ["Metrique - Controle - Superieur Perc. 75 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc80 = ["Metrique - Controle - Superieur Perc. 80 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc80 = ["Metrique - Controle - Superieur Perc. 80 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc85 = ["Metrique - Controle - Superieur Perc. 85 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc85 = ["Metrique - Controle - Superieur Perc. 85 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc90 = ["Metrique - Controle - Superieur Perc. 90 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc90 = ["Metrique - Controle - Superieur Perc. 90 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
                    LineCont3aSupPerc95 = ["Metrique - Controle - Superieur Perc. 95 - 3 ans",
                                           d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                           d[indiceCol["TypeBV"]], d[indiceCol["Zone"]], d[indiceCol["ZoneReference"]],
                                           d[indiceCol["Station"]], None, Perd3]
                    LineCont3aPctSupPerc95 = ["Metrique - Controle - Superieur Perc. 95 (%) - 3 ans",
                                              d[indiceCol["TypeStation"]], d[indiceCol["Localisation"]],
                                              d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                              d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, Perd3]
            elif IdentSuiv == "SURF":
                # Attention complement periode a partir de parametre
                LineSupPerc75OEIL = ["Metrique - Superieur Perc. 75 - OEIL", d[indiceCol["TypeStation"]],
                                     d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                     d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
                LinePctSupPerc75OEIL = ["Metrique - Superieur Perc. 75 (%) - OEIL", d[indiceCol["TypeStation"]],
                                        d[indiceCol["Localisation"]], d[indiceCol["TypeBV"]], d[indiceCol["Zone"]],
                                        d[indiceCol["ZoneReference"]], d[indiceCol["Station"]], None, per]
            for p in Parametres:
                i = Parametres.index(p) + indiceCol["PremierParam"]
                # methodes analytiques
                LineMeth.append(FindMethodes(DonneesMthStPer, i))
                # nombre de valeurs a la LQ
                LineSign.append(FindSigne(DonneesSiStPer, i))
                # valeurs LQ
                LineLQ.append(FindLQ(DonneesSiStPer, i, DonneesExpStPer, IdentSuiv))
                # recuperation des valeurs non vides ou nulles pour les prochaines statistiques
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
                    if IdentSuiv == "SOUT":
                        LineContSupPerc10.append("")
                        LineContPctSupPerc10.append("")
                        LineContSupPerc25.append("")
                        LineContPctSupPerc25.append("")
                        LineContSupPerc75.append("")
                        LineContPctSupPerc75.append("")
                        LineContSupPerc80.append("")
                        LineContPctSupPerc80.append("")
                        LineContSupPerc85.append("")
                        LineContPctSupPerc85.append("")
                        LineContSupPerc90.append("")
                        LineContPctSupPerc90.append("")
                        LineContSupPerc95.append("")
                        LineContPctSupPerc95.append("")
                        if Periodes.index(per) == len(Periodes) - 1:
                            Line3aSigne.append("")
                            Line3aLQ.append("")
                            Line3aN.append("")
                            Line3aSupPerc10.append("")
                            Line3aPctSupPerc10.append("")
                            Line3aSupPerc25.append("")
                            Line3aPctSupPerc25.append("")
                            Line3aSupPerc75.append("")
                            Line3aPctSupPerc75.append("")
                            Line3aSupPerc80.append("")
                            Line3aPctSupPerc80.append("")
                            Line3aSupPerc85.append("")
                            Line3aPctSupPerc85.append("")
                            Line3aSupPerc90.append("")
                            Line3aPctSupPerc90.append("")
                            Line3aSupPerc95.append("")
                            Line3aPctSupPerc95.append("")
                            # Controle
                            LineCont3aSupPerc10.append("")
                            LineCont3aPctSupPerc10.append("")
                            LineCont3aSupPerc25.append("")
                            LineCont3aPctSupPerc25.append("")
                            LineCont3aSupPerc75.append("")
                            LineCont3aPctSupPerc75.append("")
                            LineCont3aSupPerc80.append("")
                            LineCont3aPctSupPerc80.append("")
                            LineCont3aSupPerc85.append("")
                            LineCont3aPctSupPerc85.append("")
                            LineCont3aSupPerc90.append("")
                            LineCont3aPctSupPerc90.append("")
                            LineCont3aSupPerc95.append("")
                            LineCont3aPctSupPerc95.append("")
                    elif IdentSuiv == "SURF":
                        LineSupPerc75OEIL.append("")
                        LinePctSupPerc75OEIL.append("")
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
                    # valeur superieure au percentile75
                    # Eaux de surface ou eaux souterraines
                    if IdentSuiv == "SURF":
                        Perc10 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 10",
                                           i, ValsExpStPer, valN)
                        Perc25 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 25",
                                           i, ValsExpStPer, valN)
                        Perc75 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 75",
                                           i, ValsExpStPer, valN)
                        Perc80 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 80",
                                           i, ValsExpStPer, valN)
                        Perc85 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 85",
                                           i, ValsExpStPer, valN)
                        Perc90 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 90",
                                           i, ValsExpStPer, valN)
                        Perc95 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 95",
                                           i, ValsExpStPer, valN)
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
                        # comparaison avec seuil reference etude OEIL
                        Perc75OEIL = NbSupPerc(StatistGP, ["Localisation"], ["Seuil Riviere OEIL"],
                                               "Gamme de reference - Percentile 75", i, ValsExpStPer, valN)
                        LineSupPerc75OEIL.append(Perc75OEIL["NbS"])
                        LinePctSupPerc75OEIL.append(Perc75OEIL["PctS"])
                    # MacroInvertebres et Sediments
                    elif IdentSuiv == "MACROBEN" or IdentSuiv == "SEDI":
                        Perc10 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 10",
                                           i, ValsExpStPer, valN)
                        Perc25 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 25",
                                           i, ValsExpStPer, valN)
                        Perc75 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 75",
                                           i, ValsExpStPer, valN)
                        Perc80 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 80",
                                           i, ValsExpStPer, valN)
                        Perc85 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 85",
                                           i, ValsExpStPer, valN)
                        Perc90 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 90",
                                           i, ValsExpStPer, valN)
                        Perc95 = NbSupPerc(StatistGP, ["Localisation"], [StLoc], "Gamme de reference - Percentile 95",
                                           i, ValsExpStPer, valN)
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
                    # Souterrain
                    elif IdentSuiv == "SOUT":
                        Perc10 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 10", i, ValsExpStPer, valN)
                        Perc25 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 25", i, ValsExpStPer, valN)
                        Perc75 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 75", i, ValsExpStPer, valN)
                        Perc80 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 80", i, ValsExpStPer, valN)
                        Perc85 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 85", i, ValsExpStPer, valN)
                        Perc90 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 90", i, ValsExpStPer, valN)
                        Perc95 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                           "Gamme de reference - Percentile 95", i, ValsExpStPer, valN)
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
                        # Controle
                        Perc10Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 10", i, ValsExpStPer, valN)
                        Perc25Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 25", i, ValsExpStPer, valN)
                        Perc75Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 75", i, ValsExpStPer, valN)
                        Perc80Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 80", i, ValsExpStPer, valN)
                        Perc85Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 85", i, ValsExpStPer, valN)
                        Perc90Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 90", i, ValsExpStPer, valN)
                        Perc95Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                               "Gamme de reference - Percentile 95", i, ValsExpStPer, valN)
                        LineContSupPerc10.append(Perc10Cont["NbS"])
                        LineContPctSupPerc10.append(Perc10Cont["PctS"])
                        LineContSupPerc25.append(Perc25Cont["NbS"])
                        LineContPctSupPerc25.append(Perc25Cont["PctS"])
                        LineContSupPerc75.append(Perc75Cont["NbS"])
                        LineContPctSupPerc75.append(Perc75Cont["PctS"])
                        LineContSupPerc80.append(Perc80Cont["NbS"])
                        LineContPctSupPerc80.append(Perc80Cont["PctS"])
                        LineContSupPerc85.append(Perc85Cont["NbS"])
                        LineContPctSupPerc85.append(Perc85Cont["PctS"])
                        LineContSupPerc90.append(Perc90Cont["NbS"])
                        LineContPctSupPerc90.append(Perc90Cont["PctS"])
                        LineContSupPerc95.append(Perc95Cont["NbS"])
                        LineContPctSupPerc95.append(Perc95Cont["PctS"])
                        if Periodes.index(per) == len(Periodes) - 1:
                            Line3aSigne.append(FindSigne(DonneesSiStPer3, i))
                            Line3aLQ.append(FindLQ(DonneesSiStPer3, i, DonneesExpStPer3, IdentSuiv))
                            # recuperation des valeurs non vides ou nulles pour les prochaines statistiques
                            VN3 = FindValsNonVides(DonneesExpStPer3, i)
                            valN3 = VN3[0]
                            ValsExpStPer3 = VN3[1]
                            Line3aN.append(valN3)
                            if valN3 == 0:
                                Line3aSupPerc10.append("")
                                Line3aPctSupPerc10.append("")
                                Line3aSupPerc25.append("")
                                Line3aPctSupPerc25.append("")
                                Line3aSupPerc75.append("")
                                Line3aPctSupPerc75.append("")
                                Line3aSupPerc80.append("")
                                Line3aPctSupPerc80.append("")
                                Line3aSupPerc85.append("")
                                Line3aPctSupPerc85.append("")
                                Line3aSupPerc90.append("")
                                Line3aPctSupPerc90.append("")
                                Line3aSupPerc95.append("")
                                Line3aPctSupPerc95.append("")
                                # Controle
                                LineCont3aSupPerc10.append("")
                                LineCont3aPctSupPerc10.append("")
                                LineCont3aSupPerc25.append("")
                                LineCont3aPctSupPerc25.append("")
                                LineCont3aSupPerc75.append("")
                                LineCont3aPctSupPerc75.append("")
                                LineCont3aSupPerc80.append("")
                                LineCont3aPctSupPerc80.append("")
                                LineCont3aSupPerc85.append("")
                                LineCont3aPctSupPerc85.append("")
                                LineCont3aSupPerc90.append("")
                                LineCont3aPctSupPerc90.append("")
                                LineCont3aSupPerc95.append("")
                                LineCont3aPctSupPerc95.append("")
                            else:
                                Perc10 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 10", i, ValsExpStPer3, valN3)
                                Perc25 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 25", i, ValsExpStPer3, valN3)
                                Perc75 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 75", i, ValsExpStPer3, valN3)
                                Perc80 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 80", i, ValsExpStPer3, valN3)
                                Perc85 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 85", i, ValsExpStPer3, valN3)
                                Perc90 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 90", i, ValsExpStPer3, valN3)
                                Perc95 = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Reference"],
                                                   "Gamme de reference - Percentile 95", i, ValsExpStPer3, valN3)
                                Line3aSupPerc10.append(Perc10["NbS"])
                                Line3aPctSupPerc10.append(Perc10["PctS"])
                                Line3aSupPerc25.append(Perc25["NbS"])
                                Line3aPctSupPerc25.append(Perc25["PctS"])
                                Line3aSupPerc75.append(Perc75["NbS"])
                                Line3aPctSupPerc75.append(Perc75["PctS"])
                                Line3aSupPerc80.append(Perc80["NbS"])
                                Line3aPctSupPerc80.append(Perc80["PctS"])
                                Line3aSupPerc85.append(Perc85["NbS"])
                                Line3aPctSupPerc85.append(Perc85["PctS"])
                                Line3aSupPerc90.append(Perc90["NbS"])
                                Line3aPctSupPerc90.append(Perc90["PctS"])
                                Line3aSupPerc95.append(Perc95["NbS"])
                                Line3aPctSupPerc95.append(Perc95["PctS"])
                                # Controle
                                Perc10Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 10", i, ValsExpStPer3, valN3)
                                Perc25Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 25", i, ValsExpStPer3, valN3)
                                Perc75Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 75", i, ValsExpStPer3, valN3)
                                Perc80Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 80", i, ValsExpStPer3, valN3)
                                Perc85Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 85", i, ValsExpStPer3, valN3)
                                Perc90Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 90", i, ValsExpStPer3, valN3)
                                Perc95Cont = NbSupPerc(StatistGP, ["Localisation", "TypeStation"], [StLoc, "Controle"],
                                                       "Gamme de reference - Percentile 95", i, ValsExpStPer3, valN3)
                                LineCont3aSupPerc10.append(Perc10Cont["NbS"])
                                LineCont3aPctSupPerc10.append(Perc10Cont["PctS"])
                                LineCont3aSupPerc25.append(Perc25Cont["NbS"])
                                LineCont3aPctSupPerc25.append(Perc25Cont["PctS"])
                                LineCont3aSupPerc75.append(Perc75Cont["NbS"])
                                LineCont3aPctSupPerc75.append(Perc75Cont["PctS"])
                                LineCont3aSupPerc80.append(Perc80Cont["NbS"])
                                LineCont3aPctSupPerc80.append(Perc80Cont["PctS"])
                                LineCont3aSupPerc85.append(Perc85Cont["NbS"])
                                LineCont3aPctSupPerc85.append(Perc85Cont["PctS"])
                                LineCont3aSupPerc90.append(Perc90Cont["NbS"])
                                LineCont3aPctSupPerc90.append(Perc90Cont["PctS"])
                                LineCont3aSupPerc95.append(Perc95Cont["NbS"])
                                LineCont3aPctSupPerc95.append(Perc95Cont["PctS"])
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
            if IdentSuiv == "SOUT":
                # Controle
                SortieDonnees.append(LineContSupPerc10)
                SortieDonnees.append(LineContSupPerc25)
                SortieDonnees.append(LineContPctSupPerc25)
                SortieDonnees.append(LineContSupPerc75)
                SortieDonnees.append(LineContPctSupPerc75)
                SortieDonnees.append(LineContSupPerc80)
                SortieDonnees.append(LineContPctSupPerc80)
                SortieDonnees.append(LineContSupPerc85)
                SortieDonnees.append(LineContPctSupPerc85)
                SortieDonnees.append(LineContSupPerc90)
                SortieDonnees.append(LineContPctSupPerc90)
                SortieDonnees.append(LineContSupPerc95)
                SortieDonnees.append(LineContPctSupPerc95)
                if Periodes.index(per) == len(Periodes) - 1:
                    SortieDonnees.append(Line3aSigne)
                    SortieDonnees.append(Line3aLQ)
                    SortieDonnees.append(Line3aN)
                    SortieDonnees.append(Line3aSupPerc10)
                    SortieDonnees.append(Line3aPctSupPerc10)
                    SortieDonnees.append(Line3aSupPerc25)
                    SortieDonnees.append(Line3aPctSupPerc25)
                    SortieDonnees.append(Line3aSupPerc75)
                    SortieDonnees.append(Line3aPctSupPerc75)
                    SortieDonnees.append(Line3aSupPerc80)
                    SortieDonnees.append(Line3aPctSupPerc80)
                    SortieDonnees.append(Line3aSupPerc85)
                    SortieDonnees.append(Line3aPctSupPerc85)
                    SortieDonnees.append(Line3aSupPerc90)
                    SortieDonnees.append(Line3aPctSupPerc90)
                    SortieDonnees.append(Line3aSupPerc95)
                    SortieDonnees.append(Line3aPctSupPerc95)
                    # Controle
                    SortieDonnees.append(LineCont3aSupPerc10)
                    SortieDonnees.append(LineCont3aPctSupPerc10)
                    SortieDonnees.append(LineCont3aSupPerc25)
                    SortieDonnees.append(LineCont3aPctSupPerc25)
                    SortieDonnees.append(LineCont3aSupPerc75)
                    SortieDonnees.append(LineCont3aPctSupPerc75)
                    SortieDonnees.append(LineCont3aSupPerc80)
                    SortieDonnees.append(LineCont3aPctSupPerc80)
                    SortieDonnees.append(LineCont3aSupPerc85)
                    SortieDonnees.append(LineCont3aPctSupPerc85)
                    SortieDonnees.append(LineCont3aSupPerc90)
                    SortieDonnees.append(LineCont3aPctSupPerc90)
                    SortieDonnees.append(LineCont3aSupPerc95)
                    SortieDonnees.append(LineCont3aPctSupPerc95)
            elif IdentSuiv == "SURF":
                SortieDonnees.append(LineSupPerc75OEIL)
                SortieDonnees.append(LinePctSupPerc75OEIL)
            del LineMeth, LineSign, LineLQ, LineN, LineEct, LineMed, LinePerc10, LinePerc25, LinePerc90, LinePerc75, LinePerc80, LinePerc85, LinePerc95
            del DonneesSiStPer, DonneesMthStPer, DonneesExpStPer
        del DonneesMethStat, DonneesSignStat, Periodes, DonneesExpStat
    del DonneesMethodes, DonneesSignes, Parametres, StatsDts, Stats
    return SortieDonnees
