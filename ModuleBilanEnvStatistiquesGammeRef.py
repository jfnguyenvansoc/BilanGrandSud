# -*- coding: utf8 -*-
from ModuleBilanEnv import ConcatenationInformation, AnneeMinMax, FindSigne, FindLQ, FindMethodes, FindStatsCal, FindValsNonVides

import arcpy

# organisation des premieres colonnes Milieu Eaux douces
indiceCol = {"TypeData": 0, "TypeStation": 1, "Localisation": 2, "TypeBV": 3, "Zone": 4, "ZoneReference": 5,
             "Station": 6, "Date": 7, "Periode": 8, "PremierParam": 9}
#organisation des premieres colonnes Milieu Marin
indiceColMar = {"TypeData":0, "TypeStation":1, "Zone":2, "TypologieStation":3, "TypeRef":4, "Station":5, "Date":6, "Periode":7,"PremierParam":8}



##Calcul des Statistiques de Gamme de Reference
def StatistiquesGammeReference(Milieu, DonneesM, DonneesS, DonneesEx, Parametres, IdentSuiv, AnneeEtude, XMLParam=0, DonneesPerRefEau=list()):
    DonneesStatGammeRef = list()
    if Milieu == "Eaux douces":
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
    elif Milieu == "Marin":
        #CAS EAU (Colonne d'eau marine)
        if IdentSuiv == "EAU":
            arcpy.AddMessage("Oh non Gammes de reference pour la colonne d'eau marine")
            #EAU
            #ne pas inclure 2005 dans les gammes de reference suite a des problemes de methode differente Demande Pole ENV
            #gamme de reference calculee sur toute la periode jusqua annee n-1
            DonneesM = [dm for dm in DonneesM if int(dm[indiceColMar["Periode"]]) != 2005 and int(dm[indiceColMar["Periode"]]) < AnneeEtude]
            DonneesS = [ds for ds in DonneesS if int(ds[indiceColMar["Periode"]]) != 2005 and int(ds[indiceColMar["Periode"]]) < AnneeEtude]
            DonneesEx = [dx for dx in DonneesEx if int(dx[indiceColMar["Periode"]]) != 2005 and int(dx[indiceColMar["Periode"]]) < AnneeEtude]
            DonneesPerRefEau = [dpre for dpre in DonneesPerRefEau if int(dpre[indiceColMar["Periode"]]) != 2005]

            #necessite de Gamme de reference par station
            InfoStationsGamme = [[dm[indiceColMar["TypeStation"]], dm[indiceColMar["Zone"]], dm[indiceColMar["TypologieStation"]], dm[indiceColMar["TypeRef"]], dm[indiceColMar["Station"]]] for dm in DonneesM]
            #Filtre sur les stations de reference seulement Stations avec asterix dans le tableau de la methodologie du Bilan Environnement pour la colonne d'eau marine EAU
            InfoStationsGamme = [str(isgm[indiceColMar["TypeStation"]-1]) + ";" + isgm[indiceColMar["Zone"]-1] + ";" + isgm[indiceColMar["TypologieStation"]-1] + ";" + str(isgm[indiceColMar["TypeRef"]-1]) + ";" + str(isgm[indiceColMar["Station"]-1]) for
                                 isgm in InfoStationsGamme]
            InfoStationsGamme = list(set(InfoStationsGamme))
            InfoStationsGamme = [isgm.split(";") for isgm in InfoStationsGamme]
            #InfoStationsGamme = [isgm for isgm in InfoStationsGamme if isgm[indiceColMar["TypeRef"]-1] is not None and isgm[indiceColMar["TypeRef"]-1].find("Reference EAU") >= 0]
            StatEcartMoy = list(set([dpr[indiceColMar["Station"]] for dpr in DonneesPerRefEau if dpr[indiceColMar["TypeData"]].find("Metrique - Ecart Moyenne") >= 0]))
            for isg in InfoStationsGamme:
                isgi = InfoStationsGamme.index(isg)
                TypeStation = isg[indiceColMar["TypeStation"]-1]
                Zone = isg[indiceColMar["Zone"]-1]
                Typologie = isg[indiceColMar["TypologieStation"]-1]
                TypeRef = isg[indiceColMar["TypeRef"]-1]
                Station = isg[indiceColMar["Station"]-1]
                DonneesMStation = [d for d in DonneesM if d[indiceColMar["Station"]] == Station]

                PeriodeStation = AnneeMinMax(list(set([d[indiceColMar["Date"]] for d in DonneesMStation])))
                PeriodeStationMoy = list(set([int(dp[indiceColMar["Periode"]]) for dp in DonneesPerRefEau]))
                PeriodeStationMoy = str(min(PeriodeStationMoy)) + " - " + str(max(PeriodeStationMoy))

                DonneesSStation = [d for d in DonneesS if
                                   d[indiceColMar["Station"]] == Station]
                DonneesExStation = [d for d in DonneesEx if
                                    d[indiceColMar["Station"]] == Station]

                # ligne pour les statistiques de reference Amont et Grand BV
                LineStationMeth = ["Gamme de reference - methodes analytiques", TypeStation, Zone, Typologie, TypeRef, Station,
                                   None, PeriodeStation]
                LineStationLQ = ["Gamme de reference - Limites quantitatives", TypeStation, Zone, Typologie, TypeRef, Station,
                                 None, PeriodeStation]
                LineStationSigne = ["Gamme de reference - Nb LQ", TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationN = ["Gamme de reference - N",TypeStation, Zone, Typologie, TypeRef, Station,
                                None, PeriodeStation]
                LineStationMoy = ["Gamme de reference - Moyenne", TypeStation, Zone, Typologie, TypeRef, Station,
                                  None, PeriodeStation]
                LineStationEct = ["Gamme de reference - Ecart-Type", TypeStation, Zone, Typologie, TypeRef, Station,
                                  None, PeriodeStation]
                LineStationMed = ["Gamme de reference - Mediane",TypeStation, Zone, Typologie, TypeRef, Station,
                                  None, PeriodeStation]
                LineStationPct10 = ["Gamme de reference - Percentile 10",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct25 = ["Gamme de reference - Percentile 25", TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct75 = ["Gamme de reference - Percentile 75",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct80 = ["Gamme de reference - Percentile 80",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct85 = ["Gamme de reference - Percentile 85",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct90 = ["Gamme de reference - Percentile 90", TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationPct95 = ["Gamme de reference - Percentile 95",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                LineStationMin = ["Gamme de reference - Min", TypeStation, Zone, Typologie, TypeRef, Station,
                                  None, PeriodeStation]
                LineStationMax = ["Gamme de reference - Max", TypeStation, Zone, Typologie, TypeRef, Station,
                                  None, PeriodeStation]
                LineStationPerc90Moy = ["Gamme de reference - Percentile 90 Moyenne", TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStationMoy]
                LineStationNPerc90Moy = ["Gamme de reference - N Percentile 90 Moyenne", TypeStation, Zone, Typologie, TypeRef,
                                         Station,
                                         None, PeriodeStationMoy]
                if Station in StatEcartMoy:
                    LineStationNPerc90EcartMoy = ["Gamme de reference - N Percentile 90 Ecart Moyenne", TypeStation, Zone,
                                                  Typologie, TypeRef, Station,
                                                  None, PeriodeStationMoy]
                    LineStationPerc90EcartMoy = ["Gamme de reference - Percentile 90 Ecart Moyenne", TypeStation, Zone, Typologie, TypeRef, Station,
                                                 None, PeriodeStationMoy]

                for p in Parametres:
                    i = Parametres.index(p) + indiceColMar["PremierParam"]
                    # methodes analytiques EAU
                    LineStationMeth.append(FindMethodes(DonneesMStation,i))
                    ## VALEUR LQ
                    # EAU
                    LineStationLQ.append(FindLQ(DonneesSStation,i,DonneesExStation, IdentSuiv))
                    # NOMBRE DE VALEURS = A LA LQ
                    #EAU
                    LineStationSigne.append(FindSigne(DonneesSStation,i))
                    # DONNEES NON NULLES OU NON VIDES RECUPEREES + VALEURS MESUREES N
                    # EAU
                    VNStation = FindValsNonVides(DonneesExStation, i)
                    valNStation = VNStation[0]
                    ValsExStation = VNStation[1]
                    LineStationN.append(valNStation)
                    # EAU
                    if valNStation == 0:
                        LineStationMoy.append("")
                        LineStationEct.append("")
                        LineStationMed.append("")
                        LineStationPct10.append("")
                        LineStationPct25.append("")
                        LineStationPct75.append("")
                        LineStationPct80.append("")
                        LineStationPct85.append("")
                        LineStationPct90.append("")
                        LineStationPct95.append("")
                        LineStationMin.append("")
                        LineStationMax.append("")
                        LineStationPerc90Moy.append("")
                        LineStationNPerc90Moy.append("")
                        if Station in StatEcartMoy:
                            LineStationNPerc90EcartMoy.append("")
                            LineStationPerc90EcartMoy.append("")
                    else:
                        # Moyenne en utilisant la fonction RegleMoyenne pour effectuer moyenne par jour puis par mois et enfin sur toute la periode
                        # abandon remplace par moyenne numpy.mean sur liste de valeurs
                        CalcStation = FindStatsCal(ValsExStation)
                        LineStationMoy.append(CalcStation["moy"])
                        LineStationEct.append(CalcStation["ect"])
                        LineStationMed.append(CalcStation["med"])
                        LineStationPct10.append(CalcStation["perc10"])
                        LineStationPct25.append(CalcStation["perc25"])
                        LineStationPct75.append(CalcStation["perc75"])
                        LineStationPct80.append(CalcStation["perc80"])
                        LineStationPct85.append(CalcStation["perc85"])
                        LineStationPct90.append(CalcStation["perc90"])
                        LineStationPct95.append(CalcStation["perc95"])
                        # Min et Max
                        LineStationMin.append(CalcStation["min"])
                        LineStationMax.append(CalcStation["max"])
                        DonneesMoy = [dmo for dmo in DonneesPerRefEau if dmo[indiceColMar["Station"]] == Station
                                      and dmo[indiceColMar["TypeData"]].find("Metrique - Moyenne") >= 0]
                        VNMoyStation = FindValsNonVides(DonneesMoy, i)
                        valNMoyStation =VNMoyStation[0]
                        ValsNMoyStation = VNMoyStation[1]
                        LineStationNPerc90Moy.append(valNMoyStation)
                        if valNMoyStation == 0:
                            LineStationPerc90Moy.append("")
                        else:
                            CalcStMoy = FindStatsCal(ValsNMoyStation)
                            LineStationPerc90Moy.append(CalcStMoy["perc90"])
                        #Si station ayant ecart a la moyenne on calcule le percentile 90 sur les ecart de moyenne
                        if Station in StatEcartMoy:
                            DonneesEMoy = [dpr for dpr in DonneesPerRefEau if dpr[indiceColMar["Station"]] == Station
                                           and dpr[indiceColMar["TypeData"]].find("Metrique - Ecart Moyenne") >= 0]
                            VNEMoyStation = FindValsNonVides(DonneesEMoy, i)
                            valNEMoyStation =VNEMoyStation[0]
                            ValsNEMoyStation = VNEMoyStation[1]
                            LineStationNPerc90EcartMoy.append(valNEMoyStation)
                            if valNEMoyStation == 0:
                                LineStationPerc90EcartMoy.append("")
                            else:
                                CalcStEMoy = FindStatsCal(ValsNEMoyStation)
                                LineStationPerc90EcartMoy.append(CalcStEMoy["perc90"])

                # Insertion Donnees EAU
                DonneesStatGammeRef.append(LineStationMeth)
                DonneesStatGammeRef.append(LineStationLQ)
                DonneesStatGammeRef.append(LineStationSigne)
                DonneesStatGammeRef.append(LineStationN)
                DonneesStatGammeRef.append(LineStationMoy)
                DonneesStatGammeRef.append(LineStationEct)
                DonneesStatGammeRef.append(LineStationMed)
                DonneesStatGammeRef.append(LineStationPct10)
                DonneesStatGammeRef.append(LineStationPct25)
                DonneesStatGammeRef.append(LineStationPct75)
                DonneesStatGammeRef.append(LineStationPct80)
                DonneesStatGammeRef.append(LineStationPct85)
                DonneesStatGammeRef.append(LineStationPct90)
                DonneesStatGammeRef.append(LineStationPct95)
                DonneesStatGammeRef.append(LineStationMin)
                DonneesStatGammeRef.append(LineStationMax)
                DonneesStatGammeRef.append(LineStationPerc90Moy)
                DonneesStatGammeRef.append(LineStationNPerc90Moy)
                if Station in StatEcartMoy:
                    DonneesStatGammeRef.append(LineStationNPerc90EcartMoy)
                    DonneesStatGammeRef.append(LineStationPerc90EcartMoy)
        elif IdentSuiv == "RECOUVR32" or IdentSuiv == "RECOUVR12":
            if IdentSuiv == "RECOUVR32":
                arcpy.AddMessage("Oh non Gammes de reference pour le Substrat Vale NC")
            elif IdentSuiv == "RECOUVR12":
                arcpy.AddMessage("Oh non Gammes de reference pour le Substrat ACROPORA")
            #EAU
            #necessite de Gamme de reference par station

            InfoStationsGamme = [[dm[indiceColMar["TypeStation"]], dm[indiceColMar["Zone"]], dm[indiceColMar["TypologieStation"]], dm[indiceColMar["TypeRef"]], dm[indiceColMar["Station"]]] for dm in DonneesM]
            #Filtre sur les stations de reference seulement Stations avec asterix dans le tableau de la methodologie du Bilan Environnement pour la colonne d'eau marine EAU
            InfoStationsGamme = sorted(list(set([isgm[0] + ";" + isgm[1] + ";" + isgm[2] + ";" + isgm[3] + ";" + isgm[4] for
                                                 isgm in InfoStationsGamme])))
            InfoStationsGamme = [isgm.split(";") for isgm in InfoStationsGamme]
            InfoStationsGamme = [isgm for isgm in InfoStationsGamme if isgm[indiceColMar["TypeRef"]-1] is not None]
            for isg in InfoStationsGamme:
                isgi = InfoStationsGamme.index(isg)
                TypeStation = isg[indiceColMar["TypeStation"]-1]
                Zone = isg[indiceColMar["Zone"]-1]
                Typologie = isg[indiceColMar["TypologieStation"]-1]
                TypeRef = isg[indiceColMar["TypeRef"]-1]
                Station = isg[indiceColMar["Station"]-1]
                YearN_1 = int(datetime.datetime.now().strftime("%Y")) - 1
                DonneesMStation = [d for d in DonneesM if d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Periode"]]) < YearN_1]
                if len(DonneesMStation) > 0:
                    PeriodeStation = AnneeMinMax(list(set([d[indiceColMar["Date"]] for d in DonneesMStation])))

                    DonneesSStation = [d for d in DonneesS if
                                       d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Periode"]]) < YearN_1]
                    DonneesExStation = [d for d in DonneesEx if
                                        d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Periode"]]) < YearN_1]

                    # ligne pour les statistiques de reference Amont et Grand BV
                    LineStationMeth = ["Gamme de reference - methodes analytiques", TypeStation, Zone, Typologie, TypeRef, Station,
                                       None, PeriodeStation]
                    LineStationLQ = ["Gamme de reference - Limites quantitatives", TypeStation, Zone, Typologie, TypeRef, Station,
                                     None, PeriodeStation]
                    LineStationSigne = ["Gamme de reference - Nb LQ", TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationN = ["Gamme de reference - N",TypeStation, Zone, Typologie, TypeRef, Station,
                                    None, PeriodeStation]
                    LineStationMoy = ["Gamme de reference - Moyenne", TypeStation, Zone, Typologie, TypeRef, Station,
                                      None, PeriodeStation]
                    LineStationEct = ["Gamme de reference - Ecart-Type", TypeStation, Zone, Typologie, TypeRef, Station,
                                      None, PeriodeStation]
                    LineStationMed = ["Gamme de reference - Mediane",TypeStation, Zone, Typologie, TypeRef, Station,
                                      None, PeriodeStation]
                    LineStationPct10 = ["Gamme de reference - Percentile 10",TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct25 = ["Gamme de reference - Percentile 25", TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct75 = ["Gamme de reference - Percentile 75",TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct80 = ["Gamme de reference - Percentile 80",TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct85 = ["Gamme de reference - Percentile 85",TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct90 = ["Gamme de reference - Percentile 90", TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationPct95 = ["Gamme de reference - Percentile 95",TypeStation, Zone, Typologie, TypeRef, Station,
                                        None, PeriodeStation]
                    LineStationMin = ["Gamme de reference - Min", TypeStation, Zone, Typologie, TypeRef, Station,
                                      None, PeriodeStation]
                    LineStationMax = ["Gamme de reference - Max", TypeStation, Zone, Typologie, TypeRef, Station,
                                      None, PeriodeStation]


                    for p in Parametres:
                        i = Parametres.index(p) + indiceColMar["PremierParam"]
                        # methodes analytiques EAU
                        LineStationMeth.append(FindMethodes(DonneesMStation,i))
                        ## VALEUR LQ
                        # EAU
                        LineStationLQ.append(FindLQ(DonneesSStation,i,DonneesExStation,IdentSuiv))
                        # NOMBRE DE VALEURS = A LA LQ
                        #EAU
                        LineStationSigne.append(FindSigne(DonneesSStation,i))
                        # DONNEES NON NULLES OU NON VIDES RECUPEREES + VALEURS MESUREES N
                        # EAU
                        VNStation = FindValsNonVides(DonneesExStation, i)
                        valNStation = VNStation[0]
                        ValsExStation = VNStation[1]
                        LineStationN.append(valNStation)
                        # EAU
                        if valNStation == 0:
                            LineStationMoy.append("")
                            LineStationEct.append("")
                            LineStationMed.append("")
                            LineStationPct10.append("")
                            LineStationPct25.append("")
                            LineStationPct75.append("")
                            LineStationPct80.append("")
                            LineStationPct85.append("")
                            LineStationPct90.append("")
                            LineStationPct95.append("")
                            LineStationMin.append("")
                            LineStationMax.append("")
                        else:
                            # Moyenne en utilisant la fonction RegleMoyenne pour effectuer moyenne par jour puis par mois et enfin sur toute la periode
                            # abandon remplace par moyenne numpy.mean sur liste de valeurs
                            CalcStation = FindStatsCal(ValsExStation)
                            LineStationMoy.append(CalcStation["moy"])
                            LineStationEct.append(CalcStation["ect"])
                            LineStationMed.append(CalcStation["med"])
                            LineStationPct10.append(CalcStation["perc10"])
                            LineStationPct25.append(CalcStation["perc25"])
                            LineStationPct75.append(CalcStation["perc75"])
                            LineStationPct80.append(CalcStation["perc80"])
                            LineStationPct85.append(CalcStation["perc85"])
                            LineStationPct90.append(CalcStation["perc90"])
                            LineStationPct95.append(CalcStation["perc95"])
                            # Min et Max
                            LineStationMin.append(CalcStation["min"])
                            LineStationMax.append(CalcStation["max"])

                    # Insertion Donnees EAU
                    DonneesStatGammeRef.append(LineStationMeth)
                    DonneesStatGammeRef.append(LineStationLQ)
                    DonneesStatGammeRef.append(LineStationSigne)
                    DonneesStatGammeRef.append(LineStationN)
                    DonneesStatGammeRef.append(LineStationMoy)
                    DonneesStatGammeRef.append(LineStationEct)
                    DonneesStatGammeRef.append(LineStationMed)
                    DonneesStatGammeRef.append(LineStationPct10)
                    DonneesStatGammeRef.append(LineStationPct25)
                    DonneesStatGammeRef.append(LineStationPct75)
                    DonneesStatGammeRef.append(LineStationPct80)
                    DonneesStatGammeRef.append(LineStationPct85)
                    DonneesStatGammeRef.append(LineStationPct90)
                    DonneesStatGammeRef.append(LineStationPct95)
                    DonneesStatGammeRef.append(LineStationMin)
                    DonneesStatGammeRef.append(LineStationMax)

        else:
            arcpy.AddMessage("Ouuueeeee  Pas de Gamme de reference demandee pour le gros matou :-)")
            LineVideGamme = ["Pas de gamme de Gamme de reference attendue","","","","","",  None,""]
            for p in Parametres:
                LineVideGamme.append("")
            DonneesStatGammeRef.append(LineVideGamme)


    return DonneesStatGammeRef
