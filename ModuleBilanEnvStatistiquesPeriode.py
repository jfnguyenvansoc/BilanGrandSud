# -*- coding: utf8 -*-
from ModuleBilanEnv import FindSigne, FindLQ, FindMethodes, FindStatsCal, FindValsNonVides
from ModuleBilanEnv import NbSupPerc, Annee3ans, AnneeMinMax
import arcpy, datetime

# organisation des premieres colonnes
indiceCol = {"TypeData": 0, "TypeStation": 1, "Localisation": 2, "TypeBV": 3, "Zone": 4, "ZoneReference": 5,
             "Station": 6, "Date": 7, "Periode": 8, "PremierParam": 9}
#organisation des premieres colonnes Milieu Marin
indiceColMar = {"TypeData":0, "TypeStation":1, "Zone":2, "TypologieStation":3, "TypeRef":4, "Station":5, "Date":6, "Periode":7,"PremierParam":8}

# Fonction des Statistiques par Periode (Annee ou Saison selon le suivi)
def StatistiquesPeriode(Milieu, DonneesMethodes, DonneesSignes, DonneesExp, Parametres, StatistGP, IdentSuiv):
    if Milieu == "Eaux douces":
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
    elif Milieu == "Marin":
        #on conserve toutes les infos de station pour conserver ensuite l'ordre
        StatsDts = [str(d[indiceColMar["TypeStation"]]) + ";" + str(d[indiceColMar["Zone"]])+ ";" + str(d[indiceColMar["TypologieStation"]])+ ";" + str(d[indiceColMar["TypeRef"]])+ ";" + str(d[indiceColMar["Station"]]) for d in DonneesMethodes]
        StatsDts = list(set(StatsDts))
        StatsDts.sort()
        StatsDts = [s.split(";") for s in StatsDts]
        Stats = [s[indiceColMar["Station"]-1] for s in StatsDts]
        Stats.sort()
        SortieDonnees = list()
        #cas particulier pour les eaux paire de station suivi et ref
        if IdentSuiv == "EAU":
            PairRef = [[stref[indiceColMar["Station"]-1],stref[indiceColMar["TypeRef"]-1].replace("Reference EAU|","")] for stref in StatsDts if stref[indiceColMar["TypeRef"]-1] is not None and stref[indiceColMar["TypeRef"]-1].find("Reference EAU") >= 0]
            PairSuivi = [pr[1] for pr in PairRef]
        for s in Stats:
            TypeStation = [std[indiceColMar["TypeStation"]-1] for std in StatsDts if s == std[indiceColMar["Station"]-1]][0] #StZoneRef
            Zone = [std[indiceColMar["Zone"]-1] for std in StatsDts if s == std[indiceColMar["Station"]-1]][0] #StLoc
            TypologieStation = [std[indiceColMar["TypologieStation"]-1] for std in StatsDts if s == std[indiceColMar["Station"]-1]][0] #StBV
            TypeRef = [std[indiceColMar["TypeRef"]-1] for std in StatsDts if s == std[indiceColMar["Station"]-1]][0]
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

            DonneesMethStat = [d for d in DonneesMethodes if d[indiceColMar["Station"]] == s]
            DonneesSignStat = [d for d in DonneesSignes if d[indiceColMar["Station"]] == s]
            DonneesExpStat = [d for d in DonneesExp if d[indiceColMar["Station"]] == s]
            Periodes = [d[7] for d in DonneesMethStat]
            Periodes = list(set(Periodes))
            Periodes.sort()
            for per in Periodes:
                iPer = indiceColMar["Periode"]
                DonneesMthStPer = [d for d in DonneesMethStat if d[iPer] == per]
                DonneesSiStPer = [d for d in DonneesSignStat if d[iPer] == per]
                DonneesExpStPer = [d for d in DonneesExpStat if d[iPer] == per]
                LineMeth = ["Metrique - methodes analytiques",TypeStation, Zone, TypologieStation, TypeRef,s,None,per]
                LineSign = ["Metrique - Nb LQ",TypeStation, Zone, TypologieStation, TypeRef,s,None,per]
                LineLQ = ["Metrique - Limites quantitatives",TypeStation, Zone, TypologieStation, TypeRef,s,None,per]
                LineN = ["Metrique - N",TypeStation, Zone, TypologieStation, TypeRef,s,None,per]
                LineMoy = ["Metrique - Moyenne",TypeStation, Zone, TypologieStation,TypeRef, s,None,per]
                LineEct = ["Metrique - Ecart-type",TypeStation, Zone, TypologieStation, TypeRef,s, None,per]
                LineMed = ["Metrique - Mediane",TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc10 = ["Metrique - Percentile 10", TypeStation, Zone, TypologieStation, TypeRef, s,None,per]
                LinePerc25 = ["Metrique - Percentile 25", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc75 = ["Metrique - Percentile 75", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc80 = ["Metrique - Percentile 80", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc85 = ["Metrique - Percentile 85", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc90 = ["Metrique - Percentile 90", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LinePerc95 = ["Metrique - Percentile 95", TypeStation, Zone, TypologieStation, TypeRef, s, None, per]
                LineMin = ["Metrique - Min", TypeStation, Zone, TypologieStation,TypeRef, s, None, per]
                LineMax = ["Metrique - Max", TypeStation, Zone, TypologieStation,TypeRef, s, None, per]
                if IdentSuiv == "EAU":
                    if s in PairSuivi:
                        LineEcartMoy = ["Metrique - Ecart Moyenne",  TypeStation, Zone, TypologieStation,TypeRef, s, None, per]

                for p in Parametres:
                    i = Parametres.index(p) + indiceColMar["PremierParam"]
                    #methodes analytiques
                    LineMeth.append(FindMethodes(DonneesMthStPer,i,))
                    #nombre de valeurs a la LQ
                    LineSign.append(FindSigne(DonneesSiStPer,i))
                    #valeurs LQ
                    LineLQ.append(FindLQ(DonneesSiStPer,i,DonneesExpStPer,IdentSuiv))
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
                        if IdentSuiv == "EAU":
                            if s in PairSuivi:
                                LineEcartMoy.append("")
                    else:
                        #Moyenne en utilisant la fonction RegleMoyenne pour effectuer moyenne par jour puis par mois et enfin sur toute la periode
                        # abandon remplace par moyenne numpy.mean sur liste de valeurs
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
                        #superieur a Percentile 75 nb et pourcentage
                        if IdentSuiv == "EAU":
                            if s in PairSuivi:
                                RefSt = [stR[0] for stR in PairRef if stR[1] == s]
                                if len(RefSt) > 0:
                                    RefSt = RefSt[0]
                                    ValRef = [sd[i] for sd in SortieDonnees if sd[indiceColMar["Station"]] == RefSt and sd[indiceColMar["Periode"]] == per and sd[indiceColMar["TypeData"]].find("Metrique - Moyenne") >= 0]
                                    if len(ValRef) > 0:
                                        ValRef = ValRef[0]
                                        if ValRef == None or ValRef == "":
                                            ValRef = 0
                                        elif str(type(ValRef)) == "<type 'str'>":
                                            ValRef = float(ValRef.replace(",","."))
                                    else:
                                        ValRef = 0
                                    ValSuivi =  Calc["moy"]
                                    if ValSuivi == None or ValSuivi == "":
                                        ValSuivi = 0
                                    elif str(type(ValSuivi)) == "<type 'str'>":
                                        ValSuivi = float(ValSuivi.replace(",","."))
                                    #si un de des deux valeurs est manquante ou 0 on ne calcule pas demande Pole ENV
                                    if ValSuivi == 0 or ValRef == 0:
                                        ValEcartMoy = ""
                                    else:
                                        ValEcartMoy = ValSuivi - ValRef
                                    LineEcartMoy.append(ValEcartMoy)
                                else:
                                    LineEcartMoy.append("")

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
                if IdentSuiv == "EAU":
                    if s in PairSuivi:
                        SortieDonnees.append(LineEcartMoy)

                del LineMeth, LineSign,LineLQ, LineN, LineEct, LineMed, LinePerc10, LinePerc25, LinePerc90, LinePerc75
                del DonneesSiStPer, DonneesMthStPer, DonneesExpStPer
            del DonneesMethStat, DonneesSignStat, Periodes, DonneesExpStat
        del DonneesMethodes, DonneesSignes, Parametres,StatsDts, Stats

    return SortieDonnees

##Calcul des Statistiques sur 3 ans seulement pour EAU
def Statistiques3ans(DonneesM, DonneesS, DonneesEx, Parametres, IdentSuiv,AnneeMax):
    DonneesStat3ans = list()
    # CAS EAU (Colonne d'eau marine)
    if IdentSuiv == "EAU":
        arcpy.AddMessage("Oh non Statistiques sur 3 ans pour la colonne d'eau marine")
        # EAU
        # necessite de Gamme de reference par station

        InfoStationsGamme = [
            [dm[indiceColMar["TypeStation"]], dm[indiceColMar["Zone"]], dm[indiceColMar["TypologieStation"]],
             dm[indiceColMar["TypeRef"]], dm[indiceColMar["Station"]]] for dm in DonneesM]
        InfoStationsGamme = [str(isgm[0]) + ";" + isgm[1] + ";" + isgm[2] + ";" + str(isgm[3]) + ";" + str(isgm[4]) for isgm in InfoStationsGamme]
        InfoStationsGamme = list(set(InfoStationsGamme))
        InfoStationsGamme = [isgm.split(";") for isgm in InfoStationsGamme]

        for isg in InfoStationsGamme:
            isgi = InfoStationsGamme.index(isg)
            TypeStation = isg[indiceColMar["TypeStation"] - 1]
            Zone = isg[indiceColMar["Zone"] - 1]
            Typologie = isg[indiceColMar["TypologieStation"] - 1]
            TypeRef = isg[indiceColMar["TypeRef"] - 1]
            Station = isg[indiceColMar["Station"] - 1]
            DonneesMStation = [d for d in DonneesM if d[indiceColMar["Station"]] == Station]

            Periode3ansStation = Annee3ans(list(set([d[indiceColMar["Date"]] for d in DonneesMStation])))
            #AnneeMax = int(Periode3ansStation[Periode3ansStation.find("-")+1:len(Periode3ansStation)])

            #Filtre sur les 3 ans
            DonneesSStation = [d for d in DonneesS if
                               d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Date"]].strftime("%Y")) >= AnneeMax-2 and int(d[indiceColMar["Date"]].strftime("%Y")) <= AnneeMax]
            DonneesExStation = [d for d in DonneesEx if
                                d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Date"]].strftime("%Y")) >= AnneeMax-2 and int(d[indiceColMar["Date"]].strftime("%Y")) <= AnneeMax]
            DonneesMStation = [d for d in DonneesMStation if d[indiceColMar["Station"]] == Station and int(d[indiceColMar["Date"]].strftime("%Y")) >= AnneeMax-2 and int(d[indiceColMar["Date"]].strftime("%Y")) <= AnneeMax]
            # ligne pour les statistiques de reference Amont et Grand BV
            LineStationMeth = ["Metrique 3 ans - methodes analytiques", TypeStation, Zone, Typologie, TypeRef,
                               Station,
                               None, Periode3ansStation]
            LineStationLQ = ["Metrique 3 ans - Limites quantitatives", TypeStation, Zone, Typologie, TypeRef,
                             Station,
                             None, Periode3ansStation]
            LineStationSigne = ["Metrique 3 ans - Nb LQ", TypeStation, Zone, Typologie, TypeRef, Station,
                                None, Periode3ansStation]
            LineStationN = ["Metrique 3 ans - N", TypeStation, Zone, Typologie, TypeRef, Station,
                            None, Periode3ansStation]
            LineStationMoy = ["Metrique 3 ans - Moyenne", TypeStation, Zone, Typologie, TypeRef, Station,
                              None, Periode3ansStation]
            LineStationEct = ["Metrique 3 ans - Ecart-Type", TypeStation, Zone, Typologie, TypeRef, Station,
                              None, Periode3ansStation]
            LineStationMed = ["Metrique 3 ans - Mediane", TypeStation, Zone, Typologie, TypeRef, Station,
                              None, Periode3ansStation]
            LineStationPct10 = ["Metrique 3 ans - Percentile 10", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct25 = ["Metrique 3 ans - Percentile 25", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct75 = ["Metrique 3 ans - Percentile 75", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct80 = ["Metrique 3 ans - Percentile 75", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct85 = ["Metrique 3 ans - Percentile 75", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct90 = ["Metrique 3 ans - Percentile 90", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationPct95 = ["Metrique 3 ans - Percentile 95", TypeStation, Zone, Typologie, TypeRef,
                                Station,
                                None, Periode3ansStation]
            LineStationMin = ["Metrique 3 ans - Min", TypeStation, Zone, Typologie, TypeRef, Station,
                              None, Periode3ansStation]
            LineStationMax = ["Metrique 3 ans - Max", TypeStation, Zone, Typologie, TypeRef, Station,
                              None, Periode3ansStation]

            for p in Parametres:
                i = Parametres.index(p) + indiceColMar["PremierParam"]
                # methodes analytiques EAU
                LineStationMeth.append(FindMethodes(DonneesMStation,i))
                ## VALEUR LQ
                # EAU
                LineStationLQ.append(FindLQ(DonneesSStation,i,DonneesExStation, IdentSuiv))
                # NOMBRE DE VALEURS = A LA LQ
                # EAU
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
            DonneesStat3ans.append(LineStationMeth)
            DonneesStat3ans.append(LineStationLQ)
            DonneesStat3ans.append(LineStationSigne)
            DonneesStat3ans.append(LineStationN)
            DonneesStat3ans.append(LineStationMoy)
            DonneesStat3ans.append(LineStationEct)
            DonneesStat3ans.append(LineStationMed)
            DonneesStat3ans.append(LineStationPct10)
            DonneesStat3ans.append(LineStationPct25)
            DonneesStat3ans.append(LineStationPct75)
            DonneesStat3ans.append(LineStationPct80)
            DonneesStat3ans.append(LineStationPct85)
            DonneesStat3ans.append(LineStationPct90)
            DonneesStat3ans.append(LineStationPct95)
            DonneesStat3ans.append(LineStationMin)
            DonneesStat3ans.append(LineStationMax)
    else:
        arcpy.AddMessage("Ouuueeeee  Pas de statistique sur 3 ans demandee pour le gros matou :-)")
        LineVideGamme = ["Pas de gamme de statistique sur 3 ans attendue","","","","","",None,""]
        for p in Params:
            LineVideGamme.append("")
        DonneesStat3ans.append(LineVideGamme)

    return DonneesStat3ans
