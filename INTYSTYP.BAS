0.1 : CLS
COLOR 12, 0
PRINT TAB(20); " OPREDELENIE INTERVAL YSTYP PO X, Y "
COLOR 15, 0
1 INPUT " BBOD X, Y osnvn oryd "; Xos, Yos
1.1 INPUT " BBOD X, Y oryd "; X, Y
INPUT " Nor = "; nOr
2 INPUT " OH = "; OH
dX = X - Xos
dY = Y - Yos
Pi = 3.14159265358#
Baza = INT(SQR(dX ^ 2 + dY ^ 2) + .001)
A = ABS(ATN(dY / (dX + .001)) / Pi * 30) * 100
IF dX > 0 AND dY > 0 THEN Ygol.b = CINT(A)
IF dX < 0 AND dY > 0 THEN Ygol.b = CINT(3000 - A)
IF dX < 0 AND dY < 0 THEN Ygol.b = CINT(3000 + A)
IF dX > 0 AND dY < 0 THEN Ygol.b = CINT(6000 - A)
PRINT " Baza = "; Baza, " Ygol.b = "; Ygol.b
IF Ygol.b - 3000 < 0 THEN
        Ygol.b30 = Ygol.b + 6000 - 3000
        ELSE
        Ygol.b30 = Ygol.b + 3000
END IF
IF OH - 3000 < 0 THEN
        OH30 = OH + 6000 - 3000
        ELSE
        OH30 = OH + 3000
END IF
IF OH - 1500 < 0 THEN
        Aosiz = OH + 6000 - 1500
        ELSE
        Aosiz = OH - 1500
END IF
IF Ygol.b >= 0 AND Ygol.b <= 1500 THEN
                IF OH >= 0 AND OH < 1500 THEN
                        IF OH < Ygol.b THEN
                        Aos = OH: Aor = OH - 1500
                        ELSE
                        Aos = OH: Aor = OH + 1500
                        END IF
                ELSEIF OH >= 1500 AND OH < 3000 THEN
                        IF Aosiz < Ygol.b THEN
                        Aos = OH: Aor = OH + 1500
                        ELSE
                        Aos = OH - 1500: Aor = OH
                        END IF
                ELSEIF OH >= 3000 AND OH < 4500 THEN
                        IF OH < Ygol.b30 THEN
                        Aos = OH - 1500: Aor = OH
                        ELSE
                        Aos = OH + 1500: Aor = OH
                        END IF
                ELSEIF OH >= 4500 AND OH < 6000 THEN
                        IF Aosiz < Ygol.b30 THEN
                        Aos = OH + 1500: Aor = OH
                        ELSE
                        Aos = OH: Aor = OH - 1500
                        END IF
                ELSE
                END IF
        ELSEIF Ygol.b >= 1500 AND Ygol.b <= 3000 THEN
                IF OH >= 0 AND OH < 1500 THEN
                        IF Aosiz < Ygol.b30 THEN
                                Aos = OH + 1500: Aor = OH
                                ELSE
                                Aos = OH: Aor = OH - 1500
                        END IF
                ELSEIF OH >= 1500 AND OH < 3000 THEN
                        IF OH < Ygol.b THEN
                                Aos = OH: Aor = OH - 1500
                                ELSE
                                Aos = OH: Aor = OH - 1500
                        END IF
                ELSEIF OH >= 3000 AND OH < 4500 THEN
                        IF Aosiz < Ygol.b THEN
                                Aos = OH: Aor = OH + 1500
                                ELSE
                                Aos = OH - 1500: Aor = OH
                        END IF
                ELSEIF OH >= 4500 AND OH < 6000 THEN
                        IF OH < Ygol.b30 THEN
                                Aos = OH - 1500: Aor = OH
                                ELSE
                                Aos = OH + 1500: Aor = OH
                        END IF
                ELSE
                END IF
        ELSEIF Ygol.b >= 3000 AND Ygol.b <= 4500 THEN
                IF OH >= 0 AND OH < 1500 THEN
                        IF OH < Ygol.b30 THEN
                                Aos = OH - 1500: Aor = OH
                                ELSE
                                Aos = OH + 1500: Aor = OH
                        END IF
                ELSEIF OH >= 1500 AND OH < 3000 THEN
                        IF Aosiz < Ygol.b THEN
                                Aos = OH + 1500: Aor = OH
                                ELSE
                                Aos = OH: Aor = OH - 1500
                        END IF
                ELSEIF OH >= 3000 AND OH < 4500 THEN
                        IF OH < Ygol.b THEN
                                Aos = OH: Aor = OH - 1500
                                ELSE
                                Aos = OH: Aor = OH + 1500
                        END IF
                ELSEIF OH >= 4500 AND OH < 6000 THEN
                        IF Aosiz < Ygol.b THEN
                                Aos = OH: Aor = OH + 1500
                                ELSE
                                Aos = OH - 1500: Aor = OH
                        END IF
                ELSE
                END IF
        ELSEIF Ygol.b >= 4500 AND Ygol.b <= 6000 THEN
                IF OH >= 0 AND OH < 1500 THEN
                        IF Aosiz < Ygol.b THEN
                                Aos = OH: Aor = OH + 1500
                                ELSE
                                Aos = OH - 1500: Aor = OH
                        END IF
                ELSEIF OH >= 1500 AND OH < 3000 THEN
                        IF OH < Ygol.b30 THEN
                                Aos = OH - 1500: Aor = OH
                                ELSE
                                Aos = OH + 1500: Aor = OH
                        END IF
                ELSEIF OH >= 3000 AND OH < 4500 THEN
                        IF Aosiz < Ygol.b THEN
                                Aos = OH + 1500: Aor = OH
                                ELSE
                                Aos = OH: Aor = OH - 1500
                        END IF
                ELSEIF OH >= 4500 AND OH < 6000 THEN
                        IF OH < Ygol.b THEN
                                Aos = OH: Aor = OH - 1500
                                ELSE
                                Aos = OH: Aor = OH + 1500
                        END IF
                ELSE
                END IF
        ELSE
END IF

IF Aos > 6000 THEN Aos = Aos - 6000
IF Aos < 0 THEN Aos = Aos + 6000
IF Aor > 6000 THEN Aor = Aor - 6000
IF Aor < 0 THEN Aor = Aor + 6000
PRINT " Aos, Aor "; Aos, Aor
IF Aos > Aor AND Aos <= 4500 AND Aor >= 1500 THEN
        Alev = Aos: Aprav = Aor
        ELSEIF Aor > Aos AND Aor <= 4500 AND Aos >= 1500 THEN
        Alev = Aor: Aprav = Aos
        ELSEIF Aos >= 4500 AND Aor <= 1500 THEN
        Alev = Aor: Aprav = Aos
        ELSEIF Aor >= 4500 AND Aos <= 1500 THEN
        Alev = Aos: Aprav = Aor
        ELSEIF Aos > Aor THEN
        Alev = Aos: Aprav = Aor
        ELSE
        Alev = Aor: Aprav = Aos
END IF
PRINT " Alev, Aprav "; Alev, Aprav
'LEV
  IF Alev = Aos THEN Xl = Xos: Yl = Yos
  IF Alev = Aor THEN Xl = X: Yl = Y
'PRAV
  IF Aprav = Aos THEN Xp = Xos: Yp = Yos
  IF Aprav = Aor THEN Xp = X: Yp = Y
 
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = SQR(dxso ^ 2 + dyso ^ 2)
  aso = ABS(ATN(dyso / (dxso + .1)) / 3.141592 * 30) * 100

  IF dxso > 0 AND dyso > 0 THEN ygol.baz = INT(aso)
  IF dxso < 0 AND dyso > 0 THEN ygol.baz = INT(3000 - aso)
  IF dxso < 0 AND dyso < 0 THEN ygol.baz = INT(3000 + aso)
  IF dxso > 0 AND dyso < 0 THEN ygol.baz = INT(6000 - aso)
  IF Alev <= 1500 AND Aprav >= 4500 THEN
  fi = ABS(Alev + 6000 - Aprav)
  ELSEIF Alev >= 4500 AND Aprav <= 1500 THEN
  fi = ABS(Alev - (Aprav + 6000))
  ELSE
   fi = ABS(Alev - Aprav)
  END IF
  IF Alev <= 1500 AND ygol.baz >= 4500 THEN
  blev = ABS(Alev + 6000 - ygol.baz)
  ELSEIF Alev >= 4500 AND ygol.baz <= 1500 THEN
  blev = ABS(Alev - (ygol.baz + 6000))
  ELSE
  blev = ABS(Alev - ygol.baz)
  END IF

  IF ygol.baz - 3000 < 0 THEN
  ybazp = ygol.baz + 3000
  ELSE
  ybazp = ygol.baz - 3000
  END IF

  IF Aprav <= 1500 AND ybazp >= 4500 THEN
  bprav = ABS(Aprav + 6000 - ybazp)
  ELSEIF Aprav >= 4500 AND ybazp <= 1500 THEN
  bprav = ABS(Aprav - (ybazp + 6000))
  ELSE
  bprav = ABS(Aprav - ybazp)
  END IF
 
  Dlev = baz / (SIN(fi / 100 * 6 * 3.141592 / 180) + .001) * SIN(bprav / 100 * 6 * 3.141592 / 180)
  Dprav = baz / (SIN(fi / 100 * 6 * 3.141592 / 180) + .001) * SIN(blev / 100 * 6 * 3.141592 / 180)
  Xp = COS(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Xl
  Yp = SIN(Alev / 100 * 6 * 3.141592 / 180) * Dlev + Yl
IF Alev = OH THEN
        Ystyp = Dlev: Interval = Dprav
        ELSE
        Ystyp = Dprav: Interval = Dlev
END IF
IF Aos <> OH THEN Ystyp = Ystyp * -1
IF Aprav = Aos THEN
        Interval = Interval * -1
        ELSE
        Interval = Interval
END IF
PRINT " Ystyp = "; INT(Ystyp + .001), " Interval ="; INT(Interval + .001)
dxtP100 = 32: dxtP200 = 23: dxtP300 = 18: dxtP400 = 14: dxtP500 = 11: dxtP600 = 7: dxtP700 = 4: dxtP800 = 4: dxtP900 = 10: dxtP1000 = 15: dxtP1100 = 20
dxtY100 = 27: dxtY200 = 20: dxtY300 = 16: dxtY400 = 13: dxtY500 = 10: dxtY600 = 6: dxtY700 = 4: dxtY800 = 4: dxtY900 = 8: dxtY1000 = 12: dxtY1100 = 17
dxt1100 = 24: dxt1200 = 19: dxt1300 = 16: dxt1400 = 13: dxt1500 = 10: dxt1600 = 6: dxt1700 = 4: dxt1800 = 4: dxt1900 = 8: dxt11000 = 12: dxt11100 = 15
dxt2100 = 21: dxt2200 = 17: dxt2300 = 14: dxt2400 = 12: dxt2500 = 8: dxt2600 = 5: dxt2700 = 4: dxt2800 = 4: dxt2900 = 7: dxt21000 = 10: dxt21100 = 14
dxt3100 = 18: dxt3200 = 16: dxt3300 = 13: dxt3400 = 10: dxt3500 = 7: dxt3600 = 4: dxt3700 = 4: dxt3800 = 4: dxt3900 = 6: dxt31000 = 8: dxt31100 = 11
dxt4100 = 14: dxt4200 = 12: dxt4300 = 10: dxt4400 = 8: dxt4500 = 6: dxt4600 = 4: dxt4700 = 4: dxt4800 = 4: dxt4900 = 4: dxt41000 = 6: dxt41100 = 8

dPP100 = Ystyp * -1 / dxtP100: dPP200 = Ystyp / dxtP200 * -1: dPP300 = Ystyp / dxtP300 * -1: dPP400 = Ystyp / dxtP400 * -1: dPP500 = Ystyp / dxtP500 * -1: dPP600 = Ystyp / dxtP600 * -1: dPP700 = Ystyp / dxtP700 * -1: dPP800 = Ystyp / dxtP800 * -1
dPP900 = Ystyp / dxtP900 * -1: dPP1000 = Ystyp / dxtP1000 * -1: dPP1100 = Ystyp / dxtP1100 * -1
dPY100 = Ystyp * -1 / dxtY100: dPY200 = Ystyp / dxtY200 * -1: dPY300 = Ystyp / dxtY300 * -1: dPY400 = Ystyp / dxtY400 * -1: dPY500 = Ystyp / dxtY500 * -1: dPY600 = Ystyp / dxtY600 * -1: dPY700 = Ystyp / dxtY700 * -1: dPY800 = Ystyp / dxtY800 * -1
dPY900 = Ystyp / dxtY900 * -1: dPY1000 = Ystyp / dxtY1000 * -1: dPY1100 = Ystyp / dxtY1100 * -1
dP1100 = Ystyp * -1 / dxt1100: dP1200 = Ystyp / dxt1200 * -1: dP1300 = Ystyp / dxt1300 * -1: dP1400 = Ystyp / dxt1400 * -1: dP1500 = Ystyp / dxt1500 * -1: dP1600 = Ystyp / dxt1600 * -1: dP1700 = Ystyp / dxt1700 * -1: dP1800 = Ystyp / dxt1800 * -1
dP1900 = Ystyp / dxt1900 * -1: dP11000 = Ystyp / dxt11000 * -1: dP11100 = Ystyp / dxt11100 * -1
dP2100 = Ystyp * -1 / dxt2100: dP2200 = Ystyp / dxt2200 * -1: dP2300 = Ystyp / dxt2300 * -1: dP2400 = Ystyp / dxt2400 * -1: dP2500 = Ystyp / dxt2500 * -1: dP2600 = Ystyp / dxt2600 * -1: dP2700 = Ystyp / dxt2700 * -1: dP2800 = Ystyp / dxt2800 * -1
dP2900 = Ystyp / dxt2900 * -1: dP21000 = Ystyp / dxt21000 * -1: dP21100 = Ystyp / dxt21100 * -1
dP3100 = Ystyp * -1 / dxt3100: dP3200 = Ystyp / dxt3200 * -1: dP3300 = Ystyp / dxt3300 * -1: dP3400 = Ystyp / dxt3400 * -1: dP3500 = Ystyp / dxt3500 * -1: dP3600 = Ystyp / dxt3600 * -1: dP3700 = Ystyp / dxt3700 * -1: dP3800 = Ystyp / dxt3800 * -1
dP3900 = Ystyp / dxt3900 * -1: dP31000 = Ystyp / dxt31000 * -1: dP31100 = Ystyp / dxt31100 * -1
dP4100 = Ystyp * -1 / dxt4100: dP4200 = Ystyp / dxt4200 * -1: dP4300 = Ystyp / dxt4300 * -1: dP4400 = Ystyp / dxt4400 * -1: dP4500 = Ystyp / dxt4500 * -1: dP4600 = Ystyp / dxt4600 * -1: dP4700 = Ystyp / dxt4700 * -1: dP4800 = Ystyp / dxt4800 * -1
dP4900 = Ystyp / dxt4900 * -1: dP41000 = Ystyp / dxt41000 * -1: dP41100 = Ystyp / dxt41100 * -1

DtP100 = 5.8: DtP200 = 8.6: DtP300 = 10.8: DtP400 = 12.4: DtP500 = 13.8: DtP600 = 14.6: DtP700 = 15.2: DtP800 = 15: DtP900 = 14.4: DtP1000 = 13.2: DtP1100 = 11.4
DtY100 = 4.4: DtY200 = 6.8: DtY300 = 8.8: DtY400 = 10.4: DtY500 = 11.6: DtY600 = 12.4: DtY700 = 12.8: DtY800 = 12.6: DtY900 = 12.2: DtY1000 = 11: DtY1100 = 9.6
Dt1100 = 3.8: Dt1200 = 5.8: Dt1300 = 7.6: Dt1400 = 9.2: Dt1500 = 10.2: Dt1600 = 11: Dt1700 = 11.4: Dt1800 = 11.4: Dt1900 = 10.8: Dt11000 = 10: Dt11100 = 8.6
Dt2100 = 3: Dt2200 = 4.8: Dt2300 = 6.4: Dt2400 = 7.8: Dt2500 = 9: Dt2600 = 9.6: Dt2700 = 10: Dt2800 = 10: Dt2900 = 9.6: Dt21000 = 8.6: Dt21100 = 7.4
Dt3100 = 2.2: Dt3200 = 3.8: Dt3300 = 5.4: Dt3400 = 6.6: Dt3500 = 7.4: Dt3600 = 8: Dt3700 = 8.3: Dt3800 = 8.3: Dt3900 = 7.8: Dt31000 = 7: Dt31100 = 6
Dt4100 = 1.6: Dt4200 = 3: Dt4300 = 4: Dt4400 = 5: Dt4500 = 5.6: Dt4600 = 6.1: Dt4700 = 6.3: Dt4800 = 6.3: Dt4900 = 5.8: Dt41000 = 5.4: Dt41100 = 4.6

dYgP100 = (Interval / DtP100 * .95) * -1: dYgP200 = (Interval / DtP200 * .95) * -1: dYgP300 = (Interval / DtP300 * .95) * -1: dYgP400 = (Interval / DtP400 * .95) * -1: dYgP500 = (Interval / DtP500 * .95) * -1: dYgP600 = (Interval / DtP600 * .95) * -1: dYgP700 = (Interval / DtP700 * .95) * -1: dYgP800 = (Interval / DtP800 * .95) * -1: dYgP900 = (Interval / DtP900 * .95) * -1
dYgP1000 = (Interval / DtP1000 * .95) * -1: dYgP1100 = (Interval / DtP1100 * .95) * -1
dYgY100 = (Interval / DtY100 * .95) * -1: dYgY200 = (Interval / DtY200 * .95) * -1: dYgY300 = (Interval / DtY300 * .95) * -1: dYgY400 = (Interval / DtY400 * .95) * -1: dYgY500 = (Interval / DtY500 * .95) * -1: dYgY600 = (Interval / DtY600 * .95) * -1: dYgY700 = (Interval / DtY700 * .95) * -1: dYgY800 = (Interval / DtY800 * .95) * -1: dYgY900 = (Interval / DtY900 * .95) * -1
dYgY1000 = (Interval / DtY1000 * .95) * -1: dYgY1100 = (Interval / DtY1100 * .95) * -1
dYg1100 = (Interval / Dt1100 * .95) * -1: dYg1200 = (Interval / Dt1200 * .95) * -1: dYg1300 = (Interval / Dt1300 * .95) * -1: dYg1400 = (Interval / Dt1400 * .95) * -1: dYg1500 = (Interval / Dt1500 * .95) * -1: dYg1600 = (Interval / Dt1600 * .95) * -1: dYg1700 = (Interval / Dt1700 * .95) * -1: dYg1800 = (Interval / Dt1800 * .95) * -1: dYg1900 = (Interval / Dt1900 * .95) * -1
dYg11000 = (Interval / Dt11000 * .95) * -1: dYg11100 = (Interval / Dt11100 * .95) * -1
dYg2100 = (Interval / Dt2100 * .95) * -1: dYg2200 = (Interval / Dt2200 * .95) * -1: dYg2300 = (Interval / Dt2300 * .95) * -1: dYg2400 = (Interval / Dt2400 * .95) * -1: dYg2500 = (Interval / Dt2500 * .95) * -1: dYg2600 = (Interval / Dt2600 * .95) * -1: dYg2700 = (Interval / Dt2700 * .95) * -1: dYg2800 = (Interval / Dt2800 * .95) * -1: dYg2900 = (Interval / Dt2900 * .95) * -1
dYg21000 = (Interval / Dt21000 * .95) * -1: dYg21100 = (Interval / Dt21100 * .95) * -1
dYg3100 = (Interval / Dt3100 * .95) * -1: dYg3200 = (Interval / Dt3200 * .95) * -1: dYg3300 = (Interval / Dt3300 * .95) * -1: dYg3400 = (Interval / Dt3400 * .95) * -1: dYg3500 = (Interval / Dt3500 * .95) * -1: dYg3600 = (Interval / Dt3600 * .95) * -1: dYg3700 = (Interval / Dt3700 * .95) * -1: dYg3800 = (Interval / Dt3800 * .95) * -1: dYg3900 = (Interval / Dt3900 * .95) * -1
dYg31000 = (Interval / Dt31000 * .95) * -1: dYg31100 = (Interval / Dt31100 * .95) * -1
dYg4100 = (Interval / Dt4100 * .95) * -1: dYg4200 = (Interval / Dt4200 * .95) * -1: dYg4300 = (Interval / Dt4300 * .95) * -1: dYg4400 = (Interval / Dt4400 * .95) * -1: dYg4500 = (Interval / Dt4500 * .95) * -1: dYg4600 = (Interval / Dt4600 * .95) * -1: dYg4700 = (Interval / Dt4700 * .95) * -1: dYg4800 = (Interval / Dt4800 * .95) * -1: dYg4900 = (Interval / Dt4900 * .95) * -1
dYg41000 = (Interval / Dt41000 * .95) * -1: dYg41100 = (Interval / Dt41100 * .95) * -1
PRINT "........................................................................"
PRINT TAB(20); "<<  POPRAVKI na YSTYP INTERVAL "; " OH = "; OH; " >> "
PRINT " ZAR "; TAB(8); " POLN "; TAB(20); " YMEN "; TAB(32); " PERV "; TAB(44); " VTOR "; TAB(56); " TRET "; TAB(68); " CHET "
PRINT TAB(7); " dP / dYg "; TAB(19); " dP / dYg "; TAB(31); " dP / dYg "; TAB(43); " dP / dYg "; TAB(55); " dP / dYg "; TAB(67); " dP / dYg "
PRINT " PRIC "
PRINT " 100 "; ; TAB(7); INT(dPP100); "/"; INT(dYgP100); TAB(19); INT(dPY100); "/"; INT(dYgY100); TAB(31); INT(dP1100); "/"; INT(dYg1100); TAB(43); INT(dP2100); "/"; INT(dYg2100); TAB(55); INT(dP3100); "/"; INT(dYg3100); TAB(67); INT(dP4100); "/"; INT(dYg4100)
PRINT " 200 "; TAB(7); INT(dPP200); "/"; INT(dYgP200); TAB(19); INT(dPY200); "/"; INT(dYgY200); TAB(31); INT(dP1200); "/"; INT(dYg1200); TAB(43); INT(dP2200); "/"; INT(dYg2200); TAB(55); INT(dP3200); "/"; INT(dYg3200); TAB(67); INT(dP4200); "/"; INT(dYg4200)
PRINT " 300 "; TAB(7); INT(dPP300); "/"; INT(dYgP300); TAB(19); INT(dPY300); "/"; INT(dYgY300); TAB(31); INT(dP1300); "/"; INT(dYg1300); TAB(43); INT(dP2300); "/"; INT(dYg2300); TAB(55); INT(dP3300); "/"; INT(dYg3300); TAB(67); INT(dP4300); "/"; INT(dYg4300)
PRINT " 400 "; TAB(7); INT(dPP400); "/"; INT(dYgP400); TAB(19); INT(dPY400); "/"; INT(dYgY400); TAB(31); INT(dP1400); "/"; INT(dYg1400); TAB(43); INT(dP2400); "/"; INT(dYg2400); TAB(55); INT(dP3400); "/"; INT(dYg3400); TAB(67); INT(dP4400); "/"; INT(dYg4400)
PRINT " 500 "; TAB(7); INT(dPP500); "/"; INT(dYgP500); TAB(19); INT(dPY500); "/"; INT(dYgY500); TAB(31); INT(dP1500); "/"; INT(dYg1500); TAB(43); INT(dP2500); "/"; INT(dYg2500); TAB(55); INT(dP3500); "/"; INT(dYg3500); TAB(67); INT(dP4500); "/"; INT(dYg4500)
PRINT " 600 "; TAB(7); INT(dPP600); "/"; INT(dYgP600); TAB(19); INT(dPY600); "/"; INT(dYgY600); TAB(31); INT(dP1600); "/"; INT(dYg1600); TAB(43); INT(dP2600); "/"; INT(dYg2600); TAB(55); INT(dP3600); "/"; INT(dYg3600); TAB(67); INT(dP4600); "/"; INT(dYg4600)
PRINT " 700 "; TAB(7); INT(dPP700); "/"; INT(dYgP700); TAB(19); INT(dPY700); "/"; INT(dYgY700); TAB(31); INT(dP1700); "/"; INT(dYg1700); TAB(43); INT(dP2700); "/"; INT(dYg2700); TAB(55); INT(dP3700); "/"; INT(dYg3700); TAB(67); INT(dP4700); "/"; INT(dYg4700)
PRINT " 800 "; TAB(7); INT(dPP800); "/"; INT(dYgP800); TAB(19); INT(dPY800); "/"; INT(dYgY800); TAB(31); INT(dP1800); "/"; INT(dYg1800); TAB(43); INT(dP2800); "/"; INT(dYg2800); TAB(55); INT(dP3800); "/"; INT(dYg3800); TAB(67); INT(dP4800); "/"; INT(dYg4800)
PRINT " 900 "; TAB(7); INT(dPP900); "/"; INT(dYgP900); TAB(19); INT(dPY900); "/"; INT(dYgY900); TAB(31); INT(dP1900); "/"; INT(dYg1900); TAB(43); INT(dP2900); "/"; INT(dYg2900); TAB(55); INT(dP3900); "/"; INT(dYg3900); TAB(67); INT(dP4900); "/"; INT(dYg4900)
PRINT " 1000 "; TAB(7); INT(dPP1000); "/"; INT(dYgP1000); TAB(19); INT(dPY1000); "/"; INT(dYgY1000); TAB(31); INT(dP11000); "/"; INT(dYg11000); TAB(43); INT(dP21000); "/"; INT(dYg21000); TAB(55); INT(dP31000); "/"; INT(dYg31000); TAB(67); INT(dP41000); "/"; INT(dYg41000)
PRINT " 1100 "; TAB(7); INT(dPP1100); "/"; INT(dYgP1100); TAB(19); INT(dPY1100); "/"; INT(dYgY1100); TAB(31); INT(dP11100); "/"; INT(dYg11100); TAB(43); INT(dP21100); "/"; INT(dYg21100); TAB(55); INT(dP31100); "/"; INT(dYg31100); TAB(67); INT(dP41100); "/"; INT(dYg41100)
PRINT "........................................................................"
INPUT " Zapisat 1 - Da, 0 - Net "; vvod
IF vvod = 0 THEN GOTO 999
OPEN "popnaiy.kkk" FOR APPEND AS #1
t1$ = "......................................................................."
PRINT #1, t1$
t1$ = " <<  POPRAVKI na YSTYP INTERVAL >> "
PRINT #1, t1$
t1$ = " OH = ": t2 = OH
PRINT #1, t1$, t2
t1$ = " Nor = ": t2 = nOr
PRINT #1, t1$, t2
t1$ = " ZAR. ": t2$ = "POL": t3$ = "YNEN": t4$ = "PERV": t5$ = "VTOR": t6$ = "TRET": t7$ = "CHETV"
PRINT #1, t1$; TAB(14); t2$; TAB(27); t3$; TAB(41); t4$; TAB(54); t5$; TAB(67); t6$; TAB(80); t7$
t1$ = "PRIC": t2$ = "dP/dYg"
PRINT #1, t1$; TAB(14); t2$; TAB(27); t2$; TAB(41); t2$; TAB(54); t2$; TAB(67); t2$; TAB(80); t2$
t1 = 100: t2 = INT(dPP100): t3$ = "/": t4 = INT(dYgP100): t5 = INT(dPY100): t6$ = "/": t7 = INT(dYgY100): t8 = INT(dP1100): t9$ = "/": t10 = INT(dYg1100): t11 = INT(dP2100): t12$ = "/": t13 = INT(dYg2100): t14 = INT(dP3100): t15$ = "/": t16 = INT(dYg3100): t17 = INT(dP4100): t18$ = "/": t19 = INT(dYg4100)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 200: t2 = INT(dPP200): t3$ = "/": t4 = INT(dYgP200): t5 = INT(dPY200): t6$ = "/": t7 = INT(dYgY200): t8 = INT(dP1200): t9$ = "/": t10 = INT(dYg1200): t11 = INT(dP2200): t12$ = "/": t13 = INT(dYg2200): t14 = INT(dP3200): t15$ = "/": t16 = INT(dYg3200): t17 = INT(dP4200): t18$ = "/": t19 = INT(dYg4200)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 300: t2 = INT(dPP300): t3$ = "/": t4 = INT(dYgP300): t5 = INT(dPY300): t6$ = "/": t7 = INT(dYgY300): t8 = INT(dP1300): t9$ = "/": t10 = INT(dYg1300): t11 = INT(dP2300): t12$ = "/": t13 = INT(dYg2300): t14 = INT(dP3300): t15$ = "/": t16 = INT(dYg3300): t17 = INT(dP4300): t18$ = "/": t19 = INT(dYg4300)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 400: t2 = INT(dPP400): t3$ = "/": t4 = INT(dYgP400): t5 = INT(dPY400): t6$ = "/": t7 = INT(dYgY400): t8 = INT(dP1400): t9$ = "/": t10 = INT(dYg1400): t11 = INT(dP2400): t12$ = "/": t13 = INT(dYg2400): t14 = INT(dP3400): t15$ = "/": t16 = INT(dYg3400): t17 = INT(dP4400): t18$ = "/": t19 = INT(dYg4400)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 500: t2 = INT(dPP500): t3$ = "/": t4 = INT(dYgP500): t5 = INT(dPY500): t6$ = "/": t7 = INT(dYgY500): t8 = INT(dP1500): t9$ = "/": t10 = INT(dYg1500): t11 = INT(dP2500): t12$ = "/": t13 = INT(dYg2500): t14 = INT(dP3500): t15$ = "/": t16 = INT(dYg3500): t17 = INT(dP4500): t18$ = "/": t19 = INT(dYg4500)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 600: t2 = INT(dPP600): t3$ = "/": t4 = INT(dYgP600): t5 = INT(dPY600): t6$ = "/": t7 = INT(dYgY600): t8 = INT(dP1600): t9$ = "/": t10 = INT(dYg1600): t11 = INT(dP2600): t12$ = "/": t13 = INT(dYg2600): t14 = INT(dP3600): t15$ = "/": t16 = INT(dYg3600): t17 = INT(dP4600): t18$ = "/": t19 = INT(dYg4600)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 700: t2 = INT(dPP700): t3$ = "/": t4 = INT(dYgP700): t5 = INT(dPY700): t6$ = "/": t7 = INT(dYgY700): t8 = INT(dP1700): t9$ = "/": t10 = INT(dYg1700): t11 = INT(dP2700): t12$ = "/": t13 = INT(dYg2700): t14 = INT(dP3700): t15$ = "/": t16 = INT(dYg3700): t17 = INT(dP4700): t18$ = "/": t19 = INT(dYg4700)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 800: t2 = INT(dPP800): t3$ = "/": t4 = INT(dYgP800): t5 = INT(dPY800): t6$ = "/": t7 = INT(dYgY800): t8 = INT(dP1800): t9$ = "/": t10 = INT(dYg1800): t11 = INT(dP2800): t12$ = "/": t13 = INT(dYg2800): t14 = INT(dP3800): t15$ = "/": t16 = INT(dYg3800): t17 = INT(dP4800): t18$ = "/": t19 = INT(dYg4800)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 900: t2 = INT(dPP900): t3$ = "/": t4 = INT(dYgP900): t5 = INT(dPY900): t6$ = "/": t7 = INT(dYgY900): t8 = INT(dP1900): t9$ = "/": t10 = INT(dYg1900): t11 = INT(dP2900): t12$ = "/": t13 = INT(dYg2900): t14 = INT(dP3900): t15$ = "/": t16 = INT(dYg3900): t17 = INT(dP4900): t18$ = "/": t19 = INT(dYg4900)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 1000: t2 = INT(dPP1000): t3$ = "/": t4 = INT(dYgP1000): t5 = INT(dPY1000): t6$ = "/": t7 = INT(dYgY1000): t8 = INT(dP11000): t9$ = "/": t10 = INT(dYg11000): t11 = INT(dP21000): t12$ = "/": t13 = INT(dYg21000): t14 = INT(dP31000): t15$ = "/": t16 = INT(dYg31000): t17 = INT(dP41000): t18$ = "/": t19 = INT(dYg41000)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1 = 1100: t2 = INT(dPP1100): t3$ = "/": t4 = INT(dYgP1100): t5 = INT(dPY1100): t6$ = "/": t7 = INT(dYgY1100): t8 = INT(dP11100): t9$ = "/": t10 = INT(dYg11100): t11 = INT(dP21100): t12$ = "/": t13 = INT(dYg21100): t14 = INT(dP31100): t15$ = "/": t16 = INT(dYg31100): t17 = INT(dP41100): t18$ = "/": t19 = INT(dYg41100)
PRINT #1, t1; TAB(14); t2; t3$; t4; TAB(27); t5; t6$; t7; TAB(41); t8; t9$; t10; TAB(54); t11; t12$; t13; TAB(67); t14; t15$; t16; TAB(80); t17; t18$; t19
t1$ = "......................................................................."
PRINT #1, t1$
CLOSE #1
999 COLOR 10, 0
PRINT " BBOD 9 = END, 1 = pomen.OH, 2 = pomen.Osn.Or, 3 = pomen.Or "
PRINT " BBOD 4 = Ochist. zapis "
COLOR 15, 0
INPUT q
IF q = 9 THEN END
IF q = 1 THEN GOTO 2
IF q = 0 THEN GOTO 999
IF q = 2 THEN GOTO 0.1
IF q = 3 THEN GOTO 1.1
IF q = 4 THEN GOTO 4
GOTO 2
4 OPEN "popnaiy.kkk" FOR OUTPUT AS #1
CLOSE #1
GOTO 999

