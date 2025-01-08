DECLARE SUB OBRATZASDA ()
DECLARE SUB OBRZASA ()
CLS
5001 PRINT TAB(30); "<<TOPO>>"
PRINT " 0 - BYSOLNUI XOD (PGZ)"
PRINT " 0.1 - (PGZ)"
PRINT " 1 - OBRATNAIA ZASECHKA PO TREM YGLAM"
PRINT " 2 - OGZ"
PRINT " 3 - PZ"
PRINT " 9 - END"
INPUT W
IF W = 0 THEN GOTO 5010
IF W = .1 THEN GOTO 5110
IF W = 1 THEN GOTO 5020
IF W = 9 THEN GOTO 5900
IF W = 2 THEN GOTO 5030
IF W = 3 THEN GOTO 5040
5010 PRINT "BYSOLNUI XOD"
INPUT " X, Y "; X, Y
5011 INPUT " A, D "; A, D
   X1 = COS(A / 100 * 6 * 3.141592 / 180) * D + X
   Y1 = SIN(A / 100 * 6 * 3.141592 / 180) * D + Y
PRINT " X = "; INT(X1 + .001); " Y = "; INT(Y1 + .001)
INPUT " Prodolgit 0 - Da, 1 - Net "; prod
IF prod = 0 THEN
     X = X1: Y = Y1
     GOTO 5011
     ELSE
END IF
GOTO 5001
5020 PRINT "OBRATNAIA ZASECHKA PO TREM YGLAM "
INPUT " A - 0, dA - 1 "; A
IF A = 0 THEN
        OBRZASA
        ELSE
        OBRATZASDA
END IF
GOTO 5001

5030 PRINT "<<OGZ>>"
INPUT " Xb, Yb "; Xb, Yb
5031 INPUT " Xc, Yc "; Xc, Yc
   dxc = Xc - Xb
   dyc = Yc - Yb
   Dc = SQR(dxc ^ 2 + dyc ^ 2)
 Ac = ABS(ATN(dyc / (dxc + .1)) / 3.141592 * 30) * 100
 IF dxc > 0 AND dyc > 0 THEN Ygolc = INT(Ac)
 IF dxc < 0 AND dyc > 0 THEN Ygolc = INT(3000 - Ac)
 IF dxc < 0 AND dyc < 0 THEN Ygolc = INT(3000 + Ac)
 IF dxc > 0 AND dyc < 0 THEN Ygolc = INT(6000 - Ac)
 PRINT " D= "; INT(Dc + .001); " A= "; INT(Ygolc + .001)
 INPUT "Prodolgit 0 - Da, 1 - Net "; prod
 IF prod = 0 THEN
        GOTO 5031
        ELSE
        GOTO 5001
 END IF
5040 PRINT "PRAMAIA ZASECHKA"
INPUT " Xp, Yp "; Xp, Yp
INPUT " Xl, Yl "; Xl, Yl
5041 INPUT " PRAV A = "; Apravr
INPUT " LEV A = "; Alevr
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = SQR(dxso ^ 2 + dyso ^ 2)
  aso = ABS(ATN(dyso / (dxso + .1)) / 3.141592 * 30) * 100
  IF dxso > 0 AND dyso > 0 THEN ygol.baz = INT(aso)
  IF dxso < 0 AND dyso > 0 THEN ygol.baz = INT(3000 - aso)
  IF dxso < 0 AND dyso < 0 THEN ygol.baz = INT(3000 + aso)
  IF dxso > 0 AND dyso < 0 THEN ygol.baz = INT(6000 - aso)
  IF Alevr < 1500 AND Apravr > 4500 THEN fir = ABS(Alevr + 6000 - Apravr)
  IF Alevr > 4500 AND Apravr < 1500 THEN fir = ABS(Alevr - (Apravr + 6000))
  IF Alevr > Apravr THEN fir = ABS(Alevr - Apravr)
  IF Alevr < 1500 AND ygol.baz > 4500 THEN
        blevr = ABS(Alevr + 6000 - ygol.baz)
      ELSEIF Alevr > 4500 AND ygol.baz < 1500 THEN
        blevr = ABS(Alevr - (ygol.baz + 6000))
      ELSE
        blevr = ABS(Alevr - ygol.baz)
  END IF
  IF ygol.baz - 3000 < 0 THEN
  ybazp = ygol.baz + 3000
  ELSE
  ybazp = ygol.baz - 3000
  END IF
  IF Apravr < 1500 AND ybazp > 4500 THEN
  bpravr = ABS(Apravr + 6000 - ybazp)
  ELSEIF Apravr > 4500 AND ybazp < 1500 THEN
  bpravr = ABS(Apravr - (ybazp + 6000))
  ELSE
  bpravr = ABS(Apravr - ybazp)
  END IF
  Dlevr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(bpravr / 100 * 6 * 3.141592 / 180))
  Dpravr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(blevr / 100 * 6 * 3.141592 / 180))
  Xcsor = COS(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Xl
  Ycsor = SIN(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Yl
  X = Xcsor: Y = Ycsor
   PRINT " Xc = "; INT(X + .001); " Yc = "; INT(Y + .001)
INPUT " Prodolgit 0 - Da, 1 - Net "; prod
IF prod = 0 THEN
        GOTO 5041
        ELSE
        GOTO 5001
END IF
5110 PRINT " PGZ "
INPUT " X, Y "; X, Y
5111 PRINT " PGZ "
INPUT " A, D "; A, D
   X1 = COS(A / 100 * 6 * 3.141592 / 180) * D + X
   Y1 = SIN(A / 100 * 6 * 3.141592 / 180) * D + Y
PRINT " X="; INT(X1 + .001); " Y="; INT(Y1 + .001)
INPUT " Prodolgit 0 - Da, 1 - Net "; prod
IF prod = 0 THEN
     GOTO 5111
     ELSE
END IF
GOTO 5001

5900 END

'
SUB OBRATZASDA
PRINT "BBOD PYNKTOV SLEVA NA PRAVO"
INPUT "[A]  X, Y "; Xa, Ya
INPUT "[B]  X, Y "; Xb, Yb
INPUT "[C]  X, Y "; Xc, Yc
INPUT " s P na AB dA = "; dAab
INPUT " s P na BC dA = "; dAbc
IF Xa = 0 AND Ya = 0 THEN GOTO 5025
dAab = dAab / 100 * 6: dAbc = dAbc / 100 * 6
   dxab = Xb - Xa
   dyab = Yb - Ya
   Dab = SQR(dxab ^ 2 + dyab ^ 2)
 Ac = ABS(ATN(dyab / dxab) / 3.141592 * 180)
 IF dxab > 0 AND dyab > 0 THEN Ygolab = Ac
 IF dxab < 0 AND dyab > 0 THEN Ygolab = 180 - Ac
 IF dxab < 0 AND dyab < 0 THEN Ygolab = 180 + Ac
 IF dxab > 0 AND dyab < 0 THEN Ygolab = 360 - Ac
   dxbc = Xc - Xb
   dybc = Yc - Yb
   Dbc = SQR(dxbc ^ 2 + dybc ^ 2)
 Ac = ABS(ATN(dybc / dxbc) / 3.141592 * 180)
 IF dxbc > 0 AND dybc > 0 THEN Ygolbc = Ac
 IF dxbc < 0 AND dybc > 0 THEN Ygolbc = 180 - Ac
 IF dxbc < 0 AND dybc < 0 THEN Ygolbc = 180 + Ac
 IF dxbc > 0 AND dybc < 0 THEN Ygolbc = 360 - Ac
 IF Ygolab + 180 > 360 THEN
        Aabc = (Ygolab - 180) - Ygolbc
        ELSE
        Aabc = Ygolab + 180 - Ygolbc
 END IF
POLYSYMMA = (360 - (dAab + dAbc + Aabc)) / 2
O = Dab / SIN(dAab * 3.141592654# / 180)
O = O / (Dbc / SIN(dAbc * 3.141592654# / 180))
O = ATN(O) / 3.141592654# * 180
POLYRAZNOST = TAN(POLYSYMMA * 3.141592654# / 180)
POLYRAZNOST = POLYRAZNOST / TAN((O + 45) * 3.141592654# / 180)
POLYRAZNOST = ATN(POLYRAZNOST) / 3.141592654# * 180
Pab = POLYSYMMA + POLYRAZNOST
Pcb = POLYSYMMA - POLYRAZNOST
Abp = 180 - (Pab + dAab)
Cbp = 180 - (Pcb + dAbc)
Dap = (Dab / SIN(dAab * 3.141592654# / 180)) * SIN(Abp * 3.141592654# / 180)
Dcp = (Dbc / SIN(dAbc * 3.141592654# / 180)) * SIN(Cbp * 3.141592654# / 180)
Aap = Ygolab + Pab
IF Ygolbc + 180 > 360 THEN
        Acp = Ygolbc - 180 - Pcb
        ELSE
        Acp = Ygolbc + 180 - Pcb
 END IF
X1 = Xa + Dap * COS(Aap * 3.141592654# / 180)
Y1 = Ya + Dap * SIN(Aap * 3.141592654# / 180)
X2 = Xc + Dcp * COS(Acp * 3.141592654# / 180)
Y2 = Yc + Dcp * SIN(Acp * 3.141592654# / 180)
Xce = INT((X1 + X2) / 2): Yce = INT((Y1 + Y2) / 2)
   dxc = Xa - Xce
   dyc = Ya - Yce
   Dc = SQR(dxc ^ 2 + dyc ^ 2)
 Ac = ABS(ATN(dyc / dxc) / 3.141592654# * 30) * 100
 IF dxc > 0 AND dyc > 0 THEN Ygolc = INT(Ac)
 IF dxc < 0 AND dyc > 0 THEN Ygolc = INT(3000 - Ac)
 IF dxc < 0 AND dyc < 0 THEN Ygolc = INT(3000 + Ac)
 IF dxc > 0 AND dyc < 0 THEN Ygolc = INT(6000 - Ac)
Dpa = INT(Dc): Ygolpa = INT(Ygolc)
   dxc = Xb - Xce
   dyc = Yb - Yce
   Dc = SQR(dxc ^ 2 + dyc ^ 2)
 Ac = ABS(ATN(dyc / dxc) / 3.141592654# * 30) * 100
 IF dxc > 0 AND dyc > 0 THEN Ygolc = INT(Ac)
 IF dxc < 0 AND dyc > 0 THEN Ygolc = INT(3000 - Ac)
 IF dxc < 0 AND dyc < 0 THEN Ygolc = INT(3000 + Ac)
 IF dxc > 0 AND dyc < 0 THEN Ygolc = INT(6000 - Ac)
Dpb = INT(Dc): Ygolpb = INT(Ygolc)
   dxc = Xc - Xce
   dyc = Yc - Yce
   Dc = SQR(dxc ^ 2 + dyc ^ 2)
 Ac = ABS(ATN(dyc / dxc) / 3.141592654# * 30) * 100
 IF dxc > 0 AND dyc > 0 THEN Ygolc = INT(Ac)
 IF dxc < 0 AND dyc > 0 THEN Ygolc = INT(3000 - Ac)
 IF dxc < 0 AND dyc < 0 THEN Ygolc = INT(3000 + Ac)
 IF dxc > 0 AND dyc < 0 THEN Ygolc = INT(6000 - Ac)
Dpc = INT(Dc): Ygolpc = INT(Ygolc)

PRINT "<< KOORD. PYNKTA P >>"
PRINT " Xp = "; Xce; " Yp = "; Yce
PRINT " s P na A A = "; Ygolpa; " D = "; Dpa
PRINT " s P na B A = "; Ygolpb; " D = "; Dpb
PRINT " s P na C A = "; Ygolpc; " D = "; Dpc
5025
END SUB

'
SUB OBRZASA
PRINT "BBOD PYNKTOV SLEVA NA PRAVO"
INPUT "[A]  X, Y, A "; Xa, Ya, Aa
INPUT "[B]  X, Y, A "; Xb, Yb, Ab
INPUT "[C]  X, Y, A "; Xc, Yc, Ac
Xp = Xa: Yp = Ya: Xl = Xb: Yl = Yb: Alevr = Ab: Apravr = Aa
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = SQR(dxso ^ 2 + dyso ^ 2)
  aso = ABS(ATN(dyso / (dxso + .1)) / 3.141592 * 30) * 100
  IF dxso > 0 AND dyso > 0 THEN ygol.baz = INT(aso)
  IF dxso < 0 AND dyso > 0 THEN ygol.baz = INT(3000 - aso)
  IF dxso < 0 AND dyso < 0 THEN ygol.baz = INT(3000 + aso)
  IF dxso > 0 AND dyso < 0 THEN ygol.baz = INT(6000 - aso)
  IF Alevr < 1500 AND Apravr > 4500 THEN fir = ABS(Alevr + 6000 - Apravr)
  IF Alevr > 4500 AND Apravr < 1500 THEN fir = ABS(Alevr - (Apravr + 6000))
  IF Alevr > Apravr THEN fir = ABS(Alevr - Apravr)
  IF Alevr < 1500 AND ygol.baz > 4500 THEN
        blevr = ABS(Alevr + 6000 - ygol.baz)
      ELSEIF Alevr > 4500 AND ygol.baz < 1500 THEN
        blevr = ABS(Alevr - (ygol.baz + 6000))
      ELSE
        blevr = ABS(Alevr - ygol.baz)
  END IF
  IF ygol.baz - 3000 < 0 THEN
  ybazp = ygol.baz + 3000
  ELSE
  ybazp = ygol.baz - 3000
  END IF
  IF Apravr < 1500 AND ybazp > 4500 THEN
  bpravr = ABS(Apravr + 6000 - ybazp)
  ELSEIF Apravr > 4500 AND ybazp < 1500 THEN
  bpravr = ABS(Apravr - (ybazp + 6000))
  ELSE
  bpravr = ABS(Apravr - ybazp)
  END IF
  Dlevr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(bpravr / 100 * 6 * 3.141592 / 180))
  Dpravr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(blevr / 100 * 6 * 3.141592 / 180))
  Xcsor = COS(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Xl
  Ycsor = SIN(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Yl
  X1 = Xcsor: Y1 = Ycsor

  Xp = Xb: Yp = Yb: Xl = Xc: Yl = Yc: Alevr = Ac: Apravr = Ab
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = SQR(dxso ^ 2 + dyso ^ 2)
  aso = ABS(ATN(dyso / (dxso + .1)) / 3.141592 * 30) * 100
  IF dxso > 0 AND dyso > 0 THEN ygol.baz = INT(aso)
  IF dxso < 0 AND dyso > 0 THEN ygol.baz = INT(3000 - aso)
  IF dxso < 0 AND dyso < 0 THEN ygol.baz = INT(3000 + aso)
  IF dxso > 0 AND dyso < 0 THEN ygol.baz = INT(6000 - aso)
  IF Alevr < 1500 AND Apravr > 4500 THEN fir = ABS(Alevr + 6000 - Apravr)
  IF Alevr > 4500 AND Apravr < 1500 THEN fir = ABS(Alevr - (Apravr + 6000))
  IF Alevr > Apravr THEN fir = ABS(Alevr - Apravr)
  IF Alevr < 1500 AND ygol.baz > 4500 THEN
        blevr = ABS(Alevr + 6000 - ygol.baz)
      ELSEIF Alevr > 4500 AND ygol.baz < 1500 THEN
        blevr = ABS(Alevr - (ygol.baz + 6000))
      ELSE
        blevr = ABS(Alevr - ygol.baz)
  END IF
  IF ygol.baz - 3000 < 0 THEN
  ybazp = ygol.baz + 3000
  ELSE
  ybazp = ygol.baz - 3000
  END IF
  IF Apravr < 1500 AND ybazp > 4500 THEN
  bpravr = ABS(Apravr + 6000 - ybazp)
  ELSEIF Apravr > 4500 AND ybazp < 1500 THEN
  bpravr = ABS(Apravr - (ybazp + 6000))
  ELSE
  bpravr = ABS(Apravr - ybazp)
  END IF
  Dlevr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(bpravr / 100 * 6 * 3.141592 / 180))
  Dpravr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(blevr / 100 * 6 * 3.141592 / 180))
  Xcsor = COS(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Xl
  Ycsor = SIN(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Yl
  X2 = Xcsor: Y2 = Ycsor

  Xp = Xa: Yp = Ya: Xl = Xc: Yl = Yc: Alevr = Ac: Apravr = Aa
  dxso = Xp - Xl: dyso = Yp - Yl
  baz = SQR(dxso ^ 2 + dyso ^ 2)
  aso = ABS(ATN(dyso / (dxso + .1)) / 3.141592 * 30) * 100
  IF dxso > 0 AND dyso > 0 THEN ygol.baz = INT(aso)
  IF dxso < 0 AND dyso > 0 THEN ygol.baz = INT(3000 - aso)
  IF dxso < 0 AND dyso < 0 THEN ygol.baz = INT(3000 + aso)
  IF dxso > 0 AND dyso < 0 THEN ygol.baz = INT(6000 - aso)
  IF Alevr < 1500 AND Apravr > 4500 THEN fir = ABS(Alevr + 6000 - Apravr)
  IF Alevr > 4500 AND Apravr < 1500 THEN fir = ABS(Alevr - (Apravr + 6000))
  IF Alevr > Apravr THEN fir = ABS(Alevr - Apravr)
  IF Alevr < 1500 AND ygol.baz > 4500 THEN
        blevr = ABS(Alevr + 6000 - ygol.baz)
      ELSEIF Alevr > 4500 AND ygol.baz < 1500 THEN
        blevr = ABS(Alevr - (ygol.baz + 6000))
      ELSE
        blevr = ABS(Alevr - ygol.baz)
  END IF
  IF ygol.baz - 3000 < 0 THEN
  ybazp = ygol.baz + 3000
  ELSE
  ybazp = ygol.baz - 3000
  END IF
  IF Apravr < 1500 AND ybazp > 4500 THEN
  bpravr = ABS(Apravr + 6000 - ybazp)
  ELSEIF Apravr > 4500 AND ybazp < 1500 THEN
  bpravr = ABS(Apravr - (ybazp + 6000))
  ELSE
  bpravr = ABS(Apravr - ybazp)
  END IF
  Dlevr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(bpravr / 100 * 6 * 3.141592 / 180))
  Dpravr = ABS(baz / (SIN(fir / 100 * 6 * 3.141592 / 180) + .001) * SIN(blevr / 100 * 6 * 3.141592 / 180))
  Xcsor = COS(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Xl
  Ycsor = SIN(Alevr / 100 * 6 * 3.141592 / 180) * Dlevr + Yl
  X3 = Xcsor: Y3 = Ycsor

  Xsr = INT((X1 + X2 + X3) / 3 + .001): Ysr = INT((Y1 + Y2 + Y3) / 3 + .001)
PRINT " X= "; Xsr; " Y= "; Ysr

END SUB

