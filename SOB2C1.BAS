DECLARE SUB DYGINTER (pr AS SINGLE, Dt AS SINGLE, zar AS SINGLE)
DECLARE SUB DPYS (pr AS SINGLE, dxtus AS SINGLE, zar AS SINGLE)
DECLARE SUB IYPOREIK (inte1 AS SINGLE, inte2 AS SINGLE, inte3 AS SINGLE, inte4 AS SINGLE, inte5 AS SINGLE, inte6 AS SINGLE, Ys1 AS SINGLE, Ys2 AS SINGLE, Ys3 AS SINGLE, Ys4 AS SINGLE, Ys5 AS SINGLE, Ys6 AS SINGLE)
10 CLS : PRINT TAB(30); "BLOKNOT SOBA"
11 INPUT " Kto strel. 0 - Batareia, 1 - Vzvod "; ktostr
        IF ktostr = 0 THEN
                PRINT " OTN "
                INPUT " 1or = "; t1
                INPUT " 2or = "; t2
                INPUT " 3or = "; t3
                INPUT " 4or = "; t4
                INPUT " 5or = "; t5
                INPUT " 6or = "; t6
             ELSEIF ktostr = 1 THEN
                INPUT " 1or = "; t1
                INPUT " 2or = "; t2
                INPUT " 3or = "; t3
        END IF
   INPUT " Osnovnoe 3 - 3or, 2 - 2or "; osnov
   INPUT " Vo Vvodit 1 - Da: 0 - Net "; vv
   IF vv = 0 THEN
        GOTO 12
      ELSE
           IF ktostr = 0 THEN
                INPUT "1or Vo="; v01
                INPUT "2or Vo="; v02
                INPUT "3or Vo="; v03
                INPUT "4or Vo="; v04
                INPUT "5or Vo="; v05
                INPUT "6or Vo="; v06
             ELSE
                INPUT "1or Vo="; v01
                INPUT "2or Vo="; v02
                INPUT "3or Vo="; v03
        END IF
   END IF
   INPUT "Vo-Osnovnogo="; v0os
   raznob1 = v0os - v01: raznob2 = v0os - v02: raznob3 = v0os - v03
   raznob4 = v0os - v04: raznob5 = v0os - v05: raznob6 = v0os - v06
12 INPUT " INTERVAL YSTYP Schitat 1 - Da, 0 - Net "; vvod
 IF vvod = 1 THEN
        INPUT " BBOD 0 - VRYCNYU, 1 - Raschet po REIKE "; vvod
        IF vvod = 1 THEN
                IYPOREIK inte1, inte2, inte3, inte4, inte5, inte6, Ys1, Ys2, Ys3, Ys4, Ys5, Ys6
                ELSE
                PRINT " Inrtval Ystyp otnositelno Osnovnogo !!! "
                INPUT " 1or Inte, Ystyp = "; inte1, Ys1
                INPUT " 2or Inte, Ystyp = "; inte2, Ys2
                INPUT " 3or Inte, Ystyp = "; inte3, Ys3
                INPUT " 4or Inte, Ystyp = "; inte4, Ys4
                INPUT " 5or Inte, Ystyp = "; inte5, Ys5
                INPUT " 6or Inte, Ystyp = "; inte6, Ys6
        END IF
        ELSE
 END IF
12.5 INPUT "Zel N="; zeln$
20 INPUT "Snariad 0-OF, 1-3SH, 2-C4, 3-D4"; snar
30 INPUT "Vzrivatel 0-RGM, 1-AR5, 2-V90, 3-DTM75, 4-T90"; vzriv
        IF vzriv = 0 THEN INPUT "Ystanovka vzrivatelia 0-Oskol, 1-Fygas, 2-Zamedlen, 3-Osk.i.Fyg, 4-Osk.Fyg.Zamed"; ystvzr
        IF vzriv <> 0 THEN INPUT "Trybka ="; tryb
40 INPUT "Zariad 0-Poln, 5-Ymensh, 1-Perv, 2-Vtor, 3-Tret, 4-Chetvert"; zar
50 INPUT "Pricel,Skachek = "; Pric, skach
IF Pric = 0 THEN Pric = 1
60 INPUT "Dovorot="; dovor
70 INPUT "Veer 0-Soedinit, 1-Razdelit"; veer
        INPUT "Del. yglom="; veer1
   IF veer = 1 THEN INPUT "Ystanovki 0-Odna, 1-Dve"; yst
   IF yst = 1 THEN INPUT "Skolko dovernyt"; dover
71  pr = Pric
DPYS pr, dxtus, zar
    Pr1 = pr + (Ys1 / (dxtus + .001) * -1): Pr2 = pr + (Ys2 / (dxtus + .001) * -1): Pr3 = pr + (Ys3 / (dxtus + .001) * -1): Pr4 = pr + (Ys4 / (dxtus + .001) * -1): Pr5 = pr + (Ys5 / (dxtus + .001) * -1): Pr6 = pr + (Ys6 / (dxtus + .001) * -1)
111 OPEN "raznoboi" FOR INPUT AS #1
112 IF EOF(1) THEN GOTO 113
INPUT #1, r1, r2, r3, r4, r5, r6, r7
IF r1 <= Pric AND r1 + 50 >= Pric THEN GOTO 112.1
GOTO 112
112.1 Pricr = r1
IF zar = 0 THEN koef = r2
IF zar = 5 THEN koef = r3
IF zar = 1 THEN koef = r4
IF zar = 2 THEN koef = r5
IF zar = 3 THEN koef = r6
IF zar = 4 THEN koef = r7
GOTO 113
113 CLOSE #1
Pricc = Pric + 50
111.1 OPEN "raznoboi" FOR INPUT AS #1
112.111 IF EOF(1) THEN GOTO 113.1
INPUT #1, r1, r2, r3, r4, r5, r6, r7
IF r1 <= Pricc AND r1 + 50 >= Pricc THEN GOTO 112.11
GOTO 112.111
112.11 Pricrr = r1
IF zar = 0 THEN koef1 = r2
IF zar = 5 THEN koef1 = r3
IF zar = 1 THEN koef1 = r4
IF zar = 2 THEN koef1 = r5
IF zar = 3 THEN koef1 = r6
IF zar = 4 THEN koef1 = r7
GOTO 113.1
113.1 CLOSE #1
koef = (koef1 - koef) / 50 * (Pric - Pricr) + koef
dv0Pr1 = raznob1 * koef: dv0Pr2 = raznob2 * koef: dv0Pr3 = raznob3 * koef
dv0Pr4 = raznob4 * koef: dv0Pr5 = raznob5 * koef: dv0Pr6 = raznob6 * koef
Pr1 = Pr1 + dv0Pr1: Pr2 = Pr2 + dv0Pr2: Pr3 = Pr3 + dv0Pr3: Pr4 = Pr4 + dv0Pr4
Pr5 = Pr5 + dv0Pr5: Pr6 = Pr6 + dv0Pr6
    IF t1 + dovor > 6000 THEN
        Ygl1 = (t1 + dovor) - 6000
        ELSEIF t1 + dovor < 0 THEN
        Ygl1 = t1 + 6000 + dovor
        ELSE
        Ygl1 = t1 + dovor
    END IF
    IF t2 + dovor > 6000 THEN
        Ygl2 = (t2 + dovor) - 6000
        ELSEIF t2 + dovor < 0 THEN
        Ygl2 = t2 + 6000 + dovor
        ELSE
        Ygl2 = t2 + dovor
    END IF
    IF t3 + dovor > 6000 THEN
        Ygl3 = (t3 + dovor) - 6000
        ELSEIF t3 + dovor < 0 THEN
        Ygl3 = t3 + 6000 + dovor
        ELSE
        Ygl3 = t3 + dovor
    END IF
    IF t4 + dovor > 6000 THEN
        Ygl4 = (t4 + dovor) - 6000
        ELSEIF t4 + dovor < 0 THEN
        Ygl4 = t4 + 6000 + dovor
        ELSE
        Ygl4 = t4 + dovor
    END IF
    IF t5 + dovor > 6000 THEN
        Ygl5 = (t5 + dovor) - 6000
        ELSEIF t5 + dovor < 0 THEN
        Ygl5 = t5 + 6000 + dovor
        ELSE
        Ygl5 = t5 + dovor
    END IF
    IF t6 + dovor > 6000 THEN
        Ygl6 = (t6 + dovor) - 6000
        ELSEIF t6 + dovor < 0 THEN
        Ygl6 = t6 + 6000 + dovor
        ELSE
        Ygl6 = t6 + dovor
    END IF
    IF ktostr = 1 THEN Pr4 = 0: Ygl4 = 0: Pr5 = 0: Ygl5 = 0: Pr6 = 0: Ygl6 = 0
    IF osnov <> 2 AND veer = 0 THEN
        Ygl1 = Ygl1 - (veer1 * 2)
        ELSEIF osnov <> 2 AND veer = 1 THEN
        Ygl1 = Ygl1 + (veer1 * 2)
        ELSEIF osnov = 2 AND veer = 0 THEN
        Ygl1 = Ygl1 - veer1
        ELSEIF osnov = 2 AND veer = 1 THEN
        Ygl1 = Ygl1 + veer1
    END IF
    IF osnov <> 2 AND veer = 0 THEN
        Ygl2 = Ygl2 - veer1
        ELSEIF osnov <> 2 AND veer = 1 THEN
        Ygl2 = Ygl2 + veer1
        ELSEIF osnov = 2 THEN
        Ygl2 = Ygl2
    END IF
    IF osnov <> 2 THEN
        Ygl3 = Ygl3
        ELSEIF osnov = 2 AND veer = 0 THEN
        Ygl3 = Ygl3 + veer1
        ELSEIF osnov = 2 AND veer = 1 THEN
        Ygl3 = Ygl3 - veer1
    END IF
    IF osnov <> 2 AND veer = 0 THEN
        Ygl4 = Ygl4 + veer1
        ELSEIF osnov <> 2 AND veer = 1 THEN
        Ygl4 = Ygl4 - veer1
        ELSEIF osnov = 2 AND veer = 0 THEN
        Ygl4 = Ygl4 + (veer1 * 2)
        ELSEIF osnov = 2 AND veer = 1 THEN
        Ygl4 = Ygl4 - (veer1 * 2)
    END IF
    IF osnov <> 2 AND veer = 0 THEN
        Ygl5 = Ygl5 + (veer1 * 2)
        ELSEIF osnov <> 2 AND veer = 1 THEN
        Ygl5 = Ygl5 - (veer1 * 2)
        ELSEIF osnov = 2 THEN
        Ygl5 = 0
    END IF
    IF osnov <> 2 AND veer = 0 THEN
        Ygl6 = Ygl6 + (veer1 * 3)
        ELSEIF osnov <> 2 AND veer = 1 THEN
        Ygl6 = Ygl6 - (veer1 * 3)
        ELSEIF osnov = 2 THEN
        Ygl6 = 0
    END IF
DYGINTER pr, Dt, zar
Ygl1 = Ygl1 + (inte1 / (Dt + .001) * .95 * -1)
Ygl2 = Ygl2 + (inte2 / (Dt + .001) * .95 * -1)
Ygl3 = Ygl3 + (inte3 / (Dt + .001) * .95)
Ygl4 = Ygl4 + (inte4 / (Dt + .001) * .95)
Ygl5 = Ygl5 + (inte5 / (Dt + .001) * .95)
Ygl6 = Ygl6 + (inte6 / (Dt + .001) * .95)
    IF yst = 1 THEN Ygl12 = Ygl1 + dover: Ygl22 = Ygl2 + dover: Ygl32 = Ygl3 + dover: Ygl42 = Ygl4 + dover: Ygl52 = Ygl5 + dover: Ygl62 = Ygl6 + dover
    IF skach <> 0 THEN Pr12 = Pr1 + skach: Pr13 = Pr1 - skach: Pr22 = Pr2 + skach: Pr23 = Pr2 - skach
    IF skach <> 0 THEN Pr32 = Pr3 + skach: Pr33 = Pr3 - skach: Pr42 = Pr4 + skach: Pr43 = Pr4 - skach: Pr52 = Pr5 + skach: Pr53 = Pr5 - skach
    IF skach <> 0 THEN Pr62 = Pr6 + skach: Pr63 = Pr6 - skach
IF q = 0 THEN korr.D = 0: korr.Yg = 0
IF q = 1 THEN korr.D = 0: korr.Yg = 0
80 PRINT "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
   PRINT "KOMANDA"; TAB(10); "PRICEL"; TAB(20); "DOVOROT"; TAB(40); "Pric"; TAB(50); "Yglomer"; TAB(60); "Po.Panor."; TAB(70); "KONTROL"
   PRINT "Zel"; zeln$; TAB(10); Pric; TAB(20); "OH="; dovor; TAB(35); "1or"; TAB(40); CINT(Pr1 + .001); TAB(50); CINT(Ygl1 + .001); TAB(60); ot1; TAB(70); kont1
   IF snar = 0 THEN
        PRINT "Sn=OF"; TAB(10);
        ELSEIF snar = 1 THEN
        PRINT "Sn=3SH"; TAB(10);
        ELSEIF snar = 2 THEN
        PRINT "Sn=C4"; TAB(10);
        ELSE
        PRINT "Sn=D4"; TAB(10);
   END IF
PRINT TAB(10); "KORR."; TAB(20); "KORR."; TAB(40); CINT(Pr12 + .001); TAB(50); CINT(Ygl12)
   IF vzriv = 0 THEN
        PRINT "Vz=RGM"; TAB(10);
        ELSEIF vzriv = 1 THEN
        PRINT "Vz=AR-5"; TAB(10);
        ELSEIF vzriv = 2 THEN
        PRINT "Vz=V-90"; TAB(10);
        ELSEIF vzriv = 3 THEN
        PRINT "Vz=DTM75"; TAB(10);
        ELSE
        PRINT "Vz=T-90"; TAB(10);
   END IF
PRINT TAB(10); korr.D; TAB(20); korr.Yg; TAB(40); CINT(Pr13 + .001)
   IF vzriv <> 0 THEN
        PRINT "N="; tryb; TAB(35);
        ELSE
                IF ystvzr = 0 THEN
                        PRINT "YVz=Os"; TAB(35);
                        ELSEIF ystvzr = 1 THEN
                        PRINT "YVz=Fyg"; TAB(35);
                        ELSEIF ystvzr = 2 THEN
                        PRINT "YVz=Za"; TAB(35);
                        ELSEIF ystvzr = 3 THEN
                        PRINT "O-F"; TAB(35);
                        ELSE
                        PRINT "O-F-Z"; TAB(35);
                END IF
   END IF
   PRINT "2or"; TAB(40); CINT(Pr2 + .001); TAB(50); CINT(Ygl2); TAB(60); ot2; TAB(70); kont2
   IF zar = 0 THEN
        PRINT "Z=POLN"; TAB(40);
        ELSEIF zar = 5 THEN
        PRINT "Z=YMEN"; TAB(40);
        ELSEIF zar = 1 THEN
        PRINT "Z=PERV"; TAB(40);
        ELSEIF zar = 2 THEN
        PRINT "Z=VTOR"; TAB(40);
        ELSEIF zar = 3 THEN
        PRINT "Z=TRET"; TAB(40);
        ELSE
        PRINT "Z=CHET"; TAB(40);
   END IF
PRINT CINT(Pr22 + .001); TAB(50); CINT(Ygl22)
   IF veer = 0 THEN
        PRINT "VS="; veer1; TAB(40);
        ELSE
        PRINT "VR="; veer1; TAB(40);
   END IF
PRINT CINT(Pr23 + .001)
   IF yst = 1 THEN
        PRINT "2YY="; dover; TAB(35);
        ELSE
        PRINT "0"; TAB(35);
   END IF
   PRINT "3or"; TAB(40); CINT(Pr3 + .001); TAB(50); CINT(Ygl3); TAB(60); ot3; TAB(70); kont3
   IF skach <> 0 THEN
        PRINT "Sk="; skach; TAB(40);
        ELSE
        PRINT "0"; TAB(40);
   END IF
PRINT CINT(Pr32 + .001); TAB(50); CINT(Ygl32)
PRINT TAB(40); CINT(Pr33 + .001)
PRINT TAB(35); "4or"; TAB(40); CINT(Pr4 + .001); TAB(50); CINT(Ygl4); TAB(60); ot4; TAB(70); kont4
PRINT TAB(40); CINT(Pr42 + .001); TAB(50); CINT(Ygl42)
PRINT TAB(40); CINT(Pr43 + .001)
PRINT TAB(35); "5or"; TAB(40); CINT(Pr5 + .001); TAB(50); CINT(Ygl5); TAB(60); ot5; TAB(70); kont5
PRINT TAB(40); CINT(Pr52 + .001); TAB(50); CINT(Ygl52)
PRINT TAB(40); CINT(Pr53 + .001)
PRINT TAB(35); "6or"; TAB(40); CINT(Pr6 + .001); TAB(50); CINT(Ygl6); TAB(60); ot6; TAB(70); kont6
PRINT TAB(40); CINT(Pr62 + .001); TAB(50); CINT(Ygl62)
PRINT TAB(40); CINT(Pr63 + .001)
85
PRINT "BBOD 0 - Novaia.Zel: 1 - Smena.OP: 3 - KOREKT: 9 - END: "
PRINT "BBOD 4 - KONTROL: 5 - STOI.Zapis.zel: 6 - Plan.Zel:  "
PRINT "BBOD 7 - Ochist.zeli: 10 - Posmotret Zeli 11 - Zapis SOBA: "
90 INPUT q
   IF q = 0 OR q = 1 THEN
         Pr1 = 0: Pr12 = 0: Pr13 = 0: Ygl1 = 0: Ygl12 = 0
         Pr2 = 0: Pr22 = 0: Pr23 = 0: Ygl2 = 0: Ygl22 = 0
         Pr3 = 0: Pr32 = 0: Pr33 = 0: Ygl3 = 0: Ygl32 = 0
         Pr4 = 0: Pr42 = 0: Pr43 = 0: Ygl4 = 0: Ygl42 = 0
         Pr5 = 0: Pr52 = 0: Pr53 = 0: Ygl5 = 0: Ygl52 = 0
         Pr6 = 0: Pr62 = 0: Pr63 = 0: Ygl6 = 0: Ygl62 = 0
      ELSE
   END IF
   IF q = 0 OR q = 1 THEN
        yst = 0: Ygl12 = 0: Ygl22 = 0: Ygl32 = 0: Ygl42 = 0: Ygl52 = 0: Ygl62 = 0
      ELSE
        yst = yst
   END IF
   IF q = 0 THEN GOTO 12.5
   IF q = 9 THEN END
   IF q = 1 THEN GOTO 10
   IF q = 3 THEN GOTO 100
   IF q = 4 THEN GOTO 110
   IF q = 5 THEN GOTO 120
   IF q = 7 THEN GOTO 130
   IF q = 6 THEN GOTO 140
   IF q = 10 THEN GOTO 150
   IF q = 11 THEN GOTO 80
GOTO 85
100 INPUT "0-Vsem: 1-Kagdomy:"; vska
        IF vska = 0 THEN
                INPUT "KORR. Pric="; dPr
                INPUT "KORR. Yglom="; dYgl
              ELSE
                IF ktostr = 0 THEN
                        INPUT " 1or.KORR. Pric, Yglom = "; dPr1, dYgl1
                        INPUT " 2or.KORR. Pric, Yglom = "; dPr2, dYgl2
                        INPUT " 3or.KORR. Pric, Yglom = "; dPr3, dYgl3
                        INPUT " 4or.KORR. Pric, Yglom = "; dPr4, dYgl4
                        INPUT " 5or.KORR. Pric, Yglom = "; dPr5, dYgl5
                        INPUT " 6or.KORR. Pric, Yglom = "; dPr6, dYgl6
                   ELSE
                        INPUT " 1or.KORR. Pric, Yglom = "; dPr1, dYgl1
                        INPUT " 2or.KORR. Pric, Yglom = "; dPr2, dYgl2
                        INPUT " 3or.KORR. Pric, Yglom = "; dPr3, dYgl3

               END IF
        END IF
IF vska = 0 THEN
        Pr1 = dPr + Pr1: Ygl1 = dYgl + Ygl1: Pr2 = dPr + Pr2: Ygl2 = dYgl + Ygl2
        Pr3 = dPr + Pr3: Ygl3 = dYgl + Ygl3: Pr4 = dPr + Pr4: Ygl4 = dYgl + Ygl4
        Pr5 = dPr + Pr5: Ygl5 = dYgl + Ygl5: Pr6 = dPr + Pr6: Ygl6 = dYgl + Ygl6
   
        Pr12 = dPr + Pr12: Ygl12 = dYgl + Ygl12: Pr22 = dPr + Pr22: Ygl22 = dYgl + Ygl22
        Pr32 = dPr + Pr32: Ygl32 = dYgl + Ygl32: Pr42 = dPr + Pr42: Ygl42 = dYgl + Ygl42
        Pr52 = dPr + Pr52: Ygl52 = dYgl + Ygl52: Pr62 = dPr + Pr62: Ygl62 = dYgl + Ygl62
   
        Pr13 = dPr + Pr13: Pr23 = dPr + Pr23: Pr33 = dPr + Pr33
        Pr43 = dPr + Pr43: Pr15 = dPr + Pr53: Pr63 = dPr + Pr63
    ELSE
        Pr1 = dPr1 + Pr1: Ygl1 = dYgl1 + Ygl1: Pr2 = dPr2 + Pr2: Ygl2 = dYgl2 + Ygl2
        Pr3 = dPr3 + Pr3: Ygl3 = dYgl3 + Ygl3: Pr4 = dPr4 + Pr4: Ygl4 = dYgl4 + Ygl4
        Pr5 = dPr5 + Pr5: Ygl5 = dYgl5 + Ygl5: Pr6 = dPr6 + Pr6: Ygl6 = dYgl6 + Ygl6

        Pr12 = dPr1 + Pr12: Ygl12 = dYgl1 + Ygl12: Pr22 = dPr2 + Pr22: Ygl22 = dYgl2 + Ygl22
        Pr32 = dPr3 + Pr32: Ygl32 = dYgl3 + Ygl32: Pr42 = dPr4 + Pr42: Ygl42 = dYgl4 + Ygl42
        Pr52 = dPr5 + Pr52: Ygl52 = dYgl5 + Ygl52: Pr62 = dPr6 + Pr62: Ygl62 = dYgl6 + Ygl62

        Pr13 = dPr1 + Pr13: Pr23 = dPr2 + Pr23: Pr33 = dPr3 + Pr33
        Pr43 = dPr4 + Pr43: Pr53 = dPr5 + Pr53: Pr63 = dPr6 + Pr63
END IF
IF skach = 0 THEN Pr12 = 0: Pr13 = 0: Pr22 = 0: Pr23 = 0: Pr32 = 0: Pr33 = 0: Pr42 = 0: Pr43 = 0
IF skach = 0 THEN Pr52 = 0: Pr53 = 0: Pr62 = 0: Pr63 = 0
IF yst = 0 THEN Ygl12 = 0: Ygl22 = 0: Ygl32 = 0: Ygl42 = 0: Ygl52 = 0: Ygl62 = 0

korr.D = korr.D + dPr: korr.Yg = korr.Yg + dYgl

GOTO 80
110 IF ktostr = 0 THEN
        INPUT " Otm.po.1or = "; ot1
        INPUT " Otm.po.2or = "; ot2
        INPUT " Otm.po.3or = "; ot3
        INPUT " Otm.po.4or = "; ot4
        INPUT " Otm.po.5or = "; ot5
        INPUT " Otm.po.6or = "; ot6
      ELSE
        INPUT " Otm.po.1or = "; ot1
        INPUT " Otm.po.2or = "; ot2
        INPUT " Otm.po.3or = "; ot3
    END IF
IF yst = 0 THEN
        dovsym1 = Ygl1 - t1
   ELSE
        dovsym1 = Ygl12 - t1
END IF
IF yst = 0 THEN
        dovsym2 = Ygl2 - t2
   ELSE
        dovsym2 = Ygl22 - t2
END IF
IF yst = 0 THEN
        dovsym3 = Ygl3 - t3
   ELSE
        dovsym3 = Ygl32 - t3
END IF
IF yst = 0 THEN
        dovsym4 = Ygl4 - t4
   ELSE
        dovsym4 = Ygl42 - t4
END IF
IF yst = 0 THEN
        dovsym5 = Ygl5 - t5
   ELSE
        dovsym5 = Ygl52 - t5
END IF
IF yst = 0 THEN
        dovsym6 = Ygl6 - t6
   ELSE
        dovsym6 = Ygl62 - t6
END IF
    kont1 = ot1 + dovsym1: kont2 = ot2 + dovsym2: kont3 = ot3 + dovsym3
    kont4 = ot4 + dovsym4: kont5 = ot5 + dovsym5: kont6 = ot6 + dovsym6
IF q = 0 THEN kont1 = 0: kont2 = 0: kont3 = 0: kont4 = 0: kont5 = 0: kont6 = 0

GOTO 80
120
Pric = Pric + korr.D: dovor = dovor + korr.Yg
z1$ = zeln$: z2 = Pric: z3 = dovor
z4 = Pr1: z5 = Pr12: z6 = Pr13: z7 = Ygl1: z8 = Ygl12
z9 = Pr2: z10 = Pr22: z11 = Pr23: z12 = Ygl2: z13 = Yg22
z14 = Pr3: z15 = Pr32: z16 = Pr33: z17 = Ygl3: z18 = Yg32
z19 = Pr4: z20 = Pr42: z21 = Pr43: z22 = Ygl4: z23 = Yg42
z24 = Pr5: z25 = Pr52: z26 = Pr53: z27 = Ygl5: z28 = Yg52
z29 = Pr6: z30 = Pr62: z31 = Pr63: z32 = Ygl6: z33 = Yg62
z34 = zar: z35 = snar: z36 = vzriv: z37 = ystvzr: z38 = tryb: z39 = skach: z40 = veer: z41 = veer1
z42 = yst: z43 = dover
OPEN "planzeli" FOR APPEND AS #1
WRITE #1, z1$, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15, z16, z17, z18, z19, z20, z21, z22, z23, z24, z25, z26, z27, z28, z29, z30, z31, z32, z33, z34, z35, z36, z37, z38, z39, z40, z41, z42, z43
CLOSE #1
GOTO 80
130
OPEN "planzeli" FOR OUTPUT AS #1
CLOSE #1
GOTO 80
140 INPUT "N-Zeli"; zeln$
    OPEN "planzeli" FOR INPUT AS #1
141 IF EOF(1) THEN GOTO 142
    INPUT #1, z1$, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15, z16, z17, z18, z19, z20, z21, z22, z23, z24, z25, z26, z27, z28, z29, z30, z31, z32, z33, z34, z35, z36, z37, z38, z39, z40, z41, z42, z43
    IF z1$ = zeln$ THEN
            Pric = z2: dovor = z3
            Pr1 = z4: Pr12 = z5: Pr13 = z6: Ygl1 = z7: Ygl12 = z8
            Pr2 = z9: Pr22 = z10: Pr23 = z11: Ygl2 = z12: Ygl22 = z13
            Pr3 = z14: Pr32 = z15: Pr33 = z16: Ygl3 = z17: Ygl32 = z18
            Pr4 = z19: Pr42 = z20: Pr43 = z21: Ygl4 = z22: Ygl42 = z23
            Pr5 = z24: Pr52 = z25: Pr53 = z26: Ygl5 = z27: Ygl52 = z28
            Pr6 = z29: Pr62 = z30: Pr63 = z31: Ygl6 = z32: Ygl62 = z33
            zar = z34: snar = z35: vzriv = z36: ystvzr = z37: tryb = z38: skach = z39: veer = z40: veer1 = z41
            yst = z42: dover = z43
            GOTO 142
         ELSE
            GOTO 141
    END IF
142 CLOSE #1
GOTO 80
150 PRINT TAB(25); "<< PLANOVUE ZELI >>"
OPEN "planzeli" FOR INPUT AS #1
DO WHILE NOT EOF(1)
FOR i = 1 TO 20
IF EOF(1) THEN GOTO 155
INPUT #1, z1$, z2, z3, z4, z5, z6, z7, z8, z9, z10, z11, z12, z13, z14, z15, z16, z17, z18, z19, z20, z21, z22, z23, z24, z25, z26, z27, z28, z29, z30, z31, z32, z33, z34, z35, z36, z37, z38, z39, z40, z41, z42, z43
PRINT z1$
NEXT i
SLEEP
LOOP
155 CLOSE #1
GOTO 85

SUB DPYS (pr AS SINGLE, dxtus AS SINGLE, zar AS SINGLE)
 OPEN "pric.tus" FOR INPUT AS #1
200 IF EOF(1) THEN GOTO 250
INPUT #1, t1, t2, t3, t4, t5, t6, t7
IF pr <= t1 AND pr + 50 > t1 THEN
        IF zar = 0 THEN dxtus = t2
        IF zar >= 5 THEN dxtus = t3
        IF zar = 1 THEN dxtus = t4
        IF zar = 2 THEN dxtus = t5
        IF zar = 3 THEN dxtus = t6
        IF zar = 4 THEN dxtus = t7
        GOTO 250
        ELSE
        GOTO 200
END IF
250 CLOSE #1
END SUB

SUB DYGINTER (pr AS SINGLE, Dt AS SINGLE, zar AS SINGLE)
 OPEN "dt.tus" FOR INPUT AS #1
300 IF EOF(1) THEN GOTO 350
INPUT #1, t1, t2, t3, t4, t5, t6, t7
IF pr <= t1 AND pr + 50 > t1 THEN
        IF zar = 0 THEN Dt = t2
        IF zar >= 5 THEN Dt = t3
        IF zar = 1 THEN Dt = t4
        IF zar = 2 THEN Dt = t5
        IF zar = 3 THEN Dt = t6
        IF zar = 4 THEN Dt = t7
        GOTO 350
        ELSE
        GOTO 300
END IF
350 CLOSE #1

END SUB

SUB IYPOREIK (inte1 AS SINGLE, inte2 AS SINGLE, inte3 AS SINGLE, inte4 AS SINGLE, inte5 AS SINGLE, inte6 AS SINGLE, Ys1 AS SINGLE, Ys2 AS SINGLE, Ys3 AS SINGLE, Ys4 AS SINGLE, Ys5 AS SINGLE, Ys6 AS SINGLE)
PRINT TAB(25); "INTERVAL YSTYP"
INPUT " 1 Or. Yg-r. po OSN , tus. po REIKE "; osn1, reika1
INPUT " 2 Or. Yg-r. po OSN , tus. po REIKE "; osn2, reika2
INPUT " 3 Or. Yg-r. po OSN , tus. po REIKE "; osn3, reika3
INPUT " 4 Or. Yg-r. po OSN , tus. po REIKE "; osn4, reika4
INPUT " 5 Or. Yg-r. po OSN , tus. po REIKE "; osn5, reika5
INPUT " 6 Or. Yg-r. po OSN , tus. po REIKE "; osn6, reika6
INPUT " Osnovnoe 2e - 2, 3e - 3 "; osn
D1 = 2 / (reika1 + .001) * 1000
D2 = 2 / (reika2 + .001) * 1000
D3 = 2 / (reika3 + .001) * 1000
D4 = 2 / (reika4 + .001) * 1000
D5 = 2 / (reika5 + .001) * 1000
D6 = 2 / (reika6 + .001) * 1000
Ys1 = (osn1 - 4500) * D1 / 1000 * 1.05
IF osn = 3 THEN
    Ys2 = (osn2 - 4500) * D2 / 1000 * 1.05
    ELSE
END IF
IF osn = 2 OR osn = 0 THEN
    Ys3 = (osn3 - 1500) * D3 / 1000 * 1.05
    ELSE
END IF
Ys4 = (osn4 - 1500) * D4 / 1000 * 1.05
Ys5 = (osn5 - 1500) * D5 / 1000 * 1.05
Ys6 = (osn6 - 1500) * D6 / 1000 * 1.05
inte1 = SQR(ABS(D1 ^ 2 - Ys1 ^ 2))
inte2 = SQR(ABS(D2 ^ 2 - Ys2 ^ 2))
inte3 = SQR(ABS(D3 ^ 2 - Ys3 ^ 2))
inte4 = SQR(ABS(D4 ^ 2 - Ys4 ^ 2))
inte5 = SQR(ABS(D5 ^ 2 - Ys5 ^ 2))
inte6 = SQR(ABS(D6 ^ 2 - Ys6 ^ 2))
IF osn1 = 0 THEN inte1 = 0: Ys1 = 0
IF osn2 = 0 THEN inte2 = 0: Ys2 = 0
IF osn3 = 0 THEN inte3 = 0: Ys3 = 0
IF osn4 = 0 THEN inte4 = 0: Ys4 = 0
IF osn5 = 0 THEN inte5 = 0: Ys5 = 0
IF osn6 = 0 THEN inte6 = 0: Ys6 = 0
PRINT " 1 Or INTERVAL = "; INT(inte1 + .001), " YSTYP = "; INT(Ys1 + .001)
PRINT " 2 Or INTERVAL = "; INT(inte2 + .001), " YSTYP = "; INT(Ys2 + .001)
PRINT " 3 Or INTERVAL = "; INT(inte3 + .001), " YSTYP = "; INT(Ys3 + .001)
PRINT " 4 Or INTERVAL = "; INT(inte4 + .001), " YSTYP = "; INT(Ys4 + .001)
PRINT " 5 Or INTERVAL = "; INT(inte5 + .001), " YSTYP = "; INT(Ys5 + .001)
PRINT " 6 Or INTERVAL = "; INT(inte6 + .001), " YSTYP = "; INT(Ys6 + .001)

END SUB

