DECLARE SUB DPrr (pr AS SINGLE, Ys AS SINGLE)
DECLARE SUB dINT (pr AS SINGLE, dinte AS SINGLE)
1 : CLS
PRINT TAB(25); "INTERVAL YSTYP"
INPUT " 1 Or. Yg-r. po OSN , tus. po REIKE "; osn1, reika1
INPUT " 2 Or. Yg-r. po OSN , tus. po REIKE "; osn2, reika2
INPUT " 3 Or. Yg-r. po OSN , tus. po REIKE "; osn3, reika3
INPUT " 4 Or. Yg-r. po OSN , tus. po REIKE "; osn4, reika4
INPUT " 5 Or. Yg-r. po OSN , tus. po REIKE "; osn5, reika5
INPUT " 6 Or. Yg-r. po OSN , tus. po REIKE "; osn6, reika6
INPUT " Osnovnoe 2e - 0, 3e - 1 "; osn
D1 = 2 / (reika1 + .001) * 1000
D2 = 2 / (reika2 + .001) * 1000
D3 = 2 / (reika3 + .001) * 1000
D4 = 2 / (reika4 + .001) * 1000
D5 = 2 / (reika5 + .001) * 1000
D6 = 2 / (reika6 + .001) * 1000
Ys1 = (osn1 - 4500) * D1 / 1000 * 1.05
IF osn = 1 THEN
    Ys2 = (osn2 - 4500) * D2 / 1000 * 1.05
    ELSE
END IF
IF osn = 0 THEN
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
INPUT " Popravky v YGLOMER Scitat - 1, Net - 0 "; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
        PRINT " Popravka v YGLOMER "
        PRINT " 1 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte1
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
INPUT " 2 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8800
        ELSE
PRINT " 2 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte2
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
END IF
INPUT " 3 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8800
        ELSE
PRINT " 3 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte3
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
END IF
INPUT " 4 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8800
        ELSE
PRINT " 4 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte4
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
END IF
INPUT " 5 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8800
        ELSE
PRINT " 5 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte5
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
END IF
INPUT " 6 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8800
        ELSE
PRINT " 6 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
        pr = 100: dinte = inte6
        dINT pr, dinte
        pr = 200
        dINT pr, dinte
        pr = 300
        dINT pr, dinte
        pr = 400
        dINT pr, dinte
        pr = 500
        dINT pr, dinte
        pr = 600
        dINT pr, dinte
        pr = 700
        dINT pr, dinte
END IF

END IF

8800
'1 Or dP
INPUT " Popravky v PRICEL Scitat - 1, Net - 0 "; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 1 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys1
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
'2 Or
INPUT " 2 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 2 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys2
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
END IF
'3 Or
INPUT " 3 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 3 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys3
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
END IF
'4 Or
INPUT " 4 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 4 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys4
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
END IF
'5 Or
INPUT " 5 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 5 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys5
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
END IF
'6 Or
INPUT " 6 Or scitat - 1, Net - 0"; vvod
IF vvod = 0 THEN
        GOTO 8900
        ELSE
PRINT " 6 Or        Pol, Ymen, Perv, Vtor, Tret, Chet"
pr = 100: Ys = Ys6
DPrr pr, Ys
pr = 200
DPrr pr, Ys
pr = 300
DPrr pr, Ys
pr = 400
DPrr pr, Ys
pr = 500
DPrr pr, Ys
pr = 600
DPrr pr, Ys
pr = 700
DPrr pr, Ys
END IF
END IF
8900
PRINT "*************************************************************************"
PRINT " BBOD 1 - RASCHET, 9 - END "
INPUT q
IF q = 1 THEN GOTO 1

SUB dINT (pr AS SINGLE, dinte AS SINGLE)
OPEN "dt.tus" FOR INPUT AS #1
8200 IF EOF(1) THEN GOTO 8220
INPUT #1, t1, t2, t3, t4, t5, t6, t7
IF t1 = pr THEN GOTO 8230
GOTO 8200
8230 dINTpol = dinte / t2 * .95: dINTymen = dinte / t3 * .95: dINTperv = dinte / t4 * .95: dINTvtor = dinte / t5 * .95: dINTtret = dinte / t6 * .95: dINTchet = dinte / t7 * .95
8220 CLOSE #1
PRINT " PRICEL "; pr; INT(dINTpol); INT(dINTymen); INT(dINTperv); INT(dINTvtor); INT(dINTtret); INT(dINTchet)
END SUB

SUB DPrr (pr AS SINGLE, Ys AS SINGLE)
OPEN "pric.tus" FOR INPUT AS #1
8100 IF EOF(1) THEN GOTO 8120
INPUT #1, t1, t2, t3, t4, t5, t6, t7
IF t1 = pr THEN GOTO 8130
GOTO 8100
8130 dPpol = Ys / t2 * -1: dPymen = Ys / t3 * -1: dPperv = Ys / t4 * -1: dPvtor = Ys / t5 * -1: dPtret = Ys / t6 * -1: dPchet = Ys / t7 * -1
8120 CLOSE #1
PRINT " PRICEL "; pr; INT(dPpol); INT(dPymen); INT(dPperv); INT(dPvtor); INT(dPtret); INT(dPchet)
END SUB

