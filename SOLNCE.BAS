1 : CLS
PRINT TAB(30); " << RASCHET DY NA SOLNCE >> "
        INPUT " SHIROTA : Gradys, Minyt "; Dgr, Dmin
        INPUT " DOLGOTA : Gradys, Minyt "; SHgr, SHmin
        INPUT " N - Zonu = "; zona
        INPUT " DATA : Den, Mes, God "; Den, mes, God
        INPUT " GOD Vusokosnu - 1, Net - 0 "; tipgod
Pi = 3.141592654#
Dmin = (100 / 60 * Dmin) / 100
Dgr = Dgr + Dmin
SHmin = (100 / 60 * SHmin) / 100
SHgr = SHgr + SHmin
L0 = 6 * zona - 3
f = (SHgr - L0) * SIN(Dgr * Pi / 180)
IF God = 2018 THEN
        Kg = .5542
        ELSEIF God = 2019 THEN
        Kg = .312
        ELSE
        Kg = .0698
END IF
IF mes = 1 OR mes = 0 THEN
        OPEN "ianvar.sss" FOR INPUT AS #1
        ELSEIF mes = 2 THEN
        OPEN "fevral.sss" FOR INPUT AS #1
        ELSEIF mes = 3 THEN
        OPEN "mart.sss" FOR INPUT AS #1
        ELSEIF mes = 4 THEN
        OPEN "aprel.sss" FOR INPUT AS #1
        ELSEIF mes = 5 THEN
        OPEN "mai.sss" FOR INPUT AS #1
        ELSEIF mes = 6 THEN
        OPEN "iyn.sss" FOR INPUT AS #1
        ELSEIF mes = 7 THEN
        OPEN "iyl.sss" FOR INPUT AS #1
        ELSEIF mes = 8 THEN
        OPEN "avgyst.sss" FOR INPUT AS #1
        ELSEIF mes = 9 THEN
        OPEN "sentiabr.sss" FOR INPUT AS #1
        ELSEIF mes = 10 THEN
        OPEN "oktiabr.sss" FOR INPUT AS #1
        ELSEIF mes = 11 THEN
        OPEN "noiabr.sss" FOR INPUT AS #1
        ELSE
        OPEN "dekabr.sss" FOR INPUT AS #1
END IF
9011 IF EOF(1) THEN GOTO 9012
INPUT #1, t1, t2, t3, t4
IF tipgod = 0 THEN
        IF Den = t1 THEN
                t0 = t3: b0 = t4: GOTO 9012
                ELSE
                GOTO 9011
        END IF
        ELSE
        IF Den = t2 THEN
                t0 = t3: b0 = t4: GOTO 9012
                ELSE
                GOTO 9011
        END IF
END IF
9012 CLOSE #1
IF tipgod = 0 THEN
        IF mes = 1 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 2 AND Den = 29 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 3 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 4 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 5 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 6 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 7 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 8 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 9 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 10 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 11 AND Den = 31 THEN
                Den1 = 1: mes1 = mes + 1
                ELSEIF mes = 12 AND Den = 31 THEN
                Den1 = 1: mes1 = 1
                ELSE
                Den1 = Den + 1: mes1 = mes
        END IF
        ELSEIF tipgod = 1 THEN
                IF mes = 1 AND Den = 30 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 2 AND Den = 28 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 3 AND Den = 31 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 4 AND Den = 30 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 5 AND Den = 31 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 6 AND Den = 30 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 7 AND Den = 31 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 8 AND Den = 31 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 9 AND Den = 30 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 10 AND Den = 31 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 11 AND Den = 30 THEN
                        Den1 = 1: mes1 = mes + 1
                        ELSEIF mes = 12 AND Den = 31 THEN
                        Den1 = 1: mes1 = 1
                        ELSE
                        Den1 = Den + 1: mes1 = mes
                END IF
        ELSE
END IF
IF mes1 = 1 OR mes1 = 0 THEN
        OPEN "ianvar.sss" FOR INPUT AS #1
        ELSEIF mes1 = 2 THEN
        OPEN "fevral.sss" FOR INPUT AS #1
        ELSEIF mes1 = 3 THEN
        OPEN "mart.sss" FOR INPUT AS #1
        ELSEIF mes1 = 4 THEN
        OPEN "aprel.sss" FOR INPUT AS #1
        ELSEIF mes1 = 5 THEN
        OPEN "mai.sss" FOR INPUT AS #1
        ELSEIF mes1 = 6 THEN
        OPEN "iyn.sss" FOR INPUT AS #1
        ELSEIF mes1 = 7 THEN
        OPEN "iyl.sss" FOR INPUT AS #1
        ELSEIF mes1 = 8 THEN
        OPEN "avgyst.sss" FOR INPUT AS #1
        ELSEIF mes1 = 9 THEN
        OPEN "sentiabr.sss" FOR INPUT AS #1
        ELSEIF mes1 = 10 THEN
        OPEN "oktiabr.sss" FOR INPUT AS #1
        ELSEIF mes1 = 11 THEN
        OPEN "noiabr.sss" FOR INPUT AS #1
        ELSEIF mes1 = 12 THEN
        OPEN "dekabr.sss" FOR INPUT AS #1
        ELSE
END IF
9011.1 IF EOF(1) THEN GOTO 9012.1
INPUT #1, t1, t2, t3, t4
IF tipgod = 0 THEN
        IF Den1 = t1 THEN
                t01 = t3: b01 = t4: GOTO 9012.1
                ELSE
                GOTO 9011.1
        END IF
        ELSE
        IF Den1 = t2 THEN
                t01 = t3: b01 = t4: GOTO 9012.1
                ELSE
                GOTO 9011.1
        END IF
END IF
9012.1 CLOSE #1
PRINT "*******************************************"
9010.1    INPUT " VREMIA (Moskovskoe) : Chas, Min "; Chas, Min
t = t0 + (t01 - t0) * Kg + ((t01 + 360 - t0) / 1440) * (60 * Chas + Min) + SHgr
b = b0 + (b01 - b0) * Kg + ((b01 - b0) / 1440) * (60 * Chas + Min)
IF t >= 360 THEN
        t = t - 360
        ELSE
        t = t
END IF
tga = SIN(t * Pi / 180) / (SIN(Dgr * Pi / 180) * COS(t * Pi / 180) - TAN(b * Pi / 180) * COS(Dgr * Pi / 180))
A = ATN(tga) / Pi * 180
IF t >= 180 AND tga >= 0 THEN
        Asv = A
        ELSEIF t >= 180 AND tga < 0 THEN
        Asv = A + 180
        ELSEIF t < 180 AND tga > 0 THEN
        Asv = A + 180
        ELSEIF t < 180 AND tga < 0 THEN
        Asv = A + 360
        ELSE
END IF
Asv = Asv - f
Asv = CINT((Asv + .001) / 6 * 100)
PRINT "<< DY NA SOLNCE >>"
PRINT " DY = "; Asv
PRINT "*******************************************"
INPUT " Prodolgit? 0 - Da, 1 - Net "; vvod
IF vvod = 0 THEN GOTO 9010.1
PRINT "-------------------------------------------"
INPUT " BBOD 1 - Pomeniat koordinatu, 0 - END "; vvod
IF vvod >= 1 THEN GOTO 1
END

