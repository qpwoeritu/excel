# Porovnanie výpočtov VBA programu s knižnicou pandapower

**Dátum analýzy:** 2026-09-18
**Porovnávané:** VBA moduly v tomto repozitári (HEAD vetvy) vs. [pandapower](https://github.com/e2nIEE/pandapower) (vetva `develop`, e2nIEE/pandapower — Fraunhofer IEE / Univerzita Kassel)

---

## 1. Zhrnutie

| Oblasť | VBA program | pandapower | Verdikt |
|---|---|---|---|
| Metóda load flow | Plný Newton-Raphson, polárne súradnice | Newton-Raphson (PYPOWER formulácia), polárne súradnice | ✅ **Zhoda metodiky** |
| Výkonové rovnice a Jacobián | Štandardné H/N/M/L členy | Identické rovnice | ✅ **Zhoda** |
| Model vedenia | π-článok, B/2 na oboch koncoch | π-článok, B/2 na oboch koncoch | ✅ **Zhoda** |
| Model transformátora | `ys/a²`, `ys`, `−ys/a`, odbočka na primári, Ym na primári | Rovnaká admitančná schéma; predvolený T-model (Ym v strede), voliteľne π | ⚠️ **Malá odchýlka** (umiestnenie Ym) |
| Štítkové parametre trafa (Zk, Rk, Xk, G0, B0) | Vzorce v hárku `transformatory` | Rovnaké vzorce v `create_transformer_from_parameters` | ✅ **Zhoda** |
| Spínače | Vetvy s R = X = 1e-6 Ω | Zlučovanie uzlov (bus fusion), žiadna umelá impedancia | ❌ **Nezhoda** — numerické riziko |
| PV uzly / generátory | ΔV = 0 riadok, bez Q-limitov | Redukovaný Jacobián + `enforce_q_lims` (PV→PQ) | ⚠️ **Čiastočná zhoda** |
| Skraty | Len Ik3'', pevné c = 1,1, bez korekcií | Plný IEC 60909: 3f/2f/1f, max/min, K_T/K_G, ip, Ith | ❌ **Nezhoda** — viď kap. 4 |
| Per-unit sústava | S_base + hladiny kV, Zbase = U²/S | Interné p.u. na rovnakom princípe | ✅ **Zhoda** (s výnimkou v kap. 5.1) |

**Celkový záver:** Load flow je metodicky **v súlade** s pandapower — pri rovnakých vstupných dátach a konvergencii by mali výsledky napätí a tokov súhlasiť v rámci zaokrúhlenia (výstupy sa zaokrúhľujú na 2 desatinné miesta, `modIO.bas:1376-1379`). Skratový výpočet je **zjednodušená verzia IEC 60909** a s pandapower (ani s normou) sa v plnom rozsahu nezhoduje; rozdiely sú kvantifikované v kap. 4 a veľmi pravdepodobne vysvetľujú položku `to_do` #14 („skraty nesedia asi o 20 %").

---

## 2. Metodika porovnania

Porovnával sa VBA kód (11 modulov, jediný vstupný bod `runCALC` v `modMain.bas:29`; režim 1 = load flow, 2 = skraty z `index!G5`) proti zdrojovému kódu a dokumentácii pandapower. pandapower obsahuje moduly: `pf` (load flow), `shortcircuit` (IEC 60909), `opf`, `estimation`, `topology`, `timeseries`, `control`, `contingency`, `protection`, `diagnostic`, `converter`, `std_types` — overené priamo v strome repozitára `e2nIEE/pandapower/pandapower/`.

---

## 3. Load flow — detailné porovnanie

### 3.1 Výkonové rovnice — zhoda

`modNR.bas:46-47`:

```vba
Pcalc(i) = Pcalc(i) + ViVk * (Gik * costh + Bik * sinth)
Qcalc(i) = Qcalc(i) + ViVk * (Gik * sinth - Bik * costh)
```

je presne štandardná polárna formulácia P_i = Σ V_i·V_k·(G_ik·cos θ_ik + B_ik·sin θ_ik), Q_i = Σ V_i·V_k·(G_ik·sin θ_ik − B_ik·cos θ_ik), ktorú používa aj pandapower (PYPOWER `newtonpf`).

### 3.2 Jacobián — zhoda

Diagonálne členy `modNR.bas:140-143`:

```vba
H = -Qcalc(i) - B(i, i) * Vi * Vi        ' dP/dθ
n = Pcalc(i) / Vi + G(i, i) * Vi         ' dP/dV
m = Pcalc(i) - G(i, i) * Vi * Vi         ' dQ/dθ
L = Qcalc(i) / Vi - B(i, i) * Vi         ' dQ/dV
```

aj mimodiagonálne členy (`modNR.bas:164-172`) sa zhodujú so štandardnými deriváciami. Stavový vektor je `[Δθ; ΔV]` (nie ΔV/V) — pandapower/PYPOWER aktualizuje `Vm` tiež priamo, takže formulácie sú ekvivalentné.

Rozdiel v realizácii PV uzlov: VBA ponecháva PV uzol v sústave a vynucuje ΔV = 0 jednotkovým riadkom (`modNR.bas:151-153`, `modNR.bas:89-92`); pandapower PV uzly zo sústavy pre V-neznáme vylučuje (menší Jacobián). **Výsledok je identický**, VBA verzia je len o niečo väčšia sústava.

### 3.3 Modely prvkov

- **Vedenie** — π-článok s jB/2 na oboch koncoch (`modYBus.bas:26-38`): zhodné s pandapower (`line` s `r_ohm_per_km`, `x_ohm_per_km`, `c_nf_per_km`). Pozn.: vetva s R = X = 0 sa v VBA **potichu preskočí** (rozpojenie), pandapower nulovú impedanciu odmietne diagnostikou.
- **Transformátor** — `Y_ii = ys/a² + Ym`, `Y_jj = ys`, `Y_ij = −ys/a` (`modYBus.bas:46-52`). pandapower má rovnakú schému s odbočkou; **rozdiel:** pandapower predvolene používa T-model (magnetizačná vetva v strede série), VBA dáva Ym celú na primár (ekvivalent voľby `trafo_model="pi"` v pandapower). Pre bežné hodnoty i0 ≈ 1 % je rozdiel vo výsledkoch zanedbateľný (< 0,1 % na tokoch). Odbočka je čisto reálny pomer — bez fázového posunu (pandapower podporuje aj `shift_degree`, dôležité len pri prepojených slučkách cez rôzne hodinové uhly).
- **Kompenzácia** — čistá susceptancia na diagonále (`modYBus.bas:96-103`), ekvivalent `shunt` v pandapower. Zhoda.
- **Motory VN** — konštantná admitancia G + jB (`modYBus.bas:107-113`); pandapower modeluje motor v load flow ako záťaž (konšt. P/Q alebo `const_z_percent`). Odlišná filozofia (konšt. Z vs. konšt. P), pri napätiach blízkych 1 p.u. malé rozdiely, pri poklese napätia sa modely rozchádzajú.
- **Generátory** — režim PQ (injekcia) aj EMF (fantómový PV uzol za Ra + jXs, `modYBus.bas:116-123`). pandapower EMF model nemá — `gen` je priamo PV uzol na svorkách. VBA EMF model je fyzikálne bohatší (konštantné budenie), ale **bez Q-limitov**; pandapower má `enforce_q_lims` (automatické PV→PQ pri dosiahnutí Qmin/Qmax) — viď backlog.

### 3.4 Numerika

VBA: hustá Gaussova eliminácia s čiastočným pivotovaním (`modNR.bas:187-265`), O(n³) na iteráciu, plná n×n Y matica. pandapower: riedke matice (scipy sparse). Pri ~400 uzloch to vo VBA funguje, ale kombinácia so spínačmi 1e-6 Ω (admitancie ~1e8 p.u.) výrazne zhoršuje podmienenosť matice — pozorovateľné aj v priebehu konvergencie (prvá iterácia s mismatch ~5e6 p.u.). pandapower tento problém nemá, lebo zopnuté spínače **zlučujú uzly** namiesto vkladania malej impedancie. Súvisí s `to_do` #21 (viacero vypínačov ukazuje identický prúd) — pri 1e-6 Ω je rozdiel napätí na spínači pod hranicou presnosti double pri veľkých admitanciách.

---

## 4. Skraty — porovnanie s IEC 60909 a pandapower (priorita)

### 4.1 Čo počíta VBA

`modShortCircuit.bas` počíta počiatočný rázový skratový prúd Ik3'' vo **všetkých uzloch naraz** z diagonály Z_th = Ysc⁻¹ (komplexná Gauss-Jordan inverzia):

```vba
' modShortCircuit.bas:173-174
Ik_result(i) = (1.1 / Z_th) * (SBase_MVA / (Sqr(3) * Un))
```

Sieťový napájač zo zadaného Ik'' v `uzly!J`:

```vba
' modShortCircuit.bas:130
Z_grid_abs = (1.1 * SBase_MVA) / (Sqr(3) * Un * Ik_slack)
```

Oba vzorce sú **správne odvodené z IEC 60909** (ekvivalentný napäťový zdroj c·Un/√3; Z_Q = c·U_nQ/(√3·I_kQ'') prepočítané do p.u.). Aj zanedbanie priečnych kapacít vedení a pasívnych záťaží v maticiach Ysc je **v súlade s normou** (metóda ekvivalentného zdroja ich pre Ik'' zanedbáva). Základná kostra výpočtu je teda korektná.

### 4.2 Čo chýba oproti IEC 60909 / pandapower

pandapower `shortcircuit.calc_sc` (overené v `pandapower/shortcircuit/calc_sc.py`) podporuje: poruchy **3ph / 2ph / 1ph**, prípady **max / min**, veličiny **ikss, ip, ith**, vetvové výsledky (`branch_results`), poruchovú impedanciu (`r_fault_ohm`, `x_fault_ohm`), toleranciu NN siete (`lv_tol_percent` 6/10 %) a superpozičnú metódu s predporuchovým napätím (`use_pre_fault_voltage`). VBA oproti tomu:

| # | Položka | VBA | IEC 60909 / pandapower | Vplyv na Ik'' |
|---|---|---|---|---|
| 1 | Napäťový faktor c | Pevne 1,1 všade | c_max = 1,10 (VN), 1,05/1,10 (NN); c_min = 0,95/1,00; volí sa podľa hladiny a prípadu max/min | Pre VN max. prípad zhoda; **min. prípad chýba úplne** |
| 2 | Korekcia impedancie trafa K_T | Chýba | K_T = 0,95·c_max/(1 + 0,6·x_T) (IEC 60909-0, čl. 6.3.3) — typicky zníži Z_T o ~4–8 % | Ik'' za trafom **podhodnotený** rádovo o toľko, o koľko Z_T dominuje |
| 3 | Korekcia generátora K_G | Chýba | K_G = (U_n/U_rG)·c_max/(1 + x_d''·sin φ) | Podhodnotenie príspevku generátorov |
| 4 | R/X sieťového napájača | Čisto reaktančný (R = 0), `modShortCircuit.bas:131` | Odporúčané R_Q/X_Q = 0,1 (X_Q = 0,995·Z_Q) | Malý vplyv na \|Ik''\|, väčší na uhol a ip |
| 5 | Motory | Z = j·Xk čisto reaktančné (`modShortCircuit.bas:101-110`) | R_M/X_M = 0,10 (VN motory ≥ 1 MW), 0,15 (< 1 MW), 0,42 (NN skupiny) | Mierne nadhodnotenie príspevku motorov |
| 6 | Nárazový prúd ip | Chýba | ip = κ·√2·Ik'', κ = 1,02 + 0,98·e^(−3R/X) | Chýba veličina pre dimenzovanie |
| 7 | Tepelný prúd Ith, vypínací Ib | Chýba | Ith = Ik''·√(m+n), Ib pre generátorovo blízke skraty | Chýbajú veličiny |
| 8 | 2-fázové a 1-fázové skraty | Chýba | Súmerné zložky (netická, spätná, nulová sústava) | Chýba funkcionalita |
| 9 | Vetvové príspevky skratu | Chýba (len uzlové Ik'') | `branch_results=True` — prúdy vetvami pri skrate | Chýba pre selektivitu ochrán |

### 4.3 Pravdepodobné vysvetlenie ~20 % odchýlky (`to_do` #14)

Samotná kostra vzorcov je správna, takže systematická odchýlka ~20 % s najväčšou pravdepodobnosťou vzniká **kumuláciou** položiek 2, 3 a 5 vyššie (všetky pôsobia rovnakým smerom pri skratoch za transformátormi: chýbajúce K_T aj K_G prúd podhodnocujú), prípadne v kombinácii s:

- **impedanciami zadanými na nesprávnej napäťovej báze** — Xd'' a Xk motorov sa zadávajú v Ω a delia Zbase z-uzla (`modIO.bas`); ak sú štítkové hodnoty vztiahnuté na inú stranu trafa, chyba je kvadrát prevodu,
- porovnávaním s referenčným nástrojom, ktorý používa iný prípad (min vs. max) alebo iné c.

**Odporúčaný postup overenia:** exportovať sieť do pandapower (kap. 6), spustiť `calc_sc(net, fault="3ph", case="max")` a porovnať Ik'' po uzloch. pandapower aplikuje K_T/K_G/R-X korekcie automaticky — rozdiel oproti VBA výsledkom priamo ukáže, ktorá korekcia koľko percent nesie.

### 4.4 Návrh úprav VBA — ✅ IMPLEMENTOVANÉ

Všetkých 5 bodov je implementovaných (podrobnosti a nové vstupy/výstupy: [`skraty_iec60909.md`](skraty_iec60909.md)); vzorce numericky validované proti pandapower 3.5.4 so zhodou 0,00 % pre Ik'' aj ip v max aj min prípade:

1. ✅ c-faktor podľa hladiny + prípad max/min (`index!G6`, `GetVoltageFactorC` v `modUtils.bas`).
2. ✅ K_T pre transformátory (len max prípad, podľa IEC 60909-0 čl. 6.3.3) a K_G pre generátory (nové stĺpce `generatory!T` = Sn_G, `U` = cosφ_r).
3. ✅ R/X pre sieťový napájač (`data!K15`, default 0,1) a motory (R zo stĺpca L, inak 0,1); motory sa v min prípade zanedbávajú.
4. ✅ Nárazový prúd ip metódou B (κ = min(1,15·κ_b, 2,0)) do `uzly!N`.
5. ✅ Vetvové príspevky skratu pre zvolený uzol (`index!G7`) do nového hárku `skrat_vetvy`.

---

## 5. Ostatné zistené odchýlky a poznámky

### 5.1 Zbase vedení z „to" uzla

`modIO.bas:1474-1479` — Zbase vedenia sa počíta z bázy **koncového** uzla (priznaný pozostatok v komentári), zatiaľ čo Ibase pre výstupný prúd toho istého vedenia sa berie z **počiatočného** uzla (`modIO.bas:1357-1360`). Prejaví sa len pri vedení spájajúcom uzly rôznych napäťových hladín bez trafa (čo je samo osebe podozrivá topológia), ale je to vnútorná nekonzistencia. pandapower viaže vedenie vždy na jednu hladinu (kontroluje `vn_kv` oboch uzlov).

### 5.2 Straty vedenia vrátane nabíjacieho prúdu

`modIO.bas:1354`: `Ploss = |I_total|²·R`, kde I_total obsahuje aj kapacitný prúd π-článku (`modIO.bas:1337-1339`). Korektnejšie (a zhodne s pandapower `res_line.pl_mw`) je `Ploss = P_ij + P_ji` alebo `|I_series|²·R`. Pri krátkych kábloch VN je rozdiel malý, pri dlhších vedeniach s väčším B merateľný.

### 5.3 Definícia ΔU %

`modIO.bas:1344`: `ΔU% = (V_from_pu − V_to_pu)·100` — rozdiel p.u. modulov, čo je percento **z bázového napätia**, nie z napätia počiatočného uzla. Je to konzistentná interná konvencia, len ju netreba zamieňať s definíciou úbytku vztiahnutou na Un alebo V_from (pandapower ΔU nepočíta, udáva `vm_pu` oboch koncov).

### 5.4 Vstupná robustnosť

`ParseDouble` pri nečitateľnej bunke potichu vráti 0 (`modUtils.bas`) — v archívnych dátach existuje minimálne jedna poškodená bunka (`"0-,35"`), ktorá sa tak stala nulou bez varovania. pandapower má modul `diagnostic` (kontrola vstupov: nulové impedancie, odpojené prvky, nezmyselné hodnoty, prekryvy) — analóg by mal vo VBA veľkú hodnotu (viď backlog).

### 5.5 Znamienková konvencia

VBA: P/Q v `uzly!F/G` sú **injekcie** (odbery záporné). pandapower: `load.p_mw` je kladný **odber**, `sgen/gen.p_mw` kladná **dodávka**. Pri exporte (kap. 6) treba otočiť znamienko odberov.

---

## 6. Export do pandapower na krížovú validáciu (priorita)

Najrýchlejšia cesta k nezávislému overeniu všetkých výsledkov (load flow aj skratov). Navrhované mapovanie hárkov na prvky pandapower:

| Hárok VBA | Prvok pandapower | Poznámky k prevodu |
|---|---|---|
| `uzly` | `create_bus` (vn_kv = priradená hladina) + `create_load` / `create_sgen` | odbery: P<0 → `load` s p_mw = −P; slack → `create_ext_grid` |
| `uzly!J` (Ik'' slacku) | `ext_grid.s_sc_max_mva = √3·Un·Ik''`, `rx_max = 0.1` | pandapower c-faktor aplikuje interne |
| `vedenia` | `create_line_from_parameters` | R, X v Ω → `length_km=1`, `r_ohm_per_km=R`, `x_ohm_per_km=X`; B [S] → `c_nf_per_km = B/(2π·50)·1e9` |
| `transformatory` | `create_transformer_from_parameters` | `vk_percent=uk`, `vkr_percent = ΔPk[kW]/(10·Sn[MVA])`, `pfe_kw=ΔP0`, `i0_percent`, `tap_side="hv"`, `tap_step_percent=2`, `tap_neutral=9` (podľa vzorca `1+(odb·2−18)/100`) |
| `spinace` | `create_switch` (bus–bus, `closed` podľa stavu) | **žiadna 1e-6 Ω impedancia** — pandapower uzly zlúči |
| `reaktory`, `dif_reaktory` | `create_impedance` (R, X → p.u. na zvolenej Sn) | alternatívne `create_series_reactor_as_impedance` |
| `kompenzácia` | `create_shunt` | `q_mvar = −Un²/X_net` (kapacitná dodávka záporná v konvencii odberu) — overiť znamienko proti VBA `CompB` |
| `motoryVN` | load flow: `create_load` s `const_z_percent=100`; skraty: `create_motor` | VBA model konšt. admitancie ≙ const_z |
| `generatory` | PQ režim: `create_sgen`; EMF režim: `create_gen` (vm_pu = V_ref/Un) + sc parametre (`xdss_pu`) | EMF fantóm sa nemapuje 1:1 — porovnávať svorkové veličiny |

Vzorový skript je v prílohe [`docs/pandapower_export_navrh.py`](pandapower_export_navrh.py) — načíta CSV exporty hárkov, postaví sieť, spustí `runpp` a `calc_sc` a vypíše porovnanie s výsledkami VBA. Očakávané zhody pri korektnom prevode:

- **napätia uzlov:** zhoda < 0,5 % (rozdiely: model spínačov, motorov konšt. Z vs. const_z, zaokrúhlenie na 2 des. miesta),
- **toky a straty vedení:** zhoda < 1 % (pozor na definíciu strát, kap. 5.2),
- **Ik3'':** očakávaný **systematický rozdiel** rádovo v jednotkách až ~20 % tam, kde dominujú transformátory/generátory (chýbajúce K_T/K_G vo VBA) — presne to, čo treba na uzavretie `to_do` #14.

---

## 7. Backlog ďalších funkcionalít podľa vzoru pandapower

Zoradené podľa odhadovaného pomeru prínos/prácnosť pre tento VBA program:

1. **Plný IEC 60909 skrat** (kap. 4.4) — priorita používateľa; rieši `to_do` #14.
2. **Export do pandapower** (kap. 6) — priorita používateľa; trvalý validačný mechanizmus.
3. **Kontrola limitov / loading_percent** — prúdové limity vedení z katalógu v hárku `data` (In už existuje), zaťaženie tráf v % Sn, napäťové medze uzlov, farebné vyznačenie prekročení; ekvivalent `res_line.loading_percent` v pandapower. Rieši `to_do` #20.
4. **Diagnostika vstupov** (vzor `pandapower.diagnostic`) — hlásiť nečitateľné bunky (namiesto tichej 0 v `ParseDouble`), nulové impedancie, duplicitné mená uzlov, prvky odkazujúce na neexistujúce uzly. Lacné, vysoká hodnota.
5. **Q-limity generátorov** (`enforce_q_lims`) — Qmin/Qmax stĺpce v `generatory`, PV→PQ prepnutie v NR slučke.
6. **Zlučovanie uzlov pre spínače** — odstrániť 1e-6 Ω vetvy, zlepšiť podmienenosť; prúd spínačom sa dopočíta z tokov susedných vetiev. Rieši `to_do` #21 a extrémne prvé mismatche.
7. **Prepočet R na prevádzkovú teplotu** (`to_do` #9) — pandapower má TDPF / teplotný koeficient (`alpha`, `temperature_degree_celsius`).
8. **Automatická regulácia odbočiek trafa** (vzor `control.TrafoController`) — iterovať odbočku na udržanie napätia regulovaného uzla v pásme.
9. **DC load flow** — lineárny odhad tokov P bez iterácií (rýchla predbežná kontrola veľkých zmien).
10. **N-1 kontingencie** (vzor `contingency`) — slučka cez výpadky vedení/tráf s kontrolou limitov z bodu 3.
11. **Časové rady / scenáre** (vzor `timeseries`) — tabuľka scenárov záťaží, dávkový beh, obálky výsledkov.
12. Väčšie témy skôr pre pandapower samotný než pre VBA: optimal power flow, odhad stavu (state estimation), nesymetrický 3-fázový chod, ochrany. Tu je racionálnejšie použiť priamo pandapower nad exportom z bodu 2 než ich reimplementovať vo VBA.

---

## 8. Krížové odkazy na `to_do` hárok

| to_do | Položka | Vysvetlenie v tomto reporte |
|---|---|---|
| #9 | prepočet R na 75 °C | backlog bod 7 |
| #11 | či straty zahŕňajú 3 fázy | áno — p.u. sústava je trojfázová (S_base je 3-fáz. výkon); jediné explicitné „3·" je pri motoroch (`modIO.bas`, prepočet z A a Ω), čo je konzistentné |
| #14 | skraty nesedia ~20 % | kap. 4.2–4.3 — chýbajúce K_T/K_G korekcie + kontrola báz impedancií |
| #16 | sumárny prúd uzlov dvojnásobný? | metóda sčíta len kladné (vtekajúce) toky (`modNR.bas:523-684`) — pri uzle, kadiaľ výkon len preteká, je to prietok, nie dvojnásobok; overiť voči `res_bus` pandapower po exporte |
| #20 | limity veličín | backlog bod 3 |
| #21 | vypínače s identickým prúdom | kap. 3.4 — dôsledok 1e-6 Ω modelu; backlog bod 6 |

---

*Všetky odkazy `súbor:riadok` boli overené proti aktuálnemu stavu vetvy. Tvrdenia o pandapower vychádzajú zo zdrojového kódu repozitára e2nIEE/pandapower (moduly `pf`, `shortcircuit/calc_sc.py`, štruktúra balíka) a z normy IEC 60909-0.*
