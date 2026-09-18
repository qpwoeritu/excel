# Návrh: redukcia uzlov pre spínače (bus fusion podľa pandapower)

**Stav: NÁVRH NA ODSÚHLASENIE — zatiaľ neimplementované.**

## 1. Motivácia

Spínače sa dnes modelujú ako vetvy s R = X = 1e-6 Ω (`modYBus.bas`, `modShortCircuit.bas`), čo spôsobuje:

- **admitancie ~1e8 p.u.** v Y matici → zlá podmienenosť, obrovský prvý mismatch NR (~5e6 p.u.),
- **neurčité rozdelenie prúdov** medzi paralelnými spínačmi — delí sa nepatrný rozdiel napätí nepatrnou impedanciou (`to_do` #21: „viacero vypínačov ukazuje identický prúd"),
- zbytočne **veľkú maticu**: ~388 spínačov drží ~388 uzlov navyše; pri hustej O(n³) algebre je to hlavný výkonový faktor.

pandapower tento problém rieši **zlúčením uzlov** (bus fusion) — návrh preberá jej mechanizmus.

## 2. Ako to robí pandapower (overené v zdrojáku 3.5.4)

| Mechanizmus | pandapower | Zdroj |
|---|---|---|
| Kritérium fúzie | zopnutý bus-bus spínač **a** `z_ohm ≤ 0` | `build_bus.py:69`, `:171` (`net._fused_bb_switches`) |
| Spínač s impedanciou | `z_ohm > 0` → modeluje sa ako impedančná vetva (nefúzuje sa) | `build_bus.py:266` (`_impedance_bb_switches`) |
| Algoritmus | union-find (`DisjointSet`) nad dvojicami uzlov zopnutých spínačov | `build_bus.py:133`, `:205` |
| Reprezentant supernodu | pri fúzii sa musí zachovať PV/slack uzol; aktívny (s injekciou) má prednosť pred nulovým | `build_bus.py:190-220` |
| Výsledky uzlov | všetky uzly supernodu dostanú napätie/výsledky spoločného ppc-uzla (bus lookup) | `results_bus` |
| Prúdy spínačov | **fúzované spínače prúd nemajú** (`res_switch.i_ka` = NaN; počíta sa len pre impedančné) | `results_branch.py:664-688` |

## 3. Návrh pre VBA program

### 3.1 Kritérium fúzie (ekvivalent z_ohm)

Zopnutý spínač (`SwStatus > 0`) sa **fúzuje**, ak `|Z| = √(R²+X²) ≤ prah`; inak ostáva impedančnou vetvou ako doteraz (ekvivalent pandapower `z_ohm > 0`).

- Nový vstup **`data!K16` = prah fúzie [Ω]**, prázdne → **0,001 Ω** (dnešné 1e-6 Ω spínače spadnú pod prah; zámerne zadaná kontaktná/prechodová impedancia nad prah ostane vetvou).
- Rozopnutý spínač: bez zmeny (žiadna hrana, prúd 0).

### 3.2 Union-find a supernody (`modReduce.bas` — nový modul)

- `BuildNodeReduction(...)`: union-find nad dvojicami (SwFrom, SwTo) fúzovaných spínačov → mapa `BusToNode(1..nBuses)` a počet elektrických uzlov `nNodes`.
- **Reprezentant supernodu** (priorita podľa pandapower): slack > PV > uzol s nenulovou injekciou > prvý v poradí.
- **Kontroly konzistencie (tvrdé chyby s menami uzlov):**
  - všetky uzly supernodu musia mať rovnakú napäťovú hladinu (fúzia cez hladiny = dátová chyba),
  - max. 1 slack v supernode; PV uzly s rôznym V_ref v jednom supernode = chyba,
- **Agregácia na node úrovni:** `Pspec/Qspec` = súčet cez uzly supernodu; typ = podľa reprezentanta; `Vmag` štart = reprezentant; `BusBaseKV` = spoločná hladina.

### 3.3 Zapojenie do výpočtov — bez zásahu do jadier

Kľúčový princíp: **redukcia sa robí v `modMain` prekódovaním vstupov, jadrá (BuildYBus, RunNRPhase, BuildShortCircuitMatrix, SolveShortCircuit) bežia nezmenené** — len dostanú node-level polia:

1. konce prvkov: `FromNode(k) = BusToNode(FromBus(k))` atď. (vedenia, trafá, reaktory, dif. reaktory, nefúzované spínače, kompenzácie, motory, generátory),
2. prvok s oboma koncami v tom istom supernode sa v Y matematicky vyruší (`+ys, +ys, −ys, −ys` na tej istej diagonále) a jeho prúd vyjde 0 — správne (je vyskratovaný prípojnicou); do hárku `report` sa zapíše upozornenie,
3. fúzované spínače sa do Y **nevkladajú vôbec**,
4. **expanzia výsledkov:** `V_bus(i) = V_node(BusToNode(i))`, `Ik_bus(i) = Ik_node(...)`, `ip` rovnako → **celý existujúci post-processing (zápisy vedení/tráf/reaktorov, uzly!H:N, SLD) beží bez zmeny** na bus-level poliach.

### 3.4 Prúdy fúzovaných spínačov — nad rámec pandapower

pandapower prúdy fúzovaných spínačov nepočíta; pre nás sú výstupom (`spinace!N`, SLD tagy `Q_`). Návrh: **KCL rozklad vo vnútri supernodu** (`SolveSwitchCurrents` v `modReduce.bas`):

- graf supernodu: vrcholy = pôvodné uzly, hrany = fúzované spínače,
- injekcia každého vrchola = súčet komplexných prúdov nespínačových prvkov pripojených do daného uzla (z už spočítaných výsledkov: vedenia, trafá, reaktory, kompenzácie, motory, generátory, záťaže `conj(S/V)`),
- **strom** (typický prípad – rad odpojovač+vypínač v poli): prúd každej hrany = jednoznačný rezový súčet injekcií (DFS) → deterministické prúdy, rieši `to_do` #21,
- **slučka** (paralelné spínače / zopnutý kruh v prípojnici): prúd nie je fyzikálne jednoznačný → lokálne rozdelenie malej sústavy podľa zadaných R/X spínačov (len v rámci supernodu — numericky bezpečné) + varovanie do `report` so zoznamom dotknutých spínačov,
- v skratovom móde rovnaký rozklad s prúdmi príspevkov → report `skrat_vetvy` pri poruche na prípojnici ukáže rozpad prúdu po jednotlivých poliach **cez ich spínače**.

### 3.5 Skraty a priame príspevky

- Redukcia platí zhodne pre `BuildShortCircuitMatrix` (menšia Ysc, rovnaké vzorce).
- `uzly!J/N`: každý uzol dostane Ik''/ip svojho supernodu (uzly tej istej prípojnice prirodzene rovnaké — dnes to platí len približne cez 1e-6 Ω).
- Priame príspevky (`skrat_vetvy`): incidencia sa testuje na supernode (`BusToNode(i) = faultNode`) → report ukáže **všetky polia prípojnice**, nie len jeden spínač — presne rozpad, ktorý požadujete.

### 3.6 Prechodový režim

- **`index!G8` = režim spínačov:** prázdne/`redukcia` (nový, default) / `impedancia` (doterajšie správanie) — umožní A/B porovnanie na reálnej sieti počas validácie; po overení možno legacy vetvu odstrániť.

## 4. Dotknuté súbory

| Súbor | Zmena |
|---|---|
| `modReduce.bas` (nový) | union-find, kontroly, agregácia, expanzia, `SolveSwitchCurrents` (~300–400 riadkov) |
| `modMain.bas` | po FÁZE 1: `BuildNodeReduction`; prekódovanie polí na node-level; expanzia výsledkov; volanie `SolveSwitchCurrents`; režim z `index!G8` |
| `modIO.bas` | čítanie `data!K16` (prah) a `index!G8` (režim); upozornenia do `report` |
| `modNR.bas` | `WriteSwitchCurrents` nahradené výstupom z `SolveSwitchCurrents` (v režime redukcia) |
| `modShortCircuit.bas` | bez zmeny jadra; incidencia priamych príspevkov cez `BusToNode` |
| `modYBus.bas`, `modTopology.bas` | bez zmeny (dostávajú node-level polia; BFS topológie ostáva na pôvodných uzloch) |

## 5. Očakávané prínosy

- Y matica bez členov ~1e8 p.u. → normálna podmienenosť, prvý NR mismatch v bežných rádoch, pravdepodobne menej iterácií,
- redukcia ~400 uzlov na odhadom ~50–150 elektrických uzlov (pri 388 prevažne zopnutých spínačoch) → **rádovo 10–60× rýchlejšia** hustá algebra (O(n³)) v NR aj v inverzii skratovej matice, menšie výpisy `Y_matica`/`SC_matica`,
- deterministické prúdy spínačov (strom) — uzatvára `to_do` #21,
- lepší report priamych príspevkov (celá prípojnica ako jeden uzol poruchy).

## 6. Validácia (pred nasadením)

1. **pandapower prototyp** (rozšírenie `proto_*.py`): sieť s bus-bus spínačmi — pandapower fúzuje automaticky; porovnať napätia, Ik'' a ip s VBA-replikou redukcie (očakávaná zhoda ako doteraz ~0 %),
2. **A/B na reálnom zošite:** režim `impedancia` vs `redukcia` — napätia a Ik'' sa smú líšiť len zanedbateľne (<0,01 %), časy fáz klesnú; prúdy spínačov porovnať s rezovými súčtami ručne na 2–3 poliach,
3. kontrolné súčty: v každom supernode Σ injekcií = 0 (KCL) — automatická kontrola s toleranciou, porušenie → varovanie.

## 7. Etapy implementácie

- **E1:** redukcia + load flow + skraty (Ik''/ip, priame príspevky) + A/B režim — jadro prínosu,
- **E2:** prúdy fúzovaných spínačov cez KCL (strom + slučky s varovaním),
- **E3 (po validácii):** default `redukcia`, prípadné odstránenie legacy vetvy a 1e-6 Ω hodnôt z hárku `spinace`.
