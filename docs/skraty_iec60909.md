# Skratové výpočty podľa IEC 60909 — popis implementácie

Implementácia odporúčaní z kap. 4.4 reportu [`porovnanie_pandapower.md`](porovnanie_pandapower.md). Vzorce boli pred zápisom do VBA numericky validované proti pandapower 3.5.4 (`calc_sc`, `kappa_method="B"`, `topology="meshed"`) na testovacej sieti napájač–trafo–vedenie–motor–generátor: **zhoda Ik'' aj ip 0,00 %** v max aj min prípade.

## Nové vstupy

| Bunka/stĺpec | Význam | Prázdne → predvoľba |
|---|---|---|
| `index!G6` | prípad výpočtu: `max` / `min` | `max` |
| `index!G7` | meno uzla poruchy pre vetvové príspevky | vetvy sa nepočítajú |
| `data!K13` | Ik''max sieťového napájača [kA] | fallback: `uzly!J` v riadku slacku (pôvodné správanie; pozor, stĺpec J sa každým behom prepisuje výsledkami) |
| `data!K14` | Ik''min sieťového napájača [kA] | pri prípade `min`: použije sa K13 (varovanie do hárku `report`) |
| `data!K15` | pomer R/X napájača | 0,1 (IEC odporúčanie) |
| `generatory!T` | Sn_G [MVA] (menovitý zdanlivý výkon) | K_G = 1 (varovanie pri čiastočne vyplnených údajoch) |
| `generatory!U` | cosφ_r [-] (menovitý účinník) | K_G = 1 |

## Nové výstupy

- **`uzly!N`** — nárazový skratový prúd **ip [kA]** pre každý uzol (hlavičku bunky N2 si doplní používateľ; kód hlavičky na vstupné karty nezapisuje).
- **hárok `skrat_vetvy`** — pri vyplnenom `index!G7`: prúdy všetkých vetiev (vedenia, trafá, spínače, reaktory, dif. reaktory) a príspevky motorov a generátorov pri skrate vo zvolenom uzle, zoradené zostupne; hlavička s uzlom, prípadom, c, Ik'' a ip.
- **SLD tag** — nová premenná uzla `Ikp` (napr. `N_R100_Ikp_R`) vypíše `ip= 12,3 kA`. (Názov `Ip` je obsadený primárnym prúdom trafa.)
- **hárok `report`** — varovania výpočtu (chýbajúce údaje pre K_T/K_G, fallback Ik''min a pod.), riadky s prefixom „Varovanie:".

## Použité vzorce (odkazy na IEC 60909-0:2016)

| Veličina | Vzorec | Poznámka |
|---|---|---|
| Napäťový faktor c (tab. 1) | Un ≤ 1 kV: cmax 1,05 / cmin 0,95; Un > 1 kV: cmax 1,10 / cmin 1,00 | `GetVoltageFactorC` v `modUtils.bas`; volí sa podľa hladiny uzla a prípadu z `index!G6` |
| Ik'' | Ik'' = c·Sbase/(√3·Un·Zth_pu) | ako doteraz, ale c už nie je pevné 1,1 |
| Sieťový napájač (čl. 6.2) | Z_Q = c·Un/(√3·Ik''); X_Q = Z_Q/√(1+(R/X)²), R_Q = (R/X)·X_Q | doteraz čisto reaktančný |
| Korekcia trafa K_T (čl. 6.3.3) | K_T = 0,95·cmax/(1 + 0,6·x_T), x_T = Xk·Sn/U1n² | **len v max prípade** (v min sa podľa normy neaplikuje); cmax podľa NN strany; Sn a U1n sa čítajú zo stĺpcov E a F hárku `transformatory` |
| Korekcia generátora K_G (čl. 6.6.2) | K_G = (Un/U_rG)·cmax/(1 + xd''·sinφ_r), xd'' = Xd''[Ω]·Sn_G/U_rG² | Z_GK = K_G·(R_G + jXd''); U_rG = V_ref (stĺpec Q); aplikuje sa v oboch prípadoch |
| Fiktívny odpor generátora | R_Gf = 0,07·Xd'' ak R_G nie je zadaný | VN generátory S_rG < 100 MVA; potrebný pre korektné ip |
| Motory VN (čl. 6.5) | Z_M zo stĺpca P; R_M zo stĺpca L, inak R_M/X_M = 0,1 | **v min prípade sa príspevky motorov zanedbávajú** |
| Nárazový prúd ip (čl. 8) | ip = κ·√2·Ik''; κ = min(1,15·κ_b, 2,0) pre VN (1,8 pre NN), κ_b = 1,02 + 0,98·e^(−3·R/X) | metóda B pre zauzlené siete (konzervatívna); R/X z Théveninovej impedancie v mieste skratu |
| Vetvové príspevky | I_f = c_f/Z_ff; V_i = c_f − Z_if·I_f; I_vetvy = (V_i − V_j)·y_s | metóda ekvivalentného napäťového zdroja; trafo cez ys/a², ys/a; motory/generátory I = V·y |

## Spätná kompatibilita a zmeny výsledkov

- Bez vyplnenia nových vstupov beží výpočet ako doteraz v **max prípade** — ale výsledky sa **zámerne zmenia** oproti starej verzii: K_T sa počíta z existujúcich stĺpcov E/F hárku `transformatory` (typicky zníži impedanciu trafa o ~2–5 % → vyšší Ik'' za trafom), napájač a motory dostali reálnu zložku R/X = 0,1 a generátory fiktívny odpor. To je očakávané správanie — presne tieto chýbajúce korekcie boli identifikované ako pravdepodobná príčina ~20 % odchýlky (`to_do` #14).
- Stĺpec `uzly!J` naďalej slúži ako vstup Ik'' pre slack (fallback), `data!K13/K14` má prednosť a neprepisuje sa.
- K_G = 1, kým sa nevyplnia nové stĺpce T/U — príspevky generátorov sa teda bez nich počítajú po starom (len s fiktívnym R_Gf).

## Známe zjednodušenia (zámerne mimo rozsahu)

- Len trojfázový súmerný skrat (bez 2f/1f — vyžadovali by súostavy zložiek).
- Min prípad nekoriguje odpory vedení na koncovú teplotu (pandapower `endtemp_degree`); Ik''min je preto mierne konzervatívny. Súvisí s `to_do` #9.
- Bez Ib (vypínací prúd) a Ith (tepelný ekvivalent).
- κ metódou B (konzervatívna); presnejšia metóda C (ekvivalentná frekvencia) by vyžadovala druhé zostavenie a inverziu matice pri f = 20 Hz.

## Krížová validácia

Po exporte siete (viď [`pandapower_export_navrh.py`](pandapower_export_navrh.py)) porovnávať s:

```python
sc.calc_sc(net, fault="3ph", case="max",  # resp. "min"
           ip=True, topology="meshed", kappa_method="B")
```

pandapower predvolene používa `kappa_method="C"` — pre porovnanie ip s touto implementáciou treba explicitne `"B"` a `topology="meshed"`; Ik'' je od metódy κ nezávislý.
