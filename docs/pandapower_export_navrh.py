# -*- coding: utf-8 -*-
"""
NÁVRH (príloha reportu docs/porovnanie_pandapower.md, kap. 6)

Krížová validácia VBA load-flow / skratových výpočtov cez pandapower.

Vstup: CSV exporty hárkov zošita (uložiť každý hárok ako CSV so zachovaním
stĺpcov popísaných v reporte, oddeľovač ';', desatinná čiarka je ošetrená).
Očakávané súbory v adresári `export/`:
    uzly.csv, vedenia.csv, transformatory.csv, spinace.csv,
    reaktory.csv, dif_reaktory.csv, kompenzacia.csv, motoryVN.csv, generatory.csv

Inštalácia:  pip install pandapower pandas
Overené proti pandapower >= 2.14 (API create_* / runpp / calc_sc).

POZNÁMKA: Toto je návrh na odsúhlasenie, nie hotový nástroj — čísla stĺpcov
zodpovedajú rozloženiu hárkov popísanému v reporte a pred prvým behom ich
treba skontrolovať proti aktuálnemu zošitu.
"""

import math

import pandas as pd
import pandapower as pp
import pandapower.shortcircuit as sc

F_HZ = 50.0
SN_IMPEDANCE_MVA = 100.0  # vzťažný výkon pre create_impedance (reaktory)


def _num(x):
    """VBA ParseDouble ekvivalent: toleruje čiarku, prázdne -> None (nie 0!)."""
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return None
    s = str(x).strip().replace(",", ".")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        raise ValueError(f"Nečitateľná hodnota: {x!r}")  # zámerne tvrdo, viď report kap. 5.4


def build_net(d="export"):
    net = pp.create_empty_network(f_hz=F_HZ, sn_mva=1.0)
    bus_idx = {}

    # --- uzly: B meno, C typ, D |V| kV, F P MW, G Q Mvar (odbery záporné), J Ik'' kA
    uzly = pd.read_csv(f"{d}/uzly.csv", sep=";")
    for _, r in uzly.iterrows():
        name = str(r["meno"]).strip()
        vn = _num(r["V_kV"])          # báza uzla = najbližšia hladina; tu priamo zadané kV
        b = pp.create_bus(net, vn_kv=vn, name=name)
        bus_idx[name] = b
        typ = str(r.get("typ", "PQ")).strip().lower()
        p, q = _num(r.get("P_MW")) or 0.0, _num(r.get("Q_Mvar")) or 0.0
        if typ == "slack":
            ik = _num(r.get("Ik_kA"))  # uzly!J — vstup pre skratový napájač
            kw = {}
            if ik:
                kw = dict(s_sc_max_mva=math.sqrt(3) * vn * ik, rx_max=0.1)
            pp.create_ext_grid(net, bus=b, vm_pu=1.0, **kw)
        elif p or q:
            # VBA konvencia: injekcie (odber záporný) -> pandapower load = kladný odber
            pp.create_load(net, bus=b, p_mw=-p, q_mvar=-q, name=name)

    # --- vedenia: R, X [ohm] celkové, B [S] celkové -> na 1 km
    ved = pd.read_csv(f"{d}/vedenia.csv", sep=";")
    for _, r in ved.iterrows():
        if int(r.get("status", 1)) == 0:
            continue
        c_nf = (_num(r.get("B_S")) or 0.0) / (2 * math.pi * F_HZ) * 1e9
        pp.create_line_from_parameters(
            net, from_bus=bus_idx[str(r["od"]).strip()], to_bus=bus_idx[str(r["do"]).strip()],
            length_km=1.0, r_ohm_per_km=_num(r["R_ohm"]), x_ohm_per_km=_num(r["X_ohm"]),
            c_nf_per_km=c_nf, max_i_ka=_num(r.get("In_A", 0)) / 1000 if _num(r.get("In_A")) else 1.0,
            name=str(r["meno"]).strip(),
        )

    # --- transformatory: štítkové hodnoty (viď report, kap. 6)
    trf = pd.read_csv(f"{d}/transformatory.csv", sep=";")
    for _, r in trf.iterrows():
        sn = _num(r["Sn_MVA"])
        pp.create_transformer_from_parameters(
            net, hv_bus=bus_idx[str(r["od"]).strip()], lv_bus=bus_idx[str(r["do"]).strip()],
            sn_mva=sn, vn_hv_kv=_num(r["U1n_kV"]), vn_lv_kv=_num(r["U2n_kV"]),
            vk_percent=_num(r["uk_perc"]),
            vkr_percent=_num(r["dPk_kW"]) / (10.0 * sn),      # ΔPk[kW]/(10·Sn[MVA])
            pfe_kw=_num(r["dP0_kW"]), i0_percent=_num(r["i0_perc"]),
            tap_side="hv", tap_neutral=9, tap_step_percent=2.0,
            tap_min=int(_num(r.get("odb_min")) or 1), tap_max=int(_num(r.get("odb_max")) or 17),
            tap_pos=int(_num(r.get("odb_akt")) or 9),
            name=str(r["meno"]).strip(),
        )

    # --- spinace: bus-bus switch, BEZ 1e-6 ohm impedancie (pandapower uzly zlúči)
    sw = pd.read_csv(f"{d}/spinace.csv", sep=";")
    for _, r in sw.iterrows():
        pp.create_switch(net, bus=bus_idx[str(r["od"]).strip()],
                         element=bus_idx[str(r["do"]).strip()], et="b",
                         closed=int(r.get("status", 1)) > 0, name=str(r["meno"]).strip())

    # --- reaktory + dif_reaktory: sériové impedancie (R, X v ohmoch -> p.u. na Sn)
    for fname in ("reaktory", "dif_reaktory"):
        df = pd.read_csv(f"{d}/{fname}.csv", sep=";")
        for _, r in df.iterrows():
            fb, tb = bus_idx[str(r["od"]).strip()], bus_idx[str(r["do"]).strip()]
            zb = net.bus.vn_kv.at[fb] ** 2 / SN_IMPEDANCE_MVA
            pp.create_impedance(net, from_bus=fb, to_bus=tb,
                                rft_pu=_num(r["R_ohm"]) / zb, xft_pu=_num(r["X_ohm"]) / zb,
                                sn_mva=SN_IMPEDANCE_MVA, name=str(r["meno"]).strip())

    # --- kompenzacia: shunt; q_mvar < 0 = kapacitná dodávka (overiť znamienko!)
    komp = pd.read_csv(f"{d}/kompenzacia.csv", sep=";")
    for _, r in komp.iterrows():
        if int(r.get("status", 1)) != 1:
            continue
        b = bus_idx[str(r["uzol"]).strip()]
        x_net = (_num(r["XC_ohm"]) or 0.0) - (_num(r["XL_ohm"]) or 0.0)
        un = net.bus.vn_kv.at[b]
        pp.create_shunt(net, bus=b, q_mvar=-(un ** 2) / x_net, p_mw=0.0)

    # --- motoryVN: load flow ako konšt. Z záťaž; skratový príspevok cez sgen/motor
    mot = pd.read_csv(f"{d}/motoryVN.csv", sep=";")
    for _, r in mot.iterrows():
        if int(r.get("status", 1)) != 1:
            continue
        b = bus_idx[str(r["uzol"]).strip()]
        un = net.bus.vn_kv.at[b]
        g, bsh = _num(r.get("G_S")) or 0.0, _num(r.get("B_S")) or 0.0
        # G+jB [S] pri Un -> P,Q odber (B záporné = induktívne)
        pp.create_load(net, bus=b, p_mw=g * un ** 2, q_mvar=-bsh * un ** 2,
                       const_z_percent=100.0, name=str(r["meno"]).strip())

    # --- generatory: PQ -> sgen, EMF -> gen (PV uzol na svorkách)
    gen = pd.read_csv(f"{d}/generatory.csv", sep=";")
    for _, r in gen.iterrows():
        if int(r.get("status", 1)) != 1:
            continue
        b = bus_idx[str(r["uzol"]).strip()]
        if str(r.get("mod", "PQ")).strip().upper() == "EMF":
            pp.create_gen(net, bus=b, p_mw=_num(r["P_MW"]),
                          vm_pu=_num(r["Vref_kV"]) / net.bus.vn_kv.at[b],
                          name=str(r["meno"]).strip())
        else:
            pp.create_sgen(net, bus=b, p_mw=_num(r["P_MW"]), q_mvar=_num(r["Qref_Mvar"]),
                           name=str(r["meno"]).strip())

    return net


def main():
    net = build_net()

    # 1) load flow — porovnať s uzly!H/I a výsledkovými stĺpcami vedení/tráf
    pp.runpp(net, calculate_voltage_angles=True, init="flat", tolerance_mva=1e-6)
    print("=== Napätia uzlov (vm_pu, va_degree) ===")
    print(net.res_bus.join(net.bus["name"]).to_string())
    print("\n=== Vedenia (p_from_mw, q_from_mvar, i_ka, pl_mw) ===")
    print(net.res_line.join(net.line["name"]).to_string())

    # 2) skraty IEC 60909 — porovnať s uzly!J (Ik'') a uzly!N (ip) z VBA behu (mode 2).
    # VBA počíta ip metódou B pre zauzlené siete -> pre porovnanie ip treba
    # kappa_method="B" a topology="meshed" (pandapower default je metóda C);
    # Ik'' je od metódy kappa nezávislý.
    sc.calc_sc(net, fault="3ph", case="max", ip=True, branch_results=False,
               topology="meshed", kappa_method="B")
    print("\n=== Ik'' [kA] a ip [kA] po uzloch (IEC 60909, case=max) ===")
    print(net.res_bus_sc.join(net.bus["name"]).to_string())


if __name__ == "__main__":
    main()
