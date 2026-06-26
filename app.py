# Check Modulação Mínima
        base["Check Modulação Mínima"] = "-"
        _mod_min_cc = pd.to_numeric(base["Modulação Mínima CCEE"], errors="coerce")
        _mask_min_valid = _mask_calcular_min & _mod_min_cc.notna()
        
        # Se ambos forem "-", o check é OK
        _mask_ambos_traco_min = (base["Modulação Mínima"].astype(str).str.strip() == "-") & (base["Modulação Mínima CCEE"].astype(str).str.strip() == "-")
        base.loc[_mask_ambos_traco_min, "Check Modulação Mínima"] = "OK"

        if _mask_min_valid.any():
            _diff_min = pd.to_numeric(base.loc[_mask_min_valid, "Modulação Mínima"]) - _mod_min_cc.loc[_mask_min_valid]
            # Aplica as regras apenas onde não foi definido como OK pelo critério dos traços
            _mask_min_calc = _mask_min_valid & (base["Check Modulação Mínima"] != "OK")
            base.loc[_mask_min_calc, "Check Modulação Mínima"] = "OK"
            base.loc[_mask_min_calc & (_diff_min > _tol_mod), "Check Modulação Mínima"] = "Book maior"
            base.loc[_mask_min_calc & (_diff_min < -_tol_mod), "Check Modulação Mínima"] = "CCEE maior"

        # Check Modulação Máxima
        base["Check Modulação Máxima"] = "-"
        _mod_max_cc = pd.to_numeric(base["Modulação Máxima CCEE"], errors="coerce")
        _mask_max_valid = _mask_calcular_max & _mod_max_cc.notna()
        
        # Se ambos forem "-", o check é OK
        _mask_ambos_traco_max = (base["Modulação Máxima"].astype(str).str.strip() == "-") & (base["Modulação Máxima CCEE"].astype(str).str.strip() == "-")
        base.loc[_mask_ambos_traco_max, "Check Modulação Máxima"] = "OK"

        if _mask_max_valid.any():
            _diff_max = pd.to_numeric(base.loc[_mask_max_valid, "Modulação Máxima"]) - _mod_max_cc.loc[_mask_max_valid]
            _mask_max_calc = _mask_max_valid & (base["Check Modulação Máxima"] != "OK")
            base.loc[_mask_max_calc, "Check Modulação Máxima"] = "OK"
            base.loc[_mask_max_calc & (_diff_max > _tol_mod), "Check Modulação Máxima"] = "Book maior"
            base.loc[_mask_max_calc & (_diff_max < -_tol_mod), "Check Modulação Máxima"] = "CCEE maior"

        # Check Modulação Tipo
        base["Check Modulação"] = "-"
        # Se ambos forem "-", o check é OK
        _mask_ambos_traco_tipo = (base["Modulação WBC"].astype(str).str.strip() == "-") & (base["Modulação CCEE"].astype(str).str.strip() == "-")
        base.loc[_mask_ambos_traco_tipo, "Check Modulação"] = "OK"

        _mask_tipo_valid = _mask_tem_contrato & (~base["Modulação CCEE"].astype(str).str.strip().isin(["", "-", "None", "nan"]))
        if _mask_tipo_valid.any():
            _mask_div_tipo = base["Modulação WBC"].astype(str).str.strip().str.upper() != base["Modulação CCEE"].astype(str).str.strip().str.upper()
            _mask_tipo_calc = _mask_tipo_valid & (base["Check Modulação"] != "OK")
            base.loc[_mask_tipo_calc, "Check Modulação"] = "OK"
            base.loc[_mask_tipo_calc & _mask_div_tipo, "Check Modulação"] = "Divergente"
