# -*- coding: utf-8 -*-
"""
===================================================================================================
PDKS ve Puantaj Konfigürasyon Yöneticisi (config/settings.py)
===================================================================================================
rules.json dosyasından fabrika çalışma kurallarını, tolerans sürelerini ve bölüm primlerini yükler.
Dosya bulunamazsa güvenli varsayılan değerleri (fallback) devreye alır.
"""

import json
import os
from typing import Dict, List, Optional, Tuple

DEFAULT_CONFIG_PATH = os.path.join(os.path.dirname(__file__), "rules.json")
RULES_FILE_PATH = DEFAULT_CONFIG_PATH

DEFAULT_RULES = {
    "version": "1.0",
    "name": "Fide Konserve Vardiya ve Prim Kuralları",
    "day_shift": {
        "start_time": "08:00",
        "tolerance_time": "08:20",
        "end_time": "17:00",
        "standard_hours": 7.5,
        "lunch_break_hours": 1.5,
        "late_cut_minutes": 30,
        "late_cut_hours": 0.5,
        "eight_four_window": {"start": "15:50", "end": "16:25"}
    },
    "evening_shift": {
        "start_time": "16:00",
        "tolerance_time": "16:20",
        "end_time": "24:00",
        "standard_hours": 7.5,
        "break_hours": 0.5
    },
    "night_shift": {
        "start_time": "24:00",
        "end_time": "08:00",
        "standard_hours": 7.5,
        "break_hours": 0.5
    },
    "night_shift_20": {
        "start_time": "20:00",
        "tolerance_time": "20:20",
        "standard_end_time": "04:00",
        "standard_hours": 7.5,
        "break_hours": 0.5,
        "description": "20:00 - 08:00 Gece / 12 Saatlik Vardiya (04:00 sonrası fazla mesai)"
    },
    "overtime_policy": {
        "grace_period_minutes": 25,
        "step_hours": 0.5
    },
    "department_bonuses": {},
    "missing_punch_policy": {
        "morning_cutoff": "12:30",
        "evening_cutoff": "20:00",
        "default_morning_hours": 7.5,
        "default_evening_start": "08:00"
    }
}

def load_rules(config_path: Optional[str] = None) -> Dict:
    """Konfigürasyon dosyasını yükler, hata durumunda varsayılan kuralları döndürür."""
    path = config_path or DEFAULT_CONFIG_PATH
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data
        except Exception as e:
            print(f"Uyarı: {path} yüklenemedi ({e}), varsayılan kurallar devrede.")
    return DEFAULT_RULES

def save_rules(rules_data: Dict, config_path: Optional[str] = None):
    """Kuralları JSON dosyasına yazar."""
    path = config_path or DEFAULT_CONFIG_PATH
    with open(path, "w", encoding="utf-8") as f:
        json.dump(rules_data, f, ensure_ascii=False, indent=2)

def calculate_department_bonus(
    dept_name: str,
    worked_hours: float,
    rules: Optional[Dict] = None
) -> Tuple[float, str]:
    """
    Personelin bölümüne ve fiili çalışma saatine göre hak ettiği prim saatini ve açıklamasını döndürür.
    Döndürdüğü değer: (bonus_hours, bonus_note)
    """
    if worked_hours <= 0:
        return 0.0, ""
        
    cfg = rules or load_rules()
    bonuses_cfg = cfg.get("department_bonuses", {})
    
    # Türkçe karakter duyarsız arama için bölüm adını normalize et
    norm_dept = str(dept_name).upper()
    norm_dept = norm_dept.replace("İ", "I").replace("Ü", "U").replace("Ş", "S").replace("Ğ", "G").replace("Ç", "C").replace("Ö", "O").strip()
    
    for key, b_info in bonuses_cfg.items():
        keywords = b_info.get("match_keywords", [])
        matched = any(kw in norm_dept for kw in keywords)
        if matched:
            threshold = float(b_info.get("min_hours_threshold", 0.0))
            if worked_hours >= threshold:
                b_hours = float(b_info.get("bonus_hours", 0.0))
                b_name = b_info.get("name", key)
                tot_h = worked_hours + b_hours
                note = f"{b_name}: Fiili {worked_hours:.1f}".replace(".", ",") + f"s -> Bonuslu {tot_h:.1f}".replace(".", ",") + f"s (+{b_hours:.1f}".replace(".", ",") + "s Prim)"
                return b_hours, note
                
    return 0.0, ""
