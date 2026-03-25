import re
import io
from dataclasses import dataclass
from typing import Any, Dict, List, Optional
from collections import defaultdict

import pandas as pd
import streamlit as st

st.set_page_config(page_title="AP Metrics App", layout="wide")
st.title("NWC — AP Metrics (Auto-calc)")


# ============================================================
# Metric Model
# ============================================================

@dataclass
class MetricSpec:
    lever: str
    kpi: str
    components: List[str]
    calc_type: str  # identity | ratio_pct | multi_ratio_pct
    numerator: Optional[str] = None
    denominator: Optional[str] = None
    numerators: Optional[List[str]] = None
    multi_denominator: Optional[str] = None
    notes: Optional[str] = None
    metric_name: Optional[str] = None
    region: Optional[str] = None


# ============================================================
# ORIGINAL AP METRIC CATALOG
# ============================================================

AP_METRICS: List[MetricSpec] = [
    # ---------- Eliminating Early Payments ----------
    MetricSpec("Eliminating Early Payments", "Avg AP as % of Purchases (Overall) (Industry / Region): Overall",
               ["AP", "Purchases (Overall)"], "ratio_pct", "AP", "Purchases (Overall)"),
    MetricSpec("Eliminating Early Payments", "Avg AP as % of Purchases (Overall) (Industry / Region): Americas",
               ["AP", "Purchases (Overall)"], "ratio_pct", "AP", "Purchases (Overall)"),
    MetricSpec("Eliminating Early Payments", "Avg AP as % of Purchases (Overall) (Industry / Region): APAC",
               ["AP", "Purchases (Overall)"], "ratio_pct", "AP", "Purchases (Overall)"),
    MetricSpec("Eliminating Early Payments", "Avg AP as % of Purchases (Overall) (Industry / Region): EMEA",
               ["AP", "Purchases (Overall)"], "ratio_pct", "AP", "Purchases (Overall)"),

    MetricSpec("Eliminating Early Payments", "Early Payments as % of Purchases (Overall) (Industry / Region): Overall",
               ["Early Payments", "Purchases (Overall)"], "ratio_pct", "Early Payments", "Purchases (Overall)"),
    MetricSpec("Eliminating Early Payments", "Early Payments as % of Purchases (Overall) (Industry / Region): Americas",
               ["Early Payments (Americas)", "Purchases (Overall) (Americas)"], "ratio_pct",
               "Early Payments (Americas)", "Purchases (Overall) (Americas)"),
    MetricSpec("Eliminating Early Payments", "Early Payments as % of Purchases (Overall) (Industry / Region): APAC",
               ["Early Payments (APAC)", "Purchases (Overall) (APAC)"], "ratio_pct",
               "Early Payments (APAC)", "Purchases (Overall) (APAC)"),
    MetricSpec("Eliminating Early Payments", "Early Payments as % of Purchases (Overall) (Industry / Region): EMEA",
               ["Early Payments (EMEA)", "Purchases (Overall) (EMEA)"], "ratio_pct",
               "Early Payments (EMEA)", "Purchases (Overall) (EMEA)"),

    MetricSpec("Eliminating Early Payments", "Early payments by # early Days Purchases (Overall) Industry buckets",
               [
                   "Early Payments 1-2 Days Purchases (Overall)",
                   "Early Payments 3-5 Days Purchases (Overall)",
                   "Early Payments 6-10 Days Purchases (Overall)",
                   "Early Payments 11-15 Days Purchases (Overall)",
                   "Early Payments 16-30 Days Purchases (Overall)",
                   "Early Payments 30+ Days Purchases (Overall)",
                   "Purchases (Overall)",
               ],
               "multi_ratio_pct",
               numerators=[
                   "Early Payments 1-2 Days Purchases (Overall)",
                   "Early Payments 3-5 Days Purchases (Overall)",
                   "Early Payments 6-10 Days Purchases (Overall)",
                   "Early Payments 11-15 Days Purchases (Overall)",
                   "Early Payments 16-30 Days Purchases (Overall)",
                   "Early Payments 30+ Days Purchases (Overall)",
               ],
               multi_denominator="Purchases (Overall)",
               notes="Auto-calculates each bucket % = bucket / Purchases(Overall) * 100."
               ),

    MetricSpec("Eliminating Early Payments", "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (Overall)",
               ["Cash Benefit (Overall)", "Purchases (Overall)"], "ratio_pct", "Cash Benefit (Overall)", "Purchases (Overall)"),
    MetricSpec("Eliminating Early Payments", "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (Americas)",
               ["Cash Benefit (Americas)", "Purchases (Americas)"], "ratio_pct", "Cash Benefit (Americas)", "Purchases (Americas)"),
    MetricSpec("Eliminating Early Payments", "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (APAC)",
               ["Cash Benefit (APAC)", "Purchases (APAC)"], "ratio_pct", "Cash Benefit (APAC)", "Purchases (APAC)"),
    MetricSpec("Eliminating Early Payments", "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (EMEA)",
               ["Cash Benefit (EMEA)", "Purchases (EMEA)"], "ratio_pct", "Cash Benefit (EMEA)", "Purchases (EMEA)"),

    # ---------- Lengthen Payment Terms ----------
    MetricSpec("Lengthen Payment Terms", "DPO (Industry / Region): Overall", ["DPO (Overall)"], "identity", numerator="DPO (Overall)"),
    MetricSpec("Lengthen Payment Terms", "DPO (Industry / Region): Americas", ["DPO (Americas)"], "identity", numerator="DPO (Americas)"),
    MetricSpec("Lengthen Payment Terms", "DPO (Industry / Region): APAC", ["DPO (APAC)"], "identity", numerator="DPO (APAC)"),
    MetricSpec("Lengthen Payment Terms", "DPO (Industry / Region): EMEA", ["DPO (EMEA)"], "identity", numerator="DPO (EMEA)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) (Region: Overall)",
               ["Avg PT Days Purchases (Overall)"], "identity", numerator="Avg PT Days Purchases (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) (Region: Americas)",
               ["Avg PT Days Purchases (Americas)"], "identity", numerator="Avg PT Days Purchases (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) (Region: APAC)",
               ["Avg PT Days Purchases (APAC)"], "identity", numerator="Avg PT Days Purchases (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) (Region: EMEA)",
               ["Avg PT Days Purchases (EMEA)"], "identity", numerator="Avg PT Days Purchases (EMEA)"),

    MetricSpec("Lengthen Payment Terms", "% Purchases (Overall) by Payment Industry buckets (Region: Overall)",
               [
                   "1-15 Days Purchases (Overall)",
                   "16-30 Days Purchases (Overall)",
                   "31-45 Days Purchases (Overall)",
                   "46-60 Days Purchases (Overall)",
                   "61-90 Days Purchases (Overall)",
                   "91-120 Days Purchases (Overall)",
                   "120 Days+ Purchases (Overall)",
                   "Purchases (Overall)",
               ],
               "multi_ratio_pct",
               numerators=[
                   "1-15 Days Purchases (Overall)",
                   "16-30 Days Purchases (Overall)",
                   "31-45 Days Purchases (Overall)",
                   "46-60 Days Purchases (Overall)",
                   "61-90 Days Purchases (Overall)",
                   "91-120 Days Purchases (Overall)",
                   "120 Days+ Purchases (Overall)",
               ],
               multi_denominator="Purchases (Overall)",
               notes="Auto-calculates each bucket % = bucket / Purchases(Overall) * 100."),
    MetricSpec("Lengthen Payment Terms", "% Purchases (Overall) by Payment Industry buckets (Region: Americas)",
               [
                   "1-15 Days Purchases (Americas)",
                   "16-30 Days Purchases (Americas)",
                   "31-45 Days Purchases (Americas)",
                   "46-60 Days Purchases (Americas)",
                   "61-90 Days Purchases (Americas)",
                   "91-120 Days Purchases (Americas)",
                   "120 Days+ Purchases (Americas)",
                   "Purchases (Americas)",
               ],
               "multi_ratio_pct",
               numerators=[
                   "1-15 Days Purchases (Americas)",
                   "16-30 Days Purchases (Americas)",
                   "31-45 Days Purchases (Americas)",
                   "46-60 Days Purchases (Americas)",
                   "61-90 Days Purchases (Americas)",
                   "91-120 Days Purchases (Americas)",
                   "120 Days+ Purchases (Americas)",
               ],
               multi_denominator="Purchases (Americas)"),
    MetricSpec("Lengthen Payment Terms", "% Purchases (Overall) by Payment Industry buckets (Region: APAC)",
               [
                   "1-15 Days Purchases (APAC)",
                   "16-30 Days Purchases (APAC)",
                   "31-45 Days Purchases (APAC)",
                   "46-60 Days Purchases (APAC)",
                   "61-90 Days Purchases (APAC)",
                   "91-120 Days Purchases (APAC)",
                   "120 Days+ Purchases (APAC)",
                   "Purchases (APAC)",
               ],
               "multi_ratio_pct",
               numerators=[
                   "1-15 Days Purchases (APAC)",
                   "16-30 Days Purchases (APAC)",
                   "31-45 Days Purchases (APAC)",
                   "46-60 Days Purchases (APAC)",
                   "61-90 Days Purchases (APAC)",
                   "91-120 Days Purchases (APAC)",
                   "120 Days+ Purchases (APAC)",
               ],
               multi_denominator="Purchases (APAC)"),
    MetricSpec("Lengthen Payment Terms", "% Purchases (Overall) by Payment Industry buckets (Region: EMEA)",
               [
                   "1-15 Days Purchases (EMEA)",
                   "16-30 Days Purchases (EMEA)",
                   "31-45 Days Purchases (EMEA)",
                   "46-60 Days Purchases (EMEA)",
                   "61-90 Days Purchases (EMEA)",
                   "91-120 Days Purchases (EMEA)",
                   "120 Days+ Purchases (EMEA)",
                   "Purchases (EMEA)",
               ],
               "multi_ratio_pct",
               numerators=[
                   "1-15 Days Purchases (EMEA)",
                   "16-30 Days Purchases (EMEA)",
                   "31-45 Days Purchases (EMEA)",
                   "46-60 Days Purchases (EMEA)",
                   "61-90 Days Purchases (EMEA)",
                   "91-120 Days Purchases (EMEA)",
                   "120 Days+ Purchases (EMEA)",
               ],
               multi_denominator="Purchases (EMEA)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: overall)",
               ["Average PT days Top 10 (Overall)"], "identity", numerator="Average PT days Top 10 (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: overall)",
               ["Average PT days Top 15 (Overall)"], "identity", numerator="Average PT days Top 15 (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: overall)",
               ["Average PT days Top 20 (Overall)"], "identity", numerator="Average PT days Top 20 (Overall)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: Americas)",
               ["Average PT days Top 10 (Americas)"], "identity", numerator="Average PT days Top 10 (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: Americas)",
               ["Average PT days Top 15 (Americas)"], "identity", numerator="Average PT days Top 15 (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: Americas)",
               ["Average PT days Top 20 (Americas)"], "identity", numerator="Average PT days Top 20 (Americas)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: APAC)",
               ["Average PT days Top 10 (APAC)"], "identity", numerator="Average PT days Top 10 (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: APAC)",
               ["Average PT days Top 15 (APAC)"], "identity", numerator="Average PT days Top 15 (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: APAC)",
               ["Average PT days Top 20 (APAC)"], "identity", numerator="Average PT days Top 20 (APAC)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: EMEA)",
               ["Average PT days Top 10 (EMEA)"], "identity", numerator="Average PT days Top 10 (EMEA)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: EMEA)",
               ["Average PT days Top 15 (EMEA)"], "identity", numerator="Average PT days Top 15 (EMEA)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: EMEA)",
               ["Average PT days Top 20 (EMEA)"], "identity", numerator="Average PT days Top 20 (EMEA)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: overall)",
               ["Average PT days Top 80% (Overall)"], "identity", numerator="Average PT days Top 80% (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: overall)",
               ["Average PT days Next 15% (Overall)"], "identity", numerator="Average PT days Next 15% (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: overall)",
               ["Average PT days Last 5% (Overall)"], "identity", numerator="Average PT days Last 5% (Overall)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: Americas)",
               ["Average PT days Top 80% (Americas)"], "identity", numerator="Average PT days Top 80% (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: Americas)",
               ["Average PT days Next 15% (Americas)"], "identity", numerator="Average PT days Next 15% (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: Americas)",
               ["Average PT days Last 5% (Americas)"], "identity", numerator="Average PT days Last 5% (Americas)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: APAC)",
               ["Average PT days Top 80% (APAC)"], "identity", numerator="Average PT days Top 80% (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: APAC)",
               ["Average PT days Next 15% (APAC)"], "identity", numerator="Average PT days Next 15% (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: APAC)",
               ["Average PT days Last 5% (APAC)"], "identity", numerator="Average PT days Last 5% (APAC)"),

    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: EMEA)",
               ["Average PT days Top 80% (EMEA)"], "identity", numerator="Average PT days Top 80% (EMEA)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: EMEA)",
               ["Average PT days Next 15% (EMEA)"], "identity", numerator="Average PT days Next 15% (EMEA)"),
    MetricSpec("Lengthen Payment Terms", "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: EMEA)",
               ["Average PT days Last 5% (EMEA)"], "identity", numerator="Average PT days Last 5% (EMEA)"),

    MetricSpec("Lengthen Payment Terms", "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): Overall",
               ["Cash Benefit <30 PT days (Overall)", "Purchases <30 PT days (Overall)"], "ratio_pct",
               "Cash Benefit <30 PT days (Overall)", "Purchases <30 PT days (Overall)"),
    MetricSpec("Lengthen Payment Terms", "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): Americas",
               ["Cash Benefit <30 PT days (Americas)", "Purchases <30 PT days (Americas)"], "ratio_pct",
               "Cash Benefit <30 PT days (Americas)", "Purchases <30 PT days (Americas)"),
    MetricSpec("Lengthen Payment Terms", "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): APAC",
               ["Cash Benefit <30 PT days (APAC)", "Purchases <30 PT days (APAC)"], "ratio_pct",
               "Cash Benefit <30 PT days (APAC)", "Purchases <30 PT days (APAC)"),
    MetricSpec("Lengthen Payment Terms", "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): EMEA",
               ["Cash Benefit <30 PT days (EMEA)", "Purchases <30 PT days (EMEA)"], "ratio_pct",
               "Cash Benefit <30 PT days (EMEA)", "Purchases <30 PT days (EMEA)"),

    # ---------- Harmonizing Payment terms ----------
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Overall",
               ["Distinct PT Days Top 80% (Overall)"], "identity", numerator="Distinct PT Days Top 80% (Overall)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Overall",
               ["Distinct PT Days Next 15% (Overall)"], "identity", numerator="Distinct PT Days Next 15% (Overall)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Overall",
               ["Distinct PT Days Last 5% (Overall)"], "identity", numerator="Distinct PT Days Last 5% (Overall)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Americas",
               ["Distinct PT Days Top 80% (Americas)"], "identity", numerator="Distinct PT Days Top 80% (Americas)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Americas",
               ["Distinct PT Days Next 15% (Americas)"], "identity", numerator="Distinct PT Days Next 15% (Americas)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Americas",
               ["Distinct PT Days Last 5% (Americas)"], "identity", numerator="Distinct PT Days Last 5% (Americas)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): APAC",
               ["Distinct PT Days Top 80% (APAC)"], "identity", numerator="Distinct PT Days Top 80% (APAC)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): APAC",
               ["Distinct PT Days Next 15% (APAC)"], "identity", numerator="Distinct PT Days Next 15% (APAC)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): APAC",
               ["Distinct PT Days Last 5% (APAC)"], "identity", numerator="Distinct PT Days Last 5% (APAC)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): EMEA",
               ["Distinct PT Days Top 80% (EMEA)"], "identity", numerator="Distinct PT Days Top 80% (EMEA)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): EMEA",
               ["Distinct PT Days Next 15% (EMEA)"], "identity", numerator="Distinct PT Days Next 15% (EMEA)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): EMEA",
               ["Distinct PT Days Last 5% (EMEA)"], "identity", numerator="Distinct PT Days Last 5% (EMEA)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): Overall",
               ["Distinct PT Days Top 10 (Overall)"], "identity", numerator="Distinct PT Days Top 10 (Overall)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): Overall",
               ["Distinct PT Days Top 15 (Overall)"], "identity", numerator="Distinct PT Days Top 15 (Overall)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): Overall",
               ["Distinct PT Days Top 20 (Overall)"], "identity", numerator="Distinct PT Days Top 20 (Overall)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): Americas",
               ["Distinct PT Days Top 10 (Americas)"], "identity", numerator="Distinct PT Days Top 10 (Americas)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): Americas",
               ["Distinct PT Days Top 15 (Americas)"], "identity", numerator="Distinct PT Days Top 15 (Americas)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): Americas",
               ["Distinct PT Days Top 20 (Americas)"], "identity", numerator="Distinct PT Days Top 20 (Americas)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): APAC",
               ["Distinct PT Days Top 10 (APAC)"], "identity", numerator="Distinct PT Days Top 10 (APAC)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): APAC",
               ["Distinct PT Days Top 15 (APAC)"], "identity", numerator="Distinct PT Days Top 15 (APAC)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): APAC",
               ["Distinct PT Days Top 20 (APAC)"], "identity", numerator="Distinct PT Days Top 20 (APAC)"),

    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): EMEA",
               ["Distinct PT Days Top 10 (EMEA)"], "identity", numerator="Distinct PT Days Top 10 (EMEA)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): EMEA",
               ["Distinct PT Days Top 15 (EMEA)"], "identity", numerator="Distinct PT Days Top 15 (EMEA)"),
    MetricSpec("Harmonizing Payment terms", "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): EMEA",
               ["Distinct PT Days Top 20 (EMEA)"], "identity", numerator="Distinct PT Days Top 20 (EMEA)"),

    MetricSpec("Harmonizing Payment terms", "Std devn of PT Days Purchases (Overall) offered (Region: Overall)",
               ["Standard Deviation (Overall)"], "identity", numerator="Standard Deviation (Overall)"),
    MetricSpec("Harmonizing Payment terms", "Std devn of PT Days Purchases (Overall) offered (Region: Americas)",
               ["Standard Deviation (Americas)"], "identity", numerator="Standard Deviation (Americas)"),
    MetricSpec("Harmonizing Payment terms", "Std devn of PT Days Purchases (Overall) offered (Region: APAC)",
               ["Standard Deviation (APAC)"], "identity", numerator="Standard Deviation (APAC)"),
    MetricSpec("Harmonizing Payment terms", "Std devn of PT Days Purchases (Overall) offered (Region: EMEA)",
               ["Standard Deviation (EMEA)"], "identity", numerator="Standard Deviation (EMEA)"),

    MetricSpec("Harmonizing Payment terms", "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): Overall",
               ["Cash Benefit 2 (Overall)", "Purchases (Overall)"], "ratio_pct", "Cash Benefit 2 (Overall)", "Purchases (Overall)"),
    MetricSpec("Harmonizing Payment terms", "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): Americas",
               ["Cash Benefit 2 (Americas)", "Purchases (Americas)"], "ratio_pct", "Cash Benefit 2 (Americas)", "Purchases (Americas)"),
    MetricSpec("Harmonizing Payment terms", "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): APAC",
               ["Cash Benefit 2 (APAC)", "Purchases (APAC)"], "ratio_pct", "Cash Benefit 2 (APAC)", "Purchases (APAC)"),
    MetricSpec("Harmonizing Payment terms", "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): EMEA",
               ["Cash Benefit 2 (EMEA)", "Purchases (EMEA)"], "ratio_pct", "Cash Benefit 2 (EMEA)", "Purchases (EMEA)"),

    # ---------- Supplier Discount Optimization ----------
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Overall",
               ["Average discount % Top 80% (Overall)"], "identity", numerator="Average discount % Top 80% (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Overall",
               ["Average discount % Next 15% (Overall)"], "identity", numerator="Average discount % Next 15% (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Overall",
               ["Average discount % Last 5% (Overall)"], "identity", numerator="Average discount % Last 5% (Overall)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Americas",
               ["Average discount % Top 80% (Americas)"], "identity", numerator="Average discount % Top 80% (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Americas",
               ["Average discount % Next 15% (Americas)"], "identity", numerator="Average discount % Next 15% (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Americas",
               ["Average discount % Last 5% (Americas)"], "identity", numerator="Average discount % Last 5% (Americas)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): APAC",
               ["Average discount % Top 80% (APAC)"], "identity", numerator="Average discount % Top 80% (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): APAC",
               ["Average discount % Next 15% (APAC)"], "identity", numerator="Average discount % Next 15% (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): APAC",
               ["Average discount % Last 5% (APAC)"], "identity", numerator="Average discount % Last 5% (APAC)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): EMEA",
               ["Average discount % Top 80% (EMEA)"], "identity", numerator="Average discount % Top 80% (EMEA)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): EMEA",
               ["Average discount % Next 15% (EMEA)"], "identity", numerator="Average discount % Next 15% (EMEA)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): EMEA",
               ["Average discount % Last 5% (EMEA)"], "identity", numerator="Average discount % Last 5% (EMEA)"),

    MetricSpec("Supplier Discount Optimization", "Avg discount % offered (Top vendors / Industry / Region): Overall",
               ["Average discount percentage (Overall)"], "identity", numerator="Average discount percentage (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount % offered (Top vendors / Industry / Region): Americas",
               ["Average discount percentage (Americas)"], "identity", numerator="Average discount percentage (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount % offered (Top vendors / Industry / Region): APAC",
               ["Average discount percentage (APAC)"], "identity", numerator="Average discount percentage (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount % offered (Top vendors / Industry / Region): EMEA",
               ["Average discount percentage (EMEA)"], "identity", numerator="Average discount percentage (EMEA)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 10 vendors (Industry / Region): Overall",
               ["Average discount % Top 10 (Overall)"], "identity", numerator="Average discount % Top 10 (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 15 vendors (Industry / Region): Overall",
               ["Average discount % Top 15 (Overall)"], "identity", numerator="Average discount % Top 15 (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 20 vendors (Industry / Region): Overall",
               ["Average discount % Top 20 (Overall)"], "identity", numerator="Average discount % Top 20 (Overall)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 10 vendors (Industry / Region): Americas",
               ["Average discount % Top 10 (Americas)"], "identity", numerator="Average discount % Top 10 (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 15 vendors (Industry / Region): Americas",
               ["Average discount % Top 15 (Americas)"], "identity", numerator="Average discount % Top 15 (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 20 vendors (Industry / Region): Americas",
               ["Average discount % Top 20 (Americas)"], "identity", numerator="Average discount % Top 20 (Americas)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 10 vendors (Industry / Region): APAC",
               ["Average discount % Top 10 (APAC)"], "identity", numerator="Average discount % Top 10 (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 15 vendors (Industry / Region): APAC",
               ["Average discount % Top 15 (APAC)"], "identity", numerator="Average discount % Top 15 (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 20 vendors (Industry / Region): APAC",
               ["Average discount % Top 20 (APAC)"], "identity", numerator="Average discount % Top 20 (APAC)"),

    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 10 vendors (Industry / Region): EMEA",
               ["Average discount % Top 10 (EMEA)"], "identity", numerator="Average discount % Top 10 (EMEA)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 15 vendors (Industry / Region): EMEA",
               ["Average discount % Top 15 (EMEA)"], "identity", numerator="Average discount % Top 15 (EMEA)"),
    MetricSpec("Supplier Discount Optimization", "Average discount % (Overall) offered by top 20 vendors (Industry / Region): EMEA",
               ["Average discount % Top 20 (EMEA)"], "identity", numerator="Average discount % Top 20 (EMEA)"),

    MetricSpec("Supplier Discount Optimization", "Avg discount terms as % of WACC (Industry / Region): Overall",
               ["Average Discount (Overall)", "WACC (Overall)"], "ratio_pct", "Average Discount (Overall)", "WACC (Overall)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount terms as % of WACC (Industry / Region): Americas",
               ["Average Discount (Americas)", "WACC (Americas)"], "ratio_pct", "Average Discount (Americas)", "WACC (Americas)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount terms as % of WACC (Industry / Region): APAC",
               ["Average Discount (APAC)", "WACC (APAC)"], "ratio_pct", "Average Discount (APAC)", "WACC (APAC)"),
    MetricSpec("Supplier Discount Optimization", "Avg discount terms as % of WACC (Industry / Region): EMEA",
               ["Average Discount (EMEA)", "WACC (EMEA)"], "ratio_pct", "Average Discount (EMEA)", "WACC (EMEA)"),

    # ---------- Cheque runs ----------
    MetricSpec("Cheque runs", "% payments byCheque",
               ["Payments made through Cheque", "Purchases"], "ratio_pct", "Payments made through Cheque", "Purchases"),
    MetricSpec("Cheque runs", "# monthly cheque runs",
               ["Monthly Cheque runs"], "identity", numerator="Monthly Cheque runs"),

    # ---------- TBD ----------
    MetricSpec("TBD", "No. of suppliers in top 5 largest spend categories",
               ["Number of suppliers"], "identity", numerator="Number of suppliers"),
]


# ============================================================
# KPI naming map from current AP app
# ============================================================

KPI_NAMING_MAP: Dict[str, Dict[str, str]] = {
    "# monthly cheque runs": {"wanted": "# monthly cheque runs", "region": "Overall"},
    "% payments byCheque": {"wanted": "% payments by Cheque", "region": "Overall"},

    "Avg AP as % of Purchases (Overall) (Industry / Region): Overall": {"wanted": "Avg AP as % of Purchases", "region": "Overall"},
    "Avg AP as % of Purchases (Overall) (Industry / Region): Americas": {"wanted": "Avg AP as % of Purchases", "region": "Americas"},
    "Avg AP as % of Purchases (Overall) (Industry / Region): APAC": {"wanted": "Avg AP as % of Purchases", "region": "APAC"},
    "Avg AP as % of Purchases (Overall) (Industry / Region): EMEA": {"wanted": "Avg AP as % of Purchases", "region": "EMEA"},

    "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (Overall)": {"wanted": "Cash Benefit from Eliminating Early Payments as % of Purchases", "region": "Overall"},
    "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (Americas)": {"wanted": "Cash Benefit from Eliminating Early Payments as % of Purchases", "region": "Americas"},
    "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (APAC)": {"wanted": "Cash Benefit from Eliminating Early Payments as % of Purchases", "region": "APAC"},
    "Cash Benefit from Eliminating Early Payments as % of Purchases (Overall) (EMEA)": {"wanted": "Cash Benefit from Eliminating Early Payments as % of Purchases", "region": "EMEA"},

    "Early Payments as % of Purchases (Overall) (Industry / Region): Overall": {"wanted": "Early Payments as % of Purchases", "region": "Overall"},
    "Early Payments as % of Purchases (Overall) (Industry / Region): Americas": {"wanted": "Early Payments as % of Purchases", "region": "Americas"},
    "Early Payments as % of Purchases (Overall) (Industry / Region): APAC": {"wanted": "Early Payments as % of Purchases", "region": "APAC"},
    "Early Payments as % of Purchases (Overall) (Industry / Region): EMEA": {"wanted": "Early Payments as % of Purchases", "region": "EMEA"},

    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 11-15 Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (11-15 days)", "region": "Overall"},
    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 1-2 Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (1-2 days)", "region": "Overall"},
    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 16-30 Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (16-30 days)", "region": "Overall"},
    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 30+ Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (30+ days)", "region": "Overall"},
    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 3-5 Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (3-5 days)", "region": "Overall"},
    "Early payments by # early Days Purchases (Overall) Industry buckets — Early Payments 6-10 Days Purchases (Overall) (%)": {"wanted": "Early payments by # early days Industry buckets (6-10 days)", "region": "Overall"},

    "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (A: Top 80%)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (A: Top 80%)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (A: Top 80%)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (A: Top 80%)", "region": "EMEA"},

    "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (B: Next 15%)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (B: Next 15%)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (B: Next 15%)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (B: Next 15%)", "region": "EMEA"},

    "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (C: Last 5%)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (C: Last 5%)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (C: Last 5%)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (C: Last 5%)", "region": "EMEA"},

    "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (Top 10 Suppliers)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (Top 10 Suppliers)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (Top 10 Suppliers)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top 10 vendors (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (Top 10 Suppliers)", "region": "EMEA"},

    "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (Top 15 Suppliers)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (Top 15 Suppliers)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (Top 15 Suppliers)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top 15 vendors (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (Top 15 Suppliers)", "region": "EMEA"},

    "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): Overall": {"wanted": "# distinct PT days offered by top X vendors (Top 20 Suppliers)", "region": "Overall"},
    "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): Americas": {"wanted": "# distinct PT days offered by top X vendors (Top 20 Suppliers)", "region": "Americas"},
    "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): APAC": {"wanted": "# distinct PT days offered by top X vendors (Top 20 Suppliers)", "region": "APAC"},
    "# distinct PT Days Purchases (Overall) offered by top 20 vendors (Industry / Region): EMEA": {"wanted": "# distinct PT days offered by top X vendors (Top 20 Suppliers)", "region": "EMEA"},

    "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): Overall": {"wanted": "Cash Benefit from Harmonizing Payment Terms as % of Purchases", "region": "Overall"},
    "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): Americas": {"wanted": "Cash Benefit from Harmonizing Payment Terms as % of Purchases", "region": "Americas"},
    "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): APAC": {"wanted": "Cash Benefit from Harmonizing Payment Terms as % of Purchases", "region": "APAC"},
    "Cash Benefit as % of Purchases (Overall)2 (Industry / Region): EMEA": {"wanted": "Cash Benefit from Harmonizing Payment Terms as % of Purchases", "region": "EMEA"},

    "Std devn of PT Days Purchases (Overall) offered (Region: Overall)": {"wanted": "Std devn of PT days offered", "region": "Overall"},
    "Std devn of PT Days Purchases (Overall) offered (Region: Americas)": {"wanted": "Std devn of PT days offered", "region": "Americas"},
    "Std devn of PT Days Purchases (Overall) offered (Region: APAC)": {"wanted": "Std devn of PT days offered", "region": "APAC"},
    "Std devn of PT Days Purchases (Overall) offered (Region: EMEA)": {"wanted": "Std devn of PT days offered", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 1-15 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (1-15 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 1-15 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (1-15 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 1-15 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (1-15 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 1-15 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (1-15 days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 120 Days+ Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (120+ days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 120 Days+ Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (120+ days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 120 Days+ Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (120+ days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 120 Days+ Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (120+ days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 16-30 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (16-30 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 16-30 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (16-30 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 16-30 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (16-30 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 16-30 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (16-30 days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 31-45 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (31-45 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 31-45 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (31-45 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 31-45 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (31-45 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 31-45 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (31-45 days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 46-60 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (46-60 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 46-60 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (46-60 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 46-60 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (46-60 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 46-60 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (46-60 days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 61-90 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (61-90 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 61-90 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (61-90 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 61-90 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (61-90 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 61-90 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (61-90 days)", "region": "EMEA"},

    "% Purchases (Overall) by Payment Industry buckets (Region: Overall) — 91-120 Days Purchases (Overall) (%)": {"wanted": "% Purchases by Payment Industry buckets (91-120 days)", "region": "Overall"},
    "% Purchases (Overall) by Payment Industry buckets (Region: Americas) — 91-120 Days Purchases (Americas) (%)": {"wanted": "% Purchases by Payment Industry buckets (91-120 days)", "region": "Americas"},
    "% Purchases (Overall) by Payment Industry buckets (Region: APAC) — 91-120 Days Purchases (APAC) (%)": {"wanted": "% Purchases by Payment Industry buckets (91-120 days)", "region": "APAC"},
    "% Purchases (Overall) by Payment Industry buckets (Region: EMEA) — 91-120 Days Purchases (EMEA) (%)": {"wanted": "% Purchases by Payment Industry buckets (91-120 days)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) (Region: Overall)": {"wanted": "Avg PT days", "region": "Overall"},
    "Avg PT Days Purchases (Overall) (Region: Americas)": {"wanted": "Avg PT days", "region": "Americas"},
    "Avg PT Days Purchases (Overall) (Region: APAC)": {"wanted": "Avg PT days", "region": "APAC"},
    "Avg PT Days Purchases (Overall) (Region: EMEA)": {"wanted": "Avg PT days", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: overall)": {"wanted": "Avg PT days top X Vendor / others (A: Top 80%)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (A: Top 80%)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (A: Top 80%)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) Vendor / others A (Top: 80%)(Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (A: Top 80%)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: overall)": {"wanted": "Avg PT days top X Vendor / others (B: Next 15%)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (B: Next 15%)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (B: Next 15%)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) Vendor / others B (Next: 15%)(Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (B: Next 15%)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: overall)": {"wanted": "Avg PT days top X Vendor / others (C: Last 5%)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (C: Last 5%)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (C: Last 5%)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) Vendor / others C (Last: 5%)(Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (C: Last 5%)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: overall)": {"wanted": "Avg PT days top X Vendor / others (Top 10 Suppliers)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (Top 10 Suppliers)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (Top 10 Suppliers)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) top 10 Vendor / others (Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (Top 10 Suppliers)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: overall)": {"wanted": "Avg PT days top X Vendor / others (Top 15 Suppliers)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (Top 15 Suppliers)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (Top 15 Suppliers)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) top 15 Vendor / others (Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (Top 15 Suppliers)", "region": "EMEA"},

    "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: overall)": {"wanted": "Avg PT days top X Vendor / others (Top 20 Suppliers)", "region": "Overall"},
    "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: Americas)": {"wanted": "Avg PT days top X Vendor / others (Top 20 Suppliers)", "region": "Americas"},
    "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: APAC)": {"wanted": "Avg PT days top X Vendor / others (Top 20 Suppliers)", "region": "APAC"},
    "Avg PT Days Purchases (Overall) top 20 Vendor / others (Region: EMEA)": {"wanted": "Avg PT days top X Vendor / others (Top 20 Suppliers)", "region": "EMEA"},

    "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): Overall": {"wanted": "Cash Benefit from Lengthening Payment Terms as % of Purchases in <30 PT days Industry bucket", "region": "Overall"},
    "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): Americas": {"wanted": "Cash Benefit from Lengthening Payment Terms as % of Purchases in <30 PT days Industry bucket", "region": "Americas"},
    "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): APAC": {"wanted": "Cash Benefit from Lengthening Payment Terms as % of Purchases in <30 PT days Industry bucket", "region": "APAC"},
    "Cash Benefit as % Purchases (Overall) in <30 PT Days Purchases (Overall) Industry bucket (Industry / Region): EMEA": {"wanted": "Cash Benefit from Lengthening Payment Terms as % of Purchases in <30 PT days Industry bucket", "region": "EMEA"},

    "DPO (Industry / Region): Overall": {"wanted": "DPO", "region": "Overall"},
    "DPO (Industry / Region): Americas": {"wanted": "DPO", "region": "Americas"},
    "DPO (Industry / Region): APAC": {"wanted": "DPO", "region": "APAC"},
    "DPO (Industry / Region): EMEA": {"wanted": "DPO", "region": "EMEA"},

    "No. of suppliers in top 5 largest spend categories": {"wanted": "No. of suppliers in top 5 largest spend categories", "region": "Overall"},

    "Avg discount % offered (Top vendors / Industry / Region): Overall": {"wanted": "Avg discount % offered", "region": "Overall"},
    "Avg discount % offered (Top vendors / Industry / Region): Americas": {"wanted": "Avg discount % offered", "region": "Americas"},
    "Avg discount % offered (Top vendors / Industry / Region): APAC": {"wanted": "Avg discount % offered", "region": "APAC"},
    "Avg discount % offered (Top vendors / Industry / Region): EMEA": {"wanted": "Avg discount % offered", "region": "EMEA"},

    "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (A: Top 80%)", "region": "Overall"},
    "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (A: Top 80%)", "region": "Americas"},
    "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (A: Top 80%)", "region": "APAC"},
    "Average discount % (Overall) offered by top X A: (Top: 80%) vendors (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (A: Top 80%)", "region": "EMEA"},

    "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (B: Next 15%)", "region": "Overall"},
    "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (B: Next 15%)", "region": "Americas"},
    "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (B: Next 15%)", "region": "APAC"},
    "Average discount % (Overall) offered by top X vendors B: (Next 15%) (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (B: Next 15%)", "region": "EMEA"},

    "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (C: Last 5%)", "region": "Overall"},
    "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (C: Last 5%)", "region": "Americas"},
    "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (C: Last 5%)", "region": "APAC"},
    "Average discount % (Overall) offered by top X vendors C: (Last 5%) (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (C: Last 5%)", "region": "EMEA"},

    "Average discount % (Overall) offered by top 10 vendors (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (Top 10 Suppliers)", "region": "Overall"},
    "Average discount % (Overall) offered by top 10 vendors (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (Top 10 Suppliers)", "region": "Americas"},
    "Average discount % (Overall) offered by top 10 vendors (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (Top 10 Suppliers)", "region": "APAC"},
    "Average discount % (Overall) offered by top 10 vendors (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (Top 10 Suppliers)", "region": "EMEA"},

    "Average discount % (Overall) offered by top 15 vendors (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (Top 15 Suppliers)", "region": "Overall"},
    "Average discount % (Overall) offered by top 15 vendors (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (Top 15 Suppliers)", "region": "Americas"},
    "Average discount % (Overall) offered by top 15 vendors (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (Top 15 Suppliers)", "region": "APAC"},
    "Average discount % (Overall) offered by top 15 vendors (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (Top 15 Suppliers)", "region": "EMEA"},

    "Average discount % (Overall) offered by top 20 vendors (Industry / Region): Overall": {"wanted": "Avg discount % offered by top X vendors (Top 20 Suppliers)", "region": "Overall"},
    "Average discount % (Overall) offered by top 20 vendors (Industry / Region): Americas": {"wanted": "Avg discount % offered by top X vendors (Top 20 Suppliers)", "region": "Americas"},
    "Average discount % (Overall) offered by top 20 vendors (Industry / Region): APAC": {"wanted": "Avg discount % offered by top X vendors (Top 20 Suppliers)", "region": "APAC"},
    "Average discount % (Overall) offered by top 20 vendors (Industry / Region): EMEA": {"wanted": "Avg discount % offered by top X vendors (Top 20 Suppliers)", "region": "EMEA"},

    "Avg discount terms as % of WACC (Industry / Region): Overall": {"wanted": "Avg discount terms as % of WACC", "region": "Overall"},
    "Avg discount terms as % of WACC (Industry / Region): Americas": {"wanted": "Avg discount terms as % of WACC", "region": "Americas"},
    "Avg discount terms as % of WACC (Industry / Region): APAC": {"wanted": "Avg discount terms as % of WACC", "region": "APAC"},
    "Avg discount terms as % of WACC (Industry / Region): EMEA": {"wanted": "Avg discount terms as % of WACC", "region": "EMEA"},
}


# ============================================================
# HELPERS
# ============================================================

def safe_float(x: Any):
    try:
        s = str(x).strip().replace(",", "")
        if s == "":
            return None
        s = re.sub(r"[^\d\.\-]", "", s)
        if s == "":
            return None
        return float(s)
    except Exception:
        return None


def calc_metric(spec: MetricSpec, inputs: Dict[str, Any]):
    outputs = {}

    if spec.calc_type == "identity":
        outputs["Value"] = safe_float(inputs.get(spec.numerator))
        return outputs

    if spec.calc_type == "ratio_pct":
        num = safe_float(inputs.get(spec.numerator))
        den = safe_float(inputs.get(spec.denominator))
        outputs["Value (%)"] = None if num is None or den is None or den == 0 else (num / den) * 100
        return outputs

    if spec.calc_type == "multi_ratio_pct":
        den = safe_float(inputs.get(spec.multi_denominator))
        for n in spec.numerators or []:
            num = safe_float(inputs.get(n))
            outputs[f"{n} (%)"] = None if num is None or den is None or den == 0 else (num / den) * 100
        return outputs

    return outputs


def infer_region(kpi: str) -> str:
    if kpi in KPI_NAMING_MAP:
        return KPI_NAMING_MAP[kpi]["region"]
    for region in ["Overall", "Americas", "APAC", "EMEA"]:
        if f": {region}" in kpi or f"({region})" in kpi:
            return region
    return "Overall"


def infer_metric_name(spec: MetricSpec) -> str:
    if spec.kpi in KPI_NAMING_MAP:
        return KPI_NAMING_MAP[spec.kpi]["wanted"].strip()

    name = spec.kpi
    name = re.sub(r"\s*\(Industry\s*/\s*Region\):\s*(Overall|Americas|APAC|EMEA)$", "", name)
    name = re.sub(r"\s*\(Region:\s*(Overall|Americas|APAC|EMEA)\)$", "", name)
    name = re.sub(r"\s*\(Overall\)\s*\((Overall|Americas|APAC|EMEA)\)$", "", name)
    name = re.sub(r"\s*:\s*(Overall|Americas|APAC|EMEA)$", "", name)
    return name.strip()


def output_metric_name(spec: MetricSpec, out_label: str) -> str:
    full_key = f"{spec.kpi} — {out_label}"
    if full_key in KPI_NAMING_MAP:
        return KPI_NAMING_MAP[full_key]["wanted"].strip()
    if spec.kpi in KPI_NAMING_MAP:
        return KPI_NAMING_MAP[spec.kpi]["wanted"].strip()
    return spec.metric_name or spec.kpi


for spec in AP_METRICS:
    if spec.metric_name is None:
        spec.metric_name = infer_metric_name(spec)
    if spec.region is None:
        spec.region = infer_region(spec.kpi)


# ============================================================
# SESSION STATE
# ============================================================

if "demographics" not in st.session_state:
    st.session_state.demographics = {
        "Company": "",
        "Industry": "",
        "Industry L2": "",
        "Primary Region": "",
        "Currency": "",
        "FY / Period": "",
    }

if "kpi_inputs" not in st.session_state:
    st.session_state.kpi_inputs = {}

if "metric_comments" not in st.session_state:
    st.session_state.metric_comments = {}


# ============================================================
# SIDEBAR
# ============================================================

menu = st.sidebar.radio("Menu", ["Demographics", "KPI Components & Value"])


# ============================================================
# DEMOGRAPHICS PAGE
# ============================================================

if menu == "Demographics":

    st.subheader("Demographics")
    st.caption("Fill in the client details below. These will be included in every row of the exported Excel.")

    col1, col2 = st.columns(2)

    with col1:
        st.session_state.demographics["Company"] = st.text_input(
            "Company", st.session_state.demographics["Company"],
            placeholder="e.g. American Airlines"
        )
        st.session_state.demographics["Industry"] = st.text_input(
            "Industry", st.session_state.demographics["Industry"],
            placeholder="e.g. Advanced Manufacturing & Services"
        )
        st.session_state.demographics["Industry L2"] = st.text_input(
            "Industry L2", st.session_state.demographics["Industry L2"],
            placeholder="e.g. Airlines, Logistics & Transport"
        )

    with col2:
        st.session_state.demographics["Primary Region"] = st.selectbox(
            "Primary Region",
            options=["", "AMER", "APAC", "EMEA", "Global"],
            index=["", "AMER", "APAC", "EMEA", "Global"].index(
                st.session_state.demographics["Primary Region"]
            ) if st.session_state.demographics["Primary Region"] in ["", "AMER", "APAC", "EMEA", "Global"] else 0
        )
        st.session_state.demographics["Currency"] = st.text_input(
            "Currency", st.session_state.demographics["Currency"],
            placeholder="e.g. USD"
        )
        st.session_state.demographics["FY / Period"] = st.text_input(
            "FY / Period", st.session_state.demographics["FY / Period"],
            placeholder="e.g. FY2024"
        )

    st.divider()
    st.markdown("**Preview**")
    preview_df = pd.DataFrame([{
        "Function": "AP",
        "Company": st.session_state.demographics["Company"],
        "Industry": st.session_state.demographics["Industry"],
        "Industry L2": st.session_state.demographics["Industry L2"],
        "Primary Region": st.session_state.demographics["Primary Region"],
        "Currency": st.session_state.demographics["Currency"],
        "FY / Period": st.session_state.demographics["FY / Period"],
    }])
    st.dataframe(preview_df, use_container_width=True, hide_index=True)


# ============================================================
# KPI PAGE
# ============================================================

else:
    export_rows = []
    component_export_rows = []

    metrics_by_lever = defaultdict(list)
    for spec in AP_METRICS:
        metrics_by_lever[spec.lever].append(spec)

    for lever, specs in metrics_by_lever.items():
        st.markdown(f"## {lever}")

        for spec in specs:
            display_label = f"{spec.metric_name}  |  {spec.region}"

            with st.expander(display_label):
                col1, col2, col3 = st.columns(3)
                col1.markdown(f"**Lever:** {spec.lever}")
                col2.markdown(f"**Metric:** {spec.metric_name}")
                col3.markdown(f"**Region:** {spec.region}")

                if spec.notes:
                    st.info(spec.notes)

                st.divider()

                kpi_key = f"{spec.lever}|||{spec.kpi}"

                if kpi_key not in st.session_state.kpi_inputs:
                    st.session_state.kpi_inputs[kpi_key] = {comp: "" for comp in spec.components}

                inputs = {}
                for comp in spec.components:
                    st.session_state.kpi_inputs[kpi_key][comp] = st.text_input(
                        comp,
                        value=st.session_state.kpi_inputs[kpi_key].get(comp, ""),
                        key=f"in::{kpi_key}::{comp}"
                    )
                    inputs[comp] = st.session_state.kpi_inputs[kpi_key][comp]

                if kpi_key not in st.session_state.metric_comments:
                    st.session_state.metric_comments[kpi_key] = ""

                st.session_state.metric_comments[kpi_key] = st.text_area(
                    "Comment",
                    value=st.session_state.metric_comments[kpi_key],
                    key=f"comment::{kpi_key}",
                    placeholder="Describe how you extracted the metric components..."
                )

                outputs = calc_metric(spec, inputs)

                if len(outputs) == 1:
                    for out_label, v in outputs.items():
                        st.text_input(
                            out_label,
                            value="" if v is None else f"{v:,.6f}",
                            disabled=True,
                            key=f"out::{kpi_key}::{out_label}"
                        )
                else:
                    multi_df = pd.DataFrame([
                        {
                            "Output": output_metric_name(spec, out_label),
                            "Value": "" if v is None else f"{v:,.6f}"
                        }
                        for out_label, v in outputs.items()
                    ])
                    st.dataframe(multi_df, use_container_width=True, hide_index=True)

                for out_label, v in outputs.items():
                    if "(%)" in out_label or out_label == "Value (%)":
                        unit = "%"
                    else:
                        unit = "Abs."

                    region_value = KPI_NAMING_MAP.get(
                        f"{spec.kpi} — {out_label}",
                        KPI_NAMING_MAP.get(spec.kpi, {"region": spec.region})
                    )["region"]

                    metric_name = output_metric_name(spec, out_label)
                    comment_value = st.session_state.metric_comments.get(kpi_key, "")

                    export_rows.append({
                        "Function": "AP",
                        "Lever": spec.lever,
                        "Company": st.session_state.demographics.get("Company", ""),
                        "Industry": st.session_state.demographics.get("Industry", ""),
                        "Industry L2": st.session_state.demographics.get("Industry L2", ""),
                        "Primary Region": st.session_state.demographics.get("Primary Region", ""),
                        "Currency": st.session_state.demographics.get("Currency", ""),
                        "FY / Period": st.session_state.demographics.get("FY / Period", ""),
                        "Region": region_value,
                        "KPI": metric_name,
                        "Value": "" if v is None else v,
                        "Unit": unit,
                        "Comment": comment_value,
                    })

                    for comp in spec.components:
                        component_export_rows.append({
                            "Function": "AP",
                            "Lever": spec.lever,
                            "Company": st.session_state.demographics.get("Company", ""),
                            "Industry": st.session_state.demographics.get("Industry", ""),
                            "Industry L2": st.session_state.demographics.get("Industry L2", ""),
                            "Primary Region": st.session_state.demographics.get("Primary Region", ""),
                            "Currency": st.session_state.demographics.get("Currency", ""),
                            "FY / Period": st.session_state.demographics.get("FY / Period", ""),
                            "Region": region_value,
                            "KPI": metric_name,
                            "Component": comp,
                            "Component Value": inputs.get(comp, ""),
                            "Calculated Value": "" if v is None else v,
                            "Unit": unit,
                            "Comment": comment_value,
                        })

    st.markdown("---")
    st.subheader("Export")

    export_df = pd.DataFrame(export_rows, columns=[
        "Function", "Lever", "Company", "Industry", "Industry L2",
        "Primary Region", "Currency", "FY / Period", "Region",
        "KPI", "Value", "Unit", "Comment"
    ])

    component_export_df = pd.DataFrame(component_export_rows, columns=[
        "Function", "Lever", "Company", "Industry", "Industry L2",
        "Primary Region", "Currency", "FY / Period", "Region",
        "KPI", "Component", "Component Value", "Calculated Value",
        "Unit", "Comment"
    ])

    st.markdown("### KPI-Level Export Preview")
    st.dataframe(export_df, use_container_width=True, hide_index=True)

    kpi_excel = io.BytesIO()
    with pd.ExcelWriter(kpi_excel, engine="xlsxwriter") as writer:
        export_df.to_excel(writer, index=False, sheet_name="AP KPIs")
    kpi_excel.seek(0)

    st.download_button(
        "Download KPI Excel",
        data=kpi_excel.getvalue(),
        file_name="ap_kpis_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_ap_excel"
    )

    st.markdown("### Component-Level Export Preview")
    st.dataframe(component_export_df, use_container_width=True, hide_index=True)

    component_excel = io.BytesIO()
    with pd.ExcelWriter(component_excel, engine="xlsxwriter") as writer:
        component_export_df.to_excel(writer, index=False, sheet_name="AP KPI Components")
    component_excel.seek(0)

    st.download_button(
        "Download Component Excel",
        data=component_excel.getvalue(),
        file_name="ap_kpi_components_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="download_ap_component_excel"
    )
