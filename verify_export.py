#!/usr/bin/env python3
"""
SA360 Export Verification Script
=================================
Run this against your raw SA360 export to answer one question:
"Are the campaigns missing from my Non-Testing report actually in the export,
or are they missing from the file entirely?"

Usage:
    streamlit run verify_export.py

Then upload your SA360 export file (same one you use for the WoW report).
"""

import pandas as pd
import streamlit as st
from io import BytesIO

# ── These are the 241 campaigns identified as missing from the Non-Testing output ──
MISSING_CAMPAIGNS = [
    "COXBU_SEM_RTG_NC_Tier1_General_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_PST_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_RTG_NC_Tier1_Internet_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_General_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Fiber_Nonbr_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_TV_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Fiber_Nonbr_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Internet_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Collaboration_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier1_Internet_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_Tier1_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Internet_Brand_Phrase_Google",
    "COXBU_SEM_Employee_NC_IFP_Audience_Brand_Exact Phrase_Google",
    "COXBU_SEM_STD_NC_Tier2_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_Phone/Voice_Nonbr_Phrase_Google_Broad Test 1",
    "COXBU_SEM_RTG_NC_Tier1_TV_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Internet_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Internet_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_Tier2_General_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_TV_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_PST_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier1_Collaboration_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier3_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_Internet_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_General_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Fiber_Nonbr_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Internet_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier3_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_PST_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_TV_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_PST_Brand_Phrase_Google_Broad Test 3",
    "COXBU_SEM_STD_NC_OOF_TV_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_PST_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier1_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Internet_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_OOF_General_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_TV_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier2_Internet_Brand_Phrase_Google Broad AI Max Test",
    "COXBU_SEM_STD_NC_Tier1_Internet_Brand_Phrase_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_IFP_Internet_Nonbr_Exact Broad_Google",
    "COXBU_SEM_STD_NC_Tier2_TV_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_Internet_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_TV_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier1_PST_Brand_Phrase_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_OOF_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_PST_Brand_Phrase_Google Broad AI Max Test 2",
    "COXBU_SEM_STD_NC_OOF_Internet_Brand_Phrase_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_OOF_PST_Brand_Phrase_Google _PricingAdTest1",
    "COXBU_SEM_RTG_NC_Tier2_PST_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_Tier1_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier1_TV_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Internet_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_General_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_Collaboration_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_Collaboration_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_PST_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier2_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Internet_Brand_Phrase_Google_Broad Test 4",
    "COXBU_SEM_STD_NC_OOF_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_PST_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_Tier1_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_General_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier2_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_PST_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_RTG_NC_Tier2_Phone/Voice_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Internet_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier3_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_OOF_TV_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_Internet_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier3_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_PST_Brand_Phrase_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_Tier2_Internet_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_RTG_NC_Tier3_PST_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_OOF_General_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier1_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_PST_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_OOF_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier2_TV_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_Tier3_PST_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_OOF_PST_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_OOF_Collaboration_Brand_Phrase_Google",
    "COXBU_SEM_Employee_NC_IFP_Audience_Nonbr_Exact Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_STD_NC_OOF_Cloud Solutions_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_Net Assurance_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier1_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_RTG_NC_OOF_PST_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_Tier1_Cybersecurity_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier3_TV_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Conquesting_Nonbr_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier1_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_STD_NC_Tier3_Internet_Brand_Exact_Google _PricingAdTest",
    "COXBU_SEM_RTG_NC_Tier3_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_PST_Brand_Phrase_Google _PricingAdTest",
    "COXBU_SEM_STD_NC_Tier2_Phone/Voice_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier3_Internet_Nonbr_Phrase_Google_Broad Test 2",
    "COXBU_SEM_STD_NC_Tier2_Net Assurance_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_Internet_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_STD_NC_Tier2_TV_Brand_Phrase_Google",
    "COXBU_SEM_STD_NC_Tier3_Internet_Brand_Exact_Google _PricingAdTest1",
    "COXBU_SEM_RTG_NC_Tier2_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier2_Collaboration_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier2_Cybersecurity_Brand_Phrase_Google",
    "COXBU_SEM_RTG_NC_Tier3_TV_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_Tier1_Conquesting_Nonbr_Exact_Google",
    "COXBU_SEM_STD_NC_Tier3_Cloud Solutions_Brand_Exact_Google",
    "COXBU_SEM_STD_NC_OOF_Phone/Voice_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier3_Collaboration_Brand_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_TV_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier1_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_TV_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier2_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier1_Collaboration_Brand_Phrase_Bing Old",
    "COXBU_SEM_RTG_NC_Tier1_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier3_PST_Brand_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier3_Internet_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Net Assurance_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Collaboration_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier1_PST_Brand_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Collaboration_Brand_Exact_Bing Old",
    "COXBU_SEM_RTG_NC_Las Vegas_Geo_OOF_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier1_Internet_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier1_Collaboration_Brand_Exact_Bing Old",
    "COXBU_SEM_RTG_NC_Tier2_Net Assurance_Brand_Phrase_Bing Old",
    "COXBU_SEM_RTG_NC_Tier3_Collaboration_Brand_Exact_Bing Old",
    "COXBU_SEM_STD_NC_Tier1_Internet_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_OOF_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Net Assurance_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_Cloud Solutions_Brand_Exact_Bing Old",
    "COXBU_SEM_STD_NC_Tier1_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_Phone/Voice_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier3_Collaboration_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier2_PST_Brand_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Net Assurance_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_PST_Brand_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Phone/Voice_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Phoenix_Geo_OOF_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Collaboration_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_OOF_TV_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Internet_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Las Vegas_Geo_OOF_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier2_Phone/Voice_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Phoenix_Geo_OOF_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Net Assurance_Brand_Exact_Bing Old",
    "COXBU_SEM_STD_NC_OOF_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Collaboration_Brand_Exact_Bing Old",
    "COXBU_SEM_RTG_NC_OOF_Phone/Voice_Brand_Phrase_Bing Old",
    "COXBU_SEM_RTG_NC_Tier1_PST_Brand_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier2_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_PST_Brand_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier3_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier2_PST_Brand_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Collaboration_Brand_Phrase_Bing Old",
    "COXBU_SEM_STD_NC_Tier3_Collaboration_Brand_Phrase_Bing",
    "COXBU_SEM_STD_NC_Phoenix_Geo_OOF_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier1_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Conquesting_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Phone/Voice_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_Collaboration_Brand_Phrase_Bing Old",
    "COXBU_SEM_RTG_NC_OOF_Internet_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier3_Phone/Voice_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier2_Collaboration_Brand_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Internet_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier2_Internet_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Collaboration_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Fiber_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_OOF_Fiber_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier1_Phone/Voice_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier1_Fiber_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_Internet_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier3_TV_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Las Vegas_Geo_OOF_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier3_Conquesting_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Fiber_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_OOF_TV_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_OOF_Collaboration_Brand_Exact_Bing Old",
    "COXBU_SEM_STD_NC_Tier3_Fiber_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Net Assurance_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_OOF_Phone/Voice_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_Phoenix_Geo_OOF_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier2_Collaboration_Brand_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Phone/Voice_Brand_Exact_Bing Old",
    "COXBU_SEM_RTG_NC_Las Vegas_Geo_OOF_Nonbr_Exact_Bing",
    "COXBU_SEM_STD_NC_Tier1_Fiber_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Phoenix_Geo_Tier3_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_Tier3_Internet_Nonbr_Phrase_Bing",
    "COXBU_SEM_RTG_NC_OOF_Fiber_Nonbr_Phrase_Bing",
    "COXBU_SEM_STD_NC_OOF_Collaboration_Nonbr_Exact_Bing",
    "COXBU_SEM_RTG_NC_Tier2_Net Assurance_Brand_Exact_Bing Old",
    "COXBU_SEM_STD_NC_Tier3_Net Assurance_Brand_Phrase_Bing Old",
    "COXBU_SEM_STD_NC_OOF_Phone/Voice_Brand_Exact_Bing Old",
]

COLUMN_ALIASES = {
    'Campaign': ['Campaign', 'Campaign Name', 'campaign'],
    'Week (Mon to Sun)': [
        'Week (Mon to Sun)', 'Week (Mon - Sun)',
        'Week (mon to sun)', 'Week (mon - sun)',
        'Week', 'week',
    ],
    'Cost':   ['Cost', 'Spend', 'cost', 'spend'],
    'Clicks': ['Clicks', 'clicks'],
    'Impr.':  ['Impr.', 'Impressions', 'Impr', 'impressions', 'impr.'],
    'Labels on Campaign: Directly Applied': [
        'Labels on Campaign: Directly Applied',
        'Labels on campaign: Directly applied',
        'Labels', 'Campaign Labels',
    ],
}

def normalize_columns(df):
    df.columns = df.columns.str.strip()
    rename_map = {}
    for std, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in df.columns and alias != std:
                rename_map[alias] = std
                break
    return df.rename(columns=rename_map) if rename_map else df

def load_file(f):
    name = f.name.lower()
    if name.endswith(('.xlsx', '.xls')):
        for skip in [2, 0, 1, 3]:
            try:
                df = pd.read_excel(f, skiprows=skip)
                df = normalize_columns(df)
                if 'Campaign' in df.columns:
                    return df, None
            except Exception:
                pass
            finally:
                f.seek(0)
        return None, "Could not parse Excel file."
    for enc, sep, skip in [
        ('utf-16','\t',2),('utf-8',',',2),('utf-8','\t',2),
        ('utf-8',',',0),('latin-1',',',0),('latin-1','\t',0),
    ]:
        try:
            f.seek(0)
            df = pd.read_csv(f, encoding=enc, sep=sep, skiprows=skip)
            df = normalize_columns(df)
            if 'Campaign' in df.columns:
                return df, None
        except Exception:
            pass
    return None, "Could not parse file."


# ── Streamlit UI ──────────────────────────────────────────────────────────────

st.set_page_config(page_title="Export Verification", layout="wide")
st.title("🔍 SA360 Export Verification")
st.markdown(
    "Upload your SA360 export to verify whether the 241 missing Non-Testing "
    "campaigns are **in the file but filtered out** or **not in the file at all**."
)

uploaded = st.file_uploader("Upload SA360 export (.csv or .xlsx)", type=["csv","xlsx","xls"])

if uploaded:
    df, err = load_file(uploaded)
    if err:
        st.error(err)
        st.stop()

    all_campaigns = set(df['Campaign'].dropna().unique())
    missing_set   = set(MISSING_CAMPAIGNS)

    found_in_file    = missing_set & all_campaigns
    not_in_file      = missing_set - all_campaigns

    st.divider()

    # ── Summary metrics ──
    col1, col2, col3 = st.columns(3)
    col1.metric("Total campaigns in export",    f"{len(all_campaigns):,}")
    col2.metric("Missing campaigns IN the file", f"{len(found_in_file):,}",
                help="In the export but not appearing in Non-Testing output — script filtering issue")
    col3.metric("Missing campaigns NOT in file", f"{len(not_in_file):,}",
                help="Genuinely absent from the export — export scope issue")

    st.divider()

    # ── Verdict ──
    if len(not_in_file) == 0:
        st.success(
            "✅ **All 241 missing campaigns ARE in the export file.** "
            "This is a script filtering problem — the campaigns exist but are being "
            "excluded or missed by the classification/label logic."
        )
    elif len(found_in_file) == 0:
        st.error(
            "❌ **None of the 241 missing campaigns are in the export file.** "
            "This is an export scope problem — the file you're using does not contain "
            "these campaigns at all. You and the team are not using the same export."
        )
    else:
        st.warning(
            f"⚠️ **Mixed result:** {len(found_in_file)} campaigns are in the file "
            f"but being filtered out (script issue), and {len(not_in_file)} are "
            f"genuinely missing from the export (scope issue)."
        )

    st.divider()

    # ── Campaigns found in file — dig into why they're excluded ──
    if found_in_file:
        st.subheader(f"📋 {len(found_in_file)} campaigns in file but missing from output")
        st.caption("These exist in your export but aren't making it into Non-Testing — check their labels and customer type below.")

        found_df = df[df['Campaign'].isin(found_in_file)].copy()
        found_df['Customer Type'] = found_df['Campaign'].apply(
            lambda c: 'CC' if '_CC_' in str(c) else 'NC'
        )
        found_df['Brand/NB'] = found_df['Campaign'].apply(
            lambda c: 'NB' if '_Nonbr_' in str(c) else 'Brand'
        )

        labels_col = 'Labels on Campaign: Directly Applied'
        display_cols = ['Campaign', 'Customer Type', 'Brand/NB']
        if labels_col in found_df.columns:
            display_cols.append(labels_col)

        unique_found = found_df[display_cols].drop_duplicates('Campaign')

        # Show why each is excluded
        nc_specific_labels = [
            '2026 VBB Google Campaigns', 'CBB NB Internet STD Campaigns',
            '2026 UpMarket Campaigns', '2026 CBB NB Remaining Google Campaigns',
            'MSFT NB Max Clicks Campaigns',
        ]
        exclude_patterns = ['discovery', 'master', 'midlife']
        exclude_accounts = ['rapidscale']

        def explain_exclusion(row):
            reasons = []
            if row['Customer Type'] == 'CC':
                reasons.append('classified as CC')
            if labels_col in row.index:
                for lbl in nc_specific_labels:
                    if lbl.lower() in str(row.get(labels_col, '')).lower():
                        reasons.append(f'has label: {lbl}')
            for pat in exclude_patterns:
                if pat.lower() in str(row['Campaign']).lower():
                    reasons.append(f'campaign name contains "{pat}"')
            return '; '.join(reasons) if reasons else 'No exclusion reason found — may be a week filter issue'

        unique_found['Exclusion Reason'] = unique_found.apply(explain_exclusion, axis=1)
        st.dataframe(unique_found, use_container_width=True, hide_index=True)

        # Download
        buf = BytesIO()
        unique_found.to_excel(buf, index=False)
        st.download_button(
            "⬇️ Download found-but-excluded campaigns",
            buf.getvalue(),
            "found_but_excluded.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # ── Campaigns not in file at all ──
    if not_in_file:
        st.subheader(f"❌ {len(not_in_file)} campaigns not in export at all")
        st.caption("These campaigns do not exist anywhere in your export file — export scope mismatch with reference report.")

        # Group by engine
        bing_missing   = [c for c in not_in_file if '_Bing' in c]
        google_missing = [c for c in not_in_file if '_Google' in c]
        other_missing  = [c for c in not_in_file if '_Bing' not in c and '_Google' not in c]

        c1, c2, c3 = st.columns(3)
        c1.metric("Google campaigns missing", len(google_missing))
        c2.metric("Bing campaigns missing",   len(bing_missing))
        c3.metric("Other",                    len(other_missing))

        with st.expander("See full list of missing campaigns"):
            missing_df = pd.DataFrame({'Campaign': sorted(not_in_file)})
            missing_df['Engine'] = missing_df['Campaign'].apply(
                lambda c: 'Google' if '_Google' in c else ('Bing' if '_Bing' in c else 'Other')
            )
            st.dataframe(missing_df, use_container_width=True, hide_index=True)

            buf2 = BytesIO()
            missing_df.to_excel(buf2, index=False)
            st.download_button(
                "⬇️ Download missing campaigns list",
                buf2.getvalue(),
                "not_in_export.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
