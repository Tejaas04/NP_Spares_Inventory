import streamlit as st
import pandas as pd
import openpyxl
import os
import numpy as np
from datetime import datetime, timedelta
from PIL import Image
import json
import hashlib
from cryptography.fernet import Fernet
import base64
import io
import time

# === CONFIGURATION ===
DOCUMENT_FOLDER = "documents"
DATA_FOLDER = "secure_data"
os.makedirs(DOCUMENT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

# === UPDATED PROFESSIONAL LIGHT MODE CSS ===
PROFESSIONAL_CSS = """
<style>
    /* Professional Light Theme Only */
    :root {
        --bg-primary: #f8fafc;
        --bg-secondary: #f1f5f9;
        --bg-card: #ffffff;
        --bg-sidebar: #ffffff;
        --text-primary: #1e293b;
        --text-secondary: #475569;
        --text-muted: #64748b;
        --border-color: #e2e8f0;
        --border-light: #f1f5f9;
        --primary-color: #3b82f6;
        --primary-hover: #2563eb;
        --success-color: #10b981;
        --success-light: #d1fae5;
        --warning-color: #f59e0b;
        --warning-light: #fef3c7;
        --danger-color: #ef4444;
        --danger-light: #fee2e2;
        --info-color: #0ea5e9;
        --info-light: #e0f2fe;
        --shadow-sm: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
        --shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
        --shadow-lg: 0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06);
        --radius-sm: 6px;
        --radius: 8px;
        --radius-lg: 10px;
    }
    
    /* Apply theme to entire app */
    .stApp {
        background-color: var(--bg-primary) !important;
        color: var(--text-primary) !important;
        font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
    }
    
    /* Main Header - LARGER */

    .main-header {
        font-size: 2.2rem !important;
        font-weight: 800 !important;
        color: var(--text-primary) !important;
        text-align: center !important;
        margin-bottom: 1rem !important;
        padding: 1.2rem 0 !important;
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%) !important;
        color: white !important;
        text-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
        border-radius: 0 0 12px 12px;
        box-shadow: var(--shadow-lg) !important;
        letter-spacing: -0.5px;
        margin-top: -1rem !important;
    }
    
    .sub-header {
        font-size: 1.3rem !important;
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        padding: 0.8rem 1.2rem !important;
        margin: 1.2rem 0 0.8rem 0 !important;
        background: var(--bg-card) !important;
        border-radius: var(--radius-sm) !important;
        box-shadow: var(--shadow-sm) !important;
        border-left: 5px solid var(--primary-color) !important;
        border-top: 1px solid var(--border-color) !important;
        border-right: 1px solid var(--border-color) !important;
        border-bottom: 1px solid var(--border-color) !important;
    }
    
    /* Action Buttons - LARGER and more prominent */
    .action-button-row {
        background: var(--bg-card) !important;
        padding: 1rem !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        margin: 1.2rem 0 !important;
        border: 1px solid var(--border-color) !important;
    }
    
    .action-button {
        font-size: 1.1rem !important;
        font-weight: 700 !important;
        padding: 0.8rem 1.2rem !important;
        height: auto !important;
        min-height: 50px !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        gap: 8px !important;
    }
    
    .action-button-primary {
        background: linear-gradient(135deg, var(--primary-color), var(--primary-hover)) !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;

    }
    
    .action-button-primary:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 6px 16px rgba(59, 130, 246, 0.4) !important;
    }
    
    .action-button-secondary {
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
    }
    
    .action-button-secondary:hover {
        background: var(--bg-secondary) !important;
        border-color: var(--primary-color) !important;
        transform: translateY(-2px) !important;
    }
    
    /* Metric Cards */
    .metric-card {
        background: var(--bg-card) !important;
        padding: 1rem !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        margin-bottom: 0.8rem !important;
        border: 1px solid var(--border-color) !important;
        transition: all 0.3s ease;
    }
    
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: var(--shadow-lg) !important;
    }
    
    .metric-compact {
        text-align: center !important;
        padding: 0.8rem !important;
    }
    
    .metric-value {
        font-size: 1.8rem !important;
        font-weight: 800 !important;
        color: var(--text-primary) !important;
        margin: 0.2rem 0 !important;
        line-height: 1.2;
    }
    
    .metric-label {
        font-size: 0.85rem !important;
        color: var(--text-secondary) !important;
        font-weight: 600 !important;
        text-transform: uppercase !important;
        letter-spacing: 1px !important;
    }
    
    /* Sidebar */
    section[data-testid="stSidebar"] {
        background-color: var(--bg-sidebar) !important;
        color: var(--text-primary) !important;

        border-right: 1px solid var(--border-color) !important;
    }
    
    .sidebar-header {
        font-size: 1.1rem !important;
        font-weight: 800 !important;
        color: var(--text-primary) !important;
        text-align: center !important;
        margin-bottom: 1.2rem !important;
        padding: 0.9rem !important;
        background: linear-gradient(135deg, var(--bg-secondary), var(--bg-card)) !important;
        border-radius: var(--radius-sm) !important;
        border: 1px solid var(--border-color) !important;
        box-shadow: var(--shadow-sm) !important;
    }
    
    /* Filter Section - IMPROVED LAYOUT with PROPER ALIGNMENT */
    .filter-section {
        background: var(--bg-card) !important;
        padding: 1.5rem !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        margin-bottom: 1.5rem !important;
        border: 2px solid var(--border-color) !important;
    }
    
    /* FIXED: Selectbox alignment and visibility */
    .stSelectbox [data-baseweb="select"] > div {
        min-width: 180px !important;
        max-width: 100% !important;
        overflow: visible !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
    }
    
    .stSelectbox [data-baseweb="select"] span {
        max-width: 160px !important;
        overflow: hidden !important;
        white-space: nowrap !important;
        text-overflow: ellipsis !important;
        display: inline-block !important;
    }
    
    /* Make selectbox dropdown text visible */
    [data-baseweb="select"] div[role="listbox"] div {
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }
    
    /* Filter grid layout */
    .filter-grid {
        display: grid !important;
        grid-template-columns: repeat(4, 1fr) !important;
        gap: 1.2rem !important;
        align-items: end !important;
    }
    
    /* Fix for selectbox and text input labels */
    .filter-section .stSelectbox label, 

    .filter-section .stTextInput label, 
    .filter-section .stCheckbox label {
        font-size: 0.95rem !important;
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        margin-bottom: 0.6rem !important;
        display: block !important;
    }
    
    .stSelectbox [data-baseweb="select"] > div,
    .stTextInput input,
    .stTextInput input:focus {
        border-radius: 6px !important;
        border: 2px solid var(--border-color) !important;
        font-size: 0.95rem !important;
        padding: 0.6rem 0.8rem !important;
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        min-height: 44px !important;
        width: 100% !important;
    }
    
    .stSelectbox [data-baseweb="select"] > div:hover,
    .stTextInput input:hover {
        border-color: var(--primary-color) !important;
    }
    
    .stCheckbox > label {
        font-size: 0.95rem !important;
        font-weight: 600 !important;
        color: var(--text-primary) !important;
        padding-top: 0.5rem !important;
    }
    
    /* Table Container */
    .table-container {
        background: var(--bg-card) !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        padding: 1.2rem !important;
        margin: 1.2rem 0 !important;
        border: 1px solid var(--border-color) !important;
        font-size: 0.95rem !important;
    }
    
    /* Forms */
    .modern-form {
        background: var(--bg-card) !important;
        padding: 1.5rem !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        margin: 1.2rem 0 !important;
        border: 1px solid var(--border-color) !important;
    }
    
    /* Buttons */
    .stButton button {
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;

        padding: 0.5rem 1rem !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
        font-size: 0.95rem !important;
        margin: 0.2rem !important;
        height: 42px !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton button:hover {
        background: var(--bg-secondary) !important;
        border-color: var(--primary-color) !important;
        transform: translateY(-2px);
    }
    
    /* Primary Button */
    .stButton button[kind="primary"] {
        background: var(--primary-color) !important;
        color: white !important;
        border: none !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
    }
    
    .stButton button[kind="primary"]:hover {
        background: var(--primary-hover) !important;
        transform: translateY(-3px);
        box-shadow: 0 4px 12px rgba(59, 130, 246, 0.3) !important;
    }
    
    /* Input Fields */
    .stTextInput input, 
    .stNumberInput input, 
    .stSelectbox div,
    .stDateInput input,
    .stSelectbox select {
        background-color: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
        border-radius: 6px !important;
        font-size: 0.95rem !important;
        padding: 0.6rem 0.9rem !important;
        transition: border-color 0.3s ease !important;
        min-height: 44px !important;
    }
    
    .stTextInput input:focus, 
    .stNumberInput input:focus, 
    .stSelectbox div:focus,
    .stDateInput input:focus {
        border-color: var(--primary-color) !important;
        box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.1) !important;
    }
    
    /* DataFrames */
    .dataframe {
        font-size: 0.95rem !important;
    }
    
    /* Expanders */

    .streamlit-expanderHeader {
        background: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
        border-radius: 6px !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        margin-bottom: 0.8rem !important;
        padding: 0.8rem 1rem !important;
    }
    
    /* Welcome Message */
    .welcome-message {
        text-align: center !important;
        padding: 3rem !important;
        background: var(--bg-card) !important;
        border-radius: var(--radius-lg) !important;
        box-shadow: var(--shadow) !important;
        border: 1px solid var(--border-color) !important;
        margin: 2rem 0 !important;
    }
    
    /* Copyright Footer */
    .copyright-footer {
        text-align: center !important;
        padding: 1.5rem !important;
        margin-top: 3rem !important;
        color: var(--text-secondary) !important;
        font-size: 0.9rem !important;
        border-top: 2px solid var(--primary-color) !important;
        background: linear-gradient(to right, var(--bg-secondary), var(--bg-card)) !important;
        font-style: italic;
        border-radius: 8px 8px 0 0;
    }
    
    /* Alert/Notification Boxes */
    .stAlert {
        background-color: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 2px solid var(--border-color) !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow-sm) !important;
        font-size: 0.95rem !important;
        padding: 1rem !important;
    }
    
    /* Table cells */
    .table-container td, .table-container th {
        padding: 0.5rem 0.8rem !important;
        font-size: 0.9rem !important;
    }
    
    /* Captions */
    .stCaption {
        color: var(--text-secondary) !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
    }
    
    /* Low stock highlight - Updated for <= 1 */

    .low-stock {
        color: var(--danger-color) !important;
        font-weight: 800 !important;
        background-color: rgba(239, 68, 68, 0.15) !important;
        padding: 4px 10px !important;
        border-radius: 5px !important;
        font-size: 0.9rem !important;
        border: 1px solid rgba(239, 68, 68, 0.3) !important;
    }
    
    /* Tab styling */
    .stTabs [data-baseweb="tab"] {
        font-size: 1rem !important;
        padding: 0.8rem 1.5rem !important;
        font-weight: 600 !important;
    }
    
    /* Financial summary toggle */
    .financial-toggle {
        background: var(--bg-secondary) !important;
        padding: 1rem !important;
        border-radius: 8px !important;
        border: 2px solid var(--border-color) !important;
        margin-bottom: 1.2rem !important;
    }
    
    /* Dashboard section header */
    .dashboard-header {
        font-size: 1.2rem !important;
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        margin: 0.8rem 0 !important;
        padding-bottom: 0.6rem !important;
        border-bottom: 3px solid var(--border-color) !important;
    }
    
    /* Custom styling for form fields to fix labels */
    div[data-testid="stSelectbox"] > label,
    div[data-testid="stTextInput"] > label,
    div[data-testid="stNumberInput"] > label,
    div[data-testid="stDateInput"] > label {
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        font-size: 0.95rem !important;
        margin-bottom: 0.6rem !important;
        display: block !important;
    }
    
    /* Fix for the selectbox dropdown arrow */
    .stSelectbox [data-baseweb="select"] [data-baseweb="select__arrow"] {
        color: var(--text-primary) !important;
    }
    
    /* Custom scrollbar */
    ::-webkit-scrollbar {
        width: 10px;
        height: 10px;
    }
    
    ::-webkit-scrollbar-track {

        background: var(--bg-secondary);
        border-radius: 4px;
    }
    
    ::-webkit-scrollbar-thumb {
        background: var(--border-color);
        border-radius: 4px;
        border: 2px solid var(--bg-secondary);
    }
    
    ::-webkit-scrollbar-thumb:hover {
        background: var(--primary-color);
    }
    
    /* FIXED: Clean group badge without blue background */
    .group-badge {
        background: var(--bg-secondary) !important;
        color: var(--text-secondary) !important;
        padding: 4px 12px !important;
        border-radius: 6px !important;
        font-size: 0.85rem !important;
        font-weight: 600 !important;
        border: 1px solid var(--border-color) !important;
        display: inline-block;
    }
    
    /* Status indicators */
    .status-indicator {
        display: inline-flex;
        align-items: center;
        gap: 5px;
        font-size: 0.85rem;
        font-weight: 600;
    }
    
    .status-dot {
        width: 10px;
        height: 10px;
        border-radius: 50%;
        display: inline-block;
    }
    
    .status-green { background-color: var(--success-color); }
    .status-yellow { background-color: var(--warning-color); }
    .status-red { background-color: var(--danger-color); }
    .status-blue { background-color: var(--primary-color); }
    
    /* Delete confirmation styling */
    .delete-confirmation {
        background: linear-gradient(135deg, #fee2e2, #fecaca) !important;
        border: 2px solid var(--danger-color) !important;
        border-radius: 10px !important;
        padding: 1.5rem !important;
        margin: 1rem 0 !important;
    }
    
    /* Edit form styling */
    .edit-form-container {
        background: linear-gradient(135deg, #f0f9ff, #e0f2fe) !important;
        border: 2px solid var(--info-color) !important;

        border-radius: 10px !important;
        padding: 1.5rem !important;
        margin: 1rem 0 !important;
    }
    
    /* FIXED: Better alignment for filter columns */
    .filter-column {
        margin-bottom: 0 !important;
        padding-bottom: 0 !important;
    }
    
    /* Ensure filter inputs are properly aligned */
    .filter-section > div {
        margin-bottom: 0.5rem !important;
    }
    
    /* Fix for checkbox alignment */
    .filter-section .stCheckbox {
        padding-top: 0.5rem !important;
        margin-top: 0 !important;
    }

    /* ============================================ */
    /* NEW: ITEM CARD VIEW (mobile-friendly list)    */
    /* ============================================ */
    .item-card {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow-sm) !important;
        padding: 0.9rem 1rem !important;
        margin-bottom: 0.8rem !important;
        transition: all 0.25s ease;
    }
    .item-card:hover {
        box-shadow: var(--shadow) !important;
        transform: translateY(-2px);
    }
    .item-card.low-stock-card {
        border-left: 5px solid var(--danger-color) !important;
        background: linear-gradient(135deg, var(--bg-card), var(--danger-light)) !important;
    }
    .item-card-header {
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        margin-bottom: 0.6rem !important;
        gap: 0.5rem;
    }
    .item-card-title {
        font-size: 1.05rem !important;
        font-weight: 700 !important;
        color: var(--text-primary) !important;
        line-height: 1.3;
    }
    .item-card-grid {
        display: grid !important;
        grid-template-columns: repeat(4, 1fr) !important;
        gap: 0.6rem !important;
        text-align: center !important;
    }
    .item-stat-label {
        font-size: 0.68rem !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
        color: var(--text-muted) !important;
        font-weight: 700 !important;
        margin-bottom: 0.15rem !important;
    }
    .item-stat-value {
        font-size: 1rem !important;
        font-weight: 800 !important;
        color: var(--text-primary) !important;
    }
    .item-stat-value.danger { color: var(--danger-color) !important; }
    .item-stat-value.success { color: var(--success-color) !important; }
    .item-card-footer {
        font-size: 0.78rem !important;
        color: var(--text-muted) !important;
        margin-top: 0.6rem !important;
        padding-top: 0.5rem !important;
        border-top: 1px dashed var(--border-color) !important;
        text-align: right !important;
    }

    /* ============================================ */
    /* NEW: ALERT BANNER (low stock notifications)   */
    /* ============================================ */
    .alert-banner {
        background: linear-gradient(135deg, #fff7ed, #fee2e2) !important;
        border: 2px solid var(--danger-color) !important;
        border-radius: var(--radius) !important;
        padding: 1rem 1.2rem !important;
        margin: 1rem 0 !important;
        box-shadow: var(--shadow) !important;
        animation: pulse-border 2.5s ease-in-out infinite;
    }
    @keyframes pulse-border {
        0%, 100% { border-color: var(--danger-color); }
        50% { border-color: #fca5a5; }
    }
    .alert-banner-title {
        font-weight: 800 !important;
        font-size: 1rem !important;
        color: var(--danger-color) !important;
        margin-bottom: 0.4rem !important;
        display: flex !important;
        align-items: center !important;
        gap: 0.4rem !important;
    }
    .alert-chip {
        display: inline-block !important;
        background: var(--bg-card) !important;
        border: 1px solid var(--danger-color) !important;
        color: var(--danger-color) !important;
        font-weight: 700 !important;
        font-size: 0.8rem !important;
        padding: 0.2rem 0.65rem !important;
        border-radius: 999px !important;
        margin: 0.2rem 0.25rem 0 0 !important;
    }

    /* ============================================ */
    /* NEW: QUICK ADJUST +/- BUTTONS                 */
    /* ============================================ */
    .quick-adjust-row .stButton button {
        height: 38px !important;
        min-height: 38px !important;
        font-size: 0.95rem !important;
        font-weight: 800 !important;
        padding: 0.2rem !important;
    }

    /* ============================================ */
    /* NEW: ANALYTICS / INSIGHT CARDS                */
    /* ============================================ */
    .insight-card {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: var(--radius) !important;
        box-shadow: var(--shadow) !important;
        padding: 1rem 1.2rem !important;
        margin-bottom: 1rem !important;
    }
    .insight-title {
        font-weight: 800 !important;
        font-size: 1rem !important;
        color: var(--text-primary) !important;
        margin-bottom: 0.6rem !important;
        display: flex !important;
        align-items: center !important;
        gap: 0.4rem !important;
    }
    .activity-row {
        display: flex !important;
        justify-content: space-between !important;
        align-items: center !important;
        padding: 0.5rem 0.2rem !important;
        border-bottom: 1px solid var(--border-light) !important;
        font-size: 0.88rem !important;
    }
    .activity-row:last-child { border-bottom: none !important; }
    .activity-badge {
        font-size: 0.7rem !important;
        font-weight: 800 !important;
        padding: 0.15rem 0.5rem !important;
        border-radius: 6px !important;
        text-transform: uppercase !important;
        letter-spacing: 0.5px !important;
    }
    .activity-badge.sale { background: var(--success-light) !important; color: var(--success-color) !important; }
    .activity-badge.stock { background: var(--info-light) !important; color: var(--info-color) !important; }

    /* ============================================ */
    /* MOBILE RESPONSIVE OVERHAUL (<= 768px)         */
    /* ============================================ */
    @media (max-width: 768px) {
        .main-header {
            font-size: 1.35rem !important;
            padding: 0.9rem 0.5rem !important;
            border-radius: 0 0 10px 10px;
        }
        .sub-header {
            font-size: 1rem !important;
            padding: 0.6rem 0.9rem !important;
        }
        /* Stack metric cards 2 per row instead of squeezing 5 */
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: wrap !important;
            gap: 0.5rem !important;
        }
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
            min-width: calc(50% - 0.5rem) !important;
            flex: 1 1 calc(50% - 0.5rem) !important;
        }
        .metric-card {
            padding: 0.7rem !important;
            margin-bottom: 0.5rem !important;
        }
        .metric-value {
            font-size: 1.3rem !important;
        }
        .metric-label {
            font-size: 0.62rem !important;
            letter-spacing: 0.5px !important;
        }
        /* Action buttons full width, stacked */
        .action-button {
            font-size: 0.85rem !important;
            min-height: 44px !important;
            padding: 0.5rem 0.6rem !important;
        }
        .stButton button {
            font-size: 0.85rem !important;
            padding: 0.45rem 0.6rem !important;
            height: auto !important;
            min-height: 42px !important;
            white-space: normal !important;
        }
        /* Filter section: single column stacking */
        .filter-section {
            padding: 0.9rem !important;
        }
        .filter-grid {
            grid-template-columns: 1fr !important;
        }
        /* Item card grid: 2 columns on mobile */
        .item-card-grid {
            grid-template-columns: repeat(2, 1fr) !important;
            gap: 0.5rem !important;
        }
        .item-card-title {
            font-size: 0.95rem !important;
        }
        /* Forms: shrink padding */
        .modern-form, .edit-form-container, .delete-confirmation {
            padding: 1rem !important;
        }
        /* Selectbox text doesn't need huge width on mobile */
        .stSelectbox [data-baseweb="select"] > div {
            min-width: 0 !important;
        }
        .stSelectbox [data-baseweb="select"] span {
            max-width: 100% !important;
        }
        .welcome-message {
            padding: 1.5rem 1rem !important;
        }
        .welcome-message > div {
            grid-template-columns: 1fr !important;
        }
        .copyright-footer {
            font-size: 0.78rem !important;
            padding: 1rem !important;
        }
    }
</style>
"""

# === ENHANCED SECURE DATA MANAGEMENT (Supabase-backed, persistent) ===
from supabase import create_client

class SecureDataManager:
    """
    Drop-in replacement for the old file-based encrypted storage.
    Data is now stored as encrypted JSON blobs in a Supabase table
    called 'app_storage' (columns: key text primary key, value text).
    This means data survives app restarts/redeploys (zero data loss),
    unlike local files on free hosting which get wiped.
    """
    def __init__(self):
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        self.client = create_client(url, key)
        self.table = "app_storage"
        self.cipher_suite = self._get_cipher_suite()

    def _get_cipher_suite(self):
        # Encryption key is also stored in Supabase so it's shared
        # across all instances/redeploys of the app.
        existing = self._fetch_raw("encryption_key")
        if existing:
            key = existing.encode()
        else:
            key = Fernet.generate_key()
            self._store_raw("encryption_key", key.decode())
        return Fernet(key)

    def _fetch_raw(self, key):
        try:
            res = self.client.table(self.table).select("value").eq("key", key).execute()
            if res.data:
                return res.data[0]["value"]
            return None
        except Exception as e:
            st.error(f"Error connecting to database: {str(e)}")
            return None

    def _store_raw(self, key, value):
        try:
            self.client.table(self.table).upsert({"key": key, "value": value}).execute()
            return True
        except Exception as e:
            st.error(f"Error saving data: {str(e)}")
            return False

    def save_data(self, data, file_type='inventory'):
        try:
            storage_key = "inventory" if file_type == 'inventory' else "users"
            json_data = json.dumps(data, default=str)
            encrypted_data = self.cipher_suite.encrypt(json_data.encode())
            return self._store_raw(storage_key, base64.b64encode(encrypted_data).decode())
        except Exception as e:
            st.error(f"Error saving data: {str(e)}")
            return False

    def load_data(self, file_type='inventory'):
        try:
            storage_key = "inventory" if file_type == 'inventory' else "users"
            raw = self._fetch_raw(storage_key)

            if raw is None:
                return {} if file_type == 'inventory' else self._create_default_users()

            encrypted_data = base64.b64decode(raw)
            decrypted_data = self.cipher_suite.decrypt(encrypted_data)
            return json.loads(decrypted_data.decode())
        except Exception as e:
            st.error(f"Error loading data: {str(e)}")
            return {} if file_type == 'inventory' else self._create_default_users()

    def _create_default_users(self):
        default_users = {}
        self.save_data(default_users, 'users')
        return default_users

    def _hash_password(self, password):
        return hashlib.sha256(password.encode()).hexdigest()

class UserManager:
    def __init__(self, secure_manager):
        self.secure_manager = secure_manager
        self.users = self.secure_manager.load_data('users')
    
    def login(self, username, password):
        if username in self.users:
            hashed_password = self.secure_manager._hash_password(password)
            if self.users[username]["password"] == hashed_password:
                return True, "Login successful", self.users[username]["role"]
        return False, "Invalid username or password", None
    
    def create_user(self, username, password, role="user"):
        if username in self.users:
            return False, "Username already exists"
        
        self.users[username] = {
            "password": self.secure_manager._hash_password(password),
            "role": role,
            "created_at": datetime.now().isoformat()
        }
        
        if self.secure_manager.save_data(self.users, 'users'):
            return True, "User created successfully"
        return False, "Error creating user"

class InventoryDataManager:
    def __init__(self, secure_manager):
        self.secure_manager = secure_manager
        self.REQUIRED_COLUMNS_STOCK = ["Group", "Item Name", "Quantity", "MRP", "Purchase Price", "Date"]
        self.REQUIRED_COLUMNS_SALES = ["Group", "Item Name", "Sold Quantity", "Sale Price", "Purchase Price", "Date"]
    
    def load_inventory_data(self):
        data = self.secure_manager.load_data('inventory')

        
        if not data:
            data = {
                "stock": [],
                "sales": [],
                "metadata": {
                    "created_at": datetime.now().isoformat(),
                    "last_modified": datetime.now().isoformat(),
                    "version": "1.0"
                }
            }
        
        stock_df = pd.DataFrame(data.get("stock", []))
        sales_df = pd.DataFrame(data.get("sales", []))
        
        stock_df = self._ensure_columns(stock_df, self.REQUIRED_COLUMNS_STOCK)
        sales_df = self._ensure_columns(sales_df, self.REQUIRED_COLUMNS_SALES)
        
        for df in [stock_df, sales_df]:
            if 'Date' in df.columns and not df.empty:
                df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
        
        return stock_df, sales_df, data.get("metadata", {})
    
    def save_inventory_data(self, stock_df, sales_df, metadata=None):
        if metadata is None:
            metadata = {
                "last_modified": datetime.now().isoformat(),
                "modified_by": st.session_state.get('username', 'unknown')
            }
        
        stock_dict = stock_df.to_dict('records')
        sales_dict = sales_df.to_dict('records')
        
        data = {
            "stock": stock_dict,
            "sales": sales_dict,
            "metadata": metadata
        }
        
        return self.secure_manager.save_data(data, 'inventory')
    
    def _ensure_columns(self, df, required_columns):
        for col in required_columns:
            if col not in df.columns:
                if col == "Date":
                    df[col] = datetime.now().date()
                elif col in ["MRP", "Sale Price", "Purchase Price"]:
                    df[col] = 0.0
                else:
                    df[col] = "" if col in ["Group", "Item Name"] else 0
        return df[required_columns]

# === SIMPLE CSV EXPORT ===
def export_group_to_csv(group_data, group_name):
    """Export group data to CSV format"""
    return group_data.to_csv(index=False)

# === EXCEL EXPORT (multi-sheet workbook) ===
def export_to_excel(remaining_df, summary_df=None):
    """Export inventory (and optional summary) to a styled Excel workbook in memory"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        remaining_df.to_excel(writer, index=False, sheet_name='Inventory')
        if summary_df is not None and not summary_df.empty:
            summary_df.to_excel(writer, index=False, sheet_name='Group Summary')
    buffer.seek(0)
    return buffer.getvalue()

# === QUICK STOCK ADJUSTMENT HELPERS (one-tap mobile actions) ===
def quick_add_stock(inventory_manager, group, item_name, purchase_price, mrp):
    """Instantly add 1 unit of stock for an existing item (one-tap restock)"""
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    new_row = pd.DataFrame(
        [[group, item_name, 1, mrp, purchase_price, datetime.now().date()]],
        columns=inventory_manager.REQUIRED_COLUMNS_STOCK
    )
    updated_stock_df = pd.concat([stock_df, new_row], ignore_index=True)
    if inventory_manager.save_inventory_data(updated_stock_df, sales_df):
        st.toast(f"📦 +1 stock added for {item_name}", icon="✅")
        time.sleep(0.3)
        st.rerun()

def quick_record_sale(inventory_manager, group, item_name, purchase_price, mrp):
    """Instantly record a sale of 1 unit at MRP for an existing item (one-tap sale)"""
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    new_sale = pd.DataFrame(
        [[group, item_name, 1, mrp, purchase_price, datetime.now().date()]],
        columns=inventory_manager.REQUIRED_COLUMNS_SALES
    )
    updated_sales_df = pd.concat([sales_df, new_sale], ignore_index=True)

    stock_mask = (stock_df['Group'] == group) & (stock_df['Item Name'] == item_name)
    if not stock_df[stock_mask].empty:
        stock_df.loc[stock_mask, 'Date'] = datetime.now().date()

    if inventory_manager.save_inventory_data(stock_df, updated_sales_df):
        st.toast(f"💰 Sold 1 × {item_name} @ ₹{mrp:.0f}", icon="✅")
        time.sleep(0.3)
        st.rerun()

# === PROFESSIONAL LOGIN PAGE ===
def show_login_page(user_manager):

    st.markdown(PROFESSIONAL_CSS, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown(f"""
            <div style='background: var(--bg-card); padding: 2.5rem; border-radius: 12px; box-shadow: var(--shadow-lg); border: 2px solid var(--border-color);'>
        """, unsafe_allow_html=True)
        
        st.markdown("""
            <div style='text-align: center; margin-bottom: 2rem;'>
                <h1 style='color: var(--text-primary); margin-bottom: 0.5rem; font-size: 2rem; font-weight: 800;'>🔧 N.P. SPARES</h1>
                <p style='color: var(--text-secondary); font-size: 1rem; font-weight: 500;'>Inventory Management System</p>
            </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["🔐 Login", "👤 Create Account"])
        
        with tab1:
            with st.form("login_form"):
                st.markdown('<div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1rem;">Login to Your Account</div>', unsafe_allow_html=True)
                
                username = st.text_input("Username", placeholder="Enter your username", key="login_username")
                password = st.text_input("Password", type="password", placeholder="Enter your password", key="login_password")
                
                col_btn1, col_btn2 = st.columns([2, 1])
                with col_btn1:
                    if st.form_submit_button("Login", use_container_width=True, type="primary"):
                        if username and password:
                            with st.spinner("Logging in..."):
                                success, message, role = user_manager.login(username, password)
                                if success:
                                    st.session_state.authenticated = True
                                    st.session_state.username = username
                                    st.session_state.user_role = role
                                    st.success(f"✅ Welcome, {username}!")
                                    time.sleep(0.5)
                                    st.rerun()
                                else:
                                    st.error(f"❌ {message}")
                        else:
                            st.warning("⚠️ Please enter both fields")
        
        with tab2:
            with st.form("signup_form"):
                st.markdown('<div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1rem;">Create New Account</div>', unsafe_allow_html=True)
                
                new_username = st.text_input("Username", placeholder="Enter username", key="signup_username")
                new_password = st.text_input("Password", type="password", placeholder="Enter password", key="signup_password")
                confirm_password = st.text_input("Confirm Password", type="password", placeholder="Confirm password", key="confirm_password")
                
                if st.form_submit_button("Create Account", use_container_width=True, type="primary"):
                    if not new_username or not new_password:
                        st.error("❌ Username and password required!")
                    elif new_password != confirm_password:

                        st.error("❌ Passwords don't match!")
                    elif len(new_username) < 3:
                        st.error("❌ Username must be 3+ characters!")
                    elif len(new_password) < 6:
                        st.error("❌ Password must be 6+ characters!")
                    else:
                        with st.spinner("Creating account..."):
                            success, message = user_manager.create_user(new_username, new_password)
                            if success:
                                st.success("✅ Account created! Please login.")
                                time.sleep(1)
                                st.rerun()
                            else:
                                st.error(f"❌ {message}")
        
        st.markdown("</div>", unsafe_allow_html=True)

# === ENHANCED DASHBOARD METRICS ===
def show_dashboard_metrics(stock_df, sales_df, remaining_df):
    """Show dashboard metrics without financial data toggle"""
    
    # Basic metrics
    total_groups = stock_df['Group'].nunique() if not stock_df.empty else 0
    unique_items = stock_df['Item Name'].nunique() if not stock_df.empty else 0
    total_active_items = remaining_df['Remaining'].sum() if not remaining_df.empty else 0
    total_sales_qty = sales_df['Sold Quantity'].sum() if not sales_df.empty else 0
    
    # Updated low stock count (<= 1)
    low_stock_count = len(remaining_df[remaining_df['Remaining'] <= 1]) if not remaining_df.empty else 0
    
    # Row 1: Basic Metrics
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
        st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--primary-color);">🏷️</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{total_groups}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">TOTAL GROUPS</div>', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)
    
    with col2:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
        st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--success-color);">📦</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{unique_items}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">UNIQUE ITEMS</div>', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)
    
    with col3:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
        st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--info-color);">📊</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{total_active_items}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">IN STOCK</div>', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)
    

    with col4:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
        st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--warning-color);">💰</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value">{total_sales_qty}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">ITEMS SOLD</div>', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)
    
    with col5:
        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
        st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
        st.markdown(f'<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--danger-color)">⚠️</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="metric-value" style="color: var(--danger-color)">{low_stock_count}</div>', unsafe_allow_html=True)
        st.markdown('<div class="metric-label">LOW STOCK ITEMS</div>', unsafe_allow_html=True)
        st.markdown('</div></div>', unsafe_allow_html=True)

# === FAST EDIT FORM WITH DATE UPDATE ===
def show_edit_form(inventory_manager, group, item_name):
    """Show fast edit form with preserved filters and date update"""
    
    # Preserve filter states before editing
    if 'filter_state' not in st.session_state:
        st.session_state.filter_state = {
            'selected_group': st.session_state.get('selected_group_filter', "All Groups"),
            'search_term': st.session_state.get('search_term_filter', ""),
            'low_stock': st.session_state.get('low_stock_filter', False),
            'sort_by': st.session_state.get('sort_by_filter', "Item Name")
        }
    
    st.markdown('<div class="edit-form-container">', unsafe_allow_html=True)
    
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    
    item_mask = (stock_df['Group'] == group) & (stock_df['Item Name'] == item_name)
    if not stock_df[item_mask].empty:
        item_data = stock_df[item_mask].iloc[0]
        
        st.markdown(f'<div style="font-size: 1.4rem; font-weight: 700; margin-bottom: 1.5rem; color: var(--text-primary);">✏️ Edit Item: {item_name}</div>', unsafe_allow_html=True)
        
        with st.form(f"edit_form_{group}_{item_name}"):
            col1, col2 = st.columns(2)
            
            with col1:
                st.text_input("Group", value=item_data['Group'], disabled=True, key=f"group_{group}_{item_name}")
                new_item_name = st.text_input("Item Name", value=item_data['Item Name'], key=f"item_{group}_{item_name}")
                new_quantity = st.number_input("Quantity", min_value=0, value=int(item_data['Quantity']), key=f"qty_{group}_{item_name}")
            
            with col2:
                new_mrp = st.number_input("MRP (₹)", min_value=0.0, value=float(item_data.get('MRP', 0.0)), step=0.01, key=f"mrp_{group}_{item_name}")
                new_purchase_price = st.number_input("Purchase Price (₹)", min_value=0.0, 
                                                     value=float(item_data.get('Purchase Price', 0.0)), step=0.01, key=f"price_{group}_{item_name}")
                new_date = st.date_input("Date", value=datetime.now().date(), key=f"date_{group}_{item_name}")
            

            col_btn1, col_btn2 = st.columns(2)
            
            with col_btn1:
                save_clicked = st.form_submit_button("💾 Save Changes", use_container_width=True, type="primary")
            
            with col_btn2:
                cancel_clicked = st.form_submit_button("❌ Cancel", use_container_width=True)
            
            if save_clicked:
                if new_item_name and new_item_name.strip():
                    stock_df.loc[item_mask, 'Item Name'] = new_item_name
                    stock_df.loc[item_mask, 'Quantity'] = new_quantity
                    stock_df.loc[item_mask, 'MRP'] = new_mrp
                    stock_df.loc[item_mask, 'Purchase Price'] = new_purchase_price
                    stock_df.loc[item_mask, 'Date'] = new_date  # Update date to current date
                    
                    if new_item_name != item_name:
                        sales_mask = (sales_df['Group'] == group) & (sales_df['Item Name'] == item_name)
                        if not sales_df[sales_mask].empty:
                            sales_df.loc[sales_mask, 'Item Name'] = new_item_name
                    
                    if inventory_manager.save_inventory_data(stock_df, sales_df):
                        st.success("✅ Item updated successfully!")
                        time.sleep(0.5)
                        
                        # Restore filter state
                        if 'filter_state' in st.session_state:
                            st.session_state.selected_group_filter = st.session_state.filter_state['selected_group']
                            st.session_state.search_term_filter = st.session_state.filter_state['search_term']
                            st.session_state.low_stock_filter = st.session_state.filter_state['low_stock']
                            st.session_state.sort_by_filter = st.session_state.filter_state['sort_by']
                            del st.session_state.filter_state
                        
                        st.session_state.editing_item = None
                        st.rerun()
                else:
                    st.error("❌ Item name required!")
            
            if cancel_clicked:
                # Restore filter state
                if 'filter_state' in st.session_state:
                    st.session_state.selected_group_filter = st.session_state.filter_state['selected_group']
                    st.session_state.search_term_filter = st.session_state.filter_state['search_term']
                    st.session_state.low_stock_filter = st.session_state.filter_state['low_stock']
                    st.session_state.sort_by_filter = st.session_state.filter_state['sort_by']
                    del st.session_state.filter_state
                
                st.session_state.editing_item = None
                st.rerun()
    else:
        st.error("Item not found!")
        if st.button("← Back to Inventory", use_container_width=True):
            st.session_state.editing_item = None
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# === DELETE CONFIRMATION (FIXED) ===
def show_delete_confirmation(inventory_manager, group, item_name):
    """Show delete confirmation with proper state management"""

    
    # Preserve filter states before deleting
    if 'delete_filter_state' not in st.session_state:
        st.session_state.delete_filter_state = {
            'selected_group': st.session_state.get('selected_group_filter', "All Groups"),
            'search_term': st.session_state.get('search_term_filter', ""),
            'low_stock': st.session_state.get('low_stock_filter', False),
            'sort_by': st.session_state.get('sort_by_filter', "Item Name")
        }
    
    st.markdown('<div class="delete-confirmation">', unsafe_allow_html=True)
    
    st.markdown('<h3 style="color: var(--danger-color); margin-bottom: 1rem;">🗑️ DELETE CONFIRMATION</h3>', unsafe_allow_html=True)
    st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600;">Item: <span style="color: var(--text-primary);">{item_name}</span></p>', unsafe_allow_html=True)
    st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600;">Group: <span style="color: var(--text-primary);">{group}</span></p>', unsafe_allow_html=True)
    
    st.warning("⚠️ This action will permanently delete ALL stock and sales data for this item!")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("✅ CONFIRM DELETE", use_container_width=True, type="primary", key="confirm_delete_btn"):
            stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
            
            # Delete from stock
            stock_mask = ~((stock_df['Group'] == group) & (stock_df['Item Name'] == item_name))
            stock_df = stock_df[stock_mask]
            
            # Delete from sales
            sales_mask = ~((sales_df['Group'] == group) & (sales_df['Item Name'] == item_name))
            sales_df = sales_df[sales_mask]
            
            if inventory_manager.save_inventory_data(stock_df, sales_df):
                st.success("✅ Item deleted successfully!")
                time.sleep(0.5)
                
                # Restore filter state
                if 'delete_filter_state' in st.session_state:
                    st.session_state.selected_group_filter = st.session_state.delete_filter_state['selected_group']
                    st.session_state.search_term_filter = st.session_state.delete_filter_state['search_term']
                    st.session_state.low_stock_filter = st.session_state.delete_filter_state['low_stock']
                    st.session_state.sort_by_filter = st.session_state.delete_filter_state['sort_by']
                    del st.session_state.delete_filter_state
                
                st.session_state.deleting_item = None
                st.rerun()
    
    with col2:
        if st.button("❌ CANCEL", use_container_width=True, key="cancel_delete_btn"):
            # Restore filter state
            if 'delete_filter_state' in st.session_state:
                st.session_state.selected_group_filter = st.session_state.delete_filter_state['selected_group']
                st.session_state.search_term_filter = st.session_state.delete_filter_state['search_term']
                st.session_state.low_stock_filter = st.session_state.delete_filter_state['low_stock']
                st.session_state.sort_by_filter = st.session_state.delete_filter_state['sort_by']
                del st.session_state.delete_filter_state
            

            st.session_state.deleting_item = None
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)

# === TOTAL INVENTORY VIEW ===
def show_total_inventory(inventory_manager):
    """Show total inventory with group-wise download options"""
    st.markdown('<div class="sub-header">📊 Total Inventory</div>', unsafe_allow_html=True)
    
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    
    def calculate_remaining(stock_df, sales_df):
        if stock_df.empty:
            return pd.DataFrame(columns=['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 'Sold Quantity', 'Remaining', 'Date'])
        
        stock_total = stock_df.groupby(['Group', 'Item Name']).agg({
            'Quantity': 'sum',
            'MRP': 'first',
            'Purchase Price': 'first',
            'Date': 'max'
        }).reset_index()
        
        if sales_df.empty:
            sales_total = pd.DataFrame(columns=['Group', 'Item Name', 'Sold Quantity'])
            sales_total['Sold Quantity'] = sales_total['Sold Quantity'].astype(int)
        else:
            sales_total = sales_df.groupby(['Group', 'Item Name'])['Sold Quantity'].sum().reset_index()
        
        merged = pd.merge(stock_total, sales_total, on=['Group', 'Item Name'], how='left')
        merged['Sold Quantity'] = merged['Sold Quantity'].fillna(0)
        merged['Remaining'] = merged['Quantity'] - merged['Sold Quantity']
        merged['Low Stock'] = merged['Remaining'] <= 1  # Updated to <= 1
        
        result = merged[['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 
                         'Sold Quantity', 'Remaining', 'Low Stock', 'Date']]
        
        for col in ['Quantity', 'Sold Quantity', 'Remaining']:
            if col in result.columns:
                result[col] = result[col].astype('int32')
        
        return result
    
    remaining_df = calculate_remaining(stock_df, sales_df)
    
    if remaining_df.empty:
        st.info("📭 No inventory data available.")
        return
    
    # Group-wise download section
    st.markdown('<div class="filter-section">', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1rem;">📥 Download Options</div>', unsafe_allow_html=True)
    
    groups = sorted(remaining_df['Group'].unique().tolist())
    
    if groups:
        cols = st.columns(3)
        for idx, group in enumerate(groups):

            col_idx = idx % 3
            with cols[col_idx]:
                group_data = remaining_df[remaining_df['Group'] == group]
                csv_data = export_group_to_csv(group_data, group)
                st.download_button(
                    label=f"📄 {group}",
                    data=csv_data,
                    file_name=f"inventory_{group}_{datetime.now().strftime('%Y%m%d')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    
    st.markdown("---")

    # Pre-compute group summary (used for Excel export and the table below)
    summary = pd.DataFrame()
    if not remaining_df.empty:
        summary = remaining_df.groupby('Group').agg({
            'Item Name': 'count',
            'Quantity': 'sum',
            'Sold Quantity': 'sum',
            'Remaining': 'sum',
            'Low Stock': 'sum'
        }).rename(columns={
            'Item Name': 'Items',
            'Low Stock': 'Low Stock'
        }).reset_index()
        summary = summary.sort_values('Items', ascending=False)

    col1, col2 = st.columns(2)
    with col1:
        all_csv_data = remaining_df.to_csv(index=False)
        st.download_button(
            label="📥 Download All Data (CSV)",
            data=all_csv_data,
            file_name=f"inventory_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )
    with col2:
        excel_data = export_to_excel(remaining_df, summary)
        st.download_button(
            label="📊 Download All Data (Excel)",
            data=excel_data,
            file_name=f"inventory_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Show summary statistics
    st.markdown('<div class="table-container">', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1rem;">📈 Inventory Summary by Group</div>', unsafe_allow_html=True)
    
    if not summary.empty:
        st.dataframe(summary, use_container_width=True, hide_index=True)
    
    st.markdown('</div>', unsafe_allow_html=True)

# === FIXED INVENTORY OVERVIEW WITH CLEAN FILTERS AND DATE UPDATES ===
def show_inventory_overview(inventory_manager):
    """Show clean inventory overview with properly aligned filters and clean group badges"""
    st.markdown('<div class="sub-header">📦 Inventory Overview</div>', unsafe_allow_html=True)
    
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    
    def calculate_remaining(stock_df, sales_df):
        if stock_df.empty:
            return pd.DataFrame(columns=['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 'Sold Quantity', 'Remaining', 'Date'])
        

        stock_total = stock_df.groupby(['Group', 'Item Name']).agg({
            'Quantity': 'sum',
            'MRP': 'first',
            'Purchase Price': 'first',
            'Date': 'max'
        }).reset_index()
        
        if sales_df.empty:
            sales_total = pd.DataFrame(columns=['Group', 'Item Name', 'Sold Quantity'])
            sales_total['Sold Quantity'] = sales_total['Sold Quantity'].astype(int)
        else:
            sales_total = sales_df.groupby(['Group', 'Item Name'])['Sold Quantity'].sum().reset_index()
        
        merged = pd.merge(stock_total, sales_total, on=['Group', 'Item Name'], how='left')
        merged['Sold Quantity'] = merged['Sold Quantity'].fillna(0)
        merged['Remaining'] = merged['Quantity'] - merged['Sold Quantity']
        
        # Get latest date from sales for each item
        if not sales_df.empty:
            latest_sale_dates = sales_df.groupby(['Group', 'Item Name'])['Date'].max().reset_index()
            latest_sale_dates = latest_sale_dates.rename(columns={'Date': 'Latest_Sale_Date'})
            merged = pd.merge(merged, latest_sale_dates, on=['Group', 'Item Name'], how='left')
            
            # Use the latest date between stock date and sale date
            merged['Date'] = merged.apply(
                lambda row: max(row['Date'], row['Latest_Sale_Date']) if pd.notna(row['Latest_Sale_Date']) else row['Date'], 
                axis=1
            )
            merged = merged.drop('Latest_Sale_Date', axis=1)
        
        result = merged[['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 
                         'Sold Quantity', 'Remaining', 'Date']]
        
        for col in ['Quantity', 'Sold Quantity', 'Remaining']:
            if col in result.columns:
                result[col] = result[col].astype('int32')
        
        return result
    
    remaining_df = calculate_remaining(stock_df, sales_df)
    
    if remaining_df.empty:
        st.info("📭 No inventory data available.")
        return
    
    # === CLEAN FILTER SECTION - Properly aligned ===
    st.markdown('<div class="filter-section">', unsafe_allow_html=True)
    st.markdown('<div style="font-size: 1.1rem; font-weight: 600; margin-bottom: 1.2rem; color: var(--text-primary);">🔍 FILTER OPTIONS</div>', unsafe_allow_html=True)
    
    # Initialize filter states if not exists
    if 'selected_group_filter' not in st.session_state:
        st.session_state.selected_group_filter = "All Groups"
    if 'search_term_filter' not in st.session_state:
        st.session_state.search_term_filter = ""
    if 'low_stock_filter' not in st.session_state:
        st.session_state.low_stock_filter = False
    if 'sort_by_filter' not in st.session_state:
        st.session_state.sort_by_filter = "Item Name"
    

    # FIXED: Create clean filter layout with proper alignment
    # Use a 2x2 grid for better alignment
    row1_col1, row1_col2 = st.columns(2)
    row2_col1, row2_col2 = st.columns(2)
    
    with row1_col1:
        groups = ["All Groups"] + sorted(remaining_df['Group'].unique().tolist())
        
        # Get current index for selected group
        current_group = st.session_state.selected_group_filter
        group_index = 0
        if current_group in groups:
            group_index = groups.index(current_group)
        
        selected_group = st.selectbox(
            "**Filter by Group**",
            groups,
            index=group_index,
            key="group_filter_selectbox",
            help="Select a group to filter items"
        )
        
        # Update session state
        if selected_group != st.session_state.selected_group_filter:
            st.session_state.selected_group_filter = selected_group
    
    with row1_col2:
        sort_options = ["Item Name", "Remaining", "Sold Quantity", "Group", "Date"]
        
        # Get current index for sort by
        current_sort = st.session_state.sort_by_filter
        sort_index = 0
        if current_sort in sort_options:
            sort_index = sort_options.index(current_sort)
        
        sort_by = st.selectbox(
            "**Sort By**", 
            sort_options,
            index=sort_index,
            key="overview_sort_by_selectbox",
            help="Select sorting criteria"
        )
        
        if sort_by != st.session_state.sort_by_filter:
            st.session_state.sort_by_filter = sort_by
    
    with row2_col1:
        search_term = st.text_input(
            "**Search Items**", 
            value=st.session_state.search_term_filter,
            placeholder="Enter item name...", 
            key="overview_search_input",
            help="Search items by name"
        )
        
        if search_term != st.session_state.search_term_filter:
            st.session_state.search_term_filter = search_term
    
    with row2_col2:
        low_stock_only = st.checkbox(

            "**Show Low Stock Only**", 
            value=st.session_state.low_stock_filter,
            key="overview_low_stock_checkbox",
            help="Show items with stock ≤ 1"
        )
        
        if low_stock_only != st.session_state.low_stock_filter:
            st.session_state.low_stock_filter = low_stock_only
    
    # Clear filters button - centered
    st.markdown("---")
    col_clear = st.columns([1, 2, 1])
    with col_clear[1]:
        if st.button("🗑️ Clear All Filters", use_container_width=True, key="clear_filters_btn"):
            st.session_state.selected_group_filter = "All Groups"
            st.session_state.search_term_filter = ""
            st.session_state.low_stock_filter = False
            st.session_state.sort_by_filter = "Item Name"
            st.rerun()
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Filter data using current filter states
    display_df = remaining_df.copy()
    
    # Apply group filter
    if st.session_state.selected_group_filter != "All Groups":
        display_df = display_df[display_df['Group'] == st.session_state.selected_group_filter]
    
    # Apply search filter
    if st.session_state.search_term_filter:
        display_df = display_df[display_df['Item Name'].str.contains(
            st.session_state.search_term_filter, case=False, na=False
        )]
    
    # Apply low stock filter
    if st.session_state.low_stock_filter:
        display_df = display_df[display_df['Remaining'] <= 1]
    
    # Apply sorting
    if st.session_state.sort_by_filter:
        ascending = True if st.session_state.sort_by_filter in ["Item Name", "Group"] else False
        display_df = display_df.sort_values(by=st.session_state.sort_by_filter, ascending=ascending)
    
    if display_df.empty:
        st.info("📭 No items match your filters")
        return

    st.caption(f"📊 Showing {len(display_df):,} of {len(remaining_df):,} items")

    # Display inventory as mobile-friendly cards
    for idx, row in display_df.iterrows():
        remaining = int(row['Remaining'])
        is_low = remaining <= 1
        date_str = row['Date'].strftime('%d-%m-%Y') if pd.notna(row['Date']) else 'N/A'
        remaining_class = "danger" if is_low else "success"
        card_class = "item-card low-stock-card" if is_low else "item-card"
        footer_extra = " &nbsp;•&nbsp; ⚠️ <strong>Low Stock!</strong>" if is_low else ""

        st.markdown(f"""
            <div class="{card_class}">
                <div class="item-card-header">
                    <span class="item-card-title">{row['Item Name']}</span>
                    <span class="group-badge">{row['Group']}</span>
                </div>
                <div class="item-card-grid">
                    <div>
                        <div class="item-stat-label">Stock</div>
                        <div class="item-stat-value">{int(row['Quantity'])}</div>
                    </div>
                    <div>
                        <div class="item-stat-label">Sold</div>
                        <div class="item-stat-value">{int(row['Sold Quantity'])}</div>
                    </div>
                    <div>
                        <div class="item-stat-label">Remaining</div>
                        <div class="item-stat-value {remaining_class}">{remaining}</div>
                    </div>
                    <div>
                        <div class="item-stat-label">Cost / MRP</div>
                        <div class="item-stat-value">₹{row['Purchase Price']:.0f} / ₹{row['MRP']:.0f}</div>
                    </div>
                </div>
                <div class="item-card-footer">🕒 Updated {date_str}{footer_extra}</div>
            </div>
        """, unsafe_allow_html=True)

        btn_col1, btn_col2, btn_col3, btn_col4 = st.columns(4)
        with btn_col1:
            if st.button("➕ Stock", key=f"qadd_{row['Group']}_{row['Item Name']}_{idx}",
                       help="One-tap: add 1 unit to stock", use_container_width=True):
                quick_add_stock(inventory_manager, row['Group'], row['Item Name'],
                                row['Purchase Price'], row['MRP'])
        with btn_col2:
            if st.button("➖ Sale", key=f"qsale_{row['Group']}_{row['Item Name']}_{idx}",
                       help="One-tap: record a sale of 1 unit at MRP", use_container_width=True,
                       disabled=(remaining <= 0)):
                quick_record_sale(inventory_manager, row['Group'], row['Item Name'],
                                  row['Purchase Price'], row['MRP'])
        with btn_col3:
            if st.button("✏️ Edit", key=f"edit_{row['Group']}_{row['Item Name']}_{idx}",
                       help="Edit Item", use_container_width=True):
                st.session_state.editing_item = (row['Group'], row['Item Name'])
        with btn_col4:
            if st.button("🗑️ Delete", key=f"delete_{row['Group']}_{row['Item Name']}_{idx}",
                       help="Delete Item", use_container_width=True):
                st.session_state.deleting_item = (row['Group'], row['Item Name'])

    # Summary statistics
    total_items = len(display_df)
    total_remaining = display_df['Remaining'].sum()
    low_stock_count = len(display_df[display_df['Remaining'] <= 1])
    
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Items", f"{total_items:,}", delta=None)
    with col2:
        st.metric("Total Stock", f"{total_remaining:,}", delta=None)
    with col3:
        st.metric("Low Stock Items", f"{low_stock_count:,}", delta=None, delta_color="inverse")
    
    st.caption(f"📊 Showing {len(display_df):,} of {len(remaining_df):,} items")

# === WELCOME MESSAGE ===
def show_welcome_message():
    """Show a clean welcome message"""
    st.markdown("""
        <div class="welcome-message">
            <h3 style="color: var(--text-primary); margin-bottom: 1rem; font-size: 1.8rem; font-weight: 800;">🔧 Welcome to N.P. SPARES</h3>
            <p style="color: var(--text-secondary); margin-bottom: 1.5rem; font-size: 1.1rem; font-weight: 500;">Professional Inventory Management System</p>
            <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 1.2rem; text-align: left;">
                <div style="background: var(--bg-secondary); padding: 1.2rem; border-radius: 8px; border: 2px solid var(--border-color);">
                    <h4 style="color: var(--text-primary); margin-bottom: 0.8rem; font-size: 1.1rem; font-weight: 700;">📦 Quick Actions</h4>
                    <ul style="color: var(--text-secondary); font-size: 0.95rem; margin: 0; padding-left: 1.2rem;">
                        <li>Add new stock items</li>
                        <li>Record sales transactions</li>
                        <li>One-tap restock & quick sale buttons</li>
                        <li>View inventory overview</li>
                        <li>Export inventory data (CSV & Excel)</li>
                    </ul>
                </div>
                <div style="background: var(--bg-secondary); padding: 1.2rem; border-radius: 8px; border: 2px solid var(--border-color);">
                    <h4 style="color: var(--text-primary); margin-bottom: 0.8rem; font-size: 1.1rem; font-weight: 700;">📊 Dashboard Features</h4>
                    <ul style="color: var(--text-secondary); font-size: 0.95rem; margin: 0; padding-left: 1.2rem;">
                        <li>Real-time stock tracking</li>
                        <li>Low stock alerts</li>
                        <li>Revenue & profit trend charts</li>
                        <li>Top-selling items & recent activity</li>
                        <li>Group-wise management</li>
                        <li>Secure cloud data storage</li>
                    </ul>
                </div>
            </div>
            <p style="color: var(--text-muted); font-size: 0.9rem; margin-top: 1.5rem;">
                Use the action buttons above to get started with inventory management
            </p>
        </div>
    """, unsafe_allow_html=True)

# === NEW: ANALYTICS & INSIGHTS DASHBOARD ===
def show_analytics(inventory_manager):
    """Show charts for sales trends, top sellers, stock distribution and recent activity"""
    st.markdown('<div class="sub-header">📈 Analytics & Insights</div>', unsafe_allow_html=True)

    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()

    if stock_df.empty and sales_df.empty:
        st.info("📭 No data available yet. Add some stock or record a sale to see analytics here.")
        return

    # --- Sales Trend (Revenue over time) ---
    st.markdown('<div class="insight-card">', unsafe_allow_html=True)
    st.markdown('<div class="insight-title">📈 Revenue Trend</div>', unsafe_allow_html=True)
    if not sales_df.empty:
        sales_trend = sales_df.copy()
        sales_trend['Revenue'] = sales_trend['Sold Quantity'] * sales_trend['Sale Price']
        sales_trend['Profit'] = sales_trend['Sold Quantity'] * (sales_trend['Sale Price'] - sales_trend['Purchase Price'])
        trend = sales_trend.groupby('Date')[['Revenue', 'Profit']].sum().reset_index()
        trend = trend.sort_values('Date').set_index('Date')
        st.line_chart(trend, use_container_width=True)
        st.caption("Daily revenue (₹) and profit (₹) from recorded sales")
    else:
        st.info("No sales recorded yet — record a sale to see your revenue trend.")
    st.markdown('</div>', unsafe_allow_html=True)

    # --- Top Sellers & Stock Value by Group ---
    chart_col1, chart_col2 = st.columns(2)

    with chart_col1:
        st.markdown('<div class="insight-card">', unsafe_allow_html=True)
        st.markdown('<div class="insight-title">🏆 Top Selling Items</div>', unsafe_allow_html=True)
        if not sales_df.empty:
            top_items = sales_df.groupby('Item Name')['Sold Quantity'].sum().sort_values(ascending=False).head(8)
            st.bar_chart(top_items, use_container_width=True)
        else:
            st.info("No sales data yet.")
        st.markdown('</div>', unsafe_allow_html=True)

    with chart_col2:
        st.markdown('<div class="insight-card">', unsafe_allow_html=True)
        st.markdown('<div class="insight-title">📦 Stock Value by Group</div>', unsafe_allow_html=True)
        if not stock_df.empty:
            stock_value = stock_df.copy()
            stock_value['Value'] = stock_value['Quantity'] * stock_value['Purchase Price']
            group_value = stock_value.groupby('Group')['Value'].sum().sort_values(ascending=False)
            st.bar_chart(group_value, use_container_width=True)
        else:
            st.info("No stock data yet.")
        st.markdown('</div>', unsafe_allow_html=True)

    # --- Recent Activity Feed ---
    st.markdown('<div class="insight-card">', unsafe_allow_html=True)
    st.markdown('<div class="insight-title">🕒 Recent Activity</div>', unsafe_allow_html=True)

    activities = []
    for _, r in stock_df.iterrows():
        activities.append({
            'Date': r['Date'],
            'Type': 'stock',
            'Text': f"➕ Added {int(r['Quantity'])} × {r['Item Name']} <span class=\"group-badge\">{r['Group']}</span>"
        })
    for _, r in sales_df.iterrows():
        activities.append({
            'Date': r['Date'],
            'Type': 'sale',
            'Text': f"💰 Sold {int(r['Sold Quantity'])} × {r['Item Name']} @ ₹{r['Sale Price']:.0f}"
        })

    if activities:
        activity_df = pd.DataFrame(activities)
        activity_df = activity_df.sort_values('Date', ascending=False).head(12)
        for _, a in activity_df.iterrows():
            badge_class = "sale" if a['Type'] == 'sale' else "stock"
            badge_text = "SALE" if a['Type'] == 'sale' else "STOCK"
            date_str = a['Date'].strftime('%d-%m-%Y') if pd.notna(a['Date']) else 'N/A'
            st.markdown(f"""
                <div class="activity-row">
                    <span>{a['Text']}</span>
                    <span><span class="activity-badge {badge_class}">{badge_text}</span> &nbsp; {date_str}</span>
                </div>
            """, unsafe_allow_html=True)
    else:
        st.info("No activity recorded yet.")
    st.markdown('</div>', unsafe_allow_html=True)

# === PROFESSIONAL MAIN APPLICATION ===
def show_main_application(inventory_manager, user_manager):

    st.markdown(PROFESSIONAL_CSS, unsafe_allow_html=True)
    
    # Handle editing/deleting first
    if st.session_state.get('editing_item'):
        group, item_name = st.session_state.editing_item
        show_edit_form(inventory_manager, group, item_name)
        return
    
    if st.session_state.get('deleting_item'):
        group, item_name = st.session_state.deleting_item
        show_delete_confirmation(inventory_manager, group, item_name)
        return
    
    # Load data
    stock_df, sales_df, metadata = inventory_manager.load_inventory_data()
    
    def calculate_remaining(stock_df, sales_df):
        if stock_df.empty:
            return pd.DataFrame(columns=['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 'Sold Quantity', 'Remaining', 'Date'])
        
        stock_total = stock_df.groupby(['Group', 'Item Name']).agg({
            'Quantity': 'sum',
            'MRP': 'first',
            'Purchase Price': 'first',
            'Date': 'max'
        }).reset_index()
        
        if sales_df.empty:
            sales_total = pd.DataFrame(columns=['Group', 'Item Name', 'Sold Quantity'])
            sales_total['Sold Quantity'] = sales_total['Sold Quantity'].astype(int)
        else:
            sales_total = sales_df.groupby(['Group', 'Item Name'])['Sold Quantity'].sum().reset_index()
        
        merged = pd.merge(stock_total, sales_total, on=['Group', 'Item Name'], how='left')
        merged['Sold Quantity'] = merged['Sold Quantity'].fillna(0)
        merged['Remaining'] = merged['Quantity'] - merged['Sold Quantity']
        
        # Get latest date from sales for each item
        if not sales_df.empty:
            latest_sale_dates = sales_df.groupby(['Group', 'Item Name'])['Date'].max().reset_index()
            latest_sale_dates = latest_sale_dates.rename(columns={'Date': 'Latest_Sale_Date'})
            merged = pd.merge(merged, latest_sale_dates, on=['Group', 'Item Name'], how='left')
            
            # Use the latest date between stock date and sale date
            merged['Date'] = merged.apply(
                lambda row: max(row['Date'], row['Latest_Sale_Date']) if pd.notna(row['Latest_Sale_Date']) else row['Date'], 
                axis=1
            )
            merged = merged.drop('Latest_Sale_Date', axis=1)
        
        result = merged[['Group', 'Item Name', 'Quantity', 'MRP', 'Purchase Price', 'Sold Quantity', 'Remaining', 'Date']]
        
        for col in ['Quantity', 'Sold Quantity', 'Remaining']:
            if col in result.columns:
                result[col] = result[col].astype('int32')
        
        return result
    
    remaining_df = calculate_remaining(stock_df, sales_df)

    
    # === SIDEBAR ===
    with st.sidebar:
        st.markdown('<div class="sidebar-header">🔧 CONTROL PANEL</div>', unsafe_allow_html=True)
        
        # User Info
        st.markdown(f"**👤 {st.session_state.username}**")
        st.markdown(f"**🎯 {st.session_state.user_role}**")
        
        st.markdown("---")
        
        # System Info
        st.markdown("**📊 SYSTEM INFO**")
        st.markdown(f'<div class="status-indicator"><span class="status-dot status-green"></span> Stock Records: {len(stock_df):,}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="status-indicator"><span class="status-dot status-yellow"></span> Sales Records: {len(sales_df):,}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="status-indicator"><span class="status-dot status-blue"></span> Unique Items: {stock_df["Item Name"].nunique() if not stock_df.empty else 0:,}</div>', unsafe_allow_html=True)
        
        # Financial Summary Toggle
        st.markdown("---")
        st.markdown("**💰 FINANCIAL SETTINGS**")
        
        # Initialize financial toggle in session state
        if 'show_financial_summary' not in st.session_state:
            st.session_state.show_financial_summary = True
        
        show_financial = st.checkbox(
            "Show Financial Metrics",
            value=st.session_state.show_financial_summary,
            key="sidebar_financial_toggle"
        )
        
        if show_financial != st.session_state.show_financial_summary:
            st.session_state.show_financial_summary = show_financial
            st.rerun()
        
        # Quick Actions
        st.markdown("---")
        st.markdown("**⚡ QUICK ACTIONS**")
        if st.button("🔄 Refresh Dashboard", use_container_width=True, type="primary"):
            st.rerun()
        
        # Logout
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True, type="primary"):
            for key in list(st.session_state.keys()):
                if key not in ['_']:
                    del st.session_state[key]
            st.rerun()
    
    # === MAIN INTERFACE ===
    st.markdown('<h1 class="main-header">🔧 N.P. SPARES</h1>', unsafe_allow_html=True)
    
    # Dashboard Metrics
    show_dashboard_metrics(stock_df, sales_df, remaining_df)

    # === NEW: LOW STOCK ALERT BANNER ===
    if not remaining_df.empty:
        low_stock_items = remaining_df[remaining_df['Remaining'] <= 1].sort_values('Remaining')
        if not low_stock_items.empty:
            chips = "".join(
                f'<span class="alert-chip">{r["Item Name"]} ({int(r["Remaining"])} left)</span>'
                for _, r in low_stock_items.head(15).iterrows()
            )
            more_count = len(low_stock_items) - 15
            more_text = f' <span class="alert-chip">+{more_count} more</span>' if more_count > 0 else ""
            st.markdown(f"""
                <div class="alert-banner">
                    <div class="alert-banner-title">⚠️ Low Stock Alert — {len(low_stock_items)} item(s) need restocking</div>
                    <div>{chips}{more_text}</div>
                </div>
            """, unsafe_allow_html=True)
    
    # Show Financial Summary if enabled
    if st.session_state.get('show_financial_summary', True):

        st.markdown("---")
        st.markdown('<div class="dashboard-header">💰 FINANCIAL SUMMARY</div>', unsafe_allow_html=True)
        
        # Calculate financial metrics
        total_investment = 0
        if not stock_df.empty:
            stock_df['Total Cost'] = stock_df['Quantity'] * stock_df['Purchase Price']
            total_investment = stock_df['Total Cost'].sum()
        
        total_sales_amount = 0
        total_profit = 0
        if not sales_df.empty:
            sales_df['Total Sale Value'] = sales_df['Sold Quantity'] * sales_df['Sale Price']
            total_sales_amount = sales_df['Total Sale Value'].sum()
            
            sales_df['Total Cost of Sold'] = sales_df['Sold Quantity'] * sales_df['Purchase Price']
            total_cost_of_sold = sales_df['Total Cost of Sold'].sum()
            
            total_profit = total_sales_amount - total_cost_of_sold
        
        # Financial Metrics in a row
        col6, col7, col8 = st.columns(3)
        
        with col6:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
            st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: #8b5cf6;">💼</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">₹{total_investment:,.0f}</div>', unsafe_allow_html=True)
            st.markdown('<div class="metric-label">TOTAL INVESTMENT</div>', unsafe_allow_html=True)
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with col7:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
            st.markdown('<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: var(--success-color);">📈</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">₹{total_sales_amount:,.0f}</div>', unsafe_allow_html=True)
            st.markdown('<div class="metric-label">TOTAL SALES</div>', unsafe_allow_html=True)
            st.markdown('</div></div>', unsafe_allow_html=True)
        
        with col8:
            profit_color = "var(--success-color)" if total_profit >= 0 else "var(--danger-color)"
            profit_icon = "📈" if total_profit >= 0 else "📉"
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown('<div class="metric-compact">', unsafe_allow_html=True)
            st.markdown(f'<div style="font-size: 1.8rem; margin-bottom: 0.3rem; color: {profit_color};">{profit_icon}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value" style="color: {profit_color}">₹{total_profit:,.0f}</div>', unsafe_allow_html=True)
            st.markdown('<div class="metric-label">TOTAL PROFIT</div>', unsafe_allow_html=True)
            st.markdown('</div></div>', unsafe_allow_html=True)
    
    # Action Buttons - LARGER and more prominent
    st.markdown("---")
    st.markdown('<div style="font-size: 1.1rem; font-weight: 700; color: var(--text-primary); margin-bottom: 0.8rem;">📋 QUICK ACTIONS</div>', unsafe_allow_html=True)
    
    # Create a container for action buttons
    with st.container():

        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            button_type = "primary" if st.session_state.get('show_add_stock') else "secondary"
            button_label = "➕ ADD STOCK"
            if st.button(button_label, 
                        use_container_width=True, 
                        type=button_type,
                        key="add_stock_main"):
                st.session_state.show_add_stock = not st.session_state.get('show_add_stock', False)
                if st.session_state.show_add_stock:
                    st.session_state.show_record_sale = False
                    st.session_state.show_inventory_overview = False
                    st.session_state.show_total_inventory = False
                    st.session_state.show_analytics = False
                st.rerun()
        
        with col2:
            button_type = "primary" if st.session_state.get('show_record_sale') else "secondary"
            button_label = "💰 RECORD SALE"
            if st.button(button_label, 
                        use_container_width=True,
                        type=button_type,
                        key="record_sale_main"):
                st.session_state.show_record_sale = not st.session_state.get('show_record_sale', False)
                if st.session_state.show_record_sale:
                    st.session_state.show_add_stock = False
                    st.session_state.show_inventory_overview = False
                    st.session_state.show_total_inventory = False
                    st.session_state.show_analytics = False
                st.rerun()
        
        with col3:
            button_type = "primary" if st.session_state.get('show_inventory_overview') else "secondary"
            button_label = "📦 VIEW INVENTORY"
            if st.button(button_label, 
                        use_container_width=True,
                        type=button_type,
                        key="view_inventory_main"):
                st.session_state.show_inventory_overview = not st.session_state.get('show_inventory_overview', False)
                if st.session_state.show_inventory_overview:
                    st.session_state.show_add_stock = False
                    st.session_state.show_record_sale = False
                    st.session_state.show_total_inventory = False
                    st.session_state.show_analytics = False
                st.rerun()
        
        with col4:
            button_type = "primary" if st.session_state.get('show_total_inventory') else "secondary"
            button_label = "📊 TOTAL INVENTORY"
            if st.button(button_label, 
                        use_container_width=True,
                        type=button_type,
                        key="total_inventory_main"):
                st.session_state.show_total_inventory = not st.session_state.get('show_total_inventory', False)
                if st.session_state.show_total_inventory:
                    st.session_state.show_add_stock = False
                    st.session_state.show_record_sale = False
                    st.session_state.show_inventory_overview = False
                    st.session_state.show_analytics = False
                st.rerun()

        with col5:
            button_type = "primary" if st.session_state.get('show_analytics') else "secondary"
            button_label = "📈 ANALYTICS"
            if st.button(button_label,
                        use_container_width=True,
                        type=button_type,
                        key="analytics_main"):
                st.session_state.show_analytics = not st.session_state.get('show_analytics', False)
                if st.session_state.show_analytics:
                    st.session_state.show_add_stock = False
                    st.session_state.show_record_sale = False
                    st.session_state.show_inventory_overview = False
                    st.session_state.show_total_inventory = False
                st.rerun()
    
    # Add Stock Form
    if st.session_state.get('show_add_stock'):

        with st.expander("➕ ADD NEW STOCK", expanded=True):
            with st.form("add_stock_form", clear_on_submit=True):
                col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
                
                with col1:
                    existing_groups = sorted(stock_df['Group'].unique().tolist()) if not stock_df.empty else []
                    group_options = ["Create New Group"] + existing_groups
                    group_selection = st.selectbox("Group", group_options, key="group_select")
                    
                    if group_selection == "Create New Group":
                        new_group = st.text_input("New Group Name", placeholder="Enter new group name", key="new_group")
                        group = new_group
                    else:
                        group = group_selection
                
                with col2:
                    all_existing_items = sorted(stock_df['Item Name'].unique().tolist()) if not stock_df.empty else []
                    item_options = ["Create New Item"] + all_existing_items
                    item_selection = st.selectbox("Item", item_options, key="item_select")
                    
                    if item_selection == "Create New Item":
                        item_name = st.text_input("New Item Name", placeholder="Enter new item name", key="new_item")
                    else:
                        item_name = item_selection
                
                with col3:
                    quantity = st.number_input("Quantity", min_value=1, value=1, key="quantity")
                    purchase_price = st.number_input("Cost (₹)", min_value=0.0, value=0.0, step=0.01, key="purchase_price")
                
                with col4:
                    mrp = st.number_input("MRP (₹)", min_value=0.0, value=0.0, step=0.01, key="mrp")
                    stock_date = st.date_input("Date", value=datetime.now().date(), key="stock_date")
                
                if st.form_submit_button("➕ ADD STOCK", use_container_width=True, type="primary"):
                    if group and group.strip() and item_name and item_name.strip():
                        new_row = pd.DataFrame([[group, item_name, quantity, mrp, purchase_price, stock_date]], 
                                             columns=inventory_manager.REQUIRED_COLUMNS_STOCK)
                        updated_stock_df = pd.concat([stock_df, new_row], ignore_index=True)
                        
                        if inventory_manager.save_inventory_data(updated_stock_df, sales_df):
                            st.success(f"✅ Added {quantity} '{item_name}' to '{group}'!")
                            time.sleep(0.5)
                            st.rerun()
                    else:
                        st.error("❌ Group and item name required")
    
    # Record Sale Form with DATE UPDATE
    elif st.session_state.get('show_record_sale'):
        with st.expander("💰 RECORD SALES", expanded=True):
            col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
            
            with col1:
                groups = [""] + sorted(stock_df['Group'].unique().tolist()) if not stock_df.empty else [""]
                sale_group = st.selectbox("Select Group", groups, key="sale_group")
            
            with col2:
                if sale_group:
                    items = stock_df[stock_df['Group'] == sale_group]['Item Name'].unique().tolist()
                    sale_item = st.selectbox("Select Item", [""] + sorted(items), key="sale_item")
                    

                    if sale_item:
                        item_data = stock_df[(stock_df['Group'] == sale_group) & (stock_df['Item Name'] == sale_item)]
                        if not item_data.empty:
                            purchase_price = float(item_data.iloc[0]['Purchase Price'])
                            mrp = float(item_data.iloc[0]['MRP'])
                            st.caption(f"Cost: ₹{purchase_price:.2f}")
                            st.caption(f"MRP: ₹{mrp:.2f}")
                else:
                    sale_item = st.selectbox("Select Item", [""], key="sale_item_empty")
                    st.info("Select a group first")
            
            with col3:
                sale_qty = st.number_input("Quantity", min_value=1, value=1, key="sale_qty")
                sale_date = st.date_input("Sale Date", value=datetime.now().date(), key="sale_date")
            
            with col4:
                sale_price = st.number_input("Sale Price (₹)", min_value=0.0, value=0.0, step=0.01, key="sale_price")
            
            if st.button("💰 RECORD SALE", use_container_width=True, type="primary", key="record_sale"):
                if sale_group and sale_item:
                    item_stock = remaining_df[
                        (remaining_df['Group'] == sale_group) & 
                        (remaining_df['Item Name'] == sale_item)
                    ]
                    
                    if not item_stock.empty and item_stock.iloc[0]['Remaining'] >= sale_qty:
                        # Create new sale record
                        new_sale = pd.DataFrame([[sale_group, sale_item, sale_qty, sale_price, purchase_price, sale_date]], 
                                              columns=inventory_manager.REQUIRED_COLUMNS_SALES)
                        updated_sales_df = pd.concat([sales_df, new_sale], ignore_index=True)
                        
                        # UPDATE: Update stock date to sale date
                        stock_mask = (stock_df['Group'] == sale_group) & (stock_df['Item Name'] == sale_item)
                        if not stock_df[stock_mask].empty:
                            stock_df.loc[stock_mask, 'Date'] = sale_date
                        
                        if inventory_manager.save_inventory_data(stock_df, updated_sales_df):
                            st.success(f"✅ Sale recorded: {sale_qty} {sale_item} at ₹{sale_price:.2f}!")
                            time.sleep(0.5)
                            st.rerun()
                    else:
                        available = item_stock.iloc[0]['Remaining'] if not item_stock.empty else 0
                        st.error(f"❌ Insufficient stock! Available: {available}")
                else:
                    st.error("❌ Select group and item")
    
    # Inventory Overview
    elif st.session_state.get('show_inventory_overview'):
        show_inventory_overview(inventory_manager)
    
    # Total Inventory View
    elif st.session_state.get('show_total_inventory'):
        show_total_inventory(inventory_manager)
    
    # Analytics & Insights View
    elif st.session_state.get('show_analytics'):
        show_analytics(inventory_manager)
    
    # Welcome Message
    else:
        show_welcome_message()
    
    # Copyright Footer
    current_year = datetime.now().year

    st.markdown(f"""
        <div class="copyright-footer">
            <p>Prepared with ❤️ for <strong style="color: var(--primary-color);">N.P. SPARES</strong> by Rohit</p>
            <p style="font-size: 0.8rem; opacity: 0.8; margin-top: 0.5rem;">
                © {current_year} N.P. SPARES | Professional Inventory Management System v3.0
            </p>
        </div>
    """, unsafe_allow_html=True)

# === MAIN APP ===
def main():
    # Initialize managers
    secure_manager = SecureDataManager()
    user_manager = UserManager(secure_manager)
    inventory_manager = InventoryDataManager(secure_manager)
    
    # App configuration
    st.set_page_config(
        page_title="N.P. SPARES - Inventory Management", 
        layout="wide",
        page_icon="🔧",
        initial_sidebar_state="expanded"
    )
    
    # Session state initialization with filter states
    defaults = {
        'authenticated': False,
        'username': None,
        'user_role': None,
        'editing_item': None,
        'deleting_item': None,
        'show_add_stock': False,
        'show_record_sale': False,
        'show_inventory_overview': False,
        'show_total_inventory': False,
        'show_analytics': False,
        'show_financial_summary': True,
        # Filter states for retention
        'selected_group_filter': "All Groups",
        'search_term_filter': "",
        'low_stock_filter': False,
        'sort_by_filter': "Item Name"
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value
    
    # Login or Main App
    if not st.session_state.authenticated:
        show_login_page(user_manager)
    else:
        show_main_application(inventory_manager, user_manager)

if __name__ == "__main__":
    main()
