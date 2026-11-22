"""
pySTREAM Web Interface using Dash
Modern DTU Wind & Energy Systems Theme
Runs the pySTREAM energy scenario modelling tool in the browser.
Created based on pySTREAM.py.
"""

import os
import sys
import copy
import shutil
import base64
import io
import time
import tempfile
import uuid
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objects as go
from dash import Dash, html, dcc, dash_table, Input, Output, State, callback_context, no_update
from flask import session, has_request_context

# Import existing Backend and Time_Series logic directly
import Backend
from Backend import (
    PROGRAM_DIR, DATA_DIR,
    Base_Year, SCData, DEFAULT_SCData,
    df_FEC_BY, df_Power_BY, df_Heat_BY, df_Transport_BY,
    df_TransportFuel_BY, df_transport_consumption_BY,
    df_transport_consumption_fuels_BY, df_fuel_CO2_BY, df_fuelType_CO2_BY,
    df_FuelPrice_BY, df_capital_costs_BY, df_om_costs_BY, df_fuel_costs_BY, df_co2_costs_BY,
    co2_price_BY, df_fuelType_consumption_BY, df_Fc_BY, df_beta_BY,
    # SC initial data and functions
    load_SCdata, BY_calculations, set_SC_transport_demand,
    SC_calculations, FEC_dataframe, calculate_transport, calculate_FR_fuel,
    calculate_heat_demand, final_heatPower_calculate, fuel_composition,
    resources, calculate_CO2, capital_costs, om_costs, co2_costs,
    fuel_costs, transport_costs,
    df_Transport_SC, df_TransportImprovement_SC, df_EnergySavings_SC,
    df_TechEff_SC, df_FuelRefinery_SC, df_FuelRefineryCost_SC,
    df_SectorsDetails_SC, df_TransportCosts_SC
)
from Time_Series import run_timeSeries

# ==============================================================================
# DTU & DTU Wind Color Palette
# ==============================================================================
dtu_red = '#990000'          # DTU Corporate Red
dtu_navy = '#2F3EEA'         # Primary Electric Cobalt
dtu_blue = '#1E24B5'         # Deep Cobalt (Darker contrast for buttons)
dtu_energy_blue = '#0284C7'  # Vibrant Cyan/Sky (Clean contrast for subheaders & secondary accents)
dtu_light_blue = '#EEF2FF'   # Soft Periwinkle Tint (High contrast for editable cells & banners)
dtu_sky = '#C7D2FE'          # Crisp Light Lavender (High visibility on the #2F3EEA sidebar)

bg_page = '#F8FAFC'          # Modern clean slate background
bg_card = '#FFFFFF'          # Pure white card surface
border_subtle = '#E2E8F0'    # Neutral border color
text_main = '#0F172A'        # Dark Slate Charcoal
text_muted = '#64748B'       # Slate gray for secondary text

# Modern chart colors aligned with renewable & energy topics
PLOT_COLORS = [
    '#006699', '#299ED6', '#00A389', '#10B981', '#F59E0B',
    '#D97706', '#990000', '#6366F1', '#8B5CF6', '#EC4899', '#64748B'
]
TIMESERIES_COLORS = [
    '#006699', '#00A389', '#299ED6', '#10B981', '#38BDF8',
    '#F59E0B', '#F97316', '#990000', '#6366F1', '#A855F7',
    '#EC4899', '#14B8A6', '#0284C7', '#84CC16', '#EAB308', '#64748B'
]

SECTORS_LIST = ['Residential', 'Tertiary', 'Industry', 'Agriculture']

SCENARIO_SHEETS = [
    'Scenario Year', 'Final Energy Consumption', 'Electricity Production', 'District Heat Boiler Production', 'Transport',
    'Fuels Transportation', 'Transport Work', 'Technology Improvements', 'Energy Savings', 'Technology Efficiencies',
    'Fuel Refinery Fuel Input Factor', 'Fuel Price', 'Fuel Refinery Cost', 'Sectors Details', 'CO2 Price'
]
SCENARIO_DATA_NAMES = [
    'Scenario_Year', 'df_beta_SC', 'df_Power_SC', 'df_Heat_SC', 'df_Transport_SC', 'df_TransportFuel_SC', 'df_TransportWork_SC',
    'df_TransportImprovement_SC', 'df_EnergySavings_SC', 'df_TechEff_SC', 'df_FuelRefinery_SC', 'df_FuelPrice_SC',
    'df_FuelRefineryCost_SC', 'df_SectorsDetails_SC', 'co2_price_SC', 'df_TransportCosts_SC'
]

# --- Session-Aware Scenario State ---
class ScenarioState:
    def __init__(self, session_id="default"):
        self.session_id = session_id
        self.last_accessed = time.time()
        # Per-session working copy of Scenario Data
        self.session_dir = os.path.join(tempfile.gettempdir(), f"pystream_{session_id}")
        os.makedirs(self.session_dir, exist_ok=True)
        self.sc_file = os.path.join(self.session_dir, "Scenario Data.xlsx")
        if not os.path.exists(self.sc_file):
            shutil.copyfile(DEFAULT_SCData, self.sc_file)
        self.reset_to_current()

    def reset_to_current(self):
        self.last_accessed = time.time()
        Backend.SCData = self.sc_file
        data = load_SCdata()
        for name, value in zip(SCENARIO_DATA_NAMES, data):
            setattr(self, name, copy.deepcopy(value))
        
        self.SC_NEC_tFE_el, self.SC_NEC_tFE_heat, self.df_beta_BY, self.engr = BY_calculations()
        Backend.engr = self.engr
        set_SC_transport_demand(self.df_Transport_SC, self.engr)
        
        self.df_hour = None
        self.technologies_used = None
        self.recalculate()

    def recalculate(self):
        self.last_accessed = time.time()
        Backend.SCData = self.sc_file
        Backend.engr = self.engr
        Backend.df_beta_SC = self.df_beta_SC
        Backend.df_EnergySavings_SC = self.df_EnergySavings_SC
        Backend.df_TechEff_SC = self.df_TechEff_SC
        SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES = SC_calculations(self.SC_NEC_tFE_el, self.SC_NEC_tFE_heat)
        self.df_FEC_SC, self.df_FEC_SC_total = FEC_dataframe(SC_FEC_ES, SC_FEC_tFE, SC_FEC_tES)
        
        self.df_transport_consumption_SC, self.df_transport_consumption_fuels_SC = calculate_transport(
            'Scenario Year', self.df_FEC_SC, self.df_Transport_SC, self.df_TransportWork_SC, self.df_TransportFuel_SC
        )
        self.df_fuel_secondary_consumption_SC = calculate_FR_fuel(
            self.df_FEC_SC, self.df_transport_consumption_fuels_SC, self.df_FuelRefinery_SC
        )
        self.total_DH_demand_SC = calculate_heat_demand(
            self.df_FEC_SC, self.df_fuel_secondary_consumption_SC, self.df_FuelRefinery_SC
        )
        (self.total_electricity_consumption_SC, self.gridloss_SC, self.df_Heat_SC['Amount'],
         self.total_electricity_demand_SC, self.el_transport_SC, self.el_DH_SC, self.el_FR_SC) = final_heatPower_calculate(
            self.df_FEC_SC, self.df_transport_consumption_fuels_SC, self.df_fuel_secondary_consumption_SC,
            self.total_DH_demand_SC, self.df_Power_SC
        )
        (self.Fc_DH_SC, self.Fc_El_SC, self.total_fuel_consumption_SC,
         self.DH_boiler_fuel_consumption_SC, self.df_Power_SC['Total fuel consumption']) = fuel_composition(
            self.df_Power_SC['Fuel'], self.df_Power_SC['Electricity production (condensing or nonCHP)'],
            self.df_Power_SC['Electric efficiency (extract unit: condensing mode)'],
            self.df_Power_SC['Fuel consumption in CHP-mode'], self.df_Power_SC['CHP heat production'],
            self.df_Power_SC['Heat efficiency in CHP mode (used for allocation)'],
            self.df_Power_SC['Mode'], self.df_Heat_SC['Fuel'], self.df_Heat_SC['Amount'],
            self.df_Heat_SC['Fuel efficiency']
        )
        (self.df_Fc_p_SC, self.df_Fc_SC, self.df_fuelType_consumption_SC,
         self.df_fuelCategory_consumption_SC) = resources(
            self.df_FEC_SC, self.df_transport_consumption_fuels_SC, self.Fc_El_SC, self.Fc_DH_SC,
            self.total_fuel_consumption_SC, self.df_FuelRefinery_SC, self.DH_boiler_fuel_consumption_SC,
            self.df_fuel_secondary_consumption_SC, self.gridloss_SC, self.total_electricity_consumption_SC,
            self.df_Power_SC['CHP heat production'].sum(), self.df_Heat_SC['Amount'].sum()
        )
        self.df_fuel_CO2_SC, self.df_fuelType_CO2_SC = calculate_CO2(
            self.df_Fc_SC, self.df_Power_SC, self.df_TechEff_SC
        )
        self.df_capital_costs_SC = capital_costs(
            self.df_Power_SC, self.df_Heat_SC, self.df_Fc_SC, self.df_FEC_SC,
            self.df_SectorsDetails_SC, self.df_FuelRefineryCost_SC, self.df_fuel_secondary_consumption_SC,
            transport_costs(self.df_Transport_SC, self.df_TransportWork_SC, self.df_TransportFuel_SC, self.df_TransportCosts_SC)
        )
        self.df_om_costs_SC = om_costs(
            self.df_Power_SC, self.df_Heat_SC, self.df_Fc_SC, self.df_FEC_SC,
            self.df_SectorsDetails_SC, self.df_FuelRefineryCost_SC, self.df_fuel_secondary_consumption_SC
        )
        self.df_co2_costs_SC = co2_costs(self.co2_price_SC, self.df_fuel_CO2_SC)
        self.df_FuelPrice_SC, self.df_fuel_costs_SC = fuel_costs(
            self.df_FuelPrice_SC, self.df_FuelRefinery_SC, self.df_FuelRefineryCost_SC, self.df_Fc_SC
        )
        self.df_hour = None
        self.technologies_used = None

    def save_scenario_file(self, file_path=None):
        self.last_accessed = time.time()
        target = file_path if (file_path and file_path != SCData) else self.sc_file
        workbook = openpyxl.load_workbook(target)
        # Year
        sheet_year = workbook['Scenario Year']
        sheet_year.cell(row=2, column=1, value=self.Scenario_Year)

        # Power
        power_shares = self.df_Power_SC['Shares'].tolist()
        sheet_power = workbook['Electricity Production']
        r = 2
        for val in power_shares:
            sheet_power.cell(row=r, column=2, value=val)
            r += 1

        # DH
        DH_shares = self.df_Heat_SC['Shares'].tolist()
        sheet_DH = workbook['District Heat Boiler Production']
        r = 2
        for val in DH_shares:
            sheet_DH.cell(row=r, column=2, value=val)
            r += 1

        # Heat Sectors
        heatSectors_shares = self.df_beta_SC.to_numpy().tolist()
        sheet_heatSectors = workbook['Final Energy Consumption']
        c = 2
        for i in range(self.df_beta_SC.shape[1]):
            r = 2
            for j in range(self.df_beta_SC.shape[0]):
                sheet_heatSectors.cell(row=r, column=c, value=heatSectors_shares[j][i])
                r += 1
            c += 1

        # Transport Fuels
        transportFuels_shares = self.df_TransportFuel_SC.to_numpy().tolist()
        sheet_transportFuels = workbook['Fuels Transportation']
        c = 2
        for i in range(self.df_TransportFuel_SC.shape[1]):
            r = 2
            for j in range(self.df_TransportFuel_SC.shape[0]):
                sheet_transportFuels.cell(row=r, column=c, value=transportFuels_shares[j][i])
                r += 1
            c += 1

        # Transport Work
        transportWork_shares = self.df_TransportWork_SC.to_numpy().tolist()
        sheet_transportWork = workbook['Transport Work']
        c = 2
        for i in range(self.df_TransportWork_SC.shape[1]):
            r = 2
            for j in range(self.df_TransportWork_SC.shape[0]):
                sheet_transportWork.cell(row=r, column=c, value=transportWork_shares[j][i])
                r += 1
            c += 1

        # Fuel Price
        fuel_prices = self.df_FuelPrice_SC['Market price [€/GJ]'].tolist()
        sheet_fuelPrices = workbook['Fuel Price']
        r = 2
        for val in fuel_prices:
            sheet_fuelPrices.cell(row=r, column=2, value=val)
            r += 1

        # CO2 Price
        sheet_co2Price = workbook['CO2 Price']
        sheet_co2Price.cell(row=2, column=1, value=self.co2_price_SC)

        workbook.save(target)

# --- Per-User Session Management ---
user_sessions = {}

def cleanup_old_sessions():
    """Remove user sessions inactive for more than 4 hours to keep memory clean."""
    now = time.time()
    expired = [sid for sid, s in list(user_sessions.items()) if (now - getattr(s, 'last_accessed', now)) > 14400]
    for sid in expired:
        s = user_sessions.pop(sid, None)
        if s and hasattr(s, 'session_dir') and os.path.exists(s.session_dir):
            try:
                shutil.rmtree(s.session_dir, ignore_errors=True)
            except Exception:
                pass

def get_state():
    cleanup_old_sessions()
    sid = "default_session"
    if has_request_context():
        if 'session_id' not in session:
            session['session_id'] = str(uuid.uuid4())
            session.permanent = True
        sid = session['session_id']
    
    if sid not in user_sessions:
        user_sessions[sid] = ScenarioState(sid)
    
    curr = user_sessions[sid]
    curr.last_accessed = time.time()
    Backend.SCData = curr.sc_file
    if hasattr(curr, 'engr'):
        Backend.engr = curr.engr
    if hasattr(curr, 'df_beta_SC'):
        Backend.df_beta_SC = curr.df_beta_SC
    if hasattr(curr, 'df_EnergySavings_SC'):
        Backend.df_EnergySavings_SC = curr.df_EnergySavings_SC
    if hasattr(curr, 'df_TechEff_SC'):
        Backend.df_TechEff_SC = curr.df_TechEff_SC
    return curr

class StateProxy:
    def __getattr__(self, name):
        return getattr(get_state(), name)

    def __setattr__(self, name, value):
        setattr(get_state(), name, value)

state = StateProxy()

# --- Plot Helpers (Modern DTU Wind Aesthetic) ---
def create_sankey_figure(df_fuelType_consumption, df_Fc, title):
    label = [
        'Bioenergy', 'Coal', 'Oil', 'Geothermal', 'Hydro', 'Waste', 'Natural Gas', 'Nuclear', 'Solar', 'Wave', 'Wind', # 0-10
        'Fuel Refinery', 'Electricity', 'District heating', # 11, 12, 13
        'Transport', 'Residential', 'Tertiary', 'Industry', 'Agriculture' # 14, 15, 16, 17, 18
    ]
    source_targets = {
        'Fuel Refinery': 11, 'Electricity production': 12, 'Heat production': 12, 'DH boilers': 13,
        'Transport': 14, 'Residential': 15, 'Tertiary': 16, 'Industry': 17, 'Agriculture': 18
    }
    source_colors = [
        'rgba(16, 185, 129, 0.45)', 'rgba(100, 116, 139, 0.45)', 'rgba(71, 85, 105, 0.45)', 'rgba(249, 115, 22, 0.45)',
        'rgba(14, 165, 233, 0.45)', 'rgba(245, 158, 11, 0.45)', 'rgba(34, 197, 94, 0.45)', 'rgba(99, 102, 241, 0.45)',
        'rgba(234, 179, 8, 0.45)',  'rgba(6, 182, 212, 0.45)',  'rgba(0, 102, 153, 0.55)'
    ]
    source, target, values, color_link = [], [], [], []
    for i, fuel_type in enumerate(label[:11]):
        for column, node in source_targets.items():
            value = df_fuelType_consumption.loc[fuel_type, column]
            if value > 0:
                source.append(i)
                target.append(node)
                values.append(value)
                color_link.append(source_colors[i])

    FR = df_Fc.loc[['Biodiesel', 'Biogas', 'Diesel', 'Ethanol', 'Gasoline', 'Hydrogen', 'Methanol', 'UPGR Biogas'],
                   ['Electricity production', 'Heat production', 'DH boilers', 'Transport', 'Residential', 'Tertiary', 'Industry', 'Agriculture']].sum().tolist()
    FR_EL = df_Fc.loc['Electricity', 'Fuel Refinery']
    FR_DH = df_Fc.loc['District Heating', 'Fuel Refinery']

    El = df_fuelType_consumption.loc['Electricity', ['DH boilers', 'Transport', 'Residential', 'Tertiary', 'Industry', 'Agriculture']].tolist()
    if FR_EL >= 0:
        El.append(FR_EL)
        FR.append(0)
    else:
        FR.append(abs(FR_EL))
        El.append(0)

    heat_CHP = df_fuelType_consumption.loc['District Heating', 'Heat production']
    El[0] += abs(heat_CHP)

    DH = df_fuelType_consumption.loc['District Heating', ['Residential', 'Tertiary', 'Industry', 'Agriculture']].tolist()
    if FR_DH >= 0:
        DH.append(FR_DH)
        FR.append(0)
    else:
        FR.append(abs(FR_DH))
        DH.append(0)

    source += [11]*10 + [12]*7 + [13]*5
    target += [12, 12, 13, 14, 15, 16, 17, 18, 12, 13,
               13, 14, 15, 16, 17, 18, 11,
               15, 16, 17, 18, 11]
    values += FR + El + DH
    color_link += ['rgba(13, 148, 136, 0.45)']*10 + ['rgba(0, 102, 153, 0.45)']*7 + ['rgba(153, 0, 0, 0.45)']*5
    color_node = [
        '#10B981', '#64748B', '#475569', '#F97316', '#0EA5E9', '#F59E0B', '#22C55E', '#6366F1', '#EAB308', '#06B6D4', '#006699',
        '#0D9488', '#003366', '#990000'
    ] + [dtu_navy]*5

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(label=label, pad=18, thickness=16, color=color_node),
        link=dict(source=source, target=target, value=values, color=color_link)
    )])
    fig.update_layout(
        title_text=title,
        title_font=dict(size=16, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=50, b=20),
        autosize=True
    )
    return fig

def plotly_stacked_bar(df_data, title, yaxis_title="[TJ]"):
    fig = go.Figure()
    for i, idx in enumerate(df_data.index):
        color = PLOT_COLORS[i % len(PLOT_COLORS)]
        fig.add_trace(go.Bar(
            name=str(idx), x=list(df_data.columns), y=df_data.loc[idx],
            marker_color=color
        ))
    fig.update_layout(
        barmode='stack', title=title,
        title_font=dict(size=15, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        yaxis_title=yaxis_title,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=40, r=20, t=40, b=40),
        autosize=True,
        legend=dict(orientation="v", yanchor="middle", y=0.5, xanchor="left", x=1.02)
    )
    fig.update_yaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')
    fig.update_xaxes(zerolinecolor='#E2E8F0')
    return fig

def plotly_bar(series_data, title, yaxis_title="[PJ]", horizontal=False):
    fig = go.Figure()
    colors = [PLOT_COLORS[i % len(PLOT_COLORS)] for i in range(len(series_data))]
    if horizontal:
        fig.add_trace(go.Bar(
            y=list(series_data.index), x=list(series_data.values),
            orientation='h', marker_color=colors
        ))
        fig.update_layout(xaxis_title=yaxis_title)
        fig.update_xaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')
        fig.update_yaxes(zerolinecolor='#E2E8F0')
    else:
        fig.add_trace(go.Bar(
            x=list(series_data.index), y=list(series_data.values),
            marker_color=colors
        ))
        fig.update_layout(yaxis_title=yaxis_title)
        fig.update_yaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')
        fig.update_xaxes(zerolinecolor='#E2E8F0')
        
    fig.update_layout(
        title=title,
        title_font=dict(size=15, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=40, r=20, t=40, b=60),
        autosize=True
    )
    return fig

def plotly_pie(series_data, title, colors=None):
    if colors is None:
        colors = PLOT_COLORS[:len(series_data)]
    fig = go.Figure(data=[go.Pie(
        labels=list(series_data.index), values=list(series_data.values),
        marker=dict(colors=colors, line=dict(color='#FFFFFF', width=1.5)),
        textinfo='label+percent', insidetextorientation='radial',
        hole=0.35 # Modern donut chart style
    )])
    fig.update_layout(
        title=title,
        title_font=dict(size=14, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(l=20, r=20, t=40, b=20),
        height=320,
        showlegend=False
    )
    return fig

# --- Dash App Setup ---
app = Dash(__name__, suppress_callback_exceptions=True)
app.title = "pySTREAM | DTU Wind & Energy Systems"
server = app.server
server.secret_key = os.environ.get('SECRET_KEY', 'pystream-dtu-wind-secret-key-2026')

@server.before_request
def ensure_session_cookie():
    if 'session_id' not in session:
        session['session_id'] = str(uuid.uuid4())
        session.permanent = True

# ==============================================================================
# Modern CSS Styles & Design System
# ==============================================================================
HEADER_STYLE = {
    'fontFamily': 'Inter, system-ui, -apple-system, sans-serif',
    'color': dtu_navy,
    'fontWeight': '700',
    'fontSize': '24px',
    'letterSpacing': '-0.02em',
    'marginBottom': '6px'
}

SUBHEADER_STYLE = {
    'fontFamily': 'Inter, system-ui, -apple-system, sans-serif',
    'color': dtu_energy_blue,
    'fontWeight': '600',
    'fontSize': '15px',
    'marginBottom': '12px'
}

SIDEBAR_STYLE = {
    "position": "fixed",
    "top": 0,
    "left": 0,
    "bottom": 0,
    "width": "250px",
    "padding": "24px 14px",
    "backgroundColor": dtu_navy,
    "overflowY": "auto",
    "zIndex": 100,
    "boxShadow": "2px 0 10px rgba(0, 40, 70, 0.15)",
    "fontFamily": "Inter, system-ui, -apple-system, sans-serif"
}

CONTENT_STYLE = {
    "marginLeft": "265px",
    "padding": "24px 32px",
    "backgroundColor": bg_page,
    "minHeight": "100vh",
    "fontFamily": "Inter, system-ui, -apple-system, sans-serif",
    "boxSizing": "border-box"
}

NAV_BTN_STYLE = {
    'width': '100%',
    'backgroundColor': 'transparent',
    'color': '#E2E8F0',
    'border': 'none',
    'borderRadius': '6px',
    'padding': '10px 14px',
    'marginBottom': '4px',
    'fontWeight': '500',
    'textAlign': 'left',
    'cursor': 'pointer',
    'fontSize': '13px',
    'transition': 'all 0.15s ease',
    'display': 'flex',
    'alignItems': 'center'
}

ACTION_BTN_STYLE = {
    'backgroundColor': dtu_blue,
    'color': '#FFFFFF',
    'border': 'none',
    'borderRadius': '6px',
    'padding': '8px 16px',
    'fontWeight': '600',
    'cursor': 'pointer',
    'fontSize': '13px',
    'boxShadow': '0 1px 2px rgba(0, 0, 0, 0.05)',
    'transition': 'all 0.15s ease'
}

CARD_STYLE = {
    'backgroundColor': bg_card,
    'borderRadius': '8px',
    'border': f'1px solid {border_subtle}',
    'padding': '18px 20px',
    'marginBottom': '20px',
    'boxShadow': '0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.02)'
}

TABLE_STYLE_HEADER = {
    'backgroundColor': dtu_navy,
    'color': '#FFFFFF',
    'fontWeight': '600',
    'fontSize': '12px',
    'textAlign': 'left',
    'padding': '8px 10px',
    'borderBottom': f'2px solid {dtu_energy_blue}'
}

TABLE_STYLE_CELL = {
    'backgroundColor': '#FFFFFF',
    'color': text_main,
    'fontSize': '12px',
    'padding': '7px 10px',
    'border': f'1px solid {border_subtle}',
    'textAlign': 'left'
}

# --- Dynamic Session Layout ---
def serve_layout():
    return html.Div([
        dcc.Location(id="url", refresh=False),
        
        # Left Sidebar Menu (DTU Wind Identity)
        html.Div([
            # DTU Branding Header
            html.Div([
                html.Div([
                    html.Span("DTU", style={
                        'backgroundColor': dtu_red, 'color': '#FFFFFF', 'fontWeight': '800',
                        'padding': '3px 8px', 'borderRadius': '4px', 'fontSize': '14px',
                        'letterSpacing': '0.05em', 'marginRight': '8px'
                    }),
                    html.Span("Wind", style={'color': '#FFFFFF', 'fontSize': '18px', 'fontWeight': '700'}),
                ], style={'display': 'flex', 'alignItems': 'center', 'marginBottom': '4px'}),
                html.Div("pySTREAM 2.0 Web", style={'color': dtu_sky, 'fontSize': '13px', 'fontWeight': '600', 'letterSpacing': '0.02em'}),
                html.Div("Energy System Simulation Tool", style={'color': '#94A3B8', 'fontSize': '10px', 'marginTop': '2px'})
            ], style={'marginBottom': '24px', 'padding': '4px 6px 14px 6px', 'borderBottom': '1px solid rgba(255,255,255,0.1)'}),
            
            # Navigation Links
            html.Div([
                dcc.Link(html.Button("📊 Graphs & Energy Streams", style=NAV_BTN_STYLE), href="/graphs"),
                dcc.Link(html.Button("⚡ Edit Consumption", style=NAV_BTN_STYLE), href="/consumption"),
                dcc.Link(html.Button("🏭 Edit Production", style=NAV_BTN_STYLE), href="/production"),
                dcc.Link(html.Button("📅 Base Year Data", style=NAV_BTN_STYLE), href="/by"),
                dcc.Link(html.Button("🔮 Scenario Data", style=NAV_BTN_STYLE), href="/scenario"),
                dcc.Link(html.Button("🚗 Transport Data", style=NAV_BTN_STYLE), href="/transport"),
                dcc.Link(html.Button("🌱 CO2 Emissions", style=NAV_BTN_STYLE), href="/co2"),
                dcc.Link(html.Button("💶 Economy & Costs", style=NAV_BTN_STYLE), href="/economy"),
                dcc.Link(html.Button("📈 Time Series Simulation", style=NAV_BTN_STYLE), href="/timeseries"),
            ], style={'marginBottom': '24px'}),
            
            # Scenario File Controls in Bottom Card
            html.Div([
                html.Div("Scenario Management", style={'color': '#FFFFFF', 'fontWeight': '600', 'fontSize': '11px', 'textTransform': 'uppercase', 'letterSpacing': '0.05em', 'marginBottom': '10px'}),
                dcc.Upload(
                    id='upload-scenario',
                    children=html.Button("📂 Open scenario (.xlsx)", style={
                        **NAV_BTN_STYLE, 'backgroundColor': 'rgba(255,255,255,0.08)', 'color': '#FFFFFF', 'fontSize': '11px', 'padding': '6px 10px', 'marginBottom': '6px'
                    }),
                    multiple=False
                ),
                html.Button("💾 Download scenario (.xlsx)", id="btn-download-sc", style={
                    **NAV_BTN_STYLE, 'backgroundColor': 'rgba(255,255,255,0.08)', 'color': '#FFFFFF', 'fontSize': '11px', 'padding': '6px 10px', 'marginBottom': '6px'
                }),
                dcc.Download(id="download-scenario-xlsx"),
                html.Button("🔄 Reset to default", id="btn-reset-sc", style={
                    **NAV_BTN_STYLE, 'backgroundColor': 'rgba(153, 0, 0, 0.4)', 'color': '#FFCCCC', 'fontSize': '11px', 'padding': '6px 10px', 'marginTop': '4px'
                }),
            ], style={'backgroundColor': 'rgba(0, 0, 0, 0.2)', 'borderRadius': '6px', 'padding': '12px', 'marginTop': 'auto','marginBottom': '24px'}),
            # Acknowledgment Footer
            html.Div([
                html.Div("Acknowledgment", style={'color': '#FFFFFF', 'fontWeight': '600', 'fontSize': '11px', 'textTransform': 'uppercase', 'letterSpacing': '0.05em', 'marginBottom': '10px'}),
                html.P([
                    "pySTREAM developed by Ida Cecilie Hoielt: ",
                    html.A(
                        "GitHub repository",
                        href="https://github.com/idahoielt/pySTREAM",
                        target="_blank",
                        style={'color': dtu_sky, 'textDecoration': 'underline', 'fontWeight': '600'}
                    ),
                ], style={'color': '#E2E8F0', 'fontSize': '11px', 'marginBottom': '8px'}),
                html.P([
                    "STREAM developed by DTU Management: ",
                    html.A(
                        "DTU STREAM website",
                        href="https://www.esymodels.man.dtu.dk/da/stream",
                        target="_blank",
                        style={'color': dtu_sky, 'textDecoration': 'underline', 'fontWeight': '600'}
                    )
                ], style={'color': '#E2E8F0', 'fontSize': '11px', 'marginBottom': '4px'}),
            ], style={'backgroundColor': 'rgba(0, 0, 0, 0.25)', 'borderRadius': '6px', 'padding': '12px', 'marginTop': 'auto'}),
        ], style=SIDEBAR_STYLE),

        # Main content area
        html.Div([
            # Top Header Bar with Scenario Year & Global Recalculate
            html.Div([
                html.Div([
                    html.Span("Scenario Target Year: ", style={'fontWeight': '600', 'color': dtu_navy, 'fontSize': '14px', 'marginRight': '8px'}),
                    dcc.Input(
                        id="input-scenario-year", type="number", value=state.Scenario_Year,
                        style={'width': '85px', 'padding': '6px 8px', 'borderRadius': '4px', 'border': f'1px solid {border_subtle}', 'fontSize': '14px', 'fontWeight': '600', 'marginRight': '10px'}
                    ),
                    html.Button("Set Year", id="btn-update-year", style={**ACTION_BTN_STYLE, 'backgroundColor': dtu_energy_blue}),
                ], style={'display': 'flex', 'alignItems': 'center'}),
                html.Div([
                    html.Button("⚡ Recalculate & Save Scenario", id="btn-global-update", style={
                        **ACTION_BTN_STYLE, 'backgroundColor': dtu_navy, 'fontSize': '13px', 'padding': '8px 20px', 'border': f'1px solid {dtu_energy_blue}'
                    })
                ])
            ], style={
                'display': 'flex', 'justifyContent': 'space-between', 'alignItems': 'center',
                'padding': '12px 20px', 'backgroundColor': bg_card, 'borderRadius': '8px',
                'border': f'1px solid {border_subtle}', 'marginBottom': '18px',
                'boxShadow': '0 1px 3px rgba(0,0,0,0.03)'
            }),

            # Alert banner for messages/errors
            html.Div(id="alert-banner", style={'marginBottom': '14px'}),

            # Dynamic Page Content
            html.Div(id="page-content")
        ], style=CONTENT_STYLE)
    ])

app.layout = serve_layout

# --- Page Builders ---
def build_graphs_page(active_tab="sankey"):
    return html.Div([
        html.H2("Graphs & Energy Balance", style=HEADER_STYLE),
        html.P("Compare system flows, emissions, and sector profiles between Base Year and Scenario Year.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        # Segmented Filter Bar
        html.Div([
            dcc.RadioItems(
                id="graph-sub-tabs",
                options=[
                    {'label': ' 🌊 Sankey Energy Streams', 'value': 'sankey'},
                    {'label': ' 🏢 Sector Technologies (FEC)', 'value': 'fec_stacked'},
                    {'label': ' ⚡ All Sectors Total Consumption', 'value': 'fec_all'},
                    {'label': ' 💨 CO2 Emissions', 'value': 'co2'},
                    {'label': ' 🔌 Electricity Generation', 'value': 'el'},
                    {'label': ' ♨️ District Heating Boilers', 'value': 'dh'},
                    {'label': ' 🛢️ Primary Fuel Consumption', 'value': 'fuel'},
                ],
                value=active_tab,
                inline=True,
                inputStyle={"marginRight": "5px", "marginLeft": "14px"},
                style={'fontWeight': '600', 'color': dtu_navy, 'fontSize': '13px', 'display': 'flex', 'flexWrap': 'wrap', 'alignItems': 'center'}
            )
        ], style={**CARD_STYLE, 'padding': '12px 16px', 'backgroundColor': dtu_light_blue, 'borderColor': '#C7D2FE'}),
        
        # Two-Column Chart Layout
        html.Div([
            # Left: Base Year
            html.Div([
                html.Div([
                    html.Span("Base Year", style={'fontWeight': '700', 'color': dtu_navy, 'fontSize': '14px'}),
                    html.Span(f" ({Base_Year})", style={'color': text_muted, 'fontSize': '13px'})
                ], style={'marginBottom': '10px', 'borderBottom': f'1px solid {border_subtle}', 'paddingBottom': '6px'}),
                dcc.Graph(id="graph-by-display", style={'height': '520px'}, config={'responsive': True})
            ], style={
                'flex': '1 1 500px', 'minWidth': '420px', 'backgroundColor': bg_card,
                'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'
            }),
            # Right: Scenario Year
            html.Div([
                html.Div([
                    html.Span("Scenario Target Year", style={'fontWeight': '700', 'color': dtu_energy_blue, 'fontSize': '14px'}),
                    html.Span(f" ({state.Scenario_Year})", style={'color': text_muted, 'fontSize': '13px'})
                ], style={'marginBottom': '10px', 'borderBottom': f'1px solid {border_subtle}', 'paddingBottom': '6px'}),
                dcc.Graph(id="graph-sc-display", style={'height': '520px'}, config={'responsive': True})
            ], style={
                'flex': '1 1 500px', 'minWidth': '420px', 'backgroundColor': bg_card,
                'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'
            }),
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '20px', 'alignItems': 'stretch'})
    ])

def build_consumption_page():
    # Heat table preparation
    fuels_heat = ['Electricity', 'District heating', 'Coal', 'Oil', 'Natural gas', 'Biomass', 'Solar', 'Electricity']
    tech_types_heat = list(state.df_beta_SC[1:].index.values)
    heat_rows = []
    for i in range(len(tech_types_heat)):
        row = {'Tech Type': tech_types_heat[i], 'Fuels': fuels_heat[i]}
        for s in SECTORS_LIST:
            val = state.df_beta_SC.loc[tech_types_heat[i], s]
            row[s] = f"{val * 100:.2f}"
        heat_rows.append(row)

    heat_columns = [{'name': 'Technology Type', 'id': 'Tech Type', 'editable': False},
                    {'name': 'Fuels', 'id': 'Fuels', 'editable': False}] + \
                   [{'name': f"{s} (%)", 'id': s, 'editable': True} for s in SECTORS_LIST]

    # Transport fuels table preparation
    transport_fuels_rows = []
    transport_fuels_types = [t for t in state.df_TransportFuel_SC.index if t != 'Bike']
    for i, t in enumerate(transport_fuels_types):
        group = 'Person' if i < 4 else ('Goods' if i < 8 else 'Fishery & Agriculture')
        row = {'Sector': group, 'Type': t}
        for col in state.df_TransportFuel_SC.columns:
            row[col] = f"{state.df_TransportFuel_SC.loc[t, col] * 100:.2f}"
        transport_fuels_rows.append(row)

    tf_columns = [{'name': 'Sector', 'id': 'Sector', 'editable': False},
                  {'name': 'Transport Type', 'id': 'Type', 'editable': False}] + \
                 [{'name': f"{c} (%)", 'id': c, 'editable': True} for c in state.df_TransportFuel_SC.columns]

    # Allocation & Utilisation tables
    alloc_rows = []
    for i, idx in enumerate(state.df_TransportWork_SC.index):
        group = 'Person' if i < 5 else 'Goods'
        alloc_rows.append({
            'Group': group,
            'Type': idx,
            'Allocation': f"{state.df_TransportWork_SC.loc[idx, 'Allocation of transport work'] * 100:.2f}",
            'Utilisation': f"{state.df_TransportWork_SC.loc[idx, 'Utilisation degree'] * 100:.2f}"
        })
    alloc_cols = [
        {'name': 'Group', 'id': 'Group', 'editable': False},
        {'name': 'Transport Type', 'id': 'Type', 'editable': False},
        {'name': 'Allocation (%)', 'id': 'Allocation', 'editable': True},
        {'name': 'Utilisation (%)', 'id': 'Utilisation', 'editable': True}
    ]

    return html.Div([
        html.H2(f"Edit Final Energy Consumption - {state.Scenario_Year}", style=HEADER_STYLE),
        html.P("Double-click any cell to adjust sector technology shares or transportation details. Ensure sector sums equal 100%.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        # Row 1: Heating Consumption & Transport Work
        html.Div([
            html.Div([
                html.H4("Heating Energy Consumption Shares", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-heat-consumption',
                    columns=heat_columns,
                    data=heat_rows,
                    editable=True,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    style_data_conditional=[{'if': {'column_editable': True}, 'backgroundColor': dtu_light_blue}]
                ),
                html.Div(id='heat-sums-display', style={'marginTop': '8px', 'fontSize': '12px', 'fontWeight': 'bold', 'color': dtu_navy}),
            ], style={'flex': '3 1 520px', 'minWidth': '380px', 'boxSizing': 'border-box'}),
            
            # Transport Work (Allocation & Utilisation)
            html.Div([
                html.H4("Transport Work & Utilisation Degrees", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-transport-work',
                    columns=alloc_cols,
                    data=alloc_rows,
                    editable=True,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    style_data_conditional=[{'if': {'column_editable': True}, 'backgroundColor': dtu_light_blue}]
                ),
                html.Div(id='work-sums-display', style={'marginTop': '8px', 'fontSize': '12px', 'fontWeight': 'bold', 'color': dtu_navy})
            ], style={'flex': '2 1 340px', 'minWidth': '300px', 'boxSizing': 'border-box'})
        ], style={**CARD_STYLE, 'display': 'flex', 'flexWrap': 'wrap', 'gap': '20px'}),

        # Row 2: Transport Fuels
        html.Div([
            html.H4("Transport Fuel Shares (%)", style=SUBHEADER_STYLE),
            dash_table.DataTable(
                id='table-transport-fuels',
                columns=tf_columns,
                data=transport_fuels_rows,
                editable=True,
                style_header=TABLE_STYLE_HEADER,
                style_cell=TABLE_STYLE_CELL,
                style_table={'overflowX': 'auto'},
                style_data_conditional=[{'if': {'column_editable': True}, 'backgroundColor': dtu_light_blue}]
            ),
            html.Div(id='tf-sums-display', style={'marginTop': '8px', 'fontSize': '12px', 'fontWeight': 'bold', 'color': dtu_navy}),
        ], style=CARD_STYLE),

        html.Button("💾 Save Consumption Changes", id="btn-save-consumption", style={**ACTION_BTN_STYLE, 'padding': '10px 24px', 'fontSize': '14px'})
    ])

def build_production_page():
    # Electricity rows
    el_rows = []
    for idx in state.df_Power_SC.index[:-1]:
        share = state.df_Power_SC.loc[idx, 'Shares']
        amount = state.df_Power_SC.loc[idx, 'Amount']
        el_rows.append({
            'Method': idx,
            'Shares': f"{share * 100:.2f}",
            'Amount': f"{amount:.2f}"
        })
    el_cols = [
        {'name': 'Production Method', 'id': 'Method', 'editable': False},
        {'name': 'Share (%)', 'id': 'Shares', 'editable': True},
        {'name': 'Amount [PJ]', 'id': 'Amount', 'editable': False}
    ]

    # DH rows
    dh_rows = []
    for idx in state.df_Heat_SC.index[:-1]:
        share = state.df_Heat_SC.loc[idx, 'Shares']
        amount = state.df_Heat_SC.loc[idx, 'Amount']
        dh_rows.append({
            'Method': idx,
            'Shares': f"{share * 100:.2f}",
            'Amount': f"{amount:.2f}"
        })
    dh_cols = [
        {'name': 'Production Method', 'id': 'Method', 'editable': False},
        {'name': 'Share (%)', 'id': 'Shares', 'editable': True},
        {'name': 'Amount [PJ]', 'id': 'Amount', 'editable': False}
    ]

    # Resources table
    res_2 = state.df_Fc_p_SC['Total primary fuel consumption'].drop(['Electricity', 'District Heating', 'Imported electricity', 'Imported heat'], errors='ignore')
    if 'Oil' in res_2.index and 'Crude Oil' in res_2.index:
        res_2['Oil'] = res_2['Oil'] + res_2['Crude Oil']
        res_2 = res_2.drop('Crude Oil', errors='ignore')
    
    res_rows = []
    for fuel_idx in res_2.index:
        avail = 0.0
        if hasattr(Backend, 'df_Resources_SC') and fuel_idx in Backend.df_Resources_SC['Type'].values:
            avail = Backend.df_Resources_SC.loc[Backend.df_Resources_SC['Type'] == fuel_idx, 'Amount'].values[0]
        used = res_2[fuel_idx]
        res_rows.append({'Resource': fuel_idx, 'Available [PJ]': f"{avail:.0f}", 'Used [PJ]': f"{used:.0f}"})

    res_cols = [
        {'name': 'Resource', 'id': 'Resource'},
        {'name': 'Available [PJ]', 'id': 'Available [PJ]'},
        {'name': 'Used [PJ]', 'id': 'Used [PJ]'}
    ]

    # Pie charts
    el_pie_data = state.df_fuelCategory_consumption_SC.iloc[1:]['Electricity production']
    fig_el_pie = plotly_pie(el_pie_data, "Electricity Production Fuel Categories")

    dh_pie_series = state.df_fuelCategory_consumption_SC.iloc[1:3]['Heat production'] + state.df_fuelCategory_consumption_SC.iloc[1:3]['DH boilers']
    fig_dh_pie = plotly_pie(dh_pie_series, "District Heat Production Categories (incl. CHP)")

    return html.Div([
        html.H2(f"Edit Energy Production - {state.Scenario_Year}", style=HEADER_STYLE),
        html.P("Configure capacity and generation shares for power and district heat technologies.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            # Column 1: Electricity
            html.Div([
                html.H4("Electricity Production", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-el-production',
                    columns=el_cols,
                    data=el_rows,
                    editable=True,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    style_data_conditional=[{'if': {'column_editable': True}, 'backgroundColor': dtu_light_blue}]
                ),
                html.Div([
                    html.Div(f"Demand Coverage: {state.df_Power_SC['Shares'][:-1].sum() * 100:.2f}%"),
                    html.Div(f"Imports/Exports: {state.df_Power_SC.iloc[-1]['Shares'] * 100:.2f}%"),
                    html.Div(f"Total Demand: {state.total_electricity_consumption_SC:.2f} PJ"),
                    html.Div(f"Total Production: {state.df_Power_SC['Amount'][:-1].sum():.2f} PJ"),
                ], style={'marginTop': '12px', 'padding': '10px', 'backgroundColor': dtu_light_blue, 'borderRadius': '6px', 'fontSize': '12px', 'fontWeight': '600', 'color': dtu_navy})
            ], style={'flex': '1 1 300px', 'minWidth': '260px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'}),
            
            # Column 2: District Heating
            html.Div([
                html.H4("District Heat Boiler Production", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-dh-production',
                    columns=dh_cols,
                    data=dh_rows,
                    editable=True,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    style_data_conditional=[{'if': {'column_editable': True}, 'backgroundColor': dtu_light_blue}]
                ),
                html.Div([
                    html.Div(f"Boiler Production: {state.df_Heat_SC['Shares'][:-1].sum() * 100:.2f}% ({state.df_Heat_SC['Amount'][:-1].sum():.1f} PJ)"),
                    html.Div(f"CHP Heat: {state.df_Power_SC['CHP heat production'].sum():.1f} PJ"),
                    html.Div(f"Total District Heating: {state.df_Heat_SC['Amount'][:-1].sum() + state.df_Power_SC['CHP heat production'].sum():.1f} PJ"),
                ], style={'marginTop': '12px', 'padding': '10px', 'backgroundColor': dtu_light_blue, 'borderRadius': '6px', 'fontSize': '12px', 'fontWeight': '600', 'color': dtu_navy})
            ], style={'flex': '1 1 300px', 'minWidth': '260px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'}),

            # Column 3: Resources
            html.Div([
                html.H4("Primary Resources Balance", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-resources',
                    columns=res_cols,
                    data=res_rows,
                    editable=False,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'}
                )
            ], style={'flex': '1 1 300px', 'minWidth': '260px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '18px'}),

        html.Div([
            html.Button("💾 Save Production Changes", id="btn-save-production", style={**ACTION_BTN_STYLE, 'padding': '10px 24px', 'fontSize': '14px', 'marginTop': '16px'}),
        ]),

        # Category charts below
        html.Div([
            html.Div([dcc.Graph(figure=fig_el_pie, config={'responsive': True})], style={'flex': '1 1 380px', 'minWidth': '320px', 'boxSizing': 'border-box'}),
            html.Div([dcc.Graph(figure=fig_dh_pie, config={'responsive': True})], style={'flex': '1 1 380px', 'minWidth': '320px', 'boxSizing': 'border-box'})
        ], style={**CARD_STYLE, 'marginTop': '20px', 'display': 'flex', 'flexWrap': 'wrap', 'gap': '16px'})
    ])

def build_data_view_page(is_base_year=True):
    year_label = str(Base_Year) if is_base_year else str(state.Scenario_Year)
    df_fec = df_FEC_BY if is_base_year else state.df_FEC_SC
    df_power = df_Power_BY if is_base_year else state.df_Power_SC
    df_heat = df_Heat_BY if is_base_year else state.df_Heat_SC

    # Stacked bar plot for FEC
    fig_fec = plotly_stacked_bar(df_fec.iloc[:, :4], f"Final Energy Consumption per Sector - {year_label}")

    # FEC table
    fec_data = []
    for idx in df_fec.index:
        row = {'Tech Type': idx}
        for col in df_fec.columns:
            row[col] = f"{df_fec.loc[idx, col]:.1f}"
        fec_data.append(row)
    fec_cols = [{'name': 'Technology Type', 'id': 'Tech Type'}] + [{'name': str(c), 'id': str(c)} for c in df_fec.columns]

    # Power & Heat tables
    power_data = [{'Method': idx, 'Amount [PJ]': f"{df_power.loc[idx, 'Amount']:.2f}"} for idx in df_power.index]
    power_cols = [{'name': 'Power Method', 'id': 'Method'}, {'name': 'Amount [PJ]', 'id': 'Amount [PJ]'}]

    heat_data = [{'Method': idx, 'Amount [PJ]': f"{df_heat.loc[idx, 'Amount']:.2f}"} for idx in df_heat.index]
    heat_cols = [{'name': 'Heat Method', 'id': 'Method'}, {'name': 'Amount [PJ]', 'id': 'Amount [PJ]'}]

    return html.Div([
        html.H2(f"{'Base Year' if is_base_year else 'Scenario Target'} Overview ({year_label})", style=HEADER_STYLE),
        html.P("Summary data sheets for consumption and generation.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            # Left: FEC Graph & Table
            html.Div([
                html.H4("Final Energy Consumption (FEC) [TJ]", style=SUBHEADER_STYLE),
                dcc.Graph(figure=fig_fec, style={'height': '380px'}, config={'responsive': True}),
                dash_table.DataTable(
                    columns=fec_cols,
                    data=fec_data,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    page_size=10
                )
            ], style={'flex': '3 1 450px', 'minWidth': '380px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'}),
            
            # Right: Production tables
            html.Div([
                html.H4("Electricity & Heat Production [PJ]", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    columns=power_cols,
                    data=power_data,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    page_size=8
                ),
                html.Div(style={'height': '15px'}),
                dash_table.DataTable(
                    columns=heat_cols,
                    data=heat_data,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    page_size=8
                )
            ], style={'flex': '2 1 320px', 'minWidth': '280px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '20px'})
    ])

def build_transport_page():
    fig_by = plotly_stacked_bar(df_TransportFuel_BY[df_TransportFuel_BY.index != 'Bike'].T, f"Fuel Shares - Base Year ({Base_Year})", yaxis_title="Share")
    fig_sc = plotly_stacked_bar(state.df_TransportFuel_SC[state.df_TransportFuel_SC.index != 'Bike'].T, f"Fuel Shares - Scenario ({state.Scenario_Year})", yaxis_title="Share")

    def fuel_table(df_fuel):
        cols = [{'name': 'Fuel', 'id': 'Fuel'}, {'name': 'Used [TJ]', 'id': 'Used'}, {'name': 'Share', 'id': 'Share'}]
        data = []
        for idx in df_fuel.index:
            used = df_fuel.loc[idx, 'Total transport consumption']
            share = df_fuel.loc[idx, 'Share of total transport consumption']
            data.append({'Fuel': idx, 'Used': f"{used:.0f}", 'Share': f"{share * 100:.2f}%"})
        return cols, data

    cols_by, data_by = fuel_table(df_transport_consumption_fuels_BY)
    cols_sc, data_sc = fuel_table(state.df_transport_consumption_fuels_SC)

    return html.Div([
        html.H2("Transportation Sector Data", style=HEADER_STYLE),
        html.P("Fuel mix and total transport energy consumption across all modes.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            html.Div([
                html.H4(f"Base Year ({Base_Year})", style=SUBHEADER_STYLE),
                dcc.Graph(figure=fig_by, style={'height': '360px'}, config={'responsive': True}),
                dash_table.DataTable(columns=cols_by, data=data_by, style_header=TABLE_STYLE_HEADER, style_cell=TABLE_STYLE_CELL, style_table={'overflowX': 'auto'})
            ], style={'flex': '1 1 450px', 'minWidth': '380px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'}),

            html.Div([
                html.H4(f"Scenario Target Year ({state.Scenario_Year})", style=SUBHEADER_STYLE),
                dcc.Graph(figure=fig_sc, style={'height': '360px'}, config={'responsive': True}),
                dash_table.DataTable(columns=cols_sc, data=data_sc, style_header=TABLE_STYLE_HEADER, style_cell=TABLE_STYLE_CELL, style_table={'overflowX': 'auto'})
            ], style={'flex': '1 1 450px', 'minWidth': '380px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '20px'})
    ])

def build_co2_page():
    def make_co2_table(df_co2):
        cols = [{'name': 'Fuel', 'id': 'Fuel'}]
        for c in df_co2.columns[:-2]:
            cols.append({'name': str(c), 'id': str(c)})
        cols.append({'name': 'Total no CCS', 'id': 'Total emissions excl. CCS'})
        cols.append({'name': 'Total w. CCS', 'id': 'Total emissions incl. CCS'})

        data = []
        for idx in df_co2.index:
            row = {'Fuel': idx}
            for c in df_co2.columns:
                val = df_co2.loc[idx, c]
                row[c] = f"{val:.3f}" if (not np.isnan(val) and abs(val) > 1e-4) else ""
            data.append(row)
        return cols, data

    cols_by, data_by = make_co2_table(df_fuel_CO2_BY)
    cols_sc, data_sc = make_co2_table(state.df_fuel_CO2_SC)

    return html.Div([
        html.H2("Greenhouse Gas & CO2 Emissions", style=HEADER_STYLE),
        html.P("Sector-by-sector breakdown of emissions with and without Carbon Capture and Storage (CCS).", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            html.H4(f"Base Year ({Base_Year}) Emissions [million tonnes CO2]", style=SUBHEADER_STYLE),
            dash_table.DataTable(columns=cols_by, data=data_by, style_header=TABLE_STYLE_HEADER, style_cell=TABLE_STYLE_CELL, style_table={'overflowX': 'auto'})
        ], style={**CARD_STYLE, 'marginBottom': '20px'}),
        
        html.Div([
            html.H4(f"Scenario Target Year ({state.Scenario_Year}) Emissions [million tonnes CO2]", style=SUBHEADER_STYLE),
            dash_table.DataTable(columns=cols_sc, data=data_sc, style_header=TABLE_STYLE_HEADER, style_cell=TABLE_STYLE_CELL, style_table={'overflowX': 'auto'})
        ], style=CARD_STYLE)
    ])

def build_economy_page():
    price_rows = []
    for fuel in state.df_FuelPrice_SC.index:
        editable = state.df_FuelPrice_SC.loc[fuel, 'Editable [y/n?]'] == 'y'
        price = state.df_FuelPrice_SC.loc[fuel, 'Market price [€/GJ]']
        price_rows.append({
            'Fuel': fuel,
            'Price': f"{price:.2f}",
            'Editable': editable
        })
    price_cols = [
        {'name': 'Fuel', 'id': 'Fuel', 'editable': False},
        {'name': 'Market Price [€/GJ]', 'id': 'Price', 'editable': True}
    ]

    heading_costs = ['Electricity', 'District Heat', 'Fuel Refinery', 'Residential', 'Tertiary', 'Industry', 'Agriculture', 'Transport']
    costs_df = pd.DataFrame(np.vstack((
        state.df_capital_costs_SC.sum(axis=0),
        state.df_om_costs_SC.sum(axis=0),
        state.df_fuel_costs_SC.sum(axis=0),
        state.df_co2_costs_SC.sum(axis=0)
    )), columns=heading_costs, index=['Capital costs', 'O&M costs', 'Fuel costs', 'CO2 costs'])

    fig_costs = plotly_stacked_bar(costs_df, f"Annual System Cost Structure ({state.Scenario_Year})", yaxis_title="[million €]")

    return html.Div([
        html.H2("System Economy & Annual Costs", style=HEADER_STYLE),
        html.P("Fuel market pricing assumptions and total annualized system expenditure.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            # Left: Fuel Prices & CO2 Price
            html.Div([
                html.H4("Fuel & CO2 Prices (Scenario)", style=SUBHEADER_STYLE),
                dash_table.DataTable(
                    id='table-fuel-prices',
                    columns=price_cols,
                    data=price_rows,
                    editable=True,
                    style_header=TABLE_STYLE_HEADER,
                    style_cell=TABLE_STYLE_CELL,
                    style_table={'overflowX': 'auto'},
                    style_data_conditional=[{'if': {'filter_query': '{Editable} eq false'}, 'backgroundColor': '#F1F5F9'}]
                ),
                html.Div([
                    html.Span("CO2 Price [€/t CO2]: ", style={'fontWeight': '600', 'color': dtu_navy, 'fontSize': '13px'}),
                    dcc.Input(
                        id="input-co2-price", type="number", value=state.co2_price_SC,
                        style={'width': '95px', 'marginLeft': '10px', 'padding': '6px', 'borderRadius': '4px', 'border': f'1px solid {border_subtle}'}
                    ),
                ], style={'marginTop': '16px'}),
                html.Button("💾 Save Price Changes", id="btn-save-prices", style={**ACTION_BTN_STYLE, 'marginTop': '16px'})
            ], style={'flex': '2 1 340px', 'minWidth': '280px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'}),
            
            # Right: Costs Plot
            html.Div([
                html.H4("Annual System Costs Breakdown", style=SUBHEADER_STYLE),
                dcc.Graph(figure=fig_costs, style={'height': '420px'}, config={'responsive': True})
            ], style={'flex': '3 1 480px', 'minWidth': '380px', 'backgroundColor': bg_card, 'padding': '16px', 'borderRadius': '8px', 'border': f'1px solid {border_subtle}', 'boxSizing': 'border-box'})
        ], style={'display': 'flex', 'flexWrap': 'wrap', 'gap': '20px'})
    ])

def build_timeseries_page():
    return html.Div([
        html.H2(f"Time Series & Dispatch Simulation - {state.Scenario_Year}", style=HEADER_STYLE),
        html.P("Calculates hourly power system dispatch, load duration curves, and intermittent renewable balancing.", style={'color': text_muted, 'fontSize': '13px', 'marginBottom': '16px'}),
        
        html.Div([
            html.Button("⚡ Run 8760h Time Series Simulation", id="btn-run-timeseries", style={
                **ACTION_BTN_STYLE, 'backgroundColor': dtu_navy, 'fontSize': '14px', 'padding': '10px 24px', 'border': f'1px solid {dtu_energy_blue}'
            })
        ], style={'marginBottom': '20px'}),

        dcc.Loading(
            id="loading-timeseries",
            type="circle",
            color=dtu_energy_blue,
            children=html.Div(id="timeseries-graphs-container")
        )
    ])

# --- Router Callback ---
@app.callback(
    Output("page-content", "children"),
    Input("url", "pathname"),
    Input("alert-banner", "children")
)
def display_page(pathname, alert_msg):
    if pathname in ["/", "/graphs"]:
        return build_graphs_page()
    elif pathname == "/consumption":
        return build_consumption_page()
    elif pathname == "/production":
        return build_production_page()
    elif pathname == "/by":
        return build_data_view_page(is_base_year=True)
    elif pathname == "/scenario":
        return build_data_view_page(is_base_year=False)
    elif pathname == "/transport":
        return build_transport_page()
    elif pathname == "/co2":
        return build_co2_page()
    elif pathname == "/economy":
        return build_economy_page()
    elif pathname == "/timeseries":
        return build_timeseries_page()
    return html.Div([html.H3("404: Page not found")])

# --- Graphs Page Sub-Tab Callback ---
@app.callback(
    Output("graph-by-display", "figure"),
    Output("graph-sc-display", "figure"),
    Input("graph-sub-tabs", "value")
)
def update_graphs_tab(selected_tab):
    if selected_tab == "sankey":
        fig_by = create_sankey_figure(df_fuelType_consumption_BY, df_Fc_BY, f"Energy streams - {Base_Year}")
        fig_sc = create_sankey_figure(state.df_fuelType_consumption_SC, state.df_Fc_SC, f"Energy streams - {state.Scenario_Year}")
        return fig_by, fig_sc
    elif selected_tab == "fec_stacked":
        fig_by = plotly_stacked_bar(df_FEC_BY.iloc[:, :4], f"FEC per Sector - {Base_Year}")
        fig_sc = plotly_stacked_bar(state.df_FEC_SC.iloc[:, :4], f"FEC per Sector - {state.Scenario_Year}")
        return fig_by, fig_sc
    elif selected_tab == "fec_all":
        sec_by = df_FEC_BY.iloc[:, :4].sum() * 0.001
        sec_by.loc['Transport'] = df_transport_consumption_BY.sum().values[0] * 0.001
        sec_sc = state.df_FEC_SC.iloc[:, :4].sum() * 0.001
        sec_sc.loc['Transport'] = state.df_transport_consumption_SC.sum().values[0] * 0.001
        fig_by = plotly_bar(sec_by, f"FEC per Sector (PJ) - {Base_Year}")
        fig_sc = plotly_bar(sec_sc, f"FEC per Sector (PJ) - {state.Scenario_Year}")
        return fig_by, fig_sc
    elif selected_tab == "co2":
        new_names = {
            'Electricity production no CCS': 'Electricity no CCS',
            'Electricity production w. CCS': 'Electricity w. CCS',
            'Total emissions incl. CCS': 'Total w. CCS',
            'Total emissions excl. CCS': 'Total no CCS',
        }
        co2_by = df_fuelType_CO2_BY.sum().rename(new_names)
        co2_sc = state.df_fuelType_CO2_SC.sum().rename(new_names)
        fig_by = plotly_bar(co2_by, f"CO2 per Sector (mill tons) - {Base_Year}", yaxis_title="mill tons")
        fig_sc = plotly_bar(co2_sc, f"CO2 per Sector (mill tons) - {state.Scenario_Year}", yaxis_title="mill tons")
        return fig_by, fig_sc
    elif selected_tab == "el":
        fig_by = plotly_bar(df_Power_BY['Amount'], f"Electricity Production - {Base_Year}")
        fig_sc = plotly_bar(state.df_Power_SC['Amount'], f"Electricity Production - {state.Scenario_Year}")
        return fig_by, fig_sc
    elif selected_tab == "dh":
        fig_by = plotly_bar(df_Heat_BY['Amount'], f"District Heat Production - {Base_Year}")
        fig_sc = plotly_bar(state.df_Heat_SC['Amount'], f"District Heat Production - {state.Scenario_Year}")
        return fig_by, fig_sc
    elif selected_tab == "fuel":
        f_by = df_fuelType_consumption_BY['Total primary fuel consumption'].drop(['Electricity', 'District Heating'], errors='ignore')
        f_sc = state.df_fuelType_consumption_SC['Total primary fuel consumption'].drop(['Electricity', 'District Heating'], errors='ignore')
        fig_by = plotly_bar(f_by, f"Primary Fuel Consumption - {Base_Year}")
        fig_sc = plotly_bar(f_sc, f"Primary Fuel Consumption - {state.Scenario_Year}")
        return fig_by, fig_sc
    return go.Figure(), go.Figure()

# --- Dedicated Save Callbacks for Individual Pages ---

# 1. Consumption Page Save Callback
@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-save-consumption", "n_clicks"),
    State("table-heat-consumption", "data"),
    State("table-transport-work", "data"),
    State("table-transport-fuels", "data"),
    prevent_initial_call=True
)
def save_consumption(n_clicks, heat_data, work_data, tf_data):
    if not n_clicks:
        return no_update
    try:
        if heat_data:
            for row in heat_data:
                tech = row['Tech Type']
                for s in SECTORS_LIST:
                    val = float(str(row[s]).rstrip('%').replace(',', '.')) / 100.0
                    state.df_beta_SC.loc[tech, s] = val

            for s in SECTORS_LIST:
                total = state.df_beta_SC[s].iloc[1:].sum()
                if abs(total - 1.0) > 0.01:
                    return html.Div(
                        f"⚠️ Heating sum for {s} is {total*100:.1f}% (must be 100%). Changes not saved.",
                        style={'color': dtu_red, 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#FEF2F2', 'borderRadius': '6px', 'border': '1px solid #FCA5A5'}
                    )

        if work_data:
            for row in work_data:
                idx = row['Type']
                state.df_TransportWork_SC.loc[idx, 'Allocation of transport work'] = float(str(row['Allocation']).rstrip('%').replace(',', '.')) / 100.0
                state.df_TransportWork_SC.loc[idx, 'Utilisation degree'] = float(str(row['Utilisation']).rstrip('%').replace(',', '.')) / 100.0

        if tf_data:
            for row in tf_data:
                t = row['Type']
                for col in state.df_TransportFuel_SC.columns:
                    if col in row:
                        state.df_TransportFuel_SC.loc[t, col] = float(str(row[col]).rstrip('%').replace(',', '.')) / 100.0

        state.recalculate()
        state.save_scenario_file(SCData)
        return html.Div(
            "✓ Consumption changes successfully saved and scenario recalculated!",
            style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
        )
    except Exception as e:
        return html.Div(
            f"❌ Error saving consumption: {str(e)}",
            style={'color': dtu_red, 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#FEF2F2', 'borderRadius': '6px', 'border': '1px solid #FCA5A5'}
        )

# 2. Production Page Save Callback
@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-save-production", "n_clicks"),
    State("table-el-production", "data"),
    State("table-dh-production", "data"),
    prevent_initial_call=True
)
def save_production(n_clicks, el_data, dh_data):
    if not n_clicks:
        return no_update
    try:
        if el_data:
            for row in el_data:
                m = row['Method']
                val = float(str(row['Shares']).rstrip('%').replace(',', '.')) / 100.0
                state.df_Power_SC.loc[m, 'Shares'] = val
            state.df_Power_SC.loc[state.df_Power_SC.index[-1], 'Shares'] = round(1 - state.df_Power_SC.iloc[:-1]['Shares'].sum(), 10)
            state.df_Power_SC['Amount'] = state.df_Power_SC['Shares'] * state.total_electricity_consumption_SC

        if dh_data:
            for row in dh_data:
                m = row['Method']
                val = float(str(row['Shares']).rstrip('%').replace(',', '.')) / 100.0
                state.df_Heat_SC.loc[m, 'Shares'] = val
            state.df_Heat_SC.loc[state.df_Heat_SC.index[-1], 'Shares'] = round(1 - state.df_Heat_SC.iloc[:-1]['Shares'].sum(), 10)

        state.recalculate()
        state.save_scenario_file(SCData)
        return html.Div(
            "✓ Production changes successfully saved and scenario recalculated!",
            style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
        )
    except Exception as e:
        return html.Div(
            f"❌ Error saving production: {str(e)}",
            style={'color': dtu_red, 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#FEF2F2', 'borderRadius': '6px', 'border': '1px solid #FCA5A5'}
        )

# 3. Economy Page Save Callback
@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-save-prices", "n_clicks"),
    State("table-fuel-prices", "data"),
    State("input-co2-price", "value"),
    prevent_initial_call=True
)
def save_prices(n_clicks, price_data, co2_val):
    if not n_clicks:
        return no_update
    try:
        if price_data:
            for row in price_data:
                f = row['Fuel']
                if row.get('Editable'):
                    val = float(str(row['Price']).replace(',', '.'))
                    state.df_FuelPrice_SC.loc[f, 'Market price [€/GJ]'] = val
        if co2_val is not None:
            state.co2_price_SC = float(co2_val)

        state.recalculate()
        state.save_scenario_file(SCData)
        return html.Div(
            "✓ Price changes successfully saved and scenario recalculated!",
            style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
        )
    except Exception as e:
        return html.Div(
            f"❌ Error saving prices: {str(e)}",
            style={'color': dtu_red, 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#FEF2F2', 'borderRadius': '6px', 'border': '1px solid #FCA5A5'}
        )

# 4. Top Banner & Global Update Callbacks
@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-update-year", "n_clicks"),
    State("input-scenario-year", "value"),
    prevent_initial_call=True
)
def update_year(n_clicks, new_year):
    if not n_clicks:
        return no_update
    if new_year is None or new_year <= Base_Year or new_year > 2100:
        return html.Div(
            f"⚠️ Invalid year: Must be between {Base_Year} and 2100.",
            style={'color': dtu_red, 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#FEF2F2', 'borderRadius': '6px', 'border': '1px solid #FCA5A5'}
        )
    state.Scenario_Year = int(new_year)
    workbook = openpyxl.load_workbook(state.sc_file)
    workbook['Scenario Year'].cell(row=2, column=1, value=state.Scenario_Year)
    workbook.save(state.sc_file)
    
    Backend.SCData = state.sc_file
    state.SC_NEC_tFE_el, state.SC_NEC_tFE_heat, state.df_beta_BY, state.engr = BY_calculations()
    Backend.engr = state.engr
    set_SC_transport_demand(state.df_Transport_SC, state.engr)
    state.recalculate()
    state.save_scenario_file()
    return html.Div(
        f"✓ Scenario Target Year updated to {state.Scenario_Year} and calculations refreshed!",
        style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
    )

@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-global-update", "n_clicks"),
    prevent_initial_call=True
)
def global_update(n_clicks):
    if not n_clicks:
        return no_update
    state.recalculate()
    state.save_scenario_file()
    return html.Div(
        "✓ Scenario successfully recalculated and saved!",
        style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
    )

@app.callback(
    Output("alert-banner", "children", allow_duplicate=True),
    Input("btn-reset-sc", "n_clicks"),
    prevent_initial_call=True
)
def reset_to_default(n_clicks):
    if not n_clicks:
        return no_update
    shutil.copyfile(DEFAULT_SCData, state.sc_file)
    state.reset_to_current()
    return html.Div(
        "✓ Scenario reset to DTU default benchmark!",
        style={'color': '#065F46', 'fontWeight': '600', 'padding': '10px 14px', 'backgroundColor': '#D1FAE5', 'borderRadius': '6px', 'border': '1px solid #6EE7B7'}
    )

# --- Time Series Execution Callback ---
@app.callback(
    Output("timeseries-graphs-container", "children"),
    Input("btn-run-timeseries", "n_clicks"),
    prevent_initial_call=True
)
def compute_timeseries(n_clicks):
    if not n_clicks:
        return no_update
    df_hour, technologies_used = run_timeSeries(
        state.df_Power_SC, state.df_FuelPrice_SC, state.df_FEC_SC,
        state.total_electricity_demand_SC, state.el_transport_SC, state.el_DH_SC, state.el_FR_SC
    )
    state.df_hour = df_hour
    state.technologies_used = technologies_used

    # 1. Load duration curve plot
    sorted_cons = df_hour['Consumption after renewables (unflexible and flexible)'].sort_values(ascending=False).to_numpy()
    fig_dur = go.Figure()
    fig_dur.add_trace(go.Scatter(y=sorted_cons, mode='lines', name='Residual Duration Curve', line=dict(color=dtu_navy, width=2.5)))
    
    # Add flexible technologies horizontal bars
    bar_h = technologies_used['Avg'].tolist()
    bar_w = technologies_used['Hours'].tolist()
    tech_labels = technologies_used.index.tolist()
    cum_h = 0
    for i in range(len(tech_labels)):
        y_pos = cum_h + bar_h[i]/2
        cum_h += bar_h[i]
        fig_dur.add_trace(go.Bar(
            y=[y_pos], x=[bar_w[i]], orientation='h',
            name=tech_labels[i], marker_color=PLOT_COLORS[i % len(PLOT_COLORS)], opacity=0.75
        ))
    fig_dur.update_layout(
        title="Hourly Load Duration Curve & Flexible Dispatch",
        title_font=dict(size=15, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        xaxis_title="Hours [h]", yaxis_title="Average Power [MW]",
        paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=420
    )
    fig_dur.update_yaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')
    fig_dur.update_xaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')

    # 2. Average Week
    df_copy = df_hour.copy()
    df_copy['Exported electricity'] = -abs(df_copy['Exported electricity'])
    df_copy['HourOfWeek'] = df_copy.index % 168
    pos_techs = [t for t in state.df_Power_SC.index if t in df_copy.columns]
    avg_week = df_copy.groupby('HourOfWeek')[pos_techs].mean()
    avg_week_exp = df_copy.groupby('HourOfWeek')['Exported electricity'].mean()

    fig_week = go.Figure()
    for i, tech in enumerate(pos_techs):
        fig_week.add_trace(go.Scatter(
            x=list(avg_week.index), y=avg_week[tech], name=tech,
            stackgroup='one', line=dict(width=0.5, color=TIMESERIES_COLORS[i % len(TIMESERIES_COLORS)])
        ))
    fig_week.add_trace(go.Scatter(
        x=list(avg_week_exp.index), y=avg_week_exp, name='Exported electricity',
        line=dict(color=dtu_red, width=1.8, dash='dot')
    ))
    fig_week.update_layout(
        title="Average Weekly Power Dispatch by Technology (168h Profile)",
        title_font=dict(size=15, family="Inter, system-ui, sans-serif", color=dtu_navy),
        title_x=0.02,
        xaxis=dict(tickmode='array', tickvals=list(range(12, 169, 24)), ticktext=['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']),
        yaxis_title="Power [MW]", paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)', height=420
    )
    fig_week.update_yaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')
    fig_week.update_xaxes(gridcolor='#F1F5F9', zerolinecolor='#E2E8F0')

    return html.Div([
        html.Div([dcc.Graph(figure=fig_dur, config={'responsive': True})], style={**CARD_STYLE, 'marginBottom': '20px'}),
        html.Div([dcc.Graph(figure=fig_week, config={'responsive': True})], style=CARD_STYLE)
    ])

# --- Excel Download & Upload Callbacks ---
@app.callback(
    Output("download-scenario-xlsx", "data"),
    Input("btn-download-sc", "n_clicks"),
    prevent_initial_call=True
)
def download_scenario_excel(n_clicks):
    if not n_clicks:
        return no_update
    state.save_scenario_file()
    return dcc.send_file(state.sc_file, filename=f"pySTREAM_Scenario_{state.Scenario_Year}.xlsx")

@app.callback(
    Output("url", "pathname"),
    Input("upload-scenario", "contents"),
    State("upload-scenario", "filename"),
    prevent_initial_call=True
)
def upload_scenario_excel(contents, filename):
    if contents is None:
        return no_update
    content_type, content_string = contents.split(',')
    decoded = base64.b64decode(content_string)
    
    with open(state.sc_file, 'wb') as f:
        f.write(decoded)
    
    state.reset_to_current()
    return "/scenario"

# --- Main Entry Point ---
# Expose the Flask WSGI server instance for gunicorn / Hugging Face Spaces / production servers
server = app.server

if __name__ == '__main__':
    # Hugging Face Spaces defaults to port 7860, local/other platforms default to 8050
    # port = int(os.environ.get('PORT', 7860))
    # print(f"Starting pySTREAM Web Edition on http://0.0.0.0:{port}/")
    # app.run(host='0.0.0.0', port=port, debug=False)
    port = int(os.environ.get('PORT', 8050))
    print(f"Starting pySTREAM Web Edition on http://127.0.0.1:{port}/")
    app.run(host='127.0.0.1', port=port, debug=False)
