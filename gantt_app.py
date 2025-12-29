import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime
import io
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.backends.backend_pdf import PdfPages
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib import colors as rl_colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet

# Page config
st.set_page_config(
    page_title="Production Gantt Chart", 
    layout="wide", 
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
    <style>
    .main {background-color: #f8f9fa;}
    .stMarkdown, .stText, p, span, label {color: #1e1e1e !important;}
    h1, h2, h3, h4, h5, h6 {color: #2c3e50 !important;}
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white !important;
        font-weight: 600;
        border: none;
        padding: 12px;
        border-radius: 10px;
    }
    .stFileUploader {
        background-color: white;
        border: 2px solid #667eea;
        border-radius: 10px;
        padding: 20px;
    }
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
    }
    .stMetric label {color: #2c3e50 !important; font-weight: 600 !important;}
    .stDownloadButton>button {background-color: #28a745; color: white !important;}
    </style>
""", unsafe_allow_html=True)

# Title
st.markdown("<h1 style='text-align: center; color: #2c3e50;'>📊 Production Schedule - Gantt Chart</h1>", unsafe_allow_html=True)
st.markdown("---")

# File uploader
st.markdown("<div style='background-color: white; padding: 20px; border-radius: 10px; border: 2px solid #667eea;'>", unsafe_allow_html=True)
st.markdown("<h3 style='color: #2c3e50;'>📁 Ανέβασε το Excel αρχείο σου</h3>", unsafe_allow_html=True)
st.markdown("<p style='color: #666;'>Το αρχείο πρέπει να έχει φύλλο 'schedule' με στήλες: Description, Start Time, End Time</p>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Επιλογή αρχείου", 
    type=['xlsx', 'xls'],
    label_visibility="collapsed"
)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

def create_gantt_chart_matplotlib(df_sorted):
    """Create Gantt chart using matplotlib for PDF export"""
    fig, ax = plt.subplots(figsize=(14, max(8, len(df_sorted) * 0.4)))
    
    colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A4C93', 
              '#06A77D', '#D90368', '#F08700', '#0E9594', '#8B2635']
    
    y_pos = range(len(df_sorted))
    
    for idx, (i, row) in enumerate(df_sorted.iterrows()):
        start = mdates.date2num(row['Start Time'])
        end = mdates.date2num(row['End Time'])
        duration = end - start
        color = colors[idx % len(colors)]
        
        ax.barh(idx, duration, left=start, height=0.8, 
                color=color, edgecolor='white', linewidth=2)
    
    # Format
    ax.set_yticks(y_pos)
    ax.set_yticklabels(df_sorted['displayLabel'].tolist(), fontsize=9)
    ax.invert_yaxis()
    
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%d/%m %H:%M'))
    ax.xaxis.set_major_locator(mdates.AutoDateLocator())
    plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha='right')
    
    ax.set_xlabel('Χρονική Περίοδος', fontsize=12, fontweight='bold')
    ax.set_ylabel('Ενέργειες / Υλικά', fontsize=12, fontweight='bold')
    ax.set_title('Production Schedule - Gantt Chart', fontsize=16, fontweight='bold', pad=20)
    
    ax.grid(True, axis='x', alpha=0.3, linestyle='--')
    ax.set_facecolor('#fafafa')
    
    plt.tight_layout()
    return fig

if uploaded_file is not None:
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        available_sheets = excel_file.sheet_names
        
        # Sidebar
        st.sidebar.markdown("<h2 style='color: #2c3e50;'>⚙️ Ρυθμίσεις</h2>", unsafe_allow_html=True)
        
        # Sheet selection
        if 'schedule' in [s.lower() for s in available_sheets]:
            default_sheet = [s for s in available_sheets if s.lower() == 'schedule'][0]
            default_index = available_sheets.index(default_sheet)
        else:
            default_index = 0
        
        selected_sheet = st.sidebar.selectbox("Επιλογή Φύλλου:", available_sheets, index=default_index)
        
        # Read sheet
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        st.success(f"✅ Φύλλο '{selected_sheet}' φορτώθηκε επιτυχώς με {len(df)} γραμμές!")
        
        # Preview
        with st.expander("📋 Προεπισκόπηση Δεδομένων"):
            st.dataframe(df.head(10), use_container_width=True)
        
        # Check columns
        required_cols = ['Description', 'Start Time', 'End Time']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Λείπουν οι στήλες: {', '.join(missing_cols)}")
            st.info(f"📌 Διαθέσιμες στήλες: {', '.join(df.columns.tolist())}")
        else:
            # Convert to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'])
            df['End Time'] = pd.to_datetime(df['End Time'])
            
            # Filter
            df = df.dropna(subset=['Description', 'Start Time', 'End Time'])
            
            # Add IDs
            df['uniqueId'] = range(1, len(df) + 1)
            df['displayLabel'] = df['uniqueId'].astype(str) + '. ' + df['Description'].astype(str)
            
            # Sorting
            sort_order = st.sidebar.radio("Ταξινόμηση κατά End Time:", ["Αύξουσα", "Φθίνουσα"])
            ascending = True if sort_order == "Αύξουσα" else False
            df_sorted = df.sort_values('End Time', ascending=ascending).reset_index(drop=True)
            
            # Update labels
            df_sorted['displayLabel'] = range(1, len(df_sorted) + 1)
            df_sorted['displayLabel'] = df_sorted['displayLabel'].astype(str) + '. ' + df_sorted['Description'].astype(str)
            
            # Shift filter
            if 'Shift' in df.columns:
                shifts = ['Όλα'] + sorted(df_sorted['Shift'].dropna().unique().tolist())
                selected_shift = st.sidebar.selectbox("Φίλτρο Shift:", shifts)
                if selected_shift != 'Όλα':
                    df_sorted = df_sorted[df_sorted['Shift'] == selected_shift]
            
            # Calculate duration
            df_sorted['Duration_hours'] = (df_sorted['End Time'] - df_sorted['Start Time']).dt.total_seconds() / 3600
            
            # Stats
            st.markdown("<br>", unsafe_allow_html=True)
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("🎯 Σύνολο Ενεργειών", len(df_sorted))
            with col2:
                total_hours = df_sorted['Duration_hours'].sum()
                st.metric("⏱️ Συνολικές Ώρες", f"{total_hours:.1f}")
            with col3:
                if len(df_sorted) > 0:
                    date_range = (df_sorted['End Time'].max() - df_sorted['Start Time'].min()).days
                    st.metric("📅 Διάρκεια (ημέρες)", date_range)
            
            st.markdown("---")
            
            # Create Plotly Gantt Chart for display
            fig = go.Figure()
            
            colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A4C93', 
                      '#06A77D', '#D90368', '#F08700', '#0E9594', '#8B2635']
            
            for idx, row in df_sorted.iterrows():
                color = colors[idx % len(colors)]
                
                start_ts = row['Start Time'].value / 1000000
                end_ts = row['End Time'].value / 1000000
                duration_ms = end_ts - start_ts
                
                hover_text = f"<b>{row['Description']}</b><br>"
                hover_text += f"Start: {row['Start Time'].strftime('%d/%m/%Y %H:%M')}<br>"
                hover_text += f"End: {row['End Time'].strftime('%d/%m/%Y %H:%M')}<br>"
                hover_text += f"Duration: {row['Duration_hours']:.2f} ώρες<br>"
                
                if 'Shift' in df.columns and pd.notna(row.get('Shift')):
                    hover_text += f"Shift: {row['Shift']}<br>"
                if 'Qnt' in df.columns and pd.notna(row.get('Qnt')):
                    hover_text += f"Quantity: {row['Qnt']}<br>"
                if 'Capacity/hr' in df.columns and pd.notna(row.get('Capacity/hr')):
                    hover_text += f"Capacity/hr: {row['Capacity/hr']}<br>"
                if 'Prod. Time' in df.columns and pd.notna(row.get('Prod. Time')):
                    hover_text += f"Prod. Time: {row['Prod. Time']}<br>"
                
                fig.add_trace(go.Bar(
                    x=[duration_ms],
                    y=[row['displayLabel']],
                    base=row['Start Time'],
                    orientation='h',
                    marker=dict(color=color, line=dict(color='white', width=2)),
                    name=row['Description'],
                    hovertemplate=hover_text + '<extra></extra>',
                    showlegend=False
                ))
            
            # Layout
            fig.update_layout(
                title={
                    'text': 'Production Schedule - Gantt Chart',
                    'x': 0.5,
                    'xanchor': 'center',
                    'font': {'size': 26, 'color': '#1e1e1e', 'family': 'Arial Black'}
                },
                xaxis_title='Χρονική Περίοδος',
                yaxis_title='Ενέργειες / Υλικά',
                xaxis=dict(
                    type='date',
                    tickformat='%d/%m %H:%M',
                    tickangle=-45,
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='#d0d0d0',
                    tickfont=dict(size=12, color='#1e1e1e'),
                    title_font=dict(size=14, color='#1e1e1e')
                ),
                yaxis=dict(
                    autorange='reversed',
                    categoryorder='array',
                    categoryarray=df_sorted['displayLabel'].tolist(),
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='#d0d0d0',
                    tickfont=dict(size=11, color='#1e1e1e'),
                    title_font=dict(size=14, color='#1e1e1e')
                ),
                height=max(600, len(df_sorted) * 40),
                hovermode='closest',
                plot_bgcolor='#fafafa',
                paper_bgcolor='white',
                margin=dict(l=300, r=50, t=100, b=100),
                font=dict(color='#1e1e1e')
            )
            
            # Display
            st.plotly_chart(fig, use_container_width=True)
            
            # Download buttons
            st.markdown("---")
            col1, col2, col3 = st.columns([2, 1, 1])
            
            with col2:
                # CSV Export
                df_export = df_sorted.copy()
                df_export = df_export.drop(columns=['uniqueId', 'displayLabel'], errors='ignore')
                csv = df_export.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Κατέβασε CSV",
                    data=csv,
                    file_name='gantt_schedule.csv',
                    mime='text/csv',
                )
            
            with col3:
                # PDF Export with matplotlib chart
                try:
                    # Create matplotlib chart
                    chart_fig = create_gantt_chart_matplotlib(df_sorted)
                    
                    # Save to PDF
                    pdf_buffer = io.BytesIO()
                    with PdfPages(pdf_buffer) as pdf:
                        pdf.savefig(chart_fig, bbox_inches='tight', dpi=300)
                    
                    plt.close(chart_fig)
                    
                    pdf_bytes = pdf_buffer.getvalue()
                    pdf_buffer.close()
                    
                    st.download_button(
                        label="📄 Κατέβασε PDF",
                        data=pdf_bytes,
                        file_name='gantt_schedule.pdf',
                        mime='application/pdf',
                    )
                except Exception as e:
                    st.error(f"PDF export error: {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
        
    except Exception as e:
        st.error(f"❌ Σφάλμα: {str(e)}")
        import traceback
        st.code(traceback.format_exc())

else:
    # Instructions
    st.markdown("""
    <div style='background-color: #d1ecf1; padding: 25px; border-radius: 10px; border-left: 5px solid #0c5460;'>
        <h3 style='color: #0c5460; margin-top: 0;'>📝 Οδηγίες Χρήσης</h3>
        <p style='color: #0c5460; font-size: 16px; line-height: 1.6;'>
            <strong>1.</strong> Ανέβασε Excel (.xlsx) με φύλλο <strong>"schedule"</strong><br>
            <strong>2.</strong> Απαραίτητες στήλες: <code>Description</code>, <code>Start Time</code>, <code>End Time</code><br>
            <strong>3.</strong> Προαιρετικές: <code>Shift</code>, <code>Qnt</code>, <code>Capacity/hr</code>, <code>Prod. Time</code>
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    with st.expander("💡 Παράδειγμα Δεδομένων"):
        sample_df = pd.DataFrame({
            'Shift': ['Morning', 'Morning', 'Evening'],
            'Start Time': ['29/12/2025 06:00', '29/12/2025 08:00', '29/12/2025 14:00'],
            'End Time': ['29/12/2025 08:00', '29/12/2025 10:00', '29/12/2025 18:00'],
            'Description': ['ΒΡΩΜΗ ΣΕ ΣΑΚΙ', 'Αλλαγή Υλικού', 'ΚΑΛΑΜΠΟΚΙ'],
            'Qnt': [1000, 0, 1500],
            'Capacity/hr': [500, 0, 600]
        })
        st.dataframe(sample_df, use_container_width=True)

st.markdown("---")
st.markdown("<div style='text-align: center; color: #666;'>Production Gantt Chart | Streamlit & Plotly</div>", unsafe_allow_html=True)
