import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime

# Page config
st.set_page_config(
    page_title="Production Gantt Chart", 
    layout="wide", 
    page_icon="📊",
    initial_sidebar_state="expanded"
)

# IMPROVED Custom CSS for better visibility
st.markdown("""
    <style>
    /* Main background */
    .main {
        background-color: #f8f9fa;
    }
    
    /* Text colors */
    .stMarkdown, .stText, p, span, label {
        color: #1e1e1e !important;
    }
    
    /* Headers */
    h1, h2, h3, h4, h5, h6 {
        color: #2c3e50 !important;
    }
    
    /* Buttons */
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white !important;
        font-weight: 600;
        border: none;
        padding: 12px;
        border-radius: 10px;
    }
    
    /* File uploader */
    .stFileUploader {
        background-color: white;
        border: 2px solid #667eea;
        border-radius: 10px;
        padding: 20px;
    }
    
    /* Dataframe */
    .stDataFrame {
        background-color: white;
        border-radius: 10px;
        padding: 10px;
    }
    
    /* Metrics */
    .stMetric {
        background-color: white;
        padding: 15px;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
    }
    
    .stMetric label {
        color: #2c3e50 !important;
        font-weight: 600 !important;
    }
    
    .stMetric .metric-value {
        color: #667eea !important;
        font-size: 28px !important;
        font-weight: 700 !important;
    }
    
    /* Sidebar */
    .css-1d391kg, [data-testid="stSidebar"] {
        background-color: #f0f2f6;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background-color: white;
        border: 1px solid #e0e0e0;
        border-radius: 8px;
        color: #2c3e50 !important;
        font-weight: 600 !important;
    }
    
    /* Success/Error boxes */
    .stSuccess {
        background-color: #d4edda;
        color: #155724 !important;
        border-radius: 8px;
        padding: 15px;
    }
    
    .stError {
        background-color: #f8d7da;
        color: #721c24 !important;
        border-radius: 8px;
        padding: 15px;
    }
    
    .stInfo {
        background-color: #d1ecf1;
        color: #0c5460 !important;
        border-radius: 8px;
        padding: 15px;
    }
    
    /* Select boxes and inputs */
    .stSelectbox, .stRadio {
        background-color: white;
    }
    
    /* Better contrast for all text */
    div[data-testid="stMarkdownContainer"] p {
        color: #2c3e50 !important;
    }
    
    /* Download button */
    .stDownloadButton>button {
        background-color: #28a745;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

# Title with better styling
st.markdown("<h1 style='text-align: center; color: #2c3e50;'>📊 Production Schedule - Gantt Chart</h1>", unsafe_allow_html=True)
st.markdown("---")

# File uploader with clear instructions
st.markdown("<div style='background-color: white; padding: 20px; border-radius: 10px; border: 2px solid #667eea;'>", unsafe_allow_html=True)
st.markdown("<h3 style='color: #2c3e50;'>📁 Ανέβασε το Excel αρχείο σου</h3>", unsafe_allow_html=True)
st.markdown("<p style='color: #666;'>Το αρχείο πρέπει να έχει φύλλο 'schedule' με στήλες: Description, Start Time, End Time</p>", unsafe_allow_html=True)
uploaded_file = st.file_uploader(
    "Επιλογή αρχείου", 
    type=['xlsx', 'xls'],
    help="Το αρχείο πρέπει να έχει φύλλο 'schedule' με στήλες: Description, Start Time, End Time",
    label_visibility="collapsed"
)
st.markdown("</div>", unsafe_allow_html=True)
st.markdown("<br>", unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Show available sheets
        available_sheets = excel_file.sheet_names
        
        # Sidebar for controls
        st.sidebar.markdown("<h2 style='color: #2c3e50;'>⚙️ Ρυθμίσεις</h2>", unsafe_allow_html=True)
        
        # Sheet selection
        if 'schedule' in [s.lower() for s in available_sheets]:
            default_sheet = [s for s in available_sheets if s.lower() == 'schedule'][0]
            default_index = available_sheets.index(default_sheet)
        else:
            default_index = 0
        
        selected_sheet = st.sidebar.selectbox(
            "Επιλογή Φύλλου:",
            available_sheets,
            index=default_index
        )
        
        # Read the selected sheet
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)
        
        st.success(f"✅ Φύλλο '{selected_sheet}' φορτώθηκε επιτυχώς με {len(df)} γραμμές!")
        
        # Show data preview
        with st.expander("📋 Προεπισκόπηση Δεδομένων"):
            st.dataframe(df.head(10), use_container_width=True)
        
        # Check required columns
        required_cols = ['Description', 'Start Time', 'End Time']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            st.error(f"❌ Λείπουν οι στήλες: {', '.join(missing_cols)}")
            st.info(f"📌 Διαθέσιμες στήλες: {', '.join(df.columns.tolist())}")
        else:
            # Convert to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'])
            df['End Time'] = pd.to_datetime(df['End Time'])
            
            # Filter out rows with missing data
            df = df.dropna(subset=['Description', 'Start Time', 'End Time'])
            
            # Add unique ID for each row
            df['uniqueId'] = range(1, len(df) + 1)
            df['displayLabel'] = df['uniqueId'].astype(str) + '. ' + df['Description'].astype(str)
            
            # Sorting
            sort_order = st.sidebar.radio(
                "Ταξινόμηση κατά End Time:",
                ["Αύξουσα", "Φθίνουσα"]
            )
            
            ascending = True if sort_order == "Αύξουσα" else False
            df_sorted = df.sort_values('End Time', ascending=ascending).reset_index(drop=True)
            
            # Update display labels after sorting
            df_sorted['displayLabel'] = range(1, len(df_sorted) + 1)
            df_sorted['displayLabel'] = df_sorted['displayLabel'].astype(str) + '. ' + df_sorted['Description'].astype(str)
            
            # Shift filter
            if 'Shift' in df.columns:
                shifts = ['Όλα'] + sorted(df_sorted['Shift'].dropna().unique().tolist())
                selected_shift = st.sidebar.selectbox("Φίλτρο Shift:", shifts)
                if selected_shift != 'Όλα':
                    df_sorted = df_sorted[df_sorted['Shift'] == selected_shift]
            
            # Calculate duration - Fix for Timedelta serialization
            df_sorted['Duration_hours'] = (df_sorted['End Time'] - df_sorted['Start Time']).dt.total_seconds() / 3600
            
            # Statistics with better styling
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
            
            # Create Gantt Chart with high contrast colors
            fig = go.Figure()
            
            # High contrast color palette
            colors = ['#2E86AB', '#A23B72', '#F18F01', '#C73E1D', '#6A4C93', 
                      '#06A77D', '#D90368', '#F08700', '#0E9594', '#8B2635']
            
            for idx, row in df_sorted.iterrows():
                color = colors[idx % len(colors)]
                
                # Hover text
                hover_text = f"<b>{row['Description']}</b><br>"
                hover_text += f"Start: {row['Start Time'].strftime('%d/%m/%Y %H:%M')}<br>"
                hover_text += f"End: {row['End Time'].strftime('%d/%m/%Y %H:%M')}<br>"
                hover_text += f"Duration: {row['Duration_hours']:.2f} ώρες<br>"
                
                if 'Shift' in df.columns and pd.notna(row['Shift']):
                    hover_text += f"Shift: {row['Shift']}<br>"
                if 'Qnt' in df.columns and pd.notna(row['Qnt']):
                    hover_text += f"Quantity: {row['Qnt']}<br>"
                if 'Capacity/hr' in df.columns and pd.notna(row['Capacity/hr']):
                    hover_text += f"Capacity/hr: {row['Capacity/hr']}<br>"
                if 'Prod. Time' in df.columns and pd.notna(row['Prod. Time']):
                    hover_text += f"Prod. Time: {row['Prod. Time']}<br>"
                
                fig.add_trace(go.Bar(
                    x=[pd.Timedelta(hours=row['Duration_hours'])],
                    y=[row['displayLabel']],
                    base=row['Start Time'],
                    orientation='h',
                    marker=dict(
                        color=color, 
                        line=dict(color='white', width=2)
                    ),
                    name=row['Description'],
                    hovertemplate=hover_text + '<extra></extra>',
                    showlegend=False
                ))
            
            # Layout with better contrast
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
                    title_font=dict(size=14, color='#1e1e1e', family='Arial')
                ),
                yaxis=dict(
                    autorange='reversed',
                    categoryorder='array',
                    categoryarray=df_sorted['displayLabel'].tolist(),
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='#d0d0d0',
                    tickfont=dict(size=11, color='#1e1e1e'),
                    title_font=dict(size=14, color='#1e1e1e', family='Arial')
                ),
                height=max(600, len(df_sorted) * 40),
                hovermode='closest',
                plot_bgcolor='#fafafa',
                paper_bgcolor='white',
                margin=dict(l=300, r=50, t=100, b=100),
                font=dict(color='#1e1e1e')
            )
            
            # Display chart
            st.plotly_chart(fig, use_container_width=True)
            
            # Download filtered data
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            with col2:
                csv = df_sorted.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Κατέβασε CSV",
                    data=csv,
                    file_name='gantt_schedule.csv',
                    mime='text/csv',
                )
        
    except Exception as e:
        st.error(f"❌ Σφάλμα: {str(e)}")
        st.info("Βεβαιώσου ότι το αρχείο είναι έγκυρο Excel και έχει τις σωστές στήλες.")

else:
    # Instructions with better visibility
    st.markdown("""
    <div style='background-color: #d1ecf1; padding: 25px; border-radius: 10px; border-left: 5px solid #0c5460;'>
        <h3 style='color: #0c5460; margin-top: 0;'>📝 Οδηγίες Χρήσης</h3>
        <p style='color: #0c5460; font-size: 16px; line-height: 1.6;'>
            <strong>1.</strong> Ανέβασε αρχείο Excel (.xlsx) που περιέχει φύλλο με όνομα <strong>"schedule"</strong><br>
            <strong>2.</strong> <strong>Απαραίτητες στήλες:</strong><br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Description</code> - Περιγραφή ενέργειας/υλικού<br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Start Time</code> - Ημερομηνία και ώρα έναρξης<br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>End Time</code> - Ημερομηνία και ώρα λήξης<br>
            <strong>3.</strong> <strong>Προαιρετικές στήλες:</strong><br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Shift</code> - Βάρδια (Morning, Evening, Night)<br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Qnt</code> - Ποσότητα<br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Capacity/hr</code> - Χωρητικότητα ανά ώρα<br>
            &nbsp;&nbsp;&nbsp;&nbsp;• <code>Prod. Time</code> - Χρόνος παραγωγής
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("<br>", unsafe_allow_html=True)
    
    # Features box
    st.markdown("""
    <div style='background-color: white; padding: 25px; border-radius: 10px; border: 1px solid #e0e0e0;'>
        <h3 style='color: #2c3e50; margin-top: 0;'>✨ Χαρακτηριστικά</h3>
        <ul style='color: #2c3e50; font-size: 15px; line-height: 1.8;'>
            <li>📊 Διαδραστικό Gantt Chart</li>
            <li>🔍 Hover για λεπτομέρειες</li>
            <li>📈 Αυτόματες στατιστικές</li>
            <li>🎯 Φίλτρα και ταξινόμηση</li>
            <li>📥 Export σε CSV</li>
            <li>🎨 Κάθε ενέργεια με μοναδικό χρώμα</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    # Sample data
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

# Footer
st.markdown("---")
st.markdown(
    "<div style='text-align: center; color: #666; font-size: 14px;'>Production Gantt Chart | Powered by Streamlit & Plotly</div>",
    unsafe_allow_html=True
)
