import streamlit as st
import plotly.graph_objects as go
import pandas as pd
from datetime import datetime

# Page config
st.set_page_config(page_title="Production Gantt Chart", layout="wide", page_icon="📊")

# Custom CSS
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        font-weight: 600;
        border: none;
        padding: 12px;
        border-radius: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.title("📊 Production Schedule - Gantt Chart")
st.markdown("---")

# File uploader
uploaded_file = st.file_uploader(
    "Ανέβασε το Excel αρχείο σου (με φύλλο 'schedule')", 
    type=['xlsx', 'xls'],
    help="Το αρχείο πρέπει να έχει φύλλο 'schedule' με στήλες: Description, Start Time, End Time"
)

if uploaded_file is not None:
    try:
        # Read Excel file
        excel_file = pd.ExcelFile(uploaded_file)
        
        # Show available sheets
        available_sheets = excel_file.sheet_names
        
        # Sidebar for controls
        st.sidebar.header("⚙️ Ρυθμίσεις")
        
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
            st.info("Διαθέσιμες στήλες: " + ", ".join(df.columns.tolist()))
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
            
            # Calculate duration
            df_sorted['Duration'] = df_sorted['End Time'] - df_sorted['Start Time']
            df_sorted['Duration_hours'] = df_sorted['Duration'].dt.total_seconds() / 3600
            
            # Statistics
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
            
            # Create Gantt Chart
            fig = go.Figure()
            
            colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                      '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
            
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
                    x=[row['Duration']],
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
                    'font': {'size': 24, 'color': '#2c3e50'}
                },
                xaxis_title='Χρονική Περίοδος',
                yaxis_title='Ενέργειες / Υλικά',
                xaxis=dict(
                    type='date',
                    tickformat='%d/%m %H:%M',
                    tickangle=-45,
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='lightgray'
                ),
                yaxis=dict(
                    autorange='reversed',
                    categoryorder='array',
                    categoryarray=df_sorted['displayLabel'].tolist(),
                    showgrid=True,
                    gridwidth=1,
                    gridcolor='lightgray'
                ),
                height=max(600, len(df_sorted) * 40),
                hovermode='closest',
                plot_bgcolor='white',
                paper_bgcolor='white',
                margin=dict(l=300, r=50, t=80, b=100)
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
    # Instructions
    st.info("👆 Ανέβασε το Excel αρχείο σου για να δημιουργηθεί το Gantt Chart")
    
    st.markdown("""
    ### 📝 Οδηγίες:
    
    1. **Ανέβασε αρχείο Excel (.xlsx)** που περιέχει φύλλο με όνομα **"schedule"**
    2. **Απαραίτητες στήλες:**
       - `Description` - Περιγραφή ενέργειας/υλικού
       - `Start Time` - Ημερομηνία και ώρα έναρξης
       - `End Time` - Ημερομηνία και ώρα λήξης
    3. **Προαιρετικές στήλες:**
       - `Shift` - Βάρδια (Morning, Evening, Night)
       - `Qnt` - Ποσότητα
       - `Capacity/hr` - Χωρητικότητα ανά ώρα
       - `Prod. Time` - Χρόνος παραγωγής
    
    ### ✨ Χαρακτηριστικά:
    - 📊 Διαδραστικό Gantt Chart
    - 🔍 Hover για λεπτομέρειες
    - 📈 Αυτόματες στατιστικές
    - 🎯 Φίλτρα και ταξινόμηση
    - 📥 Export σε CSV
    - 🎨 Κάθε ενέργεια με μοναδικό χρώμα
    """)
    
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
    "<div style='text-align: center; color: gray;'>Production Gantt Chart | Powered by Streamlit & Plotly</div>",
    unsafe_allow_html=True
)