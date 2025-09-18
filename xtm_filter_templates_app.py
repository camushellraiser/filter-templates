import streamlit as st
import json
import plotly.express as px
import pandas as pd
from io import BytesIO
from openpyxl.styles import Font, PatternFill
from streamlit_agraph import agraph, Node, Edge, Config

# -------------------------------
# Cargar templates desde el JSON local
# -------------------------------
def load_templates():
    with open("Filter Templates.json", "r", encoding="utf-8") as f:
        return json.load(f)

templates = load_templates()

# -------------------------------
# Título
# -------------------------------
st.title("📖 XTM Filter Template Guide")
st.markdown("Interactive guide to help Project Managers choose the correct filter template in **XTM**.")

# -------------------------------
# Sidebar: filtros
# -------------------------------
st.sidebar.header("🔍 Filters")
search_term = st.sidebar.text_input("Search by name/keyword", "")
categories = sorted(list(set([t.get("category", "Uncategorized") for t in templates])))
selected_category = st.sidebar.selectbox("Category", ["All"] + categories, index=0)

# -------------------------------
# Filtrado de templates
# -------------------------------
filtered = []
for t in templates:
    if search_term.lower() in t["name"].lower() or search_term.lower() in t.get("description", "").lower():
        if selected_category == "All" or t.get("category", "Uncategorized") == selected_category:
            filtered.append(t)

st.subheader("📂 Filter Templates")
st.markdown(f"**Showing {len(filtered)} templates**")

# -------------------------------
# Mostrar resultados (expansores)
# -------------------------------
if not filtered:
    st.error("No results found. Try adjusting filters.")
else:
    for t in filtered:
        with st.expander(f"📌 {t['name']}"):
            st.write(f"**Description:** {t.get('description','')}")
            st.write(f"**Recommended Usage:** {t.get('recommended_usage','')}")
            st.write(f"**Suggested Use:** {t.get('suggested_use','')}")
            st.write(f"**Category:** {t.get('category','Uncategorized')}")

# -------------------------------
# Función para exportar a Excel con estilo
# -------------------------------
def to_excel_styled(dataframe: pd.DataFrame) -> BytesIO:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # --- Sheet 1: Templates ---
        dataframe.to_excel(writer, index=False, sheet_name="Templates")
        ws1 = writer.sheets["Templates"]

        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for cell in ws1[1]:
            cell.fill = header_fill
            cell.font = header_font

        for col in ws1.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws1.column_dimensions[col_letter].width = max_length + 2

        # --- Sheet 2: Stats by Category ---
        stats = dataframe.groupby("category").size().reset_index(name="Count")
        stats.to_excel(writer, index=False, sheet_name="Stats by Category")
        ws2 = writer.sheets["Stats by Category"]

        for cell in ws2[1]:
            cell.fill = header_fill
            cell.font = header_font

        for col in ws2.columns:
            max_length = 0
            col_letter = col[0].column_letter
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            ws2.column_dimensions[col_letter].width = max_length + 2

    output.seek(0)
    return output

# -------------------------------
# Exportar resultados
# -------------------------------
if filtered:
    df = pd.DataFrame(filtered)

    st.download_button(
        label="📥 Download as CSV",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="filter_templates.csv",
        mime="text/csv"
    )

    excel_file = to_excel_styled(df)
    st.download_button(
        label="📥 Download as Excel (styled + stats)",
        data=excel_file,
        file_name="filter_templates.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# -------------------------------
# Gráfico: distribución por categoría
# -------------------------------
st.subheader("📊 Templates by Category")
df_plot = [{"Category": t.get("category","Uncategorized")} for t in templates]
fig = px.sunburst(df_plot, path=["Category"], title="Distribution of Templates by Category")
st.plotly_chart(fig, use_container_width=True)

# -------------------------------
# Flowchart dinámico desde JSON
# -------------------------------
st.subheader("🌳 Decision Flowchart (auto-generated)")

nodes = [Node(id="Content", label="Content Type", size=25)]
edges = []

cats = {}
for t in templates:
    cat = t.get("category", "Uncategorized")
    if cat not in cats:
        nodes.append(Node(id=cat, label=cat, size=20))
        edges.append(Edge(source="Content", target=cat))
        cats[cat] = []
    cats[cat].append(t["name"])

for cat, items in cats.items():
    for name in items:
        nodes.append(Node(id=name, label=name, size=15))
        edges.append(Edge(source=cat, target=name))

config = Config(width=850, height=600, directed=True, physics=True, hierarchical=True)
agraph(nodes=nodes, edges=edges, config=config)
