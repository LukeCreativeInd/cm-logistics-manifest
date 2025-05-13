import streamlit as st

st.set_page_config(page_title="CM Logistics Manifest Generator", layout="centered")

st.image("CM_Logistics_Top_Logo.png", use_container_width=True)

st.title("CM Logistics Manifest Generator")
st.markdown("### Which customer group do you want to create a manifest for?")

col1, col2, col3 = st.columns(3)

with col1:
    if st.button("Clean Eats Australia"):
        from clean_eats import run_clean_eats
        run_clean_eats()

with col2:
    if st.button("Made Active"):
        from made_active import run_made_active
        run_made_active()

with col3:
    if st.button("Elite Meals"):
        from elite_meals import run_elite_meals
        run_elite_meals()
